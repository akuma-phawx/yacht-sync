"""Extract invoice totals from PDFs with the Gemini API."""

from __future__ import annotations

import re
import time
from pathlib import Path

from google import genai
from google.genai import types
from pydantic import BaseModel, field_validator
from rich.console import Console

from config import Settings
from services.matcher import ExtractedInvoice, parse_amount

_console = Console()
# Free tier for this model allows 5 requests per minute.
_MIN_INTERVAL_SECONDS = 15.0
_MAX_ATTEMPTS = 6

_PROMPT = (
    "This is a supplier invoice PDF. It may be in Greek and it may be a scan. "
    "Extract the printed invoice fields. "
    "total_amount is the final amount payable (grand total), as a number. "
    "net_amount is the amount before VAT. vat_amount is the VAT amount. "
    "invoice_number is the invoice identifier. invoice_date is YYYY-MM-DD when readable. "
    "Use null for any field that is not visible. Do not invent a total that is not printed."
)


class InvoiceExtraction(BaseModel):
    supplier_name: str | None = None
    invoice_number: str | None = None
    invoice_date: str | None = None
    currency: str | None = None
    net_amount: float | None = None
    vat_amount: float | None = None
    total_amount: float
    description: str | None = None

    @field_validator(
        "supplier_name",
        "invoice_number",
        "invoice_date",
        "currency",
        "description",
        mode="before",
    )
    @classmethod
    def _clean_text(cls, value: object) -> str | None:
        if value is None:
            return None
        text = str(value).strip()
        return text or None

    @field_validator("net_amount", "vat_amount", "total_amount", mode="before")
    @classmethod
    def _clean_number(cls, value: object) -> float | None:
        if value is None or value == "":
            return None
        if isinstance(value, bool):
            raise ValueError("amount must be numeric")
        if isinstance(value, (int, float)):
            return float(value)
        return float(parse_amount(value))


class PdfExtractor:
    def __init__(self, client: genai.Client, model: str, min_interval: float = _MIN_INTERVAL_SECONDS) -> None:
        self._client = client
        self._model = model
        self._min_interval = min_interval
        self._next_request_at = 0.0

    @classmethod
    def from_settings(cls, settings: Settings) -> PdfExtractor:
        client = genai.Client(api_key=settings.gemini_api_key)
        return cls(client, settings.gemini_model)

    def extract(self, path: Path) -> InvoiceExtraction:
        data = path.read_bytes()
        last_error: BaseException | None = None
        for attempt in range(1, _MAX_ATTEMPTS + 1):
            self._wait_for_slot()
            try:
                extraction = self._generate(data)
            except Exception as exc:
                self._reserve_slot()
                last_error = exc
                if not _is_transient(exc) or attempt == _MAX_ATTEMPTS:
                    raise RuntimeError(f"Could not extract {path.name}: {exc}") from exc
                delay = _retry_delay(exc)
                reason = "rate limit" if _is_rate_limit(exc) else "temporary error"
                _console.print(
                    f"[yellow]Gemini {reason} on {path.name}. Waiting {delay:.0f}s "
                    f"({attempt}/{_MAX_ATTEMPTS}).[/yellow]"
                )
                time.sleep(delay)
                continue
            self._reserve_slot()
            return extraction
        raise RuntimeError(f"Could not extract {path.name}: {last_error}")

    def _wait_for_slot(self) -> None:
        remaining = self._next_request_at - time.monotonic()
        if remaining > 0:
            time.sleep(remaining)

    def _reserve_slot(self) -> None:
        self._next_request_at = time.monotonic() + self._min_interval

    def _generate(self, data: bytes) -> InvoiceExtraction:
        response = self._client.models.generate_content(
            model=self._model,
            contents=[
                types.Part.from_bytes(data=data, mime_type="application/pdf"),
                _PROMPT,
            ],
            config=types.GenerateContentConfig(
                temperature=0,
                response_mime_type="application/json",
                response_schema=InvoiceExtraction,
            ),
        )
        parsed = getattr(response, "parsed", None)
        if isinstance(parsed, InvoiceExtraction):
            extraction = parsed
        elif isinstance(parsed, dict):
            extraction = InvoiceExtraction.model_validate(parsed)
        else:
            text = getattr(response, "text", None) or ""
            if not text.strip():
                raise RuntimeError("Gemini returned an empty extraction")
            extraction = InvoiceExtraction.model_validate_json(text)
        return extraction


def to_extracted_invoice(filename: str, extraction: InvoiceExtraction) -> ExtractedInvoice:
    return ExtractedInvoice(
        source_name=filename,
        total_amount=parse_amount(extraction.total_amount),
        invoice_number=extraction.invoice_number,
        net_amount=None if extraction.net_amount is None else parse_amount(extraction.net_amount),
        vat_amount=None if extraction.vat_amount is None else parse_amount(extraction.vat_amount),
        supplier_name=extraction.supplier_name,
        invoice_date=extraction.invoice_date,
        currency=extraction.currency,
        description=extraction.description,
    )


def retry_after_seconds(exc: BaseException) -> float | None:
    text = str(exc)
    match = re.search(r"retry in ([0-9]+(?:\.[0-9]+)?)s", text, re.IGNORECASE)
    if match:
        return float(match.group(1))
    match = re.search(r"retryDelay['\"]?\s*[:=]\s*['\"]?(\d+(?:\.\d+)?)s", text, re.IGNORECASE)
    if match:
        return float(match.group(1))
    return None


def _retry_delay(exc: BaseException) -> float:
    suggested = retry_after_seconds(exc)
    if _is_rate_limit(exc):
        return max(suggested or 0, 45) + 1
    if suggested is not None:
        return suggested + 1
    return 4


def _is_rate_limit(exc: BaseException) -> bool:
    code = getattr(exc, "code", None)
    if code is None:
        code = getattr(exc, "status_code", None)
    if code == 429:
        return True
    text = str(exc)
    return "RESOURCE_EXHAUSTED" in text or "generate_content_free_tier_requests" in text


def _is_transient(exc: BaseException) -> bool:
    if isinstance(exc, (TimeoutError, ConnectionError)):
        return True
    code = getattr(exc, "code", None)
    if code is None:
        code = getattr(exc, "status_code", None)
    if code in {408, 429, 500, 502, 503, 504}:
        return True
    return type(exc).__name__ in {
        "ServerError",
        "APIConnectionError",
        "ConnectError",
        "ReadTimeout",
        "RemoteProtocolError",
    }
