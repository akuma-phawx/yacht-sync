# Yacht invoice helper

This folder is a small program. You give it one Excel file of outstanding boat payments and the invoice PDFs that go with those payments. It reads the invoices, matches each payment to the right PDF by the amount, and writes a spreadsheet file called `output.csv` that you can open in Excel.

It does not upload anything to Google Drive. It does not write to Google Sheets. It only creates a file on this computer.

You do not need to understand programming to use it. You do need to follow the setup once, then the same short command every time you have a new batch of invoices.

---

## What you put in, and what you get out

You put files in the folder named `input_files`:

- Exactly **one** Excel file. It must end in `.xlsx`. This is the outstanding-payments list for one boat.
- As many invoice PDFs as you have. They must end in `.pdf`. The names can be anything. Greek names are fine. You replace these files every time you run a new batch.

You get one result file, `output.csv`, in the main `yacht-sync` folder (next to this README, not inside `input_files`). Each run **replaces** that file. If you want to keep an old result, rename it before you run again, for example `output-18-sep.csv`.

---

## The Excel file it expects

The first row of the Excel file must be the column titles. The program looks for these titles by name, so the columns can be in any order, but the names must match:

- `SUPPLIER NAME`
- `INVOICE NUMBER`
- `DESCRIPTION`
- `AMOUNT ORIGINAL`
- `CURRENCY ORIGIN`
- `CREATE DATE`
- `INVOICE/QUESTION DATE`
- `WORK/ORDER STATUS`
- `BOAT CATEGORY`
- `BOAT SUBCATEGORY`

Other columns in that file (bank details, IBAN, crew member, and so on) are ignored. They are not copied into the result.

The boat name is **not** a column in the Excel file. The program reads it from the Excel file's name, when the name looks like this:

`World_s_End-18-Sep-2026-06_38.xlsx`

That means:

- Boat: World's End (`_s_` becomes `'s`, other `_` become spaces)
- The date in the name is only used to recognise this pattern. It is not a column in the CSV.

If the Excel file has a different kind of name, the program stops and asks for the boat name. See [When the Excel file has a plain name](#when-the-excel-file-has-a-plain-name).

There must be **one** `.xlsx` file in `input_files`. Zero files, or two files, is an error. Close the Excel file in Excel before you run the program, so Windows is not holding it open (those temporary files start with `~$` and are ignored, but a file that is open can still cause trouble).

---

## What the program does, in order

1. It looks in `input_files` and finds the one Excel file and every PDF.
2. It reads each payment row from the Excel file.
3. It sends each PDF to Google's Gemini service, which reads the scan and returns the invoice total, the amount before VAT, and the VAT. The PDFs are usually scans, so a normal computer cannot read the numbers off them by itself.
4. It waits about 15 seconds between invoices. That is on purpose. The free Gemini API only allows a few reads per minute. A batch of 15 invoices takes several minutes. Leave the window open.
5. It matches payments to PDFs by amount. See the next section.
6. It writes `output.csv` and prints a table in the black window: matched, unmatched, and any PDFs that were left over.

---

## How a payment is matched to an invoice

The important number on the Excel row is `AMOUNT ORIGINAL`. The important number on the PDF is the final amount to pay (the grand total).

- If one unused PDF has the same total, within 1 cent, that PDF belongs to that row.
- If two PDFs have the same total, the program prefers the one whose invoice number appears in the Excel cell `INVOICE NUMBER`, or in the PDF's file name.
- If it still cannot tell them apart, it leaves the row unmatched rather than guessing.
- Some Excel rows list several invoice numbers in one cell, for example `524/525/585/586` or `32 & 266` or `940/316`. That means one payment is the sum of several invoices. The program looks for that many PDFs whose totals add up to `AMOUNT ORIGINAL`, again within 1 cent. It will not try this for more than 6 invoices in one cell.
- Each PDF is used at most once.
- A single PDF whose total equals the whole payment wins over a group of PDFs that add up to the same number.

**MATCHED** in the black window means a PDF (or a group of PDFs) was found.

**UNMATCHED** means no PDF total lined up. The row is still written to the CSV so you can see it. The Notes column ends with `UNMATCHED`, and the Invoice column is empty.

**Unused PDFs** are invoices that were read but did not fit any payment. Their names are printed at the end. They do not appear in the Invoice column.

---

## What each column in `output.csv` means

Open `output.csv` in Excel. The columns are the same ones used on the Master Ongoing Task Management sheet.

| Column | What is in it |
| --- | --- |
| Tasks | The Excel `DESCRIPTION` (for example "2 batteries 45AH"). |
| Supplier | The Excel `SUPPLIER NAME`. |
| Start Date | The Excel `CREATE DATE`, written as year-month-day, for example `2026-09-16`. |
| Delivery Date | The Excel `INVOICE/QUESTION DATE`, same date style. |
| Owner's Approval | Always empty. The Excel file has nothing to put here. |
| Status | The Excel `WORK/ORDER STATUS` (for example `Completed`). |
| Estimated Cost | The Excel `AMOUNT ORIGINAL`. |
| Net Price | The amount before VAT, read from the matched PDF. If several PDFs matched, this is their sum. If a matched PDF has no net amount, this cell is left empty. Empty when the row is unmatched. |
| VAT 24% | The VAT read from the matched PDF. If the PDF shows a net and a total but no VAT line, VAT is total minus net. If several PDFs matched, this is their sum. Empty if any matched PDF is missing a net amount, and empty when the row is unmatched. |
| Final Cost | The matched PDF total (or the sum, if several PDFs matched). If the row is unmatched, this is the Excel `AMOUNT ORIGINAL`. |
| Notes | Boat name, then category / subcategory, then the invoice number from Excel. Example: `World's End \| Maintenance / Electrician \| 3112`. Unmatched rows also end with `UNMATCHED`. |
| Invoice | The file name of the matched PDF. If several PDFs matched, the names are separated by a semicolon, for example `524.pdf; 525.pdf`. This is the name of the file in `input_files`, not a web link. |

---

## Gemini, your Google account, and the API key

This is the part that confuses people.

Paying for **Gemini Pro** (or Google AI Pro) on your personal Google account, and having a lot of Google Drive storage, does **not** pay for this program.

Those subscriptions are for the Gemini chat website and for Google products like Gmail and Drive. This program does not use that chat. It uses the **Gemini API**, which is a separate developer service. The program proves who it is with a key stored in a file named `.env`.

The error `free_tier` or `RESOURCE_EXHAUSTED` or `429` means that **API key** is on the free API plan. The free plan allows about 5 invoice reads per minute for the model this program uses (`gemini-3.8-flash`). The program already waits between invoices so a normal batch can finish. It will print a yellow line like `Waiting 46s` and then continue. Do not close the window.

If you want higher limits, you turn on **Cloud Billing** for the Google project that owns the API key, in [Google AI Studio](https://aistudio.google.com/). That billing is separate from the Gemini Pro subscription. Google's own note is here: [Gemini API billing](https://ai.google.dev/gemini-api/docs/billing).

---

## One-time setup on a Windows PC

Do this once on the computer that will run the program. If someone copied this folder from a Mac, ignore any folder named `.venv` that came with it. A `.venv` made on a Mac does not work on Windows. You will create a new one below.

### 1. Install Python

1. Go to [https://www.python.org/downloads/](https://www.python.org/downloads/) and download Python 3.11 or newer.
2. Run the installer.
3. On the first screen, tick **Add python.exe to PATH**. Then click Install.

Check it. Open **Command Prompt** (press the Windows key, type `cmd`, press Enter) and type:

```text
python --version
```

You should see something like `Python 3.13.x`. If Windows says it does not recognise `python`, the installer was not added to PATH. Run the installer again and tick that box.

### 2. Open the folder in Command Prompt

The folder that contains this README must be the folder you are "in". If the folder is on the Desktop and is named `yacht-sync`:

```text
cd %USERPROFILE%\Desktop\yacht-sync
```

`cd` means "change directory", which is "go into this folder". If the folder lives somewhere else, type `cd ` (with a space) and then drag the `yacht-sync` folder from File Explorer onto the black window. Press Enter.

You can check you are in the right place:

```text
dir
```

You should see `main.py`, `README.md`, and a folder named `input_files`.

### 3. Create the program's private Python and install its pieces

Copy these three lines, one at a time, and press Enter after each:

```text
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
```

The first line creates a folder named `.venv`. That is a private copy of Python for this program only. The second line downloads the libraries the program needs. It can take a few minutes. You only do this again if you move the project to a new computer.

### 4. Create the API key file

1. In a browser, sign in to Google and open [Google AI Studio API keys](https://aistudio.google.com/apikey).
2. Create an API key and copy it. It is a long string of letters and numbers. Treat it like a password.
3. In the `yacht-sync` folder, create a new text file named exactly `.env`

Windows will try to name it `.env.txt`. To stop that:

- Open Notepad.
- Paste this one line, with your real key after the equals sign and no spaces:

```text
GEMINI_API_KEY=paste-the-key-here
```

- Click File, Save As.
- Go to the `yacht-sync` folder.
- Set "Save as type" to **All files**.
- Set the file name to `.env`
- Save.

Do not put quotes around the key. Do not send this file to anyone. Do not put it in email or chat.

You do not need a second line. The program already uses the model `gemini-3.8-flash`. Only add a second line if someone tells you to change the model:

```text
GEMINI_MODEL=gemini-3.8-flash
```

---

## Every time you have a new batch

1. Empty the `input_files` folder, or replace what is in it.
2. Put in **one** `.xlsx` file and the invoice `.pdf` files for that batch. Nothing else is read (Word files, photos, and so on are ignored).
3. If you want to keep the previous result, rename `output.csv` first.
4. Open Command Prompt, go to the `yacht-sync` folder (the `cd` command from setup), and run:

```text
.venv\Scripts\python main.py --input input_files --output output.csv
```

5. Wait. You will see lines like `Extracting 3/15` and the PDF's name. A batch of 15 invoices often takes several minutes because of the wait between Gemini calls.
6. When it finishes, it prints a table and a line like `Wrote 16 row(s) to output.csv, 2 unmatched.`
7. Open `output.csv` in Excel.

If Excel shows strange characters in Greek file names: use Data, From Text/CSV, and choose **UTF-8** as the file origin. The file is already saved in a form Excel usually understands.

---

## When the Excel file has a plain name

If the file is named like `payments.xlsx` instead of `World_s_End-18-Sep-2026-06_38.xlsx`, add the boat name to the command. Keep the quotes if the name has a space or an apostrophe:

```text
.venv\Scripts\python main.py --input input_files --output output.csv --boat "World's End"
```

That name is what appears at the start of the Notes column.

---

## If this computer is a Mac

Setup is the same idea. In Terminal, go to the folder, then:

```text
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python main.py --input input_files --output output.csv
```

The `.env` file is created the same way: one line, `GEMINI_API_KEY=...`, saved in the `yacht-sync` folder.

---

## What the messages mean

**`Missing or blank in .env: GEMINI_API_KEY`**
The `.env` file is missing, is named `.env.txt`, is in the wrong folder, or the key line is empty. Fix the file and run the command again.

**`Expected exactly one .xlsx`**
`input_files` has no Excel file, or more than one. Leave only the one outstanding-payments file. Files named `.xlsx` count even if they are not the one you meant to use.

**`Could not read a boat and date from ...`**
The Excel file name does not look like `Boat_Name-18-Sep-2026-06_38.xlsx`. Run again with `--boat "The Boat Name"`.

**`Workbook is missing columns: ...`**
The first row of the Excel file does not have one of the titles listed above. Fix the title so it matches exactly, including spaces.

**`Row 5 is missing AMOUNT ORIGINAL`**
That payment row has no amount. Fill it in, or delete the row if it is junk.

**`Gemini rate limit ... Waiting 46s`**
The free API allowance for this minute is used up. The program is waiting and will continue by itself. Leave it alone. Closing the window throws away the work done so far, and the CSV is only written at the end.

**`Could not extract some-file.pdf`**
Gemini could not read that one PDF after several tries. Nothing is written for the batch. Check that the PDF opens normally, wait a minute, and run the command again.

**`UNMATCHED` in the table or in Notes**
The payment is in the CSV, but no invoice total matched it. Check the amount on the Excel row against the total printed on the PDF. A difference bigger than 1 cent will not match. A row that lists four invoice numbers needs four PDFs whose totals add up to that amount.

**`Unused PDFs: ...`**
Those invoices were read, but no payment row had a matching amount. They might be extras, or the Excel amount might not equal the invoice total.

---

## What not to share

- The file `.env` contains the API key. Anyone with that key can use your Gemini API allowance.
- The Excel file in `input_files` can contain bank account numbers and IBANs. Do not email that file around.
- `service_account.json`, if you see it, is an old Google login file. This program does not use it anymore. Do not send it to anyone. You can leave it where it is.

`output.csv` has supplier names, amounts, and descriptions. It does not copy the bank columns from the Excel file.

---

## The other things in this folder

You can ignore these. They are the program itself.

| Name | What it is |
| --- | --- |
| `main.py` | The command you run. It ties the steps together. |
| `input_files` | The drop folder. You change what is inside this. |
| `output.csv` | The result. Created when a run finishes. |
| `.env` | Your secret API key. Not part of the instructions you copy to other people. |
| `.venv` | The private Python installed in step 3. Do not copy this folder between a Mac and a Windows PC. |
| `requirements.txt` | The list of libraries installed in step 3. |
| `config.py` | Reads the API key from `.env`. |
| `services` | The parts that read Excel, read PDFs, match amounts, and write the CSV. |
| `tests` | Automatic checks for the matching rules. You never need to run these. |

---

## A normal run, from empty folder to CSV

1. Delete whatever is already in `input_files`.
2. Copy in `World_s_End-18-Sep-2026-06_38.xlsx` (or whatever this week's file is called) and the invoice PDFs.
3. Open Command Prompt in the `yacht-sync` folder.
4. Run `.venv\Scripts\python main.py --input input_files --output output.csv`
5. Wait until the table appears and the window says it wrote the rows.
6. Double-click `output.csv`.
