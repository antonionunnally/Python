# ARW Zeros Fixer (`ARW_zeros_v9.2.py`)

A desktop GUI tool (Tkinter + CustomTkinter) for cleaning up ARW dealer acknowledgment CSV files before they're reloaded or sent downstream. It fixes zero-value placeholders, standardizes state fields, and normalizes contract price/refund values.

## Requirements

```
pip install customtkinter
```
(`tkinter` ships with standard Python on Windows.)

## Running

```
python ARW_zeros_v9.2.py
```

This opens a GUI window titled **"ARW File Processor"**.

## Usage

1. **Input File** — click **Browse** and select the source CSV. The output path is auto-filled as `{input_name}_Fix.csv` in the same folder (editable via **Output File** → **Browse**).
2. Click **Process File** to run the transformations (see below) and write the output CSV.
3. Click **Processed File Location** to open the output folder in File Explorer.
4. Click **Clear All** to reset the form and start over.
5. Status and a running log of actions/errors are shown at the bottom of the window.

## What It Does

For every record in the input CSV, four transformations are applied in order:

1. **Zero-to-empty conversion** — if any of the following fields equal the literal string `"0"`, they're cleared to an empty string:
   `Cancellation_Date`, `Cancel_Reason_Code`, `Business_Name`, `Customer_Address_2`, `Customer_Phone`, `Customer_Email`, `Sales_Ticket_Number`, `Manufacturer_Name`, `Model_Number`, `Model_Name`, `Serial_Number`, `Product_Condition`, `Contract_Note`, `Renewal_Contract_Number`, `Change_Flag`, `Original_Contract_Number`

2. **State standardization** (`Customer_State`) —
   - Already a 2-letter uppercase abbreviation → left as-is.
   - A full state name (e.g. `"kentucky"`) → converted to its 2-letter abbreviation (`"KY"`) via a built-in state name lookup table (all 50 states + DC).
   - Anything else → uppercased as-is.

3. **Contract price fix** (`Contract_Price_Retail_Cost`) — if the value is numerically `0` (or `0.0`, etc.), it's changed to `"1"`. Non-numeric values are left unchanged.

4. **Refund amount fix** (`Contract_Refund_Amount`) — if `Transaction_Reason` is `"1"`, `"2"`, or `"5"` (Sales/Payments), `Contract_Refund_Amount` is forced to `"0"`.

## Input/Output Notes

- Input CSV is read with `utf-8-sig` encoding (handles BOM from Excel exports).
- Output CSV is written with `utf-8` encoding, same column order as the input.
- All records are processed in memory; no rows are dropped or added — only field values are modified.

## Logging

Every run appends to `arw_processor.log` (INFO level) in the working directory, and mirrors log messages to the console and the in-app log panel. Includes record counts read/processed/written and any errors encountered.

## Code Structure

| Class | Responsibility |
|---|---|
| `FileHandler` | Static CSV read/write helpers |
| `RecordProcessor` | Static per-record transformation methods (the 4 fixes above) |
| `ARWFileProcessor` | Orchestrates read → process → write over the full record set |
| `ARWFileProcessorGUI` | CustomTkinter UI wiring (file pickers, buttons, status/log) |

## Customizing

- To change which fields get zero-cleared, edit `ZERO_TO_EMPTY_FIELDS`.
- To change which transaction reason codes force a zero refund, edit `ZERO_REFUND_REASON_CODES`.
- State abbreviation mappings live in `STATE_ABBREVIATIONS`.
