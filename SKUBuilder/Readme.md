# Caddie SKU Load & Match Processor

## Overview

The Caddie SKU Load & Match Processor is a Streamlit web application designed to automate the generation of SKU (Stock Keeping Unit) loading files for the Caddie system. It takes two input files—a list of required SKUs and a comprehensive pricing sheet—and processes them to create a structured Excel file ready for system import.

The application builds two types of records for each input SKU, corresponding to 'HORN' and 'AMT' insurance codes, and automatically generates the associated sub-component ('SC') rows based on predefined asset-to-subasset mappings. It performs various calculations for pricing, fees, and commissions, and formats the output according to specific system requirements.

## Features

-   **Web-Based Interface:** Built with Streamlit for an easy-to-use, interactive experience.
-   **Dual File Input:** Accepts a 'SKU Needed' file and a 'Caddie Pricing Sheet' in either Excel (.xlsx, .xls) or CSV (.csv) formats.
-   **Automated Row Generation:** For each input SKU, it generates two primary 'N' (New) type rows (one for 'HORN' logic and one for 'AMT' logic).
-   **Sub-Component Processing:** Automatically creates the corresponding 'SC' (Sub-Component) rows for each 'N' row based on a hardcoded asset map.
-   **Complex Financial Calculations:** Precisely calculates various financial fields like `LOSS COST`, `RESERVE`, `UW_FEE`, `PREMIUM`, and `IWW MARKUP` using the `decimal` library to ensure accuracy.
-   **Dynamic Data Mapping:** Populates fields such as `Coverage\\SKU Code` and `Coverage\\SKU Description` by combining data from the input files.
-   **Limit of Liability (LoL) Determination:** Automatically assigns the correct LoL amount based on the SKU prefix.
-   **Progress Tracking:** Features an enhanced real-time progress bar that displays the processing speed (SKUs/sec), elapsed time, and estimated time remaining.
-   **Data Cleaning and Formatting:** The final output is meticulously cleaned and formatted to meet system import standards, including:
    -   Applying specific column visibility rules for 'SC' rows.
    -   Ensuring correct data types for integers and floats.
    -   Sorting the final data for consistency.
-   **Error Handling:** Provides clear feedback on missing files, columns, or data processing errors.
-   **Direct Download:** Allows the user to download the final processed data as a formatted Excel file with a dynamically generated name (e.g., `AGENTNUMBER_SKU_LOAD_MMDDYYYY.xlsx`).

## How to Run the Application

This is a Streamlit application. To run it locally, you will need to have Python and the required libraries installed.

1.  **Save the Code:** Save the provided code as a Python file (e.g., `app.py`).

2.  **Install Dependencies:** Open your terminal or command prompt and install the necessary libraries. The primary dependencies are `streamlit` and `pandas`. `xlsxwriter` is also used for Excel export.

    ```bash
    pip install streamlit pandas xlsxwriter
    ```

3.  **Run the App:** Navigate to the directory where you saved the file and run the following command:

    ```bash
    streamlit run app.py
    ```

4.  **Use the Application:** Your web browser will open a new tab with the running application.
    -   **Tab 1: Upload Files & Configure:**
        1.  Upload the 'Caddie\_SKU\_Needed' file.
        2.  Upload the 'Caddie\_Pricing Sheet' file.
        3.  Select a 'Start Date' for the generated records.
    -   **Tab 2: Generate & Download Results:**
        1.  Click the "Generate Output Data" button to start the process.
        2.  Monitor the progress bar and processing metrics.
        3.  Once complete, a summary and a preview of the output data will be displayed.
        4.  Click the "Download Excel Results" button to save the final file.

## Input Files

The application requires two input files:

### 1. Caddie SKU Needed File

This file contains the list of SKUs that need to be processed. It must contain the following columns (column names can be slightly different, as the script looks for keywords):

-   **SKU / SKU Coverage Code:** The unique identifier for the product.
-   **Agent Number:** The agent associated with the SKU.
-   **Dealer Group Number:** The dealer associated with the SKU.

### 2. Caddie Pricing Sheet File

This file is a comprehensive data source containing pricing and detailed attributes for each SKU. It should contain, but is not limited to, the following columns:

-   `SKU`
-   `Plan Name`
-   `Term`
-   `Loss Cost`
-   `Reserve`
-   `UW Fee`
-   `HIC Cost`
-   `Labor Rate`
-   `Trip Charge`
-   `Coverage Type`
-   `Region`
-   `Trade`
-   `Performance Level`
-   `IWW ASSET NAME` (Used for mapping to sub-components)

## Output File

The application generates a single Excel file (`.xlsx`) with the following characteristics:

-   **File Name:** Dynamically named in the format: `<AgentNumber>_SKU_LOAD_<Date>.xlsx`.
-   **Sheet Name:** `Caddie SKUs`.
-   **Content:** Contains the fully processed and formatted SKU data, including both 'N' and 'SC' rows, with all necessary fields populated for system import.
-   **Formatting:** Columns are ordered according to a predefined template, and numeric values are correctly formatted (e.g., integers do not have trailing decimals).

## Core Logic and Constants

The script contains several hardcoded business rules and constants:

-   **Asset to Subasset Map (`SUBCATEGORIES_CSV_CONTENT`):** A predefined CSV-like string that maps primary assets (e.g., `Split System AC`) to their sub-components (e.g., `Air Handler`, `Condensing Unit AC`). This is crucial for generating the 'SC' rows.
-   **Inherited Columns (`SC_INHERITED_COLUMNS`):** A list of columns whose values are copied directly from a parent 'N' row to its child 'SC' rows.
-   **Limit of Liability (`LOL_BY_PREFIX`):** A dictionary that defines the Limit of Liability amount based on SKU prefixes (`HSYS`, `HACC`, etc.).
-   **Financial Calculations:** The script uses specific percentages and formulas to derive values for `HIC CONTRACT FEE`, `IWW MARKUP`, and `CEDING COMMISSION` for both 'HORN' and 'AMT' logic.
-   **Hardcoded Expected Frequency:** A fixed `decimal` value (`0.005`) is used in internal severity calculations.
