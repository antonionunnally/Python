import streamlit as st
import pandas as pd
import io
import subprocess
import os
import tempfile
import smtplib
import pythoncom
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import logging
from typing import Optional, List, Dict, Any

# Optional import for win32com - graceful fallback if not available
try:
    import win32com.client

    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False
    win32com = None

# Configure logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

# Constants for column names (helps avoid typos and makes code more readable)
TRANSFER_FLAG_COL = "Transfer_Flag"
# These columns are now unconditionally dropped for all files.
JOB_RUN_REGISTRATION_COL = "job_run_registration"
INCOMING_RECORD_GUID_COL = "incoming_record_guid"
ERROR_UPDATE_DATETIME_COL = "Error_Update_Datetime"

IS_IPAY_COL = "is_ipay"
ERROR_JOB_RUN_COL = "Error_Job_Run"
ERROR_SOURCE_COL = "Error_Source"
GENESIS_JOB_RUN_COL = "Genesis_Job_Run"
STANDARD_JOB_RUN_COL = "Standard_Job_Run"
ORIGINAL_CONTRACT_NUMBER_COL = "Original_Contract_Number"
SOURCE_FILENAME_COL = "Source_Filename"  # This column is NOT removed
# Constant for the column that Source_Filename needs to be before for COSIGN. This column is NOT removed.
INCOMING_CLIENT_FILENAME_COL = "incoming_client_filename"  # This column is NOT removed
CLIENT_ACTION_COL = "Client_Action"
ERROR_MESSAGE_COL = "Error_Message"
ERROR_TYPE_COL = "Error_Type"
TRANSACTION_REASON_COL = "Transaction_Reason"
IS_ERROR_COL = "isError"
CUSTOMER_FIRST_NAME_COL = "Customer_First_Name"
CUSTOMER_ADDRESS_1_COL = "Customer_Address_1"
CUSTOMER_CITY_COL = "Customer_City"
CUSTOMER_STATE_COL = "Customer_State"
CUSTOMER_ZIP_CODE_COL = "Customer_Zip_Code"
CUSTOMER_PHONE_COL = "Customer_Phone"
CUSTOMER_EMAIL_COL = "Customer_Email"
AGENT_NUMBER_COL = "Agent_Number"
AGENT_NAME_COL = "Agent_Name"

# NEW CONSTANTS for COSIGN-specific PII (Property-related PII)
PROPERTY_ADDRESS_COL = "Property_Address"
PROPERTY_CITY_COL = "Property_City"
PROPERTY_STATE_CODE_COL = "Property_State_Code"  # Corrected typo for consistency
PROPERTY_ZIP_COL = "Property_Zip"

# List of PII columns to be removed specifically for 'COSIGN' agent
# This list is for conditional PII removal based on Agent_Number == 'COSIGN'
COSIGN_PII_COLUMNS = [
    PROPERTY_ADDRESS_COL,
    PROPERTY_CITY_COL,
    PROPERTY_STATE_CODE_COL,
    PROPERTY_ZIP_COL,
]

# Constants for commonly used strings
CSV_EXTENSION = ".csv"
GPG_EXTENSION = ".gpg"

# Page configuration
st.set_page_config(page_title="Ack File Processor", layout="wide")
st.title("📁  Client Ack File Processor")


def log_email_activity(
    agent_numbers: List[str],
    email_sent: bool,
    recipients: List[str] = None,
    subject: str = "",
    log_file: str = "email_log.csv",
):
    """Log email activity to CSV file with specified column order"""
    try:
        current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Create log entry for each agent
        log_entries = []
        for agent in agent_numbers:
            log_entry = {
                "Recipients": "; ".join(recipients) if recipients else "",
                "Agent": agent,
                "Date": current_date,
                "Subject": subject,
                "Email_Status": "Sent" if email_sent else "Failed",
            }
            log_entries.append(log_entry)

        # Create DataFrame with specific column order
        new_log_data = pd.DataFrame(log_entries)
        column_order = ["Recipients", "Agent", "Date", "Subject", "Email_Status"]
        new_log_data = new_log_data[column_order]

        # Load existing log if it exists, otherwise create a new one
        try:
            existing_log = pd.read_csv(log_file)
            # Ensure existing log has the same column structure
            if set(existing_log.columns) != set(column_order):
                logger.warning("Log file structure changed. Starting fresh log file.")
                combined_log = new_log_data
            else:
                existing_log = existing_log[column_order]
                combined_log = pd.concat(
                    [existing_log, new_log_data], ignore_index=True
                )
        except FileNotFoundError:
            combined_log = new_log_data

        # Save the combined log to CSV
        combined_log.to_csv(log_file, index=False)
        logger.info(f"Email activity logged to {log_file}")
        return True

    except Exception as e:
        logger.error(f"Failed to log email activity: {str(e)}")
        return False


class EmailIntegration:
    """Handles email integration with multiple methods"""

    def __init__(self):
        self.outlook_app = None
        self.email_method = None
        self.com_initialized = False

    def initialize_com(self) -> bool:
        """Initialize COM system"""
        if not WIN32COM_AVAILABLE:
            return False

        try:
            pythoncom.CoInitialize()
            self.com_initialized = True
            logger.info("COM system initialized successfully")
            return True
        except Exception as e:
            logger.error(f"Failed to initialize COM system: {str(e)}")
            return False

    def initialize_outlook_com(self) -> bool:
        """Initialize Outlook COM object"""
        if not WIN32COM_AVAILABLE:
            logger.error("win32com library not available. Please install pywin32.")
            return False

        try:
            if not self.com_initialized:
                if not self.initialize_com():
                    return False

            self.outlook_app = win32com.client.Dispatch("Outlook.Application")
            logger.info("Outlook COM object initialized successfully")
            return True
        except Exception as e:
            logger.error(f"Failed to initialize Outlook COM: {str(e)}")
            return False

    def send_email_com(
        self,
        recipients: List[str],
        subject: str,
        body: str,
        attachment_path: Optional[str] = None,
        cc_recipients: List[str] = None,
    ) -> bool:
        """Send email using Outlook COM automation"""
        try:
            if not self.outlook_app:
                if not self.initialize_outlook_com():
                    return False

            mail = self.outlook_app.CreateItem(0)  # 0 = olMailItem
            mail.To = "; ".join(recipients)
            if cc_recipients:
                mail.CC = "; ".join(cc_recipients)
            mail.Subject = subject
            mail.Body = body

            if attachment_path and os.path.exists(attachment_path):
                mail.Attachments.Add(attachment_path)
                logger.info(f"Attachment added: {attachment_path}")

            mail.Send()
            logger.info(f"Email sent successfully to {recipients}")
            if cc_recipients:
                logger.info(f"CC sent to {cc_recipients}")
            return True

        except Exception as e:
            logger.error(f"Failed to send email via COM: {str(e)}")
            return False

    def send_notification_email(
        self,
        recipients: List[str],
        subject: str,
        body: str,
        cc_recipients: List[str] = None,
    ) -> bool:
        """Send notification email without attachments"""
        try:
            if not self.outlook_app:
                if not self.initialize_outlook_com():
                    return False

            mail = self.outlook_app.CreateItem(0)  # 0 = olMailItem
            mail.To = "; ".join(recipients)
            if cc_recipients:
                mail.CC = "; ".join(cc_recipients)
            mail.Subject = subject
            mail.Body = body

            mail.Send()
            logger.info(f"Notification email sent successfully to {recipients}")
            if cc_recipients:
                logger.info(f"CC sent to {cc_recipients}")
            return True

        except Exception as e:
            logger.error(f"Failed to send notification email: {str(e)}")
            return False

    def cleanup(self):
        """Clean up COM resources"""
        try:
            if self.outlook_app:
                self.outlook_app = None

            if self.com_initialized and WIN32COM_AVAILABLE:
                pythoncom.CoUninitialize()
                self.com_initialized = False
                logger.info("COM system cleaned up")
        except Exception as e:
            logger.error(f"Error during COM cleanup: {str(e)}")

    def __del__(self):
        """Destructor to ensure COM cleanup"""
        self.cleanup()


# Initialize email integration
email_integration = EmailIntegration()


def load_error_mapping_file(uploaded_file) -> Optional[pd.DataFrame]:
    """
    Load and prepare the error mapping file from uploaded file
    Returns a DataFrame ready for error mapping operations
    """
    try:
        # Reset file pointer to beginning
        uploaded_file.seek(0)

        # Read the CSV file
        mapping_df = pd.read_csv(uploaded_file, dtype=str, keep_default_na=False)

        # Validate required columns
        required_cols = ["Error_Type", "Client_Error_Type", "Client Action"]
        missing_cols = [col for col in required_cols if col not in mapping_df.columns]

        if missing_cols:
            st.error(
                f"❌ Missing required columns in error mapping file: {missing_cols}"
            )
            return None

        # Clean the data
        # Remove rows where Error_Type is empty or just whitespace
        mapping_df = mapping_df[mapping_df["Error_Type"].str.strip() != ""].copy()

        # Remove completely duplicate rows
        mapping_df = mapping_df.drop_duplicates().copy()

        # Strip whitespace from key columns
        for col in ["Error_Type", "Client_Error_Type", "Client Action", "file_type_2"]:
            if col in mapping_df.columns:
                mapping_df[col] = mapping_df[col].str.strip()

        logger.info(f"Loaded {len(mapping_df)} error mapping records")
        st.success(
            f"✅ Loaded {len(mapping_df)} error mapping records from {uploaded_file.name}"
        )

        return mapping_df

    except Exception as e:
        error_msg = f"Failed to load error mapping file: {str(e)}"
        logger.error(error_msg)
        st.error(f"❌ {error_msg}")
        return None


def get_transaction_reason_text(code: Any) -> str:
    """Translate transaction reason codes to text"""
    code_map = {1: "Sales", 2: "Payments", 3: "Cancels", 5: "Sales"}
    return code_map.get(int(code) if str(code).isdigit() else code, str(code))


def apply_error_mapping(ack_df: pd.DataFrame, mapping_df: pd.DataFrame) -> pd.DataFrame:
    """
    Apply sophisticated error mapping to acknowledgment file

    Logic:
    1. Convert Transaction_Reason to text format (Sales, Payments, Cancels)
    2. Match on both Error_Type AND file_type_2 (Transaction_Reason text)
    3. If no match, try matching on Error_Type only
    4. Update Error_Type with Client_Error_Type and Client_Action with mapped values
    5. Only process rows where isError = TRUE
    """
    try:
        # Create a copy to avoid modifying original DataFrame
        result_df = ack_df.copy()

        # Check if we have the necessary columns for error mapping
        required_ack_cols = [ERROR_TYPE_COL, IS_ERROR_COL]
        missing_ack_cols = [
            col for col in required_ack_cols if col not in result_df.columns
        ]

        if missing_ack_cols:
            logger.warning(f"Missing columns for error mapping: {missing_ack_cols}")
            return result_df

        # Only process rows that are marked as errors
        error_rows = result_df[IS_ERROR_COL].astype(str).str.upper() == "TRUE"

        if not error_rows.any():
            logger.info("No error rows found to process")
            return result_df

        # Convert Transaction_Reason to text format for matching
        if TRANSACTION_REASON_COL in result_df.columns:
            result_df["Transaction_Reason_Text"] = result_df[
                TRANSACTION_REASON_COL
            ].apply(get_transaction_reason_text)

        # Prepare mapping data
        mapping_df_clean = mapping_df.copy()

        # Create comprehensive mapping strategy
        mappings_applied = 0
        unmatched_errors = []

        # Strategy 1: Match on Error_Type AND file_type_2 (Transaction_Reason_Text)
        if (
            "Transaction_Reason_Text" in result_df.columns
            and "file_type_2" in mapping_df_clean.columns
        ):
            # Create composite key for precise matching
            mapping_df_clean["composite_key"] = (
                mapping_df_clean["Error_Type"].astype(str)
                + "|"
                + mapping_df_clean["file_type_2"].astype(str)
            )
            result_df["composite_key"] = (
                result_df[ERROR_TYPE_COL].astype(str)
                + "|"
                + result_df["Transaction_Reason_Text"].astype(str)
            )

            # Create mapping dictionaries for composite key
            composite_error_mapping = dict(
                zip(
                    mapping_df_clean["composite_key"],
                    mapping_df_clean["Client_Error_Type"],
                )
            )
            composite_action_mapping = dict(
                zip(
                    mapping_df_clean["composite_key"], mapping_df_clean["Client Action"]
                )
            )

            # Apply composite key mapping to error rows
            error_mask = error_rows & result_df["composite_key"].isin(
                composite_error_mapping.keys()
            )

            if error_mask.any():
                # Update Error_Type with Client_Error_Type
                result_df.loc[error_mask, ERROR_TYPE_COL] = result_df.loc[
                    error_mask, "composite_key"
                ].map(composite_error_mapping)

                # Update Client_Action
                if CLIENT_ACTION_COL in result_df.columns:
                    result_df.loc[error_mask, CLIENT_ACTION_COL] = (
                        result_df.loc[error_mask, "composite_key"]
                        .map(composite_action_mapping)
                        .fillna("")
                    )

                mappings_applied += error_mask.sum()
                logger.info(
                    f"Applied {error_mask.sum()} precise mappings using Error_Type + Transaction_Reason"
                )

        # Strategy 2: For remaining unmapped errors, try Error_Type only matching
        remaining_errors = error_rows & ~result_df.index.isin(
            result_df[error_mask].index if "error_mask" in locals() else []
        )

        if remaining_errors.any():
            # Create simple Error_Type mapping
            simple_error_mapping = dict(
                zip(
                    mapping_df_clean["Error_Type"],
                    mapping_df_clean["Client_Error_Type"],
                )
            )
            simple_action_mapping = dict(
                zip(mapping_df_clean["Error_Type"], mapping_df_clean["Client Action"])
            )

            # Apply simple mapping to remaining error rows
            simple_match_mask = remaining_errors & result_df[ERROR_TYPE_COL].isin(
                simple_error_mapping.keys()
            )

            if simple_match_mask.any():
                # Update Error_Type with Client_Error_Type
                result_df.loc[simple_match_mask, ERROR_TYPE_COL] = result_df.loc[
                    simple_match_mask, ERROR_TYPE_COL
                ].map(simple_error_mapping)

                # Update Client_Action
                if CLIENT_ACTION_COL in result_df.columns:
                    result_df.loc[simple_match_mask, CLIENT_ACTION_COL] = (
                        result_df.loc[simple_match_mask, ERROR_TYPE_COL]
                        .map(simple_action_mapping)
                        .fillna("")
                    )

                mappings_applied += simple_match_mask.sum()
                logger.info(
                    f"Applied {simple_match_mask.sum()} fallback mappings using Error_Type only"
                )

        # Identify unmatched errors for reporting
        all_matched_mask = (
            error_mask
            if "error_mask" in locals()
            else pd.Series([False] * len(result_df), index=result_df.index)
        ) | (
            simple_match_mask
            if "simple_match_mask" in locals()
            else pd.Series([False] * len(result_df), index=result_df.index)
        )
        unmatched_mask = error_rows & ~all_matched_mask

        if unmatched_mask.any():
            unmatched_count = unmatched_mask.sum()
            unmatched_errors = (
                result_df[unmatched_mask][ERROR_TYPE_COL].unique().tolist()
            )
            logger.warning(
                f"Could not map {unmatched_count} error records with Error_Types: {unmatched_errors}"
            )

        # Clean up temporary columns
        if "composite_key" in result_df.columns:
            result_df.drop(columns=["composite_key"], inplace=True)
        if "Transaction_Reason_Text" in result_df.columns:
            result_df.drop(columns=["Transaction_Reason_Text"], inplace=True)

        # Report mapping results
        total_errors = error_rows.sum()
        st.info(
            f"📊 Error Mapping Results: {mappings_applied}/{total_errors} errors mapped successfully"
        )

        if unmatched_errors:
            st.warning(
                f"⚠️ Unmapped Error Types: {', '.join(unmatched_errors[:10])}"
                + ("..." if len(unmatched_errors) > 10 else "")
            )

        return result_df

    except Exception as e:
        error_msg = f"Failed to apply error mapping: {str(e)}"
        logger.error(error_msg)
        st.error(f"❌ {error_msg}")
        return ack_df  # Return original DataFrame on error


def process_csv_file(
    uploaded_file,
    global_source_filename: str,
    remove_pii: bool,
    error_mapping_df: Optional[pd.DataFrame] = None,
) -> pd.DataFrame:
    """
    Process a single CSV file with all transformations, including:
    - Initial column dropping (general, including job_run_registration, incoming_record_guid, Error_Update_Datetime)
    - Source_Filename insertion/update/blanking (ensured to be present and correctly valued)
    - COSIGN-specific PII removal for property-related fields AND column reordering for Source_Filename
    - Client_Action column handling
    - Error mapping
    - General PII removal (customer-related)
    - Type casting and NaN replacement
    """
    try:
        uploaded_file.seek(0)
        file_content = uploaded_file.read()

        if not file_content or not file_content.strip():
            raise ValueError(
                f"The uploaded file '{uploaded_file.name}' is empty or contains no data."
            )

        uploaded_file.seek(0)

        try:
            df = pd.read_csv(uploaded_file)
        except pd.errors.EmptyDataError:
            raise ValueError(
                f"The file '{uploaded_file.name}' appears to be empty or has no valid CSV data."
            )
        except pd.errors.ParserError as e:
            raise ValueError(f"Error parsing CSV file '{uploaded_file.name}': {str(e)}")

        if df.empty:
            raise ValueError(
                f"The CSV file '{uploaded_file.name}' was read successfully but contains no data rows."
            )

        if not df.columns.any():
            raise ValueError(f"The CSV file '{uploaded_file.name}' has no columns.")

        st.info(
            f"Successfully loaded {len(df)} rows and {len(df.columns)} columns from '{uploaded_file.name}'"
        )

        # --- Initial Column Dropping (General and Universal) ---
        # These columns are dropped regardless of agent type.
        # Source_Filename and incoming_client_filename are NOT in this list.
        columns_to_drop_general = [
            TRANSFER_FLAG_COL,
            JOB_RUN_REGISTRATION_COL,  # Dropped universally
            INCOMING_RECORD_GUID_COL,  # Dropped universally
            IS_IPAY_COL,
            ERROR_JOB_RUN_COL,
            ERROR_SOURCE_COL,
            ERROR_UPDATE_DATETIME_COL,  # Dropped universally
            GENESIS_JOB_RUN_COL,
            STANDARD_JOB_RUN_COL,
        ]

        existing_cols_to_drop_general = [
            col for col in columns_to_drop_general if col in df.columns
        ]
        if existing_cols_to_drop_general:
            df.drop(columns=existing_cols_to_drop_general, inplace=True)
            st.info(
                f"Dropped general columns: {', '.join(existing_cols_to_drop_general)}"
            )
        else:
            st.info("No general columns found for dropping.")

        # --- Handle Source_Filename column: Inserts, updates, or blanks out the source filename ---
        # Determine the desired value for the Source_Filename column.
        # It should be blank by default if no global_source_filename is provided.
        if global_source_filename:
            # If a global source filename is provided (via text input), use it
            desired_source_filename = global_source_filename + CSV_EXTENSION
            action_desc_verb = "set to"
        else:
            # If no global source filename is provided, the column values should be blank
            desired_source_filename = ""
            action_desc_verb = "blanked out"

        if SOURCE_FILENAME_COL not in df.columns:
            # If Source_Filename column doesn't exist, insert it.
            # Attempt to insert it after ORIGINAL_CONTRACT_NUMBER_COL if that column exists.
            insert_after_col = ORIGINAL_CONTRACT_NUMBER_COL
            if insert_after_col in df.columns:
                insert_index = df.columns.get_loc(insert_after_col) + 1
                df.insert(insert_index, SOURCE_FILENAME_COL, desired_source_filename)
                st.info(
                    f"Inserted '{SOURCE_FILENAME_COL}' after '{insert_after_col}' and its values were {action_desc_verb} '{desired_source_filename}'."
                )
            else:
                # If ORIGINAL_CONTRACT_NUMBER_COL doesn't exist, just add it at the end.
                df[SOURCE_FILENAME_COL] = desired_source_filename
                st.info(
                    f"Added '{SOURCE_FILENAME_COL}' at the end and its values were {action_desc_verb} '{desired_source_filename}'."
                )
        else:
            # If Source_Filename column already exists, update its values to the desired_source_filename.
            df[SOURCE_FILENAME_COL] = desired_source_filename
            st.info(
                f"'{SOURCE_FILENAME_COL}' column already existed and its values were {action_desc_verb} '{desired_source_filename}'."
            )

        # --- COSIGN Specific Processing ---
        # This block applies property-related PII removal AND column reordering
        if (
            AGENT_NUMBER_COL in df.columns
            and "COSIGN" in df[AGENT_NUMBER_COL].astype(str).str.upper().unique()
        ):
            st.info(
                "🎯 'COSIGN' Agent detected. Applying COSIGN-specific transformations."
            )

            # Create a boolean mask for rows where Agent_Number is 'COSIGN'
            cosign_mask = df[AGENT_NUMBER_COL].astype(str).str.upper() == "COSIGN"

            # 1. Remove PII information from specific columns for 'COSIGN' agent
            pii_removed_for_cosign = []
            for col in COSIGN_PII_COLUMNS:
                if col in df.columns:
                    # Ensure the column is of string type to avoid dtype issues with empty string assignment
                    df[col] = df[col].astype(str)
                    # Set PII values to empty string only for 'COSIGN' rows
                    df.loc[cosign_mask, col] = ""
                    pii_removed_for_cosign.append(col)
            if pii_removed_for_cosign:
                st.info(
                    f"Cleared PII for 'COSIGN' in columns: {', '.join(pii_removed_for_cosign)}"
                )
            else:
                st.info(
                    "No COSIGN-specific property PII columns found for conditional removal."
                )

            # 2. COSIGN-specific Column Reordering: Source_Filename needs to be before incoming_client_filename
            if (
                SOURCE_FILENAME_COL in df.columns
                and INCOMING_CLIENT_FILENAME_COL in df.columns
            ):
                source_filename_idx = df.columns.get_loc(SOURCE_FILENAME_COL)
                incoming_client_filename_idx = df.columns.get_loc(
                    INCOMING_CLIENT_FILENAME_COL
                )

                if source_filename_idx > incoming_client_filename_idx:
                    # If Source_Filename is currently after incoming_client_filename, move it
                    cols = df.columns.tolist()
                    # Remove Source_Filename from its current position
                    source_filename_col_data = cols.pop(source_filename_idx)
                    # Insert Source_Filename before incoming_client_filename
                    cols.insert(incoming_client_filename_idx, source_filename_col_data)
                    df = df[cols]  # Reindex DataFrame with new column order
                    st.info(
                        f"Reordered '{SOURCE_FILENAME_COL}' to be before '{INCOMING_CLIENT_FILENAME_COL}' for 'COSIGN' agent."
                    )
                else:
                    st.info(
                        f"'{SOURCE_FILENAME_COL}' is already before '{INCOMING_CLIENT_FILENAME_COL}' (or at same position). No reordering needed for 'COSIGN'."
                    )
            else:
                st.info(
                    f"Skipping reordering for 'COSIGN': '{SOURCE_FILENAME_COL}' or '{INCOMING_CLIENT_FILENAME_COL}' not found in the file."
                )
        else:
            st.info(
                "No 'COSIGN' agent detected or Agent_Number column missing. Skipping COSIGN-specific rules."
            )

        # Handle Client_Action column: Ensures this column exists.
        if CLIENT_ACTION_COL not in df.columns:
            if ERROR_MESSAGE_COL in df.columns:
                insert_index = df.columns.get_loc(ERROR_MESSAGE_COL) + 1
                df.insert(insert_index, CLIENT_ACTION_COL, "")
                st.info(f"Inserted {CLIENT_ACTION_COL} after {ERROR_MESSAGE_COL}")
            else:
                df[CLIENT_ACTION_COL] = ""
                st.info(f"Added {CLIENT_ACTION_COL} at the end")
        else:
            st.info("Client_Action column already exists")

        # Apply error mapping if available and configured.
        if error_mapping_df is not None:
            st.info("🔄 Applying error mapping...")
            df = apply_error_mapping(df, error_mapping_df)

        # General PII removal (controlled by `remove_pii` flag).
        # This targets customer-related PII columns.
        if remove_pii:
            pii_columns_customer = [
                CUSTOMER_FIRST_NAME_COL,
                CUSTOMER_ADDRESS_1_COL,
                CUSTOMER_CITY_COL,
                CUSTOMER_STATE_COL,
                CUSTOMER_ZIP_CODE_COL,
                CUSTOMER_PHONE_COL,
                CUSTOMER_EMAIL_COL,
            ]
            # Filter for columns that actually exist in the DataFrame
            existing_pii_columns_customer = [
                col for col in pii_columns_customer if col in df.columns
            ]
            if existing_pii_columns_customer:
                for col in existing_pii_columns_customer:
                    df[col] = df[col].astype(str)  # Explicitly cast to string
                    df[col] = ""
                st.info(
                    f"Removed general customer PII info from: {', '.join(existing_pii_columns_customer)}"
                )
            else:
                st.info("No general customer PII columns found for removal.")

        # Cast all columns to string to ensure consistent data types for output.
        for col in df.columns:
            df[col] = df[col].astype(str)

        # Replace "nan" strings with empty strings. This handles potential 'nan' values
        # introduced during CSV reading or type casting, ensuring truly empty cells.
        df = df.replace("nan", "", regex=True)

        return df

    except ValueError as e:
        st.error(str(e))
        raise
    except Exception as e:
        st.error(f"Unexpected error processing file {uploaded_file.name}: {str(e)}")
        raise


def generate_filename(
    df: pd.DataFrame, output_year: str, output_month: str, uploaded_file_name: str
) -> str:
    """Generate processed filename based on data content. Format: Ack_Agent_Number_YYYY_MM_Transaction_Reason"""

    agent_number = None
    transaction_reason_text = None

    if AGENT_NUMBER_COL in df.columns and not df[AGENT_NUMBER_COL].dropna().empty:
        agent_number = "_".join(
            [
                str(agent).strip()
                for agent in df[AGENT_NUMBER_COL].unique()
                if str(agent).strip()
            ]
        )

    if (
        TRANSACTION_REASON_COL in df.columns
        and not df[TRANSACTION_REASON_COL].dropna().empty
    ):
        unique_reasons = df[TRANSACTION_REASON_COL].unique()
        transaction_reasons = [
            get_transaction_reason_text(reason)
            for reason in unique_reasons
            if str(reason).strip()
        ]
        # Remove duplicates while preserving order
        transaction_reasons = list(dict.fromkeys(transaction_reasons))
        transaction_reason_text = "_".join(transaction_reasons)

    if agent_number and transaction_reason_text:
        processed_filename = f"Ack_{agent_number}_{output_year}_{output_month}_{transaction_reason_text}{CSV_EXTENSION}"
    else:
        processed_filename = (
            f"Ack_processed_{output_year}_{output_month}_{uploaded_file_name}"
        )

    return processed_filename


def create_default_notification_email_body() -> str:
    """Create default notification email body template"""
    body = """Hello,
an acknowledgement file has been posted to your sftp-output bucket. You’ll find it in the out folder, where all acknowledgement files are placed for retrieval.

Please review the Client Action column (letter BE) for any steps required before resubmitting your file. The acknowledgement file reflects the information that will appear on your invoice.

If action is required, kindly resubmit by next month's end cycle. When resubmitting corrections, use the original file name and append with “_corrections” (e.g., CLIENT_SALES_XXXXXXXX_corrections).

***As a reminder, the acknowledgement file communicates what was loaded and not loaded from your original submitted files. Some of the errors within the file need to be resolved by IWW, some need to be resolved by you, and some are not expected to be resolved (e.g. duplicate contract). Only the non-errors will be present in the invoice for that processing period. In order to ensure we have fully processed all records and stay aligned between our systems of record, it is very important you review your ACK file and contact us with any questions and ultimately provide corrections before the next processing cycle (15 – 25 of each month).***


If you have any questions, feel free to reach out to my email: antonio.nunnally@ironwoodwarrantygroup.com or call me at 502-814-1740

Best regards,


"""
    return body


def validate_uploaded_file(uploaded_file) -> bool:
    """
    Validate an uploaded file to ensure it's a valid CSV with data
    Returns True if valid, False otherwise
    """
    try:
        uploaded_file.seek(0)
        file_content = uploaded_file.read()

        if not file_content or not file_content.strip():
            st.error(f"❌ File '{uploaded_file.name}' is empty.")
            return False

        uploaded_file.seek(0)

        try:
            df_sample = pd.read_csv(
                uploaded_file, nrows=5
            )  # Read only first 5 rows for validation

            if not df_sample.columns.any():  # Robust check for no columns
                st.error(f"❌ File '{uploaded_file.name}' has no columns.")
                return False

            uploaded_file.seek(0)
            return True

        except pd.errors.EmptyDataError:
            st.error(
                f"❌ File '{uploaded_file.name}' appears to be empty or has no valid CSV data."
            )
            return False
        except pd.errors.ParserError as e:
            st.error(f"❌ Error parsing CSV file '{uploaded_file.name}': {str(e)}")
            return False
        except Exception as e:
            st.error(
                f"❌ Unexpected error validating file '{uploaded_file.name}': {str(e)}"
            )
            return False

    except Exception as e:
        st.error(f"❌ Error accessing file '{uploaded_file.name}': {str(e)}")
        return False


# --- UI Components ---

# Sidebar for Client List Upload and Error Mapping
with st.sidebar:
    st.header("📋 Configuration Files")

    st.subheader("Client List")
    client_list_file = st.file_uploader(
        "Upload Client_List_Data_Contact_Sheet.csv",
        type="csv",
        key="client_list_uploader",
    )

    st.subheader("Error Mapping")
    error_mapping_file = st.file_uploader(
        "Upload CLIENT_MONTH_END_PROCESSING_ERRORS.csv",
        type="csv",
        key="error_mapping_uploader",
        help="This file maps Error_Type values to Client_Error_Type and provides Client Actions",
    )

    # Load error mapping file
    error_mapping_df = None
    if error_mapping_file:
        error_mapping_df = load_error_mapping_file(error_mapping_file)

        if error_mapping_df is not None:
            with st.expander("📊 Error Mapping Preview", expanded=False):
                st.dataframe(error_mapping_df.head(10), use_container_width=True)
                st.caption(
                    f"Showing first 10 of {len(error_mapping_df)} mapping records"
                )

# File Upload Section
uploaded_files = st.file_uploader(
    "Upload one or more Ack CSV files", type="csv", accept_multiple_files=True
)

# Determine default for PII removal based on Agent_Number
default_pii_removal = "Yes"  # Default to "Yes"
if uploaded_files:
    for uploaded_file in uploaded_files:
        try:
            uploaded_file.seek(0)
            df_temp = pd.read_csv(
                uploaded_file, nrows=10
            )  # Read a few rows to check Agent_Number

            if AGENT_NUMBER_COL in df_temp.columns:
                agent_numbers = (
                    df_temp[AGENT_NUMBER_COL].astype(str).str.upper().unique()
                )
                if any(
                    agent in agent_numbers
                    for agent in ["GUARD", "PULS", "JCTV", "PWSC"]
                ):
                    default_pii_removal = "No"
                    break  # Stop processing files if GUARD or PULS is found
            uploaded_file.seek(0)  # Reset file pointer

        except Exception as e:
            st.warning(
                f"Could not read Agent_Number from {uploaded_file.name}: {str(e)}"
            )
            continue

# PII Removal Section
st.markdown("### PII Removal")
remove_pii = st.radio(
    "Remove Personal Identifiable Information (PII)?",
    ["No", "Yes"],
    index=["No", "Yes"].index(default_pii_removal),  # Set the default selection
    horizontal=True,
)
remove_pii = remove_pii == "Yes"

# Optional Fields Section
st.markdown("### Optional Fields")

# Initialize session state for global_source_filename and the clear notification message
if "global_source_filename_input" not in st.session_state:
    st.session_state.global_source_filename_input = ""
if "clear_notification_message" not in st.session_state:
    st.session_state.clear_notification_message = ""


# Define the callback function for the "Clear" button
def clear_filename_input_callback():
    st.session_state.global_source_filename_input = ""
    st.session_state.clear_notification_message = "✅ 'Source_Filename' input cleared. Processed files will now have blank 'Source_Filename' values."


# Display temporary notification if 'Source_Filename' input was cleared
if st.session_state.clear_notification_message:
    st.success(st.session_state.clear_notification_message)
    # Clear the message from session state so it only shows once
    st.session_state.clear_notification_message = ""

# Create two columns for the text input and the clear button
col_input, col_button = st.columns([0.9, 0.1])

with col_input:
    # The text input for the global source filename.
    # By setting key="global_source_filename_input", Streamlit automatically links
    # this widget's value to st.session_state.global_source_filename_input.
    # No need for the 'value' parameter if the key directly manages it.
    st.text_input(
        "Enter Source_Filename (optional, applied to all rows):",
        key="global_source_filename_input",
        help="Provide a filename prefix for the 'Source_Filename' column. Leave blank to clear existing values.",
    )

with col_button:
    # Add a small buffer space to align the button with the text input
    st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
    # Clear button for the Source_Filename input
    # Use the on_click callback to modify the session state.
    st.button(
        "Clear",
        on_click=clear_filename_input_callback,
        key="clear_source_filename_button",
    )

# The variable 'global_source_filename' used by process_csv_file will
# now hold the current value from the session state (which reflects the text input).
global_source_filename = st.session_state.global_source_filename_input


# Output File Naming Section
st.markdown("### Output Configuration")
col1, col2 = st.columns(2)
with col1:
    output_month = st.selectbox(
        "Select Month",
        ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"],
        index=0,
    )
with col2:
    output_year = st.selectbox("Select Year", ["2025", "2026", "2027"], index=0)

# Email Integration Section
st.markdown("### 📧 Email Notification")

# Load client list data if available
client_emails = {}
if client_list_file:
    try:
        client_list_df = pd.read_csv(client_list_file, keep_default_na=False, dtype=str)
        if "Account" in client_list_df.columns and "Email" in client_list_df.columns:
            # Aggregate emails by account, handling semicolon-separated emails
            for _, row in client_list_df.iterrows():
                account = (
                    str(row["Account"]).strip().upper()
                )  # Convert to uppercase for matching
                email_str = str(row["Email"]).strip()

                # Skip empty email entries
                if not email_str or email_str.lower() == "nan":
                    continue

                # Split the email string by semicolon, strip whitespace, and extend the list
                emails = [
                    email.strip()
                    for email in email_str.split(";")
                    if email.strip() and email.strip().lower() != "nan"
                ]

                if emails:  # Only add if there are valid emails
                    if account in client_emails:
                        client_emails[account].extend(emails)
                    else:
                        client_emails[account] = emails

            # Remove duplicates from each account's email list
            for account in client_emails:
                client_emails[account] = list(set(client_emails[account]))

            st.success(
                f"✅ Client list loaded successfully! Found {len(client_emails)} accounts with email addresses."
            )

            # Show client list preview
            with st.expander("📧 Client List Preview", expanded=False):
                preview_data = []
                for account, emails in list(client_emails.items())[
                    :10
                ]:  # Show first 10
                    preview_data.append(
                        {
                            "Account": account,
                            "Email Count": len(emails),
                            "Emails": "; ".join(emails[:3])
                            + ("..." if len(emails) > 3 else ""),
                        }
                    )
                if preview_data:
                    st.dataframe(pd.DataFrame(preview_data), use_container_width=True)
                    st.caption(
                        f"Showing first 10 of {len(client_emails)} accounts with emails"
                    )

        else:
            st.error("❌ Client list file must contain 'Account' and 'Email' columns.")
            client_emails = {}  # Ensure client_emails is empty

    except Exception as e:
        st.error(f"❌ Error loading client list: {e}")
        client_emails = {}  # Ensure client_emails is empty

# Determine default recipients based on client list and uploaded files
default_recipients = []
email_configured_by_default = False  # Flag to track if email is configured

if client_emails and uploaded_files:
    for uploaded_file in uploaded_files:
        try:
            uploaded_file.seek(0)
            df_temp = pd.read_csv(uploaded_file, keep_default_na=False, dtype=str)

            if AGENT_NUMBER_COL in df_temp.columns:
                agent_numbers = df_temp[AGENT_NUMBER_COL].dropna().unique()
                for agent_number in agent_numbers:
                    agent_number_str = str(agent_number).strip().upper()
                    if agent_number_str in client_emails:
                        # Flatten the list of emails and join with commas
                        emails = client_emails[agent_number_str]
                        default_recipients.extend(emails)
                        email_configured_by_default = True  # Account match found

            uploaded_file.seek(0)  # Reset file pointer
        except Exception as e:
            st.warning(
                f"Could not read agent numbers from {uploaded_file.name}: {str(e)}"
            )
            continue

    # Remove duplicates and join with commas
    default_recipients = ",".join(sorted(set(default_recipients)))

    # Set the "Configure email notification to clients" option to "yes" if default recipients are found
    send_email_default = "Yes" if email_configured_by_default else "No"
else:
    send_email_default = "No"  # Default if no client list or uploaded files

send_email = st.radio(
    "Configure email notification to clients?",
    ["No", "Yes"],
    index=["No", "Yes"].index(send_email_default),
    horizontal=True,
)

email_recipients = []
cc_recipients = [
    "kate.bowling@ironwoodwarrantygroup.com",
    "julie.messer@ironwoodwarrantygroup.com",
    "antonio.nunnally@ironwoodwarrantygroup.com",
]

if send_email == "Yes":
    if not WIN32COM_AVAILABLE:
        st.error(
            "⚠️ Outlook COM not available. Please install pywin32 for Outlook integration."
        )
        st.code("pip install pywin32")
        send_email = "No"
    else:
        st.markdown("#### Email Recipients")
        col1, col2 = st.columns(2)
        with col1:
            recipients_input = st.text_area(
                "To Recipients (one per line):",
                placeholder="client1@example.com\nclient2@example.com",
                value=default_recipients,
                help="Enter client email addresses, one per line",
            )
        with col2:
            cc_recipients_input = st.text_area(
                "CC Recipients (one per line):",
                placeholder="cc1@example.com\ncc2@example.com",
                value="\n".join(cc_recipients),
                help="Enter CC email addresses, one per line",
            )

        if recipients_input:
            email_recipients = [
                email.strip() for email in recipients_input.split("\n") if email.strip()
            ]
            st.info(f"To: {', '.join(email_recipients)}")

        if cc_recipients_input:
            cc_recipients = [
                email.strip()
                for email in cc_recipients_input.split("\n")
                if email.strip()
            ]
            st.info(f"CC: {', '.join(cc_recipients)}")

if email_recipients:
    st.markdown("#### Email Content")

    agent_names = set()
    agent_numbers = set()
    if uploaded_files:
        for uploaded_file in uploaded_files:
            try:
                uploaded_file.seek(0)
                file_content = uploaded_file.read()
                if not file_content or not file_content.strip():
                    st.warning(f"Skipping empty file: {uploaded_file.name}")
                    continue

                uploaded_file.seek(0)
                df_temp = pd.read_csv(uploaded_file, keep_default_na=False, dtype=str)

                if df_temp.empty or not df_temp.columns.any():
                    st.warning(f"Skipping file with no data: {uploaded_file.name}")
                    continue

                if AGENT_NAME_COL in df_temp.columns:
                    agents = df_temp[AGENT_NAME_COL].dropna().unique()
                    agent_names.update(
                        [
                            str(agent)
                            for agent in agents
                            if str(agent) and str(agent).strip()
                        ]
                    )
                if AGENT_NUMBER_COL in df_temp.columns:
                    agents = df_temp[AGENT_NUMBER_COL].dropna().unique()
                    agent_numbers.update(
                        [
                            str(agent)
                            for agent in agents
                            if str(agent) and str(agent).strip()
                        ]
                    )
            except pd.errors.EmptyDataError:
                st.warning(f"Skipping empty or invalid CSV file: {uploaded_file.name}")
                continue
            except Exception as e:
                st.warning(
                    f"Could not read file {uploaded_file.name} for email subject: {str(e)}"
                )
                continue

    # Date formatting for email subject
    month_name = datetime.strptime(output_month, "%m").strftime("%B")
    formatted_date = f"{month_name} {output_year}"

    if agent_names:
        agent_names_str = ", ".join(sorted(agent_names))
        default_subject = f"{agent_names_str} - Ack File Processing Complete - Files Uploaded for Review - {formatted_date}"
    elif agent_numbers:
        agent_numbers_str = ", ".join(sorted(agent_numbers))
        default_subject = f"{agent_numbers_str} - Ack File Processing Complete - Files Uploaded for Review - {formatted_date}"
    else:
        default_subject = f"Ack File Processing Complete - Files Uploaded for Review - {formatted_date}"

    custom_email_subject = st.text_input(
        "Email Subject:",
        value=default_subject,
        help="Edit the email subject as needed",
    )

    default_body = create_default_notification_email_body()
    custom_email_body = st.text_area(
        "Email Body:",
        value=default_body,
        height=300,
        help="Edit the email body as needed. You can use custom formatting and add specific details.",
    )


# --- File Processing ---
if uploaded_files:
    processed_files = []
    processed_data = []

    for uploaded_file in uploaded_files:
        st.markdown(f"---\n### Processing: `{uploaded_file.name}`")

        try:
            df = process_csv_file(
                uploaded_file, global_source_filename, remove_pii, error_mapping_df
            )

            st.dataframe(df, use_container_width=True, hide_index=True)

            processed_filename = generate_filename(
                df, output_year, output_month, uploaded_file.name
            )
            processed_files.append(processed_filename)

            csv_data = df.to_csv(index=False).encode("utf-8")

            processed_data.append(
                {
                    "filename": processed_filename,
                    "data": csv_data,
                    "encrypted_data": None,
                    "encrypted_filename": None,
                }
            )

            st.download_button(
                label="📥 Download Processed CSV",
                data=csv_data,
                file_name=processed_filename,
                mime="text/csv",
            )

        except Exception as e:
            st.error(f"❌ Error processing file {uploaded_file.name}: {str(e)}")
            st.exception(e)

    if send_email == "Yes" and email_recipients and processed_data:
        st.markdown("---\n### 📧 Email Notification")

        all_agent_numbers = set()
        all_agent_names = set()
        for data_item in processed_data:
            try:
                df_temp = pd.read_csv(
                    io.StringIO(data_item["data"].decode("utf-8")),
                    keep_default_na=False,
                    dtype=str,
                )
                if AGENT_NUMBER_COL in df_temp.columns:
                    agent_numbers = df_temp[AGENT_NUMBER_COL].dropna().unique()
                    all_agent_numbers.update(
                        [str(agent) for agent in agent_numbers if str(agent)]
                    )
                if AGENT_NAME_COL in df_temp.columns:
                    agent_names = df_temp[AGENT_NAME_COL].dropna().unique()
                    all_agent_names.update(
                        [str(agent) for agent in agent_names if str(agent)]
                    )
            except Exception as e:
                st.warning(
                    f"Could not extract agent information from {data_item['filename']}: {str(e)}"
                )

        final_subject = (
            custom_email_subject
            if "custom_email_subject" in locals()
            else default_subject
        )
        final_body = (
            custom_email_body
            if "custom_email_body" in locals()
            else create_default_notification_email_body()
        )

        with st.expander("📧 Email Preview"):
            st.write(f"**To:** {', '.join(email_recipients)}")
            if cc_recipients:
                st.write(f"**CC:** {', '.join(cc_recipients)}")
            st.write(f"**Subject:** {final_subject}")
            st.write("**Body:**")
            st.text(final_body)

        if st.button("📧 Send Notification Email", type="primary"):
            st.info("Sending notification email via Outlook...")

            email_success = email_integration.send_notification_email(
                email_recipients, final_subject, final_body, cc_recipients
            )

            if all_agent_names:
                agent_numbers_for_log = list(all_agent_names)
            elif all_agent_numbers:
                agent_numbers_for_log = list(all_agent_numbers)
            else:
                agent_numbers_for_log = ["Unknown"]

            all_recipients = email_recipients.copy()
            if cc_recipients:
                all_recipients.extend([f"CC: {cc}" for cc in cc_recipients])

            log_success = log_email_activity(
                agent_numbers_for_log, email_success, all_recipients, final_subject
            )

            if email_success:
                st.success("✅ Notification email sent successfully!")
            else:
                st.error(
                    "❌ Failed to send notification email. Check logs for details."
                )
        else:
            st.info("Click the button to send the notification email.")
    else:
        st.info("Email notification is disabled or no recipients/files available.")

else:
    st.info("Upload one or more CSV files to begin processing.")

# --- Footer Information ---
if error_mapping_df is not None:
    st.markdown("---")
    st.markdown("### 🔧 Error Mapping Status")
    st.success("✅ Error mapping is active and will be applied to uploaded files")

    # Show mapping statistics
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Mappings", len(error_mapping_df))
    with col2:
        unique_error_types = error_mapping_df["Error_Type"].nunique()
        st.metric("Unique Error Types", unique_error_types)
    with col3:
        if "file_type_2" in error_mapping_df.columns:
            unique_file_types = error_mapping_df["file_type_2"].nunique()
            st.metric("File Types", unique_file_types)
else:
    st.markdown("---")
    st.markdown("### ⚠️ Error Mapping Status")
    st.warning(
        "Error mapping file not loaded. Upload CLIENT_MONTH_END_PROCESSING_ERRORS.csv to enable error mapping."
    )

st.markdown("---")
st.markdown("*Powered by Streamlit | Enhanced with Smart Error Mapping*")
