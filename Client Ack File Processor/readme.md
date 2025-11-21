# Client Ack File Processor

A Streamlit-based application for processing client acknowledgment (Ack) files with advanced features including PII removal, error mapping, email notifications, and agent-specific transformations.

## Features

### Core Functionality
- **Batch CSV Processing**: Process multiple acknowledgment files simultaneously
- **Smart Column Management**: Automatically handles required columns (Source_Filename, Client_Action, etc.)
- **Error Mapping**: Sophisticated error type mapping with dual-strategy matching (Error_Type + Transaction_Reason or Error_Type only)
- **Dynamic Filename Generation**: Auto-generates standardized filenames based on agent number, year, month, and transaction reason

### Data Protection
- **PII Removal**: Configurable removal of personally identifiable information
  - Customer-related PII (name, address, phone, email)
  - COSIGN-specific property PII (property address, city, state, zip)
- **Agent-Specific Rules**: Special handling for COSIGN, GUARD, PULS, JCTV, and PWSC agents

### Email Integration
- **Automated Notifications**: Send acknowledgment files via Outlook COM automation
- **Client List Integration**: Auto-populate recipients from client contact sheet
- **Email Activity Logging**: Track all sent emails with timestamps and recipients
- **Customizable Templates**: Edit email subject and body before sending

### COSIGN Agent Special Processing
- Property-related PII removal for COSIGN agent rows
- Column reordering (Source_Filename positioned before incoming_client_filename)

## Installation

### Prerequisites
- Python 3.8 or higher
- Windows OS (for Outlook COM integration)

### Required Packages
```bash
pip install streamlit pandas pythoncom pywin32
```

### Optional Dependencies
For full email functionality:
```bash
pip install pywin32
```

## Usage

### Starting the Application
```bash
streamlit run Client_Ack_File_Processor_v4.2.py
```

### Configuration Files

#### 1. Client List (Required for Email)
Upload `Client_List_Data_Contact_Sheet.csv` with columns:
- `Account`: Agent/client account identifier
- `Email`: Email addresses (semicolon-separated for multiple)

#### 2. Error Mapping (Optional but Recommended)
Upload `CLIENT_MONTH_END_PROCESSING_ERRORS.csv` with columns:
- `Error_Type`: Original error type
- `Client_Error_Type`: Mapped client-friendly error type
- `Client Action`: Required action for the client
- `file_type_2`: Transaction reason (Sales, Payments, Cancels)

### Processing Workflow

1. **Upload Files**: Select one or more CSV acknowledgment files
2. **Configure PII Removal**: Choose whether to remove personally identifiable information
3. **Set Source Filename** (Optional): Enter a global source filename or leave blank
4. **Select Output Date**: Choose month and year for filename generation
5. **Configure Email**: 
   - Toggle email notifications
   - Review auto-populated recipients
   - Customize email subject and body
6. **Process**: Files are automatically processed and ready for download
7. **Send Notifications**: Click to send emails with processed files

### File Processing Rules

#### Columns Automatically Removed
- `Transfer_Flag`
- `job_run_registration`
- `incoming_record_guid`
- `is_ipay`
- `Error_Job_Run`
- `Error_Source`
- `Error_Update_Datetime`
- `Genesis_Job_Run`
- `Standard_Job_Run`

#### Columns Preserved
- `Source_Filename` (inserted/updated as specified)
- `incoming_client_filename`
- `Original_Contract_Number`
- All other data columns

#### PII Columns (Conditional Removal)
**Customer PII** (removed when PII removal is enabled):
- Customer_First_Name
- Customer_Address_1
- Customer_City
- Customer_State
- Customer_Zip_Code
- Customer_Phone
- Customer_Email

**Property PII** (removed only for COSIGN agent):
- Property_Address
- Property_City
- Property_State_Code
- Property_Zip

## Output

### Filename Format
```
Ack_[AgentNumber]_[YYYY]_[MM]_[TransactionReason].csv
```

Example: `Ack_COSIGN_2025_01_Sales.csv`

### Email Notifications
- Subject: `[Agent Name] - Ack File Processing Complete - Files Uploaded for Review - [Month Year]`
- Default CCs: kate.bowling@, julie.messer@, antonio.nunnally@ironwoodwarrantygroup.com
- Activity logged to `email_log.csv`

## Error Mapping Logic

The application uses a two-tier matching strategy:

1. **Precise Match**: Error_Type + Transaction_Reason (file_type_2)
2. **Fallback Match**: Error_Type only

For rows where `isError = TRUE`:
- Updates `Error_Type` with `Client_Error_Type`
- Populates `Client_Action` with mapped action
- Reports matching statistics

## Logging

### Email Activity Log
Location: `email_log.csv`

Columns:
- Recipients
- Agent
- Date
- Subject
- Email_Status (Sent/Failed)

## Agent-Specific Behavior

### COSIGN
- Removes property-related PII fields
- Reorders Source_Filename column
- Special column positioning rules

### GUARD, PULS, JCTV, PWSC
- Default PII removal set to "No"
- Standard processing otherwise

## Troubleshooting

### Outlook COM Errors
- Ensure Microsoft Outlook is installed
- Run `pip install pywin32` and restart
- Check Windows COM permissions

### Empty File Errors
- Verify CSV files contain data rows
- Check for proper CSV formatting
- Ensure files have column headers

### Missing Error Mappings
- Upload error mapping file before processing
- Verify mapping file has required columns
- Check for matching Error_Type values

## Version History

**v4.2** (Current)
- Enhanced error mapping with dual-strategy matching
- COSIGN agent special processing
- Improved PII handling
- Email activity logging
- Session state management for Source_Filename input

## Support

For questions or issues:
- Email: antonio.nunnally@ironwoodwarrantygroup.com
- Phone: 502-814-1740

## License

Internal use only - Ironwood Warranty Group

---

*Built with Streamlit | Enhanced with Smart Error Mapping*
