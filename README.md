# GiSTConfig Instructions

This software will generate XML files from an Excel file containing a Data Dictionary. It also creates an MS Access databse with a table for each worksheet.
Note the worksheets in the Excel data dictionary must be named with '_dd' appended to the name in order to create tables and xml files.
Alternatively, you can append the worksheet name with '__xml' and only an xml file will be created. All othere worksheets are ignored.

Before running the program, set the following variables in the `GiSTConfig.cs` file:


```csharp

//SONET NCF Pilot
//***************************
// Path to Excel file - this is the path to the Data Dictionary you want to use
readonly string excelFile = "C:\\Users\\glavoy\\Box\\SONET R01-UCSF and IDRC\\PILOTS- current 2025\\Pilot 1- Network case finding\\Final instruments\\SONET Data Dictionary NCF Surveys 2025-04-08.xlsx";

//***************************
// Path to XML file - this is where the generated xml files will be written
readonly string xmlPath = "C:\\Users\\glavoy\\Dropbox\\IDRC\\SONET2\\Applications\\sonet_ncf_pilot\\xml\\";

//***************************
// Path to log file - keeps track of errors
readonly string logfilePath = "C:\\temp\\";

//***************************
// Path to MS Access database - will be created
readonly string accessDB = "C:\\SONET\\NCF_Pilot\\MSAccessDatabase\\SONET_NCF_Pilot.mdb";

//***************************
// name of the source tables to copy
// Create a MS Access databse with the name as the database above, except with " - master" appended to the name
// For example, 'SONET_NCF_Pilot.mdb' and 'SONET_NCF_Pilot - Master.mdb'
// The software will look for this database to copy the tables from. Below is a list of tables you want to copy from the 'master' to the newly created MS Access databse.
public string[] sourceTableNames = { "bl_complete", "hhmembers", "households", "sn_complete", "tb_cases", "villages", "fingerprints", "config", "formchanges", "audittrail" };
```


# Excel Data Dictionary Instructions

## Header Structure
1. First row must be the header row

## Skip Logic Format
2. Skip must be of the format - `skiptype: if fieldname_to_check condition value, skip to fieldname_to_skip_to`
   - Begins with "skiptype: " (skiptype, colon, space)
   - **skiptype**: "postskip" or "preskip"
   - **fieldname_to_check**: which field name to check
   - **condition**: =, <, >, <=, >=, <>, 'does not contain', 'contains'
   - **value**: value to check against the value in fieldname_to_check
   - **fieldname_to_skip_to**: where to skip
   - *It is very important to have SINGLE SPACES between each part*

## Field Specifications
3. For fields that are QuestionType = text and FieldType = text, you should specify the MaxCharacters
4. DontKnow, Refuse, NA - each are specified in their own column. Must be TRUE or blank
5. Do not include DontKnow, Refuse, NA in radio/checkbox responses
6. LowerRange and UpperRange are continuous and both have to have a number or both be blank
7. DD must have the same 13 columns
8. Responses must begin with "1:" - (number, colon)

## Question Types
9. QuestionType:
   | Questionnaire | Comment |
   |---------------|---------|
   | radio | Radio Buttons - fieldtype MUST be integer |
   | combobox | Dropdown - fieldtype MUST be integer |
   | checkbox | Checkboxes - fieldtype MUST be text |
   | text | TextBox - should specify the MaxCharacters |
   | date | Date Picker |
   | information | Displays information on screen. Not saved to database |
   | automatic | Question is automatically responded to by the software. Code MUST to be written in the AddAutomatic() function |

## Field Types
10. FieldType:
    | Database | Comment |
    |----------|---------|
    | text | Short Text - Allows any character, default is 255 characters |
    | datetime | Date/Time |
    | date | Date/Time |
    | phone_num | Short Text - Allows only numbers in Text box; 10 characters in the database |
    | integer | Long Integer |
    | text_integer | Long Integer - Allows only numbers in Text box |
    | text_id | Text - Allows only numbers in Text box |
    | text_decimal | Decimal - Allows only numbers and decimal point in Text box; Precision = 13; Scale = 5 |
    | hourmin | Short Text - Allows only numbers and colon in Text box; 5 characters in the database |

## Worksheet Structure
11. All rows that are not questions must be merged cells
12. All Worksheets that contain a data dictionary must end in "_dd"

## Logic Checks
13. Logic checks:
    - **dynamic**: `if intvinit2 <> intvinit, error_message This does not match your previous entry!`
    - **fixed**: `if month = 2 'and' day = 30, error_message throw an error`
