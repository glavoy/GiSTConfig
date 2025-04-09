# GiSTConfig

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

