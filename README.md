# GiSTConfig Instructions


## Master Database and Table Copying Instructions

This software includes functionality to **generate XML files** and **create MS Access databases** based on a provided Excel file containing data dictionaries. It also supports copying reference tables from a corresponding master database.

---

### Step 1: Excel File Naming Rules

- Each **worksheet** in the Excel file must be named according to the following rules:
  
  | Worksheet Name Suffix | Action |
  |------------------------|--------|
  | `_dd`                 | Creates an MS Access **table** and an associated **XML** file |
  | `_xml`               | Creates an **XML** file **only** (no table created) |
  | *(Any other name)*    | Worksheet is **ignored** |

> ✅ Use `_dd` if you want both XML and a table in Access.  
> ✅ Use `_xml` if you only want an XML file.  
> ❌ Do not use any other naming if you expect the sheet to be processed.

---

### Step 2: Create the Master Database

- For every new database (e.g. `SONET_NCF_Pilot.mdb`), you may want to create a **corresponding master database** with the **same name**, followed by **` - Master.mdb`**.

  **Example:**
  - New Database: `SONET_NCF_Pilot.mdb`
  - Master Database: `SONET_NCF_Pilot - Master.mdb`

- The software will look for this master database during initialization to copy over any required preloaded or reference tables.

---

### Step 3: Copying Tables from Master

- When the new database is being created, the software will:
  1. Locate the corresponding **master database**.
  2. Copy specific **predefined tables** from the master database into the new database.

> ⚠️ Make sure the master database is in the **same directory** or an accessible path for the software to locate it.

- **Specify the list of source tables to be copied** in the appropriate configuration file or code section.


- ### Step 4: Setting variables
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

These are the required standards and formatting rules for preparing data dictionaries used in surveys and forms.

---

## 1. Header and Worksheet Structure

- The **first row** of each worksheet **must be the header row**.
- All **non-question rows** (e.g. section headers, instructions) **must have merged cells**.
- Worksheets that contain a data dictionary must **end with `_dd` or `_xml`** in their name.

---

## 2. Required Columns and Format

- The data dictionary must always have **13 columns**, with the following fields (order must be consistent).
- Fields for `DontKnow`, `Refuse`, and `NA` must each be in **separate columns**, and can only be **TRUE** or left **blank**.
- `MaxCharacters` must be specified for **QuestionType = text** and **FieldType = text**.
- `LowerRange` and `UpperRange` must either both be **numeric values** or both left **blank**.
- For all multiple-choice fields (`radio`, `checkbox`, `combobox`), **responses must begin with `1:`** (e.g. `1:Yes`, `2:No`).

---

## 3. Question Types

These define how a question appears to the interviewer/respondent.

| **QuestionType** | **Description** |
|------------------|------------------|
| `radio`          | Radio Buttons – FieldType must be `integer` |
| `combobox`       | Dropdown menu – FieldType must be `integer` |
| `checkbox`       | Checkboxes – FieldType must be `text` |
| `text`           | Text Box – Must include `MaxCharacters` |
| `date`           | Date Picker |
| `information`    | Displays information only – not saved to database |
| `automatic`      | Automatically answered by software – logic must be added to `AddAutomatic()` |

---

## 4. Field Types

These define how the data is stored in the database.

| **FieldType**   | **Description** |
|------------------|------------------|
| `text`           | Short Text – Accepts any characters (default 255) |
| `datetime`       | Date/Time |
| `date`           | Date only |
| `phone_num`      | Short Text – Only numbers allowed; 10 characters |
| `integer`        | Long Integer |
| `text_integer`   | Long Integer – Only numbers allowed in input |
| `text_id`        | Text – Only numeric values allowed |
| `text_decimal`   | Decimal – Allows numbers and decimal point; precision = 13, scale = 5 |
| `hourmin`        | Short Text – Only numbers and colon allowed (format HH:MM) |

---

## 5. Responses

- Multiple-choice responses must follow the format `1:Yes`, `2:No`, etc.
- Do **not** include `DontKnow`, `Refuse`, or `NA` as response options in **radio** or **checkbox** fields. These are captured in their own columns.

---

## 6. Skip Logic

Use the following format for skip patterns:

```
skiptype: if fieldname_to_check condition value, skip to fieldname_to_skip_to
```

**Rules:**
- `skiptype`: either `preskip` or `postskip`
- `condition`: one of `=`, `<`, `>`, `<=`, `>=`, `<>`, `contains`, `does not contain`
- Must use **single spaces** between each element
- Example:  
  ```
  postskip: if gender = 2, skip to pregnancy_status
  ```

---

## 7. Logic Checks

Use logic checks to validate relationships between responses.

- **Dynamic logic check (across questions):**  
  ```
  if intvinit2 <> intvinit, error_message This does not match your previous entry!
  ```

- **Fixed logic check (internal to one field):**  
  ```
  if month = 2 'and' day = 30, error_message throw an error
  ```

