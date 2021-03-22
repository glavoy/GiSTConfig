using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ADOX;
using Microsoft.Office.Interop.Excel;

namespace generatexml
{
    public partial class Main : Form
    {
        // Initialize the form
        public Main()
        {
            InitializeComponent();
        }

        // Version
        readonly string swVer = "2021-03-20";

        // Path to Excel file
        //readonly string excelFile = "D:\\Sapphire Dev Tools\\GistConfig\\excel\\GistTest.xlsx";
        // Feel free to change this to whatever you want
        //readonly string excelFile = "D:\\Sapphire Dev Tools\\Sapphire Data Dictionary\\SAPPHIRE_Treatment_HTN_2021_03_13.xlsx";
        readonly string excelFile = "D:\\Sapphire Dev Tools\\Sapphire Data Dictionary\\SAPPHIRE_Prevention_2021_03_20.xlsx";
        //readonly string excelFile = "C:\\Temp\\COVID_Surveillance_Data_Dictionary_2020_12_03.xlsx";
        //

        // Path to XML file
        //readonly string xmlPath = "D:\\Sapphire Dev Tools\\xml\\test\\";
        readonly string xmlPath = "D:\\Sapphire Dev Tools\\xml\\dcp\\";
        //readonly string xmlPath = "D:\\Sapphire Dev Tools\\xml\\htn\\";
        

        // Path to log file
        readonly string logfilePath = "D:\\Sapphire Dev Tools\\";

        // Path to MS Access database
        //readonly string accessDB = "D:\\Sapphire Dev Tools\\Access db\\test.mdb";
        readonly string accessDB = "D:\\Sapphire Dev Tools\\Access db\\dynamicprevention.mdb";
        //readonly string accessDB = "D:\\Sapphire Dev Tools\\Access db\\HTNLinkage.mdb";

        //log string
        public List<string> logstring = new List<string>();


        // Question class
        // There is a new Question object created for each question in the Excel file
        // Each Question object is added to the QuestionList List.
        public class Question
        {
            public string fieldName;
            public string questionType;
            public string fieldType;
            public string questionText;
            public int maxCharacters;
            public string responses;
            public int lowerRange;
            public int upperRange;
            public string logicCheck;
            public Boolean dontKnow;
            public Boolean refuse;
            public Boolean na;
            public string skip;
        }


        // List of Question objects
        public List<Question> QuestionList = new List<Question>();

        // Show software version when form loads
        private void Main_Load(object sender, EventArgs e)
        {
            // Show version
            labelVersion.Text = string.Concat("Version: ", swVer);
        }


        // Function when button is clicked
        private void ButtonXML_Click(object sender, EventArgs e)
        {
            try
            {


                // Use a wait cursor
                Cursor.Current = Cursors.WaitCursor;

                //start logging of any error
                logstring.Add("Checking field properties for " + excelFile);

                // Delete the Access database if it exists
                if (File.Exists(accessDB))
                {
                    File.Delete(accessDB);
                }


                // Create the Access database
                string connectionString = string.Format("Provider={0}; Data Source={1}; Jet OLEDB:Engine Type={2}",
                                                        "Microsoft.Jet.OLEDB.4.0",
                                                        accessDB,
                                                        5);
                ADOX.CatalogClass cat = new ADOX.CatalogClass();
                cat.Create(connectionString);
                cat = null;





                // Set up parameters for opening Excel file
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                //Excel.Worksheet xlWorkSheet;
                Excel.Range range;
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(@excelFile, 0, true, 5, "", "", true,
                                                  Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,
                                                  "\t", false, false, 0, true, 1, 0);


                // Read each sheet of the Excel file and generate list of questions
                foreach (Worksheet worksheet in xlWorkBook.Worksheets)
                {
                    // Data dictionaries must end in '_dd'
                    if (worksheet.Name.Substring(worksheet.Name.Length - 3) == "_dd")
                    {
                        // Get the range of used cellsin the Excel file
                        range = worksheet.UsedRange;

                        // Variable to get the total number of rows used in the Excel file
                        int numRows = range.Rows.Count;

                        // Used to determine if a row is merged or not
                        // All rows that are not questions, must be merged
                        Range rowRange = null;


                        // Clear the previous QuestionList, if it existed
                        QuestionList.Clear();


                        // Start at the second row and iterate through each row (question)
                        // and create a question object for each question.
                        // Each question object is added to the QuestionList list.
                        for (int rowCount = 2; rowCount <= numRows; rowCount++)
                        {
                            rowRange = worksheet.Cells[rowCount, 14];
                            if (!rowRange.MergeCells)
                            {
                                var curQuestion = new Question
                                {
                                    fieldName = range.Cells[rowCount, 1] != null && range.Cells[rowCount, 1].Value2 != null ? range.Cells[rowCount, 1].Value2.ToString() : "",
                                    questionType = range.Cells[rowCount, 2] != null && range.Cells[rowCount, 2].Value2 != null ? range.Cells[rowCount, 2].Value2.ToString() : "",
                                    fieldType = range.Cells[rowCount, 3] != null && range.Cells[rowCount, 3].Value2 != null ? range.Cells[rowCount, 3].Value.ToString() : "",
                                    questionText = range.Cells[rowCount, 4] != null && range.Cells[rowCount, 4].Value2 != null ? range.Cells[rowCount, 4].Value2.ToString() : "",
                                    maxCharacters = range.Cells[rowCount, 5] != null && range.Cells[rowCount, 5].Value2 != null ? (int)range.Cells[rowCount, 5].Value2 : -9,
                                    responses = range.Cells[rowCount, 6] != null && range.Cells[rowCount, 6].Value2 != null ? range.Cells[rowCount, 6].Value2.ToString() : "",
                                    lowerRange = range.Cells[rowCount, 7] != null && range.Cells[rowCount, 7].Value2 != null ? (int)range.Cells[rowCount, 7].Value2 : -9,
                                    upperRange = range.Cells[rowCount, 8] != null && range.Cells[rowCount, 8].Value2 != null ? (int)range.Cells[rowCount, 8].Value2 : -9,
                                    logicCheck = range.Cells[rowCount, 9] != null && range.Cells[rowCount, 9].Value2 != null ? range.Cells[rowCount, 9].Value2.ToString() : "",
                                    dontKnow = range.Cells[rowCount, 10] != null && range.Cells[rowCount, 10].Value2 != null ? (Boolean)range.Cells[rowCount, 10].Value2 : false,
                                    refuse = range.Cells[rowCount, 11] != null && range.Cells[rowCount, 11].Value2 != null ? (Boolean)range.Cells[rowCount, 11].Value2 : false,
                                    na = range.Cells[rowCount, 12] != null && range.Cells[rowCount, 12].Value2 != null ? (Boolean)range.Cells[rowCount, 12].Value2 : false,
                                    skip = range.Cells[rowCount, 13] != null && range.Cells[rowCount, 13].Value2 != null ? range.Cells[rowCount, 13].Value2.ToString() : ""
                                };

                                // Uncomment the following twp lines to debug the individual field issues
                                //string fName = range.Cells[rowCount, 1] != null && range.Cells[rowCount, 1].Value2 != null ? range.Cells[rowCount, 1].Value2.ToString() : "";
                                //Console.WriteLine("Creating field: " + fName + " from: " +  worksheet.Name);
                                // Add the question to the list
                                QuestionList.Add(curQuestion);
                            }
                        }

                        Console.WriteLine("Done Creating question list for "+ worksheet.Name);
                        //Check for duplicate columns in the question list b4 moving on
                        checkDuplicateColumns(QuestionList, worksheet.Name.Substring(0, worksheet.Name.Length - 3));
                        // Write to the XML file
                        WriteXML(worksheet.Name);

                        // Add table to database
                        CreateDatabase(worksheet.Name);
                        
                    }
                }

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(true, null, null);
                xlApp.Quit();
                //Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
                logstring.Add("Done Building the xml file and the database, Check the xml and db folders");
                writeLogfile();
                MessageBox.Show("Done Building the xml file and the database, Check the xml, db folders and gistlogfile for any errors");
            }



            // Error handling in caase we could not crread the Excel file
            catch (Exception ex)
            {
                Console.WriteLine("Error msg " + ex.Message);
                Console.WriteLine("Error code " + ex.HResult);
                MessageBox.Show("There is an error with the MS Excel data dictionary.");
                logstring.Add("There is an error with the MS Excel data dictionary.");
                logstring.Add("Error msg " + ex.Message);
                logstring.Add("Error code " + ex.HResult);
            }

            // Put the cursor back to normal
            Cursor.Current = Cursors.Default;
            

        }


        //////////////////////////////////////////////////////////////////////
        // Function to create and write to XML file
        //////////////////////////////////////////////////////////////////////
        private void WriteXML(string xmlfilename)
        {
            try
            {
                xmlfilename = xmlfilename.Substring(0, xmlfilename.Length - 3);

                // These are strings for the first two of lines in the xml file
                string[] xmlStart = { "<?xml version = '1.0' encoding = 'utf-8'?>", "<survey>" };

                // Open a XML file and start writing lines of text to it
                using (StreamWriter outputFile = new StreamWriter(string.Concat(xmlPath, xmlfilename, ".xml")))
                {
                    // Write the first 2 lines to the XML file
                    foreach (string line in xmlStart)
                        outputFile.WriteLine(line);

                    // Write a blank line 
                    outputFile.WriteLine("\n");

                    // Check to ensure that all questions are correctly define
                    checkQuestionType(QuestionList, xmlfilename);


                    // Iterate through each question object in the QuestionList list
                    // and write the necessary text to the XML file
                    foreach (Question question in QuestionList)
                    {
                        // Write the main part of the question
                        // Uses questionType, fieldName and fieldType
                        outputFile.WriteLine(string.Concat("\t<question type = '", question.questionType,
                                                           "' fieldname = '", question.fieldName,
                                                           "' fieldtype = '", question.fieldType, "'>"));


                        // Write the text if it is not an automatic question
                        if (question.questionType != "automatic")
                            outputFile.WriteLine(string.Concat("\t\t<text>", question.questionText, "</text>"));


                        // The maximum characters if necessary
                        if (question.maxCharacters != -9)
                            outputFile.WriteLine(string.Concat("\t\t<maxCharacters>", question.maxCharacters, "</maxCharacters>"));


                        // Upper and Lower range (numeric check)
                        if (question.lowerRange != -9)
                        {
                            outputFile.WriteLine("\t\t<numeric_check>");
                            outputFile.WriteLine(string.Concat("\t\t\t<values minvalue ='", question.lowerRange,
                                                               "' maxvalue='", question.upperRange,
                                                               "' other_values = '", question.lowerRange,
                                                               "' message = 'Number must be between ", question.lowerRange,
                                                               " and ", question.upperRange, "!'></values>"));
                            outputFile.WriteLine("\t\t</numeric_check>");
                        }

                        //  Logic Checks (Added by werick)                
                        if (question.logicCheck != "")
                        {
                            // This stores the text for the skip
                            string[] logicChecks = question.logicCheck.Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

                            // Lists to store logic checks
                            List<string> dynamicLogicCheck = new List<string>();
                            List<string> fixedLogicCheck = new List<string>();


                            // Populate the list for each type of logic checks
                            foreach (string check in logicChecks)
                            {
                                int index = check.IndexOf(@":");

                                if (check.Substring(0, index) == "dynamic")
                                    dynamicLogicCheck.Add(check);

                                if (check.Substring(0, index) == "fixed")
                                    fixedLogicCheck.Add(check);

                            }


                            // Create text dynamic
                            if (dynamicLogicCheck.Count > 0)
                            {
                                outputFile.WriteLine("\t\t<logic_check>");
                                foreach (string logic in dynamicLogicCheck)
                                {
                                    // Call the GenerateSkips() function
                                    outputFile.WriteLine(GenerateLogicChecks(logic, "dynamic"));
                                }
                                outputFile.WriteLine("\t\t</logic_check>");
                            }

                            // Create text dynamic
                            if (fixedLogicCheck.Count > 0)
                            {
                                outputFile.WriteLine("\t\t<logic_check>");
                                foreach (string logic in fixedLogicCheck)
                                {
                                    // Call the GenerateSkips() function
                                    outputFile.WriteLine(GenerateLogicChecks(logic, "fixed"));
                                }
                                outputFile.WriteLine("\t\t</logic_check>");
                            }

                        }

                        // Write responses if it is a radio or checkbox type question
                        if (question.questionType == "radio" || question.questionType == "checkbox")
                        {
                            outputFile.WriteLine("\t\t<responses>");
                            string[] responses = question.responses.Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

                            if (responses.Length == 0)
                            {
                                outputFile.WriteLine("\t\t\t<response></response>");
                            }
                            else
                            {
                                foreach (string response in responses)
                                {
                                    int index = response.IndexOf(@":");
                                    outputFile.WriteLine(string.Concat("\t\t\t<response value = '", response.Substring(0, index), "'>",
                                                                        response.Substring(index + 1).Trim(), "</response>"));
                                }
                            }

                            outputFile.WriteLine("\t\t</responses>");
                        }


                        // Skips
                        if (question.skip != "")
                        {
                            // This stores the text for the skip
                            string[] skips = question.skip.Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

                            // Lists to store preskips and postskips
                            List<string> preSkips = new List<string>();
                            List<string> postSkips = new List<string>();


                            // Populate the list for each type of skip
                            foreach (string skip in skips)
                            {
                                int index = skip.IndexOf(@":");

                                if (skip.Substring(0, index) == "preskip")
                                    preSkips.Add(skip);

                                if (skip.Substring(0, index) == "postskip")
                                    postSkips.Add(skip);
                            }


                            // Create text preskips
                            if (preSkips.Count > 0)
                            {
                                outputFile.WriteLine("\t\t<preskip>");
                                foreach (string preSkip in preSkips)
                                {
                                    // Call the GenerateSkips() function
                                    outputFile.WriteLine(GenerateSkips(preSkip, "preSkip"));
                                }
                                outputFile.WriteLine("\t\t</preskip>");
                            }


                            // Create text postskips
                            if (postSkips.Count > 0)
                            {
                                outputFile.WriteLine("\t\t<postskip>");
                                // Call the GenerateSkips() function
                                foreach (string postSkip in postSkips)
                                {
                                    outputFile.WriteLine(GenerateSkips(postSkip, "postSkip"));
                                }
                                outputFile.WriteLine("\t\t</postskip>");
                            }
                        }



                        // Don't know
                        if (question.dontKnow == true)
                            outputFile.WriteLine("\t\t<dont_know>-7</dont_know>");

                        // Refuse to answer
                        if (question.refuse == true)
                            outputFile.WriteLine("\t\t<refuse>-8</refuse>");

                        // Not applicable
                        if (question.na == true)
                            outputFile.WriteLine("\t\t<na>-6</na>");

                        // Close off the question
                        outputFile.WriteLine("\t</question>");
                        outputFile.WriteLine("\n");
                    }

                    // The last 'info' question ending every survey
                    string[] xmlEnd = {"\t<question type = 'information' fieldname = 'end_of_questions' fieldtype = 'n/a'>",
                                   "\t\t<text>Press the 'Next' button to save the data.</text >", "\t</question>" };
                    foreach (string line in xmlEnd)
                        outputFile.WriteLine(line);

                    outputFile.WriteLine("\n");
                    outputFile.WriteLine("</survey>");
                }
            }


            // Error handling in caase we could not create the XML file
            catch (Exception ex)
            {
                Console.WriteLine("Error msg " + ex.Message);
                MessageBox.Show("Could not create XML file. Ensure path is correct.");
                logstring.Add("Could not create XML file. Ensure path is correct.");
                logstring.Add("Error msg " + ex.Message + " Table " + xmlfilename);
                logstring.Add("Error code " + ex.HResult);
            }
        }



        //////////////////////////////////////////////////////////////////////
        // Function to generate the text for the skips
        //////////////////////////////////////////////////////////////////////
        private string GenerateSkips(string skip, string skipType)
        {
            // Number of initial characters depending on whether it's a preskip or postskip
            int lenSkip = skipType == "postSkip" ? 13 : 12;


            // Create a list to store the index of each 'space' in the skip text
            var spaceIndices = new List<int>();

            // Populate the spaceIndices list
            for (int i = 0; i < skip.Length; i++)
                if (skip[i] == ' ') spaceIndices.Add(i);


            // Get the name of the field to check for skip
            string fieldname_to_check = skip.Substring(lenSkip, spaceIndices[2] - spaceIndices[1] - 1);

            // Variables to store the condition and the value of the skip
            string condition;
            string value;

            // If there are 9 spaces, then we know that the condition is 'does not contain'
            if (spaceIndices.Count == 9)
            {
                // Get the condition
                condition = "does not contain";
                // Get the value
                value = skip.Substring(spaceIndices[5] + 1, spaceIndices[6] - spaceIndices[5] - 2);
            }
            // Check if the skip has 'contains'
            else if (skip.Contains("contains"))
            {
                // Get the condition
                condition = "contains";
                // Get the value
                value = skip.Substring(spaceIndices[3] + 1, spaceIndices[4] - spaceIndices[3] - 2);
            }
            // Skip does not have 'does not contain' or 'contains'
            else
            {
                // Get the condition
                condition = skip.Substring(spaceIndices[2] + 1, spaceIndices[3] - spaceIndices[2] - 1);

                // Replace '<' and '>' symbols, if necessary
                condition = condition.Replace("<", "&lt;");
                condition = condition.Replace(">", "&gt;");

                // Get the value
                value = skip.Substring(spaceIndices[3] + 1, spaceIndices[4] - spaceIndices[3] - 2);
            }

            // Get the field name to skip to
            string fieldname_to_skip_to = skip.Substring(spaceIndices[spaceIndices.Count - 1] + 1);

            // Build the string and return it
            return string.Concat("\t\t\t<skip fieldname='", fieldname_to_check, 
                                 "' condition = '", condition,
                                 "' response='", value,
                                 "' response_type='fixed' skiptofieldname ='",
                                 fieldname_to_skip_to, "'></skip>");
        }


        // Added by Werick
        //////////////////////////////////////////////////////////////////////
        // Function to generate the text for the skips
        //////////////////////////////////////////////////////////////////////
        private string GenerateLogicChecks(string logic, string logicType)
        {
            // Number of initial characters depending on whether it's a preskip or postskip
            int lenSkip = logicType == "dynamic" ? 12 : 10;
            

            // uncomment to debug this section
            //Console.WriteLine("Logic String: " + logic);

            //Split the logi string into two parts: 1 - the logic condition and the error message
            int index = logic.IndexOf(@",");

            string message_section = logic.Substring(index + 1);
            string logic_section = logic.Substring(0,index);

            // uncomment to debug this section
            //Console.WriteLine("Logic Section: " + logic_section);
            //Console.WriteLine("Error Message Section: " + message_section);


            // Create a list to store the index of each 'space' in the skip text
            var spaceIndices = new List<int>();
            var spaceIndicesLogic = new List<int>();
            var spaceIndicesMessage = new List<int>();

            // Populate the spaceIndices list
            for (int i = 0; i < logic.Length; i++)
                if (logic[i] == ' ') spaceIndices.Add(i);

            for (int i = 0; i < logic_section.Length; i++)
                if (logic_section[i] == ' ') spaceIndicesLogic.Add(i);

            for (int i = 0; i < message_section.Length; i++)
                if (message_section[i] == ' ') spaceIndicesMessage.Add(i);


            // Get the name of the field to check for skip
            string fieldname_to_check = logic.Substring(lenSkip, spaceIndices[2] - spaceIndices[1] - 1);

            // Variables to store the condition and the value of the skip
            string condition;
            string value;
            string currentresponse = "1";

            //set the current response for fixed logic types
            if (logicType == "fixed")
            {
                currentresponse = logic_section.Substring(spaceIndicesLogic[spaceIndicesLogic.Count - 1] + 1);
            }

            // If there are 6 spaces in the logic section, then we know that the condition is 'does not contain'
            if (logic_section.Contains("does not contain"))
            {
                // Get the condition
                condition = "does not contain";
                // Get the value
                value = logic_section.Substring(spaceIndicesLogic[spaceIndicesLogic.Count - 1] + 1);
            }
            // Check if the skip has 'contains'
            else if (logic_section.Contains("contains"))
            {
                // Get the condition
                condition = "contains";
                // Get the value
                value = logic_section.Substring(spaceIndicesLogic[spaceIndicesLogic.Count - 1] + 1);
                
            }

            // Check if the skip has 'and'
            else if (logic_section.Contains("'and'"))
            {               
                // Get the condition
                condition = logic.Substring(spaceIndices[2] + 1, spaceIndices[3] - spaceIndices[2] - 1);

                // Replace '<' and '>' symbols, if necessary
                condition = condition.Replace("<", "&lt;");
                condition = condition.Replace(">", "&gt;");

                // Get the value
                value = logic.Substring(spaceIndices[3] + 1, spaceIndices[4] - spaceIndices[3] - 1);


            }
            // Skip does not have 'does not contain' or 'contains'
            else
            {
                // Get the condition
                condition = logic.Substring(spaceIndices[2] + 1, spaceIndices[3] - spaceIndices[2] - 1);

                // Replace '<' and '>' symbols, if necessary
                condition = condition.Replace("<", "&lt;");
                condition = condition.Replace(">", "&gt;");

                // Get the value
                //value = logic.Substring(spaceIndices[3] + 1, spaceIndices[4] - spaceIndices[3] - 2);
                value = logic_section.Substring(spaceIndicesLogic[spaceIndicesLogic.Count - 1] + 1);
            }

            // Get the error message
            string error_message = message_section.Substring(spaceIndicesMessage[1] + 1);

            // Build the string and return it
            return string.Concat("\t\t\t<logic fieldname='", fieldname_to_check,
                                 "' condition = '", condition,
                                 "' response='", value,
                                 "' response_type='", logicType,
                                 "' currentresponse='", currentresponse,
                                 "' message ='", error_message, "'></logic>");
        }


        //////////////////////////////////////////////////////////////////////
        // Function to check any duplicate columns b4 the tables is created
        //////////////////////////////////////////////////////////////////////
        private void checkDuplicateColumns (List<Question> QList, string tblename)
        {
            List<string> list = new List<string>();
            foreach (Question question in QList)
            {
                if (question.questionType != "information")
                {
                    list.Add(question.fieldName);
                }
            }

            var duplicateKeys = list.GroupBy(x => x)
                        .Where(group => group.Count() > 1)
                        .Select(group => group.Key);
            if (list.Count != list.Distinct().Count())
            {
                Console.WriteLine("Duplicate elements are: " + String.Join(",", duplicateKeys) + " In table '" + tblename + "'");
                logstring.Add("Duplicate elements are: " + String.Join(",", duplicateKeys) + " In table '" + tblename + "'");
            } else
            {
                Console.WriteLine("No duplicate keys In table '" + tblename + "'");
                logstring.Add("No duplicate keys In table '" + tblename + "'");
            }          

        }

        /////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // Function to check any of the questions types, field types and corresponding datatypes are wrongly defined
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////
        private void checkQuestionType(List<Question> QList, string tblename)
        {
            string[] qtype = { "radio", "combobox", "checkbox", "text","date","information", "automatic"};
            string[] ftype = { "text", "datetime", "date", "phone_num", "integer", "text_integer", "text_decimal", "text_id", "n/a" };
            foreach (Question question in QList)
            {
                if (!qtype.Contains(question.questionType))
                {
                     logstring.Add("This question type " + question.questionType +  " for field '" + question.fieldName + "' in table '" + tblename + "' is not among the predefined list");
                    Console.WriteLine("This question type " + question.questionType + " for field '" + question.fieldName + "' in table '" + tblename + "' is not among the predefined list");
                }

                if (!ftype.Contains(question.fieldType))
                {
                    logstring.Add("This field type '" + question.fieldType + "' for field '" + question.fieldName + "' in table '" + tblename + "' is not among the predefined list");
                    Console.WriteLine("This field type '" + question.fieldType + "' for field '" + question.fieldName + "' in table '" + tblename + "' is not among the predefined list");
                }

                // check the corresponding data types for all radio question type to ensure they are integer type
                if (question.questionType == "radio")
                {
                    if (question.fieldType != "integer")
                    {
                        logstring.Add("Wrong field Type: The field type for field '" + question.fieldName + "' in table '" + tblename + "' must be integer");
                        Console.WriteLine("Wrong field Type: The field type for field '" + question.fieldName + "' in table '" + tblename + "' must be integer");
                    }
                }

                // check the corresponding data types for all checkbox question type to ensure they are text type
                if (question.questionType == "checkbox")
                {
                    if (question.fieldType != "text")
                    {
                        logstring.Add("Wrong field Type: The field type for field '" + question.fieldName + "' in table '" + tblename + "' must be text");
                        Console.WriteLine("Wrong field Type: The field type for field '" + question.fieldName + "' in table '" + tblename + "' must be text");
                    }
                }

                // check the corresponding data types for all date question type to ensure they are date type
                if (question.questionType == "date")
                {
                    List<string> datetypeslist = new List<string>();
                    datetypeslist.Add("date");
                    datetypeslist.Add("datetime");
                    var match = datetypeslist
                        .FirstOrDefault(stringToCheck => stringToCheck.Contains(question.fieldType));
                    if (match == null)
                    {
                        logstring.Add("Wrong field Type: The field type for field '" + question.fieldName + "' in table '" + tblename + "' must be date");
                        Console.WriteLine("Wrong field Type: The field type for field '" + question.fieldName + "' in table '" + tblename + "' must be date");
                    }
                }


                // check the duplicate responses for radio options
                if (question.questionType == "radio" | question.questionType == "checkbox")
                {
                    //split the list of responses/answers to generate the list/array
                    string[] responses = question.responses.Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

                    if (responses.Length != 0)                    
                    {
                        List<string> list = new List<string>();
                        foreach (string response in responses)
                        {
                            // using the substring function to get the list of keys for responses
                            int index = response.IndexOf(@":");
                            list.Add(response.Substring(0, index));                           
                        }
                        var duplicateKeys = list.GroupBy(x => x)
                                .Where(group => group.Count() > 1)
                                .Select(group => group.Key);
                        if (list.Count != list.Distinct().Count())
                        {
                            logstring.Add("Duplicate Response options: The responses for field '" + question.fieldName + "' in table '" + tblename + "' has duplicates "+ String.Join(",", duplicateKeys));
                            Console.WriteLine("Duplicate Response options: The responses for field '" + question.fieldName + "' in table '" + tblename + "' has duplicates " + String.Join(",", duplicateKeys));
                        }
                    }
                    
                }
            }
        }
        
        private void writeLogfile()
        {
            try
            {
                var logfilename = "gistlogfile";
                // Open a log file and start writing lines of text to it
                using (StreamWriter outputFile = new StreamWriter(string.Concat(logfilePath, logfilename, ".txt")))
                {
                    foreach(string line in logstring)
                        outputFile.WriteLine(line);
                    outputFile.WriteLine("\n");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);               
            }
            
        }

        //////////////////////////////////////////////////////////////////////
        // Function to create the Access database
        //////////////////////////////////////////////////////////////////////
        private void CreateDatabase(string tablename)
        {
            try
            {
                tablename = tablename.Substring(0, tablename.Length - 3);

                // Connection string for database
                string connectionString = string.Format("Provider={0}; Data Source={1}; Jet OLEDB:Engine Type={2}",
                                                        "Microsoft.Jet.OLEDB.4.0",
                                                        accessDB,
                                                        5);

                // Connect to the database
                ADODB.Connection cnn = new ADODB.Connection();
                ADOX.Catalog catalog = new ADOX.Catalog();

                // Open the Connection
                cnn.Open(connectionString);
                catalog.ActiveConnection = cnn;

                // Create a table
                ADOX.Table table = new ADOX.Table
                {
                    Name = tablename
                };

                
                // Create a field name for each question
                foreach (Question question in QuestionList)
                {
                    // Don't need to create a field for 'information' questions
                    if (question.questionType != "information")
                    {
                        // Create a field
                        ADOX.ColumnClass newCol = new ADOX.ColumnClass
                        {
                            // Field name
                            Name = question.fieldName,
                            // Allow column to have NULL values
                            Attributes = ColumnAttributesEnum.adColNullable
                        };

                        // Get the question type and set the column type accordingly
                        switch (question.fieldType)
                        {
                            case "text_integer":
                            case "text_id":
                            case "integer":
                                newCol.Type = ADOX.DataTypeEnum.adInteger;
                                break;
                            case "text":
                                newCol.Type = ADOX.DataTypeEnum.adVarWChar;
                                if (question.maxCharacters != -9)
                                    newCol.DefinedSize = question.maxCharacters;
                                break;
                            case "phone_num":
                                newCol.Type = ADOX.DataTypeEnum.adVarWChar;
                                newCol.DefinedSize = 10;
                                break;
                            case "text_decimal":
                                newCol.Type = ADOX.DataTypeEnum.adNumeric;
                                newCol.Precision = 13;
                                newCol.NumericScale = 5;
                                break;
                            case "date":
                            case "datetime":
                                newCol.Type = ADOX.DataTypeEnum.adDate;
                                break;
                            default:
                                newCol.Type = ADOX.DataTypeEnum.adVarWChar;
                                break;
                        }

                        // Finally, add the field to the table
                        table.Columns.Append(newCol);
                    }
                }

                // Add the table to the database
                catalog.Tables.Append(table);


                // Close the connection
                if (cnn != null && cnn.State != 0)
                    cnn.Close();

                // release memory
                catalog = null;
            }

            // Error handling in case the Access database could not be created
            catch (Exception ex)
            {
                var code = ex.HResult;
                Console.WriteLine("Error msg " + ex.Message + " Table " + tablename);               
                Console.WriteLine("Error code "+ code);
                MessageBox.Show("Could not create database. Check if it already exists.");
                MessageBox.Show(ex.Message);
                logstring.Add("Error msg " + ex.Message + " Table " + tablename);
                logstring.Add("Error code " + code);

            }
        }


    }
}
