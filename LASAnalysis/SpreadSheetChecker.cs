using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using Excel = Microsoft.Office.Interop.Excel;
using VBA = Microsoft.Vbe.Interop;

namespace LASAnalysis
{
    class SpreadSheetChecker : VisualBasicSyntaxWalker, IDisposable
    {
        public SpreadSheetChecker()
        {
            // Initiate the Excel application for checking.
            this.excelApp = new Excel.Application();

            // Set up the error file and the directory it should be in.
            this.errorFile = ConfigurationManager.AppSettings["ErrorFile"];

            // Make sure to have the directory created, if it does not exist.
            Directory.CreateDirectory(Path.GetDirectoryName(errorFile));

            // Get the setting for what operation to do.
            bool doMacroParsing = Convert.ToBoolean(ConfigurationManager.AppSettings["MacroParsing"]);
            if (doMacroParsing)
            {
                DoMacroParsing();
            }
            else
            {
                DoCorrectnessChecking();
            }            
        }

        // Override functions for VB syntax walker.
        public override void VisitAssignmentStatement(AssignmentStatementSyntax node)
        {
            base.VisitAssignmentStatement(node);

            ExpressionSyntax left = node.Left as ExpressionSyntax;
            if (left.ToString().Equals("ActiveCell.FormulaR1C1"))
            {
                ExpressionSyntax right = node.Right as ExpressionSyntax;
                string formulaString = right.ToString();
                if (formulaString.StartsWith("\"="))
                {
                    // Remove the leading and trailing quotes and replace all double quotes with single ones.
                    formulaString = formulaString.Substring(1, formulaString.Length - 2);
                    formulaString = formulaString.Replace("\"\"", "\"");
                    macroCellFormula.Add(formulaString);
                }
            }
        }

        public override void VisitExpressionStatement(ExpressionStatementSyntax node)
        {
            base.VisitExpressionStatement(node);

            // Get the first syntax node.
            foreach (SyntaxNode child in node.ChildNodes())
            {
                // Operations are of invocation expression kind.
                if (child.IsKind(SyntaxKind.InvocationExpression))
                {
                    InvocationExpressionSyntax expr = child as InvocationExpressionSyntax;
                    string exprText = expr.Expression.ToString();

                    // Split the expression on the dot token.
                    string[] exprTokens = exprText.Split('.');

                    //Debug.Assert(exprTokens.Length == 2, "Nested member access.");

                    // Add the last token as the operation.
                    macroOperations.Add(exprTokens[exprTokens.Length - 1]);
                }
            }
        }

        // Workhorse function for parsing macros in the input file sets.
        private void DoMacroParsing()
        {
            // Initialize the output file.
            this.macroInfoOutputFile = ConfigurationManager.AppSettings["MacroFile"];

            // Initialize the macro result builder.
            this.macroInfoBuilder = new StringBuilder();

            // Add headers.
            this.macroInfoBuilder.Append("File Name,Macro Name,Fomulas Used in Macro,Operations Used in Macro"
                                            + Environment.NewLine);

            // Initiaize the collections.
            this.macroOperations = new HashSet<string>();
            this.macroCellFormula = new List<string>();

            Console.WriteLine("Do not disturb, parsing VB macros in excel file.");

            foreach (string filePath in GetAllExcelMacroFiles())
            {
                ParseVBMacro(filePath);
            }

            // Write out the results.
            File.WriteAllText(macroInfoOutputFile, macroInfoBuilder.ToString());

            Console.WriteLine("Done with that horrible stuff, just put me out of misery.");
        }

        private void ParseVBMacro(string filePath)
        {
            Excel.Workbook workBook = null;

            try
            {
                if (this.excelApp != null)
                {
                    workBook = this.excelApp.Workbooks.Open(filePath);

                    // Check if we have VB macros or not.
                    if (workBook.HasVBProject)
                    {
                        // Get the project.
                        VBA.VBProject project = workBook.VBProject;
                        
                        // Process each component in project.
                        foreach (VBA.VBComponent component in project.VBComponents)
                        {
                            ParseVBComponent(filePath, component);
                        }                                              
                    }
                }
            }
            catch (Exception e)
            {
                // Dump the message on console.
                Console.WriteLine(e.Message);

                // Write the actual exception message to log file.
                File.AppendAllText(errorFile, e.Message + Environment.NewLine);
            }
            finally
            {
                // Close the workbook without saving anything.
                if (workBook != null)
                {
                    workBook.Close(false);
                }

                if (workBook != null)
                {
                    Marshal.ReleaseComObject(workBook);
                }
            }
        }

        private void ParseVBComponent(string filePath, VBA.VBComponent component)
        {
            if (component != null)
            {
                // Get the file name prefix to remove.
                string filePrefix = ConfigurationManager.AppSettings["RemoveNamePrefix"];

                VBA.vbext_ProcKind procedureType = VBA.vbext_ProcKind.vbext_pk_Proc;
                VBA.CodeModule componentCode = component.CodeModule;

                // Clear out the containers.
                macroOperations.Clear();
                macroCellFormula.Clear();

                string procedureName = "";
                for (int line = 1; line < componentCode.CountOfLines; line++)
                {
                    // Name of the macro procedure.
                    procedureName = componentCode.get_ProcOfLine(line, out procedureType);
                    
                    if (procedureName != string.Empty)
                    {
                        int procedureLines = componentCode.get_ProcCountLines(procedureName, procedureType);
                        int procedureStartLine = componentCode.get_ProcStartLine(procedureName, procedureType);

                        string procedureBody = componentCode.get_Lines(procedureStartLine, procedureLines);
                        CompilationUnitSyntax root = VisualBasicSyntaxTree.ParseText(procedureBody).GetRoot()
                                                        as CompilationUnitSyntax;

                        /* TODO: If we wanted to iterate through the statements of the sub block.
                        // Get the statements within the sub block.
                        MethodBlockSyntax macro = root.Members[0] as MethodBlockSyntax;
                        SyntaxList<StatementSyntax> statements = (SyntaxList<StatementSyntax>)macro.Statements;

                        int statementCount = 0;
                        foreach (StatementSyntax statement in statements)
                        {
                            Console.WriteLine(statement.Kind().ToString());
                            Console.WriteLine(i.ToString() + " : " + statement.ToString());
                            statementCount++;
                        }
                        */

                        // Do the parsing inside the syntax walker.
                        this.Visit(root);

                        line += procedureLines - 1;
                    }
                }

                if (macroCellFormula.Count > 0 || macroOperations.Count > 0)
                {
                    // Add the file name.
                    macroInfoBuilder.Append(filePath.Replace(filePrefix, ""));
                    macroInfoBuilder.Append("," + procedureName);

                    // Check if we have information for this component and add them.
                    if (macroCellFormula.Count > 0)
                    {
                        macroInfoBuilder.Append("," + EscapeCsvData(string.Join(", ", macroCellFormula)));
                    }
                    else
                    {
                        macroInfoBuilder.Append(",");
                    }

                    if (macroOperations.Count > 0)
                    {
                        macroInfoBuilder.Append("," + EscapeCsvData(string.Join(", ", macroOperations)));
                    }
                    else
                    {
                        macroInfoBuilder.Append(",");
                    }

                    macroInfoBuilder.Append(Environment.NewLine);
                }                   
            }
        }

        private static string EscapeCsvData(string data)
        {
            if (data.Contains("\""))
            {
                data = data.Replace("\"", "\"\"");
            }

            if (data.Contains(","))
            {
                data = String.Format("\"{0}\"", data);
            }

            if (data.Contains(Environment.NewLine))
            {
                data = String.Format("\"{0}\"", data);
            }

            return data;
        }


        private void DoCorrectnessChecking()
        {
            // Initiate the result buffers.
            this.summaryBuilder = new StringBuilder();
            this.detailedResultBuilder = new StringBuilder();

            // Set up the file names from configuration.
            this.summaryOutputFile = ConfigurationManager.AppSettings["SummaryFile"];
            this.detailedResultOutputFile = ConfigurationManager.AppSettings["DetailedFile"];

            // Set up the checking parameters.            
            this.columnToCheck = ConfigurationManager.AppSettings["ColToCheck"];
            this.startColumnIndex = Int32.Parse(ConfigurationManager.AppSettings["ColStartIndex"]);
            this.endColumnIndex = Int32.Parse(ConfigurationManager.AppSettings["ColEndIndex"]);

            // Add headers for output files.
            summaryBuilder.Append("Filename,No. of Incorrect Input" + System.Environment.NewLine);
            detailedResultBuilder.Append("Filename,Incorrect Cell,Given Input,Correct Answer" + System.Environment.NewLine);

            // Load the answer keys.
            LoadAnswerKey();

            Console.WriteLine("Fetching refracted light from Saturn's rings, and analyzing some spreadsheets...");

            // Do the actual correctness checking.
            CheckInputCorrectness();

            Console.WriteLine("Done with the spreadsheet analysis. No update from Saturn though.");
        }

        // Returns list of all .xls* files' absolute path for the input directory.
        private IEnumerable<string> GetAllExcelMacroFiles()
        {
            string supportedFiletypes = "*.xls,*.xlsx,*.xlsm";
            return Directory.GetFiles(ConfigurationManager.AppSettings["InputFilesDir"], "*.*", 
                                        SearchOption.AllDirectories).Where(s => 
                                        supportedFiletypes.Contains(Path.GetExtension(s).ToLower()));       
        }

        // Loads the answer key from the answer file into a dictionary as <cell_location, value> pairs.
        private void LoadAnswerKey()
        {
            // Get the range C2:C62 
            answerMap = new Dictionary<string, string>();
            GetCellRangeValues(ConfigurationManager.AppSettings["AnswerFile"], answerMap);
        }

        // Workhorse program, goes over every single file, opens them and checks 
        // the desired range with the answer key we loaded.
        private void CheckInputCorrectness()
        {
            Dictionary<string, string> rangeMap = new Dictionary<string, string>();
            foreach (string inputFile in GetAllExcelMacroFiles())
            {
                // Get the required range.
                GetCellRangeValues(inputFile, rangeMap);

                int incorrectCount = 0;
                Dictionary<string, string> incorrectInputMap = new Dictionary<string, string>();
                foreach (KeyValuePair<string,string> entry in answerMap)
                {
                    string answer = entry.Value.ToString().Trim();
                    string input;
                    if (!rangeMap.TryGetValue(entry.Key, out input))
                    {
                        // Input does not have this cell. No further processing needed, move to next cell.
                        incorrectCount++;
                        continue;
                    }

                    if (!answer.Equals(input.Trim()))
                    {
                        incorrectCount++;
                        incorrectInputMap.Add(entry.Key, input);
                    }
                }

                // Process the result if there is any incorrect input.
                if (incorrectCount > 0)
                {
                    ProcessCorrectnessResult(inputFile, incorrectCount, incorrectInputMap);
                }                
            }

            // Write out the result to file.
            File.WriteAllText(summaryOutputFile, summaryBuilder.ToString());
            File.WriteAllText(detailedResultOutputFile, detailedResultBuilder.ToString());
        }

        // Processes the correctness result of an input file. Given are the file path, no. of incorrect cells
        // and the cells which had the incorrect value along with that incorrect value.
        private void ProcessCorrectnessResult(string filePath, int incorrectCount, Dictionary<string, string> incorrectInputMap)
        {
            // Add to the summary.
            summaryBuilder.Append(filePath.Replace(ConfigurationManager.AppSettings["RemoveNamePrefix"], "") 
                                    + "," + incorrectCount.ToString() + Environment.NewLine);

            // Add the incorrect cell with the input and the correct answer.
            foreach (KeyValuePair<string, string> incorrectInput in incorrectInputMap)
            {
                detailedResultBuilder.Append(filePath.Replace(ConfigurationManager.AppSettings["RemoveNamePrefix"], "") 
                                            + "," + incorrectInput.Key.ToString() + ","
                                            + incorrectInput.Value.ToString() + "," 
                                            + answerMap[incorrectInput.Key.ToString()] + Environment.NewLine);
            }            
        }

        private void GetCellRangeValues(string filePath, Dictionary<string, string> rangeMap)
        {
            Excel.Workbook workBook = null;
            Excel.Worksheet sheet = null;
            
            // Clear out the dictionary.
            rangeMap.Clear();

            try
            {
                if (this.excelApp != null)
                {
                    // Load the answer key file.
                    workBook = this.excelApp.Workbooks.Open(filePath);

                    // Get the active sheet, there should be only one (as expected).
                    sheet = (Excel.Worksheet)workBook.ActiveSheet;
                    
                    for (int cellIndex = this.startColumnIndex; cellIndex <= this.endColumnIndex; cellIndex++)
                    {
                        string cellLocation = this.columnToCheck + cellIndex.ToString();
                        Excel.Range keyCell = (Excel.Range)sheet.Range[cellLocation];
                        Object keyCellValue = keyCell.Value2;
                        if (keyCellValue != null)
                        {
                            String cellValue = keyCellValue.ToString();
                            if (cellValue.Contains(","))
                            {
                                // Escape cell text with commas, so that they don't mess up the csv.
                                cellValue = "\"" + cellValue + "\"";
                            }
                            rangeMap.Add(cellLocation, cellValue);
                        }
                        else
                        {
                            rangeMap.Add(cellLocation, "");
                        }                        
                    }
                }
            }
            catch (Exception e)
            {
                // Dump the message on console.
                Console.WriteLine(e.Message);

                // Write the actual exception message to log file.
                File.AppendAllText(errorFile, e.Message + Environment.NewLine);
            }
            finally
            {
                // Close the workbook without saving anything.
                if (workBook != null)
                {
                    workBook.Close(false);
                }
                
                if (sheet != null)
                {
                    Marshal.ReleaseComObject(sheet);
                }

                if (workBook != null)
                {
                    Marshal.ReleaseComObject(workBook);
                }
            }
        }

        public void Dispose()
        {
            // Quit the Excel application and release it.
            if (excelApp != null)
            {
                excelApp.Quit();
            }

            Marshal.ReleaseComObject(excelApp);
        }

        // Result file paths, gets overwritten every time program is run.
        private string summaryOutputFile;
        private string detailedResultOutputFile;
        private string macroInfoOutputFile;

        // Error file.
        private string errorFile;

        // Map of answers with cells to value, eg. <Col#, 'value'> format.
        private Dictionary<string, string> answerMap;

        // Collections to contain VB component infromations.
        private HashSet<string> macroOperations;
        private List<string> macroCellFormula;

        // Single Excel interop app, so that we don't keep on opening stuff.
        private Excel.Application excelApp;

        // Buffers that holds the results in place until dumped on disk.
        private StringBuilder summaryBuilder;
        private StringBuilder detailedResultBuilder;
        private StringBuilder macroInfoBuilder;

        // Cell range to check.
        private string columnToCheck;
        private int startColumnIndex;
        private int endColumnIndex;
    }
}
