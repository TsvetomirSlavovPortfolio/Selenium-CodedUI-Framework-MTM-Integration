// <copyright file="Reporting.cs" company="Infosys Ltd.">
//  Copyright (c) Infosys Ltd. All rights reserved.
// </copyright>
// <summary>Reporting.cs Handles settings of Test Reports.</summary>
namespace INF.Selenium.TestAutomation.Utilities
{
    using System;
    using System.Data.OleDb;
    using System.Drawing;
    using System.IO;
    using System.Reflection;
    using System.Text;
    using System.Web.Configuration;
    using Configuration;
    using Entities;
    using Microsoft.Office.Interop.Excel;

    /// <summary>Reporting.cs Handles settings of Test Reports.</summary>
    public class Reporting
    {
        /// <summary>This Handles settings of Test Reports and styles.</summary>
        private enum ReportingStyles
        {
            /// <summary>Style header.</summary>
            StyleHeader,

            /// <summary>Style for Passed.</summary>
            StylePassed,

            /// <summary>Style for Failed.</summary>
            StyleFailed
        }

        ////public static int Iteration { get; set; }

        /// <summary>
        /// Gets or private sets Test case file path.
        /// </summary>
        /// <value>File path.</value>
        public static string FilePath { get; private set; }

        /// <summary>
        /// Gets or private sets File Path string.
        /// </summary>
        /// <value>File path string.</value>
        public static string PathString { get; private set; }

        /// <summary>
        /// This procedure will Insert TestCase results summary into Summary tab of Excel report.
        /// </summary>
        public static void InsertResultSummary()
        {
            var con = string.Format(WebConfigurationManager.ConnectionStrings[Entities.Constants.Oledb].ConnectionString, FilePath);
            try
            {
                using (var excelConnection = new OleDbConnection(con))
                {
                    using (var command = new OleDbCommand())
                    {
                        excelConnection.Open();
                        command.Connection = excelConnection;
                        command.CommandText = Entities.Constants.Queries.InsertSummary;
                        command.Parameters.AddWithValue(Entities.Constants.TestResultSummary.ApplicationParam, TestCase.Application);
                        command.Parameters.AddWithValue(Entities.Constants.TestResultSummary.TestCaseIdParam, TestCase.Name);
                        command.Parameters.AddWithValue(Entities.Constants.TestResultSummary.DescriptionSummaryParam, TestCase.Description);
                        command.Parameters.AddWithValue(Entities.Constants.TestResultSummary.ResultSummaryParam, Result.GetTestScriptResult());
                        command.Parameters.AddWithValue(Entities.Constants.TestResultSummary.ExecutionDurationParam, Timing.TestCaseDuration);
                        command.Parameters.AddWithValue(Entities.Constants.TestResultSummary.ExecutionDocumentReference, string.IsNullOrEmpty(TestCases.TestCases.Hyplink) ? Entities.Constants.Na : TestCases.TestCases.Hyplink);
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.Reporting, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }

        /// <summary>
        /// This procedure will create the Excel report file.
        /// </summary>
        /// <param name="errorMessage">Error message.</param>
        /// <returns>File path is Null or empty.</returns>
        public static bool CreateExcelFile(ref string errorMessage)
        {
            var applicationClass = new ApplicationClass();
            Workbook workbook = null;
            try
            {
                //// Check if file already exist
                if (!string.IsNullOrEmpty(FilePath))
                {
                    return true;
                }
                //// Create file
                var time = string.Empty + DateTime.Now.ToShortDateString().Replace(Entities.Constants.ForwardSlash, Entities.Constants.Hyphen) + Entities.Constants.Space + DateTime.Now.ToLongTimeString().Replace(Entities.Constants.Colon, Entities.Constants.Hyphen) + string.Empty;
                PathString = Path.Combine(TestCase.RootFilePath + Entities.Constants.TestReport + Entities.Constants.DoubleBackslash, TestCase.TestReportFileNamePrefix + Entities.Constants.Underscore + time);
                Directory.CreateDirectory(PathString);
                FilePath = new StringBuilder().Append(PathString).Append(Entities.Constants.DoubleBackslash).Append(TestCase.TestReportFileNamePrefix).Append(Entities.Constants.Space).Append(time).Append(Entities.Constants.Xlxs).ToString();

                workbook = applicationClass.Workbooks.Add();
                workbook.Worksheets.Add();
                dynamic workSheet = (Worksheet)workbook.ActiveSheet;
                workSheet.Name = Entities.Constants.WorkSheets.TestIterationsWorkSheet;

                //// Remove all other sheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Name != Entities.Constants.WorkSheets.TestIterationsWorkSheet)
                    {
                        sheet.Delete();
                    }
                }
                //// Get style header 
                var style = workbook.Styles;
                dynamic styleHeader = GetStyle(ReportingStyles.StyleHeader, ref style);

                //// Summary header
                ((Range)workSheet.Cells[1, 1]).Value = Entities.Constants.TestIteration.StartDateTime;
                ((Range)workSheet.Cells[1, 1]).Style = styleHeader;
                ((Range)workSheet.Cells[1, 2]).Value = Timing.TotalStartTime.ToString(Entities.Constants.LongDateTimeFormat);
                ((Range)workSheet.Cells[2, 1]).Value = Entities.Constants.TestIteration.EndDateTime;
                ((Range)workSheet.Cells[2, 1]).Style = styleHeader;
                ((Range)workSheet.Cells[2, 2]).Value = Timing.TotalEndTime.ToString(Entities.Constants.LongDateTimeFormat);
                ((Range)workSheet.Cells[3, 1]).Value = Entities.Constants.TestIteration.Duration;
                ((Range)workSheet.Cells[3, 1]).Style = styleHeader;
                ((Range)workSheet.Cells[3, 2]).Value = Timing.Totalduration.ToString();

                //// Test iterations header
                ((Range)workSheet.Cells[5, 1]).Value = Entities.Constants.TestIteration.Application;
                ((Range)workSheet.Cells[5, 1]).Style = styleHeader;
                ((Range)workSheet.Cells[5, 2]).Value = Entities.Constants.TestIteration.TestCaseName;
                ((Range)workSheet.Cells[5, 2]).Style = styleHeader;
                ((Range)workSheet.Cells[5, 3]).Value = Entities.Constants.TestIteration.TestCaseDescription;
                ((Range)workSheet.Cells[5, 3]).Style = styleHeader;
                ((Range)workSheet.Cells[5, 4]).Value = Entities.Constants.TestIteration.Result;
                ((Range)workSheet.Cells[5, 4]).Style = styleHeader;
                ((Range)workSheet.Cells[5, 5]).Value = Entities.Constants.TestIteration.Duration;
                ((Range)workSheet.Cells[5, 5]).Style = styleHeader;
                ((Range)workSheet.Cells[5, 6]).Value = Entities.Constants.TestIteration.DocumentReference;
                ((Range)workSheet.Cells[5, 6]).Style = styleHeader;

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.Reporting, MethodBase.GetCurrentMethod().Name);
                return false;
            }
            finally
            {
                var b = workbook;
                //// Save file
                if (b != null)
                {
                    workbook.SaveAs(FilePath, Type.Missing, string.Empty, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    File.SetAttributes(FilePath, FileAttributes.Hidden);
                    WorkBookUtility.CloseWorkBook(workbook, true);
                }

                WorkBookUtility.CloseExcel(applicationClass);
            }
        }

        /// <summary>
        /// This function creates a Excel sheet named as the current Test case name in Excel report file.
        /// </summary>
        /// <param name="errorMessage">Error message set when error in creating excel.</param>
        /// <returns>True if able to create excel successfully, false otherwise.</returns>
        public static bool CreateExcelSheet(ref string errorMessage)
        {
            var applicationClass = new ApplicationClass();
            Workbook workbook = null;

            try
            {
                workbook = WorkBookUtility.OpenWorkBook(applicationClass, FilePath);
                var style = workbook.Styles;

                //// Check length of TestCase.Name
                if (TestCase.Name.Length > 31)
                {
                    //// Add message to reporting file, sheet "TestIterations"
                    dynamic workSheet = workbook.Worksheets[1];
                    ((Range)workSheet.Cells[6, 2]).Value = Entities.Constants.Messages.TestCaseNameLimit;
                    ((Range)workSheet.Cells[6, 4]).Value = Entities.Constants.Fail;

                    style = workbook.Styles;
                    ((Range)workSheet.Cells[6, 4]).Style = GetStyle(ReportingStyles.StyleFailed, ref style);
                    throw new Exception(Entities.Constants.Messages.TestCaseNameLimit);
                }

                //// Check if TestCase.Name already exist as a sheet because then the same testcase may be added several times in TestIterations.xlsx which is not allowed
                foreach (Worksheet workSheet in workbook.Worksheets)
                {
                    if (workSheet.Name == TestCase.Name)
                    {
                        dynamic workSheetNext = workbook.Worksheets[1];
                        ((Range)workSheetNext.Cells[6, 2]).Value = string.Format(Entities.Constants.Messages.ReportAlreadyExist, FilePath, TestCase.Name);
                        ((Range)workSheetNext.Cells[6, 4]).Value = Entities.Constants.Fail;

                        style = workbook.Styles;
                        ((Range)workSheetNext.Cells[6, 4]).Style = GetStyle(ReportingStyles.StyleFailed, ref style);
                        throw new Exception(string.Format(Entities.Constants.Messages.ReportAlreadyExist, FilePath, TestCase.Name));
                    }
                }

                workbook.Worksheets.Add(After: workbook.ActiveSheet);
                var testCaseWorkSheet = (Worksheet)workbook.ActiveSheet;
                testCaseWorkSheet.Name = TestCase.Name;

                dynamic styleHeader = GetStyle(ReportingStyles.StyleHeader, ref style);
                ((Range)testCaseWorkSheet.Cells[1, 1]).Value = Entities.Constants.TestIteration.Application;
                ((Range)testCaseWorkSheet.Cells[1, 1]).Style = styleHeader;
                ((Range)testCaseWorkSheet.Cells[1, 2]).Value = TestCase.Application;
                ((Range)testCaseWorkSheet.Cells[2, 1]).Value = Entities.Constants.TestIteration.TestCaseName;
                ((Range)testCaseWorkSheet.Cells[2, 1]).Style = styleHeader;
                ((Range)testCaseWorkSheet.Cells[2, 2]).Value = TestCase.Name;
                ((Range)testCaseWorkSheet.Cells[3, 1]).Value = Entities.Constants.TestIteration.Description;
                ((Range)testCaseWorkSheet.Cells[3, 1]).Style = styleHeader;
                ((Range)testCaseWorkSheet.Cells[3, 2]).Value = TestCase.Description;

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.Reporting, MethodBase.GetCurrentMethod().Name);

                return false;
            }
            finally
            {
                var b = workbook;

                if (b != null)
                {
                    workbook.Save();
                    WorkBookUtility.CloseWorkBook(workbook);
                }

                WorkBookUtility.CloseExcel(applicationClass);
            }
        }

        /// <summary>
        /// This procedure will insert summary details and format the sheet.
        /// </summary>
        public static void InsertSummaryDetailsAndFormat()
        {
            var applicationClass = new ApplicationClass();
            Workbook workbook = null;

            try
            {
                workbook = WorkBookUtility.OpenWorkBook(applicationClass, FilePath);
                dynamic tcworksheet = (Worksheet)workbook.Sheets["TestIterations"];

                //// Set as active sheet
                tcworksheet.Activate();

                //// Set summary details
                ((Range)tcworksheet.Cells[1, 2]).Value = Timing.TotalStartTime.ToString("MM/dd/yyyy hh:mm:ss tt");
                ((Range)tcworksheet.Cells[2, 2]).Value = Timing.TotalEndTime.ToString("MM/dd/yyyy hh:mm:ss tt");
                ((Range)tcworksheet.Cells[3, 2]).Value = Timing.Totalduration.ToString();

                //// Set style for results, column 4 from row 6
                var row = 6;
                dynamic result = (Range)tcworksheet.Cells[row, 4];
                while (!string.IsNullOrEmpty(result.Value))
                {
                    //// Set style for result
                    if (result.Value == Entities.Constants.Pass)
                    {
                        var tmp = workbook.Styles;
                        result.Style = GetStyle(ReportingStyles.StylePassed, ref tmp);
                        ((Range)tcworksheet.Cells[row, 1]).EntireRow.AutoFit();
                    }
                    else if (result.Value == Entities.Constants.Fail)
                    {
                        var tmp = workbook.Styles;
                        result.Style = GetStyle(ReportingStyles.StyleFailed, ref tmp);
                    }

                    ((Range)tcworksheet.Cells[row, 1]).EntireRow.WrapText = false;
                    ((Range)tcworksheet.Cells[row, 2]).EntireRow.WrapText = false;
                    ((Range)tcworksheet.Cells[row, 3]).EntireRow.WrapText = false;
                    ((Range)tcworksheet.Cells[row, 4]).EntireRow.WrapText = false;
                    ((Range)tcworksheet.Cells[row, 5]).EntireRow.WrapText = false;
                    ((Range)tcworksheet.Cells[row, 6]).EntireRow.WrapText = false;
                    row += 1;
                    result = (Range)tcworksheet.Cells[row, 4];
                }
                //// Autofit all columns
                ((Range)tcworksheet.Cells[1, 1]).EntireColumn.AutoFit();
                ((Range)tcworksheet.Cells[1, 2]).EntireColumn.AutoFit();
                ((Range)tcworksheet.Cells[1, 3]).EntireColumn.AutoFit();
                ((Range)tcworksheet.Cells[1, 4]).EntireColumn.AutoFit();
                ((Range)tcworksheet.Cells[1, 5]).EntireColumn.AutoFit();
                ((Range)tcworksheet.Cells[1, 6]).EntireColumn.TextToColumns();
                ((Range)tcworksheet.Cells[1, 6]).EntireColumn.AutoFit();
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.Reporting, MethodBase.GetCurrentMethod().Name);
                throw;
            }
            finally
            {
                var b = workbook;
                //// Close objects
                if (b != null)
                {
                    workbook.Save();
                    WorkBookUtility.CloseWorkBook(workbook, true);
                }

                WorkBookUtility.CloseExcel(applicationClass);
            }
        }

        /// <summary>
        /// This function creates a Excel sheet named as the current Test case name in Excel report file.
        /// </summary>
        /// <param name="errorMessage">Error message set when error in creating excel.</param>
        /// <returns>True if able to create excel successfully, false otherwise.</returns>
        public bool InsertTestStepResult(ref string errorMessage)
        {
            var applicationClass = new ApplicationClass();
            Workbook workbook = null;
            try
            {
                workbook = WorkBookUtility.OpenWorkBook(applicationClass, FilePath);

                if (workbook.Worksheets == null)
                {
                    errorMessage = string.Format(Entities.Constants.Messages.WorkSheetNotExist, TestCase.Name);
                    return false;
                }

                var workSheet = (Worksheet)workbook.Worksheets[TestCase.Name];
                workSheet.Name = TestCase.Name;

                var row = 4;
                var testDataIterationNr = 0;
                var styles = workbook.Styles;

                foreach (var resultObj in Result.TestStepsResultsCollection)
                {
                    if (testDataIterationNr < Convert.ToInt32(resultObj.TestDataIterationNr))
                    {
                        row += 1;
                        //// add row to get a blank row
                        testDataIterationNr = Convert.ToInt32(resultObj.TestDataIterationNr);

                        //// Get style header 
                        dynamic styleHeader = GetStyle(ReportingStyles.StyleHeader, ref styles);

                        //// Test data iteration header
                        ((Range)workSheet.Cells[row, 1]).Value = Entities.Constants.TestResult.Iteration;
                        ((Range)workSheet.Cells[row, 1]).Style = styleHeader;
                        ((Range)workSheet.Cells[row, 2]).Value = resultObj.TestDataIterationNr;

                        //// Test steps header
                        row += 1;
                        ((Range)workSheet.Cells[row, 1]).Value = Entities.Constants.TestResult.StepNumber;
                        ((Range)workSheet.Cells[row, 1]).Style = styleHeader;
                        ((Range)workSheet.Cells[row, 2]).Value = Entities.Constants.TestResult.Description;
                        ((Range)workSheet.Cells[row, 2]).Style = styleHeader;
                        ((Range)workSheet.Cells[row, 3]).Value = Entities.Constants.TestResult.Result;
                        ((Range)workSheet.Cells[row, 3]).Style = styleHeader;
                        ((Range)workSheet.Cells[row, 4]).Value = Entities.Constants.TestResult.Comment;
                        ((Range)workSheet.Cells[row, 4]).Style = styleHeader;
                        ((Range)workSheet.Cells[row, 5]).Value = Entities.Constants.TestResult.Remarks;
                        ((Range)workSheet.Cells[row, 5]).Style = styleHeader;

                        row ++;
                    }

                    ((Range)workSheet.Cells[row, 1]).Value = resultObj.StepNr;
                    ((Range)workSheet.Cells[row, 2]).Value = resultObj.Description;
                    ((Range)workSheet.Cells[row, 3]).Value = resultObj.Result;

                    styles = workbook.Styles;
                    ((Range)workSheet.Cells[row, 3]).Style = resultObj.Result == Entities.Constants.Pass
                        ? GetStyle(ReportingStyles.StylePassed, ref styles)
                        : GetStyle(ReportingStyles.StyleFailed, ref styles);
                    ((Range)workSheet.Cells[row, 4]).Value = resultObj.Comment.Length > 255
                        ? resultObj.Comment.Substring(1, 255)
                        : resultObj.Comment;
                    ((Range)workSheet.Cells[row, 5]).Value = resultObj.Remarks;
                    row += 1;
                }

                //// Autofit
                ((Range)workSheet.Cells[1, 1]).EntireColumn.AutoFit();
                ((Range)workSheet.Cells[1, 2]).EntireColumn.AutoFit();
                ((Range)workSheet.Cells[1, 3]).EntireColumn.AutoFit();
                ((Range)workSheet.Cells[1, 4]).EntireColumn.AutoFit();
                ((Range)workSheet.Cells[1, 5]).EntireColumn.AutoFit();
                ((Range)workSheet.Cells[1, 6]).EntireColumn.AutoFit();
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.Reporting, MethodBase.GetCurrentMethod().Name);
                return false;
            }
            finally
            {
                var b = workbook;
                if (b != null)
                {
                    workbook.Save();
                }

                WorkBookUtility.CloseWorkBook(workbook, true);
                WorkBookUtility.CloseExcel(applicationClass);
            }

            return true;
        }

        /// <summary>
        /// This procedure will set style of report.
        /// </summary>
        /// <param name="reportingStyle">Reporting style.</param>
        /// <param name="styles">Available style.</param>
        /// <returns>True or False.</returns>
        private static Style GetStyle(ReportingStyles reportingStyle, ref Styles styles)
        {
            try
            {
                Style objStyle;
                switch (reportingStyle)
                {
                    case ReportingStyles.StyleHeader:
                        var styleName = Entities.Constants.StyleNames.StyleHeader;
                        foreach (var style in styles)
                        {
                            dynamic existingStyle = style;
                            if (existingStyle.Name == styleName)
                            {
                                return existingStyle;
                            }
                        }

                        objStyle = styles.Add(Entities.Constants.StyleNames.StyleHeader);
                        objStyle.Font.Bold = true;
                        objStyle.Interior.Color = Color.Gray;
                        break;
                    case ReportingStyles.StylePassed:
                        styleName = Entities.Constants.StyleNames.StyleResultPassed;
                        foreach (var style in styles)
                        {
                            dynamic existingStyle = style;
                            if (existingStyle.Name == styleName)
                            {
                                return existingStyle;
                            }
                        }

                        objStyle = styles.Add(Entities.Constants.StyleNames.StyleResultPassed);
                        objStyle.Interior.Color = Color.Green;
                        break;
                    case ReportingStyles.StyleFailed:
                        styleName = Entities.Constants.StyleNames.StyleResultFailed;
                        foreach (var style in styles)
                        {
                            dynamic existingStyle = style;
                            if (existingStyle.Name == styleName)
                            {
                                return existingStyle;
                            }
                        }

                        objStyle = styles.Add(Entities.Constants.StyleNames.StyleResultFailed);
                        objStyle.Interior.Color = Color.Red;
                        break;
                    default:
                        objStyle = styles.Add(Entities.Constants.StyleNames.StyleResultFailed);
                        objStyle.Interior.Color = Color.Empty;
                        break;
                }

                return objStyle;
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.Reporting, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }
    }
}