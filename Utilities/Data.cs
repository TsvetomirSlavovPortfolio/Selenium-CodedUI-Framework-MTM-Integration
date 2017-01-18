// <copyright file="Data.cs" company="Infosys Ltd.">
//  Copyright (c) Infosys Ltd. All rights reserved.
// </copyright>
// <summary>Data.cs Stores value from all Excel artifacts</summary>
namespace INF.Selenium.TestAutomation.Utilities
{
    using WebServiceData;
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Reflection;
    using System.Text;
    using Configuration;
    using Entities;
    using Microsoft.Office.Interop.Excel;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// All excel sheets data will be handled here.
    /// </summary>
    public class Data
    {

        /// <summary>
        /// Get Value of column referring to.
        /// </summary>
        /// <param name="context">Context of row.</param>
        /// <param name="columnName">Column Name.</param>
        /// <returns>Data row for the column referring to.</returns>
        public static string GetValue(TestContext context, string columnName)
        {
            return context.DataRow[columnName].ToString().Trim();
        }

        /// <summary>
        /// This function creates initialize test case and test data.
        /// </summary>
        /// <param name="applicationClass">Application class.</param>
        /// <value>Application class value.</value>
        public static void LoadUiControls(ApplicationClass applicationClass)
        {
            try
            {
                var workbook = WorkBookUtility.OpenWorkBook(applicationClass, TestCase.RootFilePath + ConfigurationManager.AppSettings.Get(Entities.Constants.AppSetting.UiControlFile));
                dynamic worksheet = workbook.Worksheets[1];
                var rowsCount = worksheet.UsedRange.Rows.Count;
                var cellCount = worksheet.UsedRange.Columns.Count + 1;

                for (var rowIndex = 2; rowIndex <= rowsCount; rowIndex++)
                {
                    if (string.IsNullOrEmpty(worksheet.Cells[rowIndex, 1].value))
                    {
                        break; // reading the sheet untill the first empty row
                    }

                    var objUiControl = new UiControl();
                    for (var cellIndex = 1; cellIndex <= cellCount; cellIndex++)
                    {
                        string headerValue = Convert.ToString(worksheet.Cells[1, cellIndex].value);
                        string dataValue = Convert.ToString(worksheet.Cells[rowIndex, cellIndex].value);

                        if (string.IsNullOrEmpty(headerValue))
                        {
                            TestCase.UiControls.Add(objUiControl);
                            break;
                        }

                        switch (headerValue.ToUpper())
                        {
                            case Entities.Constants.UiControls.UiControlId:
                                objUiControl.UiControlId = dataValue;
                                break;
                            case Entities.Constants.UiControls.UiControlType:
                                objUiControl.UiControlType = dataValue;
                                break;
                            case Entities.Constants.UiControls.UiControlSearchProperty:
                                objUiControl.UiControlSearchProperty = dataValue;
                                break;
                            case Entities.Constants.UiControls.UiControlSearchValue:
                                objUiControl.UiControlSearchValue = dataValue;
                                break;
                            default:
                                if (headerValue.Length > 0)
                                {
                                    throw new Exception(string.Format(Entities.Constants.Messages.UiControlSheetError, headerValue));
                                }

                                TestCase.UiControls.Add(objUiControl);
                                break;
                        }
                    }
                }
                WorkBookUtility.CloseWorkBook(workbook);
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.Data, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }

        /// <summary>
        /// Load all Verifications.
        /// </summary>
        /// <param name="applicationClass">Name of the Application under test.</param>
        /// <value>Application class value.</value>
        public static void LoadVerifications(ApplicationClass applicationClass)
        {
            try
            {
                var workbook = WorkBookUtility.OpenWorkBook(applicationClass, TestCase.RootFilePath + ConfigurationManager.AppSettings.Get(Entities.Constants.AppSetting.VerificationFile));
                dynamic worksheet = (Worksheet)workbook.Worksheets[1];
                var rowsCount = worksheet.UsedRange.Rows.Count;
                var cellCount = worksheet.UsedRange.Columns.Count + 1;
                for (var rowIndex = 2; rowIndex <= rowsCount; rowIndex++)
                {
                    if (string.IsNullOrEmpty(worksheet.Cells[rowIndex, 1].value))
                    {
                        break; //// reading the sheet untill the first empty row
                    }

                    var verification = new Verification();
                    for (var cellIndex = 1; cellIndex <= cellCount; cellIndex++)
                    {
                        string headerValue = Convert.ToString(worksheet.Cells[1, cellIndex].value);
                        string dataValue = Convert.ToString(worksheet.Cells[rowIndex, cellIndex].value);

                        if (string.IsNullOrEmpty(headerValue))
                        {
                            TestCase.Verifications.Add(verification);
                            break;
                        }

                        switch (headerValue.ToUpper())
                        {
                            case Entities.Constants.Verification.VerificationId:
                                verification.VerificationId = dataValue;
                                break;
                            case Entities.Constants.Verification.VerificationType:
                                verification.VerificationType = dataValue;
                                break;
                            case Entities.Constants.Verification.OperatorVerification:
                                verification.OperatorToUse = dataValue;
                                break;
                            default:

                                //// No more cell to get data from, add TestData to TestCase TestStepList
                                if (headerValue.Length > 0)
                                {
                                    //// headerValue should be empty if we get here, something is wrong -> end all loops
                                    throw new Exception(string.Format(Entities.Constants.Messages.VerificationError, headerValue));
                                }

                                TestCase.Verifications.Add(verification);
                                break;
                        }
                    }
                }

                WorkBookUtility.CloseWorkBook(workbook);
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.Data, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }

        /// <summary>
        /// Load Test Class.
        /// </summary>
        /// <param name="applicationClass">Name of the Application under test.</param>
        public static void LoadTestCases(ApplicationClass applicationClass)
        {
            var testdatakeyConfig = 1;
            try
            {
                var workbook = WorkBookUtility.OpenWorkBook(applicationClass, TestCase.RootFilePath + new StringBuilder().Append(ConfigurationManager.AppSettings.Get(Entities.Constants.AppSetting.TestCaseFolderName)).Append(Entities.Constants.DoubleBackslash).Append(TestCase.FileName));
                dynamic worksheet = (Worksheet)workbook.Worksheets[1];
                var rowsCount = worksheet.UsedRange.Rows.Count;
                var cellCount = worksheet.UsedRange.Columns.Count + 1;

                for (var rowindex = 2; rowindex <= rowsCount; rowindex++)
                {
                    if (string.IsNullOrEmpty(Convert.ToString(worksheet.Cells[rowindex, 1].value)))
                    {
                        break; //// reading the sheet untill the first empty row
                    }

                    var testStep = new TestStep();
                    for (var cellindex = 1; cellindex <= cellCount; cellindex++)
                    {
                        string headerValue = worksheet.Cells[1, cellindex].value;
                        string dataValue = Convert.ToString(worksheet.Cells[rowindex, cellindex].value);

                        if (string.IsNullOrEmpty(headerValue))
                        {
                            TestCase.TestStepList.Add(testStep);
                            break;
                        }

                        switch (headerValue)
                        {
                            case Entities.Constants.TestStep.TestStepNumber:
                                testStep.TestStepNumber = dataValue;
                                testStep.TestData = new Dictionary<int, string>();
                                testdatakeyConfig = 1;
                                break;
                            case Entities.Constants.TestStep.Action:
                                testStep.Action = dataValue;
                                break;
                            case Entities.Constants.TestStep.UiControlId:
                                testStep.UiControl = TestCase.UiControls.Find(f => f.UiControlId == dataValue);
                                break;
                            case Entities.Constants.TestStep.VerificationId:
                                if (dataValue != null)
                                {
                                    testStep.Verification = TestCase.Verifications.Find(f => f.VerificationId == dataValue);
                                }
                                break;
                            case Entities.Constants.TestStep.TestData:
                                if (dataValue != null && TestCase.TestDataSavedValues.ContainsKey(dataValue))
                                {
                                    var value = TestCase.TestDataSavedValues[dataValue];
                                    testStep.TestData.Add(testdatakeyConfig, value);
                                }
                                else
                                {
                                    testStep.TestData.Add(testdatakeyConfig, dataValue);
                                }

                                if (testdatakeyConfig > TestCase.TestDataCount)
                                {
                                    TestCase.TestDataCount += 1;
                                }

                                testdatakeyConfig += 1;
                                break;
                            case Entities.Constants.TestStep.Remarks:
                                testStep.Remarks = dataValue;
                                break;
                            default:
                                if (headerValue.Length > 0)
                                {
                                    throw new Exception(string.Format(Entities.Constants.Messages.TestCaseError, headerValue, TestCase.RootFilePath, TestCase.FileName));
                                }
                                TestCase.TestStepList.Add(testStep);
                                break;
                        }
                    }
                }
                WorkBookUtility.CloseWorkBook(workbook);
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.Data, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }

        /// <summary>
        /// Extracts Data from excel.
        /// </summary>
        /// <param name="errorMessage">Error message.</param>
        /// <returns>False if no error.</returns>
        public bool InitiliazeTestCaseAndTestData(ref string errorMessage)
        {
            var applicationClass = new ApplicationClass();
            try
            {
                TestConfigurations objTestConfigurations = new TestConfigurations();
                LoadAPIData loadAPIdata = new LoadAPIData();
                TestCase.UiControls = new List<UiControl>();
                TestCase.Verifications = new List<Verification>();
                ConfigStep.TestConfigNames = new List<ConfigStep>();
                TestCase.TestStepList = new List<TestStep>();

                LoadUiControls(applicationClass);

                LoadVerifications(applicationClass);

                LoadTestConfigurations(applicationClass);

                objTestConfigurations.Configuration();

                loadAPIdata.LoadAPIResponseData(applicationClass);

                LoadTestCases(applicationClass);

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.Data, MethodBase.GetCurrentMethod().Name);
                return false;
            }
            finally
            {
                WorkBookUtility.CloseExcel(applicationClass);
            }
        }

        /// <summary>
        /// Load Test Configurations.
        /// </summary>
        /// <param name="applicationClass">Name of the Application under test.</param>
        private static void LoadTestConfigurations(ApplicationClass applicationClass)
        {
            var testdatakeyConfig = 1;
            try
            {
                var workbook = WorkBookUtility.OpenWorkBook(applicationClass, TestCase.RootFilePath + ConfigurationManager.AppSettings.Get(Entities.Constants.AppSetting.TestConfigurationFile));
                dynamic worksheet = (Worksheet)workbook.Worksheets[1];
                var rowsCount = worksheet.UsedRange.Rows.Count;
                var cellCount = worksheet.UsedRange.Columns.Count + 1;

                for (var rowindex = 2; rowindex <= rowsCount; rowindex++)
                {
                    if (string.IsNullOrEmpty(worksheet.Cells[rowindex, 1].value))
                    {
                        break; //// reading the sheet untill the first empty row
                    }

                    var configStep = new ConfigStep();
                    for (var cellindex = 1; cellindex <= cellCount; cellindex++)
                    {
                        string dataValue = Convert.ToString(worksheet.Cells[rowindex, cellindex].value);
                        string headerValue = Convert.ToString(worksheet.Cells[1, cellindex].value);
                        if (string.IsNullOrEmpty(headerValue))
                        {
                            ConfigStep.TestConfigNames.Add(configStep);
                            break; //// reading the sheet untill the first empty column
                        }

                        switch (headerValue.ToUpper())
                        {
                            case Entities.Constants.TestConfiguration.SNo:
                                configStep.TestStepNo = dataValue;
                                testdatakeyConfig = 1;
                                break;
                            case Entities.Constants.TestConfiguration.Datatype:
                                configStep.TestDataType = dataValue;
                                break;
                            case Entities.Constants.TestConfiguration.VariableName:
                                configStep.TestVariableName = dataValue;
                                break;
                            case Entities.Constants.TestConfiguration.TestData:
                                if (testdatakeyConfig > ConfigStep.TestDataConfigCount)
                                {
                                    ConfigStep.TestDataConfigCount += 1;
                                }

                                testdatakeyConfig += 1;
                                configStep.TestDataValue = dataValue;
                                break;
                            default:
                                if (headerValue.Length > 0)
                                {
                                    throw new Exception(string.Format(Entities.Constants.Messages.TestConfigurationError, headerValue));
                                }

                                ConfigStep.TestConfigNames.Add(configStep);
                                break;
                        }
                    }
                }
                WorkBookUtility.CloseWorkBook(workbook);
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.Data, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }
    }
}