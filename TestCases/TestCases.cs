// <copyright file="TestCases.cs" company="Infosys Ltd.">
//  Copyright (c) Infosys Ltd. All rights reserved.
// </copyright>
// <summary>TestCases.cs class handles all test cases and its steps as per test iteration excel.</summary>
namespace INF.Selenium.TestAutomation.TestCases
{
    using System;
    using System.Drawing;
    using System.Drawing.Imaging;
    using System.IO;
    using System.Reflection;
    using System.Text;
    using System.Windows.Forms;
    using Configuration;
    using Entities;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using UI;
    using Utilities;

    /// <summary>
    /// Reads Test cases and executes.
    /// </summary>
    public class TestCases
    {
        /// <summary>
        /// Gets or private sets Hyper link for the failed test steps.
        /// </summary>
        /// <value>File path.</value>
        public static string Hyplink { get; private set; }

        /// <summary>
        /// Gets or private sets Test Session ID.
        /// </summary>
        /// <value>ID of test session.</value>
        public static Guid TestSessionId { get; private set; }

        /// <summary>
        /// Executes test steps.
        /// </summary>
        public void Execute()
        {
            try
            {
                TestSessionId = Guid.NewGuid();

                //// Initiliaze test data
                TestCase.TestDataCount = Constants.Zero;

                var errorMessage = string.Empty;
                Data objData = new Data();
                if (!objData.InitiliazeTestCaseAndTestData(ref errorMessage))
                {
                    Assert.Inconclusive(Constants.Messages.TestInitializationError, errorMessage);
                }

                //// Run the test case once for every test data that exist
                var testCaseCount = TestCase.TestDataCount;
                for (var testCaseIndex = 1; testCaseIndex <= testCaseCount; testCaseIndex++)
                {
                    var isSuccess = true;

                    var testStepCount = TestCase.TestStepList.Count;
                    for (var testStepIndex = 0; testStepIndex <= testStepCount-1; testStepIndex++)
                    {
                        var teststep = new TestStep();
                        var testStep = teststep;
                        testStep.TestDataKeyToUse = Convert.ToString(testCaseIndex);
                        testStep.Action = TestCase.TestStepList[testStepIndex].Action;
                        testStep.TestData = TestCase.TestStepList[testStepIndex].TestData;
                        testStep.TestStepNumber = TestCase.TestStepList[testStepIndex].TestStepNumber;
                        testStep.UiControl = TestCase.TestStepList[testStepIndex].UiControl;
                        testStep.Verification = TestCase.TestStepList[testStepIndex].Verification;
                        testStep.Remarks = TestCase.TestStepList[testStepIndex].Remarks;
                        WebdriverBrowser objWebdriverBrowser = new WebdriverBrowser();
                        switch (testStep.Action.ToUpper())
                        {
                            case Constants.TestStepAction.CloseWebDriverBrowsers:
                                isSuccess = objWebdriverBrowser.CloseAllWebdriver_Browsers();
                                break;
                            case Constants.TestStepAction.LaunchWebDriverBrowser:
                                isSuccess = UiActions.LaunchWebDriverBrowser(testStep);
                                break;
                            case Constants.TestStepAction.WebDriverEditUIControl:
                                isSuccess = UiActions.WebDriverEditUIControl(testStep);
                                break;
                            case Constants.TestStepAction.WebDriverAlertHandler:
                                isSuccess = UiActions.WebDriverAlertHandler(testStep);
                                break;
                            case Constants.TestStepAction.WebDriverFrameHandler:
                                isSuccess = UiActions.WebDriverFrameHandler(testStep);
                                break;
                            case Constants.TestStepAction.WebDriverSwitchToDefaultFrame:
                                isSuccess = UiActions.WebDriverSwitchToDefaultFrame(testStep);
                                break;
                            case Constants.TestStepAction.WebDriverVerify:
                                isSuccess = UiActions.WebDriverVerify(testStep);
                                break;
                            case Constants.TestStepAction.WebDriverSaveUIControl:
                                isSuccess = UiActions.WebDriverSaveUIControl(testStep);
                                break;
                            case Constants.TestStepAction.WebPaginationIteration:
                                isSuccess = UiActions.WebPaginationIteration(testStep);
                                break;
                            case Constants.TestStepAction.WaitforUI:
                                isSuccess = UiActions.WaitForUI(testStep);
                                break;
                            case Constants.TestStepAction.Sendkeys:
                                isSuccess = UiActions.SendKeys(testStep);
                                break;
                            default:
                                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, Constants.TestIterations, Constants.Fail, string.Format(Constants.Messages.NotSupported, testStep.Action), testStep.Remarks);
                                isSuccess = false;
                                break;
                        }

                        if (!isSuccess)
                        {
                            break;
                        }

                        //// Exit if one test step failed
                    }

                    if (!isSuccess && Result.GetTestScriptResult() == Constants.Fail)
                    {
                        //// Some test step failed, raise Assert.Fail after all test data iterations are completed
                        var bounds = Screen.GetBounds(Point.Empty);

                        using (var bitmap = new Bitmap(bounds.Width, bounds.Height))
                        {
                            using (var objGraphic = Graphics.FromImage(bitmap))
                            {
                                objGraphic.CopyFromScreen(Point.Empty, Point.Empty, bounds.Size);
                            }

                            bitmap.Save(
                                new StringBuilder()
                                    .Append(Reporting.PathString)
                                    .Append(Constants.DoubleBackslash)
                                    .Append(Path.GetFileNameWithoutExtension(TestCase.Name))
                                    .Append(Constants.Jpg).ToString(), 
                                ImageFormat.Jpeg);

                            Hyplink = new StringBuilder()
                                .Append(Constants.Hyperlink)
                                .Append(Reporting.PathString)
                                .Append(Constants.DoubleBackslash)
                                .Append(Path.GetFileNameWithoutExtension(TestCase.Name))
                                .Append(Constants.Jpg)
                                .Append(@""",""")
                                .Append(Path.GetFileNameWithoutExtension(TestCase.Name))
                                .Append(@""")").ToString();
                        }

                        Assert.Fail(Constants.Messages.TestCaseFailedError, Reporting.FilePath);
                    }
                    else
                    {
                        Hyplink = null;
                    }
                }

                LogHelper objLogHelper = new LogHelper();
                    objLogHelper.EventLog(
                    string.Format(
                    Constants.Messages.SuccessfullCompletion, 
                    MethodBase.GetCurrentMethod().Name), 
                    Constants.ClassName.TestCases, 
                    MethodBase.GetCurrentMethod().Name);
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.TestCases, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }
    }
}