// <copyright file="Verifications.cs" company="Infosys Ltd.">
//  Copyright (c) Infosys Ltd. All rights reserved.
// </copyright>
// <summary>Verifications.cs class handles test verifications.</summary>
namespace INF.Selenium.TestAutomation.Verifications
{
    using System;
    using System.Reflection;
    using Configuration;
    using Entities;
    using Microsoft.VisualStudio.TestTools.UITesting;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using OpenQA.Selenium;
    using UI;
    using Utilities;

    /// <summary>
    /// Test verifications.
    /// </summary>
    public class Verifications
    {
        /// <summary>
        /// Verifies that a specific Browser exist.
        /// </summary>
        /// <param name="ts">Test step.</param>
        public static void BrowserExist(TestStep ts)
        {
            try
            {
                var title = WebdriverBrowser.GetTitleFromPartOfTitle(ts.TestData.ContainsValue(ts.TestDataKeyToUse).ToString(), false);
                BrowserWindow.Locate(title);
                Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, Constants.Verification.VerifyBrowserExist, Constants.Pass, string.Concat(Constants.Verification.BrowserWithTitle, ts.TestData.ContainsValue(ts.TestDataKeyToUse), Constants.Verification.DoesExist), ts.Remarks);
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, Constants.Verification.VerifyBrowserNotExist, Constants.Fail, string.Format(Constants.Messages.DueToException, ex.Message), ts.Remarks);
                LogHelper.ErrorLog(ex, Constants.ClassName.Verifications, MethodBase.GetCurrentMethod().Name);
                ////return false;
            }
        }

        /// <summary>
        /// Verifies that a specific Browser doesn't exist.
        /// </summary>
        /// <param name="ts">Test step.</param>
        public static void BrowserNotExist(TestStep ts)
        {
            try
            {
                var title = WebdriverBrowser.GetTitleFromPartOfTitle(ts.TestData.ContainsValue(ts.TestDataKeyToUse).ToString(), false);
                BrowserWindow.Locate(title);
                Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, Constants.Verification.VerifyBrowserExist, Constants.Fail, string.Concat(Constants.Verification.BrowserWithTitle, ts.TestData.ContainsValue(ts.TestDataKeyToUse), Constants.Verification.DoesExist), ts.Remarks);
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, Constants.Verification.VerifyBrowserNotExist, Constants.Pass, string.Concat(Constants.Verification.BrowserWithTitle, ts.TestData.ContainsValue(ts.TestDataKeyToUse), Constants.Verification.DoesExist), ts.Remarks);
                LogHelper.ErrorLog(ex, Constants.ClassName.Verifications, MethodBase.GetCurrentMethod().Name);
            }
        }

        /// <summary>
        /// Verifies that a specific object Is Enabled.
        /// </summary>
        /// <param name="ts">Test step.</param>
        public static void IsEnabled(TestStep ts)
        {
            try
            {
                UI.UiActions.WaitForControlToExist(ts); //// Throws WebDriverTimeoutException
                var verifyobj = WebdriverBrowser.Driver.FindElement(By.XPath(ts.UiControl.UiControlSearchValue)).Enabled;
                Assert.IsTrue(verifyobj);
                Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, "Verify IsEnabled", Constants.Pass, Constants.Messages.CompleteSuccess, ts.Remarks);
            }
            catch (WebDriverTimeoutException ex)
            {
                Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, "Verify IsEnabled", Constants.Fail, string.Format(Constants.Messages.WebDriverTimeoutException, ex.Message), ts.Remarks);
                LogHelper.ErrorLog(ex, Constants.ClassName.Verifications, MethodBase.GetCurrentMethod().Name);
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, "Verify IsEnabled", Constants.Fail, string.Format(Constants.Messages.DueToException, ex.Message), ts.Remarks);
                LogHelper.ErrorLog(ex, Constants.ClassName.Verifications, MethodBase.GetCurrentMethod().Name);
            } 
        }

        /// <summary>
        /// Verifies that a specific object Is Disabled.
        /// </summary>
        /// <param name="ts">Test step.</param>
        public static void IsDisabled(TestStep ts)
        {
            try
            {
                UI.UiActions.WaitForControlToExist(ts); //// Throws WebDriverTimeoutException
                var verifyobj = WebdriverBrowser.Driver.FindElement(By.XPath(ts.UiControl.UiControlSearchValue)).GetAttribute("disabled");
                Assert.AreEqual("true", verifyobj);
                Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, "Verify IsDisabled", Constants.Pass, Constants.Messages.CompleteSuccess, ts.Remarks);
            }
            catch (WebDriverTimeoutException ex)
            {
                Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, "Verify IsDisabled", Constants.Fail, string.Format(Constants.Messages.WebDriverTimeoutException, ex.Message), ts.Remarks);
                LogHelper.ErrorLog(ex, Constants.ClassName.Verifications, MethodBase.GetCurrentMethod().Name);
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, "Verify IsDisabled", Constants.Fail, string.Format(Constants.Messages.DueToException, ex.Message), ts.Remarks);
                LogHelper.ErrorLog(ex, Constants.ClassName.Verifications, MethodBase.GetCurrentMethod().Name);
            }  
        }

        /// <summary>
        /// Verifies that a specific object Is Displayed.
        /// </summary>
        /// <param name="ts">Test step.</param>
        public static void IsDisplayed(TestStep ts)
        {
            try
            {
                UI.UiActions.WaitForControlToExist(ts); //// Throws WebDriverTimeoutException
                var verifyobj = WebdriverBrowser.Driver.FindElement(By.XPath(ts.UiControl.UiControlSearchValue)).Displayed;
                Assert.IsTrue(verifyobj);
                Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, "Verify IsDisplayed", Constants.Pass, Constants.Messages.CompleteSuccess, ts.Remarks);
            }
            catch (WebDriverTimeoutException ex)
            {
                Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, "Verify IsDisplayed", Constants.Fail, string.Format(Constants.Messages.WebDriverTimeoutException, ex.Message), ts.Remarks);
                LogHelper.ErrorLog(ex, Constants.ClassName.Verifications, MethodBase.GetCurrentMethod().Name);
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, "Verify IsDisplayed", Constants.Fail, string.Format(Constants.Messages.DueToException, ex.Message), ts.Remarks);
                LogHelper.ErrorLog(ex, Constants.ClassName.Verifications, MethodBase.GetCurrentMethod().Name);
            } 
        }

        /// <summary>
        /// Verifies that a specific object Is Selected.
        /// </summary>
        /// <param name="ts">Test step.</param>
        public static void IsSelected(TestStep ts)
        {
            try
            {
                UI.UiActions.WaitForControlToExist(ts); //// Throws WebDriverTimeoutException
                var verifyobj = WebdriverBrowser.Driver.FindElement(By.XPath(ts.UiControl.UiControlSearchValue)).Selected;
                Assert.IsTrue(verifyobj);
                Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, "Verify IsSelected", Constants.Pass, Constants.Messages.CompleteSuccess, ts.Remarks);
            }
            catch (WebDriverTimeoutException ex)
            {
                Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, "Verify IsSelected", Constants.Fail, string.Format(Constants.Messages.WebDriverTimeoutException, ex.Message), ts.Remarks);
                LogHelper.ErrorLog(ex, Constants.ClassName.Verifications, MethodBase.GetCurrentMethod().Name);
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, "Verify IsSelected", Constants.Fail, string.Format(Constants.Messages.DueToException, ex.Message), ts.Remarks);
                LogHelper.ErrorLog(ex, Constants.ClassName.Verifications, MethodBase.GetCurrentMethod().Name);
            }
        }

        /// <summary>
        /// Verifies the data base value.
        /// </summary>
        /// <param name="ts">Test step.</param>
        public void DatabaseValue(TestStep ts)
        {
            try
            {
                //// Check that mandatory values exist
                if (string.IsNullOrEmpty(Convert.ToString(ts.TestData[1])))
                {
                    throw new Exception(Constants.Verification.DbQryVerificationIdValue);
                }

                if (string.IsNullOrEmpty(ts.Verification.OperatorToUse))
                {
                    throw new Exception(Constants.Verification.OpVerificationIdValue);
                }

                if (string.IsNullOrEmpty(ts.TestData.ContainsValue(ts.TestDataKeyToUse).ToString()))
                {
                    throw new Exception(Constants.Verification.TestDataValue);
                }

                Db objDb = new Db();
                //// Validate query
                objDb.ValidateQuery(Convert.ToString(ts.TestData[1]));

                // Run db query         
                dynamic dataBaseValue = Db.ExecuteQuery(Convert.ToString(ts.TestData[1]), ts.Verification.OperatorToUse);
                String value = dataBaseValue;
                value = value.Trim();
                //// Verify that database value returned from query is correct 
                switch (ts.Verification.OperatorToUse.ToUpper())
                {  
                    case Constants.DbActions.IsEquals:
                        Assert.IsTrue(value.Equals(ts.Remarks.Trim(), StringComparison.OrdinalIgnoreCase), "Database value wasn't correct. Expected: " + ts.Remarks + ", Actual: " + value);
                        Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, "Verify database value", Constants.Pass, "Database value was correct .Expected: " + ts.Remarks + ", Actual: " + value, ts.Remarks);
                        break;
                    case Constants.DbActions.Contains:
                        Assert.IsTrue(value.Contains(ts.Remarks.Trim()), "Database value wasn't correct.");
                        Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, "Verify database value", Constants.Pass, "Database value was correct.", ts.Remarks);
                        break;
                    case Constants.DbActions.NotEquals:
                        Assert.IsTrue(!value.Equals(ts.Remarks.Trim(), StringComparison.OrdinalIgnoreCase), "Database value wasn't correct.");
                        Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, "Verify database value", Constants.Pass, "Database value was correct.", ts.Remarks);
                        break;
                    case Constants.DbActions.NotContains:
                        Assert.IsTrue(!value.Contains(ts.Remarks.Trim()), "Database value wasn't correct.");
                        Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, "Verify database value", Constants.Pass, "Database value was correct.", ts.Remarks);
                        break;
                    case Constants.DbActions.Insert:
                        Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, "Record Inserted to the Database", Constants.Pass, Constants.Messages.CompleteSuccess, ts.Remarks);
                        break;
                    case Constants.DbActions.Delete:
                        Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, "Record Deleted to the Database", Constants.Pass, Constants.Messages.CompleteSuccess, ts.Remarks);
                        break;
                    case Constants.DbActions.Update:
                        Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, "Record Updated to the Database", Constants.Pass, Constants.Messages.CompleteSuccess, ts.Remarks);
                        break;
                    case Constants.DbActions.Call:
                        Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, "Calling of DB procedure success", Constants.Pass, Constants.Messages.CompleteSuccess, ts.Remarks);
                        break;
                    case Constants.DbActions.Commit:
                        Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, "Commit Successful", Constants.Pass, Constants.Messages.CompleteSuccess, ts.Remarks);
                        break;
                    default:
                        throw new Exception("Operator " + ts.Verification.OperatorToUse + " is not supported");
                }
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, "Verify database value", Constants.Fail, string.Format(Constants.Messages.DueToException, ex.Message), ts.Remarks);
                LogHelper.ErrorLog(ex, Constants.ClassName.Verifications, MethodBase.GetCurrentMethod().Name);
            }
        }
    }
}
