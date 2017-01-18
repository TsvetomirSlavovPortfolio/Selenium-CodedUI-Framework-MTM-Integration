// <copyright file="UIActions.cs" company="Infosys Ltd.">
//  Copyright (c) Infosys Ltd. All rights reserved.
// </copyright>
// <summary>UIActions.cs Reads and performs the User Interface action needs to be done while under test.</summary>
namespace INF.Selenium.TestAutomation.UI
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;
    using System.Text;
    using System.Threading;
    using Configuration;
    using Entities;
    using Microsoft.Office.Interop.Excel;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using OpenQA.Selenium;
    using OpenQA.Selenium.Support.UI;
    using Utilities;
    using WebServiceData;
    using System.Diagnostics;
    
    
    /// <summary>
    /// User interface actions.
    /// </summary>
    public class UiActions
    {
        /// <summary>
        /// Web driver edit for user interface control.
        /// </summary>
        /// <param name="testStep">Test step as parameter.</param>
        /// <returns>Returns true or false.</returns>
        public static bool WebDriverEditUIControl(TestStep testStep)
        {
            try
            {
                switch (testStep.UiControl.UiControlSearchProperty.ToUpper())
                {
                    case "ID":
                        if (!string.IsNullOrEmpty(Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)])) && Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]).ToUpper() == "CLICK")
                        {
                            WaitForControlToExist(testStep); //// Throws WebDriverTimeoutException
                            WebdriverBrowser.Driver.FindElement(By.Id(testStep.UiControl.UiControlSearchValue)).Click();
                        }
                        else
                        {
                            if (testStep.UiControl.UiControlType.ToUpper() == "HTMLCOMBOBOX")
                            {
                                WaitForControlToExist(testStep); //// Throws WebDriverTimeoutException
                                var identifier = WebdriverBrowser.Driver.FindElement(By.Id(testStep.UiControl.UiControlSearchValue));
                                var select = new SelectElement(identifier);
                                select.SelectByText(Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]));
                            }
                            ////Load Dynamic Data loaded from Web Services
                            else if (Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]).Contains(Entities.Constants.DynamicWebServiceData))
                            {
                                WaitForControlToExist(testStep); //// Throws WebDriverTimeoutException
                                WebdriverBrowser.Driver.FindElement(By.XPath(testStep.UiControl.UiControlSearchValue)).Clear();
                                WebdriverBrowser.Driver.FindElement(By.XPath(testStep.UiControl.UiControlSearchValue)).SendKeys(LoadAPIData.GetSavedAPIData(testStep));

                                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebDriverEditUIControl", Entities.Constants.Pass, "Edit search property \"" + LoadAPIData.GetSavedAPIData(testStep) + "\" completed successfully", testStep.Remarks);
                                return true;
                            }
                            else
                            {
                                WaitForControlToExist(testStep); //// Throws WebDriverTimeoutException
                                WebdriverBrowser.Driver.FindElement(By.Id(testStep.UiControl.UiControlSearchValue)).Clear();
                                WebdriverBrowser.Driver.FindElement(By.Id(testStep.UiControl.UiControlSearchValue)).SendKeys(Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]));
                            }
                        }

                        break;
                    case "XPATH":
                        if (!string.IsNullOrEmpty(Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)])) && Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]).ToUpper() == "CLICK")
                        {
                            WaitForControlToExist(testStep); //// Throws WebDriverTimeoutException
                            WebdriverBrowser.Driver.FindElement(By.XPath(testStep.UiControl.UiControlSearchValue)).Click();
                        }
                        else
                        {
                            if (testStep.UiControl.UiControlType.ToUpper() == "HTMLCOMBOBOX")
                            {
                                WaitForControlToExist(testStep); //// Throws WebDriverTimeoutException
                                WaitForElementToBeEnabled(testStep);
                                var identifier = WebdriverBrowser.Driver.FindElement(By.XPath(testStep.UiControl.UiControlSearchValue));
                                var select = new SelectElement(identifier);
                                select.SelectByText(Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]));
                            }
                            else
                            {
                                if (Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]) == null)
                                {
                                    WaitForControlToExist(testStep); //// Throws WebDriverTimeoutException
                                    WebdriverBrowser.Driver.FindElement(By.XPath(testStep.UiControl.UiControlSearchValue)).Clear();
                                }
                                ////Load Dynamic Data loaded from Web Services
                                else if (Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]).Contains(Entities.Constants.DynamicWebServiceData))
                                {
                                    WaitForControlToExist(testStep); //// Throws WebDriverTimeoutException
                                    WebdriverBrowser.Driver.FindElement(By.XPath(testStep.UiControl.UiControlSearchValue)).Clear();
                                    WebdriverBrowser.Driver.FindElement(By.XPath(testStep.UiControl.UiControlSearchValue)).SendKeys(LoadAPIData.GetSavedAPIData(testStep));

                                    Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebDriverEditUIControl", Entities.Constants.Pass, "Edit search property \"" + LoadAPIData.GetSavedAPIData(testStep) + "\" completed successfully", testStep.Remarks);
                                    return true;
                                }
                                else
                                {
                                    WaitForControlToExist(testStep); //// Throws WebDriverTimeoutException
                                    WebdriverBrowser.Driver.FindElement(By.XPath(testStep.UiControl.UiControlSearchValue)).Clear();
                                    WebdriverBrowser.Driver.FindElement(By.XPath(testStep.UiControl.UiControlSearchValue)).SendKeys(Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]));
                                }
                            }
                        }

                        break;
                    default:
                        return false;
                }
                
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebDriverEditUIControl", Entities.Constants.Pass, "Edit search property \"" + Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]) + "\" completed successfully", testStep.Remarks);
                return true;
            }
            catch (WebDriverTimeoutException ex)
            {
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebDriverEditUIControl", Entities.Constants.Fail, string.Format(Entities.Constants.Messages.WebDriverTimeoutException, ex.Message), testStep.Remarks);
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClassName, MethodBase.GetCurrentMethod().Name);
                return false;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebDriverEditUIControl", Entities.Constants.Fail, string.Format(Entities.Constants.Messages.DueToException, ex.Message), testStep.Remarks);
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClassName, MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        /// <summary>
        /// Web driver alert handler to control the alert pop ups.
        /// </summary>
        /// <param name="testStep">Test step as parameter.</param>
        /// <returns>Returns true or false.</returns>
        public static bool WebDriverAlertHandler(TestStep testStep)
        {
            try
            {
                IAlert alert;
                string action;
                string testDataRequiredValue = string.Empty;
                string valueattribute = string.Empty;
                string key = string.Empty;
                string successMessage = string.Empty;
                string testDataValue = Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]);
                if(testDataValue.Contains(Entities.Constants.PipeDelimitor))
                {
                    string[] separator = { Entities.Constants.PipeDelimitor };
                    string[] testDataValueSplit = testDataValue.Split(separator, StringSplitOptions.None);                    
                    action = testDataValueSplit[0];
                    if (action.ToUpper().Contains("TYPE"))
                    {
                        testDataRequiredValue = testDataValueSplit[1];
                    }
                    else if (action.ToUpper().Contains("GETTEXT"))
                    {
                        key = Convert.ToString(testDataValueSplit[1]);
                    } 
                }
                else
                {
                    action = Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]);
                }

                WebDriverWait wait = new WebDriverWait(WebdriverBrowser.Driver, TimeSpan.FromSeconds(Convert.ToInt32(General.WaitForControlToExistTimeOut)));
                switch (action.ToUpper())
                {
                    case "ACCEPT":
                        wait.Until(ExpectedConditions.AlertIsPresent()); //// Throws WebDriverTimeoutException
                        alert = WebdriverBrowser.Driver.SwitchTo().Alert();
                        alert.Accept();
                        successMessage = "Alert accepted.";
                        break;
                    case "DISMISS":
                        wait.Until(ExpectedConditions.AlertIsPresent()); //// Throws WebDriverTimeoutException
                        alert = WebdriverBrowser.Driver.SwitchTo().Alert();
                        alert.Dismiss();
                        successMessage = "Alert dismissed.";
                        break;
                    case "TYPE":
                        wait.Until(ExpectedConditions.AlertIsPresent()); //// Throws WebDriverTimeoutException
                        alert = WebdriverBrowser.Driver.SwitchTo().Alert();
                        alert.SendKeys(testDataRequiredValue);
                        successMessage = "Entered text \'" + testDataRequiredValue + "\' in the alert box.";
                        break;
                    case "GETTEXT":
                        wait.Until(ExpectedConditions.AlertIsPresent()); //// Throws WebDriverTimeoutException
                        alert = WebdriverBrowser.Driver.SwitchTo().Alert();
                        valueattribute = alert.Text;
                        if (TestCase.TestDataSavedValues.ContainsKey(key))
                        {
                            TestCase.TestDataSavedValues.Remove(key);
                        }

                        //// Save value in defined dictionary
                        if (!string.IsNullOrEmpty(valueattribute))
                        {
                            TestCase.TestDataSavedValues.Add(key, valueattribute);
                        }

                        successMessage = "Fetched text \'" + valueattribute + "\' from alert box.";
                        break;
                    default:
                        return false;
                }

                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebDriverAlertHandler", Entities.Constants.Pass, successMessage, testStep.Remarks);
                return true;
            }
            catch (WebDriverTimeoutException ex)
            {
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebDriverAlertHandler", Entities.Constants.Fail, string.Format(Entities.Constants.Messages.WebDriverTimeoutException, ex.Message), testStep.Remarks);
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClassName, MethodBase.GetCurrentMethod().Name);
                return false;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebDriverAlertHandler", Entities.Constants.Fail, string.Format(Entities.Constants.Messages.DueToException, ex.Message), testStep.Remarks);
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClassName, MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        /// <summary>
        /// Web driver frame handler to control the frames.
        /// </summary>
        /// <param name="testStep">Test step as parameter.</param>
        /// <returns>Returns true or false.</returns>
        public static bool WebDriverFrameHandler(TestStep testStep)
        {
            try
            {
                IWebElement element;
                var action = Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]);
                var value = testStep.UiControl.UiControlSearchValue;
                WebDriverWait wait = new WebDriverWait(WebdriverBrowser.Driver, TimeSpan.FromSeconds(Convert.ToInt32(General.WaitForControlToExistTimeOut)));
                switch (action.ToUpper())
                {
                    case "SWITCHFRAMEBYINDEX":
                        int index = Convert.ToInt32(value); 
                        //WaitForControlToExist(testStep); //// Throws WebDriverTimeoutException
                        WebdriverBrowser.Driver.SwitchTo().DefaultContent();
                        WebdriverBrowser.Driver.SwitchTo().Frame(index);
                        break;
                    case "SWITCHFRAMEBYID":
                        wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt(testStep.UiControl.UiControlSearchValue)); //// Throws WebDriverTimeoutException
                        WebdriverBrowser.Driver.SwitchTo().DefaultContent();
                        WebdriverBrowser.Driver.SwitchTo().Frame(Convert.ToString(value));
                        break;
                    case "SWITCHFRAMEBYNAME":
                        wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt(testStep.UiControl.UiControlSearchValue)); //// Throws WebDriverTimeoutException
                        WebdriverBrowser.Driver.SwitchTo().DefaultContent();
                        WebdriverBrowser.Driver.SwitchTo().Frame(Convert.ToString(value));
                        break;
                    case "SWITCHFRAMEBYELEMENT":
                        WaitForControlToExist(testStep); //// Throws WebDriverTimeoutException
                        element = WebdriverBrowser.Driver.FindElement(By.XPath(testStep.UiControl.UiControlSearchValue));
                        WebdriverBrowser.Driver.SwitchTo().DefaultContent();
                        WebdriverBrowser.Driver.SwitchTo().Frame(element);
                        break;
                    default:
                        return false;
                }

                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebDriverFrameHandler", Entities.Constants.Pass, "Switched to frame \'" + value + "\' using \'" + action + "\'.", testStep.Remarks);
                return true;
            }
            catch (WebDriverTimeoutException ex)
            {
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebDriverFrameHandler", Entities.Constants.Fail, string.Format(Entities.Constants.Messages.WebDriverTimeoutException, ex.Message), testStep.Remarks);
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClassName, MethodBase.GetCurrentMethod().Name);
                return false;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebDriverFrameHandler", Entities.Constants.Fail, string.Format(Entities.Constants.Messages.DueToException, ex.Message), testStep.Remarks);
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClassName, MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        /// <summary>
        /// Web driver frame handler to switch to first frame on the page or the main document.
        /// </summary>
        /// <param name="testStep">Test step as parameter.</param>
        /// <returns>Returns true or false.</returns>
        public static bool WebDriverSwitchToDefaultFrame(TestStep testStep)
        {
            try
            {
                WebdriverBrowser.Driver.SwitchTo().DefaultContent();
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebDriverSwitchToDefaultFrame", Entities.Constants.Pass, "Switched to first frame on the page or main document containing the frames.", testStep.Remarks);
                return true;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebDriverSwitchToDefaultFrame", Entities.Constants.Fail, string.Format(Entities.Constants.Messages.DueToException, ex.Message), testStep.Remarks);
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClassName, MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        /// <summary>
        /// Web driver verify validations.
        /// </summary>
        /// <param name="testStep">Test step as parameter.</param>
        /// <value>Test step details.</value>
        /// <returns>True or false.</returns>
        public static bool WebDriverVerify(TestStep testStep)
        {
            try
            {
                if (testStep.Verification.VerificationType.ToUpper() == "DATABASEVALUE")
                {
                    Verify(testStep);
                }
                else
                {
                    switch (testStep.UiControl.UiControlSearchProperty.ToUpper())
                    {
                        case "ID":
                            if (Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]) != string.Empty && testStep.Verification.VerificationType.ToUpper() == "ISTEXT")
                            {
                                if (testStep.UiControl.UiControlType.ToUpper() == "HTMLCOMBOBOX")
                                {
                                    WaitForControlToExist(testStep); //// Throws WebDriverTimeoutException
                                    var identifier = WebdriverBrowser.Driver.FindElement(By.Id(testStep.UiControl.UiControlSearchValue));
                                    var select = new SelectElement(identifier);
                                    Assert.IsTrue(Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]).Contains(select.SelectedOption.Text));
                                    Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebDriverVerify", Entities.Constants.Pass, "Edit search property " + Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]) + " completed successfully", testStep.Remarks);
                                }
                                else
                                {
                                    WaitForControlToExist(testStep); //// Throws WebDriverTimeoutException
                                    Assert.IsTrue(Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]).Contains(WebdriverBrowser.Driver.FindElement(By.Id(testStep.UiControl.UiControlSearchValue)).Text));
                                    Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebDriverVerify", Entities.Constants.Pass, "Edit search property " + Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]) + " completed successfully", testStep.Remarks);
                                }
                            }
                            else
                            {
                                Verify(testStep);
                            }

                            break;
                        case "XPATH":
                            if (Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]) != string.Empty && testStep.Verification.VerificationType.ToUpper() == "ISTEXT")
                            {
                                if (testStep.UiControl.UiControlType.ToUpper() == "HTMLCOMBOBOX")
                                {
                                    WaitForControlToExist(testStep); //// Throws WebDriverTimeoutException
                                    var identifier = WebdriverBrowser.Driver.FindElement(By.XPath(testStep.UiControl.UiControlSearchValue));
                                    var select = new SelectElement(identifier);
                                    Assert.IsTrue(Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]).Contains(select.SelectedOption.Text));
                                    Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebDriverVerify", Entities.Constants.Pass, "Edit search property " + Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]) + " completed successfully", testStep.Remarks);
                                }
                                else
                                {
                                    WaitForControlToExist(testStep); //// Throws WebDriverTimeoutException
                                    Assert.IsTrue(Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]).Contains(WebdriverBrowser.Driver.FindElement(By.XPath(testStep.UiControl.UiControlSearchValue)).Text));
                                    Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebDriverVerify", Entities.Constants.Pass, "Edit search property " + Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]) + " completed successfully", testStep.Remarks);
                                }
                            }
                            else
                            {
                                Verify(testStep);
                            }

                            break;
                        default:
                            return true;
                    }
                }

                return true;
            }
            catch (WebDriverTimeoutException ex)
            {
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebDriverVerify", Entities.Constants.Fail, string.Format(Entities.Constants.Messages.WebDriverTimeoutException, ex.Message), testStep.Remarks);
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClassName, MethodBase.GetCurrentMethod().Name);
                return false;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebDriverVerify", Entities.Constants.Fail, string.Format(Entities.Constants.Messages.DueToException, ex.Message), testStep.Remarks);
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClassName, MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        /// <summary>
        /// Save user interface control.
        /// </summary>
        /// <param name="testStep">Test step.</param>
        /// <value>Test step value.</value>
        /// <returns>True or false.</returns>
        public static bool WebDriverSaveUIControl(TestStep testStep)
        {
            try
            {
                string valueattribute = string.Empty;
                dynamic key = Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]);
                switch (testStep.UiControl.UiControlSearchProperty.ToUpper())
                {
                    case "ID":
                        WaitForControlToExist(testStep); //// Throws WebDriverTimeoutException
                        valueattribute = WebdriverBrowser.Driver.FindElement(By.Id(testStep.UiControl.UiControlSearchValue)).Text;
                        break;
                    case "XPATH":
                        ////Pagination store count value
                        var keyValue = key;
                        if (keyValue.Contains(Entities.Constants.Pagination))
                        {
                            string[] stringSeparators = { Entities.Constants.PipeDelimitor };
                            string[] resultStrings = keyValue.Split(stringSeparators, StringSplitOptions.None);
                            if (resultStrings != null && resultStrings.Length > 0)
                            {
                                var expressionValue = resultStrings[0];
                                var paginationKey = resultStrings[1];
                                var expressionLength = expressionValue.Length;
                                WaitForControlToExist(testStep); //// Throws WebDriverTimeoutException
                                var getLabelText = WebdriverBrowser.Driver.FindElement(By.XPath(testStep.UiControl.UiControlSearchValue)).Text;
                                key = paginationKey;
                                valueattribute = getLabelText.Substring(expressionLength);
                            }
                        }
                        else
                        {
                            WaitForControlToExist(testStep); //// Throws WebDriverTimeoutException
                            valueattribute = WebdriverBrowser.Driver.FindElement(By.XPath(testStep.UiControl.UiControlSearchValue)).Text;
                        }

                        break;
                    default:
                        return true;
                }

                if (TestCase.TestDataSavedValues.ContainsKey(key))
                {
                    TestCase.TestDataSavedValues.Remove(key);
                }

                //// Save value
                if (!string.IsNullOrEmpty(valueattribute))
                {
                    TestCase.TestDataSavedValues.Add(key, valueattribute);
                }

                UiActions objUiActions = new UiActions();
                objUiActions.BufferApps();
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebDriverSaveUIControl", Entities.Constants.Pass, "Edit search property " + Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]) + " completed successfully", testStep.Remarks);
                return true;
            }
            catch (WebDriverTimeoutException ex)
            {
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebDriverSaveUIControl", Entities.Constants.Fail, string.Format(Entities.Constants.Messages.WebDriverTimeoutException, ex.Message), testStep.Remarks);
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClassName, MethodBase.GetCurrentMethod().Name);
                return false;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebDriverSaveUIControl", Entities.Constants.Fail, string.Format(Entities.Constants.Messages.DueToException, ex.Message), testStep.Remarks);
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClassName, MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        /// <summary>
        /// Web component pagination validation.
        /// </summary>
        /// <param name="testStep">Test step.</param>
        /// <value>Test step value.</value>
        /// <returns>true or false.</returns>
        public static bool WebPaginationIteration(TestStep testStep)
        {
            try
            {
                var value = testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)];
                int requiredValue = Convert.ToInt32(value);
                WaitForControlToExist(testStep); //// Throws WebDriverTimeoutException
                IWebElement element = WebdriverBrowser.Driver.FindElement(By.XPath(testStep.UiControl.UiControlSearchValue));
                for (int i = 1; i < requiredValue; i++)
                {
                    element.Click();
                    Thread.Sleep(2000);
                }

                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebPaginationIteration", Entities.Constants.Pass, "Pagination iteration completed successfully for iteration count (" + requiredValue + ")", testStep.Remarks);
                return true;
            }
            catch (WebDriverTimeoutException ex)
            {
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebPaginationIteration", Entities.Constants.Fail, string.Format(Entities.Constants.Messages.WebDriverTimeoutException, ex.Message), testStep.Remarks);
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClassName, MethodBase.GetCurrentMethod().Name);
                return false;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WebPaginationIteration", Entities.Constants.Fail, string.Format(Entities.Constants.Messages.DueToException, ex.Message), testStep.Remarks);
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClassName, MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        /// <summary>
        /// Launch web browser.
        /// </summary>
        /// <param name="testStep">Test step.</param>
        /// <value>Test step value.</value>
        /// <returns>True or false.</returns>
        public static bool LaunchWebDriverBrowser(TestStep testStep)
        {
            try
            {
                //// Launch browser
                WebdriverBrowser.Launch(Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]));
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "LaunchWebDriverBrowser", Entities.Constants.Pass, "WebDriverBrowser launched sucessfully", testStep.Remarks);
                return true;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "LaunchWebDriverBrowser", Entities.Constants.Fail, string.Format(Entities.Constants.Messages.DueToException, ex.Message), testStep.Remarks);
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClassName, MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        /// <summary>
        ///  Wait for UI to Load (pause the playback).
        /// </summary>
        /// <param name="testStep">Test step.</param>
        /// <value>Test step value.</value>
        /// <returns>True or false.</returns>
        public static bool WaitForUI(TestStep testStep)
        {
            try
            {
                //// Wait some seconds if required
                if (!string.IsNullOrEmpty(testStep.TestData.ContainsValue(testStep.TestDataKeyToUse).ToString()))
                {
                    dynamic secondsToWaitForUi = Convert.ToInt32(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]);
                    if (secondsToWaitForUi > 0)
                    {
                        Thread.Sleep(secondsToWaitForUi * 1000);
                    }
                }

                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WaitForUI", Entities.Constants.Pass, "Wait for UI to load completed successfully.", testStep.Remarks);
                return true;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WaitForUI", Entities.Constants.Fail, string.Format(Entities.Constants.Messages.DueToException, ex.Message), testStep.Remarks);
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClassName, MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        /// <summary>
        ///  Send keys to simulate keyboard actions.
        /// </summary>
        /// <param name="testStep">Test step.</param>
        /// <value>Test step value.</value>
        /// <returns>True or false.</returns>
        public static bool SendKeys(TestStep testStep)
        {
            try
            {
                var error = new StringBuilder().Append("Test data object does not exist for test step ");
                error.Append(testStep.TestStepNumber);
                error.Append(".");

                var testStepError = new StringBuilder().Append("Test data with key ");
                testStepError.Append(testStep.TestDataKeyToUse);
                testStepError.Append("does not exist for test step ");
                testStepError.Append(testStep.TestStepNumber);
                testStepError.Append(".");
                //// Send keys
                Assert.IsTrue(testStep.TestData != null, error.ToString());
                Assert.IsTrue(testStep.TestData.ContainsValue(testStep.TestDataKeyToUse).ToString() != null, testStepError.ToString());
                dynamic valuetosend = Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]).ToUpper().Replace(" ", string.Empty);

                if (valuetosend.Contains(Entities.Constants.UiActions.Alt))
                {
                    var userName = WebdriverBrowser.Driver.SwitchTo().ActiveElement();
                    userName.SendKeys(Keys.Alt + valuetosend);
                }
                else if (valuetosend.Contains(Entities.Constants.UiActions.Shift))
                {
                    var userName = WebdriverBrowser.Driver.SwitchTo().ActiveElement();
                    userName.SendKeys(Keys.Shift + valuetosend);
                }
                else if (valuetosend.Contains(Entities.Constants.UiActions.Control))
                {
                    var userName = WebdriverBrowser.Driver.SwitchTo().ActiveElement();
                    userName.SendKeys(Keys.Control + valuetosend);
                }
                else
                {
                    switch (Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]).ToUpper().Replace(" ", string.Empty))
                    {
                        case Entities.Constants.UiActions.EnterBracket:
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(Keys.Enter);
                            break;
                        case Entities.Constants.UiActions.DownBracket:
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(Keys.Down);
                            break;
                        case "{UP}":
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(Keys.Up);
                            break;
                        case "{RIGHT}":
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(Keys.Right);
                            break;
                        case "{LEFT}":
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(Keys.Left);
                            break;
                        case "{TAB}":
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(Keys.Tab);
                            break;
                        case "{PAGEUP}":
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(Keys.PageUp);
                            break;
                        case "{PAGEDOWN}":
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(Keys.PageDown);
                            break;
                        case "{END}":
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(Keys.End);
                            break;
                        case "{HOME}":
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(Keys.Home);
                            break;
                        case "{SPACE}":
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(Keys.Space);
                            break;
                        case "{F1}":
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(Keys.F1);
                            break;
                        case "{F2}":
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(Keys.F2);
                            break;
                        case "{F3}":
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(Keys.F3);
                            break;
                        case "{F4}":
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(Keys.F4);
                            break;
                        case "{F5}":
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(Keys.F5);
                            break;
                        case "{F6}":
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(Keys.F6);
                            break;
                        case "{F7}":
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(Keys.F7);
                            break;
                        case "{F8}":
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(Keys.F8);
                            break;
                        case "{F9}":
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(Keys.F9);
                            break;
                        case "{F10}":
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(Keys.F10);
                            break;
                        case "{F11}":
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(Keys.F11);
                            break;
                        case "{F12}":
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(Keys.F12);
                            break;
                        default:
                            WebdriverBrowser.Driver.SwitchTo().ActiveElement().SendKeys(valuetosend);
                            break;
                    }
                }

                var sendKeysCompleted = new StringBuilder().Append("Test data object does not exist for test step ");
                error.Append(testStep.TestData.ContainsValue(testStep.TestDataKeyToUse));
                error.Append(" completed successfully");
                error.Append(".");

                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "SendKeys", Entities.Constants.Pass, sendKeysCompleted.ToString(), testStep.Remarks);
                return true;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "SendKeys", Entities.Constants.Fail, string.Format(Entities.Constants.Messages.DueToException, ex.Message), testStep.Remarks);
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClassName, MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        /// <summary>
        ///  Buffer values to read.
        /// </summary>
        /// <param name="configStep">Configurations step.</param>
        /// <value>Configuration value.</value>
        /// <returns>True or false.</returns>
        public static bool BufferValues(ConfigStep configStep)
        {
            try
            {
                dynamic key = configStep.TestVariableName;
                dynamic valueAttribute = configStep.TestDataValue;

                if (TestCase.TestDataSavedValues.ContainsKey(key))
                {
                    TestCase.TestDataSavedValues.Remove(key);
                }

                if (!string.IsNullOrEmpty(valueAttribute))
                {
                    TestCase.TestDataSavedValues.Add(key, valueAttribute);
                }

                return true;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(configStep.TestConfigKeyToUse, configStep.TestStepNo, "SaveUIControlValueAttribute", Entities.Constants.Fail, string.Format(Entities.Constants.Messages.DueToException, ex.Message), configStep.Remarks);
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClassName, MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        /// <summary>
        ///  Waits for the element for a given time if the element is not immediately available.
        /// </summary>
        /// <param name="testStep">Test step.</param>
        /// <value>Test step value.</value>
        /// <returns>True or false.</returns>
        public static bool WaitForControlToExist(TestStep testStep)
        {
            WebDriverWait wait = new WebDriverWait(WebdriverBrowser.Driver, TimeSpan.FromSeconds(Convert.ToInt32(General.WaitForControlToExistTimeOut)));
            if (Convert.ToString(testStep.UiControl.UiControlSearchProperty.ToUpper()) == "ID")
            {
                wait.Until<IWebElement>(d => d.FindElement(By.Id(testStep.UiControl.UiControlSearchValue)));
            }
            else if (Convert.ToString(testStep.UiControl.UiControlSearchProperty.ToUpper()) == "XPATH")
            {
                wait.Until<IWebElement>(d => d.FindElement(By.XPath(testStep.UiControl.UiControlSearchValue)));
            }
            
            return true;
        }

        /// <summary>
        ///  Waits for the element for a given time if the element is not immediately available.
        /// </summary>
        /// <param name="testStep">Test step.</param>
        /// <value>Test step value.</value>
        /// <returns>True or false.</returns>
        public static bool WaitForElementToBeEnabled(TestStep testStep)
        {
            WebDriverWait wait = new WebDriverWait(WebdriverBrowser.Driver, TimeSpan.FromSeconds(Convert.ToInt32(General.WaitForControlToExistTimeOut)));
            if (Convert.ToString(testStep.UiControl.UiControlSearchProperty.ToUpper()) == "ID")
            {
                wait.Until(ExpectedConditions.ElementToBeClickable(By.Id(testStep.UiControl.UiControlSearchValue)));
            }
            else if (Convert.ToString(testStep.UiControl.UiControlSearchProperty.ToUpper()) == "XPATH")
            {
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(testStep.UiControl.UiControlSearchValue)));
            }
            return true;
        }

        /// <summary>
        ///  Verify the value.
        /// </summary>
        /// <param name="testStep">Test step.</param>
        /// <value>Test step value.</value>
        private static void Verify(TestStep testStep)
        {
            try
            {
                Verifications.Verifications objVerifications = new Verifications.Verifications();
                switch (testStep.Verification.VerificationType.ToUpper())
                {
                    case Entities.Constants.UiActions.BrowswrExist:
                        Verifications.Verifications.BrowserExist(testStep);
                        break;
                    case Entities.Constants.UiActions.BrowserNotExists:
                        Verifications.Verifications.BrowserNotExist(testStep);
                        break;
                    case Entities.Constants.UiActions.IsEnabled:
                        Verifications.Verifications.IsEnabled(testStep);
                        break;
                    case Entities.Constants.UiActions.IsDisplayed:
                        Verifications.Verifications.IsDisplayed(testStep);
                        break;
                    case Entities.Constants.UiActions.IsSelected:
                        Verifications.Verifications.IsSelected(testStep);
                        break;
                    case Entities.Constants.UiActions.DataBaseValue:
                        objVerifications.DatabaseValue(testStep);
                        break;
                    case Entities.Constants.UiActions.IsDisabled:
                        Verifications.Verifications.IsDisabled(testStep);
                        break;
                    default:
                        Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "WaitForUI", Entities.Constants.Fail, "VerificationType " + testStep.Verification.VerificationType + " is not supported", testStep.Remarks);
                        break;
                }
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "Verify", Entities.Constants.Fail, string.Format(Entities.Constants.Messages.DueToException, ex.Message), testStep.Remarks);
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClassName, MethodBase.GetCurrentMethod().Name);
            }
        }

        /// <summary>
        ///  Buffer application value.
        /// </summary>
        private void BufferApps()
        {
            var applicationClass = new ApplicationClass();
            try
            {
                TestCase.TestStepList = new List<TestStep>();
                TestCase.UiControls = new List<UiControl>();
                ConfigStep.TestConfigNames = new List<ConfigStep>();
                TestCase.Verifications = new List<Verification>();

                Data.LoadUiControls(applicationClass);

                Data.LoadVerifications(applicationClass);

                Data.LoadTestCases(applicationClass);
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClassName, MethodBase.GetCurrentMethod().Name);
            }
            finally
            {
                WorkBookUtility.CloseExcel(applicationClass);
            }
        }
    }
}