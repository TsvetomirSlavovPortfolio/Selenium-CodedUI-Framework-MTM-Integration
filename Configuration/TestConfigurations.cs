// <copyright file="TestConfigurations.cs" company="Infosys Ltd.">
//  Copyright (c) Infosys Ltd. All rights reserved.
// </copyright>
// <summary>TestConfigurations.cs class instructs framework to use the test configurations.</summary>
namespace INF.Selenium.TestAutomation.Configuration
{
    using System;
    using System.Reflection;
    using Entities;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using UI;
    using Utilities;

    /// <summary>
    /// TestConfigurations reads test configurations value from TestConfigurations excel file.
    /// </summary>
    public class TestConfigurations
    {
        /// <summary>
        /// Configuration reads test configurations value from TestConfigurations excel file.
        /// </summary>
        public void Configuration()
        {
            try
            {
                var configCount = ConfigStep.TestDataConfigCount;

                //// Run the configuration for saving the test data in mentioned variable name
                for (var configIndex = 1; configIndex <= configCount; configIndex++)
                {
                    //// Loop each testcase steps
                    var isSuccess = true;
                    var configList = ConfigStep.TestConfigNames;
                    foreach (var configData in configList)
                    {
                        switch (configData.TestDataType.ToUpper())
                        {
                            case Constants.String:
                            case Constants.Integer:
                                isSuccess = UiActions.BufferValues(configData);
                                break;
                            default:
                                configData.TestConfigKeyToUse = configIndex.ToString();
                                configData.ConfigAction = string.Empty;
                                configData.Remarks = string.Empty;
                                Result.PassStepOutandSave(configData.TestConfigKeyToUse, configData.TestStepNo, Constants.TestIterations, Constants.Fail, string.Format(Constants.Messages.NotSupported, configData.ConfigAction), configData.Remarks);
                                isSuccess = false;
                                break;
                        }

                        if (!isSuccess)
                        {
                            break;
                        }
                    }

                    if (!isSuccess && Result.GetTestScriptResult() == Constants.Fail)
                    {
                        Assert.Fail(Constants.Messages.TestCaseFailedError, Reporting.FilePath);
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.TestConfigurations, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }
    }
}