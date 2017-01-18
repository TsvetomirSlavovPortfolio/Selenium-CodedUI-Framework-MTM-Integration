// <copyright file="ExecuteTestCases.cs" company="Infosys Ltd.">
//  Copyright (c) Infosys Ltd. All rights reserved.
// </copyright>
// <summary>ExecuteTestCases.cs Executes the test case defined in a Test Method.</summary>
using INF.Selenium.TestAutomation.Configuration;
using INF.Selenium.TestAutomation.Entities;
using INF.Selenium.TestAutomation.Utilities;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace INF.Selenium.TestAutomation.TestIterations
{
    class ExecuteTestCases
    {
        /// <summary>
        /// Executes a Test Case Method.
        /// </summary>
        /// <param name="applicationName">Name of the application.</param>
        /// <param name="testCaseName">Name of the test case.</param>
        /// <param name="testCaseDescription">Description of the test case.</param>
        /// <param name="testCaseFileName">File name of the test case to be executed.</param>
        public static void ExecuteTestCase(String applicationName, string testCaseName, string testCaseDescription, string testCaseFileName)
        {
            try
            {
                TestCase.Application = applicationName;
                TestCase.Name = testCaseName;
                TestCase.Description = testCaseDescription;
                //// File name of test case to be executed with .xlsx extension added
                TestCase.FileName = testCaseFileName;

                //// Initiliaze test step result
                if (Result.TestStepsResultsCollection == null)
                {
                    Result.TestStepsResultsCollection = new System.Collections.Generic.Queue<TestResult>();
                }
                else
                {
                    Result.TestStepsResultsCollection.Clear();
                }

                //// Initiliaze test case reporting
                var errorMessage = string.Empty;
                if (!Reporting.CreateExcelSheet(ref errorMessage))
                {
                    Assert.Inconclusive(Constants.Messages.ReportSheetError, errorMessage);
                }
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.BaseTestClass, MethodBase.GetCurrentMethod().Name);
                throw;
            }

            try
            {
                TestCases.TestCases objTestCases = new TestCases.TestCases();
                objTestCases.Execute();
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.TestIterationsClassName, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }
    }
}
