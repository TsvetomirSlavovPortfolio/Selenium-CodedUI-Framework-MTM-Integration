// <copyright file="Result.cs" company="Infosys Ltd.">
//  Copyright (c) Infosys Ltd. All rights reserved.
// </copyright>
// <summary>Result.cs Generates Result based on test results.</summary>
namespace INF.Selenium.TestAutomation.Utilities
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;
    using Configuration;
    using Entities;

    /// <summary>
    /// Collection of test results.
    /// </summary>
    public class Result
    {
        /// <summary>
        /// Gets Collection of test results.
        /// </summary>
        /// <value>Test Result value.</value>
        /// 
        public static Queue<TestResult> TestStepsResultsCollection = new Queue<TestResult>();

        public static TestResult Obj = default(TestResult);
        
        /// <summary>
        /// Called at the end of every step to store results in collection.
        /// </summary>
        /// <param name="testDataIterationNr">Step Test Data Iteration Number.</param>
        /// <param name="stepNr">Step Number.</param>
        /// <param name="description">Step description.</param>
        /// <param name="result">Step pass/fail result.</param>
        /// <param name="comments">Step comments, exception details if applicable.</param>
        /// <param name="remarks">Step Remarks.</param>
        public static void PassStepOutandSave(string testDataIterationNr, string stepNr, string description, string result, string comments, string remarks)
        {
            try
            {
                var logObj = new TestResult
                {
                    TestDataIterationNr = testDataIterationNr,
                    StepNr = stepNr,
                    Result = result,
                    Description = description,
                    Comment = comments,
                    Remarks = remarks
                };
               Result objResult = new Result();
               objResult.AddTestScriptResulttoCollection(logObj);
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.Result, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }

        /// <summary>
        /// Enqueue the step result on the queue TestStepsResultsCollection.
        /// </summary>
        /// <returns> Returns Fail if any Step has Fail result, True otherwise.</returns>
        public static string GetTestScriptResult()
        {
            try
            {
                if (TestStepsResultsCollection.Count == 0)
                {
                    return Constants.Fail;
                }

                foreach (var obj in TestStepsResultsCollection)
                {
                    if (obj.Result == Constants.Fail)
                    {
                        return Constants.Fail;
                    }
                }

                return Constants.Pass;
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.Result, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }

        /// <summary>
        ///     Enqueue the step result on the queue testStepsResultsCollection.
        /// </summary>
        /// <param name="obj">Refers the test step results.</param>
        private void AddTestScriptResulttoCollection(TestResult obj)
        {
            TestStepsResultsCollection.Enqueue(obj);
        }
    }
}