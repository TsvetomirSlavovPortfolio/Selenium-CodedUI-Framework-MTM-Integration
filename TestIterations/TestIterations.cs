// <copyright file="TestIterations.cs" company="Infosys Ltd.">
//  Copyright (c) Infosys Ltd. All rights reserved.
// </copyright>
// <summary>TestIterations.cs class Reads test cases sheets and sets iterations.</summary>
namespace INF.Selenium.TestAutomation.TestIterations
{
    using Entities;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Test iterations. Defines test suite to run the set of test cases.
    /// </summary>
    [TestClass]
    public class TestIterations : BaseTestClass
    {

        #region Test Case - Amazon Demo

        /// <summary>
        /// Test Method.
        /// </summary>
        [TestCategory(Constants.UnitTest)]
        [TestMethod]
        [Timeout(TestTimeout.Infinite)]
        public void TC001_Amazondemo()
        {
            string applicationName = "AMAZON";
            string testCaseName = "TC001_Amazondemo";
            string testCaseDescription = "Browse and add products to the cart";
            string testCaseFileName = "TC001_Amazondemo.xlsx";
            ExecuteTestCases.ExecuteTestCase(applicationName, testCaseName, testCaseDescription, testCaseFileName);
        }

        #endregion

        #region Test Case - Database Demo

        /// <summary>
        /// Test Method.
        /// </summary>
        [TestCategory(Constants.UnitTest)]
        [TestMethod]
        [Timeout(TestTimeout.Infinite)]
        public void TC003_Database()
        {
            string applicationName = "AMAZON";
            string testCaseName = "TC003_Database";
            string testCaseDescription = "Create records in database";
            string testCaseFileName = "TC003_Database.xlsx";
            ExecuteTestCases.ExecuteTestCase(applicationName, testCaseName, testCaseDescription, testCaseFileName);
        }

        #endregion

        #region Test Case - Amazon Demo API 2

        /// <summary>
        /// Test Method.
        /// </summary>
        [TestCategory(Constants.UnitTest)]
        [TestMethod]
        [Timeout(TestTimeout.Infinite)]
        public void TC004_Amazondemo_API2()
        {
            string applicationName = "AMAZON";
            string testCaseName = "TC004_Amazondemo_API2";
            string testCaseDescription = "UI Demo Test Case with different data sets taken from API";
            string testCaseFileName = "TC004_Amazondemo_API2.xlsx";
            ExecuteTestCases.ExecuteTestCase(applicationName, testCaseName, testCaseDescription, testCaseFileName);
        }

        #endregion

        #region Test Case - Amazon Demo API 3

        /// <summary>
        /// Test Method.
        /// </summary>
        [TestCategory(Constants.UnitTest)]
        [TestMethod]
        [Timeout(TestTimeout.Infinite)]
        public void TC004_Amazondemo_API3()
        {
            string applicationName = "AMAZON";
            string testCaseName = "TC004_Amazondemo_API3";
            string testCaseDescription = "UI Demo Test Case with different data sets taken from API";
            string testCaseFileName = "TC004_Amazondemo_API3.xlsx";
            ExecuteTestCases.ExecuteTestCase(applicationName, testCaseName, testCaseDescription, testCaseFileName);
        }

        #endregion

        #region Test Case - TC001_Frame_AlertTest

        /// <summary>
        /// Test Method.
        /// </summary>
        [TestCategory(Constants.UnitTest)]
        [TestMethod]
        [Timeout(TestTimeout.Infinite)]
        public void TC001_Frame_AlertTest()
        {
            string applicationName = "W3 Schools";
            string testCaseName = "TC001_Frame_AlertTest";
            string testCaseDescription = "UI Demo Test Case for Frame and Alert Handler";
            string testCaseFileName = "TC001_Frame_AlertTest.xlsx";
            ExecuteTestCases.ExecuteTestCase(applicationName, testCaseName, testCaseDescription, testCaseFileName);
        }

        #endregion

    }
}