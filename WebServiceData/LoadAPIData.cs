// <copyright file="LoadAPIData.cs" company="Infosys Ltd.">
//  Copyright (c) Infosys Ltd. All rights reserved.
// </copyright>
// <summary>LoadAPIData.cs Stores and gives values from data captured from different Webservice APIs</summary>
using INF.Selenium.TestAutomation.Configuration;
using INF.Selenium.TestAutomation.Entities;
using INF.Selenium.TestAutomation.Utilities;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;

namespace INF.Selenium.TestAutomation.WebServiceData
{
    public class LoadAPIData
    {
        /// <summary>
        /// Stores the data from the APIs
        /// </summary>
        private static List<JArray> responseDataList = new List<JArray>();

        /// <summary>
        /// Stores the request URL
        /// </summary>
        private string requestURL;

        /// <summary>
        /// Webrequest object
        /// </summary>
        private HttpWebRequest webRequest;

        /// <summary>
        /// Webresponse object
        /// </summary>
        private HttpWebResponse response;

        /// <summary>
        /// StreamReader object
        /// </summary>
        private StreamReader streamReader;

        /// <summary>
        /// Stores the data from the API in a string
        /// </summary>
        private String singleAPIResponseData;

        /// /// <summary>
        /// This function Loads Web Service API Data.
        /// </summary>
        /// <param name="applicationClass">Application class.</param>
        /// <value>Application class value.</value>
        public void LoadAPIResponseData(ApplicationClass applicationClass)
        {
            try
            {
                var workbook = WorkBookUtility.OpenWorkBook(applicationClass, TestCase.RootFilePath + new StringBuilder().Append(ConfigurationManager.AppSettings.Get(Entities.Constants.AppSetting.LoadWebServiceAPIData)));
                dynamic worksheet = (Worksheet)workbook.Worksheets[1];
                var rowsCount = worksheet.UsedRange.Rows.Count;
                var cellCount = worksheet.UsedRange.Columns.Count + 1;

                for (var rowindex = 2; rowindex <= rowsCount; rowindex++)
                {
                    if (string.IsNullOrEmpty(Convert.ToString(worksheet.Cells[rowindex, 1].value)))
                    {
                        break; //// reading the sheet untill the first empty row
                    }
                    else
                    {
                        requestURL = Convert.ToString(worksheet.Cells[rowindex, 3].value);
                        webRequest = (HttpWebRequest)HttpWebRequest.Create(requestURL);
                        response = (HttpWebResponse)webRequest.GetResponse();
                        streamReader = new StreamReader(response.GetResponseStream());
                        singleAPIResponseData = streamReader.ReadToEnd();
                        if (!singleAPIResponseData.StartsWith("["))
                        {
                            singleAPIResponseData = singleAPIResponseData.Replace(singleAPIResponseData, "[" + singleAPIResponseData + "]");
                        }
                        dynamic jsonObject = JArray.Parse(singleAPIResponseData);
                        responseDataList.Add(jsonObject);
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
        /// Get Web Service API Data.
        /// </summary>
        /// <param name="testStep">Test step.</param>
        /// <value>Test step value.</value>
        /// <returns>integervalue</returns>
        public static string GetSavedAPIData(TestStep testStep)
        {
            try
            {
                dynamic requiredData;
                int index1=0;
                int index2=0;
                string key1;
                string key2;
                string key3;
                
                string testDataValue = Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]);
                string[] separator = { Entities.Constants.PipeDelimitor };
                string[] testDataValueSplit = testDataValue.Split(separator, StringSplitOptions.None);
                string[] testDataRequiredValueSplit = testDataValueSplit[1].Split('.');
                int totalValues = testDataRequiredValueSplit.Count();
                index1 = Convert.ToInt32(testDataRequiredValueSplit[0]);
                index2 = Convert.ToInt32(testDataRequiredValueSplit[1]);
                
                if (totalValues == 3)
                {
                    key1 = (testDataRequiredValueSplit[2].ToString());
                    requiredData = responseDataList[index1][index2][key1];
                    return requiredData.ToString();
                }
                else if (totalValues == 4)
                {
                    key1 = (testDataRequiredValueSplit[2].ToString());
                    key2 = (testDataRequiredValueSplit[3].ToString());
                    requiredData = responseDataList[index1][index2][key1][key2];
                    return requiredData.ToString();
                }
                else if (totalValues == 5)
                {
                    key1 = (testDataRequiredValueSplit[2].ToString());
                    key2 = (testDataRequiredValueSplit[3].ToString());
                    key3 = (testDataRequiredValueSplit[4].ToString());
                    requiredData = responseDataList[index1][index2][key1][key2][key3];
                    return requiredData.ToString();
                }
                else
                {
                    return "Not Found";
                }
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.Data, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }
    }
}
