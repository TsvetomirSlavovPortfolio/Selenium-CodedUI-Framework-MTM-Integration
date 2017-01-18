// <copyright file="LogHelper.cs" company="Infosys Ltd.">
//  Copyright (c) Infosys Ltd. All rights reserved.
// </copyright>
// <summary>LogHelper.cs class helps framework to write runtime logs</summary>
namespace INF.Selenium.TestAutomation.Configuration
{
    using System;
    using System.Configuration;
    using System.IO;
    using System.Linq;
    using System.Text;
    using Entities;

    /// <summary>LogHelper class collects runtime errors.</summary>
    public class LogHelper
    {
        /// <summary>
        /// EventLog collects runtime event errors.
        /// </summary>
        /// <param name="ex">Error Message being returned.</param>
        /// <param name="className">Error Message being returned from class file.</param>
        /// <param name="method">Error Message being returned from class file and its method.</param>
        public static void ErrorLog(Exception ex, string className = "", string method = "")
        {
            try
            {
                var logMessage = new StringBuilder().Append(TestCases.TestCases.TestSessionId)
                    .Append(Constants.Tab)
                    .Append(DateTime.Now)
                    .Append(Constants.Tab)
                    .Append(className).Append(Constants.Tab).Append(method).Append(Constants.Tab).Append(ex.Message)
                    .Append(Constants.Tab).Append(ex.StackTrace);

                WriteLog(logMessage.ToString(), Constants.ErrorLog);
            }
            catch
            {
                WriteLog(Constants.ErrorLog, Constants.ErrorLog);
            }
        }

        /// <summary>
        /// Error log collects runtime errors.
        /// </summary>
        /// <param name="message">Exception Message being returned.</param>
        /// <param name="className">Error Message being returned from class file.</param>
        /// <param name="method">Error Message being returned from class file and its method.</param>
        public void EventLog(string message, string className = "", string method = "")
        {
            try
            {
                var logMessage =
                    new StringBuilder().Append(TestCases.TestCases.TestSessionId)
                        .Append(Constants.Tab)
                        .Append(DateTime.Now)
                        .Append(Constants.Tab)
                        .Append(className)
                        .Append(Constants.Tab)
                        .Append(method)
                        .Append(Constants.Tab)
                        .Append(message);

                WriteLog(logMessage.ToString(), Constants.EventLog);
            }
            catch
            {
                WriteLog(Constants.ErrorLog, Constants.ErrorLog);
            }
        }

        /// <summary>
        /// Write log writes log in error log file.
        /// </summary>
        /// <param name="message">Exception Message being returned.</param>
        /// <param name="logFileName">File Name of log file.</param>
        private static void WriteLog(string message, string logFileName)
        {
            //// Compose a string that consists of three lines.
            var filePath = Environment.CurrentDirectory.Split(Convert.ToChar(Constants.DoubleBackslash));
            filePath =
                filePath.Where(item => item != filePath[filePath.Length - 1])
                    .Where(item => item != filePath[filePath.Length - 2])
                    .Where(item => item != filePath[filePath.Length - 3])
                    .ToArray();

            var rootPath = string.Join(Constants.DoubleBackslash, filePath);

            var rootFilePath =
                new StringBuilder().Append(
                    string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get(Constants.AppSetting.RootFilePath))
                        ? rootPath
                        : ConfigurationManager.AppSettings.Get(Constants.AppSetting.RootFilePath));

            if (!rootFilePath.ToString().Last().ToString().Equals(Constants.DoubleBackslash))
            {
                rootFilePath.Append(Constants.DoubleBackslash);
            }

            Directory.CreateDirectory(rootFilePath.Append("Logs").ToString());

            using (
                var sw =
                    File.AppendText(
                        rootFilePath.Append(Constants.DoubleBackslash)
                            .Append(logFileName)
                            .Append(Constants.Txt)
                            .ToString()))
            {
                sw.WriteLine(message);
            }
        }
    }
}