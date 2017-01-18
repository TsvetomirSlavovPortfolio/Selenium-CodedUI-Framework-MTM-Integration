// <copyright file="EmailNotification.cs" company="Infosys Ltd.">
//  Copyright (c) Infosys Ltd. All rights reserved.
// </copyright>
// <summary>EmailNotification.cs to send the email notifcation with bthe Test Report.</summary>
namespace INF.Selenium.TestAutomation.Utilities
{
    using System;
    using System.Configuration;
    using System.Data;
    using System.Data.OleDb;
    using System.IO;
    using System.Linq;
    using System.Net.Mail;
    using System.Reflection;
    using System.Text;
    using Configuration;
    using Entities;
    using TestIterations;

    /// <summary>
    /// Email Notifications.
    /// </summary>
    public class EmailNotification
    {
        /// <summary>
        /// Sends the Report after execution.
        /// </summary>
        public void SendTestReportMail()
        {
            if (ConfigurationManager.AppSettings.Get(Constants.AppSetting.EmailNotifcationRequired) == "Yes")
            {
                try
                {
                    var testReportPath = TestCase.RootFilePath +
                                         ConfigurationManager.AppSettings.Get(Constants.AppSetting.TestReportPath);
                    var directory = new DirectoryInfo(testReportPath);
                    directory = directory.GetDirectories().OrderByDescending(f => f.LastWriteTime).First();
                    var myFile = directory.GetFiles().OrderByDescending(f => f.LastWriteTime).First();
                    var fileShortName = myFile.ToString();
                    var stmpServer = ConfigurationManager.AppSettings.Get(Constants.AppSetting.SmtpServerHostAddress);
                    var port = Convert.ToInt32(ConfigurationManager.AppSettings.Get(Constants.AppSetting.SmtpServerPort));
                    var from = new MailAddress(ConfigurationManager.AppSettings.Get(Constants.AppSetting.MailFrom));
                    string[] multiToAddress = ConfigurationManager.AppSettings.Get(Constants.AppSetting.MailTo).Split(';');
                    var indexOfFileExtention = fileShortName.IndexOf(".xlsx");
                    var mailSubject = "Test Execution Report: \"" + fileShortName.Substring(0, indexOfFileExtention) + "\"";
                    var objSmtp = new SmtpClient(stmpServer);
                    objSmtp.Port = port;
                    objSmtp.UseDefaultCredentials = true;
                    var objMsg = new MailMessage();
                    objMsg.From = from;
                    foreach (var mailAddress in multiToAddress)
                    {
                        objMsg.To.Add(mailAddress);
                    }

                    objMsg.Subject = mailSubject;
                    objMsg.IsBodyHtml = true;
                    if (!testReportPath.Last().ToString().Equals(Constants.DoubleBackslash))
                    {
                        testReportPath += Constants.DoubleBackslash;
                    }

                    var filepath = testReportPath + directory + "\\" + fileShortName;
                    objMsg.Body = this.GetBody(filepath);
                    if (filepath.Contains(".xlsx"))
                    {
                        objMsg.Attachments.Add(new Attachment(filepath));
                    }

                    objSmtp.Send(objMsg);
                }
                catch (Exception ex)
                {
                    LogHelper.ErrorLog(ex, Constants.ClassName.EmailNotification, MethodBase.GetCurrentMethod().Name);
                }
            }
        }

        /// <summary>
        /// Query to fetch data from the test report.
        /// </summary>
        /// <param name="fileShortName">File Short Name as parameter.</param>
        /// <param name="dataSourceHeader">Data Source Header as parameter.</param>
        /// <param name="dataSourceBody">Data Source Body as parameter.</param>
        private void Query(string fileShortName, out DataTable dataSourceHeader, out DataTable dataSourceBody)
        {
            OleDbConnection connection;
            OleDbDataAdapter adapter;

            string fileName = fileShortName;
            string conn = string.Format(ConfigurationManager.AppSettings.Get(Constants.AppSetting.ExcelConStr), fileName);
            using (connection = new OleDbConnection(conn))
            {
                string queryHeader = "select * from [TestIterations$A1:B3]";
                using (adapter = new OleDbDataAdapter(queryHeader, connection))
                {
                    dataSourceHeader = new DataTable();
                    adapter.Fill(dataSourceHeader);
                }
            }

            conn = string.Format(ConfigurationManager.AppSettings.Get(Constants.AppSetting.ExcelConStrDefineHeader), fileName);
            using (connection = new OleDbConnection(conn))
            {
                string queryBody = "select * from [TestIterations$A5:F50]";
                using (adapter = new OleDbDataAdapter(queryBody, connection))
                {
                    dataSourceBody = new DataTable();
                    adapter.Fill(dataSourceBody);
                }
            }
        }

        /// <summary>
        /// Prepares the body of the mail.
        /// </summary>
        /// <param name="fileShortName">File Short Name as parameter.</param>
        /// /// <returns>String value.</returns>
        private string GetBody(string fileShortName)
        {
            DataTable dataSourceHeader;
            DataTable dataSourceBody;
            this.Query(fileShortName, out dataSourceHeader, out dataSourceBody);

            StringBuilder sb = new StringBuilder();

            sb.AppendLine(
@"<html>
    <body style=""font-family:Calibri"">
    <p>Hi All,</p>
    <p>The following is the running results of CCPR.</p>
    <table border=1 cellspacing=0 width=50% bordercolorlight=#333333  bordercolordark=#efefef>");
            
            foreach (DataRow row in dataSourceHeader.Rows)
            {
                if (row == null)
                {
                    continue;
                }   

                sb.AppendLine(string.Format(
@"<tr>
    <td bgcolor=#ACD6FF>{0} </td>
    <td>{0}</td>
    </tr>",
                                row[0], row[1]));
            }

            sb.AppendLine("</table>");

            var subContent = new StringBuilder();
            subContent.AppendLine(
@"<table border=1 cellspacing=0 width=80% bordercolorlight=#333333  bordercolordark=#efefef>
    <tr><th bgcolor=#ACD6FF width=10%>Application</th>
        <th bgcolor=#ACD6FF width=20%>Test case name</th>
        <th bgcolor=#ACD6FF width=20%>Test case description</th>
        <th bgcolor=#ACD6FF width=10%>Result</th>
        <th bgcolor=#ACD6FF width=10%>Duration</th>
        <th bgcolor=#ACD6FF width=10%>DocumentReference</th>
    </tr>");

            foreach (DataRow row in dataSourceBody.Rows)
            {
                if (row == null)
                {
                    continue;
                }   

                subContent.AppendLine(
                    string.Format(
@"<tr>
    <td>{0}</td>
    <td>{1}</td>
    <td>{2}</td>
    <td>{3}</td>
    <td>{4}</td>
    <td>{5}</td>
</tr>",
                    row[0], row[1], row[2], row[3], row[4], row[5]));
            }

            subContent.AppendLine("</table>");

            sb.AppendLine(string.Format("<br/><div>{0}</div>", subContent.ToString()));

            sb.AppendLine("<p>This is an auto-generated email, please DO NOT REPLY!</p>");
            sb.AppendLine("<p>Thanks!</p>");

            sb.AppendLine("</body>");
            sb.AppendLine("</html>");

            return sb.ToString();
        }
    }
}
