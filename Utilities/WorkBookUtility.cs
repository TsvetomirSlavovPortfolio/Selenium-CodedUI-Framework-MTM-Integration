// <copyright file="WorkBookUtility.cs" company="Infosys Ltd.">
//  Copyright (c) Infosys Ltd. All rights reserved.
// </copyright>
// <summary>WorkBookUtility.cs Work book related handles.</summary>
namespace INF.Selenium.TestAutomation.Utilities
{
    using System;
    using Microsoft.Office.Interop.Excel;

    /// <summary>
    /// Work book utility.
    /// </summary>
    internal static class WorkBookUtility
    {
        /// <summary>
        /// Open work book.
        /// </summary>
        /// <param name="applicationClass">Read application class.</param>
        /// <param name="fileName">File name.</param>
        /// <returns>Work book.</returns>
        public static Workbook OpenWorkBook(ApplicationClass applicationClass, string fileName)
        {
            var workbook = applicationClass.Workbooks.Open(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            return workbook;
        }

        /// <summary>
        /// Close the work book.
        /// </summary>
        /// <param name="workbook">Work book name.</param>
        /// <param name="save">Save the work book before close.</param>
        public static void CloseWorkBook(Workbook workbook, bool save = false)
        {
            // Close workbook
            workbook.Close(save);
            General objGeneral = new General();
            objGeneral.ReleaseObject(workbook);
        }

        /// <summary>
        /// Close Excel.
        /// </summary>
        /// <param name="applicationClass">Application Class.</param>
        public static void CloseExcel(ApplicationClass applicationClass)
        {
            applicationClass.Quit();
            General objGeneral = new General();
            objGeneral.ReleaseObject(applicationClass);
        }
    }
}