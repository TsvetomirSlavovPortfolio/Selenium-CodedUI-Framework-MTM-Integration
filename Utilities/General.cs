// <copyright file="General.cs" company="Infosys Ltd.">
//  Copyright (c) Infosys Ltd. All rights reserved.
// </copyright>
// <summary>General.cs class enables run time settings for Selenium and Browser.</summary>
namespace INF.Selenium.TestAutomation.Utilities
{
    using System;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using Configuration;
    using Entities;

    /// <summary>
    /// General.cs class enables run time settings for Selenium and Browser.
    /// </summary>
    public class General
    {
        /// <summary>
        /// Gets or sets Wait For Control To Exist TimeOut.
        /// </summary>
        /// <value>Time in milliseconds.</value>
        public static string WaitForControlToExistTimeOut { get; set; }

        /// <summary>
        /// Gets or sets browser type.
        /// </summary>
        /// <value>Browser type.</value>
        public static string BrowserType { get; set; }
        
        /// <summary>
        /// Releases objects.
        /// </summary>
        /// <param name="obj">Object to be released.</param>
        public void ReleaseObject(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.General, MethodBase.GetCurrentMethod().Name);
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
