// <copyright file="WebDriverBrowser.cs" company="Infosys Ltd.">
//  Copyright (c) Infosys Ltd. All rights reserved.
// </copyright>
// <summary>WebDriverBrowser.cs class handles Selenium web driver browsers.</summary>
namespace INF.Selenium.TestAutomation.Utilities
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Reflection;
    using Configuration;
    using Entities;
    using Microsoft.VisualStudio.TestTools.UITesting;
    using OpenQA.Selenium;
    using OpenQA.Selenium.Chrome;
    using OpenQA.Selenium.Firefox;
    using OpenQA.Selenium.IE;

    /// <summary>WebDriverBrowser handles Selenium web driver browsers.</summary>
    public class WebdriverBrowser
    {
        /// <summary>Title of application.</summary>
        private static readonly Dictionary<string, string> Titles = new Dictionary<string, string>();

        /// <summary>Gets or private sets Web driver.</summary>
        /// <value>Web driver status.</value>
        public static IWebDriver Driver { get; private set; }

        /// <summary>
        /// Launch Web driver Browser.
        ///  </summary> 
        /// <param name="urlString">URL string.</param>
        public static void Launch(string urlString)
        {
            try
            {
                if (General.BrowserType.ToUpper() == Constants.Browsers.CHrome)
                {
                    var driverPath = TestCase.RootFilePath + Constants.ChromeDriverPath;
                    var options = new ChromeOptions();
                    options.AddArgument("--start-maximized");
                    Driver = new ChromeDriver(driverPath, options);
                }

                if (General.BrowserType.ToUpper() == Constants.Browsers.Ie)
                {
                    var driverPath = TestCase.RootFilePath + Constants.IEDriverPath;
                    var options = new InternetExplorerOptions
                    {
                        InitialBrowserUrl = Constants.Url,
                        IntroduceInstabilityByIgnoringProtectedModeSettings = true
                    };

                    Driver = new InternetExplorerDriver(driverPath, options);
                    Driver.Manage().Window.Maximize();
                }

                if (General.BrowserType.ToUpper() == Constants.Browsers.FireFox)
                {
                    Driver = new FirefoxDriver();
                    Driver.Manage().Window.Maximize();
                }

                Driver.Navigate().GoToUrl(urlString);
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.WebdriverBrowser, MethodBase.GetCurrentMethod().Name);
            }
        }

        /// <summary>
        /// Locate a Web driver Browser with a specific title or part of a title.
        /// </summary>
        /// <param name="partoftitle">Part of title.</param>
        /// <param name="tryLocateInCashe">Try to locate application properties from cache.</param>
        /// <returns>True or False.</returns>
        public static string GetTitleFromPartOfTitle(string partoftitle, bool tryLocateInCashe = true)
        {
            try
            {
                // Check if we have search for partoftitle last time, if so return lastTitle instead of search once more -> just to speed up things
                if (tryLocateInCashe && Titles.ContainsKey(partoftitle))
                {
                    return Titles[partoftitle];
                }

                var bw = new BrowserWindow();
                bw.SearchProperties.Add(UITestControl.PropertyNames.Name, partoftitle, PropertyExpressionOperator.Contains);

                if (tryLocateInCashe)
                {
                    Titles.Add(partoftitle, bw.Name);
                }

                return bw.Name;
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.WebdriverBrowser, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }

        /// <summary>  
        /// Close all open Web driver Browser Processes that match given Web driver Browser Type in configuration file.
        /// </summary> 
        /// <returns>True or False.</returns>
        public bool CloseAllWebdriver_Browsers()
        {
            try
            {
                switch (General.BrowserType.ToUpper())
                {
                    case Constants.Browsers.Ie:
                        var processName = Constants.Browsers.Iexplore;
                        foreach (var window in Process.GetProcessesByName(processName))
                        {
                            window.Kill();
                        }

                        break;
                    case Constants.Browsers.FireFox:
                        processName = Constants.Browsers.Firefox;
                        foreach (var window in Process.GetProcessesByName(processName))
                        {
                            window.Kill();
                        }

                        break;
                    case Constants.Browsers.CHrome:
                        processName = Constants.Browsers.Chrome;
                        foreach (var window in Process.GetProcessesByName(processName))
                        {
                            window.Kill();
                        }

                        break;
                    default:
                        return false;
                }

                var b = Driver;
                if (b != null)
                {
                    Driver.Quit();
                }

                return true;
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.WebdriverBrowser, MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }
    }
}
