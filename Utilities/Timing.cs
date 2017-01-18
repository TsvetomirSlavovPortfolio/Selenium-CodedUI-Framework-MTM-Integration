// <copyright file="Timing.cs" company="Infosys Ltd.">
//  Copyright (c) Infosys Ltd. All rights reserved.
// </copyright>
// <summary>Timing.cs Start and End time of each test case execution</summary>
namespace INF.Selenium.TestAutomation.Utilities
{
    using System;

    /// <summary>
    /// Set up time.
    /// </summary>
    public static class Timing
    {
        /// <summary>
        /// Gets or sets up time for its start time.
        /// </summary>
        /// <value>Time for test starts.</value>
        public static DateTime TotalStartTime { get; set; }

        /// <summary>
        /// Gets or sets up time for its total end time.
        /// </summary>
        /// <value>Total end Time.</value>
        public static DateTime TotalEndTime { get; set; }

        /// <summary>
        /// Gets or sets Test case start time.
        /// </summary>
        /// <value>Test case start time.</value>
        public static DateTime TestCaseStartTime { get; set; }

        /// <summary>
        /// Gets or sets Duration of test.
        /// </summary>
        /// <value>Test case duration.</value>
        public static TimeSpan TestCaseDuration { get; set; }

        /// <summary>
        /// Gets or sets Total duration.
        /// </summary>
        /// <value>Total duration of all tests.</value>
        public static TimeSpan Totalduration { get; set; }
    }
}