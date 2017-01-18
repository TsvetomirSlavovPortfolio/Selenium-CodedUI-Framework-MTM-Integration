// <copyright file="ConfigStep.cs" company="Infosys Ltd.">
//  Copyright (c) Infosys Ltd. All rights reserved.
// </copyright>
// <summary>ConfigStep.cs class handles Configuration step variables.</summary>
namespace INF.Selenium.TestAutomation.Entities
{
    using System.Collections;
    using System.Collections.Generic;

    /// <summary>
    /// ConfigStep reads steps from Configuration.
    /// </summary>
    public class ConfigStep : CollectionBase
    {
        /// <summary>
        /// Gets or sets Test data config count.
        /// </summary>
        /// <value>Test Step Number.</value>
        public static int TestDataConfigCount { get; set; }

        /// <summary>
        /// Gets or sets Test configuration names.
        /// </summary>
        /// <value>Test configuration names.</value>
        public static List<ConfigStep> TestConfigNames { get; set; }

        /// <summary>
        /// Gets or sets Test step number.
        /// </summary>
        /// <value>Test step number.</value>
        public string TestStepNo { get; set; }

        /// <summary>
        /// Gets or sets Test Data type.
        /// </summary>
        /// <value>Test Data type.</value>
        public string TestDataType { get; set; }

        /// <summary>
        /// Gets or sets Test Variable Name.
        /// </summary>
        /// <value>Test Variable Name.</value>
        public string TestVariableName { get; set; }

        /// <summary>
        /// Gets or sets Test Data Value.
        /// </summary>
        /// <value>Test Data Value.</value>
        public string TestDataValue { get; set; }

        /// <summary>
        /// Gets or sets Configuration Action.
        /// </summary>
        /// <value>Configuration Action.</value>
        public string ConfigAction { get; set; }

        /// <summary>
        /// Gets or sets Test Configuration Key To Use.
        /// </summary>
        /// <value>Test Configuration Key To Use.</value>
        public string TestConfigKeyToUse { get; set; }

        /// <summary>
        /// Gets or sets Remarks.
        /// </summary>
        /// <value>Remarks column.</value>
        public string Remarks { get; set; }
    }
}
