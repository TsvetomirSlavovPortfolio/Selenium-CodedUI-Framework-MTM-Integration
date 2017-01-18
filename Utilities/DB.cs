// <copyright file="DB.cs" company="Infosys Ltd.">
//  Copyright (c) Infosys Ltd. All rights reserved.
// </copyright>
// <summary>DB.cs class helps framework to interact with data bases.</summary>
namespace INF.Selenium.TestAutomation.Utilities
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Data.Odbc;
    using System.Diagnostics.CodeAnalysis;
    using System.Linq;
    using Configuration;
    using Entities;

    /// <summary>
    /// This class enables framework to interact with data base.
    /// </summary>
    public class Db
    {
        /// <summary>
        ///  Method runs a database query using a connection to database server and database name given in app.config file and returns the value.
        /// </summary>
        /// <param name="query">Query to be passed to data base.</param>
        /// <param name="operatorUse">Operator to be used in query.</param>
        /// <returns>Returns data base result.</returns>
        [SuppressMessage("Microsoft.Security", "CA2100:Review SQL queries for security vulnerabilities", Justification = "Used during unit testing")]
        public static string ExecuteQuery(string query, string operatorUse)
        {
            var result = string.Empty;

            try
            {
                //// Run database query
                var numberOfRows = 0;
                var connection = new OdbcConnection { ConnectionString = "DSN=" + ConfigurationManager.AppSettings.Get(Constants.AppSetting.DsNName) + ";" };
                connection.Open();
                using (var sql = new OdbcCommand(query, connection))
                {
                    switch (operatorUse.ToUpper())
                    {
                        case Constants.DbActions.Insert:
                        case Constants.DbActions.Update:
                        case Constants.DbActions.Delete:
                        case Constants.DbActions.Call:
                            sql.ExecuteReader();
                            break;
                        case Constants.DbActions.IsEquals:
                        case Constants.DbActions.NotEquals:
                            using (var dr = sql.ExecuteReader())
                            {
                                if (dr.FieldCount > 1)
                                {
                                    throw new Exception("Database query must return only 1 row.");
                                }

                                if (dr.HasRows)
                                {
                                    while (dr.Read())
                                    {
                                        if (numberOfRows > 1)
                                        {
                                            throw new Exception("Database query must return only 1 row.");
                                        }

                                        result = dr.GetString(0);
                                        numberOfRows++;
                                    }
                                }
                            }

                            break;
                        default:
                            throw new Exception(string.Concat("Operator ", operatorUse, " is not supported"));
                    }
                }

                connection.Close();
                return result;
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.Db, System.Reflection.MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }

        /// <summary>
        /// Method validate a database query.
        /// </summary>
        /// <param name="query">Query to be passed to data base.</param>
        public void ValidateQuery(string query)
        {
            //// Validation is done according to SQL injection (http:////msdn.microsoft.com/en-us/library/ms161953(SQL.105).aspx) plus some more
            var injectionEntries = new List<string> { Constants.Semicolon, Constants.DoubleHyphen, Constants.BeginCommnet, Constants.EndComment, Constants.Xp, Constants.Cursor, Constants.Exec, Constants.Drop, Constants.Declare };

            try
            {
                if (injectionEntries.Any(item => query.ToUpper().Contains(item)))
                {
                    throw new Exception("Database query must not contain any of these strings [ ;,--,/*,*/,XP_,CURSOR,EXEC,DROP,DECLARE]");
                }
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.Db, System.Reflection.MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }
    }
}