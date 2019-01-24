using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace DataConnectorManager
{
    class OleDb
    {
        /// <summary>
        /// Connect to OleDB using given parameters
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <returns>Returns TRUE is connection succeeds. Returns FALSE if connection fails.</returns>
        public static bool ConnectToDatabase(DatabaseConnectionParameters dbParameters)
        {
            try
            {
                dbParameters.OLEDBConnection = new OleDbConnection(dbParameters.ConnectionString);

                if (dbParameters.OLEDBConnection.State == System.Data.ConnectionState.Closed)
                    dbParameters.OLEDBConnection.Open();

                dbParameters.LastCommandSucceeded = (dbParameters.OLEDBConnection.State == System.Data.ConnectionState.Open);

                return dbParameters.LastCommandSucceeded;
            }
            catch(Exception excp)
            {
                Logs.AddException(excp);
                dbParameters.LastCommandSucceeded = false;

                return dbParameters.LastCommandSucceeded;
            }
        }

        /// <summary>
        /// Disconnect from OleDB
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <returns></returns>
        public static bool DisconnectFromDatabase(DatabaseConnectionParameters dbParameters)
        {
            try
            {
                dbParameters.OLEDBConnection.Close();
                dbParameters.LastCommandSucceeded = (dbParameters.OLEDBConnection.State == ConnectionState.Closed);
                return dbParameters.LastCommandSucceeded;
            }
            catch (Exception excp)
            {
                Logs.AddException(excp);
                dbParameters.LastCommandSucceeded = false;
                return dbParameters.LastCommandSucceeded;
            }
        }

        /// <summary>
        /// Execute query stored in DatabaseConnectionParameters.Query
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <returns>Returns IDataReader if command succeeded. Returns NULL if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public static IDataReader ExecuteReader(DatabaseConnectionParameters dbParameters)
        {
            try
            {
                var oleDbCommand = new OleDbCommand(dbParameters.QueryString, dbParameters.OLEDBConnection);

                if (dbParameters.QueryParameters != null)
                {
                    foreach (OleDbParameter item in dbParameters.QueryParameters)
                    {
                        oleDbCommand.Parameters.Add(item);
                    }
                }

                oleDbCommand.CommandTimeout = dbParameters.CommandTimeout;
                oleDbCommand.CommandType = dbParameters.CommandType;
                dbParameters.LastCommandSucceeded = true;
                return oleDbCommand.ExecuteReader();

            }
            catch (Exception excp)
            {
                Logs.AddException(excp);
                dbParameters.LastCommandSucceeded = false;

                return null;
            }
        }

        /// <summary>
        /// Execute query stored in DatabaseConnectionParameters.Query
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public static int ExecuteNonQuery(DatabaseConnectionParameters dbParameters)
        {
            try
            {
                var oleDbCommand = new OleDbCommand(dbParameters.QueryString, dbParameters.OLEDBConnection);

                if (dbParameters.QueryParameters != null)
                {
                    foreach (OleDbParameter item in dbParameters.QueryParameters)
                    {
                        oleDbCommand.Parameters.Add(item);
                    }
                }

                oleDbCommand.CommandTimeout = dbParameters.CommandTimeout;
                oleDbCommand.CommandType = dbParameters.CommandType;
                dbParameters.LastCommandSucceeded = true;
                return oleDbCommand.ExecuteNonQuery();

            }
            catch (Exception excp)
            {
                Logs.AddException(excp);
                dbParameters.LastCommandSucceeded = false;

                return -1;
            }
        }

        /// <summary>
        /// Execute query stored in DatabaseConnectionParameters.Query
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <returns>Returns a single value as object if command succeeds. May return NULL whether command fails or not. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public static object ExecuteScalar(DatabaseConnectionParameters dbParameters)
        {
            try
            {
                var oleDbCommand = new OleDbCommand(dbParameters.QueryString, dbParameters.OLEDBConnection);

                if (dbParameters.QueryParameters != null)
                {
                    foreach (OleDbParameter item in dbParameters.QueryParameters)
                    {
                        oleDbCommand.Parameters.Add(item);
                    }
                }

                oleDbCommand.CommandTimeout = dbParameters.CommandTimeout;
                oleDbCommand.CommandType = dbParameters.CommandType;
                dbParameters.LastCommandSucceeded = true;
                return oleDbCommand.ExecuteScalar();

            }
            catch (Exception excp)
            {
                Logs.AddException(excp);
                dbParameters.LastCommandSucceeded = false;

                return null;
            }
        }

        /// <summary>
        /// Build INSERT, UPDATE, DELETE queries
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public static int BuildCommand(DatabaseConnectionParameters dbParameters)
        {
            try
            {
                var oleDbCommand = new OleDbCommand();
                oleDbCommand.CommandText = dbParameters.QueryString;
                oleDbCommand.Connection = dbParameters.OLEDBConnection;
                oleDbCommand.CommandType = dbParameters.CommandType;
                oleDbCommand.CommandTimeout = dbParameters.CommandTimeout;

                if (dbParameters.QueryParameters != null)
                {
                    foreach (OleDbParameter item in dbParameters.QueryParameters)
                    {
                        oleDbCommand.Parameters.Add(item);
                    }
                }

                var oleDbAdapter = new OleDbDataAdapter();

                switch (dbParameters.CommandBuildType)
                {
                    case CommandBuildType.Insert:
                        oleDbAdapter.InsertCommand = oleDbCommand;
                        break;
                    case CommandBuildType.Update:
                        oleDbAdapter.UpdateCommand = oleDbCommand;
                        break;
                    case CommandBuildType.Delete:
                        oleDbAdapter.DeleteCommand = oleDbCommand;
                        break;
                    default:
                        throw new Exception("Build Type missing");
                }

                dbParameters.LastCommandSucceeded = true;

                switch (dbParameters.DataContainerType)
                {
                    case DataContainerType.DataTable:
                        return oleDbAdapter.Update(dbParameters.DataTableContainer);
                    case DataContainerType.DataSet:
                        return oleDbAdapter.Update(dbParameters.DataSetContainer);
                    case DataContainerType.DataSetWithTable:
                        return oleDbAdapter.Update(dbParameters.DataSetContainer, dbParameters.DataTableContainerName);
                    case DataContainerType.DataRowsCollection:
                        return oleDbAdapter.Update(dbParameters.DataRowsCollectionContainer);
                    default:
                        throw new Exception("Data Container Type missing");
                }
            }
            catch (Exception excp)
            {
                Logs.AddException(excp);
                dbParameters.LastCommandSucceeded = false;

                return -1;
            }

        }

        /// <summary>
        /// Check if connection is open
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <returns>Returns TRUE if connection is open</returns>
        public static bool IsOpen(DatabaseConnectionParameters dbParameters)
        {
            try
            {
                return (dbParameters.OLEDBConnection != null && dbParameters.OLEDBConnection.State == ConnectionState.Open);
            }
            catch (Exception excp)
            {
                Logs.AddException(excp);
                return false;
            }
        }
    }

    class Access : OleDb { }
    class Excel : OleDb
    {
        public static string GetExtendedProperties(string fileName, bool IMEX)
        {
            string extendedProperties = "";

            if (fileName.EndsWith(".xls"))
                extendedProperties += "Excel 8.0;";
            else
                extendedProperties += "Excel 12.0;";

            extendedProperties += "HDR=YES;";
            extendedProperties += "IMEX=" + (IMEX ? 1 : 0);

            return extendedProperties;
        }
    }
}
