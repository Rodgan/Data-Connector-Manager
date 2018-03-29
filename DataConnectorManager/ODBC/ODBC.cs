using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Linq;
using System.Text;

namespace DataConnectorManager
{
    class ODBC
    {
        /// <summary>
        /// Connect to ODBC using given parameters
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <returns>Returns TRUE is connection succeeds. Returns FALSE if connection fails.</returns>
        public static bool ConnectToDatabase(DatabaseConnectionParameters dbParameters)
        {
            try
            {
                dbParameters.ODBCConnection = new OdbcConnection(dbParameters.ConnectionString);

                if (dbParameters.ODBCConnection.State == System.Data.ConnectionState.Closed)
                    dbParameters.ODBCConnection.Open();

                dbParameters.LastCommandSucceeded = (dbParameters.ODBCConnection.State == System.Data.ConnectionState.Open);

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
        /// Disconnect from ODBC
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <returns></returns>
        public static bool DisconnectFromDatabase(DatabaseConnectionParameters dbParameters)
        {
            try
            {
                dbParameters.ODBCConnection.Close();
                dbParameters.LastCommandSucceeded = (dbParameters.ODBCConnection.State == ConnectionState.Closed);
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
                var odbcCommand = new OdbcCommand(dbParameters.QueryString, dbParameters.ODBCConnection);

                if (dbParameters.QueryParameters != null)
                {
                    foreach (OdbcParameter item in dbParameters.QueryParameters)
                    {
                        odbcCommand.Parameters.Add(item);
                    }
                }

                odbcCommand.CommandTimeout = dbParameters.CommandTimeout;
                odbcCommand.CommandType = dbParameters.CommandType;
                dbParameters.LastCommandSucceeded = true;
                return odbcCommand.ExecuteReader();

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
                var odbcCommand = new OdbcCommand(dbParameters.QueryString, dbParameters.ODBCConnection);

                if (dbParameters.QueryParameters != null)
                {
                    foreach (OdbcParameter item in dbParameters.QueryParameters)
                    {
                        odbcCommand.Parameters.Add(item);
                    }
                }

                odbcCommand.CommandTimeout = dbParameters.CommandTimeout;
                odbcCommand.CommandType = dbParameters.CommandType;
                dbParameters.LastCommandSucceeded = true;
                return odbcCommand.ExecuteNonQuery();

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
                var odbcCommand = new OdbcCommand(dbParameters.QueryString, dbParameters.ODBCConnection);

                if (dbParameters.QueryParameters != null)
                {
                    foreach (OdbcParameter item in dbParameters.QueryParameters)
                    {
                        odbcCommand.Parameters.Add(item);
                    }
                }

                odbcCommand.CommandTimeout = dbParameters.CommandTimeout;
                odbcCommand.CommandType = dbParameters.CommandType;
                dbParameters.LastCommandSucceeded = true;
                return odbcCommand.ExecuteScalar();

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
                var odbcCommand = new OdbcCommand();
                odbcCommand.CommandText = dbParameters.QueryString;
                odbcCommand.Connection = dbParameters.ODBCConnection;
                odbcCommand.CommandType = dbParameters.CommandType;
                odbcCommand.CommandTimeout = dbParameters.CommandTimeout;

                if (dbParameters.QueryParameters != null)
                {
                    foreach (OdbcParameter item in dbParameters.QueryParameters)
                    {
                        odbcCommand.Parameters.Add(item);
                    }
                }

                var odbcAdapter = new OdbcDataAdapter();

                switch (dbParameters.CommandBuildType)
                {
                    case CommandBuildType.Insert:
                        odbcAdapter.InsertCommand = odbcCommand;
                        break;
                    case CommandBuildType.Update:
                        odbcAdapter.UpdateCommand = odbcCommand;
                        break;
                    case CommandBuildType.Delete:
                        odbcAdapter.DeleteCommand = odbcCommand;
                        break;
                    default:
                        throw new Exception("Build Type missing");
                }

                dbParameters.LastCommandSucceeded = true;

                switch (dbParameters.DataContainerType)
                {
                    case DataContainerType.DataTable:
                        return odbcAdapter.Update(dbParameters.DataTableContainer);
                    case DataContainerType.DataSet:
                        return odbcAdapter.Update(dbParameters.DataSetContainer);
                    case DataContainerType.DataSetWithTable:
                        return odbcAdapter.Update(dbParameters.DataSetContainer, dbParameters.DataTableContainerName);
                    case DataContainerType.DataRowsCollection:
                        return odbcAdapter.Update(dbParameters.DataRowsCollectionContainer);
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
        /// Provides an help to build a string that contains a Stored Procedure for ODBC connections, since ODBC's Stored Procedures are syntactically different from others
        /// </summary>
        /// <param name="storedProcedureName">Stored Procedure to execute</param>
        /// <param name="storedProcedureParameters">Stored Procedure Parameters</param>
        /// <returns>Returns a string that contains a Stored Procedure that can be executed with ODBC connections</returns>
        public static string BuildStoredProcedure(string storedProcedureName, IEnumerable<object> storedProcedureParameters)
        {
            // Standard Stored Procedure: {CALL spName}
            // SP with single/multiple parameters: {CALL spName (?)} / {CALL spName (?,?,?)}
            // SP with Return Value with/without parameters: {? = CALL spName} / {? = CALL spName(?,?,?)}

            List<OdbcParameter> spParameters = new List<OdbcParameter>();

            if (storedProcedureParameters != null)
            {
                foreach (OdbcParameter item in storedProcedureParameters)
                {
                    spParameters.Add(item);
                }
            }

            if (spParameters.Count() == 0)
            {
                return $"{{ CALL {storedProcedureName} }}";
            }

            var spText = "{";
            var spPlaceHoldersStartFrom = 0;

            if (spParameters[0].Direction == ParameterDirection.ReturnValue)
            {
                spText += "? = ";
                spPlaceHoldersStartFrom = 1;
            }
               
            var spParametersPlaceHolders = "";

            // placeholder = "?"
            for ( var placeHolder = spPlaceHoldersStartFrom; placeHolder < spParameters.Count(); placeHolder++)
            {
                if (placeHolder == spParameters.Count() - 1)
                    spParametersPlaceHolders += "?";
                else
                    spParametersPlaceHolders += "?,";
            }

            spText += $"CALL {storedProcedureName} ({spParametersPlaceHolders})}}";

            return spText;
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
                return (dbParameters.ODBCConnection != null && dbParameters.ODBCConnection.State == ConnectionState.Open);
            }
            catch (Exception excp)
            {
                Logs.AddException(excp);
                return false;
            }
        }
    }
}
