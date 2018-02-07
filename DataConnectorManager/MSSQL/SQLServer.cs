using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;

namespace DataConnectorManager
{
    class SQLServer
    {
        /// <summary>
        /// Connect to SQL Server using given parameters
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <returns>Returns TRUE is connection succeeds. Returns FALSE if connection fails.</returns>
        public static bool ConnectToDatabase(DatabaseConnectionParameters dbParameters)
        {
            try
            {
                dbParameters.SQLConnection = new System.Data.SqlClient.SqlConnection(dbParameters.ConnectionString);

                if (dbParameters.SQLConnection.State == System.Data.ConnectionState.Closed)
                    dbParameters.SQLConnection.Open();

                dbParameters.LastCommandSucceeded = (dbParameters.SQLConnection.State == System.Data.ConnectionState.Open);
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
        /// Execute query stored in DatabaseConnectionParameters.Query
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <returns>Returns IDataReader if command succeeded. Returns NULL if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public static IDataReader ExecuteReader(DatabaseConnectionParameters dbParameters)
        {
            try
            {
                var sqlCommand = new SqlCommand(dbParameters.QueryString, dbParameters.SQLConnection);

                if (dbParameters.QueryParameters != null)
                {
                    foreach (SqlParameter item in dbParameters.QueryParameters)
                    {
                        sqlCommand.Parameters.Add(item);
                    }
                }

                sqlCommand.CommandTimeout = dbParameters.CommandTimeout;
                sqlCommand.CommandType = dbParameters.CommandType;
                dbParameters.LastCommandSucceeded = true;
                return sqlCommand.ExecuteReader();

            }
            catch(Exception excp)
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
                var sqlCommand = new SqlCommand(dbParameters.QueryString, dbParameters.SQLConnection);

                if (dbParameters.QueryParameters != null)
                {
                    foreach (SqlParameter item in dbParameters.QueryParameters)
                    {
                        sqlCommand.Parameters.Add(item);
                    }
                }

                sqlCommand.CommandTimeout = dbParameters.CommandTimeout;
                sqlCommand.CommandType = dbParameters.CommandType;
                dbParameters.LastCommandSucceeded = true;
                return sqlCommand.ExecuteNonQuery();
                
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
                var sqlCommand = new SqlCommand(dbParameters.QueryString, dbParameters.SQLConnection);

                if (dbParameters.QueryParameters != null)
                {
                    foreach (SqlParameter item in dbParameters.QueryParameters)
                    {
                        sqlCommand.Parameters.Add(item);
                    }
                }
        
                sqlCommand.CommandTimeout = dbParameters.CommandTimeout;
                sqlCommand.CommandType = dbParameters.CommandType;
                dbParameters.LastCommandSucceeded = true;
                return sqlCommand.ExecuteScalar();

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
                var sqlCommand = new SqlCommand();
                sqlCommand.CommandText  = dbParameters.QueryString;
                sqlCommand.Connection   = dbParameters.SQLConnection;
                sqlCommand.CommandType  = dbParameters.CommandType;
                sqlCommand.CommandTimeout = dbParameters.CommandTimeout;

                if (dbParameters.QueryParameters != null)
                {
                    foreach (SqlParameter item in dbParameters.QueryParameters)
                    {
                        sqlCommand.Parameters.Add(item);
                    }
                }

                var sqlAdapter = new SqlDataAdapter();

                switch (dbParameters.CommandBuildType)
                {
                    case CommandBuildType.Insert:
                        sqlAdapter.InsertCommand = sqlCommand;
                        break;
                    case CommandBuildType.Update:
                        sqlAdapter.UpdateCommand = sqlCommand;
                        break;
                    case CommandBuildType.Delete:
                        sqlAdapter.DeleteCommand = sqlCommand;
                        break;
                    default:
                        throw new Exception("Build Type missing");
                }

                dbParameters.LastCommandSucceeded = true;

                switch (dbParameters.DataContainerType)
                {
                    case DataContainerType.DataTable:
                        return sqlAdapter.Update(dbParameters.DataTableContainer);
                    case DataContainerType.DataSet:
                        return sqlAdapter.Update(dbParameters.DataSetContainer);
                    case DataContainerType.DataSetWithTable:
                        return sqlAdapter.Update(dbParameters.DataSetContainer, dbParameters.DataTableContainerName);
                    case DataContainerType.DataRowsCollection:
                        return sqlAdapter.Update(dbParameters.DataRowsCollectionContainer);
                    default:
                        throw new Exception("Data Container Type missing");
                }
            }
            catch(Exception excp)
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
                return (dbParameters.SQLConnection != null && dbParameters.SQLConnection.State == ConnectionState.Open);
            }
            catch (Exception excp)
            {
                Logs.AddException(excp);
                return false;
            }
        }
    }
}
