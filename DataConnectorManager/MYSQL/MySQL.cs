using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MySql.Data.MySqlClient;
using System.Data;

namespace DataConnectorManager
{
    class MySQL
    {
        /// <summary>
        /// Connect to MySQL Server using given parameters
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <returns>Returns TRUE is connection succeeds. Returns FALSE if connection fails.</returns>
        public static bool ConnectToDatabase(DatabaseConnectionParameters dbParameters)
        {
            try
            {
                dbParameters.MySQLConnection = new MySqlConnection(dbParameters.ConnectionString);

                if (dbParameters.MySQLConnection.State == System.Data.ConnectionState.Closed)
                    dbParameters.MySQLConnection.Open();

                dbParameters.LastCommandSucceeded = (dbParameters.MySQLConnection.State == System.Data.ConnectionState.Open);
                return dbParameters.LastCommandSucceeded;
            }
            catch(Exception excp)
            {
                Logs.AddException(excp);
                dbParameters.LastCommandSucceeded = false;
                return false;
            }
        }

        /// <summary>
        /// Disconnect from MySQL Server
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <returns></returns>
        public static bool DisconnectFromDatabase(DatabaseConnectionParameters dbParameters)
        {
            try
            {
                dbParameters.MySQLConnection.Close();
                dbParameters.LastCommandSucceeded = (dbParameters.MySQLConnection.State == ConnectionState.Closed);
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
                var mySqlCommand = new MySqlCommand(dbParameters.QueryString, dbParameters.MySQLConnection);

                if (dbParameters.QueryParameters != null)
                {
                    foreach (MySqlParameter item in dbParameters.QueryParameters)
                    {
                        mySqlCommand.Parameters.Add(item);
                    }
                }
                mySqlCommand.CommandTimeout = dbParameters.CommandTimeout;
                mySqlCommand.CommandType = dbParameters.CommandType;
                dbParameters.LastCommandSucceeded = true;
                return mySqlCommand.ExecuteReader();
                
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
                var mySqlCommand = new MySqlCommand(dbParameters.QueryString, dbParameters.MySQLConnection);

                if (dbParameters.QueryParameters != null)
                {
                    foreach (MySqlParameter item in dbParameters.QueryParameters)
                    {
                        mySqlCommand.Parameters.Add(item);
                    }
                }
                mySqlCommand.CommandTimeout = dbParameters.CommandTimeout;
                mySqlCommand.CommandType = dbParameters.CommandType;
                dbParameters.LastCommandSucceeded = true;
                return mySqlCommand.ExecuteNonQuery();

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
                var mySqlCommand = new MySqlCommand(dbParameters.QueryString, dbParameters.MySQLConnection);


                if (dbParameters.QueryParameters != null)
                {
                    foreach (MySqlParameter item in dbParameters.QueryParameters)
                    {
                        mySqlCommand.Parameters.Add(item);
                    }
                }

                mySqlCommand.CommandTimeout = dbParameters.CommandTimeout;
                mySqlCommand.CommandType = dbParameters.CommandType;
                dbParameters.LastCommandSucceeded = true;
                return mySqlCommand.ExecuteScalar();

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
                var mySqlCommand = new MySqlCommand();
                mySqlCommand.CommandText = dbParameters.QueryString;
                mySqlCommand.Connection = dbParameters.MySQLConnection;
                mySqlCommand.CommandType = dbParameters.CommandType;
                mySqlCommand.CommandTimeout = dbParameters.CommandTimeout;

                if (dbParameters.QueryParameters != null)
                {
                    foreach (MySqlParameter item in dbParameters.QueryParameters)
                    {
                        mySqlCommand.Parameters.Add(item);
                    }
                }

                var mySqlAdapter = new MySqlDataAdapter();

                switch (dbParameters.CommandBuildType)
                {
                    case CommandBuildType.Insert:
                        mySqlAdapter.InsertCommand = mySqlCommand;
                        break;
                    case CommandBuildType.Update:
                        mySqlAdapter.UpdateCommand = mySqlCommand;
                        break;
                    case CommandBuildType.Delete:
                        mySqlAdapter.DeleteCommand = mySqlCommand;
                        break;
                    default:
                        throw new Exception("Build Type missing");
                }

                dbParameters.LastCommandSucceeded = true;

                switch (dbParameters.DataContainerType)
                {
                    case DataContainerType.DataTable:
                        return mySqlAdapter.Update(dbParameters.DataTableContainer);
                    case DataContainerType.DataSet:
                        return mySqlAdapter.Update(dbParameters.DataSetContainer);
                    case DataContainerType.DataSetWithTable:
                        return mySqlAdapter.Update(dbParameters.DataSetContainer, dbParameters.DataTableContainerName);
                    case DataContainerType.DataRowsCollection:
                        return mySqlAdapter.Update(dbParameters.DataRowsCollectionContainer);
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
                return (dbParameters.MySQLConnection != null && dbParameters.MySQLConnection.State == ConnectionState.Open);
            }
            catch(Exception excp)
            {
                Logs.AddException(excp);
                return false;
            }
        }
    }
}
