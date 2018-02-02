using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using System.Data.OleDb;
using MySql.Data.MySqlClient;

namespace DataConnectorManager
{
    public class DatabaseConnectionParameters
    {
        public DataConnectionType ConnectionType;
        public string ConnectionString;

        /// <summary>
        /// Connection Timeout Limit.
        /// NOT Working with OleDBConnection
        /// </summary>
        public int ConnectionTimeout = 5;
        /// <summary>
        /// Command Timeout Limit.
        /// </summary>
        public int CommandTimeout = 30;

        #region Query
        public DataContainerType    DataContainerType;
        public CommandBuildType            CommandBuildType;
        public CommandType          CommandType;
        public ICollection<object>  QueryParameters;
        public string               QueryString;
        public bool                 LastCommandSucceeded;
        public DataSet              DataSetContainer;
        public DataTable            DataTableContainer;
        public DataRow[]            DataRowsCollectionContainer;
        public string               DataTableContainerName;
        #endregion

        #region Connection Parameters

        #region OleDB
        public string   FilePath;
        public string   Provider;
        public bool     PersistSecurityInfo;
        #endregion

        #region SQL
        public string   Server;
        public string   Database;
        public string   UserId;
        public string   Password;
        public bool     TrustedConnection;

        public int      Port;
        public string   NetworkLibrary;

        public bool MultipleActiveResultSets;
        #endregion

        /// <summary>
        /// Reset all Connection Parameters
        /// </summary>
        public void ResetParameters()
        {
            FilePath        = "";
            Provider        = "";
            Server          = "";
            Database        = "";
            UserId          = "";
            Password        = "";
            NetworkLibrary  = "";
            MultipleActiveResultSets = false;
            TrustedConnection        = false;
            Port = 0;
            ConnectionTimeout = 5;
        }
        #endregion

        #region Connectors
        public SqlConnection SQLConnection;
        public OleDbConnection OLEDBConnection;
        public MySqlConnection MySQLConnection;
        #endregion
    }
}
