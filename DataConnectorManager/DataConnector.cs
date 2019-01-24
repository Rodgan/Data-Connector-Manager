using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataConnectorManager
{
    /// <summary>
    /// Connection Type. Each connection type needs parameters to be set
    /// </summary>
    public enum DataConnectionType
    {
        /// <summary>
        /// Standard Connection to SQL Server. Parameters to be set: Server, Database, UserId, Password. To specify an instance use the following syntax: Server\instanceName
        /// </summary>
        SQLServer_StandardSecurity,
        /// <summary>
        /// Trusted Connection to SQL Server. Parameters to be set: Server, Database, [TrustedConnection = TRUE]
        /// </summary>
        SQLServer_TrustedConnection,
        /// <summary>
        /// Standard Connection to SQL Server. Parameters to be set: Server, Port, Database, UserId, Password, [NetworkCatalog = DBMSSOCN]
        /// </summary>
        SQLServer_StandardSecurity_UseIpAddressAndPort,

        /// <summary>
        /// Standard Connection to Access using ACE.OLEDB.12.0 as Provider. Parameters to be set: FilePath, [Provider = Microsoft.ACE.OLEDB.12.0], [PersistSecurityInfo = FALSE]
        /// </summary>
        Access_ACE_OLEDB12_StandardSecurity,
        /// <summary>
        /// Encrypted Connection to Access using ACE.OLEDB.12.0 as Provider. Parameters to be set: FilePath, Password, [Provider = Microsoft.ACE.OLEDB.12.0].
        /// NOTE: Works with LEGACY ENCRYPTION METHOD only.
        /// </summary>
        Access_ACE_OLEDB12_WithPassword,

        /// <summary>
        /// Standard Connection to Access using JET.OLEDB.4.0 as Provider. Parameters to be set: FilePath, [Provider = Microsoft.JET.OLEDB.4.0], [UserId = admin], [Password = NULL]
        /// </summary>
        Access_JET_OLEDB4_StandardSecurity,
        /// <summary>
        /// Encrypted Connection to Access using JET.OLEDB.4.0 as Provider. Parameters to be set: FilePath, Password, [Provider = Microsoft.JET.OLEDB.4.0]
        /// NOTE: Works with LEGACY ENCRYPTION METHOD AND DEFAULT ENCRYPTION METHOD.
        /// </summary>
        Access_JET_OLEDB4_WithPassword,

        Excel_Ace_OLEDB12,
        Excel_Ace_OLEDB12_With_IMEX,

        Excel_Jet_OLEDB4,
        Excel_Jet_OLEDB4_With_IMEX,

        /// <summary>
        /// Standard Connection to MySQL. Parameters to be set: Server, Database, UserId, Password
        /// </summary>
        MySQL_StandardConnection,
        /// <summary>
        /// Standard Connection to MySQL. Parameters to be set: Server, Port, Database, UserId, Password
        /// </summary>
        MySQL_ServerAndPortConnection,

        /// <summary>
        /// Trusted Connection to ODBC. Parameters to be set: DSN
        /// </summary>
        ODBC_TrustedConnection,
        /// <summary>
        /// Standard Connection with login to ODBC. Parameters to be set: DSN, UserId, Password
        /// </summary>
        ODBC_StandardLogin
    }
    /// <summary>
    /// Used to determine which command will be built
    /// </summary>
    public enum CommandBuildType
    {
        None,
        Insert,
        Update,
        Delete
    }
    /// <summary>
    /// Used to determine which type of container will be used to build commands
    /// </summary>
    public enum DataContainerType
    {
        None,
        DataTable,
        DataSet,
        DataSetWithTable,
        DataRowsCollection
    }

    /// <summary>
    /// Connect and run queries on SQL, Access and MySQL with a single command
    /// </summary>
    public class DataConnector
    {
        // REMOVE BEFORE PUBLISH
        // Methods to set up for each DataConnectionType
        private void Shortcuts()
        {
            SetConnectionString(DbStoredParameters, 0); // SQL (3/3) ACCESS (4/4) MYSQL (2/2) ODBC(2/2) EXCEL(2/2)
            ConnectToDatabase(DbStoredParameters);      // SQL (3/3) ACCESS (4/4) MYSQL (2/2) ODBC(2/2) EXCEL(2/2)
            IsOpen(DbStoredParameters);                 // SQL (3/3) ACCESS (4/4) MYSQL (2/2) ODBC(2/2) EXCEL(2/2)
            BuildCommand(DbStoredParameters);           // SQL (3/3) ACCESS (4/4) MYSQL (2/2) ODBC(2/2) EXCEL(2/2)
            ExecuteReader(DbStoredParameters);          // SQL (3/3) ACCESS (4/4) MYSQL (2/2) ODBC(2/2) EXCEL(2/2)
            ExecuteNonQuery(DbStoredParameters);        // SQL (3/3) ACCESS (4/4) MYSQL (2/2) ODBC(2/2) EXCEL(2/2)
            ExecuteScalar(DbStoredParameters);          // SQL (3/3) ACCESS (4/4) MYSQL (2/2) ODBC(2/2) EXCEL(2/2)
            DisconnectFromDatabase(DbStoredParameters); // SQL (3/3) ACCESS (4/4) MYSQL (2/2) ODBC(2/2) EXCEL(2/2)
        }
        // REMOVE BEFORE PUBLISH

        /// <summary>
        /// Stored DatabaseConnectionParameters. Can be used to call methods without passing DatabaseConnectionParameters
        /// </summary>
        private DatabaseConnectionParameters DbStoredParameters;

        /// <summary>
        /// If TRUE, Build Commands will be treated as Stored Procedures. If FALSE, Build Commands will be treated as Raw Queries. DEFAULT: False
        /// </summary>
        public bool ExecuteBuildCommandAsStoredProcedure = false;

        /// <summary>
        /// Setup Connection String in DatabaseConnectionParameters
        /// </summary>
        /// <param name="dbParameters">DatabaseConnectionParameters</param>
        /// <param name="dbConnectionType">Connection Type</param>
        /// <param name="saveDatabaseConnectionParameters">If TRUE DatabaseConnectionParameters will be stored in the class instance</param>
        /// <returns>Returns Connection String - It will also be stored in DatabaseConnectionParameters.ConnectionString</returns>
        public string SetConnectionString(DatabaseConnectionParameters dbParameters, DataConnectionType dbConnectionType, bool saveDatabaseConnectionParameters = false)
        {
            dbParameters.ConnectionType = dbConnectionType;
            string connectionString;

            switch (dbConnectionType)
            {
                case DataConnectionType.SQLServer_StandardSecurity:
                    connectionString = $"Server={dbParameters.Server};Database={dbParameters.Database};User Id={dbParameters.UserId};Password={dbParameters.Password};Connection Timeout={dbParameters.ConnectionTimeout};";
                    break;
                case DataConnectionType.SQLServer_TrustedConnection:
                    dbParameters.TrustedConnection = true;
                    connectionString = $"Server={dbParameters.Server};Database={dbParameters.Database};Trusted_Connection={dbParameters.TrustedConnection};Connection Timeout={dbParameters.ConnectionTimeout};";
                    break;
                case DataConnectionType.SQLServer_StandardSecurity_UseIpAddressAndPort:
                    dbParameters.NetworkLibrary = "DBMSSOCN";
                    connectionString = $"Data Source={dbParameters.Server},{dbParameters.Port};Network Library={dbParameters.NetworkLibrary};Initial Catalog={dbParameters.Database};User ID={dbParameters.UserId};Password = {dbParameters.Password};Connection Timeout={dbParameters.ConnectionTimeout};";
                    break;
                case DataConnectionType.Access_ACE_OLEDB12_StandardSecurity:
                    dbParameters.Provider = "Microsoft.ACE.OLEDB.12.0";
                    dbParameters.PersistSecurityInfo = false;
                    connectionString = $"Provider={dbParameters.Provider};Data Source={dbParameters.FilePath};Persist Security Info = {dbParameters.PersistSecurityInfo};";
                    break;
                case DataConnectionType.Access_ACE_OLEDB12_WithPassword:
                    dbParameters.Provider = "Microsoft.ACE.OLEDB.12.0";
                    connectionString = $"Provider={dbParameters.Provider};Data Source={dbParameters.FilePath};Jet OLEDB:Database Password={dbParameters.Password};";
                    break;
                case DataConnectionType.Access_JET_OLEDB4_StandardSecurity:
                    dbParameters.Provider = "Microsoft.Jet.OLEDB.4.0";
                    dbParameters.UserId = "admin";
                    dbParameters.Password = "";
                    connectionString = $"Provider={dbParameters.Provider};Data Source={dbParameters.FilePath};User Id={dbParameters.UserId};Password={dbParameters.Password};";
                    break;
                case DataConnectionType.Access_JET_OLEDB4_WithPassword:
                    dbParameters.Provider = "Microsoft.Jet.OLEDB.4.0";
                    connectionString = $"Provider={dbParameters.Provider};Data Source={dbParameters.FilePath};Jet OLEDB:Database Password={dbParameters.Password};";
                    break;
                case DataConnectionType.MySQL_StandardConnection:
                    connectionString = $"Server={dbParameters.Server};Database={dbParameters.Database};Uid={dbParameters.UserId};Pwd={dbParameters.Password};Connection Timeout={dbParameters.ConnectionTimeout};";
                    break;
                case DataConnectionType.MySQL_ServerAndPortConnection:
                    connectionString = $"Server={dbParameters.Server};Port={dbParameters.Port};Database={dbParameters.Database};Uid={dbParameters.UserId};Pwd={dbParameters.Password};Connection Timeout={dbParameters.ConnectionTimeout};";
                    break;
                case DataConnectionType.ODBC_TrustedConnection:
                    connectionString = $"DSN={dbParameters.DSN};";
                    break;
                case DataConnectionType.ODBC_StandardLogin:
                    connectionString = $"DSN={dbParameters.DSN};Uid={dbParameters.UserId};Pwd={dbParameters.Password};";
                    break;
                case DataConnectionType.Excel_Ace_OLEDB12:
                    dbParameters.Provider = "Microsoft.ACE.OLEDB.12.0";
                    dbParameters.ExtendedProperties = Excel.GetExtendedProperties(dbParameters.FilePath, false);
                    connectionString = $"Provider={dbParameters.Provider};Data Source={dbParameters.FilePath};Extended Properties=\"{dbParameters.ExtendedProperties}\";";
                    break;
                case DataConnectionType.Excel_Ace_OLEDB12_With_IMEX:
                    dbParameters.Provider = "Microsoft.ACE.OLEDB.12.0";
                    dbParameters.ExtendedProperties = Excel.GetExtendedProperties(dbParameters.FilePath, true);
                    connectionString = $"Provider={dbParameters.Provider};Data Source={dbParameters.FilePath};Extended Properties=\"{dbParameters.ExtendedProperties}\";";
                    break;
                case DataConnectionType.Excel_Jet_OLEDB4:
                    dbParameters.Provider = "Microsoft.Jet.OLEDB.4.0";
                    dbParameters.ExtendedProperties = Excel.GetExtendedProperties(dbParameters.FilePath, false);
                    connectionString = $"Provider={dbParameters.Provider};Data Source={dbParameters.FilePath};Extended Properties=\"{dbParameters.ExtendedProperties}\";";
                    break;
                case DataConnectionType.Excel_Jet_OLEDB4_With_IMEX:
                    dbParameters.Provider = "Microsoft.Jet.OLEDB.4.0";
                    dbParameters.ExtendedProperties = Excel.GetExtendedProperties(dbParameters.FilePath, true);
                    connectionString = $"Provider={dbParameters.Provider};Data Source={dbParameters.FilePath};Extended Properties=\"{dbParameters.ExtendedProperties}\";";
                    break;
                default:
                    connectionString = "";
                    break;
            }

            dbParameters.ConnectionString = connectionString;

            if (saveDatabaseConnectionParameters)
                StoreDatabaseConnectionParameters(dbParameters, true);

            return connectionString;
        }
        /// <summary>
        /// Setup Connection String in stored DatabaseConnectionParameters
        /// </summary>
        /// <param name="dbConnectionType">Connection Type</param>
        /// <param name="saveDatabaseConnectionParameters">If TRUE DatabaseConnectionParameters will be stored in the class instance</param>
        /// <returns>Returns Connection String - It will also be stored in DatabaseConnectionParameters.ConnectionString</returns>
        public string SetConnectionString(DataConnectionType dbConnectionType, bool saveDatabaseConnectionParameters = false)
        {
            return SetConnectionString(DbStoredParameters, dbConnectionType, saveDatabaseConnectionParameters);
        }

        /// <summary>
        /// Allows user to store DatabaseConnectionParameters in DataConnector instance
        /// </summary>
        /// <param name="dbParameters"></param>
        /// <param name="overwriteOldParameters"></param>
        /// <returns></returns>
        public bool StoreDatabaseConnectionParameters(DatabaseConnectionParameters dbParameters, bool overwriteOldParameters = false)
        {
            var parametersSaved = true;

            if (DbStoredParameters == null)
                DbStoredParameters = dbParameters;
            else
            {
                if (overwriteOldParameters)
                    DbStoredParameters = dbParameters;
                else
                    parametersSaved = false;
            }

            return parametersSaved;
        }

        /// <summary>
        /// Get the message of the latest exception
        /// </summary>
        public string LastException
        {
            get
            {
                return Logs.GetLastException().Message;
            }
        }

        /// <summary>
        /// Check if last command executed in stored DatabaseConnectionParameters succeeded
        /// </summary>
        /// <returns>Returns TRUE if last command succeded. Returns FALSE if last command failed. Throw an exception if there are no stored DatabaseConnectionParameters.</returns>
        public bool LastCommandSucceeded()
        {
            return DbStoredParameters.LastCommandSucceeded;
        }
        /// <summary>
        /// Check if last command executed succeeded
        /// </summary>
        /// <returns>Returns TRUE if last command succeded. Returns FALSE if last command failed.</returns>
        public bool LastCommandSucceeded(DatabaseConnectionParameters dbParameters)
        {
            return dbParameters.LastCommandSucceeded;
        }

        /// <summary>
        /// Set Connection Timeout - Not working with OleDB
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="timeout">Timeout</param>
        public void SetConnectionTimeout(DatabaseConnectionParameters dbParameters, int timeout)
        {
            dbParameters.ConnectionTimeout = timeout;
        }
        /// <summary>
        /// Set Connection Timeout - Not working with OleDB
        /// </summary>
        /// <param name="timeout">Timeout</param>
        public void SetConnectionTimeout(int timeout)
        {
            SetConnectionTimeout(DbStoredParameters, timeout);
        }

        /// <summary>
        /// Set Command Timeout
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="timeout">Timeout</param>
        public void SetCommandTimeout(DatabaseConnectionParameters dbParameters, int timeout)
        {
            dbParameters.CommandTimeout = timeout;
        }
        /// <summary>
        /// Set Command Timeout
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="timeout">Timeout</param>
        public void SetCommandTimeout(int timeout)
        {
            SetCommandTimeout(DbStoredParameters, timeout);
        }

        /// <summary>
        /// Using this method, Build Commands will be treated as Raw Queries (this is set as default option)
        /// </summary>
        public void SetBuildCommandAsRawQuery()
        {
            ExecuteBuildCommandAsStoredProcedure = false;
        }
        /// <summary>
        /// Using this method, Build Commands will be treated as Stored Procedure
        /// </summary>
        public void SetBuildCommandAsStoredProcedure()
        {
            ExecuteBuildCommandAsStoredProcedure = true;
        }

        /// <summary>
        /// Check if connection is open
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <returns>Returns TRUE if connection is open</returns>
        public bool IsOpen(DatabaseConnectionParameters dbParameters)
        {
            switch (dbParameters.ConnectionType)
            {
                case DataConnectionType.SQLServer_StandardSecurity:
                case DataConnectionType.SQLServer_TrustedConnection:
                case DataConnectionType.SQLServer_StandardSecurity_UseIpAddressAndPort:
                    return SQLServer.IsOpen(dbParameters);

                case DataConnectionType.Access_ACE_OLEDB12_StandardSecurity:
                case DataConnectionType.Access_ACE_OLEDB12_WithPassword:
                case DataConnectionType.Access_JET_OLEDB4_StandardSecurity:
                case DataConnectionType.Access_JET_OLEDB4_WithPassword:
                    return Access.IsOpen(dbParameters);

                case DataConnectionType.MySQL_StandardConnection:
                case DataConnectionType.MySQL_ServerAndPortConnection:
                    return MySQL.IsOpen(dbParameters);

                case DataConnectionType.ODBC_TrustedConnection:
                case DataConnectionType.ODBC_StandardLogin:
                    return ODBC.IsOpen(dbParameters);

                case DataConnectionType.Excel_Jet_OLEDB4_With_IMEX:
                case DataConnectionType.Excel_Ace_OLEDB12_With_IMEX:
                case DataConnectionType.Excel_Ace_OLEDB12:
                case DataConnectionType.Excel_Jet_OLEDB4:
                    return Excel.IsOpen(dbParameters);

                default:
                    return false;
            }
        }

        /// <summary>
        /// Check if connection is open - Only Stored DatabaseConnectionParameters
        /// </summary>
        /// <returns>Returns TRUE if connection is open</returns>
        public bool IsOpen()
        {
            try
            {
                if (DbStoredParameters == null)
                    throw new Exception("No Stored Parameters");

                return IsOpen(DbStoredParameters);
            }
            catch (Exception excp)
            {
                Logs.AddException(excp);
                return false;
            }

        }

        /// <summary>
        /// Connect to Database - SetConnectionString() no needed
        /// </summary>
        /// <param name="dbParameters">Parameters</param>
        /// <param name="dbConnectionType">Connection Type</param>
        /// <param name="saveDatabaseConnectionParameters">If TRUE DatabaseConnectionParameters will be stored in the class instance</param>
        /// <returns>Returns TRUE if connection succeeds. Returns FALSE if connection fails.</returns>
        public bool ConnectToDatabase(DatabaseConnectionParameters dbParameters, DataConnectionType dbConnectionType, bool saveDatabaseConnectionParameters = false)
        {
            SetConnectionString(dbParameters, dbConnectionType, saveDatabaseConnectionParameters);
            return ConnectToDatabase(dbParameters);
        }
        /// <summary>
        /// Connect to Database - SetConnectionString() needed
        /// </summary>
        /// <param name="dbParameters">Parameters</param>
        /// <returns>Returns TRUE if connection succeeds. Returns FALSE if connection fails.</returns>
        public bool ConnectToDatabase(DatabaseConnectionParameters dbParameters)
        {

            switch (dbParameters.ConnectionType)
            {
                case DataConnectionType.SQLServer_StandardSecurity:
                case DataConnectionType.SQLServer_TrustedConnection:
                case DataConnectionType.SQLServer_StandardSecurity_UseIpAddressAndPort:
                    SQLServer.ConnectToDatabase(dbParameters);
                    break;

                case DataConnectionType.Access_ACE_OLEDB12_StandardSecurity:
                case DataConnectionType.Access_ACE_OLEDB12_WithPassword:
                case DataConnectionType.Access_JET_OLEDB4_StandardSecurity:
                case DataConnectionType.Access_JET_OLEDB4_WithPassword:
                    Access.ConnectToDatabase(dbParameters);
                    break;

                case DataConnectionType.MySQL_StandardConnection:
                case DataConnectionType.MySQL_ServerAndPortConnection:
                    MySQL.ConnectToDatabase(dbParameters);
                    break;

                case DataConnectionType.ODBC_TrustedConnection:
                case DataConnectionType.ODBC_StandardLogin:
                    ODBC.ConnectToDatabase(dbParameters);
                    break;

                case DataConnectionType.Excel_Jet_OLEDB4_With_IMEX:
                case DataConnectionType.Excel_Ace_OLEDB12_With_IMEX:
                case DataConnectionType.Excel_Ace_OLEDB12:
                case DataConnectionType.Excel_Jet_OLEDB4:
                    Excel.ConnectToDatabase(dbParameters);
                    break;

                default:
                    dbParameters.LastCommandSucceeded = false;
                    break;
            }

            return dbParameters.LastCommandSucceeded;

        }
        /// <summary>
        /// Connect to Database using stored Connection Parameters - SetConnectionString() needed. SaveDatabaseConnectionParameters() needed.
        /// </summary>
        /// <returns>Returns TRUE if connection succeeds. Returns FALSE if connection fails.</returns>
        public bool ConnectToDatabase()
        {
            try
            {
                if (DbStoredParameters == null)
                    throw new Exception("No Stored Parameters");

                return ConnectToDatabase(DbStoredParameters);
            }
            catch (Exception excp)
            {
                Logs.AddException(excp);
                return false;
            }

        }

        /// <summary>
        /// Disconnect from Database
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <returns></returns>
        public bool DisconnectFromDatabase(DatabaseConnectionParameters dbParameters)
        {
            switch (dbParameters.ConnectionType)
            {
                case DataConnectionType.SQLServer_StandardSecurity:
                case DataConnectionType.SQLServer_TrustedConnection:
                case DataConnectionType.SQLServer_StandardSecurity_UseIpAddressAndPort:
                    SQLServer.DisconnectFromDatabase(dbParameters);
                    break;

                case DataConnectionType.Access_ACE_OLEDB12_StandardSecurity:
                case DataConnectionType.Access_ACE_OLEDB12_WithPassword:
                case DataConnectionType.Access_JET_OLEDB4_StandardSecurity:
                case DataConnectionType.Access_JET_OLEDB4_WithPassword:
                    Access.DisconnectFromDatabase(dbParameters);
                    break;

                case DataConnectionType.MySQL_StandardConnection:
                case DataConnectionType.MySQL_ServerAndPortConnection:
                    MySQL.DisconnectFromDatabase(dbParameters);
                    break;

                case DataConnectionType.ODBC_TrustedConnection:
                case DataConnectionType.ODBC_StandardLogin:
                    ODBC.DisconnectFromDatabase(dbParameters);
                    break;

                case DataConnectionType.Excel_Jet_OLEDB4_With_IMEX:
                case DataConnectionType.Excel_Ace_OLEDB12_With_IMEX:
                case DataConnectionType.Excel_Ace_OLEDB12:
                case DataConnectionType.Excel_Jet_OLEDB4:
                    Excel.DisconnectFromDatabase(dbParameters);
                    break;

                default:
                    dbParameters.LastCommandSucceeded = false;
                    break;
            }

            return dbParameters.LastCommandSucceeded;
        }
        /// <summary>
        /// Disconnect from Database
        /// </summary>
        /// <returns></returns>
        public bool DisconnectFromDatabase()
        {
            try
            {
                if (DbStoredParameters == null)
                    throw new Exception("No Stored Parameters");

                return DisconnectFromDatabase(DbStoredParameters);
            }
            catch (Exception excp)
            {
                Logs.AddException(excp);
                return false;
            }
        }

        /// <summary>
        /// Set QueryParameters in given DatabaseConnectionParameters
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="queryParameters">Query Parameters</param>
        public void SetQueryAndParameters(DatabaseConnectionParameters dbParameters, IEnumerable<object> queryParameters)
        {
            dbParameters.QueryParameters = queryParameters;
        }
        /// <summary>
        /// Set QueryParameters in stored DatabaseConnectionParameters
        /// </summary>
        /// <param name="queryParameters">Query Parameters</param>
        public void SetQueryAndParameters(IEnumerable<object> queryParameters)
        {
            SetQueryAndParameters(DbStoredParameters, queryParameters);
        }
        /// <summary>
        /// Set Query and QueryParameters in given DatabaseConnectionParameters
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="query">Query to execute</param>
        /// <param name="queryParameters">Query Parameters</param>
        public void SetQueryAndParameters(DatabaseConnectionParameters dbParameters, string query, IEnumerable<object> queryParameters)
        {
            dbParameters.QueryString = query;
            dbParameters.QueryParameters = queryParameters;
        }
        /// <summary>
        /// Set Query and QueryParameters in stored DatabaseConnectionParameters
        /// </summary>
        /// <param name="query">Query to execute</param>
        /// <param name="queryParameters">Query Parameters</param>
        public void SetQueryAndParameters(string query, IEnumerable<object> queryParameters)
        {
            SetQueryAndParameters(DbStoredParameters, query, queryParameters);
        }

        /// <summary>
        /// Next command will be treated as Raw Query
        /// </summary>
        public void SetNextCommandAsRawQuery(DatabaseConnectionParameters dbParameters)
        {
            dbParameters.CommandType = System.Data.CommandType.Text;
        }
        /// <summary>
        /// Next command will be treated as Raw Query
        /// </summary>
        public void SetNextCommandAsRawQuery()
        {
            SetNextCommandAsRawQuery(DbStoredParameters);
        }

        /// <summary>
        /// Next command will be treated as Stored Procedure
        /// </summary>
        public void SetNextCommandAsStoredProcedure(DatabaseConnectionParameters dbParameters)
        {
            dbParameters.CommandType = System.Data.CommandType.StoredProcedure;
        }
        /// <summary>
        /// Next command will be treated as Stored Procedure
        /// </summary>
        public void SetNextCommandAsStoredProcedure()
        {
            SetNextCommandAsStoredProcedure(DbStoredParameters);
        }

        #region Set Data Container
        /// <summary>
        /// Set a DataSet as a container for built queries
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="dataSetContainer">DataSet</param>
        public void SetDataContainer(DatabaseConnectionParameters dbParameters, DataSet dataSetContainer)
        {
            dbParameters.DataContainerType = DataContainerType.DataSet;
            dbParameters.DataSetContainer = dataSetContainer;
        }
        /// <summary>
        /// Set a DataSet as a container for built queries
        /// </summary>
        /// <param name="dataSetContainer">DataSet</param>
        public void SetDataContainer(DataSet dataSetContainer)
        {
            SetDataContainer(DbStoredParameters, dataSetContainer);
        }

        /// <summary>
        /// Set a DataTable within a DataSet as a container for built queries
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="dataSetContainer">DataSet</param>
        /// <param name="dataTableName">DataTable Name</param>
        public void SetDataContainer(DatabaseConnectionParameters dbParameters, DataSet dataSetContainer, string dataTableName)
        {
            dbParameters.DataContainerType = DataContainerType.DataSetWithTable;
            dbParameters.DataSetContainer = dataSetContainer;
            dbParameters.DataTableContainerName = dataTableName;
        }
        /// <summary>
        /// Set a DataTable within a DataSet as a container for built queries
        /// </summary>
        /// <param name="dataSetContainer">DataSet</param>
        /// <param name="dataTableName">DataTable Name</param>
        public void SetDataContainer(DataSet dataSetContainer, string dataTableName)
        {
            SetDataContainer(DbStoredParameters, dataSetContainer, dataTableName);
        }

        /// <summary>
        /// Set a DataTable as a container for built queries
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="dataTableContainer">DataSet</param>
        public void SetDataContainer(DatabaseConnectionParameters dbParameters, DataTable dataTableContainer)
        {
            dbParameters.DataContainerType = DataContainerType.DataTable;
            dbParameters.DataTableContainer = dataTableContainer;
        }
        /// <summary>
        /// Set a DataTable as a container for built queries
        /// </summary>
        /// <param name="dataTableContainer">DataSet</param>
        public void SetDataContainer(DataTable dataTableContainer)
        {
            SetDataContainer(DbStoredParameters, dataTableContainer);
        }

        /// <summary>
        /// Set a DataRowCollection as a container for built queries
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="dataRowCollectionContainer">DataRowCollection</param>
        public void SetDataContainer(DatabaseConnectionParameters dbParameters, DataRow[] dataRowCollectionContainer)
        {
            dbParameters.DataContainerType = DataContainerType.DataRowsCollection;
            dbParameters.DataRowsCollectionContainer = dataRowCollectionContainer;
        }
        /// <summary>
        /// Set a DataRowCollection as a container for built queries
        /// </summary>
        /// <param name="dataRowCollectionContainer">DataRowCollection</param>
        public void SetDataContainer(DataRow[] dataRowCollectionContainer)
        {
            SetDataContainer(DbStoredParameters, dataRowCollectionContainer);
        }
        #endregion

        #region Main Methods
        /// <summary>
        /// Build INSERT, UPDATE, DELETE queries
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildCommand(DatabaseConnectionParameters dbParameters)
        {
            dbParameters.LastCommandSucceeded = false;

            if (ExecuteBuildCommandAsStoredProcedure)
                SetNextCommandAsStoredProcedure(dbParameters);
            else
                SetNextCommandAsRawQuery(dbParameters);

            int returnValue = -1;

            switch (dbParameters.ConnectionType)
            {
                case DataConnectionType.SQLServer_StandardSecurity:
                case DataConnectionType.SQLServer_TrustedConnection:
                case DataConnectionType.SQLServer_StandardSecurity_UseIpAddressAndPort:
                    returnValue = SQLServer.BuildCommand(dbParameters);
                    break;

                case DataConnectionType.Access_ACE_OLEDB12_StandardSecurity:
                case DataConnectionType.Access_ACE_OLEDB12_WithPassword:
                case DataConnectionType.Access_JET_OLEDB4_StandardSecurity:
                case DataConnectionType.Access_JET_OLEDB4_WithPassword:
                    returnValue = Access.BuildCommand(dbParameters);
                    break;

                case DataConnectionType.MySQL_StandardConnection:
                case DataConnectionType.MySQL_ServerAndPortConnection:
                    returnValue = MySQL.BuildCommand(dbParameters);
                    break;

                case DataConnectionType.ODBC_TrustedConnection:
                case DataConnectionType.ODBC_StandardLogin:
                    returnValue = ODBC.BuildCommand(dbParameters);
                    break;

                case DataConnectionType.Excel_Jet_OLEDB4_With_IMEX:
                case DataConnectionType.Excel_Ace_OLEDB12_With_IMEX:
                case DataConnectionType.Excel_Ace_OLEDB12:
                case DataConnectionType.Excel_Jet_OLEDB4:
                    returnValue = Excel.BuildCommand(dbParameters);
                    break;

                default:
                    break;
            }

            dbParameters.QueryParameters = null;
            return returnValue;
        }
        /// <summary>
        /// Build INSERT, UPDATE, DELETE queries
        /// </summary>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildCommand()
        {
            return BuildCommand(DbStoredParameters);
        }

        /// <summary>
        /// Execute given query
        /// </summary>
        /// <param name="dbParameters">Parameters</param>
        /// <param name="query">Query to execute</param>
        /// <returns>Returns a value (usually an IDataReader) if command succeeded. Returns NULL if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public IDataReader ExecuteReader(DatabaseConnectionParameters dbParameters, string query)
        {
            dbParameters.QueryString = query;
            return ExecuteReader(dbParameters);
        }
        /// <summary>
        /// Execute query stored in DatabaseConnectionParameters.Query
        /// </summary>
        /// <param name="dbParameters">Parameters</param>
        /// <returns>Returns a value (usually an IDataReader) if command succeeded. Returns NULL if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public IDataReader ExecuteReader(DatabaseConnectionParameters dbParameters)
        {
            // Set LastCommandSucceeded on FALSE because of the new command
            // It will be set automatically after command execution
            dbParameters.LastCommandSucceeded = false;
            IDataReader returnValue = null;

            switch (dbParameters.ConnectionType)
            {
                case DataConnectionType.SQLServer_StandardSecurity:
                case DataConnectionType.SQLServer_TrustedConnection:
                case DataConnectionType.SQLServer_StandardSecurity_UseIpAddressAndPort:
                    returnValue = SQLServer.ExecuteReader(dbParameters);
                    break;

                case DataConnectionType.Access_ACE_OLEDB12_StandardSecurity:
                case DataConnectionType.Access_ACE_OLEDB12_WithPassword:
                case DataConnectionType.Access_JET_OLEDB4_StandardSecurity:
                case DataConnectionType.Access_JET_OLEDB4_WithPassword:
                    returnValue = Access.ExecuteReader(dbParameters);
                    break;

                case DataConnectionType.MySQL_StandardConnection:
                case DataConnectionType.MySQL_ServerAndPortConnection:
                    returnValue = MySQL.ExecuteReader(dbParameters);
                    break;

                case DataConnectionType.ODBC_TrustedConnection:
                case DataConnectionType.ODBC_StandardLogin:
                    returnValue = ODBC.ExecuteReader(dbParameters);
                    break;

                case DataConnectionType.Excel_Jet_OLEDB4_With_IMEX:
                case DataConnectionType.Excel_Ace_OLEDB12_With_IMEX:
                case DataConnectionType.Excel_Ace_OLEDB12:
                case DataConnectionType.Excel_Jet_OLEDB4:
                    returnValue = Excel.ExecuteReader(dbParameters);
                    break;

                default:
                    break;
            }

            dbParameters.QueryParameters = null;
            return returnValue;
        }
        /// <summary>
        /// Execute given query using stored DatabaseConnectionParameters
        /// </summary>
        /// <param name="query">Query to execute</param>
        /// <returns>Returns a value (usually an IDataReader) if command succeeded. Returns NULL if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public IDataReader ExecuteReader(string query)
        {
            DbStoredParameters.QueryString = query;
            return ExecuteReader(DbStoredParameters);
        }

        /// <summary>
        /// Execute given query
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="query">Query to execute</param>
        /// /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int ExecuteNonQuery(DatabaseConnectionParameters dbParameters, string query)
        {
            dbParameters.QueryString = query;
            return ExecuteNonQuery(dbParameters);
        }
        /// <summary>
        /// Execute query stored in DatabaseConnectionParameters.Query
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int ExecuteNonQuery(DatabaseConnectionParameters dbParameters)
        {
            // Set LastCommandSucceeded on FALSE because of the new command
            // It will be set automatically after command execution
            dbParameters.LastCommandSucceeded = false;
            int returnValue = -1;

            switch (dbParameters.ConnectionType)
            {
                case DataConnectionType.SQLServer_StandardSecurity:
                case DataConnectionType.SQLServer_TrustedConnection:
                case DataConnectionType.SQLServer_StandardSecurity_UseIpAddressAndPort:
                    returnValue = SQLServer.ExecuteNonQuery(dbParameters);
                    break;

                case DataConnectionType.Access_ACE_OLEDB12_StandardSecurity:
                case DataConnectionType.Access_ACE_OLEDB12_WithPassword:
                case DataConnectionType.Access_JET_OLEDB4_StandardSecurity:
                case DataConnectionType.Access_JET_OLEDB4_WithPassword:
                    returnValue = Access.ExecuteNonQuery(dbParameters);
                    break;

                case DataConnectionType.MySQL_StandardConnection:
                case DataConnectionType.MySQL_ServerAndPortConnection:
                    returnValue = MySQL.ExecuteNonQuery(dbParameters);
                    break;

                case DataConnectionType.ODBC_TrustedConnection:
                case DataConnectionType.ODBC_StandardLogin:
                    returnValue = ODBC.ExecuteNonQuery(dbParameters);
                    break;

                case DataConnectionType.Excel_Jet_OLEDB4_With_IMEX:
                case DataConnectionType.Excel_Ace_OLEDB12_With_IMEX:
                case DataConnectionType.Excel_Ace_OLEDB12:
                case DataConnectionType.Excel_Jet_OLEDB4:
                    returnValue = Excel.ExecuteNonQuery(dbParameters);
                    break;

                default:
                    break;
            }

            dbParameters.QueryParameters = null;
            return returnValue;
        }
        /// <summary>
        /// Execute given query using stored DatabaseConnectionParameters
        /// </summary>
        /// <param name="query">Query to execute</param>
        /// /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int ExecuteNonQuery(string query)
        {
            DbStoredParameters.QueryString = query;
            return ExecuteNonQuery(DbStoredParameters);
        }

        /// <summary>
        /// Execute given query
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="query">Query to execute</param>
        /// <returns>Returns a single value as object if command succeeds. May return NULL whether command fails or not. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public object ExecuteScalar(DatabaseConnectionParameters dbParameters, string query)
        {
            dbParameters.QueryString = query;
            return ExecuteScalar(dbParameters);
        }
        /// <summary>
        /// Execute query stored in DatabaseConnectionParameters.Query
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <returns>Returns a single value as object if command succeeds. May return NULL whether command fails or not. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public object ExecuteScalar(DatabaseConnectionParameters dbParameters)
        {
            // Set LastCommandSucceeded on FALSE because of the new command
            // It will be set automatically after command execution
            dbParameters.LastCommandSucceeded = false;
            object returnValue = null;

            switch (dbParameters.ConnectionType)
            {
                case DataConnectionType.SQLServer_StandardSecurity:
                case DataConnectionType.SQLServer_TrustedConnection:
                case DataConnectionType.SQLServer_StandardSecurity_UseIpAddressAndPort:
                    returnValue = SQLServer.ExecuteScalar(dbParameters);
                    break;

                case DataConnectionType.Access_ACE_OLEDB12_StandardSecurity:
                case DataConnectionType.Access_ACE_OLEDB12_WithPassword:
                case DataConnectionType.Access_JET_OLEDB4_StandardSecurity:
                case DataConnectionType.Access_JET_OLEDB4_WithPassword:
                    returnValue = Access.ExecuteScalar(dbParameters);
                    break;

                case DataConnectionType.MySQL_StandardConnection:
                case DataConnectionType.MySQL_ServerAndPortConnection:
                    returnValue = MySQL.ExecuteScalar(dbParameters);
                    break;

                case DataConnectionType.ODBC_TrustedConnection:
                case DataConnectionType.ODBC_StandardLogin:
                    returnValue = ODBC.ExecuteScalar(dbParameters);
                    break;

                case DataConnectionType.Excel_Jet_OLEDB4_With_IMEX:
                case DataConnectionType.Excel_Ace_OLEDB12_With_IMEX:
                case DataConnectionType.Excel_Ace_OLEDB12:
                case DataConnectionType.Excel_Jet_OLEDB4:
                    returnValue = Excel.ExecuteScalar(dbParameters);
                    break;

                default:
                    break;
            }

            dbParameters.QueryParameters = null;
            return returnValue;
        }
        /// <summary>
        /// Execute given query using stored DatabaseConnectionParameters
        /// </summary>
        /// <param name="query">Query to execute</param>
        /// <returns>Returns a single value as object if command succeeds. May return NULL whether command fails or not. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public object ExecuteScalar(string query)
        {
            DbStoredParameters.QueryString = query;
            return ExecuteScalar(DbStoredParameters);
        }
        #endregion

        #region Shortcuts
        /// <summary>
        /// Shortcut for SELECT Query
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="query">Query to execute</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <returns>Returns a value (usually an IDataReader) if command succeeded. Returns NULL if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public IDataReader Select(DatabaseConnectionParameters dbParameters, string query, IEnumerable<object> queryParameters)
        {
            dbParameters.QueryString = query;
            dbParameters.QueryParameters = queryParameters;
            return Select(dbParameters);
        }
        /// <summary>
        /// Shortcut for SELECT Query
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="query">Query to execute</param>
        /// <returns>Returns a value (usually an IDataReader) if command succeeded. Returns NULL if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public IDataReader Select(DatabaseConnectionParameters dbParameters, string query)
        {
            dbParameters.QueryString = query;
            return Select(dbParameters);
        }
        /// <summary>
        /// Shortcut for SELECT Query
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <returns>Returns a value (usually an IDataReader) if command succeeded. Returns NULL if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public IDataReader Select(DatabaseConnectionParameters dbParameters)
        {
            SetNextCommandAsRawQuery(dbParameters);
            return ExecuteReader(dbParameters);
        }
        /// <summary>
        /// Shortcut for SELECT Query
        /// </summary>
        /// <param name="query">Query to execute</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <returns>Returns a value (usually an IDataReader) if command succeeded. Returns NULL if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public IDataReader Select(string query, IEnumerable<object> queryParameters)
        {
            DbStoredParameters.QueryString = query;
            DbStoredParameters.QueryParameters = queryParameters;
            return Select(DbStoredParameters);
        }
        /// <summary>
        /// Shortcut for SELECT Query
        /// </summary>
        /// <param name="query">Query to execute</param>
        /// <returns>Returns a value (usually an IDataReader) if command succeeded. Returns NULL if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public IDataReader Select(string query)
        {
            DbStoredParameters.QueryString = query;
            return Select(DbStoredParameters);
        }

        /// <summary>
        /// Shortcut for SELECT Query
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="query">Query to execute</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <returns>Returns a single value as object if command succeeds. May return NULL whether command fails or not. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public object SelectSingleValue(DatabaseConnectionParameters dbParameters, string query, IEnumerable<object> queryParameters)
        {
            dbParameters.QueryString = query;
            dbParameters.QueryParameters = queryParameters;
            return SelectSingleValue(dbParameters);
        }
        /// <summary>
        /// Shortcut for SELECT Query
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="query">Query to execute</param>
        /// <returns>Returns a single value as object if command succeeds. May return NULL whether command fails or not. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public object SelectSingleValue(DatabaseConnectionParameters dbParameters, string query)
        {
            dbParameters.QueryString = query;
            return SelectSingleValue(dbParameters);
        }
        /// <summary>
        /// Shortcut for SELECT Query
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <returns>Returns a single value as object if command succeeds. May return NULL whether command fails or not. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public object SelectSingleValue(DatabaseConnectionParameters dbParameters)
        {
            SetNextCommandAsRawQuery(dbParameters);
            return ExecuteScalar(dbParameters);
        }
        /// <summary>
        /// Shortcut for SELECT Query
        /// </summary>
        /// <param name="query">Query to execute</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <returns>Returns a single value as object if command succeeds. May return NULL whether command fails or not. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public object SelectSingleValue(string query, IEnumerable<object> queryParameters)
        {
            DbStoredParameters.QueryString = query;
            DbStoredParameters.QueryParameters = queryParameters;
            return SelectSingleValue(DbStoredParameters);
        }
        /// <summary>
        /// Shortcut for SELECT Query
        /// </summary>
        /// <param name="query">Query to execute</param>
        /// <returns>Returns a single value as object if command succeeds. May return NULL whether command fails or not. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public object SelectSingleValue(string query)
        {
            DbStoredParameters.QueryString = query;
            return SelectSingleValue(DbStoredParameters);
        }

        /// <summary>
        /// Shortcut for INSERT Query
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="query">Query to execute</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int Insert(DatabaseConnectionParameters dbParameters, string query, IEnumerable<object> queryParameters)
        {
            dbParameters.QueryString = query;
            dbParameters.QueryParameters = queryParameters;
            return Insert(dbParameters);
        }
        /// <summary>
        /// Shortcut for INSERT Query
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="query">Query to execute</param>
        /// /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int Insert(DatabaseConnectionParameters dbParameters, string query)
        {
            dbParameters.QueryString = query;
            return Insert(dbParameters);
        }
        /// <summary>
        /// Shortcut for INSERT Query
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int Insert(DatabaseConnectionParameters dbParameters)
        {
            SetNextCommandAsRawQuery(dbParameters);
            return ExecuteNonQuery(dbParameters);
        }
        /// <summary>
        /// Shortcut for INSERT Query
        /// </summary>
        /// <param name="query">Query to execute</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int Insert(string query, IEnumerable<object> queryParameters)
        {
            DbStoredParameters.QueryString = query;
            DbStoredParameters.QueryParameters = queryParameters;
            return Insert(DbStoredParameters);
        }
        /// <summary>
        /// Shortcut for INSERT Query
        /// </summary>
        /// <param name="query">Query to execute</param>
        /// /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int Insert(string query)
        {
            DbStoredParameters.QueryString = query;
            return Insert(DbStoredParameters);
        }

        /// <summary>
        /// Shortcut for UPDATE Query
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="query">Query to execute</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int Update(DatabaseConnectionParameters dbParameters, string query, IEnumerable<object> queryParameters)
        {
            dbParameters.QueryString = query;
            dbParameters.QueryParameters = queryParameters;
            return Update(dbParameters);
        }
        /// <summary>
        /// Shortcut for UPDATE Query
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="query">Query to execute</param>
        /// /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int Update(DatabaseConnectionParameters dbParameters, string query)
        {
            dbParameters.QueryString = query;
            return Update(dbParameters);
        }
        /// <summary>
        /// Shortcut for UPDATE Query
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int Update(DatabaseConnectionParameters dbParameters)
        {
            SetNextCommandAsRawQuery(dbParameters);
            return ExecuteNonQuery(dbParameters);
        }
        /// <summary>
        /// Shortcut for UPDATE Query
        /// </summary>
        /// <param name="query">Query to execute</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int Update(string query, IEnumerable<object> queryParameters)
        {
            DbStoredParameters.QueryString = query;
            DbStoredParameters.QueryParameters = queryParameters;
            return Update(DbStoredParameters);
        }
        /// <summary>
        /// Shortcut for UPDATE Query
        /// </summary>
        /// <param name="query">Query to execute</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int Update(string query)
        {
            DbStoredParameters.QueryString = query;
            return Update(DbStoredParameters);
        }

        /// <summary>
        /// Shortcut for DELETE Query
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="query">Query to execute</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int Delete(DatabaseConnectionParameters dbParameters, string query, IEnumerable<object> queryParameters)
        {
            dbParameters.QueryString = query;
            dbParameters.QueryParameters = queryParameters;
            return Delete(dbParameters);
        }
        /// <summary>
        /// Shortcut for DELETE Query
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="query">Query to execute</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int Delete(DatabaseConnectionParameters dbParameters, string query)
        {
            dbParameters.QueryString = query;
            return Delete(dbParameters);
        }
        /// <summary>
        /// Shortcut for DELETE Query
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int Delete(DatabaseConnectionParameters dbParameters)
        {
            SetNextCommandAsRawQuery(dbParameters);
            return ExecuteNonQuery(dbParameters);
        }
        /// <summary>
        /// Shortcut for DELETE Query
        /// </summary>
        /// <param name="query">Query to execute</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int Delete(string query, IEnumerable<object> queryParameters)
        {
            DbStoredParameters.QueryString = query;
            DbStoredParameters.QueryParameters = queryParameters;
            return Delete(DbStoredParameters);
        }
        /// <summary>
        /// Shortcut for DELETE Query
        /// </summary>
        /// <param name="query">Query to execute</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int Delete(string query)
        {
            DbStoredParameters.QueryString = query;
            return Delete(DbStoredParameters);
        }

        /// <summary>
        /// Shortcut for Stored Procedure that returns an IDataReader
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="storedProcedureName">Stored Procedure to execute</param>
        /// <param name="storedProcedureParameters">Stored Procedure Parameters</param>
        /// <returns>Returns a value (usually an IDataReader) if command succeeded. Returns NULL if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public IDataReader StoredProcedure(DatabaseConnectionParameters dbParameters, string storedProcedureName, IEnumerable<object> storedProcedureParameters)
        {
            dbParameters.QueryString = storedProcedureName;
            dbParameters.QueryParameters = storedProcedureParameters;
            return StoredProcedure(dbParameters);
        }
        /// <summary>
        /// Shortcut for Stored Procedure that returns an IDataReader
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="storedProcedureName">Stored Procedure to execute</param>
        /// <returns>Returns a value (usually an IDataReader) if command succeeded. Returns NULL if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public IDataReader StoredProcedure(DatabaseConnectionParameters dbParameters, string storedProcedureName)
        {
            dbParameters.QueryString = storedProcedureName;
            return StoredProcedure(dbParameters);
        }
        /// <summary>
        /// Shortcut for Stored Procedure that returns an IDataReader
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <returns>Returns a value (usually an IDataReader) if command succeeded. Returns NULL if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public IDataReader StoredProcedure(DatabaseConnectionParameters dbParameters)
        {
            SetNextCommandAsStoredProcedure(dbParameters);
            return ExecuteReader(dbParameters);
        }
        /// <summary>
        /// Shortcut for Stored Procedure that returns an IDataReader
        /// </summary>
        /// <param name="storedProcedureName">Stored Procedure to execute</param>
        /// <param name="storedProcedureParameters">Stored Procedure Parameters</param>
        /// <returns>Returns a value (usually an IDataReader) if command succeeded. Returns NULL if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public IDataReader StoredProcedure(string storedProcedureName, IEnumerable<object> storedProcedureParameters)
        {
            DbStoredParameters.QueryString = storedProcedureName;
            DbStoredParameters.QueryParameters = storedProcedureParameters;
            return StoredProcedure(DbStoredParameters);
        }
        /// <summary>
        /// Shortcut for Stored Procedure that returns an IDataReader
        /// </summary>
        /// <param name="storedProcedureName">Stored Procedure to execute</param>
        /// <returns>Returns a value (usually an IDataReader) if command succeeded. Returns NULL if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public IDataReader StoredProcedure(string storedProcedureName)
        {
            DbStoredParameters.QueryString = storedProcedureName;
            return StoredProcedure(DbStoredParameters);
        }

        /// <summary>
        /// Shortcut for Stored Procedure that returns the number of rows affected
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="storedProcedureName">Stored Procedure to execute</param>
        /// <param name="storedProcedureParameters">Stored Procedure Parameters</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int StoredProcedureNonQuery(DatabaseConnectionParameters dbParameters, string storedProcedureName, IEnumerable<object> storedProcedureParameters)
        {
            dbParameters.QueryString = storedProcedureName;
            dbParameters.QueryParameters = storedProcedureParameters;
            return StoredProcedureNonQuery(dbParameters);
        }
        /// <summary>
        /// Shortcut for Stored Procedure that returns the number of rows affected
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="storedProcedureName">Stored Procedure to execute</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int StoredProcedureNonQuery(DatabaseConnectionParameters dbParameters, string storedProcedureName)
        {
            dbParameters.QueryString = storedProcedureName;
            return StoredProcedureNonQuery(dbParameters);
        }
        /// <summary>
        /// Shortcut for Stored Procedure that returns the number of rows affected
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int StoredProcedureNonQuery(DatabaseConnectionParameters dbParameters)
        {
            SetNextCommandAsStoredProcedure(dbParameters);
            return ExecuteNonQuery(dbParameters);
        }
        /// <summary>
        /// Shortcut for Stored Procedure that returns the number of rows affected
        /// </summary>
        /// <param name="storedProcedureName">Stored Procedure to execute</param>
        /// <param name="storedProcedureParameters">Stored Procedure Parameters</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int StoredProcedureNonQuery(string storedProcedureName, IEnumerable<object> storedProcedureParameters)
        {
            DbStoredParameters.QueryString = storedProcedureName;
            DbStoredParameters.QueryParameters = storedProcedureParameters;
            return StoredProcedureNonQuery(DbStoredParameters);
        }
        /// <summary>
        /// Shortcut for Stored Procedure that returns the number of rows affected
        /// </summary>
        /// <param name="storedProcedureName">Stored Procedure to execute</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int StoredProcedureNonQuery(string storedProcedureName)
        {
            DbStoredParameters.QueryString = storedProcedureName;
            return StoredProcedureNonQuery(DbStoredParameters);
        }

        /// <summary>
        /// Shortcut for Stored Procedure that returns a single value
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="storedProcedureName">Stored Procedure to execute</param>
        /// <param name="storedProcedureParameters">Stored Procedure Parameters</param>
        /// <returns>Returns a single value as object if command succeeds. May return NULL whether command fails or not. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public object StoredProcedureSingleValue(DatabaseConnectionParameters dbParameters, string storedProcedureName, IEnumerable<object> storedProcedureParameters)
        {
            dbParameters.QueryString = storedProcedureName;
            dbParameters.QueryParameters = storedProcedureParameters;
            return StoredProcedureSingleValue(dbParameters);
        }
        /// <summary>
        /// Shortcut for Stored Procedure that returns a single value
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <param name="storedProcedureName">Stored Procedure to execute</param>
        /// <returns>Returns a single value as object if command succeeds. May return NULL whether command fails or not. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public object StoredProcedureSingleValue(DatabaseConnectionParameters dbParameters, string storedProcedureName)
        {
            dbParameters.QueryString = storedProcedureName;
            return StoredProcedureSingleValue(dbParameters);
        }
        /// <summary>
        /// Shortcut for Stored Procedure that returns a single value
        /// </summary>
        /// <param name="dbParameters">Connection Parameters</param>
        /// <returns>Returns a single value as object if command succeeds. May return NULL whether command fails or not. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public object StoredProcedureSingleValue(DatabaseConnectionParameters dbParameters)
        {
            SetNextCommandAsStoredProcedure(dbParameters);
            return ExecuteScalar(dbParameters);
        }
        /// <summary>
        /// Shortcut for Stored Procedure that returns a single value
        /// </summary>
        /// <param name="storedProcedureName">Stored Procedure to execute</param>
        /// <param name="storedProcedureParameters">Stored Procedure Parameters</param>
        /// <returns>Returns a single value as object if command succeeds. May return NULL whether command fails or not. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public object StoredProcedureSingleValue(string storedProcedureName, IEnumerable<object> storedProcedureParameters)
        {
            DbStoredParameters.QueryString = storedProcedureName;
            DbStoredParameters.QueryParameters = storedProcedureParameters;
            return StoredProcedureSingleValue(DbStoredParameters);
        }
        /// <summary>
        /// Shortcut for Stored Procedure that returns a single value
        /// </summary>
        /// <param name="storedProcedureName">Stored Procedure to execute</param>
        /// <returns>Returns a single value as object if command succeeds. May return NULL whether command fails or not. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public object StoredProcedureSingleValue(string storedProcedureName)
        {
            DbStoredParameters.QueryString = storedProcedureName;
            return StoredProcedureSingleValue(DbStoredParameters);
        }
        #endregion

        #region Build Commands
        /// <summary>
        /// Build INSERT command
        /// </summary>
        /// <param name="dbParameters">Connection Parameter</param>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataSetContainer">DataSet</param>
        /// <param name="dataTableName">DataTable Name</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildInsertCommand(DatabaseConnectionParameters dbParameters, string query, IEnumerable<object> queryParameters, DataSet dataSetContainer, string dataTableName)
        {
            SetDataContainer(dbParameters, dataSetContainer, dataTableName);
            SetQueryAndParameters(dbParameters, query, queryParameters);
            return BuildInsertCommand(dbParameters);
        }
        /// <summary>
        /// Build INSERT command
        /// </summary>
        /// <param name="dbParameters">Connection Parameter</param>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataSetContainer">DataSet</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildInsertCommand(DatabaseConnectionParameters dbParameters, string query, IEnumerable<object> queryParameters, DataSet dataSetContainer)
        {
            SetDataContainer(dbParameters, dataSetContainer);
            SetQueryAndParameters(dbParameters, query, queryParameters);
            return BuildInsertCommand(dbParameters);
        }
        /// <summary>
        /// Build INSERT command
        /// </summary>
        /// <param name="dbParameters">Connection Parameter</param>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataTableContainer">DataTable</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildInsertCommand(DatabaseConnectionParameters dbParameters, string query, IEnumerable<object> queryParameters, DataTable dataTableContainer)
        {
            SetDataContainer(dbParameters, dataTableContainer);
            SetQueryAndParameters(dbParameters, query, queryParameters);
            return BuildInsertCommand(dbParameters);
        }
        /// <summary>
        /// Build INSERT command
        /// </summary>
        /// <param name="dbParameters">Connection Parameter</param>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataRowCollection">DataRowCollection</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildInsertCommand(DatabaseConnectionParameters dbParameters, string query, IEnumerable<object> queryParameters, DataRow[] dataRowCollection)
        {
            SetDataContainer(dbParameters, dataRowCollection);
            SetQueryAndParameters(dbParameters, query, queryParameters);
            return BuildInsertCommand(dbParameters);
        }
        /// <summary>
        /// Build INSERT command
        /// </summary>
        /// <param name="dbParameters">Connection Parameter</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildInsertCommand(DatabaseConnectionParameters dbParameters)
        {
            dbParameters.CommandBuildType = CommandBuildType.Insert;
            return BuildCommand(dbParameters);
        }
        /// <summary>
        /// Build INSERT command
        /// </summary>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataSetContainer">DataSet</param>
        /// <param name="dataTableName">DataTable Name</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildInsertCommand(string query, IEnumerable<object> queryParameters, DataSet dataSetContainer, string dataTableName)
        {
            SetDataContainer(DbStoredParameters, dataSetContainer, dataTableName);
            SetQueryAndParameters(DbStoredParameters, query, queryParameters);
            return BuildInsertCommand(DbStoredParameters);
        }
        /// <summary>
        /// Build INSERT command
        /// </summary>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataSetContainer">DataSet</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildInsertCommand(string query, IEnumerable<object> queryParameters, DataSet dataSetContainer)
        {
            SetDataContainer(DbStoredParameters, dataSetContainer);
            SetQueryAndParameters(DbStoredParameters, query, queryParameters);
            return BuildInsertCommand(DbStoredParameters);
        }
        /// <summary>
        /// Build INSERT command
        /// </summary>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataTableContainer">DataTable</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildInsertCommand(string query, IEnumerable<object> queryParameters, DataTable dataTableContainer)
        {
            SetDataContainer(DbStoredParameters, dataTableContainer);
            SetQueryAndParameters(DbStoredParameters, query, queryParameters);
            return BuildInsertCommand(DbStoredParameters);
        }
        /// <summary>
        /// Build INSERT command
        /// </summary>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataRowCollection">DataRowCollection</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildInsertCommand(string query, IEnumerable<object> queryParameters, DataRow[] dataRowCollection)
        {
            SetDataContainer(DbStoredParameters, dataRowCollection);
            SetQueryAndParameters(DbStoredParameters, query, queryParameters);
            return BuildInsertCommand(DbStoredParameters);
        }

        /// <summary>
        /// Build UPDATE command
        /// </summary>
        /// <param name="dbParameters">Connection Parameter</param>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataSetContainer">DataSet</param>
        /// <param name="dataTableName">DataTable Name</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildUpdateCommand(DatabaseConnectionParameters dbParameters, string query, IEnumerable<object> queryParameters, DataSet dataSetContainer, string dataTableName)
        {
            SetDataContainer(dbParameters, dataSetContainer, dataTableName);
            SetQueryAndParameters(dbParameters, query, queryParameters);
            return BuildUpdateCommand(dbParameters);
        }
        /// <summary>
        /// Build UPDATE command
        /// </summary>
        /// <param name="dbParameters">Connection Parameter</param>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataSetContainer">DataSet</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildUpdateCommand(DatabaseConnectionParameters dbParameters, string query, IEnumerable<object> queryParameters, DataSet dataSetContainer)
        {
            SetDataContainer(dbParameters, dataSetContainer);
            SetQueryAndParameters(dbParameters, query, queryParameters);
            return BuildUpdateCommand(dbParameters);
        }
        /// <summary>
        /// Build UPDATE command
        /// </summary>
        /// <param name="dbParameters">Connection Parameter</param>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataTableContainer">DataTable</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildUpdateCommand(DatabaseConnectionParameters dbParameters, string query, IEnumerable<object> queryParameters, DataTable dataTableContainer)
        {
            SetDataContainer(dbParameters, dataTableContainer);
            SetQueryAndParameters(dbParameters, query, queryParameters);
            return BuildUpdateCommand(dbParameters);
        }
        /// <summary>
        /// Build UPDATE command
        /// </summary>
        /// <param name="dbParameters">Connection Parameter</param>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataRowCollection">DataRow<Collection</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildUpdateCommand(DatabaseConnectionParameters dbParameters, string query, IEnumerable<object> queryParameters, DataRow[] dataRowCollection)
        {
            SetDataContainer(dbParameters, dataRowCollection);
            SetQueryAndParameters(dbParameters, query, queryParameters);
            return BuildUpdateCommand(dbParameters);
        }
        /// <summary>
        /// Build UPDATE command
        /// </summary>
        /// <param name="dbParameters">Connection Parameter</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildUpdateCommand(DatabaseConnectionParameters dbParameters)
        {
            dbParameters.CommandBuildType = CommandBuildType.Update;
            return BuildCommand(dbParameters);
        }
        /// <summary>
        /// Build UPDATE command
        /// </summary>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataSetContainer">DataSet</param>
        /// <param name="dataTableName">DataTable Name</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildUpdateCommand(string query, IEnumerable<object> queryParameters, DataSet dataSetContainer, string dataTableName)
        {
            SetDataContainer(DbStoredParameters, dataSetContainer, dataTableName);
            SetQueryAndParameters(DbStoredParameters, query, queryParameters);
            return BuildUpdateCommand(DbStoredParameters);
        }
        /// <summary>
        /// Build UPDATE command
        /// </summary>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataSetContainer">DataSet</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildUpdateCommand(string query, IEnumerable<object> queryParameters, DataSet dataSetContainer)
        {
            SetDataContainer(DbStoredParameters, dataSetContainer);
            SetQueryAndParameters(DbStoredParameters, query, queryParameters);
            return BuildUpdateCommand(DbStoredParameters);
        }
        /// <summary>
        /// Build UPDATE command
        /// </summary>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataTableContainer">DataTable</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildUpdateCommand(string query, IEnumerable<object> queryParameters, DataTable dataTableContainer)
        {
            SetDataContainer(DbStoredParameters, dataTableContainer);
            SetQueryAndParameters(DbStoredParameters, query, queryParameters);
            return BuildUpdateCommand(DbStoredParameters);
        }
        /// <summary>
        /// Build UPDATE command
        /// </summary>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataRowCollection">DataRow<Collection</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildUpdateCommand(string query, IEnumerable<object> queryParameters, DataRow[] dataRowCollection)
        {
            SetDataContainer(DbStoredParameters, dataRowCollection);
            SetQueryAndParameters(DbStoredParameters, query, queryParameters);
            return BuildUpdateCommand(DbStoredParameters);
        }


        /// <summary>
        /// Build DELETE command
        /// </summary>
        /// <param name="dbParameters">Connection Parameter</param>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataSetContainer">DataSet</param>
        /// <param name="dataTableName">DataTable Name</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildDeleteCommand(DatabaseConnectionParameters dbParameters, string query, IEnumerable<object> queryParameters, DataSet dataSetContainer, string dataTableName)
        {
            SetDataContainer(dbParameters, dataSetContainer, dataTableName);
            SetQueryAndParameters(dbParameters, query, queryParameters);
            return BuildDeleteCommand(dbParameters);
        }
        /// <summary>
        /// Build DELETE command
        /// </summary>
        /// <param name="dbParameters">Connection Parameter</param>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataSetContainer">DataSet</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildDeleteCommand(DatabaseConnectionParameters dbParameters, string query, IEnumerable<object> queryParameters, DataSet dataSetContainer)
        {
            SetDataContainer(dbParameters, dataSetContainer);
            SetQueryAndParameters(dbParameters, query, queryParameters);
            return BuildDeleteCommand(dbParameters);
        }
        /// <summary>
        /// Build DELETE command
        /// </summary>
        /// <param name="dbParameters">Connection Parameter</param>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataTableContainer">DataTable</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildDeleteCommand(DatabaseConnectionParameters dbParameters, string query, IEnumerable<object> queryParameters, DataTable dataTableContainer)
        {
            SetDataContainer(dbParameters, dataTableContainer);
            SetQueryAndParameters(dbParameters, query, queryParameters);
            return BuildDeleteCommand(dbParameters);
        }
        /// <summary>
        /// Build DELETE command
        /// </summary>
        /// <param name="dbParameters">Connection Parameter</param>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataRowCollection">DataRowCollection</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildDeleteCommand(DatabaseConnectionParameters dbParameters, string query, IEnumerable<object> queryParameters, DataRow[] dataRowCollection)
        {
            SetDataContainer(dbParameters, dataRowCollection);
            SetQueryAndParameters(dbParameters, query, queryParameters);
            return BuildDeleteCommand(dbParameters);
        }
        /// <summary>
        /// Build DELETE command
        /// </summary>
        /// <param name="dbParameters">Connection Parameter</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildDeleteCommand(DatabaseConnectionParameters dbParameters)
        {
            dbParameters.CommandBuildType = CommandBuildType.Delete;
            return BuildCommand(dbParameters);
        }
        /// <summary>
        /// Build DELETE command
        /// </summary>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataSetContainer">DataSet</param>
        /// <param name="dataTableName">DataTable Name</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildDeleteCommand(string query, IEnumerable<object> queryParameters, DataSet dataSetContainer, string dataTableName)
        {
            SetDataContainer(DbStoredParameters, dataSetContainer, dataTableName);
            SetQueryAndParameters(DbStoredParameters, query, queryParameters);
            return BuildDeleteCommand(DbStoredParameters);
        }
        /// <summary>
        /// Build DELETE command
        /// </summary>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataSetContainer">DataSet</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildDeleteCommand(string query, IEnumerable<object> queryParameters, DataSet dataSetContainer)
        {
            SetDataContainer(DbStoredParameters, dataSetContainer);
            SetQueryAndParameters(DbStoredParameters, query, queryParameters);
            return BuildDeleteCommand(DbStoredParameters);
        }
        /// <summary>
        /// Build DELETE command
        /// </summary>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataTableContainer">DataTable</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildDeleteCommand(string query, IEnumerable<object> queryParameters, DataTable dataTableContainer)
        {
            SetDataContainer(DbStoredParameters, dataTableContainer);
            SetQueryAndParameters(DbStoredParameters, query, queryParameters);
            return BuildDeleteCommand(DbStoredParameters);
        }
        /// <summary>
        /// Build DELETE command
        /// </summary>
        /// <param name="query">Query to build</param>
        /// <param name="queryParameters">Query Parameters</param>
        /// <param name="dataRowCollection">DataRowCollection</param>
        /// <returns>Returns the number of rows affected if command succeeds. Returns -1 if command fails. Command execution success will also be stored in DatabaseConnectionParameters.LastCommandSucceeded</returns>
        public int BuildDeleteCommand(string query, IEnumerable<object> queryParameters, DataRow[] dataRowCollection)
        {
            SetDataContainer(DbStoredParameters, dataRowCollection);
            SetQueryAndParameters(DbStoredParameters, query, queryParameters);
            return BuildDeleteCommand(DbStoredParameters);
        }
        #endregion

        /// Specific Helper Methods will ALWAYS returns values from methods of other classes
        /// Specific Helper Methods name will ALWAYS begin with database tyoe name (ODBC_, ACCESS_, EXCEL_ etc...)
        #region Specific Helper Methods
        /// <summary>
        /// Provides an help to build a string that contains a Stored Procedure for ODBC connections, since ODBC's Stored Procedures are syntactically different from others
        /// </summary>
        /// <param name="storedProcedureName">Stored Procedure to execute</param>
        /// <param name="storedProcedureParameters">Stored Procedure Parameters</param>
        /// <returns>Returns a string that contains a Stored Procedure that can be executed with ODBC connections</returns>
        public string ODBC_StoredProcedureStringBuilder(string storedProcedureName, IEnumerable<object> storedProcedureParameters)
        {
            return ODBC.BuildStoredProcedure(storedProcedureName, storedProcedureParameters);
        }
        #endregion

    }
}
