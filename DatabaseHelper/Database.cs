using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace DatabaseHelper {
  public class Database : IDisposable {
    private readonly bool _useSingletone;
    private SqlConnection _connection;
    private SqlTransaction _transaction;

    private bool IsTransactionActive => _transaction != null;

    public string ConnectionString { get; private set; }

    public Database(string connectionString, bool useSingletone = false) {
      _useSingletone = useSingletone;
      ConnectionString = connectionString;
    }

    public Database(bool useSingletone = false) {
      _useSingletone = useSingletone;
      var typeName = this.GetType().Name;
      ConnectionString = ConfigurationManager.ConnectionStrings[typeName].ConnectionString;
    }

    public SqlConnection GetConnection(string connectionString) {
      if (_connection == null || !_useSingletone)
        _connection = new SqlConnection(connectionString);
      return _connection;
    }

    public SqlConnection GetConnection() {
      return GetConnection(ConnectionString);
    }

    public void BeginTransaction() {
      if (!_useSingletone) throw new Exception("Transaction is supported only in singletone mode");
      if (_transaction != null) throw new Exception("There is an active transaction");

      var connection = GetConnection();
      if (connection.State != ConnectionState.Open)
        connection.Open();
      _transaction = connection.BeginTransaction();
    }

    public void CommitTransaction() {
      ValidateTransaction();
      _transaction.Commit();
      _transaction.Dispose();
      _transaction = null;
    }

    public void RollbackTransaction() {
      ValidateTransaction();
      _transaction.Rollback();
      _transaction.Dispose();
      _transaction = null;
    }

    #region GetCommand

    public SqlCommand GetCommand(string connectionString, string commandText, CommandType commandType, params SqlParameter[] parameters) {
      SqlCommand command = GetConnection(connectionString).CreateCommand();
      command.CommandText = commandText;
      command.CommandType = commandType;
      command.Parameters.AddRange(parameters);
      if (IsTransactionActive) {
        command.Transaction = _transaction;
      }
      return command;
    }

    public SqlCommand GetCommand(string commandText, CommandType commandType, params SqlParameter[] parameters) {
      return GetCommand(this.ConnectionString, commandText, commandType, parameters);
    }

    public SqlCommand GetCommand(string connectionString, string commandText, params SqlParameter[] parameters) {
      return GetCommand(connectionString, commandText, CommandType.Text, parameters);
    }

    public SqlCommand GetCommand(string commandText, params SqlParameter[] parameters) {
      return GetCommand(this.ConnectionString, commandText, CommandType.Text, parameters);
    }

    #endregion

    #region ExecuteScalar

    public object ExecuteScalar(string connectionString, string commandText, CommandType commandType, params SqlParameter[] parameters) {
      var cmd = GetCommand(connectionString, commandText, commandType, parameters);
      try {
        cmd.Connection.Open();
        return cmd.ExecuteScalar();
      }
      finally {
        cmd.Connection.Close();
      }
    }

    public object ExecuteScalar(string commandText, CommandType commandType, params SqlParameter[] parameters) {
      return ExecuteScalar(ConnectionString, commandText, commandType, parameters);
    }

    public object ExecuteScalar(string connectionString, string commandText, params SqlParameter[] parameters) {
      return ExecuteScalar(connectionString, commandText, CommandType.Text, parameters);
    }

    public object ExecuteScalar(string commandText, params SqlParameter[] parameters) {
      return ExecuteScalar(ConnectionString, commandText, CommandType.Text, parameters);
    }

    #endregion

    #region ExecuteNonQuery

    public int ExecuteNonQuery(string connectionString, string commandText, CommandType commandType, params SqlParameter[] parameters) {
      var cmd = GetCommand(connectionString, commandText, commandType, parameters);
      try {
        OpenConnection();
        return cmd.ExecuteNonQuery();
      }
      finally {
        CloseConnection();
      }
    }


    public int ExecuteNonQuery(string commandText, CommandType commandType, params SqlParameter[] parameters) {
      return ExecuteNonQuery(ConnectionString, commandText, commandType, parameters);
    }

    public int ExecuteNonQuery(string connectionString, string commandText, params SqlParameter[] parameters) {
      return ExecuteNonQuery(connectionString, commandText, CommandType.Text, parameters);
    }

    public int ExecuteNonQuery(string commandText, params SqlParameter[] parameters) {
      return ExecuteNonQuery(ConnectionString, commandText, CommandType.Text, parameters);
    }

    #endregion

    #region ExecuteReader

    public SqlDataReader ExecuteReader(string connectionString, string commandText, CommandType commandType, params SqlParameter[] parameters) {

      var cmd = GetCommand(connectionString, commandText, commandType, parameters);
      cmd.Connection.Open();
      var reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
      return reader;
    }

    public SqlDataReader ExecuteReader(string commandText, CommandType commandType, params SqlParameter[] parameters) {
      return ExecuteReader(ConnectionString, commandText, commandType, parameters);
    }

    public SqlDataReader ExecuteReader(string connectionString, string commandText, params SqlParameter[] parameters) {
      return ExecuteReader(connectionString, commandText, CommandType.Text, parameters);
    }

    public SqlDataReader ExecuteReader(string commandText, params SqlParameter[] parameters) {
      return ExecuteReader(ConnectionString, commandText, CommandType.Text, parameters);
    }

    #endregion

    #region ExecuteNonQuery

    public DataTable GetTable(string connectionString, string commandText, CommandType commandType, params SqlParameter[] parameters) {
      var cmd = GetCommand(connectionString, commandText, commandType, parameters);
      try {
        OpenConnection();
        var reader = cmd.ExecuteReader();
        var table = new DataTable();
        table.Load(reader);
        return table;
      }
      finally {
        CloseConnection();
      }
    }

    public DataTable GetTable(string commandText, CommandType commandType, params SqlParameter[] parameters) {
      return GetTable(ConnectionString, commandText, commandType, parameters);
    }

    public DataTable GetTable(string connectionString, string commandText, params SqlParameter[] parameters) {
      return GetTable(connectionString, commandText, CommandType.Text, parameters);
    }

    public DataTable GetTable(string commandText, params SqlParameter[] parameters) {
      return GetTable(ConnectionString, commandText, CommandType.Text, parameters);
    }

    #endregion

    private void ValidateTransaction() {
      if (!_useSingletone) throw new Exception("Transaction is supported only in singletone mode");
      if (_transaction == null) throw new Exception("There is no active transaction");
    }

    private void OpenConnection() {
      if (!IsTransactionActive)
        GetConnection().Open();
    }

    private void CloseConnection() {
      if (!IsTransactionActive)
        GetConnection().Close();
    }

    public void Dispose() {
      GetConnection().Close();
      GC.SuppressFinalize(this);
    }
  }
}
