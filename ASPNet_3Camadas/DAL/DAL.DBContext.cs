using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace DAL
{
    public static class DBContext 
    {
        static public SqlConnection Connection { get; private set; }

       
        public static IDataReader GetData(string pSQLCommand, SqlConnection connection)
        {
            var cmd = prepareCommand(pSQLCommand, DBConnection.Connection);
            return cmd.ExecuteReader();
        }

        public static DataTable GetDataTable(string pSQLCommand)
        {
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(pSQLCommand, DBConnection.Connection);
                DataTable table = new DataTable();
                da.Fill(table);
                return table;
            }
            catch (Exception ex )
            {
                throw new Exceptions.DALExceptionGetDataTable( string.Format("GetDataTable() - Error {0} : ",ex.Message));
            }
        }

        public static DataSet GetDataSet(string pSQLCommand, SqlConnection connection)
        {
            //var cmd = prepareCommand(pSQLCommand, connection);
            return ExecuteDataSet(pSQLCommand, connection);
        }

        private static DataSet ExecuteDataSet(string pSQLCommand, SqlConnection connection)
        {
            DataSet ds = ds = null;
            try
            {
                SqlDataAdapter da = null;
                IDbCommand dbCmd = new SqlCommand(pSQLCommand, connection);
                da = new SqlDataAdapter();
                da.SelectCommand = (SqlCommand)dbCmd;
                ds = new DataSet();
                da.Fill(ds);
            }
            catch (Exception ex)
            {
                throw new Exceptions.DALExceptionExecuteReader(ex.Message);
            }
            return ds;
        }

        public static int Execute(string pSQLCommand, SqlConnection connection, string[] paramName = null, string[] paramValues = null)
        {
            var cmd = prepareCommand(pSQLCommand, connection, paramName, paramValues);

            var intReturn =  cmd.ExecuteNonQuery();
            cmd = null; cmd.Dispose();
            return intReturn;
        }

        public static string ExecuteScalar(string pSQLCommand, SqlConnection connection, string[] paramName = null, string[] paramValues = null)
        {
            var cmd = prepareCommand(pSQLCommand, connection, paramName, paramValues);
            var intReturn = cmd.ExecuteScalar();
            cmd = null; cmd.Dispose();
            return intReturn.ToString();
        }

        public static int CRUD(string pSQLCommand, SqlConnection connection, string[] paramName= null, string[] paramValues=null)
        {
            return Execute(pSQLCommand, connection, paramName, paramValues);
        }

        static public SqlCommand prepareCommand(string pSQLCommand, string[] paramName = null, string[] paramValues = null)
        {
            try
            {
                var cmd = new SqlCommand() { CommandText = pSQLCommand, Connection = Connection };
                cmd = FillParameters(cmd, paramName, paramValues);
                return cmd;
            }
            catch
            {
                throw new DAL.Exceptions.DALExceptionCommand();
            }
        }
        static public SqlCommand prepareCommand(string pSQLCommand, SqlConnection connection,string[] paramName = null, string[] paramValues = null)
        {
            return prepareCommand(pSQLCommand, DBConnection.Connection);
        }

        private static SqlCommand FillParameters(SqlCommand cmd ,string[] paramName, string[] paramValues)
        {
            if (paramName != null)
            {
                for (int i = 0; i < paramName.Length; i++)
                {
                    cmd.Parameters.Add(new SqlParameter(paramName[i], paramValues[i]));
                }
            }
            return cmd;
        }
    }
}
