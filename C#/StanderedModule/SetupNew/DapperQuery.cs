using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Common;
using Dapper;

namespace SetupNew
{
    public class DapperQuery
    {
        private string _connectionString = "YourConnectionStringHere";
        IDbConnection dbConnection;

        public T GetSingleData<T>(string sql , DbParameter db) 
        {
            using (IDbConnection dbConnection = new SqlConnection(_connectionString))
            {
                dbConnection.Open();
                return dbConnection.QuerySingleOrDefault<T>(sql, db);
            }
        }
        public IEnumerable<T> GetMultipleData<T>(string sql, DbParameter db)
        {
            using (IDbConnection dbConnection = new SqlConnection(_connectionString))
            {
                dbConnection.Open();
                return dbConnection.Query<T>(sql, db);
            }
        }
        public int ExecuteSingle(string sql, DbParameter db)
        {
            using (IDbConnection dbConnection = new SqlConnection(_connectionString))
            {
                dbConnection.Open();
                return dbConnection.Execute(sql, db);
            }
        }
        //public int ExecuteMultiple(IEnumerable<string> sql,IEnumerable<DbParameter> db)
        //{
        //    using (IDbConnection dbConnection = new SqlConnection(_connectionString))
        //    {
        //        dbConnection.Open();
        //        return dbConnection.Execute(sql, db);
        //    }
        //}
    }
}
