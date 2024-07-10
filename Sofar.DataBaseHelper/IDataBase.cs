using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sofar.DataBaseHelper
{
    public interface IDataBase : IDisposable
    {
        Int32 CommandTimeout { get; set; }
        void CreateDatabaseConnection();
        void CreateDatabaseCommand();
        int ExecuteNonQuery(string sql, params DbParameter[] param);
        DataTable ExecuteTable(string sql, params DbParameter[] param);
        int ExecuteUpdate(string tbName, string where, Dictionary<String, String> insertData);
        int ExecuteInsert(string tbName, Dictionary<String, String> insertData); 
        DataTable QueryTable(string tbName, string fields = "*", string where = "1", string orderBy = "", string limit = "", params DbParameter[] param);
    }
}
