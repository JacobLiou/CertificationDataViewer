using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;

namespace Sofar.DataBaseHelper
{
    public class SqlServerHelper:IDataBase
    {
        /// <summary>
        /// 连接字符串
        /// </summary>
        public string str { get; set; }
        public SqlConnection conn { get; set; }
        public SqlCommand Command { get; set; }
        public Int32 CommandTimeout { get; set; }
        /// <summary>
        /// 命令超时时间(300秒)
        /// </summary>
        public const int COMMANDTIMEOUT = 300;

        public SqlServerHelper(string strConn)
        {
            str = strConn;
        }

        public SqlServerHelper(string strCon, bool IsEncry)
        {
            try
            {
                string strcon = ConfigurationManager.ConnectionStrings[strCon].ToString();
                if (IsEncry)
                    strcon = Encrypt.DecryptPassword(strcon);
                str = strcon;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 新建数据库连接
        /// </summary>
        public void CreateDatabaseConnection()
        {
            OpenConnection();
        }

        /// <summary>
        /// 打开数据库连接
        /// </summary>
        /// <returns></returns>
        public bool OpenConnection()
        {
            try
            {
                if (conn == null)
                {
                    //创建Connection对像
                    conn = new SqlConnection(str);
                }
                //打开连接
                if (conn != null
                    && conn.State != ConnectionState.Open
                    && conn.State != ConnectionState.Connecting
                    && conn.State != ConnectionState.Executing
                    && conn.State != ConnectionState.Fetching)
                {
                    conn.Open();
                }
            }
            catch (Exception ex)
            {
                if (conn != null)
                    conn.Dispose();
                throw ex;
            }
            return true;
        }


        /// <summary>
        /// 关闭数据库连接
        /// </summary>
        /// <returns></returns>
        public bool CloseConnection()
        {
            try
            {
                if (conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }
                if (Command != null)
                {
                    Command.Dispose();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        /// <summary>
        /// 初始化数据库命令
        /// </summary>
        public void CreateDatabaseCommand()
        {
            try
            {
                Command = conn.CreateCommand();
                Command.CommandTimeout = CommandTimeout < 0 ? COMMANDTIMEOUT : CommandTimeout;
            }
            catch (Exception ex)
            {
                if (Command != null)
                    Command.Dispose();
                throw ex;
            }
        }

        /// <summary>
        /// 增删改
        /// 20180723
        /// </summary>
        /// <param name="sql">sql语句</param>
        /// <param name="param">sql参数</param>
        /// <returns>受影响的行数</returns>
        public int ExecuteNonQuery(string sql, params DbParameter[] param)
        {
            try
            {
                CreateDatabaseConnection();
                CreateDatabaseCommand();
                Command.CommandText = sql;

                if (param != null)
                {
                    Command.Parameters.AddRange(param);
                }
                return Command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 查询
        /// 20180723
        /// </summary>
        /// <param name="sql">sql语句</param>
        /// <param name="param">sql参数</param>
        /// <returns>首行首列</returns>
        public object ExecuteScalar(string sql, params SqlParameter[] param)
        {
            try
            {
                CreateDatabaseConnection();
                CreateDatabaseCommand();
                if (param != null)
                {
                    Command.Parameters.AddRange(param);
                }
                return Command.ExecuteScalar();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        /// <summary>
        /// 查询
        /// </summary>
        /// <param name="sql">sql语句</param>
        /// <param name="param">sql参数</param>
        /// <returns>一个表</returns>
        public DataTable ExecuteTable(string sql, params DbParameter[] param)
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlDataAdapter sda = new SqlDataAdapter(sql, str))
                {
                    if (param != null)
                    {
                        sda.SelectCommand.Parameters.AddRange(param);
                    }
                    sda.Fill(dt);
                }
                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        /// <summary>
        /// 分页查询
        /// </summary>
        /// <param name="tbName">表名</param>
        /// <param name="fields">查询需要的字段名："id, name, age"</param>
        /// <param name="where">查询条件："id = 1"</param>
        /// <param name="orderBy">排序："id desc"</param>
        /// <param name="limit">分页："0,10"</param>
        /// <param name="param">sql参数</param>
        /// <returns>受影响行数</returns>
        public DataTable QueryTable(string tbName, string fields = "*", string where = "1", string orderBy = "", string limit = "", params DbParameter[] param)
        {
            //排序
            if (orderBy != "")
            {
                orderBy = "ORDER BY " + orderBy;//Deom: ORDER BY id desc
            }

            //分页
            if (limit != "")
            {
                limit = "LIMIT " + limit;//Deom: LIMIT 0,10
            }

            string sql = string.Format("SELECT {0} FROM {1} WHERE {2} {3} {4}", fields, tbName, where, orderBy, limit);

            //return sql;
            return ExecuteTable(sql, param);

        }



        /// <summary>
        /// 数据插入
        /// </summary>
        /// <param name="tbName">表名</param>
        /// <param name="insertData">需要插入的数据字典</param>
        /// <returns>受影响行数</returns>
        public int ExecuteInsert(string tbName, Dictionary<String, String> insertData)
        {
            string point = "";//分隔符号(,)
            string keyStr = "";//字段名拼接字符串
            string valueStr = "";//值的拼接字符串

            List<DbParameter> param = new List<DbParameter>();
            foreach (string key in insertData.Keys)
            {
                keyStr += string.Format("{0} {1}", point, key);
                valueStr += string.Format("{0} @{1}", point, key);
                param.Add(new SqlParameter("@" + key, insertData[key]));
                point = ",";
            }
            string sql = string.Format("INSERT INTO {0} ({1}) VALUES({2})", tbName, keyStr, valueStr);

            //return sql;
            return ExecuteNonQuery(sql, param.ToArray());

        }

        /// <summary>
        /// 修改
        /// </summary>
        /// <param name="tbName">表名</param>
        /// <param name="where">更新条件：id=1</param>
        /// <param name="insertData">需要更新的数据</param>
        /// <returns>受影响行数</returns>
        public int ExecuteUpdate(string tbName, string where, Dictionary<String, String> insertData)
        {
            string point = "";//分隔符号(,)
            string kvStr = "";//键值对拼接字符串(Id=@Id)

            List<DbParameter> param = new List<DbParameter>();
            foreach (string key in insertData.Keys)
            {
                kvStr += string.Format("{0} {1}=@{2}", point, key, key);
                param.Add(new SqlParameter("@" + key, insertData[key]));
                point = ",";
            }
            string sql = string.Format("UPDATE {0} SET {1} WHERE {2}", tbName, kvStr, where);

            return ExecuteNonQuery(sql, param.ToArray());

        }

        public void Dispose()
        {
            if (conn != null)
            {
                conn.Close();
                conn.Dispose();
            }
            if (Command != null)
            {
                Command.Dispose();
            }
        }
    }
}