using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Data.SQLite;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace Sofar.DataBaseHelper
{
    public class SqlLiteHelper : IDataBase
    {
        /// <summary>
        /// 连接字符串
        /// </summary>
        public string str { get; set; }
        public SQLiteConnection conn { get; set; }
        public SQLiteCommand Command { get; set; }
        public Int32 CommandTimeout { get; set; }

        /// <summary>
        /// 命令超时时间(300秒)
        /// </summary>
        public const int COMMANDTIMEOUT = 300;
        public SqlLiteHelper(string strConn)
        {
            str = strConn;
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
                    conn = new SQLiteConnection(str);
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
                    conn.Dispose();
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
        public object ExecuteScalar(string sql, params DbParameter[] param)
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
                using (SQLiteDataAdapter sda = new SQLiteDataAdapter(sql, str))
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
                Debug.WriteLine("error:"+sql);
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
                param.Add(new SQLiteParameter("@" + key, insertData[key]));
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
                param.Add(new SQLiteParameter("@" + key, insertData[key]));
                point = ",";
            }
            string sql = string.Format("UPDATE {0} SET {1} WHERE {2}", tbName, kvStr, where);

            return ExecuteNonQuery(sql, param.ToArray());

        }

        public int ExecuteMutliQuery(List<string> sqlList)
        {
            int res = 0;

            CreateDatabaseConnection();
            using (SQLiteTransaction tran = conn.BeginTransaction())
            {
                try
                {
                    SQLiteCommand cmd = conn.CreateCommand();
                    cmd.Connection = conn;
                    foreach (string sql in sqlList)
                    {
                        cmd.CommandText = sql;
                        cmd.ExecuteNonQuery();
                    }
                    tran.Commit();
                    res = 1;
                }
                catch (Exception ex)
                {
                    res = -1;
                    tran.Rollback();
                    throw ex;
                }
                finally
                {
                    //Conn.Close();
                }
            }

            return res;
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

        #region 把DataTable中数据批量导入新的sqlite的db文件中
        public int ExecuteMutliQuery(string commandText, DataTable dtData)
        {
            int res = 0;
            CreateDatabaseConnection();
            using (SQLiteTransaction dbTrans = conn.BeginTransaction())
            {
                try
                {
                    foreach (DataRow row in dtData.Rows)
                    {
                        res += ExecuteNonQuery(dbTrans, commandText, row.ItemArray);
                    }
                    dbTrans.Commit();
                }
                catch (Exception ex)
                {
                    res = -1;
                    dbTrans.Rollback();
                    throw ex;
                }
                finally
                {
                    //Conn.Close();
                }
            }
            return res;
        }
        public int ExecuteNonQuery(SQLiteTransaction transaction, string commandText, params object[] paramList)
        {
            if (transaction == null) throw new ArgumentNullException("transaction is null");
            if (transaction != null && transaction.Connection == null) throw new ArgumentException("The transaction was rolled back or committed,please provide an open transaction.", "transaction");
            using (IDbCommand cmd = transaction.Connection.CreateCommand())
            {
                cmd.CommandText = commandText;
                AttachParameters((SQLiteCommand)cmd, cmd.CommandText, paramList);
                if (transaction.Connection.State == ConnectionState.Closed)
                    transaction.Connection.Open();
                int result = cmd.ExecuteNonQuery();
                return result;
            }
        }

        /// <summary>
        /// 增加参数到命令（自动判断类型）
        /// </summary>
        /// <param name="commandText">命令语句</param>
        /// <param name="paramList">object参数列表</param>
        /// <returns>返回SQLiteParameterCollection参数列表</returns>
        private SQLiteParameterCollection AttachParameters(SQLiteCommand cmd, string commandText, params object[] paramList)
        {
            if (paramList == null || paramList.Length == 0) return null;

            SQLiteParameterCollection coll = cmd.Parameters;
            string parmString = commandText.Substring(commandText.IndexOf("@"));
            // pre-process the string so always at least 1 space after a comma.
            parmString = parmString.Replace(",", " ,");
            // get the named parameters into a match collection
            string pattern = @"(@)\S*(.*?)\b";
            Regex ex = new Regex(pattern, RegexOptions.IgnoreCase);
            MatchCollection mc = ex.Matches(parmString);
            string[] paramNames = new string[mc.Count];
            int i = 0;
            foreach (Match m in mc)
            {
                paramNames[i] = m.Value;
                i++;
            }

            // now let's type the parameters
            int j = 0;
            Type t = null;
            foreach (object o in paramList)
            {
                t = o.GetType();

                SQLiteParameter parm = new SQLiteParameter();
                switch (t.ToString())
                {
                    case ("DBNull"):
                    case ("Char"):
                    case ("SByte"):
                    case ("UInt16"):
                    case ("UInt32"):
                    case ("UInt64"):
                        throw new SystemException("Invalid data type");
                    case ("System.DBNull"):
                        //判断是否为非空属性
                        if (paramNames[j].Equals("@ID") || paramNames[j].Equals("@NameZh") || paramNames[j].Equals("@DataType"))
                        {
                            throw new SystemException("NameZh和DataType字段值不可为空");
                        }
                        else
                        {
                            parm.DbType = DbType.String;
                            parm.ParameterName = paramNames[j];
                            parm.Value = null;
                            coll.Add(parm);
                        }
                        break;
                    case ("System.String"):
                        parm.DbType = DbType.String;
                        parm.ParameterName = paramNames[j];
                        parm.Value = (string)paramList[j];
                        coll.Add(parm);
                        break;

                    case ("System.Byte[]"):
                        parm.DbType = DbType.Binary;
                        parm.ParameterName = paramNames[j];
                        parm.Value = (byte[])paramList[j];
                        coll.Add(parm);
                        break;

                    case ("System.Int32"):
                        parm.DbType = DbType.Int32;
                        parm.ParameterName = paramNames[j];
                        parm.Value = (int)paramList[j];
                        coll.Add(parm);
                        break;

                    case ("System.Int64"):
                        parm.DbType = DbType.Int32;
                        parm.ParameterName = paramNames[j];
                        parm.Value = Convert.ToInt32(paramList[j]);
                        coll.Add(parm);
                        break;

                    case ("System.Boolean"):
                        parm.DbType = DbType.Boolean;
                        parm.ParameterName = paramNames[j];
                        parm.Value = (bool)paramList[j];
                        coll.Add(parm);
                        break;

                    case ("System.DateTime"):
                        parm.DbType = DbType.DateTime;
                        parm.ParameterName = paramNames[j];
                        parm.Value = Convert.ToDateTime(paramList[j]);
                        coll.Add(parm);
                        break;

                    case ("System.Double"):
                        parm.DbType = DbType.Double;
                        parm.ParameterName = paramNames[j];
                        parm.Value = Convert.ToDouble(paramList[j]);
                        coll.Add(parm);
                        break;

                    case ("System.Single"):
                    case ("System.Decimal"):
                        parm.DbType = DbType.Decimal;
                        parm.ParameterName = paramNames[j];
                        parm.Value = Convert.ToDecimal(paramList[j]);
                        coll.Add(parm);
                        break;

                    case ("System.Guid"):
                        parm.DbType = DbType.Guid;
                        parm.ParameterName = paramNames[j];
                        parm.Value = (System.Guid)(paramList[j]);
                        coll.Add(parm);
                        break;

                    case ("System.Object"):

                        parm.DbType = DbType.Object;
                        parm.ParameterName = paramNames[j];
                        parm.Value = paramList[j];
                        coll.Add(parm);
                        break;

                    default:
                        throw new SystemException("Value is of unknown data type");

                } // end switch

                j++;
            }
            return coll;
        }
        #endregion
    }
}