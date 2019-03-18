using System;
using System.Configuration;
using System.Data;
using Oracle.ManagedDataAccess.Client;
namespace ExcelExportFilmNo
{
    public class OracleDBHelper
    {
        private static readonly string connStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.1.31.114)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME = orcl11g)));User Id=settle;Password=settle;";
        public OracleDBHelper()
        {
        }
        /// <summary>
        /// 执行存储过程获取带有Out的游标数据集
        /// </summary>
        /// <param name="cmdText">存储过程名称</param>
        /// <param name="outParameters">输出的游标名</param>
        /// <param name="oracleParameters">所传参数（必须按照存储过程参数顺序）</param>
        /// <param name="strConn">链接字符串</param>
        /// <returns></returns>
        public static DataTable ExecToStoredProcedureGetTable(string storedProcedName, OracleParameter[] oracleParameters)
        {
            using (OracleConnection conn = new OracleConnection(connStr))
            {
                OracleCommand cmd = new OracleCommand(storedProcedName, conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddRange(oracleParameters);
                OracleDataAdapter oda = new OracleDataAdapter(cmd);
                conn.Open();
                DataSet ds = new DataSet();
                oda.Fill(ds);
                conn.Close();
                return ds.Tables[0];
            }
        }
        public static decimal ExecuteStorageWithRevalue(string storName, int returnIndex, params OracleParameter[] paras)
        {
            using (var conn = new OracleConnection(connStr))
            {
                conn.Open();
                using (var cmd = new OracleCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = storName;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 1800;
                    if (paras != null)
                    {
                        cmd.Parameters.AddRange(paras);
                    }
                    cmd.ExecuteScalar();
                }
            }
            if (paras[returnIndex].Value.ToString() == "null") return 0;
            return Convert.ToDecimal(paras[returnIndex].Value.ToString());
        }
        public static void ExecuteStorageWithoutRevalue(string storName, OracleParameter[] paras)
        {
            using (var conn = new OracleConnection(connStr))
            {
                conn.Open();
                using (var cmd = new OracleCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = storName;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 1800;
                    cmd.Parameters.AddRange(paras);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public static DataSet GetDataSetBySp(string inName, OracleParameter[] inParms)
        {
            using (var conn = new OracleConnection(connStr))
            {
                using (var cmd = new OracleCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = inName;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 1800;

                    if (inParms != null)
                        cmd.Parameters.AddRange(inParms);

                    var da = new OracleDataAdapter(cmd);
                    var ds = new DataSet();
                    da.Fill(ds);
                    return ds;
                }
            }
        }

        public static DataTable GetDataTableBySql(string sql, params OracleParameter[] paras)
        {
            DataTable tb = new DataTable();
            using (var conn = new OracleConnection(connStr))
            {
                using (var cmd = new OracleCommand(sql, conn))
                {
                    if (paras != null)
                    {
                        cmd.Parameters.AddRange(paras);
                    }
                    cmd.CommandTimeout = 1800;
                    var sda = new OracleDataAdapter(cmd);
                    sda.Fill(tb);
                }
            }
            return tb;
        }
        public static object ExecuteScaler(string sql, params OracleParameter[] paras)
        {
            object o = null;
            using (var conn = new OracleConnection(connStr))
            {
                conn.Open();

                using (var cmd = new OracleCommand(sql, conn))
                {
                    cmd.Parameters.AddRange(paras);
                    o = cmd.ExecuteScalar();
                }
            }
            return o;
        }
    }
}