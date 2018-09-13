using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;

namespace 补充结算一览表
{
    public class SqlServerHelper
    {
        private static readonly string conStr = ConfigurationManager.AppSettings["issueCon"];
        public static void ExecSqlIssue(string sql)
        {
            using (SqlConnection con = new SqlConnection(conStr))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.CommandTimeout = 1800;
                cmd.ExecuteNonQuery();
            }
        }

        public static DataTable GetDataTableBySql(string sql)
        {
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection(conStr);
            SqlCommand cmd = new SqlCommand(sql, con);
            cmd.CommandTimeout = 1800;
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(ds);
            return ds.Tables[0];
        }
        public static void InsertDatatableIssue(DataTable dt, string tablename)
        {
            DataSet ds = new DataSet();
            string conStr = "Data Source=10.10.10.80,2433;Initial Catalog=HyFilmCopyForZyNew;UID=sa;PWD=hyby@123;";
            using (SqlConnection con = new SqlConnection(conStr))
            {
                SqlBulkCopy bulkCopy = new SqlBulkCopy(con);
                bulkCopy.DestinationTableName = tablename;
                bulkCopy.BatchSize = dt.Rows.Count;
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    bulkCopy.ColumnMappings.Add(dt.Columns[i].ColumnName, dt.Columns[i].ColumnName);
                }
                con.Open();
                bulkCopy.WriteToServer(dt);
            }

        }
    }
}
