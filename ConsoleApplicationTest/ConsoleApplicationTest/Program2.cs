using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using Oracle.ManagedDataAccess.Types;
using OracleCommand = Oracle.ManagedDataAccess.Client.OracleCommand;
using OracleConnection = Oracle.ManagedDataAccess.Client.OracleConnection;
using OracleDataAdapter = Oracle.ManagedDataAccess.Client.OracleDataAdapter;
using OracleParameter = Oracle.ManagedDataAccess.Client.OracleParameter;
using System.Collections;

namespace ConsoleApplicationTest
{
    class Program2
    {
        static string conStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.1.31.114)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME = orcl11g)));User Id=settle;Password=settle;";
        static void Main00(string[] args)
        {
            DateTime t1 = DateTime.Now;
            ArrayList arr = DateFiledMonths("2018-01-01", "2018-12-05");
            decimal total = 0;
            decimal domestic = 0;
            string sql = "SELECT partition_name FROM USER_TAB_PARTITIONS WHERE TABLE_NAME='ZZTICKET' order by PARTITION_POSITION desc";
            DataTable dt = GetDataTableBySqlSettle(conStr, sql);
            Console.WriteLine("rows:" + dt.Rows.Count);
            //string totalSql = "";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (arr.Count == 0) break;
                DataTable dt2 = GetDataTableBySqlSettle(conStr, "select to_char(PLAYDATE,'yyyy-MM') as mindate from zzticket partition(" + dt.Rows[i][0].ToString() + ") where rownum=1");
                if (dt2.Rows.Count > 0 && arr.Contains(dt2.Rows[0][0].ToString()))
                {
                    arr.Remove(dt2.Rows[0][0].ToString());
                    Console.WriteLine(dt.Rows[i][0].ToString() + "：" + dt2.Rows[0][0].ToString());
                    string subSql = "select nvl(sum(boxOffice),0) as BOXOFFICE,nvl(sum(case when substr(FILMSEQCODE,0,3) in('001','002','003') then boxOffice else 0 end),0) as DOMESTIC  from zzticket partition(" + dt.Rows[i][0].ToString() + ") ";
                    ///totalSql += subSql + " union ";
                    DataTable dt3 = GetDataTableBySqlSettle(conStr, subSql);
                    if (dt3.Rows.Count > 0)
                    {
                        total += Convert.ToDecimal(dt3.Rows[0][0]);
                        domestic += Convert.ToDecimal(dt3.Rows[0][1]);
                    }
                }
            }
            //totalSql = totalSql.Substring(0, totalSql.Length - 7);
            TimeSpan ts = DateTime.Now - t1;
            Console.WriteLine(ts.TotalSeconds + " S");
            Console.WriteLine(total + "：" + domestic);
            Console.ReadLine();
        }

        public static DataTable GetDataTableBySqlSettle(string conStr, string sql)
        {
            DataSet ds = new DataSet();

            OracleConnection con = new OracleConnection(conStr);
            OracleCommand cmd = new OracleCommand(sql, con);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            da.Fill(ds);
            return ds.Tables[0];
        }

        /// <summary> 
        /// 获取两个时间段之间的月份 
        /// </summary> 
        /// <param name="startTime">开始月份</param> 
        /// <param name="endTime">结束月份</param> 
        /// <returns>月份字符串</returns> 
        public static ArrayList DateFiledMonths(string startTime, string endTime)
        {
            ArrayList arr = new ArrayList();
            try
            {
                int index = 0;
                string filed = string.Empty;
                DateTime c1 = Convert.ToDateTime(Convert.ToDateTime(startTime).ToString("yyyy-MM"));
                DateTime c2 = Convert.ToDateTime(Convert.ToDateTime(endTime).ToString("yyyy-MM"));
                if (c1 > c2)
                {
                    DateTime tmp = c1;
                    c1 = c2;
                    c2 = tmp;
                }
                while (c2 >= c1)
                {
                    index++;
                    arr.Add(c1.ToString("yyyy-MM"));
                    c1 = c1.AddMonths(1);
                }
                return arr;
            }
            catch { return null; }
        }

    }
}
