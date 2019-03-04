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
    /*
    SELECT ac.ticketsettlementid,'arr.Add('||ac.attachmentid||');',SETT.SETTLENAME,count(0) as cc
    FROM attachment ac
    inner join SETTLEMENT sett on AC.TICKETSETTLEMENTID=sett.SETTLEMENTID
    inner join FEEMONTHUSERUPLOADDETAILS de on sett.SETTLEMENTID=DE.SETTLEMENTID
    where ac.feeyear=2018 and ac.feemonth=10 and de.feemonth=201810 and ac.attatype='monthFee' and ac.isprocess in(4,7) and SETT.SETTLEMODE in(2,3)
    group by ac.ticketsettlementid,ac.attachmentid,SETT.SETTLENAME
    order by cc ; */
    class Program
    {
        static string conStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.1.31.116)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME = orcl11g)));User Id=settle;Password=settle;";
        //string conStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.10.130.23)(PORT=8903))(CONNECT_DATA=(SERVICE_NAME = settle_primary)));User Id=settle;Password=settle;";
        static void Main123(string[] args)
        {
            ArrayList arr = GetAttId1();
            using (OracleConnection conn = new OracleConnection(conStr))
            {

                conn.Open();

                for (int i = 0; i < arr.Count; i++)
                {
                    double ds = 0;
                    int total = 1;
                    for (int j = 0; j < total; j++)
                    {
                        DateTime startDate = DateTime.Now;
                        OracleParameter[] oracleParameters = new OracleParameter[1];
                        oracleParameters[0] = new OracleParameter("p_attachmentId", arr[i]);
                        OracleCommand cmd = new OracleCommand("SP_CHECKCINEMA_MT_BYATTID", conn);
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddRange(oracleParameters);
                        cmd.CommandTimeout = 1800;
                        cmd.ExecuteNonQuery();
                        DateTime endDate = DateTime.Now;
                       ds+=(endDate - startDate).TotalSeconds;
                    }
                    Console.WriteLine(arr[i] + "：" + ds / total);
                }
                conn.Close();
            }
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
        /// 执行存储过程没有返回值
        /// </summary>
        /// <param name="cmdText">存储过程名称</param>
        /// <param name="outParameters">参数</param>
        /// <param name="oracleParameters">所传参数（必须按照存储过程参数顺序）</param>
        /// <param name="strConn">链接字符串</param>
        /// <returns></returns>
        public static void ExecToStoredProcedure(string cmdText, OracleParameter[] oracleParameters, string strConn)
        {
            using (OracleConnection conn = new OracleConnection(strConn))
            {
                OracleCommand cmd = new OracleCommand(cmdText, conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddRange(oracleParameters);
                cmd.CommandTimeout = 1800;
                conn.Open();
                Console.WriteLine("3：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:ffff"));
                cmd.ExecuteNonQuery();
                Console.WriteLine("4：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:ffff"));
                Console.WriteLine(cmd.Parameters["p_returnvalue"].Value.ToString());
                conn.Close();
            }
        }
        public static ArrayList GetAttId1()
        {
            ArrayList arr = new ArrayList();
            arr.Add(13062);
            arr.Add(12973);
            arr.Add(13091);
            arr.Add(13037);
            arr.Add(12920);
            arr.Add(12827);
            arr.Add(12780);
            arr.Add(12684);
            arr.Add(12732);
            arr.Add(12976);
            arr.Add(12848);
            arr.Add(12821);
            arr.Add(12938);
            arr.Add(12878);
            arr.Add(12880);
            arr.Add(12891);
            arr.Add(12974);
            arr.Add(12800);
            arr.Add(12824);
            arr.Add(13086);
            arr.Add(13078);
            arr.Add(12907);
            arr.Add(13058);
            arr.Add(12723);
            arr.Add(12809);
            arr.Add(12760);
            arr.Add(12905);
            arr.Add(12959);
            arr.Add(12637);
            arr.Add(13082);
            arr.Add(13103);
            arr.Add(13003);
            arr.Add(13099);
            arr.Add(12879);
            arr.Add(13039);
            arr.Add(12996);
            arr.Add(12917);
            arr.Add(12993);
            arr.Add(13026);
            arr.Add(12776);
            arr.Add(13089);
            arr.Add(13045);
            arr.Add(13090);
            arr.Add(13084);
            arr.Add(12811);
            arr.Add(12852);
            arr.Add(12983);
            arr.Add(13074);
            arr.Add(13025);
            arr.Add(13098);

            return arr;
        }
    }
}
