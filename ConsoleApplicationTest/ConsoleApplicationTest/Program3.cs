using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Collections;
using System.Data.SqlClient;

namespace ConsoleApplicationTest
{
    class Program3
    {
        static void Main(string[] args)
        {
            execInsert();
            Console.ReadLine();
        }

        public static void execInsert()
        {
            for (int j = 0; j < 1; j++)
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("name");
                for (var i = 0; i < 10000; i++)
                {
                    DataRow dr = dt.NewRow();
                    dr[0] = "wss" + i;
                    dt.Rows.Add(dr);
                }

                string conStr = "Data Source=10.1.31.112;Initial Catalog=test;UID=sa;PWD=Welcome1;";
                using (SqlBulkCopy sqlRevdBulkCopy = new SqlBulkCopy(conStr))//引用SqlBulkCopy  
                {
                    sqlRevdBulkCopy.BulkCopyTimeout = 10000;
                    sqlRevdBulkCopy.DestinationTableName = "UserInfo";//数据库中对应的表名  
                    sqlRevdBulkCopy.NotifyAfter = dt.Rows.Count;//有几行数据  
                    sqlRevdBulkCopy.ColumnMappings.Add("name", "name");
                    sqlRevdBulkCopy.WriteToServer(dt);//数据导入数据库  
                    sqlRevdBulkCopy.Close();//关闭连接  
                }
                Console.WriteLine(j + ":" + dt.Rows.Count + "行");
            }



        }


    }
}
