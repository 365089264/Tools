using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Data;
using System.Data.SqlClient;
using Aspose.Cells;
namespace StringMatchFunction
{
    class ExportData2019
    {
        private static void Main(string[] args)
        {
            string sql1 = @"SELECT  [影院名称]
      ,[发运方式]
      ,[省份]
      ,[城市]
      ,[地址]
      ,[专资编码]
      ,[院线名称]
      ,[影院省份]
      ,[影院城市]
      ,[EQ_Adress]
      ,[相似度]
  FROM [test].[dbo].[A_ExportExcel20190317]
  order by [专资编码]";
            RegisterLicense();
            var workbook = new Workbook();
            var dt = GetDataTableBySql(sql1);
            var sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;
            int Colnum = dt.Columns.Count;//表格列数 
            int Rownum = dt.Rows.Count;//表格行数 
            //生成行 列名行 
            for (int i = 0; i < Colnum; i++)
            {
                cells[0, i].PutValue(dt.Columns[i].ColumnName);
            }
            //生成数据行 
            for (int i = 0; i < Rownum; i++)
            {
                for (int k = 0; k < Colnum; k++)
                {
                    cells[1 + i, k].PutValue(dt.Rows[i][k].ToString());
                }
            }
            workbook.Save("求表20190317 - v2.xls");
            Console.WriteLine("导出excel完成，结束！");
        }

        public static DataTable GetDataTableBySql(string sql)
        {
            DataSet ds = new DataSet();
            string conStr = "Data Source=10.10.10.80,2433;Initial Catalog=HyFilmCopyForZyNew;UID=sa;PWD=hyby@123;";
            SqlConnection con = new SqlConnection(conStr);
            SqlCommand cmd = new SqlCommand(sql, con);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(ds);
            return ds.Tables[0];

        }
        public static void RegisterLicense()
        {
            var license = new License();
            var commonAssembly = System.Reflection.Assembly.Load("VAV.Common");
            var s = commonAssembly.GetManifestResourceStream("VAV.Common.Aspose.Cells.lic");
            license.SetLicense(s);
        }
    }

}
