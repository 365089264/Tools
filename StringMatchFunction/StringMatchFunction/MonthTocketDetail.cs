using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using Oracle.ManagedDataAccess.Types;
using OracleCommand = Oracle.ManagedDataAccess.Client.OracleCommand;
using OracleConnection = Oracle.ManagedDataAccess.Client.OracleConnection;
using OracleDataAdapter = Oracle.ManagedDataAccess.Client.OracleDataAdapter;
using OracleParameter = Oracle.ManagedDataAccess.Client.OracleParameter;
using Aspose.Cells;

namespace StringMatchFunction
{
    class MonthTocketDetail
    {
        static bool isExportAllFenduan=false;
        static string feemonth = "201806";
        static void Main(string[] args)
        {
            if (DateTime.Now.Year != 2018 && DateTime.Now.Month != 7) {
                return;
            }
            Console.Write("输入月份（yyyymm）：");
            var mon = Console.ReadLine();
            if (!string.IsNullOrEmpty(mon) && mon.Length==6)
            {
                feemonth = mon;
            }
            else {
                return;
            }
            List<SheetMap> list = new List<SheetMap>();
            Console.Write("合计比较是否导出（y/n）：");
            var str=Console.ReadLine();
            if (str.ToLower() == "y")
            {
                list.Add(new SheetMap() { sheetName = "合计比较", isExport = true });
            }
            Console.Write("明细比例（y/n）：");
            str = Console.ReadLine();
            if (str.ToLower() == "y")
            {
                list.Add(new SheetMap() { sheetName = "明细比例", isExport = true });
            }
            Console.WriteLine("分段详细 a代表全部分段信息，b代表不一致分段信息，n代表不导出分段信息");
            Console.Write("分段详细（a/b/n）：");
            str = Console.ReadLine();
            if (str.ToLower() == "a" || str.ToLower() == "b")
            {
                if (str.ToLower() == "a") isExportAllFenduan = true;
                list.Add(new SheetMap() { sheetName = "分段详细", isExport = true });
            }
            Console.Write("结算系统没有的影片（y/n）：");
            str = Console.ReadLine();
            if (str.ToLower() == "y")
            {
                list.Add(new SheetMap() { sheetName = "结算系统没有的影片", isExport = true });
            }
            RegisterLicense();
            var workbook = new Workbook();
            //var list = GetSheetMap();
            int sheetIndex=0;
            foreach (var item in list)
            {
                DataTable dt = new DataTable();
                if (!item.isExport) continue;
                Console.WriteLine(item.sheetName + "开始导出.......");
                switch (item.sheetName) {
                    case "合计比较":
                        dt = GetTicketTotal();
                        break;
                    case "明细比例":
                        dt = GetTicketDetail();
                        break;
                    case "分段详细":
                        dt = GetTicketFenduanDetail();
                        break;
                    case "结算系统没有的影片":
                        dt = GetTicketNoFilm();
                        break;

                }
                if (dt.Rows.Count == 0) continue;
                if (sheetIndex>0)
                {
                 workbook.Worksheets.Add();
                }
                var sheet = workbook.Worksheets[sheetIndex];
                sheet.Name = item.sheetName;
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
                Console.WriteLine(item.sheetName+"导出成功！");
                sheetIndex++;
            }
            string fileName = feemonth+"月结算分段明细.xls";
            workbook.Save(fileName);
            Console.WriteLine("导出excel完成，结束！");
        }
         public static DataTable GetTicketFenduanDetail()
         {
             string sqlAllFilmNo = "select de.THEATRECODE,de.filmno,de.ROWINDEX from feemonthuseruploaddetails de where de.FEEMONTH='"+feemonth+"' order by de.THEATRECODE,de.ROWINDEX,de.FORMATFILMNO";
             string sqlTicketSubDetail = @"select  de.THEATRECODE as 影院编码,de.THEATRENAME as 影院名称,de.ROWINDEX as 序号,de.FILMNO,de.FILMNAME as 影片编码,de.SETTLEMENTID as 结算单位ID,seq.filmid as 影片ID,iss.PLAYSTARTTIME as 分段起始时间,iss.PLAYENDTIME as 分段结束时间,iss.ratiovalue/100  as 系统分账比例,vi.FILMVERSIONNAME as 版本信息
,(case when iss.ratiotype=0 then '自购' else '租赁' end) as 比例类型
,(case when exists(select iss_hall.FILMISSUEID from FILMISSUE_HALL iss_hall where iss.FILMISSUEID=iss_hall.FILMISSUEID and iss_hall.deleted=0) then '是' else '否' end ) as 是否特殊比例
from feemonthuseruploaddetails de
INNER JOIN FILMVERSION vi on de.FILMVERSIONTYPE=vi.FILMVERSIONTYPE
inner join FILMSEQ seq on de.FORMATFILMNO=seq.FILMSEQCODE and de.FILMVERSIONTYPE=seq.FILMVERSIONTYPE
inner join FILMISSUE iss on iss.SETTLEMENTID=de.SETTLEMENTID and iss.filmid=seq.filmid and iss.filmversiontype=seq.FILMVERSIONTYPE AND iss.ratiotype=de.ratiotype
 where   ";
             sqlTicketSubDetail += " de.FEEMONTH='" + feemonth + "' and ";
             if (!isExportAllFenduan) sqlTicketSubDetail += " (de.sysfzbl is null or de.sysfzbl<>de.fzbl) and";
             string sqlWhere = @" and  iss.deleted=0 and
 ((iss.PLAYSTARTTIME<=de.STARTDATE and iss.PLAYENDTIME>=de.STARTDATE) 
or  (iss.PLAYSTARTTIME<=de.ENDDATE and iss.PLAYENDTIME>=de.STARTDATE)
or (iss.PLAYSTARTTIME>=de.STARTDATE and iss.PLAYENDTIME<=de.ENDDATE)
)
order by de.THEATRECODE,de.ROWINDEX,de.FORMATFILMNO,seq.filmid,iss.PLAYSTARTTIME ";
             DataTable dtFilmNo = GetDataTableBySql2(sqlAllFilmNo);
             DataTable dtTicketSubDetail = new DataTable();
             DataTable dtTicketAllDetail = new DataTable();
             for (int i = 0; i < dtFilmNo.Rows.Count; i++)
             {
                 dtTicketSubDetail = GetDataTableBySql2(sqlTicketSubDetail + "  de.ROWINDEX=" + dtFilmNo.Rows[i][2].ToString() +
                        " and de.THEATRECODE='" + dtFilmNo.Rows[i][0].ToString() + "' and de.filmno='"
                        + dtFilmNo.Rows[i][1].ToString() + "' " + sqlWhere);
                 if (dtTicketAllDetail.Rows.Count == 0 && dtTicketSubDetail.Rows.Count != 0)
                 {
                     dtTicketAllDetail = dtTicketSubDetail;
                 }
                 else
                 {
                     foreach (DataRow dr in dtTicketSubDetail.Rows)
                     {
                         DataRow dr1 = dtTicketAllDetail.NewRow();
                         for (int j = 0; j < dtTicketSubDetail.Columns.Count; j++)
                         {
                             dr1[j] = dr[j].ToString();
                         }
                         dtTicketAllDetail.Rows.Add(dr1);
                     }
                 }
             }
             return dtTicketAllDetail;
         }
         public static DataTable GetTicketTotal()
         {
             string sql = @"select distinct des.theatrename as 影院名称,des.theatrecode as 影院编码,de.totalpf as 总票房,de.jpf as 净票房,de.fzpk as 分账片款,de.sysfzpk as 系统分账片款,de.ismatchfzpk as 分账金额差值,de.ismatchfzbl as 是否一致,b.SETTLENAME  as 结算单位名称
from feemonthuseruploadtotal de 
inner join  feemonthuseruploaddetails des on de.SETTLEMENTID=des.SETTLEMENTID
inner join settlement b on de.SETTLEMENTID=b.SETTLEMENTID";
sql+=" where de.FEEMONTH='"+feemonth+"' order by des.THEATRECODE";
             DataTable dt = GetDataTableBySql2(sql);
             return dt;
         }
         public static DataTable GetTicketDetail()
         {
             string sql = @"select de.rowindex as 序号,de.filmno as 影片编码,de.filmname as 影片名称,de.theatrename as 影院名称,de.theatrecode as 影院编码,de.devicebelong as 设备归属,de.startdate as 开始日期,de.enddate as 结束日期,de.totalcc as 总场次,de.totalrc as 总人次,de.totalpf as 总票房,de.dianzzjj as 电影专项基金,de.zzrate as 增值税率,de.sj as 税金,de.jpf as 净票房,de.fzbl as 分账比例,de.fzpk as 分账片款,de.sysfzbl as 系统分账比例,de.sysfzpk as 系统分账片款,de.ismatchfzbl as 分账比例差值,de.ismatchfzpk as 分账金额差值,b.SETTLENAME  as 结算单位名称
,(case when seq.filmid is not null then '是' else '否' end) as 系统查询到影片
,(case when v.ratioTotal is null  then 0 else v.ratioTotal end) as 不同的普通比例值数量
,(case when v.ratioTotal is null  then 0 else v.specialratioTotal end) as 不同的特殊比例值数量
from feemonthuseruploaddetails de 
inner join settlement b on de.SETTLEMENTID=b.SETTLEMENTID
left join FILMSEQ seq on de.FORMATFILMNO=seq.FILMSEQCODE and de.FILMVERSIONTYPE=seq.FILMVERSIONTYPE
left join V_feemonthuploadratiovalue v on de.THEATRECODE=v.THEATRECODE and de.ROWINDEX=v.ROWINDEX and de.FORMATFILMNO=v.FORMATFILMNO ";
sql+="where de.FEEMONTH='"+feemonth+"' order by de.THEATRECODE,de.ROWINDEX,de.FILMNAME";
             DataTable dt = GetDataTableBySql2(sql);
             return dt;
         }
         public static DataTable GetTicketNoFilm()
         {
             string sql = @"select distinct vv.FILMVERSIONNAME as 版本类型,de.FORMATFILMNO as 影片排次号,de.filmname as 影片名称
from feemonthuseruploaddetails de 
inner join V_FILMFORMATE vv on vv.filmcode=de.FORMATFILMNO
inner join settlement b on de.SETTLEMENTID=b.SETTLEMENTID
left join FILMSEQ seq on de.FORMATFILMNO=seq.FILMSEQCODE and de.FILMVERSIONTYPE=seq.FILMVERSIONTYPE
 where seq.filmid is  null  ";
             sql += " and de.FEEMONTH='" + feemonth + "'";
             DataTable dt = GetDataTableBySql2(sql);
             return dt;
         }
         public static List<SheetMap> GetSheetMap()
         {
             List<SheetMap> list = new List<SheetMap>();
             list.Add(new SheetMap() { sheetName = "合计比较", isExport = true });
             list.Add(new SheetMap() { sheetName = "明细比例", isExport = true });
             list.Add(new SheetMap() { sheetName = "分段详细", isExport = false });
             list.Add(new SheetMap() { sheetName = "结算系统没有的影片", isExport = true });
             return list;
         }
        public static string GetFileJson(string filepath)
        {
            string json = string.Empty;
            using (FileStream fs = new FileStream(filepath, FileMode.Open, System.IO.FileAccess.Read, FileShare.ReadWrite))
            {
                using (StreamReader sr = new StreamReader(fs, Encoding.GetEncoding("utf-8")))
                {
                    json = sr.ReadToEnd().ToString();
                }
            }
            return json;
        }
        public static DataTable GetDataTableBySql2(string sql)
        {
            DataSet ds = new DataSet();
            //string conStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=124.207.105.120)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME = orcl11g.us.oracle.com)));User Id=settle;Password=settle;";
            string conStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.10.130.23)(PORT=8903))(CONNECT_DATA=(SERVICE_NAME = settle_primary)));User Id=settle;Password=settle;";
            OracleConnection con = new OracleConnection(conStr);
            OracleCommand cmd = new OracleCommand(sql, con);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            da.Fill(ds);
            return ds.Tables[0];

        }

        // 从一个Json串生成对象信息
        public static object JsonToObject(string jsonString, object obj)
        {
            var serializer = new DataContractJsonSerializer(obj.GetType());
            var mStream = new MemoryStream(Encoding.UTF8.GetBytes(jsonString));
            return serializer.ReadObject(mStream);
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
