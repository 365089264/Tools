using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using Oracle.ManagedDataAccess.Client;
using Oracle.ManagedDataAccess.Types;
using OracleCommand = Oracle.ManagedDataAccess.Client.OracleCommand;
using OracleConnection = Oracle.ManagedDataAccess.Client.OracleConnection;
using OracleDataAdapter = Oracle.ManagedDataAccess.Client.OracleDataAdapter;
using OracleParameter = Oracle.ManagedDataAccess.Client.OracleParameter;
using Aspose.Cells;

namespace ConsoleTestConn23Settle
{
    class Program
    {
        static void Main(string[] args)
        {
            string startdate = "2018-07-01";
            string enddate = "2018-08-01";
            ExecSqlIssue("drop table A_FilmIssue_Issue");
            ExecSqlIssue("truncate table A_FilmIssue_Settle");
            ExecSqlIssue(string.Format(@"select iss.DivideID 一览表ID
,iss.SettleID 结算单位ID
,sett.SettleName 结算单位名称
,fi.filmname 影片名称,jfpv.VersionName 影片版本,pa.FilmNum 影片编码,iss.RationValue 比例值,(case when iss.SettleTypeNo=0 then '自购' else '租赁' end) 比例类型
,CONVERT(varchar(10), iss.StartTime, 120) 一览表开始时间,CONVERT(varchar(10), iss.EndTime, 120) 一览表结束时间
,sum(case when fh.DivideID is null then 0 else 1 end) 特殊设备总数
into A_FilmIssue_Issue
from JS_IssuedList iss
inner join JS_Publish_Batch tag  on iss.sectionid=tag.sectionid
inner join JS_SettleCompany sett on iss.SettleID=sett.SettleID
inner join Publish_Batch pb on tag.SystemNum=pb.SystemNum
inner join Film_FilmInfo fi on pb.filmno=fi.filmno
left join JS_FilmVersion jfpv on iss.Versiontype=jfpv.Versiontype
left join JS_SpecialValueHall fh on iss.DivideID=fh.DivideID and fh.DelFlag=0
left join Publish_FilmPermitAdd pa on pb.filmno=pa.filmno and jfpv.VersionID=pa.VersionID
where tag.EndTime>='{0}'  and tag.StartTime<='{1}' and   iss.DelFlag=0
group by iss.DivideID ,iss.SettleID,sett.SettleName ,fi.filmname,jfpv.VersionName,pa.FilmNum,iss.RationValue,iss.SettleTypeNo,iss.StartTime,iss.EndTime
order by iss.DivideID", startdate, enddate));

            DataTable settleDetail = GetDataTableBySqlSettle(string.Format(@"select iss.FILMISSUEID ,tag.ISSUESTAGEGUID 分段ID,to_char(iss.createtime,'yyyy-MM-dd') 创建日期,to_char(iss.OPDATE,'yyyy-MM-dd') 修改日期,iss.FILMISSUEGUID 一览表ID
,iss.SETTLEMENTGUID 结算单位ID
,sett.SETTLENAME 结算单位名称
,(case when sett.SETTLEMODE=0 then '单结' when sett.SETTLEMODE=1 then '影院全结' when sett.SETTLEMODE=2 then '管理公司' else '院线代结' end) 结算单位类型
,f.filmname 影片名称,fv.FILMVERSIONNAME 影片版本,seq.FILMSEQCODE 影片编码,iss.RATIOVALUE 比例值,(case when iss.RATIOTYPE=0 then '自购' else '租赁' end) 比例类型
,to_char(iss.PLAYSTARTTIME,'yyyy-MM-dd') 一览表开始时间,to_char(iss.PLAYENDTIME,'yyyy-MM-dd') 一览表结束时间
,sum(case when fh.FILMISSUEID is null then 0 else 1 end) 特殊设备总数
from FILMISSUE iss
inner join issuestage tag on iss.ISSUESTAGEID=tag.ISSUESTAGEID
inner join SETTLEMENT sett on iss.SETTLEMENTID=sett.SETTLEMENTID
left join filmseq seq on iss.filmversiontype=seq.filmversiontype and  iss.filmid= seq.filmid
left join film f on  iss.filmid=f.filmid
left join FILMVERSION fv on iss.FILMVERSIONTYPE=fv.FILMVERSIONTYPE
left join FILMISSUE_HALL fh on iss.FILMISSUEID=fh.FILMISSUEID
where tag.playendtime>=to_date('{0}','yyyy-MM-dd') and tag.playstarttime<=to_date('{1}','yyyy-MM-dd') and tag.deleted=0  and iss.deleted=0 
group by tag.ISSUESTAGEGUID,iss.createtime,iss.OPDATE,iss.FILMISSUEID,iss.FILMISSUEGUID ,iss.SETTLEMENTGUID,sett.SETTLENAME,sett.SETTLEMODE ,f.filmname,fv.FILMVERSIONNAME,seq.FILMSEQCODE,iss.RATIOVALUE,iss.RATIOTYPE,iss.PLAYSTARTTIME,iss.PLAYENDTIME
order by iss.FILMISSUEID", startdate, enddate));
            InsertDatatableIssue(settleDetail, "A_FilmIssue_Settle");
            List<SheetMap> list = new List<SheetMap>();
            list.Add(new SheetMap()
            {
                sheetName = "结算 查询分段明细",
                isExport = true,
                sql =string.Format(@"SELECT tag.ISSUESTAGEGUID 分段ID
,f.filmname 影片名称,fv.FILMVERSIONNAME 影片版本,seq.FILMSEQCODE 影片编码
,to_char(tag.playstarttime,'yyyy-MM-dd') 分段开始时间,to_char(tag.playendtime,'yyyy-MM-dd') 分段结束时间
FROM issuestage tag
left join issuestage_issueversion tag1 on tag.ISSUESTAGEID=tag1.ISSUESTAGEID
left join issueversion tag2 on tag1.ISSUEVERSIONTYPE= tag2.issueversiontype
left join filmseq seq on   tag2.filmversiontype=seq.filmversiontype and  tag.filmid= seq.filmid
left join film f on f.filmid= tag.filmid
left join FILMVERSION fv on tag2.FILMVERSIONTYPE=fv.FILMVERSIONTYPE
where tag.playendtime>=to_date('{0}','yyyy-MM-dd')  and tag.playstarttime<=to_date('{1}','yyyy-MM-dd') and tag.deleted=0
order by tag.ISSUESTAGEGUID", startdate, enddate) 
            });
            list.Add(new SheetMap()
            {
                sheetName = "发行 查询分段明细",
                isExport = true,
                sql = string.Format(@"select tag.SectionID 分段ID --,jfpv.PVersionName
,fi.FilmName 影片名称,jfv.VersionName 影片版本,pa.FilmNum 影片编码
,CONVERT(varchar(10), tag.StartTime, 120) 分段开始时间,CONVERT(varchar(10), tag.EndTime, 120) 分段结束时间
from [dbo].[JS_Publish_Batch]  tag
inner join JS_FilmEdtionList jf on tag.SectionID=jf.SectionID
inner join JS_FilmPublishVersion jfpv on jf.PVersionID=jfpv.PVersionID
inner join JS_FilmVersion jfv on jfpv.VersionType=jfv.VersionType
inner join Publish_Batch pb on tag.SystemNum=pb.SystemNum
inner join Film_FilmInfo fi on pb.filmno=fi.filmno
left join Publish_FilmPermitAdd pa on pb.filmno=pa.filmno and jfv.VersionID=pa.VersionID
where tag.EndTime>='{0}'  and tag.StartTime<='{1}'", startdate, enddate)
            });
            list.Add(new SheetMap()
            {
                sheetName = "结算 一览表统计",
                isExport = true,
                sql =string.Format( @"select count(distinct iss.FILMISSUEID) 一览表总数,count(distinct fh.FILMISSUEID) 特殊一览表总数,count(fh.FILMISSUE_HALLGUID)  特殊设备表总数,count(distinct fh.HALLID) 特殊设备总数
from FILMISSUE iss
inner join issuestage tag on iss.ISSUESTAGEID=tag.ISSUESTAGEID
left join filmseq seq on iss.filmversiontype=seq.filmversiontype and  iss.filmid= seq.filmid
left join film f on  iss.filmid=f.filmid
left join FILMVERSION fv on iss.FILMVERSIONTYPE=fv.FILMVERSIONTYPE
left join FILMISSUE_HALL fh on iss.FILMISSUEID=fh.FILMISSUEID
where tag.playendtime>=to_date('{0}','yyyy-MM-dd') and tag.playstarttime<=to_date('{1}','yyyy-MM-dd') and tag.deleted=0  and iss.deleted=0", startdate, enddate)
            });
            list.Add(new SheetMap()
            {
                sheetName = "发行 一览表统计",
                isExport = true,
                sql = string.Format( @"select count(distinct iss.DivideID) 一览表总数,count(distinct fh.DivideID) 特殊一览表总数,count(fh.SpecialHallID)  特殊设备表总数,count(distinct fh.deviceid) 特殊设备总数
from JS_IssuedList iss
inner join JS_Publish_Batch tag  on iss.sectionid=tag.sectionid
inner join Publish_Batch pb on tag.SystemNum=pb.SystemNum
inner join Film_FilmInfo fi on pb.filmno=fi.filmno
left join JS_FilmVersion jfpv on iss.Versiontype=jfpv.Versiontype
left join JS_SpecialValueHall fh on iss.DivideID=fh.DivideID
where tag.EndTime>='{0}'  and tag.StartTime<='{1}' and   iss.DelFlag=0", startdate, enddate)
            });
            list.Add(new SheetMap() { sheetName = "结算 7月份一览表明细", isExport = true, sql = @"select * from A_FilmIssue_Settle" });
            list.Add(new SheetMap() { sheetName = "发行 7月份一览表明细", isExport = true, sql = @"select * from A_FilmIssue_Issue" });
            list.Add(new SheetMap()
            {
                sheetName = "结算平台未被正确标记",
                isExport = true,
                sql = @"select b.*,(case when iss.DelFlag is null then '发行未找到' when iss.DelFlag=1 then '发行标记删除' else '发行可用' end) 状态
from [dbo].[A_FilmIssue_Issue] a
right join [dbo].[A_FilmIssue_Settle] b on a.[一览表ID]=b.[一览表ID]
left join JS_IssuedList iss on b.[一览表ID]=iss.DivideID
where a.[一览表ID] is null
order by iss.DelFlag"
            });
            list.Add(new SheetMap()
            {
                sheetName = "",
                isExport = true,
                sql = @"select b.*,(case when iss.DelFlag is null then '发行未找到' when iss.DelFlag=1 then '发行标记删除' else '发行可用' end) 状态
from [dbo].[A_FilmIssue_Issue] a
right join [dbo].[A_FilmIssue_Settle] b on a.[一览表ID]=b.[一览表ID]
left join JS_IssuedList iss on b.[一览表ID]=iss.DivideID
where a.[一览表ID] is null
order by iss.DelFlag"
            });
            list.Add(new SheetMap()
            {
                sheetName = "发行未推送到结算平台",
                isExport = true,
                sql = @"select a.*
from [dbo].[A_FilmIssue_Issue] a
left join [dbo].[A_FilmIssue_Settle] b on a.[一览表ID]=b.[一览表ID]
where b.[一览表ID] is null"
            });
            list.Add(new SheetMap()
            {
                sheetName = "特殊设备数量不匹配",
                isExport = true,
                sql = @"select b.*,a.[特殊设备总数]
from [dbo].[A_FilmIssue_Issue] a
inner join [dbo].[A_FilmIssue_Settle] b on a.[一览表ID]=b.[一览表ID]
where a.[特殊设备总数]<>b.[特殊设备总数]"
            });
            RegisterLicense();
            var workbook = new Workbook();

            int sheetIndex = 0;
            foreach (var item in list)
            {
                DataTable dt = new DataTable();
                if (!item.isExport) continue;
                Console.WriteLine(item.sheetName + "开始导出.......");
                switch (item.sheetName)
                {
                    case "结算 查询分段明细":
                        dt = GetDataTableBySqlSettle(item.sql);
                        break;
                    case "发行 查询分段明细":
                        dt = GetDataTableBySqlIssue(item.sql);
                        break;
                    case "结算 一览表统计":
                        dt = GetDataTableBySqlSettle(item.sql);
                        break;
                    case "发行 一览表统计":
                        dt = GetDataTableBySqlIssue(item.sql);
                        break;
                    case "结算 7月份一览表明细":
                        dt = GetDataTableBySqlIssue(item.sql);
                        break;
                    case "发行 7月份一览表明细":
                        dt = GetDataTableBySqlIssue(item.sql);
                        break;
                    case "结算平台未被正确标记":
                        dt = GetDataTableBySqlIssue(item.sql);
                        break;
                    case "发行未推送到结算平台":
                        dt = GetDataTableBySqlIssue(item.sql);
                        break;
                    case "特殊设备数量不匹配":
                        dt = GetDataTableBySqlIssue(item.sql);
                        break;


                }
                if (dt.Rows.Count == 0) continue;
                if (sheetIndex > 0)
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
                Console.WriteLine(item.sheetName + "导出成功！");
                sheetIndex++;
            }
            string fileName = "发行结算推送一览表对比" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            workbook.Save(fileName);
            Console.WriteLine("导出excel完成，结束！");


            Console.ReadLine();
        }
        public static DataTable GetDataTableBySqlSettle(string sql)
        {
            DataSet ds = new DataSet();
            //string conStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.1.31.66)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME = orcl11g)));User Id=settle;Password=settle;";
            string conStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.10.130.23)(PORT=8903))(CONNECT_DATA=(SERVICE_NAME = settle_primary)));User Id=settle;Password=settle;";
            OracleConnection con = new OracleConnection(conStr);
            OracleCommand cmd = new OracleCommand(sql, con);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            da.Fill(ds);
            return ds.Tables[0];

        }

        public static DataTable GetDataTableBySqlIssue(string sql)
        {
            DataSet ds = new DataSet();
            //string conStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.1.31.66)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME = orcl11g)));User Id=settle;Password=settle;";
            string conStr = "Data Source=10.10.10.80,2433;Initial Catalog=HyFilmCopyForZyNew;UID=sa;PWD=hyby@123;";
            SqlConnection con = new SqlConnection(conStr);
            SqlCommand cmd = new SqlCommand(sql, con);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(ds);
            return ds.Tables[0];

        }
        public static void ExecSqlIssue(string sql)
        {
            string conStr = "Data Source=10.10.10.80,2433;Initial Catalog=HyFilmCopyForZyNew;UID=sa;PWD=hyby@123;";
            using (SqlConnection con = new SqlConnection(conStr))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.ExecuteNonQuery();
            }

        }
        public static void InsertDatatableIssue(DataTable dt,string tablename)
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

                //con.Open();
                //var cmdText = "INSERT INTO " + tablename +"(" ;
                //for (var i = 0; i < dt.Columns.Count; i++)
                //{
                //    cmdText += dt.Columns[i].ColumnName;
                //    if (i != (dt.Columns.Count - 1)) cmdText += ",";
                //}
                //cmdText += ") VALUES( ";
                //for (var i = 0; i < dt.Columns.Count; i++)
                //{
                //    cmdText += "@v" + i ;
                //    if (i != (dt.Columns.Count - 1)) cmdText += ",";
                //}
                //cmdText += ")";
                //var cmd = new SqlCommand(cmdText, con);

                //for (var j = 0; j < dt.Rows.Count; j++)
                //{
                //    cmd.Parameters.Clear();
                //    for (var i = 0; i < dt.Columns.Count; i++)
                //    {
                //        var par = new SqlParameter("@v" + i, dt.Rows[j][i].ToString())
                //        {
                //        };
                //        cmd.Parameters.Add(par);
                //    }
                //    cmd.ExecuteNonQuery();
                //}
                //con.Close();
            }

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
