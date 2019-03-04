using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using 补充结算一览表.model;
using Oracle.ManagedDataAccess.Client;
using Oracle.ManagedDataAccess.Types;


namespace 补充结算一览表
{
    class Program
    {
        static void Main(string[] args)
        {
            //ExecMergeDate();
            //Console.ReadLine();
            //return;
            //结算分段明细
            List<IssueStage> issStages = DataTableSerializer.ToList<IssueStage>(GetSettleIssueTagList());
            List<Settlement> settles = DataTableSerializer.ToList<Settlement>(GetSettleSettlementList());
            List<Hall> halls = DataTableSerializer.ToList<Hall>(GetSettleHallList());

            //发行未推送到结算平台
            string sql = @"select a.*
from [dbo].[A_FilmIssue_Issue] a
left join [dbo].[A_FilmIssue_Settle] b on a.[一览表ID]=b.[一览表ID]
where b.[一览表ID] is null  and a.审核状态='通过'";
            var dt = SqlServerHelper.GetDataTableBySql(sql);
            int maxFilmIssueId = Convert.ToInt32(OracleHelper.GetSingle("SELECT max(filmissueid) FROM  filmissue"));
            int maxFilmIssueHallId = Convert.ToInt32(OracleHelper.GetSingle("SELECT max(filmissue_hallid) FROM  filmissue_hall"));
            string divideId = "";
            int specialValueHallNum = 0;
            int filmId = 0;
            int settleId = 0;
            int issueStageId = 0;
            string insertFilmissueSql = "";
            string insertFilmissue_HallSql = "";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                maxFilmIssueId += 1;
                divideId = dt.Rows[i]["一览表ID"].ToString();
                specialValueHallNum = Convert.ToInt32(dt.Rows[i]["特殊设备总数"]);
                filmId = issStages.Where(re => re.issuestageguid == dt.Rows[i]["分段ID"].ToString()).FirstOrDefault().filmid;
                settleId = settles.Where(re => re.settlementguid == dt.Rows[i]["结算单位ID"].ToString()).FirstOrDefault().settlementid;
                issueStageId = issStages.Where(re => re.issuestageguid == dt.Rows[i]["分段ID"].ToString()).FirstOrDefault().issuestageid;

                var isExists = OracleHelper.GetSingle("SELECT filmissueid FROM  filmissue where filmissueguid='" + divideId + "'");
                if (isExists != null) {
                    OracleHelper.ExecToSqlNonQuery("delete filmissue where filmissueid=" + isExists);
                    OracleHelper.ExecToSqlNonQuery("delete filmissue_hall where filmissueid=" + isExists);
                }

                OracleParameter[] arrParas = new OracleParameter[16];
                arrParas[0] = new OracleParameter(":FILMISSUEID", maxFilmIssueId);
                arrParas[1] = new OracleParameter(":FILMISSUEGUID", divideId);
                arrParas[2] = new OracleParameter(":ISSUESTAGEID", issueStageId);
                arrParas[3] = new OracleParameter(":FILMID", filmId);
                arrParas[4] = new OracleParameter(":FILMVERSIONTYPE", dt.Rows[i]["VersionType"].ToString());
                arrParas[5] = new OracleParameter(":SETTLEMENTGUID", dt.Rows[i]["结算单位ID"].ToString());
                arrParas[6] = new OracleParameter(":SETTLEMENTID", settleId);
                arrParas[7] = new OracleParameter(":FILMISSUETYPE", dt.Rows[i]["SettleType"].ToString());
                arrParas[8] = new OracleParameter(":PLAYSTARTTIME", (OracleTimeStamp)Convert.ToDateTime(dt.Rows[i]["一览表开始时间"]));
                arrParas[9] = new OracleParameter(":PLAYENDTIME", (OracleTimeStamp)Convert.ToDateTime(dt.Rows[i]["一览表结束时间"].ToString()));
                arrParas[10] = new OracleParameter(":RATIOTYPE", dt.Rows[i]["SettleTypeNo"].ToString());
                arrParas[11] = new OracleParameter(":RATIOVALUE", dt.Rows[i]["比例值"].ToString());
                arrParas[12] = new OracleParameter(":PLAYREQUIRE", dt.Rows[i]["PLAYREQUIRE"].ToString());
                arrParas[13] = new OracleParameter(":PRICEREQUIRE", dt.Rows[i]["PRICEREQUIRE"].ToString());
                arrParas[14] = new OracleParameter(":OPERATORNAME", dt.Rows[i]["OPERATORNAME"].ToString());
                arrParas[15] = new OracleParameter(":REMARK", dt.Rows[i]["REMARK"].ToString());
                insertFilmissueSql = @"INSERT INTO FILMISSUE(FILMISSUEID,FILMISSUEGUID,ISSUESTAGEID,FILMID,FILMVERSIONTYPE,SETTLEMENTGUID,SETTLEMENTID,FILMISSUETYPE,PLAYSTARTTIME,
								PLAYENDTIME,RATIOTYPE,RATIOVALUE,PLAYREQUIRE,PRICEREQUIRE,OPERATORNAME,REMARK,CREATETIME,CREATEUSERID,DELETED,OPUSERID,OPFUNCTION,OPDATE,OPTYPE)VALUES(";
                insertFilmissueSql += ":FILMISSUEID,";
                insertFilmissueSql += ":FILMISSUEGUID,";
                insertFilmissueSql += ":ISSUESTAGEID,";
                insertFilmissueSql += ":FILMID,";
                insertFilmissueSql += ":FILMVERSIONTYPE,";
                insertFilmissueSql += ":SETTLEMENTGUID,";
                insertFilmissueSql += ":SETTLEMENTID,";
                insertFilmissueSql += ":FILMISSUETYPE,";
                insertFilmissueSql += ":PLAYSTARTTIME,";
                insertFilmissueSql += ":PLAYENDTIME,";
                insertFilmissueSql += ":RATIOTYPE,";
                insertFilmissueSql += ":RATIOVALUE,";
                insertFilmissueSql += ":PLAYREQUIRE,";
                insertFilmissueSql += ":PRICEREQUIRE,";
                insertFilmissueSql += ":OPERATORNAME,";
                insertFilmissueSql += ":REMARK,";
                insertFilmissueSql += "sysdate,0,0,0,'sql',sysdate,'insert'";
                insertFilmissueSql += ")";
                OracleHelper.ExecToSqlNonQuery(insertFilmissueSql, arrParas);
                if (specialValueHallNum > 0)
                {
                    DataTable specialDt = SqlServerHelper.GetDataTableBySql(@" select SpecialHallID,EQ_TheaterCode,DeviceID from JS_SpecialValueHall where DelFlag=0 and DivideID='" + divideId + "'");
                    for (int x = 0; x < specialDt.Rows.Count; x++) {
                        maxFilmIssueHallId += 1;
                        int hallid = halls.Where(re => re.hallcode == specialDt.Rows[x]["EQ_TheaterCode"].ToString()).FirstOrDefault().hallid; ;
                        insertFilmissue_HallSql = "INSERT INTO FILMISSUE_HALL(FILMISSUE_HALLID,FILMISSUE_HALLGUID,HALLCODE,HALLID,FILMISSUEID,HALL_DEVICETYPEGUID,CREATETIME,CREATEUSERID,DELETED,OPUSERID,OPFUNCTION,OPDATE,OPTYPE)VALUES(";
                        insertFilmissue_HallSql += ":FILMISSUE_HALLID,";
                        insertFilmissue_HallSql += ":FILMISSUE_HALLGUID,";
                        insertFilmissue_HallSql += ":HALLCODE,";
                        insertFilmissue_HallSql += ":HALLID,";
                        insertFilmissue_HallSql += ":FILMISSUEID,";
                        insertFilmissue_HallSql += ":HALL_DEVICETYPEGUID,";
                        insertFilmissue_HallSql += "sysdate,0,0,0,'sql',sysdate,'insert')";
                        OracleParameter[] arrParasSpecial = new OracleParameter[6];
                        arrParasSpecial[0] = new OracleParameter(":FILMISSUE_HALLID", maxFilmIssueHallId);
                        arrParasSpecial[1] = new OracleParameter(":FILMISSUE_HALLGUID", specialDt.Rows[x]["SpecialHallID"].ToString());
                        arrParasSpecial[2] = new OracleParameter(":HALLCODE", specialDt.Rows[x]["EQ_TheaterCode"].ToString());
                        arrParasSpecial[3] = new OracleParameter(":HALLID", hallid);
                        arrParasSpecial[4] = new OracleParameter(":FILMISSUEID", maxFilmIssueId);
                        arrParasSpecial[5] = new OracleParameter(":HALL_DEVICETYPEGUID", specialDt.Rows[x]["DeviceID"].ToString());
                        OracleHelper.ExecToSqlNonQuery(insertFilmissue_HallSql, arrParasSpecial);
                    }
                }


            }

            //结算平台未被正确标记
            sql = @"select b.*,(case when iss.DelFlag is null then '发行未找到' when iss.DelFlag=1 then '发行标记删除' else '发行可用' end) 状态
from [dbo].[A_FilmIssue_Issue] a
right join [dbo].[A_FilmIssue_Settle] b on a.[一览表ID]=b.[一览表ID]
left join JS_IssuedList iss on b.[一览表ID]=iss.DivideID
where a.[一览表ID] is null and (b.结算单位类型!='影院全结' or iss.DelFlag is not null or b.一览表开始时间>='2018-08-01') and a.审核状态='通过'
order by iss.DelFlag";
            dt = SqlServerHelper.GetDataTableBySql(sql);
            string sqlFilmIssueDelere = "";
            for (int i = 0; i < dt.Rows.Count; i++)
            {

                sqlFilmIssueDelere = "update FILMISSUE set DELETED=1 where FILMISSUEID=" + dt.Rows[i]["FILMISSUEID"].ToString();
                OracleHelper.ExecToSqlNonQuery(sqlFilmIssueDelere);
                if (Convert.ToInt32(dt.Rows[i]["特殊设备总数"]) > 0)
                {
                    sqlFilmIssueDelere = "delete FILMISSUE_HALL where FILMISSUEID=" + dt.Rows[i]["FILMISSUEID"].ToString();
                    OracleHelper.ExecToSqlNonQuery(sqlFilmIssueDelere);
                }
            }

            //特殊设备数量不匹配
            sql = @"select b.*,a.[特殊设备总数] as [特殊设备总数(发行)],a.[特殊设备总数(删除)],a.[特殊设备总数]+a.[特殊设备总数(删除)] -(select COUNT(1) from [dbo].JS_IssuedSendLog c  where c.DivideID=a.一览表ID ) as [特殊设备(发行未推送)],a.[审核状态] as [审核状态(发行)],a.[是否已经推送]
from [dbo].[A_FilmIssue_Issue] a
inner join [dbo].[A_FilmIssue_Settle] b on a.[一览表ID]=b.[一览表ID]
where a.[特殊设备总数]<>b.[特殊设备总数] and a.审核状态='通过'";
            dt = SqlServerHelper.GetDataTableBySql(sql);
            int filmIssueId = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                divideId = dt.Rows[i]["一览表ID"].ToString();
                filmIssueId = Convert.ToInt32(dt.Rows[i]["FILMISSUEID"].ToString());
                var dtSpecial = SqlServerHelper.GetDataTableBySql("SELECT  [SpecialHallID] ,[EQ_TheaterCode]  ,[DeviceID] FROM [dbo].[JS_SpecialValueHall] where [DelFlag]=0 and [DivideID]='" + divideId + "'");
                var dtFilmIssueHall = OracleHelper.GetDataTableBySql("SELECT filmissue_hallguid,hallcode,filmissueid,hall_devicetypeguid FROM  filmissue_hall where FILMISSUEID=" + filmIssueId);
                List<Filmissue_Hall> settleHallList = DataTableSerializer.ToList<Filmissue_Hall>(dtFilmIssueHall);
                List<SpecialValueHall> issueHallList = DataTableSerializer.ToList<SpecialValueHall>(dtSpecial);
                for (int x = 0; x < issueHallList.Count; x++)
                {
                    Filmissue_Hall isFindHall = settleHallList.Where(re => re.hall_devicetypeguid == issueHallList[x].DeviceID).FirstOrDefault();
                    if (isFindHall != null)
                    {
                        settleHallList.Remove(isFindHall);
                    }
                    else
                    {
                        maxFilmIssueHallId += 1;
                        insertFilmissue_HallSql = "INSERT INTO FILMISSUE_HALL(FILMISSUE_HALLID,FILMISSUE_HALLGUID,HALLCODE,HALLID,FILMISSUEID,HALL_DEVICETYPEGUID,CREATETIME,CREATEUSERID,DELETED,OPUSERID,OPFUNCTION,OPDATE,OPTYPE)VALUES(";
                        insertFilmissue_HallSql += ":FILMISSUE_HALLID,";
                        insertFilmissue_HallSql += ":FILMISSUE_HALLGUID,";
                        insertFilmissue_HallSql += ":HALLCODE,";
                        insertFilmissue_HallSql += ":HALLID,";
                        insertFilmissue_HallSql += ":FILMISSUEID,";
                        insertFilmissue_HallSql += ":HALL_DEVICETYPEGUID,";
                        insertFilmissue_HallSql += "sysdate,0,0,0,'sql',sysdate,'insert')";
                        OracleParameter[] arrParasSpecial = new OracleParameter[6];
                        arrParasSpecial[0] = new OracleParameter(":FILMISSUE_HALLID", maxFilmIssueHallId);
                        arrParasSpecial[1] = new OracleParameter(":FILMISSUE_HALLGUID", issueHallList[x].SpecialHallID);
                        arrParasSpecial[2] = new OracleParameter(":HALLCODE", issueHallList[x].EQ_TheaterCode);
                        arrParasSpecial[3] = new OracleParameter(":HALLID", halls.Where(re => re.hallcode == issueHallList[x].EQ_TheaterCode).FirstOrDefault().hallid);
                        arrParasSpecial[4] = new OracleParameter(":FILMISSUEID", filmIssueId);
                        arrParasSpecial[5] = new OracleParameter(":HALL_DEVICETYPEGUID", issueHallList[x].DeviceID);
                        OracleHelper.ExecToSqlNonQuery(insertFilmissue_HallSql, arrParasSpecial);
                    }
                }
                string deleteFilmIssueHallSql = "";
                for (int x = 0; x < settleHallList.Count; x++)
                {
                    deleteFilmIssueHallSql = "delete FILMISSUE_HALL where FILMISSUEID=" + filmIssueId + " and HALL_DEVICETYPEGUID='" + settleHallList[x].hall_devicetypeguid + "'";
                    OracleHelper.ExecToSqlNonQuery(deleteFilmIssueHallSql);
                }
            }
        }

        public static void ExecMergeDate()
        {
            string startdate = "2018-07-01";
            string enddate = "2018-12-07";
            SqlServerHelper.ExecSqlIssue("drop table A_FilmIssue_Issue");
            SqlServerHelper.ExecSqlIssue("truncate table A_FilmIssue_Settle");
            SqlServerHelper.ExecSqlIssue(string.Format(@"select iss.DivideID 一览表ID
,iss.SettleID 结算单位ID
,sett.SettleName 结算单位名称
,tag.SectionID 分段ID
,fi.filmname 影片名称,jfpv.VersionName 影片版本,pa.FilmNum 影片编码,iss.RationValue 比例值,(case when iss.SettleTypeNo=0 then '自购' else '租赁' end) 比例类型
,CONVERT(varchar(10), iss.StartTime, 120) 一览表开始时间,CONVERT(varchar(10), iss.EndTime, 120) 一览表结束时间
,sum(case when fh.DivideID is null then 0 when fh.DelFlag=0 then 1 else 0 end) 特殊设备总数
,sum(case when fh.DivideID is null then 0 when fh.DelFlag=1 then 1 else 0 end) [特殊设备总数(删除)]
,(case when iss.AuditState=0 then '通过' when iss.AuditState=1 then '未通过' when iss.AuditState=3 then '未审核' else '' end ) 审核状态
,(case when iss.HasSend=0 then '不推送' when iss.HasSend=1 then '待推送' when iss.HasSend=2 then '已推送' else '' end ) 是否已经推送
,iss.VersionType
,iss.SettleType
,iss.SettleTypeNo
,iss.PlayRequire
,iss.PriceRequire
,iss.OperatorName
,iss.Remark
into A_FilmIssue_Issue
from JS_IssuedList iss
inner join JS_Publish_Batch tag  on iss.sectionid=tag.sectionid
inner join JS_SettleCompany sett on iss.SettleID=sett.SettleID
inner join Publish_Batch pb on tag.SystemNum=pb.SystemNum
inner join Film_FilmInfo fi on pb.filmno=fi.filmno
left join JS_FilmVersion jfpv on iss.Versiontype=jfpv.Versiontype
left join JS_SpecialValueHall fh on iss.DivideID=fh.DivideID  
left join Publish_FilmPermitAdd pa on pb.filmno=pa.filmno and jfpv.VersionID=pa.VersionID
where tag.SectionID='78EDC8EC-0A89-4CC6-9F4D-30B20CC1F596'   and   iss.DelFlag=0 
group by iss.DivideID ,iss.AuditState,iss.HasSend,iss.SettleID,sett.SettleName ,tag.SectionID,iss.VersionType
,iss.SettleType
,iss.SettleTypeNo
,iss.PlayRequire
,iss.PriceRequire
,iss.OperatorName
,iss.Remark,fi.filmname,jfpv.VersionName,pa.FilmNum,iss.RationValue,iss.SettleTypeNo,iss.StartTime,iss.EndTime
order by iss.DivideID", enddate));//where tag.EndTime>='{0}'   and   iss.DelFlag=0 

            DataTable settleDetail = OracleHelper.GetDataTableBySql(string.Format(@"select iss.FILMISSUEID ,tag.ISSUESTAGEGUID 分段ID,to_char(iss.createtime,'yyyy-MM-dd') 创建日期,to_char(iss.OPDATE,'yyyy-MM-dd') 修改日期,iss.FILMISSUEGUID 一览表ID
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
where tag.ISSUESTAGEGUID='78EDC8EC-0A89-4CC6-9F4D-30B20CC1F596' and tag.deleted=0  and iss.deleted=0 
group by tag.ISSUESTAGEGUID,iss.createtime,iss.OPDATE,iss.FILMISSUEID,iss.FILMISSUEGUID ,iss.SETTLEMENTGUID,sett.SETTLENAME,sett.SETTLEMODE ,f.filmname,fv.FILMVERSIONNAME,seq.FILMSEQCODE,iss.RATIOVALUE,iss.RATIOTYPE,iss.PLAYSTARTTIME,iss.PLAYENDTIME
order by iss.FILMISSUEID",  enddate));
            SqlServerHelper.InsertDatatableIssue(settleDetail, "A_FilmIssue_Settle");
        }//where tag.playendtime>=to_date('{0}','yyyy-MM-dd') and tag.deleted=0  and iss.deleted=0 

        public static DataTable GetSettleIssueTagList()
        {

            DataTable dt = OracleHelper.GetDataTableBySql(@"SELECT tag.ISSUESTAGEID,tag.ISSUESTAGEGUID 
,f.filmname ,fv.FILMVERSIONNAME ,seq.FILMSEQCODE ,tag.FILMID
,to_char(tag.playstarttime,'yyyy-MM-dd') playstarttime,to_char(tag.playendtime,'yyyy-MM-dd') playendtime
FROM issuestage tag
left join issuestage_issueversion tag1 on tag.ISSUESTAGEID=tag1.ISSUESTAGEID
left join issueversion tag2 on tag1.ISSUEVERSIONTYPE= tag2.issueversiontype
left join filmseq seq on   tag2.filmversiontype=seq.filmversiontype and  tag.filmid= seq.filmid
left join film f on f.filmid= tag.filmid
left join FILMVERSION fv on tag2.FILMVERSIONTYPE=fv.FILMVERSIONTYPE");
            return dt;
        }

        public static DataTable GetSettleSettlementList()
        {

            DataTable dt = OracleHelper.GetDataTableBySql(@"SELECT    settlementid,    settlementguid,    settlename FROM    settlement");
            return dt;
        }
        public static DataTable GetSettleHallList()
        {

            DataTable dt = OracleHelper.GetDataTableBySql(@"SELECT   hallcode,hallid FROM hall where deleted=0 ");
            return dt;
        }
    }
}
