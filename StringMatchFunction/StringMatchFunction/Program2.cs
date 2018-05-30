using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Data;
using System.Data.SqlClient;

namespace StringMatchFunction
{
    class Program2
    {
        static void Main0(string[] args)
        {
            //string filepath = "EQ_CinemaInfo.txt";
            //string json = GetFileJson(filepath);
            //var eQCin_Cinemas = (List<EQCin_Cinema>)JsonToObject(json, new List<EQCin_Cinema>());
            //filepath = "EQCin_ForwardInfo.txt";
            //json = GetFileJson(filepath);
            //var eQCin_ForwardInfos = (List<EQCin_ForwardInfo>)JsonToObject(json, new List<EQCin_ForwardInfo>());
            string sql1 = "select * from v_EQCin_CinemaFilter  where EQ_CinemaCode<>''";
            string sql2 = "select * from v_EQCin_ForwardInfoFilter ";

            var dt = Program.GetDataTableBySql(sql1);

            var eQCin_Cinemas = DataTableSerializer.ToList<EQCin_Cinema>(dt);
            dt = Program.GetDataTableBySql(sql2);
            var eQCin_ForwardInfos = DataTableSerializer.ToList<EQCin_ForwardInfo>(dt);

            var x = 0;
            Console.WriteLine(DateTime.Now.ToLongTimeString());
            List<MatchResult> reslut = new List<MatchResult>();
            float maxRate;
            float tempRate;
            string EQ_CinemaCode;
            string matchType = "";
            float maxRate2;
            float tempRate2;
            string EQ_CinemaCode2;
            string matchType2 = "";
            string remark="";
            string remark2 = "";
            for (var i = 0; i < eQCin_ForwardInfos.Count; i++)
            {
                tempRate = 0;
                maxRate = 0;
                EQ_CinemaCode = "";
                tempRate2 = 0;
                maxRate2 = 0;
                EQ_CinemaCode2 = "";
                 remark="";
                 remark2 = "";
                 var cinemaEnity = eQCin_Cinemas.Where(re => re.EQ_CinemaCode == eQCin_ForwardInfos[i].EQ_CinemaCode).FirstOrDefault();
                 if (cinemaEnity != null)
                 {
                     tempRate = Program.levenshtein(cinemaEnity.EQ_Adress, eQCin_ForwardInfos[i].EQ_ForAddress);
                     tempRate2 = Program.levenshtein(cinemaEnity.EQ_CinemaName, eQCin_ForwardInfos[i].EQ_CinemaName);

                     maxRate = tempRate;
                     EQ_CinemaCode = eQCin_ForwardInfos[i].EQ_CinemaCode;
                     matchType = "EQ_Adress";
                     remark = (tempRate2 * 100).ToString();
                     //sysnum = eQCin_ForwardInfos[j].SystemNum;


                     maxRate2 = tempRate2;
                     EQ_CinemaCode2 = eQCin_ForwardInfos[i].EQ_CinemaCode;
                     matchType2 = "EQ_CinemaName";
                     remark2 = (tempRate * 100).ToString();
                     //sysnum = eQCin_ForwardInfos[j].SystemNum;

                     Program.InsertCinameMatchInfoBySql("insert into [CinameMatchInfoSync3] values(" + eQCin_ForwardInfos[i].SystemNum + ",'" + EQ_CinemaCode + "','" + cinemaEnity.EQ_CinemaName + "','" + cinemaEnity.EQ_Adress + "','',1," + eQCin_Cinemas.Where(re => re.EQ_CinemaCode == eQCin_ForwardInfos[i].EQ_CinemaCode).Count() + ",'EQ_Adress','" + maxRate + "','EQ_CinemaName','" + maxRate2 + "')");
                     //Console.WriteLine(eQCin_Cinemas[i].EQ_CinemaCode + " : " + sysnum + " ,相似度最高" + maxRate);

                     //Console.WriteLine(eQCin_Cinemas[i].EQ_CinemaCode + " : " + sysnum + " ,相似度最高" + maxRate);
                 }
                 else
                 {
                     EQ_CinemaCode = eQCin_ForwardInfos[i].EQ_CinemaCode;
                     Program.InsertCinameMatchInfoBySql("insert into CinameMatchInfoSync3 values(" + eQCin_ForwardInfos[i].SystemNum + ",'" + EQ_CinemaCode + "','','','未找到相应专资编码',4,0,null,null,null,null)");
                 }
                //reslut.Add(option);
            }
            Console.WriteLine(DateTime.Now.ToLongTimeString());
            Console.WriteLine(x);
            Console.ReadLine();
        }
    

    }
}
