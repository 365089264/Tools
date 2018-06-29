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
    class Program
    {
        static void Maina(string[] args)
        {
            //string filepath = "EQ_CinemaInfo.txt";
            //string json = GetFileJson(filepath);
            //var eQCin_Cinemas = (List<EQCin_Cinema>)JsonToObject(json, new List<EQCin_Cinema>());
            //filepath = "EQCin_ForwardInfo.txt";
            //json = GetFileJson(filepath);
            //var eQCin_ForwardInfos = (List<EQCin_ForwardInfo>)JsonToObject(json, new List<EQCin_ForwardInfo>());
            string sql1 = "select * from v_EQCin_CinemaFilter where  EQ_CinemaCode<>'' ";
            string sql2 = "select * from v_EQCin_ForwardInfoFilter where [isDelete]=0 and [Match_CinemaCode] is   null   ";

            var dt = GetDataTableBySql(sql1);

            var eQCin_Cinemas = DataTableSerializer.ToList<EQCin_Cinema>(dt);
            dt = GetDataTableBySql(sql2);
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
                //var option = new MatchResult();
                for (var j = 0; j < eQCin_Cinemas.Count; j++)
                {
                    tempRate = levenshtein(eQCin_Cinemas[j].EQ_Adress, eQCin_ForwardInfos[i].EQ_ForAddress);
                    tempRate2 = levenshtein(eQCin_Cinemas[j].EQ_CinemaName, eQCin_ForwardInfos[i].EQ_CinemaName);
                    if (tempRate > maxRate)
                    {
                        maxRate = tempRate;
                        EQ_CinemaCode = eQCin_Cinemas[j].EQ_CinemaCode;
                        matchType = "EQ_Adress";
                        remark = (tempRate2*100).ToString();
                        //sysnum = eQCin_ForwardInfos[j].SystemNum;
                    }
                    
                    if (tempRate2 > maxRate2)
                    {
                        maxRate2 = tempRate2;
                        EQ_CinemaCode2 = eQCin_Cinemas[j].EQ_CinemaCode;
                        matchType2 = "EQ_CinemaName";
                        remark2 = (tempRate*100).ToString();
                        //sysnum = eQCin_ForwardInfos[j].SystemNum;
                    }
                }
                if (maxRate > 0)
                {
                    InsertCinameMatchInfoBySql("insert into CinameMatchInfo20150529 values(" + eQCin_ForwardInfos[i].SystemNum + ",'" + EQ_CinemaCode + "','','','" + remark + "',2,1,'EQ_Adress','" + maxRate + "')");
                    //Console.WriteLine(eQCin_Cinemas[i].EQ_CinemaCode + " : " + sysnum + " ,相似度最高" + maxRate);
                }
                if (maxRate2 > 0)
                {
                    InsertCinameMatchInfoBySql("insert into CinameMatchInfo20150529 values(" + eQCin_ForwardInfos[i].SystemNum + ",'" + EQ_CinemaCode2 + "','','','" + remark2 + "',3,1,'EQ_CinemaName','" + maxRate2 + "')");
                    //Console.WriteLine(eQCin_Cinemas[i].EQ_CinemaCode + " : " + sysnum + " ,相似度最高" + maxRate);
                }
                //reslut.Add(option);
            }
            Console.WriteLine(DateTime.Now.ToLongTimeString());
            Console.WriteLine(x);
            Console.ReadLine();
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
        public static DataTable GetDataTableBySql(string sql)
        {
            DataSet ds = new DataSet();
            string conStr = "Data Source=192.168.144.73;Initial Catalog=HyFilmCopyForZyNew;UID=sa;PWD=Welcome1;";
            SqlConnection con = new SqlConnection(conStr);
            SqlCommand cmd = new SqlCommand(sql, con);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(ds);
            return ds.Tables[0];

        }
        public static void InsertCinameMatchInfoBySql(string sql)
        {
            DataSet ds = new DataSet();
            string conStr = "Data Source=192.168.144.73;Initial Catalog=HyFilmCopyForZyNew;UID=sa;PWD=Welcome1;";
            using (SqlConnection con = new SqlConnection(conStr))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.ExecuteNonQuery();
            }

        }

        // 从一个Json串生成对象信息
        public static object JsonToObject(string jsonString, object obj)
        {
            var serializer = new DataContractJsonSerializer(obj.GetType());
            var mStream = new MemoryStream(Encoding.UTF8.GetBytes(jsonString));
            return serializer.ReadObject(mStream);
        }
        /// <summary>
        /// C#计算2个字符串的相似度
        /// </summary>
        /// <param name="str1"></param>
        /// <param name="str2"></param>
        /// <returns></returns>
        public static float levenshtein(string str1, string str2)
        {
            //计算两个字符串的长度。 
            int len1 = str1.Length;
            int len2 = str2.Length;
            //建立上面说的数组，比字符长度大一个空间 
            int[,] dif = new int[len1 + 1, len2 + 1];
            //赋初值，步骤B。 
            for (int a = 0; a <= len1; a++)
            {
                dif[a, 0] = a;
            }
            for (int a = 0; a <= len2; a++)
            {
                dif[0, a] = a;
            }
            //计算两个字符是否一样，计算左上的值 
            int temp;
            for (int i = 1; i <= len1; i++)
            {
                for (int j = 1; j <= len2; j++)
                {
                    if (str1[i - 1] == str2[j - 1])
                    {
                        temp = 0;
                    }
                    else
                    {
                        temp = 1;
                    }
                    //取三个值中最小的 
                    dif[i, j] = Math.Min(Math.Min(dif[i - 1, j - 1] + temp, dif[i, j - 1] + 1), dif[i - 1, j] + 1);
                }
            }
            //Console.WriteLine("字符串\"" + str1 + "\"与\"" + str2 + "\"的比较");

            //取数组右下角的值，同样不同位置代表不同字符串的比较 
            //Console.WriteLine("差异步骤：" + dif[len1, len2]);
            //计算相似度 
            float similarity = 1 - (float)dif[len1, len2] / Math.Max(str1.Length, str2.Length);
            //Console.WriteLine("相似度：" + similarity);
            return similarity;
        }

    }
}
