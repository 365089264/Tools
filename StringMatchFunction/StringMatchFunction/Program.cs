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
        static void Main(string[] args)
        {
            //string filepath = "EQ_CinemaInfo.txt";
            //string json = GetFileJson(filepath);
            //var eQCin_Cinemas = (List<EQCin_Cinema>)JsonToObject(json, new List<EQCin_Cinema>());
            //filepath = "EQCin_ForwardInfo.txt";
            //json = GetFileJson(filepath);
            //var eQCin_ForwardInfos = (List<EQCin_ForwardInfo>)JsonToObject(json, new List<EQCin_ForwardInfo>());
//            string sql1 = "  select EQ_CinemaCode, EQ_CinemaName,EQ_GroupName,EQ_AreaCode,EQ_Adress from EQCin_Cinema     where  EQ_CinemaCode<>'' ";
//            string sql2 = @"select a.SystemNum,a.EQ_CinemaName,a.EQ_GroupName,a.EQ_CityCode,a.EQ_ForAddress,a.ReEQ_CinemaCode as EQ_CinemaCode
//  from  [test].[dbo].[EQCin_ForwardInfo20190215] a inner join [test].[dbo].EQCin_ForwardInfo20190222 b on b.SystemNum=a.SystemNum where a.ReEQ_CinemaCode is not null  ";
            string sql1 = @"SELECT [专资编码] as EQ_CinemaCode
      ,[影院名称] as Forward_CinemaName
      ,[硬盘接收地址] as Forward_Adress
      ,[专资编码] as EQ_CinemaCode
      ,[标准影院名称] as EQ_CinemaName
      ,[专资影院名称] as EQ_ZZCinemaName
      ,[平台应原地址] as EQ_Adress
FROM [test].[dbo].[Sheet4$] a
  inner join [HyFilmCopyForZyNew].[dbo].[EQCin_Cinema] b on a.[专资编码]=b.[EQ_CinemaCode]
  where   a.[专资编码] is not  null and b.[EQ_CinemaCode] is not null 
  and a.[专资编码]  not in ('院线影管','转万达','不是奥斯卡','不是时代院线','机动硬盘') and a.[专资编码] <>''  and a.[发运方式] is null and a.[地址相似度] is  null ";

            var dt = GetDataTableBySql(sql1);

            var eQCin_Cinemas = DataTableSerializer.ToList<Forward_Cinema>(dt);
            //dt = GetDataTableBySql(sql2);
            //var eQCin_ForwardInfos = DataTableSerializer.ToList<EQCin_ForwardInfo>(dt);

            var x = 0;
            Console.WriteLine(DateTime.Now.ToLongTimeString());
            List<MatchResult> reslut = new List<MatchResult>();
            float maxRate;
            float tempRate;
            string EQ_CinemaCode;
            string matchType = "";
            float maxRate2;
            float tempRate2;
            float tempRate3;
            string EQ_CinemaCode2;
            string matchType2 = "";
            string remark = "";
            string remark2 = "";
            for (var i = 0; i < eQCin_Cinemas.Count; i++)
            {
                tempRate = 0;
                maxRate = 0;
                EQ_CinemaCode = "";
                tempRate2 = 0;
                maxRate2 = 0;
                EQ_CinemaCode2 = "";
                remark = "";
                remark2 = "";
                tempRate3 = 0;
                //EQCin_Cinema eq = eQCin_Cinemas.Where(re => re.EQ_CinemaCode == eQCin_ForwardInfos[i].EQ_CinemaCode).FirstOrDefault();
                //if (eq == null) continue;
                tempRate = levenshtein(eQCin_Cinemas[i].EQ_Adress, eQCin_Cinemas[i].Forward_Adress);
                tempRate2 = levenshtein(eQCin_Cinemas[i].EQ_CinemaName, eQCin_Cinemas[i].Forward_CinemaName);
                if (eQCin_Cinemas[i].EQ_ZZCinemaName != null)
                {
                    tempRate3 = levenshtein(eQCin_Cinemas[i].EQ_ZZCinemaName, eQCin_Cinemas[i].Forward_CinemaName);
                }
                InsertCinameMatchInfoBySql("update [test].[dbo].[Sheet4$]  set [标准影院名称相似度]=" + tempRate2 + ",[专资影院名称相似度]=" + tempRate3 + ",[地址相似度]=" + tempRate + " where [专资编码]='" + eQCin_Cinemas[i].EQ_CinemaCode + "' and [发运方式] is null");
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
            string conStr = "Data Source=10.10.10.80,2433;Initial Catalog=HyFilmCopyForZyNew;UID=sa;PWD=hyby@123;";
            SqlConnection con = new SqlConnection(conStr);
            SqlCommand cmd = new SqlCommand(sql, con);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(ds);
            return ds.Tables[0];

        }
        public static void InsertCinameMatchInfoBySql(string sql)
        {
            DataSet ds = new DataSet();
            string conStr = "Data Source=10.10.10.80,2433;Initial Catalog=HyFilmCopyForZyNew;UID=sa;PWD=hyby@123;";
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
