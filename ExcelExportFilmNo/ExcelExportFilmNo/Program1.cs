using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Data;
using System.Text.RegularExpressions;
using Aspose.Cells;
using System.IO;
using System.Data;
using System.Data.SqlClient;


namespace ExcelExportFilmNo
{
    class Program1
    {
        static void Main11(string[] args)
        {
            List<char> keywords = ImportDataTable();

            string ss = "";
            for (int i = 0; i < keywords.Count; i++)
            {
                ss += "insert into SCRPROPERTY values(" + (i + 18) + ",'" + keywords[i] + "');\r\n";
            }

            Console.WriteLine(keywords.Count);
            Console.ReadLine();

        }
        public static List<char> ImportDataTable()
        {
            List<char> keywords = new List<char>();
            DataSet ds=new DataSet();
            using (
                var con =
                    new SqlConnection(
                        "Data Source=192.168.144.73;Initial Catalog=HyFilmCopyForZyNew;UID=sa;PWD=Welcome1;"))
            {
                //SqlCommand cmd = new SqlCommand("select FilmName from [A_FilmFromIssue]", con);
                SqlCommand cmd = new SqlCommand("select distinct  FilmName from [A_FilmFromExcel]", con);
                //SqlCommand cmd = new SqlCommand("select FilmName from [A_FilmFromExcel1408]", con);
                SqlDataAdapter da=new SqlDataAdapter(cmd);
                da.Fill(ds);
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    string filename = ds.Tables[0].Rows[i][0].ToString();
                    string _s = Regex.Replace(filename, @"[\u4e00-\u9fa5]", ""); //去除汉字
                    _s = Regex.Replace(_s, @"\d", "");
                    _s = Regex.Replace(_s, "[a-z]", "", RegexOptions.IgnoreCase);
                    if (_s.Trim() != "" )
                    {
                        foreach (char x in _s)
                        {
                            if (!keywords.Contains(x))
                            {
                                keywords.Add(x);
                                Console.WriteLine(x); 
                            }
                           
                        }
                       
                    }
                }
            }
            return keywords;
        }
        
    }
}
