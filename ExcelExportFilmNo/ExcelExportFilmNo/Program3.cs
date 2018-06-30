using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using Aspose.Cells;
using System.IO;
using System.Data;
using System.Data.SqlClient;


namespace ExcelExportFilmNo
{
    class Program3
    {
        static string status = "20180606审核通过月结算表";
        static void Main(string[] args)
        {
            string ff = @"D:\" + status;
            RegisterLicense();
            List<string> excelNames = new List<string>();
            excelNames.Clear();
            string dirp = Path.Combine(ff); //有一个posed文件夹在目录文件夹下  
            DirectoryInfo mydir = new DirectoryInfo(dirp);
            foreach (FileSystemInfo fsi in mydir.GetFileSystemInfos())
            {
                if (fsi is FileInfo)
                {
                    string s = System.IO.Path.GetExtension(fsi.FullName); //获取扩展名  
                    if (s == ".xlsx" || s == ".xls")
                    {
                        excelNames.Add(fsi.Name);
                        //Console.WriteLine(fsi.Name);
                    }
                    else
                    {
                        Console.WriteLine(fsi.Name);
                    }
                }
            }
            //Console.WriteLine(excelNames.Count);
            //Console.ReadLine();
            //return;
            foreach (var excelName in excelNames)
            {
                Workbook v = new Workbook(ff+@"\" + excelName);
                foreach (Worksheet st in v.Worksheets)
                {
                    Cells cells = st.Cells;
                    try
                    {
                        ImportDataTable(cells, excelName);
                        Console.WriteLine(excelName);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                        throw;
                    }

                }
            }

            Console.WriteLine(excelNames.Count);
            //Console.ReadLine();

        }
        public static void ImportDataTable(Cells c, string excelName)
        {
            DataTable tableData = ImportWorkSheet(c, excelName);
            using (var con = new SqlConnection("Data Source=192.168.144.73;Initial Catalog=HyFilmCopyForZyNew;UID=sa;PWD=Welcome1;"))
            {
                con.Open();
                var cmdText = "INSERT INTO [A_FilmFromExcel0606] VALUES( ";
                for (var i = 0; i < tableData.Columns.Count; i++)
                {
                    cmdText += "@v" + i + ",";
                }
                cmdText += "getdate())";
                var cmd = new SqlCommand(cmdText, con);

                for (var j = 0; j < tableData.Rows.Count; j++)
                {
                    cmd.Parameters.Clear();
                    for (var i = 0; i < tableData.Columns.Count; i++)
                    {
                        var par = new SqlParameter("@v" + i, SqlDbType.VarChar)
                        {
                            Direction = ParameterDirection.Input,
                            Value = tableData.Rows[j][i].ToString()
                        };
                        cmd.Parameters.Add(par);
                    }
                    cmd.ExecuteNonQuery();
                }
                con.Close();
            }
        }

        public static DataTable ImportWorkSheet(Cells cells, string excelName)
        {
            DataTable tb = CreateTableSchema();
            int startRowIndex = 1;
            int RowIDCell = -1;
            int FilmNoCell = 2;
            int FilmNameCell = 1;
            int CinemaNameCell = 3;
            int CinemaCodeCell = 4;
            int DeviceBelongCell = 0;
            int StartDateCell = 0;
            int EndDateCell = 0;
            int TotalCCCell = 0;
            int TotalRCCell = 0;
            int TotalPFCell = 0;
            int DianZZJJCell = 0;
            int ZZRateCell = 0;
            int SJCell = 0;
            int JPFCell = 0;
            int FZBLCell = 0;
            int FZPKCell = 0;
            bool isFindStartRow = false;
            for (int j = 0; j < 10; j++)
            {
                if (isFindStartRow) break;
                for (int i = 0; i < 19; i++)
                {
                    if (cells[j, i].Value != null)
                    {
                        if (cells[j, i].Value.ToString().Contains("影片名称"))
                        {
                            FilmNameCell = i;
                            isFindStartRow = true;
                            startRowIndex = j + 1;
                            continue;
                        }
                        if (cells[j, i].Value.ToString().Contains("影片编码"))
                        {
                            FilmNoCell = i;
                            continue;
                        }
                        if (cells[j, i].Value.ToString().Contains("影院名称"))
                        {
                            CinemaNameCell = i;
                            continue;
                        }
                        if (cells[j, i].Value.ToString().Contains("影院编码"))
                        {
                            CinemaCodeCell = i;
                            continue;
                        }
                        if (cells[j, i].Value.ToString().Contains("设备归属"))
                        {
                            DeviceBelongCell = i;
                            continue;
                        }
                        if (cells[j, i].Value.ToString().Contains("开始日期"))
                        {
                            StartDateCell = i;
                            continue;
                        }
                        if (cells[j, i].Value.ToString().Contains("结束日期"))
                        {
                            EndDateCell = i;
                            continue;
                        }
                        if (cells[j, i].Value.ToString().Contains("总场次"))
                        {
                            TotalCCCell = i;
                            continue;
                        }
                        if (cells[j, i].Value.ToString().Contains("总人次"))
                        {
                            TotalRCCell = i;
                            continue;
                        }
                        if (cells[j, i].Value.ToString().Contains("总票房"))
                        {
                            TotalPFCell = i;
                            continue;
                        }
                        if (cells[j, i].Value.ToString().Contains("电影专项基金"))
                        {
                            DianZZJJCell = i;
                            continue;
                        }
                        if (cells[j, i].Value.ToString().Contains("增值税率"))
                        {
                            ZZRateCell = i;
                            continue;
                        }
                        if (cells[j, i].Value.ToString().Contains("税金"))
                        {
                            SJCell = i;
                            continue;
                        }
                        if (cells[j, i].Value.ToString().Contains("净票房"))
                        {
                            JPFCell = i;
                            continue;
                        }
                        if (cells[j, i].Value.ToString().Contains("分账比例"))
                        {
                            FZBLCell = i;
                            continue;
                        }
                        if (cells[j, i].Value.ToString().Contains("分账片款"))
                        {
                            FZPKCell = i;
                            continue;
                        }
                        if (cells[j, i].Value.ToString().Contains("序号"))
                        {
                            RowIDCell = i;
                            
                            continue;
                        }
                    }
                }

            }
            for (int i = startRowIndex; i < cells.Rows.Count; i++)
            {
                DataRow dataRow = tb.NewRow();
                if (RowIDCell == -1)
                {
                    dataRow["RowID"] = i;
                }
                else
                {
                    dataRow["RowID"] = cells[i, RowIDCell].Value;
                }
                
                dataRow["FilmNo"] = cells[i, FilmNoCell].Value;
                dataRow["FilmName"] = cells[i, FilmNameCell].Value;
                dataRow["CinemaName"] = cells[i, CinemaNameCell].Value;
                dataRow["CinemaCode"] = cells[i, CinemaCodeCell].Value;
                dataRow["DeviceBelong"] = cells[i, DeviceBelongCell].Value;
                dataRow["StartDate"] = cells[i, StartDateCell].Value;
                dataRow["EndDate"] = cells[i, EndDateCell].Value;
                dataRow["ExcelName"] = excelName;
                dataRow["TotalCC"] = cells[i, TotalCCCell].Value;
                dataRow["TotalRC"] = cells[i, TotalRCCell].Value;
                dataRow["TotalPF"] = cells[i, TotalPFCell].Value;
                dataRow["DianZZJJ"] = cells[i, DianZZJJCell].Value;
                dataRow["ZZRate"] = cells[i, ZZRateCell].Value;
                dataRow["SJ"] = cells[i, SJCell].Value;
                dataRow["JPF"] = cells[i, JPFCell].Value;
                dataRow["FZBL"] = cells[i, FZBLCell].Value;
                dataRow["FZPK"] = cells[i, FZPKCell].Value;
                dataRow["ExcelNameUser"] = status;
                dataRow["IsVisible"] = 0;
                //dataRow["rn"] = 0;
                //dataRow["mtime"] = "2018-06-03";
                tb.Rows.Add(dataRow);
            }
            return tb;
        }

        private static DataTable CreateTableSchema()
        {
            DataTable tb = new DataTable();
            tb.Columns.Add("RowID");
            tb.Columns.Add("FilmNo");
            tb.Columns.Add("FilmName");
            tb.Columns.Add("CinemaName");
            tb.Columns.Add("CinemaCode");
            tb.Columns.Add("DeviceBelong");
            tb.Columns.Add("StartDate");
            tb.Columns.Add("EndDate");
            tb.Columns.Add("ExcelName");
            tb.Columns.Add("TotalCC");
            tb.Columns.Add("TotalRC");
            tb.Columns.Add("TotalPF");
            tb.Columns.Add("DianZZJJ");
            tb.Columns.Add("ZZRate");
            tb.Columns.Add("SJ");
            tb.Columns.Add("JPF");
            tb.Columns.Add("FZBL");
            tb.Columns.Add("FZPK");
            tb.Columns.Add("ExcelNameUser");
            tb.Columns.Add("IsVisible");
            //tb.Columns.Add("rn");
            //tb.Columns.Add("mtime");
            return tb;
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
