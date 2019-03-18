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
    class Program
    {
        static string status = "Test";
        static void Main12(string[] args)
        {
            string ff = @"D:\06";
            RegisterLicense();
            List<string> excelNames = new List<string>();
            excelNames.Add("20180602182346mkkB_1.xlsx");
            excelNames.Add("20180601151914E6SK_1.xlsx");
            excelNames.Add("20180601095822fWCH_1.xlsx");
            excelNames.Add("20180601100303GPzH_1.xlsx");
            excelNames.Add("20180601151715lB2P_1.xlsx");
            excelNames.Add("20180604121311O1f5_1.xlsx");
            excelNames.Add("20180601153806Sz9N_1.xlsx");
            excelNames.Add("20180601160942rtR4_1.xlsx");
            excelNames.Add("20180601101128vB5O_1.xlsx");
            excelNames.Add("20180602230854bLV2_1.xlsx");
            excelNames.Add("20180601132415zk95_1.xlsx");
            excelNames.Add("20180605104023UwRV_1.xlsx");
            excelNames.Add("20180604104657WbBf_1.xlsx");
            excelNames.Add("20180604113037GPIU_1.xlsx");
            excelNames.Add("20180601151807vqDI_1.xlsx");
            excelNames.Add("20180602152057pspj_1.xlsx");
            excelNames.Add("20180601112730pY0j_1.xlsx");
            excelNames.Add("20180604110925YOHS_1.xlsx");
            excelNames.Add("20180601155445fnBx_1.xlsx");
            excelNames.Add("201806011443549kuG_1.xls");
            excelNames.Add("20180605094328kxsK_1.xlsx");
            excelNames.Add("201806021352390MHE_1.xlsx");
            excelNames.Add("201806041137178p8X_1.xlsx");
            excelNames.Add("20180604103439HIv5_1.xlsx");
            excelNames.Add("20180601153020XnYJ_1.xls");
            excelNames.Add("20180601132813jy7F_1.xlsx");
            excelNames.Add("20180601141956QuST_1.xlsx");
            excelNames.Add("20180601094120NIwl_1.xlsx");
            excelNames.Add("20180601115249WfOo_1.xlsx");
            excelNames.Add("20180601113523Hm6X_1.xls");
            excelNames.Add("20180601103446EDoN_1.xlsx");
            excelNames.Add("20180601104920hDy5_1.xlsx");
            excelNames.Add("201806011043525EMq_1.xlsx");
            excelNames.Add("20180601123726B558_1.xlsx");
            excelNames.Add("201806011036502ZmE_1.xlsx");
            excelNames.Add("20180601125530KwZE_1.xlsx");
            excelNames.Add("20180601162024UmWd_1.xls");
            excelNames.Add("2018060116154339lA_1.xls");
            excelNames.Add("20180601121128FbMG_1.xlsx");
            excelNames.Add("20180601122546suXy_1.xlsx");
            excelNames.Add("20180601110739i1vT_1.xlsx");
            excelNames.Add("201806041145422TLW_1.xlsx");
            excelNames.Add("20180601162713MSQP_1.xlsx");
            excelNames.Add("20180604102954SDZu_1.xlsx");
            excelNames.Add("20180602165104xpZ4_1.xlsx");
            excelNames.Add("20180601165055D5oi_1.xls");
            excelNames.Add("20180601100830Wvb3_1.xlsx");
            excelNames.Add("20180605114636Zvvl_1.xlsx");
            excelNames.Add("20180601093012G6W4_1.xlsx");
            excelNames.Add("20180604151734uce1_1.xlsx");
            excelNames.Add("20180602134509A2UI_1.xlsx");
            excelNames.Add("20180602140533KHUm_1.xlsx");
            excelNames.Add("201806021359372h0d_1.xlsx");
            excelNames.Add("20180602140407GuNh_1.xlsx");
            excelNames.Add("20180601145816chJL_1.xlsx");
            excelNames.Add("20180601145518gFnY_1.xlsx");
            excelNames.Add("20180604225813zgoj_1.xlsx");
            excelNames.Add("20180604095845XLrh_1.xlsx");
            excelNames.Add("20180601111331Rcxu_1.xls");
            excelNames.Add("20180601212003VrWy_1.xlsx");
            excelNames.Add("20180601132944hkc5_1.xlsx");
            excelNames.Add("20180603182139RWpj_1.xlsx");
            excelNames.Add("20180601142253ZPRw_1.xlsx");
            excelNames.Add("20180601100518huwx_1.xlsx");
            excelNames.Add("201806011618019Z8e_1.xlsx");
            excelNames.Add("20180601165820kVWd_1.xlsx");
            excelNames.Add("20180601100122Bcbi_1.xlsx");
            excelNames.Add("20180601091152P91T_1.xlsx");
            excelNames.Add("20180601121123aQys_1.xlsx");
            excelNames.Add("20180601123228rRM2_1.xlsx");
            excelNames.Add("20180601163403mjrq_1.xlsx");
            excelNames.Add("20180604115456omd7_1.xlsx");
            excelNames.Add("20180604101721Jljb_1.xlsx");
            excelNames.Add("20180601094133j6p3_1.xlsx");
            excelNames.Add("20180601142822aq0V_1.xlsx");
            excelNames.Add("20180601105742aIvk_1.xlsx");
            excelNames.Add("20180604105303H46g_1.xlsx");
            excelNames.Add("20180601113134oxGR_1.xlsx");
            excelNames.Add("20180601101732sikR_1.xlsx");
            excelNames.Add("20180601112534bpaU_1.xlsx");
            excelNames.Add("201806011439014ZFu_1.xlsx");
            excelNames.Add("20180601094614aAVF_1.xlsx");
            excelNames.Add("20180601163317PwjU_1.xlsx");
            excelNames.Add("20180601171210GE9C_1.xlsx");
            excelNames.Add("20180601102636aGLM_1.xlsx");
            excelNames.Add("20180601102412lD9n_1.xlsx");
            excelNames.Add("201806041116000vA5_1.xlsx");
            excelNames.Add("20180601133151Y3ns_1.xlsx");
            excelNames.Add("20180604105512ydwK_1.xlsx");
            excelNames.Add("20180602185909Papq_1.xlsx");
            excelNames.Add("20180601172430aCFm_1.xlsx");
            excelNames.Add("20180601172408L555_1.xlsx");
            excelNames.Add("20180601192155238N_1.xlsx");
            excelNames.Add("201806041450539duH_1.xlsx");
            excelNames.Add("20180601120110twFV_1.xlsx");
            excelNames.Add("20180601104958d6bQ_1.xlsx");
            excelNames.Add("20180601101730trga_1.xlsx");
            excelNames.Add("20180601095230Wqdn_1.xlsx");
            excelNames.Add("20180601174327NS11_1.xlsx");
            excelNames.Add("201806011641516tPH_1.xls");
            excelNames.Add("20180601162950w9qF_1.xlsx");
            excelNames.Add("201806011617097BqB_1.xlsx");
            excelNames.Add("20180601152452DKEQ_1.xlsx");
            excelNames.Add("20180601151612hohk_1.xlsx");
            excelNames.Add("201806011617250Ffk_1.xls");
            excelNames.Add("20180604105625i7Ux_1.xlsx");
            excelNames.Add("20180601142235kY5p_1.xlsx");
            excelNames.Add("201806011620462RcA_1.xlsx");
            excelNames.Add("20180604110815nHVw_1.xlsx");
            excelNames.Add("20180606102714K3z4_1.xlsx");
            excelNames.Add("201806011657389EAl_1.xlsx");
            excelNames.Add("20180603120043Q0gJ_1.xlsx");
            excelNames.Add("20180601094234tZJO_1.xlsx");
            excelNames.Add("201806011134302KKE_1.xlsx");
            excelNames.Add("20180601110726gITU_1.xlsx");
            excelNames.Add("20180601102928954F_1.xlsx");
            excelNames.Add("20180601103944Bjol_1.xlsx");
            excelNames.Add("20180601143703qHQ3_1.xlsx");
            excelNames.Add("20180604110130KBQs_1.xlsx");
            excelNames.Add("20180601100023q5aS_1.xlsx");
            excelNames.Add("20180601110451aW5e_1.xls");
            excelNames.Add("20180601171733YcDN_1.xlsx");
            excelNames.Add("20180601170919I8vU_1.xlsx");
            excelNames.Add("201806041549216Gr6_1.xlsx");
            excelNames.Add("20180604123645mGii_1.xlsx");
            excelNames.Add("20180602091828bWOk_1.xlsx");
            excelNames.Add("20180604155332chud_1.xlsx");
            excelNames.Add("20180601175353hg43_1.xlsx");
            excelNames.Add("20180604181037wSa7_1.xlsx");
            excelNames.Add("20180601152028yHzW_1.xlsx");
            excelNames.Add("20180601160643vhsg_1.xlsx");
            excelNames.Add("20180601145813oDpY_1.xlsx");
            excelNames.Add("20180601170735t458_1.xls");
            excelNames.Add("20180601100630L12W_1.xlsx");
            excelNames.Add("20180601101634Ffus_1.xlsx");
            excelNames.Add("201806041125374xsk_1.xlsx");
            excelNames.Add("20180605094351xMt0_1.xlsx");
            excelNames.Add("20180604132259lAgq_1.xlsx");
            excelNames.Add("20180601093825mqrH_1.xlsx");
            excelNames.Add("20180605145014LgX8_1.xlsx");
            excelNames.Add("201806011526357kwr_1.xlsx");
            excelNames.Add("20180601101019sZzX_1.xlsx");
            excelNames.Add("20180604165510uPgi_1.xlsx");
            excelNames.Add("20180604142738h1ZB_1.xlsx");
            excelNames.Add("20180601105135nQmH_1.xlsx");
            excelNames.Add("20180601100007ZnIV_1.xlsx");
            excelNames.Add("20180601102850khbJ_1.xlsx");
            excelNames.Add("20180606092501ijHD_1.xlsx");
            excelNames.Add("20180604124213Ajgs_1.xlsx");
            excelNames.Add("20180601132303itg3_1.xlsx");
            excelNames.Add("20180604140458lL5h_1.xlsx");
            excelNames.Add("20180601113308ppiN_1.xlsx");
            excelNames.Add("20180601184417k5ry_1.xlsx");
            excelNames.Add("20180604164643835c_1.xlsx");
            //Console.WriteLine(excelNames.Count);
            //Console.ReadLine();
            //return;
            var i = 1;
            foreach (var excelName in excelNames)
            {
                Workbook v = new Workbook(ff + @"\" + excelName);
                foreach (Worksheet st in v.Worksheets)
                {
                    Cells cells = st.Cells;
                    try
                    {
                        ImportDataTable(cells, excelName);
                        Console.WriteLine(i + ":" + excelName);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                        throw;
                    }

                }
                i++;
            }

            Console.WriteLine(excelNames.Count);
            //Console.ReadLine();

        }
        public static void ImportDataTable(Cells c, string excelName)
        {
            DataTable tableData = ImportWorkSheet(c, excelName);
            var conStr = "Data Source=192.168.11.213;Initial Catalog=test;UID=sa;PWD=Welcome1;";
            using (SqlBulkCopy sqlRevdBulkCopy = new SqlBulkCopy(conStr))//引用SqlBulkCopy  
            {
                sqlRevdBulkCopy.BulkCopyTimeout = 10000;
                sqlRevdBulkCopy.DestinationTableName = "A_FilmFromExcel06";//数据库中对应的表名  
                sqlRevdBulkCopy.NotifyAfter = tableData.Rows.Count;//有几行数据  
                for (var j = 0; j < tableData.Columns.Count; j++)
                {
                    sqlRevdBulkCopy.ColumnMappings.Add(tableData.Columns[j].ColumnName, tableData.Columns[j].ColumnName);
                }
                sqlRevdBulkCopy.WriteToServer(tableData);//数据导入数据库  
                sqlRevdBulkCopy.Close();//关闭连接  
            }
            //using (var con = new SqlConnection("Data Source=10.1.31.112;Initial Catalog=HyFilmCopyForZyNew20181205;UID=sa;PWD=Welcome1;"))
            //{
            //    con.Open();
            //    var cmdText = "INSERT INTO [A_FilmFromExcel06] VALUES( ";
            //    for (var i = 0; i < tableData.Columns.Count; i++)
            //    {
            //        cmdText += "@v" + i + ",";
            //    }
            //    cmdText += "getdate())";
            //    var cmd = new SqlCommand(cmdText, con);

            //    for (var j = 0; j < tableData.Rows.Count; j++)
            //    {
            //        cmd.Parameters.Clear();
            //        for (var i = 0; i < tableData.Columns.Count; i++)
            //        {
            //            var par = new SqlParameter("@v" + i, SqlDbType.VarChar)
            //            {
            //                Direction = ParameterDirection.Input,
            //                Value = tableData.Rows[j][i].ToString()
            //            };
            //            cmd.Parameters.Add(par);
            //        }
            //        cmd.ExecuteNonQuery();
            //    }
            //    con.Close();
            //}
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
                    dataRow["rowid"] = i;
                }
                else
                {
                    dataRow["rowid"] = cells[i, RowIDCell].Value;
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
                dataRow["Ctime"] = "2018-06-01";
                tb.Rows.Add(dataRow);
            }
            return tb;
        }

        private static DataTable CreateTableSchema()
        {
            DataTable tb = new DataTable();
            tb.Columns.Add("rowid");
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
            tb.Columns.Add("Ctime");
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
