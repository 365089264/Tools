using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using Aspose.Cells;
using System.IO;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using Oracle.ManagedDataAccess.Types;
using System.Diagnostics;


namespace ExcelExportFilmNo
{
    class Program3
    {
        /*
          SELECT 'excelNames.Add("'||attatitle||'");'
            FROM attachment where attatype='monthTicket' and feeyear=2018 and feemonth=10 order by createtime;
        */
        static string status = "11";
        static void Main3(string[] args)
        {
            string ff = @"D:\" + status;
            RegisterLicense();
            List<string> excelNames = new List<string>();
            excelNames.Add("20181101091736nBfN_1.xlsx");
            excelNames.Add("20181101093846BJiN_1.xlsx");
            excelNames.Add("20181101095555Yiy2_1.xlsx");
            excelNames.Add("20181101100407eyT3_1.xlsx");
            excelNames.Add("20181101100909PKCt_1.xlsx");
            excelNames.Add("20181101101938r1gq_1.xlsx");
            excelNames.Add("20181101103012IT6h_1.xlsx");
            excelNames.Add("20181101103120JaBS_1.xlsx");
            excelNames.Add("20181101103141AsYd_1.xls");
            excelNames.Add("201811011039182nTB_1.xlsx");
            excelNames.Add("20181101103957kOcf_1.xlsx");
            excelNames.Add("20181101104459LhUq_1.xlsx");
            excelNames.Add("20181101104523Al3X_1.xlsx");
            excelNames.Add("20181101105607WOLZ_1.xls");
            excelNames.Add("20181101110506I3wO_1.xlsx");
            excelNames.Add("20181101110803L0s6_1.xlsx");
            excelNames.Add("20181101111036dYjn_1.xls");
            excelNames.Add("20181101111047d3Mn_1.xlsx");
            excelNames.Add("20181101111904RNnY_1.xlsx");
            excelNames.Add("20181101112154tpmU_1.xlsx");
            excelNames.Add("20181101112336h41H_1.xlsx");
            excelNames.Add("20181101112536hQUD_1.xlsx");
            excelNames.Add("20181101112547YheH_1.xlsx");
            excelNames.Add("20181101112721DuBM_1.xlsx");
            excelNames.Add("20181101113014d5Ez_1.xls");
            excelNames.Add("20181101113313L0NH_1.xlsx");
            excelNames.Add("20181101113438hkRO_1.xlsx");
            excelNames.Add("20181101113541R8My_1.xlsx");
            excelNames.Add("201811011137072B37_1.xls");
            excelNames.Add("20181101113921rJdN_1.xlsx");
            excelNames.Add("20181101113936GSmr_1.xlsx");
            excelNames.Add("20181101114212uqok_1.xlsx");
            excelNames.Add("20181101114238XQF0_1.xlsx");
            excelNames.Add("201811011143406fXd_1.xls");
            excelNames.Add("20181101114630RqGM_1.xlsx");
            excelNames.Add("20181101114930bne0_1.xlsx");
            excelNames.Add("20181101115037OBpU_1.xlsx");
            excelNames.Add("20181101115150ysjv_1.xls");
            excelNames.Add("20181101115221vP10_1.xlsx");
            excelNames.Add("20181101115639ffO6_1.xlsx");
            excelNames.Add("201811011158446mFB_1.xlsx");
            excelNames.Add("20181101120213Yui9_1.xlsx");
            excelNames.Add("20181101120240bXa4_1.xls");
            excelNames.Add("20181101120259c0Sb_1.xls");
            excelNames.Add("20181101120906NtcE_1.xls");
            excelNames.Add("201811011215479euP_1.xlsx");
            excelNames.Add("20181101122645Oox5_1.xlsx");
            excelNames.Add("20181101123234vSxM_1.xlsx");
            excelNames.Add("20181101123337FWFz_1.xlsx");
            excelNames.Add("20181101123559HBRr_1.xlsx");
            excelNames.Add("20181101123803Z7n2_1.xlsx");
            excelNames.Add("20181101124007i30m_1.xlsx");
            excelNames.Add("20181101124034BuQg_1.xlsx");
            excelNames.Add("20181101125545wIps_1.xlsx");
            excelNames.Add("20181101131153RTpg_1.xlsx");
            excelNames.Add("20181101133023Any4_1.xlsx");
            excelNames.Add("20181101134651X9mN_1.xlsx");
            excelNames.Add("20181101140417szFv_1.xlsx");
            excelNames.Add("20181101140528JtQq_1.xlsx");
            excelNames.Add("20181101140739mjAF_1.xls");
            excelNames.Add("20181101141652Agbl_1.xlsx");
            excelNames.Add("20181101141819BJ4i_1.xlsx");
            excelNames.Add("20181101143012W17P_1.xlsx");
            excelNames.Add("20181101143138cFM0_1.xlsx");
            excelNames.Add("20181101145601nsDq_1.xlsx");
            excelNames.Add("20181101152013eGTj_1.xlsx");
            excelNames.Add("20181101154345vc68_1.xls");
            excelNames.Add("20181101155627x0Gn_1.xlsx");
            excelNames.Add("20181101160307xXXJ_1.xlsx");
            excelNames.Add("20181101160754g4Xd_1.xlsx");
            excelNames.Add("20181101162245VPx1_1.xlsx");
            excelNames.Add("20181101162527NqF2_1.xlsx");
            excelNames.Add("20181101163048TPru_1.xls");
            excelNames.Add("20181101163451Fmlv_1.xlsx");
            excelNames.Add("201811011636207l9p_1.xlsx");
            excelNames.Add("20181101171025dR59_1.xlsx");
            excelNames.Add("2018110117114792VH_1.xls");
            excelNames.Add("20181101171309RfYc_1.xlsx");
            excelNames.Add("20181101171656MGNv_1.xlsx");
            excelNames.Add("20181101171916V6bt_1.xls");
            excelNames.Add("20181101174914sDR6_1.xlsx");
            excelNames.Add("20181101181746gCHz_1.xls");
            excelNames.Add("20181101200612URuH_1.xlsx");
            excelNames.Add("20181102103135nXtA_1.xlsx");
            excelNames.Add("20181102105403jt9Q_1.xlsx");
            excelNames.Add("20181102111653vdEw_1.xlsx");
            excelNames.Add("20181102133657SHoS_1.xlsx");
            excelNames.Add("2018110214184775id_1.xlsx");
            excelNames.Add("20181102145049AYGv_1.xlsx");
            excelNames.Add("201811021510016NZA_1.xlsx");
            excelNames.Add("201811021520312fkB_1.xlsx");
            excelNames.Add("20181102152127niT7_1.xlsx");
            excelNames.Add("20181102152233v94s_1.xlsx");
            excelNames.Add("20181102152327HoJp_1.xlsx");
            excelNames.Add("20181102152413Hr8z_1.xlsx");
            excelNames.Add("20181102152510USyl_1.xlsx");
            excelNames.Add("20181102152710bzBV_1.xlsx");
            excelNames.Add("20181102165547H8t8_1.xlsx");
            excelNames.Add("20181102170649sxTZ_1.xlsx");
            excelNames.Add("20181102171836CCto_1.xlsx");
            excelNames.Add("20181103101624BZao_1.xlsx");
            excelNames.Add("20181103162727wg4U_1.xlsx");
            excelNames.Add("20181105093033zguP_1.xlsx");
            excelNames.Add("201811051010293Nsi_1.xlsx");
            excelNames.Add("20181105101049crxX_1.xlsx");
            excelNames.Add("20181105101421v3eb_1.xls");
            excelNames.Add("20181105102426lYef_1.xlsx");
            excelNames.Add("20181105110420RAHk_1.xlsx");
            excelNames.Add("20181105111202IfhB_1.xlsx");
            excelNames.Add("20181105114217KwZC_1.xlsx");
            excelNames.Add("20181105120848ABuI_1.xlsx");
            excelNames.Add("20181105150013OIwY_1.xlsx");
            excelNames.Add("20181105160308R5Y4_1.xlsx");
            excelNames.Add("201811051607110nfx_1.xlsx");
            excelNames.Add("20181105165654jFZr_1.xlsx");
            excelNames.Add("20181105170109N91C_1.xls");
            excelNames.Add("20181106085534GciH_1.xlsx");
            excelNames.Add("20181106085611xbkB_1.xlsx");
            excelNames.Add("20181106095256QwHY_1.xlsx");
            excelNames.Add("20181106104049LRhk_1.xlsx");
            excelNames.Add("20181106112306oMMJ_1.xlsx");
            excelNames.Add("20181106113321CDqE_1.xls");
            excelNames.Add("20181106113419BnYS_1.xlsx");
            excelNames.Add("20181106113427hyVj_1.xlsx");
            excelNames.Add("20181106113941GDb7_1.xlsx");
            excelNames.Add("201811061217536dkZ_1.xls");
            excelNames.Add("20181106121948d4hy_1.xlsx");
            excelNames.Add("20181106134517YJgK_1.xlsx");
            excelNames.Add("20181106141408r8BQ_1.xlsx");
            excelNames.Add("20181106141750AyQi_1.xls");
            excelNames.Add("20181106141801eeRJ_1.xlsx");
            excelNames.Add("201811061422316HbC_1.xls");
            excelNames.Add("20181106142621wsSv_1.xlsx");
            excelNames.Add("20181106142712uLUm_1.xls");
            excelNames.Add("20181106143207ekZq_1.xls");
            excelNames.Add("20181106143333fCWm_1.xlsx");
            excelNames.Add("20181106143409yxiM_1.xlsx");
            excelNames.Add("20181106145446D391_1.xlsx");
            excelNames.Add("20181106145923vucM_1.xlsx");
            excelNames.Add("20181106150823bKXD_1.xlsx");
            excelNames.Add("20181106151039SFZ8_1.xlsx");
            excelNames.Add("20181106152010nY47_1.xls");
            excelNames.Add("20181106152206PF0g_1.xls");
            excelNames.Add("20181106152458nMpK_1.xlsx");
            excelNames.Add("20181106152553Sl5m_1.xlsx");
            excelNames.Add("201811061529599322_1.xls");
            excelNames.Add("201811061545264VLC_1.xlsx");
            excelNames.Add("20181106155046m4qt_1.xlsx");
            excelNames.Add("20181106155931zLZw_1.xlsx");
            excelNames.Add("20181106160249QNHD_1.xlsx");
            excelNames.Add("20181106160504CZ3r_1.xlsx");
            excelNames.Add("20181106162929T3BG_1.xlsx");
            excelNames.Add("20181106164931pUIS_1.xlsx");
            excelNames.Add("20181106172516iIFV_1.xlsx");
            excelNames.Add("20181106174156D46W_1.xls");
            excelNames.Add("20181106175447Ip9y_1.xls");
            excelNames.Add("20181106180431ijbJ_1.xlsx");
            excelNames.Add("20181107100249cWHH_1.xlsx");
            excelNames.Add("20181107100914wj0S_1.xlsx");
            excelNames.Add("20181107101930ObAn_1.xlsx");
            excelNames.Add("2018110710533365vC_1.xls");
            excelNames.Add("201811071107429Xim_1.xls");
            excelNames.Add("20181107111119G7ia_1.xls");
            excelNames.Add("20181107111958ZxmU_1.xlsx");
            excelNames.Add("20181107115224ZgCy_1.xlsx");
            excelNames.Add("20181107140311KnIZ_1.xlsx");
            excelNames.Add("20181107140550A88H_1.xls");
            excelNames.Add("20181107144732uWhA_1.xlsx");
            excelNames.Add("20181107145938FjuO_1.xlsx");
            excelNames.Add("20181107162627mmPg_1.xlsx");
            excelNames.Add("20181107162747rnPN_1.xlsx");
            excelNames.Add("201811071711055mNf_1.xlsx");
            excelNames.Add("20181107171209x3X7_1.xlsx");
            excelNames.Add("20181107173553WfX7_1.xls");
            excelNames.Add("20181108090937FuzR_1.xls");
            excelNames.Add("201811081433030oFk_1.xlsx");
            excelNames.Add("20181108152250kTrb_1.xlsx");
            excelNames.Add("20181108194410glme_1.xlsx");
            excelNames.Add("20181109101106KCHA_1.xlsx");
            excelNames.Add("20181109103033fyf6_1.xls");
            excelNames.Add("20181109105429G4vv_1.xlsx");
            excelNames.Add("20181109112132DmE3_1.xlsx");
            excelNames.Add("20181109112201shH2_1.xlsx");
            excelNames.Add("20181109112235QOLC_1.xlsx");
            excelNames.Add("20181109112257YOCY_1.xlsx");
            excelNames.Add("20181109112317D3BH_1.xlsx");
            foreach (var excelName in excelNames)
            {
                Workbook v = new Workbook(ff + @"\" + excelName);
                Cells cells = v.Worksheets[0].Cells;
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

            Console.WriteLine(excelNames.Count);
            Console.ReadLine();

        }
        public static void ImportDataTable(Cells c, string excelName)
        {
            DataTable tableData = ImportWorkSheet(c, excelName);
            using (var con = new OracleConnection("Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.10.130.23)(PORT=8903))(CONNECT_DATA=(SERVICE_NAME = settle_primary)));User Id=settle;Password=settle;"))
            {
                con.Open();
                var cmdText = "INSERT INTO ATEST_D VALUES( '" + excelName + "'";
                for (var i = 0; i < tableData.Columns.Count; i++)
                {
                    cmdText += ",:v" + i;
                }
                cmdText += ")";
                var cmd = new OracleCommand(cmdText, con);
                cmd.ArrayBindCount = tableData.Rows.Count;
                cmd.Parameters.Clear();
                int[] v0 = new int[tableData.Rows.Count];
                string[] v1 = new string[tableData.Rows.Count];
                string[] v2 = new string[tableData.Rows.Count];
                string[] v3 = new string[tableData.Rows.Count];
                for (var j = 0; j < tableData.Rows.Count; j++)
                {
                    v0[j] = Convert.ToInt32(tableData.Rows[j][0].ToString());
                    v1[j] = tableData.Rows[j][1].ToString();
                    v2[j] = tableData.Rows[j][2].ToString();
                    v3[j] = tableData.Rows[j][3].ToString();

                }

                var par = new OracleParameter(":v0", v0) { };
                cmd.Parameters.Add(par);
                var par1 = new OracleParameter(":v1", v1) { };
                cmd.Parameters.Add(par1);
                var par2 = new OracleParameter(":v2", v2) { };
                cmd.Parameters.Add(par2);
                var par3 = new OracleParameter(":v3", v3) { };
                cmd.Parameters.Add(par3);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                cmd.ExecuteNonQuery();
                sw.Stop();
                con.Close();
            }
        }

        public static DataTable ImportWorkSheet(Cells cells, string excelName)
        {
            DataTable tb = CreateTableSchema();
            int startRowIndex = 1;
            int RowIDCell = -1;
            int FilmNoCell = 2;
            int FilmNameCell = 3;
            int FilmissuetypeCell = 1;
            bool isFindStartRow = false;
            for (int j = 0; j < 10; j++)
            {
                if (isFindStartRow) break;
                for (int i = 0; i < 19; i++)
                {
                    if (cells[j, i].Value != null)
                    {
                        if (cells[j, i].Value.ToString().Contains("影片名称") || cells[j, i].Value.ToString() == "影片")
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
                        if (cells[j, i].Value.ToString().Contains("放映版本"))
                        {
                            FilmissuetypeCell = i;
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
                dataRow["FilmName"] = cells[i, FilmNameCell].Value;
                string sourceCode = cells[i, FilmNoCell].Value == null ? null : cells[i, FilmNoCell].Value.ToString();
                if (string.IsNullOrEmpty(sourceCode) || sourceCode.Length == 12)
                {
                    dataRow["FilmNo"] = sourceCode;
                }
                else if (sourceCode.Length < 12)
                {
                    dataRow["FilmNo"] = sourceCode.PadLeft(12, '0');
                }
                else if (sourceCode.Length > 12)
                {
                    dataRow["FilmNo"] = sourceCode.Substring(0, 12);
                }
                dataRow["Filmissuetype"] = cells[i, FilmissuetypeCell].Value;
                tb.Rows.Add(dataRow);
            }
            return tb;
        }

        private static DataTable CreateTableSchema()
        {
            DataTable tb = new DataTable();
            tb.Columns.Add("RowID");
            tb.Columns.Add("FilmName");
            tb.Columns.Add("FilmNo");
            tb.Columns.Add("Filmissuetype");
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
