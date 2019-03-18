using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using Oracle.ManagedDataAccess.Types;
using OracleCommand = Oracle.ManagedDataAccess.Client.OracleCommand;
using OracleConnection = Oracle.ManagedDataAccess.Client.OracleConnection;
using OracleDataAdapter = Oracle.ManagedDataAccess.Client.OracleDataAdapter;
using OracleParameter = Oracle.ManagedDataAccess.Client.OracleParameter;

namespace ExcelExportFilmNo
{
    class DataExportExcel
    {
        static void Main(string[] args)
        {
            //全国院线
            OracleParameter[] oracleParameters = new OracleParameter[5];
            oracleParameters[0] = new OracleParameter("P_tongjiDay", OracleDbType.Date) { Value = (OracleDate)Convert.ToDateTime("2018-10-25") };
            oracleParameters[1] = new OracleParameter("P_startDay", OracleDbType.Date) { Value = (OracleDate)Convert.ToDateTime("2018-04-25") };
            oracleParameters[2] = new OracleParameter("P_endDay", OracleDbType.Date, 100) { Value = (OracleDate)Convert.ToDateTime("2018-10-25") };
            oracleParameters[3] = new OracleParameter("P_fileName", OracleDbType.Varchar2, 50) { Value = "无双" };
            oracleParameters[4] = new OracleParameter("P_CUR", OracleDbType.RefCursor) { Direction = ParameterDirection.Output };
            DataTable dt = OracleDBHelper.ExecToStoredProcedureGetTable("SP_GETREPORT_CINEMACODE", oracleParameters);
            Console.WriteLine(dt.Rows.Count);
             //* */
            Console.ReadLine();
        }
    }
}
