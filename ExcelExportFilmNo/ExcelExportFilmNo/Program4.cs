using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Cells;

namespace ExcelExportFilmNo
{
    class Program4
    {
        static void Main333(string[] args)
        {
            Workbook v = new Workbook(@"D:\院线影片分影厅查询 (8).xls");
            Cells cells = v.Worksheets[0].Cells;
            int ccI, rsI; decimal pfI;
            for (int i = 1; i < cells.Rows.Count; i++)
            {
                string cc = cells[i, 7].Value.ToString();
                string rs = cells[i, 8].Value.ToString();
                string pf = cells[i, 9].Value.ToString();

                if (!Int32.TryParse(cc, out ccI) || string.IsNullOrEmpty(cc))
                {
                    Console.WriteLine("rows:" + i + " cc:" + cc);
                }
                if (!Int32.TryParse(rs, out rsI) || string.IsNullOrEmpty(rs))
                {
                    Console.WriteLine("rows:" + i + " rs:" + rs);
                }
                if (!Decimal.TryParse(pf, out pfI) || string.IsNullOrEmpty(pf))
                {
                    Console.WriteLine("rows:" + i + " pf:" + pf);
                }
            }
            Console.WriteLine(cells.Rows.Count);
            Console.ReadLine();
        }
    }
}
