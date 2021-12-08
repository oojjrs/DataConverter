using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Data;
using System.Runtime.InteropServices;

namespace DataConverter
{
    public class ExcelFile
    {
        private string _path;

        public ExcelFile(string path)
        {
            _path = path;
        }

        public DataSet Import()
        {
            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Workbook ex = null;
            try
            {
                app = new Excel.Application();
                app.DisplayAlerts = false;
                wb = app.Workbooks.Open(_path);

                return Import(wb);
            }
            finally
            {
                if (ex != null)
                {
                    ex.Close();
                    Marshal.ReleaseComObject(ex);
                }
                if (wb != null)
                {
                    wb.Close();
                    Marshal.ReleaseComObject(wb);
                }
                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }
                GC.Collect();
            }
        }

        private DataSet Import(Excel.Workbook workbook)
        {
            var ds = new DataSet();

            foreach (Excel.Worksheet sheet in workbook.Sheets)
            {
                if (sheet.Name.StartsWith("#"))
                    continue;

                var dt = new DataTable(sheet.Name);
                var range = sheet.UsedRange;

                for (int col = 1; col <= range.Columns.Count; ++col)
                    dt.Columns.Add((sheet.Cells[1, col] as Excel.Range).Value2);

                // 첫 번째 줄은 컬럼명이기 때문에 실제 데이터는 rows count - 1
                for (int row = 0; row < range.Rows.Count - 1; ++row)
                {
                    var dr = dt.NewRow();
                    for (int col = 0; col < range.Columns.Count; ++col)
                        dr[col] = (range.Cells[row + 2, col + 1] as Excel.Range).Value2;
                    dt.Rows.Add(dr);
                }

                ds.Tables.Add(dt);
                Marshal.ReleaseComObject(sheet);
            }

            return ds;
        }
    }
}
