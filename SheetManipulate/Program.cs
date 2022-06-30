using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Runtime.CompilerServices;


namespace myProject
{
     class Excel
    {

        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

         private static void Main()
        {
            var excelProcessID = new List<int>();

            
            foreach (Process p in Process.GetProcessesByName("EXCEL"))
            {
                excelProcessID.Add(p.Id);
            }

            

            OpenFlie();
            KillProcess();


            void OpenFlie()
            {
                Excel excel = new Excel(@$"C:\Users\andre\Desktop\Lista de emails.xlsx", 1);
                
                Console.WriteLine(excel.ReadCell(0, 0));
                excel.WriteCell("C#", 1, 0); // text to write, row, column
                Console.WriteLine(excel.ReadCell(1, 0));

                string term = @$"gmail";
                int count = 1;
                do
                {
                    if (excel.ReadCell(count,1).ToString().Contains(term) == true)
                    {
                        excel.WriteCell(excel.ReadCell(count, 1),count,1);
                    }
                    count++;
                } while (count <= excel.LastRow());


                excel.SaveFile();
            }

            void KillProcess()
            {

                int count = 0;
                foreach (Process xlProcess in Process.GetProcessesByName("EXCEL"))
                {
                    foreach (int xlid in excelProcessID)
                    {
                        if (xlProcess.Id == xlid)
                        {
                            count++;
                        }
                    }

                    if (count == 0)
                    {
                        xlProcess.Kill();
                    }
                    count = 0;
                }
            }


            //int count = 0;
            //foreach (Process xlProcess in Process.GetProcessesByName("EXCEL"))
            //{
            //    foreach (int xlid in excelProcessID)
            //    {
            //        if (xlProcess.Id == xlid)
            //        {
            //            count++;
            //        }
            //    }

            //    if (count == 0)
            //    {
            //        xlProcess.Kill();
            //    }
            //    count = 0;
            //}

        }
        
      
        public Excel(string path,int Sheet)
        {

            
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
            
        }

        public string ReadCell(int i, int j)
        {
            
            i++;
            j++;

            
            if (ws.Cells[i, j].Value2 != null)
                return ws.Cells[i, j].Value2;
            else
                return "";

        }

        public string WriteCell(string k, int i, int j)
        {
            i++;
            j++;

            ws.Cells[i, j].Value2 = k;
            ws.Cells[i, j].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red); 

            return k;

            
            
        }

        public void SaveFile()
        {
            wb.SaveAs2(@$"C:\Users\andre\Desktop\testando2.xlsx");
        }

        public int LastRow()
        {
            _Excel.Range last = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            _Excel.Range range = ws.get_Range("A1", last);

            int lastUsedRow = last.Row;
            
            return lastUsedRow;
        }

        public int LastColumn()
        {
            _Excel.Range last = ws.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            _Excel.Range range = ws.get_Range("A1", last);

            int lastUsedcolumn = last.Column;

            return lastUsedcolumn;
        }

        

    }
}