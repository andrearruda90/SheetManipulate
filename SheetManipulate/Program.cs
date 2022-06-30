using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

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
            

            void OpenFlie()
            {
                Excel excel = new Excel(@$"C:\Users\andre\Desktop\testando.xlsx", 1);
                
                Console.WriteLine(excel.ReadCell(0, 0));
                excel.WriteCell("C#", 1, 0);
                Console.WriteLine(excel.ReadCell(1, 0));

                excel.SaveFile();
            }

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
            return k;

            
            
        }

        public void SaveFile()
        {
            wb.SaveAs2(@$"C:\Users\andre\Desktop\testando2.xlsx");
        }

    }
}