using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SearchExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string risposta = string.Empty;
            do
            {
                //Grant Carroll //CLOUDWICK TECHNOLOGIES INC.

                //C:\Users\NIRAV\Desktop\reg_exports\H-1B_Disclosure_Data_FY16.xlsx
                Console.Write("Please enter excel file path: ");
                string path = Console.ReadLine();

                Console.Write("\nPlease enter any string to search: ");
                string excelsearch = Console.ReadLine();
                
                Excelsearching.SearchText(excelsearch, path);                        

                Console.WriteLine("\nDo you want to continue? (Y/N)");
                risposta = Console.ReadLine();

                if (risposta.Equals("Y"))
                {
                    continue;
                }
                else if (risposta.Equals("N"))
                {
                    break;
                }
                else {
                    Console.WriteLine("\nPlease enter valid input (Y/N)!!!");
                }
                Console.ReadLine();
            } while (risposta == "Y");            

        }

      
    }

    //for the excel searhing static class and their extentions 
    public static class Excelsearching {
        /// <summary>
        /// To search text from the excel sheet
        /// </summary>
        public static void SearchText(string strsearch, string path)
        {
            string File_name = @path;
            StringBuilder sbresult = new StringBuilder();
            StringBuilder sbheader = new StringBuilder();
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook oWB;
            Microsoft.Office.Interop.Excel.Worksheet oSheet;
            try
            {
                object missing = System.Reflection.Missing.Value;
                oWB = oXL.Workbooks.Open(File_name, missing, missing, missing, missing,
                    missing, missing, missing, missing, missing, missing,
                    missing, missing, missing, missing);
                oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oWB.Worksheets[1];
                List<Microsoft.Office.Interop.Excel.Range> oRng = GetSpecifiedRange(strsearch, oSheet);
               

                int row = oSheet.Rows.CurrentRegion.EntireRow.Count;
                int col = oSheet.Columns.CurrentRegion.EntireColumn.Count;
                //for the header file details
                for (int i = 1; i < col; i++)
                {
                    sbheader.Append((oSheet.Cells[1, i] as Microsoft.Office.Interop.Excel.Range).Value + "\t");
                }

                Console.Write("Total Rows: " + row + " Total Columns: " + col);
                Console.WriteLine("\nColumn Names:\n");
                Console.WriteLine(sbheader.ToString() + "\n");
                Console.WriteLine("Filtered Row:\n");
                if (oRng != null)
                {
                    foreach (var raange in oRng)
                    {
                        if (raange != null)
                        {
                            Console.Write("\nText found, position is Row:" + raange.Row + " and column:" + raange.Column + " \n");

                            //for the row details
                            for (int i = 1; i < col; i++)
                            {
                                sbresult.Append((oSheet.Cells[raange.Row, i] as Microsoft.Office.Interop.Excel.Range).Value + "\t");
                            }

                            Console.WriteLine(sbresult.ToString() + "\n");
                            sbresult = new StringBuilder();
                        }
                        else
                        {
                            Console.Write("\nText is not found");
                        }
                    }

                    Console.WriteLine("\nTotal Matching Row Count: " + oRng.Count());
                }
                else
                {
                    Console.Write("\nText is not found");
                }
                oWB.Close(false, missing, missing);

                oSheet = null;
                oWB = null;
                oXL.Quit();
            }
            catch (Exception ex)
            {
                Console.Write("\nException occured: " + ex.Message);
            }
        }

        /// <summary>
        /// to get the specific value for the search range
        /// </summary>
        /// <param name="matchStr"></param>
        /// <param name="objWs"></param>
        /// <returns></returns>
        public static List<Microsoft.Office.Interop.Excel.Range> GetSpecifiedRange(string matchStr, Microsoft.Office.Interop.Excel.Worksheet objWs)
        {
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Range allFind = null;
            List<Microsoft.Office.Interop.Excel.Range> lifirstFind = new List<Microsoft.Office.Interop.Excel.Range>();
            Microsoft.Office.Interop.Excel.Range last = objWs.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
           
            //row and column count
            int row = objWs.Rows.CurrentRegion.EntireRow.Count;
            int col = objWs.Columns.CurrentRegion.EntireColumn.Count;
            
            //for loop for the take the list of matching rows range
            for(int i=0;i<row;i++){
                //for the 1st loop
                if (i == 0)
                {
                    allFind = objWs.get_Range("A1", last).Find(matchStr, missing,
                              Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues,
                              Microsoft.Office.Interop.Excel.XlLookAt.xlPart,
                              Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
                              Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, false, missing, missing);
                }
                else {
                    allFind = objWs.get_Range(allFind, last).Find(matchStr, missing,
                                 Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues,
                                 Microsoft.Office.Interop.Excel.XlLookAt.xlPart,
                                 Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
                                 Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, false, missing, missing);
                }

                //to check whether allfind is null or not
                if (allFind != null)
                {
                    lifirstFind.Add(allFind);
                    allFind = objWs.Cells[allFind.Row + 1, 1];//increment by 1
                }
                else {
                    break;
                }
            }
           
            return lifirstFind;
        }
    
    
    }
}
