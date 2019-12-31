using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace tempproj
{
    public class ExcelActivity
    {
        public SortedList<string, Excel.Application> eXL = new SortedList<string, Excel.Application>();
        public SortedList<string, Excel.Workbook> eWB = new SortedList<string, Excel.Workbook>();
        public SortedList<string, Excel.Worksheet> eWS = new SortedList<string, Excel.Worksheet>();
        public SortedList<string, object[,]> colvalues = new SortedList<string, object[,]>();
        public Excel.Range eRng, ID;
        public object[,] ID_values;
        public List<String> columnNames = new List<String>();

        public Exception Work(string path01, string path02, string path03)
        {
            try
            {
                Open(path01); Open(path02);
                Read_Column(path01);
                Copy_Paste(path02);
                Save(path02, path03);
                Close();
                return null;

            }
            catch (Exception e)
            {
                Close();
                return e;
            }
        }

        public void Open(string path, string sheetName = null)
        {
            eXL.Add(path, new Excel.Application());
            eWB.Add(path, eXL[path].Workbooks.Open(path));
            //eXL[path].Visible = true;
            if (sheetName == null)  //default : 현재 workbook의 첫번째 worksheet를 open
                eWS.Add(path, eWB[path].Worksheets.get_Item(1) as Excel.Worksheet);
            else                    //workbook 내에 여러 시트중 원하는 시트가 있으면 해당 시트 open
                eWS.Add(path, eWB[path].Worksheets.Item[sheetName]);
        }

        public void Read_Column(string exl)
        {

            ID = eWS[exl].Range["E6", eWS[exl].Range["E6"].End[Excel.XlDirection.xlDown]];
            ID_values = ID.Value;
            int IDcnt = ID.Count;
            IDcnt += 5;
            char curcol = 'I';

            while (true)
            {
                String colname = eWS[exl].Range[curcol + "5"].Value;
                if (colname == "소계")
                    break;
                else
                {
                    columnNames.Add(colname);
                    colvalues.Add(colname, eWS[exl].Range[curcol + "6: " + curcol + IDcnt.ToString()].Value);
                    curcol = (char)((int)curcol + 1);
                }
            }
        }

        public void Copy_Paste(string exl)
        {
            Excel.Range colrng = eWS[exl].Range["A1", eWS[exl].Range["A1"].End[Excel.XlDirection.xlToRight]];
            int colcnt = colrng.Count;
            foreach (Excel.Range item in colrng)
            {
                if (item.Text == "사원코드")
                {
                    int col = 4;
                    Console.WriteLine("Found");
                    foreach (var val in ID_values)
                    {
                        //String r = col.ToString();
                        item[col.ToString()].Value = val;
                        col++;
                    }
                    break;
                }
            }
            foreach (String colname in columnNames)
            {
                foreach (Excel.Range item in colrng)
                {
                    if (item.Text == colname)
                    {
                        int col = 4;
                        Console.WriteLine(colname + " Found");
                        foreach (var val in colvalues[colname])
                        {
                            item[col.ToString()].Value = val;
                            col++;
                        }
                        break;
                    }
                }
            }
        }

        public void Save(string exl, string path)
        {
            eWB[exl].SaveAs(path);
        }

        public void Close() //열려있는 모든 excel 객체 해제
        {
            foreach (KeyValuePair<string, Excel.Workbook> item in eWB)
            {
                item.Value.Close();
            }
            foreach (KeyValuePair<string, Excel.Worksheet> item in eWS)
            {
                ReleaseExcelObject(item.Value);
            }
            foreach (KeyValuePair<string, Excel.Workbook> item in eWB)
            {
                ReleaseExcelObject(item.Value);
            }
            foreach (KeyValuePair<string, Excel.Application> item in eXL)
            {
                ReleaseExcelObject(item.Value);
            }
        }
        private void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
