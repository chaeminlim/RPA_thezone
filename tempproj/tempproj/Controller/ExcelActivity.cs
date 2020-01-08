using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;

namespace tempproj
{
    public class ExcelActivity
    {
        public struct point
        {
            public point(int r, int c)
            {
                row = r;
                column = c;
            }
            public int row;
            public int column;
        };
        public Dictionary<string, Excel.Application> eXL = new Dictionary<string, Excel.Application>();
        public Dictionary<string, Excel.Workbook> eWB = new Dictionary<string, Excel.Workbook>();
        public Dictionary<string, Excel.Worksheet> eWS = new Dictionary<string, Excel.Worksheet>();
        public Dictionary<string, object[,]> colvalues = new Dictionary<string, object[,]>(); //key : column name , value : value of each column
        public Dictionary<string, point> colNames = new Dictionary<string, point>(); //key : column name, key : coordinate of column on excel
        public Dictionary<string, Excel.Range> colAddr = new Dictionary<string, Excel.Range>();
        public JObject mapped_table = new JObject();
        public Excel.Range eRng, ID;
        public object[,] ID_values;


        public Exception Work(string path01, string path02, string savepath, JObject json)
        {
            try
            {
                mapped_table = json;
                Open(path01); Open(path02);
                Find_Columns(path01);
                Read_Column(path01);
                Copy_Paste(path02);
                Brush(path02);
                //FIrst_Column(path01);
                Save(path02, savepath);
                Close();
                return null;
            }
            catch (Exception e)
            {
                Close();
                return e;
            }



        }

        private void Open(string path, string sheetName = null)
        {
            eXL.Add(path, new Excel.Application());
            eWB.Add(path, eXL[path].Workbooks.Open(path));
            //eXL[path].Visible = true;
            if (sheetName == null)  //default : 현재 workbook의 첫번째 worksheet를 open
                eWS.Add(path, eWB[path].Worksheets.get_Item(1) as Excel.Worksheet);
            else                    //workbook 내에 여러 시트중 원하는 시트가 있으면 해당 시트 open
                eWS.Add(path, eWB[path].Worksheets.Item[sheetName]);
        }
        private object[,] Read_Range(string exl, string start, string end)
        {
            //eWS[exl] = eWB[exl].Worksheets.Item[sheetName];
            eRng = eWS[exl].get_Range(start, end);
            return eRng.Value;
        }



        private void Read_Column(string exl)
        {
            Excel.Range usedrng = eWS[exl].UsedRange;
            int rcnt = usedrng.Rows.Count;
            int ccnt = usedrng.Columns.Count;
            int IDcnt = 0; //인원수
            Excel.Range temp, ID;



            foreach (KeyValuePair<string, Excel.Range> item in colAddr)
            {
                int offset = Find_Entry(exl, item.Value.Row, item.Value.Column); //cell이 병합됬을 경우 시작점 계산
                string col = GetExcelColumnName(item.Value.Column);
                int rstart = item.Value.Row + offset;
                int rend = rstart + rcnt - 1;
                Console.WriteLine(item.Key + " " + rstart + " " + rend);
                Console.WriteLine(col + rstart.ToString() + ":" + col + rend.ToString());
                temp = eWS[exl].Range[col + rstart.ToString() + ":" + col + rend.ToString()];
                colvalues.Add(item.Key, temp.Value);
            }

            /*
            foreach (KeyValuePair<string, Excel.Range> item in colAddr)
            {
                
                if (mapped_table[item.Key].ToString().Equals("사원코드"))
                {
                    int offset = Find_Entry(exl, item.Value.Row, item.Value.Column);
                    temp = eWS[exl].Cells[item.Value.Row+offset, item.Value.Column];
                    ID = eWS[exl].Range[temp, temp.End[Excel.XlDirection.xlDown]];
                    
                    colvalues.Add(item.Key, ID.Value);
                    IDcnt = ID.Count;
                    Console.WriteLine(IDcnt);
                    break;
                }
            }
            foreach (KeyValuePair<string, Excel.Range> item in colAddr)
            {
                if (!mapped_table[item.Key].ToString().Equals("사원코드"))
                {
                    int offset = Find_Entry(exl, item.Value.Row, item.Value.Column); //cell이 병합됬을 경우 시작점 계산
                    string col = GetExcelColumnName(item.Value.Column);
                    int rstart = item.Value.Row + offset;
                    int rend = rstart + IDcnt - 1;
                    Console.WriteLine(item.Key + " " + rstart + " " + rend);
                    Console.WriteLine(col + rstart.ToString() + ":" + col + rend.ToString());
                    temp = eWS[exl].Range[col + rstart.ToString() + ":" + col + rend.ToString()];
                    colvalues.Add(item.Key, temp.Value);
                }
            }*/
        }

        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        private int Find_Entry(string exl, int r, int c)
        {
            Excel.Range mrng = eWS[exl].Cells[r, c];
            bool rg = mrng.MergeCells;
            int rows = 1;
            if (rg)
            {
                if (eWS[exl].Cells[r, c].MergeCells != null)
                {
                    dynamic mvalue = mrng.MergeArea.Value2;
                    object[,] vals = mvalue as object[,];
                    if (vals != null)
                    {
                        rows = vals.GetLength(0);
                        //int cols = vals.GetLength(1);
                        //Console.WriteLine(rows + " " + cols);
                    }
                }
            }
            return rows;
            //eWS[exl].Range[addr].Value = "hi";
        }

        private void Find_Columns(string exl)
        {

            Excel.Range usedrng = eWS[exl].UsedRange;
            List<string> names = mapped_table.Properties().Select(p => p.Name).ToList(); //mapping table에 있는 Key값들을 List로 가져오기
            foreach (string name in names)
            {
                Excel.Range rng = usedrng.Find(name);
                if (rng != null)
                    colAddr.Add(name, rng);
            }

            foreach (KeyValuePair<string, Excel.Range> item in colAddr)
            {
                Console.WriteLine(item.Key + " " + item.Value.Row + " " + item.Value.Column);
            }
        }


        private void Copy_Paste(string exl)
        {
            Excel.Range colrng = eWS[exl].Range["A1", eWS[exl].Range["A1"].End[Excel.XlDirection.xlToRight]];
            int colcnt = colrng.Count;
            foreach (KeyValuePair<string, object[,]> item in colvalues)
            {
                foreach (Excel.Range rng in colrng)
                {
                    if (mapped_table[item.Key].ToString().Equals(rng.Value))
                    {
                        int col = 4;
                        Console.WriteLine(item.Key + " is Found");
                        foreach (var val in item.Value)
                        {
                            rng[col.ToString()].Value = val;
                            col++;
                        }
                        break;
                    }
                }
            }
        }

        private void Brush(string exl)
        {
            Excel.Range usedrng = eWS[exl].UsedRange.Rows.Offset[3];
            Stack<Excel.Range> deleted = new Stack<Excel.Range>();
            //int rcnt = usedrng.Rows.Count;
            foreach (Excel.Range item in usedrng)
            {
                string ssn = item.Cells[1, 1].Value;

                if (ssn != null)
                {
                    Regex regex = new Regex(@"^[A-Z]{3}");
                    if (!regex.IsMatch(ssn))
                    {
                        //Console.WriteLine("MissMatch   " + item.Address );
                        //item.Delete();
                        deleted.Push(item);

                    }
                }
                else
                {
                    //Console.WriteLine("null  " + item.Address);
                    //item.Delete();
                    deleted.Push(item);
                }
            }
            foreach (Excel.Range row in deleted)
            {
                row.Delete();
            }

        }

        private void Save(string exl, string savepath)
        {
            eWS[exl].SaveAs(savepath);
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
