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
        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
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
        public Dictionary<string, point> colNames = new Dictionary<string, point>(); //key : column name, key : coordinate of column on excel
        public Dictionary<string, List<Excel.Range>> colAddr = new Dictionary<string, List<Excel.Range>>();
        public Dictionary<string, List<object[,]>> colvalues = new Dictionary<string, List<object[,]>>(); //key : column name , value : value of each column
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
                Copy_Paste(path01, path02);
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



            foreach (KeyValuePair<string, List<Excel.Range>> item in colAddr)
            {
                foreach (Excel.Range addr in item.Value)
                {
                    int offset = Find_Entry(exl, addr.Address); //cell이 병합됬을 경우 시작점 계산

                    string col = GetExcelColumnName(addr.Column);
                    int rstart = addr.Row + offset;
                    int rend = rstart + rcnt - 1;
                    Console.WriteLine(item.Key + " " + rstart + " " + rend);
                    Console.WriteLine(col + rstart.ToString() + ":" + col + rend.ToString());
                    temp = eWS[exl].Range[col + rstart.ToString() + ":" + col + rend.ToString()];
                    if (colvalues.ContainsKey(item.Key))
                    {
                        colvalues[item.Key].Add(temp.Value);
                    }
                    else
                    {
                        List<object[,]> rtemp = new List<object[,]>();
                        rtemp.Add(temp.Value);
                        colvalues.Add(item.Key, rtemp);
                    }
                }
            }
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

        private int Find_Entry(string exl, string addr)
        {
            Excel.Range mrng = eWS[exl].Range[addr];
            bool rg = mrng.MergeCells;
            int rows = 1, row, col;
            if (rg)
            {
                dynamic mvalue = mrng.MergeArea.Value;
                object[,] vals = mvalue as object[,];
                if (vals != null)
                {
                    rows = vals.GetLength(0);
                    //int cols = vals.GetLength(1);
                    //Console.WriteLine(rows + " " + cols);
                }
            }
            else
            {
                if (mapped_table[mrng.Value].ToString().Equals("사원코드"))
                {
                    row = mrng.Row + 1;
                    col = mrng.Column;
                    while (eWS[exl].Cells[row, col].Value == null)
                    {
                        row++;
                    }
                }
                else
                {
                    row = mrng.Row + 1;
                    col = mrng.Column;
                    while (!(eWS[exl].Cells[row, col].Value is double))
                    {
                        row++;
                    }
                }
                rows = row - mrng.Row;
            }
            return rows;
            //eWS[exl].Range[addr].Value = "hi";
        }

        private void Find_Columns(string exl)
        {
            Excel.Range currentFind = null;
            Excel.Range usedrng = eWS[exl].UsedRange;
            int colcnt = usedrng.Columns.Count;
            List<string> names = mapped_table.Properties().Select(p => p.Name).ToList(); //mapping table에 있는 Key값들을 List로 가져오기

            foreach (string name in names)
            {
                List<Excel.Range> addrs = new List<Excel.Range>();
                Excel.Range rng = usedrng.Find(name);

                if (rng != null)
                {
                    //colAddr.Add(name, rng);
                    addrs.Add(rng);
                    currentFind = usedrng.Find(name, rng, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                        Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);
                    if (currentFind != null && rng.Address != currentFind.Address && rng.Column != currentFind.Column)
                    {
                        Console.WriteLine(name);
                        Console.WriteLine(rng.Address);
                        Console.WriteLine(currentFind.Address);
                        addrs.Add(currentFind);
                    }
                    colAddr.Add(name, addrs);
                }
            }
            foreach (KeyValuePair<string, List<Excel.Range>> item in colAddr)
            {
                Console.WriteLine(item.Key);
                foreach (Excel.Range r in item.Value)
                {
                    Console.WriteLine(r.Address);
                }
            }
        }

        private void Copy_Paste(string exl, string exl2)
        {
            Excel.Range colrng = eWS[exl2].Range["A1", eWS[exl2].Range["A1"].End[Excel.XlDirection.xlToRight]];
            int colcnt = colrng.Count;
            String key = String.Empty;
            Excel.Range rng = null;
            foreach (KeyValuePair<string, List<object[,]>> item in colvalues)
            {
                foreach (var values in item.Value) //item.Value => List<object[,]>
                {
                    int row = 4;
                    Console.WriteLine(item.Key + " is Found");
                    foreach (var val in values)
                    {
                        if (!mapped_table[item.Key].ToString().Equals("사원코드"))
                        {
                            if (!(mapped_table[item.Key] is JObject)) //1:1 mapping
                            {
                                rng = colrng.Find(mapped_table[item.Key].ToString());
                            }
                            else //Rule 3 적용
                            {
                                key = Get_Colname(exl, item.Key, row, (JObject)mapped_table[item.Key]);
                                if (key == null)
                                {
                                    row++;
                                    continue;
                                }

                                rng = colrng.Find(key);
                            }
                            if (rng[row.ToString()].Value != null)
                            {
                                if (val != null)
                                    rng[row.ToString()].Value += Math.Truncate((double)val);
                                else
                                    rng[row.ToString()].Value += 0;
                            }
                            else
                            {
                                if (val != null)
                                    rng[row.ToString()].Value = Math.Truncate((double)val);
                                else
                                    rng[row.ToString()].Value += 0;
                            }
                        }
                        else
                        {
                            rng = colrng.Find(mapped_table[item.Key].ToString());
                            rng[row.ToString()].Value = val;
                        }
                        row++;
                    }

                }
            }
        }

        private String Get_Colname(string exl, string itemkey, int row, JObject json)
        {
            String position, pos, key = null;
            int offset, dest;
            position = json["구분"].ToString();

            Excel.Range prng = eWS[exl].UsedRange.Find(position);
            offset = Find_Entry(exl, prng.Address); //cell이 병합됬을 경우 시작점 계산
            dest = prng.Row + offset + row - 4;
            pos = prng.Cells[dest].Value;

            if (itemkey.Contains("/"))
            {
                if (pos != null)
                {
                    Console.WriteLine(dest + " " + pos);
                    JArray positions = (JArray)json["값"];

                    foreach (var item in positions)
                    {
                        if (pos.Equals(item.ToString()))
                        {
                            key = json["True"].ToString();
                            Console.WriteLine(key);
                            break;
                        }
                    }
                    if (key == null)
                    {
                        key = json["False"].ToString();
                    }
                }
            }
            else
            {
                if (pos != null)
                {
                    Console.WriteLine(dest + " " + pos);
                    JArray positions = (JArray)json["값"];

                    foreach (var item in positions)
                    {
                        if (pos.Equals(item.ToString()))
                        {
                            key = json["True"].ToString();
                            Console.WriteLine(key);
                            break;
                        }
                    }
                }
            }

            return key;
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
            try
            {
                uint processId = 0;
                foreach (KeyValuePair<string, Excel.Workbook> item in eWB)
                    item.Value.Close(0);
                foreach (KeyValuePair<string, Excel.Application> item in eXL)
                {
                    GetWindowThreadProcessId(new IntPtr(item.Value.Hwnd), out processId);
                    item.Value.Quit();
                    if (processId != 0)
                    {
                        System.Diagnostics.Process excelProcess = System.Diagnostics.Process.GetProcessById((int)processId);
                        excelProcess.CloseMainWindow();
                        excelProcess.Refresh();
                        excelProcess.Kill();
                    }
                }
                foreach (KeyValuePair<string, Excel.Worksheet> item in eWS)
                    ReleaseExcelObject(item.Value);
                foreach (KeyValuePair<string, Excel.Workbook> item in eWB)
                    ReleaseExcelObject(item.Value);
                foreach (KeyValuePair<string, Excel.Application> item in eXL)
                    ReleaseExcelObject(item.Value);
                eWS.Clear();
                eWB.Clear();
                eXL.Clear();
                colNames.Clear();
                colAddr.Clear();
                colvalues.Clear();
            }
            catch (Exception) { }
        }
        private void ReleaseExcelObject(object obj)
        {
            Marshal.ReleaseComObject(obj);
            GC.Collect();
        }
    }
}