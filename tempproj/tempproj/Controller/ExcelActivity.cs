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

        public Dictionary<string, Excel.Application> eXL = new Dictionary<string, Excel.Application>();
        public Dictionary<string, Excel.Workbook> eWB = new Dictionary<string, Excel.Workbook>();
        public Dictionary<string, Excel.Sheets> eWS = new Dictionary<string, Excel.Sheets>();
        public Excel.Worksheet centerWS = null;
        public Excel.Worksheet thezoneWS = null;
        //public Dictionary<string,object[,]> colvalbyAddr = new Dictionary<string, object[,]>();
        public Dictionary<string, object[,]> colvalbyName = new Dictionary<string, object[,]>();
        List<String> addrs = new List<String>();
        public JObject mapped_table = new JObject();
        public int minbound = 4, maxbound = 0, totalrow = 0;


        public String Work(string center, string thezone, string savepath, JObject json)
        {
            String ex = null;
            try
            {
                mapped_table = json;
                Open(center);
                Open(thezone);
                thezoneWS = eWS[thezone].Item[1];
                foreach (Excel.Worksheet ws in eWS[center])
                {
                    addrs = mapped_table.Properties().Select(p => p.Name).ToList();
                    centerWS = ws;
                    Read_Column();
                    Paste();
                    Brush();


                    minbound = maxbound;
                    colvalbyName.Clear();
                    //colvalbyAddr.Clear();
                }
                Checksum(thezone);
                Save(savepath);
                return ex;
            }
            catch (NullReferenceException e)
            {
                ex = e.Message.ToString();
                Close();
                return ex;
            }
            catch (InvalidCastException)
            {
                ex = "Error : Excel File에 필요한 내용만 있는지 확인해주세요";
                Close();
                return ex;
            }
            catch (COMException)
            {
                ex = "Error : Excel 프로세스가 열려있는지 확인해주세요";
                Close();
                return ex;
            }
            catch (Exception)
            {
                ex = "Error : 관리자에게 문의해주세요";
                Close();
                return ex;
            }
        }

        private void Open(string path)
        {
            eXL.Add(path, new Excel.Application());
            eWB.Add(path, eXL[path].Workbooks.Open(path));
            //eXL[path].Visible = true;
            eWS.Add(path, eWB[path].Worksheets);
        }

        private void Check_mapping_table(string center, string thezone) //mapping table
        {
            List<string> names = mapped_table.Properties().Select(p => p.Name).ToList(); //mapping table에 있는 Key값들을 List로 가져오기
            Excel.Range check = null;
            foreach (string colname in names)
            {
                check = centerWS.UsedRange.Find(colname);
                if (check == null)
                    throw new NullReferenceException($"{center}의 mapping table에서 center쪽 {colname} (이/가) 없음");
            }
            foreach (string colname in names)
            {
                string temp = mapped_table[colname].ToString();
                string temp2 = null;
                if (mapped_table[colname].GetType() == typeof(Newtonsoft.Json.Linq.JObject))
                {
                    JObject t = (JObject)mapped_table[colname];
                    temp = t["True"].ToString();
                    temp2 = t["False"].ToString();
                    if (temp2 != null)
                    {
                        check = centerWS.UsedRange.Find(temp2);
                        if (check == null)
                            throw new NullReferenceException($"{center}의 mapping table에서 thezone쪽 {colname} (이/가) 없음");
                    }
                }
                check = centerWS.UsedRange.Find(temp);
                if (check == null)
                    throw new NullReferenceException($"{center}의 mapping table에서 thezone쪽 {colname} (이/가) 없음");
            }
        }

        private void Read_Column()
        {
            Excel.Range usedrng = centerWS.UsedRange;
            Excel.Range temp;

            Excel.Range boundary = centerWS.UsedRange.Find("일용직");
            if (boundary != null)
            {
                maxbound += boundary.Row - 1;
                totalrow = boundary.Row - 1;
            }
            else
            {
                maxbound += usedrng.Rows.Count;
                totalrow = usedrng.Rows.Count;
            }


            foreach (String addr in addrs)
            {
                String name = centerWS.Range[addr].Value;
                Console.WriteLine(name);
                int row = centerWS.Range[addr].Row;
                int col = centerWS.Range[addr].Column;
                int offset = Find_Entry(addr);
                String colname = GetExcelColumnName(col);
                String start = colname + (row + offset).ToString();
                String end = colname + totalrow.ToString();
                Console.WriteLine(start + " " + end);
                temp = centerWS.Range[start + ":" + end];
                //Console.WriteLine(v.GetLength(0) + " " + v.GetLength(1)); 
                object[,] v = temp.Value;
                if (!colvalbyName.ContainsKey(name))
                {
                    colvalbyName.Add(name, v);
                    //colvalbyAddr.Add(addr, v);
                }
                else //합산하는 경우
                {
                    for (int i = 1; i <= temp.Value.GetLength(0); i++)
                    {
                        colvalbyName[name][i, 1] = (double)colvalbyName[name][i, 1] + (double)v[i, 1];
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

        private int Find_Entry(string addr)
        {
            Excel.Range mrng = centerWS.Range[addr];
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
                if (mapped_table[addr].ToString().Equals("사원코드"))
                {
                    row = mrng.Row + 1;
                    col = mrng.Column;
                    while (centerWS.Cells[row, col].Value == null)
                    {
                        row++;
                    }
                }
                else
                {
                    row = mrng.Row + 1;
                    col = mrng.Column;
                    while (!(centerWS.Cells[row, col].Value is double))
                    {
                        row++;
                    }
                }
                rows = row - mrng.Row;
            }
            return rows;
        }


        private void Paste()
        {
            Excel.Range colrng = thezoneWS.Range["A1", thezoneWS.Range["A1"].End[Excel.XlDirection.xlToRight]];
            foreach (var item in addrs)
            {
                Excel.Range findrng = colrng.Find(mapped_table[item].ToString());
                String itemkey = centerWS.Range[item].Value.ToString();
                if (findrng == null)
                {
                    throw new NullReferenceException($"thezone에 {mapped_table[item].ToString()}(이/가) 없음");
                }
                else
                {
                    if (!(mapped_table[item] is JObject)) //1:1 mapping
                    {
                        String col = GetExcelColumnName(findrng.Column);
                        thezoneWS.Range[col + minbound.ToString() + ":" + col + maxbound].Value = colvalbyName[itemkey];
                    }
                    else
                    {
                        int row = minbound;
                        foreach (var val in colvalbyName[itemkey])
                        {
                            String key = Get_Colname(itemkey, row, (JObject)mapped_table[item]);//Rule 3 적용 함수
                            if (key == null)
                            {
                                row++;
                                continue;
                            }
                            Excel.Range rng = colrng.Find(key);

                            if (rng[row.ToString()].Value != null)
                            {
                                if (val != null)
                                    rng[row.ToString()].Value += (double)val;
                                else
                                    rng[row.ToString()].Value += 0;
                            }
                            else
                            {
                                if (val != null)
                                {
                                    //Console.WriteLine("?? : " + val);
                                    rng[row.ToString()].Value = (double)val;

                                }
                                else
                                    rng[row.ToString()].Value += 0;
                            }
                            row++;

                        }
                    }
                }
            }
        }

        private String Get_Colname(string itemkey, int row, JObject json) //For Rule 3
        {
            String position, pos, key = null;
            int offset, dest;
            position = json["구분"].ToString();

            Excel.Range prng = centerWS.UsedRange.Find(position);
            offset = Find_Entry(prng.Address); //cell이 병합됬을 경우 시작점 계산
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
                    //Console.WriteLine(dest + " " + pos);
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

        private void Brush()
        {
            Excel.Range usedrng = thezoneWS.UsedRange.Rows.Offset[3];


            int rcnt = usedrng.Rows.Count;
            int ccnt = usedrng.Columns.Count;
            Stack<Excel.Range> deleted = new Stack<Excel.Range>();
            //int rcnt = usedrng.Rows.Count;
            foreach (Excel.Range item in usedrng)
            {
                var ssn = item.Cells[1, 1].Value;

                if (ssn != null)
                {
                    Regex regex = new Regex(@"^[A-Z]{3}");
                    if (!regex.IsMatch(ssn.ToString()))
                    {
                        deleted.Push(item);
                    }
                }
                else
                {
                    deleted.Push(item);
                }
            }
            foreach (Excel.Range row in deleted)
            {
                row.Delete();
            }


        }

        private void Save(string savepath)
        {
            if (System.IO.File.Exists(savepath))
                System.IO.File.Delete(savepath);
            thezoneWS.SaveAs(savepath);
        }
        private void Checksum(string thezone)
        {
            Excel.Range usedrng = thezoneWS.UsedRange;
            Excel.Range sum = null;
            Excel.Range valrng = thezoneWS.UsedRange.Offset[3, 1];
            for (int i = usedrng.Columns.Count; i >= 1; i--)
            {
                double num = eXL[thezone].WorksheetFunction.CountA(valrng.Columns[i]);
                if (num == 0)
                {
                    usedrng.Columns[i + 1].Delete();
                    //deleted.Enqueue(usedrng.Columns[i+1]);
                }
            }
            Console.WriteLine(usedrng.Rows.Count + " " + usedrng.Columns.Count + " " + maxbound);
            int cnt = usedrng.Columns.Count;

            sum = thezoneWS.Range["B" + maxbound];
            Excel.Range categorysum = sum.Resize[Type.Missing, cnt - 1];
            categorysum.Formula = "=ROUND(SUM(B" + 4.ToString() + ":B" + (maxbound - 1).ToString() + "), 0)"; //항목 별 합계


            String forsum = GetExcelColumnName(cnt + 1);
            String forend = GetExcelColumnName(cnt);
            sum = thezoneWS.Range[forsum + "4"];
            Excel.Range employeesum = sum.Resize[maxbound - 4, Type.Missing];//offset 때문에 4를 빼줌
            employeesum.Formula = "=ROUND(SUM(B4:" + forend + "4), 0)"; //직원 당 합계

            Excel.Range totalsum = thezoneWS.Range[forsum + maxbound.ToString()];
            string formula = "=ROUND(SUM(" + forsum + "4:" + forsum + (maxbound - 1).ToString() + ", B" + maxbound.ToString() + ":" + forend + maxbound.ToString() + "), 2)"; //최종 합계
            //Console.WriteLine(formula);
            totalsum.Formula = formula;
            totalsum.NumberFormat = 0;

            valrng.NumberFormat = "0";

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
                foreach (KeyValuePair<string, Excel.Sheets> item in eWS)
                    ReleaseExcelObject(item.Value);
                foreach (KeyValuePair<string, Excel.Workbook> item in eWB)
                    ReleaseExcelObject(item.Value);
                foreach (KeyValuePair<string, Excel.Application> item in eXL)
                    ReleaseExcelObject(item.Value);
                eWS.Clear();
                eWB.Clear();
                eXL.Clear();
                colvalbyName.Clear();
                //colvalbyAddr.Clear();
                centerWS = null;
                thezoneWS = null;
                minbound = 4; maxbound = 0;
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
