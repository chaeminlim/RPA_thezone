﻿using System;
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
        public Excel.Worksheet centerWS = null;
        public Excel.Worksheet thezoneWS = null;
        public Dictionary<string, object[,]> colvalbyCenterName = new Dictionary<string, object[,]>();
        public Dictionary<string, object[,]> colvalbyThezoneName = new Dictionary<string, object[,]>();
        List<String> addrs = new List<String>();
        public List<JObject> mapping_table = new List<JObject>();
        /*  maxbound : 더존 파일의 끝 행
            minbound : 더존 파일에 붙여넣을 때의 시작 행
            totalrow : center worksheet의 끝 행
            columncnt : 필요없는 column을 지운 후의 column 개수   
        */
        public int minbound = 4, maxbound = 0, totalrow = 0, columncnt = 0;
        public List<int> SumFlag = new List<int>();//DeleteNullCol에서 사용할 합계가 적혀있는 Row정보


        public String Work(string center, string thezone, string savepath, List<JObject> json)
        {
            String ex = null;
            try
            {
                List<JObject> center_table = json;
                int table = 0;
                //mapping_table = json;
                List<JObject> center_tables = json;

                Open(center);
                Open(thezone);
                thezoneWS = eWB[thezone].Worksheets.Item[1];
                foreach (Excel.Worksheet ws in eWB[center].Worksheets)
                {

                    if (ws.Visible == Excel.XlSheetVisibility.xlSheetHidden)
                    {
                        continue;
                    }
                    mapping_table = center_tables[table]["배치표"].ToObject<List<JObject>>();
                    centerWS = ws;

                    Extract(center);

                    Paste();

                    DeleteWithSsn(thezone);

                    Checksum(thezone);

                    DeleteNullRow(thezone);
                   
                    minbound = maxbound + 5;
                    colvalbyCenterName.Clear();
                    colvalbyThezoneName.Clear();
                    table++;
                }
                DeleteNullCol(thezone);
                round();

                Save(savepath);


                Close();
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
                ex = "mapping table 내용을 다시 한번 체크해주세요.";
                Close();
                return ex;
            }
            catch (COMException)
            {
                ex = "Error : Excel 프로세스가 열려있는지 확인해주세요";
                Close();
                return ex;
            }
            catch (ArgumentOutOfRangeException)
            {
                ex = "mapping table과 워크시트 사이의 개수가 안맞습니다.";
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

        #region Utility
        private int GetlastRow(string exl, string company, int lastrow)
        {
            Excel.Application eTemp = new Excel.Application();
            if (company.Equals("center"))
            {
                while (eXL[exl].WorksheetFunction.CountA(centerWS.Rows[lastrow]) == 0)
                {
                    lastrow--;
                }
            }
            else
            {
                while (eXL[exl].WorksheetFunction.CountA(thezoneWS.Rows[lastrow]) == 0)
                {
                    lastrow--;
                }
            }
            return lastrow;
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
                }
            }
            else
            {
                if (mapping_table[0]["셀위치"].ToString().Equals(addr))//addr이 사번을 나타내는 좌표인지 확인
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
        private String Get_Colname(string itemkey, int row, JObject json) //For Rule 3
        {
            String position, pos, key = null;
            int offset, dest;
            position = json["구분"].ToString();

            Excel.Range prng = centerWS.Range[position];
            offset = Find_Entry(prng.Address); //cell이 병합됬을 경우 시작점 계산
            dest = prng.Row + offset + (row - 4);
            pos = centerWS.Cells[dest, prng.Column].Value;
            if (itemkey.Contains("/"))
            {
                if (pos != null)
                {
                    JArray positions = (JArray)json["값"];

                    foreach (var item in positions)
                    {
                        if (pos.Equals(item.ToString()))
                        {
                            key = json["True"].ToString();

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
                    JArray positions = (JArray)json["값"];

                    foreach (var item in positions)
                    {
                        if (pos.Equals(item.ToString()))
                        {
                            key = json["True"].ToString();
                            break;
                        }
                    }
                }
            }

            return key;
        }
        #endregion


        #region Excel_Functions
        private void Open(string path)
        {
            eXL.Add(path, new Excel.Application());

            eWB.Add(path, eXL[path].Workbooks.Open(path));
        }
        private void Extract(string center)
        {
            Excel.Range usedrng = centerWS.UsedRange;
            Excel.Range temp;
            List<string> filter = new List<string>();
            filter.Add("일용직"); filter.Add("교육비"); filter.Add("사업소득");
            Excel.Range boundary = null;
            int minboundary = 10000;
            foreach (var item in filter)
            {
                boundary = centerWS.UsedRange.Find(item);
                if (boundary != null)
                {
                    if (minboundary > boundary.Row)
                        minboundary = boundary.Row;
                }
            }

            if (minboundary != 10000)
            {
                int lastrow = GetlastRow(center, "center", minboundary - 1);
                maxbound += lastrow;
                totalrow = lastrow;
            }
            else
            {
                int lastrow = GetlastRow(center, "center", usedrng.Rows.Count);
                maxbound += lastrow;
                totalrow = lastrow;
            }
            //항상 mapping table의 첫번째는 사번임을 가정
            String s = (centerWS.Range[mapping_table[0]["셀위치"]].Row + Find_Entry(mapping_table[0]["셀위치"].ToString())).ToString();

            foreach (JObject category in mapping_table)
            {
                String name = centerWS.Range[category["셀위치"]].Value;
                if (name == null)
                {
                    throw new NullReferenceException($"{center}의 mapping table을 확인해주세요");
                }
                int col = centerWS.Range[category["셀위치"]].Column;
                String colname = GetExcelColumnName(col);
                String start = colname + s;
                String end = colname + totalrow.ToString();
                temp = centerWS.Range[start + ":" + end];
                object[,] v = temp.Value;

                if (!colvalbyCenterName.ContainsKey(name))
                {
                    colvalbyCenterName.Add(name, v);
                }
                foreach (var thezonecol in category["더존이름"]) //여러 곳에 mapping 될 경우
                {
                    if (!colvalbyThezoneName.ContainsKey(thezonecol.ToString()))
                    {
                        if (!(thezonecol is JObject))
                        {

                            object[,] thezonevalues;
                            thezonevalues = (object[,])v.Clone();
                            colvalbyThezoneName.Add(thezonecol.ToString(), thezonevalues);//thezone은 1:1mapping되는 column만 저장
                                                                                          //Console.WriteLine(mapping_table[addr].ToString() + " 노합산");

                        }
                    }
                    else //합산하는 경우
                    {

                        if (!(thezonecol is JObject))
                        {
                            for (int i = 1; i <= temp.Value.GetLength(0); i++)
                            {
                                if (colvalbyThezoneName[thezonecol.ToString()][i, 1] != null && v[i, 1] != null)
                                {
                                    double thezonevalue = Convert.ToDouble(colvalbyThezoneName[thezonecol.ToString()][i, 1]);
                                    double centervalue = Convert.ToDouble(v[i, 1]);
                                    colvalbyThezoneName[thezonecol.ToString()][i, 1] = thezonevalue + centervalue;
                                }
                                else if (colvalbyThezoneName[thezonecol.ToString()][i, 1] == null && v[i, 1] != null)
                                {
                                    colvalbyThezoneName[thezonecol.ToString()][i, 1] = v[i, 1];
                                }
                            }
                        }
                    }
                }

            }
        }

       
        private void Paste()
        {
            Excel.Range colrng = thezoneWS.Range["A1", thezoneWS.Range["A1"].End[Excel.XlDirection.xlToRight]];
            foreach (JObject category in mapping_table)
            {
                foreach (var thezonecol in category["더존이름"])
                {
                    if (!(thezonecol is JObject)) //1:1 mapping
                    {
                        Excel.Range findrng = colrng.Find(thezonecol.ToString());

                        if (findrng == null)
                        {
                            throw new NullReferenceException($"thezone에 {thezonecol.ToString()}(이/가) 없음");
                        }
                        String col = GetExcelColumnName(findrng.Column);
                        thezoneWS.Range[col + minbound.ToString() + ":" + col + maxbound].Value = colvalbyThezoneName[thezonecol.ToString()];

                    }
                    else
                    {
                        int row = minbound;
                        String itemkey = centerWS.Range[category["셀위치"]].Value.ToString();
                        foreach (var val in colvalbyCenterName[itemkey])
                        {
                            String key = Get_Colname(itemkey, row, (JObject)thezonecol);//Rule 3 적용 함수
                            if (key == null)
                            {
                                row++;
                                continue;
                            }
                            Excel.Range rng = colrng.Find(key);
                            if (rng == null)
                            {
                                throw new NullReferenceException($"thezone에 {key.ToString()}(이/가) 없음");
                            }
                            if (rng[row.ToString()].Value != null)
                            {
                                if (val != null)
                                    rng[row.ToString()].Value += Convert.ToDouble(val);
                                else
                                    rng[row.ToString()].Value += 0;
                            }
                            else
                            {
                                if (val != null)
                                {
                                    rng[row.ToString()].Value = Convert.ToDouble(val);

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

        private void DeleteWithSsn(string thezone)
        {
            Excel.Range usedrng = thezoneWS.UsedRange.Rows.Offset[minbound - 1];
            int deletedcnt = 0;

            int rcnt = usedrng.Rows.Count;
            int ccnt = usedrng.Columns.Count;

            Stack<Excel.Range> deleted = new Stack<Excel.Range>();
            foreach (Excel.Range item in usedrng)
            {
                var ssn = item.Cells[1, 1].Value;

                if (ssn != null)
                {
                    Regex newschool = new Regex(@"^[A-Z]{3}");
                    Regex oldschool = new Regex(@"^\d{5}");
                    if (!newschool.IsMatch(ssn.ToString()) && !oldschool.IsMatch(ssn.ToString()))
                    {
                        deleted.Push(item);
                        deletedcnt++;
                    }
                }
                else
                {
                    deleted.Push(item);
                    deletedcnt++;
                }
            }
            foreach (Excel.Range row in deleted)
            {
                row.Delete();
            }
            
        }



        private void Checksum(string thezone)
        {
            Excel.Range usedrng = thezoneWS.UsedRange;
            Excel.Range sum = null;
            int colcnt = 84;


            ////////직원별 합계
            String forsum = GetExcelColumnName(colcnt + 1);
            String fortotalsum = GetExcelColumnName(colcnt + 2);
            String forend = GetExcelColumnName(colcnt);

            sum = thezoneWS.Range[forsum + minbound.ToString()];
            Excel.Range employeesum = sum.Resize[totalrow - 5, Type.Missing];//offset 때문에 4를 빼줌 + 개수만큼 늘려서 1을 더 빼줌
            employeesum.Formula = "=SUM(B" + minbound.ToString() + ":" + "BH" + minbound.ToString() + ")"; //직원 당 합계
            employeesum.Copy();
            employeesum.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone);
            employeesum.NumberFormat = "#,##0";



            ////////category별 합계
            sum = thezoneWS.Range["B" + maxbound];
            Excel.Range categorysum = sum.Resize[Type.Missing, colcnt];
            categorysum.Formula = "=SUM(B" + minbound.ToString() + ":B" + (maxbound - 1).ToString() + ")"; //항목 별 합계
            categorysum.Copy();
            categorysum.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone);
            categorysum.NumberFormat = "#,##0";

        }
        private void DeleteNullRow(string thezone)
        {
            Stack<Excel.Range> deleted = new Stack<Excel.Range>();
            Excel.Range usedrow = thezoneWS.UsedRange.Rows.Offset[minbound - 1];
            int cnt = 0;
            foreach (Excel.Range cell in usedrow)
            {
                var v = cell.Columns[84 + 1].Value;

                if (v == null || Convert.ToDouble(v) == 0)
                {
                    deleted.Push(cell);
                }
            }
            foreach (Excel.Range row in deleted)
            {
                row.Delete();
            }
            maxbound = GetlastRow(thezone, "thezone", maxbound);
            SumFlag.Add(maxbound);
        }
        private void DeleteNullCol(string thezone)
        {
            Excel.Range usedrng = thezoneWS.UsedRange;
            for (int i = usedrng.Columns.Count; i >= 2; i--)
            {
                bool iszero = true;
                foreach (int sumrow in SumFlag)
                {
                    int num = Convert.ToInt32(usedrng.Cells[sumrow, i].Value);
                    if (num != 0)
                    {
                        iszero = false;
                    }
                }
                if (iszero)
                {
                    usedrng.Columns[i].Delete();
                }
            }
        }
        private void round()
        {
            int lastUsedRow = thezoneWS.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            // Find the last real column
            int lastUsedColumn = thezoneWS.Cells.Find("*", System.Reflection.Missing.Value,
                                           System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                           Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious,
                                           false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
            String valcolend = GetExcelColumnName(lastUsedColumn - 1);
            Excel.Range valrng = thezoneWS.Range["B4:" + valcolend + lastUsedRow.ToString()];
            valrng.NumberFormat = "#,##0";
            Object[,] val = valrng.Value;
            for (int i = 1; i < val.GetLength(0); i++)
            {
                for (int j = 1; j < val.GetLength(1); j++)
                {
                    if (val[i, j] != null)
                    {
                        val[i, j] = Math.Round(Convert.ToDouble(val[i, j]), 0);
                    }
                }
            }
            valrng.Value = val;
        }
        private void Save(string savepath)
        {
            if (System.IO.File.Exists(savepath))
                System.IO.File.Delete(savepath);
            thezoneWS.SaveAs(savepath);
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
                //foreach (KeyValuePair<string, Excel.Sheets> item in eWS)
                    //ReleaseExcelObject(item.Value);
                foreach (KeyValuePair<string, Excel.Workbook> item in eWB)
                    ReleaseExcelObject(item.Value);
                foreach (KeyValuePair<string, Excel.Application> item in eXL)
                    ReleaseExcelObject(item.Value);
                //eWS.Clear();
                eWB.Clear();
                eXL.Clear();
                colvalbyCenterName.Clear();
                colvalbyThezoneName.Clear();
                SumFlag.Clear();
                centerWS = null;
                thezoneWS = null;
                minbound = 4; maxbound = 0; columncnt = 0; totalrow = 0;
            }
            catch (Exception) { }
        }
        private void ReleaseExcelObject(object obj)
        {
            Marshal.ReleaseComObject(obj);
            GC.Collect();
        }
        #endregion
    }
}
