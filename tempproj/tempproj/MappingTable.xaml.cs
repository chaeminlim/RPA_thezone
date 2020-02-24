using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace tempproj
{
    /// <summary>
    /// MappingTable.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MappingTable : Window
    {
        public ObservableCollection<ComboBoxItem> cbItems { get; set; }
        public ObservableCollection<ComboBoxItem> TheZoneItems { get; set; }
        private Tuple<String, String, String, String, dynamic> CurrentSelectedObjInfo;
        public ComboBoxItem SelectedcbItem { get; set; }
        public ComboBoxItem SelectedtzItem { get; set; }
        public JObject CurrentJson { get; set; }
        public TextBox JsonTextBlock { get; set; }
        private string path = @"./MappingInfo.json";
        private int EventFlag { get; set; }

        public MappingTable()
        {
            
            InitializeComponent();
            DataContext = this;

            cbItems = new ObservableCollection<ComboBoxItem>();
            TheZoneItems = new ObservableCollection<ComboBoxItem>();

            EventFlag = 3;
        }


        #region 콤보박스관련
        #region 콤보박스 처리(not eventHandler)

        private void ClientTypeComboBox_Loaded(object sender, RoutedEventArgs e)
        {
            if(cbItems == null)
            {
                return;
            }
            cbItems.Clear();

            using (StreamReader file = new StreamReader(path, Encoding.GetEncoding("UTF-8")))
            using (JsonTextReader reader = new JsonTextReader(file))
            {
                JObject fullObj = (JObject)JToken.ReadFrom(reader);
                List<JObject> companyList = fullObj["회사목록"].ToObject<List<JObject>>();

                foreach (JObject companyJsonObj in companyList)
                {
                    cbItems.Add(new ComboBoxItem { Content = companyJsonObj["회사명"].ToString() });
                }
                reader.Close();
            }

        }

        private void ClientTypeComboBox_ReLoad(String companyNameParam)
        {

            using (StreamReader file = new StreamReader(path, Encoding.GetEncoding("UTF-8")))
            using (JsonTextReader reader = new JsonTextReader(file))
            {
                JObject fullObj = (JObject)JToken.ReadFrom(reader);
                List<JObject> companyList = fullObj["회사목록"].ToObject<List<JObject>>();

                cbItems.Clear();

                foreach (JObject companyJsonObj in companyList)
                {
                    cbItems.Add(new ComboBoxItem { Content = companyJsonObj["회사명"].ToString() });
                }
                reader.Close();
            }

            foreach (ComboBoxItem cbi in cbItems)
            {
                if (cbi.Content.ToString() == companyNameParam)
                {
                    SelectedcbItem = cbi;
                    break;
                }
            }
        }
        private void ClientTypeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                UpdateJsonTreeView();
            }
            catch(System.NullReferenceException)
            {
                JsonTreeView.Items.Clear();
            }
        }

        

        private void UpdateJsonTreeView(String companyNameParam = null)
        {
            JsonTreeView.Items.Clear();
            
            using (StreamReader file = new StreamReader(path, Encoding.GetEncoding("UTF-8")))
            using (JsonTextReader reader = new JsonTextReader(file))
            {
                JObject currentCompanyInfo = null;

                JObject fullObj = (JObject)JToken.ReadFrom(reader);
                List<JObject> elemList = fullObj["회사목록"].ToObject<List<JObject>>();

                foreach (JObject companyInfo in elemList)
                {
                    String tempCompName = "";
                    if (companyNameParam == null)
                        tempCompName = (string)SelectedcbItem.Content;
                    else
                    {
                        tempCompName = companyNameParam;
                    }
                   
                    if (companyInfo["회사명"].ToString() == tempCompName)
                    {
                        currentCompanyInfo = companyInfo;
                        break;
                    }
                }

                if(currentCompanyInfo == null)
                {
                    return;
                }

                reader.Close();
                CurrentJson = currentCompanyInfo;
            }

            String companyName = CurrentJson["회사명"].ToString();
            List<JObject> sheetList = CurrentJson["시트"].ToObject<List<JObject>>();
            List<TreeViewItem> sheetInfoTreeViewItemsList = new List<TreeViewItem>();

            TreeViewItem treeViewItem = new TreeViewItem
            {
                Header = companyName
            };

            sheetInfoTreeViewItemsList.Clear();
            foreach (JObject sheet in sheetList)
            {
                String sheetName = sheet["시트이름"].ToString();
                List<JObject> mappingList = sheet["배치표"].ToObject<List<JObject>>();

                TreeViewItem sheetTreeViewItem = new TreeViewItem
                {
                    Header = "시트이름 : " + sheetName,
                    Tag = sheetName
                };
                List<TreeViewItem> mappingInfoTreeViewItemsList = new List<TreeViewItem>();
                int i = 1;
                foreach (JObject mapping in mappingList)
                {
                    List<TreeViewItem> cellInfoTreeViewItemsList = new List<TreeViewItem>();
                    
                    String cellPoint = mapping["셀위치"].ToString();
                    String cellName = mapping["셀이름"].ToString();
                    var thezoneName = mapping["더존이름"];

                    TreeViewItem cellPointTreeViewItem = new TreeViewItem
                    {
                        Header = "셀위치 : " + cellPoint
                    };

                    TreeViewItem mappingInfoTreeViewItem = new TreeViewItem
                    {
                        Header = "배치정보" + i++ +  " : " + cellName,
                        Tag = new Tuple<String, String, String, dynamic>(sheetName, cellName, cellPoint, thezoneName)
                    };

                    cellPointTreeViewItem.Selected += CellPointTreeViewItem_Selected;
                    cellInfoTreeViewItemsList.Add(cellPointTreeViewItem);

                    try
                    {
                        String toZones = "더존이름 : ";
                        foreach (JValue toZone in mapping["더존이름"])
                        {
                            toZones += ", " + toZone.ToString();
                        }
                        TreeViewItem toZoneTreeViewItem = new TreeViewItem
                        {
                            Header = toZones
                        };

                        toZoneTreeViewItem.Selected += CellPointTreeViewItem_Selected;
                        cellInfoTreeViewItemsList.Add(toZoneTreeViewItem);
                    }
                    catch (Exception)
                    {
                        List<JObject> toZones = mapping["더존이름"].ToObject<List<JObject>>();
                        foreach (JObject toZone in toZones)
                        {
                            String dividPoint = toZone["구분"].ToString();
                            String values = "";
                            foreach (JValue value in toZone["값"])
                            {
                                values += ", " + value.ToString();
                            }
                            String trueVal = toZone["True"].ToString();
                            String falseVal = toZone["False"].ToString();

                            TreeViewItem toZoneTreeViewItem = new TreeViewItem
                            {
                                Header = "구분 : " + dividPoint + "\n값 : " + values + "\nTrue :" + trueVal + "\nFalse " + falseVal
                            };
                            toZoneTreeViewItem.Selected += CellPointTreeViewItem_Selected;
                            cellInfoTreeViewItemsList.Add(toZoneTreeViewItem);

                        }
                    }
                    
                    mappingInfoTreeViewItem.ItemsSource = cellInfoTreeViewItemsList;
                    mappingInfoTreeViewItem.Selected += MappingInfoTreeViewItem_Selected;
                    mappingInfoTreeViewItemsList.Add(mappingInfoTreeViewItem);
                }

                TreeViewItem addCellButtonTreeViewItem = new TreeViewItem()
                {
                    Header = "셀 추가하기",
                    Tag =  sheetName
                };
                addCellButtonTreeViewItem.Selected += AddCellButtonTreeViewItem_Selected;
                mappingInfoTreeViewItemsList.Add(addCellButtonTreeViewItem);
                sheetTreeViewItem.IsExpanded = true;
                sheetTreeViewItem.ItemsSource = mappingInfoTreeViewItemsList;
                sheetTreeViewItem.Selected += SheetTreeViewItem_Selected;
                sheetInfoTreeViewItemsList.Add(sheetTreeViewItem);
            }
            TreeViewItem addSheetButtonTreeViewItem = new TreeViewItem()
            {
                Header = "시트 추가하기"
            };
            addSheetButtonTreeViewItem.Selected += AddSheetButtonTreeViewItem_Selected;
            sheetInfoTreeViewItemsList.Add(addSheetButtonTreeViewItem);
            treeViewItem.IsExpanded = true;
            treeViewItem.ItemsSource = sheetInfoTreeViewItemsList;
            treeViewItem.Selected += TreeViewItem_Selected;
            JsonTreeView.Items.Add(treeViewItem);

        }






        #endregion
        #region 이벤트 핸들러
        private void TreeViewItem_Selected(object sender, RoutedEventArgs e)
        {
            if (EventFlag >= 3)
            {
                CurrentSelectedObjInfo = new Tuple<String, String, String, String, dynamic>("COMPANY", (String)((TreeViewItem)sender).Header, "", "", "");
                BtnDeleteJson.IsEnabled = true;
                EditSheetTabItem.IsEnabled = false;
                EditMappingTabItem.IsEnabled = false;
                EditCompanyTabItem.IsSelected = true;
            }
            EventFlag = 3;
        }

        private void SheetTreeViewItem_Selected(object sender, RoutedEventArgs e)
        {
            if (EventFlag >= 2)
            {
                CurrentSelectedObjInfo = new Tuple<String, String, String, String, dynamic>("SHEET", (String)((TreeViewItem)sender).Tag, "", "", "");
                BtnDeleteJson.IsEnabled = true;
                EditSheetTabItem.IsEnabled = false;
                EditMappingTabItem.IsEnabled = false;
                EditCompanyTabItem.IsSelected = true;
            }
            EventFlag = 2;
        }

        private void MappingInfoTreeViewItem_Selected(object sender, RoutedEventArgs e)
        {
            if (EventFlag >= 1)
            {
                CurrentSelectedObjInfo = new Tuple<String, String, String, String , dynamic>("MAPPING", ((Tuple<String, String, String, dynamic>)((TreeViewItem)sender).Tag).Item1, ((Tuple<String, String, String, dynamic>)((TreeViewItem)sender).Tag).Item2,
                    ((Tuple<String, String, String, dynamic>)((TreeViewItem)sender).Tag).Item3, ((Tuple<String, String, String, dynamic>)((TreeViewItem)sender).Tag).Item4);
                BtnDeleteJson.IsEnabled = true;
                EditSheetTabItem.IsEnabled = false;
                EditMappingTabItem.IsEnabled = true;
                EditMappingTabItem.IsSelected = true;
                

                divisionCheckBox.IsChecked = false;
                sheetNameTextBox.Clear();
                cellPointTextBox.Clear();
                cellNameTextBox.Clear();
                divisionTextBox.Clear();
                theZoneTrueListBox.Items.Clear();
                theZoneFalseListBox.Items.Clear();
                valueCheckListBox.Items.Clear();
                theZoneFalseTextBox.Clear();
                theZoneTrueTextBox.Clear();
                valueCheckTextBox.Clear();

                sheetNameTextBox.Text = CurrentSelectedObjInfo.Item2;
                cellNameTextBox.Text = CurrentSelectedObjInfo.Item3;
                cellPointTextBox.Text = CurrentSelectedObjInfo.Item4;
                var thezone = CurrentSelectedObjInfo.Item5;
                if (thezone[0] is JObject)
                {
                    JObject mappingByPos = (JObject)thezone[0];
                    String divisionText = mappingByPos["구분"].ToString();
                    List<String> pos = mappingByPos["값"].ToObject<List<String>>();
                    String trueText = mappingByPos["True"].ToString();
                    String falseText = mappingByPos["False"].ToString();
                    divisionCheckBox.IsChecked = true;
                    divisionTextBox.Text = divisionText;
                    theZoneTrueListBox.Items.Add(trueText);
                    theZoneFalseListBox.Items.Add(falseText);
                    foreach (var item in pos)
                    {
                        valueCheckListBox.Items.Add(item);
                    }

                }
                else
                {
                    foreach (String name in thezone)
                    {
                        theZoneTrueListBox.Items.Add(name);
                    }
                }
            }
            EventFlag = 1;

        }
        private void CellPointTreeViewItem_Selected(object sender, RoutedEventArgs e)
        {
            EventFlag = 0;
            BtnDeleteJson.IsEnabled = false;
            EditSheetTabItem.IsEnabled = false;
            EditMappingTabItem.IsEnabled = false;
        }

        

        private void AddCellButtonTreeViewItem_Selected(object sender, RoutedEventArgs e)
        {
            CurrentSelectedObjInfo = null;
            EventFlag = 0;
            BtnDeleteJson.IsEnabled = false;
            EditSheetTabItem.IsEnabled = false;
            EditMappingTabItem.IsEnabled = true;
            EditMappingTabItem.IsSelected = true;


            divisionCheckBox.IsChecked = false;
            sheetNameTextBox.Clear();
            cellPointTextBox.Clear();
            cellNameTextBox.Clear();
            divisionTextBox.Clear();
            theZoneTrueListBox.Items.Clear();
            theZoneFalseListBox.Items.Clear();
            theZoneFalseTextBox.Clear();
            theZoneTrueTextBox.Clear();
            valueCheckTextBox.Clear();
        }

        private void AddSheetButtonTreeViewItem_Selected(object sender, RoutedEventArgs e)
        {
            EventFlag = 0;
            BtnDeleteJson.IsEnabled = false;
            EditSheetTabItem.IsEnabled = true;
            EditSheetTabItem.IsSelected = true;
            EditMappingTabItem.IsEnabled = false;

        }

        #endregion
        #endregion

        #region 삭제 버튼
        
        private void BtnDeleteJson_Click(object sender, RoutedEventArgs e)
        {
            String jsonType = CurrentSelectedObjInfo.Item1;
            String jsonComp = CurrentSelectedObjInfo.Item2;
            String jsonSheet = CurrentSelectedObjInfo.Item2;
            
            String jsonName = CurrentSelectedObjInfo.Item3;
            

            // JObject CurrentJson
            if(jsonType == "COMPANY")
            {
                using (StreamReader file = new StreamReader(path, Encoding.GetEncoding("UTF-8")))
                using (JsonTextReader reader = new JsonTextReader(file))
                {
                    JObject fullObj = (JObject)JToken.ReadFrom(reader);
                    //List<JObject> elemList = fullObj["회사목록"].ToObject<List<JObject>>();
                    foreach (JObject companyInfo in fullObj["회사목록"])
                    {
                        if (companyInfo["회사명"].ToString() == jsonComp)
                        {
                            companyInfo.Remove();
                            break;
                        }
                    }

                    string output = JsonConvert.SerializeObject(fullObj, Newtonsoft.Json.Formatting.Indented);
                    reader.Close();
                    File.WriteAllText(path, output, Encoding.GetEncoding("UTF-8"));
                }
                ClientTypeComboBox_ReLoad("");
                UpdateJsonTreeView("");
            }
            else if (jsonType == "SHEET")
            {

                using (StreamReader file = new StreamReader(path, Encoding.GetEncoding("UTF-8")))
                using (JsonTextReader reader = new JsonTextReader(file))
                {
                    JObject fullObj = (JObject)JToken.ReadFrom(reader);
                    //List<JObject> elemList = fullObj["회사목록"].ToObject<List<JObject>>();
                    foreach (JObject companyInfo in fullObj["회사목록"])
                    {
                        if (companyInfo["회사명"].ToString() == CurrentJson["회사명"].ToString())
                        {
                            foreach(JObject sheetInfo in companyInfo["시트"])
                            {
                                if (sheetInfo["시트이름"].ToString() == jsonSheet)
                                {
                                    sheetInfo.Remove();
                                    break;
                                }
                            }
                            break;
                        }
                    }

                    string output = JsonConvert.SerializeObject(fullObj, Newtonsoft.Json.Formatting.Indented);
                    reader.Close();
                    File.WriteAllText(path, output, Encoding.GetEncoding("UTF-8"));
                }
                UpdateJsonTreeView();
            }
            else if (jsonType == "MAPPING")
            {
                using (StreamReader file = new StreamReader(path, Encoding.GetEncoding("UTF-8")))
                using (JsonTextReader reader = new JsonTextReader(file))
                {
                    JObject fullObj = (JObject)JToken.ReadFrom(reader);

                    foreach (JObject companyInfo in fullObj["회사목록"])
                    {
                        if (companyInfo["회사명"].ToString() == CurrentJson["회사명"].ToString())
                        {
                            foreach (JObject sheetInfo in companyInfo["시트"])
                            {
                                if (sheetInfo["시트이름"].ToString() == jsonSheet)
                                {
                                    JToken temp = sheetInfo["배치표"];
                                    foreach (JObject mappingInfo in sheetInfo["배치표"])
                                    {
                                        if(mappingInfo["셀이름"].ToString() == jsonName)
                                        {
                                            mappingInfo.Remove();
                                            break;
                                        }

                                    }
                                    break;
                                }
                            }
                            break;
                        }
                    }

                    string output = JsonConvert.SerializeObject(fullObj, Newtonsoft.Json.Formatting.Indented);
                    reader.Close();
                    File.WriteAllText(path, output, Encoding.GetEncoding("UTF-8"));
                }
                UpdateJsonTreeView();
            }
            else
            {
                //error
            }


        }

        #endregion

        #region 정보 추가 기능
        #region 추가기능 
        private void EditCompanyYesButton_Click(object sender, RoutedEventArgs e)
        {
            
            if(CompanyNameTextBox.Text == "")
            {
                MessageBox.Show("회사명을 입력하세요.");
                return;
            }
            String tempCompName = CompanyNameTextBox.Text;

            dynamic companyObj = new JObject();
            companyObj.회사명 = CompanyNameTextBox.Text;
            companyObj.시트 = new JArray();

            using (StreamReader file = new StreamReader(path, Encoding.GetEncoding("UTF-8")))
            using (JsonTextReader reader = new JsonTextReader(file))
            {
                JObject fullObj = (JObject)JToken.ReadFrom(reader);
                ((JArray)fullObj["회사목록"]).Add(companyObj);

                string output = JsonConvert.SerializeObject(fullObj, Newtonsoft.Json.Formatting.Indented);
                Console.WriteLine(output);
                reader.Close();
                File.WriteAllText(path, output, Encoding.GetEncoding("UTF-8"));
            }

            ClientTypeComboBox_ReLoad(tempCompName);
            UpdateJsonTreeView(tempCompName);
        }

        private void EditSheetYesButton_Click(object sender, RoutedEventArgs e)
        {
            if (SheetNameTextBox.Text == "")
            {
                MessageBox.Show("시트이름을 입력하세요.");
                return;
            }

            dynamic sheetObj = new JObject();
            sheetObj.시트이름 = SheetNameTextBox.Text;
            sheetObj.배치표 = new JArray();
            String tempCompName = CurrentJson["회사명"].ToString();

            using (StreamReader file = new StreamReader(path, Encoding.GetEncoding("UTF-8")))
            using (JsonTextReader reader = new JsonTextReader(file))
            {
                JObject fullObj = (JObject)JToken.ReadFrom(reader);
                foreach (JObject companyInfo in fullObj["회사목록"])
                {
                    if (companyInfo["회사명"].ToString() == CurrentJson["회사명"].ToString())
                    {
                        ((JArray)companyInfo["시트"]).Add(sheetObj);
                        break;
                    }
                }

                string output = JsonConvert.SerializeObject(fullObj, Newtonsoft.Json.Formatting.Indented);
                reader.Close();
                File.WriteAllText(path, output, Encoding.GetEncoding("UTF-8"));
            }

            ClientTypeComboBox_ReLoad(tempCompName);
            UpdateJsonTreeView(tempCompName);
        }


        private void EditMappingYesButton_Click(object sender, RoutedEventArgs e)
        {
            
            String cellPoint = cellPointTextBox.Text;
            String cellName = cellNameTextBox.Text;
            String division = divisionTextBox.Text;
            String sheetName = sheetNameTextBox.Text;
            
            String tempCompName = CurrentJson["회사명"].ToString();
            
            List<String> valueCheckList = new List<String>();
            List<String> theZoneTrueList = new List<String>();
            List<String> theZoneFalseList = new List<String>();

            foreach (String listBoxItem in valueCheckListBox.Items)
                valueCheckList.Add(listBoxItem);
            
            foreach (String listBoxItem in theZoneTrueListBox.Items) {
                Console.WriteLine(listBoxItem);
                theZoneTrueList.Add(listBoxItem);
            }
            foreach (String listBoxItem in theZoneFalseListBox.Items)
                theZoneFalseList.Add(listBoxItem);
            
            dynamic mappingObj = new JObject();
            
            if ((bool)divisionCheckBox.IsChecked)
            {
                //mappingObj = new JObject();
                mappingObj.셀위치 = cellPoint;
                mappingObj.셀이름 = cellName;
                dynamic mappingtzInfo = new JObject();
                mappingtzInfo.구분 = division;
                mappingtzInfo.값 = new JArray(valueCheckList);

                try
                {
                    mappingtzInfo.True = theZoneTrueList[0];
                    mappingtzInfo.False = theZoneFalseList[0];
                }
                catch (System.ArgumentOutOfRangeException)
                {
                    MessageBox.Show("값을 올바르게 입력하세요.\n" +
                        "리스트 박스에 값이 추가되지 않았습니다.");
                    return;
                }

                mappingObj.더존이름 = new JArray(mappingtzInfo);
            }
            else // 일반경우
            {
                //mappingObj = new JObject();
                mappingObj.셀위치 = cellPoint;
                mappingObj.셀이름 = cellName;
                mappingObj.더존이름 = new JArray(theZoneTrueList);

            }
            
            using (StreamReader file = new StreamReader(path, Encoding.GetEncoding("UTF-8")))
            using (JsonTextReader reader = new JsonTextReader(file))
            {
                JObject fullObj = (JObject)JToken.ReadFrom(reader);
                foreach (JObject companyInfo in fullObj["회사목록"])
                {
                    if (companyInfo["회사명"].ToString() == CurrentJson["회사명"].ToString())
                    {
                        foreach (JObject sheetInfo in companyInfo["시트"])
                        {
                            if (sheetInfo["시트이름"].ToString() == sheetName)
                            {
                                bool flag = true;
                                if (CurrentSelectedObjInfo != null)
                                {
                                    foreach (JObject mappingInfo in sheetInfo["배치표"])
                                    { 
                                    
                                        if (mappingInfo["셀위치"].ToString().Equals(CurrentSelectedObjInfo.Item4))
                                        {
                                            flag = false;
                                            mappingInfo["셀위치"] = mappingObj.셀위치;
                                            mappingInfo["셀이름"] = mappingObj.셀이름;
                                            mappingInfo["더존이름"] = mappingObj.더존이름;
                                            //foreach (String thezoneName in theZoneTrueList)
                                            //{
                                            //    mappingInfo["더존이름"].ToObject<List<String>>().Add(thezoneName);
                                            //}
                                            break;
                                        }
                                    }
                                }
                                if(flag)
                                {
                                    ((JArray)sheetInfo["배치표"]).Add(mappingObj);
                                    break;
                                }
                                
                            }
                            else
                            {
                                MessageBox.Show("시트 정보가 잘못되었습니다.");
                                return;
                            }
                        }
                        break;
                    }
                }
                
                string output = JsonConvert.SerializeObject(fullObj, Newtonsoft.Json.Formatting.Indented);
                reader.Close();
                File.WriteAllText(path, output, Encoding.GetEncoding("UTF-8"));
            }

            ClientTypeComboBox_ReLoad(tempCompName);
            UpdateJsonTreeView(tempCompName);
        }

        private void btnvalueCheckListBox_Click(object sender, RoutedEventArgs e)
        {
            valueCheckListBox.Items.Add(valueCheckTextBox.Text);
            valueCheckTextBox.Clear();
        }

        private void btnTheZoneTrueListBox_Click(object sender, RoutedEventArgs e)
        {
            theZoneTrueListBox.Items.Add(theZoneTrueTextBox.Text);
            theZoneTrueTextBox.Clear();
        }

        private void btnTheZoneFalseListBox_Click(object sender, RoutedEventArgs e)
        {
            theZoneFalseListBox.Items.Add(theZoneFalseTextBox.Text);
            theZoneFalseTextBox.Clear();
        }

        private void btnClearTheZoneFalseListBox_Click(object sender, RoutedEventArgs e)
        {
            theZoneFalseListBox.Items.Clear();
        }

        private void btnClearvalueCheckListBox_Click(object sender, RoutedEventArgs e)
        {
            valueCheckListBox.Items.Clear();
        }

        private void btnClearTheZoneTrueListBox_Click(object sender, RoutedEventArgs e)
        {
            theZoneTrueListBox.Items.Clear();
        }
        #endregion
        #region 취소기능

        private void EditCompanyNoButton_Click(object sender, RoutedEventArgs e)
        {
            CompanyNameTextBox.Clear();
        }

        private void EditSheetNoButton_Click(object sender, RoutedEventArgs e)
        {
            SheetNameTextBox.Clear();
        }
        
        private void EditMappingNoButton_Click(object sender, RoutedEventArgs e)
        {
            cellPointTextBox.Clear();
            cellNameTextBox.Clear();
            theZoneTrueTextBox.Clear();
            divisionTextBox.Clear();
            valueCheckTextBox.Clear();
            theZoneFalseTextBox.Clear();
            theZoneTrueListBox.Items.Clear();
            theZoneFalseListBox.Items.Clear();
        }


        #endregion

        #endregion
    }

    #region 기존코드
    /*
    private void btn_AddRow_Click(object sender, RoutedEventArgs e)
    {
        #region 기존코드
        /*
        String JKey = FromTextBox.Text;
        String JValue;
        try
        {
            JValue = (String)SelectedtzItem.Content;
        }
        catch (System.NullReferenceException)
        {
            JValue = "";
        }
        String JType = TypeTextBox.Text;
        String JTypeName = TypeNameTextBox.Text;
        String JTrue = TrueTextBox.Text;
        String JFalse = FalseTextBox.Text;
        try
        {
            MappingDataGrid.Items.Add(new MappingDataMember(false, JKey, JValue, JType, JTypeName, JTrue, JFalse));
        }
        catch (NullReferenceException)
        {
            StatusLabel.Content = "선택되지 않았습니다.";
            return;
        }
        catch (System.ArgumentException)
        {

            StatusLabel.Content = "중복되는 컬럼이 존재합니다.";
            return;
        }

    }
    private void btn_deleteRow_Click(object sender, RoutedEventArgs e)
    {
        #region 기존코드
        /*
        while (MappingDataGrid.SelectedItems.Count >= 1)
        {
            MappingDataGrid.Items.Remove(MappingDataGrid.SelectedItem);
        }


    }

    private void btn_Save_Click(object sender, RoutedEventArgs e)
    {
        #region 기존코드
        /*
        try
        {
            using (StreamReader file = new StreamReader(path, Encoding.GetEncoding("UTF-8")))
            using (JsonTextReader reader = new JsonTextReader(file))
            {
                CurrentJsonObj.RemoveAll();
                foreach (MappingDataMember m in MappingDataGrid.Items)
                {
                    if(m.ValueString != "" && m.TypeString != "")
                    {
                        MessageBox.Show("To값과 Type값은 동시에 있을수 없습니다");
                        break;
                    }
                    if (m.ValueString == "")
                    {
                        JArray jt = new JArray();
                        string[] sa = m.TypeNameString.Split(',');
                        foreach (string s in sa)
                            jt.Add(s);

                        JObject t = new JObject();
                        t.Add("구분", m.TypeString);
                        t.Add("값", jt);
                        t.Add("True", m.TrueString);
                        t.Add("False", m.FalseString);

                        CurrentJsonObj.Add(m.KeyString, t);
                    }
                    else
                    {
                        CurrentJsonObj.Add(m.KeyString, m.ValueString);
                    }
                }

                JObject object1 = (JObject)JToken.ReadFrom(reader);
                object1[SelectedcbItem.Content] = CurrentJsonObj;
                string output = Newtonsoft.Json.JsonConvert.SerializeObject(object1, Newtonsoft.Json.Formatting.Indented);
                reader.Close();

                File.WriteAllText(path, output, Encoding.GetEncoding("UTF-8"));
            }
            MessageBox.Show("적용이 완료되었습니다");
            UpdateList();
        }
        catch (System.ArgumentException)
        {
            MessageBox.Show("같은 from값은 넣을수 없습니다");
        }

        #endregion
    }
    */
    #endregion  
}
