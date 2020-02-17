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
        private Tuple<String, String, String> CurrentSelectedObjInfo;
        public ComboBoxItem SelectedcbItem { get; set; }
        public ComboBoxItem SelectedtzItem { get; set; }
        public JObject CurrentJson { get; set; }
        public TextBox JsonTextBlock { get; set; }
        private string path = @"../../../MappingInfo.json";
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
        private void ClientTypeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            UpdateJsonTreeView();
        }

        private void UpdateJsonTreeView()
        {
            this.JsonTreeView.Items.Clear();
            using (StreamReader file = new StreamReader(path, Encoding.GetEncoding("UTF-8")))
            using (JsonTextReader reader = new JsonTextReader(file))
            {
                JObject currentCompanyInfo = null;

                JObject fullObj = (JObject)JToken.ReadFrom(reader);
                List<JObject> elemList = fullObj["회사목록"].ToObject<List<JObject>>();

                foreach (JObject companyInfo in elemList)
                {
                    if (companyInfo["회사명"].ToString() == (string)SelectedcbItem.Content)
                    {
                        currentCompanyInfo = companyInfo;
                        break;
                    }
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

                    TreeViewItem cellPointTreeViewItem = new TreeViewItem
                    {
                        Header = "셀위치 : " + cellPoint
                    };

                    TreeViewItem mappingInfoTreeViewItem = new TreeViewItem
                    {
                        Header = "배치정보" + i++ +  " : " + cellName,
                        Tag = new Tuple<String, String>(cellName, sheetName) 
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
                            Header = toZones,
                            
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
                                Header = "구분 : " + dividPoint + ", " + values + ", True :" + trueVal + ", False " + falseVal
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
                    Header = "셀 추가하기"
                };
                addCellButtonTreeViewItem.Selected += AddCellButtonTreeViewItem_Selected;
                mappingInfoTreeViewItemsList.Add(addCellButtonTreeViewItem);
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
                CurrentSelectedObjInfo = new Tuple<String, String, String>("COMPANY", (String)((TreeViewItem)sender).Header, "");
                BtnDeleteJson.IsEnabled = false;
            }
            EventFlag = 3;
        }

        private void SheetTreeViewItem_Selected(object sender, RoutedEventArgs e)
        {
            if (EventFlag >= 2)
            {
                CurrentSelectedObjInfo = new Tuple<String, String, String>("SHEET", (String)((TreeViewItem)sender).Tag, null);
                BtnDeleteJson.IsEnabled = true;
            }
            EventFlag = 2;
        }

        private void MappingInfoTreeViewItem_Selected(object sender, RoutedEventArgs e)
        {
            if (EventFlag >= 1)
            {
                CurrentSelectedObjInfo = new Tuple<String, String, String>("MAPPING", ((Tuple<String, String>)((TreeViewItem)sender).Tag).Item1, ((Tuple<String, String>)((TreeViewItem)sender).Tag).Item2);
                BtnDeleteJson.IsEnabled = true;
            }
            EventFlag = 1;

        }
        private void CellPointTreeViewItem_Selected(object sender, RoutedEventArgs e)
        {
            EventFlag = 0;
            BtnDeleteJson.IsEnabled = false;
        }

        private void AddCellButtonTreeViewItem_Selected(object sender, RoutedEventArgs e)
        {
            EventFlag = 0;
            BtnDeleteJson.IsEnabled = false;
            Console.WriteLine("셀추가핸들러");
        }

        private void AddSheetButtonTreeViewItem_Selected(object sender, RoutedEventArgs e)
        {
            EventFlag = 0;
            BtnDeleteJson.IsEnabled = false;
            Console.WriteLine("시트추가핸들러");
        }

        #endregion
        #endregion


        #region 버튼 처리기
        
        private void BtnDeleteJson_Click(object sender, RoutedEventArgs e)
        {
            String jsonType = CurrentSelectedObjInfo.Item1;
            String jsonName = CurrentSelectedObjInfo.Item2;
            String jsonSheet = CurrentSelectedObjInfo.Item3;

            // JObject CurrentJson
            if (jsonType == "SHEET")
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
                                if (sheetInfo["시트이름"].ToString() == jsonName)
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
            }
            else
            {
                //error
            }

            UpdateJsonTreeView();
        }

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
