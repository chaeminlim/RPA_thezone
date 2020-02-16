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
        private List<JObject> CurrentJsonObjList;
        public ComboBoxItem SelectedcbItem { get; set; }
        public ComboBoxItem SelectedtzItem { get; set; }
        public JObject CurrentJson { get; set; }
        public TextBox JsonTextBlock { get; set; }
        private string path = @"../../../MappingInfo.json";

        public MappingTable()
        {
            
            InitializeComponent();
            DataContext = this;

            cbItems = new ObservableCollection<ComboBoxItem>();
            TheZoneItems = new ObservableCollection<ComboBoxItem>();
        }

        



        public class MappingDataMember
        {
            public MappingDataMember(bool isIncluded, string keystring, string valuestring, string typestring, string typenamestring, string truestring, string falsestring)
            {
                IsIncluded = isIncluded;
                KeyString = keystring;
                ValueString = valuestring;
                TypeString = typestring;
                TypeNameString = typenamestring;
                TrueString = truestring;
                FalseString = falsestring;
            }

            public bool IsIncluded
            {
                get; set;
            }
            public string KeyString
            {
                get; set;
            }
            public string ValueString
            {
                get; set;
            }
            public string TypeString
            {
                get; set;
            }
            public string TypeNameString
            {
                get; set;
            }
            public string TrueString
            {
                get; set;
            }
            public string FalseString
            {
                get; set;
            }
        }

        #region 콤보박스관련
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

            UpdateJsonTreeView();
        }

        private void UpdateJsonTreeView()
        {
            this.JsonTreeView.Items.Clear();

            String companyName = CurrentJson["회사명"].ToString();
            List<JObject> sheetList = CurrentJson["시트"].ToObject<List<JObject>>();

            List<TreeViewItem> childTreeViewItemsList = new List<TreeViewItem>();
            List<TreeViewItem> grandChildTreeViewItemsList = new List<TreeViewItem>();
            TreeViewItem treeViewItem = new TreeViewItem
            {
                Header = companyName
            };
            foreach (JObject sheet in sheetList)
            {
                String sheetName = sheet["시트이름"].ToString();
                List<JObject> mappingList = sheet["배치표"].ToObject<List<JObject>>();

                TreeViewItem childTreeViewItem = new TreeViewItem
                {
                    Header = sheetName
                };

                childTreeViewItemsList.Add(childTreeViewItem);
                treeViewItem.ItemsSource = childTreeViewItemsList;
                foreach (JObject mapping in mappingList)
                {
                    String cellPoint = mapping["셀위치"].ToString();
                    String cellName = mapping["셀이름"].ToString();

                    TreeViewItem GrandchildTreeViewItem1 = new TreeViewItem
                    {
                        Header = cellPoint
                    };
                    TreeViewItem GrandchildTreeViewItem2 = new TreeViewItem
                    {
                        Header = cellName
                    };
                    grandChildTreeViewItemsList.Add(GrandchildTreeViewItem1);
                    grandChildTreeViewItemsList.Add(GrandchildTreeViewItem2);
                    childTreeViewItem.ItemsSource = grandChildTreeViewItemsList;

                    /*
                    try
                    {
                        List<JObject> toZone = mapping["더존이름"].ToObject<List<JObject>>();
                        

                    }
                    catch (System.ArgumentException)
                    {
                        List<String> toZone = mapping["더존이름"].ToObject<List<String>>();

                    }
                    */
                }
            } 

            JsonTreeView.Items.Add(treeViewItem);

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
