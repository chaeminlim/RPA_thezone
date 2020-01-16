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
        private JObject CurrentJsonObj;
        public ComboBoxItem SelectedcbItem { get; set; }
        public ComboBoxItem SelectedtzItem { get; set; }

        private string path = @"MappingInfo.json";
        public MappingTable()
        {
            InitializeComponent();
            DataContext = this;

            cbItems = new ObservableCollection<ComboBoxItem>();
            TheZoneItems = new ObservableCollection<ComboBoxItem>();

            string result = string.Empty;

            try
            {
                using (StreamReader file = new StreamReader(path, Encoding.GetEncoding("UTF-8")))
                using (JsonTextReader reader = new JsonTextReader(file))
                {
                    JObject object1 = (JObject)JToken.ReadFrom(reader);
                    List<string> clientList = object1.Properties().Select(p => p.Name).ToList();

                    foreach (string clientName in clientList)
                    {
                        if(clientName == "TheZone")
                        {
                            JArray thez = (JArray)object1[clientName];

                            foreach (string tz in thez)
                            {
                                TheZoneItems.Add(new ComboBoxItem { Content = tz });
                            }
                        }
                        else
                        {
                            cbItems.Add(new ComboBoxItem { Content = clientName });
                        }
                        
                    }

                    reader.Close();
                }
            }
            catch (System.IO.FileNotFoundException)
            {
                StatusLabel.Content = "파일을 찾을 수 없습니다.";
                return;
            }
        }

        private void ClientTypeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            using (StreamReader file = new StreamReader(path, Encoding.GetEncoding("UTF-8")))
            using (JsonTextReader reader = new JsonTextReader(file))
            {
                JObject object1 = (JObject)JToken.ReadFrom(reader);

                JObject elem = JObject.Parse(object1.SelectToken((string)SelectedcbItem.Content).ToString());

                CurrentJsonObj = elem;

                UpdateList();

                reader.Close();
            }
        }

        private void UpdateList()
        {
            List<string> mappingNameList = CurrentJsonObj.Properties().Select(p => p.Name).ToList();
            List<JToken> mappingValueList = CurrentJsonObj.Properties().Select(p => p.Value).ToList();

            MappingDataGrid.Items.Clear();

            for (int i = 0; i < mappingNameList.Count; i++)
            {
                try
                {
                    MappingDataGrid.Items.Add(new MappingDataMember(true, mappingNameList[i], (string)mappingValueList[i], "", "", "", ""));
                }
                catch (System.ArgumentException)
                {
                    JObject temp = (JObject)mappingValueList[i];
                    string ty = (string)temp["구분"];
                    JArray tyn = (JArray)temp["값"];
                    var tynn = new StringBuilder();
                    foreach (string s in tyn)
                    {
                        tynn.Append(s + ",");
                    }
                    tynn.Remove(tynn.Length-1, 1);
                    Console.WriteLine(tynn);
                    string tr = (string)temp["True"];
                    string fa = (string)temp["False"];
                    MappingDataGrid.Items.Add(new MappingDataMember(true, mappingNameList[i], "", ty, tynn.ToString(), tr, fa));
                }
            }
        }

        private void btn_AddRow_Click(object sender, RoutedEventArgs e)
        {
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
            while (MappingDataGrid.SelectedItems.Count >= 1)
            {
                MappingDataGrid.Items.Remove(MappingDataGrid.SelectedItem);
            }
        }

        private void btn_Save_Click(object sender, RoutedEventArgs e)
        {
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
    }
}
