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
        private JObject CurrentJsonObj;
        public ComboBoxItem SelectedcbItem { get; set; }

        private string path = @"..\..\..\MappingInfo.json";
        public MappingTable()
        {
            InitializeComponent();
            DataContext = this;

            cbItems = new ObservableCollection<ComboBoxItem>();

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
                        cbItems.Add(new ComboBoxItem { Content = clientName });
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
            List<string> mappingValueList = CurrentJsonObj.Properties().Select(p => (string)p.Value).ToList();

            MappingDataGrid.Items.Clear();

            for (int i = 0; i < mappingNameList.Count; i++)
            {
                MappingDataGrid.Items.Add(new MappingDataMember(mappingNameList[i], mappingValueList[i], true));
            }
        }

        private void btn_AddRow_Click(object sender, RoutedEventArgs e)
        {
            String JKey = FromTextBox.Text;
            String JValue = ToTextBox.Text;

            try
            {
                MappingDataGrid.Items.Add(new MappingDataMember(JKey, JValue, false));
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
                        CurrentJsonObj.Add(m.KeyString, m.ValueString);
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
            public MappingDataMember(string keystring, string valuestring, bool isIncluded)
            {
                KeyString = keystring;
                ValueString = valuestring;
                IsIncluded = isIncluded;
            }

            public string KeyString
            {
                get; set;
            }
            public string ValueString
            {
                get; set;
            }
            public bool IsIncluded
            {
                get; set;
            }

        }
    }
}
