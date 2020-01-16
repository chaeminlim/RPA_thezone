﻿using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using tempproj.Controller;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.ObjectModel;
using System.Windows.Threading;

namespace tempproj
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        private MainController MainControllerObject;
        private ContextController contextController;
        private ExcelActivity excelActivity;
        private List<ExcelWorkQueueDataStruct> ExcelWorkQueue;
        private string path = @"MappingInfo.json";
        private double scr = double.MaxValue;

        public MainWindow()
        {
            InitializeComponent();
            InitAnnoucement();
            InitContext();
            InitExcelContext();
        }

        public void UpdateWindow()
        {
            // 화면 객체의 변경사항을 즉시 업데이트한다
            System.Windows.Threading.Dispatcher.CurrentDispatcher.Invoke(
                      System.Windows.Threading.DispatcherPriority.Background,
                      new System.Threading.ThreadStart(delegate { }));
        }
        public void WriteDebugLine(string text)
        {
            DebugConsoleBlock.Text += text + Environment.NewLine;
            scrollv.ScrollToVerticalOffset(scr);
            UpdateWindow();
        }
        private void InitContext()
        {
            excelActivity = new ExcelActivity();
            contextController = new ContextController();
        }

        private void InitAnnoucement()
        {
            AnnouncementTextBlock.Text = @"<프로그램 사용법>
1. 업로드 자료 불러오기 버튼을 눌러 급여수당일괄업로드 엑셀 파일을 지정하세요.
2. 급여 자료 불러오기 버튼을 눌러 작업할 엑셀 파일을 불러오세요.
3. 콤보 박스를 이용해 자료에 맞는 각 회사 이름을 지정하세요.
4. 작업 시작하기를 눌러 작업을 진행하세요. 작업이 완료되었음을 알리는 메세지창이 나타나기 전까지 프로그램을 종료하지 마세요.
5. 작업 대상의 급여자료 파일들은 프로그램 시작 전 닫아주세요.
";
        }



          
        private void ClearListBox_Click(object sender, RoutedEventArgs e)
        {
            ClearAllCurrentQueueData();
        }

        private void btnStartWorkflow_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Excel 파일을 올바르게 지정하였습니까?\n 프로그램이 시작되면 아무 작업도 수행하지 마세요.", "Question", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
            {
            }
            else
            {
                contextController.ClearExcelPath();

                ExcelListView.SelectAll();
                //WorkflowXmlListView.SelectAll();
                //contextController.SetWorkflowXmlPath((string)WorkflowXmlListView.SelectedItem);

                foreach (string excel in ExcelWorkEndView.SelectedItems)
                {
                    contextController.AddExcelPath(excel);
                }

                StartWorkflow();
            }
        }

        private void StartWorkflow()
        {
            WorkflowController wf = new WorkflowController();

            try
            {
                String[] xmllines = System.IO.File.ReadAllLines(contextController.GetWorkflowXmlPath());
                foreach (String line in xmllines)
                {
                    int ErrCode = wf.DoActionXml(line);
                    if (ErrCode == 0 || ErrCode >= 2)
                    {
                        WriteDebugLine("Error Code" + ErrCode);
                        MessageBox.Show("Error Code " + ErrCode + ".\n 프로그램을 중단합니다.");
                        
                        return;
                    }
                }

            }
            catch (Exception)
            {
                MessageBox.Show("xmlFile이 로드되지 않았습니다.", "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
        }


        private void Recorder_Click(object sender, RoutedEventArgs e)
        {
            pwdBox pwdBox = new pwdBox();
            pwdBox.ShowDialog();
            
            if (pwdBox.valid == 1)
            {
                Recorder recorder = new Recorder(contextController);
                recorder.ShowDialog();
            }
            else
            {
                return;
            }
            
        }

        private void btnLoadXml_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;


            if (openFileDialog.ShowDialog() == false)
                return;

            /*foreach (string filename in openFileDialog.FileNames)
            {
                WorkflowXmlListView.Items.Add(filename);
            }*/
        }

        private void BtnStartExcelWork_Click(object sender, RoutedEventArgs e)
        {
            // do stm
            if (MessageBox.Show("Excel 작업을 시작하시겠습니까?", "Question", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
            {
            }
            else
            {
                ExcelTemplateView.SelectAll();
                string templatePath = (string)ExcelTemplateView.SelectedItem;
                if (templatePath == null)
                {
                    MessageBox.Show("탬플릿 파일이 선택되지 않았습니다.", "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (ExcelWorkQueue.Count == 0)
                {
                    MessageBox.Show("회사 파일이 선택되지 않았습니다.", "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                foreach (ExcelWorkQueueDataStruct dataStruct in ExcelWorkQueue)
                {
                    if (dataStruct.jObject == null)
                    {
                        MessageBox.Show("회사가 선택되지 않았거나,\n중복된 파일을 선택했습니다.", "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                }

                int c = 1;
                List<string> workFailList = new List<string>();

                foreach (ExcelWorkQueueDataStruct dataStruct in ExcelWorkQueue)
                {

                    string extension = System.IO.Path.GetExtension(templatePath);
                    string savePath = System.IO.Path.GetFullPath(dataStruct.PathInfo);
                    string filename = System.IO.Path.GetFileNameWithoutExtension(dataStruct.PathInfo);

                    StringBuilder temp = new StringBuilder(savePath.Remove(savePath.LastIndexOf("\\"), savePath.Length - savePath.LastIndexOf("\\")));
                    temp.Append("\\작업");
                    DirectoryInfo di = new DirectoryInfo(temp.ToString());
                    if (di.Exists == false) di.Create();
                    temp.Append("\\" + filename + "_수정" + extension);

                    savePath = temp.ToString();
                    Console.WriteLine(savePath);

                    WriteDebugLine("작업중입니다.. (" + c + "/" + ExcelWorkQueue.Count + ")");

                    string ErrorCode = "";
                    ErrorCode = excelActivity.Work(dataStruct.PathInfo, templatePath, savePath, dataStruct.jObject);

                    if (ErrorCode != null)
                    {
                        //MessageBox.Show(ErrorCode, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                        WriteDebugLine(ErrorCode);
                        WriteDebugLine("작업이 비정상적으로 종료되었습니다.\n");
                        c++;
                        workFailList.Add(filename);
                        ExcelWorkFailView.Items.Add(savePath);
                        UpdateWindow();
                        continue;
                    }
                    else
                    {
                        ExcelWorkEndView.Items.Add(savePath);

                        WriteDebugLine("작업이 끝났습니다. (" + c + "/" + ExcelWorkQueue.Count + ")\n");
                        c++;
                    }
                }

                ClearAllCurrentQueueData(0);
                string temps = "";
                if (workFailList.Count == 0)
                    temps = "없음";
                foreach(string s in workFailList)
                {
                    temps += s;
                    temps += '\n';
                }
                MessageBox.Show("작업이 끝났습니다.\n" +
                    "작업 실패 리스트 : " +
                    temps , "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                WriteDebugLine("================================\n");
            }
        }

        private void ClearAllCurrentQueueData()
        {
            ExcelWorkEndView.Items.Clear();
            ExcelWorkQueue.Clear();
            ExcelTemplateView.Items.Clear();
            ExcelListView.Items.Clear();
        }
        private void ClearAllCurrentQueueData(int i)
        {
            if(i == 0)
            {
                ExcelWorkQueue.Clear();
                ExcelTemplateView.Items.Clear();
                ExcelListView.Items.Clear();
            }
        }

        private void BtnOpenTemplateFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel 97-2003 (*.xls)|*.xls|Excel (*.xlsx)|*.xlsx";


            if (openFileDialog.ShowDialog() == false)
                return;

            ExcelTemplateView.Items.Clear();

            foreach (string filename in openFileDialog.FileNames)
            {
                ExcelTemplateView.Items.Add(filename);
            }
        }

        private void btn_MappingTable_Click(object sender, RoutedEventArgs e)
        {
            MappingTable mappingTable = new MappingTable();
            try
            {
                mappingTable.ShowDialog();
            }
            catch (System.InvalidOperationException)
            {
                MessageBox.Show("더블클릭은 안됩니다");
                mappingTable.Close();
            } //더블클릭 Exception 방지
        }

        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls";
            openFileDialog.Multiselect = true;


            if (openFileDialog.ShowDialog() == false)
                return;
            ///


            List<string> clientNames = new List<string>();
            
            try
            {
                using (StreamReader file = new StreamReader(path, Encoding.GetEncoding("UTF-8")))
                using (JsonTextReader reader = new JsonTextReader(file))
                {
                    JObject object1 = (JObject)JToken.ReadFrom(reader);
                    List<string> clientList = object1.Properties().Select(p => p.Name).ToList();

                    foreach (string clientName in clientList)
                    {
                        clientNames.Add(clientName);
                    }

                    reader.Close();
                }

            }
            catch (System.IO.FileNotFoundException)
            {
                return;
            }


            foreach (string filename in openFileDialog.FileNames)
            {
                if (ExcelListView.Items.Contains(filename))
                {
                    continue;
                }
                ExcelWorkQueueDataStruct dataStructObj = new ExcelWorkQueueDataStruct(filename, clientNames);
                
                ExcelListView.Items.Add(dataStructObj);
                ExcelWorkQueue.Add(dataStructObj);

            }
        }

        class ExcelWorkQueueDataStruct
        {
            public string PathInfo { get; set; }
            public ObservableCollection<ComboBoxItem> cbItems { get; set; }
            public JObject jObject { get; set; }
            public ExcelWorkQueueDataStruct(string path, List<string> clientNames)
            {
                PathInfo = path;
                jObject = null;
                cbItems = new ObservableCollection<ComboBoxItem>();

                foreach (string s in clientNames)
                {
                    cbItems.Add(new ComboBoxItem { Content = s });
                }
            }

            
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {


            foreach(ExcelWorkQueueDataStruct d in ExcelWorkQueue)
            {
                if (d.PathInfo == (string)((ComboBox)sender).Tag)
                {
                    d.jObject = GetJObj((string)((ComboBoxItem)((ComboBox)sender).SelectedItem).Content);
                    break;
                }
            }
        }

        private JObject GetJObj(string key)
        {
            
            using (StreamReader file = new StreamReader(path, Encoding.GetEncoding("UTF-8")))
            using (JsonTextReader reader = new JsonTextReader(file))
            {
                JObject object1 = (JObject)JToken.ReadFrom(reader);

                JObject elem = JObject.Parse(object1.SelectToken(key).ToString());
                
                reader.Close();
                return elem;
            }
        }

        private void InitExcelContext()
        {
            ExcelWorkQueue = new List<ExcelWorkQueueDataStruct>();
        }

        private void ComboBox_Loaded(object sender, RoutedEventArgs e)
        {
            foreach (ExcelWorkQueueDataStruct d in ExcelWorkQueue)
            {
                if (d.PathInfo == (string)((ComboBox)sender).Tag)
                {
                    string filename = System.IO.Path.GetFileNameWithoutExtension(d.PathInfo);
                    filename = filename.Replace(" ", "");
                    Console.WriteLine(filename);

                    foreach(ComboBoxItem cbi in d.cbItems)
                    {
                        string comp = (string)cbi.Content;
                        int l, c=0;
                        if (filename.Length > comp.Length) l = comp.Length;
                        else l = filename.Length;

                        for (int i = 0; i < l; i++)
                            if (filename[i] == comp[i]) c++;
                        
                        if(c >= l/2)
                        {
                            ((ComboBox)sender).SelectedItem = cbi;
                            break;
                        }
                    }
                }
            }
        }
    }
}
