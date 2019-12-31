using System;
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

namespace tempproj
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        private MainController MainControllerObject;
        private ContextController contextController;
        
        public MainWindow()
        {
            InitializeComponent();
            InitAnnoucement();
            InitContext();
        }

        public void WriteDebugLine(string text)
        {
            DebugConsoleBlock.Dispatcher.BeginInvoke(new Action(() =>
            {
                DebugConsoleBlock.Text += text + Environment.NewLine;
            }));
        }
        private void InitContext()
        {
            contextController = new ContextController();
        }

        private void InitAnnoucement()
        {
            AnnouncementTextBlock.Text = @"<프로그램 사용법>
이 프로그램은 관리자 권한으로 실행되어야 합니다.
1. OpenFile 버튼을 눌러 작업할 엑셀 파일을 지정하세요.
2. 상단 리스트 박스에 파일 경로가 업로드 되었다면 파일이 지정된 것입니다.
3. 작업을 진행하기 전에 엑셀 파일을 닫아주어야 합니다.
";
        }


        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls";
            openFileDialog.Multiselect = true;


            if (openFileDialog.ShowDialog() == false)
                return;

            foreach(string filename in openFileDialog.FileNames)
            {
                ExcelListView.Items.Add(filename);
                
            }
        }
          
        private void ClearListBox_Click(object sender, RoutedEventArgs e)
        {
            ExcelListView.Items.Clear();
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
                WorkflowXmlListView.SelectAll();
                contextController.SetWorkflowXmlPath((string)WorkflowXmlListView.SelectedItem);

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

                    Thread.Sleep(1000);
                }

            }
            catch (Exception e)
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

            foreach (string filename in openFileDialog.FileNames)
            {
                WorkflowXmlListView.Items.Add(filename);
            }
        }

        private void BtnStartExcelWork_Click(object sender, RoutedEventArgs e)
        {
            // do stm
            if (MessageBox.Show("Excel 작업을 시작하시겠습니까?", "Question", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
            {
            }
            else
            {
                WriteDebugLine("Job done");
            }
        }
    }
}
