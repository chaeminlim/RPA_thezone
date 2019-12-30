using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using tempproj.Controller;

namespace tempproj
{
    /// <summary>
    /// Recorder.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Recorder : Window
    {
        private RecorderController recorderController;
        private ContextController contextController;

        public Recorder(ContextController contextController)
        {
            InitializeComponent();
            this.contextController = contextController;
            contextController.SetRecorder(this);
            recorderController = new RecorderController(contextController, this);
            recorderController.Install();   
        }

        private void RecordStart_Click(object sender, RoutedEventArgs e)
        {
            recorderController.Start();
        }

        private void RecordStop_Click(object sender, RoutedEventArgs e)
        {
            recorderController.Stop();
        }

        private void Try_Click(object sender, RoutedEventArgs e)
        {
            recorderController.StartRecorded();
        }

        private void RecordClear_Click(object sender, RoutedEventArgs e)
        {
            recorderController.Clear();
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            string[] lines = contextController.RecorderXmlList.ToArray();

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "XML files(.xml)|*.xml|all Files(*.*)|*.*";

            if (saveFileDialog.ShowDialog() == false)
                return;


            using (StreamWriter outputFile = new StreamWriter(saveFileDialog.FileName))
            {
                foreach (string line in lines)
                {
                    outputFile.WriteLine(line);
                }
            }
        }
    }
}
