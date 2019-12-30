using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace tempproj.Controller
{
    public class ContextController
    {
        private List<String> ExcelPathList;
        private string WorkflowXmlPath = "";
        private Queue<String> RecorderXmlQueue;
        public List<String> RecorderXmlList;
        private Recorder recorder;

        public ContextController()
        {
            ExcelPathList = new List<String>();
            RecorderXmlQueue = new Queue<String>();
            RecorderXmlList = new List<String>();
        }

        public void SetRecorder(Recorder recorder)
        {
            this.recorder = recorder;
        }

        private void UpdateRecorderListView()
        {
            recorder.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate
            {
                recorder.RecorderListView.Items.Clear();
                foreach (string item in RecorderXmlList)
                {
                    recorder.RecorderListView.Items.Add(item);
                }
            }));
        }
        public void AddExcelPath(string path)
        {
            if(ExcelPathList.Contains(path))
                return;
            else
                ExcelPathList.Add(path);    
        }

        public void AddRecorderXmlQueue(String xmlline)
        {
            RecorderXmlQueue.Enqueue(xmlline);
            RecorderXmlList.Add(xmlline);
            UpdateRecorderListView();
        }
        public void ClearRecorderXmlQueue()
        {
            RecorderXmlQueue.Clear();
            RecorderXmlList.Clear();
            UpdateRecorderListView();
        }
        public int CountRecorderXmlQueue()
        {
            return RecorderXmlQueue.Count;
        }
        public string DequeueRecorderXmlQueue()
        {
            return RecorderXmlQueue.Dequeue();
        }
        public void ClearExcelPath()
        {
            ExcelPathList.Clear();
        }

        public void SetWorkflowXmlPath(string path)
        {
            WorkflowXmlPath = path;
        }

        public string GetWorkflowXmlPath()
        {
            return WorkflowXmlPath;
        }

        

    }
}
