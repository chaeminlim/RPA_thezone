using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Automation;
using System.Windows.Controls;

namespace tempproj
{
    class MainController
    {
        public List<Process> Processes { get; set; }
        private TreeView TreeView { get; set; }
        private ListView ListView { get; set; }
        private ListView ListView2 { get; set; }
        
        public MainController()
        {
        }


        public AutomationTreeElementWrapper GetRootInit()
        {
            Queue<AutomationElement> aeQueue = new Queue<AutomationElement>();
            Condition conditions = Condition.TrueCondition;

            AutomationElement root = AutomationElement.RootElement;
            AutomationTreeElementWrapper rootWrapper = new AutomationTreeElementWrapper(root);
            AutomationElementCollection aec =  root.FindAll(TreeScope.Children, conditions);
            
            foreach(AutomationElement ae in aec)
            {
                rootWrapper.AddChild(new AutomationTreeElementWrapper(ae));
            }
            #region get from process list
            /// get from process list
            ///
            /*
            Process[] allProcesses = Process.GetProcesses();
            foreach(Process proc in allProcesses)
            {
                Condition tempCondition = new PropertyCondition(AutomationElement.ProcessIdProperty, proc.Id);
                AutomationElement ae = AutomationElement.RootElement.FindFirst(TreeScope.Element | TreeScope.Children, tempCondition);
                if (ae != null)
                    aeQueue.Enqueue(ae);
            }
            */
            #endregion
            return rootWrapper;
        }

        public void MakeTree()
        {
            AutomationTreeElementWrapper rootWrapper =  GetRootInit();

            TreeWalker walker = TreeWalker.RawViewWalker;

            foreach(AutomationTreeElementWrapper treeElem in rootWrapper.ChildList)
            {
                try
                {
                    TraverseElement(walker, treeElem);
                }
                catch (Exception)
                {
                    continue;
                }
            }
        }

        private static void TraverseElement(TreeWalker walker, AutomationTreeElementWrapper automationElementWrapper)
        {
            Queue<AutomationTreeElementWrapper> elementQueue = new Queue<AutomationTreeElementWrapper>();

            elementQueue.Enqueue(automationElementWrapper);
            while(elementQueue.Count > 0)
            {
                AutomationTreeElementWrapper atew = elementQueue.Dequeue();
                AutomationElement child = walker.GetFirstChild(atew.AE);

                while (child != null)
                {
                    AutomationTreeElementWrapper caew = new AutomationTreeElementWrapper(child);
                    atew.AddChild(caew);
                    elementQueue.Enqueue(caew);
                    child = walker.GetNextSibling(child);
                }
            }

        }

        public void PrintSelected(TreeViewItem treeViewItem)
        {
            AutomationElement ae = (AutomationElement)(treeViewItem).Tag;

            
            try
            {
                Process p = Process.GetProcessById(ae.Current.ProcessId);

                ListView.Items.Clear();

                ListView.Items.Add("프로세스 이름: " + p.ProcessName);
                ListView.Items.Add("프로세스 모듈 이름: " + p.MainModule.ModuleName);
                ListView.Items.Add("파일 경로" + p.MainModule.FileName);

                ListView.Items.Add("요소명: " + ae.Current.Name);
                ListView.Items.Add("가속화 키: " + ae.Current.AcceleratorKey);
                ListView.Items.Add("액세스 키: " + ae.Current.AccessKey);
                ListView.Items.Add("자동화 요소 ID: " + ae.Current.AutomationId);
                ListView.Items.Add("클래스 이름: " + ae.Current.ClassName);
                ListView.Items.Add("컨트롤 유형: " + ae.Current.ControlType.ProgrammaticName);
                ListView.Items.Add("Framework ID: " + ae.Current.FrameworkId);
                ListView.Items.Add("포커스 소유 : " + ae.Current.HasKeyboardFocus);
                ListView.Items.Add("도움말: " + ae.Current.HelpText);
                ListView.Items.Add("컨텐츠 여부: " + ae.Current.IsContentElement);
                ListView.Items.Add("컨트롤 여부: " + ae.Current.IsControlElement);
                ListView.Items.Add("활성화 여부: " + ae.Current.IsEnabled);
                ListView.Items.Add("포커스 소유 가능 여부: " + ae.Current.IsKeyboardFocusable);
                ListView.Items.Add("화면 비표시 여부: " + ae.Current.IsOffscreen);
                ListView.Items.Add("내용 보화(패스워드) 여부: " + ae.Current.IsPassword);
                ListView.Items.Add("IsRequiredForForm: " + ae.Current.IsRequiredForForm);
                ListView.Items.Add("아이템 상태: " + ae.Current.ItemStatus);
                ListView.Items.Add("아이템 형식: " + ae.Current.ItemType);
                ListView.Items.Add("사각영역: " + ae.Current.BoundingRectangle);

            }
            catch (ElementNotAvailableException)
            {

            }
            SelectedItemController(ae);


        }

        private void SelectedItemController(AutomationElement SelectedItem)
        {
            AutomationPattern[] patterns = SelectedItem.GetSupportedPatterns();  //주어진 트리노드의 컨트롤유형을 배열형태로 저장


            ListView2.Items.Clear();  //listview2부분을 초기화한뒤
            foreach (AutomationPattern pattern in patterns)//각 컨트롤 유형에 대해서 반복문 실행
            {
                //    listView2.Items.Add("ProgrammaticName: " + pattern.ProgrammaticName);
                //    listView2.Items.Add("PatternName: " + Automation.PatternName(pattern));

                /*
                PatternControl pc = new PatternControl();  //컨트롤유형과 관련한 객체를 생성해서
                pc.PatternController(pattern, SelectedItem, ListView2);  //각 컨트롤유형에 따라 가능한 자동화형태를 listview2에 출력
                */
            }

        }
    }
}
