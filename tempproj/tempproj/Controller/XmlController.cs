using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Automation.Provider;
using System.Windows.Documents;
using System.Xml;

namespace tempproj
{
    class XmlController
    {
        public static string MakeXmlFile(Stack<AutomationElement> automationElements, int type, string inputText)
        {
            XmlDocument tree = new XmlDocument();
            XmlElement root = tree.CreateElement("UI");
            tree.AppendChild(root);


            XmlElement Action = tree.CreateElement("Action");
            if (type == 0)
                Action.SetAttribute("Type", "NULL");
            else if (type == 1)
                Action.SetAttribute("Type", "Click");
            else if (type == 2)
                Action.SetAttribute("Type", "DBClick");
            else if (type == 3)
            {
                Action.SetAttribute("Type", "TextInput");
                Action.SetAttribute("Value", inputText);
            }
            root.AppendChild(Action);


            //
            XmlElement xmlElement = tree.CreateElement("Window");
            AutomationElement wae = automationElements.Pop();
            Process p = Process.GetProcessById(wae.Current.ProcessId);

            xmlElement.SetAttribute("App", p.MainModule.ModuleName);
            xmlElement.SetAttribute("Class", wae.Current.ClassName);

            root.AppendChild(xmlElement);
            //
            while (automationElements.Count > 0)
            {
                AutomationElement ae = automationElements.Pop();
                xmlElement = tree.CreateElement("Element");
                xmlElement.SetAttribute("Name", ae.Current.Name);
                xmlElement.SetAttribute("Class", ae.Current.ClassName);
                xmlElement.SetAttribute("ControlType", ae.Current.ControlType.ProgrammaticName);
                xmlElement.SetAttribute("AutomationId", ae.Current.AutomationId);
                root.AppendChild(xmlElement);
            }



            StringWriter sw = new StringWriter();
            XmlTextWriter tx = new XmlTextWriter(sw);
            tree.WriteTo(tx);

            string strData = sw.ToString();
            return strData;

        }
        public static Stack<AutomationElement> MakeStack(AutomationElement ae)
        {

            Stack<AutomationElement> automationElements = new Stack<AutomationElement>();  //계층구조를 저장할 스택선언
            TreeWalker walker = TreeWalker.RawViewWalker;  // treewalker 선언
            automationElements.Push(ae);  //해당노드를 스택에 먼저 push

            AutomationElement parent = walker.GetParent(ae);  //treewalker를 통해 바로위의 부모를 찾는다.
            while (walker.GetParent(parent) != null)  //최상위 부모를 찾을때까지 반복문 실행
            {
                automationElements.Push(parent);  //각 계층구조를 스택에 저장
                parent = walker.GetParent(parent);
            }

            return automationElements;
        }

        public (int, AutomationElement) XmlFinder(string xmlData)
        {
            XmlDocument xml = new XmlDocument();
            xml.LoadXml(xmlData);
            XmlElement root = xml.DocumentElement;
            XmlNodeList nodes = root.ChildNodes;
            Queue<XmlNode> xmlQueue = new Queue<XmlNode>();
            int depth = nodes.Count - 1; // action 정보 제거

            foreach (XmlNode node in nodes)
            {
                xmlQueue.Enqueue(node);
            }
            xmlQueue.Dequeue(); // action정보 제거

            Queue<AutomationElement> windowQueue = WindowXMlFinder(xmlQueue.Dequeue(), depth);
            if (windowQueue.Count == 0)
            {
                // 프로세스 정보가 뒤바뀌는 에러
            }
            Queue<AutomationElement> elemQueue1 = windowQueue;
            depth -= 1;

            while (xmlQueue.Count > 0)
            {
                XmlNode xmlNode = xmlQueue.Dequeue();
                // debug
                Console.WriteLine("loop start");
                foreach (var elem in elemQueue1)
                {
                    Console.WriteLine(elem.Current.Name);
                }
                Console.WriteLine("loop end");
                // debug end
                elemQueue1 = ElementXMlFinder(xmlNode, elemQueue1, depth);

                depth -= 1;

            }

            if (elemQueue1.Count != 1)
                return (elemQueue1.Count, null);
            else
                return (elemQueue1.Count, elemQueue1.Dequeue());
        }
        #region dummy
        //private AutomationElement FinderFromPoint(Queue<XmlNode> xmlQueue, int depth)
        //{

        //    XmlNode xmlNode = null;
        //    String Name = null; String Class = null;
        //    while(xmlQueue.Count > 0)
        //    {
        //        xmlNode = xmlQueue.Dequeue();
        //    }
        //    foreach(XmlAttribute xmlAttribute in xmlNode.Attributes)
        //    {
        //        if(xmlAttribute.Name == "Name")
        //        {
        //            Name = xmlAttribute.Value;
        //        }
        //        else if (xmlAttribute.Name == "Class")
        //        {
        //            Class = xmlAttribute.Value;
        //        }
        //    }

        //    AutomationElement ae_root = AutomationElement.RootElement;
        //    Rect rect = ae_root.Current.BoundingRectangle;
        //    Point point = new Point();
        //    AutomationElement ae1, ae2 = null, result = null;

        //    try
        //    {
        //        Parallel.For(0, (int)rect.Height/10,
        //       (x, loopState) =>
        //       {
        //           Console.WriteLine("Searching...");
        //           point.X = 10*x;
        //           for (int y = 0; y < rect.Width; y += 10)
        //           {
        //               point.Y = y;
        //               ae1 = AutomationElement.FromPoint(point);
        //               if ((ae1 != ae2) && (ae1 != null))
        //               {

        //                   ae2 = ae1;
        //                   if (ae1.Current.Name == Name && ae1.Current.ClassName == Class)
        //                   {
        //                       result = ae1;
        //                       loopState.Stop();

        //                   }
        //               }
        //           }
        //       });
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine(e);
        //    }


        //    if (result != null)
        //    {
        //        Console.WriteLine("탐색이 끝났습니다... 찾았습니다...");
        //        return result;
        //    }
        //    else
        //    {
        //        Console.WriteLine("탐색이 끝났습니다... 못찾았습니다...");
        //        return result;
        //    }
        //    //for (double x = 0; x < rect.Height; x += 5)
        //    //{
        //    //    Console.WriteLine("탐색중입니다...");
        //    //    point.X = x;
        //    //    for (double y = 0; y < rect.Width; y+= 5)
        //    //    {
        //    //        point.Y = y;
        //    //        ae1 = AutomationElement.FromPoint(point);
        //    //        if ((ae1 != ae2) && (ae1 != null))
        //    //        {

        //    //            ae2 = ae1;
        //    //            if (ae1.Current.Name == Name && ae1.Current.ClassName == Class)
        //    //            {
        //    //                Console.WriteLine("탐색이 끝났습니다... 찾았습니다...");
        //    //                return ae1;
        //    //            }
        //    //        }
        //    //    }
        //    //}

        //}
        #endregion

        #region enumchildwindows
        private delegate bool EnumWindowProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumChildWindows(IntPtr window, EnumWindowProc callback, IntPtr lParam);

        public List<IntPtr> GetAllChildHandles(IntPtr hWnd)
        {
            List<IntPtr> childHandles = new List<IntPtr>();

            GCHandle gcChildhandlesList = GCHandle.Alloc(childHandles);
            IntPtr pointerChildHandlesList = GCHandle.ToIntPtr(gcChildhandlesList);

            try
            {
                EnumWindowProc childProc = new EnumWindowProc(EnumWindow);
                EnumChildWindows(hWnd, childProc, pointerChildHandlesList);
            }
            finally
            {
                gcChildhandlesList.Free();
            }

            return childHandles;
        }

        private bool EnumWindow(IntPtr hWnd, IntPtr lParam)
        {
            GCHandle gcChildhandlesList = GCHandle.FromIntPtr(lParam);

            if (gcChildhandlesList == null || gcChildhandlesList.Target == null)
            {
                return false;
            }

            List<IntPtr> childHandles = gcChildhandlesList.Target as List<IntPtr>;
            childHandles.Add(hWnd);

            return true;
        }
        #endregion


        public Queue<AutomationElement> GetRootInit()
        {
            Queue<AutomationElement> aeQueue = new Queue<AutomationElement>();

            //System.Windows.Automation.Condition conditions = new PropertyCondition(AutomationElement.IsEnabledProperty, true);
            //Condition conditions = new PropertyCondition(AutomationElement.IsEnabledProperty, true);
            System.Windows.Automation.Condition conditions = System.Windows.Automation.Condition.TrueCondition;
            AutomationElement root = AutomationElement.RootElement;
            AutomationElementCollection aec = root.FindAll(TreeScope.Children, conditions);
            foreach (AutomationElement ae in aec)
            {
                aeQueue.Enqueue(ae);
            }
            /*
            Process[] allProcesses = Process.GetProcesses();
            foreach (Process proc in allProcesses)
            {
                System.Windows.Automation.Condition tempCondition = new PropertyCondition(AutomationElement.ProcessIdProperty, proc.Id);
                AutomationElement ae = AutomationElement.RootElement.FindFirst(TreeScope.Element | TreeScope.Children, tempCondition);
                if (ae != null)
                    aeQueue.Enqueue(ae);
            }
            */
            return aeQueue;
        }
        #region FindWindowEx

        //////////
        //[DllImport("user32.dll", EntryPoint = "FindWindowEx", CharSet = CharSet.Auto)]
        //static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        //static List<IntPtr> GetAllChildrenWindowHandles(IntPtr hParent,int maxCount)
        //{
        //    List<IntPtr> result = new List<IntPtr>();
        //    int ct = 0;
        //    IntPtr prevChild = IntPtr.Zero;
        //    IntPtr currChild = IntPtr.Zero;
        //    while (true && ct < maxCount)
        //    {
        //        currChild = FindWindowEx(hParent, prevChild, null, null);
        //        if (currChild == IntPtr.Zero) break;
        //        result.Add(currChild);
        //        prevChild = currChild;
        //        ++ct;
        //    }
        //    return result;
        //}
        //////////
        #endregion

        private Queue<AutomationElement> WindowXMlFinder(XmlNode windowNode, int depth)
        {

            Queue<AutomationElement> windowQueue = GetRootInit();
            Queue<AutomationElement> Filter1 = new Queue<AutomationElement>();
            Queue<AutomationElement> Filter2 = new Queue<AutomationElement>();

            foreach (XmlAttribute attribute in windowNode.Attributes)
            {

                if (attribute.Name == "App")
                {
                    while (windowQueue.Count > 0)
                    {
                        AutomationElement ae = windowQueue.Dequeue();
                        Process p = Process.GetProcessById(ae.Current.ProcessId);

                        try
                        {
                            if (p.MainModule.ModuleName == attribute.Value)
                            {
                                Filter1.Enqueue(ae);
                            }

                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e);
                        }
                    }
                }
                else if (attribute.Name == "Class")
                {
                    while (Filter1.Count > 0)
                    {
                        AutomationElement ae = Filter1.Dequeue();
                        if (ae.Current.ClassName == attribute.Value)
                        {
                            Filter2.Enqueue(ae);
                        }

                    }
                }
            }

            return FindChild(Filter2, depth);
        }

        private void travserse(AutomationElement ae)
        {
            AutomationElement child = TreeWalker.RawViewWalker.GetFirstChild(ae);
            if (child == null)
                Console.WriteLine("자식없음");
            while (child != null)
            {
                Console.WriteLine(child.Current.Name);
                travserse(child);
                child = TreeWalker.RawViewWalker.GetNextSibling(child);
            }
        }

        private Queue<AutomationElement> FindChild(Queue<AutomationElement> elemQueue, int depth)
        {
            if (depth == 1)
            {
                return elemQueue;
            }
            else
            {
                Queue<AutomationElement> returnQueue = new Queue<AutomationElement>();
                while (elemQueue.Count > 0)
                {
                    AutomationElement ae = elemQueue.Dequeue();
                    AutomationElement child = TreeWalker.RawViewWalker.GetFirstChild(ae);
                    if (child == null)
                    {
                        continue;
                    }
                    else
                    {
                        while (child != null)
                        {
                            returnQueue.Enqueue(child);
                            child = TreeWalker.RawViewWalker.GetNextSibling(child);
                        }
                    }
                }


                return returnQueue;
            }

        }
        private Queue<AutomationElement> ElementXMlFinder(XmlNode elemNode, Queue<AutomationElement> elementQueue, int depth)
        {
            Queue<AutomationElement> Filter1 = new Queue<AutomationElement>();
            Queue<AutomationElement> Filter2 = new Queue<AutomationElement>();
            Queue<AutomationElement> Filter3 = new Queue<AutomationElement>();
            Queue<AutomationElement> Filter4 = new Queue<AutomationElement>();

            foreach (XmlAttribute attribute in elemNode.Attributes)
            {
                if (attribute.Name == "Name")
                {
                    while (elementQueue.Count > 0)
                    {
                        AutomationElement ae = elementQueue.Dequeue();

                        if (ae.Current.Name == attribute.Value)
                        {
                            Filter1.Enqueue(ae);
                        }
                    }
                }
                else if (attribute.Name == "Class")
                {
                    while (Filter1.Count > 0)
                    {
                        AutomationElement ae = Filter1.Dequeue();
                        if (ae.Current.ClassName == attribute.Value)
                        {
                            Filter2.Enqueue(ae);
                        }
                    }
                }
                else if (attribute.Name == "ControlType")
                {
                    while (Filter2.Count > 0)
                    {
                        AutomationElement ae = Filter2.Dequeue();
                        if (ae.Current.ControlType.ProgrammaticName == attribute.Value)
                        {
                            Filter3.Enqueue(ae);
                        }
                    }
                }
                else if (attribute.Name == "AutomationId")
                {
                    while (Filter3.Count > 0)
                    {
                        AutomationElement ae = Filter3.Dequeue();
                        if (ae.Current.AutomationId == attribute.Value)
                        {
                            Filter4.Enqueue(ae);
                        }
                    }
                }

            }

            return FindChild(Filter4, depth);
        }
    }
}
