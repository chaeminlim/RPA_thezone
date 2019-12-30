using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;

namespace tempproj
{
    class WorkflowController
    {
        private XmlController xmlController;
        private RecorderController recorderController;

        public WorkflowController()
        {
            xmlController = new XmlController();
            recorderController = new RecorderController();
        }
        
        public void ExecuteProgram()
        {

        }

        public int DoActionXml(String strXmlData)
        {
            int num, tempNum;
            AutomationElement ae;
            (num, ae) = xmlController.XmlFinder(strXmlData);
            
            switch (num)
            {
                case 0:
                    (tempNum, ae) = HandleNotFound(strXmlData);
                    break;
                case 1:
                    tempNum = 1;
                    break;
                default:
                    tempNum = 2;
                    break;

            }
            
            if (tempNum == 0 || tempNum >= 2)
            {
                return tempNum;
            }

            recorderController.WindowControl(strXmlData, ae);
            recorderController.DoAction(strXmlData, ae);

            return 1;
        }

        private (int, AutomationElement) HandleNotFound(string strXmlData)
        {
            int num;
            AutomationElement ae;

            for (int i = 10; i > 0; i--)
            {
                (num, ae) = xmlController.XmlFinder(strXmlData);

                if (num == 0)
                {
                    Thread.Sleep(5000);
                }
                else if (num == 1)
                {
                    return (1, ae);
                }
                else
                {
                    return (2, null);
                }
            }

            return (0, null);
        }
    }
}
