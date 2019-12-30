using System.Windows.Automation;
using System.Windows.Controls;
using System.Collections.Generic;
namespace tempproj
{
    class AutomationTreeElementWrapper
    {
        public List<AutomationTreeElementWrapper> ChildList
        {
            get; set;
        }
        
        public AutomationElement AE
        {
            get;
            private set;
        }
        public string Header
        {
            get; set;
        }


        public AutomationTreeElementWrapper(AutomationElement ae = null)
        {
            this.ChildList = new List<AutomationTreeElementWrapper>();
            this.AE = ae;

            if (ae.Current.Name != "")
                this.Header = ae.Current.Name;
            else
                this.Header = "No Name";
        }
        

        public void AddChild(AutomationTreeElementWrapper aew)
        {
            this.ChildList.Add(aew);
        }

    }
}
