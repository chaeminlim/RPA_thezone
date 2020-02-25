using System;
using System.Collections.Generic;
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

namespace tempproj
{
    /// <summary>
    /// pwdBox.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class pwdBox : Window
    {
        public int valid;
        public pwdBox()
        {
            InitializeComponent();
        }

        private void btnEnter_Click(object sender, RoutedEventArgs e)
        {
            if (pwBox.Password == "1111")
            {
                valid = 1;
            }
            else
                valid = 0;

            this.Close();
        }
    }
}
