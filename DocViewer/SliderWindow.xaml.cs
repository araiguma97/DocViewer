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

namespace DocViewer
{
    /// <summary>
    /// SliderWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class SliderWindow : Window
    {
        public double threshold;

        public SliderWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            MainWindow mw = (MainWindow)this.Owner;
            mw.callAnalyzer(slider.Value);
        }
    }
}
