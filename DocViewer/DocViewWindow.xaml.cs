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
using System.Windows.Xps.Packaging;

namespace DocViewer
{
    /// <summary>
    /// DocViewWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class DocViewWindow : Window
    {
        public DocViewWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            String xpsFileName = @".\out\out.xps";

            // docxをxpsに変換（Office Word使用)
            System.Diagnostics.Process p =
                System.Diagnostics.Process.Start("Doc2XpsConverter.exe");
            p.WaitForExit();

            try
            {
                // xpsを表示
                XpsDocument xd = new XpsDocument(xpsFileName, System.IO.FileAccess.Read);
                FixedDocumentSequence fds = xd.GetFixedDocumentSequence();
                xd.Close();
                this.documentViewer.Document = fds;
            }
            catch (Exception)
            {
                MessageBox.Show("Wordファイルの表示に失敗しました。",
                "エラー",
                MessageBoxButton.OK, MessageBoxImage.Error);
                this.Close();
                return;
            }
        }
    }
}
