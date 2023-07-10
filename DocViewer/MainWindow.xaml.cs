using Microsoft.Win32;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Xps.Packaging;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Media;
using System.Text;
using System.Data;
using System.Collections.Generic;
using Microsoft.VisualBasic.FileIO;

namespace DocViewer
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        // スライドタブ選択時の左パネル表示用オブジェクト
        private StackPanel[] slidePanels;
        private Border[] slideBorders;
        private TextBlock[] slideNoBlocks;
        private TextBlock[] slideTitleBlocks;
        private TextBlock[] slideTextBlocks;

        // 論文タブ選択時の左パネル表示用オブジェクト
        private StackPanel[] docSentencePanels;
        private TextBlock[] docSentenceBlocks;

        // 右パネル表示用オブジェクト
        private TextBlock[] reviewBlocks;
        private TextBlock[] explainBlocks;
        private StackPanel[] simSlidePanels;

        // 論文タブ選択時の右パネル表示用オブジェクト
        private Border[] simSlideBorders;
        private TextBlock[] simSlideNoBlocks;
        private TextBlock[] simSlideTitleBlocks;
        private TextBlock[] simSlideTextBlocks;

        // スライドタイトル・テキスト配列（後で構造体にする）
        private String[] titles;
        private String[] texts;

        // 順方向検討要素・逆方向検討要素リスト
        List<List<String>> reviewInfos = new List<List<String>>();
        List<List<String>> reverseReviewInfos = new List<List<String>>();

        String rowXmlFileName = @".\out\out.xml";
        String loadXmlFileName = @".\out\load.xml";
        String pptFileName;

        // Boolean loadSlideFinish = false;
        Boolean estimateReviewFinish = false;
        Double reviewThreshold = 0.2; // 検討要素判定の閾値

        public MainWindow()
        {
            // ウィンドウが読み込まれたら
            Loaded += (s, e) =>
            {
                // 古い出力ファルダの削除
                if (System.IO.Directory.Exists("./out") == true)
                    System.IO.Directory.Delete("./out", true);
            };
        }

        private void loadSlideMenu_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog();
            dialog.Title = "ファイルを開く";
            dialog.Filter = "プレゼンテーション (*.pptx)|*.pptx|全てのファイル (*.*)|*.*";

            if (dialog.ShowDialog() == true)
            {
                pptFileName = dialog.FileName;

                System.Diagnostics.Process p;
                p = System.Diagnostics.Process.Start("java", "-jar ppt2doc.jar \"" + dialog.FileName + "\"");
                p.WaitForExit();
                System.IO.File.Delete(loadXmlFileName);
                System.IO.File.Copy(rowXmlFileName, loadXmlFileName);
                loadXml();
                updateSlideOutlinePanel(-1);
            }
        }

        private void estimateReviewMenu_Click(object sender, RoutedEventArgs e)
        {
            reviewInfos.Clear();
            reverseReviewInfos.Clear();

            var dialog = new OpenFileDialog();
            dialog.Title = "ファイルを開く";
            dialog.Filter = "Word 文書 (*.docx)|*.docx|全てのファイル (*.*)|*.*";

            if (dialog.ShowDialog() == true)
            {
                System.Diagnostics.Process p;
                p = System.Diagnostics.Process.Start("java", "-jar ppt2doc.jar \"" + loadXmlFileName + "\" \"" + dialog.FileName + "\"");
                p.WaitForExit();

                // 検討要素CSVファイルの読み込み
                try
                {
                    TextFieldParser parser1 = new TextFieldParser("./out/review1.csv", Encoding.GetEncoding("Shift_JIS"));
                    parser1.TextFieldType = FieldType.Delimited;
                    parser1.SetDelimiters(",");// ","区切り

                    while (parser1.EndOfData == false)
                    {
                        List<String> reviewInfo = new List<String>();
                        reviewInfo.AddRange(parser1.ReadFields());
                        reviewInfos.Add(reviewInfo);
                    }

                    TextFieldParser parser2 = new TextFieldParser("./out/review2.csv", Encoding.GetEncoding("Shift_JIS"));
                    parser2.TextFieldType = FieldType.Delimited;
                    parser2.SetDelimiters(",");// ","区切り

                    while (parser2.EndOfData == false)
                    {
                        List<String> reverseReviewInfo = new List<String>();
                        reverseReviewInfo.AddRange(parser2.ReadFields());
                        reverseReviewInfos.Add(reverseReviewInfo);
                    }

                    estimateReviewFinish = true;
                    updateSlideOutlinePanel(-1);
                    updateDocOutlinePanel(-1);
                }
                catch (FileNotFoundException)
                {
                    MessageBox.Show("CSVファイルの読み込みに失敗しました。",
                    "エラー",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }
        }
        
        private void loadXml() {
            XmlReader xr = new XmlReader();
            if (xr.readXml(loadXmlFileName) == true)
            {
                titles = xr.getTitles();
                texts = xr.getTexts();
            }
            else
            {
                MessageBox.Show("XMLファイルの読み込みに失敗しました。",
                "エラー",
                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /*
         * 左パネル（スライドタブ選択時）の表示
         */
        private void updateSlideOutlinePanel(int selectedNo)
        {
            slideOutlinePanel.Children.Clear();
            detailsPanel.Children.Clear();

            slidePanels = new StackPanel[titles.Length];
            slideBorders = new Border[titles.Length];
            slideNoBlocks = new TextBlock[titles.Length];
            slideTitleBlocks = new TextBlock[titles.Length];
            slideTextBlocks = new TextBlock[titles.Length];

            for (int i = 0; i < titles.Length; i++)
            {
                // スライド用パネル
                slidePanels[i] = new StackPanel();
                slidePanels[i].Name = "slidePanels" + i.ToString();
                slidePanels[i].MouseLeftButtonUp += new MouseButtonEventHandler(slidePanels_LeftButtonUp);

                // 罫線
                slideBorders[i] = new Border();
                slideBorders[i].Margin = new Thickness(10);
                slideBorders[i].Padding = new Thickness(10);
                slideBorders[i].BorderThickness = new Thickness(1);
                slideBorders[i].BorderBrush = Brushes.Black;
                slideBorders[i].Child = slidePanels[i];

                // スライド番号
                slideNoBlocks[i] = new TextBlock();
                slideNoBlocks[i].Name = "slideNoBlock" + i.ToString();
                slideNoBlocks[i].Margin = new Thickness(5);
                
                // スライドパネル被選択時
                if (i == selectedNo)
                {
                    slideBorders[i].Background = Brushes.LightGray;
                    updateDocOutlinePanel(-1);
                    updateSlideDetailsPanel(selectedNo);
                }
                if (estimateReviewFinish == true)
                {
                    // スライドに検討要素がある場合
                    Boolean reviewExsists = false;
                    foreach (List<String> reviewInfo in reviewInfos)
                    {
                        if (int.Parse(reviewInfo[0]) == i && double.Parse(reviewInfo[3]) < reviewThreshold)
                        {
                            reviewExsists = true;
                            break;
                        } 
                    }

                    if (reviewExsists == true)
                    {
                        slideNoBlocks[i].Text = (i + 1).ToString() + "（検討要素が見つかりました）";
                    }
                    else
                    {
                        slideNoBlocks[i].Text = (i + 1).ToString();
                    }
                }
                else
                {
                    slideNoBlocks[i].Text = (i + 1).ToString();
                }
                
                slidePanels[i].Children.Add(slideNoBlocks[i]);

                // タイトル
                if (titles[i] != null)
                {
                    slideTitleBlocks[i] = new TextBlock();
                    slideTitleBlocks[i].Name = "slideTitleBlock" + i.ToString();
                    slideTitleBlocks[i].Margin = new Thickness(5);
                    slideTitleBlocks[i].FontSize = 14;
                    slideTitleBlocks[i].FontWeight = FontWeights.Bold;
                    slideTitleBlocks[i].TextWrapping = TextWrapping.Wrap;
                    slideTitleBlocks[i].Text = titles[i];
                    slidePanels[i].Children.Add(slideTitleBlocks[i]);
                }

                // テキスト
                if (titles[i] != null)
                {
                    slideTextBlocks[i] = new TextBlock();
                    slideTextBlocks[i].Name = "slideTextBlock" + i.ToString();
                    slideTextBlocks[i].Margin = new Thickness(5);
                    slideTextBlocks[i].TextWrapping = TextWrapping.Wrap;
                    slideTextBlocks[i].Text = texts[i];
                    slidePanels[i].Children.Add(slideTextBlocks[i]);
                }

                slideOutlinePanel.Children.Add(slideBorders[i]);
            }
        }

        /*
         * 左パネル（論文タブ選択時）の表示
         */
        private void updateDocOutlinePanel(int selectedNo) {
            docOutlinePanel.Children.Clear();
            detailsPanel.Children.Clear();

            docSentencePanels = new StackPanel[reverseReviewInfos.Count];
            docSentenceBlocks = new TextBlock[reverseReviewInfos.Count];

            for (int i = 0; i < reverseReviewInfos.Count; i++)
            {
                // 論文センテンスパネル
                docSentencePanels[i] = new StackPanel();
                docSentencePanels[i].Name = "docSentencePanels" + i.ToString();
                docSentencePanels[i].MouseLeftButtonUp += new MouseButtonEventHandler(docSentencePanels_LeftButtonUp);
                
                // テキスト
                if (reverseReviewInfos[i][0] != null)
                {
                    docSentenceBlocks[i] = new TextBlock();
                    docSentenceBlocks[i].Name = "docSentenceBlock" + i.ToString();
                    docSentenceBlocks[i].Margin = new Thickness(5);
                    docSentenceBlocks[i].TextWrapping = TextWrapping.Wrap;
                    docSentenceBlocks[i].Text = reverseReviewInfos[i][0];
                    docSentencePanels[i].Children.Add(docSentenceBlocks[i]);
                }

                // スライドに検討要素がある場合
                if (estimateReviewFinish == true && double.Parse(reverseReviewInfos[i][3]) < reviewThreshold)
                {
                    docSentenceBlocks[i].Foreground = Brushes.Red;
                }

                // 論文センテンスパネル被選択時
                if (i == selectedNo)
                {
                    docSentencePanels[i].Background = Brushes.LightGray;
                    updateSlideOutlinePanel(-1);
                    updateDocDetailsPanel(selectedNo);
                }

                docOutlinePanel.Children.Add(docSentencePanels[i]);
            }
        }

        /*
         * 右パネル（スライドタブ選択時）の表示
         */
        private void updateSlideDetailsPanel(int selectedNo)
        {
            detailsPanel.Children.Clear();

            StackPanel testPanel = new StackPanel();

            // 検討要素のタイトル
            TextBlock reviewTitleBlock = new TextBlock();
            reviewTitleBlock = new TextBlock();
            reviewTitleBlock.Margin = new Thickness(10);
            reviewTitleBlock.FontSize = 14;
            reviewTitleBlock.FontWeight = FontWeights.Bold;
            reviewTitleBlock.TextWrapping = TextWrapping.Wrap;
            reviewTitleBlock.Text = "検討要素";
            testPanel.Children.Add(reviewTitleBlock);

            // 検討要素が存在したかのフラグ
            Boolean reviewExsists = false;

            reviewBlocks = new TextBlock[reviewInfos.Count];
            for (int i = 0; i < reviewInfos.Count; i++)
            {
                if (int.Parse(reviewInfos[i][0]) == selectedNo)
                {
                    reviewBlocks[i] = new TextBlock();
                    reviewBlocks[i].Margin = new Thickness(10);
                    reviewBlocks[i].TextWrapping = TextWrapping.Wrap;
                    testPanel.Children.Add(reviewBlocks[i]);

                    Run slideSentenceRun = new Run();
                    slideSentenceRun.Text = reviewInfos[i][1];
                    slideSentenceRun.FontWeight = FontWeights.Bold;

                    if (double.Parse(reviewInfos[i][3]) < reviewThreshold)
                    {
                        reviewBlocks[i].Text = "- 選択されたスライド内の文「";
                        reviewBlocks[i].Inlines.Add(slideSentenceRun);
                        reviewBlocks[i].Inlines.Add("」の内容は，論文に反映されてない可能性があります（類似度: " + Double.Parse(reviewInfos[i][3]).ToString("F3") + "）");
                        reviewBlocks[i].Foreground = Brushes.Red;
                    } else {
                        reviewBlocks[i].Text = "- 選択されたスライド内の文「";
                        reviewBlocks[i].Inlines.Add(slideSentenceRun);
                        reviewBlocks[i].Inlines.Add("」の内容は，論文に反映されています（類似度: " + Double.Parse(reviewInfos[i][3]).ToString("F3") + "）");
                    }

                    reviewExsists = true;
                }
            }

            if (estimateReviewFinish == false)
            {
                TextBlock reviewBlock = new TextBlock();
                reviewBlock = new TextBlock();
                reviewBlock.Margin = new Thickness(10);
                reviewBlock.TextWrapping = TextWrapping.Wrap;
                reviewBlock.Text = "検討要素推定が行われていません。";
                testPanel.Children.Add(reviewBlock);
            }
            else if (reviewExsists == false)
            {
                TextBlock reviewBlock = new TextBlock();
                reviewBlock = new TextBlock();
                reviewBlock.Margin = new Thickness(10);
                reviewBlock.TextWrapping = TextWrapping.Wrap;
                reviewBlock.Text = "検討要素は見つかりませんでした。";
                testPanel.Children.Add(reviewBlock);
            }

            detailsPanel.Children.Add(testPanel);
        }

        /*
         * 右パネル（論文タブ選択時）の表示
         */
        private void updateDocDetailsPanel(int selectedNo)
        {
            detailsPanel.Children.Clear();

            StackPanel testPanel = new StackPanel();

            // 検討要素のタイトル
            TextBlock reviewTitleBlock = new TextBlock();
            reviewTitleBlock = new TextBlock();
            reviewTitleBlock.Margin = new Thickness(10);
            reviewTitleBlock.FontSize = 14;
            reviewTitleBlock.FontWeight = FontWeights.Bold;
            reviewTitleBlock.TextWrapping = TextWrapping.Wrap;
            reviewTitleBlock.Text = "対応関係・検討要素";
            testPanel.Children.Add(reviewTitleBlock);

            Boolean reviewExsists = false;
            simSlidePanels = new StackPanel[reverseReviewInfos.Count];
            explainBlocks = new TextBlock[reverseReviewInfos.Count];
            simSlideBorders = new Border[reverseReviewInfos.Count];
            simSlideNoBlocks = new TextBlock[reverseReviewInfos.Count];
            simSlideTitleBlocks = new TextBlock[reverseReviewInfos.Count];
            simSlideTextBlocks = new TextBlock[reverseReviewInfos.Count];
            reviewBlocks = new TextBlock[reverseReviewInfos.Count];
            for (int i = 0; i < reverseReviewInfos.Count; i++)
            {
                if (i == selectedNo)
                {
                    XmlReader xr = new XmlReader();
                    xr.readXml(loadXmlFileName);

                    explainBlocks[i] = new TextBlock();
                    explainBlocks[i].Margin = new Thickness(10);
                    explainBlocks[i].TextWrapping = TextWrapping.Wrap;

                    // スライド用パネル
                    simSlidePanels[i] = new StackPanel();
                    simSlidePanels[i].Name = "simSlidePanels" + i.ToString();

                    // 罫線
                    simSlideBorders[i] = new Border();
                    simSlideBorders[i].Margin = new Thickness(10, 0, 10, 0);
                    simSlideBorders[i].Padding = new Thickness(10);
                    simSlideBorders[i].BorderThickness = new Thickness(1);
                    simSlideBorders[i].BorderBrush = Brushes.Black;
                    simSlideBorders[i].Child = simSlidePanels[i];

                    // スライド番号
                    simSlideNoBlocks[i] = new TextBlock();
                    simSlideNoBlocks[i].Name = "simSlideNoBlock" + i.ToString();
                    simSlideNoBlocks[i].Margin = new Thickness(5);
                    simSlideNoBlocks[i].Text = (int.Parse(reverseReviewInfos[i][1]) + 1).ToString();
                    simSlidePanels[i].Children.Add(simSlideNoBlocks[i]);

                    // タイトル
                    if (xr.getTitles()[int.Parse(reverseReviewInfos[i][1])] != null)
                    {
                        simSlideTitleBlocks[i] = new TextBlock();
                        simSlideTitleBlocks[i].Name = "simSlideTitleBlock" + i.ToString();
                        simSlideTitleBlocks[i].Margin = new Thickness(5);
                        simSlideTitleBlocks[i].FontSize = 14;
                        simSlideTitleBlocks[i].FontWeight = FontWeights.Bold;
                        simSlideTitleBlocks[i].TextWrapping = TextWrapping.Wrap;
                        simSlideTitleBlocks[i].Text = xr.getTitles()[int.Parse(reverseReviewInfos[i][1])];
                        simSlidePanels[i].Children.Add(simSlideTitleBlocks[i]);
                    }

                    if (xr.getTexts()[int.Parse(reverseReviewInfos[i][1])] != null)
                    {
                        simSlideTextBlocks[i] = new TextBlock();
                        simSlideTextBlocks[i].Name = "simSlideTextBlock" + i.ToString();
                        simSlideTextBlocks[i].Margin = new Thickness(5);
                        simSlideTextBlocks[i].TextWrapping = TextWrapping.Wrap;
                        simSlideTextBlocks[i].Text = xr.getTexts()[int.Parse(reverseReviewInfos[i][1])];
                        simSlidePanels[i].Children.Add(simSlideTextBlocks[i]);
                    }

                    Run slideSentenceRun = new Run();
                    slideSentenceRun.Text = reverseReviewInfos[i][2];
                    slideSentenceRun.FontWeight = FontWeights.Bold;

                    reviewBlocks[i] = new TextBlock();
                    reviewBlocks[i].Margin = new Thickness(10);
                    reviewBlocks[i].TextWrapping = TextWrapping.Wrap;
                    reviewBlocks[i].Text = "一番似ている文は，「";
                    reviewBlocks[i].Inlines.Add(slideSentenceRun);
                    reviewBlocks[i].Inlines.Add("」でした（類似度： " + Double.Parse(reverseReviewInfos[i][3]).ToString("F3") + "）");

                    if (double.Parse(reverseReviewInfos[i][3]) < reviewThreshold)
                    {
                        explainBlocks[i].Text = "選択された文の内容は、スライド資料内に存在しない可能性があります。";
                        testPanel.Children.Add(explainBlocks[i]);
                    }
                    else {
                        explainBlocks[i].Text = "選択された文は、以下のスライドと対応関係がある可能性があります。";
                        testPanel.Children.Add(explainBlocks[i]);
                        testPanel.Children.Add(simSlideBorders[i]);
                    }
                    testPanel.Children.Add(reviewBlocks[i]);

                    reviewExsists = true;
                }
            }

            if (estimateReviewFinish == false) {
                TextBlock reviewBlock = new TextBlock();
                reviewBlock = new TextBlock();
                reviewBlock.Margin = new Thickness(10);
                reviewBlock.TextWrapping = TextWrapping.Wrap;
                reviewBlock.Text = "検討要素推定が行われていません。";
                testPanel.Children.Add(reviewBlock);
            }
            else if (reviewExsists == false)
            {
                TextBlock reviewBlock = new TextBlock();
                reviewBlock = new TextBlock();
                reviewBlock.Margin = new Thickness(10);
                reviewBlock.TextWrapping = TextWrapping.Wrap;
                reviewBlock.Text = "検討要素は見つかりませんでした。";
                testPanel.Children.Add(reviewBlock);
            }

            detailsPanel.Children.Add(testPanel);
        }

        // メニュー〔論文下地生成〕のクリックイベント
        private void combineBySimMenu_Click(object sender, RoutedEventArgs e)
        {
            SliderWindow sw = new SliderWindow();
            sw.Owner = this;
            sw.Show();
        }

        // メニュー〔論文下地表示〕のクリックイベント
        private void viewDocMenu_Click(object sender, RoutedEventArgs e)
        {
            DocViewWindow dvw = new DocViewWindow();
            dvw.Show();
        }

        // 重要度計算プログラムの呼び出し
        public void callAnalyzer(double threshold)
        {
            System.Diagnostics.Process p;
            p = System.Diagnostics.Process.Start("java", "-jar ppt2doc.jar " + threshold + " \"" + loadXmlFileName + "\"");
            p.WaitForExit();
            System.IO.File.Delete(loadXmlFileName);
            System.IO.File.Copy(rowXmlFileName, loadXmlFileName);
            loadXml();
            updateSlideOutlinePanel(-1);
            updateDocOutlinePanel(-1);
        }

        // スライドパネルのクリック（マウス左ボタン押下）イベント
        private void slidePanels_LeftButtonUp(object sender, RoutedEventArgs e)
        {
            var slidePanel = (StackPanel)sender;
            int no = int.Parse(Regex.Replace(slidePanel.Name, @"[^0-9]", ""));
            updateSlideOutlinePanel(no);
        }

        // 論文センテンスパネルのクリック（マウス左ボタン押下）イベント
        private void docSentencePanels_LeftButtonUp(object sender, RoutedEventArgs e)
        {
            var docSentencePanel = (StackPanel)sender;
            int no = int.Parse(Regex.Replace(docSentencePanel.Name, @"[^0-9]", ""));
            updateDocOutlinePanel(no);
        }

        // タブ選択イベント
        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            detailsPanel.Children.Clear();
        }
    }
}
