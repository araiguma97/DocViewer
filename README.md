# DocViewer

「DocViewer」は、以下のクラスで構成されるC#（.NET Framework）プログラムです。「Ppt2Doc」の結果をGUI上で表示します。各クラスの詳細は、各クラスのコードやコメントを参照してください。

- MainWindowクラス、SliderWindowクラス、DocViewWindowクラス：それぞれ、メインウィンドウ、スライド統合時の類似度指定用スライダーウィンドウ、論文下地表示ウィンドウのクラスです。
- XmlReaderクラス、XmlWriterクラス：「Ppt2Doc」内の同名のクラスと同様です。
- Doc2XpsConverter.exe：「out.docx」を「out.xps」に変換するC#（.NET Framework）プログラムです。コードは、『[C#でWordファイルをPDFに変換する](https://blog.jhashimoto.net/entry/20120604/1338801745)』にあるコードを書き換えたものです。
