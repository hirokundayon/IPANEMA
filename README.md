# IPANEMA
## 概要  
Microsoft ExcelでGoogle Chromeを自動実行するプログラムです。  
次の順番でGoogle Chromeを自動実行します。  
1. Webブラウザ起動
1. Googleで「OSC 名古屋 2012 レポート」を検索
1. OSC2012 Nagoya のレポートを表示させる  
[https://www.ospn.jp/press/20120601osc2012-nagoya-report.html](https://www.ospn.jp/press/20120601osc2012-nagoya-report.html)  
1. メモリーカードエラー写真をクリック  
1. さらにクリックして拡大表示  
1. Webブラウザ停止  

## 必要なもの
+ Microsoft Excel  
   * Micrsoft Office 2016で検証しました。  
   * 32bit版でも64bit版でも動作します。  
+ [ChromeDriver](https://sites.google.com/a/chromium.org/chromedriver/)  
    * Google Chromeの操作に使用します。  
+ [VBA-JSON](https://github.com/VBA-tools/VBA-JSON)  
    * JSONのパーズに使用します。  
+ [VBA-Dictionary](https://github.com/VBA-tools/VBA-Dictionary)  
    * VBA-JSONが使用します。  
