#Scotch

Scotch POIを使用 Wrapperです & BDD ライブラリです。
エクセルによる帳票出力などを行う場合に使用する事を想定して作成しています。

1. POIを簡単に(OOP的に)使用できるように
2. 出力ファイルをBDD的にテストする 


# sample
   ScWorkbook tw = new ScWorkbook();
   ScSheet t = tw.createSheet("test");
   ScCell c3 = e.getCellAt("C3");

# TODO
* 全然できてない、、、足りない機能などなど実装

