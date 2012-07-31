#Scotch
==========================

Scotch POIを使用 Wrapperです & BDD ライブラリです。
エクセルによる帳票出力などを行う場合に使用する事を想定して作成しています。

1. POIを簡単に(OOP的に)使用できるように
2. 出力ファイルをBDD的にテストする 

# Sample
セル取得
```java
   ScWorkbook wb = new ScWorkbook();
   ScSheet s = wb.createSheet("test");
   ScCell c3 = s.getCellAt("C3");

```

ワークブックのテスト
```java
// セルのスタイルの確認
new シート(s).のセル("B10").のスタイル().の横位置が右詰め();

// セルの値の確認
new シート(TestExpected.getFirstSheetOf("compareTarget.xls")).のセル("H18").から("F21").の文字列が("テスト");

// フォントの確認
new シート(TestExpected.getFirstSheetOf("expected.xls")).のセル("C10").のフォントは("ＭＳ Ｐゴシック");

// 罫線の確認
new シート(TestExpected.getFirstSheetOf("compareTarget.xls")).のセル("H6").から("I8").のそれぞれののセルが罫線で囲まれている();

// etc...

```

# TODO
* いろいろ足りない機能などなど実装、、、(列の幅のテストやらなんやら)


