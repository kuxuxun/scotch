#Scotch
==========================

Scotch POIを使用 Wrapperです & BDD ライブラリです。
エクセルによる帳票出力などを行う場合に使用する事を想定して作成しています。

1. POIを簡単に(OOP的に)使用できるように
2. 出力ファイルをBDD的にテストする 


# Sample
ワークブック作成
```java
   ScWorkbook wb = new ScWorkbook();
   ScSheet s = wb.createSheet("test");
   ScCell c3 = s.getCellAt("C3");
```

ワークブックのテスト
```java
new シート(s).のセル("B10").のスタイル().の横位置が右詰め();

new シート(TestExpected.getFirstSheetOf("compareTarget.xls")).のセル("H18").から("F21").の文字列が("テスト");
```



# TODO
* 全然できてない、、、足りない機能などなど実装


