VbaDoxygen
==========

このスクリプトはExcelVBAのDoxygen文章を作成します。

使い方
------
このスクリプトを使用するにはマクロのセキュリティで [Visual Basic プロジェクトへのアクセスを信頼する]にチェックを付与する必要があります。

CScript ExcelVBADoxygen.vbs "C:\dev\VbaDoxygen\Sample.xlsm" "C:\dev\VbaDoxygen\output" "C:\Program Files\doxygen\bin\doxygen.exe"

64bitOSで32bitExcelを操作する場合
C:\Windows\SysWOW64\CScript.exe ExcelVBADoxygen.vbs "C:\dev\VbaDoxygen\Sample.xlsm" "C:\dev\VbaDoxygen\output" "C:\Program Files\doxygen\bin\doxygen.exe"

第一引数：Excelのパス
第二引数：出力フォルダ
第三引数：doxygen.exeへのフルパス

ファイルの説明
------
Sample.xlsm  テスト出力用のエクセルファイルです


vbfilter.exe 
VBFilterをExe化したものです。いくつか異なるものが存在しますが、今回は下記をExe化しました
「だらろぐ」様　vbfilter.pyを改造してみた
http://r-satsuki.air-nifty.com/blog/2008/02/vbfilter_61f1.html

python27.dll
w9xpopen.exe
vbfilter.exeを動かすのに必要です。

doxyfile_template 
doxygenファイルの元になるものです。
