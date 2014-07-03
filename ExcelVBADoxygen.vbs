'このスクリプトはExcelVBAのDoxygen文章を作成します
'このスクリプトを使用するにはマクロのセキュリティで [Visual Basic プロジェクトへのアクセスを信頼する]にチェックを付与する必要があります
'第一引数：Excelのパス
'第二引数：出力フォルダ
'第三引数：doxygen.exeへのフルパス
'
' 例：
' 　ExcelVBADoxygen "C:\dev\VbaDoxygen\Sample.xlsm" "C:\dev\VbaDoxygen\output" "C:\Program Files\doxygen\bin\doxygen.exe"
' ファイルの説明
' Sample.xlsm  テスト出力用のエクセルファイルです
' output       テスト出力の結果が格納されています
' python27.dll
' w9xpopen.exe
' vbfilter.exe VBFilterをExe化したものです。いくつか異なるものが存在しますが、今回は下記をExe化しました
' 「だらろぐ」様　vbfilter.pyを改造してみた
'　 http://r-satsuki.air-nifty.com/blog/2008/02/vbfilter_61f1.html
' doxyfile_template doxygenファイルの元になるものです。
' 
' doxyfile_template で置換を行なう文字
' <@ProjectName>      プロジェクト名をあらわします。このスクリプトではExcelのファイル名を指定してます。
' <@InputPath>        Excelから出力したVBAのソースを格納しているフォルダを指定します
' <@InputFilterPath>  vbfilter.exeのフルパスが指定されます
' <@OutputDirectory>  出力先のフォルダが指定されます
Option Explicit
Dim fileSrc		' Excel
Dim dirDst		' Doxygenの保存先のディレクトリ
Dim dirDstSrc	' エクスポートしたファイルを格納するディレクトリ
Dim fso
Dim doxytemplatePath
Dim doxy
Dim doxyPath
Dim doxyBin
Dim shell
Set shell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
doxyBin = WScript.Arguments(2)

' Excelの取得.	
fileSrc = WScript.Arguments(0)

' 保存先のディレクトリを取得
dirDst  = WScript.Arguments(1)
dirDstSrc = dirDst & "\src"

If Not fso.FolderExists(dirDstSrc) Then
	fso.CreateFolder( dirDstSrc )
End If

Call ExportExcelVBA(fileSrc, dirDstSrc)

doxytemplatePath = fso.GetFile(WScript.ScriptFullName).ParentFolder & "\doxyfile_template"
doxy = ReadFile(doxytemplatePath)

doxyPath = dirDst & "\doxygen"

'
doxy = Replace( doxy, "<@ProjectName>", fso.GetFileName(fileSrc) )
doxy = Replace( doxy, "<@InputPath>", dirDstSrc )
doxy = Replace( doxy, "<@InputFilterPath>",fso.GetFile(WScript.ScriptFullName).ParentFolder & "\vbfilter.exe")
doxy = Replace( doxy, "<@OutputDirectory>", dirDst )
Call WriteFile(doxyPath, doxy)
Call shell.Run( """" & doxyBin & """" & " """ & doxyPath &  """", 1, True )

Set fso = Nothing
Set shell = Nothing

'* ExcelからVBAのコードを抽出する
'* @param[in] fileSrc Excelファイルのパス
'* @param[in] dirDst  ソースコードを出力する先
'*
Private Sub ExportExcelVBA(Byval fileSrc, Byval dirDst)
	Dim fso			' FileSystemObject
	Dim fo			' 出力ファイル
	Dim xl			' Excelオブジェクト
	Dim wbk			' ワークブック
	Dim cmp			' VBProject.VBComponents
	

	Dim sFormat		' 拡張子
	Dim fileDst 	' 保存先のファイル

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set xl = CreateObject("Excel.Application")
	Set wbk = xl.Workbooks.Open(fileSrc)

	xl.DisplayAlerts = False

	On Error Resume Next

	For Each cmp In wbk.VBProject.VBComponents
		
		Select Case cmp.Type
			Case 1 
				sFormat = "bas"
			Case 2
				sFormat = "cls"
			case 100
				sFormat = "cls"
			Case 3
				sFormat = "frm"
			Case Else
				sFormat = "unkwon" & cmp.Type
		End Select
		If sFormat <> "" Then
			fileDst = dirDst + "\" + cmp.Name + "." + sFormat
			Set fo = fso.CreateTextFile(fileDst, True)
			
			If cmp.CodeModule.CountOfLines > 0 Then
				fo.WriteLine "Attribute VB_Name = """ & cmp.Name & """"
				fo.WriteLine cmp.CodeModule.Lines(1, cmp.CodeModule.CountOfLines)
			End If
			fo.WriteLine ""
			fo.Close
		End If
	Next

	wbk.Close
	Set wbk = Nothing
	xl.Quit
	Set xl = Nothing
	Set fo = Nothing
	Set fso = Nothing
End Sub

'*
'* テキストファイルの作成
'* @param[in] sFile 出力先のパス
'* @param[in] sData 出力内容
'*
Private Sub WriteFile( Byval sFile , Byval sData )
	'定数の宣言
	Const ForWriting = 2 '書きこみ（上書きモード）

	Dim objFileSys
	Dim objOutFile

	'
	Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")

	'
	Set objOutFile = objFileSys.OpenTextFile(sFile ,ForWriting,True)

	'
	objOutFile.Write sData

	'
	objOutFile.Close

	'
	Set objFileSys = Nothing
	Set objOutFile = Nothing

End Sub

'*
'* テキストファイルの読み込み
'* @param[in] ファイルのパス
'* @return ファイルの内容
'*
Private Function ReadFile( Byval sFile )
	'定数の宣言
	Const ForReading = 1
	
	Dim objFileSys
	Dim objFile

	'
	Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")

	'
	Set objFile = objFileSys.OpenTextFile(sFile ,ForReading,True)


    Do Until objFile.AtEndOfStream
        ReadFile = ReadFile & objFile.ReadLine & vbCrLf
    Loop
	'
	objFile.Close

	'
	Set objFileSys = Nothing
	Set objFile = Nothing	
End Function
