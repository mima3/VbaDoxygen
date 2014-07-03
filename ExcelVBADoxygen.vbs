'���̃X�N���v�g��ExcelVBA��Doxygen���͂��쐬���܂�
'���̃X�N���v�g���g�p����ɂ̓}�N���̃Z�L�����e�B�� [Visual Basic �v���W�F�N�g�ւ̃A�N�Z�X��M������]�Ƀ`�F�b�N��t�^����K�v������܂�
'�������FExcel�̃p�X
'�������F�o�̓t�H���_
'��O�����Fdoxygen.exe�ւ̃t���p�X
'
' ��F
' �@ExcelVBADoxygen "C:\dev\VbaDoxygen\Sample.xlsm" "C:\dev\VbaDoxygen\output" "C:\Program Files\doxygen\bin\doxygen.exe"
' �t�@�C���̐���
' Sample.xlsm  �e�X�g�o�͗p�̃G�N�Z���t�@�C���ł�
' output       �e�X�g�o�͂̌��ʂ��i�[����Ă��܂�
' python27.dll
' w9xpopen.exe
' vbfilter.exe VBFilter��Exe���������̂ł��B�������قȂ���̂����݂��܂����A����͉��L��Exe�����܂���
' �u����낮�v�l�@vbfilter.py���������Ă݂�
'�@ http://r-satsuki.air-nifty.com/blog/2008/02/vbfilter_61f1.html
' doxyfile_template doxygen�t�@�C���̌��ɂȂ���̂ł��B
' 
' doxyfile_template �Œu�����s�Ȃ�����
' <@ProjectName>      �v���W�F�N�g��������킵�܂��B���̃X�N���v�g�ł�Excel�̃t�@�C�������w�肵�Ă܂��B
' <@InputPath>        Excel����o�͂���VBA�̃\�[�X���i�[���Ă���t�H���_���w�肵�܂�
' <@InputFilterPath>  vbfilter.exe�̃t���p�X���w�肳��܂�
' <@OutputDirectory>  �o�͐�̃t�H���_���w�肳��܂�
Option Explicit
Dim fileSrc		' Excel
Dim dirDst		' Doxygen�̕ۑ���̃f�B���N�g��
Dim dirDstSrc	' �G�N�X�|�[�g�����t�@�C�����i�[����f�B���N�g��
Dim fso
Dim doxytemplatePath
Dim doxy
Dim doxyPath
Dim doxyBin
Dim shell
Set shell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
doxyBin = WScript.Arguments(2)

' Excel�̎擾.	
fileSrc = WScript.Arguments(0)

' �ۑ���̃f�B���N�g�����擾
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

'* Excel����VBA�̃R�[�h�𒊏o����
'* @param[in] fileSrc Excel�t�@�C���̃p�X
'* @param[in] dirDst  �\�[�X�R�[�h���o�͂����
'*
Private Sub ExportExcelVBA(Byval fileSrc, Byval dirDst)
	Dim fso			' FileSystemObject
	Dim fo			' �o�̓t�@�C��
	Dim xl			' Excel�I�u�W�F�N�g
	Dim wbk			' ���[�N�u�b�N
	Dim cmp			' VBProject.VBComponents
	

	Dim sFormat		' �g���q
	Dim fileDst 	' �ۑ���̃t�@�C��

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
'* �e�L�X�g�t�@�C���̍쐬
'* @param[in] sFile �o�͐�̃p�X
'* @param[in] sData �o�͓��e
'*
Private Sub WriteFile( Byval sFile , Byval sData )
	'�萔�̐錾
	Const ForWriting = 2 '�������݁i�㏑�����[�h�j

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
'* �e�L�X�g�t�@�C���̓ǂݍ���
'* @param[in] �t�@�C���̃p�X
'* @return �t�@�C���̓��e
'*
Private Function ReadFile( Byval sFile )
	'�萔�̐錾
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
