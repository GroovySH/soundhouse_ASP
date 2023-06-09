<%@ LANGUAGE="VBScript" %>
<%
'�l�b�g�n�E�X�˂��ƃn�E�X�l�b�g�͂���
'�T�E���h�n�E�X
 Option Explicit
%>
<!--#include file="../common/ADOVBS.inc"-->
<!--#include file="../common/system_common.inc"-->
<!--#include file="../common/shop_common_functions.inc"-->
<!--#include file="../common/bfunctions1.asp"-->
<!--#include file="./Static/ManualDownload.inc"-->

<%
'========================================================================
'
'	�}�j���A���_�E�����[�h�p �ÓIHTML�e�L�X�g�t�@�C������
'
'�X�V����
' 2012/02/10 GV #1233 �V�K�쐬
'
'========================================================================

On Error Resume Next

Const THIS_PAGE_NAME = "MakeManualDownload.asp"

Dim wMsg						' �G���[���b�Z�[�W (�{�y�[�W�ō쐬���郁�b�Z�[�W)

Dim wStatus
Dim wCreateFileCount
Dim wCreateFile

'========================================================================

wMsg = ""
wCreateFileCount = 0
wCreateFile = ""

Call main()

If Err.Description <> "" Then
	wMsg = Err.Description
End If

'--- �������ʂ��o��
If wMsg = "" Then
	wStatus = "����I��"
	wMsg = "�쐬���� : " & wCreateFileCount & "�� <br>"
Else
	wStatus = "�ُ픭��"
	wMsg = "�G���[���e = " & wMsg
End If

'========================================================================
'
'	Function	Main
'
'========================================================================
Function main()

Dim vFilePath

' �}�j���A���_�E�����[�h�p �ÓIHTML�e�L�X�g�t�@�C���쐬
If fMakeManualDownloadHTMLFile(vFilePath, wMsg) = False Then
	Exit Function
End If

' �쐬����
wCreateFileCount = 1

' �t�@�C������Ҕ�
wCreateFile = vFilePath

End Function

'========================================================================
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html lang="ja">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
<meta http-equiv="Content-Style-Type" content="text/css">
<meta http-equiv="Content-Script-Type" content="text/javascript">
<title>�}�j���A���_�E�����[�h�p �ÓIHTML�e�L�X�g�t�@�C���̐���</title>
<style type="text/css">
<!--
table {
	border-collapse: collapse;
}
td, th {
	border: 1px solid black;
	padding: 3px;
}
.label {
	background-color: #0033cc;
	color: #FFFFFF;
	text-align: center;
}
-->
</style>
</head>
<body>
<h1>�}�j���A���_�E�����[�h�p �ÓIHTML�e�L�X�g�t�@�C���̐���</h1>
<table>
	<tr>
		<td class='label'>
		��������
		</td>
		<td>
		<% = wStatus %>
		</td>
	</tr>
	<tr>
		<td class='label'>
		���b�Z�[�W
		</td>
		<td>
		<% = wMsg %>
		</td>
	</tr>
<% If Len(wCreateFile) > 0 Then %>
	<tr>
		<td class='label'>
		�쐬�����t�@�C��
		</td>
		<td>
		<% = Replace(wCreateFile, vbNewLine, "<br>") %>
		</td>
	</tr>
<% End If %>
</table>
</body>
</html>