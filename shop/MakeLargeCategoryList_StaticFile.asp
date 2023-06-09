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
<!--#include file="./LargeCategoryList/LargeCategoryList.inc"-->
<%
'========================================================================
'
'	��J�e�S���[�ꗗ�p �ÓI�� html�t�@�C������
'
'�X�V����
' 2012/03/13 GV #1224 �V�K�쐬
'
'========================================================================
On Error Resume Next

Const THIS_PAGE_NAME = "MakeLargeCategoryList_StaticFile.asp"

Dim Connection
Dim wMsg						' �G���[���b�Z�[�W (�{�y�[�W�ō쐬���郁�b�Z�[�W)

Dim wStatus
Dim wCreateFileCount
Dim wCreateFile

'========================================================================

wMsg = ""
wCreateFileCount = 0
wCreateFile = ""

Call connect_db()

Call main()

If Err.Description <> "" Then
	wMsg = Err.Description
End If

Call close_db()

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
'	Function	Connect database
'
'========================================================================
Function connect_db()

'--- Connect database
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

End function

'========================================================================
'
'	Function	Main
'
'========================================================================
Function main()

Dim vRS
Dim vSQL
Dim vLargeCategoryCode
Dim vLargeCategoryName
Dim vFilePath


'--- ��J�e�S���[ ���o��
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.��J�e�S���[�R�[�h "
vSQL = vSQL & "    , a.��J�e�S���[�� "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    ��J�e�S���[ a WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "    a.Web��J�e�S���[�t���O = 'Y' "
vSQL = vSQL & "ORDER BY "
vSQL = vSQL & "    a.��J�e�S���[�R�[�h "

'@@@@@@@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, Connection, adOpenStatic

If vRS.EOF = True Then
	vRS.Close
	Set vRS = Nothing
	wMsg = "��J�e�S���[�̏�񂪁A�f�[�^�x�[�X�ɑ��݂��܂���B"
	Exit Function
End If

wCreateFileCount = 0

Do Until vRS.EOF

	'--- ��J�e�S���[�R�[�h
	vLargeCategoryCode = vRS("��J�e�S���[�R�[�h")
	vLargeCategoryName = vRS("��J�e�S���[��")

	'--- ��J�e�S���[ �ÓI���pHTML�e�L�X�g�t�@�C���쐬
	If fMakeLargeCategoryStaticHTMLFile(vLargeCategoryCode, vFilePath, wMsg) = False Then
		Exit Function
	End If

	' �쐬�����ƃt�@�C������Ҕ�
	wCreateFileCount = wCreateFileCount + 1
	wCreateFile = wCreateFile & vLargeCategoryName & "(" & vLargeCategoryCode & ") = " & vFilePath & vbNewLine

	vRS.MoveNext

Loop

vRS.Close
Set vRS = Nothing

End Function

'========================================================================
'
'	Function	Close database
'
'========================================================================
Function close_db()

Connection.Close
Set Connection = Nothing

End function

'========================================================================
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html lang="ja">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
<meta http-equiv="Content-Style-Type" content="text/css">
<meta http-equiv="Content-Script-Type" content="text/javascript">
<title>��J�e�S���[�ꗗ�p �ÓI���\��HTML�e�L�X�g�t�@�C������</title>
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
<h1>��J�e�S���[�ꗗ�p �ÓI���\��HTML�e�L�X�g�t�@�C���̐���</h1>
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