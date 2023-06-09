<%@ LANGUAGE="VBScript" %>
<%
'ネットハウスねっとハウスネットはうす
'サウンドハウス
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
'	大カテゴリー一覧用 静的化 htmlファイル生成
'
'更新履歴
' 2012/03/13 GV #1224 新規作成
'
'========================================================================
On Error Resume Next

Const THIS_PAGE_NAME = "MakeLargeCategoryList_StaticFile.asp"

Dim Connection
Dim wMsg						' エラーメッセージ (本ページで作成するメッセージ)

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

'--- 処理結果を出力
If wMsg = "" Then
	wStatus = "正常終了"
	wMsg = "作成件数 : " & wCreateFileCount & "件 <br>"
Else
	wStatus = "異常発生"
	wMsg = "エラー内容 = " & wMsg
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


'--- 大カテゴリー 取り出し
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.大カテゴリーコード "
vSQL = vSQL & "    , a.大カテゴリー名 "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    大カテゴリー a WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "    a.Web大カテゴリーフラグ = 'Y' "
vSQL = vSQL & "ORDER BY "
vSQL = vSQL & "    a.大カテゴリーコード "

'@@@@@@@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, Connection, adOpenStatic

If vRS.EOF = True Then
	vRS.Close
	Set vRS = Nothing
	wMsg = "大カテゴリーの情報が、データベースに存在しません。"
	Exit Function
End If

wCreateFileCount = 0

Do Until vRS.EOF

	'--- 大カテゴリーコード
	vLargeCategoryCode = vRS("大カテゴリーコード")
	vLargeCategoryName = vRS("大カテゴリー名")

	'--- 大カテゴリー 静的化用HTMLテキストファイル作成
	If fMakeLargeCategoryStaticHTMLFile(vLargeCategoryCode, vFilePath, wMsg) = False Then
		Exit Function
	End If

	' 作成件数とファイル情報を待避
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
<title>大カテゴリー一覧用 静的化表示HTMLテキストファイル生成</title>
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
<h1>大カテゴリー一覧用 静的化表示HTMLテキストファイルの生成</h1>
<table>
	<tr>
		<td class='label'>
		処理結果
		</td>
		<td>
		<% = wStatus %>
		</td>
	</tr>
	<tr>
		<td class='label'>
		メッセージ
		</td>
		<td>
		<% = wMsg %>
		</td>
	</tr>
<% If Len(wCreateFile) > 0 Then %>
	<tr>
		<td class='label'>
		作成したファイル
		</td>
		<td>
		<% = Replace(wCreateFile, vbNewLine, "<br>") %>
		</td>
	</tr>
<% End If %>
</table>
</body>
</html>