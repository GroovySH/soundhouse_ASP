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
<!--#include file="./Static/ManualDownload.inc"-->

<%
'========================================================================
'
'	マニュアルダウンロード用 静的HTMLテキストファイル生成
'
'更新履歴
' 2012/02/10 GV #1233 新規作成
'
'========================================================================

On Error Resume Next

Const THIS_PAGE_NAME = "MakeManualDownload.asp"

Dim wMsg						' エラーメッセージ (本ページで作成するメッセージ)

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
'	Function	Main
'
'========================================================================
Function main()

Dim vFilePath

' マニュアルダウンロード用 静的HTMLテキストファイル作成
If fMakeManualDownloadHTMLFile(vFilePath, wMsg) = False Then
	Exit Function
End If

' 作成件数
wCreateFileCount = 1

' ファイル情報を待避
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
<title>マニュアルダウンロード用 静的HTMLテキストファイルの生成</title>
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
<h1>マニュアルダウンロード用 静的HTMLテキストファイルの生成</h1>
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