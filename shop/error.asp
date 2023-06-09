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

<%
'========================================================================
'
'	エラーページ
'
'更新履歴
'2010/01/26 an トップへリダイレクト
'2011/08/01 an #1087 ログファイルにエラーログ出力
'2012/07/13 if-web リニューアルレイアウト調整
'
'========================================================================

'2011/08/01 an add s
On Error Resume Next
Response.buffer = true

Dim ErrDesc
Dim Connection

'=======================================================================

'---- execute main
call connect_db()
call main()
call close_db()

'========================================================================
'
'	Function	Connect database
'
'========================================================================
'
Function connect_db()

'---- Connect database
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

End function

'========================================================================
'
'	Function	main
'
'========================================================================
Function main()

Dim FS
Dim FS_Log
Dim vLogFileName

'---- セッションデータからエラーメッセージ取得
ErrDesc =  fGetSessionData(gSessionID, "ErrDesc")
'---- エラーメッセージクリア
call fSetSessionData(gSessionID, "ErrDesc", "")

if ErrDesc <> "" then

	'---- Log File open
	Set FS = CreateObject("Scripting.FileSystemObject")
	vLogFileName = "../ErrorLog/ErrorLog" & Year(Date()) & Right("0" & Month(Date()), 2) & Right("0" & Day(Date()), 2) & ".log"
	vLogFileName = Server.MapPath(vLogFileName)               'Map log file
	Set FS_Log = FS.OpenTextFile(vLogFileName, 8, true)       'Log open - Append Mode

	'---- Logging
	FS_Log.WriteLine(cf_FormatTime(Now(), "HH:MM:SS") & " " & ErrDesc )

	FS_Log.Close
	
end if

End function
'2011/08/01 an add e

'========================================================================
'
'	Function	Close database
'
'========================================================================
'
Function close_db()

Connection.close
Set Connection= Nothing   '2011/08/01 an add

End function

'========================================================================
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<meta http-equiv="Refresh" content="0;URL=<%=g_HTTP%>">
<title>エラー｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->
<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>

<!-- コンテンツstart -->
<div id="globalContents">
<!--
  <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
    <p class="home"><a href="../"><img src="../images/icon_home.gif" alt="HOME"></a></p>
    <ul id="path">
      <li class="now">エラー</li>
    </ul>
  </div></div></div>
-->
  <h1 class="title">エラー</h1>

  <p>処理中にエラーが発生しました。<br>入力値に不正な文字が含まれています。</p>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>