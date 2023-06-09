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
<!--#include file="../common/HttpsSecurity.inc"-->

<%
'========================================================================
'
'	カタログ請求
'
'更新履歴
'2008/05/14 HTTPSチェック対応
'2008/05/23 入力データチェック強化（LEFT他)
'2009/04/30 エラー時にerror.aspへ移動
'2011/08/01 an #1087 Error.aspログ出力対応
'
'========================================================================

On Error Resume Next

Dim userID
Dim msg

Dim message

Dim Connection
Dim RS

Dim w_sql
Dim w_html
Dim w_msg
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.buffer = true

'---- UserID 取り出し
userID = Session("userID")

'---- 呼び出し元からのデータ取り出し
message = ReplaceInput(Left(Request("message"), 100))

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "CatalogRequestStore.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'========================================================================
'
'	Function	Connect database
'
'========================================================================
'
Function connect_db()
Dim i

'---- Connect database
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

End function

'========================================================================
'
'	Function	Main
'
'========================================================================
'
Function main()

'---- カタログ請求情報登録
w_sql = ""
w_sql = w_sql & "SELECT *"
w_sql = w_sql & "  FROM カタログ請求"
w_sql = w_sql & " WHERE 1 = 2"
	  
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic, adLockOptimistic

'---- insert カタログ請求
RS.AddNew

RS("カタログ請求日") = now()
RS("顧客番号") = userID
RS("備考") = message
RS("最新フラグ") = "Y"

RS.Update

RS.close

message = Replace(message, vbNewLine, "<br>")

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
'
Function close_db()

Connection.close
Set Connection= Nothing    '2011/08/01 an add

End function

'========================================================================
%>

<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="Shift_JIS">
<title>カタログ請求｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->

</head>
<!--#include file="../Navi/NaviTop.inc"-->
<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>
  
  <!-- コンテンツstart -->
  <div id="globalContents">
    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="<%=g_HTTP%>"><img src="<%=g_RelLink%>images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li class="now">カタログ請求</li>
      </ul>
    </div></div></div>

    <h1 class="title">カタログ請求</h1>
    
    <p>ご依頼いただきましたカタログを送付いたします。<br>ありがとうございました。</p>

</div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>