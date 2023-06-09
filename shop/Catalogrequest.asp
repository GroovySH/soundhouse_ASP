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
'	カタログ請求ページ
'
'更新履歴
'2008/05/14 HTTPSチェック対応
'2009/04/30 エラー時にerror.aspへ移動
'2011/08/01 an #1087 Error.aspログ出力対応
'
'========================================================================

On Error Resume Next

Dim userID
Dim msg

Dim customer_nm
Dim furigana
Dim customer_email
Dim zip
Dim prefecture
Dim address
Dim telephone

Dim Connection
Dim RS

Dim w_sql
Dim w_html
Dim w_msg
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.buffer = true

'---- 呼び出し元プログラムからのメッセージ取り出し
msg = Session.contents("msg")
Session("msg") = ""

'---- UserID 取り出し

userID = Session("userID")

if userID = "" then
	w_msg = "ログインを行ってください｡"
else
	'---- Execute main
	call connect_db()
	call main()
	
	'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
	if Err.Description <> "" then
		wErrDesc = "Catalogrequest.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
		call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
	end if                                           '2011/08/01 an add e

	call close_db()
end if

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

if w_msg <> "" then
	Response.Redirect "../shop/Login.asp?called_from=catalog"
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

'---- 顧客情報取り出し
w_sql = ""
w_sql = w_sql & "SELECT a.顧客名"
w_sql = w_sql & "       , a.顧客フリガナ"
w_sql = w_sql & "       , a.顧客E_mail1"
w_sql = w_sql & "       , b.顧客郵便番号"
w_sql = w_sql & "       , b.顧客都道府県"
w_sql = w_sql & "       , b.顧客住所"
w_sql = w_sql & "       , c.顧客電話番号"
w_sql = w_sql & "  FROM Web顧客 a WITH (NOLOCK)"
w_sql = w_sql & "     , Web顧客住所 b WITH (NOLOCK)"
w_sql = w_sql & "     , Web顧客住所電話番号 c WITH (NOLOCK)"
w_sql = w_sql & " WHERE a.顧客番号 = " & userID
w_sql = w_sql & "   AND b.顧客番号 = a.顧客番号"
w_sql = w_sql & "   AND b.住所連番 = 1"
w_sql = w_sql & "   AND c.顧客番号 = a.顧客番号"
w_sql = w_sql & "   AND c.電話連番 = 1"
	  
'@@@@@@response.write(w_sql)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic

if RS.EOF = true then
	w_msg = "<p class='error'>顧客情報がありません。</p>"
	Session("msg") = w_msg
else
	customer_nm = RS("顧客名")
	furigana = RS("顧客フリガナ")
	customer_email = RS("顧客E_mail1")
	zip = RS("顧客郵便番号")
	prefecture = RS("顧客都道府県")
	address = RS("顧客住所")
	telephone = RS("顧客電話番号")
end if

RS.close

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
    <p>各種人気商品が満載されている『HOT MENU』カラーカタログを無料で発送しております。<br>※カタログはメール便での発送となるため、お届けまで1週間前後かかる場合がございます。</p>

<form action="CatalogRequestStore.asp" method="post">    
<table class="form">
  <tr>
    <th>お名前</th>
    <td><%=customer_nm%><input type="hidden" name="customer_nm" value="<%=customer_nm%>"></td>
  </tr>
  <tr>
    <th>フリガナ</th>
    <td><%=furigana%><input type="hidden" name="furigana" value="<%=furigana%>"></td>
  </tr>
  <tr>
    <th>メールアドレス</th>
    <td><%=customer_email%><input type="hidden" name="e_mail" value="<%=customer_email%>"></td>
  </tr>
  <tr>
    <th>郵便番号</th>
    <td><%=zip%><input type="hidden" name="zip" value="<%=zip%>"></td>
  </tr>
  <tr>
    <th>住所</th>
    <td><%=prefecture%><%=address%><input type="hidden" name="address" value="<%=prefecture%> <%=address%>"></td>
  </tr>
  <tr>
    <th>電話番号</th>
    <td><%=telephone%><input type="hidden" name="telephone" value="<%=telephone%>"></td>
  </tr>
  <tr>
    <th>その他のカタログ</th>
    <td><input type="text" name="message" size="70" maxlength="100"><div>その他ご希望のカタログがある際は、ジャンル、モデル、メーカー等をご記入ください。（100文字以内）</div></td>
  </tr>
  <tr>
    <th>HOT MENU 希望</th>
    <td><input type="checkbox" id="i_HOTMENU" name="i_HOTMENU" value="Y" checked><label for="i_HOTMENU">希望する</label></td>
  </tr>
</table>
<p class="btnBox"><input type="submit" value="送信" class="opover"></p>
</form>


</div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>