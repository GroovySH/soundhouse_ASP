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
'	プレゼント応募
'
'	更新履歴
'2008/05/12 改行コードインジェクション対策（i_toパラメータ削除）
'2008/05/13 クロスサイトリクエストフォジェリー対策 Keyパラメータチェック
'2009/04/30 エラー時にerror.aspへ移動
'2011/03/02 hn SetSecureKeyの位置変更
'2011/08/01 an #1087 Error.aspログ出力対応
'
'========================================================================

On Error Resume Next

Dim userID
Dim msg

Dim customer_nm
Dim furigana
Dim e_mail
Dim zip
Dim prefecture
Dim address
Dim telephone

DIm Skey

Dim Connection
Dim RS_customer

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
	w_msg = "<p class='error'>ログインを行ってください｡</p>"
else
	call connect_db()
	call main()
	
	'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
	if Err.Description <> "" then
		wErrDesc = "PresentOubo.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
		call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
	end if                                           '2011/08/01 an add e

	call close_db()
end if

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

if w_msg <> "" then
	Session("msg") = w_msg
	Response.Redirect g_HTTPS & "shop/Login.asp?called_from=present"
end if

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
'	Function	main proc
'
'========================================================================
'
Function main()

'---- セキュリティーキーセット 
Skey = SetSecureKey()

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
w_sql = w_sql & "   AND c.住所連番 = 1"
w_sql = w_sql & "   AND c.電話連番 = 1"

'@@@@@@response.write(w_sql)

Set RS_customer = Server.CreateObject("ADODB.Recordset")
RS_customer.Open w_sql, Connection, adOpenStatic

if RS_customer.EOF = true then
	w_msg = "<p class='error'>顧客情報がありません。</p>"
else
	customer_nm = RS_customer("顧客名")
	furigana = RS_customer("顧客フリガナ")
	e_mail = RS_customer("顧客E_mail1")
	zip = RS_customer("顧客郵便番号")
	prefecture = RS_customer("顧客都道府県")
	address = RS_customer("顧客住所")
	telephone = RS_customer("顧客電話番号")
end if

RS_customer.close

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
<title>プレゼント応募｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->

<body>

<!--#include file="../Navi/NaviTop.inc"-->

<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>
  
  <!-- コンテンツstart -->
  <div id="globalContents">
    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="<%=g_HTTP%>"><img src="<%=g_RelLink%>images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li class="now">プレゼント応募</li>
      </ul>
    </div></div></div>

    <h1 class="title">プレゼント応募</h1>
    
<form name="f_data" action="PresentOuboSend.asp" method="post">
    
<table class="form">
  <tr>
    <th>お名前</th>
    <td><%=customer_nm%></td>
  </tr>
  <tr>
    <th>フリガナ</th>
    <td><%=furigana%></td>
  </tr>
  <tr>
    <th>郵便番号</th>
    <td><%=zip%></td>
  </tr>
  <tr>
    <th>住所</th>
    <td><%=prefecture%><%=address%></td>
  </tr>
  <tr>
    <th>メールアドレス</th>
    <td><%=e_mail%></td>
  </tr>
  <tr>
    <th>サウンドハウス購入歴</th>
    <td>
    	<input type="radio" id="0" name="purchase" value="初めて"><label for="0">初めて</label>
		<input type="radio" id="1_2" name="purchase" value="1～2回"><label for="1_2">1～2回</label>
		<input type="radio" id="3_9" name="purchase" value="3～9回"><label for="3_9">3～9回</label>
        <input type="radio" id="10" name="purchase" value="10回以上"><label for="10">10回以上</label>
    </td>
  </tr>
  <tr>
    <th>コメント</th>
    <td><textarea name="comment" cols="70" rows="5"></textarea></td>
  </tr>
</table>

<p class="btnBox"><input type="submit" value="送信" class="opover"></p>

<input type="hidden" name="customer_nm" value="<%=customer_nm%>">
<input type="hidden" name="furigana" value="<%=furigana%>">
<input type="hidden" name="zip" value="<%=zip%>">
<input type="hidden" name="prefecture" value="<%=prefecture%>">
<input type="hidden" name="address" value="<%=address%>">
<input type="hidden" name="telephone" value="<%=telephone%>">
<input type="hidden" name="e_mail" value="<%=e_mail%>">
<input type="hidden" name="Skey" value="<%=Skey%>">
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