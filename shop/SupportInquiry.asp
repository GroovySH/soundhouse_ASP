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
'	修理･商品サポートページ
'
'更新履歴
'2008/05/12 改行コードインジェクション対策（i_toパラメータ削除）
'2008/05/13 クロスサイトリクエストフォジェリー対策 Keyパラメータセット
'2009/04/30 エラー時にerror.aspへ移動
'2010/10/04 an リニューアル対応。依頼データをDBに登録するように変更
'2011/02/21 hn RtnURL使用時はg_HTTP/g_HTTPSを使用するように変更（PCIDSS)
'2011/03/02 hn SetSecureKeyの位置変更
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/06/29 if-web リニューアルレイアウト調整
'
'========================================================================

On Error Resume Next

Dim userID

Dim MakerName       '2010/10/4 an add
Dim ProductName     '2010/10/4 an add
Dim Warranty        '2010/10/4 an add
Dim SerialNo        '2010/10/4 an add
Dim WhenPurchased   '2010/10/4 an add
Dim Comment         '2010/10/4 an add

Dim wCustomerName
Dim wZip
Dim wPrefecture
Dim wAddress
Dim wTelephone
Dim wFax
Dim wEmail

Dim Skey

Dim Connection
Dim RS

Dim wSQL
Dim wHTML
Dim wMSG
Dim wNoData
Dim wErrDesc   '2011/08/01 an add

'========================================================================

'---- 呼び出し元プログラムからのエラーメッセージ取り出し  '2010/10/4 an add
wMSG = Session("msg")
Session("msg") = ""

'---- 顧客番号取り出し
userID = Session("userID")

'---- エラー時は入力データを受け取って再表示   '2010/10/4 an add
MakerName = ReplaceInput(Left(Request("MakerName"),25))
ProductName = ReplaceInput(Left(Request("ProductName"),50))
Warranty = ReplaceInput(Left(Request("Warranty"),2))
SerialNo = ReplaceInput(Left(Request("SerialNo"),40))
WhenPurchased = ReplaceInput(Left(Request("WhenPurchased"),10))
Comment = ReplaceInput(Left(Request("Comment"),500))

'---- ログインしていなければログイン画面へ
if userID = "" then
	Response.Redirect g_HTTPS & "shop/LoginCheck.asp?RtnURL=" & g_HTTPS & "shop/SupportInquiry.asp"	'2011/02/21 hn mod
end if

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "SupportInquiry.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" OR wNoData = "Y" then  'LoginCheckしているはずなので顧客情報が取得できなければエラー
	Response.Redirect g_HTTP & "shop/Error.asp"
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
'	Function	main
'
'				userIDがCookieにあれば会員情報を表示
'
'========================================================================
Function main()

'---- セキュリティーキーセット 
Skey = SetSecureKey()

wNoData = ""

'--------- select customer
wSQL = ""
wSQL = wSQL & "SELECT a.顧客番号"
wSQL = wSQL & "     , a.顧客名"
wSQL = wSQL & "     , a.顧客E_mail1"
wSQL = wSQL & "     , b.顧客郵便番号"
wSQL = wSQL & "     , b.顧客都道府県"
wSQL = wSQL & "     , b.顧客住所"
wSQL = wSQL & "     , c.顧客電話番号"
wSQL = wSQL & "  FROM Web顧客 a WITH (NOLOCK)"
wSQL = wSQL & "     , Web顧客住所 b WITH (NOLOCK)"
wSQL = wSQL & "     , Web顧客住所電話番号 c WITH (NOLOCK)"
wSQL = wSQL & " WHERE b.顧客番号 = a.顧客番号" 
wSQL = wSQL & "   AND c.顧客番号 = b.顧客番号" 
wSQL = wSQL & "   AND c.住所連番 = b.住所連番" 
wSQL = wSQL & "   AND b.住所連番 = 1" 
wSQL = wSQL & "   AND c.電話連番 = 1" 
wSQL = wSQL & "   AND a.顧客番号 = " & userID 
		
'@@@@@response.write(wSQL & "<BR>")

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

if RS.EOF = true then
	wNoData = "Y"
	exit function
else
	wCustomerName = RS("顧客名")
	wZip = RS("顧客郵便番号")
	wPrefecture = RS("顧客都道府県")
	wAddress = RS("顧客住所")
	wTelephone = RS("顧客電話番号")
	wEmail = RS("顧客E_mail1")
end if

RS.close

end function

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
<title>修理･商品サポート｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/inquiry.css" type="text/css">
<script type="text/javascript">
//
// ====== 	Function:	check if some data was entered other than spaces
//		Parm:		p_val		Check value
//		Return value:	If entered --> True,  Not entered --> False
//
function check_required(p_val){
	if (p_val == ""){return(false);}
	for(i=0; i<p_val.length; i++){
		if (p_val.substring(i, i+1)!=" " && p_val.substring(i, i+1)!="　"){
			return(true);
		}
	}
	return(false);
}
//
// ====== 	Function:	post on submit
//
function post_onSubmit(){
	var vChecked = false;
	if (check_required(document.f_data.MakerName.value) == false){
		alert("\nメーカーを入力してください。");
		document.f_data.MakerName.focus();
		return false;
 	}
	if (check_required(document.f_data.ProductName.value) == false){
		alert("\n商品名を入力してください。");
		document.f_data.ProductName.focus();
		return false;
 	}
	if ((document.f_data.Warranty[0].checked == false) && (document.f_data.Warranty[1].checked == false)){
		alert("\n保証書のあり/なしをチェックしてください。");
		return false;
 	}
	if (document.f_data.WhenPurchased[0].selected == true){
		alert("\nご購入後期間を選択してください。");
		return false;
 	}
	if (check_required(document.f_data.Comment.value) == false){
		alert("\n内容を入力してください。");
		document.f_data.Comment.focus();
		return false;
 	}
	return true;
}
//========================================================================
</script>
</head>
<body>
<!--#include file="../Navi/NaviTop.inc"-->
<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>
  
  <!-- コンテンツstart -->
  <div id="globalContents">
    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="<%=g_HTTP%>"><img src="<%=g_RelLink%>images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li class="now">修理･商品サポート</li>
      </ul>
    </div></div></div>

    <h1 class="title">修理･商品サポート</h1>

<!-- エラーメッセージ -->
<% if wMSG <> "" then %>
<ul class="error">
  <li><%=wMSG %></li>
</ul>
<% end if %>

<form name="f_data" id="inquiry" action="SupportInquiryConfirm.asp" method="post" onSubmit="return post_onSubmit();">
  <table>
    <tr>
      <th>メーカー<span>*</span></th>
      <td><input type="text" name="MakerName" value="<%=MakerName%>" size="50" maxlength="25"></td>
    </tr>
    <tr>
      <th>商品名<span>*</span></th>
      <td><input type="text" name="ProductName" value="<%=ProductName%>" size="65" maxlength="50"></td>
    </tr>
    <tr>
      <th>保証書<span>*</span></th>
      <td>
        <label><input type="radio" name="Warranty" value="あり"<% if Warranty = "あり" then %> checked="checked"<% end if %>>あり</label>　
        <label><input type="radio" name="Warranty" value="なし"<% if Warranty = "なし" then %> checked="checked"<% end if %>>なし</label>　<span>(保証書が無い場合、保証を受けられない場合があります)</span>
      </td>
    </tr>
    <tr>
      <th>シリアル番号</th>
      <td><input name="SerialNo" type="text" value="<%=SerialNo%>" size="60" maxlength="40"><br><span>(保証書、本体に記載がない場合は必要ありません)</span></td>
    </tr>
    <tr>
      <th>ご購入後期間<span>*</span></th>
      <td>
        <select name="WhenPurchased">
          <option value=""<% if WhenPurchased = "" then%> selected="selected"<% end if %>>選択してください 
          <option value="一週間以内"<% if WhenPurchased = "一週間以内" then%> selected="selected"<% end if %>>一週間以内 
          <option value="一年未満"<% if WhenPurchased = "一年未満" then%> selected="selected"<% end if %>>一年未満
          <option value="一年以上/不明"<% if WhenPurchased = "一年以上/不明" then%> selected="selected"<% end if %>>一年以上/不明
        </select>
      </td>
    </tr>
    <tr>
      <th>内容<span>*</span><br>（500文字まで）</th>
      <td><textarea name="Comment" rows="5" cols="55"><%=Comment%></textarea></td>
    </tr>
    <tr>
      <th>お名前</th>
      <td><%=wCustomerName%></td>
    </tr>
    <tr>
      <th>住所</th>
      <td>
        〒<%=wZip%><br>
        <%=wPrefecture%><%=wAddress%>
      </td>
    </tr>
    <tr>
      <th>電話番号</th>
      <td><%=wTelephone%></td>
    </tr>
    <tr>
      <th>Fax番号</th>
      <td><%=wFax%></td>
    </tr>
    <tr>
      <th>メールアドレス</th>
      <td><%=wEmail%></td>
    </tr>
  </table>
  <p>「*」のついている項目は必須入力項目です。</p>
  <input type="hidden" name="Skey" value="<%=Skey%>">
  <p class="btnBox"><input type="submit" value="内容を確認する" class="opover"></p>
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