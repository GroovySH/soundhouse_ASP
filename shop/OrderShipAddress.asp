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

'	届先ページ
'変更履歴
'

'========================================================================
On Error Resume Next
Response.Expires = -1			' Do not cache

'---- Session情報
Dim wUserID
Dim wUserName
Dim wMsg

'---- 受け渡し情報を受取る変数
Dim ship_name
Dim ship_zip
Dim ship_prefecture
Dim ship_address
dim ship_telephone

'---- DB
Dim Connection
'=======================================================================
'	受け渡し情報取り出し
'=======================================================================
'---- Session変数
wUserID = Session("UserID")
wUserName = Session("userName")
wMsg = Session("msg")

'---- 受け渡し情報取り出し
ship_name = Left(ReplaceInput(Trim(Request("ship_name"))), 30)
ship_zip = Left(ReplaceInput(Trim(Request("ship_zip"))), 10)
ship_prefecture = Left(ReplaceInput(Trim(Request("ship_prefecture"))), 4)
ship_address = Left(ReplaceInput(Trim(Request("ship_address"))), 40)
ship_telephone = Left(ReplaceInput(Trim(Request("ship_telephone"))), 20)

Session("msg") = ""

'---- セッション切れチェック
If wUserID = ""Then
	Response.Redirect g_HTTP
End If

'========================================================================
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<title>お届け先の登録｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/StyleOrder.css?20120629a" type="text/css">
<script type="text/javascript">
//=====================================================================
//	住所検索 onClick
//=====================================================================
function address_search_onClick(){

	var addrWin;

	if (document.f_data.ship_zip.value == ""){
		alert("郵便番号を入力してください。");
		return;
	}
 
	AddrWin = window.open("../comasp/address_search.asp?zip=" + document.f_data.ship_zip.value + "&name_prefecture=i_ship_prefecture&name_address=ship_address","AddrSearch","width=200,height=100");

}
//=====================================================================
//	ラジオボタン、ドロップダウンリストを以前に選択した状態にする
//=====================================================================
function preset_values(){

	// 住所検索処理からの呼び出し時は都道府県のみをセット
	for (var i=0; i<document.f_data.ship_prefecture.options.length; i++){
		if (document.f_data.ship_prefecture.options[i].value == document.f_data.i_ship_prefecture.value){
			document.f_data.ship_prefecture.options[i].selected = true;
			break;
		}
	}
	return;

}
//=====================================================================
//	次へボタン onClick
//=====================================================================
function Next_onClick() {
	document.f_data.action = "OrderShipAddressStore.asp";
	document.f_data.submit();
}
//=====================================================================
//	キャンセルボタン onClick
//=====================================================================
function Cancel_onClick() {
	document.f_data.action = "OrderInfoEnter.asp";
	document.f_data.submit();
}
</script>
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->
<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>

<!-- コンテンツstart -->
<div id="globalContents">

  <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
    <p class="home"><a href="../"><img src="../images/icon_home.gif" alt="HOME"></a></p>
    <ul id="path">
      <li>お届け先、お支払方法の選択</li>
      <li class="now">お届け先の登録</li>
    </ul>
  </div></div></div>

  <h1 class="title">お届け先、お支払方法の選択</h1>
  <ol id="step">
    <li><img src="images/step01.gif" alt="1.ショッピングカート" width="170" height="50"></li>
    <li><img src="images/step02_now.gif" alt="2.お届け先、お支払方法の選択" width="170" height="50"></li>
    <li><img src="images/step03.gif" alt="3.ご注文内容の確認" width="170" height="50"></li>
    <li><img src="images/step04.gif" alt="4.ご注文完了" width="170" height="50"></li>
  </ol>

  <p class="error"><% = wMsg %></p>

  <h2 class="cart_title">お届け先の登録</h2>

  <form name="f_data" method="post">
    <table id="shipAddress">
      <tr>
        <th>お名前</th>
        <td><input type="text" name="ship_name" id="ship_name" size=30 maxlength=60 value="<% = ship_name %>"></td>
      </tr>
      <tr>
        <th>住所</th>
        <td>
          〒<input type="text" name="ship_zip" id="ship_zip" size="10" maxlength="8" value="<% = ship_zip %>"><span>（半角）</span>
          <a href="JavaScript:address_search_onClick();" class="tipBtn">住所検索</a><span>郵便番号を入力してボタンを押してください｡</span><br>
          <select name="ship_prefecture">
            <option value="">都道府県</option>
            <option value="北海道">北海道</option>
            <option value="青森県">青森県</option>
            <option value="秋田県">秋田県</option>
            <option value="岩手県">岩手県</option>
            <option value="宮城県">宮城県</option>
            <option value="山形県">山形県</option>
            <option value="福島県">福島県</option>
            <option value="栃木県">栃木県</option>
            <option value="新潟県">新潟県</option>
            <option value="群馬県">群馬県</option>
            <option value="埼玉県">埼玉県</option>
            <option value="茨城県">茨城県</option>
            <option value="千葉県">千葉県</option>
            <option value="東京都">東京都</option>
            <option value="神奈川県">神奈川県</option>
            <option value="山梨県">山梨県</option>
            <option value="長野県">長野県</option>
            <option value="岐阜県">岐阜県</option>
            <option value="富山県">富山県</option>
            <option value="石川県">石川県</option>
            <option value="静岡県">静岡県</option>
            <option value="愛知県">愛知県</option>
            <option value="三重県">三重県</option>
            <option value="奈良県">奈良県</option>
            <option value="和歌山県">和歌山県</option>
            <option value="福井県">福井県</option>
            <option value="滋賀県">滋賀県</option>
            <option value="京都府">京都府</option>
            <option value="大阪府">大阪府</option>
            <option value="兵庫県">兵庫県</option>
            <option value="岡山県">岡山県</option>
            <option value="鳥取県">鳥取県</option>
            <option value="島根県">島根県</option>
            <option value="広島県">広島県</option>
            <option value="山口県">山口県</option>
            <option value="香川県">香川県</option>
            <option value="徳島県">徳島県</option>
            <option value="愛媛県">愛媛県</option>
            <option value="高知県">高知県</option>
            <option value="福岡県">福岡県</option>
            <option value="佐賀県">佐賀県</option>
            <option value="大分県">大分県</option>
            <option value="熊本県">熊本県</option>
            <option value="宮崎県">宮崎県</option>
            <option value="長崎県">長崎県</option>
            <option value="鹿児島県">鹿児島県</option>
            <option value="沖縄県">沖縄県</option>
          </select>
          <input type="text" name="ship_address" id="ship_address" size="60" maxlength="80" value="<% = ship_address %>"><br><span>会社名、マンション/ビル名、部屋番号、ｘｘ様方、等は忘れずご記入ください。</span>
        </td>
      </tr>
      <tr>
        <th>電話番号</th>
        <td><input type="text" name="ship_telephone" id="ship_telephone" size="30" maxlength="20" value="<% = ship_telephone %>"><span>（半角数字）</span></td>
      </tr>
    </table>

    <div id="btn_box">
      <ul class="btn">
        <li><a href="javascript:Cancel_onClick();"><img src="images/btn_back.png" alt="戻る" width="151" height="32" class="opover"></a></li>
        <li class="last"><a href="javascript:Next_onClick();"><img src="images/btn_next.png" alt="次へ" width="151" height="32" class="opover"></a></li>
      </ul>
    </div>

    <input type="hidden" name="i_ship_prefecture" value="<% = ship_prefecture %>">
  </form>

<!--/#contents --></div>
	<div id="globalSide">
	<!--#include file="../Navi/NaviSide.inc"-->
	<!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript">
	preset_values();
</script>
</body>
</html>