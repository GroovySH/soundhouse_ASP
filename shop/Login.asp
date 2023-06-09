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
'	ログインページ
'
'	更新履歴
'2004/12/20 呼び出しもとURL設定追加
'2006/09/18 LoginFl追加　ログイン情報保持
'2007/08/13 エラーメッセージ表示変更
'2008/05/12 パスワードリセットをHTTPSに変更
'2008/05/14 HTTPSチェック対応
'2010/07/30 st RtnURLがある場合はそのまま呼び出し元へリダイレクト'
'2011/04/20 an #843 ログイン時、Emailの代わりにユーザーIDを使用
'2011/05/09 an #843関連 「ユーザーIDを忘れた時は」追加
'2011/05/11 an 「パスワードを忘れた時は」はユーザーID/電話番号の入力に変更

'========================================================================

Dim member_email  '2011/04/20 an del, 2011/05/09 an re-add
Dim telephone     '2011/05/09 an add
Dim MemberID      '2011/05/11 an add
Dim telephone_password  '2011/05/11 an add
Dim msg

Dim called_from
Dim logoff_fl
Dim userID
Dim RtnURL

Dim w_html
Dim w_msg

'========================================================================

'gHTTPSPage = true		'HTTPSページ

Response.buffer = true

'---- 呼び出し元プログラムからのメッセージ取り出し

msg = Session.contents("msg")
Session("msg") = ""
'userID = Session("userID")

called_from = ReplaceInput(Request("called_from"))
logoff_fl = ReplaceInput(Request("logoff_fl"))
RtnURL = replace(ReplaceInput(Request("RtnURL")), "＆", "&")		'呼び出し元URL '2010/07/30 st mod
member_email = ReplaceInput_NoCRLF(Left(Request("member_email"),60))  '2011/05/09 an add エラー時に受け取り
telephone = ReplaceInput_NoCRLF(Left(Request("telephone"),20))        '2011/05/09 an add
MemberID = ReplaceInput_NoCRLF(Left(Request("MemberID"),60))      '2011/05/11 an add エラー時に受け取り
telephone_password = ReplaceInput_NoCRLF(Left(Request("telephone_password"),20)) '2011/05/11 an add

if logoff_fl = "Y" then
	'---- 顧客番号, 顧客名Cookieをクリア
	Session("userID") = ""
	Session("userName") = ""
	'Session("userEmail") = ""  '2011/04/20 an del
	Session("LoginFl") = ""

	Response.Redirect g_HTTP		' Log out 後　TOPへ戻る
'else
	'---- 以前に入力したEmailを取り出し
	'member_email = Session.contents("member_email")
	'if userID <> "" AND member_email = "" then
	'	member_email = Session("userEmail")
	'end if

	'Session("member_email") = ""
end if

'========================================================================

%>

<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="Shift_JIS">
<title>ログイン｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link href="<%=g_HTTPS%>/shop/style/login.css?20120718" rel="stylesheet" type="text/css">

<script type="text/javascript">
//
//	Login onSubmit
//
function Login_onSubmit(p_form){

	if (p_form.MemberID.value == ""){
		alert("ユーザーIDを入力してください。");
		p_form.MemberID.focus();
		return false;
	}
	if (p_form.member_password.value == ""){
		alert("パスワードを入力してください。");
		p_form.member_password.focus();
		return false;
	}
		return true;
}

//
//	Password onSubmit
//
function Password_onSubmit(p_form){

	if (p_form.MemberID.value == ""){
		alert("ユーザーIDを入力してください。");
		p_form.MemberID.focus();
		return false;	
	}
	if (p_form.telephone_password.value == ""){
		alert("電話番号を入力してください。");
		p_form.telephone_password.focus();
		return false;
	}
		return true;
}

//
//	UserID onSubmit
//
function UserID_onSubmit(p_form){

	if (p_form.member_email.value == ""){
		alert("メールアドレスを入力してください。");
		p_form.member_email.focus();
		return false;	
	}
	if (p_form.telephone.value == ""){
		alert("電話番号を入力してください。");
		p_form.telephone.focus();
		return false;
	}
		return true;
}

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
        <li class="now">ログイン</li
      ></ul>
    </div></div></div>


    <h1 class="title">ログイン</h1>

	<% if msg <> "" then %>
  	<p class="error"><%=msg%></p>
	<% end if %>
    
    <ul id="login">
      <li>
      	<form name="fLogin" method="post" action="<%=g_HTTPS%>shop/LoginCheck.asp" onSubmit="return Login_onSubmit(this);">
    	<h2>会員登録をされている方</h2>
    	<p>ユーザーID・パスワードを入力し[ログイン]ボタンを押してください。<br><a href="#login00">※ログインできない場合はこちら</a></p>
        <table class="form">
            <tr>
              <th>ユーザーID</th>
              <td><input name="MemberID" type="text">半角英数字</td>
            </tr>
            <tr>
              <th>パスワード</th>
              <td><input name="member_password" type="password">半角英数字</td>
            </tr>
          </table>
          <p class="btnBox"><input type="submit" value="ログイン" class="opover"></p>
          <input name="called_from" type="hidden" value="<%=called_from%>">
          <input name="RtnURL" type="hidden" value="<%=RtnURL%>">
        </form>
      </li>
      <li>
    	<h2>会員登録をされていない方</h2>
    	<p>会員登録は無料でカンタンです！ご登録いただければ、次回より住所の入力が必要ありません。<br>また、WEB会員だけのメールニュース等お得な情報がいっぱいです。</p>
        <p class="btnBox"><a href="<%=g_HTTPS%>Member/MemberAgreement.asp?called_from=navi" class="opover">ご登録はこちら</a></p>
      </li>
      <li class="forget">
    	<h3>パスワードを忘れた時は</h3>
        <form name="fForgotPassword" method="post" action="<%=g_HTTPS%>Member/MemberPasswordSend.asp?called_from=<%=called_from%>" onSubmit="return Password_onSubmit(this);">
    	<p>パスワードを忘れた方は、ご登録のユーザーID・電話番号を入力し[パスワードリセット]ボタンを押してください。<br>ご登録のメールアドレス宛にメール送付されますので、ご案内内容をご確認ください。</p>
        <table class="form">
            <tr>
              <th>ユーザーID</th>
              <td><input name="MemberID" type="text" value="<%=MemberID%>">半角英数字</td>
            </tr>
            <tr>
              <th>電話番号</th>
              <td><input name="telephone_password" type="text" value="<%=telephone_password%>">半角数字</td>
            </tr>
          </table>
          <p class="btnBox"><input type="submit" value="パスワードリセット" class="opover"></p>
          <input name="called_from" type="hidden" value="<%=called_from%>">
          <input name="i_function" type="hidden" value="send">
        </form>
      </li>
      <li class="forget">
    	<h3>ユーザーIDを忘れた時は</h3>
        <form name="fForgotUserID" method="post" action="<%=g_HTTPS%>Member/MemberUserIDSend.asp?called_from=<%=called_from%>" onSubmit="return UserID_onSubmit(this);">
    	<p>ユーザーIDを忘れた方は、ご登録のメールアドレス・電話番号を入力し[ユーザーID確認]ボタンを押してください。<br>ご登録のメールアドレス宛にユーザーIDをお知らせしますので、ご確認ください。</p>
        <table class="form">
            <tr>
              <th>メールアドレス</th>
              <td><input name="member_email" type="text" value="<%=member_email%>">半角英数字</td>
            </tr>
            <tr>
              <th>電話番号</th>
              <td><input name="telephone" type="text" value="<%=telephone%>">半角数字</td>
            </tr>
          </table>
          <p class="btnBox"><input type="submit" value="ユーザーID確認" class="opover"></p>
        </form>
      </li>
    </ul>
    
    <div id="login00">
    	<h4>ログインできない場合は</h4>
        <p>ログインできない場合、以下の項目をご確認ください。</p>
        <h5>【エラーメッセージ】</h5>
        <p>画面上部に赤字のエラーメッセージが表示される場合は、メッセージ内容に従いご入力内容の修正をお願いします。</p>
        <h5>【日時設定】</h5>
        <p>お使いのパソコンの日時設定が正しい時刻になっているかご確認ください。</p>
        <h5>【ブラウザの再起動】</h5>
        <p>お使いのブラウザの開いている全てのウインドウを一旦閉じ、再度開いてご確認ください。</p>
        <h5>【クッキーの削除】</h5>
        <p>上記を確認してもログインができない場合は、クッキーのクリアをお試しください。</p>
        <ul>
        	<li><a href="<%=g_HTTP%>guide/qanda14.asp#ie">Internet Explorerをご利用の方</a></li>
            <li><a href="<%=g_HTTP%>guide/qanda14.asp#ff">Firefoxをご利用の方</a></li>
            <li><a href="<%=g_HTTP%>guide/qanda14.asp#sf">Safariをご利用の方</a></li>
            <li><a href="<%=g_HTTP%>guide/qanda14.asp#cr">Chromeをご利用の方</a></li>
        </ul>
        <p>※キャッシュ/クッキーのクリア後はブラウザを再起動してください。</p>
        <h5>【お問い合わせ先】</h5>
        <p>ログインについてのお問い合わせは、<a href="<%=g_HTTPS%>shop/Inquiry.asp">お問い合わせフォ－ム</a>または下記へお願いいたします。</p>
        <ul>
        	<li>TEL：0476-89-1111</li>
            <li>FAX：0476-89-2222</li>
            <li>MAIL：<a href="mailto:shop@soundhouse.co.jp">shop@soundhouse.co.jp</a></li>
        </ul>
    </div>


</div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>