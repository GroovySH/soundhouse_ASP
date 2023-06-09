<%@ LANGUAGE="VBScript" %>
<%
'ネットハウスねっとハウスネットはうす
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
'	プレゼントフォームの送信
'
'更新履歴
'2005/05/13 OBJ_NewMail.BodyPart.Charset = "iso-2022-jp"をセット
'2005/09/29 メールサブジェクトにコンタクト管理用情報を追加
'2008/05/12 改行コードインジェクション対策（i_toパラメータ削除）
'2008/05/13 クロスサイトリクエストフォジェリー対策 Keyパラメータチェック
'2009/04/30 エラー時にerror.aspへ移動
'
'========================================================================

On Error Resume Next

Dim userID
Dim msg

Dim customer_nm
Dim furigana
Dim zip
Dim prefecture
Dim address
Dim telephone
Dim e_mail
Dim purchase
Dim comment

Dim w_sql
Dim w_html
Dim w_msg

'========================================================================

Response.buffer = true

'---- セキュリティーキーチェック
if Session("Skey") <> ReplaceInput(Request("SKey")) then
	Response.redirect "PresentOubo.asp"
end if

'---- UserID 取り出し
userID = Session("userID")

'---- 呼び出し元からのデータ取り出し
customer_nm = ReplaceInput(Request("customer_nm"))
furigana = ReplaceInput(Request("furigana"))
zip = ReplaceInput(Request("zip"))
prefecture = ReplaceInput(Request("prefecture"))
address = ReplaceInput(Request("address"))
telephone = ReplaceInput(Request("telephone"))
e_mail = ReplaceInput(Request("e_mail"))
purchase = ReplaceInput(Request("purchase"))
comment = ReplaceInput(Request("comment"))

'---- Execute main
call main()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'========================================================================
'
'	Function	Main
'
'========================================================================
'
Function main()

Dim i
Dim v_body
Dim v_item
Dim OBJ_NewMail

'---- edit body
v_body = ""
v_body = v_body & "受付日時：" & now() & vbNewLine & vbNewLine
v_body = v_body & "顧客番号：" & userID & vbNewLine & vbNewLine
v_body = v_body & "名前　　：" & customer_nm & vbNewLine
v_body = v_body & "ふりがな：" & furigana & vbNewLine
v_body = v_body & "住所　　：" & zip & " " & prefecture & address & vbNewLine
v_body = v_body & "電話番号：" & telephone & vbNewLine
v_body = v_body & "Ｅメール：" & e_mail & vbNewLine
v_body = v_body & "購入歴　：" & purchase & vbNewLine
v_body = v_body & "コメント：" & comment & vbNewLine

'@@@@@response.write(v_body)

'---- send e-mail
Set OBJ_NewMail = Server.CreateObject("CDO.Message") 

OBJ_NewMail.from = "present@soundhouse.co.jp"
OBJ_NewMail.to = "present@soundhouse.co.jp"
OBJ_NewMail.subject = "プレゼント応募 " & customer_nm & " [" & userID & "/Web-Emax/プレゼント応募]"
OBJ_NewMail.TextBody = v_body
OBJ_NewMail.BodyPart.Charset = "iso-2022-jp"

OBJ_NewMail.Send

Set OBJ_NewMail = Nothing

End function

'========================================================================
%>

<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="Shift_JIS">
<title>プレゼント応募｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
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
        <li class="now">プレゼント応募</li>
      </ul>
    </div></div></div>

    <h1 class="title">プレゼント応募</h1>
    <p>プレゼントのご応募を承りました。<br>ありがとうございました。</p>
    
</div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>