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
'	友達にすすめる送信
'
'更新履歴
'2007/08/23 商品アクセス件数登録（友達にすすめる）
'2007/09/10 商品アクセス件数登録を年月別に変更
'2009/04/30 エラー時にerror.aspへ移動
'2011/08/01 an #1087 Error.aspログ出力対応
'
'========================================================================

On Error Resume Next

Dim userEmail
Dim UserName

Dim Item
Dim ToAddr
Dim FromName
Dim Message1
Dim Message

Dim wItem
Dim MakerCd
Dim ProductCd

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim wBody

Dim Connection
Dim RS

Dim wSQL
Dim wHTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.buffer = true

'---- UserID 取り出し

'---- 呼び出し元からのデータ取り出し
Item = ReplaceInput(Request("Item"))
ToAddr = ReplaceInput_NoCRLF(Request("ToAddr"))  '2011/08/01 an mod
FromName = ReplaceInput(Request("FromName"))
Message = ReplaceInput(Request("Message"))
Message1 = ReplaceInput(Request("Message1"))

wItem = Split(Item, "^")
MakerCd = Trim(wItem(0))
ProductCd = Trim(wItem(1))

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "TellaFriendSend.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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

Dim i
Dim vYYYYMM
Dim OBJ_NewMail
Dim RSv

'---- おすすめメール作成（顧客へ)
call getCntlMst("Web","Email","一般トレーラ", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)

wBody = Message1 & vbNewLine & Message & vbNewLine & vbNewLine & wItemChar1

Set OBJ_NewMail = Server.CreateObject("CDO.Message") 

OBJ_NewMail.from = "shop@soundhouse.co.jp"
OBJ_NewMail.to = ToAddr

OBJ_NewMail.subject = FromName & "　様から、おすすめメールが届いています"
OBJ_NewMail.TextBody = wBody
OBJ_NewMail.BodyPart.Charset = "iso-2022-jp"

OBJ_NewMail.Send

Set OBJ_NewMail = Nothing

'---- 商品アクセス件数登録（友達にすすめる）
vYYYYMM = Year(Now()) & Right("0" & Month(Now()),2)
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM 商品アクセス件数"
wSQL = wSQL & " WHERE メーカーコード = '" & MakerCd & "'"
wSQL = wSQL & "   AND 商品コード = '" & ProductCd & "'"
wSQL = wSQL & "   AND 年月 = '" & vYYYYMM & "'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RSv.EOF = true then
	RSv.AddNew

	RSv("メーカーコード") = MakerCd
	RSv("商品コード") = ProductCd
	RSv("年月") = vYYYYMM
	RSv("友達にお勧め件数") = 1
else
	RSv("友達にお勧め件数") = RSv("友達にお勧め件数") + 1
end if

RSv.Update
RSv.close

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
<title>お友達にすすめる｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/ask.css" type="text/css">
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
        <li class="now">お友達にすすめる</li>
      </ul>
    </div></div></div>

    <h1 class="title">お友達にすすめる</h1>
    <p>以下の内容でメールを送信しました。<br>ありがとうございました。</p>
    
    <table class="form">
      <tr>
        <th>宛先</th>
        <td><%=ToAddr%></td>
      </tr>
      <tr>
        <th>件名</th>
        <td><%=FromName%> 様から、おすすめメールが届いています</td>
      </tr>
      <tr>
        <th>メッセージ</th>
        <td><p><%=Replace(wBody, vbNewLine, "<br>")%></p></td>
      </tr>
    </table>
    <p><a href="ProductDetail.asp?Item=<%=Item%>">商品ページへ戻る</a></p>

</div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<div class="tooltip"><p>ASK</p></div>
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript" src="jslib/ask.js"></script>
</body>
</html>