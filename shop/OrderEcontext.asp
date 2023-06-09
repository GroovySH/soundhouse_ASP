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
'	コンビニ支払リクエスト処理 (eContext)
'
'------------------------------------------------------------------------
'	更新履歴
'2008/04/28 リベート対応のため合計金額のみセット
'2008/05/14 HTTPSチェック対応
'2009/04/30 エラー時にerror.aspへ移動
'2010/09/27 hn eContextからの戻り値のチェック強化
'2011/04/14 hn SessionID関連変更
'2011/08/01 an #1087 Error.aspログ出力対応
'
'========================================================================

On Error Resume Next

Dim userID
Dim msg

Dim CustomerTel
Dim CustomerName
Dim CustomerEmail
Dim Shipping
Dim OrderTotal
Dim SalesTax
Dim OrderDate
Dim ItemTotal
Dim OrderNo
Dim wSalesTaxRate
Dim wPrice

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim Connection
Dim RS

Dim OBJeContext
Dim eContextRtn
Dim eContextNo
Dim eConF
Dim eConK

Dim wSQL
Dim wHTML
Dim wMSG
Dim wNextURL
Dim wErrDesc   '2011/08/01 an add

'=======================================================================

userID = Session("UserID")

Session("msg") = ""
wMSG = ""

'---- execute main process
call ConnectDB()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "OrderEcontext.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'---- エラーが無いときは注文登録処理ページ、エラーがあれば確認ページへ
if wMSG = "" then
''response.write("NO=" & eContextNo & "F=" & eConF & "  k=" & eConK)
	Response.Redirect "OrderSubmit.asp?OrderNo=" & OrderNo & "&eConF=" & eConF & "&eConK=" & eConK
else
	Session("msg") = wMSG
	Response.Redirect "OrderInfoEnter.asp"
end if

'========================================================================
'========================================================================
'
'	Function	Connect database
'
'========================================================================
'
Function ConnectDB()

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

Dim vTemp

'---- 消費税率取出し
call getCntlMst("共通","消費税率","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'消費税率
wSalesTaxRate = Clng(wItemNum1)

'---- 仮受注取り出し
wSQL = ""
wSQL = wSQL & "SELECT a.商品合計金額"
wSQL = wSQL & "     , a.送料"
wSQL = wSQL & "     , a.代引手数料"
wSQL = wSQL & "     , a.コンビニ支払手数料"
wSQL = wSQL & "     , a.外税合計金額"
wSQL = wSQL & "     , a.受注合計金額"
wSQL = wSQL & "     , a.顧客電話番号"
wSQL = wSQL & "     , a.顧客E_mail"
wSQL = wSQL & "     , a.eContext受付番号"
wSQL = wSQL & "     , b.顧客名"
wSQL = wSQL & "  FROM 仮受注 a"
wSQL = wSQL & "     , Web顧客 b"
wSQL = wSQL & " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
wSQL = wSQL & "   AND b.顧客番号 = a.顧客番号"
	  
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS.EOF = true then
	wMSG = "<font color='#ff0000'>NoData</font>"
	exit function
end if

CustomerTel = cf_numeric_only(RS("顧客電話番号"))
CustomerName = RS("顧客名")
CustomerEmail = RS("顧客E_mail")
'Shipping = (RS("送料") + RS("代引手数料") + RS("コンビニ支払手数料")) * (100 + wSalesTaxRate) / 100
OrderTotal = RS("受注合計金額")
'SalesTax = RS("外税合計金額")
OrderDate = cf_FormatDate(Now(), "YYYY/MM/DD") & " " & cf_FormatTime(Now(), "HH:MM:SS")
'ItemTotal = RS("受注合計金額") - Shipping
ItemTotal = OrderTotal

OrderNo = GetOrderNo()              '注文番号

'---- eContext リクエスト
Set OBJeContext = CreateObject("FormPost.Https")

vTemp = OBJeContext.init()

vTemp = OBJeContext.set("shopID", g_eContext_ID)
vTemp = OBJeContext.set("orderID", OrderNo)
vTemp = OBJeContext.set("sessionID", "1")
vTemp = OBJeContext.set("telNo", CustomerTel)
vTemp = OBJeContext.set("kanjiName1_1", CustomerName)
vTemp = OBJeContext.set("kanjiName1_2", "　")
vTemp = OBJeContext.set("email", CustomerEmail)
vTemp = OBJeContext.set("paymentFlg", "0")
vTemp = OBJeContext.set("shippmentFlg", "2")
'vTemp = OBJeContext.set("commission", Shipping)
vTemp = OBJeContext.set("commission", 0)
vTemp = OBJeContext.set("ordAmount", OrderTotal)
vTemp = OBJeContext.set("ordAmountbfTax", "0")
'vTemp = OBJeContext.set("ordAmountTax", SalesTax)
vTemp = OBJeContext.set("ordAmountTax", 0)
vTemp = OBJeContext.set("ordItemNo", "1")
vTemp = OBJeContext.set("orderDate", OrderDate)
vTemp = OBJeContext.set("siteInfo", "領収書備考を記述")

vTemp = OBJeContext.set("itemName1", "ご注文一式(税込み)")
vTemp = OBJeContext.set("unitPrice1", ItemTotal)
vTemp = OBJeContext.set("ordUnit1", "1")
vTemp = OBJeContext.set("unitChar1", "式")
vTemp = OBJeContext.set("dtlAmount1", ItemTotal)
vTemp = OBJeContext.set("goodsCode1", "0")

'---- リクエスト
vTemp = OBJeContext.send3(g_eContext_URL)

'---- 戻り値のチェック
call checkError(vTemp)

'---- 受注情報にeContext受付番号をセット
if wMSG = "" then
	RS("eContext受付番号") = eContextNo
	RS.update
end if


'----
vTemp = OBJeContext.finally()
Set OBJeContext = Nothing

RS.close

End Function

'========================================================================
'
'	Function	受注番号取り出し
'
'========================================================================
'
Function GetOrderNo()

Dim vRS_Cntl

'---- コントロールマスタ取り出し
wSQL = ""
wSQL = wSQL & "SELECT item_num1"
wSQL = wSQL & "  FROM コントロールマスタ"
wSQL = wSQL & " WHERE sub_system_cd = '共通'"
wSQL = wSQL & "   AND item_cd = '番号'"
wSQL = wSQL & "   AND item_sub_cd = 'Web受注'"
	  
Set vRS_Cntl = Server.CreateObject("ADODB.Recordset")
vRS_Cntl.Open wSQL, Connection, adOpenStatic, adLockOptimistic

vRS_Cntl("item_num1") = Clng(vRS_Cntl("item_num1")) + 1
GetOrderNo = vRS_Cntl("item_num1")

vRS_Cntl.update
vRS_Cntl.close

End function

'========================================================================
'
'	Function	カードエラーチェック
'
'========================================================================
'
Function checkError(pRtn)

Dim vTemp

eContextRtn = Split(pRtn,chr(10))			'戻り値を行単位に分割
vTemp = Split(eContextRtn(0)," ")			'line1を' 'で分割

'---- 戻りコードチェック　　正常
'2010/09/27 hn mod s
If (vTemp(0) = "1") Then
	eContextNo = Replace(eContextRtn(1), chr(13), "")		'受付番号
	eConF = Replace(eContextRtn(2), chr(13), "")				'振り込み票URL
	eConK = Replace(eContextRtn(7), chr(13), "")				'決済選択用URL
end if

if (vTemp(0) = "1") AND eContextNo <> "" AND eConF <> "" AND eConK <> "" then
	wMSG = ""
else
	'---- エラー
	wMSG = "<font color='#ff0000'>" _
				& "申し訳ございませんが､処理中にエラーが発生しました｡<br>" _
				& "Code: " & eContextRtn(0) & "<br>" _
				& "もう一度ご注文いただくか、他のお支払方法を選択ください｡<br>よくあるご質問は<a href='" & G_HTTP & "information/t_qanda.htm#card'>こちら</a>" _
				& "</font>"
end if
'2010/09/27 hn mod e

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

%>
