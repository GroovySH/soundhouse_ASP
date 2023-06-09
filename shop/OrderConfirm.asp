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
'	ご注文内容の確認ページ
'
'2012/06/15 ok デザイン変更のため旧版を元に新規作成
'2013/10/21 GV # 大型商品の表示
'
'========================================================================
On Error Resume Next
Response.Expires = -1			' Do not cache
Response.buffer = true

'---- Session情報
Dim wUserID
Dim wUserName
Dim wMsg
Dim Skey

'---- 受け渡し情報を受取る変数

'---- 顧客情報
Dim wCustomerNm
Dim wCustomerZip
Dim wCustomerPref
Dim wCustomerAddress
Dim wCustomerTel
Dim wCustomerKabusokuAm

'---- 仮受注情報
Dim wShipNm
Dim wShipZip
Dim wShipPrefecture
Dim wShipAddress
Dim wShipTel
Dim wFreightForwarder
Dim wDeliveryDt
Dim wDeliveryTm
Dim wEigyoushoDome
Dim wIkkatsu
Dim wPaymentMethod
Dim wFurikomiNm
Dim wRitouFl
Dim wRebateFl


'2013/10/21 GV # add start
'---- 大型商品
Dim wLargeItemHtml
Dim wLargeItemFl
Dim wNonLargeItemFl
wLargeItemHtml = ""
wLargeItemFl = "N"
wNonLargeItemFl = "N"
'2013/10/21 GV # add end

'---- 金額
Dim wPrdctAmTotal
Dim wPrdctAmTotalNoTax
Dim wShippingNoTax
Dim wCodAm
Dim wTax
Dim wOrderAmTotal
Dim wSokoCnt
Dim wSalesTaxRate
Dim wKoguchi
Dim wErrDesc   '2011/08/01 an add
Dim wTotal_NoDaibikiFee  '2012/03/03 an add

'---- DB
Dim Connection

'---- HTML
Dim wProductHtml
Dim wHaisouHtml
Dim wPaymentHtml
Dim wReceiptHtml

'=======================================================================
'	受け渡し情報取り出し
'=======================================================================
'---- Session変数
wUserID = Session("UserID")
wUserName = Session("userName")
wMsg = Session.Contents("msg")

'---- 受け渡し情報取り出し
Session("msg") = ""

'---- セッション切れチェック
If wUserID = "" Then
	Response.Redirect g_HTTP
End If

'=======================================================================
'	Execute main
'=======================================================================
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "OrderConfirm.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

If Err.Description <> "" Then
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

If wMsg <> "" Then
	Session("msg") = wMsg
	Server.Transfer "OrderInfoEnter.asp"
End If

'========================================================================
'
'	Function	Connect database
'
'========================================================================
Function connect_db()

'---- Connect database
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
Function close_db()

Connection.Close
Set Connection= Nothing    '2011/08/01 an add

End Function

'========================================================================
'
'	Function	main proc
'
'========================================================================
Function main()

Dim vItemChar1
Dim vItemChar2
Dim vItemNum1
Dim vItemNum2
Dim vItemDate1
Dim vItemDate2

'---- セキュリティーキーセット
Skey = SetSecureKey()

'---- 消費税率取出し
Call getCntlMst("共通","消費税率","1", vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)
wSalesTaxRate = Clng(vItemNum1)

Call getCustomer()				'顧客情報の取り出し
Call getOrder()					'仮受注情報の取り出し、更新

End Function

'========================================================================
'
'	Function	顧客情報の取り出し
'
'========================================================================
Function getCustomer()

Dim vSQL
Dim RSv

'---- 顧客情報取り出し
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    a.顧客名"
vSQL = vSQL & "  , a.入金過不足金額"
vSQL = vSQL & "  , b.顧客郵便番号"
vSQL = vSQL & "  , b.顧客都道府県"
vSQL = vSQL & "  , b.顧客住所"
vSQL = vSQL & "  , c.顧客電話番号"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    Web顧客 a WITH (NOLOCK)"
vSQL = vSQL & "  , Web顧客住所 b WITH (NOLOCK)"
vSQL = vSQL & "  , Web顧客住所電話番号 c WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "        a.顧客番号 = " & wUserID
vSQL = vSQL & "    AND b.顧客番号 = a.顧客番号"
vSQL = vSQL & "    AND b.住所連番 = 1"
vSQL = vSQL & "    AND c.顧客番号 = a.顧客番号"
vSQL = vSQL & "    AND c.住所連番 = 1"
vSQL = vSQL & "    AND c.電話連番 = 1"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

If RSv.EOF = True Then
	wMsg = wMsg & "顧客情報がありません。<BR />"
Else
	wCustomerNm = RSv("顧客名")
	wCustomerZip = RSv("顧客郵便番号")
	wCustomerPref = RSv("顧客都道府県")
	wCustomerAddress = RSv("顧客住所")
	wCustomerTel = RSv("顧客電話番号")
	wCustomerKabusokuAm = RSv("入金過不足金額")
End If

RSv.Close

End Function

'========================================================================
'
'	Function	受注情報の取り出し 更新（送料)
'
'========================================================================
Function getOrder()

Dim vSQL
Dim RSv
Dim vLoanDownPayment
Dim vPaymentMethodDisp	'2011/03/22
Dim wLargeItemHTMLBuff	'2013/10/21 GV # add

'---- 仮受注情報取り出し
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    *"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    仮受注"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    SessionID = '" & gSessionID & "'" 

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

If RSv.EOF = True Then
	Exit Function
End If

'---- 届先
wShipNm = RSv("届先名前")
wShipZip = RSv("届先郵便番号")
wShipPrefecture = RSv("届先都道府県")
wShipAddress = RSv("届先住所")
wShipTel = RSv("届先電話番号")

'---- 配送情報
Select Case Trim(RSv("運送会社コード"))
	Case "1"
		wFreightForwarder = "佐川急便"
	Case "2"
		wFreightForwarder = "ヤマト運輸"
	Case "3"
		wFreightForwarder = "福山通運"
	Case "4"
		wFreightForwarder = "自社配送"
	Case "5"                                            '2011/06/29 an add
		wFreightForwarder = "西濃運輸"
	Case Else
		wMsg = wMsg & "配送情報がありません。<BR />"	'2011/04/11 hn add
End Select

wDeliveryDt= RSv("指定納期")
wDeliveryTm = RSv("時間指定")

If Trim(RSv("営業所止めフラグ")) = "Y" Then
	wEigyoushoDome = "運送会社営業所止め"
Else
	wEigyoushoDome = ""
End If

Select Case Trim(RSv("一括出荷フラグ"))
	Case "Y"
		wIkkatsu = "商品が全て揃ってから一括出荷いたします｡"
	Case "N"
		wIkkatsu = "在庫のある商品から出荷いたします｡"
	Case Else
		wIkkatsu = ""
End Select

'離島チェック
Call check_ritou(wShipZip)
If wRitouFl = "Y" Then
	wIkkatsu = "お届先が遠隔地のため、商品が全て揃ってから一括出荷いたします｡"
End If

'---- 支払情報
wPaymentMethod = RSv("支払方法")
wFurikomiNm = RSv("振込名義人")
wRebateFl = RSv("リベート使用フラグ")
vLoanDownPayment = ""

vPaymentMethodDisp = wPaymentMethod
if vPaymentMethodDisp = "コンビニ支払" then
	vPaymentMethodDisp = "ネットバンキング・ゆうちょ・コンビニ払い"
end if

if wPaymentMethod = "銀行振込" OR wPaymentMethod = "代引き" OR wPaymentMethod = "ローン" OR wPaymentMethod = "コンビニ支払" then
else
	wMsg = wMsg & "支払方法がありません。<BR />"
end if

'---- 受注明細情報表示 + 送料計算
Call display_order_detail() '受注明細情報表示+送料計算

'---- HTML出力
'------ 配送指定
wHaisouHtml = ""
wHaisouHtml = wHaisouHtml & "    <tr>" & vbNewLine
wHaisouHtml = wHaisouHtml & "      <td colspan='3' class='preview'>" & vbNewLine
wHaisouHtml = wHaisouHtml & "        <dl class='delivery'>" & vbNewLine
wHaisouHtml = wHaisouHtml & "          <dt>配送指定</dt>" & vbNewLine
wHaisouHtml = wHaisouHtml & "          <dd>〒" & wShipZip & "<br>" & wShipPrefecture & wShipAddress & "<br>Tel. " & wShipTel & "<br>" & wShipNm & " 様</dd>" & vbNewLine
wHaisouHtml = wHaisouHtml & "        </dl>" & vbNewLine
wHaisouHtml = wHaisouHtml & "      </td>" & vbNewLine
wHaisouHtml = wHaisouHtml & "      <td colspan='3' class='preview'>" & vbNewLine

If IsNull(wDeliveryDt) = False Then
	wHaisouHtml = wHaisouHtml & "        <dd>配達日指定　" & wDeliveryDt & "</dd>" & vbNewLine
End If
If wDeliveryTm <> "" Then
	wHaisouHtml = wHaisouHtml & "        <dd>時間帯指定　" &  wDeliveryTm & "</dd>" & vbNewLine
End If
If wEigyoushoDome <> "" Then
	wHaisouHtml = wHaisouHtml & "        <dd>" & wEigyoushoDome & "</dd>" & vbNewLine
End If
If wIkkatsu <> "" Then
	wHaisouHtml = wHaisouHtml & "        <dd>" & wIkkatsu & "</dd>" & vbNewLine
End If
wHaisouHtml = wHaisouHtml & "      </td>" & vbNewLine
wHaisouHtml = wHaisouHtml & "    </tr>" & vbNewLine

'------ 支払方法
If wPaymentMethod = "ローン" Then
	If RSv("ローン頭金ありフラグ") = "Y" Then
		vLoanDownPayment = "(頭金あり)"
	Else
		vLoanDownPayment = "(頭金なし)"
	End If
End If

wPaymentHtml = ""
wPaymentHtml = wPaymentHtml & "    <tr>" & vbNewLine
wPaymentHtml = wPaymentHtml & "      <td colspan='6' class='preview'>" & vbNewLine
wPaymentHtml = wPaymentHtml & "        <dl class='delivery'>" & vbNewLine
wPaymentHtml = wPaymentHtml & "          <dt>お支払い方法</dt>" & vbNewLine

If wRebateFl = "Y" And wOrderAmTotal = 0 Then
	wPaymentHtml = wPaymentHtml & "          <dd>お支払い不要</dd>" & vbNewLine
Else
	wPaymentHtml = wPaymentHtml & "          <dd>" & vPaymentMethodDisp & vLoanDownPayment & "</dd>" & vbNewLine
End If

Select Case wPaymentMethod
	Case "銀行振込"
		wPaymentHtml = wPaymentHtml & "          <dd>振込名義　" & wFurikomiNm & "</dd>" & vbNewLine
	Case "ローン"
		If RSv("ローン頭金ありフラグ") = "Y" Then
			wPaymentHtml = wPaymentHtml & "          <dd>ローン頭金　　　" & FormatNumber(Ccur(RSv("ローン頭金")), 0) & "</dd>" & vbNewLine
		End If
		If RSv("希望ローン回数") <> 0 Then
			wPaymentHtml = wPaymentHtml & "          <dd>希望ローン回数　" & RSv("希望ローン回数") & "</dd>" & vbNewLine
		End If
		If RSv("ローン金額") <> "0" Then
			wPaymentHtml = wPaymentHtml & "          <dd>月額支払金額　　" & FormatNumber(Ccur(RSv("ローン金額")), 0) & "</dd>" & vbNewLine
		End If
		If RSv("オンラインローン申込フラグ") = "Y" Then
			wPaymentHtml = wPaymentHtml & "          <dd>オンラインでローンの申し込みを行う。（" & RSv("ローン会社") & "）</dd>" & vbNewLine
		End If
End Select

wPaymentHtml = wPaymentHtml & "        </dl>" & vbNewLine
wPaymentHtml = wPaymentHtml & "      </td>" & vbNewLine
wPaymentHtml = wPaymentHtml & "    </tr>" & vbNewLine

'------ 領収書
wReceiptHtml = ""
If RSv("領収書発行フラグ") = "Y" Then
	wReceiptHtml = wReceiptHtml & "    <tr>" & vbNewLine
	wReceiptHtml = wReceiptHtml & "      <td colspan='6' class='preview'>" & vbNewLine
	wReceiptHtml = wReceiptHtml & "        <dl class='delivery'>" & vbNewLine
	wReceiptHtml = wReceiptHtml & "          <dt>領収書</dt>" & vbNewLine
	wReceiptHtml = wReceiptHtml & "          <dd>領収書宛先：" & RSv("領収書宛先") & " 様　　領収書但し書き：" & RSv("領収書但し書き") & "</dd>" & vbNewLine
	wReceiptHtml = wReceiptHtml & "        </dl>" & vbNewLine
	wReceiptHtml = wReceiptHtml & "      </td>" & vbNewLine
	wReceiptHtml = wReceiptHtml & "    </tr>" & vbNewLine
End If

'2013/10/21 GV # add start
'------ 大型商品
If wLargeItemFl = "Y" Then
	wLargeItemHtml = "    <tr>" & vbNewLine
	wLargeItemHtml = wLargeItemHtml & "      <td colspan='6' class='preview'>" & vbNewLine
	wLargeItemHtml = wLargeItemHtml & "        <dl class='delivery'>" & vbNewLine
	wLargeItemHtml = wLargeItemHtml & "          <dt>大型商品について</dt>" & vbNewLine

	'大型商品ではない商品と混在している場合
	If wNonLargeItemFl = "Y" Then
		Call getCntlMst("大型貨物","遅延メッセージ","2", wLargeItemHTMLBuff, null, null, null, null, null)
	Else
		'単品
		Call getCntlMst("大型貨物","遅延メッセージ","1", wLargeItemHTMLBuff, null, null, null, null, null)
	End If

	wLargeItemHtml = wLargeItemHtml & "<dd style='color:red;'>" & wLargeItemHTMLBuff & "</dd>" & vbNewLine
	wLargeItemHtml = wLargeItemHtml & "        </dl>" & vbNewLine
	wLargeItemHtml = wLargeItemHtml & "      </td>" & vbNewLine
	wLargeItemHtml = wLargeItemHtml & "    </tr>" & vbNewLine
End If
'2013/10/21 GV # add end

'---- 仮受注情報更新
RSv("出荷倉庫数") = wSokoCnt
RSv("商品合計金額") = wPrdctAmTotalNoTax
RSv("送料") = wShippingNoTax
RSv("代引手数料") = wCodAm
RSv("コンビニ支払手数料") = 0
RSv("外税合計金額") = wTax
RSv("受注合計金額") = wOrderAmTotal
RSv("離島フラグ") = wRitouFl

'---- リベート金額
If wRebateFl = "Y" Then
	RSv("過不足相殺金額") = wCustomerKabusokuAm
Else
	RSv("過不足相殺金額") = 0
End If

RSv("消費税率") = wSalesTaxRate
RSv("最終更新日") = Now()

RSv.Update
RSv.Close

'---- 代引きで20万円以上の注文はエラー
If wPaymentMethod = "代引き" And wOrderAmTotal > 200000 Then
	wMsg = wMsg & "代引きの場合、1回のご注文は20万円が限度となります。ご注文内容又はお支払方法を変更して下さい｡<br />"
End If

'---- コンビニ支払いで30万円以上の注文はエラー
If wPaymentMethod = "コンビニ支払" And wOrderAmTotal > 300000 Then
	wMsg = wMsg & "ネットバンキング・ゆうちょ・コンビニ払いの場合、1回のご注文は30万円が限度となります。ご注文内容又はお支払方法を変更して下さい｡<br />"		'2011/03/22 hn mod
End If

End Function

'========================================================================
'
'	Function	離島チェック
'
'		parm:		配送先郵便番号
'		return:	離島なら　wRitouFl = Y
'						離島以外　wRitouFl = N
'				離島中の離島なら　wRitouRitouFl = Y
'						離島中の離島以外　wRitouRitouFl = N
'
'========================================================================
Function check_ritou(pZip)

Dim vZip
Dim vSQL
Dim RSv

vZip = Replace(pZip, "-", "")

If vZip = "" Then
	wRitouFl  = "N"
	Exit Function
End If

'---- 離島チェック
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    郵便番号"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    離島 WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    郵便番号 = '" & vZip & "'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

If RSv.EOF = True Then
	wRitouFl = "N"
Else
	wRitouFl = "Y"
End If

RSv.Close

End Function

'========================================================================
'
'	Function	代引手数料計算
'
'		・(商品金額+送料)*消費税に応じた手数料率をコントロールマスタから取り出す
'		・代引手数料を無料にする商品金額合計(vItemNum1)を取り出す
'
'		parm:	(商品金額+送料)*消費税
'		return: 代引手数料
'
'========================================================================
Function calc_cod_am(p_total_am)

Dim i
Dim vTotalAm()
Dim vCodAm()
Dim vItemChar1
Dim vItemChar2
Dim vItemNum1
Dim vItemNum2
Dim vItemDate1
Dim vItemDate2

'---- '代引手数料
Call getCntlMst("受注","代引手数料","1", vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)
Call cf_unstring(vItemChar1, vTotalAm, ",")
Call cf_unstring(vItemChar2, vCodAm, ",")

wTotal_NoDaibikiFee = vItemNum1

For i = 0 to UBound(vTotalAm) - 1
	If CCur(vTotalAm(i)) > CCur(p_total_am) Then
		Exit For
	End If
Next

calc_cod_am = vCodAm(i)

End function

'========================================================================
'
'	Function	受注明細内容表示
'
'========================================================================
Function display_order_detail()

Dim vProductNm
Dim vPrice
Dim vBeforeRebateAm
Dim vSQL
Dim RSv
Dim vProdTermFl
Dim vInventoryCd
Dim vInventoryImage
Dim strLargeItem	'2013/10/21 GV # add

'---- 仮受注明細情報取り出し
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    a.*"
vSQL = vSQL & "  , b.本支店コード"
vSQL = vSQL & "  , b.送料区分"
vSQL = vSQL & "  , b.特定商品個口"
vSQL = vSQL & "  , b.重量商品送料"
vSQL = vSQL & "  , b.マスターカートン数"
vSQL = vSQL & "  , b.倉庫指定なしフラグ"
vSQL = vSQL & "  , b.ASK商品フラグ"
vSQL = vSQL & "  , b.取扱中止日"
vSQL = vSQL & "  , b.廃番日"
vSQL = vSQL & "  , b.完売日"
vSQL = vSQL & "  , b.希少数量"
vSQL = vSQL & "  , b.セット商品フラグ"
vSQL = vSQL & "  , b.メーカー直送取寄区分"
vSQL = vSQL & "  , b.Web納期非表示フラグ"
vSQL = vSQL & "  , b.入荷予定未定フラグ"
vSQL = vSQL & "  , b.B品フラグ"
vSQL = vSQL & "  , b.個数限定数量"
vSQL = vSQL & "  , b.個数限定受注済数量"
vSQL = vSQL & "  , b.空輸禁止フラグ "					'2013/10/21 GV # add
vSQL = vSQL & "  , b.代引不可フラグ "					'2013/10/21 GV # add
vSQL = vSQL & "  , c.引当可能数量"
vSQL = vSQL & "  , c.引当可能入荷予定日"
vSQL = vSQL & "  , c.B品引当可能数量"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    仮受注明細 a WITH (NOLOCK)"
vSQL = vSQL & "  , Web商品 b WITH (NOLOCK)"
vSQL = vSQL & "  , Web色規格別在庫 c WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "        a.SessionID = '" & gSessionID & "'"
vSQL = vSQL & "    AND b.メーカーコード = a.メーカーコード"
vSQL = vSQL & "    AND b.商品コード = a.商品コード"
vSQL = vSQL & "    AND c.メーカーコード = a.メーカーコード"
vSQL = vSQL & "    AND c.商品コード = a.商品コード"
vSQL = vSQL & "    AND c.色 = a.色"
vSQL = vSQL & "    AND c.規格 = a.規格"
vSQL = vSQL & " ORDER BY"
vSQL = vSQL & "     a.受注明細番号"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

wProductHtml = ""

'---- 明細HTML作成
If RSv.EOF = True Then
	'-- データがない
	wProductHtml = wProductHtml & "    <tr>" & vbNewLine
	wProductHtml = wProductHtml & "      <td align=""center""><b>カートに商品がありません。</b></td>" & vbNewLine
	wProductHtml = wProductHtml & "    </tr>" & vbNewLine
Else
	'-- データがある
	'----- 見出し
	wProductHtml = wProductHtml & "    <tr>" & vbNewLine
	wProductHtml = wProductHtml & "      <th class='maker'>メーカー</th>" & vbNewLine
	wProductHtml = wProductHtml & "      <th class='name'>商品名</th>" & vbNewLine
	wProductHtml = wProductHtml & "      <th class='stock'>在庫</th>" & vbNewLine
	wProductHtml = wProductHtml & "      <th class='price'>単価</th>" & vbNewLine
	wProductHtml = wProductHtml & "      <th class='number'>数量</th>" & vbNewLine
	wProductHtml = wProductHtml & "      <th class='amount'>金額(税込)</th>" & vbNewLine
	wProductHtml = wProductHtml & "    </tr>"

	wPrdctAmTotal = 0
	wPrdctAmTotalNoTax = 0

	'---- 商品明細部
	Do Until RSv.EOF = True

		'---- 2013.10.21 GV # add start
		'---- 大型商品の表示
		strLargeItem = ""
		If (((IsNull(RSv("空輸禁止フラグ")) = False) And (RSv("空輸禁止フラグ") = "Y")) And _
			((IsNull(RSv("代引不可フラグ")) = False) And (RSv("代引不可フラグ") = "Y")) And _
			(RSv("送料区分") = "重量商品")) Then
			strLargeItem = strLargeItem & "<br><span style='color:red;'>大型商品</span>"
			wLargeItemFl = "Y"
		Else
			wNonLargeItemFl = "Y"
		End If
		'---- 2013.10.21 GV # add start

		vProductNm = RSv("商品名")
		If Trim(RSv("色")) <> "" Then
			vProductNm = vProductNm & "/" & RSv("色")
		End If
		If Trim(RSv("規格")) <> "" Then
			vProductNm = vProductNm & "/" & RSv("規格")
		End If

		vPrice = calcPrice(RSv("受注単価"), wSalesTaxRate)
		wPrdctAmTotal = wPrdctAmTotal + (vPrice * RSv("受注数量"))
		wPrdctAmTotalNoTax = wPrdctAmTotalNoTax + (Fix(RSv("受注単価")) * RSv("受注数量"))

		wProductHtml = wProductHtml & "    <tr>" & vbNewLine
		wProductHtml = wProductHtml & "      <td>" & RSv("メーカー名") & "</td>" & vbNewLine
'---- 2013/10/21 GV # mod start
'		wProductHtml = wProductHtml & "      <td><a href='" & g_HTTP & "shop/ProductDetail.asp?Item=" & RSv("メーカーコード") & "^" & Server.URLEncode(RSv("商品コード")) & "^" & RSv("色") & "^" & RSv("規格") & "' alt=''>" & vProductNm & "</a></td>" & vbNewLine
		wProductHtml = wProductHtml & "      <td><a href='" & g_HTTP & "shop/ProductDetail.asp?Item=" & RSv("メーカーコード") & "^" & Server.URLEncode(RSv("商品コード")) & "^" & RSv("色") & "^" & RSv("規格") & "' alt=''>" & vProductNm & "</a>" & strLargeItem & "</td>" & vbNewLine
'---- 2013/10/21 GV # mod end
		
		'------------- 在庫
		vProdTermFl = "N"
		If IsNull(RSv("取扱中止日")) = False Then	'取扱中止
			vProdTermFl = "Y"
		End If
		If IsNull(RSv("廃番日")) = False And RSv("引当可能数量") <= 0 Then	'廃番で在庫無し
			vProdTermFl = "Y"
		End If
		If IsNull(RSv("完売日")) = False Then		'完売商品
			vProdTermFl = "Y"
		End If

		'---- 在庫状況
		vInventoryCd = GetInventoryStatus(RSv("メーカーコード"), RSv("商品コード"), RSv("色"), RSv("規格"), RSv("引当可能数量"), RSv("希少数量"), RSv("セット商品フラグ"), RSv("メーカー直送取寄区分"), RSv("引当可能入荷予定日"), vProdTermFl)

		'---- 在庫状況、色を最終セット
		Call GetInventoryStatus2(RSv("引当可能数量"), RSv("Web納期非表示フラグ"), RSv("入荷予定未定フラグ"), RSv("廃番日"), RSv("B品フラグ"), RSv("B品引当可能数量"), RSv("個数限定数量"), RSv("個数限定受注済数量"), vProdTermFl, vInventoryCd, vInventoryImage)

		'----- 在庫状況表示
		If IsNull(RSv("取扱中止日")) = False Or _
		   IsNull(RSv("完売日")) = False Or _
		   (RSv("B品フラグ") = "Y" And RSv("B品引当可能数量") <= 0) Or _
		   (IsNull(RSv("廃番日")) = False And RSv("引当可能数量") <= 0) Then
			wProductHtml = wProductHtml & "      <td><span class='stock'>&nbsp</span></td>" & vbNewLine
		Else
			'---- 完売御礼でない場合のみ、在庫状況を表示
			wProductHtml = wProductHtml & "      <td><span class='stock'><img src='images/" & vInventoryImage & "' alt=''>" & vInventoryCd & "</span></td>" & vbNewLine
		End If

		wProductHtml = wProductHtml & "      <td>" & FormatNumber(vPrice, 0) & "円</td>"& vbNewLine
		wProductHtml = wProductHtml & "      <td>" & RSv("受注数量") & "</td>" & vbNewLine
		wProductHtml = wProductHtml & "      <td>" & FormatNumber(vPrice * RSv("受注数量"), 0) & "円</td>" & vbNewLine
		wProductHtml = wProductHtml & "    </tr>" & vbNewLine

		RSv.MoveNext
	Loop

	'---- 商品合計金額
	wProductHtml = wProductHtml & "    <tr>" & vbNewLine
	wProductHtml = wProductHtml & "      <td colspan='6'>" & vbNewLine
	wProductHtml = wProductHtml & "        <dl class='total'>" & vbNewLine
	wProductHtml = wProductHtml & "          <dt>商品合計（税込）</dt><dd>" & FormatNumber(wPrdctAmTotal, 0) & "円</dd>" & vbNewLine

	'---- 送料計算
	Call fCalcShipping(gSessionID, "通常", wShippingNoTax, wFreightForwarder, wSokoCnt, wKoguchi)
	vPrice = Fix(wShippingNoTax * (100 + wSalesTaxRate) / 100)

	If wRitouFl = "Y" Then
		wProductHtml = wProductHtml & "          <dt>送料（税込）（遠隔地）</dt><dd>" & FormatNumber(vPrice, 0) & "円</dd>" & vbNewLine
	Else
		wProductHtml = wProductHtml & "          <dt>送料（税込）</dt><dd>" & FormatNumber(vPrice, 0) & "円</dd>" & vbNewLine
	End If

	'---- 代引手数料計算
	wCodAm = 0
	If wPaymentMethod = "代引き" THen
		wCodAm = calc_cod_am((wPrdctAmTotal + wShippingNoTax) * (wSalesTaxRate + 100) / 100) * wSokoCnt

		if CCur(wPrdctAmTotal) >= CCur(wTotal_NoDaibikiFee) then
			wCodAm = 0
		end if
	End If

	'---- リベート金額計算
	'---- 手数料は無視してチェックし、支払い金額が0円になったら支払方法法をなしに（表示のみ）
	wOrderAmTotal = wPrdctAmTotal + ((wShippingNoTax + wCodAm) * (wSalesTaxRate + 100) / 100)
	vBeforeRebateAm = wOrderAmTotal

	If wRebateFl = "Y" Then
		' 入金過不足金額 ≧ 代引手数料なしの受注金額（商品金額＋送料）
		If wCustomerKabusokuAm >= (wOrderAmTotal - (wCodAm * (wSalesTaxRate + 100) / 100)) Then
			wCustomerKabusokuAm = (wOrderAmTotal - (wCodAm * (wSalesTaxRate + 100) / 100))
			vBeforeRebateAm = wCustomerKabusokuAm
			wOrderAmTotal = 0
			wCodAm = 0
		Else
			wOrderAmTotal = wOrderAmTotal - wCustomerKabusokuAm
		End If
	End If

	'---- 代引手数料
	If wPaymentMethod = "代引き" Then
		vPrice = Fix(wCodAm * (100 + wSalesTaxRate) / 100)
		If wCodAm = 0 Then
			wProductHtml = wProductHtml & "          <dt>代引手数料（税込）</dt><dd>" & "無料</dd>" & vbNewLine
		Else
			wProductHtml = wProductHtml & "          <dt>代引手数料（税込）</dt><dd>" & FormatNumber(vPrice, 0) & "円</dd>" & vbNewLine
		End If
	End If

	'------------- 購入合計
	wProductHtml = wProductHtml & "          <dt>ご購入合計金額（税込）</dt><dd>" & FormatNumber(vBeforeRebateAm,0) & "円</dd>" & vbNewLine

	'------------- 消費税
	wTax = vBeforeRebateAm - (wPrdctAmTotalNoTax + wShippingNoTax + wCodAm)
	wProductHtml = wProductHtml & "          <dt class='normalweight'>内消費税</dt><dd>" & FormatNumber(wTax, 0) & "円</dd>" & vbNewLine

	'---- リベート
	If wRebateFl = "Y" Then
		vPrice = wCustomerKabusokuAm * -1
		wProductHtml = wProductHtml & "          <dt class='credit'>クレジット／過不足金</dt><dd>" &  FormatNumber(vPrice,0) & "円</dd>" & vbNewLine

		'------------- 支払合計
		wProductHtml = wProductHtml & "          <dt>お支払い合計金額（税込）</dt><dd>" & FormatNumber(wOrderAmTotal, 0) & "円</dd>" & vbNewLine

		wProductHtml = wProductHtml & "        </dl>" & vbNewLine

		'------------- リベート使用メッセージ
		wProductHtml = wProductHtml & "        <div class='contact'>上記クレジット/過不足金は、このご注文・お見積りのみに充当されます。<br>キャンセルしてご利用にならない場合は弊社営業宛までご連絡ください。</div>" & vbNewLine
	Else
		wProductHtml = wProductHtml & "        </dl>" & vbNewLine
	End If

	wProductHtml = wProductHtml & "      </td>" & vbNewLine
	wProductHtml = wProductHtml & "    </tr>" & vbNewLine

End If

RSv.Close

End function
'========================================================================
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<title>ご注文内容の確認｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/shop.css" type="text/css">
<link rel="stylesheet" href="style/StyleOrder.css?20120629a" type="text/css">
<script type="text/javascript">
//=====================================================================
//	Next onClick
//=====================================================================
function next_onClick(){

	if (document.f_data.payment_method.value == "クレジットカード"){
		document.f_data.action = "OrderCardEnter.asp";
	}else{
		document.f_data.action = "OrderProcessing.asp";
	}
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
      <li class="now">ご注文内容の確認</li>
    </ul>
  </div></div></div>

  <h1 class="title">ご注文内容の確認</h1>
  <ol id="step">
    <li><img src="images/step01.gif" alt="1.ショッピングカート" width="170" height="50"></li>
    <li><img src="images/step02.gif" alt="2.お届け先、お支払方法の選択" width="170" height="50"></li>
    <li><img src="images/step03_now.gif" alt="3.ご注文内容の確認" width="170" height="50"></li>
    <li><img src="images/step04.gif" alt="4.ご注文完了" width="170" height="50"></li>
  </ol>

  <h2 class="cart_title">カート内容</h2>
  <table id="cart" class="confirm">
            <!---- 注文商品一覧 start ---->
<% = wProductHtml %>
            <!---- 注文商品一覧 end ---->
            <!-- 配送指定 start -->
<% = wHaisouHtml %>
            <!-- 配送指定 end -->
            <!-- 支払方法 start -->
<% = wPaymentHtml %>
            <!-- 支払方法 end -->
            <!-- 領収書 start -->
<% = wReceiptHtml %>
            <!-- 領収書 end -->
            <!-- 大型商品 start -->
<% = wLargeItemHtml %>
            <!-- 大型商品 end -->
  </table>

  <div id="btn_box">
    <ul class="btn">
      <li><a href="OrderInfoEnter.asp"><img src="images/btn_fix.png" alt="内容を変更する" class="opover"></a></li>
      <li class="last"><a href="JavaScript:next_onClick();"><img src="images/btn_send.png" alt="送　信" class="opover"></a></li>
    </ul>
  </div>

  <p class="caution">※送信ボタンを2度押さないようにお願いします。</p>
  <ul class="info left">
    <li><a href="../guide/change.asp">ご注文商品のキャンセル・返品について</a></li>
    <li><a href="../guide/nouki.asp">商品の納期についてはこちら</a></li>
  </ul>

  <form method="post" name="f_data" action="">
    <input type="hidden" name="OrderTotalAm" value="<% = wOrderAmTotal %>">
    <input type="hidden" name="payment_method" value="<% = wPaymentMethod %>">
    <input type="hidden" name="Skey" value="<% = Skey %>">
  </form>

<!--/#contents --></div>
	<div id="globalSide">
	<!--#include file="../Navi/NaviSide.inc"-->
	<!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>