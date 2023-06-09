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
'	オーダーページ
'更新履歴
'2004/12/20 カード有効期限に/がないデータの対策
'2005/01/04 不特定商品の注残出荷をまとめる変更(不特定一括対応)
'2005/01/27 見積機能削除に関した修正
'2005/04/05 カード情報を保存するチェックボックス追加
'2005/04/25 このページからリンクしている画面を別Windowで開くように変更
'2005/06/20 オリコローン追加
'2005/06/28 通信欄の備考変更
'2005/06/29 通信欄を削除
'2005/07/06 ｢離島｣を｢遠隔地｣に変更
'2005/09/20 領収書但し書きコメント変更
'2005/11/17 カード入力欄を1つにまとめる
'2005/11/18 InputのValueに｢’｣がない記述を訂正　Valueに空白が入るとそこで切れるため
'2006/01/09 ローンラジオボタンを初期チェックしないように変更
'2006/06/14 オリコローンを削除
'2006/06/28 Hiddenでもっていたカード番号を削除
'2006/06/29 オリコローン復活
'2006/06/30 カードコメント変更
'2006/08/11 オリコローン削除
'2006/10/24 コンビニ決済追加
'2006/12/08 コンビニ決済コメント変更
'2006/12/15 領収書コメント変更
'2007/01/12 「コンビニ支払」を「コンビニ/郵便局支払」に表示変更
'2007/01/22 住所検索処理からの呼び出し時は都道府県のみをセット
'2007/03/20 佐川時間指定を変更
'2007/08/14 カードエラー時のメッセージ変更
'2007/09/10 希望時間帯にタイトル変更
'2007/12/11 一括、分割出荷選択を常に表示
'2008/04/14 リベート機能追加、カード情報取り出し部分削除
'2008/05/14 HTTPSチェック対応
'2008/05/23 入力データチェック強化（LEFT, Numeric, EOF他)
'2008/08/28 「コンビニ/郵便局支払」を「コンビニエンスストア/ゆうちょ銀行支払い/Pay-easy」に表示変更
'2008/09/02 「コンビニエンスストア/ゆうちょ銀行支払い/Pay-easy」を「ネットバンキング・郵貯・コンビニ払い」に表示変更
'2009/04/21 JACCSオンラインローン追加
'
'========================================================================

On Error Resume Next

Dim w_sessionID
Dim userID
Dim userName
Dim msg

Dim CardErrorCd

Dim customer_nm
Dim furigana
Dim customer_email
Dim zip
Dim prefecture
Dim address
Dim telephone
Dim fax

Dim payment_method
Dim furikomi_nm
Dim loan_downpayment_fl
Dim loan_downpayment_am
Dim loan_term
Dim loan_am
Dim loan_apply_fl
Dim loan_company

Dim ship_address_no
Dim ship_name
Dim ship_zip
Dim ship_prefecture
Dim ship_address
Dim ship_telephone
Dim ship_invoice_fl
Dim freight_forwarder
Dim delivery_mm
Dim delivery_dd
Dim delivery_tm
Dim eigyousho_dome_fl

Dim receipt_fl
Dim receipt_nm
Dim receipt_memo
Dim receipt_nm_org
Dim receipt_memo_org

Dim CustomerClass
Dim KabusokuAm
Dim RebateFl

Dim ikkatsu_fl

Dim i_tokuchuu_fl
Dim i_toriyose_fl
Dim i_daibiki_fuka_fl

Dim wSalesTaxRate
Dim wPrice
Dim wNoData
Dim wShipAddressHTML
Dim wOrderProductHTML

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim Connection
Dim RS
Dim RS_customer
Dim RS_order

Dim w_sql
Dim w_html
Dim wMSG

'========================================================================

Response.Expires = -1			' Do not cache

'---- UserID 取り出し
userID = Session("userID")
userName = Session("userName")
w_sessionID = Session.SessionID

'---- Get input data
msg = Session.contents("msg")
wMSG = Session.contents("msg")
Session("msg") = ""
CardErrorCd = ReplaceInput(Request("CardErrorCd"))

if msg = "CardError1" then
	wMSG = "<font size='+1'><b>カードでの処理ができませんでした。</b></font><br><br>"
	wMSG = wMSG & "下記のエラーコードをご参照の上、再度処理をやり直して頂くか、別のカード他のお支払方法にて再度ご注文をお願いします。<br>"
	wMSG = wMSG & "尚、カード会社に直接御問い合わせの際は、エラーの内容、ご注文された時刻（" & fFormatDate(Now()) & " " & fFormatTime(Now()) & "）も併せてお伝えください。<br><br>"
	wMSG = wMSG & "<font size='+1'><b>エラーコード:" & CardErrorCd & "</b></font><br><br>"
end if

if msg = "CardError2" then
	wMSG = "<font size='+1'><b>カードの処理が正常に実行できませんでした。</b></font><br><br>"
	wMSG = wMSG & "ご入力されたカード情報をご確認の上、再度ご注文頂くか、別のカード、または他のお支払方法にてご注文願います。<br>"
	wMSG = wMSG & "尚、どうしても御注文が実行できない場合は、下記のエラーコードを明記の上、弊社システム担当までお問い合わせください。<br><br>"
	wMSG = wMSG & "<font size='+1'><b>エラーコード:" & CardErrorCd & "</b></font><br><br>"
end if

'---- Execute main
call connect_db()
call main()
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
'---- 消費税率取出し
call getCntlMst("共通","消費税率","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'消費税率
wSalesTaxRate = Clng(wItemNum1)

payment_method = ReplaceInput(Request("payment_method"))
ship_address_no = ReplaceInput(Request("ship_address_no"))

if isNumeric(ship_address_no) = false then
	ship_address_no = ""
end if

wNoData = false
call get_customer()				'顧客情報の取り出し
call get_order()					'仮受注情報の取り出し
call createOrderHtml()		'注文商品一覧HTML作成
call get_todokesaki()			'顧客届先情報の取り出し

End function

'========================================================================
'
'	Function	顧客情報の取り出し
'
'========================================================================
'
Function get_customer()

'---- 顧客情報取り出し
w_sql = ""
w_sql = w_sql & "SELECT a.顧客名"
w_sql = w_sql & "       , a.顧客フリガナ"
w_sql = w_sql & "       , a.顧客E_mail1"
w_sql = w_sql & "       , a.入金過不足金額"
w_sql = w_sql & "       , a.顧客クラス"
w_sql = w_sql & "       , b.顧客郵便番号"
w_sql = w_sql & "       , b.顧客都道府県"
w_sql = w_sql & "       , b.顧客住所"
w_sql = w_sql & "       , c.顧客電話番号"
w_sql = w_sql & "      , d.顧客電話番号 AS FAX"
w_sql = w_sql & "  FROM Web顧客 a WITH (NOLOCK)"
w_sql = w_sql & "     , Web顧客住所 b WITH (NOLOCK) LEFT JOIN Web顧客住所電話番号 d WITH (NOLOCK)"
w_sql = w_sql & "                                          ON d.顧客番号 = b.顧客番号"
w_sql = w_sql & "                                         AND d.住所連番 = b.住所連番"
w_sql = w_sql & "                                         AND d.電話区分 = 'FAX'"
w_sql = w_sql & "     , Web顧客住所電話番号 c WITH (NOLOCK)"
w_sql = w_sql & " WHERE a.顧客番号 = " & userID
w_sql = w_sql & "   AND b.顧客番号 = a.顧客番号"
w_sql = w_sql & "   AND b.住所連番 = 1"
w_sql = w_sql & "   AND c.顧客番号 = a.顧客番号"
w_sql = w_sql & "   AND c.住所連番 = 1"
w_sql = w_sql & "   AND c.電話連番 = 1"
	  
'@@@@@response.write(w_sql)

Set RS_customer = Server.CreateObject("ADODB.Recordset")
RS_customer.Open w_sql, Connection, adOpenStatic

if RS_customer.EOF = true then
	wMSG = "<center><font color='#ff0000'>顧客情報がありません。</font></center>"
	Session("msg") = wMSG
else
	customer_nm = RS_customer("顧客名")
	furigana = RS_customer("顧客フリガナ")
	customer_email = RS_customer("顧客E_mail1")
	KabusokuAm = RS_customer("入金過不足金額")
	CustomerClass = RS_customer("顧客クラス")
	zip = RS_customer("顧客郵便番号")
	prefecture = RS_customer("顧客都道府県")
	address = RS_customer("顧客住所")
	telephone = RS_customer("顧客電話番号")

	if isNull(RS_customer("FAX")) = true then
		fax = ""
	else
		fax = RS_customer("FAX")
	end if

end if

RS_customer.close

End function

'========================================================================
'
'	Function	受注情報の取り出し
'
'========================================================================
'
Function get_order()

'----仮受注データ取り出し
w_sql = ""
w_sql = w_sql & "SELECT a.支払方法"
w_sql = w_sql & "     , a.振込名義人"
w_sql = w_sql & "     , a.ローン頭金ありフラグ"
w_sql = w_sql & "     , a.ローン頭金"
w_sql = w_sql & "     , a.希望ローン回数"
w_sql = w_sql & "     , a.ローン金額"
w_sql = w_sql & "     , a.オンラインローン申込フラグ"
w_sql = w_sql & "     , a.ローン会社"
w_sql = w_sql & "     , a.見積備考"
w_sql = w_sql & "     , a.届先住所連番"
w_sql = w_sql & "     , a.届先名前"
w_sql = w_sql & "     , a.届先郵便番号"
w_sql = w_sql & "     , a.届先都道府県"
w_sql = w_sql & "     , a.届先住所"
w_sql = w_sql & "     , a.届先電話番号"
w_sql = w_sql & "     , a.届先納品書送付可フラグ"
w_sql = w_sql & "     , a.運送会社コード"
w_sql = w_sql & "     , a.指定納期"
w_sql = w_sql & "     , a.時間指定"
w_sql = w_sql & "     , a.営業所止めフラグ"
w_sql = w_sql & "     , a.一括出荷フラグ"
w_sql = w_sql & "     , a.領収書発行フラグ"
w_sql = w_sql & "     , a.領収書宛先"
w_sql = w_sql & "     , a.領収書但し書き"
w_sql = w_sql & "     , a.リベート使用フラグ"
w_sql = w_sql & "     , b.受注明細番号"
w_sql = w_sql & "     , b.メーカーコード"
w_sql = w_sql & "     , b.商品コード"
w_sql = w_sql & "     , b.色"
w_sql = w_sql & "     , b.規格"
w_sql = w_sql & "     , b.メーカー名"
w_sql = w_sql & "     , b.商品名"
w_sql = w_sql & "     , b.受注数量"
w_sql = w_sql & "     , b.受注単価" 
w_sql = w_sql & "     , b.受注金額" 
w_sql = w_sql & "     , c.メーカー直送取寄区分"
w_sql = w_sql & "     , c.代引不可フラグ" 
w_sql = w_sql & "     , d.引当可能数量"
w_sql = w_sql & "  FROM 仮受注 a WITH (NOLOCK)"
w_sql = w_sql & "     , 仮受注明細 b WITH (NOLOCK)"
w_sql = w_sql & "     , Web商品 c WITH (NOLOCK)"
w_sql = w_sql & "     , Web色規格別在庫 d WITH (NOLOCK)"
w_sql = w_sql & " WHERE a.SessionID = '" & w_sessionID & "'"
w_sql = w_sql & "   AND b.SessionID = a.SessionID"
w_sql = w_sql & "   AND c.メーカーコード = b.メーカーコード"
w_sql = w_sql & "   AND c.商品コード = b.商品コード"
w_sql = w_sql & "   AND d.メーカーコード = b.メーカーコード"
w_sql = w_sql & "   AND d.商品コード = b.商品コード"
w_sql = w_sql & "   AND d.色 = b.色"
w_sql = w_sql & "   AND d.規格 = b.規格"
w_sql = w_sql & " ORDER BY b.受注明細番号"

'@@@@@@response.write(w_sql)

Set RS_order = Server.CreateObject("ADODB.Recordset")
RS_order.Open w_sql, Connection, adOpenStatic

if RS_order.EOF = false then
'---- ヘッダ情報セット
	payment_method = RS_order("支払方法")
	furikomi_nm = RS_order("振込名義人")

	loan_downpayment_fl = RS_order("ローン頭金ありフラグ")
	loan_downpayment_am = RS_order("ローン頭金")
	loan_term = RS_order("希望ローン回数")
	loan_am = RS_order("ローン金額")
	loan_apply_fl = RS_order("オンラインローン申込フラグ")
	loan_company = RS_order("ローン会社")

	if ship_address_no = "" then
		ship_address_no = RS_order("届先住所連番")
	end if

	ship_name = RS_order("届先名前")
	ship_zip = RS_order("届先郵便番号")
	ship_prefecture = RS_order("届先都道府県")
	ship_address = RS_order("届先住所")
	ship_telephone = RS_order("届先電話番号")
	ship_invoice_fl = RS_order("届先納品書送付可フラグ")

	freight_forwarder = RS_order("運送会社コード")
	if freight_forwarder = "" then
		freight_forwarder = "1"		'佐川 初期値
	end if

	if isNull(RS_order("指定納期")) = false then
		delivery_mm = cf_NumToChar(DatePart("m", RS_order("指定納期")),2)
		delivery_dd = cf_NumToChar(DatePart("d", RS_order("指定納期")),2)
	end if

	delivery_tm = RS_order("時間指定")

	eigyousho_dome_fl = RS_order("営業所止めフラグ")
	ikkatsu_fl = RS_order("一括出荷フラグ")

	payment_method = RS_order("支払方法")

	receipt_fl = RS_order("領収書発行フラグ")
	receipt_nm = RS_order("領収書宛先")

	if receipt_fl = "Y" then
		if receipt_nm = "" then
			receipt_nm = customer_nm
		end if
		receipt_memo = RS_order("領収書但し書き")
		if receipt_memo = "" then
			receipt_memo = "音響機器代として"
		end if
	end if
	receipt_nm_org = customer_nm
	receipt_memo_org = "音響機器代として"

	RebateFl = RS_order("リベート使用フラグ")

end if

End function

'========================================================================
'
'	Function	注文商品一覧HTML作成
'
'========================================================================
'
Function CreateOrderHtml()

Dim v_dataCnt
Dim v_product_nm
Dim vTotalAm

v_dataCnt = 0
vTotalAm = 0
w_html = ""
i_toriyose_fl = ""
i_tokuchuu_fl = ""
i_daibiki_fuka_fl = ""

'---- 明細HTML作成
if RS_order.EOF = true then
	w_html = w_html & "<table width='100%' border='0' cellspacing='1' cellpadding='0'>" & vbNewLine
	w_html = w_html & "<tr class='honbun'><td align='center'><b>カートに商品がありません。</b></td></tr>" & vbNewLine
	w_html = w_html & "</table>" & vbNewLine
	wOrderProductHTML = w_html
	wNoData = true
	exit function
end if

'----- 見出し
w_html = w_html & "<table width='100%' border='0' cellspacing='1' cellpadding='0'>" & vbNewLine
w_html = w_html & "  <tr align='center' bgcolor='#d3d3d3' class='honbun'>" & vbNewLine
w_html = w_html & "    <td>メーカー</td>" & vbNewLine
w_html = w_html & "    <td>商品名</td>" & vbNewLine
w_html = w_html & "    <td>単価(税込)</td>" & vbNewLine
w_html = w_html & "    <td>数量</td>" & vbNewLine
w_html = w_html & "    <td>金額(税込)</td>" & vbNewLine
w_html = w_html & "  </tr>" & vbNewLine

Do Until RS_order.EOF = true
	'------------- メーカー、商品名
	v_product_nm = RS_order("商品名")
	if Trim(RS_order("色")) <> "" then
		v_product_nm = v_product_nm & "/" & RS_order("色")
	end if
	if Trim(RS_order("規格")) <> "" then
		v_product_nm = v_product_nm & "/" & RS_order("規格")
	end if
	w_html = w_html & "  <tr>" & vbNewLine
	w_html = w_html & "    <td align='left' width='170' nowrap class='honbun'>" & RS_order("メーカー名") & "</td>" & vbNewLine
	w_html = w_html & "    <td align='left' nowrap><a href='" & g_HTTP & "shop/ProductDetail.asp?Item=" & RS_order("メーカーコード") & "^" & Server.URLEncode(RS_order("商品コード")) & "^" & RS_order("色") & "^" & RS_order("規格") & "' class='link'>" & v_product_nm & "</a></td>" & vbNewLine

		'------------- 単価、数量、金額
	wPrice = calcPrice(RS_order("受注単価"), wSalesTaxRate)
	vTotalAm = vTotalAm + (wPrice * RS_order("受注数量"))

	w_html = w_html & "    <td align='right' width='100' class='honbun'>" & FormatNumber(wPrice,0) & "円</td>" & vbNewLine
	w_html = w_html & "    <td align='right' width='70' class='honbun'>" & RS_order("受注数量") & "</td>" & vbNewLine
	w_html = w_html & "    <td align='right' width='130' class='honbun'>" & FormatNumber(wPrice*RS_order("受注数量"),0) & "円</td>" & vbNewLine
	w_html = w_html & "   </tr>" & vbNewLine

'----その他情報セット
	if RS_order("引当可能数量") <= 0 then		'要発注
		i_toriyose_fl = "Y"
	end if
	if RS_order("メーカー直送取寄区分") = "特注" then		'特別注文
		i_toriyose_fl = "Y"
		i_tokuchuu_fl = "Y"
	end if
	if RS_order("代引不可フラグ") = "Y" then		'代引き不可
		i_daibiki_fuka_fl = "Y"
	end if

	v_dataCnt = v_dataCnt + 1
	RS_order.MoveNext
Loop

'----商品合計金額
w_html = w_html & "  <tr bgcolor='#d3d3d3' class='honbun'>" & vbNewLine
w_html = w_html & "    <td height='2' colspan='5' align='left'><img src='images/blank.gif' width='1' height='2'></td>" & vbNewLine
w_html = w_html & "  </tr>" & vbNewLine
 
w_html = w_html & "  <tr class='honbun'>" & vbNewLine
w_html = w_html & "    <td align='left'><a href='" & g_HTTP & "shop/Order.asp'><img src='images/OrderItemUpdate.gif' width='120' height='19' border='0' align='absmiddle' alt='注文商品の変更'></a></td>" & vbNewLine
w_html = w_html & "    <td align='left'></td>" & vbNewLine
w_html = w_html & "    <td colspan='2' align='right'><b>商品合計(税込)</b></td>" & vbNewLine
w_html = w_html & "    <td align='right'><b>" & FormatNumber(vTotalAm,0) & "円</b></td>" & vbNewLine
w_html = w_html & "  </tr>" & vbNewLine

'----リベート金額表示
if KabusokuAm > 0 AND CustomerClass = "一般顧客"then
	w_html = w_html & "  <tr bgcolor='#d3d3d3' class='honbun'>" & vbNewLine
	w_html = w_html & "    <td height='2' colspan='5' align='left'><img src='images/blank.gif' width='1' height='2'></td>" & vbNewLine
	w_html = w_html & "  </tr>" & vbNewLine
	 
	w_html = w_html & "  <tr class='honbun'>" & vbNewLine
	w_html = w_html & "    <td align='left'></td>" & vbNewLine
	w_html = w_html & "    <td align='left'></td>" & vbNewLine
	w_html = w_html & "    <td colspan='2' align='right'><b>クレジット/過不足金</b></td>" & vbNewLine
	w_html = w_html & "    <td align='right'><b>" & FormatNumber(KabusokuAm,0) & "円</b></td>" & vbNewLine
	w_html = w_html & "  </tr>" & vbNewLine

	w_html = w_html & "  <tr class='honbun'>" & vbNewLine
	w_html = w_html & "    <td align='left'></td>" & vbNewLine
	w_html = w_html & "    <td align='left'></td>" & vbNewLine
	w_html = w_html & "    <td colspan='3' align='left'><input type='checkbox' name='RebateFl' value='Y' "

	if RebateFl = "Y" then
		w_html = w_html & "CHECKED"
	end if

	w_html = w_html & "><b>お支払いにクレジット/過不足金を使用する</b>" & vbNewLine
	w_html = w_html & "  </tr>" & vbNewLine
end if

w_html = w_html & "</table>" & vbNewLine

if v_dataCnt = 1 then
	i_toriyose_fl = "N"		'データが1件しかない場合は取寄せ時の一括メッセージ不要
end if

RS_order.close
wOrderProductHTML = w_html

End Function

'========================================================================
'
'	Function	顧客届先情報の取り出し
'
'========================================================================
'
Function get_todokesaki()

'---- 顧客届先情報取り出し
w_sql = ""
w_sql = w_sql & "SELECT b.住所連番" 
w_sql = w_sql & "       , b.住所名称" 
w_sql = w_sql & "       , b.顧客郵便番号" 
w_sql = w_sql & "       , b.顧客都道府県" 
w_sql = w_sql & "       , b.顧客住所" 
w_sql = w_sql & "       , c.顧客電話番号" 
w_sql = w_sql & "  FROM Web顧客住所 b WITH (NOLOCK)" 
w_sql = w_sql & "     , Web顧客住所電話番号 c WITH (NOLOCK)" 
w_sql = w_sql & " WHERE b.顧客番号 = " & userID 
w_sql = w_sql & "   AND c.顧客番号 = b.顧客番号" 
w_sql = w_sql & "   AND c.住所連番 = b.住所連番" 
w_sql = w_sql & "   AND c.電話区分 = '電話'"
w_sql = w_sql & " ORDER BY b.住所連番"
	  
'@@@@@@response.write(w_sql)

Set RS_customer = Server.CreateObject("ADODB.Recordset")
RS_customer.Open w_sql, Connection, adOpenStatic

wShipAddressHTML = ""

Do while RS_customer.EOF = false
	wShipAddressHTML = wShipAddressHTML _
							& "<option value='" & RS_customer("住所連番") & "'>" _
							& RS_customer("住所名称") _
							& "　〒" & RS_customer("顧客郵便番号") _
							& " " & RS_customer("顧客都道府県") & RS_customer("顧客住所") _
							& "　" & RS_customer("顧客電話番号") & vbNewLine
	
	RS_customer.MoveNext
Loop

RS_customer.close

wShipAddressHTML = "<select name='ship_address_no'>" & vbNewLine _
									& wShipAddressHTML _
									& "</select>"

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
'
Function close_db()

Connection.close

End function

'========================================================================
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<title>サウンドハウス  注文</title>

<!--#include file="../Navi/NaviStyle.inc"-->

<script language="JavaScript">
//=====================================================================
//	注文 onClick
//=====================================================================
function order_onClick(pEstimate){

	document.f_data.estimate_fl.value = pEstimate;
	document.f_data.action = "OrderInfoInsert.asp";
	document.f_data.submit();
}

//=====================================================================
//	届先変更 onClick
//=====================================================================
function ship_address_onClick(){

	if (document.f_data.ship_address_no.selectedIndex == 0){
		document.f_data.action = "../member/Member.asp?called_from=order";
	}else{
		document.f_data.action = "../member/MemberShipaddress.asp?called_from=order";
	}
 
	document.f_data.submit();
}

//=====================================================================
//	住所検索 onClick
//=====================================================================
function address_search_onClick(){

	var addrWin;

	if (document.f_data.ship_zip.value == ""){
		alert("郵便番号を入力して下さい。");
		return;
	}
 
	AddrWin = window.open("../comasp/address_search.asp?zip=" + document.f_data.ship_zip.value +"&name_prefecture=i_ship_prefecture&name_address=ship_address","AddrSearch","width=200,height=100");

}

//=====================================================================
//	カード
//=====================================================================
function card_onClick(){

	//希望ローン情報をクリア
	document.f_data.loan_downpayment_am.value = "";
	document.f_data.loan_term.options[0].selected = true;
	document.f_data.loan_am.value = "";
}

//=====================================================================
//	ローン
//=====================================================================
function loan_onClick(){

//	receipt_onClick();
}

//=====================================================================
//	希望ローン回数
//=====================================================================
function loan_term_onChange(){

	//カード支払いの場合選択不可
	if (document.f_data.payment_method[3].checked == true){
		alert("カードでご購入の際，希望ローン回数の設定はできません。");
		document.f_data.loan_term.options[0].selected = true;
	}
}

//=====================================================================
//	ラジオボタン、ドロップダウンリストを以前に選択した状態にする
//=====================================================================
function preset_values(pPref){

// 住所検索処理からの呼び出し時は都道府県のみをセット
	if (pPref == "pref"){
		for (var i=0; i<document.f_data.ship_prefecture.options.length; i++){
			if (document.f_data.ship_prefecture.options[i].value == document.f_data.i_ship_prefecture.value)		{
				document.f_data.ship_prefecture.options[i].selected = true;
				break;
			}
		}
		return;
	}

//	支払方法
	for (var i=0; i<document.f_data.payment_method.length; i++){
		if (document.f_data.payment_method[i].value == document.f_data.i_payment_method.value){
			document.f_data.payment_method[i].checked = true;
			break;
		}
	}

// ローン頭金
	if (document.f_data.i_payment_method.value == "ローン"){
		if (document.f_data.i_loan_downpayment_fl.value == "Y"){
			document.f_data.loan_downpayment_fl[1].checked = true;
		}
		if (document.f_data.i_loan_downpayment_fl.value == "N"){
			document.f_data.loan_downpayment_fl[0].checked = true;
		}
	}

// ローン回数/月額
	if (document.f_data.i_payment_method.value == "ローン"){
		if (document.f_data.loan_am.value != "0"){
			document.f_data.loan_term_payment[1].checked = true;
		}
	}

// ローン回数
		for (var i=0; i<document.f_data.loan_term.options.length; i++){
			if (document.f_data.loan_term.options[i].value == document.f_data.i_loan_term.value){
				document.f_data.loan_term.options[i].selected = true;
				break;
			}
		}

//	オンラインローン申込		030829 add
	if (document.f_data.i_payment_method.value == "ローン"){
		if (document.f_data.i_loan_apply_fl.value == "Y"){
			document.f_data.loan_apply_fl[0].checked = true;
		}
		if (document.f_data.i_loan_apply_fl.value == "N"){
			document.f_data.loan_apply_fl[1].checked = true;
		}
	}

//	オンラインローン会社
	if (document.f_data.i_payment_method.value == "ローン"){
		if (document.f_data.i_loan_apply_fl.value == "Y"){
			if (document.f_data.i_loan_company.value == "セントラル"){
				document.f_data.loan_company[0].checked = true;
			}
			if (document.f_data.i_loan_company.value == "ジャックス"){
				document.f_data.loan_company[1].checked = true;
			}
		}
	}

// 届先一覧
	for (var i=0; i<document.f_data.ship_address_no.options.length; i++){
		if (document.f_data.ship_address_no.options[i].value == document.f_data.i_ship_address_no.value){
			document.f_data.ship_address_no.options[i].selected = true;
			break;
		}
	}

// 都道府県
	for (var i=0; i<document.f_data.ship_prefecture.options.length; i++){
		if (document.f_data.ship_prefecture.options[i].value == document.f_data.i_ship_prefecture.value)		{
			document.f_data.ship_prefecture.options[i].selected = true;
			break;
		}
	}

// 納品書送付
	if (document.f_data.i_ship_invoice_fl.value == "Y"){
		document.f_data.ship_invoice_fl[0].checked = true;
	}
	if (document.f_data.i_ship_invoice_fl.value == "N"){
		document.f_data.ship_invoice_fl[1].checked = true;
	}

//	運送会社
	for (var i=0; i<document.f_data.freight_forwarder.options.length; i++){
		if (document.f_data.freight_forwarder.options[i].value == document.f_data.i_freight_forwarder.value){
			document.f_data.freight_forwarder.options[i].selected = true;
			break;
		}
	}

//	配達日
	freight_forwarder_onChange();			

	for (var i=0; i<document.f_data.delivery_mm.options.length; i++){
		if (document.f_data.delivery_mm.options[i].value == document.f_data.i_delivery_mm.value){
			document.f_data.delivery_mm.options[i].selected = true;
			break;
		}
	}
	for (var i=0; i<document.f_data.delivery_dd.options.length; i++){
		if (document.f_data.delivery_dd.options[i].value == document.f_data.i_delivery_dd.value){
			document.f_data.delivery_dd.options[i].selected = true;
			break;
		}
	}

//	時間指定
	for (var i=0; i<document.f_data.delivery_tm.options.length; i++){
		if (document.f_data.delivery_tm.options[i].value == document.f_data.i_delivery_tm.value){
			document.f_data.delivery_tm.options[i].selected = true;
			break;
		}
	}

// 営業所止め
	if (document.f_data.i_eigyousho_dome_fl.value == "Y"){
		document.f_data.eigyousho_dome_fl.checked = true;
	}

// 一括出荷
	if (document.f_data.i_toriyose_fl.value == "Y"){
		if (document.f_data.ikkatsu_fl.length >= 2){
			if (document.f_data.i_ikkatsu_fl.value == "Y"){
				document.f_data.ikkatsu_fl[0].checked = true;
			}
			if (document.f_data.i_ikkatsu_fl.value == "N"){
				document.f_data.ikkatsu_fl[1].checked = true;
			}
		}
	}

// 領収証
	if (document.f_data.receipt_fl.type != "hidden"){
		if (document.f_data.i_receipt_fl.value == "N"){
			document.f_data.receipt_fl[0].checked = true;
		}
		if (document.f_data.i_receipt_fl.value == "Y"){
			document.f_data.receipt_fl[1].checked = true;
		}
	}

}

//=====================================================================
//	運送会社を変更されたら、時間帯指定ドロップダウンを運送会社に合わせて変更
//=====================================================================
function freight_forwarder_onChange(){

	for (var i=0; i<document.f_data.freight_forwarder.options.length; i++){
		if (document.f_data.freight_forwarder.options[i].selected == true){
			if (document.f_data.freight_forwarder.options[i].text == "佐川急便"){
				document.f_data.delivery_tm.options.length = 6;
				document.f_data.delivery_tm.options[0].value = "";
				document.f_data.delivery_tm.options[1].value = "午前中";
				document.f_data.delivery_tm.options[2].value = "12時から14時まで";
				document.f_data.delivery_tm.options[3].value = "14時から16時まで";
				document.f_data.delivery_tm.options[4].value = "16時から18時まで";
				document.f_data.delivery_tm.options[5].value = "18時から21時まで";
				document.f_data.delivery_tm.options[0].text = "";
				document.f_data.delivery_tm.options[1].text = "午前中";
				document.f_data.delivery_tm.options[2].text = "12時から14時まで";
				document.f_data.delivery_tm.options[3].text = "14時から16時まで";
				document.f_data.delivery_tm.options[4].text = "16時から18時まで";
				document.f_data.delivery_tm.options[5].text = "18時から21時まで";
			}
			if (document.f_data.freight_forwarder.options[i].text == "ヤマト運輸"){
				document.f_data.delivery_tm.options.length = 7;
				document.f_data.delivery_tm.options[0].value = "";
				document.f_data.delivery_tm.options[1].value = "午前中";
				document.f_data.delivery_tm.options[2].value = "12時から14時";
				document.f_data.delivery_tm.options[3].value = "14時から16時";
				document.f_data.delivery_tm.options[4].value = "16時から18時";
				document.f_data.delivery_tm.options[5].value = "18時から20時";
				document.f_data.delivery_tm.options[6].value = "20時から21時";
				document.f_data.delivery_tm.options[0].text = "";
				document.f_data.delivery_tm.options[1].text = "午前中";
				document.f_data.delivery_tm.options[2].text = "12時から14時";
				document.f_data.delivery_tm.options[3].text = "14時から16時";
				document.f_data.delivery_tm.options[4].text = "16時から18時";
				document.f_data.delivery_tm.options[5].text = "18時から20時";
				document.f_data.delivery_tm.options[6].text = "20時から21時";
			}
		}
	}
	document.f_data.delivery_tm.options[0].selected = true;
}

</script>

</head>

<body background="../Navi/Images/back_ground.gif" bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<!--#include file="../Navi/NaviTop.inc"-->

<table width="940" height="26" border="0" cellpadding="0" cellspacing="0">
  <tr>

<!--#include file="../Navi/NaviLeft.inc"-->

    <td width="798" align="left" valign="top" bgcolor="#ffffff">

<!------------ ページメイン部分の記述 START ------------>

<!-- エラーメッセージ -->
<% if msg <> "" then %>

<table width="99%" border="1" cellspacing="0" cellpadding="3" bordercolor="#999999" bordercolorlight="#999999" bordercolordark="#ffffff">
  <tr align="center" valign="top">
    <td align="left" bgcolor="#D2FFFF">
      <font color = "#ff0000">
      <%=wMSG%>
      </font>

	<% if msg = "CardError1" OR msg = "CardError2" then %>
      <b>◆よくあるカードエラーについて</b><br>

      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td bgcolor="#000000">
            <table width="100%" border="0" cellpadding="2" cellspacing="1" class="ctTable">
              <tr>
                <td bgcolor="#CCCCCC" class="honbun"><p class="bold">エラーコード</p></td>
                <td bgcolor="#CCCCCC" class="honbun"><p class="bold">理由・対処方法</p></td>
              </tr>
              <tr>
                <td bgcolor="#FFFFFF" class="honbun">S0010G12</td>
                <td bgcolor="#FFFFFF" class="honbun">このカードはご利用できませんでした。<br>ご利用できない詳細理由に関しましては、カード会社へお問い合わせください。</td>
              </tr>
              <tr>
                <td bgcolor="#FFFFFF" class="honbun">S0010G65</td>
                <td bgcolor="#FFFFFF" class="honbun">入力されたカード番号に誤りがある可能性があります。<br>再度入力されるか、カード会社へお問い合わせください。</td>
              </tr>
              <tr>
                <td bgcolor="#FFFFFF" class="honbun">S0010G83</td>
                <td bgcolor="#FFFFFF" class="honbun">入力された有効期限に誤りがある可能性があります。<br>再度入力されるか、カード会社へお問い合わせください。</td>
              </tr>
              <tr>
                <td bgcolor="#FFFFFF" class="honbun">S102000C</td>
                <td bgcolor="#FFFFFF" class="honbun">3Dセキュア認証中にキャンセルをされたか、入力されたパスワードが認証できませんでした。</td>
              </tr>
              <tr>
                <td bgcolor="#FFFFFF" class="honbun">S20210A2</td>
                <td bgcolor="#FFFFFF" class="honbun">カード番号を誤入力された可能性があります。<br>再度入力されるか、カード会社へお問い合わせください。</td>
              </tr>
              <tr>
                <td bgcolor="#FFFFFF" class="honbun">S2022017</td>
                <td bgcolor="#FFFFFF" class="honbun">ご注文の際、一定時間が経過した為、タイムアウトされた可能性が考えられます。<br>お手数ですが、はじめからやり直してください。</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
	<% end if %>

    </td>
  </tr>
</table>

<% end if %>

      <form method="post" name="f_data">
      <table width="790" border="0" cellspacing="0" cellpadding="2">
        <tr class="honbun">
          <td width="5" height="5"></td>
          <td></td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top"><b><%=customer_nm%> 様</b></td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
          <font color="#666666">
          ・確認メールは自動的に送信されます。売買契約は商品の発送をもって成立となります。<br>
          ・携帯電話などの受信制限があるアドレスでご登録された場合、ご注文情報が受信できない場合がございます。あらかじめご了承下さい。<br>
          ・ご注文商品についての問合せは、メールやお電話にて承っておりますのでご注文前にご確認頂けます様お願いします。<br>
          ・カートに入れた商品以外をご希望の際は、予めメールやお電話にてお問合せ下さい。
          </font>
          </td>
        </tr>

<!---- 注文内容 -------------------->
        <tr>
          <td>&nbsp;</td>
          <td align="left" valign="top"><span class="midashi">注文内容の確認</span></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="left" valign="top">

<!--注文商品一覧-->
<%=wOrderProductHTML%>

          </td>
        </tr>

<!---- 顧客情報 -------------------->
        <tr>
          <td>&nbsp;</td>
          <td align="left" valign="top"><span class="midashi">お客様情報の確認</span></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <table width="770" border="0" cellspacing="1" cellpadding="0">
              <tr class="honbun">
                <td align="left" valign="top"><b>お名前</b></td>
                <td align="left" valign="top"><%=customer_nm%></td>
              </tr class="honbun">
              <tr class="honbun">
                <td align="left" valign="top"><b>フリガナ</b></td>
                <td align="left" valign="top"><%=furigana%></td>
              </tr>
              <tr class="honbun">
                <td align="left" valign="top"><b>〒</b></td>
                <td align="left" valign="top"><%=zip%></td>
              </tr>
              <tr class="honbun">
                <td align="left" valign="top"><b>住所</b></td>
                <td align="left" valign="top"><%=prefecture%><%=address%></td>
              </tr>
              <tr class="honbun">
                <td align="left" valign="top"><b>電話番号</b></td>
                <td align="left" valign="top"><%=telephone%></td>
              </tr>
              <tr class="honbun">
                <td align="left" valign="top"><b>FAX番号</b></td>
                <td align="left" valign="top"><%=fax%></td>
              </tr>
              <tr class="honbun">
                <td align="left" valign="top"><b>e-mail</b></td>
                <td align="left" valign="top"><%=customer_email%></td>
              </tr>
            </table>
          </td>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top" colspan="2"><a href="../member/member.asp?called_from=order"><img src="images/MemberUpdate.gif" width="120" height="19" border="0" alt='お客様情報の変更'></a></td>
        </tr>

<!---- お支払い方法 -------------------->
        <tr>
          <td>&nbsp;</td>
          <td align="left" valign="top"><span class="midashi">お支払方法の選択</span></td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <input type="radio" name="payment_method" value="銀行振込"><b>銀行振込</b>　振込人名義
            <input type="text" name="furikomi_nm" size=45  maxlength=60 value="<%=furikomi_nm%>"><br>
            <img src="images/blank.gif" alt="" width="85" height="5" align="left"><font color='#666666'>（振込名義がお客様のお名前と異なる場合のみご記入下さい。)</font>
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <input type="radio" name="payment_method" value="コンビニ支払"><b>ネットバンキング・ゆうちょ・コンビニ払い</b>
<table class="honbun">
  <tr>
    <td width=30></td>
    <td><font color='#666666'>・ローソン、ファミリーマート、サークルK、サンクス、セイコーマート、ゆうちょ銀行、ネットバンキング でお支払いただけます｡</font></td>
  </tr>
  <tr>
    <td width=30></td>
    <td><font color='#666666'>・ご注文後、金額が変更となるご注文の変更は承る事が出来ません。在庫の無い商品などをご注文の際は、事前にお問合せください。<br>
・E-MAILアドレスが携帯の場合は、必要事項が確認できない場合がある為、パソコンからのご利用をおすすめします。</font></td>
  </tr>
  <tr>
    <td width=30></td>
    <td><font color='#ff0000'>・後ほど、お見積書をご案内致しますので、こちらをご確認後にお振込み頂けます様お願い致します。</font></td>
  </tr>
</table>
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <input type="radio" name="payment_method" value="代引き"><b>代金引換</b>　<font color='#666666'>代金引換でのご購入の場合、商品の発送は一括出荷となります。また、お支払いは現金のみの受付となります。</font>
          </td>
        </tr>

        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top"><!-- <disabled="disabled"> -->
            <input type="radio" name="payment_method" value="クレジットカード" onClick="card_onClick();"><b>クレジットカード</b><br>
			<font color="#FF0000">現在クレジットカードの利用を停止させて頂いております。<br>
			ご利用のお客様には、ご迷惑をお掛けいたしますが、何卒ご了承くださいます様お願い申し上げます。</font>
			<br>
          </td>
        </tr>

		<tr class='honbun'>
		  <td width=30></td>
		  <td colspan=2><font color='#666666'>・クレジットカードでのご購入の場合、一括払いのみのお取扱となります。<br>
		  ・ご本人名義のカードのみご利用頂けます。<br>
		  ・カード会社の登録内容と今回ご登録データに相違があった場合、ご注文を承ることが出来ない場合もあります。<br>
		  ・クレジットカードは注文日に決済されます。在庫が無く、納期がかかる場合、商品をお届けする前に代金が引き落とされる事もあります。<br>　あらかじめご了承ください。</font></td>
		</tr>

        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <input type="radio" name="payment_method" value="ローン" onClick="loan_onClick();"><b>ローン</b>　
            <input type="radio" name="loan_downpayment_fl" value="N">頭金無し　/　
            <input type="radio" name="loan_downpayment_fl" value="Y">頭金あり　 頭金
            <input type="text" name="loan_downpayment_am" size=10 maxlength=6 value="<%=loan_downpayment_am%>">円
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <img src="images/blank.gif" alt="" width="20" height="20" align="left">
            <input type="radio" name="loan_apply_fl" value="Y">オンラインでローンを申し込む。

<!--        <input type="hidden" name="loan_company" value="セントラル">  -->

            <img src="images/blank.gif" alt="" width="40" height="20" align="left">

            <input type="radio" name="loan_company" value="ジャックス" checked>ジャックス　
            <input type="radio" name="loan_company" value="セントラル">セディナ　<br>
            <font color='#ff0000'>オンラインローンの場合､お申込後のご注文内容の変更を承ることができません。<br>
            ご注文内容と､オンラインローン申込フォームの内容をご確認の上ご注文ください。</font>
          </td>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <img src="images/blank.gif" alt="" width="20" height="20" align="left">
            <input type="radio" name="loan_apply_fl" value="N">オンラインを使用しない。(ローン回数または月額を指定願います。) 
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <img src="images/blank.gif" alt="" width="40" height="20" align="left">
            <input type="radio" name="loan_term_payment" value="T">希望ローン回数
            <select name="loan_term" onChange="loan_term_onChange();">
              <option value="0">
              <option value="1">1
              <option value="2">2
              <option value="3">3
              <option value="6">6
              <option value="10">10
              <option value="12">12
              <option value="15">15
              <option value="18">18
              <option value="20">20
              <option value="24">24
              <option value="30">30
              <option value="36">36
              <option value="42">42
              <option value="48">48
              <option value="54">54
              <option value="60">60
            </select>　/　
            <input type="radio" name="loan_term_payment" value="P">月額支払金額
            <input type="text" name="loan_am" size=10 maxlength=6 value="<%=loan_am%>">円
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <img src="images/blank.gif" alt="" width="40" height="20" align="left"><font color="#666666">・ローン会社によりご希望のお支払い回数を指定できない場合がございます。 </font>
          </td>
        </tr>

<!---- その他指定 -------------------->
        <tr>
          <td>&nbsp;</td>
          <td align="left" valign="top"><span class="midashi">その他指定</span></td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            ご登録住所以外へのお届け、配達日指定、配達時間帯指定、領収証の発行等特別なご指定がある場合は、以下の希望の項目を入力して[注文]ボタン押して下さい。
          </td>
        </tr>
<% if wNoData = false then %>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <a href="JavaScript:order_onClick('N');"><img src="images/Order.gif" width="120" height="19" border="0" alt='注文画面へ進む'></a>
          </td>
        </tr>
<% end if %>

<!---- 配送先指定-------------------->
        <tr>
          <td>&nbsp;</td>
          <td align="left" valign="top"><span class="midashi">お届け先の変更</span></td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">お届先を一覧の中から選択する。一覧に無い場合は下の欄へ入力して下さい。 </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top"><%=wShipAddressHTML%></td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <table width="600" border="0" cellspacing="2" cellpadding="0">
              <tr class="honbun">
                <td width="80"><b>お名前</b></td>
                <td><input type="text" name="ship_name" size=30 maxlength=60 value="<%=ship_name%>"></td>
              </tr>
              <tr class="honbun">
                <td width="80"><b>住所</b></td>
                <td>〒<input type="text" name="ship_zip" size="10" maxlength="8" value="<%=ship_zip%>">（半角）<a href="JavaScript:address_search_onClick();"><img src="images/AddressSearch.gif" width="120" height="19" border="0" alt='住所検索'></a>&nbsp;郵便番号を入力してボタンを押して下さい｡</td>
              </tr>
              <tr class="honbun">
                <td width="80"></td>
                <td>
                  <select name="ship_prefecture">
                    <option value="" SELECTED>都道府県
                    <option value="北海道">北海道
                    <option value="青森県">青森県
                    <option value="秋田県">秋田県
                    <option value="岩手県">岩手県
                    <option value="宮城県">宮城県
                    <option value="山形県">山形県
                    <option value="福島県">福島県
                    <option value="栃木県">栃木県
                    <option value="新潟県">新潟県
                    <option value="群馬県">群馬県
                    <option value="埼玉県">埼玉県
                    <option value="茨城県">茨城県
                    <option value="千葉県">千葉県
                    <option value="東京都">東京都
                    <option value="神奈川県">神奈川県
                    <option value="山梨県">山梨県
                    <option value="長野県">長野県
                    <option value="岐阜県">岐阜県
                    <option value="富山県">富山県
                    <option value="石川県">石川県
                    <option value="静岡県">静岡県
                    <option value="愛知県">愛知県
                    <option value="三重県">三重県
                    <option value="奈良県">奈良県
                    <option value="和歌山県">和歌山県
                    <option value="福井県">福井県
                    <option value="滋賀県">滋賀県
                    <option value="京都府">京都府
                    <option value="大阪府">大阪府
                    <option value="兵庫県">兵庫県
                    <option value="岡山県">岡山県
                    <option value="鳥取県">鳥取県
                    <option value="島根県">島根県
                    <option value="広島県">広島県
                    <option value="山口県">山口県
                    <option value="香川県">香川県
                    <option value="徳島県">徳島県
                    <option value="愛媛県">愛媛県
                    <option value="高知県">高知県
                    <option value="福岡県">福岡県
                    <option value="佐賀県">佐賀県
                    <option value="大分県">大分県
                    <option value="熊本県">熊本県
                    <option value="宮崎県">宮崎県
                    <option value="長崎県">長崎県
                    <option value="鹿児島県">鹿児島県
                    <option value="沖縄県">沖縄県
                  </select>
                  <input type="text" name="ship_address" size="60" maxlength="80" value="<%=ship_address%>"><br>会社名、ビル名、部屋番号、ｘｘ様方、等は忘れずご記入下さい。
                </td>
              </tr>
              <tr class="honbun">
                <td width="80"><b>電話番号</b></td>
                <td><input type="text" name="ship_telephone" size="30" maxlength="20" value="<%=ship_telephone%>">（半角数字）</td>
              </tr>
              <tr class="honbun">
                <td width="80"></td>
                <td align="left" valign="top">
                  <input type="radio" name="ship_invoice_fl" value="Y" checked>お届先に納品書を送付して良い　　　
                  <input type="radio" name="ship_invoice_fl" value="N">お届先に納品書を送付しない
                </td>
              </tr>
            </table>          
          </td>
        </tr>

<!---- 配送方法指定 -------------------->
        <tr>
          <td>&nbsp;</td>
          <td align="left" valign="top"><span class="midashi">運送会社・配送日時の指定</span></td>
        </tr>

        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
  <% if payment_method <> "代引き" then %>
            <table width="600" border="0" cellspacing="0" cellpadding="0">
              <tr class="honbun">
                <td width="80"><b>出荷指定</b></td>
                <td>
                  <input type="radio" name="ikkatsu_fl" value="Y" >商品が全て揃ってから一括出荷する　/
                  <input type="radio" name="ikkatsu_fl" value="N" checked>在庫商品のみを先に出荷する
                </td>
              </tr>
            </table>
  <% else %>
            <table width="600" border="0" cellspacing="0" cellpadding="0">
              <tr class="honbun">
                <td width="80"><b>出荷指定</b></td>
                <td>
                  <input type="hidden" name="ikkatsu_fl" value="Y">商品が全て揃ってから一括出荷となります｡
                </td>
              </tr>
            </table>
  <% end if %>
          </td>
        </tr>

        <tr>
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <table width="600" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="80" class="honbun"><b>運送会社</b></td>
                <td>
                  <span class="honbun">
                  <select name="freight_forwarder" onChange="freight_forwarder_onChange();">
                    <option value="1" SELECTED>佐川急便
                    <option value="2">ヤマト運輸
                  </select>
                  配送方法についての情報は</span><a href="<%=g_HTTP%>guide/kaimono.asp#haisou" target='_new' class='link'>｢こちら｣</a><span class="honbun">をご覧下さい｡ </span>
                </td>
              </tr>
            </table>          
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <table width="600" border="0" cellspacing="0" cellpadding="0">
              <tr class="honbun">
                <td width="80"><b>配達希望日</b></td>
                <td>
                  <select name="delivery_mm">
                    <option value="" SELECTED>
                    <option value="01">1
                    <option value="02">2
                    <option value="03">3
                    <option value="04">4
                    <option value="05">5
                    <option value="06">6
                    <option value="07">7
                    <option value="08">8
                    <option value="09">9
                    <option value="10">10
                    <option value="11">11
                    <option value="12">12
                  </select>月
                  <select name="delivery_dd">
                    <option value="" SELECTED>
                    <option value="01">1
                    <option value="02">2
                    <option value="03">3
                    <option value="04">4
                    <option value="05">5
                    <option value="06">6
                    <option value="07">7
                    <option value="08">8
                    <option value="09">9
                    <option value="10">10
                    <option value="11">11
                    <option value="12">12
                    <option value="13">13
                    <option value="14">14
                    <option value="15">15
                    <option value="16">16
                    <option value="17">17
                    <option value="18">18
                    <option value="19">19
                    <option value="20">20
                    <option value="21">21
                    <option value="22">22
                    <option value="23">23
                    <option value="24">24
                    <option value="25">25
                    <option value="26">26
                    <option value="27">27
                    <option value="28">28
                    <option value="29">29
                    <option value="30">30
                    <option value="31">31
                  </select>
                </td>
              </tr>
            </table>          
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <table width="600" border="0" cellspacing="0" cellpadding="0">
              <tr class="honbun">
                <td width="80"><b>希望時間帯</b></td>
                <td>
                  <select name="delivery_tm">
                    <option value="" SELECTED>
                  </select>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <table width="700" border="0" cellspacing="0" cellpadding="0">
              <tr class="honbun">
                <td width="80"><b>営業所止め</b></td>
                <td>
                  <input type="checkbox" name="eigyousho_dome_fl" value="Y">運送会社営業所止め　<font color="#666666">(営業所止めの場合、お届先住所の担当営業所への留め置きとなります。)</font>
                </td>
              </tr>
            </table>          
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <font color="#666666">・天候、交通状況ならびに配達業者の都合によりご希望に添えない場合がございます。予めご了承ください。<br>
          ・佐川急便の場合、時間帯指定は平日ならびにお届け先が都市部にお住まいの個人宅の場合に限り可能です。<br>
          ・お届先が遠隔地の場合は､自動的に一括出荷となります。<br>
          ・一部お取扱いできない地域がございます。詳細は担当営業にお問合せ下さい。</font>
          </td>
        </tr>

<!---- その他 ------------->
        <tr>
          <td>&nbsp;</td>
          <td align="left" valign="top"><span class="midashi">領収証</span></td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <input type="radio" name="receipt_fl" value="N" checked><b>不要</b>　
            <input type="radio" name="receipt_fl" value="Y"><b>必要</b>
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <table width="500" border="0" cellspacing="0" cellpadding="0">
              <tr class="honbun">
                <td width="80"><b>領収証宛先</b></td>
                <td><input type="text" name="receipt_nm" size=30 maxlength=60 value="<%=receipt_nm%>">様</td>
              </tr>
            </table>          
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <table width="500" border="0" cellspacing="0" cellpadding="0">
              <tr class="honbun">
                <td width="80"><b>但し書き</b></td>
                <td><input type="text" name="receipt_memo" size=30 maxlength=50 value="<%=receipt_memo%>"></td>
              </tr>
            </table>          
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <font color="#ff0000">・お支払方法が以下の場合、サウンドハウスの領収証は発行致しません。<br>
            1.　代金引換<br>
            2.　ローン<br>
            3.　コンビニ/郵便局支払<br>
            ・宛名、但書きは、指定欄に入力頂いた内容のまま作成致します。
          </td>
        </tr>
        <input type="hidden" name="receipt_nm_org" value="<%=receipt_nm_org%>">
        <input type="hidden" name="receipt_memo_org" value="<%=receipt_memo_org%>">

<!---- 注文 -------------------->
<% if wNoData = false then %>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <a href="JavaScript:order_onClick('N');"><img src="images/Order.gif" width="120" height="19" border="0" alt='注文画面へ進む'></a>
          </td>
        </tr>
<% end if %>

        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="center" valign="top">&nbsp;</td>
        </tr>
      </table>

      <input type="hidden" name="estimate_fl" value="">
      <input type="hidden" name="customer_email" value="<%=customer_email%>">
      <input type="hidden" name="telephone" value="<%=telephone%>">
      <input type="hidden" name="i_payment_method" value="<%=payment_method%>">
      <input type="hidden" name="i_loan_downpayment_fl" value="<%=loan_downpayment_fl%>">
      <input type="hidden" name="i_loan_term" value="<%=loan_term%>">
      <input type="hidden" name="i_loan_apply_fl" value="<%=loan_apply_fl%>">
      <input type="hidden" name="i_loan_company" value="<%=loan_company%>">
      <input type="hidden" name="i_ship_address_no" value="<%=ship_address_no%>">
      <input type="hidden" name="i_ship_prefecture" value="<%=ship_prefecture%>">
      <input type="hidden" name="i_ship_invoice_fl" value="<%=ship_invoice_fl%>">
      <input type="hidden" name="i_freight_forwarder" value="<%=freight_forwarder%>">
      <input type="hidden" name="i_delivery_mm" value="<%=delivery_mm%>">
      <input type="hidden" name="i_delivery_dd" value="<%=delivery_dd%>">
      <input type="hidden" name="i_delivery_tm" value="<%=delivery_tm%>">
      <input type="hidden" name="i_eigyousho_dome_fl" value="<%=eigyousho_dome_fl%>">
      <input type="hidden" name="i_ikkatsu_fl" value="<%=ikkatsu_fl%>">
      <input type="hidden" name="i_receipt_fl" value="<%=receipt_fl%>">
      <input type="hidden" name="i_tokuchuu_fl" value="<%=i_tokuchuu_fl%>">
      <input type="hidden" name="i_toriyose_fl" value="<%=i_toriyose_fl%>">
      <input type="hidden" name="i_daibiki_fuka_fl" value="<%=i_daibiki_fuka_fl%>">
      </form>

<!------------ ページメイン部分の記述 END ------------>

    </td>
  </tr>
</table>

<!--#include file="../Navi/NaviBottom.inc"-->

<!--#include file="../Navi/NaviScript.inc"-->

</body>
</html>

<script language="JavaScript">

	preset_values();

</script>
