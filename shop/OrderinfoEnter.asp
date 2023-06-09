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
'	お届け先、お支払い方法の選択
'
'2012/06/14 ok デザイン変更のため旧版を元に新規作成
'2012/08/08 nt 嵩重量品時の画面制御を追加
'2012/08/25 nt 佐川代引き禁止地域の制御機能を追加
'2012/09/05 ok 日時指定表示変更
'2012/10/25 ok 振込以外で「領収書必要」クリックでポップアップ表示追加
'2014/08/05 GV 納品書表示変更対応
'2014/08/20 GV 領収書が必要な場合に納品書不要を選択できないよう修正
'
'========================================================================
On Error Resume Next
Response.Expires = -1			' Do not cache
Response.buffer = true

'---- Session情報
Dim wUserID
Dim wMsg
Dim wErrMsg

'---- 受け渡し情報を受取る変数

'---- Web顧客情報
Dim wCustomerNm
Dim wCustomerKn
Dim wCustomerEmail
Dim wCustomerKabusokuAm
Dim wCustomerClass
Dim wCustomerZip
Dim wCustomerPref
Dim wCustomerAddress
Dim wCustomerTel
Dim wCustomerRitouFl   '2012/08/08 nt add
Dim wCustomerSagawaLTFl '2012/08/25 nt add

'---- 仮受注
Dim wPaymentMethod
Dim wShipAddressNo
Dim wFurikomiNm
Dim wShipInvoiceFl
Dim wFreightForwarder
Dim wIkkatsuFl
Dim wDeliveryMM
Dim wDeliveryDD
Dim wDeliveryTM
Dim wEigyoushoDomeFl
Dim wReceiptFl
Dim wReceiptNm
Dim wReceiptMemo
Dim wToriyoseFl
Dim wTokuchuuFl
Dim wDaibikiFukaFl
Dim wRebateFl
Dim wKuyuKinshiFl   '2012/08/08 nt add
Dim wSagawaLTFl     '2012/08/25 nt add

Dim wNoData
Dim wShipAddressHTML
Dim wErrDesc   '2011/08/01 an add

'---- お届け先リスト
Dim wAddressNoHTML					'住所連番
Dim wZipHTML						'郵便番号
Dim wAddressHTML					'住所
Dim wTelephoneNoHTML				'電話番号
Dim wAddressNameHTML				'お届け先氏名
Dim wRitouFlHTML					'離島フラグ                     2012/08/08 nt add
Dim wKuyuKinshiFlHTML				'空輸禁止フラグ：嵩重量品フラグ 2012/08/08 nt add
Dim wSagawaLTHTML					'佐川制限フラグ                 2012/08/25 nt add

Dim wShowInvoice					'2014/08/05 GV add
Dim wInvoiceDisabled				'2014/08/05 GV add
Dim wShipAddressNo1Data				'2014/08/05 GV add
Dim wSelectedShipAddressData		'2014/08/05 GV add
Dim wInvoiceChecked					'2014/08/05 GV add


'---- 配送時間指定    2011/06/29 an add
Dim wDeliveryTime01
Dim wDeliveryTime02
Dim wDeliveryTime03
Dim wDeliveryTime04
Dim wDeliveryTime05

'---- DB
Dim Connection

'=======================================================================
'	受け渡し情報取り出し
'=======================================================================
'---- Session変数
wUserID = Session("userID")
wMsg = Session.contents("msg")

'---- 受け渡し情報取り出し

Session("msg") = ""

'=======================================================================
'	Execute main
'=======================================================================
Call connect_db()
Call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "OrderinfoEnter.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
end if                                           '2011/08/01 an add e

Call close_db()

If Err.Description <> "" Then
	Response.Redirect g_HTTP & "shop/Error.asp"
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
'	Function	Main
'
'========================================================================
Function main()

wNoData = False

wInvoiceChecked = array("", "", "")		'2014/08/05 GV add

Call get_customer()				'顧客情報の取り出し
Call get_order()				'仮受注情報の取り出し
Call get_todokesaki()			'顧客届先情報の取り出し
Call get_DeliveryTime()			'配送時間帯をコントロールマスタから取り出し

End Function

'========================================================================
'
'	Function	顧客情報の取り出し
'
'========================================================================
Function get_customer()

Dim RSv
Dim vSQL
Dim bobj

'---- 顧客情報取り出し
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.顧客名 "
vSQL = vSQL & "    , a.顧客フリガナ "
vSQL = vSQL & "    , a.振込名義人 "
vSQL = vSQL & "    , a.顧客E_mail1 "
vSQL = vSQL & "    , a.入金過不足金額 "
vSQL = vSQL & "    , a.顧客クラス "
vSQL = vSQL & "    , b.顧客郵便番号 "
vSQL = vSQL & "    , b.顧客都道府県 "
vSQL = vSQL & "    , b.顧客住所 "
vSQL = vSQL & "    , c.顧客電話番号 "
vSQL = vSQL & "    , CASE WHEN d.郵便番号 IS NOT NULL THEN 'Y' ELSE 'N' END AS 離島フラグ "		'2012/08/08 nt add
vSQL = vSQL & "    , CASE WHEN e.郵便番号 IS NOT NULL THEN 'Y' ELSE 'N' END AS 佐川制限フラグ "	'2012/08/25 nt add
vSQL = vSQL & "FROM "
vSQL = vSQL & "    Web顧客                          a WITH (NOLOCK) "
vSQL = vSQL & "      INNER JOIN Web顧客住所         b WITH (NOLOCK) "
vSQL = vSQL & "        ON     b.顧客番号 = a.顧客番号 "
vSQL = vSQL & "           AND b.住所連番 = 1 "
vSQL = vSQL & "      INNER JOIN Web顧客住所電話番号 c WITH (NOLOCK) "
vSQL = vSQL & "        ON     c.顧客番号 = a.顧客番号 "
vSQL = vSQL & "           AND c.住所連番 = b.住所連番 "
vSQL = vSQL & "           AND c.電話連番 = 1 "
vSQL = vSQL & "      LEFT  JOIN ( SELECT '住所' AS 'AddrTypeHouse' ) t1 "
vSQL = vSQL & "        ON     b.住所区分 = t1.AddrTypeHouse "
'2012/08/08 nt add Start
vSQL = vSQL & "      LEFT  JOIN 離島 d "
vSQL = vSQL & "        ON     REPLACE(b.顧客郵便番号, '-', '') = d.郵便番号 "
'2012/08/08 nt add End
'2012/08/25 nt add Start
vSQL = vSQL & "      LEFT  JOIN 佐川制限 e "
vSQL = vSQL & "        ON     REPLACE(b.顧客郵便番号, '-', '') = e.郵便番号 AND e.代引不可フラグ='Y' "
'2012/08/25 nt add End
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        t1.AddrTypeHouse IS NOT NULL "
vSQL = vSQL & "    AND a.顧客番号 = " & wUserID & " "

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

If RSv.EOF = True Then
	wErrMsg = "顧客情報がありません。"
Else
	wCustomerNm = RSv("顧客名")
	If RSv("振込名義人") <> "" Then
		wFurikomiNm = RSv("振込名義人")
	Else
		wFurikomiNm = RSv("顧客フリガナ")
	End If

	'---- 半角を全角に変換		'2011/09/09 hn add
	Set bobj = Server.CreateObject("basp21")
	wFurikomiNm = bobj.StrConv(wFurikomiNm,4)

	wCustomerEmail = RSv("顧客E_mail1")
	wCustomerKabusokuAm = RSv("入金過不足金額")
	wCustomerClass = RSv("顧客クラス")
	wCustomerZip = RSv("顧客郵便番号")
	wCustomerPref = RSv("顧客都道府県")
	wCustomerAddress = RSv("顧客住所")
	wCustomerTel = RSv("顧客電話番号")

	'2012/08/08 nt add Start
	wCustomerRitouFl =  RSv("離島フラグ")
	'2012/08/08 nt add End
	'2012/08/25 nt add Start
	wCustomerSagawaLTFl =  RSv("佐川制限フラグ")
	'2012/08/25 nt add End
End If

RSv.Close

End Function

'========================================================================
'
'	Function	受注情報の取り出し
'
'========================================================================
Function get_order()

Dim RSv
Dim vSQL

'----仮受注データ取り出し
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.支払方法 "
vSQL = vSQL & "    , a.振込名義人 "
vSQL = vSQL & "    , a.届先住所連番 "
vSQL = vSQL & "    , a.届先名前 "
vSQL = vSQL & "    , a.届先郵便番号 "
vSQL = vSQL & "    , a.届先都道府県 "
vSQL = vSQL & "    , a.届先住所 "
vSQL = vSQL & "    , a.届先電話番号 "
vSQL = vSQL & "    , a.届先納品書送付可フラグ "
vSQL = vSQL & "    , a.運送会社コード "
vSQL = vSQL & "    , a.指定納期 "
vSQL = vSQL & "    , a.時間指定 "
vSQL = vSQL & "    , a.営業所止めフラグ "
vSQL = vSQL & "    , a.一括出荷フラグ "
vSQL = vSQL & "    , a.領収書発行フラグ "
vSQL = vSQL & "    , a.領収書宛先 "
vSQL = vSQL & "    , a.領収書但し書き "
vSQL = vSQL & "    , a.リベート使用フラグ "
vSQL = vSQL & "    , c.引当可能数量 "
vSQL = vSQL & "    , d.メーカー直送取寄区分 "
vSQL = vSQL & "    , d.代引不可フラグ "
'2012/08/08 nt add Start
vSQL = vSQL & "    , d.空輸禁止フラグ "
'2012/08/08 nt add End
vSQL = vSQL & "FROM "
vSQL = vSQL & "    仮受注                       AS a WITH (NOLOCK) "
vSQL = vSQL & "      INNER JOIN 仮受注明細      AS b WITH (NOLOCK) "
vSQL = vSQL & "        ON     b.SessionID      = a.SessionID "
vSQL = vSQL & "      INNER JOIN Web色規格別在庫 AS c WITH (NOLOCK) "
vSQL = vSQL & "        ON     c.メーカーコード = b.メーカーコード "
vSQL = vSQL & "           AND c.商品コード     = b.商品コード "
vSQL = vSQL & "           AND c.色             = b.色 "
vSQL = vSQL & "           AND c.規格           = b.規格 "
vSQL = vSQL & "      INNER JOIN Web商品         AS d WITH (NOLOCK) "
vSQL = vSQL & "        ON     d.メーカーコード = c.メーカーコード "
vSQL = vSQL & "           AND d.商品コード     = c.商品コード "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        a.SessionID = '" & gSessionID & "' "
vSQL = vSQL & "ORDER BY "
vSQL = vSQL & "      b.受注明細番号 "

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

If RSv.EOF = False Then

	'---- ヘッダ情報セット
	wShipAddressNo = RSv("届先住所連番")
	If isNumeric(wShipAddressNo) = False Then
		wShipAddressNo = 1
	ElseIf wShipAddressNo <= 0 Then
		wShipAddressNo = 1
	End If

	wShipInvoiceFl = RSv("届先納品書送付可フラグ")
	wPaymentMethod = RSv("支払方法")

	'---- 仮受注に振込名義人情報があれば上書き   '2011/09/09 an mod s
	if RSv("振込名義人") <> "" then
		wFurikomiNm = RSv("振込名義人")
	end if                                       '2011/09/09 an mod e

	wIkkatsuFl = RSv("一括出荷フラグ")

	wFreightForwarder = RSv("運送会社コード")
	If wFreightForwarder = "" Then
		wFreightForwarder = "5"		'西濃 初期値  '2011/06/29 an mod
	End If

	If isNull(RSv("指定納期")) = False Then
		wDeliveryMM = cf_NumToChar(DatePart("m", RSv("指定納期")),2)
		wDeliveryDD = cf_NumToChar(DatePart("d", RSv("指定納期")),2)
	End If

	wDeliveryTM = RSv("時間指定")

	wEigyoushoDomeFl = RSv("営業所止めフラグ")

	wReceiptFl = RSv("領収書発行フラグ")
	wReceiptNm = RSv("領収書宛先")
	wReceiptMemo = RSv("領収書但し書き")
	If wReceiptFl = "Y" Then
		If wReceiptNm = "" Then
			wReceiptNm = wCustomerNm
		End If
		If wReceiptMemo = "" Then
			wReceiptMemo = "音響機器代として"
		End If
	End If

	wRebateFl = RSv("リベート使用フラグ")

	wToriyoseFl = "N"
	wTokuchuuFl = "N"
	wDaibikiFukaFl = "N"

	'Do While RSv.EOF		'2011/03/04 na del
	Do Until RSv.EOF	'2011/03/04 na mod

		If RSv("引当可能数量") <= 0 Then				'要発注
			wToriyoseFl = "Y"
		End If
		If RSv("メーカー直送取寄区分") = "特注" Then	'特別注文
			wToriyoseFl = "Y"
			wTokuchuuFl = "Y"
		End If
		If RSv("代引不可フラグ") = "Y" Then				'代引き不可
			wDaibikiFukaFl = "Y"
		End If

		'2012/08/08 nt add Start
		'---- 画面制御情報をセット（空輸禁止フラグ：嵩重量品フラグ）
		If RSv("空輸禁止フラグ") = "Y" Then
			wKuyuKinshiFl = "Y"
		End If
		'2012/08/08 nt add End

		RSv.MoveNext

	Loop

Else

	wNoData = True

End If

RSv.Close

End Function

'========================================================================
'
'	Function	顧客届先情報の取り出し
'
'	Note		お届け先選択用のドロップダウンリストを生成
'
'========================================================================
Function get_todokesaki()

Dim RSv
Dim vSQL

'---- 顧客届先情報取り出し
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "    a.住所連番 "
vSQL = vSQL & "  , a.住所名称 "
vSQL = vSQL & "  , a.顧客郵便番号 "
vSQL = vSQL & "  , a.顧客都道府県 "
vSQL = vSQL & "  , a.顧客住所 "
vSQL = vSQL & "  , b.顧客電話番号 "
vSQL = vSQL & "  , CASE WHEN c.郵便番号 IS NOT NULL THEN 'Y' ELSE 'N' END AS 離島フラグ "		'2012/08/08 nt add
vSQL = vSQL & "  , CASE WHEN d.郵便番号 IS NOT NULL THEN 'Y' ELSE 'N' END AS 佐川制限フラグ "	'2012/08/25 nt add
vSQL = vSQL & "FROM "
vSQL = vSQL & "    Web顧客住所                      a WITH (NOLOCK) "
vSQL = vSQL & "      INNER JOIN Web顧客住所電話番号 b WITH (NOLOCK) "
vSQL = vSQL & "        ON     b.顧客番号 = a.顧客番号 "
vSQL = vSQL & "           AND b.住所連番 = a.住所連番 "
vSQL = vSQL & "      LEFT  JOIN ( SELECT '電話' AS 'PhoneTypeTel' ) t1 "
vSQL = vSQL & "        ON     b.電話区分 = t1.PhoneTypeTel "
'2012/08/08 nt add Start
vSQL = vSQL & "      LEFT  JOIN 離島 c "
vSQL = vSQL & "        ON     REPLACE(a.顧客郵便番号, '-', '') = c.郵便番号 "
'2012/08/08 nt add End
'2012/08/25 nt add Start
vSQL = vSQL & "      LEFT  JOIN 佐川制限 d "
vSQL = vSQL & "        ON     REPLACE(a.顧客郵便番号, '-', '') = d.郵便番号 AND d.代引不可フラグ='Y' "
'2012/08/25 nt add End
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        t1.PhoneTypeTel IS NOT NULL "
vSQL = vSQL & "    AND a.削除日 IS NULL "
vSQL = vSQL & "    AND a.顧客番号 = " & wUserID & " "
vSQL = vSQL & "ORDER BY "
vSQL = vSQL & "    a.住所連番 "

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

wAddressNoHTML = "0"						'住所連番
wZipHTML = "''"								'郵便番号
wAddressHTML = "''"							'住所
wTelephoneNoHTML = "''"						'電話番号
wAddressNameHTML = "'お届け先を変更する'"	'お届け先氏名

wRitouFlHTML = "''"							'離島フラグ 2012/08/08 nt add
wKuyuKinshiFlHTML = "''"					'空輸禁止フラグ：嵩重量品フラグ 2012/08/08 nt add
wSagawaLTHTML = "''"						'佐川制限フラグ 2012/08/25 nt add

wShipAddressHTML = "<option value=""0"">お届け先を変更する</option>"

Do While RSv.EOF = False
	'2014/08/05 GV add start
	If RSv("住所連番") = 1 Then
		wShipAddressNo1Data = RSv("顧客都道府県") & RSv("顧客住所")
	End If
	
	If RSv("住所連番") = wShipAddressNo Then
		wSelectedShipAddressData =  RSv("顧客都道府県") & RSv("顧客住所")
	End If
	'2014/08/05 GV add end

	wShipAddressHTML = wShipAddressHTML & _
					   "<option value=""" & RSv("住所連番") & """>" & _
						   RSv("住所名称") & _
						   " 〒" & RSv("顧客郵便番号") & _
						   " " & RSv("顧客都道府県") & RSv("顧客住所") & _
						   " Tel. " & RSv("顧客電話番号") & _
						   "</option>" & vbNewLine

	' JavaScript用のお届け先情報リストを作成
	wAddressNoHTML = wAddressNoHTML & "," & Replace(Replace(RSv("住所連番"),vbCR,""),vbLF,"") & ""
	wZipHTML = wZipHTML & ",'〒" & Replace(Replace(RSv("顧客郵便番号"),vbCR,""),vbLF,"") & "'"
	wAddressHTML = wAddressHTML & ",'" & Replace(Replace(RSv("顧客都道府県") & RSv("顧客住所"),vbCR,""),vbLF,"") & "'"
	wTelephoneNoHTML = wTelephoneNoHTML & ",'" & Replace(Replace(RSv("顧客電話番号"),vbCR,""),vbLF,"") & "'"
	wAddressNameHTML = wAddressNameHTML & ",'" & Replace(Replace(RSv("住所名称"),vbCR,""),vbLF,"") & "'"
	wRitouFlHTML = wRitouFlHTML & ",'" & Replace(Replace(RSv("離島フラグ"),vbCR,""),vbLF,"") & "'"			'2012/08/08 nt add
	wKuyuKinshiFlHTML = wKuyuKinshiFlHTML & ",'" & Replace(Replace(wKuyuKinshiFl,vbCR,""),vbLF,"") & "'"	'2012/08/08 nt add
	wSagawaLTHTML = wSagawaLTHTML & ",'" & Replace(Replace(RSv("佐川制限フラグ"),vbCR,""),vbLF,"") & "'"	'2012/08/25 nt add

	RSv.MoveNext

Loop

RSv.Close

wShipAddressHTML = "<select name=""select_ship_address_no"" id=""select_ship_address_no"" onChange=""changeShipAddress();"">" & vbNewLine & _
				   wShipAddressHTML & _
				   "</select>"

End Function

'========================================================================
'
'	Function	配送時間指定情報取得
'
'========================================================================
Function get_DeliveryTime()

Dim vItemChar1
Dim vItemChar2
Dim vItemNum1
Dim vItemNum2
Dim vItemDate1
Dim vItemDate2

'---- 西濃時間指定01
call getCntlMst("受注","時間指定_西濃","01", vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)
wDeliveryTime01 = vItemChar1
'---- 西濃時間指定02
call getCntlMst("受注","時間指定_西濃","02", vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)
wDeliveryTime02 = vItemChar1
'---- 西濃時間指定03
call getCntlMst("受注","時間指定_西濃","03", vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)
wDeliveryTime03 = vItemChar1
'---- 西濃時間指定04
call getCntlMst("受注","時間指定_西濃","04", vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)
wDeliveryTime04 = vItemChar1
'---- 西濃時間指定05
call getCntlMst("受注","時間指定_西濃","05", vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)
wDeliveryTime05 = vItemChar1

End Function

'========================================================================
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<title>お届け先、お支払い方法の選択｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/shop.css?20120629" type="text/css">
<link rel="stylesheet" href="style/StyleOrder.css?20120831" type="text/css">
<link rel="stylesheet" href="style/jquery.fancybox-1.3.4.css" type="text/css">
<script type="text/javascript">
//=====================================================================
//	OrderInfoInsert.aspへSubmit
//=====================================================================
function OrderSubmit(pCmd){

	document.f_data.cmd.value = pCmd;
	document.f_data.action = "OrderInfoInsert.asp";
	document.f_data.submit();

}

//=====================================================================
//	ラジオボタン、ドロップダウンリストを以前に選択した状態にする
//=====================================================================
function preset_values(){

	// 支払方法
	for (var i=0; i<document.f_data.payment_method.length; i++){
		if (document.f_data.payment_method[i].value == document.f_data.i_payment_method.value){
			document.f_data.payment_method[i].checked = true;
			break;
		}
	}

	// 支払方法変更
	checkPaymentMethod();

	// 届先一覧
	for (var i=0; i<document.f_data.select_ship_address_no.options.length; i++){
		if (document.f_data.select_ship_address_no.options[i].value == document.f_data.i_ship_address_no.value){
			document.f_data.select_ship_address_no.options[i].selected = true;
			break;
		}
	}
	changeShipAddress();

	//2014/08/05
	// 納品書送付
	if (document.f_data.ship_invoice_fl) {
		if (document.f_data.i_ship_invoice_fl.value == "Y"){
			document.f_data.ship_invoice_fl[2].checked = true;
		}
		if (document.f_data.i_ship_invoice_fl.value == "N"){
			document.f_data.ship_invoice_fl[1].checked = true;
		}
		if (document.f_data.i_ship_invoice_fl.value == "X"){
			document.f_data.ship_invoice_fl[0].checked = true;
		}
	}

//2011/06/01 if-web del start
	// 運送会社

//	for (var i=0; i<document.f_data.freight_forwarder.length; i++){
//		if (document.f_data.freight_forwarder[i].value == document.f_data.i_freight_forwarder.value){
//			document.f_data.freight_forwarder[i].checked = true;
//			break;
//		}
//	}
//2011/06/01 if-web del end

	// 運送会社指定変更
	//sel_FreightForwarder();

	// 日時指定あり、なし
	if (((document.f_data.i_delivery_mm.value != "") && (document.f_data.i_delivery_dd.value != "")) || (document.f_data.i_delivery_tm.value != "")) {
		document.f_data.delivery_fl[1].checked = true;
	} else {
		document.f_data.delivery_fl[0].checked = true;
	}

	// 日時指定変更
	checkDeliveryDate();

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

	// 時間指定
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
	if (document.f_data.i_ikkatsu_fl.value == "Y"){
		document.f_data.ikkatsu_fl[0].checked = true;
	}
	if (document.f_data.i_ikkatsu_fl.value == "N"){
		document.f_data.ikkatsu_fl[1].checked = true;
	}

	// 領収証
	if (document.f_data.receipt_fl.type != "hidden"){
		if (document.f_data.i_receipt_fl.value == "N"){
			document.f_data.receipt_fl[0].checked = true;
			// 2014/08/20
			document.getElementById('ship_invoice_fl_x').disabled = false;
		}
		if (document.f_data.i_receipt_fl.value == "Y"){
			document.f_data.receipt_fl[1].checked = true;
			// 2014/08/20
			if (document.getElementById('ship_invoice_fl_x').checked == true) {
				document.getElementById('ship_invoice_fl_n').checked = true;
			}
			document.getElementById('ship_invoice_fl_x').disabled = true;
		}
	}

	// 領収書変更
	checkReceipt();

	// 過不足金を使用する
	if (document.f_data.i_rebate_fl.value == "Y"){
		document.f_data.RebateFl.checked = true;
	}
}

//=====================================================================
//	運送会社 選択変更時  2011/06/01 mod （佐川のみに）
//=====================================================================
function sel_FreightForwarder(){

	// 佐川急便
//	if (document.getElementById('freight_forwarder_1').checked == true){
//		document.f_data.delivery_tm.options.length = 6;
//		document.f_data.delivery_tm.options[0].value = "";
//		document.f_data.delivery_tm.options[1].value = "午前中";
//		document.f_data.delivery_tm.options[2].value = "12時から14時まで";
//		document.f_data.delivery_tm.options[3].value = "14時から16時まで";
//		document.f_data.delivery_tm.options[4].value = "16時から18時まで";
//		document.f_data.delivery_tm.options[5].value = "18時から21時まで";
//		document.f_data.delivery_tm.options[0].text = "";
//		document.f_data.delivery_tm.options[1].text = "午前中";
//		document.f_data.delivery_tm.options[2].text = "12時から14時まで";
//		document.f_data.delivery_tm.options[3].text = "14時から16時まで";
//		document.f_data.delivery_tm.options[4].text = "16時から18時まで";
//		document.f_data.delivery_tm.options[5].text = "18時から21時まで";
//	}

	// ヤマト運輸
//	if (document.getElementById('freight_forwarder_2').checked == true){
//		document.f_data.delivery_tm.options.length = 7;
//		document.f_data.delivery_tm.options[0].value = "";
//		document.f_data.delivery_tm.options[1].value = "午前中";
//		document.f_data.delivery_tm.options[2].value = "12時から14時";
//		document.f_data.delivery_tm.options[3].value = "14時から16時";
//		document.f_data.delivery_tm.options[4].value = "16時から18時";
//		document.f_data.delivery_tm.options[5].value = "18時から20時";
//		document.f_data.delivery_tm.options[6].value = "20時から21時";
//		document.f_data.delivery_tm.options[0].text = "";
//		document.f_data.delivery_tm.options[1].text = "午前中";
//		document.f_data.delivery_tm.options[2].text = "12時から14時";
//		document.f_data.delivery_tm.options[3].text = "14時から16時";
//		document.f_data.delivery_tm.options[4].text = "16時から18時";
//		document.f_data.delivery_tm.options[5].text = "18時から20時";
//		document.f_data.delivery_tm.options[6].text = "20時から21時";
//	}

//	document.f_data.delivery_tm.options[0].selected = true;

}

//=====================================================================
//	支払方法 選択変更時
//=====================================================================
function checkPaymentMethod(){

	// 銀行振込の場合、振込人名義入力可、領収書必要選択可
	if(document.getElementById('radio_ginkou').checked == true){
		document.getElementById('furikomi_nm').disabled = false;
		document.getElementById('receipt_fl_y').disabled = false;

		$("#receipt1").css("display","inline");
		$("#receipt2").css("display","none");

	}else{
		document.getElementById('furikomi_nm').disabled = true;
		document.getElementById('receipt_fl_y').disabled = true;
		document.getElementById('furikomi_nm').value='';
		document.getElementById('receipt_fl_n').checked=true;
		// 領収書変更
		checkReceipt();

		$("#receipt1").css("display","none");
		$("#receipt2").css("display","inline");
	}

	// 代引の場合、在庫商品から出荷不可
	if(document.getElementById('radio_daibiki').checked == true){
		document.getElementById('ikkatsu_fl_n').disabled = true;
	}else{
		document.getElementById('ikkatsu_fl_n').disabled = false;
	}

	// 2014/08/05
	if(document.f_data.radio_daibiki.checked){
		$("#ship_invoice").css("display", "none");
		document.getElementById('ship_invoice_fl_x').disabled = true;
		document.getElementById('ship_invoice_fl_y').disabled = true;
		document.getElementById('ship_invoice_fl_n').disabled = true;
	} else {
		if (document.f_data.ship_address_no.value != 1) {
			if ($("#i_ship_address_no1").val() == $("#i_selected_ship_address").val()) {
				$("#ship_invoice").css("display", "none");
				document.getElementById('ship_invoice_fl_x').disabled = true;
				document.getElementById('ship_invoice_fl_y').disabled = true;
				document.getElementById('ship_invoice_fl_n').disabled = true;
			} else {
				if ($("#ship_invoice").css("display") == 'none') {
					$("#ship_invoice").css("display", "inline");
					// 2014/08/20
					document.getElementById('ship_invoice_fl_y').disabled = false;
					document.getElementById('ship_invoice_fl_n').disabled = false;
					if (document.getElementById('receipt_fl_n').checked == true) {
						document.getElementById('ship_invoice_fl_x').disabled = false;
						document.getElementById('ship_invoice_fl_x').checked  = true;
					} else {
						if (document.getElementById('ship_invoice_fl_x').checked == true) {
							document.getElementById('ship_invoice_fl_n').checked = true;
						}
					}
				}
			}
		} else {
			$("#ship_invoice").css("display", "none");
		}
	}

}

//=====================================================================
//	日時指定 選択変更時
//=====================================================================
function checkDeliveryDate(){

	// 日時指定ありの場合、日付、時間選択可
	if(document.getElementById('delivery_fl_y').checked == true){
		// 2012/09/06 ok Add
		if(document.getElementById('delivery_mm').disabled){
			$('a[href=#delivery_time]').click();
		}

		document.getElementById('delivery_mm').disabled = false;
		document.getElementById('delivery_dd').disabled = false;
		document.getElementById('delivery_tm').disabled = false;
	}else{
		document.getElementById('delivery_mm').disabled = true;
		document.getElementById('delivery_dd').disabled = true;
		document.getElementById('delivery_tm').disabled = true;
		document.getElementById('delivery_mm').options[0].selected=true;
		document.getElementById('delivery_dd').options[0].selected=true;
		document.getElementById('delivery_tm').options[0].selected=true;
	}

}

//=====================================================================
//	領収書指定 選択変更時
//=====================================================================
function checkReceipt(){

	// 領収書必要の場合、領収書宛先、領収書但し書き入力可
	if(document.getElementById('receipt_fl_y').checked==true){
		document.getElementById('receipt_nm').disabled=false;
		document.getElementById('receipt_memo').disabled=false;

		// 2014/08/20
		if (document.getElementById('ship_invoice_fl_x').checked == true) {
			document.getElementById('ship_invoice_fl_n').checked = true;
		}
		document.getElementById('ship_invoice_fl_x').disabled=true;
	}else{
		document.getElementById('receipt_nm').disabled=true;
		document.getElementById('receipt_memo').disabled=true;
		document.getElementById('receipt_nm').value='';
		document.getElementById('receipt_memo').value='';
		document.getElementById('ship_invoice_fl_x').disabled=false;	// 2014/08/20
	}

}

//=====================================================================
//	お届け先変更時
//=====================================================================

function changeShipAddress(){

	var i;
	var vAddressNo = new Array(<% = wAddressNoHTML %>);
	var vZip = new Array(<% = wZipHTML %>);
	var vAddress = new Array(<% = wAddressHTML %>);
	var vTelephoneNo = new Array(<% = wTelephoneNoHTML %>);
	var vAddressName = new Array(<% = wAddressNameHTML %>);

	//2012/08/08 nt add Start
	var flag;
	var vRitouFl = new Array(<% = wRitouFlHTML %>);
	var vKuyuKinshiFl = new Array(<% = wKuyuKinshiFlHTML %>);
	//2012/08/08 nt add End

	//2012/08/25 nt add Start
	var vSagawaLTFlg = new Array(<% = wSagawaLTHTML %>);
	//2012/08/25 nt add End

	var idx = document.f_data.select_ship_address_no.selectedIndex;
	if(idx <= 0){
		return;
	}
	var AddrNo = document.f_data.select_ship_address_no.options[idx].value;
	if(AddrNo <= 0){
		return;
	}

	for ( i=0; i<vAddressNo.length; i++){
		if(AddrNo==vAddressNo[i]){
			idx = i;
			break;
		}
	}

	//2014/08/05
	document.getElementById('i_selected_ship_address').value = vAddress[idx];
	if(document.f_data.radio_daibiki.checked){
		$("#ship_invoice").css("display", "none");
		document.getElementById('ship_invoice_fl_x').disabled = true;
		document.getElementById('ship_invoice_fl_y').disabled = true;
		document.getElementById('ship_invoice_fl_n').disabled = true;
	} else {
		if (AddrNo != 1) {
			if ($("#i_ship_address_no1").val() == $("#i_selected_ship_address").val()) {
				$("#ship_invoice").css("display", "none");
			} else {
				if ($("#ship_invoice").css("display") == 'none') {
					$("#ship_invoice").css("display", "inline");
					// 2014/08/20
					document.getElementById('ship_invoice_fl_y').disabled = false;
					document.getElementById('ship_invoice_fl_n').disabled = false;
					if (document.getElementById('receipt_fl_n').checked == true) {
						document.getElementById('ship_invoice_fl_x').disabled = false;
						document.getElementById('ship_invoice_fl_x').checked  = true;
					} else {
						if (document.getElementById('ship_invoice_fl_x').checked == true) {
							document.getElementById('ship_invoice_fl_n').checked = true;
						}
					}
				}
			}
		} else {
			$("#ship_invoice").css("display", "none");
			document.getElementById('ship_invoice_fl_x').disabled = true;
			document.getElementById('ship_invoice_fl_y').disabled = true;
			document.getElementById('ship_invoice_fl_n').disabled = true;
		}
	}



	document.f_data.ship_address_no.value = AddrNo;
	document.getElementById('ShipZip').innerHTML = vZip[idx];
	document.getElementById('ShipAddress').innerHTML = vAddress[idx];
	document.getElementById('ShipTel').innerHTML = 'Tel. ' + vTelephoneNo[idx];
	document.getElementById('ShipName').innerHTML = vAddressName[idx] + ' 様';
	document.f_data.select_ship_address_no.length = vAddressNo.length - 1;

	//2012/08/08 nt add Start
	//嵩重量品対応
	flag=0;
	if (vRitouFl[idx] == "Y" && vKuyuKinshiFl[idx] == "Y") {
		//離島フラグ：Y + 嵩重量品フラグ：Yの場合、「代引」選択不可
		document.getElementById('radio_daibiki').disabled=true;

		//「代引」が選択されていた場合、「銀行振込」へラジオボタンのcheckedを変更
		if(document.f_data.radio_daibiki.checked){
			flag = 1;

			if (flag = 1) {
				document.getElementById('radio_ginkou').checked=true;
			}
		}

		//表示・非表示を切替え
		$("#lDaibiki2").css("display","inline");
		$("#lDaibiki3").css("display","none");
		$("#lDaibiki").css("display","none");

	//2012/08/25 nt add Start
	//佐川制限対応
	}else if (vSagawaLTFlg[idx] == "Y") {
		//佐川代引き制限フラグ：Yの場合、「代引」選択不可
		document.getElementById('radio_daibiki').disabled=true;

		//「代引」が選択されていた場合、「銀行振込」へラジオボタンのcheckedを変更
		if(document.f_data.radio_daibiki.checked){
			flag = 1;

			if (flag = 1) {
				document.getElementById('radio_ginkou').checked=true;
			}
		}

		//表示・非表示を切替え
		$("#lDaibiki3").css("display","inline");
		$("#lDaibiki2").css("display","none");
		$("#lDaibiki").css("display","none");

	//2012/08/25 nt add End
	}else if (vRitouFl[idx] != "Y" && vKuyuKinshiFl[idx] == "Y") {
		//嵩重量品フラグ：Yのみの場合、「代引」選択可
		document.getElementById('radio_daibiki').disabled=false;

		//表示・非表示を切替え
		$("#lDaibiki2").css("display","none");
		$("#lDaibiki3").css("display","none");
		$("#lDaibiki").css("display","inline");

	//2012/08/25 nt add Start
	//佐川制限対応
	}else if (vSagawaLTFlg[idx] != "Y") {
		//佐川代引き制限フラグ：Yでない場合、「代引」選択可
		document.getElementById('radio_daibiki').disabled=false;

		//表示・非表示を切替え
		$("#lDaibiki2").css("display","none");
		$("#lDaibiki3").css("display","none");
		$("#lDaibiki").css("display","inline");

	//2012/08/25 nt add End
	}else{
		//上記以外は「代引」選択可
		document.getElementById('radio_daibiki').disabled=false;
	}
	//2012/08/08 nt add End

	idx = 0;
	for ( i=0; i<vAddressNo.length; i++){
		if(AddrNo!=vAddressNo[i]){
			document.f_data.select_ship_address_no.options[idx].value = vAddressNo[i];
			document.f_data.select_ship_address_no.options[idx].text = vAddressName[i] + '　' + vZip[i] + ' ' + vAddress[i] + '　' + vTelephoneNo[i];
			idx++;
		}
	}
	document.f_data.select_ship_address_no.options[0].selected = true;
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
      <li class="now">お届け先、お支払い方法の選択</li>
    </ul>
  </div></div></div>

<% If wMsg <> "" Then %>
  <p class="error"><% = wMsg %></p>
<% End If %>

<% If wErrMsg <> "" Then %>
  <p class="error"><% = wErrMsg %></p>
<% End If %>

  <h1 class="title">お届け先、お支払い方法の選択</h1>
  <ol id="step">
    <li><img src="images/step01.gif" alt="1.ショッピングカート" width="170" height="50"></li>
    <li><img src="images/step02_now.gif" alt="2.お届け先、お支払方法の選択" width="170" height="50"></li>
    <li><img src="images/step03.gif" alt="3.ご注文内容の確認" width="170" height="50"></li>
    <li><img src="images/step04.gif" alt="4.ご注文完了" width="170" height="50"></li>
  </ol>

  <h2 class="cart_title">お届け先</h2>
  <form action="JavaScript:OrderSubmit('next')" method="post" name="f_data" >
    <table id="address">
      <tr>
        <td class="main">
          <div class="box_l">
            <p>
              <div id="ShipZip"><% = "〒" & wCustomerZip %></div>
              <div id="ShipAddress"><% = wCustomerPref & wCustomerAddress %></div>
              <div id="ShipTel"><% = "Tel. " & wCustomerTel %></div>
              <div id="ShipName"><% = wCustomerNm & " 様" %></div>
            </p>
          </div>
          <div class="box_r">
<% = wShipAddressHTML %>
            <p class="change"><a href="JavaScript:OrderSubmit('address');">新しい住所を登録する</a></p>
          </div>
        </td>
      </tr>
<%
' 2014/08/05 GV mod start
'1)代引き(送り先が本人でなくても）は表示しない
'2)届け先が本人と違場合、表示
'3)登録住所と入力された新住所（連番1以外）が異なる場合、表示
'  (登録住所と同じ内容を毎回入力するお客様への対応のため)
wShowInvoice = "none;"

If (wShipInvoiceFl = "X") Then
	wInvoiceChecked(0) = " checked"
ElseIf (wShipInvoiceFl = "N") Then
	wInvoiceChecked(1) = " checked"
ElseIf (wShipInvoiceFl = "Y") Then
	wInvoiceChecked(2) = " checked"
Else
	wInvoiceChecked(0) = " checked"
End If

If (wPaymentMethod = "代引き") Then
	wShowInvoice = "none;"
	wInvoiceDisabled = " disabled='disabled'"
Else
	If (wShipAddressNo <> "1") Then
'		If wShipAddressNo1Data = (wCustomerPref & wCustomerAddress) Then
'			wShowInvoice = "none;"
'			wInvoiceDisabled = " disabled='disabled'"
'		Else
			wShowInvoice = "inline;"
			wInvoiceDisabled = ""
'		End If
	Else
		wShowInvoice = "none;"
		wInvoiceDisabled = " disabled='disabled'"
	End If
End If
' 2014/08/05 GV mod end
%>
      <tr id="ship_invoice" style="display:<%=wShowInvoice%>">
        <td class="left">納品書の送付先を選択してください。<br>
          <input type="radio" name="ship_invoice_fl" id="ship_invoice_fl_x" value="X"<%=wInvoiceChecked(0)%><%=wInvoiceDisabled%>><label for="ship_invoice_fl_x">不要</label><br>
          <input type="radio" name="ship_invoice_fl" id="ship_invoice_fl_n" value="N"<%=wInvoiceChecked(1)%><%=wInvoiceDisabled%>><label for="ship_invoice_fl_n">購入者（別途郵送手配）...プレゼント、贈り物の際はこちらをご選択ください。</label><br>
          <input type="radio" name="ship_invoice_fl" id="ship_invoice_fl_y" value="Y"<%=wInvoiceChecked(2)%><%=wInvoiceDisabled%>><label for="ship_invoice_fl_y">届け先（荷物へ同梱）...ご自身のお買い物の際にはこちらをご選択ください。</label>
        </td>
      </tr>
    </table>
    <p class="kome">※商品をご購入いただく際は、「<a href="#notes">ご注文にあたり</a>」を必ずご確認ください。</p>

    <h2 class="cart_title">お支払い方法の選択</h2>
    <table id="pay">
      <tr>
        <th>お支払い方法</th>
        <th>配送方法</th>
        <th>ご希望配達日時</th>
        <th>領収書</th>
      </tr>
      <tr>
        <td>
          <ul class="select">
            <li onClick="checkPaymentMethod();">
              <input id="radio_ginkou" name="payment_method" type="radio" value="銀行振込" checked><label for="radio_ginkou">銀行振込</label>
              <span>振込人名義</span><span><input name="furikomi_nm" type="text" id="furikomi_nm" value="<% = wFurikomiNm %>">様</span>
            </li>
            <li onClick="checkPaymentMethod();">
              <input id="radio_netbank" name="payment_method" type="radio" value="コンビニ支払"><label for="radio_netbank">ネットバンキング</label>
              <span><label for="radio_netbank">（Pay-easy<img src="images/payeasy.gif" alt="Pay-easy">）</label></span>
              <span><label for="radio_netbank">ゆうちょ</label></span><span><label for="radio_netbank">コンビニ払い</label></span>
            </li>
            <li onClick="checkPaymentMethod();">

              <!-- 2012/08/25 nt mod Start -->
              <!-- <input id="radio_daibiki" name="payment_method" type="radio" value="代引き"><label for="radio_daibiki">代金引換</label> -->
                   <input id="radio_daibiki" name="payment_method" type="radio" value="代引き">
                   <label for="radio_daibiki" id="lDaibiki">代金引換</label>
                   <label for="radio_daibiki" id="lDaibiki2" style="display:none;"><a href="#hmethods" class="fancybox" id="aDaibiki">代金引換</a></label>
                   <label for="radio_daibiki" id="lDaibiki3" style="display:none;"><a href="#hmethods2" class="fancybox" id="aDaibiki">代金引換</a></label>
              <!-- 2012/08/25 nt mod End -->
            </li>
            <li onClick="checkPaymentMethod();"><input id="radio_loan" name="payment_method" type="radio" value="ローン"><label for="radio_loan">ローン</label></li>
          </ul>
        </td>
        <td>
          <ul class="select">
            <li><input id="ikkatsu_fl_y" name="ikkatsu_fl" type="radio" value="Y" checked><label for="ikkatsu_fl_y">一括出荷</label></li>
            <li><input id="ikkatsu_fl_n" name="ikkatsu_fl" type="radio" value="N"><label for="ikkatsu_fl_n">在庫商品から出荷</label></li>
          </ul>
        </td>
        <td>
          <ul class="select">
            <li onClick="checkDeliveryDate();"><input id="delivery_fl_n" name="delivery_fl" type="radio" value="N"><label for="delivery_fl_n">指定なし</label></li>
            <li onClick="checkDeliveryDate();">

              <!-- 2012/08/08 nt mod Start -->
              <!-- <input id="delivery_fl_y" name="delivery_fl" type="radio" value="Y"><label for="delivery_fl_y">指定あり</label> -->
              <% if (wKuyuKinshiFl = "Y") then %>
                   <input id="delivery_fl_y" name="delivery_fl" type="radio" value="Y" disabled>
                   <label for="delivery_fl_y"><a href="#hdelivery_time" class="fancybox">指定あり</a></label>
              <% else %>
                   <input id="delivery_fl_y" name="delivery_fl" type="radio" value="Y">
                   <label for="delivery_fl_y">指定あり</label>
              <% end if %>
              <!-- 2012/08/08 nt mod End -->

              <span>
                <select id="delivery_mm" name="delivery_mm" disabled>
                  <option value=""></option>
                  <option value="01">1</option>
                  <option value="02">2</option>
                  <option value="03">3</option>
                  <option value="04">4</option>
                  <option value="05">5</option>
                  <option value="06">6</option>
                  <option value="07">7</option>
                  <option value="08">8</option>
                  <option value="09">9</option>
                  <option value="10">10</option>
                  <option value="11">11</option>
                  <option value="12">12</option>
                </select>月
                <select id="delivery_dd" name="delivery_dd" disabled>
                  <option value=""></option>
                  <option value="01">1</option>
                  <option value="02">2</option>
                  <option value="03">3</option>
                  <option value="04">4</option>
                  <option value="05">5</option>
                  <option value="06">6</option>
                  <option value="07">7</option>
                  <option value="08">8</option>
                  <option value="09">9</option>
                  <option value="10">10</option>
                  <option value="11">11</option>
                  <option value="12">12</option>
                  <option value="13">13</option>
                  <option value="14">14</option>
                  <option value="15">15</option>
                  <option value="16">16</option>
                  <option value="17">17</option>
                  <option value="18">18</option>
                  <option value="19">19</option>
                  <option value="20">20</option>
                  <option value="21">21</option>
                  <option value="22">22</option>
                  <option value="23">23</option>
                  <option value="24">24</option>
                  <option value="25">25</option>
                  <option value="26">26</option>
                  <option value="27">27</option>
                  <option value="28">28</option>
                  <option value="29">29</option>
                  <option value="30">30</option>
                  <option value="31">31</option>
                </select>日
              </span>
              <span>時間</span>
              <span>
                <select id="delivery_tm" name="delivery_tm" style="width:115px;" disabled>
                  <option value="" selected="selected"></option>
                  <option value="<%=wDeliveryTime01%>"><%=wDeliveryTime01%></option>
                  <option value="<%=wDeliveryTime02%>"><%=wDeliveryTime02%></option>
                  <option value="<%=wDeliveryTime03%>"><%=wDeliveryTime03%></option>
                  <option value="<%=wDeliveryTime04%>"><%=wDeliveryTime04%></option>
                  <option value="<%=wDeliveryTime05%>"><%=wDeliveryTime05%></option>
                </select>
              </span>
            </li>
            <li><input type="checkbox" id="eigyousho_dome_fl" name="eigyousho_dome_fl" value="Y"><label for="eigyousho_dome_fl">運送会社営業所止め</label></li>
          </ul>
        </td>
        <td>
          <ul class="select">
            <li onClick="checkReceipt();"><input id="receipt_fl_n" name="receipt_fl" type="radio" value="N" checked><label for="receipt_fl_n">不要</label></li>
            <li onClick="checkReceipt();">
              <input id="receipt_fl_y" name="receipt_fl" type="radio" value="Y"><label for="receipt_fl_y" id="receipt1">必要</label><label for="receipt_fl_y" id="receipt2" style="display:none;"><a href="#receipt" class="fancybox">必要</a></label>
              <span>領収書宛名</span><span><input name="receipt_nm" type="text" id="receipt_nm" value="<% = wReceiptNm %>" disabled>様</span>
              <span>領収書但し書き</span><span><input type="text" name="receipt_memo" id="receipt_memo" value="<% = wReceiptMemo %>" disabled></span>
            </li>
          </ul>
        </td>
      </tr>
      <tr>
      	<td class="detail"><a href="#payment" class="fancybox">詳細はこちら</a></td>
        <td class="detail"><a href="#methods" class="fancybox">詳細はこちら</a></td>
        <td class="detail"><a href="#delivery_time" class="fancybox">詳細はこちら</a></td>
        <td class="detail"><a href="#receipt" class="fancybox">詳細はこちら</a></td>
      </tr>
    </table>

<% If wCustomerKabusokuAm > 0 And wCustomerClass = "一般顧客" Then %>
    <dl class="excess">
      <dt>クレジット／過不足金</dt>
      <dd><% = FormatNumber(wCustomerKabusokuAm, 0) %>円</dd>
    </dl>
    <p class="excess_use"><input type="checkbox" name='RebateFl' id="RebateFl" value="Y"><label for="RebateFl">お支払いにクレジット/過不足金を使用する</label></p>
<% End If %>

<% If wNoData = False Then %>
    <div id="btn_box">
      <ul class="btn next">
        <li><a href="JavaScript:OrderSubmit('next');"><img src="images/btn_next.png" alt="次へ" class="opover"></a></li>
      </ul>
    </div>
<% End If %>
    <input type="hidden" name="cmd" value="">
    <input type="hidden" name="customer_kn" value="<% = wCustomerKn %>">
    <input type="hidden" name="customer_email" value="<% = wCustomerEmail %>">
    <input type="hidden" name="telephone" value="<% = wCustomerTel %>">
    <input type="hidden" name="KabusokuAm" value="<% = wCustomerKabusokuAm %>">
    <input type="hidden" name="i_rebate_fl" value="<% = wRebateFl %>">
    <input type="hidden" name="i_payment_method" value="<% = wPaymentMethod %>">
    <input type="hidden" name="i_ship_address_no" value="<% = wShipAddressNo %>">
    <input type="hidden" name="i_ship_invoice_fl" value="<% = wShipInvoiceFl %>">
    <input type="hidden" name="i_ikkatsu_fl" value="<% = wIkkatsuFl %>">
    <input type="hidden" name="i_freight_forwarder" value="<% = wFreightForwarder %>">
    <input type="hidden" name="i_delivery_mm" value="<% = wDeliveryMM %>">
    <input type="hidden" name="i_delivery_dd" value="<% = wDeliveryDD %>">
    <input type="hidden" name="i_delivery_tm" value="<% = wDeliveryTM %>">
    <input type="hidden" name="i_eigyousho_dome_fl" value="<% = wEigyoushoDomeFl %>">
    <input type="hidden" name="i_receipt_fl" value="<% = wReceiptFl %>">
    <input type="hidden" name="i_tokuchuu_fl" value="<% = wTokuchuuFl %>">
    <input type="hidden" name="i_daibiki_fuka_fl" value="<% = wDaibikiFukaFl %>">
    <input type="hidden" name="ship_address_no" value="<% = wShipAddressNo %>">
    <input type="hidden" name="freight_forwarder" value="5">
  </form>
<% '2014/08/05 GV add start %>
  <form id='f_shipping_data' name='f_shipping_data'>
    <input type="hidden" name="i_ship_address_no1" id="i_ship_address_no1" value="<% = wShipAddressNo1Data %>">
    <input type="hidden" name="i_selected_ship_address" id="i_selected_ship_address" value="<% = wSelectedShipAddressData %>">
  </form>
<% '2014/08/05 GV add end %>
  <ul class="info left">
    <li><a href="#cancel" class="fancybox">ご注文商品のキャンセル・返品について</a></li>
    <li><a href="#delivery" class="fancybox">商品の納期についてはこちら</a></li>
  </ul>

  <h2 id="notes" class="cart_title">ご注文にあたり</h2>
  <h3 style="font-weight:bold;">当日配達サービスについて</h3>
  <p>お届け先は関東一都七県（東京、神奈川、千葉、埼玉、茨城、栃木、群馬、山梨）に限定されております。<br>また、以下の場合は当日配達対象外となります。</p>
  <ul id="attention" style="margin:.5em;">
    <li>大型商品を含むご注文</li>
    <li>営業所止めのご注文</li>
    <li>配達日時指定がされているご注文</li>
    <li>お支払い方法「代金引換」以外のご注文</li>
  </ul>
  <p class="notice"><span style="color:red">※</span>当日配送サービスでは、配達時間帯のご希望はお受けすることができません。</p>

  <ul id="attention" style="border:3px solid #ccc; line-height:1.8;background-color:#f0f0f0;padding:.8em;">
    <li>確認メールは自動的に送信されます。売買契約は商品の発送をもって成立となります。</li>
    <li>携帯電話などの受信制限があるアドレスでご登録された場合、 ご注文情報が受信できない場合がございます。</li>
    <li>ご注文商品についてのお問い合わせは、メールやお電話にて承っておりますのでご注文前にご確認いただけますようお願いいたします。</li>
    <li>カートに入れた商品以外をご希望の際は、あらかじめメールやお電話にてお問い合わせください。</li>
  </ul>

  <div style="display:none;">

    <!-- お支払方法 -->
    <div id="payment">
      <h2>お支払い方法について</h2>
      <div>
        <ul>
          <li>銀行振込
            <ul>
              <li>振込人名義欄は会員登録のお名前と異なる場合のみご記入ください。</li>
              <li>振込手数料はお客様の負担とさせていただきます。</li>
              <li>後ほど、お見積書をご案内いたしますので、こちらをご確認後にお振込みいただきますようお願いいたします。</li>
            </ul>
          </li>
          <li>ネットバンキング・ゆうちょ・コンビニ払い
            <ul>
              <li>ローソン、ファミリーマート、サークルK、サンクス、セイコーマート、ゆうちょ銀行、ネットバンキング でお支払いただけます。</li>
              <li>ご注文後、金額が変更となるご注文の変更は承ることができません。在庫の無い商品などをご注文の際は、事前にお問い合わせください。</li>
              <li>E-MAILアドレスが携帯の場合は、必要事項が確認できない場合があるため、パソコンからのご利用をおすすめします。</li>
              <li>後ほど、お見積書をご案内いたしますので、こちらをご確認後にお振込みいただきますようお願いいたします。</li>
            </ul>
          </li>
          <li>代金引換
            <ul>
              <li>代金引換でのご購入の場合、商品の発送は一括出荷となります。また、お支払いは現金のみの受付となります。</li>
            </ul>
          </li>
          <li>ローン
            <ul>
              <li>オンラインローンの場合､お申込後のご注文内容の変更を承ることができません。</li>
              <li>ご注文内容と､オンラインローン申込フォームの内容をご確認の上ご注文ください。</li>
              <li>ジャックスでお申し込みの場合は、頭金なしとなります。</li>
            </ul>
          </li>
        </ul>
        <p class="info"><a href="http://guide.soundhouse.co.jp/guide/oshiharai.asp" target="_blank">お支払い方法について詳しくはこちらをご覧ください。</a></p>
      </div>
    </div>

    <!-- 配送方法 -->
    <div id="methods">
      <h2>配送方法について</h2>
      <ul>
        <%
        '<li>ご指定が無い場合は佐川急便で発送いたします。</li>
        %>
        <li>
          <%
          '沖縄など離島へのお届けはサイズの小さい商品に限りヤマト運輸で発送いたします。<br>また、
	      %>
          大型の商品の場合はお届けまでに1週間程度お時間をいただく場合があります。
        </li>
        <li>発送は、ご注文商品が全て揃った時点でまとめて発送する「一括出荷」または在庫のあるものから都度発送する「在庫商品から出荷」のいずれかの方法をご注文の際に選択できます。</li>
        <li>複数商品をご注文いただいた場合、商品によっては同じご注文でも別配送となる場合があります。</li>
        <li>代金引換での発送は一括発送のみとなります。
          <ul>
            <li>お取り寄せ商品が含まれる場合、商品が揃い次第の一括出荷となります。</li>
            <%
            '<li>ヤマト運輸での代金引換の発送は1個口のみとなります。</li>
            %>
          </ul>
        </li>
      </ul>
    </div>

    <!-- 配送会社・配送日時 -->
    <div id="delivery_time">
      <h2><!--運送会社・-->配送日時の指定</h2>
      <ul>
        <li>天候、交通状況ならびに配達業者の都合によりご希望に添えない場合がございます。日程には余裕を持ってご注文ください。</li>
        <%
        '<li>
        '  佐川急便の場合、時間帯指定は平日ならびにお届け先が
        '  都市部にお住まいの個人宅の場合に限り可能です。
        '</li>
        %>
        <li>一部、配達日時のご希望をお受けできない地域がございます。詳細は電話もしくはメールにてお問い合わせください。</li>
        <li>運送会社営業所止めを指定された場合、お届先住所に該当する宅配便会社の規定営業所への留め置きとなります。</li>
        <li>いかなる場合においても配達遅延から生じる損害(事業利益の損失、事業の遅延・中断、事業情報の損失またはその他の金銭的損害等)に関して、サウンドハウスは一切の責任を負いません。</li>
        <li>代金引換の場合、配送日指定は10日以内の日付を指定してください。</li>
      </ul>
    </div>

    <!-- 領収書 -->
    <div id="receipt">
      <h2>領収書の発行について</h2>
      <ul>
        <li>領収書は、納品書の最後に印刷されております。お手数ですが、切り離してご使用ください。</li>
        <li>お支払い方法が以下の場合、サウンドハウスの領収書は発行いたしません。
          <ol>
            <li>代金引換</li>
            <li>ローン</li>
            <li>コンビニ/郵便局支払</li>
          </ol>
        </li>
        <li>宛名、但書きは、指定欄に入力いただいた内容のまま作成いたします。</li>
        <li>1件のご注文に対して、領収書は1枚のみ発行させていただきます。</li>
      </ul>
      <p class="info"><a href="http://guide.soundhouse.co.jp/guide/kaimono.asp#ryousyuu" target="_blank">領収書について詳しくはこちらをご覧ください。</a></p>
    </div>

    <!-- キャンセル・返品 -->
    <div id="cancel">
      <h2>購入された商品の交換、キャンセルについて</h2>
      <div>
        <p>サウンドハウスでは、原則としてお客様のご都合によるキャンセル・返品は承ることができません。商品の詳細な仕様や納期など、ご不明な点は事前にお問い合わせの上、ご注文いただきますようお願い申し上げます。<br>ただし、商品到着後7日以内にお申し出いただければ、下記の条件に該当する場合のみ交換・キャンセルを承ります。</p>
        <h3>交換・キャンセルが可能なケース</h3>
        <ol>
          <li>破損、誤送の場合
            <p>万一商品が破損していたり、ご注文と異なる商品が届いた場合は、商品到着後7日以内にご連絡ください。担当スタッフより破損・誤送商品の引き取り、再送の手順について詳しく説明させていただきます。</p>
          </li>
          <li>初期不良の場合
            <p>ご購入いただいた商品に不具合があり、商品到着後7日以内にご連絡いただければ、弊社にて商品の状態を確認後、別商品への交換、もしくはご注文のキャンセルを承ります。代金は、お客様が指定された口座に返金、もしくはお預かり金として次回ご注文いただく際に相殺いたします。<br>ただし、下記に記載されている6番の「お客様の使用環境や機材の相性が理由で商品が作動しない場合」を除きます。</p>
          </li>
          <li>お客様都合による場合
            <p>お届けした商品のキャンセル、または他商品への交換ご希望の場合、該当商品が未開封、未使用であり、なおかつ商品を受け取った日から7日以内にご連絡いただければ、キャンセル、または他商品への交換を承ります。また、箱、梱包を開封されている場合でも、未使用であれば、商品代金の15%をお支払いいただくことにより、キャンセルまたは他商品への交換を承ります。
            </p>
          </li>
        </ol>
        <h3>交換・キャンセルができないケース</h3>
        <p>次の場合は、商品到着後7日以内であっても交換およびキャンセルはお受けできません。</p>
        <ol>
          <li>商品を開封して使用した場合</li>
          <li>ソフトウェアを開封した場合</li>
          <li>メーカーが返品を受け付けない場合</li>
          <li>体に直接身につける商品の場合</li>
          <li>お客様の指定によるお取り寄せ商品や特注品</li>
          <li>お客様の使用環境や機材の相性が理由で商品が作動しない場合</li>
          <li>お客様の元で汚損、破損が生じた商品</li>
          <li>商品に付属するオリジナルの外箱および梱包材がすべて揃っていない場合</li>
        </ol>
        <p>※詳細は弊社カスタマーサポートまでお問い合わせください。</p>
        <h3>商品の交換、キャンセルの手順</h3>
        <ol>
          <li>商品の交換、またはキャンセルをご希望の場合、事前に詳細を確認した上で弊社発行のRA（返品承認）番号が必要となります。商品を送付する前に必ず弊社のカスタマーサポートまでご連絡ください。内容確認の上、RA(返品承認)番号を発行します。</li>
          <li>商品をお送りいただく際は、運送会社の送り状の備考欄にRA番号をご記入ください。商品の外箱そのものには記載しないようお願いいたします。RA番号が記載されていない商品は受領できませんので、あらかじめご了承ください。</li>
          <li>保証期間内の修理などで弊社が送料を負担する場合、指定の運送会社（通常は佐川急便）以外は有料となります。お客様が元払いにて商品を発送される場合は、どの運送会社でもご利用可能です。</li>
          <li>商品を送付される場合は、必ず外箱、マニュアル、保証書等、全ての付属品を受け取り時と同じ状態のまま送付してください。</li>
          <li>商品到着後、弊社もしくはメーカーにて商品の検品を行い、交換、またはキャンセルの処理を進めさせていただきます。同一商品と交換できない場合は、差額を調整した上で他商品への交換、もしくは返金にて対応する場合もございます。</li>
          <li>お客様へ商品の代金を返金する場合、弊社に商品が到着し、検品、確認を行った後に、お客様ご指定の銀行口座へ振込みにて返金いたします。</li>
          <li>返品・交換のお申し出から2週間以内にご返送いただけない場合は、一旦キャンセル扱いとさせていただきますのであらかじめご了承ください。</li>
        </ol>
      </div>
    </div>

    <!-- 納期について -->
    <div id="delivery">
      <h2>商品の納期について</h2>
      <ul>
        <li>ウェブサイト上、およびご注文やお見積り時点でご案内しております納期につきましては、あくまでも予定となっており、諸事情により変更となる場合がございます。</li>
        <li>商品の納期につきましては、メールやお電話でのお問い合わせも承っております。指定日までに納品が必要なご注文は、遠慮なく事前にご相談ください。</li>
        <li>なお、納期遅延によって生じた問題につきましては、当社では一切の責を負うことができません。あらかじめご了承ください。</li>
      </ul>
    </div>

    <!-- 2012/08/08 nt add Start -->
    <!-- 嵩重量品 代金引換について -->
    <div id="hmethods">
      <ul>
        <p align="center">お客様のご注文は、重量、もしくは大きさが規定値を超える商品を含む為、代金引換以外のお支払方法を選択してください。</p>
      </ul>
    </div>

    <!-- 嵩重量品 配達日時指定について -->
    <div id="hdelivery_time">
      <ul>
        <p align="center">お客様のご注文は、重量、もしくは大きさが規定値を超える商品を含む為、配達日時指定なしでのお届けとなります。</p>
      </ul>
    </div>
    <!-- 2012/08/08 nt add End -->

    <!-- 2012/08/25 nt add Start -->
    <!-- 佐川制限 代金引換について -->
    <div id="hmethods2">
      <ul>
        <p align="center">お客様のご注文は、代金引換をお受けできない地域の為、代金引換以外のお支払方法を選択してください。</p>
      </ul>
    </div>
    <!-- 2012/08/08 nt add End -->

  </div>

<!--/#contents --></div>
	<div id="globalSide">
	<!--#include file="../Navi/NaviSide.inc"-->
	<!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript" src="jslib/jquery.fancybox-1.3.4.pack.js"></script>
<script type="text/javascript">
$(function(){
	$(".fancybox").fancybox({
	'scrolling'		: 'no',
	'titleShow'		: false
	});
});
preset_values();
</script>
</body>
</html>