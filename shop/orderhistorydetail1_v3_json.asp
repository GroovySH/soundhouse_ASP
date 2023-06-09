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
<!--#include file="../3rdParty/aspJSON1.17.asp"-->
<%
'========================================================================
'
'	購入履歴一覧ページ
'
'
'変更履歴
'2014/09/16 GV 新規作成
'2015.05.07 GV 過不足相殺金額を追加
'2016.02.05 GV 不要処理を削除。(Web注文変更キャンセル機能)
'2016.06.01 GV 非表示フラグの有無を追加。
'2018.12.21 GV PayPal対応。
'2020.02.05 GV 請求書DL対応。
'2020.03.18 GV 請求書DL対応。
'2020.06.30 GV 欲しい物リスト対応。(#2841)
'
'========================================================================
'On Error Resume Next

Dim ConnectionEmax

Dim wErrMsg						' エラーメッセージ (他のページから渡されるメッセージ)
Dim wDispMsg					' 通常メッセージ(エラー以外) (他のページから渡されるメッセージ)
Dim wErrDesc
Dim wMsg						' エラーメッセージ (本ページで作成するメッセージ)

Dim oJSON						' JSONオブジェクト
Dim wCustomerNo					' 顧客番号
Dim wOrderNo					' 受注番号
Dim wGiftCustomerNo				' ギフト顧客番号
Dim wGiftNo						' ギフト番号
Dim wOrderGift					' ギフト注文フラグ

'=======================================================================
'	受け渡し情報取り出し & 初期設定
'=======================================================================
' Getパラメータ
wCustomerNo = ReplaceInput(Trim(Request("cno")))
wOrderNo = ReplaceInput(Trim(Request("ono")))

'ギフト注文フラグ
wOrderGift = ReplaceInput_NoCRLF(Trim(Request("gift")))
If ((IsNull(wOrderGift) = True) Or (UCase(wOrderGift) <> "Y")) Then
	wOrderGift = "N"
Else
	wOrderGift = "Y"
End If

'=======================================================================
'	Execute main
'=======================================================================
Call connect_db()

Call main()

'---- エラーメッセージをセッションデータに登録   ' member系の他のページ処理にならう
If Err.Description <> "" Then
'	wErrDesc = THIS_PAGE_NAME & " " & Replace(Replace(Err.Description, vbCR, " "), vbLF, " ")
'	Call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
End If

Call close_db()

If Err.Description <> "" Then

End If


'========================================================================
'
'	Function	Connect database
'
'========================================================================
Function connect_db()

Set ConnectionEmax = Server.CreateObject("ADODB.Connection")
ConnectionEmax.Open g_connectionEmax

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
Function close_db()

ConnectionEmax.close
Set ConnectionEmax= Nothing

End function

'========================================================================
'
'	Function	Main
'
'========================================================================
Function main()

Dim vSQL
Dim vRS

Dim orderDate
Dim shippingDate
Dim estimateDate
Dim one_time_todokesaki1
Dim one_time_todokesaki2
Dim final_nouki_date_time
Dim receiptName
Dim receiptNote
Dim totalOrderAmount2
Dim usedPoint
Dim furikomiMeigi
Dim pos
Dim storeStop
Dim webModCancelFlg
Dim deleteDate
Dim hide
Dim paymentMethodDetail '2018.12.21 GV add
Dim receiptFlag '2020.02.05 GV add
Dim receiptDate '2020.02.05 GV add
Dim displayReceiptDate '2020.03.18 GV add
Dim giftCustomerNo '2021.06.30 GV add
Dim giftNo '2021.06.30 GV add

Set oJSON = New aspJSON


one_time_todokesaki1 = ""
one_time_todokesaki2 = ""
final_nouki_date_time = ""
receiptName = ""
receiptNote = ""
totalOrderAmount2 = 0
hide = ""

'-- 非表示フラグが存在しているか
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "  count(受注番号) as cnt "
vSQL = vSQL & " FROM "
vSQL = vSQL & "   受注非表示リスト ov WITH (NOLOCK) "
vSQL = vSQL & " WHERE "
vSQL = vSQL & " ov.受注番号 = " & wOrderNo & " "
Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then
	If (CDbl(vRS("cnt")) > 0) Then
		hide = "Y"
	End If
End If


'--- ヘッダ部分の情報取出し
vSQL = ""
vSQL = vSQL & "SELECT TOP 1 "
vSQL = vSQL & "      a.受注番号 "
vSQL = vSQL & "    , a.顧客番号 "
vSQL = vSQL & "    , a.見積日 "
vSQL = vSQL & "    , a.受注日 "
vSQL = vSQL & "    , a.出荷完了日 "
vSQL = vSQL & "    , a.受注形態 "
vSQL = vSQL & "    , a.支払方法 "
vSQL = vSQL & "    , a.商品合計金額 "
vSQL = vSQL & "    , a.送料 "
vSQL = vSQL & "    , a.代引手数料 "
vSQL = vSQL & "    , a.受注合計金額 "
vSQL = vSQL & "    , a.合計金額 "
vSQL = vSQL & "    , a.外税合計金額 "
vSQL = vSQL & "    , a.利用ポイント "
vSQL = vSQL & "    , a.一括出荷フラグ "
vSQL = vSQL & "    , a.領収書宛先 "
vSQL = vSQL & "    , a.領収書但し書き "
vSQL = vSQL & "    , a.Web受注変更開始日 "
vSQL = vSQL & "    , a.消費税率 "
vSQL = vSQL & "    , a.運送会社コード "
vSQL = vSQL & "    , a.担当者コード "
vSQL = vSQL & "    , a.振込名義人 "
vSQL = vSQL & "    , a.営業所止めフラグ "
vSQL = vSQL & ", (CASE "
vSQL = vSQL & "     WHEN a.受注形態 = 'ギフト' THEN '' "
vSQL = vSQL & "     ELSE b.今回限り届先郵便番号 END "
vSQL = vSQL & "   ) AS 今回限り届先郵便番号 "
vSQL = vSQL & ", (CASE "
vSQL = vSQL & "     WHEN a.受注形態 = 'ギフト' THEN '' "
vSQL = vSQL & "     ELSE b.今回限り届先都道府県 END "
vSQL = vSQL & "   ) AS 今回限り届先都道府県 "
vSQL = vSQL & ", (CASE "
vSQL = vSQL & "     WHEN a.受注形態 = 'ギフト' THEN '' "
vSQL = vSQL & "     ELSE b.今回限り届先住所 END "
vSQL = vSQL & "   ) AS 今回限り届先住所 "
vSQL = vSQL & ", (CASE "
vSQL = vSQL & "     WHEN a.受注形態 = 'ギフト' THEN gift_c.ハンドルネーム "
vSQL = vSQL & "     ELSE b.今回限り届先名前 END "
vSQL = vSQL & "   ) AS 今回限り届先名前 "
vSQL = vSQL & ", (CASE "
vSQL = vSQL & "     WHEN a.受注形態 = 'ギフト' THEN '' "
vSQL = vSQL & "     ELSE b.今回限り届先電話番号 END "
vSQL = vSQL & "   ) AS 今回限り届先電話番号 "
vSQL = vSQL & "    , b.今回限り届先郵便番号 AS ORG_今回限り届先郵便番号 "
vSQL = vSQL & "    , b.今回限り届先都道府県 AS ORG_今回限り届先都道府県 "
vSQL = vSQL & "    , b.今回限り届先住所 AS ORG_今回限り届先住所 "
vSQL = vSQL & "    , b.今回限り届先名前 AS ORG_今回限り届先名前 "
vSQL = vSQL & "    , b.今回限り届先電話番号 AS ORG_今回限り届先電話番号 "
vSQL = vSQL & "    , b.最終指定納期 "
vSQL = vSQL & "    , b.最終時間指定 "
vSQL = vSQL & "    , a.過不足相殺金額 " ' 2015.05.07 GV add
vSQL = vSQL & "    , a.削除日 "
vSQL = vSQL & "    , a.Web注文変更キャンセル中フラグ "
vSQL = vSQL & "    , a.支払方法詳細 " '2018.12.21 GV add
vSQL = vSQL & "    , a.領収書番号 " '2020.02.05 GV add
vSQL = vSQL & "    , a.領収書発行日 " '2020.02.05 GV add
vSQL = vSQL & "    , (CASE WHEN a.最終入金日 IS NULL THEN a.受注日 " '2020.03.18 GV add
vSQL = vSQL & "            ELSE a.最終入金日 " '2020.03.18 GV add
vSQL = vSQL & "       END) AS 領収日 " '2020.03.18 GV add
vSQL = vSQL & "    , a.ギフト顧客番号 "
vSQL = vSQL & "    , a.ギフト番号 "

vSQL = vSQL & "FROM "
vSQL = vSQL & "      " & gLinkServer & "受注     a WITH (NOLOCK) "
vSQL = vSQL & " INNER JOIN " & gLinkServer & "受注明細 b WITH (NOLOCK) "
vSQL = vSQL & "   ON b.受注番号 = a.受注番号 "

vSQL = vSQL & " LEFT JOIN " & gLinkServer & "顧客 gift_c WITH (NOLOCK) "
vSQL = vSQL & "   ON gift_c.顧客番号 = a.ギフト顧客番号 "

vSQL = vSQL & "WHERE "
If (wOrderGift = "N") Then
	vSQL = vSQL & "        a.受注番号 = " & wOrderNo
	vSQL = vSQL & "    AND a.顧客番号 = " & wCustomerNo & " "
Else
	vSQL = vSQL & "        a.ギフト番号 = " & wOrderNo
	vSQL = vSQL & "    AND a.ギフト顧客番号 = " & wCustomerNo & " "
End If

'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then

	' リスト追加
	oJSON.data.Add "data" ,oJSON.Collection()

	' 受注日
	If (IsNull(vRS("受注日"))) Then
		orderDate = ""
	Else
		orderDate = CStr(Trim(vRS("受注日")))
	End If

	' 見積日
	If (IsNull(vRS("見積日"))) Then
		estimateDate = ""
	Else
		estimateDate = CStr(Trim(vRS("見積日")))
	End If

	' 出荷完了日
	If (IsNull(vRS("出荷完了日"))) Then
		shippingDate = ""
	Else
		shippingDate = CStr(Trim(vRS("出荷完了日")))
	End If

	' 今回限り届先
	one_time_todokesaki1 = vRS("今回限り届先郵便番号") & "^" &_
		vRS("今回限り届先都道府県") & "^" &_
		vRS("今回限り届先住所") & "^" &_
		vRS("今回限り届先名前") & "^" &_
		vRS("今回限り届先電話番号")

	one_time_todokesaki1 = Replace(one_time_todokesaki1, """", "”")

	' ORG_今回限り届先
	one_time_todokesaki2 = vRS("ORG_今回限り届先郵便番号") & "^" &_
		vRS("ORG_今回限り届先都道府県") & "^" &_
		vRS("ORG_今回限り届先住所") & "^" &_
		vRS("ORG_今回限り届先名前") & "^" &_
		vRS("ORG_今回限り届先電話番号")

	one_time_todokesaki2 = Replace(one_time_todokesaki2, """", "”")


	' 最終指定納期と時刻
	final_nouki_date_time = vRS("最終指定納期")&"_"&vRS("最終時間指定") 

	' 領収書宛先
	If (IsNull(vRS("領収書宛先"))) Then
		receiptName = ""
	Else
		receiptName = CStr(Trim(vRS("領収書宛先")))
		receiptName = Replace(receiptName, """", "”")
	End If

	' 領収書但し書き
	If (IsNull(vRS("領収書但し書き"))) Then
		receiptNote = ""
	Else
		receiptNote = CStr(Trim(vRS("領収書但し書き")))
		receiptNote = Replace(receiptNote, """", "”")
	End If

	' 合計金額
	If (IsNull(vRS("合計金額"))) Then
		totalOrderAmount2 = 0
	Else
		totalOrderAmount2 = CDbl(vRS("合計金額"))
	End If

	' 利用ポイント
	If (IsNull(vRS("利用ポイント"))) Then
		usedPoint = 0
	Else
		usedPoint = CDbl(vRS("利用ポイント"))
	End If

	' 振込名義人
	If (IsNull(vRS("振込名義人"))) Then
		furikomiMeigi = ""
	Else
		furikomiMeigi = CStr(Trim(vRS("振込名義人")))
		furikomiMeigi = Replace(furikomiMeigi, """", "”")
	End If

	'営業所止めフラグ
	If (IsNull(vRS("営業所止めフラグ"))) Then
		storeStop = ""
	Else
		storeStop = CStr(Trim(vRS("営業所止めフラグ")))
	End If

	'Web注文変更キャンセル中フラグ
	If (IsNull(vRS("Web注文変更キャンセル中フラグ"))) Then
		webModCancelFlg = "N"
	Else
		If (Trim(vRS("Web注文変更キャンセル中フラグ")) <> "Y") Then
			webModCancelFlg = "N"
		Else
			webModCancelFlg = "Y"
		End If
	End If

	' 削除日
	If (IsNull(vRS("削除日"))) Then
		deleteDate = ""
	Else
		deleteDate = CStr(Trim(vRS("削除日")))
		webModCancelFlg = "N"
	End If

	' 2018.12.21 GV add start
	'支払い方法詳細
	If (IsNull(vRS("支払方法詳細"))) Then
		paymentMethodDetail = ""
	Else
		paymentMethodDetail = CStr(vRS("支払方法詳細"))
	End If
	' 2018.12.21 GV add end

	'2020.02.05 GV add start
	'領収書発行フラグ
	receiptFlag = getReceiptFlag(vRS("支払方法"), wOrderNo)

	'領収書発行日
	If (IsNull(vRS("領収書発行日"))) Then
		receiptDate = ""
	Else
		receiptDate = CStr(Trim(vRS("領収書発行日")))
	End If
	'2020.02.05 GV add end

	'2020.03.18 GV add start
	'領収日
	If (IsNull(vRS("領収日"))) Then
		displayReceiptDate = ""
	Else
		displayReceiptDate = CStr(Trim(vRS("領収日")))
	End If
	'2020.03.18 GV add end

	' ギフト顧客番号 2021.06.30 GV add
	If (IsNull(vRS("ギフト顧客番号"))) Then
		giftCustomerNo = 0
	Else
		giftCustomerNo =CStr(vRS("ギフト顧客番号"))
	End If

	' ギフト番号 2021.06.30 GV add
	If (IsNull(vRS("ギフト番号"))) Then
		giftNo = 0
	Else
		giftNo = CStr(vRS("ギフト番号"))
	End If


	With oJSON.data("data")
		.Add "o_no", CStr(Trim(vRS("受注番号")))
		.Add "est_dt", estimateDate
		.Add "o_dt", orderDate
		.Add "ship_comp_dt", shippingDate
		.Add "o_type", CStr(Trim(vRS("受注形態")))
		.Add "pay_method",  CStr(Trim(vRS("支払方法")))
		.Add "pay_method_detail", paymentMethodDetail ' 2018.12.21 GV add
		.Add "furikomi_nm", furikomiMeigi
		.Add "tax_am", CDbl(Trim(vRS("外税合計金額")))
		.Add "total_item_am", CDbl(Trim(vRS("商品合計金額")))
		.Add "ff_charge", CDbl(vRS("送料")) 
		.Add "cod_charge", CDbl(vRS("代引手数料"))
		.Add "kabusoku_am", CDbl(Trim(vRS("過不足相殺金額"))) ' 2015.05.07 GV add
		.Add "total_order_am", CDbl(vRS("受注合計金額"))
		.Add "total_order_am2", totalOrderAmount2 ' 合計金額
		.Add "used_pt", usedPoint ' 利用ポイント
		.Add "comb_ship_flg", CStr(Trim(vRS("一括出荷フラグ")))
		.Add "receipt_name", receiptName
		.Add "receipt_note", receiptNote
'		.Add "web_order_modify_start_date", vRS("Web受注変更開始日")
		.Add "tax_rate", CDbl(vRS("消費税率"))
		.Add "ff_cd", CStr(vRS("運送会社コード"))
		'.Add "tantou_cd", CStr(vRS("担当者コード"))
		.Add "one_time_todokesaki1", one_time_todokesaki1
		.Add "one_time_todokesaki2", one_time_todokesaki2
		.Add "nouki_dt", final_nouki_date_time
		.Add "store_stop", storeStop
		.Add "modifying", webModCancelFlg
		.Add "del_dt", deleteDate
		.Add "hide_ari", hide ' 2016.06.01 GV add
		.Add "receipt_flg", receiptFlag '2020.02.05 GV add
		.Add "receipt_no", CStr(Trim(vRS("領収書番号"))) '2020.02.05 GV add
		.Add "receipt_dt", receiptDate '2020.02.05 GV add
		.Add "display_receipt_dt", displayReceiptDate '2020.03.18 GV add
		.Add "gift_cst_no" , giftCustomerNo 'ギフト顧客番号 2021.06.30 GV add
		.Add "gift_no" , giftNo 'ギフト番号 2021.06.30 GV add
	End With
End If

'レコードセットを閉じる
vRS.Close

'レコードセットのクリア
Set vRS = Nothing

' -------------------------------------------------
' JSONデータの返却
' -------------------------------------------------
' ヘッダ出力
Response.AddHeader "Content-Type", "application/json"
' JSONデータの出力
Response.Write oJSON.JSONoutput()

End Function
%>
