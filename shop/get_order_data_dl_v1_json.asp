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
'	DL用Emax受注　取得API
'
'
'変更履歴
'2019.08.19 GV 新規作成。(DL販売対応)(#1959)
'
'========================================================================
'On Error Resume Next

Dim ConnectionEmax

Dim wErrMsg						' エラーメッセージ (他のページから渡されるメッセージ)
Dim wDispMsg					' 通常メッセージ(エラー以外) (他のページから渡されるメッセージ)
Dim wErrDesc
Dim wMsg						' エラーメッセージ (本ページで作成するメッセージ)
Dim oJSON						' JSONオブジェクト
Dim orderNo						' 受注番号
Dim customerNo					' 顧客番号

'=======================================================================
'	受け渡し情報取り出し & 初期設定
'=======================================================================
' Getパラメータ
customerNo = ReplaceInput(Trim(Request("cno")))
orderNo = ReplaceInput(Trim(Request("ono")))

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

Dim i
Dim j
Dim orderDate
Dim customerEmail
Dim estimateNote
Dim estimateDate
Dim storeStopFlg
Dim furikomiMeigi
Dim ccTotalAm
Dim ccCreditNo
Dim ccSlipNo
Dim receiptFlg
Dim receiptName
Dim customerTel
Dim receiptNote
Dim adCd
Dim econNo
Dim econPayUrl
Dim econTranUrl
Dim orderTel
Dim orderFax
Dim usedPoint
Dim totalOrderAmount2
Dim depositAmount ' 入金合計金額
Dim depositFlag   ' 入金完了フラグ
Dim coupon   ' クーポン
Dim wPaymentMethodDetail
Dim delDate '削除日
Dim makerName
Dim itemName
Dim iro
Dim kikaku

Set oJSON = New aspJSON

'--- ヘッダ部分の情報取出し
vSQL = ""
vSQL = vSQL & "SELECT TOP 1 "
vSQL = vSQL & "  o.受注番号 "							' order_no
vSQL = vSQL & " ,o.顧客番号 "							' customer_no
vSQL = vSQL & " ,o.顧客E_mail "							' email
vSQL = vSQL & " ,o.支払方法 "							' payment_method
vSQL = vSQL & " ,o.運送会社コード "						' freight_forwarder_cd
vSQL = vSQL & " ,o.見積備考 "							' estimate_note
vSQL = vSQL & " ,o.商品合計金額 "						' total_item_amount
vSQL = vSQL & " ,o.送料 "								' freight_charge
vSQL = vSQL & " ,o.代引手数料 "							' daibiki_charge
vSQL = vSQL & " ,o.外税合計金額 "						' total_tax_amount
vSQL = vSQL & " ,o.受注合計金額 "						' total_order_amount
vSQL = vSQL & " ,o.受注日 "								' order_date
vSQL = vSQL & " ,o.見積日 "								' input_date
vSQL = vSQL & " ,o.営業所止めフラグ "					' store_stop_flag
vSQL = vSQL & " ,o.一括出荷フラグ "						' combined_shipping_flag
vSQL = vSQL & " ,o.振込名義人 "							' furikomi_meigi
vSQL = vSQL & " ,o.領収書発行フラグ "					' receipt_flag
vSQL = vSQL & " ,o.領収書宛先 "							' receipt_name
vSQL = vSQL & " ,o.注文者電話番号 "						' customer_tel
vSQL = vSQL & " ,o.領収書但し書き "						' receipt_note
vSQL = vSQL & " ,o.広告コード "							' ad_cd
vSQL = vSQL & " ,o.消費税率 "							' tax_rate
vSQL = vSQL & " ,o.eContext受付番号 "					' e_context_no
vSQL = vSQL & " ,o.eContext支払方法URL "				' e_context_payment_method_url
vSQL = vSQL & " ,o.eContext振込票URL "					' e_context_transfer_url
vSQL = vSQL & " ,o.受注形態 "							' order_type
vSQL = vSQL & " ,o.過不足相殺金額 "						' kabusoku_sousai_amount
vSQL = vSQL & " ,o.注文者名前 "							' order_name
vSQL = vSQL & " ,o.注文者郵便番号 "						' order_postal_cd
vSQL = vSQL & " ,o.注文者都道府県 "						' order_prefecture
vSQL = vSQL & " ,o.注文者住所 "							' order_address
vSQL = vSQL & " ,o.注文者電話番号 "						' order_tel
vSQL = vSQL & " ,o.注文者FAX "							' order_fax
vSQL = vSQL & " ,o.利用ポイント "						' used_point
vSQL = vSQL & " ,o.合計金額 "							' total_used_point_order_amount
vSQL = vSQL & " ,o.入金合計金額 "						' work_order.old_deposit_amount
vSQL = vSQL & " ,o.入金完了フラグ "						' work_order.old_deposit_flag
vSQL = vSQL & " ,o.クーポン "							' coupon
vSQL = vSQL & " ,o.認証アシストフラグ "					' cc_assist_flag 2016.06.22 GV add
vSQL = vSQL & " ,o.セキュア3Dフラグ "					' cc_3d_secure_flag 2016.11.17 GV add
vSQL = vSQL & " ,o.セキュア3D結果コード "				' cc_3d_secure_result_cd 2016.11.17 GV add
vSQL = vSQL & " ,o.支払方法詳細 "						' payment_method_detail 2018.12.21 GV add
vSQL = vSQL & " ,o.削除日 "								' delete_date 2019.08.19 GV add

vSQL = vSQL & "FROM "
vSQL = vSQL & "      " & gLinkServer & "受注     o WITH (NOLOCK) "

vSQL = vSQL & "WHERE "
vSQL = vSQL & "      o.受注番号 = " & orderNo & " "
vSQL = vSQL & "  AND o.顧客番号 = " & customerNo & " "

'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then

	' 受注追加
	oJSON.data.Add "order" ,oJSON.Collection()

	' --------------------
	' 受注日
	If (IsNull(vRS("受注日"))) Then
		orderDate = ""
	Else
		orderDate = CStr(Trim(vRS("受注日")))
	End If

	'顧客E_mail
	If (IsNull(vRS("顧客E_mail"))) Then
		customerEmail = ""
	Else
		customerEmail = CStr(Trim(vRS("顧客E_mail")))
	End If

	'見積備考
	If (IsNull(vRS("見積備考"))) Then
		estimateNote = ""
	Else
		estimateNote = CStr(Trim(vRS("見積備考")))
	End If

	' 見積日
	If (IsNull(vRS("見積日"))) Then
		estimateDate = ""
	Else
		estimateDate = CStr(Trim(vRS("見積日")))
	End If

	'営業所止めフラグ
	If (IsNull(vRS("営業所止めフラグ"))) Then
		storeStopFlg = ""
	Else
		storeStopFlg = CStr(Trim(vRS("営業所止めフラグ")))
	End If

	' 振込名義人
	If (IsNull(vRS("振込名義人"))) Then
		furikomiMeigi = ""
	Else
		furikomiMeigi = CStr(Trim(vRS("振込名義人")))
		'furikomiMeigi = Replace(furikomiMeigi, """", "”")
	End If


	'領収書発行フラグ
	If (IsNull(vRS("領収書発行フラグ"))) Then
		receiptFlg = ""
	Else
		receiptFlg = CStr(Trim(vRS("領収書発行フラグ")))
	End If

	' 領収書宛先
	If (IsNull(vRS("領収書宛先"))) Then
		receiptName = ""
	Else
		receiptName = CStr(Trim(vRS("領収書宛先")))
		receiptName = Replace(receiptName, """", "”")
	End If

	'注文者電話番号
	If (IsNull(vRS("注文者電話番号"))) Then
		customerTel = ""
	Else
		customerTel = CStr(Trim(vRS("注文者電話番号")))
	End If

	' 領収書但し書き
	If (IsNull(vRS("領収書但し書き"))) Then
		receiptNote = ""
	Else
		receiptNote = CStr(Trim(vRS("領収書但し書き")))
		receiptNote = Replace(receiptNote, """", "”")
	End If

	'広告コード
	If (IsNull(vRS("広告コード"))) Then
		adCd = ""
	Else
		adCd = CStr(Trim(vRS("広告コード")))
	End If

	'eContext受付番号
	If (IsNull(vRS("eContext受付番号"))) Then
		econNo = ""
	Else
		econNo = CStr(Trim(vRS("eContext受付番号")))
	End If

	'eContext支払方法URL
	If (IsNull(vRS("eContext支払方法URL"))) Then
		econPayUrl = ""
	Else
		econPayUrl = CStr(Trim(vRS("eContext支払方法URL")))
	End If

	'eContext振込票URL
	If (IsNull(vRS("eContext振込票URL"))) Then
		econTranUrl = ""
	Else
		econTranUrl = CStr(Trim(vRS("eContext振込票URL")))
	End If

	'注文者電話番号
	If (IsNull(vRS("注文者電話番号"))) Then
		orderTel = ""
	Else
		orderTel = CStr(Trim(vRS("注文者電話番号")))
	End If

	'注文者FAX
	If (IsNull(vRS("注文者FAX"))) Then
		orderFax = ""
	Else
		orderFax = CStr(Trim(vRS("注文者FAX")))
	End If

	' 利用ポイント
	If (IsNull(vRS("利用ポイント"))) Then
		usedPoint = 0
	Else
		usedPoint = CDbl(vRS("利用ポイント"))
	End If

	' 合計金額
	If (IsNull(vRS("合計金額"))) Then
		totalOrderAmount2 = 0
	Else
		totalOrderAmount2 = CDbl(vRS("合計金額"))
	End If

	'入金完了フラグ
	If (IsNull(vRS("入金完了フラグ"))) Then
		depositFlag = ""
	Else
		depositFlag = CStr(Trim(vRS("入金完了フラグ")))
	End If

	' 入金合計金額
	If (IsNull(vRS("入金合計金額"))) Then
		depositAmount = 0
	Else
		depositAmount = CDbl(vRS("入金合計金額"))
	End If

	'クーポン
	If (IsNull(vRS("クーポン"))) Then
		coupon = ""
	Else
		coupon = CStr(Trim(vRS("クーポン")))
		coupon = Replace(coupon, """", "”")
	End If


	'支払方法詳細 2020.03.10 GV add
	If (IsNull(vRS("支払方法詳細"))) Then
		wPaymentMethodDetail = ""
	Else
		wPaymentMethodDetail = CStr(Trim(vRS("支払方法詳細")))
	End If

	' 削除日 2019.08.19 GV add
	If (IsNull(vRS("削除日"))) Then
		delDate = ""
	Else
		delDate = CStr(Trim(vRS("削除日")))
	End If

	With oJSON.data("order")
		.Add "o_no", CStr(Trim(vRS("受注番号")))
		.Add "cstm_mail", customerEmail
		.Add "pay_method",  CStr(Trim(vRS("支払方法")))
		.Add "pay_method_detail",  wPaymentMethodDetail '2020.03.10 GV add
		.Add "ff_cd", CStr(vRS("運送会社コード"))
		.Add "est_nt", estimateNote
		.Add "total_item_am", CDbl(Trim(vRS("商品合計金額")))
		.Add "ff_charge", CDbl(vRS("送料")) 
		.Add "cod_charge", CDbl(vRS("代引手数料"))
		.Add "tax_am", CDbl(vRS("外税合計金額"))
		.Add "total_order_am", CDbl(vRS("受注合計金額"))
		.Add "est_dt", estimateDate '見積日(input_date)
		.Add "store_stop", storeStopFlg '営業所止めフラグ(store_stop_flag)
		.Add "comb_ship_flg", CStr(Trim(vRS("一括出荷フラグ"))) ' combined_shipping_flag
		.Add "furikomi_nm", furikomiMeigi '振込名義人(furikomi_meigi)
		.Add "receipt_flg", receiptFlg '領収書発行フラグ(receipt_flag)
		.Add "receipt_nm", receiptName '領収書宛先(receipt_name)
		.Add "cstm_tel", customerTel '注文者電話番号(customer_tel)
		.Add "receipt_nt", receiptNote '領収書但し書き(receipt_note)
		.Add "ad_cd", adCd '広告コード(ad_cd)
		.Add "tax_rate", CDbl(vRS("消費税率"))
		.Add "econ_no", econNo 'eContext受付番号(e_context_no)
		.Add "econ_pay", econPayUrl 'eContext支払方法URL(e_context_payment_method_url)
		.Add "econ_tran", econTranUrl 'eContext振込票URL(e_context_transfer_url)
		.Add "o_type", CStr(Trim(vRS("受注形態")))
		.Add "kabusoku_am", CDbl(Trim(vRS("過不足相殺金額"))) ' 2015.05.07 GV add
		.Add "o_nm", CStr(Trim(vRS("注文者名前")))
		.Add "o_zip", CStr(Trim(vRS("注文者郵便番号")))
		.Add "o_pref", CStr(Trim(vRS("注文者都道府県")))
		.Add "o_addr", CStr(Trim(vRS("注文者住所")))
		.Add "o_tel", orderTel '注文者電話番号(order_tel)
		.Add "o_fax", orderFax '注文者FAX(order_fax)
		.Add "used_pt", usedPoint ' 利用ポイント
		.Add "total_order_am2", totalOrderAmount2 ' 合計金額(total_used_point_order_amount)
		.Add "deposit_flg", depositFlag ' 入金完了フラグ(work_order.old_deposit_flag)
		.Add "deposit_am", depositAmount ' 入金合計金額(work_order.old_deposit_amount)
		.Add "coupon", coupon ' クーポン(work_order.coupon)
		.Add "o_dt", orderDate '受注日 2016.09.06 GV add
		.Add "del_dt", delDate '削除日 2019.08.19 GV add
	End With
End If

'レコードセットを閉じる
vRS.Close

'レコードセットのクリア
'Set vRS = Nothing

' -----------------------------------
' 受注明細部分
' -----------------------------------
' JSONオブジェクト追加
oJSON.data.Add "detail" ,oJSON.Collection()

vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "  od.受注番号 "
vSQL = vSQL & ", od.今回限り届先名前 "
vSQL = vSQL & ", od.受注明細番号 "
vSQL = vSQL & ", od.メーカーコード "
vSQL = vSQL & ", mk.メーカー名 "
vSQL = vSQL & ", od.商品コード "
vSQL = vSQL & ", od.商品名 "
vSQL = vSQL & ", od.色 "
vSQL = vSQL & ", od.規格 "
vSQL = vSQL & ", iz.商品ID "
vSQL = vSQL & ", od.メーカー直送フラグ "
vSQL = vSQL & ", od.受注単価 "
vSQL = vSQL & ", od.受注金額 "
vSQL = vSQL & ", od.受注数量 "
vSQL = vSQL & ", od.出荷指示合計数量 "
vSQL = vSQL & ", od.受注明細備考 "
vSQL = vSQL & ", od.適正在庫数量 "
vSQL = vSQL & " FROM "
vSQL = vSQL & " 受注明細 od WITH (NOLOCK) "

vSQL = vSQL & "INNER JOIN 受注 o WITH (NOLOCK) "
vSQL = vSQL & "  ON o.受注番号 = od.受注番号 "

vSQL = vSQL & "INNER JOIN 色規格別在庫 iz WITH (NOLOCK) "
vSQL = vSQL & "   ON iz.メーカーコード = od.メーカーコード "
vSQL = vSQL & "  AND iz.商品コード = od.商品コード "
vSQL = vSQL & "  AND iz.色 = od.色 "
vSQL = vSQL & "  AND iz.規格 = od.規格 "

vSQL = vSQL & "INNER JOIN メーカー mk WITH (NOLOCK) "
vSQL = vSQL & "   ON mk.メーカーコード = od.メーカーコード "


vSQL = vSQL & "WHERE "
vSQL = vSQL & "      o.受注番号 = " & orderNo & " "
vSQL = vSQL & "  AND o.顧客番号 = " & customerNo & " "

vSQL = vSQL & "ORDER BY "
vSQL = vSQL & "  od.受注明細番号 ASC "


Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic
'
''レコードが存在している場合
If vRS.EOF = False Then
	j = 0

	For i = 0 To (vRS.RecordCount - 1)
		makerName = Replace(Trim(vRS("メーカー名")), """", "”")
		makerName = CStr(makerName)

		itemName = Replace(Trim(vRS("商品名")), """", "”")
		itemName = CStr(itemName)

		iro = Replace(Trim(vRS("色")), """", "”")
		iro = CStr(iro)

		kikaku = Replace(Trim(vRS("規格")), """", "”")
		kikaku = CStr(kikaku)

		'--- 明細行生成
		With oJSON.data("detail")
			.Add j ,oJSON.Collection()
			With .item(j)
				.Add "od_no" ,CStr(Trim(vRS("受注明細番号")))
				.Add "m_cd" ,CStr(Trim(vRS("メーカーコード")))
				.Add "m_name" ,makerName
				.Add "i_cd" ,CStr(Trim(vRS("商品コード")))
				.Add "i_name" ,itemName
				.Add "iro" ,iro
				.Add "kikaku" ,kikaku
				.Add "i_id" ,CStr(Trim(vRS("商品ID")))
				.Add "i_suu", CDbl(vRS("受注数量")) 
			End With
		End With

		' 次のレコード行へ移動
		vRS.MoveNext

		If vRS.EOF Then
			Exit For
		End If

		j = j + 1
	Next
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
Response.AddHeader "X-Content-Type-Options", "nosniff"

' JSONデータの出力
Response.Write oJSON.JSONoutput()

End Function
%>
