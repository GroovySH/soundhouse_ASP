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
'	Emax受注　取得API
'
'
'変更履歴
'2016/03/29 GV 新規作成
'2016.06.22 GV 認証アシスト使用フラグ対応。
'2016.09.06 GV キャンセル時の引当数戻し処理の改修対応。
'2016.11.17 GV 3Dセキュア対応
'
'========================================================================
'On Error Resume Next

Dim ConnectionEmax

Dim wErrMsg						' エラーメッセージ (他のページから渡されるメッセージ)
Dim wDispMsg					' 通常メッセージ(エラー以外) (他のページから渡されるメッセージ)
Dim wErrDesc
Dim wMsg						' エラーメッセージ (本ページで作成するメッセージ)
Dim wUserID

Dim oJSON						' JSONオブジェクト
Dim wOrderNo					' 受注番号

'=======================================================================
'	受け渡し情報取り出し & 初期設定
'=======================================================================
' Getパラメータ
wUserID = ReplaceInput(Trim(Request("cno")))
wOrderNo = ReplaceInput(Trim(Request("ono")))

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
Dim customerEmail
Dim estimateNote
Dim estimateDate
Dim shipSodFlg
Dim noukiDt
Dim noukiTm
Dim storeStopFlg
Dim furikomiMeigi
Dim ccTotalAm
Dim ccCreditNo
Dim ccSlipNo
Dim receiptFlg
Dim receiptName
Dim ritouFlg
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
Dim ccAssist '認証アシストフラグ 2016.06.22 GV add
Dim cc3dSecure '3Dセキュアフラグ 2016.11.17 GV add
Dim cc3dSecureResult '3Dセキュア結果コード 2016.11.17 GV add

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
vSQL = vSQL & " ,od.届先住所連番 "						' todokesaki_address_renban
vSQL = vSQL & " ,od.今回限り届先名前 "					' todokesaki_name
vSQL = vSQL & " ,od.今回限り届先郵便番号 "				' todokesaki_postal_cd
vSQL = vSQL & " ,od.今回限り届先都道府県 "				' todokesaki_prefecture
vSQL = vSQL & " ,od.今回限り届先住所 "					' todokesaki_address
vSQL = vSQL & " ,od.今回限り届先電話番号 "				' todokesaki_tel
vSQL = vSQL & " ,od.今回限り届先納品書送付可フラグ "	' todokesaki_nouhinsho_send_flag
vSQL = vSQL & " ,od.最終指定納期 "						' nouki_date
vSQL = vSQL & " ,od.最終時間指定 "						' nouki_time
vSQL = vSQL & " ,o.営業所止めフラグ "					' store_stop_flag
vSQL = vSQL & " ,o.一括出荷フラグ "						' combined_shipping_flag
vSQL = vSQL & " ,o.振込名義人 "							' furikomi_meigi
vSQL = vSQL & " ,cc.カード支払金額 "					' card_total_amount
vSQL = vSQL & " ,cc.カード与信確認番号 "				' card_credit_no
vSQL = vSQL & " ,cc.カードネット伝票番号 "				' card_net_slip_no
vSQL = vSQL & " ,o.領収書発行フラグ "					' receipt_flag
vSQL = vSQL & " ,o.領収書宛先 "							' receipt_name
vSQL = vSQL & " ,od.離島フラグ "						' ritou_flag
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


vSQL = vSQL & "FROM "
vSQL = vSQL & "      " & gLinkServer & "受注     o WITH (NOLOCK) "
vSQL = vSQL & "INNER JOIN " & gLinkServer & "受注明細 od WITH (NOLOCK) "
vSQL = vSQL & "   ON od.受注番号 = o.受注番号 "
vSQL = vSQL & "  AND od.受注数量 > 0 "

vSQL = vSQL & "LEFT JOIN " & gLinkServer & "受注カード情報 cc WITH (NOLOCK) "
vSQL = vSQL & "  ON cc.受注番号 = o.受注番号 "

vSQL = vSQL & "WHERE "
vSQL = vSQL & "      o.受注番号 = " & wOrderNo & " "
vSQL = vSQL & "  AND o.顧客番号 = " & wUserID & " "

vSQL = vSQL & " ORDER BY "
vSQL = vSQL & "        od.受注明細番号 ASC "


'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then

	' リスト追加
	oJSON.data.Add "data" ,oJSON.Collection()

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

	'今回限り届先納品書送付可フラグ
	If (IsNull(vRS("今回限り届先納品書送付可フラグ"))) Then
		shipSodFlg = ""
	Else
		shipSodFlg = CStr(Trim(vRS("今回限り届先納品書送付可フラグ")))
	End If

	' 配送日
	If (IsNull(vRS("最終指定納期"))) Then
		noukiDt = ""
	Else
		noukiDt = CStr(Trim(vRS("最終指定納期")))
	End If

	' 配送時間帯
	If (IsNull(vRS("最終指定納期"))) Then
		noukiTm = ""
	Else
		noukiTm = CStr(Trim(vRS("最終時間指定"))) 
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

	' カード支払金額
	If (IsNull(vRS("カード支払金額"))) Then
		ccTotalAm = ""
	Else
		ccTotalAm = CStr(CDbl(vRS("カード支払金額")))
	End If

	'カード与信確認番号
	If (IsNull(vRS("カード与信確認番号"))) Then
		ccCreditNo = ""
	Else
		ccCreditNo = CStr(Trim(vRS("カード与信確認番号")))
	End If

	'カードネット伝票番号
	If (IsNull(vRS("カードネット伝票番号"))) Then
		ccSlipNo = ""
	Else
		ccSlipNo = CStr(Trim(vRS("カードネット伝票番号")))
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

	'離島フラグ
	If (IsNull(vRS("離島フラグ"))) Then
		ritouFlg = ""
	Else
		ritouFlg = CStr(Trim(vRS("離島フラグ")))
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

	'2016.06.22 GV add start
	'認証アシストフラグ
	If (IsNull(vRS("認証アシストフラグ"))) Then
		ccAssist = ""
	Else
		ccAssist = CStr(Trim(vRS("認証アシストフラグ")))
	End If
	'2016.06.22 GV add end

	'2016.11.17 GV add start
	'3Dセキュアフラグ
	If (IsNull(vRS("セキュア3Dフラグ"))) Then
		cc3dSecure = ""
	Else
		cc3dSecure = CStr(Trim(vRS("セキュア3Dフラグ")))
	End If

	'3Dセキュア結果コード
	If (IsNull(vRS("セキュア3D結果コード"))) Then
		cc3dSecureResult = ""
	Else
		cc3dSecureResult = CStr(Trim(vRS("セキュア3D結果コード")))
	End If
	'2016.11.17 GV add end

	With oJSON.data("data")
		.Add "o_no", CStr(Trim(vRS("受注番号")))
		.Add "cstm_mail", customerEmail
		.Add "pay_method",  CStr(Trim(vRS("支払方法")))
		.Add "ff_cd", CStr(vRS("運送会社コード"))
		.Add "est_nt", estimateNote
		.Add "total_item_am", CDbl(Trim(vRS("商品合計金額")))
		.Add "ff_charge", CDbl(vRS("送料")) 
		.Add "cod_charge", CDbl(vRS("代引手数料"))
		.Add "tax_am", CDbl(vRS("外税合計金額"))
		.Add "total_order_am", CDbl(vRS("受注合計金額"))
		.Add "est_dt", estimateDate '見積日(input_date)
		.Add "ship_addr_no", CDbl(vRS("届先住所連番"))
		.Add "ship_sod_flg", shipSodFlg '今回限り届先納品書送付可フラグ(todokesaki_nouhinsho_send_flag)
		.Add "nouki_dt", noukiDt '最終指定納期(nouki_date)
		.Add "nouki_tm", noukiTm '最終時間指定(nouki_time)
		.Add "store_stop", storeStopFlg '営業所止めフラグ(store_stop_flag)
		.Add "comb_ship_flg", CStr(Trim(vRS("一括出荷フラグ"))) ' combined_shipping_flag
		.Add "furikomi_nm", furikomiMeigi '振込名義人(furikomi_meigi)
		.Add "cc_pay_am", ccTotalAm 'カード支払金額(card_total_amount)
		.Add "cc_c_no", ccCreditNo 'カード与信確認番号(card_credit_no)
		.Add "cc_slip", ccSlipNo 'カードネット伝票番号(card_net_slip_no)
		.Add "receipt_flg", receiptFlg '領収書発行フラグ(receipt_flag)
		.Add "receipt_nm", receiptName '領収書宛先(receipt_name)
		.Add "ritou", ritouFlg '離島フラグ(ritou_flag)
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
		.Add "cc_assist", ccAssist ' 認証アシストフラグ(work_order.old_cc_assist_flag) 2016.06.22 GV add
		.Add "o_dt", orderDate '受注日 2016.09.06 GV add
		.Add "cc_3d_secure", cc3dSecure ' 3Dセキュアフラグ 2016.11.17 GV add
		.Add "cc_3d_secure_result", cc3dSecureResult ' 3Dセキュア結果コード 2016.11.17 GV add
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
Response.AddHeader "X-Content-Type-Options", "nosniff"

' JSONデータの出力
Response.Write oJSON.JSONoutput()

End Function
%>
