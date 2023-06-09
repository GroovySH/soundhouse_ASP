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
'2020.02.27 GV orderhistorylist_v2_json.aspをベースに新規作成。(PayPay対応)(#2405)
'2020.06.01 GV 修正。
'2020.06.03 GV PayPay返金改修。(#2440)
'2021.01.04 GV メンテナンスポータルPayPay返金機能改修。(#2647)
'2021.01.21 GV メンテナンスポータルPayPay返金機能改修。(#2662)
'
'========================================================================
'On Error Resume Next

Const PAGE_SIZE = 20			' 1ページあたりの表示行数

Dim ConnectionEmax

Dim wErrMsg						' エラーメッセージ (他のページから渡されるメッセージ)
Dim wDispMsg					' 通常メッセージ(エラー以外) (他のページから渡されるメッセージ)
Dim wErrDesc
Dim wMsg						' エラーメッセージ (本ページで作成するメッセージ)
Dim wCustomerNo					' 顧客番号
Dim wOrderNo					' 受注番号
Dim wFlg						' 実行フラグ
Dim wIPage						' 表示するページ位置 (パラメータ)
Dim estimateStartDate			' 検索期間自
Dim estimateEndDate				' 検索期間至
Dim oJSON						' JSONオブジェクト
Dim wOrderHidden				' 非表示フラグ
Dim wOrderCancelled				' キャンセル注文フラグ
Dim wOrderShipping				' 未発送注文フラグ
Dim wSlipNo						' 送り状番号
Dim wReceipt					' 領収書
Dim wDepositTerm				' 入金確認期限（日）
Dim wPaypayPaymentId			' PayPay決済番号(カード与信確認番号)

'=======================================================================
'	受け渡し情報取り出し & 初期設定
'=======================================================================
wFlg = True

' Getパラメータ
' 顧客番号
wCustomerNo = ReplaceInput_NoCRLF(Trim(Request("cno")))
' 数値のみチェック (ASPは全角でも数字ならTrueを返す)
'If (IsNumeric(wCustomerNo) = False) Or (cf_checkNumeric(wCustomerNo) = False) Then
'	wFlg = False
'End If
If (IsNull(wCustomerNo) = False) And (wCustomerNo <> "") Then
	If (IsNumeric(wCustomerNo) = False) Or (cf_checkNumeric(wCustomerNo) = False) Then
		wFlg = False
	End If
End If



'ページ番号
wIPage = ReplaceInput_NoCRLF(Trim(Request("page")))
' 数値のみチェック (ASPは全角でも数字ならTrueを返す)
If (IsNumeric(wIPage) = False) Or (cf_checkNumeric(wIPage) = False) Then
	wIPage = 1
Else
	wIPage = CLng(wIPage)
End If

'検索期間自
estimateStartDate = ReplaceInput_NoCRLF(Trim(Request("est_from")))
estimateStartDate = CStr(estimateStartDate)

'検索期間至
estimateEndDate = ReplaceInput_NoCRLF(Trim(Request("est_to")))
estimateEndDate = CStr(estimateEndDate)


'受注番号
wOrderNo = ReplaceInput_NoCRLF(Trim(Request("ono")))
' 数値のみチェック (ASPは全角でも数字ならTrueを返す)
If (IsNumeric(wOrderNo) = False) Or (cf_checkNumeric(wOrderNo) = False) Then
	wOrderNo = null
Else
	wOrderNo = CLng(wOrderNo)
End If

'PayPay決済番号
wPaypayPaymentId = ReplaceInput_NoCRLF(Trim(Request("pay_id")))
wPaypayPaymentId = CStr(wPaypayPaymentId)

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
Dim vWHERE
Dim i
Dim j
Dim vRS
Dim orderDate
Dim deleteDate
Dim orderName
Dim customerName
Dim shippingCompDate
Dim allCount
Dim orderTotalAm2
Dim usedPoint
Dim dateTerm
Dim orderType
Dim webModCancelFlg
Dim depositFlag ' 入金完了フラグ
Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2
Dim wPaymentMethodDetail
Dim ccTotalAm
Dim ccCreditNo
Dim ccSlipNo
Dim isEstimateDateExist
Dim totalAmAtAuth 'オーソリ時受注合計金額 2021.01.04 GV add

'2020.06.01 GV add
Dim totalAmAtOrder
Dim usedPointAtOrder
Dim kabusokuAmAtOrder
'2020.06.01 GV add

Dim shipSuuSum '2020.06.03 GV add

Set oJSON = New aspJSON

' 初期化
i = 0
j = 0
allCount = 0
dateTerm = ""

' 受注形態(カンマ区切りで指定)
orderType = ""
orderType = orderType & "  'インターネット'"
'orderType = orderType & " ,'E-mail'"
'orderType = orderType & " ,'FAX'"
'orderType = orderType & " ,'携帯'"
'orderType = orderType & " ,'電話'"
'orderType = orderType & " ,'郵送'"
'orderType = orderType & " ,'来店'"
orderType = orderType & " ,'スマートフォン'"

'コントロールマスタから見積もり有効期限を取得 2018.01.12 GV add
call getEmaxCntlMst("受注","入金確認待ち期限","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
If (IsNull(wItemNum1)) Then
	wDepositTerm = 10
Else
	wDepositTerm = wItemNum1
End If


' 入力値が正常の場合
If (wFlg = True) Then
	' ---------------------------
	'検索期間自
	If ((IsNull(estimateStartDate) = True) Or (estimateStartDate = "")) Then
	ElseIf ((IsNull(estimateStartDate) = False) Or (estimateStartDate <> "")) Then
		dateTerm = dateTerm & " AND o1.見積日 >= '" & estimateStartDate & " 00:00:00' "
	End If

	'検索期間至
	If ((IsNull(estimateEndDate) = True) Or (estimateEndDate = "")) Then
	ElseIf ((IsNull(estimateEndDate) = False) Or (estimateEndDate <> "")) Then
		dateTerm = dateTerm & " AND o1.見積日 <= '" & estimateEndDate & " 23:59:59' "
	End If

	' ---------------------------
	' 総数を取得
	vSQL = ""
	vSQL = vSQL & "SELECT count(o.受注番号) AS cnt "
	vSQL = vSQL & "FROM "
	vSQL = vSQL & " (SELECT DISTINCT "
	vSQL = vSQL & "   o1.顧客番号 "
	vSQL = vSQL & "  ,o1.受注番号 "
	vSQL = vSQL & "  ,o1.見積日 "
	vSQL = vSQL & "  ,o1.削除日 "

	vSQL = vSQL & "  FROM 受注 AS o1 WITH (NOLOCK) "
	vSQL = vSQL & "      LEFT JOIN 受注カード情報 AS cc WITH (NOLOCK) "
	vSQL = vSQL & "        ON cc.受注番号 = o1.受注番号 "

'	vSQL = vSQL & "  WHERE o1.顧客番号 = " & wCustomerNo & " "
'	vSQL = vSQL & "    AND o1.受注形態 IN (" & orderType & ") "
	vSQL = vSQL & "  WHERE 1=1 "

	' 顧客番号
	If (IsNull(wCustomerNo) = False) And (wCustomerNo <> "") Then
		vSQL = vSQL & " AND o1.顧客番号 = " & wCustomerNo
	End If

	' 受注番号
	If (IsNull(wOrderNo) = False) And (wOrderNo <> "") Then
		vSQL = vSQL & " AND o1.受注番号 = " & wOrderNo
	End If

	vSQL = vSQL & "    AND o1.受注形態 IN (" & orderType & ") "
	vSQL = vSQL & "    AND o1.支払方法 = 'クレジットカード' "
'	vSQL = vSQL & "    AND o1.支払方法詳細 = '03' " ' 03は店舗用
	vSQL = vSQL & "    AND o1.支払方法詳細 = '05' "


	'PayPay決済番号
	'If (IsNull(wPaypayPaymentId) = False) Or (wPaypayPaymentId <> "") Then
	If (wPaypayPaymentId <> "") Then
		vSQL = vSQL & " AND cc.カード与信確認番号 = '" & wPaypayPaymentId & "' "  ' card_credit_no
	End If

	vSQL = vSQL & dateTerm

	vSQL = vSQL & " ) AS o "

	'検索期間を結合
	'@@@@Response.Write vSQL & "<br>"


	Set vRS = Server.CreateObject("ADODB.Recordset")
	vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	'レコードが存在している場合
	If vRS.EOF = False Then
		allCount = vRS("cnt")
	End If

	'--- 該当顧客の受注一覧取り出し
	vSQL = ""
	vSQL = vSQL & "SELECT "
	vSQL = vSQL & "  o.* "
	vSQL = vSQL & "  , (CASE WHEN o.受注日 IS NOT NULL AND o.削除日 IS NULL THEN 'Y' ELSE 'N' END ) AS 注残 "
	vSQL = vSQL & "FROM ("
	vSQL = vSQL & "  SELECT "
	vSQL = vSQL & "    * "
	vSQL = vSQL & "  FROM ("
	vSQL = vSQL & "    SELECT "
	vSQL = vSQL & "      ROW_NUMBER() OVER(ORDER BY o2.見積日 DESC) AS RN "
	vSQL = vSQL & "      ,o2.* "
	vSQL = vSQL & "    FROM ("
	vSQL = vSQL & "      SELECT DISTINCT "
	vSQL = vSQL & "         o1.受注番号 "
	vSQL = vSQL & "        , o1.顧客番号 "
	vSQL = vSQL & "        , o1.注文者名前 "
	vSQL = vSQL & "        , c.顧客名 "
	vSQL = vSQL & "        , o1.受注日 "
	vSQL = vSQL & "        , o1.見積日 "
	vSQL = vSQL & "        , o1.削除日 "
	vSQL = vSQL & "        , o1.出荷完了日 "
	vSQL = vSQL & "        , o1.消費税率 "
	vSQL = vSQL & "        , o1.受注形態 "
	vSQL = vSQL & "        , o1.支払方法 "
	vSQL = vSQL & "        , o1.支払方法詳細 "
	vSQL = vSQL & "        , o1.受注合計金額 "
	vSQL = vSQL & "        , o1.合計金額 "
'2020.06.01 GV add start
	vSQL = vSQL & "        , o1.受注時送料 "
	vSQL = vSQL & "        , o1.受注時代引手数料 "
	vSQL = vSQL & "        , o1.受注時過不足相殺金額 "
	vSQL = vSQL & "        , o1.受注時利用ポイント "
	vSQL = vSQL & "        , o1.受注時合計金額 "
'2020.06.01 GV add end
	vSQL = vSQL & "        , o1.利用ポイント "
	vSQL = vSQL & "        , o1.Web注文変更キャンセル中フラグ "
	vSQL = vSQL & "        , o1.その他合計金額 "
	vSQL = vSQL & "        , o1.入金完了フラグ "
'	vSQL = vSQL & "        ,cc.カード支払金額 "			' card_total_amount
	vSQL = vSQL & "        ,cc.カード与信確認番号 "		' card_credit_no
	vSQL = vSQL & "        ,cc.カードネット伝票番号 "	' card_net_slip_no
'	vSQL = vSQL & "        ,(SELECT TOP 1 "
'	vSQL = vSQL & "            cc1.カード与信確認番号 "
'	vSQL = vSQL & "          FROM 受注カード情報 AS cc1 WITH (NOLOCK) "
'	vSQL = vSQL & "          WHERE "
'	vSQL = vSQL & "            cc1.受注番号 = o1.受注番号 "
'	vSQL = vSQL & "         ) as カード与信確認番号 "
'	vSQL = vSQL & "        ,(SELECT TOP 1 "
'	vSQL = vSQL & "            cc2.カードネット伝票番号 "
'	vSQL = vSQL & "          FROM 受注カード情報 AS cc2 WITH (NOLOCK) "
'	vSQL = vSQL & "          WHERE "
'	vSQL = vSQL & "            cc2.受注番号 = o1.受注番号 "
'	vSQL = vSQL & "         ) as カードネット伝票番号 "
' 2020.06.03 GV add start
	vSQL = vSQL & "        ,(SELECT "
'	vSQL = vSQL & "            sum(od1.出荷指示合計数量)  "
	vSQL = vSQL & "            sum(od1.出荷合計数量)  "
	vSQL = vSQL & "          FROM "
	vSQL = vSQL & "            受注明細 AS od1 WITH (NOLOCK) "
	vSQL = vSQL & "          WHERE "
	vSQL = vSQL & "            od1.受注番号 = o1.受注番号 "
'	vSQL = vSQL & "         ) as 出荷指示合計数量 "
	vSQL = vSQL & "         ) as 出荷合計数量 "
' 2020.06.03 GV add end
' 2021.01.04 GV add start
	vSQL = vSQL & "        ,(SELECT "
	vSQL = vSQL & "            オーソリ時受注合計金額  "
	vSQL = vSQL & "          FROM "
	vSQL = vSQL & "            受注カード情報 AS oc1 WITH (NOLOCK) "
	vSQL = vSQL & "          WHERE "
	vSQL = vSQL & "            oc1.受注番号 = o1.受注番号 "
	vSQL = vSQL & "            AND oc1.変更区分 = '削除' "
	vSQL = vSQL & "            AND oc1.オーソリ時受注合計金額 > 0 " '2021.01.12 GV add
	vSQL = vSQL & "         ) as オーソリ時受注合計金額 "
' 2021.01.04 GV add end

	vSQL = vSQL & "      FROM "
	vSQL = vSQL & "        受注 AS o1 WITH (NOLOCK)  "
	vSQL = vSQL & "      INNER JOIN 顧客 AS c WITH (NOLOCK) "
	vSQL = vSQL & "        ON c.顧客番号 = o1.顧客番号 "
	vSQL = vSQL & "      LEFT JOIN 受注カード情報 AS cc WITH (NOLOCK) "
	vSQL = vSQL & "        ON cc.受注番号 = o1.受注番号 "
	vSQL = vSQL & "      WHERE 1 = 1 "
'	vSQL = vSQL & "        o1.顧客番号 = " & wCustomerNo

	' 顧客番号
	If (IsNull(wCustomerNo) = False) And (wCustomerNo <> "") Then
		vSQL = vSQL & "     AND o1.顧客番号 = " & wCustomerNo
	End If

	' 受注番号
	If (IsNull(wOrderNo) = False) And (wOrderNo <> "") Then
		vSQL = vSQL & "     AND o1.受注番号 = " & wOrderNo
	End If

	vSQL = vSQL & "        AND o1.受注形態 IN (" & orderType & " ) "
'	vSQL = vSQL & "        AND o1.支払方法 = 'クレジットカード' AND o1.支払方法詳細 = '03' " ' 03 は店舗用
	vSQL = vSQL & "        AND o1.支払方法 = 'クレジットカード' AND o1.支払方法詳細 = '05' "

	' 受注番号
'	If (IsNull(wOrderNo) = False) Or (wOrderNo <> "") Then
'		vSQL = vSQL & "        AND o1.受注番号 = " & wOrderNo
'	End If

	'PayPay決済番号
	'If (IsNull(wPaypayPaymentId) = False) Or (wPaypayPaymentId <> "") Then
	If (wPaypayPaymentId <> "") Then
		vSQL = vSQL & "        AND cc.カード与信確認番号 = '" & wPaypayPaymentId & "' "  ' card_credit_no
	End If

	vSQL = vSQL & dateTerm
	vSQL = vSQL & "    ) AS o2 "
	vSQL = vSQL & "  ) as o3 "
'	vSQL = vSQL & "  WHERE RN BETWEEN 1 AND 10 "
	vSQL = vSQL & "  WHERE RN BETWEEN " & ((PAGE_SIZE * (wIPage - 1)) + 1) & " AND " & (PAGE_SIZE * wIPage)
	vSQL = vSQL & ") AS o "
	vSQL = vSQL & "  ORDER BY "
	vSQL = vSQL & "  見積日 DESC "
	vSQL = vSQL & "  ,受注番号 desc "

	'@@@@Response.Write(vSQL) & "<br>"

	Set vRS = Server.CreateObject("ADODB.Recordset")
	vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	'レコードが存在している場合
	If vRS.EOF = False Then

		' 全件数をJSONデータにセット
		oJSON.data.Add "count" ,allCount

		' ページ番号
		oJSON.data.Add "page" ,wIPage

		' ページあたりの行数
		oJSON.data.Add "page_size" ,PAGE_SIZE

		' リスト追加
		oJSON.data.Add "list" ,oJSON.Collection()

		For i = 0 To (vRS.RecordCount - 1)
			' 受注日
			If (IsNull(vRS("受注日"))) Then
				orderDate = ""
			Else
				orderDate = CStr(Trim(vRS("受注日")))
			End If

			' 出荷完了日
			If (IsNull(vRS("出荷完了日"))) Then
				shippingCompDate = ""
			Else
				shippingCompDate = CStr(Trim(vRS("出荷完了日")))
			End If

			' 削除日
			If (IsNull(vRS("削除日"))) Then
				deleteDate = ""
			Else
				deleteDate = CStr(Trim(vRS("削除日")))
			End If

			'注文者名前
			If (IsNull(vRS("注文者名前"))) Then
				orderName = ""
			Else
				orderName = CStr(Trim(vRS("注文者名前")))
			End If

			'顧客名
			If (IsNull(vRS("顧客名"))) Then
				customerName = ""
			Else
				customerName = CStr(Trim(vRS("顧客名")))
			End If

			If (IsNull(vRS("合計金額"))) Then
				orderTotalAm2 = 0
			Else
				orderTotalAm2 = CDbl(vRS("合計金額"))
			End If

			If (IsNull(vRS("受注時過不足相殺金額"))) Then
				kabusokuAmAtOrder = 0
			Else
				kabusokuAmAtOrder = CDbl(vRS("受注時過不足相殺金額"))
			End If

			If (IsNull(vRS("受注時利用ポイント"))) Then
				usedPointAtOrder = 0
			Else
				usedPointAtOrder = CDbl(vRS("受注時利用ポイント"))
			End If

			If (IsNull(vRS("受注時合計金額"))) Then
				totalAmAtOrder = 0
			Else
				totalAmAtOrder = CDbl(vRS("受注時合計金額"))
			End If


			'20201.01.04 GV add start
			If (IsNull(vRS("オーソリ時受注合計金額"))) Then
				totalAmAtAuth = 0
			Else
				totalAmAtAuth = CDbl(vRS("オーソリ時受注合計金額"))
			End If
			'20201.01.04 GV add end

			' 利用ポイント
			If (IsNull(vRS("利用ポイント"))) Then
				usedPoint = 0
			Else
				usedPoint = CDbl(vRS("利用ポイント"))
			End If

			'入金完了フラグ
			If (IsNull(vRS("入金完了フラグ"))) Then
				depositFlag = ""
			Else
				depositFlag = CStr(Trim(vRS("入金完了フラグ")))
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

			'支払い方法詳細
			If (IsNull(vRS("支払方法詳細"))) Then
				wPaymentMethodDetail = ""
			Else
				wPaymentMethodDetail = CStr(vRS("支払方法詳細"))
			End If

			' カード支払金額
			'If (IsNull(vRS("カード支払金額"))) Then
			'	ccTotalAm = ""
			'Else
			'	ccTotalAm = CStr(CDbl(vRS("カード支払金額")))
			'End If

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

			'出荷合計数量 2020.06.02 GV add
			If (IsNull(vRS("出荷合計数量"))) Then
				shipSuuSum = 0
			Else
				shipSuuSum = CDbl(vRS("出荷合計数量"))
			End If


			'--- 明細行生成
			With oJSON.data("list")
				.Add j ,oJSON.Collection()
				With .item(j)
					.Add "c_no", CStr(Trim(vRS("顧客番号")))
					.Add "o_no" ,CStr(Trim(vRS("受注番号")))
					.Add "o_dt" ,orderDate '受注日
					.Add "est_dt" ,CStr(Trim(vRS("見積日")))
					.Add "ship_comp_dt" , shippingCompDate  '出荷完了日
					.Add "del_dt" ,deleteDate '削除日
					.Add "o_nm" ,orderName '注文者名前
					.Add "cst_nm" ,customerName '顧客名
					.Add "o_type" ,CStr(Trim(vRS("受注形態")))
					.Add "pay_method" ,get_paymetMethodWord(vRS("支払方法"))
					.Add "pay_method_detail" ,wPaymentMethodDetail '受注方法明細
					.Add "total_order_am", CDbl(vRS("受注合計金額")) 
					.Add "total_order_am2",  orderTotalAm2  ' 合計金額
					.Add "ff_charge_o", CDbl(vRS("受注時送料"))  '2020.06.01 GV add
					.Add "cod_charge_o", CDbl(vRS("受注時代引手数料")) '2020.06.01 GV add
					.Add "kabusoku_am_o", kabusokuAmAtOrder '2020.06.01 GV add
					.Add "used_pt_o", usedPointAtOrder  ' 受注時利用ポイント  '2020.06.01 GV add
					.Add "total_am_o", totalAmAtOrder '受注時合計金額 '2020.06.01 GV add
					.Add "used_pt", usedPoint  ' 利用ポイント
					.Add "cc_c_no", ccCreditNo 'カード与信確認番号(card_credit_no)
					.Add "cc_slip", ccSlipNo 'カードネット伝票番号(card_net_slip_no)
					.Add "o_zan", CStr(Trim(vRS("注残"))) 
					.Add "tax_rate", CDbl(vRS("消費税率"))
					.Add "ship_suu_sum", shipSuuSum '出荷合計数量
					.Add "deposit", depositFlag '入金完了フラグ 2016.06.03 GV add
					.Add "total_am_auth", totalAmAtAuth 'オーソリ時受注合計金額 2021.01.04 GV add
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
End If

' -------------------------------------------------
' JSONデータの返却
' -------------------------------------------------
' ヘッダ出力
Response.AddHeader "Content-Type", "application/json; charset=shift_jis"
Response.AddHeader "Cache-Control", "no-cache,must-revalidate"
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "X-Content-Type-Options", "nosniff"

' JSONデータの出力
Response.Write oJSON.JSONoutput()

End Function


'========================================================================
'
'	Function	日付けのフォーマット (YYYY年MM月DD日)
'
'========================================================================
Function formatDateYYYYMMDD(pdatDate)

Dim vDate

If IsNull(pdatDate) = True Then
	' Null は計算不能
	Exit Function
End If

If IsDate(pdatDate) = False Then
	' 日付けでなければ計算不能
	Exit Function
End If

vDate = DatePart("yyyy", pdatDate) & "年"

If DatePart("m", pdatDate) <= 9 Then
	vDate = vDate & "0" & DatePart("m", pdatDate)
Else
	vDate = vDate & DatePart("m", pdatDate)
End If

vDate = vDate & "月"

If DatePart("d", pdatDate) <= 9 Then
	vDate = vDate & "0" & DatePart("d", pdatDate)
Else
	vDate = vDate & DatePart("d", pdatDate)
End If

vDate = vDate & "日"

formatDateYYYYMMDD = vDate

End Function

'========================================================================
'
'	Function	表示用支払い方法文字の生成
'
'	Note
'	  支払方法              表示文字
'	──────────────────────
'	  コンビニ支払       → "コンビニ払い"
'	  ネットバンキング   → "コンビニ払い"
'	  ゆうちょ           → "コンビニ払い"
'	  ローン(頭金あり)   → "ローン"
'	  ローン(頭金なし)   → "ローン"
'	  ローン(頭金無し)   → "ローン"
'	  銀行振込           → "銀行振込"
'	  代引き             → "代金引換"
'	  現金               → (支払方法そのまま)
'	  売掛               → (支払方法そのまま)
'	  アマゾン           → (支払方法そのまま)
'	  クレジットカード   → (支払方法そのまま)
'
'========================================================================
Function get_paymetMethodWord(pstrPaymetMethod)

Dim vDisplayWord

If IsNull(pstrPaymetMethod) = True Then
	' Null は判定不能
	Exit Function
End If

If pstrPaymetMethod = "代引き" Then
	vDisplayWord = "代金引換"
ElseIf pstrPaymetMethod = "コンビニ支払" Then
	vDisplayWord = "コンビニ払い"
ElseIf pstrPaymetMethod = "ネットバンキング" Then
	vDisplayWord = "コンビニ払い"
ElseIf pstrPaymetMethod = "ゆうちょ" Then
	vDisplayWord = "コンビニ払い"
ElseIf pstrPaymetMethod = "銀行振込" Then
	vDisplayWord = "銀行振込"
ElseIf InStr(pstrPaymetMethod, "ローン") > 0 Then
	vDisplayWord = "ローン"
Else
	vDisplayWord = pstrPaymetMethod
End If

get_paymetMethodWord = vDisplayWord

End Function

'========================================================================
'
'	Function	Emaxのコントロールマスタからデータ取得
'
'========================================================================

Function getEmaxCntlMst(pSubSystemCd, pItemCd, pItemSubCd, pItemChar1, pItemChar2, pItemNum1, pItemNum2, pItemDate1, pItemDate2)

Dim RS_cntl
Dim v_sql

'---- コントロールマスタ取り出し

v_sql = ""
v_sql = v_sql & "SELECT a.*"
v_sql = v_sql & "  FROM コントロールマスタ a WITH (NOLOCK)"
v_sql = v_sql & " WHERE a.sub_system_cd = '" & pSubSystemCd & "'"
v_sql = v_sql & "   AND a.item_cd = '" & pItemCd & "'"
v_sql = v_sql & "   AND a.item_sub_cd = '" & pItemSubCd & "'"

'@@@@@@response.write(v_sql)

Set RS_cntl = Server.CreateObject("ADODB.Recordset")
RS_cntl.Open v_sql, ConnectionEmax, adOpenStatic

If RS_cntl.EOF <> True Then
	pItemChar1 = RS_cntl("item_char1")
	pItemChar2 = RS_cntl("item_char2")
	pItemNum1 = RS_cntl("item_num1")
	pItemNum2 = RS_cntl("item_num2")
	pItemDate1 = RS_cntl("item_date1")
	pItemDate2 = RS_cntl("item_date2")
End If

RS_cntl.Close

End Function
'========================================================================
%>
