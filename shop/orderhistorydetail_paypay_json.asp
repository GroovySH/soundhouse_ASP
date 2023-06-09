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
'	PayPay支払一覧ページ
'
'
'変更履歴
'2020.02.27 GV 新規作成。(PayPay対応)(#2405)
'2020.06.01 GV 修正。
'
'========================================================================
'On Error Resume Next

Const PAGE_SIZE = 10			' 購入履歴情報の1ページあたりの表示行数

Dim ConnectionEmax

Dim wErrMsg						' エラーメッセージ (他のページから渡されるメッセージ)
Dim wDispMsg					' 通常メッセージ(エラー以外) (他のページから渡されるメッセージ)
Dim wErrDesc
Dim wMsg						' エラーメッセージ (本ページで作成するメッセージ)
Dim wCustomerNo					' 顧客番号
Dim wOrderNo					' 受注番号
Dim wFlg						' 実行フラグ
Dim wIPage						' 表示するページ位置 (パラメータ)
Dim wYear						' 検索機関
Dim oJSON						' JSONオブジェクト
Dim wOrderHidden				' 非表示フラグ
Dim wOrderCancelled				' キャンセル注文フラグ
Dim wOrderShipping				' 未発送注文フラグ
Dim wSlipNo						' 送り状番号
Dim wReceipt					' 領収書
Dim wDepositTerm				' 入金確認期限（日）

'=======================================================================
'	受け渡し情報取り出し & 初期設定
'=======================================================================
wFlg = True

' Getパラメータ
' 顧客番号
wCustomerNo = ReplaceInput_NoCRLF(Trim(Request("cno")))
' 数値のみチェック (ASPは全角でも数字ならTrueを返す)
If (IsNumeric(wCustomerNo) = False) Or (cf_checkNumeric(wCustomerNo) = False) Then
	wFlg = False
End If


'ページ番号
wIPage = ReplaceInput_NoCRLF(Trim(Request("page")))
' 数値のみチェック (ASPは全角でも数字ならTrueを返す)
If (IsNumeric(wIPage) = False) Or (cf_checkNumeric(wIPage) = False) Then
	wIPage = 1
Else
	wIPage = CLng(wIPage)
End If

'検索期間
wYear = ReplaceInput_NoCRLF(Trim(Request("year")))
' 数値のみチェック (ASPは全角でも数字ならTrueを返す)
If (IsNumeric(wYear) = False) Or (cf_checkNumeric(wYear) = False) Then
	wYear = null
Else
	wYear = CLng(wYear)
End If

'受注番号
wOrderNo = ReplaceInput_NoCRLF(Trim(Request("ono")))
' 数値のみチェック (ASPは全角でも数字ならTrueを返す)
If (IsNumeric(wOrderNo) = False) Or (cf_checkNumeric(wOrderNo) = False) Then
	wOrderNo = null
Else
	wOrderNo = CLng(wOrderNo)
End If

'非表示フラグ
wOrderHidden = ReplaceInput_NoCRLF(Trim(Request("hide")))
If ((IsNull(wOrderHidden) = True) Or (UCase(wOrderHidden) <> "Y")) Then
	wOrderHidden = "N"
Else
	wOrderHidden = "Y"
End If

'キャンセル注文フラグ
wOrderCancelled = ReplaceInput_NoCRLF(Trim(Request("cancelled")))
If ((IsNull(wOrderCancelled) = True) Or (UCase(wOrderCancelled) <> "Y")) Then
	wOrderCancelled = "N"
Else
	wOrderCancelled = "Y"
End If

'未発送注文フラグ
wOrderShipping = ReplaceInput_NoCRLF(Trim(Request("shipping")))
If ((IsNull(wOrderShipping) = True) Or (UCase(wOrderShipping) <> "Y")) Then
	wOrderShipping = "N"
Else
	wOrderShipping = "Y"
End If

'送り状番号
wSlipNo = ReplaceInput_NoCRLF(Trim(Request("slip")))
' 数値のみチェック (ASPは全角でも数字ならTrueを返す)
If (IsNumeric(wSlipNo) = False) Or (cf_checkNumeric(wSlipNo) = False) Then
	wSlipNo = null
End If


'領収書用
wReceipt = ReplaceInput_NoCRLF(Trim(Request("receipt")))
If ((IsNull(wReceipt) = True) Or (UCase(wReceipt) <> "Y")) Then
	wReceipt = "N"
Else
	wReceipt = "Y"
End If

If (wReceipt = "Y") Then
	'非表示フラグを無効
	wOrderHidden = "N"
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
Dim vWHERE
Dim i
Dim j
Dim vRS
Dim orderDate
Dim deleteDate
Dim shippingCompDate
Dim shippingDate
Dim shippingSuu
Dim itemPicSmall
Dim makerName
Dim makerChokusou
Dim itemName
Dim iro
Dim kikaku
Dim allCount
Dim orderTotalAm2
Dim usedPoint
Dim shipNo
Dim slipNo
Dim modifyFlag  '変更可能フラグ
Dim cancelFlag  'キャンセル可能フラグ
Dim modifyNg    '変更NG理由
Dim cancelNg    'キャンセルNG理由
Dim dateTerm
Dim maxDate
Dim ngReason
Dim ffCd
Dim orderType
Dim modifiable
Dim setItemFlag
Dim promote
Dim estMemo
Dim buy
Dim webModCancelFlg
Dim webOutline
Dim btnOn 'ボタン表示フラグ
Dim depositFlag ' 入金完了フラグ 2016.06.03 GV add
Dim receiptFlag '領収書発行フラグ 2020.02.05 GV add
Dim receiptDate '領収書発行日 2020.02.05 GV add
' 2018.01.12 GV add start
Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2
' 2018.01.12 GV end
Dim isOtherAmountOk '2018.12.03 GV add

' 2020/.06.01 GV add start
Dim totalAmAtOrder ' 受注時合計金額
Dim usedPointAtOrder


Dim wPaymentMethodDetail '2018.12.21 GV add

Set oJSON = New aspJSON

' 初期化
i = 0
j = 0
allCount = 0
modifyFlag = "Y"
cancelFlag = "Y"
modifiable = "Y"
maxDate = ""
ngReason = ""
ffCd = ""
promote = ""
isOtherAmountOk = True '2018.12.03 GV add

' 受注形態(カンマ区切りで指定)
orderType = ""
orderType = orderType & "  'E-mail'"
orderType = orderType & " ,'FAX'"
orderType = orderType & " ,'インターネット'"
orderType = orderType & " ,'携帯'"
orderType = orderType & " ,'電話'"
orderType = orderType & " ,'郵送'"
orderType = orderType & " ,'来店'"
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
	'--- 該当顧客の受注一覧取り出し
	vSQL = ""
	vSQL = vSQL & "SELECT "
	vSQL = vSQL & "  o.顧客番号 "
	vSQL = vSQL & " ,o.受注番号 "
	vSQL = vSQL & " ,o.受注日 "
	vSQL = vSQL & " ,o.見積日 "
	vSQL = vSQL & " ,o.削除日 "
	vSQL = vSQL & " ,o.出荷完了日 "
	vSQL = vSQL & " ,o.受注形態 "
	vSQL = vSQL & " ,o.支払方法 "
	vSQL = vSQL & " ,o.消費税率 "
	vSQL = vSQL & " ,o.受注合計金額 "
	vSQL = vSQL & " ,o.合計金額 "
	vSQL = vSQL & " ,o.受注時合計金額 " ' 2020.06.01 GV add
	vSQL = vSQL & " ,o.受注時利用ポイント " '2020.06.01 GV add
	vSQL = vSQL & " ,o.利用ポイント "
	vSQL = vSQL & " ,o.Web注文変更キャンセル中フラグ "
	vSQL = vSQL & " ,o.その他合計金額 "
	vSQL = vSQL & ",(SELECT "
	vSQL = vSQL & "    count(*) FROM 受注その他明細 other1 WITH(NOLOCK) "
	vSQL = vSQL & "  WHERE "
	vSQL = vSQL & "        other1.受注番号 = o.受注番号 " 
	vSQL = vSQL & "    AND 受注その他コード <> 'COUPON' "
	vSQL = vSQL & ") as その他明細件数 "
	vSQL = vSQL & " ,ISNULL(o.配送情報明細指定フラグ, 'N') AS  配送情報明細指定フラグ "
	vSQL = vSQL & " ,o.入金完了フラグ "
	vSQL = vSQL & " ,o.領収書番号 "
	vSQL = vSQL & " ,o.領収書発行日 "
	vSQL = vSQL & " ,o.支払方法詳細 "
	vSQL = vSQL & " , od.今回限り届先名前 "
	vSQL = vSQL & " , od.受注明細番号 "
	vSQL = vSQL & " , od.メーカーコード "
	vSQL = vSQL & " , m.メーカー名 "
	vSQL = vSQL & " , od.商品コード "
	vSQL = vSQL & " , i.商品名 "
	vSQL = vSQL & " , od.色 "
	vSQL = vSQL & " , od.規格 "
	vSQL = vSQL & " , iz.商品ID "
	vSQL = vSQL & " , i.商品画像ファイル名_小 "
	vSQL = vSQL & " , i.Web商品フラグ "
	vSQL = vSQL & " , i.セット商品フラグ "
	vSQL = vSQL & " ,CASE "
	vSQL = vSQL & "    WHEN i.B品フラグ = 'Y' THEN iz.B品引当可能数量 "
	vSQL = vSQL & "    ELSE iz.引当可能数量 "
	vSQL = vSQL & "	   END AS 在庫数 "
	vSQL = vSQL & " ,i.取扱中止日 "
	vSQL = vSQL & " ,i.終了日 "

	vSQL = vSQL & " , od.メーカー直送フラグ "
	vSQL = vSQL & " , od.受注単価 "
	vSQL = vSQL & " , od.受注金額 "
	vSQL = vSQL & " , od.受注数量 "
	vSQL = vSQL & " , od.受注時数量 " '2020.06.01 GV add
	vSQL = vSQL & " , od.出荷指示合計数量 "
	vSQL = vSQL & " , (CASE "
	vSQL = vSQL & "     WHEN o.受注日 IS NOT NULL AND o.削除日 IS NULL THEN 'Y' "
	vSQL = vSQL & "     ELSE 'N' END "
	vSQL = vSQL & "   ) AS 注残 "
	'vSQL = vSQL & " , iz.適正在庫数量 "
	vSQL = vSQL & " , ISNULL(od.適正在庫数量, 0) AS 適正在庫数量 "   '注文時の適正在庫数量
'	vSQL = vSQL & " ,od.送り状番号 "
'	vSQL = vSQL & " ,od.出荷番号 "
'	vSQL = vSQL & " ,od.出荷数量 "
'	vSQL = vSQL & " ,od.出荷日 "
'	vSQL = vSQL & " ,od.運送会社コード "
	vSQL = vSQL & ", od.受注明細備考 "


	vSQL = vSQL & " FROM "
	vSQL = vSQL & " 受注 AS o WITH (NOLOCK) "

	vSQL = vSQL & " INNER JOIN 受注明細 od WITH (NOLOCK) "
	vSQL = vSQL & "    ON od.受注番号 = o.受注番号 "
	vSQL = vSQL & "   AND od.セット品親明細番号 = 0 "

	vSQL = vSQL & "INNER JOIN 色規格別在庫 iz WITH (NOLOCK) "
	vSQL = vSQL & "   ON iz.メーカーコード = od.メーカーコード "
	vSQL = vSQL & "  AND iz.商品コード = od.商品コード "
	vSQL = vSQL & "  AND iz.色 = od.色 "
	vSQL = vSQL & "  AND iz.規格 = od.規格 "

	vSQL = vSQL & "INNER JOIN 商品 i WITH (NOLOCK) "
	vSQL = vSQL & "   ON i.メーカーコード = iz.メーカーコード "
	vSQL = vSQL & "  AND i.商品コード = iz.商品コード "

	vSQL = vSQL & "INNER JOIN メーカー m WITH (NOLOCK) "
	vSQL = vSQL & "   ON m.メーカーコード = i.メーカーコード "

	vSQL = vSQL & " WHERE "
	vSQL = vSQL & "       o.顧客番号 =  " & wCustomerNo
	vSQL = vSQL & "   AND o.受注番号 = " & wOrderNo

	vSQL = vSQL & " ORDER BY "
	vSQL = vSQL & "   受注明細番号 asc"


	'@@@@Response.Write(vSQL) & "<br>"


	Set vRS = Server.CreateObject("ADODB.Recordset")
	vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	'レコードが存在している場合
	If vRS.EOF = False Then

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

			If (IsNull(vRS("商品画像ファイル名_小"))) Then
				itemPicSmall = ""
			Else
				itemPicSmall = CStr(vRS("商品画像ファイル名_小"))
			End If


			If (IsNull(vRS("セット商品フラグ"))) Then
				setItemFlag = ""
			Else
				setItemFlag = CStr(vRS("セット商品フラグ"))
			End If

			makerName = Replace(Trim(vRS("メーカー名")), """", "”")
			makerName = CStr(makerName)

			itemName = Replace(Trim(vRS("商品名")), """", "”")
			itemName = CStr(itemName)

			iro = Replace(Trim(vRS("色")), """", "”")
			iro = CStr(iro)

			kikaku = Replace(Trim(vRS("規格")), """", "”")
			kikaku = CStr(kikaku)

			If (IsNull(vRS("メーカー直送フラグ"))) Then
				makerChokusou = ""
			Else
				makerChokusou = CStr(vRS("メーカー直送フラグ"))
			End If

			If (IsNull(vRS("合計金額"))) Then
				orderTotalAm2 = 0
			Else
				orderTotalAm2 = CDbl(vRS("合計金額"))
			End If

			If (IsNull(vRS("受注時合計金額"))) Then
				totalAmAtOrder = 0
			Else
				totalAmAtOrder = CDbl(vRS("受注時合計金額"))
			End If

			If (IsNull(vRS("受注時利用ポイント"))) Then
				usedPointAtOrder = 0
			Else
				usedPointAtOrder = CDbl(vRS("受注時利用ポイント"))
			End If

			' 利用ポイント
			If (IsNull(vRS("利用ポイント"))) Then
				usedPoint = 0
			Else
				usedPoint = CDbl(vRS("利用ポイント"))
			End If

			'2016.06.03 GV add start
			'入金完了フラグ
			If (IsNull(vRS("入金完了フラグ"))) Then
				depositFlag = ""
			Else
				depositFlag = CStr(Trim(vRS("入金完了フラグ")))
			End If
			'2016.06.03 GV add start

			'2020.02.05 GV add start
			'領収書発行フラグ
			receiptFlag = getReceiptFlag(vRS("支払方法"), CStr(Trim(vRS("受注番号"))))

			'領収書発行日
			If (IsNull(vRS("領収書発行日"))) Then
				receiptDate = ""
			Else
				receiptDate = CStr(Trim(vRS("領収書発行日")))
			End If
			'2020.02.05 GV add end

			'販促品判定
			promote = "N"
			If (CDbl(Trim(vRS("受注単価"))) = 0) Then
				'受注明細備考に「販促品」と含まれる場合、
				'estMemo = InStr(Trim(vRS("受注明細備考")), "販促品")
				'If (IsNull(estMemo) = False) And (IsNumeric(estMemo)) And (estMemo > 0) Then
				If (InStr(Trim(vRS("受注明細備考")), "販促品") > 0) Then
					promote = "Y"
				ElseIf (InStr(Trim(vRS("商品コード")), "HOTMENU") > 0) Then
					promote = "Y"
				End If
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

			' 2018.12.21 GV add start
			'支払い方法詳細
			If (IsNull(vRS("支払方法詳細"))) Then
				wPaymentMethodDetail = ""
			Else
				wPaymentMethodDetail = CStr(vRS("支払方法詳細"))
			End If
			' 2018.12.21 GV add end

			' ---------------------------------------------------


			'--- 明細行生成
			With oJSON.data("list")
				.Add j ,oJSON.Collection()
				With .item(j)
					.Add "o_no" ,CStr(Trim(vRS("受注番号")))
					.Add "o_dt" ,orderDate '受注日
					.Add "est_dt" ,CStr(Trim(vRS("見積日")))
					.Add "ship_comp_dt" , shippingCompDate  '出荷完了日
					.Add "del_dt" ,deleteDate '削除日
					.Add "o_type" ,CStr(Trim(vRS("受注形態")))
					.Add "pay_method" ,get_paymetMethodWord(vRS("支払方法"))
					.Add "pay_method_detail" ,wPaymentMethodDetail '2018.12.21 GV add
					.Add "tax_rate", CDbl(vRS("消費税率")) 
					.Add "total_order_am", CDbl(vRS("受注合計金額")) 
					.Add "total_order_am2",  orderTotalAm2  ' 合計金額
					.Add "total_am_o",  totalAmAtOrder  ' 受注時合計金額 2020.06.01 GV add
					.Add "used_pt_o", usedPointAtOrder  ' 受注時利用ポイント 2020.06.01 GV add
					.Add "used_pt", usedPoint  ' 利用ポイント
					.Add "ship_name" ,CStr(Trim(vRS("今回限り届先名前")))
					.Add "web_flg", CStr(vRS("Web商品フラグ"))
					.Add "set_flg", setItemFlag
					.Add "od_no" ,CStr(Trim(vRS("受注明細番号")))
					.Add "m_cd" ,CStr(Trim(vRS("メーカーコード")))
					.Add "m_name" ,makerName
					.Add "i_cd" ,CStr(Trim(vRS("商品コード")))
					.Add "i_name" ,itemName
					.Add "iro" ,iro
					.Add "kikaku" ,kikaku
					.Add "i_id" ,CStr(Trim(vRS("商品ID")))
					.Add "i_pic", itemPicSmall
					.Add "m_chokusou", makerChokusou
					.Add "i_tanka", CDbl(Trim(vRS("受注単価")))
					.Add "i_am", CDbl(Trim(vRS("受注金額")))
					.Add "i_suu", CDbl(vRS("受注数量")) 
					.Add "i_suu_o", CDbl(vRS("受注時数量")) 
					.Add "ship_inst_suu", CDbl(vRS("出荷指示合計数量"))
					.Add "o_zan", CStr(Trim(vRS("注残"))) 
					.Add "t_zaiko_suu", CDbl(vRS("適正在庫数量"))
					.Add "promote" , promote '販促品判定
					.Add "modifying", webModCancelFlg
					.Add "deposit", depositFlag '入金完了フラグ 2016.06.03 GV add
				End With
			End With

			' 次のレコード行へ移動
			vRS.MoveNext

			If vRS.EOF Then
				Exit For
			End If

			j = j + 1
		Next

		'受注番号指定の場合
		'If (wOrderNo <> "") Then
		'	' 変更可能かセット
		'	oJSON.data.Add "modifiable" ,modifiable
		'End If
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
