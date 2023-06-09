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
'
'========================================================================
'On Error Resume Next

Const PAGE_SIZE = 20						' 購入履歴情報の1ページあたりの表示行数

Dim Connection
Dim ConnectionEmax

Dim wErrMsg						' エラーメッセージ (他のページから渡されるメッセージ)
Dim wDispMsg					' 通常メッセージ(エラー以外) (他のページから渡されるメッセージ)
Dim wErrDesc
Dim wMsg						' エラーメッセージ (本ページで作成するメッセージ)
Dim wUserID


Dim wIPage						' 表示するページ位置 (パラメータ)
Dim oJSON						' JSONオブジェクト
Dim listType					' 取得するリストタイプ（1...見積もり中/出荷準備中、2...購入履歴）


'=======================================================================
'	受け渡し情報取り出し & 初期設定
'=======================================================================
' Getパラメータ
wUserID = ReplaceInput(Trim(Request("cno")))
listType = ReplaceInput(Trim(Request("list_type")))
wIPage = ReplaceInput(Trim(Request("page")))	' ページ位置

'ページ番号
If wIPage = "" Or IsNumeric(wIPage) = False Then
	wIPage = 1
Else
	wIPage = CLng(wIPage)
End If


' 取得するリストタイプ
If (listType = "") Or (IsNumeric(listType) = False) Then
	listType = 1
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

Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

Set ConnectionEmax = Server.CreateObject("ADODB.Connection")
ConnectionEmax.Open g_connectionEmax

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
Function close_db()

Connection.close
Set Connection= Nothing

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
Dim i
Dim j
Dim vRS
Dim vParam
Dim vTitleWord
Dim vTitleWordSave
Dim vOrderDateLabel
Dim vHistoryCount
Dim vHTML
Dim orderDate
Dim shippingDate

Set oJSON = New aspJSON

' イテレータ初期化
i = 0
j = 0


'--- 該当顧客の受注一覧取り出し1 (見積中・出荷準備中)
If (listType = 1) Then
	vSQL = ""
	vSQL = vSQL & "SELECT "
	vSQL = vSQL & "      a.受注番号 "
	vSQL = vSQL & "    , a.受注日 "
	vSQL = vSQL & "    , a.見積日 "
	vSQL = vSQL & "    , a.出荷完了日 "
	vSQL = vSQL & "    , a.受注形態 "
	vSQL = vSQL & "    , a.支払方法 "
	vSQL = vSQL & "    , a.Web受注変更開始日 "
	vSQL = vSQL & "FROM "
	vSQL = vSQL & "    " & gLinkServer & "受注 a WITH (NOLOCK) "
	vSQL = vSQL & "WHERE "
	vSQL = vSQL & "        a.削除日     IS NULL "
	vSQL = vSQL & "    AND a.出荷完了日 IS NULL "
	'vSQL = vSQL & "    AND a.受注形態 in ('E-mail','FAX','インターネット','携帯','電話','郵送','来店')"	'2012/11/24 ok Del
	vSQL = vSQL & "    AND a.受注形態 in ('E-mail','FAX','インターネット','携帯','電話','郵送','来店','スマートフォン')"	'2012/11/24 ok Add
	vSQL = vSQL & "    AND a.顧客番号   = " & wUserID & " "
	vSQL = vSQL & "ORDER BY "
	vSQL = vSQL & "      CASE WHEN a.受注日 IS NULL "
	vSQL = vSQL & "          THEN 1 "
	vSQL = vSQL & "          ELSE 2 "
	vSQL = vSQL & "      END "
	vSQL = vSQL & "    , 見積日 DESC "

	'@@@@Response.Write(vSQL)

	Set vRS = Server.CreateObject("ADODB.Recordset")
	vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	If vRS.EOF = False Then

		' リスト追加
		oJSON.data.Add "list" ,oJSON.Collection()

		Do Until vRS.EOF = True
			'--- 出荷状況(タイトル) の判定
			vTitleWord = make_titleWord(vRS("受注日"), vRS("出荷完了日"))

			'--- 注文日列のタイトルラベル決定
			If vTitleWord = "お見積" Then
				vOrderDateLabel = "お見積日"
			ElseIf vTitleWord = "出荷準備中" Then
				vOrderDateLabel = "ご注文日"
			ElseIf vTitleWord = "ご購入履歴" Then
				vOrderDateLabel = "ご注文日"
			Else
				vOrderDateLabel = "ご注文日"
			End If

			' 受注日
			If (IsNull(vRS("受注日"))) Then
				orderDate = ""
			Else
				orderDate = CStr(Trim(vRS("受注日")))
			End If


			' 出荷完了日
			If (IsNull(vRS("出荷完了日"))) Then
				shippingDate = ""
			Else
				shippingDate = CStr(Trim(vRS("出荷完了日")))
			End If


			With oJSON.data("list")
				.Add j ,oJSON.Collection()
				With .item(j)
'					.Add "title" ,vTitleWord
'					.Add "list" ,oJSON.Collection()
'					With .item("list")
						.Add "order_date" ,orderDate
						.Add "estimate_date" ,CStr(Trim(vRS("見積日")))
						.Add "order_no" ,CStr(Trim(vRS("受注番号")))
						.Add "order_type" ,CStr(Trim(vRS("受注形態")))
						.Add "payment_method" ,get_paymetMethodWord(vRS("支払方法"))
						.Add "shipping_date" , shippingDate
'					End With
				End With
			End With

			vRS.MoveNext
			j = j + 1
		Loop
	End If

	'レコードセットを閉じる
	vRS.Close
End If

If (listType = 2) Then
'--- 該当顧客の受注一覧取り出し2 (ご購入履歴)
	vSQL = ""
	vSQL = vSQL & "SELECT "
	vSQL = vSQL & "      a.受注番号 "
	vSQL = vSQL & "    , a.受注日 "
	vSQL = vSQL & "    , a.見積日 "
	vSQL = vSQL & "    , a.出荷完了日 "
	vSQL = vSQL & "    , a.受注形態 "
	vSQL = vSQL & "    , a.支払方法 "
	vSQL = vSQL & "    , a.Web受注変更開始日 "
	vSQL = vSQL & "FROM "
	vSQL = vSQL & "    " & gLinkServer & "受注 a WITH (NOLOCK) "
	vSQL = vSQL & "WHERE "
	vSQL = vSQL & "        a.削除日     IS NULL "
	vSQL = vSQL & "    AND a.出荷完了日 IS NOT NULL "
	'vSQL = vSQL & "    AND a.受注形態 in ('E-mail','FAX','インターネット','携帯','電話','郵送','来店')"	'2012/11/24 ok Del
	vSQL = vSQL & "    AND a.受注形態 in ('E-mail','FAX','インターネット','携帯','電話','郵送','来店','スマートフォン')"	'2012/11/24 ok Add
	vSQL = vSQL & "    AND a.顧客番号   = " & wUserID & " "
	vSQL = vSQL & "ORDER BY "
	vSQL = vSQL & "    見積日 DESC "

	'@@@@Response.Write(vSQL)

	Set vRS = Server.CreateObject("ADODB.Recordset")
	vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	If vRS.EOF = False Then

		' 全件数をJSONデータにセット
		oJSON.data.Add "count" ,vRS.RecordCount

		' リスト追加
		oJSON.data.Add "list" ,oJSON.Collection()

		'--- 出荷状況(タイトル) 文字生成
		vTitleWord = "ご購入履歴"
		vOrderDateLabel = "ご注文日"

		'--- 指定ページを表示する為のレコード位置付け(SearchListの処理に倣う)
		vRS.PageSize = PAGE_SIZE
		If wIPage > ((vRS.RecordCount + (PAGE_SIZE - 1)) / PAGE_SIZE) Then		'MAXページを超える場合は最終ページへ
			wIPage = Fix(vRS.RecordCount / PAGE_SIZE)
		End If

		' レコード位置の位置付け
		vRS.AbsolutePage = wIPage

		For i = 0 To (vRS.PageSize - 1)

			' 受注日
			orderDate = vRS("受注日")
			If (IsNull(vRS("受注日"))) Then
				orderDate = ""
			Else
				orderDate = CStr(Trim(vRS("受注日")))
			End If

			' 出荷完了日
			If (IsNull(vRS("出荷完了日"))) Then
				shippingDate = ""
			Else
				shippingDate = CStr(Trim(vRS("出荷完了日")))
			End If


			'--- 明細行生成
			With oJSON.data("list")
				.Add j ,oJSON.Collection()
				With .item(j)
'					.Add "title" ,vTitleWord
'					.Add "list" ,oJSON.Collection()
'					With .item("list")
						.Add "order_date" ,orderDate
						.Add "estimate_date" ,CStr(Trim(vRS("見積日")))
						.Add "order_no" ,CStr(Trim(vRS("受注番号")))
						.Add "order_type" ,CStr(Trim(vRS("受注形態")))
						.Add "payment_method" ,get_paymetMethodWord(vRS("支払方法"))
						.Add "shipping_date" , shippingDate
'					End With
				End With
			End With


			vRS.MoveNext

			If vRS.EOF Then
				Exit For
			End If

			j = j + 1
		Next
	End If

	'レコードセットを閉じる
	vRS.Close
End If

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
'	Function	購入履歴のタイトル文字生成
'
'========================================================================
Function make_titleWord(pdatOrderDate, pdatShipCompleteDate)

Dim vTitleWord

If IsNull(pdatOrderDate) Then
	'--- 受注日がNullの場合
	vTitleWord = "お見積"
ElseIf IsNull(pdatOrderDate) = False And IsNull(pdatShipCompleteDate) Then
	'--- 受注日がNullでなく、出荷完了日がNullの場合
	vTitleWord = "出荷準備中"
ElseIf IsNull(pdatShipCompleteDate) = False Then
	'--- 出荷完了日がNullの場合
	vTitleWord = "ご購入履歴"
Else
	vTitleWord = "ご購入履歴"
End If

make_TitleWord = vTitleWord

End Function

'========================================================================
%>
