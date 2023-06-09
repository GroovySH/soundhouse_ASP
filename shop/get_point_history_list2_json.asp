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
'	購入履歴一覧ページにおけるポイント情報を取得
'
'
'変更履歴
'2016.02.10 GV 新規作成
'
'========================================================================
'On Error Resume Next

'Const PAGE_SIZE = 20			' 購入履歴情報の1ページあたりの表示行数

Dim Connection
Dim ConnectionEmax

Dim wErrMsg						' エラーメッセージ (他のページから渡されるメッセージ)
Dim wDispMsg					' 通常メッセージ(エラー以外) (他のページから渡されるメッセージ)
Dim wErrDesc
Dim wMsg						' エラーメッセージ (本ページで作成するメッセージ)
Dim wCustomerNo					' 顧客番号
Dim wOrderNo					' 受注番号
Dim oJSON						' JSONオブジェクト
Dim wPage						' 表示するページ位置 (パラメータ)
Dim wPageSize					'購入履歴情報の1ページあたりの表示行数

'=======================================================================
'	受け渡し情報取り出し & 初期設定
'=======================================================================
' Getパラメータ
wCustomerNo = ReplaceInput(Trim(Request("cno")))
wOrderNo = ReplaceInput(Trim(Request("ono")))
wPage = ReplaceInput(Trim(Request("page")))	' ページ位置
wPageSize = ReplaceInput(Trim(Request("page_size")))

'ページ番号
If wPage = "" Or IsNumeric(wPage) = False Then
	wPage = 1
Else
	wPage = CLng(wPage)
End If

'ページサイズ
If wPageSize = "" Or IsNumeric(wPageSize) = False Then
	wPageSize = 10
Else
	wPageSize = CLng(wPageSize)
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
'Dim vPaymentMethod
Dim vKubun
Dim vOrderDate
Dim vPointDate
Dim vPoint
'Dim vOrderType
Dim vOrderNo
Dim vPointZan
Dim vAllPage
Dim vOffset
Dim vAdjust
Dim vPointExpire

Set oJSON = New aspJSON

' 初期化
i         = 0
j         = 0
vPointZan = 0
vAllPage  = 0
vOffset   = 0
vAdjust   = 1

'--- 該当顧客のポイント明細の取り出し
vSQL = ""
' OLD
'vSQL = vSQL & "SELECT "
'vSQL = vSQL & "  T1.ポイント日付 AS ご注文日, "
'vSQL = vSQL & "  T1.受注番号 AS ご注文番号, "
'vSQL = vSQL & "  MAX(T2.受注形態) AS ご注文方法, "
'vSQL = vSQL & "  MAX(T2.支払方法) AS お支払方法, "
'vSQL = vSQL & "  T1.ポイント区分 AS ポイント利用獲得, "
'vSQL = vSQL & "  T1.ポイント日付 AS ポイント獲得日, "
'vSQL = vSQL & "  SUM(T1.ポイント) AS ポイント "
'vSQL = vSQL & "FROM "
'vSQL = vSQL & "  ポイント明細 T1 WITH (NOLOCK) "
'vSQL = vSQL & "  LEFT JOIN 受注 T2 WITH (NOLOCK) "
'vSQL = vSQL & "    ON (T2.受注番号 = T1.受注番号 "
'vSQL = vSQL & "    AND T2.顧客番号 = T1.顧客番号) "
'vSQL = vSQL & "WHERE "
'vSQL = vSQL & "  T1.顧客番号=" & wCustomerNo
'vSQL = vSQL & " AND T1.ポイント日付 IS NOT NULL "
'vSQL = vSQL & "GROUP BY "
'vSQL = vSQL & "  T1.ポイント区分, "
'vSQL = vSQL & "  T1.ポイント日付, "
'vSQL = vSQL & "  T1.受注番号 "
'vSQL = vSQL & "HAVING "
'vSQL = vSQL & "  SUM(T1.ポイント) <> '0' "
'vSQL = vSQL & "ORDER BY "
'vSQL = vSQL & "  ご注文日, "
'vSQL = vSQL & "  ご注文番号, "
'vSQL = vSQL & "  ポイント獲得日, "
'vSQL = vSQL & "  ポイント利用獲得  DESC"

'ここから
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "  CONVERT(NVARCHAR, T1.ポイント日付, 111) AS ご注文日, "
vSQL = vSQL & "  T1.受注番号 AS ご注文番号, "
'vSQL = vSQL & "  MAX(T2.受注形態) AS ご注文方法, "
'vSQL = vSQL & "  MAX(T2.支払方法) AS お支払方法, "
vSQL = vSQL & "  T1.ポイント区分 AS ポイント利用獲得, "
vSQL = vSQL & "  T1.ポイント日付 AS ポイント獲得日, "
vSQL = vSQL & "  T1.ポイント期限, "
vSQL = vSQL & "  SUM(T1.ポイント) AS ポイント, "
vSQL = vSQL & "  CASE T1.ポイント区分 "
vSQL = vSQL & "    WHEN '利用' THEN 0 "
vSQL = vSQL & "    ELSE 1 "
vSQL = vSQL & "  END AS POINT_SORT "
vSQL = vSQL & "FROM "
vSQL = vSQL & "  ポイント明細 T1 WITH (NOLOCK) "
'vSQL = vSQL & "  LEFT JOIN 受注 T2 WITH (NOLOCK) "
'vSQL = vSQL & "    ON (T2.受注番号 = T1.受注番号 "
'vSQL = vSQL & "    AND T2.顧客番号 = T1.顧客番号) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "  T1.顧客番号= " & wCustomerNo
vSQL = vSQL & " AND T1.ポイント日付 IS NOT NULL "
vSQL = vSQL & "GROUP BY "
vSQL = vSQL & "  T1.ポイント区分, "
vSQL = vSQL & "  T1.ポイント日付, "
vSQL = vSQL & "  T1.受注番号 "
vSQL = vSQL & "  ,T1.ポイント期限 "
vSQL = vSQL & "HAVING "
vSQL = vSQL & "  SUM(T1.ポイント) <> '0' "
vSQL = vSQL & "ORDER BY "
vSQL = vSQL & "  ご注文日, "
vSQL = vSQL & "  ご注文番号, "
vSQL = vSQL & "  POINT_SORT ASC"


'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic
If vRS.EOF = False Then

	' 全件数をJSONデータにセット
	oJSON.data.Add "count" ,vRS.RecordCount

	' リスト追加
	oJSON.data.Add "list" ,oJSON.Collection()

	'--- 指定ページを表示する為のレコード位置付け(SearchListの処理に倣う)
	'vRS.PageSize = PAGE_SIZE
	'If wPage > ((vRS.RecordCount + (PAGE_SIZE - 1)) / PAGE_SIZE) Then		'MAXページを超える場合は最終ページへ
	'	wPage = Fix(vRS.RecordCount / PAGE_SIZE)
	'vRS.PageSize = wPageSize
'	If wPage > ((vRS.RecordCount + (wPageSize - 1)) / wPageSize) Then		'MAXページを超える場合は最終ページへ
'		wPage = Round((vRS.RecordCount / wPageSize) + 0.5)
'	End If
	' レコード数からのページ数
	vAllPage = Round((vRS.RecordCount / wPageSize) + 0.5)

	' レコード位置の位置付け
	'vRS.AbsolutePage = wPage

	vOffset = vRS.RecordCount - (wPage * wPageSize)

	'最後のページの場合
	If wPage = vAllPage Then
		vOffset = 0
		vAdjust = 2
	Else
		vAdjust = 1
	End If

	'For i = 0 To (vRS.PageSize - 1)
	For i = 0 To (vRS.RecordCount - 1)

		' --------------------------------
		' ポイント
		If (IsNull(vRS("ポイント"))) Then
			vPoint = 0
		Else
			vPoint = CStr(Trim(vRS("ポイント")))
		End If

		'ポイント残を累積
		vPointZan = vPointZan + CLng(vPoint)

		' --------------------------------
		'必要なレコード位置の場合に、JSONデータを生成する
		'If (wPage <= vAllPage) And (i >= (wPageSize * (wPage - 1))) And (i <= (wPageSize * wPage - 1)) Then
		If (wPage <= vAllPage) And (i >= (vOffset)) And (i <= (vOffset + wPageSize - vAdjust)) Then
			'ご注文日(ポイント日付)
			If (IsNull(vRS("ご注文日"))) Then
				vOrderDate = ""
			Else
				vOrderDate = CStr(Trim(vRS("ご注文日")))
			End If

			'ご注文番号(受注番号)
			If (IsNull(vRS("ご注文番号"))) Then
				vOrderNo = ""
			Else
				vOrderNo = CStr(Trim(vRS("ご注文番号")))
			End If

			'ご注文方法(MAX(T2.受注形態))
			'If (IsNull(vRS("ご注文方法"))) Then
			'	vOrderType = ""
			'Else
			'	vOrderType = CStr(Trim(vRS("ご注文方法")))
			'End If

			'お支払方法(MAX(T2.支払方法))
			'If (IsNull(vRS("お支払方法"))) Then
			'	vPaymentMethod = ""
			'Else
			'	vPaymentMethod = CStr(Trim(vRS("お支払方法")))
			'End If

			'ポイント利用獲得(ポイント区分)
			If (IsNull(vRS("ポイント利用獲得"))) Then
				vKubun = ""
			Else
				vKubun = CStr(Trim(vRS("ポイント利用獲得")))
			End If

			'ポイント獲得日(ポイント日付)
			If (IsNull(vRS("ポイント獲得日"))) Then
				vPointDate = ""
			Else
				vPointDate = CStr(Trim(vRS("ポイント獲得日")))
			End If

			'ポイント期限
			If (IsNull(vRS("ポイント期限"))) Then
				vPointExpire = ""
			Else
				vPointExpire = CStr(Trim(vRS("ポイント期限")))
			End If


			' リスト追加
			With oJSON.data("list")
				.Add j, oJSON.Collection()
				With .item(j)
					.Add "order_date" ,formatDateYYYYMMDD(vOrderDate)
					.Add "order_no" ,vOrderNo
					'.Add "order_type" ,vOrderType
					'.Add "payment_method" ,vPaymentMethod
					.Add "kubun" ,vKubun
					.Add "point_date" ,formatDateYYYYMMDD(vPointDate)
					.Add "point_expire" ,formatDateYYYYMMDD(vPointExpire)
					.Add "point" ,vPoint
					.Add "point_zan" ,vPointZan
				End With
			End With
		End If

		' イテレータをインクリメント
		j = j + 1

		vRS.MoveNext

		If vRS.EOF Then
			Exit For
		End If
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
Response.AddHeader "Content-Type", "application/json; charset=shift_jis"
Response.AddHeader "Cache-Control", "no-cache,must-revalidate"
Response.AddHeader "Pragma", "no-cache"

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

vDate = DatePart("yyyy", pdatDate) & "/"

If DatePart("m", pdatDate) <= 9 Then
	vDate = vDate & "0" & DatePart("m", pdatDate)
Else
	vDate = vDate & DatePart("m", pdatDate)
End If

vDate = vDate & "/"

If DatePart("d", pdatDate) <= 9 Then
	vDate = vDate & "0" & DatePart("d", pdatDate)
Else
	vDate = vDate & DatePart("d", pdatDate)
End If

vDate = vDate & ""

formatDateYYYYMMDD = vDate

End Function

'========================================================================
%>
