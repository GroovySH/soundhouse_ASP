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
'2021.08.04 GV 新規作成。(get_point_history_list_v3_jsonを踏襲)(#2859)
'2021.08.27 GV 見積もり状態の注文は除外するよう改修。(#2909)
'2021.09.03 GV 獲得予定ポイントのクエリ修正。(#2921)
'2022.01.07 GV 備考欄改修対応。(#3040)
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
Dim wStatus						'PHPから送信されたポイント区分パラメータ
Dim wPointKubun					'クエリに組み込むポイント区分

'=======================================================================
'	受け渡し情報取り出し & 初期設定
'=======================================================================
' Getパラメータ
wCustomerNo = ReplaceInput(Trim(Request("cno")))
wOrderNo = ReplaceInput(Trim(Request("ono")))
wPage = ReplaceInput(Trim(Request("page")))	' ページ位置
wPageSize = ReplaceInput(Trim(Request("page_size")))
wStatus = ReplaceInput(Trim(Request("status")))

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

'ステータス
If wStatus = "" Or IsNumeric(wStatus) = False Then
	wStatus = 0
Else
	wStatus = CLng(wStatus)
End If

' ポイント区分
If wStatus = 1 Then
	wPointKubun = "獲得予定"
ElseIf wStatus = 2 Then
	wPointKubun = "獲得"
ElseIf wStatus = 3 Then
	wPointKubun = "利用"
ElseIf wStatus = 4 Then
	wPointKubun = "失効"
ElseIf wStatus = 99 Then
	wPointKubun = "'返品', '返金', '調整', '不足'"
Else
	wPointKubun = ""
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
Dim i 'レコードセットのループイテレータ
Dim j 'JSONのイテレータ
Dim vRS
'Dim vPaymentMethod
Dim vKubun1
Dim vKubun2
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
Dim vOnoFlg
Dim maxLoop '2016.07.14 GV add
Dim addCnt  '2016.07.14 GV add
Dim makerName
Dim itemName
Dim itemId
Dim shipCompDate
Dim kakutokuYoteiPoint
Dim standardDate
Dim webItem

Set oJSON = New aspJSON

' 初期化
i         = 0
j         = 0
vPointZan = 0
vAllPage  = 0
vOffset   = 0
vAdjust   = 1
vOnoFlg   = False
maxLoop   = 0 '2016.07.14 GV add
addCnt    = 0 '2016.07.14 GV add
makerName = ""
itemName = ""
itemId = ""
shipCompDate = ""
kakutokuYoteiPoint = 0
webItem = ""

'--- 該当顧客のポイント明細の取り出し
vSQL = ""
' 全体の獲得予定ポイント
vSQL = vSQL & "SELECT "
vSQL = vSQL & "  sum(T1.ポイント) AS ポイント "
vSQL = vSQL & "FROM "
vSQL = vSQL & "  ポイント明細 T1 WITH (NOLOCK) "
vSQL = vSQL & "INNER JOIN 受注 T2 WITH (NOLOCK) "'2021.08.27 GV add
vSQL = vSQL & "  ON  T2.受注番号 = T1.受注番号 "'2021.08.27 GV add
vSQL = vSQL & "  AND T2.顧客番号 = T1.顧客番号 "'2021.08.27 GV add
vSQL = vSQL & "WHERE "
vSQL = vSQL & "  T1.顧客番号 = " & wCustomerNo
vSQL = vSQL & " AND "
vSQL = vSQL & "  T1.ポイント区分 = '獲得' "
vSQL = vSQL & "AND "
vSQL = vSQL & "  T1.ポイント日付 IS NULL "
vSQL = vSQL & "AND " '2021.09.02 GV add
vSQL = vSQL & "  ((T1.使用受注番号 NOT LIKE 'RA%') OR (T1.使用受注番号 IS NULL)) " '2021.09.02 GV add
vSQL = vSQL & "AND "
vSQL = vSQL & "  T2.受注日 IS NOT NULL " '2021.08.27 GV add

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic
If vRS.EOF = False Then
		If (IsNull(vRS("ポイント"))) Then
			kakutokuYoteiPoint =  0
		Else
			kakutokuYoteiPoint = CStr(Trim(vRS("ポイント")))
		End If
End If

'レコードセットを閉じる
vRS.Close

'レコードセットのクリア
Set vRS = Nothing

'JSONデータに追加
oJSON.data.Add "obtain_yotei_pt" ,kakutokuYoteiPoint



'--------------------------------------------------------
' ここから 2021.07.09 GV
' 旧ver. は get_point_history_list_v3_json.asp を参考
'--------------------------------------------------------
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "  x.基準日 "
vSQL = vSQL & " ,x.表示用基準日 "
vSQL = vSQL & " ,x.出荷完了日 "
vSQL = vSQL & " ,x.受注番号 "
vSQL = vSQL & " ,x.受注明細番号 "
vSQL = vSQL & " ,x.メーカー名 "
vSQL = vSQL & " ,x.商品名 "
vSQL = vSQL & " ,x.色 "
vSQL = vSQL & " ,x.規格 "
vSQL = vSQL & " ,x.商品ID "
vSQL = vSQL & " ,x.Web商品フラグ "
vSQL = vSQL & " ,x.ポイント利用獲得 "
vSQL = vSQL & " ,x.ポイント区分 "
vSQL = vSQL & " ,x.使用受注番号 "
vSQL = vSQL & " ,x.処理区分 "
vSQL = vSQL & " ,x.ポイント獲得日 "
vSQL = vSQL & " ,x.ポイント期限 "
vSQL = vSQL & " ,x.ポイント "
vSQL = vSQL & " ,x.POINT_SORT "
vSQL = vSQL & "FROM "
vSQL = vSQL & "(SELECT "
vSQL = vSQL & "  T2.見積日 AS 基準日 "
vSQL = vSQL & " ,CONVERT(NVARCHAR, T2.見積日, 111) AS 表示用基準日 "
vSQL = vSQL & " ,CONVERT(NVARCHAR, T2.出荷完了日, 111) as 出荷完了日 "
vSQL = vSQL & " ,T1.受注番号 AS 受注番号 "
vSQL = vSQL & " ,T3.受注明細番号 "
vSQL = vSQL & " ,mk.メーカー名 "
vSQL = vSQL & " ,T3.商品名 "
vSQL = vSQL & " ,T3.色 "
vSQL = vSQL & " ,T3.規格 "
vSQL = vSQL & " ,z.商品ID "
vSQL = vSQL & " ,i.Web商品フラグ "
vSQL = vSQL & " ,T1.ポイント区分 AS ポイント利用獲得 "
vSQL = vSQL & " ,(CASE "
vSQL = vSQL & "     WHEN T1.ポイント日付 IS NULL AND T1.ポイント区分 = '獲得' "
vSQL = vSQL & "       THEN '獲得予定' "
vSQL = vSQL & "     WHEN  T1.ポイント日付 IS NOT NULL AND T1.ポイント区分 = '獲得' "
vSQL = vSQL & "       THEN '獲得' "
vSQL = vSQL & "     ELSE T1.ポイント区分 "
vSQL = vSQL & "   END) AS ポイント区分 "
vSQL = vSQL & " ,T1.使用受注番号 "
vSQL = vSQL & " ,(CASE "
vSQL = vSQL & "     WHEN T1.ポイント日付 IS NULL THEN "
vSQL = vSQL & "       CASE "
vSQL = vSQL & "         WHEN LEFT(T1.使用受注番号, 2) = 'RA' THEN '-' "
vSQL = vSQL & "         ELSE '処理中' "
vSQL = vSQL & "       END "
vSQL = vSQL & "     ELSE '処理済' "
vSQL = vSQL & "   END) AS 処理区分 "
vSQL = vSQL & " ,T1.ポイント日付 AS ポイント獲得日 "
vSQL = vSQL & " ,T1.ポイント期限 "
vSQL = vSQL & " ,SUM(T1.ポイント) AS ポイント "
vSQL = vSQL & " ,CASE T1.ポイント区分 WHEN '利用' THEN 0 ELSE 1 END AS POINT_SORT "
vSQL = vSQL & "FROM "
vSQL = vSQL & "  ポイント明細 T1 WITH (NOLOCK) "
vSQL = vSQL & "LEFT JOIN 受注 T2 WITH (NOLOCK) "
vSQL = vSQL & "  ON  T2.受注番号 = T1.受注番号 "
vSQL = vSQL & "  AND T2.顧客番号 = T1.顧客番号 "
vSQL = vSQL & "LEFT JOIN 受注明細 T3 WITH (NOLOCK) "
vSQL = vSQL & " ON  T3.受注番号     = T1.受注番号 "
vSQL = vSQL & " AND T3.受注明細番号 = T1.受注明細番号 "
vSQL = vSQL & "LEFT JOIN メーカー mk WITH (NOLOCK) "
vSQL = vSQL & " ON mk.メーカーコード = T3.メーカーコード "
vSQL = vSQL & "LEFT JOIN 色規格別在庫 z WITH (NOLOCK) "
vSQL = vSQL & "  ON z.メーカーコード = T3.メーカーコード "
vSQL = vSQL & " AND z.商品コード     = T3.商品コード "
vSQL = vSQL & " AND z.色             = T3.色 "
vSQL = vSQL & " AND z.規格           = T3.規格 "
vSQL = vSQL & "LEFT JOIN 商品 i WITH (NOLOCK) "
vSQL = vSQL & "  ON i.メーカーコード = z.メーカーコード "
vSQL = vSQL & " AND i.商品コード     = z.商品コード "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "  T1.顧客番号 = " & wCustomerNo
vSQL = vSQL & "  AND T1.受注番号 > 0 "
vSQL = vSQL & "  AND T2.受注日 IS NOT NULL " '2021.08.27 GV add

'2021.09.02 GV add start
vSQL = vSQL & "  AND "
vSQL = vSQL & "  (CASE "
vSQL = vSQL & "     WHEN "
vSQL = vSQL & "       LEFT(T1.使用受注番号, 2) = 'RA' "
vSQL = vSQL & "       THEN "
vSQL = vSQL & "         CASE "
vSQL = vSQL & "           WHEN T1.ポイント日付 IS NOT NULL THEN 1 "
vSQL = vSQL & "           ELSE 0 "
vSQL = vSQL & "         END "
vSQL = vSQL & "     ELSE 1 "
vSQL = vSQL & "   END) = 1 "
'2021.09.02 GV add end

vSQL = vSQL & "GROUP BY "
vSQL = vSQL & " T1.受注番号,T1.ポイント区分,T1.ポイント日付,T1.使用受注番号,T1.ポイント期限 "
vSQL = vSQL & " ,T2.受注日, T2.見積日, T2.出荷完了日, T3.受注明細番号 "
vSQL = vSQL & " ,T3.商品名, T3.色, T3.規格,z.商品ID, i.Web商品フラグ "
vSQL = vSQL & " ,mk.メーカー名 "
vSQL = vSQL & "HAVING "
vSQL = vSQL & "  SUM(T1.ポイント) <> '0' "
vSQL = vSQL & "UNION "
vSQL = vSQL & "SELECT "
vSQL = vSQL & "  T1.ポイント日付 AS 基準日 "
vSQL = vSQL & " ,CONVERT(NVARCHAR, T1.ポイント日付, 111) as 表示用基準日 "
vSQL = vSQL & " ,CONVERT(NVARCHAR, T1.ポイント日付, 111) as 出荷完了日 "
vSQL = vSQL & " ,T1.受注番号 AS 受注番号 "
vSQL = vSQL & " ,T1.受注明細番号 "
vSQL = vSQL & " ,'' AS メーカー名 "
'vSQL = vSQL & " ,T1.備考 AS 商品名 " ' 2022.01.07 GV mod
vSQL = vSQL & " ,CASE T1.ポイント区分 WHEN '調整' THEN '' ELSE T1.備考 END AS 商品名 " ' 2022.01.07 GV mod
vSQL = vSQL & " ,'' AS 色 "
vSQL = vSQL & " ,'' AS 規格 "
vSQL = vSQL & " ,NULL AS 商品ID "
vSQL = vSQL & " ,'' AS Web商品フラグ "
vSQL = vSQL & " ,T1.ポイント区分 AS ポイント利用獲得 "
vSQL = vSQL & " ,(CASE "
vSQL = vSQL & "     WHEN T1.ポイント日付 IS NULL AND T1.ポイント区分 = '獲得' "
vSQL = vSQL & "       THEN '獲得予定' "
vSQL = vSQL & "     WHEN  T1.ポイント日付 IS NOT NULL AND T1.ポイント区分 = '獲得' "
vSQL = vSQL & "       THEN '獲得' "
vSQL = vSQL & "     ELSE T1.ポイント区分 "
vSQL = vSQL & "   END) AS ポイント区分 "
vSQL = vSQL & " ,T1.使用受注番号 "
vSQL = vSQL & " ,(CASE "
vSQL = vSQL & "     WHEN T1.ポイント日付 IS NULL THEN "
vSQL = vSQL & "       CASE "
vSQL = vSQL & "         WHEN LEFT(T1.使用受注番号, 2) = 'RA' THEN '-' "
vSQL = vSQL & "         ELSE '処理中' "
vSQL = vSQL & "       END "
vSQL = vSQL & "     ELSE '処理済' "
vSQL = vSQL & "   END) AS 処理区分 "
vSQL = vSQL & " ,T1.ポイント日付 AS ポイント獲得日 "
vSQL = vSQL & " ,T1.ポイント期限 "
vSQL = vSQL & " ,SUM(T1.ポイント) AS ポイント "
vSQL = vSQL & " ,CASE T1.ポイント区分 WHEN '利用' THEN 0 ELSE 1 END AS POINT_SORT "
vSQL = vSQL & "FROM "
vSQL = vSQL & "  ポイント明細 T1 WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "  T1.顧客番号 =" & wCustomerNo
vSQL = vSQL & "  AND T1.受注番号 = 0 "

'2021.09.02 GV add start
vSQL = vSQL & " AND "
vSQL = vSQL & " (CASE "
vSQL = vSQL & "     WHEN "
vSQL = vSQL & "       LEFT(T1.使用受注番号, 2) = 'RA' "
vSQL = vSQL & "       THEN "
vSQL = vSQL & "         CASE "
vSQL = vSQL & "           WHEN T1.ポイント日付 IS NOT NULL THEN 1 "
vSQL = vSQL & "           ELSE 0 "
vSQL = vSQL & "         END "
vSQL = vSQL & "     ELSE 1 "
vSQL = vSQL & "   END) = 1 "
'2021.09.02 GV add end

vSQL = vSQL & "GROUP BY "
vSQL = vSQL & "  T1.ポイント日付,T1.受注番号,T1.受注明細番号,T1.備考,T1.ポイント区分,T1.使用受注番号,ポイント期限 "
vSQL = vSQL & "HAVING "
vSQL = vSQL & "  SUM(T1.ポイント) <> '0' "
vSQL = vSQL & ") as x "
vSQL = vSQL & "WHERE "
vSQL = vSQL & " 1 = 1 "

If wPointKubun <> "" THen
	If wStatus = 99 Then
		vSQL = vSQL & " AND x.ポイント区分 IN (" & wPointKubun & ") "
	Else
		vSQL = vSQL & " AND x.ポイント区分 = '" & wPointKubun & "' "
	End If
End If


vSQL = vSQL & "ORDER BY "
vSQL = vSQL & "  x.基準日, x.受注番号, x.受注明細番号 DESC, x.POINT_SORT ASC "
'--------------------------------------------------------
' ここまで 2021.07.09 GV
'--------------------------------------------------------

'@@@@Response.Write(vSQL)


Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic
If vRS.EOF = False Then

	' 全件数をJSONデータにセット
	oJSON.data.Add "cnt" ,vRS.RecordCount

	If (wOrderNo = "") Then
		' リスト追加
		oJSON.data.Add "list" ,oJSON.Collection()
	End If

	'--- 指定ページを表示する為のレコード位置付け(SearchListの処理に倣う)
	' レコード数からのページ数
	vAllPage = Round((vRS.RecordCount / wPageSize) + 0.5)

	' レコード位置の位置付け
	'vRS.AbsolutePage = wPage

	vOffset = vRS.RecordCount - (wPage * wPageSize)

'Response.Write "vRS.RecordCount=" & vRS.RecordCount & "<br>"
'Response.Write "wPageSize=" & wPageSize & "<br>"
'Response.Write "wPage=" & wPage & "<br>"
'Response.Write "vAllPage=" & vAllPage & "<br>"
'Response.Write "vOffset=" & vOffset & "<br>"
'Response.Write "a=" & fix(vRS.RecordCount / wPageSize) & "<br>"

	'最後のページの場合
	If wPage = vAllPage Then
		vOffset = 0
		vAdjust = 2
		'レコード数 - (ページサイズ * fix(レコード数 / ページサイズ))
		maxLoop = vRS.RecordCount - (wPageSize * fix(vRS.RecordCount / wPageSize)) - 1
	Else
		vAdjust = 1
		maxLoop = wPageSize - 1
	End If
'Response.Write "vOffset=" & vOffset & "<br>"
'Response.Write "vAdjust=" & vAdjust & "<br>"
'Response.Write "maxLoop=" & maxLoop & "<br>"

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
		'2016.07.14 GV mod start
		'vPointZan = vPointZan + CLng(vPoint)
		If (CStr(Trim(vRS("処理区分"))) <> "処理中") Then
				vPointZan = vPointZan + CLng(vPoint)
		End If
		'2016.07.14 GV mod end

		' --------------------------------

		'受注番号の指定がある場合、そこまでのポイント残を取得する
		If (wOrderNo <> "") Then
			If (wOrderNo = CStr(Trim(vRS("受注番号")))) Then
				vOnoFlg = True
			End If

			' 指定の受注番号までループ到達かつ異なる受注番号になった場合
			If (vOnoFlg = True) And (wOrderNo <> CStr(Trim(vRS("受注番号")))) Then
				'累積ポイント残を１つ前にもどす
				vPointZan = vPointZan - CLng(vPoint)

				' リスト追加
				oJSON.data.Add "o_no" ,wOrderNo
				oJSON.data.Add "pt_zan" ,vPointZan

				'ループ脱出
				Exit For
			End If
		Else
		'必要なレコード位置の場合に、JSONデータを生成する
			If (wPage <= vAllPage) And (i >= (vOffset)) And (addCnt <= (maxLoop)) Then
				'基準日
				standardDate = CStr(Trim(vRS("表示用基準日")))
	
				'出荷完了日
				If (IsNull(vRS("出荷完了日"))) Then
					shipCompDate = ""
				Else
					shipCompDate = CStr(Trim(vRS("出荷完了日")))
				End If

				'受注番号
				If (IsNull(vRS("受注番号"))) Then
					vOrderNo = ""
				Else
					vOrderNo = CStr(Trim(vRS("受注番号")))
				End If

				'メーカー名
				If (IsNull(vRS("メーカー名"))) Then
					makerName = ""
				Else
					makerName = CStr(Trim(vRS("メーカー名")))
				End If

				'商品名
				If (IsNull(vRS("商品名"))) Then
					itemName = ""
				Else
					itemName = CStr(Trim(vRS("商品名")))

					'色
					If (IsNull(vRS("色"))) Then
					ElseIF Trim(vRS("色")) <> "" Then
						itemName = itemName & " / " & CStr(Trim(vRS("色")))
					End If

					'規格
					If (IsNull(vRS("規格"))) Then
					ElseIF Trim(vRS("規格")) <> "" Then
						itemName = itemName & " / " & CStr(Trim(vRS("規格")))
					End If
				End If

				'商品ID
				If (IsNull(vRS("商品ID"))) Then
					itemId = ""
				Else
					itemId = CStr(Trim(vRS("商品ID")))
				End If

				'Web商品フラグ
				If (IsNull(vRS("Web商品フラグ"))) Then
					webItem = ""
				Else
					webItem = CStr(Trim(vRS("Web商品フラグ")))
				End If


				'ポイント利用獲得(ポイント区分)
				If (IsNull(vRS("ポイント利用獲得"))) Then
					vKubun1 = ""
				Else
					vKubun1 = CStr(Trim(vRS("ポイント利用獲得")))
				End If

				If (CStr(Trim(vRS("処理区分"))) = "処理中") Then
					vKubun1 = "処理中"
				End If

				'ポイント利用獲得(ポイント区分)
				If (IsNull(vRS("ポイント区分"))) Then
					vKubun2 = ""
				Else
					vKubun2 = CStr(Trim(vRS("ポイント区分")))
				End If

				'ポイント獲得日
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
						.Add "std_dt" ,standardDate
						.Add "ship_comp_dt" ,formatDateYYYYMMDD(shipCompDate)
						.Add "o_dt" ,formatDateYYYYMMDD(vOrderDate)
						.Add "o_no" ,vOrderNo
						'.Add "order_type" ,vOrderType
						'.Add "payment_method" ,vPaymentMethod
						.Add "kubun1" ,vKubun1
						.Add "kubun2" ,vKubun2
						.Add "m_name" ,makerName
						.Add "i_name" ,itemName
						.Add "i_id" ,itemId
						.Add "web_flag" ,webItem
						.Add "pt_dt" ,formatDateYYYYMMDD(vPointDate)
						.Add "pt_expire" ,formatDateYYYYMMDD(vPointExpire)
						.Add "pt" ,vPoint
						.Add "pt_zan" ,vPointZan
					End With
				End With

				addCnt = addCnt + 1
			End If

			' イテレータをインクリメント
			j = j + 1
		End If

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
