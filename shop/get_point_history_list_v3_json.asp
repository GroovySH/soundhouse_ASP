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
'2016.03.11 GV 新規作成。(Web注文変更キャンセル機能)
'2016.07.14 GV ポイント処理中を追加。
'2016.12.01 GV 代引きの場合の抽出条件を改修。
'2020.06.06 GV 代引きの場合の抽出条件を改修。
'2020.06.25 GV 代引き以外も「処理中」を表示。(#2458)
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
Dim i 'レコードセットのループイテレータ
Dim j 'JSONのイテレータ
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
Dim vOnoFlg
Dim maxLoop '2016.07.14 GV add
Dim addCnt  '2016.07.14 GV add

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

'--------------------------------------------------------
' コメントアウト ここから 2020.06.25 GV
'--------------------------------------------------------
'ここから
'vSQL = ""
'vSQL = vSQL & "SELECT "
'' 2017.07.14 GV mod start
''vSQL = vSQL & "  CONVERT(NVARCHAR, T1.ポイント日付, 111) AS ご注文日, "

''vSQL = vSQL & "  CASE "
''vSQL = vSQL & "     WHEN T1.ポイント日付 IS NOT NULL "
''vSQL = vSQL & "       THEN CONVERT(NVARCHAR, T1.ポイント日付, 111) "
''vSQL = vSQL & "     WHEN T1.ポイント日付 IS NULL "
''vSQL = vSQL & "       THEN CONVERT(NVARCHAR, T2.出荷完了日, 111) "
''vSQL = vSQL & "  END AS ご注文日, "
'vSQL = vSQL & "CASE "
''  --ポイント日付が設定されており、獲得の場合
'vSQL = vSQL & "  WHEN (T1.ポイント日付 IS NOT NULL) AND (T1.ポイント区分 = '獲得') "
'vSQL = vSQL & "    THEN "
'vSQL = vSQL & "      CASE "
'vSQL = vSQL & "        WHEN T2.支払方法 = '代引き' "
'vSQL = vSQL & "          THEN t2.出荷完了日 "
'vSQL = vSQL & "        ELSE T1.ポイント日付 "
'vSQL = vSQL & "      END "
''  -- ポイント日付が設定されており、獲得以外の場合
'vSQL = vSQL & "  WHEN T1.ポイント日付 IS NOT NULL AND T1.ポイント区分 <> '獲得' "
'vSQL = vSQL & "    THEN T1.ポイント日付 "
'vSQL = vSQL & "  ELSE "
''    -- 上記以外（ポイント日付がNULL、出荷完了日が設定）
'vSQL = vSQL & "    CASE "
'vSQL = vSQL & "      WHEN T2.支払方法 = '代引き' " '処理中
'vSQL = vSQL & "        THEN CONVERT(NVARCHAR, T2.出荷完了日, 111) "
'vSQL = vSQL & "      ELSE  "
'vSQL = vSQL & "        CONVERT(NVARCHAR, T1.ポイント日付, 111) "
'vSQL = vSQL & "      END "
'vSQL = vSQL & "  end as ご注文日, "
'' 2017.07.14 GV mod end


'vSQL = vSQL & "  T1.受注番号 AS ご注文番号, "
''vSQL = vSQL & "  MAX(T2.受注形態) AS ご注文方法, "
''vSQL = vSQL & "  MAX(T2.支払方法) AS お支払方法, "

'vSQL = vSQL & "  T1.ポイント区分 AS ポイント利用獲得, "
'' 2017.07.14 GV add start
'vSQL = vSQL & "  CASE "
'vSQL = vSQL & "     WHEN T1.ポイント日付 IS NULL THEN '処理中' "
'vSQL = vSQL & "     ELSE '処理済' "
'vSQL = vSQL & "  END AS 処理区分, "
'' 2017.07.14 GV add end

'vSQL = vSQL & "  T1.ポイント日付 AS ポイント獲得日, "
'vSQL = vSQL & "  T1.ポイント期限, "
'vSQL = vSQL & "  SUM(T1.ポイント) AS ポイント, "
'vSQL = vSQL & "  CASE T1.ポイント区分 "
'vSQL = vSQL & "    WHEN '利用' THEN 0 "
'vSQL = vSQL & "    ELSE 1 "
'vSQL = vSQL & "  END AS POINT_SORT "
'vSQL = vSQL & "FROM "
'vSQL = vSQL & "  ポイント明細 T1 WITH (NOLOCK) "

''2016.07.14 GV mod start
'vSQL = vSQL & "  LEFT JOIN 受注 T2 WITH (NOLOCK) "
'vSQL = vSQL & "    ON T2.受注番号 = T1.受注番号 "
'vSQL = vSQL & "    AND T2.顧客番号 = T1.顧客番号 "
''2016.07.14 GV mod end

'vSQL = vSQL & "WHERE "
'vSQL = vSQL & "  T1.顧客番号= " & wCustomerNo
''vSQL = vSQL & " AND T1.ポイント日付 IS NOT NULL " ' 2016.07.14 GV mod
'' 2016.07.14 GV add start
'vSQL = vSQL & "  AND T2.削除日 IS NULL "
'vSQL = vSQL & " AND ((T1.ポイント日付 IS NOT NULL) "
''2016.12.01 GV mod start
''vSQL = vSQL & "   OR (T1.ポイント日付 IS NULL AND T2.支払方法 = '代引き' AND T2.出荷完了日 IS NOT NULL)) "
''vSQL = vSQL & "   OR (T1.ポイント日付 IS NULL AND T2.支払方法 = '代引き' AND T2.出荷完了日 IS NOT NULL AND T2.最終入金日 IS NULL)) " ' 2020.06.06 GV mod
'vSQL = vSQL & "   OR (T1.ポイント日付 IS NULL AND T2.支払方法 = '代引き' AND T2.出荷完了日 IS NOT NULL)) " ' 2020.06.06 GV add
''2016.12.01 GV mod end
'' 2016.07.14 GV add end

'vSQL = vSQL & "GROUP BY "
'vSQL = vSQL & "  T1.ポイント区分, "
'vSQL = vSQL & "  T1.ポイント日付, "
'vSQL = vSQL & "  T1.受注番号 "
'vSQL = vSQL & "  ,T1.ポイント期限 "
'vSQL = vSQL & "  ,T2.出荷完了日 " ' 2016.07.14 GV add
'vSQL = vSQL & "  ,T2.支払方法 " ' 2016.07.14 GV add
'vSQL = vSQL & "HAVING "
'vSQL = vSQL & "  SUM(T1.ポイント) <> '0' "
'vSQL = vSQL & "ORDER BY "
'vSQL = vSQL & "  ご注文日, "
'vSQL = vSQL & "  ご注文番号, "
'vSQL = vSQL & "  POINT_SORT ASC"
'--------------------------------------------------------
' コメントアウト ここまで 2020.06.25 GV
'--------------------------------------------------------

'--------------------------------------------------------
' 新規 ここから 2020.06.25 GV
'--------------------------------------------------------
vSQL = vSQL & "SELECT "
vSQL = vSQL & "  x.ご注文日 "
vSQL = vSQL & ",x.ご注文番号 "
vSQL = vSQL & ",x.ポイント利用獲得 "
vSQL = vSQL & ",x.処理区分 "
vSQL = vSQL & ",x.ポイント獲得日 "
vSQL = vSQL & ",x.ポイント期限 "
vSQL = vSQL & ",sum(x.ポイント) AS ポイント "
vSQL = vSQL & ",x.POINT_SORT "
vSQL = vSQL & " FROM ( "
' Xのクエリ
vSQL = vSQL & "SELECT "
vSQL = vSQL & " CASE "
vSQL = vSQL & "   WHEN (T1.ポイント日付 IS NOT NULL) AND (T1.ポイント区分 = '獲得') "
'2020.11.16 takeuchi MOD start
'vSQL = vSQL & "     THEN "
'vSQL = vSQL & "       CASE "
'vSQL = vSQL & "         WHEN T2.支払方法 = '代引き' THEN t2.出荷完了日 "
'vSQL = vSQL & "         ELSE T1.ポイント日付 "
'vSQL = vSQL & "       END "
vSQL = vSQL & "     THEN T1.ポイント日付 "
'2020.11.16 takeuchi MOD end
vSQL = vSQL & "   WHEN T1.ポイント日付 IS NOT NULL AND T1.ポイント区分 <> '獲得' "
vSQL = vSQL & "     THEN T1.ポイント日付 "
vSQL = vSQL & "   ELSE "
vSQL = vSQL & "     CASE "
vSQL = vSQL & "       WHEN T2.支払方法 = '代引き' "
vSQL = vSQL & "         THEN CONVERT(NVARCHAR, T2.出荷完了日, 111) "
vSQL = vSQL & "       ELSE "
vSQL = vSQL & "         CONVERT(NVARCHAR, T1.ポイント日付, 111) "
vSQL = vSQL & "     END "
vSQL = vSQL & " END AS ご注文日 "
vSQL = vSQL & ", T1.受注番号 AS ご注文番号 "
vSQL = vSQL & ", T1.ポイント区分 AS ポイント利用獲得 "
'vSQL = vSQL & ", CASE WHEN T1.ポイント日付 IS NULL THEN '処理中' ELSE '処理済' END AS 処理区分 "

vSQL = vSQL & ",T1.使用受注番号 "
vSQL = vSQL & ",CASE "
vSQL = vSQL & "   WHEN T1.ポイント日付 IS NULL THEN "
vSQL = vSQL & "     CASE "
vSQL = vSQL & "       WHEN LEFT(T1.使用受注番号, 2) = 'RA' THEN ポイント区分 "
vSQL = vSQL & "       ELSE '処理中' "
vSQL = vSQL & "     END "
vSQL = vSQL & "   ELSE '処理済' END AS 処理区分 "

vSQL = vSQL & ", T1.ポイント日付 AS ポイント獲得日 "
vSQL = vSQL & ", T1.ポイント期限 "
vSQL = vSQL & ", SUM(T1.ポイント) AS ポイント "
vSQL = vSQL & ", CASE T1.ポイント区分 WHEN '利用' THEN 0 ELSE 1 END AS POINT_SORT "
vSQL = vSQL & "FROM ポイント明細 T1 WITH (NOLOCK) "
vSQL = vSQL & "LEFT JOIN 受注 T2 WITH (NOLOCK) "
vSQL = vSQL & "  ON T2.受注番号 = T1.受注番号 "
vSQL = vSQL & "  AND T2.顧客番号 = T1.顧客番号 "
vSQL = vSQL & "LEFT JOIN 受注明細 T3 WITH (NOLOCK) "
vSQL = vSQL & " ON T3.受注番号 = T1.受注番号 AND T3.受注明細番号 = T1.受注明細番号 "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "  T1.顧客番号= " & wCustomerNo
vSQL = vSQL & "  AND T2.削除日 IS NULL "
vSQL = vSQL & "  AND (T1.ポイント日付 IS NOT NULL) "
vSQL = vSQL & "GROUP BY "
vSQL = vSQL & "  T1.ポイント区分, T1.ポイント日付, T1.受注番号 ,T1.ポイント期限, T1.使用受注番号, T2.出荷完了日, T2.支払方法 "
vSQL = vSQL & "HAVING "
vSQL = vSQL & "  SUM(T1.ポイント) <> '0' "
vSQL = vSQL & "UNION "
vSQL = vSQL & "SELECT "
vSQL = vSQL & " CASE "
vSQL = vSQL & "   WHEN (T1.ポイント日付 IS NOT NULL) AND (T1.ポイント区分 = '獲得') "
vSQL = vSQL & "     THEN "
vSQL = vSQL & "       CASE "
vSQL = vSQL & "         WHEN T2.支払方法 = '代引き' THEN t2.出荷完了日 "
vSQL = vSQL & "         ELSE T1.ポイント日付 "
vSQL = vSQL & "       END "
vSQL = vSQL & "   WHEN T1.ポイント日付 IS NOT NULL AND T1.ポイント区分 <> '獲得' "
vSQL = vSQL & "     THEN T1.ポイント日付 "
vSQL = vSQL & "   ELSE "
'2020.11.16 takeuchi MOD start
'vSQL = vSQL & "     CASE "
'vSQL = vSQL & "       WHEN T2.支払方法 = '代引き' "
'vSQL = vSQL & "         THEN CONVERT(NVARCHAR, T2.出荷完了日, 111) "
'vSQL = vSQL & "       ELSE "
'vSQL = vSQL & "         CONVERT(NVARCHAR, TA.出荷日, 111) "
'vSQL = vSQL & "     END "
vSQL = vSQL & "     GETDATE() "
'2020.11.16 takeuchi MOD end
vSQL = vSQL & " END AS ご注文日 "
vSQL = vSQL & ", T1.受注番号 AS ご注文番号 "
vSQL = vSQL & ", T1.ポイント区分 AS ポイント利用獲得 "
'vSQL = vSQL & ", CASE WHEN T1.ポイント日付 IS NULL THEN '処理中' ELSE '処理済' END AS 処理区分 "

vSQL = vSQL & ",T1.使用受注番号 "
vSQL = vSQL & ",CASE "
vSQL = vSQL & "   WHEN T1.ポイント日付 IS NULL THEN "
vSQL = vSQL & "     CASE "
vSQL = vSQL & "       WHEN LEFT(T1.使用受注番号, 2) = 'RA' THEN ポイント区分 "
vSQL = vSQL & "       ELSE '処理中' "
vSQL = vSQL & "     END "
vSQL = vSQL & "   ELSE '処理済' END AS 処理区分 "

vSQL = vSQL & ", T1.ポイント日付 AS ポイント獲得日 "
vSQL = vSQL & ", T1.ポイント期限 "
vSQL = vSQL & ", SUM(T1.ポイント) AS ポイント "
vSQL = vSQL & ", CASE T1.ポイント区分 WHEN '利用' THEN 0 ELSE 1 END AS POINT_SORT "
vSQL = vSQL & "FROM ポイント明細 T1 WITH (NOLOCK) "
vSQL = vSQL & "   , 受注 T2 WITH (NOLOCK) "
vSQL = vSQL & "   , 受注明細 T3 WITH (NOLOCK) "
vSQL = vSQL & "     inner join (SELECT T4.受注番号 AS 受注番号, T4.受注明細番号 AS 受注明細番号, MAX(出荷日) AS 出荷日 "
vSQL = vSQL & "                   FROM 出荷明細 T4 WITH (NOLOCK) ,出荷 T5 WITH (NOLOCK) "
vSQL = vSQL & "                  WHERE T4.出荷番号 = T5.出荷番号 "
vSQL = vSQL & "                  GROUP BY T4.受注番号 ,T4.受注明細番号) TA "
vSQL = vSQL & "             on  T3.受注番号 = TA.受注番号 AND T3.受注明細番号 = TA.受注明細番号 "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "  T1.顧客番号= " & wCustomerNo
vSQL = vSQL & "  AND T2.受注番号 = T1.受注番号 AND T2.顧客番号 = T1.顧客番号 "
vSQL = vSQL & "  AND T3.受注番号 = T1.受注番号 AND T3.受注明細番号 = T1.受注明細番号 "
vSQL = vSQL & "  AND T2.削除日 IS NULL "
vSQL = vSQL & "  AND ((T1.ポイント日付 IS NULL AND T2.支払方法 <> '代引き' AND T3.受注数量 = T3.出荷合計数量 AND T3.受注数量 > 0) "
vSQL = vSQL & "       OR "
vSQL = vSQL & "       (T1.ポイント日付 IS NULL AND T2.支払方法 = '代引き' AND T2.出荷完了日 IS NOT NULL)) "
vSQL = vSQL & "GROUP BY "
vSQL = vSQL & "  T1.ポイント区分, T1.ポイント日付, T1.受注番号, T1.ポイント期限, T1.使用受注番号, T2.出荷完了日, T2.支払方法, TA.出荷日 "
vSQL = vSQL & "HAVING "
vSQL = vSQL & "  SUM(T1.ポイント) <> '0' "
vSQL = vSQL & ") AS x "
vSQL = vSQL & " GROUP BY "
vSQL = vSQL & "  x.ご注文日, x.ご注文番号, x.ポイント利用獲得, x.処理区分, x.ポイント獲得日, x.ポイント期限, x.POINT_SORT "
vSQL = vSQL & "ORDER BY "
vSQL = vSQL & "  x.ご注文日, x.ご注文番号, x.POINT_SORT ASC"
'--------------------------------------------------------
' 新規 ここまで 2020.06.25 GV
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

			If (wOrderNo = CStr(Trim(vRS("ご注文番号")))) Then
				vOnoFlg = True
			End If

			' 指定の受注番号までループ到達かつ異なる受注番号になった場合
			If (vOnoFlg = True) And (wOrderNo <> CStr(Trim(vRS("ご注文番号")))) Then
				'累積ポイント残を１つ前にもどす
				vPointZan = vPointZan - CLng(vPoint)

				' リスト追加
				oJSON.data.Add "o_no" ,wOrderNo
				oJSON.data.Add "pt_zan" ,vPointZan

				
				'With oJSON.data("list")
				'	.Add j, oJSON.Collection()
				'	With .item(j)
				'		'.Add "o_dt" ,formatDateYYYYMMDD(vOrderDate)
				'		.Add "o_no" ,wOrderNo
				'		'.Add "kubun" ,vKubun
				'		'.Add "pt_dt" ,formatDateYYYYMMDD(vPointDate)
				'		'.Add "pt_expire" ,formatDateYYYYMMDD(vPointExpire)
				'		'.Add "pt" ,vPoint
				'		.Add "pt_zan" ,vPointZan
				'	End With
				'End With

				'ループ脱出
				Exit For
			End If
		Else
		'必要なレコード位置の場合に、JSONデータを生成する
			'If (wPage <= vAllPage) And (i >= (vOffset)) And (i <= (vOffset + wPageSize - vAdjust)) Then
			If (wPage <= vAllPage) And (i >= (vOffset)) And (addCnt <= (maxLoop)) Then
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

				'ポイント利用獲得(ポイント区分)
				If (IsNull(vRS("ポイント利用獲得"))) Then
					vKubun = ""
				Else
					vKubun = CStr(Trim(vRS("ポイント利用獲得")))
				End If

				'2016.07.14 GV mod start
				If (CStr(Trim(vRS("処理区分"))) = "処理中") Then
					vKubun = "処理中"
				End If
				'2016.07.14 GV mod end


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
						.Add "o_dt" ,formatDateYYYYMMDD(vOrderDate)
						.Add "o_no" ,vOrderNo
						'.Add "order_type" ,vOrderType
						'.Add "payment_method" ,vPaymentMethod
						.Add "kubun" ,vKubun
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
