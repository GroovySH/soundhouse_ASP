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

<%
'========================================================================
'
'	レビューポイントページ
'
'更新履歴
'2010/04/26 an 新規作成
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/01/20 GV データ取得 SELECT文へ LACクエリー案を適用
'2012/01/23 GV 「商品レビュー」テーブルから「商品レビュー集計」テーブル使用に変更 (CreateReviewPointHTML()プロシージャ)
'2012/08/08 if-web リニューアルレイアウト調整
'2014/03/19 GV 消費税増税に伴う2重表示対応
'
'========================================================================

On Error Resume Next

Dim LargeCategoryCd

Dim wSalesTaxRate
Dim wLargeCategoryName
Dim wMidCategoryName
Dim wNoData

Dim wLargeCategoryHTML
Dim wReviewPointHTML

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim Connection
Dim RS
Dim wSQL
Dim wHTML
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.buffer = true

'---- Get input data
LargeCategoryCd = ReplaceInput(Trim(Request("LargeCategoryCd")))

'---- 大カテゴリーコードの指定がない場合
if LargeCategoryCd = "" then
	LargeCategoryCd = "1"
end if

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "RankingReviewPoint.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
end if                                           '2011/08/01 an add e

call close_db()

'---- 想定外の大カテゴリーコードを指定された場合もエラー
if wNoData = "Y" OR Err.Description <> "" then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'========================================================================
'
'	Function	Connect database
'
'========================================================================
'
Function connect_db()

'---- Connect database
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

End function

'========================================================================
'
'	Function	Main
'
'========================================================================
'
Function main()

'---- 消費税率取出し
call getCntlMst("共通","消費税率","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'消費税率
wSalesTaxRate = Clng(wItemNum1)

'---- 大カテゴリー一覧作成
call CreateLargeCategoryHTML()

if wNoData <> "Y" then  '想定外の大カテゴリーを指定されてNoDataの場合はエラー
	'---- レビューポイントランキング作成
	call CreateReviewPointHTML()
end if

End Function

'========================================================================
'
'	Function	大カテゴリー一覧表示
'
'========================================================================
'
Function CreateLargeCategoryHTML()

Dim vCount

'---- 全大カテゴリーを取り出し
wSQL = ""
wSQL = wSQL & "SELECT a.大カテゴリーコード"
wSQL = wSQL & "     , a.大カテゴリー名"
wSQL = wSQL & "  FROM 大カテゴリー a WITH (NOLOCK)"
wSQL = wSQL & " WHERE a.Web大カテゴリーフラグ = 'Y'"
wSQL = wSQL & " ORDER BY a.表示順"

'@@@@@@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

wHTML = ""
vCount = 0
wHTML = wHTML & "  <p id='cat_select'>"  & vbNewLine

Do Until RS.EOF = true

	vCount = vCount + 1
	wHTML = wHTML & "<a href='RankingReviewPoint.asp?LargeCategoryCd=" & RS("大カテゴリーコード") & "'>" & RS("大カテゴリー名") & "</a>"

	if RS("大カテゴリーコード") = LargeCategoryCd then
		wLargeCategoryName = RS("大カテゴリー名")  'レビューポイント一覧のタイトルで使用
	end if

	RS.MoveNext

	'後ろにデータがあれば仕切り線を表示
	if RS.EOF = false then
		wHTML = wHTML & "｜"

		if vCount = 8 then
			wHTML = wHTML & "<br>"  & vbNewLine
		end if
	end if

Loop

if wLargeCategoryName = "" then
	wNoData = "Y" '想定外の大カテゴリーを指定された場合
else
	wHTML = wHTML & vbNewLine
	wHTML = wHTML & "  </p>"  & vbNewLine

	wLargeCategoryHTML = wHTML
end if

RS.close

End Function

'========================================================================
'
'	Function	レビューポイントランキング一覧
'
'========================================================================
'
Function CreateReviewPointHTML()

Dim RSv
Dim vPrice
Dim vPriceNoTax				'2014/03/19 GV add
Dim vItem
Dim vRank

Dim vMakerProduct
Dim vProductName
Dim vProdTermFl '販売終了商品フラグ
Dim vInventoryCd
Dim vInventoryImage

Dim vBGColor

'---- 大カテゴリーごとのレビューポイントTOP25
wSQL = ""
' 2012/01/20 GV Mod Start
'wSQL = wSQL & "SELECT DISTINCT TOP 25"
'wSQL = wSQL & "     (SELECT CAST(AVG(CAST(ISNULL(h.評価,0) AS decimal(1,0))) AS decimal(2,1))"
'wSQL = wSQL & "        FROM 商品レビュー h WITH (NOLOCK)"
'wSQL = wSQL & "       WHERE h.メーカーコード = b.メーカーコード"
'wSQL = wSQL & "         AND h.商品コード = b.商品コード"
'wSQL = wSQL & "     ) AS レビュー評価平均"
'wSQL = wSQL & "     ,(SELECT COUNT(*)"
'wSQL = wSQL & "         FROM 商品レビュー i WITH (NOLOCK)"
'wSQL = wSQL & "        WHERE i.メーカーコード = b.メーカーコード"
'wSQL = wSQL & "          AND i.商品コード = b.商品コード"
'wSQL = wSQL & "     ) AS レビューコメント数"
'wSQL = wSQL & "     ,(SELECT SUM(j.評価)"
'wSQL = wSQL & "         FROM 商品レビュー j WITH (NOLOCK)"
'wSQL = wSQL & "        WHERE j.メーカーコード = b.メーカーコード"
'wSQL = wSQL & "          AND j.商品コード = b.商品コード"
'wSQL = wSQL & "     ) AS レビュー評価合計"
'wSQL = wSQL & "     , b.メーカーコード"
'wSQL = wSQL & "     , b.カテゴリーコード"
'wSQL = wSQL & "     , b.商品コード"
'wSQL = wSQL & "     , b.商品名"
'wSQL = wSQL & "     , b.取扱中止日"
'wSQL = wSQL & "     , b.完売日"
'wSQL = wSQL & "     , b.廃番日"
'wSQL = wSQL & "     , b.B品フラグ"
'wSQL = wSQL & "     , b.ASK商品フラグ"
'wSQL = wSQL & "     , b.希少数量"
'wSQL = wSQL & "     , b.個数限定数量"
'wSQL = wSQL & "     , b.個数限定受注済数量"
'wSQL = wSQL & "     , b.セット商品フラグ"
'wSQL = wSQL & "     , b.メーカー直送取寄区分"
'wSQL = wSQL & "     , b.入荷予定未定フラグ"
'wSQL = wSQL & "     , b.Web納期非表示フラグ"
'wSQL = wSQL & "     , c.色"
'wSQL = wSQL & "     , c.規格"
'wSQL = wSQL & "     , c.B品引当可能数量"
'wSQL = wSQL & "     , c.引当可能数量"
'wSQL = wSQL & "     , c.引当可能入荷予定日"
'wSQL = wSQL & "     , c.商品ID"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN b.個数限定数量 > b.個数限定受注済数量 THEN b.個数限定単価"
'wSQL = wSQL & "         WHEN b.B品フラグ = 'Y' THEN b.B品単価"
'wSQL = wSQL & "       ELSE b.販売単価"
'wSQL = wSQL & "       END AS 実販売単価"
'wSQL = wSQL & "     , d.メーカー名"
'wSQL = wSQL & "     , e.カテゴリー名"
'wSQL = wSQL & "     , g.大カテゴリー名"
'
''色規格があるかどうか 2007/05/30
'wSQL = wSQL & "     , (SELECT COUNT(*)"
'wSQL = wSQL & "          FROM Web色規格別在庫 k WITH (NOLOCK)"
'wSQL = wSQL & "         WHERE k.メーカーコード = b.メーカーコード"
'wSQL = wSQL & "           AND k.商品コード = b.商品コード"
'wSQL = wSQL & "           AND (k.色 != '' OR k.規格 != '')"
'wSQL = wSQL & "           AND k.終了日 IS NULL"
'wSQL = wSQL & "       ) AS 色規格CNT"
'
'wSQL = wSQL & " FROM 商品レビュー a WITH (NOLOCK)"
'wSQL = wSQL & "    , Web商品 b WITH (NOLOCK) "
'wSQL = wSQL & "    , Web色規格別在庫 c WITH (NOLOCK) "
'wSQL = wSQL & "    , メーカー d WITH (NOLOCK)  "
'wSQL = wSQL & "    , カテゴリー e WITH (NOLOCK)"
'wSQL = wSQL & "    , 中カテゴリー f WITH (NOLOCK)  "
'wSQL = wSQL & "    , 大カテゴリー g WITH (NOLOCK) "
'wSQL = wSQL & " WHERE b.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND b.商品コード = a.商品コード"
'wSQL = wSQL & "   AND c.メーカーコード = b.メーカーコード"
'wSQL = wSQL & "   AND c.商品コード = b.商品コード"
'wSQL = wSQL & "   AND d.メーカーコード = b.メーカーコード"
'wSQL = wSQL & "   AND e.カテゴリーコード = b.カテゴリーコード"
'wSQL = wSQL & "   AND f.中カテゴリーコード = e.中カテゴリーコード"
'wSQL = wSQL & "   AND g.大カテゴリーコード = f.大カテゴリーコード"
'wSQL = wSQL & "   AND g.大カテゴリーコード = '" & LargeCategoryCd & "'"
'wSQL = wSQL & "   AND b.Web商品フラグ = 'Y'"
'wSQL = wSQL & "   AND c.色 = ''"
'wSQL = wSQL & "   AND c.規格 = ''"
'wSQL = wSQL & " ORDER BY"
'wSQL = wSQL & "    レビュー評価平均 DESC"
'wSQL = wSQL & "  , レビューコメント数 DESC"
'wSQL = wSQL & "  , d.メーカー名"
'wSQL = wSQL & "  , b.商品名"
' 2012/01/23 GV Mod Start
'wSQL = wSQL & "SELECT DISTINCT TOP 25 "
'wSQL = wSQL & "     (SELECT CAST(AVG(CAST(ISNULL(h.評価, 0) AS DECIMAL(1, 0))) AS DECIMAL(2, 1)) "
'wSQL = wSQL & "        FROM 商品レビュー h WITH (NOLOCK) "
'wSQL = wSQL & "       WHERE     h.メーカーコード = b.メーカーコード "
'wSQL = wSQL & "             AND h.商品コード = b.商品コード "
'wSQL = wSQL & "     ) AS レビュー評価平均 "
'wSQL = wSQL & "     ,(SELECT COUNT(i.ID) "
'wSQL = wSQL & "         FROM 商品レビュー i WITH (NOLOCK) "
'wSQL = wSQL & "        WHERE     i.メーカーコード = b.メーカーコード "
'wSQL = wSQL & "              AND i.商品コード = b.商品コード "
'wSQL = wSQL & "     ) AS レビューコメント数 "
'wSQL = wSQL & "     ,(SELECT SUM(j.評価) "
'wSQL = wSQL & "         FROM 商品レビュー j WITH (NOLOCK) "
'wSQL = wSQL & "        WHERE     j.メーカーコード = b.メーカーコード "
'wSQL = wSQL & "              AND j.商品コード = b.商品コード "
'wSQL = wSQL & "     ) AS レビュー評価合計 "
'wSQL = wSQL & "     , b.メーカーコード "
'wSQL = wSQL & "     , b.カテゴリーコード "
'wSQL = wSQL & "     , b.商品コード "
'wSQL = wSQL & "     , b.商品名 "
'wSQL = wSQL & "     , b.取扱中止日 "
'wSQL = wSQL & "     , b.完売日 "
'wSQL = wSQL & "     , b.廃番日 "
'wSQL = wSQL & "     , b.B品フラグ "
'wSQL = wSQL & "     , b.ASK商品フラグ "
'wSQL = wSQL & "     , b.希少数量 "
'wSQL = wSQL & "     , b.個数限定数量 "
'wSQL = wSQL & "     , b.個数限定受注済数量 "
'wSQL = wSQL & "     , b.セット商品フラグ "
'wSQL = wSQL & "     , b.メーカー直送取寄区分 "
'wSQL = wSQL & "     , b.入荷予定未定フラグ "
'wSQL = wSQL & "     , b.Web納期非表示フラグ "
'wSQL = wSQL & "     , c.色 "
'wSQL = wSQL & "     , c.規格 "
'wSQL = wSQL & "     , c.B品引当可能数量 "
'wSQL = wSQL & "     , c.引当可能数量 "
'wSQL = wSQL & "     , c.引当可能入荷予定日 "
'wSQL = wSQL & "     , c.商品ID "
'wSQL = wSQL & "     , CASE "
'wSQL = wSQL & "         WHEN b.個数限定数量 > b.個数限定受注済数量 THEN b.個数限定単価 "
'wSQL = wSQL & "         WHEN b.B品フラグ = 'Y' THEN b.B品単価 "
'wSQL = wSQL & "       ELSE b.販売単価 "
'wSQL = wSQL & "       END AS 実販売単価 "
'wSQL = wSQL & "     , d.メーカー名 "
'wSQL = wSQL & "     , e.カテゴリー名 "
'wSQL = wSQL & "     , g.大カテゴリー名 "
'wSQL = wSQL & "     , (SELECT COUNT(k.商品コード) "
'wSQL = wSQL & "          FROM Web色規格別在庫 k WITH (NOLOCK) "
'wSQL = wSQL & "         WHERE     k.メーカーコード = b.メーカーコード "
'wSQL = wSQL & "               AND k.商品コード = b.商品コード "
'wSQL = wSQL & "               AND (k.色 != '' OR k.規格 != '') "
'wSQL = wSQL & "               AND k.終了日 IS NULL "
'wSQL = wSQL & "       ) AS 色規格CNT "
'wSQL = wSQL & "FROM "
'wSQL = wSQL & "    商品レビュー                 a WITH (NOLOCK) "
'wSQL = wSQL & "      INNER JOIN Web商品         b WITH (NOLOCK) "
'wSQL = wSQL & "        ON     b.メーカーコード     = a.メーカーコード "
'wSQL = wSQL & "           AND b.商品コード         = a.商品コード "
'wSQL = wSQL & "      INNER JOIN Web色規格別在庫 c WITH (NOLOCK) "
'wSQL = wSQL & "        ON     c.メーカーコード     = b.メーカーコード "
'wSQL = wSQL & "           AND c.商品コード         = b.商品コード "
'wSQL = wSQL & "      INNER JOIN メーカー        d WITH (NOLOCK) "
'wSQL = wSQL & "        ON     d.メーカーコード     = b.メーカーコード "
'wSQL = wSQL & "      INNER JOIN カテゴリー      e WITH (NOLOCK) "
'wSQL = wSQL & "        ON     e.カテゴリーコード   = b.カテゴリーコード "
'wSQL = wSQL & "      INNER JOIN 中カテゴリー    f WITH (NOLOCK) "
'wSQL = wSQL & "        ON     f.中カテゴリーコード = e.中カテゴリーコード "
'wSQL = wSQL & "      INNER JOIN 大カテゴリー    g WITH (NOLOCK) "
'wSQL = wSQL & "        ON     g.大カテゴリーコード = f.大カテゴリーコード "
'wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' ) t1 "
'wSQL = wSQL & "        ON     b.Web商品フラグ      = t1.ShohinWebY "
'wSQL = wSQL & "      LEFT JOIN ( SELECT ''  AS 'Iro' )        t2 "
'wSQL = wSQL & "        ON     c.色               = t2.Iro "
'wSQL = wSQL & "      LEFT JOIN ( SELECT ''  AS 'Kikaku' )     t3 "
'wSQL = wSQL & "        ON     c.規格             = t3.Kikaku "
'wSQL = wSQL & "WHERE "
'wSQL = wSQL & "        t1.ShohinWebY IS NOT NULL "
'wSQL = wSQL & "    AND t2.Iro        IS NOT NULL "
'wSQL = wSQL & "    AND t3.Kikaku     IS NOT NULL "
'wSQL = wSQL & "    AND g.大カテゴリーコード = '" & LargeCategoryCd & "' "
'wSQL = wSQL & "ORDER BY "
'wSQL = wSQL & "      レビュー評価平均 DESC "
'wSQL = wSQL & "    , レビューコメント数 DESC "
'wSQL = wSQL & "    , d.メーカー名 "
'wSQL = wSQL & "    , b.商品名 "

wSQL = wSQL & "SELECT DISTINCT TOP 25 "
wSQL = wSQL & "       a.レビュー評価平均 "
wSQL = wSQL & "     , a.レビュー件数 "
wSQL = wSQL & "     , b.メーカーコード "
wSQL = wSQL & "     , b.カテゴリーコード "
wSQL = wSQL & "     , b.商品コード "
wSQL = wSQL & "     , b.商品名 "
wSQL = wSQL & "     , b.取扱中止日 "
wSQL = wSQL & "     , b.完売日 "
wSQL = wSQL & "     , b.廃番日 "
wSQL = wSQL & "     , b.B品フラグ "
wSQL = wSQL & "     , b.ASK商品フラグ "
wSQL = wSQL & "     , b.希少数量 "
wSQL = wSQL & "     , b.個数限定数量 "
wSQL = wSQL & "     , b.個数限定受注済数量 "
wSQL = wSQL & "     , b.セット商品フラグ "
wSQL = wSQL & "     , b.メーカー直送取寄区分 "
wSQL = wSQL & "     , b.入荷予定未定フラグ "
wSQL = wSQL & "     , b.Web納期非表示フラグ "
wSQL = wSQL & "     , c.色 "
wSQL = wSQL & "     , c.規格 "
wSQL = wSQL & "     , c.B品引当可能数量 "
wSQL = wSQL & "     , c.引当可能数量 "
wSQL = wSQL & "     , c.引当可能入荷予定日 "
wSQL = wSQL & "     , c.商品ID "
wSQL = wSQL & "     , CASE "
wSQL = wSQL & "         WHEN b.個数限定数量 > b.個数限定受注済数量 THEN b.個数限定単価 "
wSQL = wSQL & "         WHEN b.B品フラグ = 'Y' THEN b.B品単価 "
wSQL = wSQL & "       ELSE b.販売単価 "
wSQL = wSQL & "       END AS 実販売単価 "
wSQL = wSQL & "     , d.メーカー名 "
wSQL = wSQL & "     , e.カテゴリー名 "
wSQL = wSQL & "     , g.大カテゴリー名 "
wSQL = wSQL & "     , (SELECT COUNT(k.商品コード) "
wSQL = wSQL & "          FROM Web色規格別在庫 k WITH (NOLOCK) "
wSQL = wSQL & "         WHERE     k.メーカーコード = b.メーカーコード "
wSQL = wSQL & "               AND k.商品コード = b.商品コード "
wSQL = wSQL & "               AND (k.色 != '' OR k.規格 != '') "
wSQL = wSQL & "               AND k.終了日 IS NULL "
wSQL = wSQL & "       ) AS 色規格CNT "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    商品レビュー集計             a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN Web商品         b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.メーカーコード     = a.メーカーコード "
wSQL = wSQL & "           AND b.商品コード         = a.商品コード "
wSQL = wSQL & "      INNER JOIN Web色規格別在庫 c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.メーカーコード     = b.メーカーコード "
wSQL = wSQL & "           AND c.商品コード         = b.商品コード "
wSQL = wSQL & "      INNER JOIN メーカー        d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.メーカーコード     = b.メーカーコード "
wSQL = wSQL & "      INNER JOIN カテゴリー      e WITH (NOLOCK) "
wSQL = wSQL & "        ON     e.カテゴリーコード   = b.カテゴリーコード "
wSQL = wSQL & "      INNER JOIN 大カテゴリー    g WITH (NOLOCK) "
wSQL = wSQL & "        ON     g.大カテゴリーコード = a.大カテゴリーコード "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' ) t1 "
wSQL = wSQL & "        ON     b.Web商品フラグ      = t1.ShohinWebY "
wSQL = wSQL & "      LEFT JOIN ( SELECT ''  AS 'Iro' )        t2 "
wSQL = wSQL & "        ON     c.色               = t2.Iro "
wSQL = wSQL & "      LEFT JOIN ( SELECT ''  AS 'Kikaku' )     t3 "
wSQL = wSQL & "        ON     c.規格             = t3.Kikaku "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        t1.ShohinWebY IS NOT NULL "
wSQL = wSQL & "    AND t2.Iro        IS NOT NULL "
wSQL = wSQL & "    AND t3.Kikaku     IS NOT NULL "
wSQL = wSQL & "    AND a.大カテゴリーコード = '" & LargeCategoryCd & "' "
wSQL = wSQL & "ORDER BY "
wSQL = wSQL & "      a.レビュー評価平均 DESC "
wSQL = wSQL & "    , a.レビュー件数 DESC "
wSQL = wSQL & "    , d.メーカー名 "
wSQL = wSQL & "    , b.商品名 "
' 2012/01/23 GV Mod End
' 2012/01/20 GV Mod End

'@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

wHTML = ""
wHTML = wHTML & "<!--  container START  -->" & vbNewLine
wHTML = wHTML & "  <div id='container'>" & vbNewLine
wHTML = wHTML & "    <h1>" & wLargeCategoryName & "</h1>" & vbNewLine
wHTML = wHTML & "    <!-- s_box TH START -->" & vbNewLine
wHTML = wHTML & "    <div id='s_box_th_rp'>" & vbNewLine
wHTML = wHTML & "      <div id='th_no'>順位</div>" & vbNewLine
wHTML = wHTML & "      <div id='th_prod'>メーカー　商品</div>" & vbNewLine
wHTML = wHTML & "      <div id='th_cat'>カテゴリー</div>" & vbNewLine
wHTML = wHTML & "      <div id='th_point'>レビューポイント</div>" & vbNewLine
wHTML = wHTML & "      <div id='th_stock'>在庫状況</div>" & vbNewLine
wHTML = wHTML & "      <div id='th_cart'>カート</div>" & vbNewLine
wHTML = wHTML & "    </div>" & vbNewLine
wHTML = wHTML & "    <!-- s_box TH END -->" & vbNewLine

if RSv.EOF = true then
	wHTML = wHTML & "<p>レビューが投稿されていません。</p>" & vbNewLine   '商品レビューデータが全くない場合
	wHTML = wHTML & "</div>"
else

	vRank = 0    '順位のカウンタ

	Do Until RSv.EOF = true

		vPrice = FormatNumber(calcPrice(RSv("実販売単価"), wSalesTaxRate),0)
		vPriceNoTax = FormatNumber(RSv("実販売単価"),0)								'2014/03/19 GV add
		vItem = Server.URLEncode(RSv("メーカーコード") & "^" & RSv("商品コード") & "^" & "^")
		vRank = vRank + 1  '順位

		'---- 偶数と奇数で背景色を変更
		if vRank Mod 2 <> 0 then
			vBGColor = "s_box1"
		else
			vBGColor = "s_box2"
		end if

		'---- 在庫状況表示のため、終了チェック
		vProdTermFl = "N"
		if isNull(RSv("取扱中止日")) = false then		'取扱中止
			vProdTermFl = "Y"
		end if
		if isNull(RSv("廃番日")) = false AND RSv("引当可能数量") <= 0 then		'廃番で在庫無し
			vProdTermFl = "Y"
		end if
		if isNull(RSv("完売日")) = false then		'完売商品
			vProdTermFl = "Y"
		end if
		if RSv("B品フラグ") = "Y" AND RSv("B品引当可能数量") <= 0 then    'B品で在庫なし
			vProdTermFl = "Y"
		end if

		'---- 在庫状況作成
		if RSv("色規格CNT") = 0 then
			vInventoryCd = GetInventoryStatus(RSv("メーカーコード"),RSv("商品コード"),RSv("色"),RSv("規格"),RSv("引当可能数量"),RSv("希少数量"),RSv("セット商品フラグ"),RSv("メーカー直送取寄区分"),RSv("引当可能入荷予定日"),vProdTermFl)

			'---- 在庫状況、色を最終セット
			call GetInventoryStatus2(RSv("引当可能数量"), RSv("Web納期非表示フラグ"), RSv("入荷予定未定フラグ"), RSv("廃番日"), RSv("B品フラグ"), RSv("B品引当可能数量"), RSv("個数限定数量"), RSv("個数限定受注済数量"), vProdTermFl, vInventoryCd, vInventoryImage)

		end if

		wHTML = wHTML & "    <!-- s_box START -->" & vbNewLine
		wHTML = wHTML & "    <div class='" & vBGColor & "'>" & vbNewLine
		wHTML = wHTML & "      <div class='s_box_height'>" & vbNewLine
		wHTML = wHTML & "        <div class='rp_no'>" & vbNewLine

		'---- 1～3位は王冠表示
		if vRank <= 3 then
			wHTML = wHTML & "          <div class='crown_pad'>" & vbNewLine
			wHTML = wHTML & "            <img height='30' src='images/ranking/ico_no" & vRank & "crown.gif' alt='' width='41'>" & vbNewLine
			wHTML = wHTML & "          </div>" & vbNewLine
		'---- 4～25位は順位表示
		else
			wHTML = wHTML & "          " & vRank & vbNewLine
		end if

		wHTML = wHTML & "        </div>" & vbNewLine
		wHTML = wHTML & "        <div class='rp_prod'>" & vbNewLine
		wHTML = wHTML & "          <a href='ProductDetail.asp?Item=" & vItem & "'>" & vbNewLine

		'--- メーカー名＋商品名が長くて2行になる場合は"..."で省略
		vMakerProduct = RSv("メーカー名") & " " & RSv("商品名")
		if Len(vMakerProduct) > 33 then
			vProductName = Left(RSv("商品名"), 30-Len(RSv("メーカー名"))) &  "..."
		else
			vProductName = RSv("商品名")
		end if

		wHTML = wHTML & "            <span class='txt_maker'>" & RSv("メーカー名") & "</span>&nbsp;" & vbNewLine
		wHTML = wHTML & "            <span class='txt_product'>" & vProductName & "</span><br>" & vbNewLine
		wHTML = wHTML & "            <span class='txt_price_h'>衝撃特価		：</span>"  & vbNewLine
		wHTML = wHTML & "            <span class='txt_price_d'>"

		'---- ASK商品はASK表示→<a>を入れ子にできないのでリンクはなし
		if RSv("ASK商品フラグ") = "Y" then
			wHTML = wHTML & "ASK"
		else
'2014/03/19 GV mod start ---->
'			wHTML = wHTML & FormatNumber(vPrice,0) & "円(税込)"
			wHTML = wHTML & FormatNumber(vPriceNoTax,0) & "円(税抜)</span>"
			wHTML = wHTML & "　<span class='txt_price_t'>(税込&nbsp;" & FormatNumber(vPrice,0) & "円)</span>"
		end if

'		wHTML = wHTML & "</span>　" & vbNewLine
'2014/03/19 GV mod end   <----
		wHTML = wHTML & "          </a>" & vbNewLine
		wHTML = wHTML & "        </div>" & vbNewLine
		wHTML = wHTML & "        <div class='rp_cat'>" & vbNewLine
		wHTML = wHTML & "          <a href='SearchList.asp?i_type=c&s_category_cd=" & RSv("カテゴリーコード")  & "'><strong>" & RSv("カテゴリー名") & "</strong></a>" & vbNewLine
		wHTML = wHTML & "        </div>" & vbNewLine
		wHTML = wHTML & "        <div class='rp_point'>" & vbNewLine
		wHTML = wHTML & "          <div class='pt8'>" & vbNewLine
' 2012/01/23 GV Mod Start
'		wHTML = wHTML & "            " & CreateReviewImg(RSv("レビューコメント数"),RSv("レビュー評価合計")) &  "<br>" & vbNewLine
		wHTML = wHTML & "            " & CreateReviewImg(RSv("レビュー件数"), RSv("レビュー評価平均")) &  "<br>" & vbNewLine
' 2012/01/23 GV Mod End
		wHTML = wHTML & "          </div>" & vbNewLine
		wHTML = wHTML & "        </div>" & vbNewLine
		wHTML = wHTML & "        <div class='rp_stock'>" & vbNewLine
		wHTML = wHTML & "          <div class='pt12'>" & vbNewLine

		if RSv("色規格CNT") = 0 then
			wHTML = wHTML & "            <img height='10' src='images/" & vInventoryImage & "' width='10' alt=''> " & vInventoryCd & vbNewLine
		end if

		wHTML = wHTML & "          </div>" & vbNewLine
		wHTML = wHTML & "        </div>" & vbNewLine
		wHTML = wHTML & "        <div class='rp_cart'>" & vbNewLine
		wHTML = wHTML & "        <form name='f_item' method='post' action='OrderPreInsert.asp' onSubmit='return order_onClick(this);'>" & vbNewLine

		'---- 色規格なし
		if RSv("色規格CNT") = 0 then
			if vProdTermFl = "Y" then
				wHTML = wHTML & "            <img src='images/icon_sold.gif' alt='完売'><br>" & vbNewLine
			else
				wHTML = wHTML & "            <input type='hidden' name='qt' value='1'>" & vbNewLine
				wHTML = wHTML & "            <input type='hidden' name='maker_cd' value='" & RSv("メーカーコード") & "'>" & vbNewLine
				wHTML = wHTML & "            <input type='hidden' name='product_cd' value='" & RSv("商品コード") & "'>" & vbNewLine
				wHTML = wHTML & "            <input type='hidden' name='category_cd' value='" & RSv("カテゴリーコード") & "'>" & vbNewLine
				wHTML = wHTML & "            <input type='image' src='images/btn_cart.png' style='vertical-align:middle' alt='カートへ'><br>" & vbNewLine
			end if

		'----色規格あり
		else
			if vProdTermFl = "Y" then
				wHTML = wHTML & "            <img src='images/icon_sold.gif' alt='完売'><br>" & vbNewLine
			else
				wHTML = wHTML & "            <input type='hidden' name='qt' value='0'>" & vbNewLine
				wHTML = wHTML & "            <a href='ProductDetail.asp?Item=" & vItem & "'><img src='images/btn_detail.png' alt='詳細を見る'></a><br>" & vbNewLine
			end if
		end if

		wHTML = wHTML & "              商品ID:" & RSv("商品ID") & vbNewLine

		wHTML = wHTML & "        </form>" & vbNewLine
		wHTML = wHTML & "        </div>" & vbNewLine
		wHTML = wHTML & "      </div>" & vbNewLine
		wHTML = wHTML & "    </div>" & vbNewLine
		wHTML = wHTML & "    <!-- s_box END -->" & vbNewLine



		RSv.MoveNext

	Loop

	wHTML = wHTML & "  </div>" & vbNewLine
	wHTML = wHTML & "  <!-- container END -->" & vbNewLine

end if

wReviewPointHTML = wHTML

RSv.close

End Function

'========================================================================
'
'	Function	商品レビュー画像作成
'
'========================================================================
'
' 2012/01/23 GV Mod Start
'Function CreateReviewImg(pRatingNum,pRatingSum)
Function CreateReviewImg(pRatingNum, pAvgRating)
' 2012/01/23 GV Mod End

Dim vAvgRating
Dim v1Cnt
Dim v0Cnt
Dim vHalfCnt
Dim vReview
Dim i

vReview = ""

if pRatingNum > 0 then
	'---- レビュー評価の平均を計算
' 2012/01/23 GV Mod Start
	vAvgRating = Round(pAvgRating, 1)
' 2012/01/23 GV Mod End
	v1Cnt = Fix(vAvgRating)
	if (vAvgRating - v1Cnt) >= 0.5 then
		vHalfCnt = 1
	else
		vHalfCnt = 0
	end if
	v0Cnt = 5 - v1Cnt - vHalfCnt

	'---- 総合評価表示
	For i=1 to v1Cnt
		vReview = vReview & "<img src='images/review_icon10.png' alt=''>"
	Next
	if vHalfcnt = 1 then
		vReview = vReview & "<img src='images/review_icon05.png' alt=''>"
	end if
	For i=1 to v0Cnt
		vReview = vReview & "<img src='images/review_icon00.png' alt=''>"
	Next
end if

CreateReviewImg = vReview

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
'
Function close_db()

Connection.close
Set Connection= Nothing    '2011/08/01 an add

End function

'========================================================================
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<title>レビューポイント｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/Ranking.css?20140401a" type="text/css">
<script type="text/javascript">
//
// ====== 	Function:	order_onClick
//
function order_onClick(pForm){
	if (pForm.qt.value == ""){
		pForm.qt.value = 0;
	}else{
		if (numericCheck(pForm.qt.value) == false){
			pForm.qt.value = 0;
		}
	}
	if (pForm.qt.value > 0){
		return true;
	}else{
		return false;
	}
}
</script>
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->

<div id="globalMain">
	<span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>

<!-- コンテンツstart -->
<div id="globalContents">
<!--
  <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
    <p class="home"><a href="../"><img src="../images/icon_home.gif" alt="HOME"></a></p>
    <ul id="path">
      <li class="now">レビューポイント</li>
    </ul>
  </div></div></div>
  <h1 class="title">レビューポイント</h1>
-->

<!-- Mainpage START -->
<div id="ranking_key_main_flame">

<!-- Menu START -->
  <div id="ranking_key_top_menu">
    <div class="top_menu_parts">
      <a href="BestSellerList.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image15','','images/ranking/ts_btn_on.jpg',1)"><img src="images/ranking/ts_btn_off.jpg" alt="" name="Image15" width="114" height="80" />
      </a>
    </div>
    <div class="top_menu_parts">
      <a href="RankingSearchWord.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image163','','images/ranking/sk_btn_on.jpg',1)">
        <img src="images/ranking/sk_btn_off.jpg" alt="" name="Image163" width="114" height="80" />
      </a>
    </div>
    <!--
    <div class="top_menu_parts">
      <a href="RankingAccess.asp?RankType=<%=Server.URLEncode("商品ビュー")%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image12','','images/ranking/noc_btn_on.jpg',1)">
        <img src="images/ranking/noc_btn_off.jpg" alt="" name="Image12" width="114" height="80" />
      </a>
    </div>
    -->
    <div class="top_menu_parts">
      <a href="RankingAccess.asp?RankType=<%=Server.URLEncode("友達にお勧め")%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image13','','images/ranking/rtaf_btn_on.jpg',1)"><img src="images/ranking/rtaf_btn_off.jpg" alt="" name="Image13" width="114" height="80" />
      </a>
    </div>
    <div class="top_menu_parts">
      <a href="RankingAccess.asp?RankType=<%=Server.URLEncode("欲しいものリスト")%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image14','','images/ranking/wl_btn_on.jpg',1)"><img src="images/ranking/wl_btn_off.jpg" alt="" name="Image14" width="114" height="80" />
      </a>
    </div>
    <div class="top_menu_parts">
      <a href="RankingReview.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image16','','images/ranking/nor_btn_on.jpg',1)"><img src="images/ranking/nor_btn_off.jpg" alt="" name="Image16" width="114" height="80" />
      </a>
    </div>
    <div class="top_menu_parts">
      <a href="RankingReviewPoint.asp">
      <img src="images/ranking/rr_btn_on.jpg" alt="" name="Image17" width="113" height="80" />
      </a>
    </div>
  </div>
<!-- 大カテゴリー一覧 -->
<%=wLargeCategoryHTML%>

<!-- レビューポイントTOP 25 一覧 -->
<%=wReviewPointHTML%>

</div>
<!-- Mainpage END -->

  </div>
  <div id="globalSide">
    <!--#include file="../Navi/NaviSide.inc"-->
  </div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>