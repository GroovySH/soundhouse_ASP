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
<!--#include file="../common/SalesCommon.inc"-->
<!--#include file="../common/SearchListCommon.inc"-->
<%
'========================================================================
'
'    数量限定バーゲン
'
'更新履歴
'2009/09/28 an SearchList.aspから分離して新規作成
'2010/03/18 an 並び替え種別変更（新商品、おすすめ評価順の追加、不要並び替えの削除）
'2010/05/13 an レビュー評価平均の計算条件「ショップコメント IS NULL」を削除
'              レビュー評価平均計算時にCAST(decimal)を追加し、小数点以下を考慮
'2010/06/08 an 左NAVIの絞り込み条件変更対応
'2010/06/29 an 複数の大カテゴリーに所属する商品が重複表示される不具合を修正
'2010/07/09 st レビュー画像作成関数 CreateReviewImg を削除
'2010/07/12 st 並び替え条件変更・追加（おすすめ評価⇒評価順に変更、評価件数順の追加）
'2010/07/16 an 在庫有ソートの不具合修正（B品、個数限定の条件を上に）
'2011/02/23 GV(dy) #826 送料完全無料表示の対応
'2011/05/25 hn 色規格ありでも在庫情報を表示するように変更
'2011/06/09 hn 廃番で在庫なし＋発注なし　の時に完売とするように変更
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/01/10 GV Webセール商品テーブルを駆動表とする対応
'2012/07/11 ok リニューアル新デザイン変更
'2012/10/22 ok 並び替えにお得順追加
'2012/11/09 ok 絞り込み用ソート順を表示順に変更
'2014/03/19 GV 消費税増税に伴う2重表示対応
'
'========================================================================

On Error Resume Next


'Dim s_mid_category_cd       '2010/06/08 an del
'Dim s_category_cd           '2010/06/08 an del
Dim s_maker_cd
Dim s_product_cd
Dim sPriceFrom
Dim sPriceTo
Dim sSeriesCd
Dim i_page
Dim i_sort
Dim i_page_size
Dim i_ListType

Dim wSalesTaxRate
Dim wHikaku
Dim wTemp
Dim LargeCategoryCd          '2010/06/08 an add
Dim MidCategoryCd            '2010/06/08 an add
Dim CategoryCd               '2010/06/08 an add

Dim wListHTML
Dim wCountHTML
Dim wMakerHTML
Dim wNaviMakerHTML
Dim wNaviCategoryHTML
Dim wNaviLargeCategoryHTML   '2010/06/08 an add
Dim wNaviMidCategoryHTML     '2010/06/08 an add
Dim wNaviPricerangeHTML
Dim wMakerInfoHTML
Dim wNoDataHTML              '2010/06/08 an add

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim Connection
Dim RS

Dim wSQL
Dim wSQL2              '2010/06/08 an add
Dim wSQLMaker
Dim wSQLCategory
Dim wSQLMidCategory    '2010/06/08 an add
Dim wSQLLargeCategory  '2010/06/08 an add
Dim wSQLPricerange

Dim wNoData
Dim wHTML
Dim wFootprintHTML   '2010/06/08 an add
Dim wErrDesc   '2011/08/01 an add
Dim wTitle   '2012/07/11 ok add


'========================================================================

Response.buffer = true

'---- Get input data
s_product_cd = ReplaceInput(Trim(Request("s_product_cd")))
s_maker_cd = ReplaceInput(Trim(Request("s_maker_cd")))
LargeCategoryCd = ReplaceInput(Trim(Request("s_large_category_cd")))   '2010/06/08 an add
MidCategoryCd = ReplaceInput(Trim(Request("s_mid_category_cd")))       '2010/06/08 an add
CategoryCd = ReplaceInput(Trim(Request("s_category_cd")))              '2010/06/08 an mod
sPriceFrom = ReplaceInput(Trim(Request("sPriceFrom")))
sPriceTo = ReplaceInput(Trim(Request("sPriceTo")))
sSeriesCd = ReplaceInput(Trim(Request("sSeriesCd")))
i_page = ReplaceInput(Trim(Request("i_page")))
i_sort = ReplaceInput(Trim(Request("i_sort")))
i_page_size = ReplaceInput(Trim(Request("i_page_size")))
i_ListType = ReplaceInput(Trim(Request("i_ListType")))

if ISNumeric(sPriceFrom) = false then
    sPriceFrom = 0
end if
if ISNumeric(sPriceTo) = false then
    sPriceTo = 9999999
end if

sPriceFrom = CCur(sPriceFrom)
sPriceTo = CCur(sPriceTo)

if sPriceTo < sPriceFrom then
    wTemp = sPriceFrom
    sPriceFrom = sPriceTo
    sPriceTo = wTemp
end if

'---- 比較Cookie取り出し
wHikaku = Session("compare")

'---- 表示タイプ取り出し    09/05/26
if i_ListType = "" then
    i_ListType = Session("ListType")
    if i_ListType = "" then
        i_ListType = "type1"
        Session("ListType") = i_ListType
    end if
else
    Session("ListType") = i_ListType
end if

'---- ページサイズ設定
if i_page = "" then
    i_page = 1
else
    i_page = Clng(i_page)
end if

if i_page_size = "" then
    i_page_size = Session("PageSize")
end if

if i_page_size = "" then
    i_page_size = g_page_size
else
    i_page_size = Clng(i_page_size)
end if

if i_ListType = "type1" then
    if i_page_size = 12 then i_page_size = 10
    if i_page_size = 32 then i_page_size = 30
    if i_page_size = 52 then i_page_size = 50
else
    if i_page_size = 10 then i_page_size = 12
    if i_page_size = 30 then i_page_size = 32
    if i_page_size = 50 then i_page_size = 52
end if

Session("PageSize") = i_page_size

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "BargainSale.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then
    Response.Redirect g_HTTP & "shop/Error.asp"
end if

'========================================================================
'
'    Function    Connect database
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
'    Function    Main
'
'========================================================================
'
Function main()

Dim vPointer

'---- 消費税率取出し
call getCntlMst("共通","消費税率","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)            '消費税率
wSalesTaxRate = Clng(wItemNum1)

'---- 該当商品取り出し
call GetProducts()

'---- パン屑リスト作成
call fCreateFootprintHTML("BargainSale.asp","衝撃特価品")  '2010/06/08 an add

'-----
if RS.EOF = true then
	wNoData = "Y"
    wNoDataHTML = wNoDataHTML & "該当する商品が見つかりません。" & vbNewLine
else

	'----- ListHTML作成
	call fCreateSearchListHTML(RS, i_page_size, i_page, i_ListType, wSalesTaxRate, wListHTML)

	'----- お得順設定	'2012/10/22 ok Add
	wListHTML = Replace(wListHTML,"選択してください</option>","選択してください</option>" & vbNewLIne & "<option value='Nesage_DESC'>お得順</option>")

	'---- メーカー情報作成
	if s_maker_cd <> "" then
	    call CreateMakerInfo()
	end if

	'---- NAVI Left Sale用HTML作成     '2010/06/08 an mod s
	if LargeCategoryCd <> "" then
		'----- 大カテゴリーコード指定時、中カテゴリー一覧作成　NAVI用
		call fCreateSalesNAVIMidCategoryHTML(wSQLMidCategory, wNaviMidCategoryHTML)
	else
		if MidCategoryCd <> "" then
			'----- 中カテゴリーコード指定時、カテゴリー一覧作成　NAVI用
			call fCreateSalesNAVICategoryHTML(wSQLCategory, wNaviCategoryHTML)
		else
			if CategoryCd <> "" then
				'----- カテゴリーコード指定時、メーカー一覧作成
				call fCreateSalesNAVIMakerHTML(wSQLMaker, wNaviMakerHTML)
			else
				'---- 指定なしの場合、大カテゴリー一覧作成
				call fCreateSalesNAVILargeCategoryHTML(wSQLLargeCategory, wNaviLargeCategoryHTML)
			end if
		end if
	end if

	'----- 価格帯選択作成　NAVI用
	call fCreateSalesNAVIPriceRangeHTML(wSQLPriceRange, wSalesTaxRate, wNaviPriceRangeHTML)    '2010/06/08 an mod e


end if

RS.close

End Function


'========================================================================
'
'    Function    該当商品取り出し
'
'========================================================================
'
Function GetProducts()

Dim v_order

'---- SQL作成
wSQL = ""
' 20120110 GV Mod Start
'wSQL = wSQL & "SELECT DISTINCT"    '2005/07/19
'wSQL = wSQL & "       a.メーカーコード"            '2005/07/19
'wSQL = wSQL & "     , a.商品コード"
'wSQL = wSQL & "     , a.商品名"
'wSQL = wSQL & "     , a.商品概略Web"
'wSQL = wSQL & "     , a.送料区分"
'wSQL = wSQL & "     , a.特定商品個口"
'wSQL = wSQL & "     , a.重量商品送料"
'wSQL = wSQL & "     , a.商品画像ファイル名_小"
'wSQL = wSQL & "     , a.商品備考"
'wSQL = wSQL & "     , a.標準単価"
'wSQL = wSQL & "     , a.販売単価"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN a.B品フラグ = 'Y' THEN a.B品単価"   '2010/07/16 an mod
'wSQL = wSQL & "         WHEN a.個数限定数量 > a.個数限定受注済数量 THEN a.個数限定単価"   '2010/07/16 an mod
'wSQL = wSQL & "         ELSE a.販売単価"
'wSQL = wSQL & "       END AS 実販売単価"
'wSQL = wSQL & "     , a.個数限定単価"
'wSQL = wSQL & "     , a.個数限定数量"
'wSQL = wSQL & "     , a.個数限定受注済数量"
'wSQL = wSQL & "     , a.オープン価格フラグ"
'wSQL = wSQL & "     , a.メーカー直送取寄区分"
'wSQL = wSQL & "     , a.ASK商品フラグ"
'wSQL = wSQL & "     , a.取扱中止日"
'wSQL = wSQL & "     , a.廃番日"
'wSQL = wSQL & "     , a.終了日"
'wSQL = wSQL & "     , a.希少数量"
'wSQL = wSQL & "     , a.セット商品フラグ"
'wSQL = wSQL & "     , a.カテゴリーコード"
'wSQL = wSQL & "     , a.直輸入品フラグ"
'wSQL = wSQL & "     , a.試聴フラグ"
'wSQL = wSQL & "     , a.試聴URL"
'wSQL = wSQL & "     , a.動画フラグ"
'wSQL = wSQL & "     , a.動画URL"
'wSQL = wSQL & "     , a.Web納期非表示フラグ"
'wSQL = wSQL & "     , a.入荷予定未定フラグ"
'wSQL = wSQL & "     , a.商品スペック使用不可フラグ"
'wSQL = wSQL & "     , a.B品単価"
'wSQL = wSQL & "     , a.完売日"
'wSQL = wSQL & "     , a.発売日"
'wSQL = wSQL & "     , a.前回単価変更日"
'wSQL = wSQL & "     , a.前回販売単価"
'wSQL = wSQL & "     , a.B品フラグ"
'wSQL = wSQL & "     , a.初回登録日"   '2010/03/18 an add
'wSQL = wSQL & "     , a.送料完全無料商品フラグ"				' 2011/02/23 GV Add
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN (a.お勧め商品表示順 = 0) THEN 999999"
'wSQL = wSQL & "         ELSE a.お勧め商品表示順"
'wSQL = wSQL & "       END AS お勧め商品表示順"
'wSQL = wSQL & "     , b.色"
'wSQL = wSQL & "     , b.規格"
'wSQL = wSQL & "     , b.引当可能数量"
'wSQL = wSQL & "     , b.発注数量"								'2011/06/09 hn add
'wSQL = wSQL & "     , b.引当可能入荷予定日"
'wSQL = wSQL & "     , b.B品引当可能数量"
'wSQL = wSQL & "     , b.商品ID"
'wSQL = wSQL & "     , b.適正在庫数量"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN a.B品フラグ = 'Y' THEN b.B品引当可能数量"    '2010/07/16 an mod
'wSQL = wSQL & "         WHEN a.個数限定数量 > 0 AND (a.個数限定数量-a.個数限定受注済数量) > 0 THEN a.個数限定数量-a.個数限定受注済数量"    '2010/07/16 an mod
'wSQL = wSQL & "         WHEN b.引当可能数量 <= 0 THEN -1"                                     '2010/03/18 an mod s
'wSQL = wSQL & "         WHEN b.引当可能数量 > 0 AND b.引当可能数量 <= a.希少数量 THEN 0"
'wSQL = wSQL & "         WHEN b.引当可能数量 > 0 AND b.引当可能数量 > a.希少数量 THEN 9999"
'wSQL = wSQL & "         ELSE 99999"
'wSQL = wSQL & "       END AS 在庫有無"                                                        '2010/03/18 an mod e
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN a.完売日 IS NULL AND a.取扱中止日 IS NULL THEN 1"
'wSQL = wSQL & "         ELSE 2"
'wSQL = wSQL & "       END AS 完売区分"
'wSQL = wSQL & "     , c.メーカー名"
'
'wSQL = wSQL & "     , (SELECT COUNT(*)"
'wSQL = wSQL & "          FROM 商品スペック s WITH (NOLOCK)"
'wSQL = wSQL & "         WHERE s.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "           AND s.商品コード = a.商品コード"
'wSQL = wSQL & "       ) AS 商品スペックCNT"
'
''色規格があるかどうか 2007/05/30
'wSQL = wSQL & "     , (SELECT COUNT(*)"
'wSQL = wSQL & "          FROM Web色規格別在庫 t WITH (NOLOCK)"
'wSQL = wSQL & "         WHERE t.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "           AND t.商品コード = a.商品コード"
'wSQL = wSQL & "           AND (t.色 != '' OR t.規格 != '')"
'wSQL = wSQL & "           AND t.終了日 IS NULL"
'wSQL = wSQL & "       ) AS 色規格CNT"
'
''---- 色規格の合計引当可能数量	2011/06/09 hn mod
'wSQL = wSQL & "     , ISNULL((SELECT SUM(引当可能数量)"
'wSQL = wSQL & "                 FROM Web色規格別在庫 u WITH (NOLOCK)"
'wSQL = wSQL & "                WHERE u.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "                  AND u.商品コード = a.商品コード"
'wSQL = wSQL & "                  AND u.終了日 IS NULL"
'wSQL = wSQL & "                  AND u.引当可能数量 > 0),0)"
'wSQL = wSQL & "       AS 色規格合計引当可能数量"
'
''---- 色規格の合計発注数量	2011/06/09 hn add
'wSQL = wSQL & "     , ISNULL((SELECT SUM(発注数量)"
'wSQL = wSQL & "                FROM Web色規格別在庫 w WITH (NOLOCK)"
'wSQL = wSQL & "               WHERE w.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "                 AND w.商品コード = a.商品コード"
'wSQL = wSQL & "                 AND w.終了日 IS NULL"
'wSQL = wSQL & "                 AND w.発注数量 > 0),0)"
'wSQL = wSQL & "       AS 色規格合計発注数量"
'
''レビュー評価の平均  '2010/03/18 an add
'wSQL = wSQL & "     , (SELECT CAST(AVG(CAST(ISNULL(v.評価,0) AS decimal(1,0))) AS decimal(2,1)) "   '2010/05/13 an changed
'wSQL = wSQL & "          FROM 商品レビュー v WITH (NOLOCK)"
'wSQL = wSQL & "         WHERE v.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "           AND v.商品コード = a.商品コード"
''''wSQL = wSQL & "           AND v.ショップコメント日 IS NULL"   '2010/05/13 an del
'wSQL = wSQL & "       ) AS レビュー評価平均"
'
''レビュー評価件数  '2010/07/12 st add
'wSQL = wSQL & "     , (SELECT COUNT(*) "
'wSQL = wSQL & "          FROM 商品レビュー w WITH (NOLOCK)"
'wSQL = wSQL & "         WHERE w.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "           AND w.商品コード = a.商品コード"
'wSQL = wSQL & "       ) AS レビュー数"
'
''---- FROM
'wSQL2 = ""
'wSQL2 = wSQL2 & "  FROM Web商品 a WITH (NOLOCK)"
'wSQL2 = wSQL2 & "     , Web色規格別在庫 b WITH (NOLOCK)"
'wSQL2 = wSQL2 & "     , メーカー c WITH (NOLOCK)"
'wSQL2 = wSQL2 & "     , カテゴリー d WITH (NOLOCK)"
'wSQL2 = wSQL2 & "     , 中カテゴリー e WITH (NOLOCK)"
'wSQL2 = wSQL2 & "     , 大カテゴリー h WITH (NOLOCK)"    '2010/06/08 an add
'wSQL2 = wSQL2 & "     , 商品カテゴリー f WITH (NOLOCK)"        '2005/07/19
'
''---- WHERE
'wSQL2 = wSQL2 & " WHERE a.Web商品フラグ = 'Y'"
'wSQL2 = wSQL2 & "   AND b.メーカーコード = a.メーカーコード"
'wSQL2 = wSQL2 & "   AND b.商品コード = a.商品コード"
'wSQL2 = wSQL2 & "   AND b.色 = ''"    '2007/05/30 add
'wSQL2 = wSQL2 & "   AND b.規格 = ''"    '2007/05/30 add
'wSQL2 = wSQL2 & "   AND c.メーカーコード = a.メーカーコード"
'
'wSQL2 = wSQL2 & "   AND d.カテゴリーコード = f.カテゴリーコード "        '2005/07/19
'wSQL2 = wSQL2 & "   AND e.中カテゴリーコード = d.中カテゴリーコード"        '2005/07/19
'wSQL2 = wSQL2 & "   AND f.メーカーコード = a.メーカーコード"        '2005/07/19
'wSQL2 = wSQL2 & "   AND f.商品コード = a.商品コード "        '2005/07/19
'wSQL2 = wSQL2 & "   AND h.大カテゴリーコード = e.大カテゴリーコード"   '2010/06/08 an add
'
''---- カテゴリーを指定して絞り込み
'if CategoryCd <> "" then   '2010/06/17 an mod
'    wSQL2 = wSQL2 & "  AND f.カテゴリーコード = '" & CategoryCd & "'"    '2005/07/19 '2010/06/17 an mod
'end if
'
''---- 中カテゴリーを指定して絞り込み
'if MidCategoryCd <> "" then                                                                          '2010/06/08 an add s
'    wSQL2 = wSQL2 & "   AND e.中カテゴリーコード = '" & MidCategoryCd & "'"
'end if
'
''---- 大カテゴリーを指定して絞り込み
'if LargeCategoryCd <> "" then
'    wSQL2 = wSQL2 & "   AND e.大カテゴリーコード = '" & LargeCategoryCd & "'"
'end if                                                                                                '2010/06/08 an add e
'
'wSQL2 = wSQL2 & "   AND ((a.個数限定数量 > a.個数限定受注済数量 AND a.個数限定数量 > 0) OR a.廃番日 IS NOT NULL)"
'
'if s_maker_cd <> "" then
'    wSQL2 = wSQL2 & "   AND c.メーカーコード = '" & s_maker_cd & "'"
'end if
'
'if trim(Request("sPriceFrom")) = "" AND Trim(Request("sPriceTo")) = "" then
'else
'    wSQL2 = wSQL2 & "   AND (a.販売単価 * (" & wSalesTaxRate & " + 100) / 100) BETWEEN " & sPriceFrom & " AND " & sPriceTo
'end if
'
'if s_product_cd <> "" then
'    if instr(s_product_cd, "%") > 0 then
'        wSQL2 = wSQL2 & "   AND a.商品コード LIKE '" & s_product_cd & "'"
'    else
'        wSQL2 = wSQL2 & "   AND a.商品コード = '" & s_product_cd & "'"
'    end if
'end if
'
'if sSeriesCd <> "" then
'    wSQL2 = wSQL2 & "   AND a.シリーズコード = '" & sSeriesCd & "'"
'end if
'
''---- ORDER BY  2010/03/18 an mod
'v_order = ""
'
'Select Case i_sort
'    Case "Price_ASC"
'    	v_order = v_order & " ORDER BY 実販売単価, c.メーカー名, a.商品名"
'    Case "Price_DESC"
'    	v_order = v_order & " ORDER BY 実販売単価 DESC, c.メーカー名, a.商品名"
'    Case "MakerName_ASC"
'    	v_order = v_order & " ORDER BY c.メーカー名, a.商品名"
'    Case "ProductName_ASC"
'    	v_order = v_order & " ORDER BY a.商品名"
'    Case "NewArrivals"
'    	v_order = v_order & " ORDER BY a.発売日 DESC, c.メーカー名, a.商品名"
'    Case "Reviews"
'    	v_order = v_order & " ORDER BY レビュー評価平均 DESC, c.メーカー名, a.商品名"
'    Case "ReviewCount"
'        v_order = " ORDER BY レビュー数 DESC, c.メーカー名, a.商品名"        '2010/07/12 st add
'    Case "Zaiko_DESC"
'    	v_order = v_order & " ORDER BY 完売区分, 在庫有無 DESC, c.メーカー名, a.商品名"
'    Case Else
'    	v_order = v_order & " ORDER BY お勧め商品表示順, c.メーカー名, a.商品名"
'End Select
'
''---- 該当商品一覧SQL
'wSQL = wSQL & wSQL2 & v_order
'
''---- 絞り込み用メーカーSQL
'wSQLMaker = "SELECT c.メーカー名, c.メーカーコード, COUNT(DISTINCT a.商品コード) AS 商品件数"
'wSQLMaker = wSQLMaker & wSQL2     '2010/06/08 an mod
'wSQLMaker = wSQLMaker & " GROUP BY c.メーカー名, c.メーカーコード"
'wSQLMaker = wSQLMaker & " ORDER BY 1"
'
''---- 絞り込み用カテゴリーSQL
'wSQLCategory = "SELECT e.中カテゴリー名日本語, d.カテゴリー名, d.カテゴリーコード, COUNT(DISTINCT a.商品コード) AS 商品件数"
'wSQLCategory = wSQLCategory & wSQL2
'wSQLCategory = wSQLCategory & " GROUP BY e.中カテゴリー名日本語, d.カテゴリー名, d.カテゴリーコード"
'wSQLCategory = wSQLCategory & " ORDER BY 1"
'
''---- 絞り込み用中カテゴリーSQL                                                                  '2010/06/08 an add s
'wSQLMidCategory = ""
'wSQLMidCategory = "SELECT h.大カテゴリー名, e.中カテゴリー名日本語, e.中カテゴリーコード, COUNT(DISTINCT a.商品コード) AS 商品件数"
'wSQLMidCategory = wSQLMidCategory & wSQL2
'wSQLMidCategory = wSQLMidCategory & " GROUP BY h.大カテゴリー名, e.中カテゴリー名日本語, e.中カテゴリーコード"
'wSQLMidCategory = wSQLMidCategory & " ORDER BY 1"
'
''---- 絞り込み用大カテゴリーSQL
'wSQLLargeCategory = ""
'wSQLLargeCategory = "SELECT h.大カテゴリー名, h.大カテゴリーコード, COUNT(DISTINCT a.商品コード) AS 商品件数"
'wSQLLargeCategory = wSQLLargeCategory & wSQL2
'wSQLLargeCategory = wSQLLargeCategory & " GROUP BY h.大カテゴリー名, h.大カテゴリーコード"
'wSQLLargeCategory = wSQLLargeCategory & " ORDER BY 1"                                            '2010/06/08 an add e
'
''---- 絞り込み用価格帯SQL
'wSQLPricerange = "SELECT MAX(a.販売単価) AS MAX販売単価, MIN(a.販売単価) AS MIN販売単価"
'wSQLPricerange = wSQLPricerange & wSQL2      '2010/06/08 an mod
'--- Webセール商品 テーブル対応 ここから ---
wSQL = wSQL & "SELECT DISTINCT "
wSQL = wSQL & "      a.メーカーコード "
wSQL = wSQL & "    , a.商品コード "
wSQL = wSQL & "    , a.商品名 "
wSQL = wSQL & "    , a.商品概略Web "
wSQL = wSQL & "    , a.送料区分 "
wSQL = wSQL & "    , a.特定商品個口 "
wSQL = wSQL & "    , a.重量商品送料 "
wSQL = wSQL & "    , a.商品画像ファイル名_小 "
wSQL = wSQL & "    , a.商品備考 "
wSQL = wSQL & "    , a.標準単価 "
wSQL = wSQL & "    , a.販売単価 "
wSQL = wSQL & "    , a.個数限定単価               AS 実販売単価 "
wSQL = wSQL & "    , a.実販売単価                 AS 個数限定単価 "
wSQL = wSQL & "    , a.個数限定数量 "
wSQL = wSQL & "    , a.個数限定受注済数量 "
wSQL = wSQL & "    , a.オープン価格フラグ "
wSQL = wSQL & "    , a.メーカー直送取寄区分 "
wSQL = wSQL & "    , a.ASK商品フラグ "
wSQL = wSQL & "    , a.取扱中止日 "
wSQL = wSQL & "    , a.廃番日 "
wSQL = wSQL & "    , a.終了日 "
wSQL = wSQL & "    , a.希少数量 "
wSQL = wSQL & "    , a.セット商品フラグ "
'wSQL = wSQL & "    , b.カテゴリーコード "
wSQL = wSQL & "    , a.カテゴリーコード "	'2011/01/14 na mod
wSQL = wSQL & "    , a.直輸入品フラグ "
wSQL = wSQL & "    , a.試聴フラグ "
wSQL = wSQL & "    , a.試聴URL "
wSQL = wSQL & "    , a.動画フラグ "
wSQL = wSQL & "    , a.動画URL "
wSQL = wSQL & "    , a.Web納期非表示フラグ "
wSQL = wSQL & "    , a.入荷予定未定フラグ "
wSQL = wSQL & "    , a.商品スペック使用不可フラグ "
wSQL = wSQL & "    , a.B品単価 "
wSQL = wSQL & "    , a.完売日 "
wSQL = wSQL & "    , a.発売日 "
wSQL = wSQL & "    , a.前回単価変更日 "
wSQL = wSQL & "    , a.前回販売単価 "
wSQL = wSQL & "    , a.B品フラグ "
wSQL = wSQL & "    , a.初回登録日 "
wSQL = wSQL & "    , a.送料完全無料商品フラグ "
wSQL = wSQL & "    , a.特価表示順                 AS お勧め商品表示順 "
wSQL = wSQL & "    , a.色 "
wSQL = wSQL & "    , a.規格 "
wSQL = wSQL & "    , a.引当可能数量 "
wSQL = wSQL & "    , a.発注数量 "
wSQL = wSQL & "    , a.引当可能入荷予定日 "
wSQL = wSQL & "    , a.B品引当可能数量 "
wSQL = wSQL & "    , a.商品ID "
wSQL = wSQL & "    , a.適正在庫数量 "
wSQL = wSQL & "    , a.在庫有無 "
wSQL = wSQL & "    , a.完売区分 "
wSQL = wSQL & "    , a.メーカー名 "
wSQL = wSQL & "    , a.商品スペックCNT "
wSQL = wSQL & "    , a.色規格CNT "
wSQL = wSQL & "    , a.色規格合計引当可能数量 "
wSQL = wSQL & "    , a.色規格合計発注数量 "
wSQL = wSQL & "    , a.レビュー評価平均 "
wSQL = wSQL & "    , a.レビュー数 "
wSQL = wSQL & "    , CASE "							'2012/10/22 ok Add
wSQL = wSQL & "        WHEN a.ASK商品フラグ != 'Y' THEN "
wSQL = wSQL & "         CASE "
wSQL = wSQL & "           WHEN a.個数限定数量 > a.個数限定受注済数量         AND (a.販売単価 - a.個数限定単価) > 0 THEN (a.販売単価 - a.個数限定単価) / a.販売単価 "
wSQL = wSQL & "           WHEN a.B品フラグ = 'Y'                             AND (a.販売単価 - a.B品単価) > 0 THEN (a.販売単価 - a.B品単価) / a.販売単価 "
wSQL = wSQL & "           WHEN DATEADD(d, 60, a.前回単価変更日) >= GETDATE() AND (a.前回販売単価 - a.販売単価) > 0 THEN (a.前回販売単価 - a.販売単価) / a.前回販売単価 "
wSQL = wSQL & "           ELSE 0 "
wSQL = wSQL & "         END "
wSQL = wSQL & "        ELSE 0 "
wSQL = wSQL & "      END AS 値下げ率 "

'--- FROM
wSQL2 = ""
wSQL2 = wSQL2 & "FROM "
wSQL2 = wSQL2 & "      Webセール商品       a WITH (NOLOCK) "

'wSQL2 = wSQL2 & "        LEFT JOIN Web商品 b WITH (NOLOCK) "				'2011/01/14 na del
'wSQL2 = wSQL2 & "          ON     b.メーカーコード = a.メーカーコード "
'wSQL2 = wSQL2 & "             AND b.商品コード     = a.商品コード "

'--- WHERE
wSQL2 = wSQL2 & "WHERE "
wSQL2 = wSQL2 & "       a.セール区分番号 = 3 "						' 数量限定 (BargainSalse) セール区分番号 : 3

'--- カテゴリーを指定して絞り込み
If CategoryCd <> "" Then
    wSQL2 = wSQL2 & "    AND a.カテゴリーコード = '" & CategoryCd & "' "
End If

'--- 中カテゴリーを指定して絞り込み
If MidCategoryCd <> "" Then
    wSQL2 = wSQL2 & "    AND a.中カテゴリーコード = '" & MidCategoryCd & "' "
End If

'--- 大カテゴリーを指定して絞り込み
If LargeCategoryCd <> "" Then
    wSQL2 = wSQL2 & "    AND a.大カテゴリーコード = '" & LargeCategoryCd & "' "
End If

If s_maker_cd <> "" Then
    wSQL2 = wSQL2 & "    AND a.メーカーコード = '" & s_maker_cd & "' "
End If

If Trim(Request("sPriceFrom")) = "" And Trim(Request("sPriceTo")) = "" Then
Else
'2014/03/19 GV mod start ---->
'価格帯の検索条件を税抜き
'    wSQL2 = wSQL2 & "    AND (a.販売単価 * (" & wSalesTaxRate & " + 100) / 100) BETWEEN " & sPriceFrom & " AND " & sPriceTo & " "
    wSQL2 = wSQL2 & "    AND a.販売単価 BETWEEN " & sPriceFrom & " AND " & sPriceTo & " "
'2014/03/19 GV mod end   <----
End If

If s_product_cd <> "" Then
    If Instr(s_product_cd, "%") > 0 Then
        wSQL2 = wSQL2 & "    AND a.商品コード LIKE '" & s_product_cd & "' "
    Else
        wSQL2 = wSQL2 & "    AND a.商品コード = '" & s_product_cd & "' "
    End If
End If

If sSeriesCd <> "" Then
    wSQL2 = wSQL2 & "    AND a.シリーズコード = '" & sSeriesCd & "'"
End If

Select Case i_sort
    Case "Price_ASC"
        v_order = " ORDER BY 実販売単価, a.メーカー名, a.商品名 "
    Case "Price_DESC"
        v_order = " ORDER BY 実販売単価 DESC, a.メーカー名, a.商品名 "
    Case "MakerName_ASC"
        v_order = " ORDER BY a.メーカー名, a.商品名 "
    Case "ProductName_ASC"
        v_order = " ORDER BY a.商品名 "
    Case "NewArrivals"
        v_order = " ORDER BY a.発売日 DESC, a.メーカー名, a.商品名 "
    Case "Reviews"
        v_order = " ORDER BY a.レビュー評価平均 DESC, a.メーカー名, a.商品名 "
    Case "ReviewCount"
        v_order = " ORDER BY a.レビュー数 DESC, a.メーカー名, a.商品名 "
    Case "Zaiko_DESC"
        v_order = " ORDER BY a.完売区分, a.在庫有無 DESC, a.メーカー名, a.商品名 "
	Case "Nesage_DESC"								'2012/10/22 ok Add
		v_order = " ORDER BY 値下げ率 DESC, a.メーカー名, a.商品名 "
    Case Else
        v_order = " ORDER BY a.特価表示順, a.メーカー名, a.商品名 "
End Select

'--- 該当商品一覧SQL
wSQL = wSQL & wSQL2 & v_order

'--- 絞り込み用メーカーSQL
wSQLMaker = "SELECT a.メーカー名, a.メーカーコード, COUNT(DISTINCT a.商品コード) AS 商品件数 " _
          & wSQL2 & " "_
          & "GROUP BY a.メーカー名, a.メーカーコード " _
          & "ORDER BY 1"

'--- 絞り込み用カテゴリーSQL	'2012/11/09 ok Mod
wSQLCategory = "SELECT a.中カテゴリー名日本語, a.カテゴリー名, a.カテゴリーコード, COUNT(DISTINCT a.商品コード) AS 商品件数 " _
             & wSQL2 & " " _
             & "GROUP BY a.中カテゴリー名日本語, a.カテゴリー名, a.カテゴリーコード "
'             & "ORDER BY 1"
wSQLCategory = "SELECT a.* FROM (" & wSQLCategory & ") a INNER JOIN カテゴリー b WITH (NOLOCK) ON " _
             & "a.カテゴリーコード = b.カテゴリーコード ORDER BY b.表示順 "

'--- 絞り込み用中カテゴリーSQL	'2010/06/08 an add s	'2012/11/09 ok Mod
wSQLMidCategory = "SELECT a.大カテゴリー名, a.中カテゴリー名日本語, a.中カテゴリーコード, COUNT(DISTINCT a.商品コード) AS 商品件数 " _
                & wSQL2 & " " _
                & "GROUP BY a.大カテゴリー名, a.中カテゴリー名日本語, a.中カテゴリーコード "
'                & "ORDER BY 1"
wSQLMidCategory = "SELECT a.* FROM (" & wSQLMidCategory & ") a INNER JOIN 中カテゴリー b WITH (NOLOCK) ON " _
                & "a.中カテゴリーコード = b.中カテゴリーコード ORDER BY b.表示順 "

'--- 絞り込み用大カテゴリーSQL	'2012/11/09 ok Mod
wSQLLargeCategory = "SELECT a.大カテゴリー名, a.大カテゴリーコード, COUNT(DISTINCT a.商品コード) AS 商品件数 " _
                  & wSQL2 & " " _
                  & "GROUP BY a.大カテゴリー名, a.大カテゴリーコード "
'                  & "ORDER BY 1"
wSQLLargeCategory = "SELECT a.* FROM (" & wSQLLargeCategory & ") a INNER JOIN 大カテゴリー b WITH (NOLOCK) ON " _
                  & "a.大カテゴリーコード = b.大カテゴリーコード ORDER BY b.表示順 "

'--- 絞り込み用価格帯SQL
wSQLPricerange = "SELECT MAX(a.販売単価) AS MAX販売単価, MIN(a.販売単価) AS MIN販売単価 " _
               & wSQL2
' 20120110 GV Mod End

'@@@@response.write("<br>s_maker_cd=" & s_maker_cd & "<br>s_category_cd=" & s_category_cd & "<br>" & w_where)
'@@@@response.write(wSQL)
'@@@@response.write("<br><br>" & wSQLMaker)
'@@@@response.write("<br><br>" & wSQLCategory)
'@@@@response.write("<br><br>" & wSQLMidCategory)
'@@@@response.write("<br><br>" & wSQLLargeCategory)

Set RS = Server.CreateObject("ADODB.Recordset")

RS.Open wSQL, Connection, adOpenStatic

End Function

'========================================================================
'
'    Function    メーカー情報作成
'
'========================================================================
'
Function CreateMakerInfo()

Dim RSv
Dim i

'---- メーカー情報取り出し
wSQL = ""
wSQL = wSQL & "SELECT a.メーカー名"
wSQL = wSQL & "     , a.メーカーホームページURL"
wSQL = wSQL & "     , a.メーカーロゴファイル名"
wSQL = wSQL & "     , a.メーカー紹介"
wSQL = wSQL & "  FROM メーカー a WITH (NOLOCK)"
wSQL = wSQL & " WHERE a.メーカーコード = '" & s_maker_cd & "'"

'@@@@@@@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

wHTML = ""
'---- メーカー情報
wHTML = wHTML & "    <div id='category_box'>" & vbNewLine
wHTML = wHTML & "      <p class='logo'>"

If RSv("メーカーロゴファイル名") <> "" Then
    wHTML = wHTML & "<img src='maker_img/" & RSv("メーカーロゴファイル名") & "' alt='" & RSv("メーカー名") & "'>"
end if
wHTML = wHTML & "</p>" & vbNewLine
wHTML = wHTML & "      <p class='txt'>" & Replace(RSv("メーカー紹介"), vbNewLine, "<br>") & "</p>" & vbNewLine
wHTML = wHTML & "    </div>" & vbNewLine

RSv.Close

'2012/07/11 ok Del Start
'---- メーカー売れ筋ランキング 取り出し
'wSQL = ""
'wSQL = wSQL & "SELECT TOP 5"
'wSQL = wSQL & "       a.メーカーコード"
'wSQL = wSQL & "     , a.商品コード"
'wSQL = wSQL & "     , b.商品名"
'wSQL = wSQL & "     , c.メーカー名"
'wSQL = wSQL & "  FROM "
'wSQL = wSQL & "       売筋商品 a WITH (NOLOCK)"
'wSQL = wSQL & "     , Web商品 b WITH (NOLOCK)"
'wSQL = wSQL & "     , メーカー c WITH (NOLOCK)"
'wSQL = wSQL & "     , カテゴリー d WITH (NOLOCK)"
'wSQL = wSQL & "     , Web色規格別在庫 g WITH (NOLOCK)"
'wSQL = wSQL & " WHERE "
'wSQL = wSQL & "       b.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND b.商品コード = a.商品コード"
'wSQL = wSQL & "   AND c.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND d.カテゴリーコード = a.カテゴリーコード"
'wSQL = wSQL & "   AND g.メーカーコード = b.メーカーコード"
'wSQL = wSQL & "   AND g.商品コード = b.商品コード"
'wSQL = wSQL & "   AND b.終了日 IS NULL"
'wSQL = wSQL & "   AND g.終了日 IS NULL"
'wSQL = wSQL & "   AND b.Web商品フラグ = 'Y'"
'wSQL = wSQL & "   AND d.売れ筋ランキング表示フラグ = 'Y'"
'wSQL = wSQL & "   AND a.メーカーコード = '" & s_maker_cd & "'"
'wSQL = wSQL & "   AND a.年月 = (SELECT MAX(年月) FROM 売筋商品)"
'wSQL = wSQL & " ORDER BY"
'wSQL = wSQL & "       a.受注数量 DESC"
'
''@@@@@@@@@@response.write(wSQL)
'
'Set RSv = Server.CreateObject("ADODB.Recordset")
'RSv.Open wSQL, Connection, adOpenStatic
'
'wHTML = wHTML & "    <td valign=top>" & vbNewLine
'
'if RSv.EOF = false then
'    wHTML = wHTML & "      <table width=215 border=0 cellspacing=0 cellpadding=0>" & vbNewLine
'    wHTML = wHTML & "        <tr>" & vbNewLine
'    wHTML = wHTML & "          <td align='center' nowrap='nowrap' bgcolor='#ccccff' class='honbun'>順位</td>" & vbNewLine
'    wHTML = wHTML & "          <td bgcolor='#eeeeee' class='honbun'><h3 style='font-size:100%;font-weight:normal;margin: 0px 0px 0px 0px'>" & RSv("メーカー名") & "</h3></td>" & vbNewLine
'    wHTML = wHTML & "        </tr>" & vbNewLine
'
'    i = 0
'    '----ランキング作成
'    Do until RSv.EOF = true
'        i = i + 1
'        wHTML = wHTML & "        <tr>" & vbNewLine
'        wHTML = wHTML & "          <td align='center' nowrap='nowrap' class='honbun'>" & i & ".</td>" & vbNewLine
'        wHTML = wHTML & "          <td><a href='ProductDetail.asp?item=" & Server.URLEncode(RSv("メーカーコード") & "^" & RSv("商品コード")) & "' class='link'>" & RSv("商品名") & "</a></td>" & vbNewLine
'        wHTML = wHTML & "        </tr>" & vbNewLine
'
'        RSv.MoveNext
'    Loop
'
'    RSv.Close
'
'    wHTML = wHTML & "      </table>" & vbNewLine
'end if
'
'wHTML = wHTML & "    </td>" & vbNewLine
'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & "</table>" & vbNewLine
'wHTML = wHTML & "<div style='padding: 10px 0px 0px 0px;'></div>" & vbNewLine
'2012/07/11 ok Del End

wMakerInfoHTML = wHTML

End Function

'========================================================================
'
'    Function    Close database
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
<title>衝撃特価品 <%=wTitle%> 一覧｜サウンドハウス</title>
<meta name="Description" content="サウンドハウスがおすすめするとってもお買い得な商品を衝撃特価でご提供。数量限定のためご注文はお早めに！">
<meta name="keywords" content="衝撃特価品,アウトレット,数量限定バーゲン,得割市場,プライスダウン情報">
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/shop.css?20121108" type="text/css">
<link rel="stylesheet" href="Style/searchlist.css?20121201" type="text/css">
<link rel="stylesheet" href="style/ask.css?20140401a" type="text/css">
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->

<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>
  <!-- コンテンツstart -->
  <div id="globalContents">

    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="../"><img src="../images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
<%=wFootprintHTML%>
      </ul>
    </div></div></div>

<!-- タブバナー部分の記述 START -->
    <div class="tab" id="bargainsale">
      <ul>
        <li><a href="SpecialPriceSale.asp"><img src="images/tab_bargainsalea.png" alt="得割市場"></a></li>
        <li><a href="Outlet.asp"><img src="images/tab_outlet.png" alt="わけあり市場"></a></li>
        <li><a href="PriceDownSale.asp"><img src="images/tab_pricedown.png" alt="プライスダウン"></a></li>
      </ul>
    </div>
<!-- タブバナー部分の記述 END -->

<!-- メーカー情報 (メーカーで検索時) -->
<%=wMakerInfoHTML%>

<!-- 商品一覧 -->
<% = wListHTML %>

  <!--/#contents --></div>

<!-- 絞込検索用Form -->
<form name="f_search" method="get" action="BargainSale.asp">
  <input type="hidden" name="s_maker_cd" value="<%=s_maker_cd%>">
  <input type="hidden" name="s_category_cd" value="<%=CategoryCd%>">
  <input type="hidden" name="s_mid_category_cd" value="<%=MidCategoryCd%>">
  <input type="hidden" name="s_large_category_cd" value="<%=LargeCategoryCd%>">
  <input type="hidden" name="s_product_cd" value="<%=s_product_cd%>">
  <input type="hidden" name="sSeriesCd" value="<%=sSeriesCd%>">
  <input type="hidden" name="sPriceFrom" value="<%=sPriceFrom%>">
  <input type="hidden" name="sPriceTo" value="<%=sPriceTo%>">
  <input type="hidden" name="i_page" value="1">
  <input type="hidden" name="i_sort" value="<%=i_sort%>">
  <input type="hidden" name="i_page_size" value="<%=i_page_size%>">
  <input type="hidden" name="i_ListType" value="<%=i_ListType%>">
</form>
      <div id="globalSide">
<%
'----NAVI用パラメータセット
NAVISearchListMakerListHTML = wNaviMakerHTML
NAVISearchListCategoryListHTML = wNaviCategoryHTML
NAVISearchListPriceRangeListHTML = wNaviPriceRangeHTML
NAVISearchListLargeCategoryListHTML = wNaviLargeCategoryHTML
NAVISearchListMidCategoryListHTML = wNaviMidCategoryHTML
%>
	<!--#include file="../Navi/NaviSideSale.inc"-->
	<!--#include file="../Navi/NaviSide.inc"-->
	<!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/Navibottom.inc"-->
<div class="tooltip"><p>ASK</p></div>
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript" src="jslib/ask.js?20140401a"></script>
<script type="text/javascript" src="jslib/SearchList.js?20121108" charset="Shift_JIS"></script>
<script type="text/javascript" src="../jslib/jquery.tinyscrollbar.min.js"></script>
<script type="text/javascript">
$(function(){
    $('#scrollbar1').tinyscrollbar();
});
<% if wNoData <> "Y" then%>

    preset_values();

<% end if %>


//
//    Search onClick
//
function Search_onClick(pMakerCd, pCategoryCd, pMidCategoryCd, pLargeCategoryCd, pPriceFrom, pPriceTo){
    document.f_search.s_maker_cd.value = pMakerCd;
    document.f_search.s_category_cd.value = pCategoryCd;
    document.f_search.s_mid_category_cd.value = pMidCategoryCd;
    document.f_search.s_large_category_cd.value = pLargeCategoryCd;
    document.f_search.sPriceFrom.value = pPriceFrom;
    document.f_search.sPriceTo.value = pPriceTo;
    document.f_search.submit();
}
</script>
</body>
</html>
