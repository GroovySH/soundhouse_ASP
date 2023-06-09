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
'	アクセスランキング　(商品ビュー、友達に勧める、ウィッシュリスト)
'
'	更新履歴
'2007/10/16 前月ランキング表示に変更
'2009/04/30 エラー時にerror.aspへ移動
'2010/02/18 an ASK商品パラメータにServer.URLEncodeを行なう
'2010/05/10 an リニューアル対応（カートボタン、在庫情報、商品ID表示追加）
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/01/19 GV データ取得 SELECT文へ LACクエリー案を適用
'2012/01/19 GV NOLOCKオプション付与漏れ対応
'2012/01/23 GV 「商品レビュー」テーブルから「商品レビュー集計」テーブル使用に変更 (CreateReviewImg()プロシージャ)
'2012/08/07 if-web リニューアルレイアウト調整
'2014/03/19 GV 消費税増税に伴う2重表示対応
'
'========================================================================

On Error Resume Next

Dim RankType

Dim wSalesTaxRate
Dim wYYYYMM

Dim wRank
Dim wItem
Dim wPrice
Dim wPriceNoTax		'2014/03/19 GV add


Dim wProdTermFl '販売終了商品フラグ
Dim wInventoryCd
Dim wInventoryImage

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
Dim wTop3HTML
Dim wUnder4HTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add

'========================================================================

'---- 送信データーの取り出し
RankType = ReplaceInput(Request("RankType"))

if RankType = "" then
	RankType = "商品ビュー"
end if

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "RankingAccess.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then
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

Dim i
Dim vPrevMonth

'---- 消費税率取出し
call getCntlMst("共通","消費税率","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'消費税率
wSalesTaxRate = Clng(wItemNum1)

'---- 前月
vPrevMonth = DateAdd("m", -1, Date())
wYYYYMM = Year(vPrevMonth) &  Right("0" & Month(vPrevMonth), 2)

'---- ランキング取り出し
wSQL = ""
' 2012/01/19 GV Mod Start
'wSQL = wSQL & "SELECT TOP 20"
'wSQL = wSQL & "       a.メーカーコード"
'wSQL = wSQL & "     , a.商品コード"
'wSQL = wSQL & "     , b.商品名"
'wSQL = wSQL & "     , b.商品画像ファイル名_小"
'wSQL = wSQL & "     , b.お勧め商品コメント"
'wSQL = wSQL & "     , b.商品概略Web"
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
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN (b.個数限定数量 > b.個数限定受注済数量 AND b.個数限定数量 > 0) THEN b.個数限定単価"
'wSQL = wSQL & "         ELSE b.販売単価"
'wSQL = wSQL & "       END AS 販売単価"
'wSQL = wSQL & "     , c.メーカー名"
'wSQL = wSQL & "     , d.カテゴリーコード"
'wSQL = wSQL & "     , d.カテゴリー名"
'wSQL = wSQL & "     , e.色"
'wSQL = wSQL & "     , e.規格"
'wSQL = wSQL & "     , e.B品引当可能数量"
'wSQL = wSQL & "     , e.引当可能数量"
'wSQL = wSQL & "     , e.引当可能入荷予定日"
'wSQL = wSQL & "     , e.商品ID"
'
''色規格があるかどうか 2007/05/30
'wSQL = wSQL & "     , (SELECT COUNT(*)"
'wSQL = wSQL & "          FROM Web色規格別在庫 f WITH (NOLOCK)"
'wSQL = wSQL & "         WHERE f.メーカーコード = b.メーカーコード"
'wSQL = wSQL & "           AND f.商品コード = b.商品コード"
'wSQL = wSQL & "           AND (f.色 != '' OR f.規格 != '')"
'wSQL = wSQL & "           AND f.終了日 IS NULL"
'wSQL = wSQL & "       ) AS 色規格CNT"
'
'wSQL = wSQL & "  FROM 商品アクセス件数 a WITH (NOLOCK)"
'wSQL = wSQL & "     , Web商品 b WITH (NOLOCK)"
'wSQL = wSQL & "     , メーカー c WITH (NOLOCK)"
'wSQL = wSQL & "     , カテゴリー d WITH (NOLOCK)"
'wSQL = wSQL & "     , Web色規格別在庫 e WITH (NOLOCK)"
'
'wSQL = wSQL & " WHERE b.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND b.商品コード = a.商品コード"
'wSQL = wSQL & "   AND c.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND d.カテゴリーコード = b.カテゴリーコード"
'wSQL = wSQL & "   AND e.メーカーコード = b.メーカーコード"
'wSQL = wSQL & "   AND e.商品コード = b.商品コード"
'wSQL = wSQL & "   AND b.Web商品フラグ = 'Y'"
'wSQL = wSQL & "   AND e.色 = ''"
'wSQL = wSQL & "   AND e.規格 = ''"
'wSQL = wSQL & "   AND a.年月 = '" & wYYYYMM & "'"
'
'wSQL = wSQL & " ORDER BY"
'
'if RankType = "商品ビュー" then
'	wSQL = wSQL & "       a.ページビュー件数 DESC"
'end if
'if RankType = "友達にお勧め" then
'	wSQL = wSQL & "       a.友達にお勧め件数 DESC"
'end if
'if RankType = "欲しいものリスト" then
'	wSQL = wSQL & "       a.ウィッシュリスト件数 DESC"
'end if
'
'wSQL = wSQL & "     , c.メーカー名"
'wSQL = wSQL & "     , b.商品名"
wSQL = wSQL & "SELECT TOP 20 "
wSQL = wSQL & "      a.メーカーコード "
wSQL = wSQL & "    , a.商品コード "
wSQL = wSQL & "    , b.商品名 "
wSQL = wSQL & "    , b.商品画像ファイル名_小 "
wSQL = wSQL & "    , b.お勧め商品コメント "
wSQL = wSQL & "    , b.商品概略Web "
wSQL = wSQL & "    , b.取扱中止日 "
wSQL = wSQL & "    , b.完売日 "
wSQL = wSQL & "    , b.廃番日 "
wSQL = wSQL & "    , b.B品フラグ "
wSQL = wSQL & "    , b.ASK商品フラグ "
wSQL = wSQL & "    , b.希少数量 "
wSQL = wSQL & "    , b.個数限定数量 "
wSQL = wSQL & "    , b.個数限定受注済数量 "
wSQL = wSQL & "    , b.セット商品フラグ "
wSQL = wSQL & "    , b.メーカー直送取寄区分 "
wSQL = wSQL & "    , b.入荷予定未定フラグ "
wSQL = wSQL & "    , b.Web納期非表示フラグ "
wSQL = wSQL & "    , CASE "
wSQL = wSQL & "        WHEN (b.個数限定数量 > b.個数限定受注済数量 AND b.個数限定数量 > 0) THEN b.個数限定単価 "
wSQL = wSQL & "        ELSE b.販売単価 "
wSQL = wSQL & "      END AS 販売単価 "
wSQL = wSQL & "    , c.メーカー名 "
wSQL = wSQL & "    , d.カテゴリーコード "
wSQL = wSQL & "    , d.カテゴリー名 "
wSQL = wSQL & "    , e.色 "
wSQL = wSQL & "    , e.規格 "
wSQL = wSQL & "    , e.B品引当可能数量 "
wSQL = wSQL & "    , e.引当可能数量 "
wSQL = wSQL & "    , e.引当可能入荷予定日 "
wSQL = wSQL & "    , e.商品ID "
wSQL = wSQL & "    , (SELECT COUNT(f.商品コード) "
wSQL = wSQL & "         FROM Web色規格別在庫 f WITH (NOLOCK) "
wSQL = wSQL & "        WHERE     f.メーカーコード = b.メーカーコード "
wSQL = wSQL & "              AND f.商品コード = b.商品コード "
wSQL = wSQL & "              AND (   f.色   != '' "
wSQL = wSQL & "                   OR f.規格 != '') "
wSQL = wSQL & "              AND f.終了日 IS NULL "
wSQL = wSQL & "      ) AS 色規格CNT "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    商品アクセス件数             a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN Web商品         b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.メーカーコード   = a.メーカーコード "
wSQL = wSQL & "           AND b.商品コード       = a.商品コード "
wSQL = wSQL & "      INNER JOIN メーカー        c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.メーカーコード   = a.メーカーコード "
wSQL = wSQL & "      INNER JOIN カテゴリー      d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.カテゴリーコード = b.カテゴリーコード "
wSQL = wSQL & "      INNER JOIN Web色規格別在庫 e WITH (NOLOCK) "
wSQL = wSQL & "        ON     e.メーカーコード   = b.メーカーコード "
wSQL = wSQL & "           AND e.商品コード       = b.商品コード "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' ) t1 "
wSQL = wSQL & "        ON     b.Web商品フラグ    = t1.ShohinWebY "
wSQL = wSQL & "      LEFT JOIN ( SELECT ''  AS 'Iro' )        t2 "
wSQL = wSQL & "        ON     e.色               = t2.Iro "
wSQL = wSQL & "      LEFT JOIN ( SELECT ''  AS 'Kikaku' )     t3 "
wSQL = wSQL & "        ON     e.規格             = t3.Kikaku "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        t1.ShohinWebY   IS NOT NULL "
wSQL = wSQL & "    AND t2.Iro          IS NOT NULL "
wSQL = wSQL & "    AND t3.Kikaku       IS NOT NULL "
wSQL = wSQL & "    AND a.年月 = '" & wYYYYMM & "' "
wSQL = wSQL & "ORDER BY "
If RankType     = "商品ビュー" Then
	wSQL = wSQL & "      a.ページビュー件数 DESC "
ElseIf RankType = "友達にお勧め" Then
	wSQL = wSQL & "      a.友達にお勧め件数 DESC "
ElseIf RankType = "欲しいものリスト" Then
	wSQL = wSQL & "      a.ウィッシュリスト件数 DESC "
End If
wSQL = wSQL & "    , c.メーカー名 "
wSQL = wSQL & "    , b.商品名 "
' 2012/01/19 GV Mod End

'@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

if RS.EOF = true then
	exit function
end if

wRank = 0
wHTML = ""
wTop3HTML = ""
wUnder4HTML = ""

Do Until RS.EOF = true
	wRank = wRank + 1
	wItem = Server.URLEncode(RS("メーカーコード") & "^" & RS("商品コード") & "^" & "^")
	wPrice = calcPrice(RS("販売単価"), wSalesTaxRate)
	wPriceNoTax = RS("販売単価")							'2014/03/19 GV add

	'---- 在庫状況表示のため、終了チェック
	wProdTermFl = "N"
	if isNull(RS("取扱中止日")) = false then		'取扱中止
		wProdTermFl = "Y"
	end if
	if isNull(RS("廃番日")) = false AND RS("引当可能数量") <= 0 then		'廃番で在庫無し
		wProdTermFl = "Y"
	end if
	if isNull(RS("完売日")) = false then		'完売商品
		wProdTermFl = "Y"
	end if
	if RS("B品フラグ") = "Y" AND RS("B品引当可能数量") <= 0 then    'B品で在庫なし
		wProdTermFl = "Y"
	end if

	'---- 在庫状況作成
	if RS("色規格CNT") = 0 then
		wInventoryCd = GetInventoryStatus(RS("メーカーコード"),RS("商品コード"),RS("色"),RS("規格"),RS("引当可能数量"),RS("希少数量"),RS("セット商品フラグ"),RS("メーカー直送取寄区分"),RS("引当可能入荷予定日"),wProdTermFl)

		'---- 在庫状況、色を最終セット
		call GetInventoryStatus2(RS("引当可能数量"), RS("Web納期非表示フラグ"), RS("入荷予定未定フラグ"), RS("廃番日"), RS("B品フラグ"), RS("B品引当可能数量"), RS("個数限定数量"), RS("個数限定受注済数量"), wProdTermFl, wInventoryCd, wInventoryImage)

	end if

	'---- 1～3位
	if wRank <= 3 then
		call CreateTop3ItemHTML()
	'---- 4～20位
	else
		call CreateUnder4ItemHTML()
	end if

	RS.MoveNext
Loop

RS.Close

End function


'========================================================================
'
'	Function	ランキング上位3商品の表示
'
'========================================================================

Function CreateTop3ItemHTML()

Dim vBGColor

'---- 偶数行と奇数行で表示色を変える
if wRank Mod 2 <> 0 then
	vBGColor = "bg_color1"   '奇数行
else
	vBGColor = "bg_color2"   '偶数行
end if

wHTML = wHTML & "<!-- bigbox1 START -->" & vbNewLine
wHTML = wHTML & "    <div class='rankingaccess_bigbox'>" & vbNewLine
wHTML = wHTML & "      <div class='left " & vBGColor & "'>" & vbNewLine
wHTML = wHTML & "        <div class='crown_box'><img src='images/ranking/ico_no" & wRank & "crown.gif' alt='' width='41' height='30'></div>" & vbNewLine
wHTML = wHTML & "      </div>" & vbNewLine
wHTML = wHTML & "      <div class='centerleft'>" & vbNewLine
wHTML = wHTML & "        <a href='ProductDetail.asp?Item=" & wItem & "'>" & vbNewLine
if RS("商品画像ファイル名_小") <> "" then
	wHTML = wHTML & "          <img src='prod_img/" & RS("商品画像ファイル名_小") & "' alt='" & RS("メーカー名") & " " & RS("商品名") & "'>" & vbNewLine
else
	wHTML = wHTML & "          <img src='prod_img/n/nopict.jpg' alt='" & RS("メーカー名") & " " & RS("商品名") & "'>" & vbNewLine
end if
wHTML = wHTML & "        </a>" & vbNewLine
wHTML = wHTML & "      </div>" & vbNewLine
wHTML = wHTML & "      <div class='center'>" & vbNewLine
wHTML = wHTML & "        <h2>" & vbNewLine
wHTML = wHTML & "          <a href='ProductDetail.asp?Item=" & wItem & "'>" & vbNewLine
wHTML = wHTML & "            <span class='txt_maker'>" & RS("メーカー名") & "</span>&nbsp;" & vbNewLine
wHTML = wHTML & "            <span class='txt_product'>" & RS("商品名") & "</span>" & vbNewLine
wHTML = wHTML & "          </a>" & vbNewLine
wHTML = wHTML & "        </h2>" & vbNewLine
wHTML = wHTML & "        <h3>" & vbNewLine
wHTML = wHTML & "          <a href='SearchList.asp?i_type=c&s_category_cd=" & RS("カテゴリーコード") & "'>" & RS("カテゴリー名") & "</a>" & vbNewLine
wHTML = wHTML & "        </h3>" & vbNewLine
wHTML = wHTML & "        <div class='bg'>" & vbNewLine
wHTML = wHTML & "          <div class='price_box'>" & vbNewLine
wHTML = wHTML & "            衝撃特価： "

if RS("ASK商品フラグ") = "Y" then
	wHTML = wHTML & "ASK" & vbNewLine
else
'2014/03/19 GV mod start ---->
'	wHTML = wHTML & "<strong>" & FormatNumber(wPrice,0) & "円(税込)</strong>　" & vbNewLine
	wHTML = wHTML & "<strong>" & FormatNumber(wPriceNoTax,0) & "円(税抜)</strong>　" & vbNewLine
	wHTML = wHTML & "(税込&nbsp;" & FormatNumber(wPrice,0) & "円)　" & vbNewLine
'2014/03/19 GV mod end   <----
end if

wHTML = wHTML & "          </div>" & vbNewLine
wHTML = wHTML & "          <div class='notes_box'>" & vbNewLine

'----- 商品レビュー
wHTML = wHTML & "            " & CreateReviewImg(RS("メーカーコード"), RS("商品コード"))

wHTML = wHTML & "          </div>" & vbNewLine
wHTML = wHTML & "        </div>" & vbNewLine

'---- 商品説明
wHTML = wHTML & "        <p>"
if RS("お勧め商品コメント") <> "" then
	wHTML = wHTML & Replace(RS("お勧め商品コメント"), vbNewLine, "<br>")
else
	wHTML = wHTML & Replace(RS("商品概略Web"), vbNewLine, "<br>")
end if
wHTML = wHTML & "        </p>" & vbNewLine

wHTML = wHTML & "      </div>" & vbNewLine
wHTML = wHTML & "      <div class='rankingaccess_shopbox'>" & vbNewLine
wHTML = wHTML & "        <form name='f_item' method='post' action='OrderPreInsert.asp' onSubmit='return order_onClick(this);'>" & vbNewLine
'wHTML = wHTML & "        <div class='right'>" & vbNewLine

'---- 色規格なし
if RS("色規格CNT") = 0 then
	if wProdTermFl = "Y" then
		wHTML = wHTML & "            <img src='images/icon_sold.gif'><br>" & vbNewLine
	else
		wHTML = wHTML & "            <input type='hidden' name='qt' value='1'>" & vbNewLine
		wHTML = wHTML & "            <input type='hidden' name='maker_cd' value='" & RS("メーカーコード") & "'>" & vbNewLine
		wHTML = wHTML & "            <input type='hidden' name='product_cd' value='" & RS("商品コード") & "'>" & vbNewLine
		wHTML = wHTML & "            <input type='hidden' name='category_cd' value='" & RS("カテゴリーコード") & "'>" & vbNewLine
		wHTML = wHTML & "            <input type='image' src='images/btn_cart.png' style='vertical-align:middle' alt='カートへ' class='opover'><br>" & vbNewLine
	end if

'----色規格あり
else
	if wProdTermFl = "Y" then
		wHTML = wHTML & "            <img src='images/icon_sold.gif'><br>" & vbNewLine
	else
		wHTML = wHTML & "            <input type='hidden' name='qt' value='0'>" & vbNewLine
		wHTML = wHTML & "            <a href='ProductDetail.asp?Item=" & wItem & "'><img src='images/btn_detail.png'></a><br>" & vbNewLine
	end if
end if

'wHTML = wHTML & "        </div>" & vbNewLine
wHTML = wHTML & "        <div class='shopid'>商品ID:" & RS("商品ID") & "</div>" & vbNewLine
wHTML = wHTML & "        </form>" & vbNewLine
wHTML = wHTML & "        <div class='itemstock'><img src='images/" & wInventoryImage & "' alt=''> " & wInventoryCd & "</div>" & vbNewLine
wHTML = wHTML & "      </div>" & vbNewLine
wHTML = wHTML & "    </div>" & vbNewLine
wHTML = wHTML & "<!-- bigbox END -->" & vbNewLine

wTop3HTML = wHTML

End function

'========================================================================
'
'	Function	ランキング4位以下の商品の表示
'
'========================================================================

Function CreateUnder4ItemHTML()

Dim vBGColor
Dim vMakerProduct
Dim vProductName

wHTML = ""

'---- 4位の手前に項目名の行を表示
if wRank = 4 then
	wHTML = wHTML & "    <!-- s_box TH START -->" & vbNewLine
	wHTML = wHTML & "    <div id='s_box_th_ra'>" & vbNewLine
	wHTML = wHTML & "      <div id='th_no'>順位</div>" & vbNewLine
	wHTML = wHTML & "      <div id='th_prod'><div class='cell'>メーカー　商品</div></div>" & vbNewLine
	wHTML = wHTML & "      <div id='th_cat'>カテゴリー</div>" & vbNewLine
	wHTML = wHTML & "      <div id='th_point'>レビューポイント</div>" & vbNewLine
	wHTML = wHTML & "      <div id='th_stock'>在庫状況</div>" & vbNewLine
	wHTML = wHTML & "      <div id='th_cart'>カート</div>" & vbNewLine
	wHTML = wHTML & "    </div>" & vbNewLine
	wHTML = wHTML & "    <!-- s_box TH END -->" & vbNewLine
end if

'---- 偶数行と奇数行で表示色を変える
if wRank Mod 2 <> 0 then
	vBGColor = "s_box1"
else
	vBGColor = "s_box2"
end if

wHTML = wHTML & "    <!-- s_box START -->" & vbNewLine
wHTML = wHTML & "    <div class='" & vBGColor & "'>" & vbNewLine
wHTML = wHTML & "      <div class='s_box_height'>" & vbNewLine
wHTML = wHTML & "        <div class='num_box'>" & wRank & "</div>" & vbNewLine
wHTML = wHTML & "        <div class='text_box'>" & vbNewLine
wHTML = wHTML & "          <a href='ProductDetail.asp?Item=" & wItem & "'>" & vbNewLine

'--- メーカー名＋商品名が長くて2行になる場合は"..."で省略
vMakerProduct = RS("メーカー名") & " " & RS("商品名")
if Len(vMakerProduct) > 33 then

	vProductName = Left(RS("商品名"), 30-Len(RS("メーカー名"))) &  "..."
else
	vProductName = RS("商品名")
end if

wHTML = wHTML & "            <span class='txt_maker'>" & RS("メーカー名") & "</span>&nbsp;" & vbNewLine
wHTML = wHTML & "            <span class='txt_product'>" & vProductName & "</span><br>" & vbNewLine
wHTML = wHTML & "            <span class='txt_price_h'>衝撃特価：</span>" & vbNewLine

if RS("ASK商品フラグ") = "Y" then
	wHTML = wHTML & "ASK"
else
'2014/03/19 GV mod start ---->
'	wHTML = wHTML & "            <span class='txt_price_d'>" & FormatNumber(wPrice,0) & "円(税込)</span>" & vbNewLine
	wHTML = wHTML & "            <span class='txt_price_d'>" & FormatNumber(wPriceNoTax,0) & "円(税抜)</span>　"
	wHTML = wHTML & "<span class='txt_price_t'>(税込&nbsp;"&FormatNumber(wPrice,0) & "円)</span>" & vbNewLine
'2014/03/19 GV mod end   <----
end if
wHTML = wHTML & "          </a>" & vbNewLine
wHTML = wHTML & "        </div>" & vbNewLine
wHTML = wHTML & "        <div class='cat_box'>" & vbNewLine
wHTML = wHTML & "          <a href='SearchList.asp?i_type=c&s_category_cd=" & RS("カテゴリーコード")  & "'>" & RS("カテゴリー名") & "</a>" & vbNewLine
wHTML = wHTML & "        </div>" & vbNewLine

'----- 商品レビュー
wHTML = wHTML & "        <div class='note_box'>" & vbNewLine
wHTML = wHTML & "          <div class='pt8'>" & vbNewLine
wHTML = wHTML & "            " & CreateReviewImg(RS("メーカーコード"), RS("商品コード")) & vbNewLine
wHTML = wHTML & "          </div>" & vbNewLine
wHTML = wHTML & "        </div>" & vbNewLine

'----- 在庫表示
wHTML = wHTML & "        <div class='stock_box'>" & vbNewLine
wHTML = wHTML & "          <div class='pt12'>" & vbNewLine
wHTML = wHTML & "            <img height='10' src='images/" & wInventoryImage & "' width='10' alt=''> " & wInventoryCd & vbNewLine
wHTML = wHTML & "          </div>" & vbNewLine
wHTML = wHTML & "        </div>" & vbNewLine

'---- カート
wHTML = wHTML & "        <div class='cart_box'>" & vbNewLine
wHTML = wHTML & "        <form name='f_item' method='post' action='OrderPreInsert.asp' onSubmit='return order_onClick(this);'>" & vbNewLine

'---- 色規格なし
if RS("色規格CNT") = 0 then
	if wProdTermFl = "Y" then
		wHTML = wHTML & "            <img src='images/icon_sold.gif' alt='完売'><br>" & vbNewLine
	else
		wHTML = wHTML & "            <input type='hidden' name='qt' value='1'>" & vbNewLine
		wHTML = wHTML & "            <input type='hidden' name='maker_cd' value='" & RS("メーカーコード") & "'>" & vbNewLine
		wHTML = wHTML & "            <input type='hidden' name='product_cd' value='" & RS("商品コード") & "'>" & vbNewLine
		wHTML = wHTML & "            <input type='hidden' name='category_cd' value='" & RS("カテゴリーコード") & "'>" & vbNewLine
		wHTML = wHTML & "            <input type='image' src='images/btn_cart.png' style='vertical-align:middle' alt='カートへ' class='opover'><br>" & vbNewLine
	end if

'----色規格あり
else
	if wProdTermFl = "Y" then
		wHTML = wHTML & "            <img src='images/icon_sold.gif' alt='完売'><br>" & vbNewLine
	else
		wHTML = wHTML & "            <input type='hidden' name='qt' value='0'>" & vbNewLine
		wHTML = wHTML & "            <a href='ProductDetail.asp?Item=" & wItem & "'><img src='images/btn_detail.png' alt='詳細を見る'></a><br>" & vbNewLine
	end if
end if

wHTML = wHTML & "          商品ID:" & RS("商品ID") & vbNewLine
wHTML = wHTML & "        </form>" & vbNewLine
wHTML = wHTML & "        </div>" & vbNewLine
wHTML = wHTML & "      </div>" & vbNewLine
wHTML = wHTML & "    </div>" & vbNewLine
wHTML = wHTML & "<!-- s_box END -->" & vbNewLine

wUnder4HTML = wUnder4HTML & wHTML

End function

'========================================================================
'
'	Function	商品レビュー画像作成
'
'========================================================================
'
Function CreateReviewImg(pMakerCd, pProductCd)

Dim vAvgRating
Dim v1Cnt
Dim v0Cnt
Dim vHalfCnt
Dim vTotalCnt
Dim vReview
Dim RSv
Dim i

'---- Select 商品レビュー 平均，件数 取得
' 2012/01/23 GV Mod Start
'wSQL = ""
'wSQL = wSQL & "SELECT SUM(a.評価) AS 評価合計"
'wSQL = wSQL & "     , COUNT(a.ID) AS レビュー数"
'wSQL = wSQL & "  FROM 商品レビュー a WITH (NOLOCK) "				' 2012/01/19 GV Mod (NOLOCK オプション付与)
'wSQL = wSQL & " WHERE a.メーカーコード = '" & pMakerCd & "'"
'wSQL = wSQL & "   AND a.商品コード = '" & pProductCd & "'"
'
''@@@@response.write(wSQL)
'
'Set RSv = Server.CreateObject("ADODB.Recordset")
'RSv.Open wSQL, Connection, adOpenStatic
'
'if RSv("レビュー数") = 0 then
'	CreateReviewImg = ""
'	exit function
'end if
'
'vAvgRating = Round(RSv("評価合計")/RSv("レビュー数"), 1)
'v1Cnt = Fix(vAvgRating)
'if (vAvgRating - v1Cnt) >= 0.5 then
'	vHalfCnt = 1
'else
'	vHalfCnt = 0
'end if
'v0Cnt = 5 - v1Cnt - vHalfCnt
'
'vTotalCnt = RSv("レビュー数")
'Rsv.Close

CreateReviewImg = ""

'---- Select 商品レビュー 平均，件数 取得
wSQL = ""
wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.レビュー評価平均 "
wSQL = wSQL & "    , a.レビュー件数 "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    商品レビュー集計 a WITH (NOLOCK) "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        a.メーカーコード = '" & pMakerCd & "' "
wSQL = wSQL & "    AND a.商品コード     = '" & Replace(pProductCd, "'", "''") & "' "

'@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

If RSv.EOF Then
	RSv.Close
	Set RSv = Nothing
	Exit Function
End If

vAvgRating = Round(RSv("レビュー評価平均"), 1)
vTotalCnt = RSv("レビュー件数")

Rsv.Close
Set RSv = Nothing

If vTotalCnt = 0 Then
	Exit Function
End If

v1Cnt = Fix(vAvgRating)
If (vAvgRating - v1Cnt) >= 0.5 Then
	vHalfCnt = 1
Else
	vHalfCnt = 0
End If
v0Cnt = 5 - v1Cnt - vHalfCnt
' 2012/01/23 GV Mod End

'--- 総合評価編集
For i = 1 to v1Cnt
	vReview = vReview & "<img src='images/review_icon10.png' alt=''>"
Next
If vHalfcnt = 1 Then
	vReview = vReview & "<img src='images/review_icon05.png' alt=''>"
End If
For i=1 to v0Cnt
	vReview = vReview & "<img src='images/review_icon00.png' alt=''>"
Next

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
<title>ランキング｜サウンドハウス</title>
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
      <li class="now">検索キーワード</li>
    </ul>
  </div></div></div>
  <h1 class="title">検索キーワード</h1>
-->

<!-- Mainpage START -->
<div id="ranking_key_main_flame">
  <div id="shukei">（集計：<%=Left(wYYYYMM,4)%>年<%=right(wYYYYMM,2)%>月）</div>
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
        <img src="images/ranking/<% if RankType="商品ビュー" then%>noc_btn_on.jpg<% else %>noc_btn_off.jpg<% end if%>" alt="" name="Image12" width="114" height="80" />
      </a>
    </div>
    -->
    <div class="top_menu_parts">
      <a href="RankingAccess.asp?RankType=<%=Server.URLEncode("友達にお勧め")%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image13','','images/ranking/rtaf_btn_on.jpg',1)"><img src="images/ranking/<% if RankType="友達にお勧め" then%>rtaf_btn_on.jpg<% else %>rtaf_btn_off.jpg<% end if%>" alt="" name="Image13" width="114" height="80" />
      </a>
    </div>
    <div class="top_menu_parts">
      <a href="RankingAccess.asp?RankType=<%=Server.URLEncode("欲しいものリスト")%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image14','','images/ranking/wl_btn_on.jpg',1)"><img src="images/ranking/<% if RankType="欲しいものリスト" then%>wl_btn_on.jpg<% else %>wl_btn_off.jpg<% end if%>" alt="" name="Image14" width="114" height="80" />
      </a>
    </div>
    <div class="top_menu_parts">
      <a href="RankingReview.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image16','','images/ranking/nor_btn_on.jpg',1)"><img src="images/ranking/nor_btn_off.jpg" alt="" name="Image16" width="114" height="80" />
      </a>
    </div>
    <div class="top_menu_parts">
      <a href="RankingReviewPoint.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image17','','images/ranking/rr_btn_on.jpg',1)"><img src="images/ranking/rr_btn_off.jpg" alt="" name="Image17" width="113" height="80" /></a>
    </div>
  </div>

<!-- Menu END -->
<!--  container START  -->
  <div id="container">

<%=wTop3HTML%>
<%=wUnder4HTML%>

  </div>
<!-- container END -->
</div>

  </div>
  <div id="globalSide">
    <!--#include file="../Navi/NaviSide.inc"-->
  </div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>