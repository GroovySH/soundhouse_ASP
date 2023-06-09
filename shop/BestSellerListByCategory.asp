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
'	ベストセラー商品(カテゴリー別)
'
'	更新履歴
'2007/04/13 販売単価に個数限定単価を考慮
'2007/05/08 ハッカーセーフ対応
'2007/05/25 シリーズ対応
'2007/05/29 順位表示を1位、2位。。。に変更
'2009/04/30 エラー時にerror.aspへ移動
'2010/02/18 an ASK商品パラメータにServer.URLEncodeを行なう
'2011/08/01 an #1087 Error.aspログ出力対応
'2011/10/19 hn 1063 ASK表示方法変更
'2012/01/19 GV データ取得 SELECT文へ LACクエリー案を適用
'2012/01/20 GV データ取得 SELECT文から売筋商品テーブルの最新月のデータのみ抽出する条件を削除
'2012/07/11 ok リニューアル新デザイン変更
'2012/07/23 nt 存在しないカテゴリーコード指定時のエラー画面リダイレクトを追加
'2014/03/19 GV 消費税増税に伴う2重表示対応
'
'========================================================================

On Error Resume Next

Dim CategoryCd

Dim wCategoryName
Dim wMidCategoryCd
Dim wMidCategoryName
Dim wLargeCategoryCd
Dim wLargeCategoryName

Dim wSalesTaxRate
Dim wPrice
Dim wRank

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim Connection
Dim RS

Dim wProductList

Dim wSQL
Dim wHTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add
Dim wNoData    '2012/07/23 nt add

'========================================================================

'---- 送信データーの取り出し
CategoryCd = ReplaceInput(Request("CategoryCd"))
Response.Status="301 Moved Permanently" 
Response.AddHeader "Location", "http://www.soundhouse.co.jp/best_seller/category/" & CategoryCd

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "BestSellerListByCategory.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
end if                                           '2011/08/01 an add e

call close_db()

'2012/07/23 nt add
If wNoData = "Y" Or Err.Description <> "" Then
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

'2012/07/23 nt del
'if Err.Description <> "" then
'	Response.Redirect g_HTTP & "shop/Error.asp"
'end if

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
Dim v_item

'---- 消費税率取出し
call getCntlMst("共通","消費税率","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'消費税率
wSalesTaxRate = Clng(wItemNum1)

'---- ベストセラー商品取り出し
wSQL = ""
' 2012/01/19 GV Mod Start
'wSQL = wSQL & "SELECT a.メーカーコード"
'wSQL = wSQL & "     , a.商品コード"
'wSQL = wSQL & "     , '' AS シリーズコード"
'wSQL = wSQL & "     , b.色"
'wSQL = wSQL & "     , b.規格"
'wSQL = wSQL & "     , a.商品名"
'wSQL = wSQL & "     , a.商品画像ファイル名_小"
'wSQL = wSQL & "     , a.お勧め商品コメント"
'wSQL = wSQL & "     , a.商品概略Web"
'wSQL = wSQL & "     , a.ASK商品フラグ"
'wSQL = wSQL & "     , a.試聴フラグ"
'wSQL = wSQL & "     , a.試聴URL"
'wSQL = wSQL & "     , a.動画フラグ"
'wSQL = wSQL & "     , a.動画URL"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN (a.個数限定数量 > a.個数限定受注済数量 AND a.個数限定数量 > 0) THEN a.個数限定単価"
'wSQL = wSQL & "         ELSE a.販売単価"
'wSQL = wSQL & "       END AS 販売単価"
'wSQL = wSQL & "     , c.メーカー名"
'wSQL = wSQL & "     , e.順位"
'wSQL = wSQL & "     , f.カテゴリー名"
'wSQL = wSQL & "     , g.中カテゴリーコード"
'wSQL = wSQL & "     , g.中カテゴリー名日本語"
'wSQL = wSQL & "     , h.大カテゴリーコード"
'wSQL = wSQL & "     , h.大カテゴリー名"
'wSQL = wSQL & "  FROM Web商品 a WITH (NOLOCK)"
'wSQL = wSQL & "     , Web色規格別在庫 b WITH (NOLOCK)"
'wSQL = wSQL & "     , メーカー c WITH (NOLOCK)"
'wSQL = wSQL & "     , 商品カテゴリー d WITH (NOLOCK)"
'wSQL = wSQL & "     , 売筋商品 e WITH (NOLOCK)"
'wSQL = wSQL & "     , カテゴリー f WITH (NOLOCK)"
'wSQL = wSQL & "     , 中カテゴリー g WITH (NOLOCK)"
'wSQL = wSQL & "     , 大カテゴリー h WITH (NOLOCK)"
'wSQL = wSQL & " WHERE b.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND b.商品コード = a.商品コード"
'wSQL = wSQL & "   AND c.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND d.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND d.商品コード = a.商品コード"
'wSQL = wSQL & "   AND e.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND e.商品コード = a.商品コード"
'wSQL = wSQL & "   AND f.カテゴリーコード = d.カテゴリーコード"
'wSQL = wSQL & "   AND g.中カテゴリーコード = f.中カテゴリーコード"
'wSQL = wSQL & "   AND h.大カテゴリーコード = g.大カテゴリーコード"
'wSQL = wSQL & "   AND d.カテゴリーコード = '" & CategoryCd & "'"
'wSQL = wSQL & "   AND a.取扱中止日 IS NULL"
'wSQL = wSQL & "   AND ((a.廃番日 IS NULL) OR (a.廃番日 IS NOT NULL AND b.引当可能数量 > 0)) "
'wSQL = wSQL & "   AND a.Web商品フラグ = 'Y'"
'wSQL = wSQL & "   AND b.終了日 IS NULL"
'wSQL = wSQL & "   AND e.年月 = (SELECT MAX(年月) FROM 売筋商品)"
'
'wSQL = wSQL & " UNION "
'
'wSQL = wSQL & "SELECT a.メーカーコード"
'wSQL = wSQL & "     , '' AS 商品コード"
'wSQL = wSQL & "     , a.シリーズコード"
'wSQL = wSQL & "     , '' AS 色"
'wSQL = wSQL & "     , '' AS 規格"
'wSQL = wSQL & "     , a.シリーズ名 AS 商品名"
'wSQL = wSQL & "     , a.シリーズ画像ファイル名 AS 商品画像ファイル名_小"
'wSQL = wSQL & "     , a.シリーズ備考 AS お勧め商品コメント"
'wSQL = wSQL & "     , '' AS 商品概略Web"
'wSQL = wSQL & "     , '' AS ASK商品フラグ"
'wSQL = wSQL & "     , '' AS 試聴フラグ"
'wSQL = wSQL & "     , '' AS 試聴URL"
'wSQL = wSQL & "     , '' AS 動画フラグ"
'wSQL = wSQL & "     , '' AS 動画URL"
'wSQL = wSQL & "     , '' AS 販売単価"
'wSQL = wSQL & "     , c.メーカー名"
'wSQL = wSQL & "     , e.順位"
'wSQL = wSQL & "     , f.カテゴリー名"
'wSQL = wSQL & "     , g.中カテゴリーコード"
'wSQL = wSQL & "     , g.中カテゴリー名日本語"
'wSQL = wSQL & "     , h.大カテゴリーコード"
'wSQL = wSQL & "     , h.大カテゴリー名"
'wSQL = wSQL & "  FROM シリーズ a WITH (NOLOCK)"
'wSQL = wSQL & "     , メーカー c WITH (NOLOCK)"
'wSQL = wSQL & "     , 売筋商品 e WITH (NOLOCK)"
'wSQL = wSQL & "     , カテゴリー f WITH (NOLOCK)"
'wSQL = wSQL & "     , 中カテゴリー g WITH (NOLOCK)"
'wSQL = wSQL & "     , 大カテゴリー h WITH (NOLOCK)"
'wSQL = wSQL & " WHERE c.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND e.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND e.シリーズコード = a.シリーズコード"
'wSQL = wSQL & "   AND f.カテゴリーコード = a.カテゴリーコード"
'wSQL = wSQL & "   AND g.中カテゴリーコード = f.中カテゴリーコード"
'wSQL = wSQL & "   AND h.大カテゴリーコード = g.大カテゴリーコード"
'wSQL = wSQL & "   AND e.カテゴリーコード = '" & CategoryCd & "'"
'wSQL = wSQL & "   AND e.年月 = (SELECT MAX(年月) FROM 売筋商品)"
'
'wSQL = wSQL & " ORDER BY"
'wSQL = wSQL & "       e.順位"
'wSQL = wSQL & "     , メーカー名"
'wSQL = wSQL & "     , 商品名"
'wSQL = wSQL & "     , 色"
'wSQL = wSQL & "     , 規格"
wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.メーカーコード "
wSQL = wSQL & "    , a.商品コード "
wSQL = wSQL & "    , '' AS シリーズコード "
'wSQL = wSQL & "    , b.色 "  2012/07/11 nt del
'wSQL = wSQL & "    , b.規格 "  2012/07/11 nt del
wSQL = wSQL & "    , a.商品名 "
wSQL = wSQL & "    , a.商品画像ファイル名_小 "
wSQL = wSQL & "    , a.お勧め商品コメント "
wSQL = wSQL & "    , a.商品概略Web "
wSQL = wSQL & "    , a.ASK商品フラグ "
wSQL = wSQL & "    , a.試聴フラグ "
wSQL = wSQL & "    , a.試聴URL "
wSQL = wSQL & "    , a.動画フラグ "
wSQL = wSQL & "    , a.動画URL "
'wSQL = wSQL & "    , CASE "
'wSQL = wSQL & "        WHEN (a.個数限定数量 > a.個数限定受注済数量 AND a.個数限定数量 > 0) THEN a.個数限定単価 "
'wSQL = wSQL & "        ELSE a.販売単価 "
'wSQL = wSQL & "      END AS 販売単価 "
wSQL = wSQL & "    , a.販売単価 "
wSQL = wSQL & "    , a.個数限定数量 "
wSQL = wSQL & "    , a.個数限定受注済数量 "
wSQL = wSQL & "    , a.個数限定単価 "
wSQL = wSQL & "    , a.B品フラグ "
wSQL = wSQL & "    , a.B品単価 "
wSQL = wSQL & "    , c.メーカー名 "
wSQL = wSQL & "    , e.順位 "
wSQL = wSQL & "    , f.カテゴリー名 "
wSQL = wSQL & "    , g.中カテゴリーコード "
wSQL = wSQL & "    , g.中カテゴリー名日本語 "
wSQL = wSQL & "    , h.大カテゴリーコード "
wSQL = wSQL & "    , h.大カテゴリー名 "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    Web商品                      a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN Web色規格別在庫 b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.メーカーコード     = a.メーカーコード "
wSQL = wSQL & "           AND b.商品コード         = a.商品コード "
wSQL = wSQL & "      INNER JOIN メーカー        c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.メーカーコード     = a.メーカーコード "
wSQL = wSQL & "      INNER JOIN 商品カテゴリー  d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.メーカーコード     = a.メーカーコード "
wSQL = wSQL & "           AND d.商品コード         = a.商品コード "
wSQL = wSQL & "      INNER JOIN 売筋商品        e WITH (NOLOCK) "
wSQL = wSQL & "        ON     e.メーカーコード     = a.メーカーコード "
wSQL = wSQL & "           AND e.商品コード         = a.商品コード "
wSQL = wSQL & "           AND e.カテゴリーコード   = d.カテゴリーコード"	'2012/7/11 ok add
wSQL = wSQL & "      INNER JOIN カテゴリー      f WITH (NOLOCK) "
wSQL = wSQL & "        ON     f.カテゴリーコード   = d.カテゴリーコード "
wSQL = wSQL & "      INNER JOIN 中カテゴリー    g WITH (NOLOCK) "
wSQL = wSQL & "        ON     g.中カテゴリーコード = f.中カテゴリーコード "
wSQL = wSQL & "      INNER JOIN 大カテゴリー    h WITH (NOLOCK) "
wSQL = wSQL & "        ON     h.大カテゴリーコード = g.大カテゴリーコード "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' )   t1 "
wSQL = wSQL & "        ON     a.Web商品フラグ    = t1.ShohinWebY "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        t1.ShohinWebY  IS NOT NULL "
wSQL = wSQL & "    AND a.取扱中止日   IS NULL "
wSQL = wSQL & "    AND (   (a.廃番日 IS NULL) "
wSQL = wSQL & "         OR (    a.廃番日 IS NOT NULL "
wSQL = wSQL & "             AND b.引当可能数量 > 0)) "
wSQL = wSQL & "    AND b.終了日 IS NULL "
'wSQL = wSQL & "    AND e.年月 = (SELECT MAX(年月) FROM 売筋商品) "				' 2012/01/20 GV Del
wSQL = wSQL & "    AND d.カテゴリーコード = '" & CategoryCd & "' "

'---- 色・規格データ不要のため、GROUP BY句を追加(2012/07/11 nt add)
wSQL = wSQL & "GROUP BY "
wSQL = wSQL & "           a.メーカーコード "
wSQL = wSQL & "         , a.商品コード "
wSQL = wSQL & "         , a.商品名 "
wSQL = wSQL & "         , a.商品画像ファイル名_小 "
wSQL = wSQL & "         , a.お勧め商品コメント "
wSQL = wSQL & "         , a.商品概略Web "
wSQL = wSQL & "         , a.ASK商品フラグ "
wSQL = wSQL & "         , a.試聴フラグ "
wSQL = wSQL & "         , a.試聴URL "
wSQL = wSQL & "         , a.動画フラグ "
wSQL = wSQL & "         , a.動画URL "
wSQL = wSQL & "         , a.個数限定数量"
wSQL = wSQL & "         , a.個数限定受注済数量"
wSQL = wSQL & "         , a.個数限定単価 "
wSQL = wSQL & "         , a.販売単価 "
wSQL = wSQL & "         , a.B品フラグ "
wSQL = wSQL & "         , a.B品単価 "
wSQL = wSQL & "         , c.メーカー名 "
wSQL = wSQL & "         , e.順位 "
wSQL = wSQL & "         , f.カテゴリー名 "
wSQL = wSQL & "         , g.中カテゴリーコード "
wSQL = wSQL & "         , g.中カテゴリー名日本語 "
wSQL = wSQL & "         , h.大カテゴリーコード "
wSQL = wSQL & "         , h.大カテゴリー名 "

wSQL = wSQL & "UNION "

wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.メーカーコード "
wSQL = wSQL & "    , '' AS 商品コード "
wSQL = wSQL & "    , a.シリーズコード "
'wSQL = wSQL & "    , '' AS 色 "  2012/07/11 nt del
'wSQL = wSQL & "    , '' AS 規格 "  2012/07/11 nt del
wSQL = wSQL & "    , a.シリーズ名 AS 商品名 "
wSQL = wSQL & "    , a.シリーズ画像ファイル名 AS 商品画像ファイル名_小 "
wSQL = wSQL & "    , a.シリーズ備考 AS お勧め商品コメント "
wSQL = wSQL & "    , '' AS 商品概略Web "
wSQL = wSQL & "    , '' AS ASK商品フラグ "
wSQL = wSQL & "    , '' AS 試聴フラグ "
wSQL = wSQL & "    , '' AS 試聴URL "
wSQL = wSQL & "    , '' AS 動画フラグ "
wSQL = wSQL & "    , '' AS 動画URL "
wSQL = wSQL & "    , '' AS 販売単価 "
wSQL = wSQL & "    , '' AS 個数限定数量 "
wSQL = wSQL & "    , '' AS 個数限定受注済数量 "
wSQL = wSQL & "    , '' AS 個数限定単価 "
wSQL = wSQL & "    , '' AS B品フラグ "
wSQL = wSQL & "    , '' AS B品単価 "
wSQL = wSQL & "    , c.メーカー名 "
wSQL = wSQL & "    , e.順位 "
wSQL = wSQL & "    , f.カテゴリー名 "
wSQL = wSQL & "    , g.中カテゴリーコード "
wSQL = wSQL & "    , g.中カテゴリー名日本語 "
wSQL = wSQL & "    , h.大カテゴリーコード "
wSQL = wSQL & "    , h.大カテゴリー名 "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    シリーズ                  a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN メーカー     c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.メーカーコード     = a.メーカーコード "
wSQL = wSQL & "      INNER JOIN 売筋商品     e WITH (NOLOCK) "
wSQL = wSQL & "        ON     e.メーカーコード     = a.メーカーコード "
wSQL = wSQL & "           AND e.シリーズコード     = a.シリーズコード "
wSQL = wSQL & "      INNER JOIN カテゴリー   f WITH (NOLOCK) "
wSQL = wSQL & "        ON     f.カテゴリーコード   = a.カテゴリーコード "
wSQL = wSQL & "      INNER JOIN 中カテゴリー g WITH (NOLOCK) "
wSQL = wSQL & "        ON     g.中カテゴリーコード = f.中カテゴリーコード "
wSQL = wSQL & "      INNER JOIN 大カテゴリー h WITH (NOLOCK) "
wSQL = wSQL & "        ON     h.大カテゴリーコード = g.大カテゴリーコード "
wSQL = wSQL & "WHERE "
'wSQL = wSQL & "        e.年月 = (SELECT MAX(年月) FROM 売筋商品) "				' 2012/01/20 GV Del
wSQL = wSQL & "        e.カテゴリーコード = '" & CategoryCd & "' "

'---- 色・規格データ不要のため、GROUP BY句を追加(2012/07/11 nt add)
wSQL = wSQL & "GROUP BY "
wSQL = wSQL & "           a.メーカーコード "
wSQL = wSQL & "         , a.シリーズコード  "
wSQL = wSQL & "         , a.シリーズ名"
wSQL = wSQL & "         , a.シリーズ画像ファイル名"
wSQL = wSQL & "         , a.シリーズ備考"
wSQL = wSQL & "         , c.メーカー名 "
wSQL = wSQL & "         , e.順位 "
wSQL = wSQL & "         , f.カテゴリー名 "
wSQL = wSQL & "         , g.中カテゴリーコード "
wSQL = wSQL & "         , g.中カテゴリー名日本語 "
wSQL = wSQL & "         , h.大カテゴリーコード "
wSQL = wSQL & "         , h.大カテゴリー名 "

wSQL = wSQL & "ORDER BY "
wSQL = wSQL & "      e.順位 "
wSQL = wSQL & "    , メーカー名 "
wSQL = wSQL & "    , 商品名 "
'wSQL = wSQL & "    , 色 " 2012/07/11 nt del
'wSQL = wSQL & "    , 規格 " 2012/07/11 nt del
' 2012/01/19 GV Mod End

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

'@@@@@		response.write(wSQL)

'2012/07/23 nt add
if RS.EOF = True then
	wNoData = "Y"
else
'if RS.EOF = false then 2012/07/23 nt del
	wCategoryName = RS("カテゴリー名")
	wMidCategoryCd = RS("中カテゴリーコード")
	wMidCategoryName = RS("中カテゴリー名日本語")
	wLargeCategoryCd = RS("大カテゴリーコード")
	wLargeCategoryName = RS("大カテゴリー名")
end if

wRank = 0
wHTML = ""
wHTML = wHTML & "    <ul class='item_list listtype bestseller'>" & vbNewLine

Do Until RS.EOF = true
'2012/07/23 ok Mod Start
'	wHTML = wHTML & "  <tr valign='top'>" & vbNewLine

	'---- 商品
	call CreateProductHTML()

'	if RS.EOF = false then
	'---- 中央 商品
'		call CreateProductHTML()

'		if RS.EOF = false then
		'---- 右 商品
'			call CreateProductHTML()
'		end if
'	end if

'	wHTML = wHTML & "  </tr>" & vbNewLine
	RS.MoveNext
'	if RS.EOF = true then
'		Exit Do
'	end if
Loop

'wHTML = wHTML & "</table>" & vbNewLine
wHTML = wHTML & "    </ul>" & vbNewLine
'2012/07/23 ok Mod End
wProductList = wHTML

RS.Close

End function

'========================================================================
'
'	Function	個別商品HTML作成
'
'========================================================================
'
Function CreateProductHTML()
Dim vComment
Dim vOldProductCd
Dim vIroKikakuFl
Dim vWidth1
Dim vWidth2
Dim vItemCnt
Dim vItemList()
Dim vSoundMovie

wRank = wRank + 1

'2012/07/11 ok Mod Start
If wRank < 4 Then
	wHTML = wHTML & "      <li class='toprank'>" & vbNewLine
Else
	wHTML = wHTML & "      <li>" & vbNewLine
End If

wPrice = calcPrice(RS("販売単価"), wSalesTaxRate)

wHTML = wHTML & "        <ul class='detail'>" & vbNewLine
wHTML = wHTML & "          <li class='rank'><img src='../top_images/ranking/rank" & Right("0" & Cstr(wRank),2) & ".png' alt='" & wRank & "位'></li>" & vbNewLine

'---- メーカー名，商品番号，画像
wHTML = wHTML & "          <li class='prod_name'><strong>" & RS("メーカー名") & " / <a href='ProductDetail.asp?item=" & RS("メーカーコード") & "^" & Server.URLEncode(RS("商品コード")) & "'>" & RS("商品名") & "</a></strong></li>" & vbNewLine

'wHTML = wHTML & "          <li>販売価格：" & FormatNumber(wPrice,0) & "円（税込）</li>" & vbNewLine

wHTML = wHTML & "                        <li class='price'>"
If RS("ASK商品フラグ") <> "Y" Then
	'---- B品単価
	If RS("B品フラグ") = "Y" Then
		wPrice = calcPrice(RS("B品単価"), wSalesTaxRate)
'2014/03/19 GV mod start ---->
'		wHTML = wHTML & "わけあり品特価：<span>" & FormatNumber(wPrice,0) & "円</span>(税込)"
		wHTML = wHTML & "わけあり品特価：<span>" & FormatNumber(RS("B品単価"),0) & "円</span>(税抜)<br>"
		wHTML = wHTML & "(税込&nbsp;<span>" & FormatNumber(wPrice,0) & "円</span>)"
'2014/03/19 GV mod end   <----
	'---- 個数限定単価
	ElseIf RS("個数限定数量") > RS("個数限定受注済数量") AND RS("個数限定数量") > 0 Then
		wPrice = calcPrice(RS("個数限定単価"), wSalesTaxRate)
'2014/03/19 GV mod start ---->
'		wHTML = wHTML & "限定特価：<span>" & FormatNumber(wPrice,0) & "円</span>(税込)"
		wHTML = wHTML & "限定特価：<span>" & FormatNumber(RS("個数限定単価"),0) & "円</span>(税抜)<br>"
		wHTML = wHTML & "(税込&nbsp;<span>" & FormatNumber(wPrice,0) & "円</span>)"
'2014/03/19 GV mod end   <----
	'---- 通常商品
	Else
'2014/03/19 GV mod start ---->
'		wHTML = wHTML & "衝撃特価：<span>" & FormatNumber(wPrice,0) & "円</span>(税込)"
		wHTML = wHTML & "衝撃特価：<span>" & FormatNumber(RS("販売単価"),0) & "円</span>(税抜)<br>"
		wHTML = wHTML & "(税込&nbsp;<span>" & FormatNumber(wPrice,0) & "円</span>)"
'2014/03/19 GV mod end   <----
	End If
Else
	'---- B品単価
	If RS("B品フラグ") = "Y" Then
		wPrice = calcPrice(RS("B品単価"), wSalesTaxRate)
'2014/03/19 GV mod start ---->
'		wHTML = wHTML & "わけあり品特価：<a class='tip'>ASK<span>" & FormatNumber(wPrice,0) & "円(税込)</span></a>"
		wHTML = wHTML & "わけあり品特価：<a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RS("B品単価"),0) & "円(税抜)</span><br>"
		wHTML = wHTML & "<span class='inc-tax'>(税込&nbsp;" & FormatNumber(wPrice,0) & "円)</span></a>"
'2014/03/19 GV mod end   <----
	'---- 個数限定単価
	ElseIf RS("個数限定数量") > RS("個数限定受注済数量") AND RS("個数限定数量") > 0 Then
		wPrice = calcPrice(RS("個数限定単価"), wSalesTaxRate)
'2014/03/19 GV mod start ---->
'		wHTML = wHTML & "限定特価：<a class='tip'>ASK<span>" & FormatNumber(wPrice,0) & "円(税込)</span></a>"
		wHTML = wHTML & "限定特価：<a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RS("個数限定単価"),0) & "円(税抜)</span><br>"
		wHTML = wHTML & "<span class='inc-tax'>(税込&nbsp;" & FormatNumber(wPrice,0) & "円)</span></a>"
'2014/03/19 GV mod end   <----
	'---- 通常商品
	Else
'2014/03/19 GV mod start ---->
'		wHTML = wHTML & "衝撃特価：<a class='tip'>ASK<span>" & FormatNumber(wPrice,0) & "円(税込)</span></a>"
		wHTML = wHTML & "衝撃特価：<a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RS("販売単価"),0) & "円(税抜)</span><br>"
		wHTML = wHTML & "<span class='inc-tax'>(税込&nbsp;" & FormatNumber(wPrice,0) & "円)</span></a>"
'2014/03/19 GV mod end   <----
	End If
End If
wHTML = wHTML & "</li>" & vbNewLine

wHTML = wHTML & "          <li class='photo'><a href='ProductDetail.asp?item=" & RS("メーカーコード") & "^" & Server.URLEncode(RS("商品コード")) & "'>"
If RS("商品画像ファイル名_小") <> "" Then
	wHTML = wHTML & "<img src='prod_img/" & RS("商品画像ファイル名_小") & "' alt='" & Replace(RS("メーカー名") & " / " & RS("商品名"),"'","&#39;") & "' class='opover'>"
End If
wHTML = wHTML & "</a></li>" & vbNewLine
'---- 商品説明
If RS("お勧め商品コメント") <> "" Then
	wHTML = wHTML & "        <li>" & Replace(RS("お勧め商品コメント"), vbNewLine, "<br>") & "</li>" & vbNewLine
Else
	wHTML = wHTML & "        <li>" & Replace(RS("商品概略Web"), vbNewLine, "<br>") & "</li>" & vbNewLine
End If
wHTML = wHTML & "        </ul>" & vbNewLine

wHTML = wHTML & "        <div class='other_detail'>" & vbNewLine
wHTML = wHTML & "          <ul>" & vbNewLine
wHTML = wHTML & "            <li><a href='ProductDetail.asp?item=" & RS("メーカーコード") & "^" & Server.URLEncode(RS("商品コード")) &"'><img src='images/btn_detail.png' alt='詳細を見る' class='opover'></a></li>"
wHTML = wHTML & "          </ul>" & vbNewLine
wHTML = wHTML & "        </div>" & vbNewLine
wHTML = wHTML & "      </li>" & vbNewLine
'2012/07/11 ok End Start

'2012/07/11 nt del Start
''----試聴リンク
'vSoundMovie = ""
'if RS("試聴フラグ") = "Y" AND RS("試聴URL") <> "" then
'	vItemCnt = cf_unstring(RS("試聴URL"), vItemList, ",")
'	if vItemCnt > 1 then
'		vSoundMovie = vSoundMovie & "<a href='JavaScript:void(0);' onClick=""window.open('SoundMoviePopUp.asp?item=" & RS("メーカーコード") & "^" & Server.URLEncode(RS("商品コード")) & "','SoundMovie', 'width=201 height=200 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');"" class='link'><img src='images/Shichou.gif' width='18' height='18' border='0' alt='試聴する'></a>&nbsp;"
'	else
'		if InStr(LCase(RS("試聴URL")), "http://") > 0 then
'			vSoundMovie = "<a href='" & RS("試聴URL") & "' target='_blank'><img src='images/Shichou.gif' width='18' height='18' border='0' alt='試聴する'></a>&nbsp;&nbsp;"
'		else
'			vSoundMovie = "<a href='" & g_HTTP & RS("試聴URL") & "' target='_blank'><img src='images/Shichou.gif' width='18' height='18' border='0' alt='試聴する'></a>&nbsp;&nbsp;"
'		end if
'	end if
'end if
'
''----動画リンク
'if RS("動画フラグ") = "Y" AND RS("動画URL") <> "" then
'	vItemCnt = cf_unstring(RS("動画URL"), vItemList, ",")
'	if vItemCnt > 1 then
'		vSoundMovie = vSoundMovie & "<a href='JavaScript:void(0);' onClick=""window.open('SoundMoviePopUp.asp?item=" & RS("メーカーコード") & "^" & Server.URLEncode(RS("商品コード")) & "','SoundMovie', 'width=201 height=200 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');"" class='link'><img src='images/Movie.jpg' width='18' height='18' border='0' alt='動画を見る'></a>&nbsp;"
'	else
'		if InStr(LCase(RS("動画URL")), "http://") > 0 then
'			vSoundMovie = vSoundMovie & "<a href='" & RS("動画URL") & "' target='_blank'><img src='images/Movie.jpg' width='18' height='18' border='0' alt='動画を見る'></a>&nbsp;"
'		else
'			vSoundMovie = vSoundMovie & "<a href='" & g_HTTP & RS("動画URL") & "' target='_blank'><img src='images/Movie.jpg' width='18' height='18' border='0' alt='動画を見る'></a>&nbsp;"
'		end if
'	end if
'end if
'
'if vSoundMovie <> "" then
'	wHTML = wHTML & "        <tr align='left' valign='middle'>" & vbNewLine
'	wHTML = wHTML & "          <td height='25' colspan='2' class='honbun'>" & vbNewLine
'	wHTML = wHTML & "            サンプル： " & vSoundMovie & vbNewLine
'	wHTML = wHTML & "          </td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine
'end if
'
'vOldProductCd = RS("メーカーコード") & "+" & RS("商品コード")
'
''---- 同一商品終了まで繰り返し (色規格がある場合のみ繰り返し)
'Do until vOldProductCd <> RS("メーカーコード") & "+" & RS("商品コード")
'	'---- 色規格, 単価
'	wHTML = wHTML & "        <tr>" & vbNewLine
'	wHTML = wHTML & "          <td width='" & vWidth1 & "' height='25' align='right' nowrap>" & vbNewLine
'	if RS("色") <> "" OR RS("規格") <> "" then
'		wHTML = wHTML & "            <a href='ProductDetail.asp?item=" & RS("メーカーコード") & "^" & Server.URLEncode(RS("商品コード")) & "^" & RS("色") & "^" & RS("規格") & "' class='link'>"
'		if RS("色") <> "" then
'			wHTML = wHTML & RS("色") & " "
'		end if
'		if RS("規格") <> "" then
'			wHTML = wHTML & RS("規格")
'		end if
'		wHTML = wHTML & "</a>"
'	end if
'
'	wPrice = calcPrice(RS("販売単価"), wSalesTaxRate)
'
'	if RS("商品コード") <> "" then
'		if RS("ASK商品フラグ") = "Y" then
'
''2011/10/19 hn  mod s
''			wHTML = wHTML & "         <a href='JavaScript:void(0);' onClick=""askWin=window.open('AskPrice.asp?MakerName=" & Server.URLEncode(RS("メーカー名")) & "&ProductName=" & Server.URLEncode(RS("商品名")) & "&Price=" & wPrice & "' ,'ask', 'width=250 height=80 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');"" class='link'>ASK</a>" & vbNewLine
'			wHTML = wHTML & "                        <a class='tip'>ASK<span>" & FormatNumber(wPrice,0) & "円(税込)</span></a>" & vbNewLine
''2011/10/19 hn mod e
'
'		else
'			wHTML = wHTML & "            <span class='honbun'>" & FormatNumber(wPrice,0) & "円(税込)</span>" & vbNewLine
'		end if
'	else
'		wHTML = wHTML & "            &nbsp;" & vbNewLine
'	end if
'
'	wHTML = wHTML & "          </td>" & vbNewLine
'
'	'---- 詳細ボタン，カートボタン
'	wHTML = wHTML & "          <td width='" & vWidth2 & "' nowrap height='25' align='center' valign='middle'>" & vbNewLine
'
'	' 通常商品
'	if RS("商品コード") <> "" then
'		if vIroKikakuFl = false then
'			wHTML = wHTML & "            <a href='ProductDetail.asp?item=" & RS("メーカーコード") & "^" & Server.URLEncode(RS("商品コード")) & "^" & RS("色") & "^" & RS("規格") & "'><img src='images/Shousai.gif' width='50' height='19' border='0' align='middle'></a>" & vbNewLine
'		end if
'		wHTML = wHTML & "            <a href='OrderPreInsert.asp?maker_cd=" & RS("メーカーコード") & "&product_cd=" & Server.URLEncode(RS("商品コード")) & "&iro=" & RS("色") & "&kikaku=" & RS("規格") & "&qt=1'><img src='images/CartBlue.gif' width='30' height='19' border='0' align='middle'></a>" & vbNewLine
'
'	' シリーズ
'	else
'		wHTML = wHTML & "            <a href='SearchList.asp?i_type=se&sSeriesCd=" & RS("シリーズコード") & "'><img src='images/Shousai.gif' width='50' height='19' border='0' align='middle'></a>"
'		wHTML = wHTML & "            <img src='images/blank.gif' width='30' height='19' border='0' align='middle'>" & vbNewLine
'	end if
'
'	wHTML = wHTML & "          </td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine
'
'	RS.MoveNext
'	if RS.EOF = true then
'		Exit Do
'	end if
'Loop
'
'wHTML = wHTML & "      </table>" & vbNewLine
'wHTML = wHTML & "    </td>" & vbNewLine
'2012/07/11 nt del End

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
<meta name="Description" content="サウンドハウス「<%=wCategoryName%>」の売れ筋（ベストセラー）TOP10商品をご案内します。">
<meta name="keyword" content="<%=wCategoryName%>">
<title><%=wCategoryName%>の売れ筋TOP10｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/shop.css" type="text/css">
<link rel="stylesheet" href="Style/searchlist.css" type="text/css">
<link rel="stylesheet" href="style/ask.css?20140401a" type="text/css">
</head>
<body>
<!--#include file="../Navi/NaviTop.inc"-->
<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>
  <!-- コンテンツstart -->
  <div id="globalContents">
    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="../"><img src="../images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='LargeCategoryList.asp?LargeCategoryCd=<%=wLargeCategoryCd%>' itemprop='url'><span itemprop='title'><%=wLargeCategoryName%></span></a></span></li>
        <li><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='MidCategoryList.asp?MidCategoryCd=<%=wMidCategoryCd%>' itemprop='url'><span itemprop='title'><%=wMidCategoryName%></span></a></span></li>
        <li><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='SearchList.asp?i_type=c&s_category_cd=<%=CategoryCd%>' itemprop='url'><span itemprop='title'><%=wCategoryName%></span></a></span></li>
        <li class="now"><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><span itemprop='title'>売れ筋TOP10</span></span></li>
      </ul>
    </div></div></div>

    <h1 class="title"><%=wCategoryName%>の売れ筋TOP10</h1>

<!-- ベストセラー商品一覧-->
<%=wProductList%>

    <!--/#contents --></div>
  <div id="globalSide">
  <!--#include file="../Navi/NaviSide.inc"-->
  <!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
<div class="tooltip"><p>ASK</p></div>
<script type="text/javascript" src="jslib/ask.js?20140401a"></script>
<script type="text/javascript" src="jslib/SearchList.js?20120321" charset="Shift_JIS"></script>
<script type="text/javascript" src="../jslib/jquery.tinyscrollbar.min.js"></script>
<script type="text/javascript">
$(function(){
    $('#scrollbar1').tinyscrollbar();
});
</script>
</body>
</html>