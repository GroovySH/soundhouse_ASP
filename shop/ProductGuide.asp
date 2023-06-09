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
<!--#include file="../common/SearchListCommon.inc"-->

<%
'========================================================================
'
'	おすすめ商品ページ自動生成
'
'	更新履歴
'2005/02/16 個数限定数量単価取り出し時の条件強化　個数限定数量＞0を追加
'2005/02/21 新製品の抽出条件に発売日を使用するよう変更
'2005/03/02 タイトルにもカテゴリーコードを表示するように変更
'2005/03/15 シリーズ商品表示追加
'2005/03/18 ASK商品単価表示変更
'2005/07/21 シリーズ明細表示からメーカー名を削除（シリーズ名でメーカ名が表示されているため)
'2005/08/10 色規格がある場合は、シリーズと同じ形式の表示を行う｡
'2005/09/07 商品抽出時サブカテゴリーも考慮
'2005/09/08 メインカテゴリー商品を先に、サブカテゴリー商品を後に表示するよう変更
'2005/11/01 試聴・Movieポップアップページへのリンク追加
'2006/01/10 試聴、動画のリンクにhttpが含まれている場合は外部リンクとする
'2006/01/23 シリーズ表示時のソート順を販売単価に変更、シリーズ商品を先に表示
'2006/04/05 パンくず追加
'2006/12/08 タイトル変更
'2007/01/22 在庫状況表示追加
'2007/03/15 パラメータに対してReplaceInputを追加
'2007/05/30 色規格あり商品対応
'2008/10/23 色規格なしおすすめ商品が引当可能なのに表示されなかった不具合対応
'2008/10/23 (変更依頼#503)個数限定数量の表示を次のように変更
'						4以下 現行どおり/5-9 限定5個/10-14 限定10個/15-19 限定15個/20以上、限定20個
'2008/12/24 在庫状況セット関数化
'2010/02/12 hn ASK商品パラメータにServer.URLEncodeを行なう
'2010/11/10 an シリーズ商品のソート順を修正→表示順指定がないときに表示が崩れないように
'2011/08/01 an #1087 Error.aspログ出力対応
'2011/10/19 hn 1063 ASK表示方法変更
'2012/01/19 GV データ取得 SELECT文へ LACクエリー案を適用
'2012/07/14 nt リニューアル用にデータ取得 SELECT文およびasp画面出力を修正
'2012/07/23 nt 存在しないカテゴリーコード指定時のエラー画面リダイレクトを追加
'
'========================================================================

On Error Resume Next

Dim CategoryCd

Dim wCategoryName
Dim wMidCategoryCd
Dim wMidCategoryName
Dim wLargeCategoryCd
Dim wLargeCategoryName
Dim wCategoryComment

Dim wSalesTaxRate
Dim wPrice

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim Connection
Dim RS

Dim wRedirectURL
Dim wProductList
Dim wProductList2
Dim wSaleItemHTML
Dim wNewItemHTML
Dim wKanrenCtegoryLinkHTML

Dim wSQL
Dim wHTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add
Dim wNoData '2012/07/23 nt add

'========================================================================

Response.buffer = true

'---- 送信データーの取り出し
CategoryCd = ReplaceInput(Request("CategoryCd"))
Response.Status="301 Moved Permanently" 
Response.AddHeader "Location", "http://www.soundhouse.co.jp/products/guide/?s_category_cd=" & CategoryCd

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "ProductGuide.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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

if wRedirectURL <> "" then
	Response.Redirect wRedirectURL
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
Dim v_item

'---- 消費税率取出し
call getCntlMst("共通","消費税率","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'消費税率
wSalesTaxRate = Clng(wItemNum1)

'---- カテゴリー情報取り出し
wSQL = ""
' 2012/01/19 GV Mod Start
'wSQL = wSQL & "SELECT a.カテゴリー名"
'wSQL = wSQL & "     , a.お勧めカテゴリーコメント"
'wSQL = wSQL & "     , a.お勧めカテゴリーURL"
'wSQL = wSQL & "     , b.中カテゴリーコード"
'wSQL = wSQL & "     , b.中カテゴリー名日本語"
'wSQL = wSQL & "     , c.大カテゴリーコード"
'wSQL = wSQL & "     , c.大カテゴリー名"
'wSQL = wSQL & "  FROM カテゴリー a WITH (NOLOCK)"
'wSQL = wSQL & "     , 中カテゴリー b WITH (NOLOCK)"
'wSQL = wSQL & "     , 大カテゴリー c WITH (NOLOCK)"
'wSQL = wSQL & " WHERE b.中カテゴリーコード = a.中カテゴリーコード"
'wSQL = wSQL & "   AND c.大カテゴリーコード = b.大カテゴリーコード"
'wSQL = wSQL & "   AND a.カテゴリーコード = '" & CategoryCd & "'"
wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.カテゴリー名 "
wSQL = wSQL & "    , a.お勧めカテゴリーコメント "
wSQL = wSQL & "    , a.お勧めカテゴリーURL "
wSQL = wSQL & "    , b.中カテゴリーコード "
wSQL = wSQL & "    , b.中カテゴリー名日本語 "
wSQL = wSQL & "    , c.大カテゴリーコード "
wSQL = wSQL & "    , c.大カテゴリー名 "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    カテゴリー                a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN 中カテゴリー b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.中カテゴリーコード = a.中カテゴリーコード "
wSQL = wSQL & "      INNER JOIN 大カテゴリー c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.大カテゴリーコード = b.大カテゴリーコード "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        a.カテゴリーコード = '" & CategoryCd & "' "
' 2012/01/19 GV Mod End

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

wRedirectURL = ""

'2012/07/23 nt add
if RS.EOF = True then
	wNoData = "Y"
	RS.close	'2012/07/23 nt add
else
'if RS.EOF = false then	'2012/07/23 nt del
	if RS("お勧めカテゴリーURL") <> "" then
		wRedirectURL = RS("お勧めカテゴリーURL")
		exit function
	else
		if RS("お勧めカテゴリーコメント") <> "" then
' 2012/07/18 nt Mod Start
			wCategoryComment = wCategoryComment & Replace(RS("お勧めカテゴリーコメント"), vbNewLine, "<br>") & vbnewline
'			wCategoryComment = "            <table width='610' border='1' cellspacing='0' cellpadding='2' bordercolor='#999999' bordercolorlight='#999999' bordercolordark='#ffffff' >" & vbnewline
'			wCategoryComment = wCategoryComment & "              <tr align='center' valign='top'>" & vbnewline
'			wCategoryComment = wCategoryComment & "                <td align='left' bgcolor='#ffffee' class='honbun' >" & vbnewline
'			wCategoryComment = wCategoryComment & Replace(RS("お勧めカテゴリーコメント"), vbNewLine, "<br>") & vbnewline
'			wCategoryComment = wCategoryComment & "                </td>" & vbnewline
'			wCategoryComment = wCategoryComment & "              </tr>" & vbnewline
'			wCategoryComment = wCategoryComment & "            </table>" & vbnewline
' 2012/07/18 nt Mod End
		else
			wCategoryComment = ""
		end if
	end if
	wCategoryName = RS("カテゴリー名")
	wMidCategoryCd = RS("中カテゴリーコード")
	wMidCategoryName = RS("中カテゴリー名日本語")
	wLargeCategoryCd = RS("大カテゴリーコード")
	wLargeCategoryName = RS("大カテゴリー名")

	RS.close	'2012/07/23 nt add

	'---- お勧め商品情報取り出し
	call CreateProductList()	'2012/07/23 nt add

	'---- お勧め商品情報取り出し2（シリーズ商品)	'2005/03/14
	call CreateProductList2()	'2012/07/23 nt add

end if

'RS.close	'2012/07/23 nt del

'---- お勧め商品情報取り出し
'call CreateProductList()	'2012/07/23 nt del

'---- お勧め商品情報取り出し2（シリーズ商品)	'2005/03/14
'call CreateProductList2()	'2012/07/23 nt del

'---- 衝撃特価品情報取り出し
'call CreateSaleItemHTML	'2012/07/18 nt del

'---- 新商品情報取り出し
'Call CreateNewItemHTML	'2012/07/18 nt del

'---- 関連カテゴリーリンク作成
'Call CreateCategoryLinkHTML()	'2012/07/18 nt del

End Function

'========================================================================
'
'	Function	お勧め商品情報編集
'
'========================================================================
'
Function CreateProductList()

'---- 商品Recordset作成
wSQL = ""
' 2012/01/19 GV Mod Start
'wSQL = wSQL & "SELECT a.メーカーコード"
'wSQL = wSQL & "     , a.商品コード"
'wSQL = wSQL & "     , b.色"
'wSQL = wSQL & "     , b.規格"
'wSQL = wSQL & "     , a.商品名"
'wSQL = wSQL & "     , a.商品画像ファイル名_小"
'wSQL = wSQL & "     , a.お勧め商品コメント"
'wSQL = wSQL & "     , a.ASK商品フラグ"
'wSQL = wSQL & "     , a.試聴フラグ"
'wSQL = wSQL & "     , a.試聴URL"
'wSQL = wSQL & "     , a.動画フラグ"
'wSQL = wSQL & "     , a.動画URL"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN (a.個数限定数量 > a.個数限定受注済数量 AND a.個数限定数量 > 0) THEN a.個数限定単価"
'wSQL = wSQL & "         ELSE a.販売単価"
'wSQL = wSQL & "       END AS 販売単価"
'wSQL = wSQL & "     , a.個数限定数量"
'wSQL = wSQL & "     , a.個数限定受注済数量"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN (a.個数限定数量 > a.個数限定受注済数量 AND a.個数限定数量 > 0) THEN 'Y'"
'wSQL = wSQL & "         ELSE 'N'"
'wSQL = wSQL & "       END AS 個数限定単価フラグ"
'wSQL = wSQL & "     , a.希少数量"
'wSQL = wSQL & "     , a.セット商品フラグ"
'wSQL = wSQL & "     , a.メーカー直送取寄区分"
'wSQL = wSQL & "     , a.Web納期非表示フラグ"
'wSQL = wSQL & "     , a.廃番日"
'wSQL = wSQL & "     , a.B品フラグ"
'wSQL = wSQL & "     , a.入荷予定未定フラグ"
'wSQL = wSQL & "     , b.引当可能入荷予定日"
'wSQL = wSQL & "     , b.引当可能数量"
'wSQL = wSQL & "     , b.B品引当可能数量"
'wSQL = wSQL & "     , c.メーカー名"
'wSQL = wSQL & "  FROM Web商品 a WITH (NOLOCK)"
'wSQL = wSQL & "     , Web色規格別在庫 b WITH (NOLOCK)"
'wSQL = wSQL & "     , メーカー c WITH (NOLOCK)"
'wSQL = wSQL & "     , 商品カテゴリー d WITH (NOLOCK)"
'wSQL = wSQL & " WHERE b.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND b.商品コード = a.商品コード"
'wSQL = wSQL & "   AND c.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND d.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND d.商品コード = a.商品コード"
'wSQL = wSQL & "   AND d.カテゴリーコード = '" & CategoryCd & "'"
'wSQL = wSQL & "   AND a.お勧め商品フラグ = 'Y'"
'wSQL = wSQL & "   AND a.取扱中止日 IS NULL"
'wSQL = wSQL & "   AND ((a.廃番日 IS NULL AND b.終了日 IS NULL) OR (a.廃番日 IS NOT NULL AND b.引当可能数量 > 0)) "
'wSQL = wSQL & "   AND a.Web商品フラグ = 'Y'"
'wSQL = wSQL & "   AND b.終了日 IS NULL"
'wSQL = wSQL & " ORDER BY"
'wSQL = wSQL & "       d.カテゴリー区分"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "       	WHEN a.お勧め商品表示順 = 0 THEN 99999"
'wSQL = wSQL & "       	ELSE a.お勧め商品表示順"
'wSQL = wSQL & "       END"
'wSQL = wSQL & "     , c.メーカー名"
'wSQL = wSQL & "     , a.商品名"
'wSQL = wSQL & "     , b.色"
'wSQL = wSQL & "     , b.規格"
wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.メーカーコード "
wSQL = wSQL & "    , a.商品コード "
'wSQL = wSQL & "    , b.色 " 2012/07/17 natori del
'wSQL = wSQL & "    , b.規格 " 2012/07/17 natori del
wSQL = wSQL & "    , a.商品名 "
wSQL = wSQL & "    , a.商品画像ファイル名_小 "
wSQL = wSQL & "    , a.お勧め商品コメント "
'wSQL = wSQL & "    , a.ASK商品フラグ " 2012/07/17 natori del
'wSQL = wSQL & "    , a.試聴フラグ " 2012/07/17 natori del
'wSQL = wSQL & "    , a.試聴URL " 2012/07/17 natori del
'wSQL = wSQL & "    , a.動画フラグ " 2012/07/17 natori del
'wSQL = wSQL & "    , a.動画URL " 2012/07/17 natori del
'wSQL = wSQL & "    , CASE "
'wSQL = wSQL & "        WHEN (a.個数限定数量 > a.個数限定受注済数量 AND a.個数限定数量 > 0) THEN a.個数限定単価 "
'wSQL = wSQL & "        ELSE a.販売単価 "
'wSQL = wSQL & "      END AS 販売単価 " 2012/07/17 natori del
'wSQL = wSQL & "    , a.個数限定数量 " 2012/07/17 natori del
'wSQL = wSQL & "    , a.個数限定受注済数量 " 2012/07/17 natori del
'wSQL = wSQL & "    , CASE "
'wSQL = wSQL & "        WHEN (a.個数限定数量 > a.個数限定受注済数量 AND a.個数限定数量 > 0) THEN 'Y' "
'wSQL = wSQL & "        ELSE 'N' "
'wSQL = wSQL & "      END AS 個数限定単価フラグ " 2012/07/17 natori del
'wSQL = wSQL & "    , a.希少数量 " 2012/07/17 natori del
'wSQL = wSQL & "    , a.セット商品フラグ " 2012/07/17 natori del
'wSQL = wSQL & "    , a.メーカー直送取寄区分 " 2012/07/17 natori del
'wSQL = wSQL & "    , a.Web納期非表示フラグ " 2012/07/17 natori del
'wSQL = wSQL & "    , a.廃番日 " 2012/07/17 natori del
'wSQL = wSQL & "    , a.B品フラグ " 2012/07/17 natori del
'wSQL = wSQL & "    , a.入荷予定未定フラグ " 2012/07/17 natori del
'wSQL = wSQL & "    , b.引当可能入荷予定日 " 2012/07/17 natori del
'wSQL = wSQL & "    , b.引当可能数量 " 2012/07/17 natori del
'wSQL = wSQL & "    , b.B品引当可能数量 " 2012/07/17 natori del
wSQL = wSQL & "    , c.メーカー名 "
wSQL = wSQL & "    , COUNT(b.色) AS 色規格数 "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    Web商品                      a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN Web色規格別在庫 b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.メーカーコード = a.メーカーコード "
wSQL = wSQL & "           AND b.商品コード     = a.商品コード "
wSQL = wSQL & "      INNER JOIN メーカー        c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.メーカーコード = a.メーカーコード "
wSQL = wSQL & "      INNER JOIN 商品カテゴリー  d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.メーカーコード = a.メーカーコード "
wSQL = wSQL & "           AND d.商品コード     = a.商品コード "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' )   t1 "
wSQL = wSQL & "        ON     a.Web商品フラグ    = t1.ShohinWebY "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'RecommendY' )   t2 "
wSQL = wSQL & "        ON     a.お勧め商品フラグ = t2.RecommendY "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        t1.ShohinWebY  IS NOT NULL "
wSQL = wSQL & "    AND t2.RecommendY  IS NOT NULL "
wSQL = wSQL & "    AND a.取扱中止日 IS NULL "
wSQL = wSQL & "    AND (    (    a.廃番日 IS NULL "
wSQL = wSQL & "              AND b.終了日 IS NULL) "
wSQL = wSQL & "         OR  (    a.廃番日 IS NOT NULL "
wSQL = wSQL & "              AND b.引当可能数量 > 0)) "
wSQL = wSQL & "    AND b.終了日 IS NULL "
wSQL = wSQL & "    AND d.カテゴリーコード = '" & CategoryCd & "' "
'---- 色・規格などデータ不要のため、GROUP BY句を追加(2012/7/17 natori add)
wSQL = wSQL & "GROUP BY "
wSQL = wSQL & "      a.メーカーコード "
wSQL = wSQL & "    , a.商品コード "
wSQL = wSQL & "    , a.商品名 "
wSQL = wSQL & "    , a.商品画像ファイル名_小 "
wSQL = wSQL & "    , a.お勧め商品コメント "
wSQL = wSQL & "    , c.メーカー名 "
wSQL = wSQL & "    , d.カテゴリー区分 "
wSQL = wSQL & "    , a.お勧め商品表示順 "
wSQL = wSQL & "ORDER BY "
wSQL = wSQL & "      d.カテゴリー区分 "
wSQL = wSQL & "    , CASE "
wSQL = wSQL & "        WHEN a.お勧め商品表示順 = 0 THEN 99999 "
wSQL = wSQL & "        ELSE                             a.お勧め商品表示順 "
wSQL = wSQL & "      END "
wSQL = wSQL & "    , c.メーカー名 "
wSQL = wSQL & "    , a.商品名 "
'wSQL = wSQL & "    , b.色 " 2012/07/17 natori del
'wSQL = wSQL & "    , b.規格 " 2012/07/17 natori del
' 2012/01/19 GV Mod End

'@@@@@@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

wHTML = ""
'wHTML = wHTML & "<table border='0' cellspacing='1' cellpadding='0'>" & vbNewLine	2012/07/17 natori del

Do Until RS.EOF = true
	wHTML = wHTML & "<ul class='item_list gridtype productguide'>" & vbNewLine

	'---- 左 商品
	Call CreateProductHTML

	if RS.EOF = false then
	'---- 中央左 商品
		call CreateProductHTML()

		if RS.EOF = false then
		'---- 中央右 商品
			call CreateProductHTML()

			if RS.EOF = false then
			'---- 右 商品
				call CreateProductHTML()
			end if
		end if
	end if

	wHTML = wHTML & "</ul>"
Loop

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
'2012/07/17 natori del Start
'Dim vIroKikakuFl
'Dim vWidth1
'Dim vWidth2
'Dim vItemCnt
'Dim vItemList()
'Dim vSoundMovie
'Dim vInventoryCd
'Dim vInventoryImage

'vIroKikakuFl = false
'vWidth1 = 110
'vWidth2 = 85

'if Trim(RS("色")) <> "" OR Trim(RS("規格")) <> "" then
'	vIroKikakuFl = true
'	vWidth1 = 160
'	vWidth2 = 35
'end if

'wHTML = wHTML & "    <td>" & vbNewLine
'wHTML = wHTML & "      <table width='200' border='1' cellspacing='0' cellpadding='0' bordercolor='#999999' bordercolorlight='#999999' bordercolordark='#ffffff'>" & vbNewLine

'---- メーカー名，商品番号
'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td height='39' colspan='2' bgcolor='#eeeeee'>" & vbNewLine
'wHTML = wHTML & "            <span class='honbun'><b>" & RS("メーカー名") & "</b></span> "
'if vIroKikakuFl = false then
'	wHTML = wHTML & "<a href='ProductDetail.asp?item=" & RS("メーカーコード") & "^" & Server.URLEncode(RS("商品コード")) & "' class='link'><b>" & RS("商品名") & "</b></a>" & vbNewLine
'else
'	wHTML = wHTML & "<span class='honbun'><b>" & RS("商品名") & "</b></span>" & vbNewLine
'end if
'wHTML = wHTML & "          </td>" & vbNewLine
'wHTML = wHTML & "        </tr>" & vbNewLine

'---- 商品画像
'wHTML = wHTML & "        <tr align='center'>" & vbNewLine
'wHTML = wHTML & "          <td height='100' colspan='2'>" & vbNewLine
'if vIroKikakuFl = false then
'	wHTML = wHTML & "            <a href='ProductDetail.asp?item=" & RS("メーカーコード") & "^" & Server.URLEncode(RS("商品コード")) & "' class='link'>"
'end if

'if RS("商品画像ファイル名_小") <> "" then
'	wHTML = wHTML & "<img src='prod_img/" & RS("商品画像ファイル名_小") & "' width='198' height='99' border='0'>"
'else
'	wHTML = wHTML & "<img src='images/blank.gif' width='198' height='99' border='0'>"
'end if
'if vIroKikakuFl = false then
'	wHTML = wHTML & "</a>" & vbNewLine
'end if

'wHTML = wHTML & "          </td>" & vbNewLine
'wHTML = wHTML & "        </tr>" & vbNewLine
'2012/07/17 natori del End
'---- お勧め商品説明
if Trim(RS("お勧め商品コメント")) <> "" Then
	vComment = Replace(RS("お勧め商品コメント"), vbNewLine, "<br>")
End If
if vComment = "" then
	vComment = "&nbsp;"
end if

wHTML = wHTML & " <li>" & vbNewLine
wHTML = wHTML & "  <div class='photo'><a href='ProductDetail.asp?item=" & RS("メーカーコード") & "^" & Server.URLEncode(RS("商品コード")) & "'>"
If RS("商品画像ファイル名_小") <> "" Then
	wHTML = wHTML & "<img src='prod_img/" & RS("商品画像ファイル名_小") & "' alt='" & Replace(RS("メーカー名") & " / " & RS("商品名"),"'","&#39;") & "' class='opover'>"
End If
wHTML = wHTML & "</a></div>" & vbNewLine
wHTML = wHTML & "  <ul class='detail'>" & vbNewLine
wHTML = wHTML & "   <li><strong>" & RS("メーカー名") & "</strong></li>" & vbNewLine
wHTML = wHTML & "   <li class='prod_name'><strong><a href='ProductDetail.asp?item=" & RS("メーカーコード") & "^" & Server.URLEncode(RS("商品コード")) & "'>" & RS("商品名") & "</a></strong></li>" & vbNewLine
wHTML = wHTML & "   <li>" & vComment & "</li>"
wHTML = wHTML & "  </ul>" & vbNewLine
wHTML = wHTML & "  <div class='other_detail'>" & vbNewLine
wHTML = wHTML & "   <a href='ProductDetail.asp?item=" & RS("メーカーコード") & "^" & Server.URLEncode(RS("商品コード")) & "'><img src='images/btn_detail.png' alt='詳細' class='opover'></a>" & vbNewLine
If RS("色規格数") = 1 Then
	wHTML = wHTML & "   <a href='OrderPreInsert.asp?maker_cd=" & RS("メーカーコード") & "&amp;product_cd=" & Server.URLEncode(RS("商品コード")) & "&amp;qt=1'><img src='images/btn_cart.png' alt='カートに入れる' class='opover'></a>" & vbNewLine
End If
wHTML = wHTML & "  </div>" & vbNewLine
wHTML = wHTML & " </li>" & vbNewLine

RS.MoveNext

'2012/07/17 natori del Start
'wHTML = wHTML & "        <tr align='left' valign='top'>" & vbNewLine
'wHTML = wHTML & "          <td height='75' colspan='2' class='honbun'>" & vbNewLine
'wHTML = wHTML & vComment & vbNewLine
'wHTML = wHTML & "          </td>" & vbNewLine
'wHTML = wHTML & "        </tr>" & vbNewLine

'----試聴リンク
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

'----動画リンク
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

'if vSoundMovie <> "" then
'	wHTML = wHTML & "        <tr align='left' valign='middle'>" & vbNewLine
'	wHTML = wHTML & "          <td height='25' colspan='2' class='honbun'>" & vbNewLine
'	wHTML = wHTML & "            サンプル： " & vSoundMovie & vbNewLine
'	wHTML = wHTML & "          </td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine
'end if

'vOldProductCd = RS("メーカーコード") & "+" & RS("商品コード")

'---- 同一商品終了まで繰り返し (色規格がある場合のみ繰り返し)
'Do until vOldProductCd <> RS("メーカーコード") & "+" & RS("商品コード")
	'---- 色規格
'	wHTML = wHTML & "        <tr>" & vbNewLine
'	wHTML = wHTML & "          <td  height='25' align='right' nowrap colspan='2'>" & vbNewLine
'	if Trim(RS("色")) <> "" OR Trim(RS("規格")) <> "" then
'		wHTML = wHTML & "            <a href='ProductDetail.asp?item=" & RS("メーカーコード") & "^" & Server.URLEncode(RS("商品コード")) & "^" & RS("色") & "^" & RS("規格") & "' class='link'>"
'		if RS("色") <> "" then
'			wHTML = wHTML & RS("色") & " "
'		end if
'		if RS("規格") <> "" then
'			wHTML = wHTML & RS("規格")
'		end if
'		wHTML = wHTML & "</a>" & vbNewLine
'	end if

	'---- 単価
'	wPrice = calcPrice(RS("販売単価"), wSalesTaxRate)

'	if RS("ASK商品フラグ") = "Y" then
'2011/10/19 hn mod s
'		wHTML = wHTML & "         <span class='honbun'>衝撃特価：<a href='JavaScript:void(0);' onClick=""askWin=window.open('AskPrice.asp?MakerName=" & Server.URLEncode(RS("メーカー名")) & "&ProductName=" & Server.URLEncode(RS("商品名")) & "&Price=" & wPrice & "' ,'ask', 'width=250 height=80 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');"" class='link'>ASK</a></span>" & vbNewLine

'		wHTML = wHTML & "            <span class='honbun'>衝撃特価：<a class='tip'>ASK<span>" & FormatNumber(wPrice,0) & "円(税込)</span></a></span>" & vbNewLine
'2011/10/19 hn mod e

'	else
'		wHTML = wHTML & "            <span class='honbun'>衝撃特価：<b>" & FormatNumber(wPrice,0) & "円(税込)</b></span>" & vbNewLine

'	end if

'	wHTML = wHTML & "          </td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine

'----- 在庫状況
'	vInventoryCd = GetInventoryStatus(RS("メーカーコード"),RS("商品コード"),RS("色"),RS("規格"),RS("引当可能数量"),RS("希少数量"),RS("セット商品フラグ"),RS("メーカー直送取寄区分"),RS("引当可能入荷予定日"),"N")

	'---- 在庫状況、色を最終セット
'	call GetInventoryStatus2(RS("引当可能数量"), RS("Web納期非表示フラグ"), RS("入荷予定未定フラグ"), RS("廃番日"), RS("B品フラグ"), RS("B品引当可能数量"), RS("個数限定数量"), RS("個数限定受注済数量"), "N", vInventoryCd, vInventoryImage)

'	wHTML = wHTML & "        <tr>" & vbNewLine
'	wHTML = wHTML & "          <td width='" & vWidth1 & "' height='25' align='right' nowrap class='honbun'><img src='images/" & vInventoryImage & "' width=10 height=10> " & vInventoryCd & "</td>" & vbNewLine

	'---- 詳細ボタン，カートボタン
'	wHTML = wHTML & "          <td width='" & vWidth2 & "' nowrap height='25' align='center' valign='middle'>" & vbNewLine
'	if vIroKikakuFl = false then
'		wHTML = wHTML & "            <a href='ProductDetail.asp?item=" & RS("メーカーコード") & "^" & Server.URLEncode(RS("商品コード")) & "^" & RS("色") & "^" & RS("規格") & "'><img src='images/Shousai.gif' width='50' height='19' border='0' align='middle'></a>" & vbNewLine
'	end if
'	wHTML = wHTML & "            <a href='OrderPreInsert.asp?maker_cd=" & RS("メーカーコード") & "&product_cd=" & Server.URLEncode(RS("商品コード")) & "&iro=" & RS("色") & "&kikaku=" & RS("規格") & "&qt=1'><img src='images/CartBlue.gif' width='30' height='19' border='0' align='middle'></a>" & vbNewLine
'	wHTML = wHTML & "          </td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine

'	RS.MoveNext
'	if RS.EOF = true then
'		Exit Do
'	end if
'Loop

'wHTML = wHTML & "      </table>" & vbNewLine
'wHTML = wHTML & "    </td>" & vbNewLine
'2012/07/17 natori del End

End function

'========================================================================
'
'	Function	お勧め商品情報編集2 （シリーズ商品)
'
'========================================================================
'
Function CreateProductList2()

'---- 商品Recordset作成
wSQL = ""
' 2012/01/19 GV Mod Start
'wSQL = wSQL & "SELECT a.メーカーコード"
'wSQL = wSQL & "     , a.商品コード"
'wSQL = wSQL & "     , a.商品名"
'wSQL = wSQL & "     , a.商品画像ファイル名_小"
'wSQL = wSQL & "     , a.お勧め商品コメント"
'wSQL = wSQL & "     , a.ASK商品フラグ"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN (a.個数限定数量 > a.個数限定受注済数量 AND a.個数限定数量 > 0) THEN a.個数限定単価"
'wSQL = wSQL & "         ELSE a.販売単価"
'wSQL = wSQL & "       END AS 販売単価"
'wSQL = wSQL & "     , c.メーカー名"
'wSQL = wSQL & "     , e.シリーズコード"
'wSQL = wSQL & "     , e.シリーズ名"
'wSQL = wSQL & "     , e.シリーズ画像ファイル名"
'wSQL = wSQL & "     , e.お勧めシリーズ備考"
'wSQL = wSQL & "     , e.お勧めシリーズ表示順"
'wSQL = wSQL & "  FROM Web商品 a WITH (NOLOCK)"
'wSQL = wSQL & "     , Web色規格別在庫 b WITH (NOLOCK)"
'wSQL = wSQL & "     , メーカー c WITH (NOLOCK)"
'wSQL = wSQL & "     , 商品カテゴリー d WITH (NOLOCK)"
'wSQL = wSQL & "     , シリーズ e WITH (NOLOCK)"
'wSQL = wSQL & " WHERE b.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND b.商品コード = a.商品コード"
'wSQL = wSQL & "   AND c.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND d.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND d.商品コード = a.商品コード"
'wSQL = wSQL & "   AND e.シリーズコード = a.シリーズコード"
'wSQL = wSQL & "   AND d.カテゴリーコード = '" & CategoryCd & "'"
'wSQL = wSQL & "   AND e.お勧めシリーズフラグ = 'Y'"
'wSQL = wSQL & "   AND a.取扱中止日 IS NULL"
'wSQL = wSQL & "   AND ((a.廃番日 IS NULL) OR (a.廃番日 IS NOT NULL AND b.引当可能数量 > 0)) "
'wSQL = wSQL & "   AND a.Web商品フラグ = 'Y'"
'wSQL = wSQL & " ORDER BY"
''wSQL = wSQL & "       d.カテゴリー区分"       '2010/11/10 an del
'wSQL = wSQL & "       e.お勧めシリーズ表示順"
'wSQL = wSQL & "     , e.シリーズコード"        '2010/11/10 an add
'wSQL = wSQL & "     , 販売単価"
wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.メーカーコード "
wSQL = wSQL & "    , a.商品コード "
wSQL = wSQL & "    , a.商品名 "
wSQL = wSQL & "    , a.商品画像ファイル名_小 "
wSQL = wSQL & "    , a.お勧め商品コメント "
wSQL = wSQL & "    , a.ASK商品フラグ "
wSQL = wSQL & "    , CASE "
wSQL = wSQL & "        WHEN (a.個数限定数量 > a.個数限定受注済数量 AND a.個数限定数量 > 0) THEN a.個数限定単価 "
wSQL = wSQL & "        ELSE a.販売単価 "
wSQL = wSQL & "      END AS 販売単価 "
wSQL = wSQL & "    , c.メーカー名 "
wSQL = wSQL & "    , e.シリーズコード "
wSQL = wSQL & "    , e.シリーズ名 "
wSQL = wSQL & "    , e.シリーズ画像ファイル名 "
wSQL = wSQL & "    , e.お勧めシリーズ備考 "
wSQL = wSQL & "    , e.お勧めシリーズ表示順 "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    Web商品                      a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN Web色規格別在庫 b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.メーカーコード = a.メーカーコード "
wSQL = wSQL & "           AND b.商品コード     = a.商品コード "
wSQL = wSQL & "      INNER JOIN メーカー        c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.メーカーコード = a.メーカーコード "
wSQL = wSQL & "      INNER JOIN 商品カテゴリー  d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.メーカーコード = a.メーカーコード "
wSQL = wSQL & "           AND d.商品コード     = a.商品コード "
wSQL = wSQL & "      INNER JOIN シリーズ        e WITH (NOLOCK) "
wSQL = wSQL & "        ON     e.シリーズコード = a.シリーズコード "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' )         t1 "
wSQL = wSQL & "        ON     a.Web商品フラグ    = t1.ShohinWebY "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'RecommendSeriesY' )   t2 "
wSQL = wSQL & "        ON     e.お勧めシリーズフラグ = t2.RecommendSeriesY "
wSQL = wSQL & "WHERE"
wSQL = wSQL & "        t1.ShohinWebY       IS NOT NULL "
wSQL = wSQL & "    AND t2.RecommendSeriesY IS NOT NULL "
wSQL = wSQL & "    AND a.取扱中止日 IS NULL "
wSQL = wSQL & "    AND (    (    a.廃番日 IS NULL) "
wSQL = wSQL & "         OR  (    a.廃番日 IS NOT NULL "
wSQL = wSQL & "              AND b.引当可能数量 > 0)) "
wSQL = wSQL & "    AND d.カテゴリーコード = '" & CategoryCd & "' "
wSQL = wSQL & "ORDER BY "
wSQL = wSQL & "      e.お勧めシリーズ表示順 "
wSQL = wSQL & "    , e.シリーズコード "
wSQL = wSQL & "    , 販売単価 "
' 2012/01/19 GV Mod End

'@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

wHTML = ""
'wHTML = wHTML & "<table border='0' cellspacing='1' cellpadding='0'>" & vbNewLine

Do Until RS.EOF = true
	wHTML = wHTML & "<ul class='item_list gridtype productguide'>" & vbNewLine

	'---- 左 商品
	Call CreateProduct2HTML

	if RS.EOF = false then
	'---- 中央左 商品
		call CreateProduct2HTML()

		if RS.EOF = false then
		'---- 中央右 商品
			call CreateProduct2HTML()

			if RS.EOF = false then
			'---- 右 商品
				call CreateProduct2HTML()
			end if
		end if
	end if

	wHTML = wHTML & "</ul>" & vbNewLine
Loop

wProductList2 = wHTML

RS.Close

End function

'========================================================================
'
'	Function	個別商品HTML作成2 （シリーズ商品）
'
'========================================================================
'
Function CreateProduct2HTML()
Dim vComment
Dim vOldSeriesCd

'---- お勧めシリーズ説明
vComment = Replace(RS("お勧めシリーズ備考"), vbNewLine, "<br>")
if vComment = "" then
	vComment = "&nbsp;"
end if

wHTML = wHTML & "<li>" & vbNewLine
wHTML = wHTML & " <div class='photo'><a href='SearchList.asp?i_type=se&amp;sSeriesCd=" & RS("シリーズコード") & "'>"
If RS("シリーズ画像ファイル名") <> "" Then
	wHTML = wHTML & "<img src='prod_img/" & RS("シリーズ画像ファイル名") & "' alt='" & Replace(RS("メーカー名") & " / " & RS("シリーズ名"),"'","&#39;") & "'" & " class='opover'>"
End If
wHTML = wHTML & "</a></div>" & vbNewLine
wHTML = wHTML & " <ul class='detail'>" & vbNewLine
wHTML = wHTML & "  <li><strong>" & RS("メーカー名") & "</strong></li>" & vbNewLine
wHTML = wHTML & "  <li class='prod_name'><strong><a href='SearchList.asp?i_type=se&amp;sSeriesCd=" & RS("シリーズコード") & "'>" & RS("シリーズ名") & "</a>" & vbNewLine & "</strong></li>" & vbNewLine
wHTML = wHTML & "  <li>" & vComment & "</li>" & vbNewLine
wHTML = wHTML & " </ul>" & vbNewLine
wHTML = wHTML & " <div class='other_detail'><a href='SearchList.asp?i_type=se&amp;sSeriesCd=" & RS("シリーズコード") & "'><img src='images/btn_alllist.png' alt='一覧' class='opover'></a></div>" & vbNewLine
wHTML = wHTML & "</li>" & vbNewLine

'---- 同一シリーズ終了まで繰り返し（同一シリーズ集約）
vOldSeriesCd = RS("シリーズコード")
Do until vOldSeriesCd <> RS("シリーズコード")
	RS.MoveNext
	if RS.EOF = true then
		Exit Do
	end if
Loop

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
<meta name="description" content="<%=wCategoryComment%>">
<meta name="keywords" content="<%=wLargeCategoryName%>,<%=wMidCategoryName%>,<%=wCategoryName%>">
<title><%=wCategoryName%>のおすすめ商品｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/shop.css" type="text/css">
<link rel="stylesheet" href="Style/searchlist.css?20120811" type="text/css">
<link rel="stylesheet" href="style/ask.css" type="text/css">
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
				<li><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='LargeCategoryList.asp?LargeCategoryCd=<%=wLargeCategoryCd%>' itemprop='url'><span itemprop='title'><%=wLargeCategoryName%></span></a></span></li>
				<li><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='MidCategoryList.asp?MidCategoryCd=<%=wMidCategoryCd%>' itemprop='url'><span itemprop='title'><%=wMidCategoryName%></span></a></span></li>
				<li><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='SearchList.asp?i_type=c&amp;s_category_cd=<%=CategoryCd%>' itemprop='url'><span itemprop='title'><%=wCategoryName%></span></a></span></li>
				<li class="now"><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><span itemprop='title'>おすすめ商品</span></span></li>
			</ul>
			</div></div></div>

			<h1 class="title"><%=wCategoryName%> のおすすめ商品</h1>
			<p><%=wCategoryComment%></p>

			<!-- お勧めシリーズ商品一覧-->
			<%=wProductList2%>

			<!-- お勧め商品一覧-->
			<%=wProductList%>
		<!--/#contents -->
		</div>
		<div id="globalSide">
			<!--#include file="../Navi/NaviSide.inc"-->
		<!--/#globalSide -->
		</div>
	<!--/#main -->
	</div>
	<!--#include file="../Navi/Navibottom.inc"-->
	<!--#include file="../Navi/NaviScript.inc"-->
	<div class="tooltip"><p>ASK</p></div>
	<script type="text/javascript" src="jslib/ask.js"></script>
	<script type="text/javascript" src="jslib/SearchList.js?20120321" charset="Shift_JIS"></script>
	<script type="text/javascript" src="../jslib/jquery.tinyscrollbar.min.js"></script>
	<script type="text/javascript">
		$(function(){
		    $('#scrollbar1').tinyscrollbar();
		});
	</script>
</body>
</html>