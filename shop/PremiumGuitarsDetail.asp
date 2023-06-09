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
'	商品詳細ページ
'更新履歴
'2009/09/08 an 新規作成
'2011/04/14 hn SessionID関連変更
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/01/20 an SELECT文へLACクエリー案を適用
'2012/02/20 na 「商品アクセス数」レスポンス対策のため停止
'2014/03/19 GV 消費税増税に伴う2重表示対応

'========================================================================

On Error Resume Next

Dim wUserID

Dim maker_cd
Dim product_cd

Dim item
Dim item_list()
Dim item_cnt

Dim wMakerName
Dim wProductName
Dim wMakerCode
Dim wCategoryCode
Dim wKoukeiMakerCd
Dim wKoukeiProductCd
Dim wTokucho

Dim wLogoHTML
Dim wProductHTML
Dim wTokuchoHTML
Dim wPictureHTML
Dim wSpecHTML
Dim wOthersHTML
Dim wCartHTML

Dim Connection
Dim RS

Dim wMinimumPrice
Dim wMidCategoryCd
Dim wSalesTaxRate
Dim wProdTermFl
Dim wPrice
Dim wPriceNoTax			'2014/03/19 GV add
Dim wNoData

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim wSQL
Dim wHTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.buffer = true

wUserID = Session("UserID")

'---- Get input data
item = ReplaceInput(Trim(Request("item")))

'メーカーコード、商品コードに分解
if item <> "" then 
	item_cnt = cf_unstring(item, item_list, "^")
	maker_cd = item_list(0)
	product_cd = item_list(1)
end if

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "PremiumGuitarsDetail.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

if Err.Description <> "" then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

call close_db()

'---- 該当商品なしのとき
if wNoData = "Y" then
	Response.Redirect "SearchNotFound.asp"
end if

'---- 後継機種がある場合はその商品を表示
if wKoukeiMakerCd <> "" then
	Response.Redirect "SearchList.asp?i_type=successor&s_maker_cd=" & wKoukeiMakerCd & "&s_product_cd=" & Server.URLEncode(wKoukeiProductCd)
end if


'========================================================================
'
'	Function	Connect database
'
'========================================================================
'
Function connect_db()
Dim i

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

Dim vPointer

'---- 対象カテゴリーコード、最低単価取出し
call getCntlMst("商品","PuremiumGuitar","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
wMinimumPrice = Clng(wItemNum1)
wMidCategoryCd = wItemChar1

'---- 消費税率取出し
call getCntlMst("共通","消費税率","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
wSalesTaxRate = Clng(wItemNum1)

'---- 商品情報取り出し
call GetProduct()
if wMSG <> "" OR wKoukeiMakerCd <> "" OR wNoData = "Y" then
	exit function
end if

'---- メーカーロゴ、商品名、商品画像大
call CreateLogoHTML()

'---- メーカー/商品名
Call CreateProductHTML()

'---- 特徴
call CreateTokuchoHTML()

'---- 商品画像小
call CreatePictureHTML()

'---- スペック
call CreateSpecificationHTML()

'---- その他の商品を見る
call CreateOthersHTML()

'----- カート情報HTML作成（共通関数）
wCartHTML = fCreateCartHtml()

'----- 商品アクセス件数登録 '2012/02/20 na レスポンス対策のため停止
'call SetAccessCount()

RS.Close

End function

'========================================================================
'
'	Function	商品情報取り出し
'
'========================================================================
'
Function GetProduct()

Dim vInventoryCd
Dim vInventoryImage
Dim vSetCount
Dim vRowspan
Dim vWidth
Dim vHeight

Dim vProdPic(4)

'---- 商品情報取り出し
wSQL = ""
wSQL = wSQL & "SELECT DISTINCT a.商品コード"
wSQL = wSQL & "     , a.商品名"
wSQL = wSQL & "     , a.販売単価"
wSQL = wSQL & "     , a.個数限定単価"
wSQL = wSQL & "     , a.個数限定数量"
wSQL = wSQL & "     , a.個数限定受注済数量"
wSQL = wSQL & "     , a.商品備考"
wSQL = wSQL & "     , a.商品概略Web"
wSQL = wSQL & "     , a.商品画像ファイル名_大"
wSQL = wSQL & "     , a.商品画像ファイル名_小2"
wSQL = wSQL & "     , a.商品画像ファイル名_小3"
wSQL = wSQL & "     , a.商品画像ファイル名_小4"
wSQL = wSQL & "     , a.ASK商品フラグ"
wSQL = wSQL & "     , a.終了日"
wSQL = wSQL & "     , a.取扱中止日"
wSQL = wSQL & "     , a.廃番日"
wSQL = wSQL & "     , a.Web商品フラグ"
wSQL = wSQL & "     , a.カテゴリーコード"
wSQL = wSQL & "     , a.メーカーコード"
wSQL = wSQL & "     , a.希少数量"
wSQL = wSQL & "     , a.セット商品フラグ"
wSQL = wSQL & "     , a.メーカー直送取寄区分"
wSQL = wSQL & "     , a.直輸入品フラグ"
wSQL = wSQL & "     , a.お勧め商品コメント"
wSQL = wSQL & "     , a.関連記事タイトル1"
wSQL = wSQL & "     , a.関連記事URL1"
wSQL = wSQL & "     , a.関連記事タイトル2"
wSQL = wSQL & "     , a.関連記事URL2"
wSQL = wSQL & "     , a.関連記事タイトル3"
wSQL = wSQL & "     , a.関連記事URL3"
wSQL = wSQL & "     , a.関連記事タイトル4"
wSQL = wSQL & "     , a.関連記事URL4"
wSQL = wSQL & "     , a.Web納期非表示フラグ"
wSQL = wSQL & "     , a.後継機種メーカーコード"
wSQL = wSQL & "     , a.後継機種商品コード"
wSQL = wSQL & "     , a.入荷予定未定フラグ"
wSQL = wSQL & "     , a.商品スペック使用不可フラグ"
wSQL = wSQL & "     , a.B品単価"
wSQL = wSQL & "     , a.完売日"
wSQL = wSQL & "     , a.B品フラグ"
wSQL = wSQL & "     , a.商品備考インサートURL1"
wSQL = wSQL & "     , a.商品備考インサートURL2"
wSQL = wSQL & "     , a.商品備考インサートサイズW1"
wSQL = wSQL & "     , a.商品備考インサートサイズH1"
wSQL = wSQL & "     , a.商品備考インサートサイズW2"
wSQL = wSQL & "     , a.商品備考インサートサイズH2"
wSQL = wSQL & "     , b.メーカー名"
wSQL = wSQL & "     , b.メーカー名カナ"
wSQL = wSQL & "     , b.メーカーロゴファイル名"
wSQL = wSQL & "     , b.メーカーホームページURL"
wSQL = wSQL & "     , b.詳細情報タイトル1"
wSQL = wSQL & "     , b.詳細情報URL1"
wSQL = wSQL & "     , b.詳細情報Web表示可フラグ1"
wSQL = wSQL & "     , b.詳細情報タイトル2"
wSQL = wSQL & "     , b.詳細情報URL2"
wSQL = wSQL & "     , b.詳細情報Web表示可フラグ2"
wSQL = wSQL & "     , b.詳細情報タイトル3"
wSQL = wSQL & "     , b.詳細情報URL3"
wSQL = wSQL & "     , b.詳細情報Web表示可フラグ3"
wSQL = wSQL & "     , b.詳細情報タイトル4"
wSQL = wSQL & "     , b.詳細情報URL4"
wSQL = wSQL & "     , b.詳細情報Web表示可フラグ4"
wSQL = wSQL & "     , c.カテゴリー名"
wSQL = wSQL & "     , d.中カテゴリーコード"
wSQL = wSQL & "     , d.中カテゴリー名日本語"
wSQL = wSQL & "     , e.大カテゴリーコード"
wSQL = wSQL & "     , e.大カテゴリー名"
wSQL = wSQL & "     , f.色"
wSQL = wSQL & "     , f.規格"
wSQL = wSQL & "     , f.引当可能数量"
wSQL = wSQL & "     , f.引当可能入荷予定日"
wSQL = wSQL & "     , f.B品引当可能数量"
wSQL = wSQL & "     , f.色規格商品画像ファイル名1"
wSQL = wSQL & "     , f.色規格商品画像ファイル名2"
wSQL = wSQL & "     , f.色規格商品画像ファイル名3"
wSQL = wSQL & "     , f.色規格商品画像ファイル名4"

'wSQL = wSQL & "  FROM Web商品 a WITH (NOLOCK)"     '2012/01/20 an mod s
'wSQL = wSQL & "     , メーカー b WITH (NOLOCK)"
'wSQL = wSQL & "     , カテゴリー c WITH (NOLOCK)"
'wSQL = wSQL & "     , 中カテゴリー d WITH (NOLOCK) "
'wSQL = wSQL & "     , 大カテゴリー e WITH (NOLOCK) "
'wSQL = wSQL & "     , Web色規格別在庫 f WITH (NOLOCK)"
'wSQL = wSQL & " WHERE b.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND c.カテゴリーコード = a.カテゴリーコード"
'wSQL = wSQL & "   AND d.中カテゴリーコード = c.中カテゴリーコード"
'wSQL = wSQL & "   AND e.大カテゴリーコード = d.大カテゴリーコード"
'wSQL = wSQL & "   AND a.Web商品フラグ = 'Y'"
'wSQL = wSQL & "   AND a.メーカーコード = '" & maker_cd & "'"
'wSQL = wSQL & "   AND a.商品コード = '" & product_cd & "'"
'wSQL = wSQL & "   AND f.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND f.商品コード = a.商品コード"
'wSQL = wSQL & "   AND f.色 = ''"
'wSQL = wSQL & "   AND f.規格 = ''"
'wSQL = wSQL & "   AND f.終了日 IS NULL"

wSQL = wSQL & "  FROM Web商品                   a WITH (NOLOCK)"
wSQL = wSQL & "      INNER JOIN メーカー        b WITH (NOLOCK)"
wSQL = wSQL & "        ON    b.メーカーコード = a.メーカーコード"
wSQL = wSQL & "      INNER JOIN カテゴリー      c WITH (NOLOCK)"
wSQL = wSQL & "        ON     c.カテゴリーコード = a.カテゴリーコード"
wSQL = wSQL & "      INNER JOIN 中カテゴリー    d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.中カテゴリーコード = c.中カテゴリーコード"
wSQL = wSQL & "      INNER JOIN 大カテゴリー    e WITH (NOLOCK) "
wSQL = wSQL & "        ON     e.大カテゴリーコード = d.大カテゴリーコード"
wSQL = wSQL & "      INNER JOIN Web色規格別在庫 f WITH (NOLOCK)"
wSQL = wSQL & "        ON     f.メーカーコード = a.メーカーコード"
wSQL = wSQL & "          AND  f.商品コード = a.商品コード"
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' ) t1 "
wSQL = wSQL & "        ON     a.Web商品フラグ      = t1.ShohinWebY "
wSQL = wSQL & "      LEFT JOIN ( SELECT ''  AS 'Iro' )        t2 "
wSQL = wSQL & "        ON     f.色               = t2.Iro "
wSQL = wSQL & "      LEFT JOIN ( SELECT ''  AS 'Kikaku' )     t3 "
wSQL = wSQL & "        ON     f.規格             = t3.Kikaku "
wSQL = wSQL & " WHERE "
wSQL = wSQL & "        t1.ShohinWebY   IS NOT NULL "
wSQL = wSQL & "    AND t2.Iro          IS NOT NULL "
wSQL = wSQL & "    AND t3.Kikaku       IS NOT NULL "
wSQL = wSQL & "    AND a.メーカーコード = '" & maker_cd & "'"
wSQL = wSQL & "    AND a.商品コード = '" & product_cd & "'"
wSQL = wSQL & "    AND f.終了日 IS NULL"    '2012/01/20 an mod e

'@@@@@@@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

if RS.EOF = true then
	wPictureHTML = "<p class='error'>該当商品はありません。</p>"
	wMSG = "no data"
	wNoData = "Y"
	exit function
end if

'---- 終了チェック
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

'---- 取扱中止または、廃番在庫無しで後継機種チェック あれば後継機種を表示
if wProdTermFl = "Y" AND RS("後継機種メーカーコード") <> "" then
	wKoukeiMakerCd = RS("後継機種メーカーコード")
	wKoukeiProductCd = RS("後継機種商品コード")
	exit function
end if

'----- メーカー名、商品名､ タイトル
wMakerName = RS("メーカー名")
wProductName = RS("商品名")
wMakerCode = RS("メーカーコード")

End Function

'========================================================================
'
'	Function	メーカーロゴ、商品名、商品画像大 HTML作成
'
'========================================================================
'
Function CreateLogoHTML()

wLogoHTML = ""
wLogoHTML = wLogoHTML & "<div id='pgProductNameBox'>" & vbNewLine
If RS("メーカーロゴファイル名") <> "" Then
	wLogoHTML = wLogoHTML & "  <img src='maker_img/" & RS("メーカーロゴファイル名") & "'>" & vbNewLine
End If
wLogoHTML = wLogoHTML & "  <h1><span>" & wMakerName & "</span>" & wProductName  & "</h1>" & vbNewLine
wLogoHTML = wLogoHTML & "  <a href='javascript:void(0);' onclick='history.back();' class='tipBtn'>BACK</a>" & vbNewLine
wLogoHTML = wLogoHTML & "</div>" & vbNewLine
wLogoHTML = wLogoHTML & "<div id='pgLargeImage'><img name='LargeImage' src='prod_img/"

if Trim(RS("色規格商品画像ファイル名1")) <> "" then
	wLogoHTML = wLogoHTML & RS("色規格商品画像ファイル名1") & "' alt=''></div>" & vbNewLine
else
	if RS("商品画像ファイル名_大") <> "" then
		wLogoHTML = wLogoHTML & RS("商品画像ファイル名_大") & "' alt=''></div>" & vbNewLine
	else
		wLogoHTML = wLogoHTML & "n/nopict.jpg' alt=''></div>" & vbNewLine
	end if
end if

End Function

'========================================================================
'
'	Function	商品情報 HTML作成
'
'========================================================================
'
Function CreateProductHTML()

Dim vInventoryCd
Dim vInventoryImage

'----- 在庫状況
vInventoryCd = GetInventoryStatus(RS("メーカーコード"),RS("商品コード"),RS("色"),RS("規格"),RS("引当可能数量"),RS("希少数量"),RS("セット商品フラグ"),RS("メーカー直送取寄区分"),RS("引当可能入荷予定日"),wProdTermFl)

'---- 在庫状況、色を最終セット
call GetInventoryStatus2(RS("引当可能数量"), RS("Web納期非表示フラグ"), RS("入荷予定未定フラグ"), RS("廃番日"), RS("B品フラグ"), RS("B品引当可能数量"), RS("個数限定数量"), RS("個数限定受注済数量"), wProdTermFl, vInventoryCd, vInventoryImage)

wProductHTML = ""
wProductHTML = wProductHTML & "        <h2>メーカー/商品名</h2>" & vbNewLine
wProductHTML = wProductHTML & "        <dl class='pgDetailBox'>" & vbNewLine
wProductHTML = wProductHTML & "          <dt>メーカー</dt>" & vbNewLine
wProductHTML = wProductHTML & "          <dd>" &  wMakerName & " ( " & RS("メーカー名カナ") & " )</dd>" & vbNewLine
wProductHTML = wProductHTML & "          <dt>商品名</dt>" & vbNewLine
wProductHTML = wProductHTML & "          <dd>" & wProductName & "</dd>" & vbNewLine
wProductHTML = wProductHTML & "          <dt>カテゴリー</dt>" & vbNewLine
wProductHTML = wProductHTML & "          <dd>" & RS("カテゴリー名") & "</dd>" & vbNewLine
wProductHTML = wProductHTML & "          <dt>販売価格</dt>" & vbNewLine
wProductHTML = wProductHTML & "          <dd>"

if RS("個数限定数量") > RS("個数限定受注済数量") AND RS("個数限定数量") > 0 then
	wPrice = calcPrice(RS("個数限定単価"), wSalesTaxRate)
	wPriceNoTax = RS("個数限定単価")						'2014/03/19 GV add
else
	wPrice = calcPrice(RS("販売単価"), wSalesTaxRate)
	wPriceNoTax = RS("販売単価")						'2014/03/19 GV add
end if

'2014/03/19 GV mod start ---->
'wProductHTML = wProductHTML & FormatNumber(wPrice,0) & "円(税込)</dd>" & vbNewLine
wProductHTML = wProductHTML & FormatNumber(wPriceNoTax,0) & "円(税抜)</dd>" & vbNewLine

wProductHTML = wProductHTML & "          <dt>&nbsp;</dt>" & vbNewLine
wProductHTML = wProductHTML & "<dd>" & FormatNumber(wPrice,0) & "円(税込)</dd>" & vbNewLine
'2014/03/19 GV mod end   <----
wProductHTML = wProductHTML & "          <dt>在庫状況</dt>" & vbNewLine

wProductHTML = wProductHTML & "          <dd><img src='images/" & vInventoryImage & "' width='10' height='10' style='vertical-align:baseline;'> " & vInventoryCd & "</dd>" & vbNewLine
wProductHTML = wProductHTML & "          <dt></dt>" & vbNewLine
wProductHTML = wProductHTML & "          " & vbNewLine
wProductHTML = wProductHTML & "          <dd id='cartBox'><form name='f_data' method='post' action='OrderPreInsert.asp' onSubmit='return order_onClick(this);'><input type='text' name='qt' size='2' maxsize='3' value='1'>" & vbNewLine

if (IsNull(RS("取扱中止日")) = false) OR (IsNull(RS("完売日")) = false) OR (RS("B品フラグ") = "Y" AND RS("B品引当可能数量") <= 0) OR (IsNull(RS("廃番日")) = false AND RS("引当可能数量") <= 0) then
	wProductHTML = wProductHTML & "<img src='images/Kanbai2.jpg' alt='完売'>" & vbNewLine
else
	wProductHTML = wProductHTML & "            <input type='image' src='images/PremiumGuitars/grey_cart.jpg' width='80' height='23' alt=''>" & vbNewLine
	wProductHTML = wProductHTML & "            <input type='hidden' name='Item' value='" & RS("メーカーコード") & "^" & RS("商品コード") & "'>" & vbNewLine
end if

wProductHTML = wProductHTML & "          </form>" & vbNewLine
wProductHTML = wProductHTML & "          </dd>" & vbNewLine
wProductHTML = wProductHTML & "        </dl>" & vbNewLine

End Function

'========================================================================
'
'	Function	特徴 HTML作成
'
'========================================================================
'
Function CreateTokuchoHTML()

wHTML = ""

'---- 特徴, 直輸入品表示
If RS("お勧め商品コメント") <> "" Or RS("直輸入品フラグ") = "Y" Then
	wTokuchoHTML = wTokuchoHTML & "        <h2>特徴</h2>" & vbNewLine
	wTokuchoHTML = wTokuchoHTML & "        <div class='pgDetailBox'>" & vbNewLine
	if RS("お勧め商品コメント") <> "" then
		wTokuchoHTML = wTokuchoHTML & RS("お勧め商品コメント") & "<br>" & vbNewLine

		'---- meta description用データ取得
		wTokucho = fDeleteHTMLTag(RS("お勧め商品コメント")) 'HTMLタグ削除
		wTokucho = replace(replace(replace(wTokucho, vbCr, ""), vbLf, ""), vbTab, "") '改行、Tabの削除

		if Len(wTokucho) > 97 then  '長い場合は100文字に省略
			wTokucho = Left(wTokucho,97) & "..."
		end if

	end if
	if RS("直輸入品フラグ") = "Y" then
		wTokuchoHTML = wTokuchoHTML & "<a href='../information/direct_import.asp'>[直輸入品]</a>" & vbNewLine
	end if
	wTokuchoHTML = wTokuchoHTML & "        </div>" & vbNewLine
End If

End Function

'========================================================================
'
'	Function	商品小画像 HTML作成
'
'========================================================================
'
Function CreatePictureHTML()

Dim vProdPic(4)
Dim i

'色規格商品画像ファイル名がある場合はそちらを優先
if Trim(RS("色規格商品画像ファイル名1")) <> "" then
	vProdPic(1) = RS("色規格商品画像ファイル名1")
	vProdPic(2) = RS("色規格商品画像ファイル名2")
	vProdPic(3) = RS("色規格商品画像ファイル名3")
	vProdPic(4) = RS("色規格商品画像ファイル名4")
else
	vProdPic(1) = RS("商品画像ファイル名_大")
	vProdPic(2) = RS("商品画像ファイル名_小2")
	vProdPic(3) = RS("商品画像ファイル名_小3")
	vProdPic(4) = RS("商品画像ファイル名_小4")
end if

wPictureHTML = ""
wPictureHTML = wPictureHTML & "      <ul id='pgSmallImage'>" & vbNewLine

for i=1 to 4
	wPictureHTML = wPictureHTML & "        <li><img src='prod_img/"
	if vProdPic(i) <> "" then
		wPictureHTML = wPictureHTML & vProdPic(i) & "' alt='" & wMakerName & " / " & wProductname & " 画像" & i & "' onMouseOver='SmallImage_onMouseOver(""prod_img/" & vProdPic(i) & """);'>"
	else
		'小画像がない場合は代替画像を表示
		wPictureHTML = wPictureHTML & "p/pg_photo_s.jpg' alt='" & wMakerName & " / " & wProductname & " 画像" & i & "'>"
	end if
	wPictureHTML = wPictureHTML & "        </li>" & vbNewLine
Next

wPictureHTML = wPictureHTML & "      </ul>" & vbNewLine

End Function

'========================================================================
'
'	Function	スペック HTML作成
'
'========================================================================
'
Function CreateSpecificationHTML()

Dim vWidth
Dim vHeight

wSpecHTML = ""

'---- スペック
wSpecHTML = wSpecHTML & "        <h2>スペック</h2>" & vbNewLine
wSpecHTML = wSpecHTML & "        <div class='pgDetailBox'>" & vbNewLine

if RS("商品備考インサートURL1") <> "" then
	if RS("商品備考インサートサイズW1") <> 0 then
		vWidth = RS("商品備考インサートサイズW1")
		if vWidth > 600 then
			vWidth = 600
		end if
	else
		vWidth = 600
	end if
	if RS("商品備考インサートサイズH1") <> 0 then
		vHeight = RS("商品備考インサートサイズH1")
		if vHeight > 290 then
			vHeight = 290
		end if
	else
		vHeight = 290
	end if
'	wSpecHTML = wSpecHTML & "<iframe marginwidth='0' marginheight='0' src='" & RS("商品備考インサートURL1") & "' frameborder='0' scrolling='no' style='PADDING-RIGHT: 0px; PADDING-LEFT: 0px; PADDING-BOTTOM: 0px; MARGIN: 0px; WIDTH: " & vWidth & "px; PADDING-TOP: 0px; HEIGHT: " & vHeight & "px'> </iframe>" & vbNewLine
end if

wSpecHTML = wSpecHTML & CreateSpecHTML(RS("カテゴリーコード"),RS("メーカーコード"),RS("商品コード"),RS("商品備考"),RS("商品スペック使用不可フラグ")) & vbNewLine

if RS("商品備考インサートURL2") <> "" then
	if RS("商品備考インサートサイズW2") <> 0 then
		vWidth = RS("商品備考インサートサイズW2")
	else
		vWidth = 600
	end if
	if RS("商品備考インサートサイズH2") <> 0 then
		vHeight = RS("商品備考インサートサイズH2")
	else
		vHeight = 300
	end if
'	wSpecHTML = wSpecHTML & "<iframe marginwidth='0' marginheight='0' src='" & RS("商品備考インサートURL2") & "' frameborder='0' scrolling='no' style='PADDING-RIGHT: 0px; PADDING-LEFT: 0px; PADDING-BOTTOM: 0px; MARGIN: 0px; WIDTH: " & vWidth & "px; PADDING-TOP: 0px; HEIGHT: " & vHeight & "px'> </iframe>" & vbNewLine
end if

wSpecHTML = wSpecHTML & "        </div>" & vbNewLine

End Function

'========================================================================
'
'	Function	その他の商品を見る HTML作成
'
'========================================================================
'
Function CreateOthersHTML

Dim RSv

'----同一メーカーの上位15品（初回登録日順） 取り出し
wSQL = ""
wSQL = wSQL & "SELECT DISTINCT TOP 15"
wSQL = wSQL & "   a.商品名"
wSQL = wSQL & " , a.商品コード"
wSQL = wSQL & " , a.初回登録日"

'wSQL = wSQL & " FROM  Web商品  a WITH (NOLOCK)"    '2012/01/20 an mod s
'wSQL = wSQL & "     , カテゴリー中カテゴリー b WITH (NOLOCK)"
'wSQL = wSQL & " WHERE (SELECT CASE"
'wSQL = wSQL & "                   WHEN x.個数限定数量 > x.個数限定受注済数量 THEN (x.個数限定単価 * (100 + " & wSalesTaxRate & " )/100)"
'wSQL = wSQL & "                   ELSE (x.販売単価 * (100 + " & wSalesTaxRate & " )/100)"
'wSQL = wSQL & "               END"
'wSQL = wSQL & "        FROM web商品 x "
'wSQL = wSQL & "        WHERE x.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "          AND x.商品コード = a.商品コード"
'wSQL = wSQL & "        ) > " & wMinimumPrice
'wSQL = wSQL & "    AND a.カテゴリーコード = b.カテゴリーコード"
'wSQL = wSQL & "    AND b.中カテゴリーコード IN (" & wMidCategoryCd & ")"
'wSQL = wSQL & "    AND a.Web商品フラグ = 'Y'"
'wSQL = wSQL & "    AND a.メーカーコード =" & RS("メーカーコード")

wSQL = wSQL & " FROM  Web商品                          a WITH (NOLOCK)"
wSQL = wSQL & "      INNER JOIN カテゴリー中カテゴリー b WITH (NOLOCK)"
wSQL = wSQL & "        ON     b.カテゴリーコード = a.カテゴリーコード"
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' ) t1 "
wSQL = wSQL & "        ON     a.Web商品フラグ      = t1.ShohinWebY "
wSQL = wSQL & " WHERE "
wSQL = wSQL & "        t1.ShohinWebY   IS NOT NULL "
wSQL = wSQL & "    AND (SELECT CASE"
wSQL = wSQL & "                   WHEN x.個数限定数量 > x.個数限定受注済数量 THEN (x.個数限定単価 * (100 + " & wSalesTaxRate & " )/100)"
wSQL = wSQL & "                   ELSE (x.販売単価 * (100 + " & wSalesTaxRate & " )/100)"
wSQL = wSQL & "               END"
wSQL = wSQL & "        FROM web商品 x WITH (NOLOCK)"
wSQL = wSQL & "        WHERE x.メーカーコード = a.メーカーコード"
wSQL = wSQL & "          AND x.商品コード = a.商品コード"
wSQL = wSQL & "        ) > " & wMinimumPrice
wSQL = wSQL & "    AND b.中カテゴリーコード IN (" & wMidCategoryCd & ")"
wSQL = wSQL & "    AND a.メーカーコード =" & RS("メーカーコード")
wSQL = wSQL & " ORDER BY a.初回登録日 DESC"    '2012/01/20 an mod e

'@@@@@@@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

	wOthersHTML = ""
	wOthersHTML = wOthersHTML & "<div id='pgOtherProduct'>" & vbNewLine
	wOthersHTML = wOthersHTML & "  <h2>その他の商品を見る</h2>" & vbNewLine
	wOthersHTML = wOthersHTML & "  <ul>" & vbNewLine
	Do Until RSv.EOF = true
		'--- 詳細表示中の商品はリストに表示しない
		if RSv("商品コード") <> product_cd then
			wOthersHTML = wOthersHTML & "    <li><a href='PremiumGuitarsDetail.asp?Item=" & RS("メーカーコード") & "^" & RSv("商品コード") & "' style='text-decoration:none; color:#cccccc;'>" & RSv("商品名") & "</a></li>" & vbNewLine
		end if
		RSv.MoveNext
	Loop
	wOthersHTML = wOthersHTML & "  </ul>" & vbNewLine
	wOthersHTML = wOthersHTML & "  <div class='more'><li><a href='PremiumGuitarsList.asp?MakerCd=" & RS("メーカーコード") & "'>" &  wMakerName & "<br>プレミアムギター一覧</a></li></div>" & vbNewLine
	wOthersHTML = wOthersHTML & "</div>" & vbNewLine
	
RSv.Close

End Function
'========================================================================
'
'	Function	最近チェックした商品に追加
'
'========================================================================
'
Function AddViewdProduct()

Dim RSv

wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM 最近チェックした商品"
wSQL = wSQL & " WHERE 顧客番号 = " & wUserID
wSQL = wSQL & "   AND メーカーコード = '" & maker_cd & "'"
wSQL = wSQL & "   AND 商品コード = '" & product_cd & "'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RSv.EOF = true then
	RSv.AddNew

	RSv("顧客番号") = wUserID
	RSv("メーカーコード") = maker_cd
	RSv("商品コード") = product_cd
end if

RSv("チェック日") = Now()

RSv.Update
RSv.close

End function

'========================================================================
'
'	Function	商品アクセスカウント登録（ページビュー）
'
'========================================================================
'
Function SetAccessCount()

Dim vYYYYMM
Dim RSv

'---- 同一セッションで1回目かどうかチェック
	wSQL = ""
	wSQL = wSQL & "SELECT *"
	wSQL = wSQL & "  FROM セッションデータ"
	wSQL = wSQL & " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
	wSQL = wSQL & "   AND 項目名 = '" & maker_cd & "^" & product_cd & "'"

	Set RSv = Server.CreateObject("ADODB.Recordset")
	RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

	if RSv.EOF = true then
		'---- セッションデータ登録
		RSv.AddNew

		RSv("SessionID") = gSessionID		'2011/04/14 hn mod
		RSv("項目名") = maker_cd & "^" & product_cd
		RSv("内容") = "ページビューチェック用"
		RSv("最終更新日") = Now()

		RSv.Update
		RSv.close

		'---- ページビュー登録
		vYYYYMM = Year(Now()) & Right("0" & Month(Now()),2)

		wSQL = ""
		wSQL = wSQL & "SELECT *"
		wSQL = wSQL & "  FROM 商品アクセス件数"
		wSQL = wSQL & " WHERE メーカーコード = '" & maker_cd & "'"
		wSQL = wSQL & "   AND 商品コード = '" & product_cd & "'"
		wSQL = wSQL & "   AND 年月 = '" & vYYYYMM & "'"

		Set RSv = Server.CreateObject("ADODB.Recordset")
		RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

		if RSv.EOF = true then
			RSv.AddNew

			RSv("メーカーコード") = maker_cd
			RSv("商品コード") = product_cd
			RSv("年月") = vYYYYMM
			RSv("ページビュー件数") = 1
		else
			RSv("ページビュー件数") = RSv("ページビュー件数") + 1
		end if

		RSv.Update
		RSv.close
	end if

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
<meta name="robots" content="noindex,nofollow">
<title>プレミアムギター <%=wMakerName%> / <%=wProductName%>｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link href='http://fonts.googleapis.com/css?family=Ovo' rel='stylesheet' type='text/css'>
<link rel="stylesheet" href="style/PremiumGuitars.css" type="text/css">
<% if wTokucho <> "" then%>
<meta name="description" content="<%=wTokucho%>">
<% end if %>
<meta name="keywords" content="<%=wLargeCategoryName%>,<%=wMidCategoryName%>,<%=wCategoryName%>,<%=wMakerName%>,<%=wProductName%>">
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
	if (pForm.qt.value <= 0){
		alert("数量を入力してからカートボタンを押してください。");
		return false;
	}
	return true;
}
//
// ====== 	Function:	SmallImage_onMouseOver
//
function SmallImage_onMouseOver(pFile){
	document.images["LargeImage"].src = pFile;
}
</script>
<style type="text/css">
#globalContents ul.sns {
	overflow: hidden;
	padding: 5px;
}

#globalContents ul.sns li {
	float: right;
	width: 100px;
	height: 20px;
}
</style>
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->

<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>

<!-- コンテンツstart -->
<div id="globalContents">
    <div id='path_box'><div id='path_box_inner01'><div id='path_box_inner02'>
    <p class='home'><a href="<%=g_HTTP%>"><img src="<%=g_RelLink%>images/icon_home.gif" alt="HOME"></a></p>
    <ul id='path'>
      <li><a href="<%=g_HTTP%>material/">SPECIAL SELECTION一覧</a></li>
      <li><a href="PremiumGuitars.asp">プレミアムギター</a></li>
      <li><a href="PremiumGuitarsList.asp?MakerCd=<%=wMakerCode%>"><%=wMakerName%></a></li>
      <li class="now"><%=wProductName%></li>
    </ul>
  </div></div></div>
    <ul class="sns">
          <li><a href="https://twitter.com/share" class="twitter-share-button" data-lang="ja">ツイート</a><script>!function(d,s,id){var js,fjs=d.getElementsByTagName(s)[0];if(!d.getElementById(id)){js=d.createElement(s);js.id=id;js.src="//platform.twitter.com/widgets.js";fjs.parentNode.insertBefore(js,fjs);}}(document,"script","twitter-wjs");</script></li>
          <li><iframe src="//www.facebook.com/plugins/like.php?href=http%3A%2F%2Fwww.soundhouse.co.jp%2Fshop%2FPremiumGuitars.asp&amp;send=false&amp;layout=button_count&amp;width=100&amp;show_faces=false&amp;action=like&amp;colorscheme=light&amp;font&amp;height=21&amp;appId=191447484218062" scrolling="no" frameborder="0" style="border:none; overflow:hidden; width:100px; height:21px;" allowTransparency="true"></iframe></li>
        </ul>

  <div id="pgContainer">
<!-- トップ画像 START -->
<div id="pgHeader">
  <div class="topbox">
    <div class="left"></div>
    <div class="right"></div>
  </div>
</div>
<!-- トップ画像 END -->

<!-- メーカー名、商品名、商品画像大 -->
<%=wLogoHTML%>

<!-- 商品画像小 START -->
<%=wPictureHTML%>
<!-- 商品画像小 END -->

<div id="pgInfoBox">
<div class="left">
    
<!-- 商品情報 START -->
<%=wProductHTML%>
<!-- 商品情報 END -->

<!-- 特徴 START -->
<%=wTokuchoHTML%>
<!-- 特徴 END -->
<!-- スペック START -->
<%=wSpecHTML%>
<!-- スペック END -->
</div>
<!-- その他の商品 START -->
<%=wOthersHTML%>
<!-- その他の商品 END -->

</div>

  <p class="arrow"><a href="#site_title"><img src="images/PremiumGuitars/white_arrow_up.gif" alt=""></a></p>
</div>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>