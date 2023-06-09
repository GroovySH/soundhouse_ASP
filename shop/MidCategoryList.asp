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
<!--#include file="./MidCategoryList/MidCategoryList.inc"-->
<%
'========================================================================
'
'	中カテゴリー一覧ページ
'
'更新履歴
'2012/08/21 ok リニューアルに伴い新デザインに新規作成
'2012/09/03 GV #1426 大・中カテゴリ画面で表示されるSALES&OUTLET欄の表示データを一意に取得・表示する
'2014/03/19 GV 消費税増税に伴う2重表示対応
'
'========================================================================
On Error Resume Next

Response.Buffer = True

Dim MidCategoryCd
Dim wLargeCategoryCd
Dim wLargeCategoryName
Dim wMidCategoryName
Dim wMidCategoryOverview
Dim wMetaTag
Dim wNoData
Dim wSalesTaxRate
Dim wErrDesc

Dim wMidCategoryComment

Dim wNaviMakerHTML				' (左)NAVI用
Dim wNaviCategoryHTML			' (左)NAVI用
Dim wNaviPricerangeHTML			' (左)NAVI用
Dim s_category_cd				' (左)NAVI用
Dim s_mid_category_cd			' (左)NAVI用
Dim s_large_category_cd			' (左)NAVI用
Dim s_maker_cd					' (左)NAVI用
Dim sPriceFrom					' (左)NAVI用
Dim sPriceTo					' (左)NAVI用

Dim wInsertHTMLPath1
Dim wInsertHTMLPath2
Dim wStaticHTML(2)
Dim wSaleAndOutletHTML
Dim wChumokuHTML

Dim Connection
Dim RS
Dim wHTML

'=======================================================================
'	受け渡し情報取り出し
'=======================================================================
MidCategoryCd = ReplaceInput(Trim(Request("MidCategoryCd")))

'=======================================================================
'	Execute main
'=======================================================================
Call connect_db()
Call main()

'---- エラーメッセージをセッションデータに登録
If Err.Description <> "" Then
	wErrDesc = "MidCategoryList.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
End If

Call close_db()

If wNoData = "Y" Or Err.Description <> "" Then
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

'========================================================================
'
'	Function	Connect database
'
'========================================================================
Function connect_db()

'---- Connect database
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

End Function

'========================================================================
'
'	Function	Close database
'
'========================================================================
Function close_db()

Connection.Close
Set Connection = Nothing

End Function

'========================================================================
'
'	Function	Main
'
'========================================================================
Function main()

Dim vItemChar1
Dim vItemChar2
Dim vItemNum1
Dim vItemNum2
Dim vItemDate1
Dim vItemDate2
Dim vSQL
Dim vSQLMaker					' (左)NAVI用
Dim vSQLCategory				' (左)NAVI用
Dim vSQLMiddleCategory			' (左)NAVI用
Dim vSQLLargeCategory			' (左)NAVI用
Dim vSQLPricerange				' (左)NAVI用
Dim vFilePath
Dim vMsg

' 消費税率
Call getCntlMst("共通", "消費税率", "1", vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)
wSalesTaxRate = CLng(vItemNum1)

' 中カテゴリ情報取得
Call MidCategoryInfo()

If wNoData <> "Y" Then
	' ぱんくず・中カテゴリーについて・カテゴリーリスト・おすすめ商品・ニュース 静的HTMLファイルの存在チェック (有効期限切れチェック)
	If fExistMidCategoryStaticMainHTMLFile(MidCategoryCd) = False Then

		' ぱんくず・中カテゴリーについて・カテゴリーリスト・おすすめ商品・ニュース 静的HTMLテキストファイル作成
		If fMakeMidCategoryStaticMainHTMLFile(MidCategoryCd, vFilePath, vMsg) = False Then
			Exit Function
		End If

	End If
End If

Call fCreateGetProductsSQL(  "cg" _
                           , "" _
                           , "" _
                           , "" _
                           , "" _
                           , "" _
                           , "" _
                           , MidCategoryCd _
                           , "" _
                           , "" _
                           , "" _
                           , "" _
                           , vSQL _
                           , vSQLMaker _
                           , vSQLCategory _
                           , vSQLMiddleCategory _
                           , vSQLLargeCategory _
                           , vSQLPricerange)
'--- NAVI用パラメータセット
s_large_category_cd = ""
s_mid_category_cd = MidCategoryCd


' 左NAVI用 メーカーー一覧作成
Call fCreateNAVIMaker2HTML(vSQLMaker, wNaviMakerHTML)

' 左NAVI用 カテゴリー一覧作成
Call fCreateNAVICategory2HTML(vSQLCategory, wNaviCategoryHTML)

' 左NAVI用 価格帯選択作成
Call fCreateNAVIPriceRange2HTML(vSQLPriceRange, wSalesTaxRate, wNaviPriceRangeHTML)

End Function

'========================================================================
'
'	Function	中カテゴリ情報取得
'
'========================================================================
Function MidCategoryInfo()

Dim vFilePath
Dim vMsg
Dim vSQL

'---- カテゴリー 取り出し
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.中カテゴリー名日本語 "
vSQL = vSQL & "    , a.中カテゴリー説明 "
vSQL = vSQL & "    , a.メタタグ "
vSQL = vSQL & "    , a.大カテゴリーコード "
vSQL = vSQL & "    , b.大カテゴリー名 "
vSQL = vSQL & "    , a.注目商品メーカーコード AS 注目商品メーカーコード1 "
vSQL = vSQL & "    , a.注目商品商品コード AS 注目商品商品コード1 "
vSQL = vSQL & "    , a.注目商品コメント AS 注目商品コメント1 "
vSQL = vSQL & "    , a.注目商品メーカーコード2 "
vSQL = vSQL & "    , a.注目商品商品コード2 "
vSQL = vSQL & "    , a.注目商品コメント2 "
vSQL = vSQL & "    , a.注目商品メーカーコード3 "
vSQL = vSQL & "    , a.注目商品商品コード3 "
vSQL = vSQL & "    , a.注目商品コメント3 "
vSQL = vSQL & "    , a.注目商品メーカーコード4 "
vSQL = vSQL & "    , a.注目商品商品コード4 "
vSQL = vSQL & "    , a.注目商品コメント4 "
vSQL = vSQL & "    , (SELECT メーカー名 FROM メーカー WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE メーカーコード = a.注目商品メーカーコード) AS 注目商品メーカー名1 "
vSQL = vSQL & "    , (SELECT 商品名 FROM Web商品 WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE メーカーコード = a.注目商品メーカーコード "
vSQL = vSQL & "                          AND 商品コード     = a.注目商品商品コード) AS 注目商品商品名1 "
vSQL = vSQL & "    , (SELECT メーカー名 FROM メーカー WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE メーカーコード = a.注目商品メーカーコード2) AS 注目商品メーカー名2 "
vSQL = vSQL & "    , (SELECT 商品名 FROM Web商品 WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE メーカーコード = a.注目商品メーカーコード2 "
vSQL = vSQL & "                          AND 商品コード     = a.注目商品商品コード2) AS 注目商品商品名2 "
vSQL = vSQL & "    , (SELECT メーカー名 FROM メーカー WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE メーカーコード = a.注目商品メーカーコード3) AS 注目商品メーカー名3 "
vSQL = vSQL & "    , (SELECT 商品名 FROM Web商品 WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE メーカーコード = a.注目商品メーカーコード3 "
vSQL = vSQL & "                          AND 商品コード     = a.注目商品商品コード3) AS 注目商品商品名3 "
vSQL = vSQL & "    , (SELECT メーカー名 FROM メーカー WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE メーカーコード = a.注目商品メーカーコード4) AS 注目商品メーカー名4 "
vSQL = vSQL & "    , (SELECT 商品名 FROM Web商品 WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE メーカーコード = a.注目商品メーカーコード4 "
vSQL = vSQL & "                          AND 商品コード     = a.注目商品商品コード4) AS 注目商品商品名4 "
vSQL = vSQL & "    , (SELECT 商品画像ファイル名_小 FROM Web商品 WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE メーカーコード = a.注目商品メーカーコード "
vSQL = vSQL & "                          AND 商品コード     = a.注目商品商品コード) AS 商品画像ファイル名_小1 "
vSQL = vSQL & "    , (SELECT 商品画像ファイル名_小 FROM Web商品 WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE メーカーコード = a.注目商品メーカーコード2 "
vSQL = vSQL & "                          AND 商品コード     = a.注目商品商品コード2) AS 商品画像ファイル名_小2 "
vSQL = vSQL & "    , (SELECT 商品画像ファイル名_小 FROM Web商品 WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE メーカーコード = a.注目商品メーカーコード3 "
vSQL = vSQL & "                          AND 商品コード     = a.注目商品商品コード3) AS 商品画像ファイル名_小3 "
vSQL = vSQL & "    , (SELECT 商品画像ファイル名_小 FROM Web商品 WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE メーカーコード = a.注目商品メーカーコード4 "
vSQL = vSQL & "                          AND 商品コード     = a.注目商品商品コード4) AS 商品画像ファイル名_小4 "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    中カテゴリー AS a WITH (NOLOCK) "
vSQL = vSQL & "  , 大カテゴリー AS b WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "    a.中カテゴリーコード = '" & MidCategoryCd & "'"
vSQL = vSQL & "    AND a.大カテゴリーコード = b.大カテゴリーコード"

'@@@@response.write(vSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open vSQL, Connection, adOpenStatic

If RS.EOF = True Then
	wNoData = "Y"
Else
	' 大カテゴリコード
	wLargeCategoryCd = RS("大カテゴリーコード")
	' 大カテゴリ名
	wLargeCategoryName = RS("大カテゴリー名")
	' 中カテゴリ名
	wMidCategoryName = RS("中カテゴリー名日本語")
	' 中カテゴリ概要
	wMidCategoryOverview = RS("中カテゴリー説明")

	wInsertHTMLPath1 = fGetInsertHTMLPath(MidCategoryCd,"1")
	wInsertHTMLPath2 = fGetInsertHTMLPath(MidCategoryCd,"2")

	' メタタグ <から始まっていない場合は無視
	If Left(RS("メタタグ"), 1) = "<" Then
		wMetaTag = RS("メタタグ")
	End If

	'----- HTML作成
	Call CreateChumokuHTML()				' 注目商品
	
	If fExistMidCategoryStaticMainHTMLFile(MidCategoryCd) = False Then
		' 一押し商品用 静的HTMLテキストファイル作成
		If fMakeMidCategoryStaticMainHTMLFile(MidCategoryCd, vFilePath, vMsg) = False Then
			Exit Function
		End If
	End If

	Call fIncludeMidCategoryStaticTextMain(MidCategoryCd)
	Call CreateSaleAndOutletHTML()

End If

RS.Close
Set RS = Nothing

End Function

'========================================================================
'
'	Function	注目商品
'
'========================================================================
Function CreateChumokuHTML()

Dim vItem
Dim i
Dim vCnt

'----- 注目商品HTML編集
wHTML = ""
wHTML = wHTML & "        <h2 class='subtitle04' id='pickup'>" & wMidCategoryName & "のピックアップアイテム" & "</h2>" & vbNewLine

wHTML = wHTML & "        <div id='pickup_box'>" & vbNewLine

vCnt=1
For i = 1 To 4 Step 1
	If GetProductFlag(RS("注目商品メーカーコード" & i),RS("注目商品商品コード" & i)) = "Y" Then
		vItem = Server.URLEncode(RS("注目商品メーカーコード" & i) & "^" & RS("注目商品商品コード" & i))
		
		If (vCnt Mod 2) = 1 Then
			If vCnt <> 1 Then
				wHTML = wHTML & "            </ul>" & vbNewLine
			End If
			wHTML = wHTML & "            <ul class='pickup col" & int(vCnt/2)+(vCnt Mod 2)  & "' >" & vbNewLine
		End If
		wHTML = wHTML & "                <li>" & vbNewLine
		wHTML = wHTML & "                    <div class='pickup_inner'><div class='pickup_inner02'>" & vbNewLine
		wHTML = wHTML & "                        <div class='item_name_box'>" & vbNewLine
		wHTML = wHTML & "                            <p class='left'><a href='ProductDetail.asp?Item=" & vItem & "'>"
		If RS("商品画像ファイル名_小" & i) <> "" Then
			wHTML = wHTML & "<img src='prod_img/" & RS("商品画像ファイル名_小" & i) & "' alt='" & RS("注目商品メーカー名" & i) & " / " & RS("注目商品商品名" & i) & "' class='opover'>"
		End If
		wHTML = wHTML & "</a></p>" & vbNewLine
		wHTML = wHTML & "                            <p class='item_name'><a href='ProductDetail.asp?Item=" & vItem & "'>" & RS("注目商品メーカー名" & i) & " / " & RS("注目商品商品名" & i) & "</a><p>" & vbNewLine
		wHTML = wHTML & "                            <p>" & GetPrice(RS("注目商品メーカーコード" & i), RS("注目商品商品コード" & i)) & "</p>" & vbNewLine
		wHTML = wHTML & "                        </div>" & vbNewLine
		wHTML = wHTML & "                        <p class='desc'>" & RS("注目商品コメント" & i) & "</p>" & vbNewLine
		wHTML = wHTML & "                    </div></div>" & vbNewLine
		wHTML = wHTML & "                </li>" & vbNewLine
		
		vCnt = vCnt+1
	End If
Next

If vCnt = 1 Then
	Exit Function
End If

wHTML = wHTML & "            </ul>" & vbNewLine
wHTML = wHTML & "        </div>" & vbNewLine

wChumokuHTML = wHTML

End Function

'========================================================================
'
'	Function	SALE&OUTLET商品
'	2012/08/21 ok Add
'========================================================================
Function CreateSaleAndOutletHTML()

Dim RSv
Dim vSQL
Dim v_price
Dim v_exprice
' 2012/09/03 GV #1426 Add Start
Dim wHTML1
Dim cnt
Dim ctr
Dim dcnt
Dim flg
Dim w_MakerCd()
Dim w_ItemCd()
Dim w_price1()
Dim w_price2()
cnt = 0
dcnt = 0
' 2012/09/03 GV #1426 Add End

'---- セール商品取り出し
vSQL = ""
vSQL = vSQL & "SELECT "
' 2012/09/03 GV #1426 Mod Start
'vSQL = vSQL & "    TOP 5 "
vSQL = vSQL & "    TOP 20 "
' 2012/09/03 GV #1426 Mod End
vSQL = vSQL & "      a.商品コード "
vSQL = vSQL & "    , a.商品名 "
vSQL = vSQL & "    , a.メーカーコード "
vSQL = vSQL & "    , a.メーカー名 "
vSQL = vSQL & "    , a.商品画像ファイル名_小 "
vSQL = vSQL & "    , a.販売単価 "
vSQL = vSQL & "    , a.前回販売単価 "
vSQL = vSQL & "    , a.ASK商品フラグ "
vSQL = vSQL & "    , a.B品フラグ "
vSQL = vSQL & "    , a.個数限定数量 "
vSQL = vSQL & "    , a.個数限定単価 "
vSQL = vSQL & "    , a.個数限定受注済数量 "
vSQL = vSQL & "    , a.前回単価変更日 "
vSQL = vSQL & "    , a.B品フラグ "
vSQL = vSQL & "    , a.B品単価 "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    Webセール商品 a WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "    a.セール区分番号 BETWEEN 1 AND 4"
vSQL = vSQL & " AND a.中カテゴリーコード = '" & MidCategoryCd & "' "
vSQL = vSQL & "ORDER BY NEWID() "

'@@@@@@@@@@Response.Write(vSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic
wHTML = ""

If RSv.EOF = false Then
	'----- セール商品HTML編集
	wHTML = wHTML & "        <h2 class='subtitle_red'>" & wMidCategoryName & "のSALE &amp; OUTLET</h2>" & vbNewLine
	wHTML = wHTML & "        <div class='box'><div class='box_inner01'>" & vbNewLine
	wHTML = wHTML & "            <ul class='list'>" & vbNewLine

	Do Until RSv.EOF = True OR dcnt > 4
' 2012/09/03 GV #1426 Add Start
		ReDim Preserve w_MakerCd(cnt)
		w_MakerCd(cnt) = RSv("メーカーコード")
		ReDim Preserve w_ItemCd(cnt)
		w_ItemCd(cnt) = RSv("商品コード")
		wHTML1 = ""
' 2012/09/03 GV #1426 Add End
		wHTML1 = wHTML1 & "                <li><a href='ProductDetail.asp?Item=" & Server.URLEncode(RSv("メーカーコード") & "^" & RSv("商品コード")) & "'>"
		If RSv("商品画像ファイル名_小") <> "" Then
			wHTML1 = wHTML1 & "<img src='prod_img/" & RSv("商品画像ファイル名_小") & "' alt='" & RSv("メーカー名") & " / " & RSv("商品名") & "' class='opover'>"
		End If
		wHTML1 = wHTML1 & RSv("メーカー名") & " / " & RSv("商品名") & "</a><span>"
		
		'---- 販売単価
		v_price = calcPrice(RSv("販売単価"), wSalesTaxRate)
		v_exprice = calcPrice(RSv("前回販売単価"), wSalesTaxRate)
		'1行目の表示（ASK商品ではない値下げ品の旧価格）
		If RSv("ASK商品フラグ") <> "Y" Then
			If RSv("B品フラグ") = "Y" OR (RSv("個数限定数量") > RSv("個数限定受注済数量") AND RSv("個数限定数量") > 0) OR ( isNULL(RSv("前回単価変更日")) = False AND DateAdd("d", 60, RSv("前回単価変更日")) >= Date() AND RSv("前回販売単価") > RSv("販売単価") AND RSv("前回販売単価") <> 0) Then

				'値下げ品の旧価格を表示
				If isNULL(RSv("前回単価変更日")) = False AND DateAdd("d", 60, RSv("前回単価変更日")) >= Date() AND RSv("前回販売単価") > RSv("販売単価") Then
'2013/03/19 GV mod start ---->
'前回単価はしばらく表示させない
'					wHTML1 = wHTML1 & FormatNumber(v_exprice,0) & "円（税込）↓<br>"
'2013/03/19 GV mod end   <----
' 2012/09/03 GV #1426 Add Start
					ReDim Preserve w_price1(cnt)
					w_price1(cnt) = FormatNumber(v_exprice,0)
' 2012/09/03 GV #1426 Add End
				'B品、限定品は販売価格を旧価格として表示
				Else
'2013/03/19 GV mod start ---->
'前回単価はしばらく表示させない
'					wHTML1 = wHTML1 & FormatNumber(v_price,0) & "円（税込）↓<br>"
'2013/03/19 GV mod end   <----
' 2012/09/03 GV #1426 Add Start
					ReDim Preserve w_price1(cnt)
					w_price1(cnt) = FormatNumber(v_price,0)
' 2012/09/03 GV #1426 Add End
				End If
' 2012/09/03 GV #1426 Add Start
			Else
				ReDim Preserve w_price1(cnt)
				w_price1(cnt) = 0
' 2012/09/03 GV #1426 Add End
			End If
' 2012/09/03 GV #1426 Add Start
		Else
			ReDim Preserve w_price1(cnt)
			w_price1(cnt) = 0
' 2012/09/03 GV #1426 Add End
		End If

		'2行目の表示（通常価格 or ASK or 値下げ後価格）
		If RSv("ASK商品フラグ") <> "Y" Then
			'---- B品単価
			If RSv("B品フラグ") = "Y" Then
				v_price = calcPrice(RSv("B品単価"), wSalesTaxRate)
'2013/03/19 GV mod start ---->
'				wHTML1 = wHTML1 & "<strong>【わけあり品特価】" & FormatNumber(v_price,0) & "円(税込)</strong>"
				wHTML1 = wHTML1 & "<strong>【わけあり品特価】" & FormatNumber(RSv("B品単価"),0) & "円(税抜)</strong><br>"
'2013/03/19 GV mod end   <----
' 2012/09/03 GV #1426 Add Start
				ReDim Preserve w_price2(cnt)
				w_price2(cnt) = FormatNumber(v_price,0)
' 2012/09/03 GV #1426 Add End
			'---- 個数限定単価
			ElseIf RSv("個数限定数量") > RSv("個数限定受注済数量") AND RSv("個数限定数量") > 0 Then
				v_price = calcPrice(RSv("個数限定単価"), wSalesTaxRate)
'2013/03/19 GV mod start ---->
'				wHTML1 = wHTML1 & "<strong>【限定特価】" & FormatNumber(v_price,0) & "円(税込)</strong>"
				wHTML1 = wHTML1 & "<strong>【限定特価】" & FormatNumber(RSv("個数限定単価"),0) & "円(税抜)</strong><br>"
'2013/03/19 GV mod end   <----
' 2012/09/03 GV #1426 Add Start
				ReDim Preserve w_price2(cnt)
				w_price2(cnt) = FormatNumber(v_price,0)
' 2012/09/03 GV #1426 Add End
			'---- 通常商品
			Else
'2013/03/19 GV mod start ---->
'				wHTML1 = wHTML1 & "<strong>【衝撃特価】" & FormatNumber(v_price,0) & "円(税込)</strong>"
				wHTML1 = wHTML1 & "<strong>【衝撃特価】" & FormatNumber(RSv("販売単価"),0) & "円(税抜)</strong><br>"
'2013/03/19 GV mod end   <----
' 2012/09/03 GV #1426 Add Start
				ReDim Preserve w_price2(cnt)
				w_price2(cnt) = FormatNumber(v_price,0)
' 2012/09/03 GV #1426 Add End
			End If

			wHTML1 = wHTML1 & "(税込&nbsp;" & FormatNumber(v_price,0) & "円)"	'2013/03/19 GV ad

			wHTML1 = wHTML1 & "</span></li>" & vbNewLine

		Else
			'---- B品単価
			If RSv("B品フラグ") = "Y" Then
				v_price = calcPrice(RSv("B品単価"), wSalesTaxRate)
'2013/03/19 GV mod start ---->
'				wHTML1 = wHTML1 & "【わけあり品特価】</span><a class='tip'>ASK<span>" & FormatNumber(v_price,0) & "円(税込)</span>"
				wHTML1 = wHTML1 & "【わけあり品特価】</span><a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RSv("B品単価"),0) & "円(税抜)</span><br>"
'2013/03/19 GV mod end   <----
' 2012/09/03 GV #1426 Add Start
				ReDim Preserve w_price2(cnt)
				w_price2(cnt) = FormatNumber(v_price,0)
' 2012/09/03 GV #1426 Add End
			'---- 個数限定単価
			ElseIf RSv("個数限定数量") > RSv("個数限定受注済数量") AND RSv("個数限定数量") > 0 Then
				v_price = calcPrice(RSv("個数限定単価"), wSalesTaxRate)
'2013/03/19 GV mod start ---->
'				wHTML1 = wHTML1 & "【限定特価】</span><a class='tip'>ASK<span>" & FormatNumber(v_price,0) & "円(税込)</span>"
				wHTML1 = wHTML1 & "【限定特価】</span><a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RSv("個数限定単価"),0) & "円(税抜)</span><br>"
'2013/03/19 GV mod end   <----
' 2012/09/03 GV #1426 Add Start
				ReDim Preserve w_price2(cnt)
				w_price2(cnt) = FormatNumber(v_price,0)
' 2012/09/03 GV #1426 Add End
			'---- 通常商品
			Else
'2013/03/19 GV mod start ---->
'				wHTML1 = wHTML1 & "【衝撃特価】</span><a class='tip'>ASK<span>" & FormatNumber(v_price,0) & "円(税込)</span>"
				wHTML1 = wHTML1 & "【衝撃特価】</span><a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RSv("販売単価"),0) & "円(税抜)</span><br>"
'2013/03/19 GV mod end   <----
' 2012/09/03 GV #1426 Add Start
				ReDim Preserve w_price2(cnt)
				w_price2(cnt) = FormatNumber(v_price,0)
' 2012/09/03 GV #1426 Add End
			End If

			wHTML1 = wHTML1 & "<span class='inc-tax'>(税込&nbsp;" & FormatNumber(v_price,0) & "円)</span>"	'2013/03/19 GV ad

			wHTML1 = wHTML1 & "</a></li>" & vbNewLine

		End If

' 2012/09/03 GV #1426 Add Start
		flg = True
		For ctr = 0 to Ubound(w_ItemCd)
			If ctr < cnt Then
				If w_MakerCd(ctr) = w_MakerCd(cnt) AND w_ItemCd(ctr) = w_ItemCd(cnt) Then
					if w_price1(ctr) = w_price1(cnt) AND w_price2(ctr) = w_price2(cnt) Then
						flg = False
						Exit For
					End If
				End If
			End If
		Next
		if flg Then
			dcnt = dcnt + 1
			wHTML = wHTML & wHTML1
		End If
		cnt = cnt + 1
' 2012/09/03 GV #1426 Add End

		RSv.MoveNext
	Loop

	wHTML = wHTML & "            </ul>" & vbNewLine
	wHTML = wHTML & "        </div></div>" & vbNewLine
End If
wSaleAndOutletHTML = wHTML

RSv.Close

End Function

'========================================================================
'
'	Function	商品価格取得
'
'========================================================================
Function GetPrice(pMakerCd, pProductCd)

Dim RSv
Dim vSQL
Dim v_price
GetPrice = ""

'---- Web商品フラグ取り出し
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.販売単価 "
vSQL = vSQL & "    , a.前回販売単価 "
vSQL = vSQL & "    , a.ASK商品フラグ "
vSQL = vSQL & "    , a.B品フラグ "
vSQL = vSQL & "    , a.個数限定数量 "
vSQL = vSQL & "    , a.個数限定単価 "
vSQL = vSQL & "    , a.個数限定受注済数量 "
vSQL = vSQL & "    , a.前回単価変更日 "
vSQL = vSQL & "    , a.B品フラグ "
vSQL = vSQL & "    , a.B品単価 "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    Web商品 a WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        a.メーカーコード = '" & pMakerCd & "' "
vSQL = vSQL & "    AND a.商品コード     = '" & pProductCd & "'"
vSQL = vSQL & "    AND a.Web商品フラグ  = 'Y'"

'@@@@@@@@@@Response.Write(vSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

If RSv.EOF = false Then
	'---- 販売単価
	v_price = calcPrice(RSv("販売単価"), wSalesTaxRate)

	If RSv("ASK商品フラグ") <> "Y" Then
		'---- B品単価
		If RSv("B品フラグ") = "Y" Then
			v_price = calcPrice(RSv("B品単価"), wSalesTaxRate)
'2014/03/19 GV mod start ---->
'			GetPrice = GetPrice & "【わけあり品特価】" & FormatNumber(v_price,0) & "円(税込)"
			GetPrice = GetPrice & "【わけあり品特価】" & FormatNumber(RSv("B品単価"),0) & "円(税抜)<br>"
			GetPrice = GetPrice & "(税込&nbsp;" & FormatNumber(v_price,0) & "円)"
'2014/03/19 GV mod end   <----
		'---- 個数限定単価
		ElseIf RSv("個数限定数量") > RSv("個数限定受注済数量") AND RSv("個数限定数量") > 0 Then
			v_price = calcPrice(RSv("個数限定単価"), wSalesTaxRate)
'2014/03/19 GV mod start ---->
'			GetPrice = GetPrice& "【限定特価】" & FormatNumber(v_price,0) & "円(税込)"
			GetPrice = GetPrice& "【限定特価】" & FormatNumber(RSv("個数限定単価"),0) & "円(税抜)<br>"
			GetPrice = GetPrice & "(税込&nbsp;" & FormatNumber(v_price,0) & "円)"
'2014/03/19 GV mod end   <----
		'---- 通常商品
		Else
'2014/03/19 GV mod start ---->
'			GetPrice = GetPrice &  FormatNumber(v_price,0) & "円(税込)"
			GetPrice = GetPrice &  FormatNumber(RSv("販売単価"),0) & "円(税抜)<br>"
			GetPrice = GetPrice & "(税込&nbsp;" & FormatNumber(v_price,0) & "円)"
'2014/03/19 GV mod end   <----
		End If
	Else
		'---- B品単価
		If RSv("B品フラグ") = "Y" Then
			v_price = calcPrice(RSv("B品単価"), wSalesTaxRate)
'2014/03/19 GV mod start ---->
'			GetPrice = GetPrice & "【わけあり品特価】<a class='tip'>ASK<span>" & FormatNumber(v_price,0) & "円(税込)</span></a>"
			GetPrice = GetPrice & "【わけあり品特価】<a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RSv("B品単価"),0) & "円(税抜)</span><br>"
			GetPrice = GetPrice & "<span class='inc-tax'>(税込&nbsp;" & FormatNumber(v_price,0) & "円)</span></a>"
'2014/03/19 GV mod end   <----
		'---- 個数限定単価
		ElseIf RSv("個数限定数量") > RSv("個数限定受注済数量") AND RSv("個数限定数量") > 0 Then
			v_price = calcPrice(RSv("個数限定単価"), wSalesTaxRate)
'2014/03/19 GV mod start ---->
'			GetPrice = GetPrice & "【限定特価】<a class='tip'>ASK<span>" & FormatNumber(v_price,0) & "円(税込)</span></a>"
			GetPrice = GetPrice & "【限定特価】<a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RSv("個数限定単価"),0) & "円(税抜)</span><br>"
			GetPrice = GetPrice & "<span class='inc-tax'>(税込&nbsp;" & FormatNumber(v_price,0) & "円)</span></a>"
'2014/03/19 GV mod end   <----
		'---- 通常商品
		Else
'2014/03/19 GV mod start ---->
'			GetPrice = GetPrice & "<a class='tip'>ASK<span>" & FormatNumber(v_price,0) & "円(税込)</span></a>"
			GetPrice = GetPrice & "<a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RSv("販売単価"),0) & "円(税抜)</span><br>"
			GetPrice = GetPrice & "<span class='inc-tax'>(税込&nbsp;" & FormatNumber(v_price,0) & "円)</span></a>"
'2014/03/19 GV mod end   <----
		End If
	End If
End If

RSv.Close

End Function

'========================================================================
'
'	Function	Web商品フラグチェック
'
'========================================================================
Function GetProductFlag(pMakerCd, pProductCd)

Dim RSv
Dim vSQL
GetProductFlag = ""

'---- Web商品フラグ取り出し
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.Web商品フラグ "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    Web商品 a WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        a.メーカーコード = '" & pMakerCd & "' "
vSQL = vSQL & "    AND a.商品コード     = '" & pProductCd & "'"

'@@@@@@@@@@Response.Write(vSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

If RSv.EOF = false Then
	GetProductFlag = RSv("Web商品フラグ")
End If

RSv.Close

End Function

'========================================================================
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<meta name="robots" content="noindex,nofollow">
<title><% = wMidCategoryName %> 一覧｜サウンドハウス</title>
<% = wMetaTag %>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/shop.css?20121116" type="text/css">
<link rel="stylesheet" href="style/categorylist.css?20140812" type="text/css">
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
				<li><span itemscope itemtype="http://data-vocabulary.org/Breadcrumb"><a href="LargeCategoryList.asp?LargeCategoryCd=<%=wLargeCategoryCd%>" itemprop="url"><span itemprop="title"><%=wLargeCategoryName%></span></a></span></li>
				<li class="now"><span itemscope itemtype="http://data-vocabulary.org/Breadcrumb"><span itemprop='title'><%=wMidCategoryName%></span></span></li>
			</ul>
		</div></div></div>
<!-- ページメイン部分の記述 START -->

<%=fIncludeInsertHTML(wInsertHTMLPath1)%>

<!-- 中カテゴリーについて・カテゴリーから選ぶ・最新ニュース・新製品 -->

<%=wStaticHTML(0)%>

<%=wChumokuHTML%>

<%=wStaticHTML(1)%>

<%=fIncludeInsertHTML(wInsertHTMLPath2)%>

<%=wSaleAndOutletHTML%>

<%=wStaticHTML(2)%>

    <!--/#contents --></div>
  
<!-- 絞込検索用Form -->
    <form name='f_search' method='get' action='SearchList.asp'>
      <input type='hidden' name='s_maker_cd' value=''>
      <input type='hidden' name='s_category_cd' value=''>
      <input type='hidden' name='s_mid_category_cd' value='<% = MidCategoryCd %>'>
      <input type='hidden' name='s_large_category_cd' value=''>
      <input type='hidden' name='s_product_cd' value=''>
      <input type='hidden' name='search_all' value=''>
      <input type='hidden' name='sSeriesCd' value=''>
      <input type='hidden' name='sPriceFrom' value=''>
      <input type='hidden' name='sPriceTo' value=''>
      <input type='hidden' name='i_type' value=''>
      <input type='hidden' name='i_sub_type' value=''>
      <input type='hidden' name='i_page' value='1'>
      <input type='hidden' name='i_sort' value=''>
      <input type='hidden' name='i_page_size' value=''>
      <input type='hidden' name='i_ListType' value=''>
    </form>

	<div id="globalSide">
<%
	' 左NAVI用パラメータセット
	NAVIMidCategoryCd = MidCategoryCd
	NAVISearchListMakerListHTML = wNaviMakerHTML
	NAVISearchListCategoryListHTML = wNaviCategoryHTML
	NAVISearchListPriceRangeListHTML = wNaviPriceRangeHTML
%>
<!--#include file="../Navi/NaviSideShop.inc"-->
<!--#include file="../Navi/NaviSide.inc"-->
    <!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/Navibottom.inc"-->
<div class="tooltip"><p>ASK</p></div>
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript" src="jslib/MidCategoryList.js"></script>
<script type="text/javascript" src="jslib/ask.js?20140401a"></script>
<script type="text/javascript" src="jslib/SearchList.js?20121108" charset="Shift_JIS"></script>
<script type="text/javascript" src="../jslib/jquery.tinyscrollbar.min.js"></script>
<script type="text/javascript">
$(function(){
    $('#scrollbar1').tinyscrollbar();
});
</script>
</body>
</html>
