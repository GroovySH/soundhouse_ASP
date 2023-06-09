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
<!--#include file="./LargeCategoryList/LargeCategoryList.inc"-->
<%
'========================================================================
'
'	大カテゴリー一覧ページ
'
'更新履歴
'2008/12/19 リニューアル　新規
'2009/04/13 常にALL表示に変更
'2009/05/20 MidCategoryListへのリンク追加
'2009/08/18 トピックス(News)の表示条件に商品記事.大カテゴリーコード=該当大カテゴリーコードを追加
'2009/11/05 an METAタグ追加機能を追加
'2010/01/26 an 存在しないカテゴリーを指定した場合は、Error.aspを表示（Error.asp側でTOPにリダイレクト）
'2010/05/17 ko-web 検索対策のためHTMLタグ（hX,p,strong）追加
'2010/12/07 an 一般記事を表示するように修正
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/01/20 GV WITH (NOLOCK) 漏れ 追加
'2012/01/20 GV 最新ニュース および 新製品 の情報を「Web商品記事TOP10」テーブルおよび「Web新製品TOP10」テーブルより取得するよう変更
'2012/01/23 GV 商品記事のソート順 記事日付 DESC → 記事番号 DESC へ変更
'2012/03/13 GV #1224 「一押し商品」の表示部以外を外部の静的テキストファイルから取り込むように変更
'2012/03/13 GV #1224 「一押し商品」の表示部以外を外部の静的テキストファイルが存在もしくは有効期限切れの場合、生成する処理追加
'2012/07/23 ok リニューアルに伴い新デザインに変更
'2012/09/03 GV #1426 大・中カテゴリ画面で表示されるSALES&OUTLET欄の表示データを一意に取得・表示する
'2014/03/19 GV 消費税増税に伴う2重表示対応
'
'========================================================================
On Error Resume Next

Dim LargeCategoryCd
'Dim ALLFl										' 2012/03/13 GV Del

Dim wSalesTaxRate

Dim wLargeCategoryName
Dim wLargeCategoryComment
Dim wMetaTag
Dim wNoData '2010/01/26 an 追加

Dim wIchioshiHTML
'Dim wMidCategoryListHTML						' 2012/03/13 GV Del
'Dim wNewsHTML									' 2012/03/13 GV Del
'Dim wNewItemHTML								' 2012/03/13 GV Del

Dim Connection
Dim RS

Dim wSQL

Dim wHTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add
Dim wInsertHTMLPath1		'2012/07/24 ok Add
Dim wInsertHTMLPath2		'2012/07/24 ok Add
Dim wStaticHTML(2)			'2012/07/24 ok Add
Dim wSaleAndOutletHTML		'2012/07/24 ok Add

'========================================================================

Response.Buffer = True

'---- Get input data
LargeCategoryCd = ReplaceInput(Trim(Request("LargeCategoryCd")))
'AllFl = ReplaceInput(Trim(Request("AllFl")))							' 2012/03/13 GV Del (未使用の為)

'AllFl = "Y"															' 2012/03/13 GV Del (未使用の為)

'---- Execute main
Call connect_db()
Call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
If Err.Description <> "" Then
	wErrDesc = "LargeCategoryList.asp" & " " & Replace(Replace(Err.Description, vbCr, " "), vbLf, " ")
	Call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
End If                                           '2011/08/01 an add e

Call close_db()

If wNoData = "Y" Or Err.Description <> "" Then '2010/01/26 an 修正
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
'	Function	Main
'
'========================================================================
Function main()

Dim vFilePath			' 2012/03/13 GV Add
Dim vMsg				' 2012/03/13 GV Add
Dim vItemChar1
Dim vItemChar2
Dim vItemNum1
Dim vItemNum2
Dim vItemDate1
Dim vItemDate2


'--- 消費税率取出し
Call getCntlMst("共通","消費税率","1", vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)			'消費税率
wSalesTaxRate = Clng(vItemNum1)

'---- 大カテゴリー 取り出し
wSQL = ""
wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.大カテゴリー名 "
wSQL = wSQL & "    , a.大カテゴリー説明 "
wSQL = wSQL & "    , a.一押しメーカーコード1 "
wSQL = wSQL & "    , a.一押し商品コード1 "
'wSQL = wSQL & "    , a.一押し画像ファイル名1 "		'2012/07/24 ok Del
wSQL = wSQL & "    , a.一押しメーカーコード2 "
wSQL = wSQL & "    , a.一押し商品コード2 "
'wSQL = wSQL & "    , a.一押し画像ファイル名2 "		'2012/07/24 ok Del
wSQL = wSQL & "    , a.一押しメーカーコード3 "
wSQL = wSQL & "    , a.一押し商品コード3 "
'wSQL = wSQL & "    , a.一押し画像ファイル名3 "		'2012/07/24 ok Del
wSQL = wSQL & "    , a.一押しメーカーコード4 "
wSQL = wSQL & "    , a.一押し商品コード4 "
'wSQL = wSQL & "    , a.一押し画像ファイル名4 "		'2012/07/24 ok Del
wSQL = wSQL & "    , a.一押しメーカーコード5 "
wSQL = wSQL & "    , a.一押し商品コード5 "
'wSQL = wSQL & "    , a.一押し画像ファイル名5 "		'2012/07/24 ok Del
wSQL = wSQL & "    , a.メタタグ "
wSQL = wSQL & "    , (SELECT メーカー名 FROM メーカー WITH (NOLOCK) "												' 2012/01/20 GV Mod  WITH (NOLOCK) 付加
wSQL = wSQL & "                        WHERE メーカーコード = a.一押しメーカーコード1) AS 一押しメーカー名1 "
wSQL = wSQL & "    , (SELECT 商品名 FROM Web商品 WITH (NOLOCK) "
wSQL = wSQL & "                        WHERE メーカーコード = a.一押しメーカーコード1 "
wSQL = wSQL & "                          AND 商品コード     = a.一押し商品コード1) AS 一押し商品名1 "
wSQL = wSQL & "    , (SELECT メーカー名 FROM メーカー WITH (NOLOCK) "												' 2012/01/20 GV Mod  WITH (NOLOCK) 付加
wSQL = wSQL & "                        WHERE メーカーコード = a.一押しメーカーコード2) AS 一押しメーカー名2 "
wSQL = wSQL & "    , (SELECT 商品名 FROM Web商品 WITH (NOLOCK) "
wSQL = wSQL & "                        WHERE メーカーコード = a.一押しメーカーコード2 "
wSQL = wSQL & "                          AND 商品コード     = a.一押し商品コード2) AS 一押し商品名2 "
wSQL = wSQL & "    , (SELECT メーカー名 FROM メーカー WITH (NOLOCK) "												' 2012/01/20 GV Mod  WITH (NOLOCK) 付加
wSQL = wSQL & "                        WHERE メーカーコード = a.一押しメーカーコード3) AS 一押しメーカー名3 "
wSQL = wSQL & "    , (SELECT 商品名 FROM Web商品 WITH (NOLOCK) "
wSQL = wSQL & "                        WHERE メーカーコード = a.一押しメーカーコード3 "
wSQL = wSQL & "                          AND 商品コード     = a.一押し商品コード3) AS 一押し商品名3 "
wSQL = wSQL & "    , (SELECT メーカー名 FROM メーカー WITH (NOLOCK) "												' 2012/01/20 GV Mod  WITH (NOLOCK) 付加
wSQL = wSQL & "                        WHERE メーカーコード = a.一押しメーカーコード4) AS 一押しメーカー名4 "
wSQL = wSQL & "    , (SELECT 商品名 FROM Web商品 WITH (NOLOCK) "
wSQL = wSQL & "                        WHERE メーカーコード = a.一押しメーカーコード4 "
wSQL = wSQL & "                          AND 商品コード     = a.一押し商品コード4) AS 一押し商品名4 "
wSQL = wSQL & "    , (SELECT メーカー名 FROM メーカー WITH (NOLOCK) "												' 2012/01/20 GV Mod  WITH (NOLOCK) 付加
wSQL = wSQL & "                        WHERE メーカーコード = a.一押しメーカーコード5) AS 一押しメーカー名5 "
wSQL = wSQL & "    , (SELECT 商品名 FROM Web商品 WITH (NOLOCK) "
wSQL = wSQL & "                        WHERE メーカーコード = a.一押しメーカーコード5 "
wSQL = wSQL & "                          AND 商品コード     = a.一押し商品コード5) AS 一押し商品名5 "
'2012/07/25 ok Add Start
wSQL = wSQL & "    , (SELECT 商品画像ファイル名_小 FROM Web商品 WITH (NOLOCK) "
wSQL = wSQL & "                        WHERE メーカーコード = a.一押しメーカーコード1 "
wSQL = wSQL & "                          AND 商品コード     = a.一押し商品コード1) AS 商品画像ファイル名_小1 "
wSQL = wSQL & "    , (SELECT 商品画像ファイル名_小 FROM Web商品 WITH (NOLOCK) "
wSQL = wSQL & "                        WHERE メーカーコード = a.一押しメーカーコード2 "
wSQL = wSQL & "                          AND 商品コード     = a.一押し商品コード2) AS 商品画像ファイル名_小2 "
wSQL = wSQL & "    , (SELECT 商品画像ファイル名_小 FROM Web商品 WITH (NOLOCK) "
wSQL = wSQL & "                        WHERE メーカーコード = a.一押しメーカーコード3 "
wSQL = wSQL & "                          AND 商品コード     = a.一押し商品コード3) AS 商品画像ファイル名_小3 "
wSQL = wSQL & "    , (SELECT 商品画像ファイル名_小 FROM Web商品 WITH (NOLOCK) "
wSQL = wSQL & "                        WHERE メーカーコード = a.一押しメーカーコード4 "
wSQL = wSQL & "                          AND 商品コード     = a.一押し商品コード4) AS 商品画像ファイル名_小4 "
wSQL = wSQL & "    , (SELECT 商品画像ファイル名_小 FROM Web商品 WITH (NOLOCK) "
wSQL = wSQL & "                        WHERE メーカーコード = a.一押しメーカーコード5 "
wSQL = wSQL & "                          AND 商品コード     = a.一押し商品コード5) AS 商品画像ファイル名_小5 "
wSQL = wSQL & "    , 一押しコメント "
'2012/07/25 ok Add End
wSQL = wSQL & "FROM "
wSQL = wSQL & "    大カテゴリー a WITH (NOLOCK) "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "    a.大カテゴリーコード = '" & LargeCategoryCd & "' "

'@@@@@@@@@@Response.Write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

If RS.EOF = True Then
	wNoData = "Y" '2010/01/26 an 修正
Else
	'----- 大カテゴリー名
	wLargeCategoryName = RS("大カテゴリー名")
	wLargeCategoryComment = RS("大カテゴリー説明")

'2012/07/24 ok Add Start
	wInsertHTMLPath1 = fGetInsertHTMLPath(LargeCategoryCd,"1")
	wInsertHTMLPath2 = fGetInsertHTMLPath(LargeCategoryCd,"2")
'2012/07/24 ok Add End

	'----- メタタグ <から始まっていない場合は無視
	If Left(RS("メタタグ"),1) = "<" Then
		wMetaTag = RS("メタタグ")
	End If

	'----- HTML作成
	Call CreateIchioshiHTML()				' 一押し商品
'	Call CreateMidCategoryListHTML()		' 中カテゴリー一覧
'	Call CreatewNewsHTML()					' トピックス News
'	Call CreateNewItemHTML()				' トピックス 新製品

' 2012/03/13 GV Add Start
	' 一押し商品用 静的HTMLファイルの存在チェック (有効期限切れチェック)
	If fExistLargeCategoryStaticHTMLFile(LargeCategoryCd) = False Then

		' 一押し商品用 静的HTMLテキストファイル作成
		If fMakeLargeCategoryStaticHTMLFile(LargeCategoryCd, vFilePath, vMsg) = False Then
			Exit Function
		End If

	End If
' 2012/03/13 GV Add End

'2012/07/24 ok Add Start
	Call fIncludeLargeCategoryStaticText(LargeCategoryCd)
	Call CreateSaleAndOutletHTML()
'2012/07/24 ok Add End
End If

RS.Close

End Function

'========================================================================
'
'	Function	一押し商品
'
'========================================================================
Function CreateIchioshiHTML()

'Dim vPrice		'2012/07/24 ok Del
Dim vItem
Dim i
Dim vCnt

'----- 一押し商品HTML編集
wHTML = ""
'2012/07/24 ok Del Start
'wHTML = wHTML & "<table width='794' border='0' cellpadding='0' cellspacing='0' id='Shop_LargeCategory_HotItem'>" & vbNewLine
'
'vPrice = getPrice(RS("一押しメーカーコード1"), RS("一押し商品コード1"))
'vPrice = calcPrice(vPrice, wSalesTaxRate)
'vItem = Server.URLEncode(RS("一押しメーカーコード1") & "^" & RS("一押し商品コード1"))
'
'wHTML = wHTML & "  <tr>" & vbNewLine
'wHTML = wHTML & "    <td rowspan='2' class='1oshi'>" & vbNewLine
'wHTML = wHTML & "      <a href='ProductDetail.asp?Item=" & vItem & "'><img src='cat_hotitem/" & RS("一押し画像ファイル名1") & "' alt='" & RS("一押しメーカー名1") & " " & RS("一押し商品名1") & "' width='406' height='320' border='0'></a><br>衝撃特価&nbsp;" & FormatNumber(vPrice,0) & "円（税込）&nbsp;" & vbNewLine
'wHTML = wHTML & "    </td>" & vbNewLine
'
'vPrice = getPrice(RS("一押しメーカーコード2"), RS("一押し商品コード2"))
'vPrice = calcPrice(vPrice, wSalesTaxRate)
'vItem = Server.URLEncode(RS("一押しメーカーコード2") & "^" & RS("一押し商品コード2"))
'
'wHTML = wHTML & "    <td>" & vbNewLine
'wHTML = wHTML & "      <a href='ProductDetail.asp?Item=" & vItem & "'><img src='cat_hotitem/" & RS("一押し画像ファイル名2") & "' alt='" & RS("一押しメーカー名2") & " " & RS("一押し商品名2") & "' width='190' height='150' border='0'></a><br>" & FormatNumber(vPrice,0) & "円（税込）&nbsp;" & vbNewLine
'wHTML = wHTML & "    </td>" & vbNewLine
'
'vPrice = getPrice(RS("一押しメーカーコード3"), RS("一押し商品コード3"))
'vPrice = calcPrice(vPrice, wSalesTaxRate)
'vItem = Server.URLEncode(RS("一押しメーカーコード3") & "^" & RS("一押し商品コード3"))
'
'wHTML = wHTML & "    <td>" & vbNewLine
'wHTML = wHTML & "      <a href='ProductDetail.asp?Item=" & vItem & "'><img src='cat_hotitem/" & RS("一押し画像ファイル名3") & "' alt='" & RS("一押しメーカー名3") & " " & RS("一押し商品名3") & "' width='190' height='150' border='0'></a><br>" & FormatNumber(vPrice,0) & "円（税込）&nbsp;" & vbNewLine
'wHTML = wHTML & "    </td>" & vbNewLine
'wHTML = wHTML & "  </tr>" & vbNewLine
'
'vPrice = getPrice(RS("一押しメーカーコード4"), RS("一押し商品コード4"))
'vPrice = calcPrice(vPrice, wSalesTaxRate)
'vItem = Server.URLEncode(RS("一押しメーカーコード4") & "^" & RS("一押し商品コード4"))
'
'wHTML = wHTML & "  <tr>" & vbNewLine
'wHTML = wHTML & "    <td>" & vbNewLine
'wHTML = wHTML & "      <a href='ProductDetail.asp?Item=" & vItem & "'><img src='cat_hotitem/" & RS("一押し画像ファイル名4") & "' alt='" & RS("一押しメーカー名4") & " " & RS("一押し商品名4") & "' width='190' height='150' border='0'></a><br>" & FormatNumber(vPrice,0) & "円（税込）&nbsp;" & vbNewLine
'wHTML = wHTML & "    </td>" & vbNewLine
'
'vPrice = getPrice(RS("一押しメーカーコード5"), RS("一押し商品コード5"))
'vPrice = calcPrice(vPrice, wSalesTaxRate)
'vItem = Server.URLEncode(RS("一押しメーカーコード5") & "^" & RS("一押し商品コード5"))
'
'wHTML = wHTML & "    <td>" & vbNewLine
'wHTML = wHTML & "      <a href='ProductDetail.asp?Item=" & vItem & "'><img src='cat_hotitem/" & RS("一押し画像ファイル名5") & "' alt='" & RS("一押しメーカー名5") & " " & RS("一押し商品名5") & "' width='190' height='150' border='0'></a><br>" & FormatNumber(vPrice,0) & "円（税込）&nbsp;" & vbNewLine
'wHTML = wHTML & "    </td>" & vbNewLine
'wHTML = wHTML & "  </tr>" & vbNewLine
'
'wHTML = wHTML & "</table>"
'2012/07/24 ok Del End

'2012/07/24 ok Add Start
wHTML = wHTML & "  <h2 class='subtitle pickup'>" & wLargeCategoryName & "のイチオシ商品"
If RS("一押しコメント") <> "" Then
	wHTML = wHTML & "<span>［" & RS("一押しコメント") & "］</span>"
End If
wHTML = wHTML & "</h2>" & vbNewLine
wHTML = wHTML & "  <ul class='rank'>" & vbNewLine

vCnt = 0
For i = 1 To 5 Step 1
	If GetProductFlag(RS("一押しメーカーコード" & i),RS("一押し商品コード" & i)) = "Y" Then
		vItem = Server.URLEncode(RS("一押しメーカーコード" & i) & "^" & RS("一押し商品コード" & i))
		wHTML = wHTML & "    <li class='rank0" & i-vCnt & "' ><a href='ProductDetail.asp?Item=" & vItem & "'>"
		If RS("商品画像ファイル名_小" & i) <> "" Then
			wHTML = wHTML & "<img src='prod_img/" & RS("商品画像ファイル名_小" & i) & "' alt='" & RS("一押しメーカー名" & i) & " / " & RS("一押し商品名" & i) & "' class='opover'>"
		End If
		wHTML = wHTML & RS("一押しメーカー名" & i) & " / " & RS("一押し商品名" & i) & "</a></li>" & vbNewLine
	Else
		vCnt = vCnt + 1
	End If
Next

wHTML = wHTML & "  </ul>" & vbNewLine
'2012/07/24 ok Add End

wIchioshiHTML = wHTML

End Function

'========================================================================
'
'	Function	単価取り出し
'
'========================================================================
Function GetPrice(pMakerCd, pProductCd)

Dim RSv

GetPrice = 0

'---- 単価取り出し
wSQL = ""
wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.販売単価 "
wSQL = wSQL & "    , a.B品単価 "
wSQL = wSQL & "    , a.個数限定単価 "
wSQL = wSQL & "    , a.B品フラグ "
wSQL = wSQL & "    , a.個数限定数量 "
wSQL = wSQL & "    , a.個数限定受注済数量 "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    Web商品 a WITH (NOLOCK) "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        a.メーカーコード = '" & pMakerCd & "' "
wSQL = wSQL & "    AND a.商品コード     = '" & pProductCd & "'"

'@@@@@@@@@@Response.Write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

If RSv.EOF = True Then
	Exit Function
End If

If RSv("B品フラグ") = "Y" Then

	'---- B品特価
	GetPrice = RSv("B品単価")

Else

	If RSv("個数限定数量") > RSv("個数限定受注済数量") And RSv("個数限定数量") > 0 Then
		'---- 個数限定単価
		GetPrice = RSv("個数限定単価")
	Else
		'---- 販売単価
		GetPrice = RSv("販売単価")
	End If

End If

End Function

' 2012/03/13 GV Del Start
''========================================================================
''
''	Function	中カテゴリー一覧
''
''========================================================================
'Function CreateMidCategoryListHTML()
'
'Dim RSv
'
''---- 中カテゴリー、カテゴリー 取り出し
'wSQL = ""
'wSQL = wSQL & "SELECT a.中カテゴリーコード"
'wSQL = wSQL & "     , a.中カテゴリー名日本語"
'wSQL = wSQL & "     , ISNULL(a.中カテゴリー画像ファイル名,'') AS 中カテゴリー画像ファイル名"
'wSQL = wSQL & "  FROM 中カテゴリー a WITH (NOLOCK)"
'wSQL = wSQL & " WHERE a.大カテゴリーコード = '" & LargeCategoryCd & "'"
'wSQL = wSQL & " ORDER BY a.表示順"
'
''@@@@@@@@@@Response.Write(wSQL)
'
'Set RSv = Server.CreateObject("ADODB.Recordset")
'RSv.Open wSQL, Connection, adOpenStatic
'
'If RSv.EOF = True Then
'	Exit Function
'End If
'
'wHTML = ""
'wHTML = wHTML & "<table width='794' border='0' cellspacing='4' cellpadding='0' id='Shop_LargeCategory_MidCat'>" & vbNewLine
'wHTML = wHTML & "<tr>" & vbNewLine
'
'Do Until RSv.EOF = True
'
'	'---- 左　編集
'	wHTML = wHTML & "    <td>" & vbNewLine
'	wHTML = wHTML & "      <table width='253' height:'100%' border='0' cellspacing='0' cellpadding='0' id='Shop_LargeCategory_SmallCat'>" & vbNewLine
'	wHTML = wHTML & "        <tr>" & vbNewLine
'	wHTML = wHTML & "          <td align='center' class='cat_left'><a href='MidCategoryList.asp?MidCategoryCd=" & RSv("中カテゴリーコード") & "'><img src='cat_img/" & RSv("中カテゴリー画像ファイル名") & "' width='50' height='50' border='0' alt='" & RSv("中カテゴリー名日本語") & "'></a></td>" & vbNewLine
'	wHTML = wHTML & "          <td class='cat_right'><a href='MidCategoryList.asp?MidCategoryCd=" & RSv("中カテゴリーコード") & "'><h3>" & RSv("中カテゴリー名日本語") & "</h3></a><br>" & vbNewLine
'
'	wHTML = wHTML & SetCategory(RSv("中カテゴリーコード"))
'
'	wHTML = wHTML & "          </td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine
'	wHTML = wHTML & "      </table>" & vbNewLine
'	wHTML = wHTML & "    </td>" & vbNewLine
'
'	RSv.MoveNext
'	If RSv.EOF = True Then
' 		wHTML = wHTML & " </tr>" & vbNewLine
'		Exit Do
'	End If
'
'	'---- 中　編集
'	wHTML = wHTML & "    <td>" & vbNewLine
'	wHTML = wHTML & "      <table width='253' height:'100%' border='0' cellspacing='0' cellpadding='0' id='Shop_LargeCategory_SmallCat'>" & vbNewLine
'	wHTML = wHTML & "        <tr>" & vbNewLine
'	wHTML = wHTML & "          <td align='center' class='cat_left'><a href='MidCategoryList.asp?MidCategoryCd=" & RSv("中カテゴリーコード") & "'><img src='cat_img/" & RSv("中カテゴリー画像ファイル名") & "' width='50' height='50' border='0' alt='" & RSv("中カテゴリー名日本語") & "'></a></td>" & vbNewLine
'	wHTML = wHTML & "          <td class='cat_right'><a href='MidCategoryList.asp?MidCategoryCd=" & RSv("中カテゴリーコード") & "' class='link'><h3>" & RSv("中カテゴリー名日本語") & "</h3></a><br>" & vbNewLine
'
'	wHTML = wHTML & SetCategory(RSv("中カテゴリーコード"))
'
'	wHTML = wHTML & "          </td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine
'	wHTML = wHTML & "      </table>" & vbNewLine
'	wHTML = wHTML & "    </td>" & vbNewLine
'
'	RSv.MoveNext
'	If RSv.EOF = True Then
' 		wHTML = wHTML & " </tr>" & vbNewLine
'		Exit Do
'	End If
'
'	'---- 右　編集
'	wHTML = wHTML & "    <td>" & vbNewLine
'	wHTML = wHTML & "      <table width='253' height:'100%' border='0' cellspacing='0' cellpadding='0' id='Shop_LargeCategory_SmallCat'>" & vbNewLine
'	wHTML = wHTML & "        <tr>" & vbNewLine
'	wHTML = wHTML & "          <td align='center' class='cat_left'><a href='MidCategoryList.asp?MidCategoryCd=" & RSv("中カテゴリーコード") & "'><img src='cat_img/" & RSv("中カテゴリー画像ファイル名") & "' width='50' height='50' border='0' alt='" & RSv("中カテゴリー名日本語") & "'></a></td>" & vbNewLine
'	wHTML = wHTML & "          <td class='cat_right'><a href='MidCategoryList.asp?MidCategoryCd=" & RSv("中カテゴリーコード") & "' class='link'><h3>" & RSv("中カテゴリー名日本語") & "</h3></a><br>" & vbNewLine
'
'	wHTML = wHTML & SetCategory(RSv("中カテゴリーコード"))
'
'	wHTML = wHTML & "          </td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine
'	wHTML = wHTML & "      </table>" & vbNewLine
'	wHTML = wHTML & "    </td>" & vbNewLine
'
'	RSv.MoveNext
'	If RSv.EOF = True Then
' 		wHTML = wHTML & " </tr>" & vbNewLine
'		Exit Do
'	End If
'
' 	wHTML = wHTML & " </tr>" & vbNewLine
'Loop
'
'wHTML = wHTML & "</table>" & vbNewLine
'wMidCategoryListHTML = wHTML
'
'RSv.Close
'
'End Function
'
''========================================================================
''
''	Function	カテゴリー一覧
''
''========================================================================
'Function SetCategory(pMidCategoryCd)
'
'Dim RSv
'Dim vHTML
'Dim i
'
''---- 中カテゴリー、カテゴリー 取り出し
'wSQL = ""
'wSQL = wSQL & "SELECT a.カテゴリーコード"
'wSQL = wSQL & "     , a.カテゴリー名"
'wSQL = wSQL & "  FROM カテゴリー a WITH (NOLOCK)"
'wSQL = wSQL & "     , カテゴリー中カテゴリー b WITH (NOLOCK)"
'wSQL = wSQL & " WHERE b.カテゴリーコード = a.カテゴリーコード"
'wSQL = wSQL & "   AND b.中カテゴリーコード = '" & pMidCategoryCd & "'"
'wSQL = wSQL & "   AND A.Webカテゴリーフラグ = 'Y'"
'wSQL = wSQL & " ORDER BY a.表示順"
'
''@@@@@@@@@@Response.Write(wSQL)
'
'Set RSv = Server.CreateObject("ADODB.Recordset")
'RSv.Open wSQL, Connection, adOpenStatic
'
'i = 0
'vHTML = ""
'
'Do Until RSv.EOF = True
'	vHTML = vHTML & "            <a href='SearchList.asp?i_type=c&s_category_cd=" & RSv("カテゴリーコード") & "'>- " & RSv("カテゴリー名") & "</a><br>" & vbNewLine
'
'	i = i + 1
'	If AllFl <> "Y" And i >= 4 Then
'		Exit Do
'	End If
'
'	RSv.MoveNext
'Loop
'
'If AllFl <> "Y" Then
'	vHTML = vHTML & "            <a href='LargeCategoryList.asp?AllFl=Y&LargeCategoryCd=" & LargeCategoryCd & "'><strong>+ 全てを見る</strong></a>"
'End If
'
'RSv.Close
'
'SetCategory = vHTML
'
'End Function
'
''========================================================================
''
''	Function	トピックス News
''
''========================================================================
'Function CreatewNewsHTML()
'
'Dim RSv
'
''---- 商品記事 取り出し
'wSQL = ""
'' 2012/01/20 GV Mod Start
''wSQL = wSQL & "SELECT TOP 5 * "
''wSQL = wSQL & "FROM "
''wSQL = wSQL & "(SELECT "
''wSQL = wSQL & "      a.記事番号 "
''wSQL = wSQL & "    , a.記事日付 "
''wSQL = wSQL & "    , a.記事タイトル "
''wSQL = wSQL & " FROM "
''wSQL = wSQL & "      商品記事 a WITH (NOLOCK) "
''wSQL = wSQL & "    , 商品記事中カテゴリー b WITH (NOLOCK) "
''wSQL = wSQL & "    , 中カテゴリー c WITH (NOLOCK) "
''wSQL = wSQL & "WHERE b.記事番号 = a.記事番号"
''wSQL = wSQL & "  AND c.中カテゴリーコード = b.中カテゴリーコード"
''wSQL = wSQL & "  AND c.大カテゴリーコード = '" & LargeCategoryCd & "'"
''wSQL = wSQL & "  AND ((getDate() BETWEEN a.表示期間From AND a.表示期間To)"
''wSQL = wSQL & "    OR (a.表示期間From IS NULL AND a.表示期間To IS NULL))"
''wSQL = wSQL & "UNION "
''wSQL = wSQL & "SELECT "
''wSQL = wSQL & "      a.記事番号 "
''wSQL = wSQL & "    , a.記事日付 "
''wSQL = wSQL & "    , a.記事タイトル "
''wSQL = wSQL & " FROM "
''wSQL = wSQL & "      商品記事 a WITH (NOLOCK)  "
''wSQL = wSQL & "WHERE (a.大カテゴリーコード = '" & LargeCategoryCd & "'"    '2010/12/07 an mod
''wSQL = wSQL & "   OR a.記事区分 = '一般記事')"                             '2010/12/07 an add
''wSQL = wSQL & "  AND ((getDate() BETWEEN a.表示期間From AND a.表示期間To)"
''wSQL = wSQL & "    OR (a.表示期間From IS NULL AND a.表示期間To IS NULL)) "
''wSQL = wSQL & ")AS inLineView "
''wSQL = wSQL & "ORDER BY 記事日付 DESC"
'wSQL = wSQL & "SELECT DISTINCT TOP 5 "
'wSQL = wSQL & "      a.記事番号 "
'wSQL = wSQL & "    , a.記事日付 "
'wSQL = wSQL & "    , a.記事タイトル "
'wSQL = wSQL & "FROM "
'wSQL = wSQL & "    Web商品記事TOP10 a WITH (NOLOCK) "
'wSQL = wSQL & "WHERE "														' 2012/01/20 GV Mod Mailにて調整依頼の為、条件変更
'wSQL = wSQL & "        (    (    a.記事区分 = '一般記事' "
'wSQL = wSQL & "              OR  a.記事区分 = '個別記事') "
'wSQL = wSQL & "         AND a.大カテゴリーコード = '" & LargeCategoryCd & "') "
'wSQL = wSQL & "    OR  (    a.記事区分 = '一般記事' "
'wSQL = wSQL & "         AND a.大カテゴリーコード = '' "
'wSQL = wSQL & "         AND a.中カテゴリーコード = '') "
'wSQL = wSQL & "ORDER BY "
'' 2012/01/23 GV Mod Start
''wSQL = wSQL & "      a.記事日付 DESC "
'wSQL = wSQL & "      a.記事番号 DESC "
'' 2012/01/23 GV Mod End
'' 2012/01/20 GV Mod End
'
''@@@@Response.Write(wSQL)
'
'Set RSv = Server.CreateObject("ADODB.Recordset")
'RSv.Open wSQL, Connection, adOpenStatic
'
''----- NewsHTML編集
'wNewsHTML = ""
'
'If RSv.EOF = False Then
'	wNewsHTML = wNewsHTML & "<table width='794' border='0' cellspacing='0' cellpadding='0'>" & vbNewLine
'
'	Do until RSv.EOF = True
'		wNewsHTML = wNewsHTML & "  <tr>" & vbNewLine
'		wNewsHTML = wNewsHTML & "    <td class='honbun'>" & fFormatDate(RSv("記事日付")) & " <a href='News.asp?NewsNo=" & RSv("記事番号") & "' class='link'>" & RSv("記事タイトル") & "</a></td>" & vbNewLine
'		wNewsHTML = wNewsHTML & "  </tr>" & vbNewLine
'		RSv.MoveNext
'	Loop
'
'	wNewsHTML = wNewsHTML & "</table>" & vbNewLine
'End If
'
'RSv.Close
'
'End Function
'
''========================================================================
''
''	Function	トピックス 新製品
''
''========================================================================
'Function CreateNewItemHTML()
'
'Dim RSv
'
''---- 新製品 取り出し
'wSQL = ""
'' 2012/01/20 GV Mod Start
''wSQL = wSQL & "SELECT DISTINCT TOP 10"
''wSQL = wSQL & "       a.発売日"
''wSQL = wSQL & "     , a.メーカーコード"
''wSQL = wSQL & "     , a.商品コード"
''wSQL = wSQL & "     , a.商品名"
''wSQL = wSQL & "     , b.メーカー名"
''wSQL = wSQL & "     , c.カテゴリー名"
''wSQL = wSQL & "  FROM Web商品 a WITH (NOLOCK)"
''wSQL = wSQL & "     , メーカー b WITH (NOLOCK)"
''wSQL = wSQL & "     , カテゴリー c WITH (NOLOCK)"
''wSQL = wSQL & "     , カテゴリー中カテゴリー d WITH (NOLOCK)"
''wSQL = wSQL & "     , 中カテゴリー e WITH (NOLOCK)"
''wSQL = wSQL & " WHERE b.メーカーコード = a.メーカーコード"
''wSQL = wSQL & "   AND c.カテゴリーコード = a.カテゴリーコード"
''wSQL = wSQL & "   AND d.カテゴリーコード = a.カテゴリーコード"
''wSQL = wSQL & "   AND e.中カテゴリーコード = d.中カテゴリーコード"
''wSQL = wSQL & "   AND e.大カテゴリーコード = '" & LargeCategoryCd & "'"
''wSQL = wSQL & "   AND a.終了日 IS NULL"
''wSQL = wSQL & "   AND a.Web商品フラグ = 'Y'"
''wSQL = wSQL & " ORDER BY a.発売日 DESC"
'wSQL = wSQL & "SELECT DISTINCT TOP 10 "
'wSQL = wSQL & "      a.発売日 "
'wSQL = wSQL & "    , a.メーカーコード "
'wSQL = wSQL & "    , a.商品コード "
'wSQL = wSQL & "    , a.商品名 "
'wSQL = wSQL & "    , a.メーカー名 "
'wSQL = wSQL & "    , a.カテゴリー名 "
'wSQL = wSQL & "FROM "
'wSQL = wSQL & "    Web新製品TOP10 a WITH (NOLOCK) "
'wSQL = wSQL & "WHERE "
'wSQL = wSQL & "        a.大カテゴリーコード = '" & LargeCategoryCd & "' "
'wSQL = wSQL & "ORDER BY "
'wSQL = wSQL & "      a.発売日 DESC "
'' 2012/01/20 GV Mod End
'
''@@@@@@@@@@Response.Write(wSQL)
'
'Set RSv = Server.CreateObject("ADODB.Recordset")
'RSv.Open wSQL, Connection, adOpenStatic
'
''----- 新製品HTML編集
'wNewItemHTML = ""
'
'If RSv.EOF = False Then
'	wNewItemHTML = wNewItemHTML & "<table width='794' border='0' cellspacing='0' cellpadding='0'>" & vbNewLine
'
'	Do until RSv.EOF = True
'		wNewItemHTML = wNewItemHTML & "  <tr>" & vbNewLine
'		wNewItemHTML = wNewItemHTML & "     <td class='honbun'>" & fFormatDate(RSv("発売日")) & " <a href='ProductDetail.asp?Item=" & RSv("メーカーコード") & "^" & Server.URLEncode(RSv("商品コード")) & "' class='link'>" & RSv("商品名") & " " & RSv("カテゴリー名") & " (" & RSv("メーカー名") & ")</a></td>" & vbNewLine
'		wNewItemHTML = wNewItemHTML & "  </tr>" & vbNewLine
'		RSv.MoveNext
'	Loop
'
'	wNewItemHTML = wNewItemHTML & "</table>" & vbNewLine
'
'End If
'
'RSv.Close
'
'End Function
' 2012/03/13 GV Del End

'========================================================================
'
'	Function	SALE&OUTLET商品
'	2012/07/24 ok Add
'========================================================================
Function CreateSaleAndOutletHTML()

Dim RSv
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
wSQL = ""
wSQL = wSQL & "SELECT "
' 2012/09/03 GV #1426 Mod Start
'wSQL = wSQL & "    TOP 5 "
wSQL = wSQL & "    TOP 20 "
' 2012/09/03 GV #1426 Mod End
wSQL = wSQL & "      a.商品コード "
wSQL = wSQL & "    , a.商品名 "
wSQL = wSQL & "    , a.メーカーコード "
wSQL = wSQL & "    , a.メーカー名 "
wSQL = wSQL & "    , a.商品画像ファイル名_小 "
wSQL = wSQL & "    , a.販売単価 "
wSQL = wSQL & "    , a.前回販売単価 "
wSQL = wSQL & "    , a.ASK商品フラグ "
wSQL = wSQL & "    , a.B品フラグ "
wSQL = wSQL & "    , a.個数限定数量 "
wSQL = wSQL & "    , a.個数限定単価 "
wSQL = wSQL & "    , a.個数限定受注済数量 "
wSQL = wSQL & "    , a.前回単価変更日 "
wSQL = wSQL & "    , a.B品フラグ "
wSQL = wSQL & "    , a.B品単価 "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    Webセール商品 a WITH (NOLOCK) "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "    a.セール区分番号 BETWEEN 1 AND 4"
wSQL = wSQL & " AND a.大カテゴリーコード = '" & LargeCategoryCd & "' "
wSQL = wSQL & "ORDER BY NEWID() "

'@@@@@@@@@@Response.Write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic
wHTML = ""

If RSv.EOF = false Then
	'----- セール商品HTML編集
	wHTML = wHTML & "<h2 class='subtitle_red'>" & wLargeCategoryName & "のSALE &amp; OUTLET</h2>" & vbNewLine
	wHTML = wHTML & "<div class='box'><div class='box_inner01'>" & vbNewLine
	wHTML = wHTML & "  <ul class='list'>" & vbNewLine

	Do Until RSv.EOF = True OR dcnt > 4
' 2012/09/03 GV #1426 Add Start
		ReDim Preserve w_MakerCd(cnt)
		w_MakerCd(cnt) = RSv("メーカーコード")
		ReDim Preserve w_ItemCd(cnt)
		w_ItemCd(cnt) = RSv("商品コード")
		wHTML1 = ""
' 2012/09/03 GV #1426 Add End
		wHTML1 = wHTML1 & "    <li><a href='ProductDetail.asp?Item=" & Server.URLEncode(RSv("メーカーコード") & "^" & RSv("商品コード")) & "'>"
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
'					wHTML1 = wHTML1 & FormatNumber(v_exprice,0) & "円（税込）↓<br>" & vbNewLine
'2013/03/19 GV mod end   <----
' 2012/09/03 GV #1426 Add Start
					ReDim Preserve w_price1(cnt)
					w_price1(cnt) = FormatNumber(v_exprice,0)
' 2012/09/03 GV #1426 Add End
				'B品、限定品は販売価格を旧価格として表示
				Else
'2013/03/19 GV mod start ---->
'前回単価はしばらく表示させない
'					wHTML1 = wHTML1 & FormatNumber(v_price,0) & "円（税込）↓<br>" & vbNewLine
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
				wHTML1 = wHTML1 & "(税込&nbsp;" & FormatNumber(v_price,0) & "円)"
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
				wHTML1 = wHTML1 & "(税込&nbsp;" & FormatNumber(v_price,0) & "円)"
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
				wHTML1 = wHTML1 & "(税込&nbsp;" & FormatNumber(v_price,0) & "円)"
'2013/03/19 GV mod end   <----
' 2012/09/03 GV #1426 Add Start
				ReDim Preserve w_price2(cnt)
				w_price2(cnt) = FormatNumber(v_price,0)
' 2012/09/03 GV #1426 Add End
			End If

			wHTML1 = wHTML1 & "</span></li>" & vbNewLine

		Else
			'---- B品単価
			If RSv("B品フラグ") = "Y" Then
				v_price = calcPrice(RSv("B品単価"), wSalesTaxRate)
'2013/03/19 GV mod start ---->
'				wHTML1 = wHTML1 & "【わけあり品特価】</span><a class='tip'>ASK<span>" & FormatNumber(v_price,0) & "円(税込)</span>"
				wHTML1 = wHTML1 & "【わけあり品特価】</span><a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RSv("B品単価"),0) & "円(税抜)</span><br>"
				wHTML1 = wHTML1 & "<span class='inc-tax'>(税込&nbsp;" & FormatNumber(v_price,0) & "円)</span>"
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
				wHTML1 = wHTML1 & "【限定特価】</span><a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(v_price,0) & "円(税抜)</span><br>"
				wHTML1 = wHTML1 & "<span class='inc-tax'>(税込&nbsp;" & FormatNumber(v_price,0) & "円)</span>"
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
				wHTML1 = wHTML1 & "<span class='inc-tax'>(税込&nbsp;" & FormatNumber(v_price,0) & "円)</span>"
'2013/03/19 GV mod end   <----
' 2012/09/03 GV #1426 Add Start
				ReDim Preserve w_price2(cnt)
				w_price2(cnt) = FormatNumber(v_price,0)
' 2012/09/03 GV #1426 Add End
			End If

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

	wHTML = wHTML & "  </ul>" & vbNewLine
	wHTML = wHTML & "</div></div>" & vbNewLine
End If
wSaleAndOutletHTML = wHTML

RSv.Close

End Function

'========================================================================
'
'	Function	Web商品フラグチェック
'
'========================================================================
Function GetProductFlag(pMakerCd, pProductCd)

Dim RSv
GetProductFlag = ""

'---- Web商品フラグ取り出し
wSQL = ""
wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.Web商品フラグ "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    Web商品 a WITH (NOLOCK) "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        a.メーカーコード = '" & pMakerCd & "' "
wSQL = wSQL & "    AND a.商品コード     = '" & pProductCd & "'"

'@@@@@@@@@@Response.Write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

If RSv.EOF = false Then
	GetProductFlag = RSv("Web商品フラグ")
End If

RSv.Close

End Function

'========================================================================
'
'	Function	Close database
'
'========================================================================
'
Function close_db()

Connection.Close
Set Connection = Nothing    '2011/08/01 an add

End Function

'========================================================================
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<meta name="robots" content="noindex,nofollow">
<title><%=wLargeCategoryName%> 一覧｜サウンドハウス</title>
<%=wMetaTag%>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/shop.css" type="text/css">
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
        <li class="now"><%=wLargeCategoryName%></li>
      </ul>
    </div></div></div>

<%=fIncludeInsertHTML(wInsertHTMLPath1)%>

<!-- 大カテゴリーについて・カテゴリーから選ぶ・最新ニュース・新製品 -->

<%=wStaticHTML(0)%>

<%=wIchioshiHTML%>

<%=wStaticHTML(1)%>

<%=fIncludeInsertHTML(wInsertHTMLPath2)%>

<%=wSaleAndOutletHTML%>

<%=wStaticHTML(2)%>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<div class="tooltip"><p>ASK</p></div>
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript" src="jslib/LargeCategoryList.js?20130805"></script>
<script type="text/javascript" src="../jslib/jquery.carouFredSel-5.5.0-packed.js"></script>
<script type="text/javascript" src="jslib/ask.js?20140401a"></script>
<script type="text/javascript">
var userAgent = window.navigator.userAgent.toLowerCase();
var appVersion = window.navigator.appVersion.toLowerCase();
if(userAgent.indexOf("msie")!=-1){
	if(appVersion.indexOf("msie 7.")!=-1){
		$("ul.cate_tab li a span").each(function(){
			if($(this).height()>20){
				$(this).css("top","4px");
			}
		});
	}
}
</script>
</body>
</html>