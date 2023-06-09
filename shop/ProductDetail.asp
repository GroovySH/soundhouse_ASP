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
<!--#include file="../3rdParty/EAgency.inc"-->
<%
'========================================================================
'
'	商品詳細ページ
'更新履歴
'2004/12/16 hn ASK, 直の場合のポップアップ画面を中止しリンクに変更（MACで問題ありのため)
'2004/12/17 hn 数量入力欄追加
'2005/01/05 hn リンク時のTarget指定を変更
'2005/01/14 hn パーツ、オプションデータ抽出にカテゴリー別テーブルも追加
'2005/01/27 hn ASKクリックでASK単価表示画面をポップアップするように変更
'              ｢この商品に関する問合せ｣ボタン追加
'2005/02/03 hn メーカー名+商品名の表示を変更 ()/の前後に半角スペース追加
'2005/02/09 hn ｢この商品のお問合せ｣からのリンク時のパラメータにURLEncode適用
'2005/02/16 個数限定数量単価取り出し時の条件強化　個数限定数量＞0を追加
'2005/02/23 メーカー名、カテゴリーからSearchListへリンク追加
'2005/03/15 関連シリーズ商品表示追加
'2005/05/12 関連記事，関連ニュースを関連記事1－4にまとめる
'2005/05/13 メーカー関連リンクを追加
'2005/06/27 商品レビューを追加
'2005/07/01 納期表示をおよその日数に変更
'2005/08/18 ｢この商品のお問合せ｣パラメータを中カテゴリー日本語に変更
'2005/10/18 Web納期非表示フラグ対応
'2005/11/01 試聴・Movieポップアップページへのリンク追加
'2005/12/01 ASK価格表示ページへのパラメータをServer.URLEncodeする
'2006/01/10 試聴、動画、関連記事、製品マニュアル、新製品紹介のリンクにhttpが含まれている場合は外部リンクとする。
'2006/03/07 単価の数字をイメージ化
'2006/04/03 パンくず追加
'2006/04/07 Web商品フラグチェックを追加
'2006/04/21 後継機種が設定されている商品は後継機種商品を表示
'2006/04/25 取扱中止、灰番で在庫無し商品のみ後継機種が設定されている商品は後継機種商品を表示
'2006/10/18 在庫状況　｢未定｣を追加
'2006/12/21 色規格別在庫の終了日チェックを追加（色規格あり商品で色規格なしが表示されてしまうため）
'2007/01/08 emタグ追加
'2007/03/01 CreateSpecHTMLのパラメータ変更
'2007/03/02 在庫状況　表示の色を変更
'2007/03/13 新NAVIに変更
'2007/04/05 ログインしていれば商品レビューを登録できるように変更(１回のみ）
'2007/04/18 B品を表示追加
'2007/04/20 購入実績がありログインしていれば商品レビューを登録できるように変更(１回のみ）
'2007/04/24 出荷通知からの商品レビュー対処
'2007/04/25 商品レビューの変更（顧客別商品レビュー一覧ボタン、ハンドルネーム）
'2007/05/08 商品備考インサートURL1,2 追加
'2007/05/14 商品備考インサートサイズ指定 追加
'2007/05/14 廃番で在庫あり商品の在庫状況を「在庫限り」とする
'2007/05/15 カート内容に送料表示追加
'2007/05/25 ショップコメントフラグ対処
'2007/05/30	色規格無しで呼び出され、該当商品が色規格ありの場合は、色規格選択ドロップダウンを表示する。
'2007/06/05 関連リンクを追加(NaviLeftShopから移動）
'2007/06/15 レビューにリンク文字入力チェック追加
'2007/06/25 商品レビューに注意コメント追加
'2007/07/05 色規格選択時に変更後画面が切り替わるまでにカートボタンを押されたとき違う色規格を送信するエラーの対処
'2007/07/11 色規格商品画像ファイル名1-4の考慮
'2007/07/17 友達に勧める、ウィッシュリストに入れるボタンを追加, カートの中身を見るを共通関数に変更
'2007/08/23 商品アクセス件数登録（ページビュー）　同一セッション同一商品で1回
'2007/09/10 商品アクセス件数登録を年月別に変更
'2007/09/12 該当商品なしのときにSearchNotFound.aspを表示
'           後継機種ありのときは、SearchList.aspで後継機種表示(i_type=successor)
'2007/10/22 レビュー書き込みチェック強化 WriteReview=Yでも購入回数チェック
'2007/11/20 カートボタンの下へ商品ID表示追加
'2007/12/13 廃番+引当可能在庫＝1の時に「在庫限り」と表示
'2007/12/27 在庫状況　表示の色を変更
'2008/01/11 個数限定単価もB品と同様の単価表示に変更
'2008/01/28 色規格が1種類しかない場合の対応
'2008/05/07 入力データチェック強化
'2008/05/21 レビュー投稿チェック時のEOFチェック強化
'2008/07/31 色規格別在庫.終了日 IS NULLデータの扱い変更
'2008/09/13 出荷通知のリンクからレビュー作成で呼ばれた時は受注番号からUserIDを取り出しレビュー入力画面を表示するように変更
'2008/09/16 (変更依頼#503)個数限定数量の表示を次のように変更
'						4以下　現行どおり/5-9 限定5個/10-14　限定10個/15-19　限定15個/20以上、限定20個
'
'2008/12/19 リニューアル ********
'2009/04/27 色規格が確定されていないときは、在庫状況を非表示
'2009/05/27 動画、試聴、メーカーリンク、マニュアル　アイコン変更
'2009/08/06 価格表示サイズを変更(Style指定）
'2009/10/26 商品備考インサートサイズH1,2の高さ制限を削除
'2009/12/17 hn レコメンド用変更（商品アクセスログ出力、レコメンド表示）
'2010/01/20 an パン屑リストの最下層リンクにカテゴリーコードを追加し、i_tyep=cmに変更。中カテゴリーへのリンクが入っていなかったので追加
'2010/01/26 hn 色規格指定ありで、1個しかない場合の不具合を修正
'2010/01/29 an B品特価→訳あり品に表記を変更
'2010/02/06 if-web 価格表示部に「特価：」追加
'2010/02/22 st 訳あり品→わけあり品に表記を変更
'2010/03/04 an レコメンド商品アクセスログ登録を有効化
'2010/03/08 hn レビューはい・いいえを画像にし、JavaScriptで実行　（Bot対策）
'2010/04/06 an レコメンドのASK商品をASK表示に修正
'2010/04/21 an レコメンドのASK商品の価格表示の間違いを修正
'2010/05/17 ko-web 検索対策のためHTMLタグ（h1,h2,h3,p,strong）追加
'2010/06/10 an SEO対策で<link>タグ追加
'2010/07/01 an HTMLレイアウト修正
'2010/08/23 an meta descripion,keywordsに商品情報を自動セットするように修正
'2010/08/30 an Twitterつぶやくボタン追加
'2010/09/27 an Twitterつぶやくボタン位置変更
'2010/11/04 GV(dy) #724 商品詳細の「在庫状況」を "完売御礼" の画像が表示されない時のみ表示するように修正
'2010/11/10 an 関連商品に限定特価、B品特価を反映するように修正。レコメンド、関連シリーズにB品特価を反映するように修正
'2010/11/26 an 廃番品でも色規格商品の在庫があれば後継機種を表示しないように修正
'2010/12/28 hn パーツオプション内　wProdTermFl→vProdTermFl に変更　wProdTermFlをつぶすため
'2011/02/18 GV(dy) #826 送料完全無料表示の対応
'2011/03/18 GV(dy) #731 Style/StyleNaviLeftShop.css のstylesheet定義を追加
'2011/06/09 hn 廃番で在庫なし＋発注なし　の時に完売とするように変更
'2011/06/15 if-web 送料完全無料表示に沖縄・離島を除く旨を追記
'2011/08/01 an #1087 Error.aspログ出力対応
'2011/09/09 an #816 レビューメンテナンス対応でレビューID表示追加
'2011/10/19 hn 1063 ASK表示方法変更
'2011/11/22 an #1150 リッチスニペット対応, OGP対応
'2012/01/10 an レコメンド商品アクセスログ出力停止
'2012/01/18 GV カスタマーレビュー用のデータ取得時、「投稿日」列での並び替えを 「ID」列での並び替えに変更
'2012/01/18 GV 関連シリーズ商品用, オプション用およびパーツ用のデータ取得 SELECT文へ LACクエリー案を適用 (あわせて WITH (NOLOCK) 付加)
'2012/01/23 GV 「商品レビュー」テーブルから「商品レビュー集計」テーブル使用に変更 (CreateReviewHTML()プロシージャ)
'2012/02/20 na 「商品アクセス数」レスポンス対策のため停止
'2012/04/10 GV レコメンド表示レイアウト変更
'2012/05/01 GV 代替商品表示機能追加
'2012/07/10 GV 商品詳細ページデザイン変更
'2012/08/01 ok 右サイドの商品情報のメーカー/商品/カテゴリー名部分をリンクありに変更
'2012/08/27 ok 関連パーツ、オプションのカテゴリー表示対応
'2012/10/30 nt 関連コンテンツ表示項目追加
'2013/05/17 GV #1507 レビュー編集機能
'2013/05/22 GV #1505 さぶみっと！レコメンド対応
'2013/08/07 if-web 旧レコメンド（チームラボ）をコメントアウト
'2013/08/14 GV 代替商品取得メソッド で、DBから取得した値がない場合の処理を追加
'2014/03/19 GV 消費税増税に伴う2重表示対応
'
'========================================================================

On Error Resume Next

Dim wUserID

Dim maker_cd
Dim product_cd
Dim iro
Dim kikaku
Dim item
Dim item_list()
Dim item_cnt

Dim ReviewAll
Dim WriteReview
Dim OrderNo
'Dim iShop				'2013/05/17 GV #1507 comment out

Dim wMakerName
Dim wMakerNameNoKana     '2010/08/30 an add
Dim wProductName
Dim wCategoryCode
Dim wTitleWithLink
Dim wKoukeiMakerCd
Dim wKoukeiProductCd
Dim wLargeCategoryCd
Dim wMidCategoryCd
Dim wCanWriteReviewFl
Dim wPrefecture
Dim wHandleName
Dim wIroKikakuSelectMsg
Dim wLargeCategoryName   '2010/08/23 an add
Dim wMidCategoryName     '2010/08/23 an add
Dim wCategoryName        '2010/08/23 an add
Dim wTokucho             '2010/08/23 an add
Dim wFreeShippingFlag		' 2011/02/18 GV Add
Dim s_category_cd        '2011/09/09 an add For NaviLeftShop
Dim wOptionPartsTitleFlag		'2012/08/29 ok Add

Dim wIroKikakuCombo

Dim wPictureHTML
Dim wKanrenLinkHTML
Dim wTokuchoHTML
Dim wSpecHTML
Dim wOptionHtml
Dim wPartsHtml
Dim wReviewHTML

Dim wProductHTML
Dim wHyoukaHTML
Dim wCartHTML
Dim wSeriesHTML
Dim wRecommendHTML
Dim wRecommendBuyHTML	' 2012/04/10 GV Add
DIm wSubItemHTML        ' 2012/05/01 GV Add
Dim wViewHTML		' 2012/07/10 GV Add

Dim Connection
Dim RS

Dim wTitle
Dim wSalesTaxRate
Dim wProdTermFl
Dim wPrice
Dim wTaxedPrice			'2014/03/19 GV add
Dim wOptionPartsFl
Dim wIroKikakuSelectedFl
Dim wNoData
Dim wIroKikakuFl       '2010/11/26 an add
Dim wIroKikakuZaikoFl  '2010/11/26 an add
Dim wIroKikakuHacchuuFl  '2011/06/09 an add
Dim wMainProdPic        '2011/11/22 an add

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim wNum		'数字画像

Dim wSQL
Dim wHTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add

Dim wSeriesCd		'2012/10/30 nt Add
Dim wContentsHTML	'2012/10/30 nt Add

Dim wTwPriceLabel
Dim wTwPriceData
Dim wTwInventoryData

'2013/05/17 GV #1505 add start
Dim wRecommendJS
Dim wRecommendBuyJS
Dim wEAProductDetailData
Dim wEAInventoryData
Dim wEAPrice
Dim wEAPriceExcTax
Dim wEAIroKikakuData
'2013/05/17 GV #1505 add end

'========================================================================

Response.buffer = true

wUserID = Session("UserID")

'---- Get input data
maker_cd = ReplaceInput(Trim(Request("maker_cd")))
product_cd = ReplaceInput(Trim(Request("product_cd")))
iro = ReplaceInput(Trim(Request("iro")))
kikaku = ReplaceInput(Trim(Request("kikaku")))
item = ReplaceInput(Trim(Request("item")))

if Trim(Request("parm")) <> "" then
	item = ReplaceInput(Trim(Request("parm")))
end if

iro = ""
kikaku = ""

if item <> "" then
	item_cnt = cf_unstring(item, item_list, "^")
	maker_cd = item_list(0)
	product_cd = item_list(1)
	if item_cnt > 2 then
		iro = item_list(2)
		if item_cnt > 3 then
			kikaku = item_list(3)
		end if
	end if
end if

'----商品レビュー用パラメータ
ReviewAll = ReplaceInput(Request("ReviewAll"))
WriteReview = ReplaceInput(UCase(Request("WriteReview")))

'2013/07/10 GV #1507 add start
'旧リンクでアクセスしてきた場合、リダイレクト
If (WriteReview = "Y") Then
	If  gPhoneType = "SP" Then
		Response.Redirect g_HTTPS & "sp/shop/LoginCheck.asp?RtnURL=" & g_HTTPS & "sp/shop/ReviewWrite.asp?Item=" & Server.URLEncode(item)
	Else
		Response.Redirect g_HTTPS & "shop/LoginCheck.asp?RtnURL=" & g_HTTPS & "shop/ReviewWrite.asp?Item=" & Server.URLEncode(item)
	End If
End If
'2013/07/10 GV #1507 add end

OrderNo = ReplaceInput(Request("OrderNo"))
if (OrderNo <> "" and isNumeric(OrderNo) = false) OR OrderNo = "" then
	OrderNo = 0
end if

'iShop = ReplaceInput(Trim(Request("iShop")))		'2013/05/17 GV #1507 comment out

'2013/05/22 GV #1505 add start
wRecommendJS = ""
wRecommendBuyJS = ""
wEAInventoryData = ""
wEAPrice = 0
wEAPriceExcTax = 0
'2013/05/22 GV #1505 add start

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "ProductDetail.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
end if                                           '2011/08/01 an add e

if Err.Description <> "" then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

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

'---- 消費税率取出し
call getCntlMst("共通","消費税率","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'消費税率
wSalesTaxRate = Clng(wItemNum1)

'---- 商品情報取り出し
call GetProduct()
if wMSG <> "" OR wKoukeiMakerCd <> "" OR wNoData = "Y" then
	exit function
end if

'---- 商品画像
call CreatePictureHTML()

'---- 関連リンク
Call CreateKanrenLinkHTML()

'---- 特徴
call CreateTokuchoHTML()

'---- スペック
call CreateSpecificationHTML()

'----- オプションHTML作成
wOptionPartsFl = false
call CreateOptionHTML()

'---- パーツHTML作成
call CreatePartsHTML()

wCanWriteReviewFl = "N"

'---- カスタマーレビュー、評価HTML作成
if RS("B品フラグ") <> "Y" then
	call CreateReviewHTML()

	'---- 出荷通知からのリンクで受注番号が渡された場合は、UserID取り出し
	if OrderNo <> 0 AND wUserID = "" then
		call GetUserID()
	end if

	'---- 商品レビュー登録済チェック
	if wProdTermFl <> "Y" AND wUserID <> "" then
		call CheckReview()
	end if
end if

'==== ここから右側 ================================
'---- メーカー/商品HTML作成
call CreateProductHTML()

'---- 評価HTML作成
'CreateReviewで一緒に作成済

'----- カート情報HTML作成（共通関数）
wCartHTML = fCreateCartHtml()

'----- 商品代替HTML作成・表示 GV 2012/05/01
Call GetSubstituteItem()

'----- レコメンド結果表示		2009/12/17
call CreateRecommendHTML()

'---- GV Add Start 2012/04/10 
call CreateRecommendBuyHTML()
'---- GV Add End 2012/04/10 

'---- 2012/10/30 nt add Start
Call CreateContentsHTML()
'---- 2012/10/30 nt add End

'----- 関連シリーズ商品HTML作成
'if RS("シリーズコード") <> "" then
'	call CreateSeriesHTML()
'end if

'=================================================
'2013/08/07 if-web del s
'----- ログインしていれば、最近チェックした商品を追加作成
'if wUserID <> "" then
'	call CreateViewedProductList()	'2012/07/10 GV Add
'	call AddViewdProduct()
'end if
'2013/08/07 if-web del e

'----- 商品アクセス件数登録	'2012/02/20 na レスポンス対策のため停止
'call SetAccessCount()

'----- レコメンド商品アクセスログ登録   2009/12/17 add 2010/03/04 an 有効化 2012/01/10 an 停止
'call AddRecommendAccessLog()

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

Dim vIroCnt
Dim vKikakuCnt
Dim vProdPic(4)

Dim RSv

'---- 色規格あり商品かどうかのチェック
wSQL = ""
wSQL = wSQL & "SELECT a.色"
wSQL = wSQL & "     , a.規格"
wSQL = wSQL & "     , a.引当可能数量"    '2010/11/26 an add
wSQL = wSQL & "     , a.発注数量"    			'2011/06/09 an add
wSQL = wSQL & "     , a.商品ID"			'2013/05/22 GV #1505 add
wSQL = wSQL & "  FROM Web色規格別在庫 a WITH (NOLOCK)"
wSQL = wSQL & " WHERE a.メーカーコード = '" & maker_cd & "'"
wSQL = wSQL & "   AND a.商品コード = '" & Replace(product_cd, "'", "''") & "'"	' 2012/01/23 GV Mod (コード内にシングルクオーテーションが存在した場合の対応)
wSQL = wSQL & "   AND a.終了日 IS NULL"
wSQL = wSQL & " ORDER BY a.色"
wSQL = wSQL & "     , a.規格"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

wIroKikakuCombo = ""
vIroCnt = 0
vKikakuCnt = 0
wIroKikakuFl = "N"        '2010/11/26 an add
wIroKikakuZaikoFl = "N"   '2010/11/26 an add
wIroKikakuHacchuuFl = "N"   '2011/06/09 an add

if RSv.EOF = false then

	if RSv.RecordCount > 1 OR Trim(RSv("色")) <> "" OR Trim(RSv("規格")) <> "" then	'2010/01/26 hn change

		wIroKikakuFl = "Y"  '色規格有   2010/11/26 an add

		'2013/05/22 GV #1505 add start
		Set wEAIroKikakuData = CreateObject("Scripting.Dictionary")
		wEAIroKikakuData.Item("code") = ""
		wEAIroKikakuData.Item("stock") = 0
		wEAIroKikakuData.Item("iro") = ""
		wEAIroKikakuData.Item("kikaku") = ""
		'2013/05/22 GV #1505 add end

		'----色規格指定で1個しかない場合は、その色規格をセットする
		if RSv.RecordCount = 1 AND (Trim(RSv("色")) <> "" OR Trim(RSv("規格")) <> "") then	'2010/01/26 hn add
			iro = Trim(RSv("色"))
			kikaku = Trim(RSv("規格"))
		end if

		wIroKikakuCombo = wIroKikakuCombo & "                    <p class='color'>" & vbNewLine
		wIroKikakuCombo = wIroKikakuCombo & "                        <select name='IroKikaku' onChange='IroKikaku_onChange(this.form);'>" & vbNewLine
		wIroKikakuCombo = wIroKikakuCombo & "                            <option value=''>選択" & vbNewLine

		Do until RSv.EOF = true

			'2013/05/22 GV #1505 add start
			If (wEAIroKikakuData.Item("code") = "") Then
				wEAIroKikakuData.Item("code") = Trim(RSv("商品ID"))
				wEAIroKikakuData.Item("iro") = Trim(RSv("色"))
				wEAIroKikakuData.Item("kikaku") = Trim(RSv("規格"))
			End If
			'2013/05/22 GV #1505 add end

			'---- 後継機種表示チェック対応  2010/11/26 an add s
			if RSv("引当可能数量") > 0 then
				wIroKikakuZaikoFl = "Y"

				'2013/05/22 GV #1505 add start
				If (wEAIroKikakuData.Item("stock") = 0) Then
					wEAIroKikakuData.Item("code") = Trim(RSv("商品ID"))
					wEAIroKikakuData.Item("iro") = Trim(RSv("色"))
					wEAIroKikakuData.Item("kikaku") = Trim(RSv("規格"))
					wEAIroKikakuData.Item("stock") = 1
				End If
				'2013/05/22 GV #1505 add end
			end if  '2010/11/26 an add s

'2011/06/09 hn add s
			if RSv("発注数量") > 0 then
				wIroKikakuHacchuuFl = "Y"
			end if  '2010/11/26 an add s
'2011/06/09 hn add e

			if Trim(RSv("色")) <> "" AND Trim(RSv("規格")) <> "" then
				if Trim(RSv("色")) = iro AND Trim(RSv("規格")) = kikaku then
					wIroKikakuCombo = wIroKikakuCombo & "                            <option value='" & Trim(RSv("色")) & "^" & Trim(RSv("規格")) & "' SELECTED>" & Trim(RSv("色")) & "/" & Trim(RSv("規格")) & vbNewLine

					'2013/05/22 GV #1505 add start
					wEAIroKikakuData.Item("code") = Trim(RSv("商品ID"))
					wEAIroKikakuData.Item("iro") = Trim(RSv("色"))
					wEAIroKikakuData.Item("kikaku") = Trim(RSv("規格"))
					'2013/05/22 GV #1505 add end
				else
					wIroKikakuCombo = wIroKikakuCombo & "                            <option value='" & Trim(RSv("色")) & "^" & Trim(RSv("規格")) & "'>" & Trim(RSv("色")) & "/" & Trim(RSv("規格")) & vbNewLine
				end if
				vIroCnt = vIroCnt + 1
				vKikakuCnt = vKikakuCnt + 1
			end if

			if Trim(RSv("色")) <> "" AND Trim(RSv("規格")) = "" then
				if Trim(RSv("色")) = iro AND Trim(RSv("規格")) = kikaku then
					wIroKikakuCombo = wIroKikakuCombo & "                            <option value='" & Trim(RSv("色")) & "^' SELECTED>" & Trim(RSv("色")) & vbNewLine

					'2013/05/22 GV #1505 add start
					wEAIroKikakuData.Item("code") = Trim(RSv("商品ID"))
					wEAIroKikakuData.Item("iro") = Trim(RSv("色"))
					wEAIroKikakuData.Item("kikaku") = Trim(RSv("規格"))
					'2013/05/22 GV #1505 add end
				else
					wIroKikakuCombo = wIroKikakuCombo & "                            <option value='" & Trim(RSv("色")) & "^'>" & Trim(RSv("色")) & vbNewLine
				end if
				vIroCnt = vIroCnt + 1
			end if

			if Trim(RSv("色")) = "" AND Trim(RSv("規格")) <> "" then
				if Trim(RSv("色")) = iro AND Trim(RSv("規格")) = kikaku then
					wIroKikakuCombo = wIroKikakuCombo & "                            <option value='^" & Trim(RSv("規格")) & "' SELECTED>" & Trim(RSv("規格")) & vbNewLine

					'2013/05/22 GV #1505 add start
					wEAIroKikakuData.Item("code") = Trim(RSv("商品ID"))
					wEAIroKikakuData.Item("iro") = Trim(RSv("色"))
					wEAIroKikakuData.Item("kikaku") = Trim(RSv("規格"))
					'2013/05/22 GV #1505 add end
				else
					wIroKikakuCombo = wIroKikakuCombo & "                            <option value='^" & Trim(RSv("規格")) & "'>" & Trim(RSv("規格")) & vbNewLine
				end if
				vKikakuCnt = vKikakuCnt + 1
			end if

			RSv.MoveNext
		Loop
		wIroKikakuCombo = wIroKikakuCombo & "                        </select>" & vbNewLine
		wIroKikakuCombo = wIroKikakuCombo & "                    </p>" & vbNewLine

		if vIroCnt > 0 AND vKikakuCnt > 0 then
			wIroKikakuCombo = Replace(wIroKikakuCombo, "選択", "色規格を選択")
			wIroKikakuSelectMsg = "色規格を選択してください"
		end if
		if vIroCnt > 0 AND vKikakuCnt = 0 then
			wIroKikakuCombo = Replace(wIroKikakuCombo, "選択", "色を選択")
			wIroKikakuSelectMsg = "色を選択してください"
		end if
		if vIroCnt = 0 AND vKikakuCnt > 0 then
			wIroKikakuCombo = Replace(wIroKikakuCombo, "選択", "規格を選択")
			wIroKikakuSelectMsg = "規格を選択してください"
		end if

	else
		wIroKikakuCombo = wIroKikakuCombo & "                    <input type='hidden' name='IroKikaku' value='" & Trim(RSv("色")) & "^" & Trim(RSv("規格")) & "'>" & vbNewLine

		iro = Trim(RSv("色"))
		kikaku = Trim(RSv("規格"))

	end if

	if RSv.RecordCount <= 1 OR iro <> "" OR kikaku <> "" then
		wIroKikakuSelectedFl = true
	else
		wIroKikakuSelectedFl = false
	end if
end if

RSv.close

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
wSQL = wSQL & "     , a.商品画像ファイル名_小"
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
wSQL = wSQL & "     , a.新製品紹介URL"
wSQL = wSQL & "     , a.製品マニュアルURL"
wSQL = wSQL & "     , a.試聴フラグ"
wSQL = wSQL & "     , a.試聴URL"
wSQL = wSQL & "     , a.動画フラグ"
wSQL = wSQL & "     , a.動画URL"
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
wSQL = wSQL & "     , a.シリーズコード"
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
wSQL = wSQL & "     , a.商品備考インサートサイズW2"
wSQL = wSQL & "     , a.商品備考インサートサイズH1"
wSQL = wSQL & "     , a.商品備考インサートサイズH2"
wSQL = wSQL & "     , a.送料完全無料商品フラグ"				' 2011/02/18 GV Add
wSQL = wSQL & "     , a.前回単価変更日"						' 2012/07/13 ok Add
wSQL = wSQL & "     , a.前回販売単価"						' 2012/07/13 ok Add
wSQL = wSQL & "     , b.メーカー名"
wSQL = wSQL & "     , b.メーカー名カナ"
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

if wIroKikakuSelectedFl = true then
	wSQL = wSQL & "     , d.色"
	wSQL = wSQL & "     , d.規格"
	wSQL = wSQL & "     , d.引当可能数量"
	wSQL = wSQL & "     , d.発注数量"								'2011/06/09 hn add
	wSQL = wSQL & "     , d.引当可能入荷予定日"
	wSQL = wSQL & "     , d.B品引当可能数量"
	wSQL = wSQL & "     , d.色規格商品画像ファイル名小"
	wSQL = wSQL & "     , d.色規格商品画像ファイル名1"
	wSQL = wSQL & "     , d.色規格商品画像ファイル名2"
	wSQL = wSQL & "     , d.色規格商品画像ファイル名3"
	wSQL = wSQL & "     , d.色規格商品画像ファイル名4"
	wSQL = wSQL & "     , d.商品ID"
else
	wSQL = wSQL & "     , '' AS 色"
	wSQL = wSQL & "     , '' AS 規格"
	wSQL = wSQL & "     , 0 AS 引当可能数量"
	wSQL = wSQL & "     , 0 AS 発注数量"								'2011/06/09 hn add
	wSQL = wSQL & "     , NULL AS 引当可能入荷予定日"
	wSQL = wSQL & "     , 0 AS B品引当可能数量"
	wSQL = wSQL & "     , '' AS 色規格商品画像ファイル名小"
	wSQL = wSQL & "     , '' AS 色規格商品画像ファイル名1"
	wSQL = wSQL & "     , '' AS 色規格商品画像ファイル名2"
	wSQL = wSQL & "     , '' AS 色規格商品画像ファイル名3"
	wSQL = wSQL & "     , '' AS 色規格商品画像ファイル名4"
	wSQL = wSQL & "     , '' AS 商品ID"
end if

wSQL = wSQL & "     , f.中カテゴリーコード"
wSQL = wSQL & "     , f.中カテゴリー名日本語"
wSQL = wSQL & "     , g.大カテゴリーコード"
wSQL = wSQL & "     , g.大カテゴリー名"
wSQL = wSQL & "     , g.オプションパーツ見出し表記フラグ"	'2012/08/29 ok Add
wSQL = wSQL & "     , h.シリーズ名"				'2012/07/13 ok add
wSQL = wSQL & "     , h.シリーズ画像ファイル名"	'2012/07/13 ok add
wSQL = wSQL & "     , h.シリーズ備考"			'2012/07/13 ok add
wSQL = wSQL & "  FROM Web商品 a WITH (NOLOCK)"
wSQL = wSQL & "     LEFT JOIN シリーズ h WITH (NOLOCK) "		'2012/07/13 ok add
wSQL = wSQL & "     ON a.シリーズコード = h.シリーズコード"		'2012/07/13 ok add
wSQL = wSQL & "     , メーカー b WITH (NOLOCK)"
wSQL = wSQL & "     , カテゴリー c WITH (NOLOCK)"

if wIroKikakuSelectedFl = true then
	wSQL = wSQL & "     , Web色規格別在庫 d WITH (NOLOCK)"
end if

wSQL = wSQL & "     , 中カテゴリー f WITH (NOLOCK) "
wSQL = wSQL & "     , 大カテゴリー g WITH (NOLOCK) "
wSQL = wSQL & " WHERE b.メーカーコード = a.メーカーコード"
wSQL = wSQL & "   AND c.カテゴリーコード = a.カテゴリーコード"
wSQL = wSQL & "   AND f.中カテゴリーコード = c.中カテゴリーコード"
wSQL = wSQL & "   AND g.大カテゴリーコード = f.大カテゴリーコード"
wSQL = wSQL & "   AND a.Web商品フラグ = 'Y'"
wSQL = wSQL & "   AND a.メーカーコード = '" & maker_cd & "'"
wSQL = wSQL & "   AND a.商品コード = '" & Replace(product_cd, "'", "''") & "'"	' 2012/01/23 GV Mod (コード内にシングルクオーテーションが存在した場合の対応)

if wIroKikakuSelectedFl = true then
	wSQL = wSQL & "   AND d.メーカーコード = a.メーカーコード"
	wSQL = wSQL & "   AND d.商品コード = a.商品コード"
	wSQL = wSQL & "   AND d.色 = '" & iro & "'"
	wSQL = wSQL & "   AND d.規格 = '" & kikaku & "'"
	wSQL = wSQL & "   AND d.終了日 IS NULL"
end if

'@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

if RS.EOF = true then
	wPictureHTML = "<center><br><br><div class='honbun'><font color='#ff0000'>　　　　該当商品は有りません。</font></div></center>"
	wMSG = "no data"
	wNoData = "Y"
	exit function
end if

'---- 終了チェック
wProdTermFl = "N"

if isNull(RS("取扱中止日")) = false then		'取扱中止
	wProdTermFl = "Y"
end if

if isNull(RS("廃番日")) = false then  '2010/11/26 an mod s
	if wIroKikakuFl = "Y" then
		if wIroKikakuZaikoFl = "N" AND wIroKikakuHacchuuFl = "N" then		'廃番色規格あり商品で、全色規格在庫なし+発注なし	2011/06/09 hn mod
			wProdTermFl = "Y"
		end if
	else
		if RS("引当可能数量") <= 0 AND RS("発注数量") <= 0 then		'廃番で在庫無し+発注なし	'2011/06/09 hn mod
			wProdTermFl = "Y"
		end if
	end if
end if                                '2010/11/26 an mod e

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
wMakerNameNoKana = RS("メーカー名")  '2010/08/30 an add
if RS("メーカー名カナ") <> "" then
	wMakerName = wMakerName & " ( " & RS("メーカー名カナ") & " ) "
end if
wProductName = RS("商品名")
if trim(RS("色")) <> "" then
	wProductName = wProductName & "/" & RS("色")
end if
if trim(RS("規格")) <> "" then
	wProductName = wProductName & "/" & RS("規格")
end if

'---- パン屑リスト 2010/01/20 an 修正
wTitleWithLink = ""
'2012/07/10 GV Mod Start
'wTitleWithLink = "<h1><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='LargeCategoryList.asp?LargeCategoryCd=" & RS("大カテゴリーコード") & "' class='link' itemprop='url'><span itemprop='title'>" & RS("大カテゴリー名") & "</span>&gt;</a></span>"   '2011/11/22 an mod s
'wTitleWithLink = wTitleWithLink & "<span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='MidCategoryList.asp?MidCategoryCd=" & RS("中カテゴリーコード") & "' class='link' itemprop='url'><span itemprop='title'>" & RS("中カテゴリー名日本語") & "</span>&gt;</a></span>"
'wTitleWithLink = wTitleWithLink & "<span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='SearchList.asp?i_type=c&s_category_cd=" & RS("カテゴリーコード") & "' class='link' itemprop='url'><span itemprop='title'>" & RS("カテゴリー名") &  "</span>&gt;</a></span>"
'wTitleWithLink = wTitleWithLink & "<span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='SearchList.asp?i_type=cm&s_maker_cd=" & RS("メーカーコード") & "&s_category_cd=" & RS("カテゴリーコード") & "' class='link' itemprop='url'><span itemprop='title'>" & wMakerName & "</span></a></span>/" & wProductName & "</h1>"   '2011/11/22 an mod e
wTitleWithLink = "<li><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='LargeCategoryList.asp?LargeCategoryCd=" & RS("大カテゴリーコード") & "' itemprop='url'><span itemprop='title'>" & RS("大カテゴリー名") & "</span></a></span></li>"
wTitleWithLink = wTitleWithLink & "<li><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='MidCategoryList.asp?MidCategoryCd=" & RS("中カテゴリーコード") & "' itemprop='url'><span itemprop='title'>" & RS("中カテゴリー名日本語") & "</span></a></span></li>"
wTitleWithLink = wTitleWithLink & "<li><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='SearchList.asp?i_type=c&s_category_cd=" & RS("カテゴリーコード") & "' itemprop='url'><span itemprop='title'>" & RS("カテゴリー名") &  "</span></a></span></li>"
wTitleWithLink = wTitleWithLink & "<li class='now'><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='SearchList.asp?i_type=cm&s_maker_cd=" & RS("メーカーコード") & "&s_category_cd=" & RS("カテゴリーコード") & "' itemprop='url'><span itemprop='title'>" & wMakerName & "</span></a></span>/" & wProductName & "</li>"
'2012/07/10 GV Mod End

'---- カテゴリーコードセーブ
wCategoryCode = RS("カテゴリーコード")
wLargeCategoryCd = RS("大カテゴリーコード")
wMidCategoryCd = RS("中カテゴリーコード")

'2012/08/29 ok Add
wOptionPartsTitleFlag = RS("オプションパーツ見出し表記フラグ")

'2012/10/30 nt Add
wSeriesCd = RS("シリーズコード")

'---- meta keywords用    2010/08/23 an add s
wLargeCategoryName = RS("大カテゴリー名")
wMidCategoryName = RS("中カテゴリー名日本語")
wCategoryName = RS("カテゴリー名")  '2010/08/23 an add e

'---- 送料完全無料フラグ
wFreeShippingFlag = RS("送料完全無料商品フラグ")			' 2011/02/18 GV Add

'2012/07/13 ok Add Start
'---- シリーズ商品
wSeriesHTML = ""
If RS("シリーズコード") <> "" Then
	wSeriesHTML = wSeriesHTML & "      <div class='detail_side_inner01'><div class='detail_side_inner02'>" & vbNewLine
	wSeriesHTML = wSeriesHTML & "        <div class='detail_side_inner_box'>" & vbNewLine
	wSeriesHTML = wSeriesHTML & "          <h4 class='detail_sub'><a href='SearchList.asp?i_type=se&sSeriesCd=" & RS("シリーズコード") & "'>" & RS("シリーズ名") & "</a></h4>" & vbNewLine
	wSeriesHTML = wSeriesHTML & "            <ul class='check_item'>" & vbNewLine
	wSeriesHTML = wSeriesHTML & "              <li>" & vbNewLine
	If RS("シリーズ画像ファイル名") <> "" Then
		wSeriesHTML = wSeriesHTML & "                <p><a href='SearchList.asp?i_type=se&sSeriesCd=" & RS("シリーズコード") & "'><img src='prod_img/" & RS("シリーズ画像ファイル名") & "' alt='" & Replace(RS("シリーズ名"),"'","&#39;") & "' class='opover'></a>" & RS("シリーズ備考") & "</p>" & vbNewLine
	End If
	wSeriesHTML = wSeriesHTML & "              </li>" & vbNewLine
	wSeriesHTML = wSeriesHTML & "          </ul>" & vbNewLine
	wSeriesHTML = wSeriesHTML & "        </div>" & vbNewLine
	wSeriesHTML = wSeriesHTML & "      </div></div>" & vbNewLine
End If
'2012/07/13 ok Add End

Set wEAProductDetailData = RS	'2013/05/17 GV #1505 add

End Function

'========================================================================
'
'	Function	商品画像 HTML作成
'
'========================================================================
'
Function CreatePictureHTML()

Dim vProdPic(4)

'----- 商品画像
wHTML = ""
'wHTML = wHTML & "<table width='602' border='0' cellspacing='0' cellpadding='0' id='Shop_product_img'>" & vbNewLine	'2012/07/10 GV Del

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

'2012/07/10 GV Mod Start
if vProdPic(1) <> "" then

	wMainProdPic = vProdPic(1)  '2011/11/22 an add

'	wHTML = wHTML & "  <tr align='left' valign='middle'>" & vbNewLine
'	wHTML = wHTML & "    <td><img name='LargeImage' src='prod_img/" & vProdPic(1) & "' alt='" & wMakerName & " / " & wProductname & "' itemprop='image' class='big'></td>" & vbNewLine     '2011/11/22 an mod
'	wHTML = wHTML & "  </tr>" & vbNewLine
	wHTML = wHTML & "  <div id='item_photo_box'>"
	wHTML = wHTML & "    <p><img src='prod_img/" & vProdPic(1) & "' alt='" & Replace(wMakerName & " / " & wProductname,"'","&#39;") & "' id='target' itemprop='image' class='opover'></p>"

end if

'if vProdPic(2) <> "" OR vProdPic(3) <> "" OR vProdPic(4) <> "" then
if vProdPic(1) <> "" OR vProdPic(2) <> "" OR vProdPic(3) <> "" OR vProdPic(4) <> "" then
'	wHTML = wHTML & "  <tr align='center' valign='middle'>" & vbNewLine
'	wHTML = wHTML & "    <td height='75' nowrap>" & vbNewLine
	wHTML = wHTML & "    <ul class='sub_box'>"
	if vProdPic(1) <> "" then
'		wHTML = wHTML & "    <img src='prod_img/" & vProdPic(1) & "' width='147' height='73' class='small1' alt='" & wMakerName & " / " & wProductname & " 画像1' onMouseOver='SmallImage_onMouseOver(""prod_img/" & vProdPic(1) & """);'>"
		wHTML = wHTML & "      <li><a class='modalImg' rel='fancybox' href='prod_img/" & vProdPic(1) & "'><img src='prod_img/" & vProdPic(1) & "' alt='" & Replace(wMakerName & " / " & wProductname,"'","&#39;") & " 画像1' class='opover'></a></li>"
	end if
	if vProdPic(2) <> "" then
'		wHTML = wHTML & "<img src='prod_img/" & vProdPic(2) & "' width='147' height='73' class='small1' alt='" & wMakerName & " / " & wProductname & " 画像2' onMouseOver='SmallImage_onMouseOver(""prod_img/" & vProdPic(2) & """);'>"
		wHTML = wHTML & "      <li><a class='modalImg' rel='fancybox' href='prod_img/" & vProdPic(2) & "'><img src='prod_img/" & vProdPic(2) & "' alt='" & Replace(wMakerName & " / " & wProductname,"'","&#39;") & " 画像2' class='opover'></a></li>"
	end if
	if vProdPic(3) <> "" then
'		wHTML = wHTML & "<img src='prod_img/" & vProdPic(3) & "' width='147' height='73' class='small1' alt='" & wMakerName & " / " & wProductname & " 画像3' onMouseOver='SmallImage_onMouseOver(""prod_img/" & vProdPic(3) & """);'>"
		wHTML = wHTML & "      <li><a class='modalImg' rel='fancybox' href='prod_img/" & vProdPic(3) & "'><img src='prod_img/" & vProdPic(3) & "' alt='" & Replace(wMakerName & " / " & wProductname,"'","&#39;") & " 画像3' class='opover'></a></li>"
	end if
	if vProdPic(4) <> "" then
'		wHTML = wHTML & "<img src='prod_img/" & vProdPic(4) & "' width='147' height='73' class='small1' alt='" & wMakerName & " / " & wProductname & " 画像4' onMouseOver='SmallImage_onMouseOver(""prod_img/" & vProdPic(4) & """);'>" & vbNewLine
		wHTML = wHTML & "      <li><a class='modalImg' rel='fancybox' href='prod_img/" & vProdPic(4) & "'><img src='prod_img/" & vProdPic(4) & "' alt='" & Replace(wMakerName & " / " & wProductname,"'","&#39;") & " 画像4' class='opover'></a></li>"
	end if
'	wHTML = wHTML & "    </td>" & vbNewLine
'	wHTML = wHTML & "  </tr>" & vbNewLine
	wHTML = wHTML & "    </ul>"
	wHTML = wHTML & "  </div>"
end if
'2012/07/10 GV Mod End

'wHTML = wHTML & "</table>" & vbNewLine	'2012/07/10 GV Del

wPictureHTML = wHTML


End Function

'========================================================================
'
'	Function	関連リンク HTML作成
'
'========================================================================
'
Function CreateKanrenLinkHTML()

Dim vURL()
Dim vURLCount
Dim vURL2
Dim wHTMLTemp
Dim i

wHTML = ""
wHTMLTemp = ""

'----動画リンク
if RS("動画フラグ") = "Y" then
	vURLCount = cf_unstring(RS("動画URL"), vURL, ",")

	if vURLCount > 1 then
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "      <a href='JavaScript:void(0);' onMouseOut=""MM_swapImgRestore()""  onMouseOver=""MM_swapImage('movie" & i & "','','images/movie_on.gif',1)"" onClick=""window.open('SoundMoviePopUp.asp?item=" & RS("メーカーコード") & "^" & Server.URLEncode(RS("商品コード")) & "','SoundMovie', 'width=201 height=200 resizable=1 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');"" class='link'><img src='images/movie_off.gif' border='0' name='movie" & i & "' alt='動画を見る'></a>" & vbNewLine
		wHTML = wHTML & "      <li><a href='JavaScript:void(0);' onClick=""window.open('SoundMoviePopUp.asp?item=" & RS("メーカーコード") & "^" & Server.URLEncode(RS("商品コード")) & "','SoundMovie', 'width=201 height=200 resizable=1 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');""><img src='images/btn_movie.png' alt='動画を見る' class='opover'></a></li>" & vbNewLine
		'2012/07/10 GV Mod End
	else
		if InStr(vURL(0), "http://") > 0 then
			vURL2 = vURL(0)
		else
			vURL2 = g_HTTP & vURL(0)
		end if
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "      <a href='" & vURL2 & "' target='_blank' onMouseOut=""MM_swapImgRestore()""  onMouseOver=""MM_swapImage('movie" & i & "','','images/movie_on.gif',1)""><img src='images/movie_off.gif' border='0' name='movie" & i & "' alt='動画を見る'></a>" & vbNewLine
		wHTML = wHTML & "      <li><a href='" & vURL2 & "' target='_blank'><img src='images/btn_movie.png' alt='動画を見る' class='opover'></a></li>" & vbNewLine
		'2012/07/10 GV Mod End
	end if

end if

'----試聴リンク
if RS("試聴フラグ") = "Y" then
	vURLCount = cf_unstring(RS("試聴URL"), vURL, ",")

	if vURLCount > 1 then
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "      <a href='JavaScript:void(0);' onMouseOut=""MM_swapImgRestore()""  onMouseOver=""MM_swapImage('audio" & i & "','','images/audio_on.gif',1)"" onClick=""window.open('SoundMoviePopUp.asp?item=" & RS("メーカーコード") & "^" & Server.URLEncode(RS("商品コード")) & "','SoundMovie', 'width=201 height=200 resizable=1 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');"" class='link'><img src='images/audio_off.gif' border='0' name='audio" & i & "' alt='試聴する'></a>" & vbNewLine
		wHTML = wHTML & "      <li><a href='JavaScript:void(0);' onClick=""window.open('SoundMoviePopUp.asp?item=" & RS("メーカーコード") & "^" & Server.URLEncode(RS("商品コード")) & "','SoundMovie', 'width=201 height=200 resizable=1 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');""><img src='images/btn_view.png' alt='試聴する' class='opover'></a></li>" & vbNewLine
		'2012/07/10 GV Mod End
	else
		if InStr(vURL(0), "http://") > 0 then
			vURL2 = vURL(0)
		else
			vURL2 = g_HTTP & vURL(0)
		end if
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "      <a href='" & vURL2 & "' target='_blank' onMouseOut=""MM_swapImgRestore()""  onMouseOver=""MM_swapImage('audio" & i & "','','images/audio_on.gif',1)""><img src='images/audio_off.gif' border='0' name='audio" & i & "' alt='試聴する'></a>" & vbNewLine
		wHTML = wHTML & "      <li><a href='" & vURL2 & "' target='_blank'><img src='images/btn_view.png' alt='試聴する' class='opover'></a></li>" & vbNewLine
		'2012/07/10 GV Mod End

	end if

end if

'---- メーカーホームページ
if RS("メーカーホームページURL") <> "" then
	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "      <a href='" & RS("メーカーホームページURL") & "'target='_blank' onMouseOut=""MM_swapImgRestore()""  onMouseOver=""MM_swapImage('maker','','images/maker_on.gif',1)""><img src='images/maker_off.gif' border='0' name='maker' alt='メーカーサイトへ'></a>" & vbNewLine
	wHTML = wHTML & "      <li><a href='" & RS("メーカーホームページURL") & "' target='_blank'><img src='images/btn_maker.png' alt='メーカーサイト' class='opover'></a></li>" & vbNewLine
	'2012/07/10 GV Mod End
end if

'---- 製品マニュアル
if RS("製品マニュアルURL") <> "" then
	if InStr(LCase(RS("製品マニュアルURL")), "http://") > 0 then
		vURL2 = Trim(RS("製品マニュアルURL"))
	else
		vURL2 = g_HTTP & Trim(RS("製品マニュアルURL"))
	end if

	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "      <a href='" & vURL2 &  "'target='_blank' onMouseOut=""MM_swapImgRestore()""  onMouseOver=""MM_swapImage('manual" & i & "','','images/manual_on.gif',1)""><img src='images/manual_off.gif' border='0' name='manual" & i & "' alt='製品マニュアル'></a>" & vbNewLine
	wHTML = wHTML & "      <li><a href='" & vURL2 & "' target='_blank'><img src='images/btn_manual.png' alt='製品マニュアル' class='opover'></a></li>" & vbNewLine
	'2012/07/10 GV Mod End
End if

'----
if wHTML <> "" then
'2012/07/10 GV Mod Start
'	wHTMLTemp = wHTMLTemp & "<table width='602' border='0' cellSpacing='0' cellPadding='0'>" & vbNewLine
'	wHTMLTemp = wHTMLTemp & "  <tr>" & vbNewLine
'	wHTMLTemp = wHTMLTemp & "    <td  height='40'>" & vbNewLine
'	wHTMLTemp = wHTMLTemp & wHTML
'	wHTMLTemp = wHTMLTemp & "    </td>" & vbNewLine
'	wHTMLTemp = wHTMLTemp & "  </tr>" & vbNewLine
	wHTMLTemp = wHTMLTemp & "  <div class='btn_box'>" & vbNewLine
	wHTMLTemp = wHTMLTemp & "    <ul class='btn'>" & vbNewLine
	wHTMLTemp = wHTMLTemp & wHTML
	wHTMLTemp = wHTMLTemp & "    </ul>" & vbNewLine
	wHTMLTemp = wHTMLTemp & "  </div>" & vbNewLine
'2012/07/10 GV Mod End
end if

'---- メーカー関連記事URL1-4
wHTML = ""

if RS("関連記事URL1") <> "" then

	if InStr(LCase(RS("関連記事URL1")), "http://") > 0 then
		vURL2 = Trim(RS("関連記事URL1"))
	else
		vURL2 = g_HTTP & Trim(RS("関連記事URL1"))
	end if

	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "      <a href='" & vURL2 & "' target='_blank'>" & RS("関連記事タイトル1") & "</a>&nbsp;&nbsp;|" & vbNewLine
	wHTML = wHTML & "      <li><a href='" & vURL2 & "' target='_blank'>" & RS("関連記事タイトル1") & "</a></li>" & vbNewLine
	'2012/07/10 GV Mod End
end if

if RS("関連記事URL2") <> "" then

	if InStr(LCase(RS("関連記事URL2")), "http://") > 0 then
		vURL2 = Trim(RS("関連記事URL2"))
	else
		vURL2 = g_HTTP & Trim(RS("関連記事URL2"))
	end if

	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "      <a href='" & vURL2 & "' target='_blank'>" & RS("関連記事タイトル2") & "</a>&nbsp;&nbsp;|" & vbNewLine
	wHTML = wHTML & "      <li><a href='" & vURL2 & "' target='_blank'>" & RS("関連記事タイトル2") & "</a></li>" & vbNewLine
	'2012/07/10 GV Mod End
end if

if RS("関連記事URL3") <> "" then

	if InStr(LCase(RS("関連記事URL3")), "http://") > 0 then
		vURL2 = Trim(RS("関連記事URL3"))
	else
		vURL2 = g_HTTP & Trim(RS("関連記事URL3"))
	end if

	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "      <a href='" & vURL2 & "' target='_blank'>" & RS("関連記事タイトル3") & "</a>&nbsp;&nbsp;|" & vbNewLine
	wHTML = wHTML & "      <li><a href='" & vURL2 & "' target='_blank'>" & RS("関連記事タイトル3") & "</a></li>" & vbNewLine
	'2012/07/10 GV Mod End
end if


if RS("関連記事URL4") <> "" then

	if InStr(LCase(RS("関連記事URL4")), "http://") > 0 then
		vURL2 = Trim(RS("関連記事URL4"))
	else
		vURL2 = g_HTTP & Trim(RS("関連記事URL4"))
	end if

	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "      <a href='" & vURL2 & "' target='_blank'>" & RS("関連記事タイトル4") & "</a>&nbsp;&nbsp;|" & vbNewLine
	wHTML = wHTML & "      <li><a href='" & vURL2 & "' target='_blank'>" & RS("関連記事タイトル4") & "</a></li>" & vbNewLine
	'2012/07/10 GV Mod End
end if

'----
if wHTML <> "" then
	'2012/07/10 GV Del Start
'	wHTML = Left(wHTML, Len(wHTML)-3)
'	if wHTMLTemp = "" then
'		wHTMLTemp = wHTMLTemp & "<table border=0 cellSpacing=0 cellPadding=0 width=602>" & vbNewLine
'	end if
	'2012/07/10 GV Del End

	'2012/07/10 GV Mod Start
'	wHTMLTemp = wHTMLTemp & "  <tr>" & vbNewLine
'	wHTMLTemp = wHTMLTemp & "    <td  height='40'>" & vbNewLine
'	wHTMLTemp = wHTMLTemp & wHTML & vbNewLine
'	wHTMLTemp = wHTMLTemp & "    </td>" & vbNewLine
'	wHTMLTemp = wHTMLTemp & "  </tr>" & vbNewLine
	wHTMLTemp = wHTMLTemp & "  <div class='other_link'>" & vbNewLine
	wHTMLTemp = wHTMLTemp & "    <ul class='link'>" & vbNewLine
	wHTMLTemp = wHTMLTemp & wHTML & vbNewLine
	wHTMLTemp = wHTMLTemp & "    </ul>" & vbNewLine
	wHTMLTemp = wHTMLTemp & "  </div>" & vbNewLine
	'2012/07/10 GV Mod End
end if

'2012/07/10 GV Del Start
'if wHTMLTemp <> "" then
'	wHTMLTemp = wHTMLTemp & "</table>" & vbNewLine
'end if
'2012/07/10 GV Del End

wKanrenLinkHTML = wHTMLTemp

End function


'========================================================================
'
'	Function	特徴 HTML作成
'
'========================================================================
'
Function CreateTokuchoHTML()

wHTML = ""

'---- 特徴, 直輸入品
if RS("お勧め商品コメント") <> "" OR RS("直輸入品フラグ") = "Y" then
'2012/07/10 GV Mod Start
'	wHTML = wHTML & "<table width='602' border='0' cellpadding='0' cellspacing='0' id='main_header'>" & vbNewLine
'	wHTML = wHTML & "  <tr>" & vbNewLine
'	wHTML = wHTML & "    <td><strong>特徴&nbsp;&nbsp;[" & RS("メーカー名") & "(" & RS("メーカー名カナ") & ")/" & RS("商品名") & "]</strong></td>" & vbNewLine
'	wHTML = wHTML & "  </tr>" & vbNewLine
'	wHTML = wHTML & "</table>" & vbNewLine
'	wHTML = wHTML & "<table width='602' border='0' cellpadding='0' cellspacing='0' id='shop_border'>" & vbNewLine
'	wHTML = wHTML & "  <tr>" & vbNewLine
'	wHTML = wHTML & "    <td>" & vbNewLine
'	wHTML = wHTML & "      <p>"
	wHTML = wHTML & "<div class='inner_box_spec'>" & vbNewLine

	if RS("お勧め商品コメント") <> "" then
		wHTML = wHTML & "<span itemprop='description'>"    '2011/11/22 an add
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & RS("お勧め商品コメント") & "<br>" & vbNewLine
		wHTML = wHTML & RS("お勧め商品コメント") & vbNewLine
		'2012/07/10 GV Mod End
		wHTML = wHTML & "</span>"                          '2011/11/22 an add

		'---- meta description用データ取得          '2010/08/23 an add s
		wTokucho = fDeleteHTMLTag(RS("お勧め商品コメント")) 'HTMLタグ削除
		wTokucho = replace(replace(replace(replace(wTokucho, vbCr, ""), vbLf, ""), vbTab, ""), """", "") '改行、Tabの削除

		if Len(wTokucho) > 97 then  '長い場合は100文字に省略
			wTokucho = Left(wTokucho,97) & "..."
		end if                                      '2010/08/23 an add e

	end if

	if RS("直輸入品フラグ") = "Y" then
'		wHTML = wHTML & "<a href='../information/direct_import.asp' class='link'>[直輸入品]</a>" & vbNewLine
		wHTML = wHTML & "  <p><a href='../information/direct_import.asp'>[直輸入品]</a></p>" & vbNewLine
	end if

'	wHTML = wHTML & "      </p>" & vbNewLine
'	wHTML = wHTML & "    </td>" & vbNewLine
'	wHTML = wHTML & "  </tr>" & vbNewLine
'	wHTML = wHTML & "</table>" & vbNewLine
	wHTML = wHTML & "</div>" & vbNewLine
'2012/07/10 GV Mod End
end if

wTokuchoHTML = wHTML

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
Dim vRemark          '2011/11/22 an add
Dim vFileExtention   '2011/11/22 an add
Dim vURL1HTML        '2011/11/22 an add
Dim vURL2HTML        '2011/11/22 an add

wHTML = ""
vURL1HTML = ""       '2011/11/22 an add
vURL2HTML = ""       '2011/11/22 an add

'---- スペック
'2012/07/10 GV Del Start
'wHTML = wHTML & "<table width='602' border='0' cellpadding='0' cellspacing='0' id='main_header'>" & vbNewLine
'wHTML = wHTML & "  <tr>" & vbNewLine
'wHTML = wHTML & "    <td><h2>スペック&nbsp;&nbsp;[" & RS("メーカー名") & "(" & RS("メーカー名カナ") & ")/" & RS("商品名") & "]</h2></td>" & vbNewLine
'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & "</table>" & vbNewLine
'2012/07/10 GV Del End

if RS("商品備考インサートURL1") <> "" then

	if InStr(LCase(RS("商品備考インサートURL1")), "http") > 0 then   '2011/11/22 an mod s
	else
		'---- txtファイルの場合
vFileExtention = LCase(Right(RS("商品備考インサートURL1"), 3))
		if LCase(Right(RS("商品備考インサートURL1"), 3)) = "txt" then

			'---- ファイルの存在確認
			vRemark = GetMapPath(RS("商品備考インサートURL1"), vFileExtention)

			if vRemark <> "" then
'				vURL1HTML = vURL1HTML & "<div class='insert'>" & vbNewLine	'2012/07/10 GV Del
				vURL1HTML = vURL1HTML & cf_read_file_all(vRemark) & vbNewLine
'				vURL1HTML = vURL1HTML & "</div>" & vbNewLine			'2012/07/10 GV Del
			end if
		'---- txtファイル以外の場合
'2012/07/19 ok Del Start txt以外の場合、幅調整が必要なため非表示とする
'		else
'
'			if RS("商品備考インサートサイズW1") <> 0 then
'				vWidth = RS("商品備考インサートサイズW1")
'				if vWidth > 600 then
'					vWidth = 600
'				end if
'			else
'				vWidth = 600
'			end if
'
'			if RS("商品備考インサートサイズH1") <> 0 then
'				vHeight = RS("商品備考インサートサイズH1")
'			else
'				vHeight = 290
'			end if
'
'			vURL1HTML = vURL1HTML & "<iframe class='insert' marginwidth='0' marginheight='0' scrolling='no' src='" & RS("商品備考インサートURL1") & "' width='" & vWidth & "' height='" & vHeight & "' frameborder='0'></iframe>"
'2012/07/19 ok Del End
		end if

		if vURL1HTML <> "" then
			if InStr(LCase(RS("商品備考インサートURL1")), "http") > 0 then
			else
'2012/07/10 GV Mod Start
'				wHTML = wHTML & "<table width='602' border='0' cellpadding='0' cellspacing='0' id='shop_border_insert'>" & vbNewLine
'				wHTML = wHTML & "  <tr>" & vbNewLine
'				wHTML = wHTML & "    <td>" & vURL1HTML & "</td>" & vbNewLine
'				wHTML = wHTML & "  </tr>" & vbNewLine
'				wHTML = wHTML & "</table>" & vbNewLine
				wHTML = wHTML & "<div class='insert_box'>" & vbNewLine
				wHTML = wHTML & vURL1HTML & vbNewLine
				wHTML = wHTML & "</div>" & vbNewLine
'2012/07/10 GV Mod End
			end if
		end if

	end if   '2011/11/22 an mod e
end if

'2012/07/10 GV Mod Start
'wHTML = wHTML & "<table width='602' border='0' cellpadding='0' cellspacing='0' id='shop_border'>" & vbNewLine
'wHTML = wHTML & "  <tr>" & vbNewLine
'wHTML = wHTML & "    <td><p>" & CreateSpecHTML(RS("カテゴリーコード"),RS("メーカーコード"),RS("商品コード"),RS("商品備考"),RS("商品スペック使用不可フラグ")) & "</p></td>" & vbNewLine
'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & "</table>" & vbNewLine
wHTML = wHTML & "<div class='inner_box_spec'>" & vbNewLine
wHTML = wHTML & CreateSpecHTML(RS("カテゴリーコード"),RS("メーカーコード"),RS("商品コード"),RS("商品備考"),RS("商品スペック使用不可フラグ")) & vbNewLine
wHTML = wHTML & "</div>" & vbNewLine
'2012/07/10 GV Mod End

if RS("商品備考インサートURL2") <> "" then

	if InStr(LCase(RS("商品備考インサートURL2")), "http") > 0 then   '2011/11/22 an mod s
	else
		'---- txtファイルの場合
		if LCase(Right(RS("商品備考インサートURL2"), 3)) = "txt" then

			'---- ファイルの存在確認
			vRemark = GetMapPath(RS("商品備考インサートURL2"), vFileExtention)

			if vRemark <> "" then
'				vURL2HTML = vURL2HTML & "<div class='insert'>" & vbNewLine	'2012/07/10 GV Del
				vURL2HTML = vURL2HTML & cf_read_file_all(vRemark) & vbNewLine
'				vURL2HTML = vURL2HTML & "</div>" & vbNewLine			'2012/07/10 GV Del
			end if
'2012/07/19 ok Del Start txt以外の場合、幅調整が必要なため非表示とする
'		else
'
'			if RS("商品備考インサートサイズW2") <> 0 then
'				vWidth = RS("商品備考インサートサイズW2")
'			else
'				vWidth = 600
'			end if
'
'			if RS("商品備考インサートサイズH2") <> 0 then
'				vHeight = RS("商品備考インサートサイズH2")
'			else
'				vHeight = 300
'			end if
'
'			vURL2HTML = vURL2HTML & "<iframe class='insert' marginwidth='0' marginheight='0' scrolling='no' src='" & RS("商品備考インサートURL2") & "' width='" & vWidth & "' height='" & vHeight & "' frameborder='0'></iframe>"
'2012/07/19 ok Del End
		end if

		if vURL2HTML <> "" then
			if InStr(LCase(RS("商品備考インサートURL2")), "http") > 0 then
			else
'2012/07/10 GV Mod Start
'				wHTML = wHTML & "<table width='602' border='0' cellpadding='0' cellspacing='0' id='shop_border_insert'>" & vbNewLine
'				wHTML = wHTML & "  <tr>" & vbNewLine
'				wHTML = wHTML & "    <td>" & vURL2HTML & "</td>" & vbNewLine
'				wHTML = wHTML & "  </tr>" & vbNewLine
'				wHTML = wHTML & "</table>" & vbNewLine
				wHTML = wHTML & "<div class='insert_box'>" & vbNewLine
				wHTML = wHTML & vURL2HTML & vbNewLine
				wHTML = wHTML & "</div>" & vbNewLine
'2012/07/10 GV Mod End
			end if
		end if

	end if   '2011/11/22 an mod e
end if

wSpecHTML = wHTML

End Function

'========================================================================
'
'	Function	オプション HTML（データ抽出）
'
'		色規格に関係なく該当商品のオプションを取り出す
'
'========================================================================
'
Function CreateOptionHTML()

'---- Select オプション商品
wSQL = ""
' 2012/01/18 GV Mod Start
'wSQL = wSQL & "SELECT c.オプションメーカーコード AS メーカーコード"
'wSQL = wSQL & "     , c.オプション商品コード AS 商品コード"
'wSQL = wSQL & "     , d.色 AS 色"
'wSQL = wSQL & "     , d.規格 AS 規格"
'wSQL = wSQL & "     , a.商品名"
''wSQL = wSQL & "     , a.販売単価"     '2010/11/10 an del
'
''2010/11/10 an add s
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN a.B品フラグ = 'Y' THEN a.B品単価"
'wSQL = wSQL & "         WHEN a.個数限定数量 > a.個数限定受注済数量 THEN a.個数限定単価"
'wSQL = wSQL & "         ELSE a.販売単価"
'wSQL = wSQL & "       END AS 実販売単価"
''2010/11/10 an add e
'
'wSQL = wSQL & "     , a.商品概略Web"
'wSQL = wSQL & "     , a.商品画像ファイル名_小"
'wSQL = wSQL & "     , a.ASK商品フラグ"
'wSQL = wSQL & "     , a.取扱中止日"
'wSQL = wSQL & "     , a.廃番日"
'wSQL = wSQL & "     , a.完売日"
'wSQL = wSQL & "     , a.希少数量"
'wSQL = wSQL & "     , a.セット商品フラグ"
'wSQL = wSQL & "     , a.メーカー直送取寄区分"
'wSQL = wSQL & "     , a.Web納期非表示フラグ"
'wSQL = wSQL & "     , a.入荷予定未定フラグ"
'wSQL = wSQL & "     , a.B品フラグ"
'wSQL = wSQL & "     , a.個数限定数量"
'wSQL = wSQL & "     , a.個数限定受注済数量"
'wSQL = wSQL & "     , b.メーカー名"
'wSQL = wSQL & "     , d.引当可能数量"
'wSQL = wSQL & "     , d.発注数量"			'2011/06/09 hn add
'wSQL = wSQL & "     , d.引当可能入荷予定日"
'wSQL = wSQL & "     , d.B品引当可能数量"
'wSQL = wSQL & "  FROM Web商品 a"
'wSQL = wSQL & "     , メーカー b"
'wSQL = wSQL & "     , オプション2 c"
'wSQL = wSQL & "     , Web色規格別在庫 d"
'wSQL = wSQL & " WHERE a.メーカーコード = c.オプションメーカーコード"
'wSQL = wSQL & "   AND a.商品コード = c.オプション商品コード"
'wSQL = wSQL & "   AND b.メーカーコード = c.オプションメーカーコード"
'wSQL = wSQL & "   AND d.メーカーコード = c.オプションメーカーコード"
'wSQL = wSQL & "   AND d.商品コード = c.オプション商品コード"
'wSQL = wSQL & "   AND d.色 = c.オプション色"
'wSQL = wSQL & "   AND d.規格 = c.オプション規格"
'wSQL = wSQL & "   AND c.メーカーコード = '" & maker_cd & "'"
'wSQL = wSQL & "   AND c.商品コード = '" & product_cd & "'"
'wSQL = wSQL & "   AND a.Web商品フラグ = 'Y'"
'
'wSQL = wSQL & " UNION "
'
'wSQL = wSQL & "SELECT c.オプションメーカーコード AS メーカーコード"
'wSQL = wSQL & "     , c.オプション商品コード AS 商品コード"
'wSQL = wSQL & "     , d.色 AS 色"
'wSQL = wSQL & "     , d.規格 AS 規格"
'wSQL = wSQL & "     , a.商品名"
''wSQL = wSQL & "     , a.販売単価"     '2010/11/10 an del
'
''2010/11/10 an add s
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN a.B品フラグ = 'Y' THEN a.B品単価"
'wSQL = wSQL & "         WHEN a.個数限定数量 > a.個数限定受注済数量 THEN a.個数限定単価"
'wSQL = wSQL & "         ELSE a.販売単価"
'wSQL = wSQL & "       END AS 実販売単価"
''2010/11/10 an add e
'
'wSQL = wSQL & "     , a.商品概略Web"
'wSQL = wSQL & "     , a.商品画像ファイル名_小"
'wSQL = wSQL & "     , a.ASK商品フラグ"
'wSQL = wSQL & "     , a.取扱中止日"
'wSQL = wSQL & "     , a.廃番日"
'wSQL = wSQL & "     , a.完売日"
'wSQL = wSQL & "     , a.希少数量"
'wSQL = wSQL & "     , a.セット商品フラグ"
'wSQL = wSQL & "     , a.メーカー直送取寄区分"
'wSQL = wSQL & "     , a.Web納期非表示フラグ"
'wSQL = wSQL & "     , a.入荷予定未定フラグ"
'wSQL = wSQL & "     , a.B品フラグ"
'wSQL = wSQL & "     , a.個数限定数量"
'wSQL = wSQL & "     , a.個数限定受注済数量"
'wSQL = wSQL & "     , b.メーカー名"
'wSQL = wSQL & "     , d.引当可能数量"
'wSQL = wSQL & "     , d.発注数量"			'2011/06/09 hn add
'wSQL = wSQL & "     , d.引当可能入荷予定日"
'wSQL = wSQL & "     , d.B品引当可能数量"
'wSQL = wSQL & "  FROM Web商品 a"
'wSQL = wSQL & "     , メーカー b"
'wSQL = wSQL & "     , カテゴリー別オプション c"
'wSQL = wSQL & "     , Web色規格別在庫 d"
'wSQL = wSQL & " WHERE a.メーカーコード = c.オプションメーカーコード"
'wSQL = wSQL & "   AND a.商品コード = c.オプション商品コード"
'wSQL = wSQL & "   AND b.メーカーコード = c.オプションメーカーコード"
'wSQL = wSQL & "   AND d.メーカーコード = c.オプションメーカーコード"
'wSQL = wSQL & "   AND d.商品コード = c.オプション商品コード"
'wSQL = wSQL & "   AND d.色 = c.オプション色"
'wSQL = wSQL & "   AND d.規格 = c.オプション規格"
'wSQL = wSQL & "   AND c.カテゴリーコード = '" & wCategoryCode & "'"
'wSQL = wSQL & "   AND a.Web商品フラグ = 'Y'"
'
'wSQL = wSQL & " ORDER BY"
'wSQL = wSQL & "       b.メーカー名"
'wSQL = wSQL & "     , a.商品名"
'wSQL = wSQL & "     , d.色"
'wSQL = wSQL & "     , d.規格"
wSQL = wSQL & "SELECT "
wSQL = wSQL & "      c.オプションメーカーコード AS メーカーコード "
wSQL = wSQL & "    , c.オプション商品コード AS 商品コード "
wSQL = wSQL & "    , d.色 AS 色 "
wSQL = wSQL & "    , d.規格 AS 規格 "
wSQL = wSQL & "    , a.商品名 "
wSQL = wSQL & "    , CASE "
wSQL = wSQL & "        WHEN a.B品フラグ = 'Y'                     THEN a.B品単価 "
wSQL = wSQL & "        WHEN a.個数限定数量 > a.個数限定受注済数量 THEN a.個数限定単価 "
wSQL = wSQL & "        ELSE                                            a.販売単価 "
wSQL = wSQL & "      END AS 実販売単価 "
wSQL = wSQL & "    , a.商品概略Web "
wSQL = wSQL & "    , a.商品画像ファイル名_小 "
wSQL = wSQL & "    , a.ASK商品フラグ "
wSQL = wSQL & "    , a.取扱中止日 "
wSQL = wSQL & "    , a.廃番日 "
wSQL = wSQL & "    , a.完売日 "
wSQL = wSQL & "    , a.希少数量 "
wSQL = wSQL & "    , a.セット商品フラグ "
wSQL = wSQL & "    , a.メーカー直送取寄区分 "
wSQL = wSQL & "    , a.Web納期非表示フラグ "
wSQL = wSQL & "    , a.入荷予定未定フラグ "
wSQL = wSQL & "    , a.B品フラグ "
wSQL = wSQL & "    , a.個数限定数量 "
wSQL = wSQL & "    , a.個数限定受注済数量 "
wSQL = wSQL & "    , b.メーカー名 "
wSQL = wSQL & "    , d.引当可能数量 "
wSQL = wSQL & "    , d.発注数量 "
wSQL = wSQL & "    , d.引当可能入荷予定日 "
wSQL = wSQL & "    , d.B品引当可能数量 "
wSQL = wSQL & "    , a.カテゴリーコード "		'2012/08/27 ok Add
wSQL = wSQL & "    , e.カテゴリー名 "			'2012/08/27 ok Add
wSQL = wSQL & "FROM "
wSQL = wSQL & "    オプション2                  c WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN Web商品         a WITH (NOLOCK) "
wSQL = wSQL & "        ON     a.メーカーコード = c.オプションメーカーコード "
wSQL = wSQL & "           AND a.商品コード     = c.オプション商品コード "
wSQL = wSQL & "      INNER JOIN メーカー        b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.メーカーコード = c.オプションメーカーコード "
wSQL = wSQL & "      INNER JOIN Web色規格別在庫 d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.メーカーコード = c.オプションメーカーコード "
wSQL = wSQL & "           AND d.商品コード     = c.オプション商品コード "
wSQL = wSQL & "           AND d.色             = c.オプション色 "
wSQL = wSQL & "           AND d.規格           = c.オプション規格 "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' )   t1 "
wSQL = wSQL & "        ON     a.Web商品フラグ    = t1.ShohinWebY "
wSQL = wSQL & "      INNER JOIN カテゴリー e WITH (NOLOCK) "			'2012/08/27 ok Add
wSQL = wSQL & "        ON     a.カテゴリーコード = e.カテゴリーコード "	'2012/08/27 ok Add
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        t1.ShohinWebY IS NOT NULL "
wSQL = wSQL & "    AND c.メーカーコード = '" & maker_cd & "' "
wSQL = wSQL & "    AND c.商品コード     = '" & Replace(product_cd, "'", "''") & "' "	' 2012/01/23 GV Mod (コード内にシングルクオーテーションが存在した場合の対応)

wSQL = wSQL & "UNION "

wSQL = wSQL & "SELECT "
wSQL = wSQL & "      c.オプションメーカーコード AS メーカーコード "
wSQL = wSQL & "    , c.オプション商品コード AS 商品コード "
wSQL = wSQL & "    , d.色 AS 色 "
wSQL = wSQL & "    , d.規格 AS 規格 "
wSQL = wSQL & "    , a.商品名 "
wSQL = wSQL & "    , CASE "
wSQL = wSQL & "        WHEN a.B品フラグ = 'Y'                     THEN a.B品単価 "
wSQL = wSQL & "        WHEN a.個数限定数量 > a.個数限定受注済数量 THEN a.個数限定単価 "
wSQL = wSQL & "        ELSE                                            a.販売単価 "
wSQL = wSQL & "      END AS 実販売単価 "
wSQL = wSQL & "    , a.商品概略Web "
wSQL = wSQL & "    , a.商品画像ファイル名_小 "
wSQL = wSQL & "    , a.ASK商品フラグ "
wSQL = wSQL & "    , a.取扱中止日 "
wSQL = wSQL & "    , a.廃番日 "
wSQL = wSQL & "    , a.完売日 "
wSQL = wSQL & "    , a.希少数量 "
wSQL = wSQL & "    , a.セット商品フラグ "
wSQL = wSQL & "    , a.メーカー直送取寄区分 "
wSQL = wSQL & "    , a.Web納期非表示フラグ "
wSQL = wSQL & "    , a.入荷予定未定フラグ "
wSQL = wSQL & "    , a.B品フラグ "
wSQL = wSQL & "    , a.個数限定数量 "
wSQL = wSQL & "    , a.個数限定受注済数量 "
wSQL = wSQL & "    , b.メーカー名 "
wSQL = wSQL & "    , d.引当可能数量 "
wSQL = wSQL & "    , d.発注数量 "
wSQL = wSQL & "    , d.引当可能入荷予定日 "
wSQL = wSQL & "    , d.B品引当可能数量 "
wSQL = wSQL & "    , a.カテゴリーコード "			'2012/08/27 ok Add
wSQL = wSQL & "    , e.カテゴリー名 "				'2012/08/27 ok Add
wSQL = wSQL & "FROM "
wSQL = wSQL & "    カテゴリー別オプション       c WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN Web商品         a WITH (NOLOCK) "
wSQL = wSQL & "        ON     a.メーカーコード = c.オプションメーカーコード "
wSQL = wSQL & "           AND a.商品コード     = c.オプション商品コード "
wSQL = wSQL & "      INNER JOIN メーカー        b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.メーカーコード = c.オプションメーカーコード "
wSQL = wSQL & "      INNER JOIN Web色規格別在庫 d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.メーカーコード = c.オプションメーカーコード "
wSQL = wSQL & "           AND d.商品コード     = c.オプション商品コード "
wSQL = wSQL & "           AND d.色             = c.オプション色 "
wSQL = wSQL & "           AND d.規格           = c.オプション規格 "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' )   t1 "
wSQL = wSQL & "        ON     a.Web商品フラグ    = t1.ShohinWebY "
wSQL = wSQL & "      INNER JOIN カテゴリー e WITH (NOLOCK) "				'2012/08/27 ok Add
wSQL = wSQL & "        ON     a.カテゴリーコード = e.カテゴリーコード "		'2012/08/27 ok Add
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        t1.ShohinWebY IS NOT NULL "
wSQL = wSQL & "    AND c.カテゴリーコード = '" & wCategoryCode & "' "

wSQL = wSQL & "ORDER BY "
wSQL = wSQL & "      a.カテゴリーコード "		'2012/08/27 ok Add
wSQL = wSQL & "    , b.メーカー名 "
wSQL = wSQL & "    , a.商品名 "
wSQL = wSQL & "    , d.色 "
wSQL = wSQL & "    , d.規格 "
' 2012/01/18 GV Mod End

'@@@@response.write(wSQL)

'2012/07/10 GV Mod Start
'call CreateOptionPartsHTML("オプション")
call CreateOptionPartsHTML("関連オプション")
'2012/07/10 GV Mod End

wOptionHTML = wHTML

End Function

'========================================================================
'
'	Function	パーツ HTML（データ抽出）
'
'		色規格に関係なく該当商品のパーツを取り出す
'
'========================================================================
'
Function CreatePartsHtml()

'---- Select パーツ
wSQL = ""
' 2012/01/18 GV Mod Start
'wSQL = wSQL & "SELECT c.パーツメーカーコード AS メーカーコード"
'wSQL = wSQL & "     , c.パーツ商品コード AS 商品コード"
'wSQL = wSQL & "     , d.色 AS 色"
'wSQL = wSQL & "     , d.規格 AS 規格"
'wSQL = wSQL & "     , a.商品名"
''wSQL = wSQL & "     , a.販売単価"   '2010/11/10 an del
'
''2010/11/10 an add s
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN a.B品フラグ = 'Y' THEN a.B品単価"
'wSQL = wSQL & "         WHEN a.個数限定数量 > a.個数限定受注済数量 THEN a.個数限定単価"
'wSQL = wSQL & "         ELSE a.販売単価"
'wSQL = wSQL & "       END AS 実販売単価"
''2010/11/10 an add e
'
'wSQL = wSQL & "     , a.商品概略Web"
'wSQL = wSQL & "     , a.商品画像ファイル名_小"
'wSQL = wSQL & "     , a.ASK商品フラグ"
'wSQL = wSQL & "     , a.取扱中止日"
'wSQL = wSQL & "     , a.廃番日"
'wSQL = wSQL & "     , a.完売日"
'wSQL = wSQL & "     , a.希少数量"
'wSQL = wSQL & "     , a.セット商品フラグ"
'wSQL = wSQL & "     , a.メーカー直送取寄区分"
'wSQL = wSQL & "     , a.Web納期非表示フラグ"
'wSQL = wSQL & "     , a.入荷予定未定フラグ"
'wSQL = wSQL & "     , a.B品フラグ"
'wSQL = wSQL & "     , a.個数限定数量"
'wSQL = wSQL & "     , a.個数限定受注済数量"
'wSQL = wSQL & "     , b.メーカー名"
'wSQL = wSQL & "     , d.引当可能数量"
'wSQL = wSQL & "     , d.発注数量"			'2011/06/09 hn add
'wSQL = wSQL & "     , d.引当可能入荷予定日"
'wSQL = wSQL & "     , d.B品引当可能数量"
'wSQL = wSQL & "  FROM Web商品 a"
'wSQL = wSQL & "     , メーカー b"
'wSQL = wSQL & "     , パーツ c"
'wSQL = wSQL & "     , Web色規格別在庫 d"
'wSQL = wSQL & " WHERE a.メーカーコード = c.パーツメーカーコード"
'wSQL = wSQL & "   AND a.商品コード = c.パーツ商品コード"
'wSQL = wSQL & "   AND b.メーカーコード = c.パーツメーカーコード"
'wSQL = wSQL & "   AND d.メーカーコード = c.パーツメーカーコード"
'wSQL = wSQL & "   AND d.商品コード = c.パーツ商品コード"
'wSQL = wSQL & "   AND d.色 = c.パーツ色"
'wSQL = wSQL & "   AND d.規格 = c.パーツ規格"
'wSQL = wSQL & "   AND c.メーカーコード = '" & maker_cd & "'"
'wSQL = wSQL & "   AND c.商品コード = '" & Replace(product_cd, "'", "''") & "'"	' 2012/01/23 GV Mod (コード内にシングルクオーテーションが存在した場合の対応)
'wSQL = wSQL & "   AND a.Web商品フラグ = 'Y'"
'
'wSQL = wSQL & " UNION "
'
'wSQL = wSQL & "SELECT c.パーツメーカーコード AS メーカーコード"
'wSQL = wSQL & "     , c.パーツ商品コード AS 商品コード"
'wSQL = wSQL & "     , d.色 AS 色"
'wSQL = wSQL & "     , d.規格 AS 規格"
'wSQL = wSQL & "     , a.商品名"
''wSQL = wSQL & "     , a.販売単価"   '2010/11/10 an del
'
''2010/11/10 an add s
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN a.B品フラグ = 'Y' THEN a.B品単価"
'wSQL = wSQL & "         WHEN a.個数限定数量 > a.個数限定受注済数量 THEN a.個数限定単価"
'wSQL = wSQL & "         ELSE a.販売単価"
'wSQL = wSQL & "       END AS 実販売単価"
''2010/11/10 an add e
'
'wSQL = wSQL & "     , a.商品概略Web"
'wSQL = wSQL & "     , a.商品画像ファイル名_小"
'wSQL = wSQL & "     , a.ASK商品フラグ"
'wSQL = wSQL & "     , a.取扱中止日"
'wSQL = wSQL & "     , a.廃番日"
'wSQL = wSQL & "     , a.完売日"
'wSQL = wSQL & "     , a.希少数量"
'wSQL = wSQL & "     , a.セット商品フラグ"
'wSQL = wSQL & "     , a.メーカー直送取寄区分"
'wSQL = wSQL & "     , a.Web納期非表示フラグ"
'wSQL = wSQL & "     , a.入荷予定未定フラグ"
'wSQL = wSQL & "     , a.B品フラグ"
'wSQL = wSQL & "     , a.個数限定数量"
'wSQL = wSQL & "     , a.個数限定受注済数量"
'wSQL = wSQL & "     , b.メーカー名"
'wSQL = wSQL & "     , d.引当可能数量"
'wSQL = wSQL & "     , d.発注数量"		'2011/06/09 hn add
'wSQL = wSQL & "     , d.引当可能入荷予定日"
'wSQL = wSQL & "     , d.B品引当可能数量"
'wSQL = wSQL & "  FROM Web商品 a"
'wSQL = wSQL & "     , メーカー b"
'wSQL = wSQL & "     , カテゴリー別パーツ c"
'wSQL = wSQL & "     , Web色規格別在庫 d"
'wSQL = wSQL & " WHERE a.メーカーコード = c.パーツメーカーコード"
'wSQL = wSQL & "   AND a.商品コード = c.パーツ商品コード"
'wSQL = wSQL & "   AND b.メーカーコード = c.パーツメーカーコード"
'wSQL = wSQL & "   AND d.メーカーコード = c.パーツメーカーコード"
'wSQL = wSQL & "   AND d.商品コード = c.パーツ商品コード"
'wSQL = wSQL & "   AND d.色 = c.パーツ色"
'wSQL = wSQL & "   AND d.規格 = c.パーツ規格"
'wSQL = wSQL & "   AND c.カテゴリーコード = '" & wCategoryCode & "'"
'wSQL = wSQL & "   AND a.Web商品フラグ = 'Y'"
'
'wSQL = wSQL & " ORDER BY"
'wSQL = wSQL & "       b.メーカー名"
'wSQL = wSQL & "     , a.商品名"
'wSQL = wSQL & "     , d.色"
'wSQL = wSQL & "     , d.規格"
wSQL = wSQL & "SELECT "
wSQL = wSQL & "      c.パーツメーカーコード AS メーカーコード "
wSQL = wSQL & "    , c.パーツ商品コード AS 商品コード "
wSQL = wSQL & "    , d.色 AS 色 "
wSQL = wSQL & "    , d.規格 AS 規格 "
wSQL = wSQL & "    , a.商品名 "
wSQL = wSQL & "    , CASE "
wSQL = wSQL & "        WHEN a.B品フラグ = 'Y'                     THEN a.B品単価 "
wSQL = wSQL & "        WHEN a.個数限定数量 > a.個数限定受注済数量 THEN a.個数限定単価 "
wSQL = wSQL & "        ELSE                                            a.販売単価 "
wSQL = wSQL & "      END AS 実販売単価 "
wSQL = wSQL & "    , a.商品概略Web "
wSQL = wSQL & "    , a.商品画像ファイル名_小 "
wSQL = wSQL & "    , a.ASK商品フラグ "
wSQL = wSQL & "    , a.取扱中止日 "
wSQL = wSQL & "    , a.廃番日 "
wSQL = wSQL & "    , a.完売日 "
wSQL = wSQL & "    , a.希少数量 "
wSQL = wSQL & "    , a.セット商品フラグ "
wSQL = wSQL & "    , a.メーカー直送取寄区分 "
wSQL = wSQL & "    , a.Web納期非表示フラグ "
wSQL = wSQL & "    , a.入荷予定未定フラグ "
wSQL = wSQL & "    , a.B品フラグ "
wSQL = wSQL & "    , a.個数限定数量 "
wSQL = wSQL & "    , a.個数限定受注済数量 "
wSQL = wSQL & "    , b.メーカー名 "
wSQL = wSQL & "    , d.引当可能数量 "
wSQL = wSQL & "    , d.発注数量 "
wSQL = wSQL & "    , d.引当可能入荷予定日 "
wSQL = wSQL & "    , d.B品引当可能数量 "
wSQL = wSQL & "    , a.カテゴリーコード "			'2012/08/27 ok Add
wSQL = wSQL & "    , e.カテゴリー名 "				'2012/08/27 ok Add
wSQL = wSQL & "FROM "
wSQL = wSQL & "    パーツ                       c WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN Web商品         a WITH (NOLOCK) "
wSQL = wSQL & "        ON     a.メーカーコード = c.パーツメーカーコード "
wSQL = wSQL & "           AND a.商品コード     = c.パーツ商品コード "
wSQL = wSQL & "      INNER JOIN メーカー        b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.メーカーコード = c.パーツメーカーコード "
wSQL = wSQL & "      INNER JOIN Web色規格別在庫 d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.メーカーコード = c.パーツメーカーコード "
wSQL = wSQL & "           AND d.商品コード     = c.パーツ商品コード "
wSQL = wSQL & "           AND d.色             = c.パーツ色 "
wSQL = wSQL & "           AND d.規格           = c.パーツ規格 "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' )   t1 "
wSQL = wSQL & "        ON     a.Web商品フラグ    = t1.ShohinWebY "
wSQL = wSQL & "      INNER JOIN カテゴリー e WITH (NOLOCK) "				'2012/08/27 ok Add
wSQL = wSQL & "        ON     a.カテゴリーコード = e.カテゴリーコード "		'2012/08/27 ok Add
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        t1.ShohinWebY IS NOT NULL "
wSQL = wSQL & "    AND c.メーカーコード = '" & maker_cd & "' "
wSQL = wSQL & "    AND c.商品コード     = '" & Replace(product_cd, "'", "''") & "' "	' 2012/01/23 GV Mod (コード内にシングルクオーテーションが存在した場合の対応)

wSQL = wSQL & " UNION "

wSQL = wSQL & "SELECT "
wSQL = wSQL & "      c.パーツメーカーコード AS メーカーコード "
wSQL = wSQL & "    , c.パーツ商品コード AS 商品コード "
wSQL = wSQL & "    , d.色 AS 色 "
wSQL = wSQL & "    , d.規格 AS 規格 "
wSQL = wSQL & "    , a.商品名 "
wSQL = wSQL & "    , CASE "
wSQL = wSQL & "        WHEN a.B品フラグ = 'Y'                     THEN a.B品単価 "
wSQL = wSQL & "        WHEN a.個数限定数量 > a.個数限定受注済数量 THEN a.個数限定単価 "
wSQL = wSQL & "        ELSE                                            a.販売単価 "
wSQL = wSQL & "      END AS 実販売単価 "
wSQL = wSQL & "    , a.商品概略Web "
wSQL = wSQL & "    , a.商品画像ファイル名_小 "
wSQL = wSQL & "    , a.ASK商品フラグ "
wSQL = wSQL & "    , a.取扱中止日 "
wSQL = wSQL & "    , a.廃番日 "
wSQL = wSQL & "    , a.完売日 "
wSQL = wSQL & "    , a.希少数量 "
wSQL = wSQL & "    , a.セット商品フラグ "
wSQL = wSQL & "    , a.メーカー直送取寄区分 "
wSQL = wSQL & "    , a.Web納期非表示フラグ "
wSQL = wSQL & "    , a.入荷予定未定フラグ "
wSQL = wSQL & "    , a.B品フラグ "
wSQL = wSQL & "    , a.個数限定数量 "
wSQL = wSQL & "    , a.個数限定受注済数量 "
wSQL = wSQL & "    , b.メーカー名 "
wSQL = wSQL & "    , d.引当可能数量 "
wSQL = wSQL & "    , d.発注数量 "
wSQL = wSQL & "    , d.引当可能入荷予定日 "
wSQL = wSQL & "    , d.B品引当可能数量 "
wSQL = wSQL & "    , a.カテゴリーコード "		'2012/08/27 ok Add
wSQL = wSQL & "    , e.カテゴリー名 "			'2012/08/27 ok Add
wSQL = wSQL & "FROM "
wSQL = wSQL & "    カテゴリー別パーツ           c WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN Web商品         a WITH (NOLOCK) "
wSQL = wSQL & "        ON     a.メーカーコード = c.パーツメーカーコード "
wSQL = wSQL & "           AND a.商品コード     = c.パーツ商品コード "
wSQL = wSQL & "      INNER JOIN メーカー        b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.メーカーコード = c.パーツメーカーコード "
wSQL = wSQL & "      INNER JOIN Web色規格別在庫 d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.メーカーコード = c.パーツメーカーコード "
wSQL = wSQL & "           AND d.商品コード     = c.パーツ商品コード "
wSQL = wSQL & "           AND d.色             = c.パーツ色 "
wSQL = wSQL & "           AND d.規格           = c.パーツ規格 "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' )   t1 "
wSQL = wSQL & "        ON     a.Web商品フラグ    = t1.ShohinWebY "
wSQL = wSQL & "      INNER JOIN カテゴリー e WITH (NOLOCK) "				'2012/08/27 ok Add
wSQL = wSQL & "        ON     a.カテゴリーコード = e.カテゴリーコード "		'2012/08/27 ok Add
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        t1.ShohinWebY IS NOT NULL "
wSQL = wSQL & "    AND c.カテゴリーコード = '" & wCategoryCode & "' "

wSQL = wSQL & "ORDER BY "
wSQL = wSQL & "      a.カテゴリーコード "			'2012/08/27 ok Add
wSQL = wSQL & "    , b.メーカー名 "
wSQL = wSQL & "    , a.商品名 "
wSQL = wSQL & "    , d.色 "
wSQL = wSQL & "    , d.規格 "
' 2012/01/18 GV Mod End

'@@@@@response.write(wSQL)

'2012/07/10 GV Mod Start
'call CreateOptionPartsHTML("パーツ")
call CreateOptionPartsHTML("関連パーツ")
'2012/07/10 GV Mod End

wPartsHtml = wHTML

End Function

'========================================================================
'
'	Function	オプション、パーツ HTML作成（共通）
'
'	Parm: pTitle(タイトル)
'
'========================================================================
'
Function CreateOptionPartsHTML(pTitle)

Dim RSv
Dim vInventoryCD
Dim vInventoryImage
Dim vProdTermFl		'2010/12/28 hn add
Dim i
Dim j
Dim vCategoryCode			'2012/08/27 ok Add

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

wHTML = ""

if RSv.EOF = false then
	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "<table width='602' border='0' cellspacing='0' cellpadding='0' id='Shop_Option_Parts_title'>" & vbNewLine
'	wHTML = wHTML & "<form name='fBuyTogether' method='post'>" & vbNewLine
'	wHTML = wHTML & "  <tr>" & vbNewLine
'	wHTML = wHTML & "    <td align='left'>&nbsp;<b>" & pTitle & "</b></td>" & vbNewLine
'	wHTML = wHTML & "    <td align='left'><a href='#top'><img src='images/goes_up.gif' width='18' height='18' border='0' align='right'></a></td>" & vbNewLine
'	wHTML = wHTML & "  </tr>" & vbNewLine
'	wHTML = wHTML & "</table>" & vbNewLine
'	wHTML = wHTML & "<table width='602' border='1' cellpadding='0' cellspacing='0' id='Shop_Option_Parts_Frame'>" & vbNewLine
	wHTML = wHTML & "<h2 class='detail_title'>" & pTitle & "</h2>" & vbNewLine
	wHTML = wHTML & "<form name='fBuyTogether' method='post'>" & vbNewLine
	'2012/07/10 GV Mod End

	vCategoryCode = ""		'2012/08/27 ok Add
	Do While RSv.EOF = false
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "  <tr>" & vbNewLine
'2012/08/27 ok Add Start
		if vCategoryCode = "" And wOptionPartsTitleFlag = "Y" Then
			wHTML = wHTML & "  <div class='headline'>" & vbNewLine
			wHTML = wHTML & "    <h3>" & RSv("カテゴリー名") & "</h3>" & vbNewLine
			wHTML = wHTML & "  </div>" & vbNewLine
			vCategoryCode = RSv("カテゴリーコード")
		End If
'2012/08/27 ok Add End

		wHTML = wHTML & "<ul class='relation'>" & vbNewLine
		'2012/07/10 GV Mod End

		'2012/07/10 GV Mod Start
'		For i=1 To 5
		For i=1 To 4
		'2012/07/10 GV Mod End
			'---- 廃番チェック
			if  (isNull(RSv("取扱中止日")) = true AND isNull(RSv("廃番日")) = true) _
			 OR (isNull(RSv("廃番日")) = false AND (RSv("引当可能数量") > 0 OR RSv("発注数量") > 0)) _
			 OR (isNull(RSv("完売日")) = false) then		'2011/06/09 hn mod
				vProdTermFl = "N"		'2010/12/28 hn mod
			else
				vProdTermFl = "Y"		'2010/12/28 hn mod
			end if

			'2012/07/10 GV Del Start
'			wHTML = wHTML & "    <td>" & vbNewLine
'			wHTML = wHTML & "      <table border='0' cellspacing='0' cellpadding='0' id='Shop_Option_Parts_product'>" & vbNewLine
'			wHTML = wHTML & "        <tr>" & vbNewLine
			'2012/07/10 GV Del End

			'---- 商品画像、商品名
			'2012/07/10 GV Mod Start
'			wHTML = wHTML & "          <td><a href='ProductDetail.asp?Item=" & Server.URLEncode(RSv("メーカーコード") & "^" & RSv("商品コード") & "^" & Trim(RSv("色")) & "^" & Trim(RSv("規格"))) & "'><img src='prod_img/" & RSv("商品画像ファイル名_小") & "' width='100' height='50' border='0'><br>" & RSv("メーカー名") & "<br>" & RSv("商品名") & "</a><br></td>" & vbNewLine
'			wHTML = wHTML & "        </tr>" & vbNewLine
			wHTML = wHTML & "  <li>" & vbNewLine
			wHTML = wHTML & "    <p><a href='ProductDetail.asp?Item=" & Server.URLEncode(RSv("メーカーコード") & "^" & RSv("商品コード") & "^" & Trim(RSv("色")) & "^" & Trim(RSv("規格"))) & "'>"
			If RSv("商品画像ファイル名_小") <> "" Then
				wHTML = wHTML & "<img src='prod_img/" & RSv("商品画像ファイル名_小") & "' alt='" & Replace(RSv("メーカー名") & " " & RSv("商品名"),"'","&#39;") & "' class='opover'>"
			Else
				wHTML = wHTML & "<img src=""prod_img/n/nopict-.jpg"" alt="""">"
			End If
			wHTML = wHTML & RSv("メーカー名") & " / " & RSv("商品名") & "</a></p>" & vbNewLine
			'2012/07/10 GV Mod End

			wHTML = wHTML & "    <div class='box'>" & vbNewLine	'2012/07/10 GV Add
			'----- 販売単価
			wPrice = calcPrice(RSv("実販売単価"), wSalesTaxRate)  '2010/11/10 an mod
'			wHTML = wHTML & "        <tr>" & vbNewLine	'2012/07/10 GV Del
			if RSv("ASK商品フラグ") = "Y" then
'2011/10/19 hn mod s
'				wHTML = wHTML & "          <td>ASK</td>" & vbNewLine
				'2012/07/10 GV Mod Start
'				wHTML = wHTML & "          <td><a class='tip'>ASK<span>"  & FormatNumber(wPrice,0) & "円(税込)</span></a></td>" & vbNewLine
'2014/03/19 GV mod start ---->
'				wHTML = wHTML & "      <p><a class='tip'>ASK<span>"  & FormatNumber(wPrice,0) & "円(税込)</span></a></p>" & vbNewLine
				wHTML = wHTML & "      <p><a class='tip'>ASK<span class='exc-tax'>"  & FormatNumber(RSv("実販売単価"),0) & "円(税抜)</span><br>"
				wHTML = wHTML & "      <span class='inc-tax'>(税込&nbsp;"  & FormatNumber(wPrice,0) & "円)</span></a></p>" & vbNewLine
'2014/03/19 GV mod end <-----
				'2012/07/10 GV Mod End
'2011/10/19 hn mod e

			else
				'2012/07/10 GV Mod Start
'				wHTML = wHTML & "          <td>" & FormatNumber(wPrice,0) & "円(税込)</td>" & vbNewLine
'2014/03/19 GV mod start ---->
'				wHTML = wHTML & "      <p>"  & FormatNumber(wPrice,0) & "円(税込)</p>" & vbNewLine
				wHTML = wHTML & "      <p>"  & FormatNumber(RSv("実販売単価"),0) & "円(税抜)</p>" & vbNewLine
				wHTML = wHTML & "      <p>(税込&nbsp;"  & FormatNumber(wPrice,0) & "円)</p>" & vbNewLine
'2014/03/19 GV mod end <-----
				'2012/07/10 GV Mod End
			end if
'			wHTML = wHTML & "        </tr>" & vbNewLine	'2012/07/10 GV Del

			'----- 在庫状況
			vInventoryCd = GetInventoryStatus(RSv("メーカーコード"),RSv("商品コード"),RSv("色"),RSv("規格"),RSv("引当可能数量"),RSv("希少数量"),RSv("セット商品フラグ"),RSv("メーカー直送取寄区分"),RSv("引当可能入荷予定日"),vProdTermFl)  		'2010/12/28 hn mod

			'---- 在庫状況、色を最終セット
			call GetInventoryStatus2(RSv("引当可能数量"), RSv("Web納期非表示フラグ"), RSv("入荷予定未定フラグ"), RSv("廃番日"), RSv("B品フラグ"), RSv("B品引当可能数量"), RSv("個数限定数量"), RSv("個数限定受注済数量"), vProdTermFl, vInventoryCd, vInventoryImage)		'2010/12/28 hn mod

			'----
			'2012/07/10 GV Mod Start
'			wHTML = wHTML & "        <tr>" & vbNewLine
'			wHTML = wHTML & "          <td><img src='images/" & vInventoryImage & "' width='10' height='10'> " & vInventoryCd & "</td>" & vbNewLine
'			wHTML = wHTML & "        </tr>" & vbNewLine
			wHTML = wHTML & "      <p class='stock'><img src='images/" & vInventoryImage & "' alt='" & vInventoryCd & "'>" & vInventoryCd & "</p>" & vbNewLine
			'2012/07/10 GV Mod End

		'----- 一緒に購入する
'			wHTML = wHTML & "        <tr>" & vbNewLine	'2012/07/10 GV Del

			if vInventoryCd = "取扱中止" then
				'2012/07/10 GV Mod Start
'				wHTML = wHTML & "          <td class='prod_cart'>&nbsp;</td>" & vbNewLine
				wHTML = wHTML & "      <p class='together'>&nbsp;</p>" & vbNewLine
				'2012/07/10 GV Mod End
			else
				'2012/07/10 GV Mod Start
'				wHTML = wHTML & "          <td class='prod_cart'><input type='checkbox' name='iBuyTogether' value='" & RSv("メーカーコード") & "^" & RSv("商品コード") & "^" & Trim(RSv("色")) & "^" & Trim(RSv("規格")) & "' id='checkbox' onClick='BuyTogether_onClick(this);'>一緒に購入する</td>" & vbNewLine
				wHTML = wHTML & "      <p class='together'><input type='checkbox' name='iBuyTogether' value='" & RSv("メーカーコード") & "^" & RSv("商品コード") & "^" & Trim(RSv("色")) & "^" & Trim(RSv("規格")) & "' onClick='BuyTogether_onClick(this);'>一緒に購入</p>" & vbNewLine
				'2012/07/10 GV Mod End

			end if

			'2012/07/10 GV Mod Start
'			wHTML = wHTML & "        </tr>" & vbNewLine
'			wHTML = wHTML & "      </table>" & vbNewLine
'			wHTML = wHTML & "    </td>" & vbNewLine
			wHTML = wHTML & "    </div>" & vbNewLine
			wHTML = wHTML & "  </li>" & vbNewLine
			'2012/07/10 GV Mod End

			RSv.MoveNext

			'---- 1行5明細以内の時は空明細を作る
			if RSv.EOF = true then
				'2012/07/10 GV Del Start
'				For j=i+1 to 5
'					wHTML = wHTML & "    <td>" & vbNewLine
'					wHTML = wHTML & "      <table border='0' cellspacing='0' cellpadding='0' id='Shop_Option_Parts_product'>" & vbNewLine
'					wHTML = wHTML & "        <tr>" & vbNewLine
'					wHTML = wHTML & "          <td>&nbsp;</td>" & vbNewLine
'					wHTML = wHTML & "        </tr>" & vbNewLine
'					wHTML = wHTML & "      </table>" & vbNewLine
'					wHTML = wHTML & "    </td>" & vbNewLine
'				Next
				'2012/07/10 GV Del End
				i = 5
'2012/08/27 ok Add Start
			else
				if vCategoryCode <> RSv("カテゴリーコード") And wOptionPartsTitleFlag = "Y" Then
					vCategoryCode = ""
					i = 5
				end if
'2012/08/27 ok Add End
			end if
		Next

		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "  </tr>" & vbNewLine
		wHTML = wHTML & "</ul>" & vbNewLine
		'2012/07/10 GV Mod End
	Loop

	wHTML = wHTML & "</form>" & vbNewLine
'	wHTML = wHTML & "</table>" & vbNewLine	'2012/07/10 GV Del

	wOptionPartsFl = true
end if

RSv.Close

End Function

'========================================================================
'
'	Function	カスタマーレビュー、評価 HTML作成
'
'========================================================================
'
Function CreateReviewHTML()

Dim vAvgRating
Dim v1Cnt
Dim v0Cnt
Dim vHalfCnt
Dim vTotalCnt
Dim vOnpu
Dim RSv
Dim i

'---- Select 商品レビュー 平均，件数 取得
wSQL = ""
' 2012/01/23 GV Mod Start
'wSQL = wSQL & "SELECT SUM(a.評価) AS 評価合計"
'wSQL = wSQL & "     , COUNT(a.ID) AS レビュー数"
'wSQL = wSQL & "  FROM 商品レビュー a WITH (NOLOCK) "				' 2012/01/18 GV Mod  WITH (NOLOCK)付加
'wSQL = wSQL & " WHERE a.メーカーコード = '" & maker_cd & "'"
'wSQL = wSQL & "   AND a.商品コード = '" & product_cd & "'"
'
''@@@@@@response.write(wSQL)
'
'Set RSv = Server.CreateObject("ADODB.Recordset")
'RSv.Open wSQL, Connection, adOpenStatic
'
'if RSv("レビュー数") = 0 then
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

wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.レビュー評価平均 "
wSQL = wSQL & "    , a.レビュー件数 "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    商品レビュー集計 a WITH (NOLOCK) "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        a.メーカーコード = '" & maker_cd & "' "
wSQL = wSQL & "    AND a.商品コード     = '" & Replace(product_cd, "'", "''") & "' "

'@@@@@@response.write(wSQL)

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

'---- 音符画像作成
vOnpu = ""
For i=1 to v1Cnt
	'2012/07/10 GV Mod Start
'	vOnpu = vOnpu & "<img src='images/onpu1.jpg' width='20' height='18'>"
	vOnpu = vOnpu & "<img src='images/review_icon10.png' alt='1'>"
	'2012/07/10 GV Mod End
Next
if vHalfcnt = 1 then
	'2012/07/10 GV Mod Strat
'	vOnpu = vOnpu & "<img src='images/onpuHalf.jpg' width='20' height='18'>"
	vOnpu = vOnpu & "<img src='images/review_icon05.png' alt='0.5'>"
	'2012/07/10 GV Mod End
end if
For i=1 to v0Cnt
	'2012/07/10 GV Mod Start
'	vOnpu = vOnpu & "<img src='images/onpu0.jpg' width='20' height='18'>"
	vOnpu = vOnpu & "<img src='images/review_icon00.png' alt='0'>"
	'2012/07/10 GV Mod End
Next

wHTML = ""

'---- 評価編集
'2012/07/10 GV Mod Start
'wHTML = wHTML & "<table width='188' border='0' cellspacing='0' cellpadding='0' id='Shop_right'>" & vbNewLine
'wHTML = wHTML & "  <tr>" & vbNewLine
'wHTML = wHTML & "    <td class='head'>評価</td>" & vbNewLine
'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & "  <tr>" & vbNewLine
'wHTML = wHTML & "    <td class='base' itemprop='review' itemscope itemtype='http://data-vocabulary.org/Review-aggregate'>" & vbNewLine        '2011/11/22 an mod
'wHTML = wHTML & "      <table width='180' border='0' cellspacing='0' cellpadding='0' id='Shop_right_product'>" & vbNewLine
'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td align='left' width='80'>おすすめ度：</td>" & vbNewLine
'wHTML = wHTML & "          <td align='left' width='100'><span itemprop='rating'>" & FormatNumber(vAvgRating,1) & "</span></td>" & vbNewLine   '2011/11/22 an mod
'wHTML = wHTML & "        </tr>" & vbNewLine
'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td colspan='2' align='center' height='26'>" & vOnpu & "</td>" & vbNewLine
'wHTML = wHTML & "        </tr>" & vbNewLine
'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td align='left'>レビュー数：</td>" & vbNewLine
'wHTML = wHTML & "          <td align='left'><span itemprop='count'>" & vTotalCnt & "</span></td>" & vbNewLine   '2011/11/22 an mod
'wHTML = wHTML & "        </tr>" & vbNewLine
'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td colspan='2' align='center' height='26'><a href='#review'><img src='images/Reviews.gif' border='0'></a></td>" & vbNewLine
'wHTML = wHTML & "        </tr>" & vbNewLine
'wHTML = wHTML & "      </table>" & vbNewLine

'wHTML = wHTML & "    </td>" & vbNewLine
'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & "</table>" & vbNewLine
wHTML = wHTML & "<div class='review'>" & vbNewLine
wHTML = wHTML & "  <p><strong>評価：</strong>" & vOnpu & "</p>" & vbNewLine
wHTML = wHTML & "  <p><a href='#review'>レビュー数：" & vTotalCnt & "</a></p>" & vbNewLine
wHTML = wHTML & "</div>" & vbNewLine
'2012/07/10 GV Mod End

wHyoukaHTML = wHTML

'----ここからカスタマーレビュー ===========
'---- 総合評価編集
wHTML = ""
'2012/07/10 GV Mod Start
'wHTML = wHTML & "<table width='602' height='50' border='0' cellspacing='0' cellpadding='0' id='Shop_review_head'>" & vbNewLine

'---- おすすめ度
'wHTML = wHTML & "  <tr>" & vbNewLine
'wHTML = wHTML & "    <td width='80' align='center'>総合評価</td>" & vbNewLine
'wHTML = wHTML & "    <td width='110' nowrap>" & vOnpu & "</td>" & vbNewLine
'wHTML = wHTML & "    <td width='50' nowrap><b>(" & FormatNumber(vAvgRating,1) & ")</b></td>" & vbNewLine
wHTML = wHTML & "<div class='comment_box'>" & vbNewLine
wHTML = wHTML & "  <ul id='totalreview' itemprop='review' itemscope itemtype='http://data-vocabulary.org/Review-aggregate'>" & vbNewLine
wHTML = wHTML & "    <li><span class='review_icon'>総合評価：" & vOnpu & "<span itemprop='rating'>(" & FormatNumber(vAvgRating,1) & ")</span></span></li>" & vbNewLine

'---- レビュー数
'wHTML = wHTML & "    <td nowrap><b>レビュー数： " & vTotalCnt & "</b></td>" & vbNewLine
'wHTML = wHTML & "    <td align='left'><a href='#top'><img src='images/goes_up.gif' width='18' height='18' border='0' align='right'></a></td>" & vbNewLine
'wHTML = wHTML & "  </tr>" & vbNewLine
wHTML = wHTML & "    <li>レビュー数：<span itemprop='count'>" & vTotalCnt & "</span></li>" & vbNewLine
wHTML = wHTML & "  </ul>" & vbNewLine
wHTML = wHTML & "</div>" & vbNewLine

'wHTML = wHTML & "</table>" & vbNewLine
'2012/07/10 GV Mod End

'---- Select 個別商品レビュー
wSQL = ""
if ReviewAll = "Y" then
	wSQL = wSQL & "SELECT "
else
	wSQL = wSQL & "SELECT TOP 5 "
end if
wSQL = wSQL & "      a.ID "
wSQL = wSQL & "    , a.投稿日 "
wSQL = wSQL & "    , a.評価 "
wSQL = wSQL & "    , a.タイトル "
wSQL = wSQL & "    , a.名前 "
wSQL = wSQL & "    , a.レビュー内容 "
wSQL = wSQL & "    , a.参考数 "
wSQL = wSQL & "    , a.不参考数 "
wSQL = wSQL & "    , a.顧客番号 "
wSQL = wSQL & "    , a.ショップコメント日 "
wSQL = wSQL & "    , a.ショップコメントタイトル "
wSQL = wSQL & "    , a.ショップコメント "
wSQL = wSQL & "    , b.顧客都道府県 "
' 2012/01/18 GV Mod Start (WITH (NOLOCK)付加)
'wSQL = wSQL & "  FROM 商品レビュー a LEFT JOIN Web顧客住所 b"
'wSQL = wSQL & "                             ON b.顧客番号 = a.顧客番号"
'wSQL = wSQL & "                            AND b.住所連番 = 1"
wSQL = wSQL & "FROM "
wSQL = wSQL & "    商品レビュー            a WITH (NOLOCK) "
wSQL = wSQL & "      LEFT JOIN Web顧客住所 b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.顧客番号 = a.顧客番号 "
wSQL = wSQL & "           AND b.住所連番 = 1 "
' 2012/01/18 GV Mod End
wSQL = wSQL & " WHERE a.メーカーコード = '" & maker_cd & "'"
wSQL = wSQL & "   AND a.商品コード = '" & Replace(product_cd, "'", "''") & "'"	' 2012/01/23 GV Mod (コード内にシングルクオーテーションが存在した場合の対応)
wSQL = wSQL & " ORDER BY"
' 2012/01/18 GV Mod Start
'wSQL = wSQL & "       a.投稿日 DESC"
wSQL = wSQL & "       a.ID DESC"
' 2012/01/18 GV Mod End

'@@@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

'wHTML = wHTML & "<table width='602' border='0' cellspacing='0' cellpadding='5' id='shop_border'>" & vbNewLine

'---- 個別レビュー編集
Do While RSv.EOF = false

	'2012/07/10 GV Mod Start
'---- おすすめ度
'	wHTML = wHTML & "  <tr class='honbun'>" & vbNewLine
'	wHTML = wHTML & "    <td width='130' nowrap class='Shop_review_th Shop_review_th_onpu'>"
'	For i=1 to RSv("評価")
'		wHTML = wHTML & "<img src='images/onpu1.jpg' width='20' height='18'>"
'	Next
'	For i=RSv("評価")+1 to 5
'		wHTML = wHTML & "<img src='images/onpu0.jpg' width='20' height='18'>"
'	Next
'	wHTML = wHTML & " (" & FormatNumber(RSv("評価"), 1) & ")" & vbNewLine
'	wHTML = wHTML & "</td>" & vbNewLine

'---- タイトル, 投稿日
'	wHTML = wHTML & "    <td width='400' class='Shop_review_th'><h3>" & RSv("タイトル") & "</h3></td>" & vbNewLine
'	wHTML = wHTML & "    <td align='right' nowrap class='Shop_review_th Shop_review_th_right'><span>レビューID：" & RSv("ID") & "</span><br>" & cf_FormatDate(RSv("投稿日"), "YYYY/MM/DD") & "</td>" & vbNewLine   '2011/09/09 an mod
'	wHTML = wHTML & "  </tr>" & vbNewLine

'---- 投稿日，おすすめ度，タイトル
	wHTML = wHTML & "<div class='comment_box'>" & vbNewLine
	wHTML = wHTML & "  <p>" & cf_FormatDate(RSv("投稿日"), "YYYY/MM/DD") & "</p>" & vbNewLine
	wHTML = wHTML & "  <p class='subject'><span class='review_icon'>"
	For i=1 to RSv("評価")
		wHTML = wHTML & "<img src='images/review_icon10.png' alt='1'>"
	Next
	wHTML = wHTML & "</span>" & RSv("タイトル") & "</p>" & vbNewLine
	'2012/07/10 GV Mod End
	'2012/07/10 GV Mod Start
'---- 投稿者名，都道府県、この人のレビューリンク、参考になった人数
'	wHTML = wHTML & "  <tr>" & vbNewLine
'	wHTML = wHTML & "    <td colspan='3' nowrap>" & vbNewLine
'	wHTML = wHTML & "      <table width='100%' cellspacing='0' cellpadding='0'>" & vbNewLine
'	wHTML = wHTML & "        <tr class='honbun'>" & vbNewLine
'	wHTML = wHTML & "          <td nowrap>投稿者名：" & RSv("名前")

'	if IsNull(RSv("顧客番号")) = false then
'		if RSv("顧客番号") <> 0 then
'			wHTML = wHTML & " 【" & RSv("顧客都道府県") & "】 <a href='ReviewAllByCustomer.asp?CNo=" & RSv("顧客番号") & "' class='link'><b>レビューを見る</b></a>"
'		end if
'	end if

'	wHTML = wHTML & "</td>" & vbNewLine

'	wHTML = wHTML & "          <td align='right' nowrap>参考になった人数：" & RSv("参考数") & "人(" & RSv("参考数") + RSv("不参考数") & "人中)</td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine
'	wHTML = wHTML & "      </table>" & vbNewLine
'	wHTML = wHTML & "    </td>" & vbNewLine
'	wHTML = wHTML & "  </tr>" & vbNewLine

'--- 投稿者名(この人のレビューリンク)，都道府県
	if IsNull(RSv("顧客番号")) = false then
		if RSv("顧客番号") <> 0 then
			wHTML = wHTML & "  <p class='postname'>投稿者名：<a href='ReviewAllByCustomer.asp?CNo=" & RSv("顧客番号") & "'>" & RSv("名前") & "</a><span>"
			wHTML = wHTML & " 【" & RSv("顧客都道府県") & "】" & vbNewLine
		end if
	end if
	wHTML = wHTML & "</span></p>"
	'2012/07/10 GV Mod End

'---- レビュー内容
	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "  <tr class='honbun'>" & vbNewLine
'	wHTML = wHTML & "    <td colspan='3' class='Shop_review_text'><p>" & Replace(RSv("レビュー内容"), vbNewline, "<br>") & "</p></td>" & vbNewLine
	wHTML = wHTML & "  <p>" & Replace(RSv("レビュー内容"), vbNewline, "<br>") & "</p>" & vbNewLine
'	wHTML = wHTML & "  </tr>" & vbNewLine
	'2012/07/10 GV Mod End

'---- レビュー内容  2010/03/08 hn changed
	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "    <tr>" & vbNewLine
'	wHTML = wHTML & "      <td colspan='3' align='right' class='Shop_review_btn_td'>" & vbNewLine
'	wHTML = wHTML & "        <div class='Shop_review_yn_wrap'>" & vbNewLine
'	wHTML = wHTML & "          <div class='Shop_review_yntxt'>参考になりましたか？</div>" & vbNewLine
'	wHTML = wHTML & "          <div class='Shop_review_ynbtn'>" & vbNewLine
'	wHTML = wHTML & "            <img src='images/btn_yes20.jpg' alt='YES' width='34' height='20' border='0' onMouseOver='this.src=""images/btn_yes20-.jpg"";' onMouseOut='this.src=""images/btn_yes20.jpg"";' onClick='ReviewSankou_onClick(""" & RSv("ID") & """,""" & item & """,""Y"");'>" & vbNewLine
'	wHTML = wHTML & "          </div>" & vbNewLine
'	wHTML = wHTML & "          <div class='Shop_review_slash'>/</div>" & vbNewLine
'	wHTML = wHTML & "          <div class='Shop_review_ynbtn'>" & vbNewLine
'	wHTML = wHTML & "            <img src='images/btn_no20.jpg' alt='NO' width='34' height='20' border='0' onMouseOver='this.src=""images/btn_no20-.jpg"";' onMouseOut='this.src=""images/btn_no20.jpg"";' onClick='ReviewSankou_onClick(""" & RSv("ID") & """,""" & item & """,""N"");'>" & vbNewLine
'	wHTML = wHTML & "          </div>" & vbNewLine
'	wHTML = wHTML & "        </div>" & vbNewLine
'	wHTML = wHTML & "      </td>" & vbNewLine
'	wHTML = wHTML & "    </tr>" & vbNewLine

'2013/05/17 GV #1507 add start
'---- 自分のコメントは編集
If (Trim(RSv("顧客番号")) = CStr(wUserID)) Then
	wHTML = wHTML & "  <p id='review_edit'><a href='" & g_HTTPS & "shop/ReviewWrite.asp?Item=" & Server.URLEncode(Item) & "'><img src='images/btn_review_edit.png' alt='このレビューを編集する' class='opover'></a></p>"
End If
'2013/05/17 GV #1507 add end

	wHTML = wHTML & "  <div class='review_other'>"	& vbNewLine
	wHTML = wHTML & "    <p class='review_id'>レビューID：" & RSv("ID") & "</p>"	& vbNewLine
	wHTML = wHTML & "    <p>参考になった人数：" & RSv("参考数") & "人(" & RSv("参考数") + RSv("不参考数") & "人中)</p>"	& vbNewLine
	wHTML = wHTML & "    <dl>"	& vbNewLine
	wHTML = wHTML & "      <dt>参考になりましたか？</dt>"	& vbNewLine
	wHTML = wHTML & "      <dd><img src='images/btn_yes20.jpg' alt='Yes' class='opover' onClick='ReviewSankou_onClick(""" & RSv("ID") & """,""" & item & """,""Y"");'></dd>" & vbNewLine
	wHTML = wHTML & "      <dd><img src='images/btn_no20.jpg' alt='NO' class='opover' onClick='ReviewSankou_onClick(""" & RSv("ID") & """,""" & item & """,""N"");'></dd>" & vbNewLine
	wHTML = wHTML & "    </dl>"	& vbNewLine
	wHTML = wHTML & "  </div>"	& vbNewLine
	'2012/07/10 GV Mod End

'---- ショップコメント 2010/03/08 an changed
	if IsNull(RSv("ショップコメント日")) = false then
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "  <tr>" & vbNewLine
'		wHTML = wHTML & "    <td colspan='3' class='Shop_review_text'>" & vbNewLine
'		wHTML = wHTML & "      <div class='Shop_review_sh_res'>" & vbNewLine
'		wHTML = wHTML & "        <div class='Shop_review_sh_res_head'><span>" & RSv("ショップコメントタイトル") & "</span> " & cf_FormatDate(RSv("ショップコメント日"), "YYYY/MM/DD") & "</div>" & vbNewLine
'		wHTML = wHTML & "        <div class='Shop_review_text'><p>" & Replace(RSv("ショップコメント"), vbNewline, "<br>") & "</p></div>" & vbNewLine
'		wHTML = wHTML & "      </div>" & vbNewLine
'		wHTML = wHTML & "    </td>" & vbNewLine
'		wHTML = wHTML & "  </tr>" & vbNewLine
		wHTML = wHTML & "  <div class='reply_box'>" & vbNewLine
		wHTML = wHTML & "    <p>" & cf_FormatDate(RSv("ショップコメント日"), "YYYY/MM/DD") & "</p><br>" & vbNewLine
'		wHTML = wHTML & "    <p class='ansewr'>" & RSv("ショップコメントタイトル") & "</p>" & vbNewLine
		wHTML = wHTML & "    <p>" & Replace(RSv("ショップコメント"), vbNewline, "<br>") & "</p>" & vbNewLine
		wHTML = wHTML & "  </div>" & vbNewLine
		'2012/07/10 GV Mod End
	end if

'2013/05/17 GV #1507 Mod Start
'使用してないのでコメントアウト
'---- ショップコメント書き込み
'	if iShop = "Y" then
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "  <tr>" & vbNewLine
'		wHTML = wHTML & "    <td colspan='3'><a href='ReviewShopComment.asp?id=" & RSv("ID") & "' class='link'><b>●ショップコメントを書く</b></a></td>" & vbNewLine
'		wHTML = wHTML & "    <a href='ReviewShopComment.asp?id=" & RSv("ID") & "' class='link'><b>●ショップコメントを書く</b></a>" & vbNewLine
'		wHTML = wHTML & "  <tr>" & vbNewLine
		'2012/07/10 GV Mod End
'	end if
'2013/05/17 GV #1507 Mod End
	wHTML = wHTML & "</div>" & vbNewLine	'class='comment_box'

'---- 区切り線
'	wHTML = wHTML & "  <tr>" & vbNewLine
'	wHTML = wHTML & "    <td colspan='3'><hr width='99%' size='1'></td>" & vbNewLine
'	wHTML = wHTML & "  <tr>" & vbNewLine & vbNewLine


	RSv.MoveNext
Loop

'wHTML = wHTML & "</table>" & vbNewLine	'2012/07/10 GV Del

RSv.Close

wHTML = wHTML & "<ul class='btn_review'>" & vbNewLine
'---- 全てのレビューを見る
if ReviewAll <> "Y" AND vTotalCnt > 5 then
	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "<table width='602' border='0' cellspacing='0' cellpadding='3'>" & vbNewLine
'	wHTML = wHTML & "  <tr>" & vbNewLine
'	wHTML = wHTML & "    <td align='right'><a href='ProductDetail.asp?ReviewAll=Y&Item=" & item & "'><img src='images/ReviewAll.gif' border='0'></a></td>" & vbNewLine
	wHTML = wHTML & "  <li><a href='ProductDetail.asp?ReviewAll=Y&Item=" & item & "'><img src='images/btn_review.png' alt='商品レビューをもっと見る' class='opover'></a></li>" & vbNewLine
'	wHTML = wHTML & "  </tr>" & vbNewLine
'	wHTML = wHTML & "</table>" & vbNewLine
	'2012/07/10 GV Mod End
end if

wReviewHTML = wHTML

End Function

'========================================================================
'
'	Function	出荷通知からのリンクで受注番号が渡された場合は、UserID取り出し
'
'========================================================================
'
Function GetUserID()

Dim RSv

'---- Select Web受注
wSQL = ""
wSQL = wSQL & "SELECT a.顧客番号"
wSQL = wSQL & "  FROM Web受注 a WITH (NOLOCK)"
wSQL = wSQL & " WHERE a.受注番号 = " & OrderNo

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

if RSv.EOF = false then
	wUserID = RSv("顧客番号")
end if

RSv.Close

End Function

'========================================================================
'
'	Function	Check Review 該当顧客が購入実績があり、この商品のレビューを投稿しているかどうかチェック
'
'========================================================================
'
Function CheckReview()

Dim RSv

'---- Select 商品レビュー
wSQL = ""
wSQL = wSQL & "SELECT a.購入回数 "
wSQL = wSQL & "     , a.ハンドルネーム "
wSQL = wSQL & "     , c.顧客都道府県 "
wSQL = wSQL & "     , b.ID "
' 2012/01/18 GV Mod Start ( WITH (NOLOCK)付加 )
'wSQL = wSQL & "  FROM Web顧客 a LEFT JOIN 商品レビュー b"
'wSQL = wSQL & "                        ON b.顧客番号 = a.顧客番号"
'wSQL = wSQL & "                       AND b.メーカーコード = '" & maker_cd & "'"
'wSQL = wSQL & "                       AND b.商品コード = '" & product_cd & "'"
'wSQL = wSQL & "     , Web顧客住所 c"
'wSQL = wSQL & " WHERE c.顧客番号 = a.顧客番号"
'wSQL = wSQL & "   AND c.住所連番 = 1"
'wSQL = wSQL & "   AND a.顧客番号 = " & wUserID
wSQL = wSQL & "FROM "
wSQL = wSQL & "    Web顧客                   a WITH (NOLOCK) "
wSQL = wSQL & "      LEFT JOIN  商品レビュー b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.顧客番号       = a.顧客番号"
wSQL = wSQL & "           AND b.メーカーコード = '" & maker_cd & "' "
wSQL = wSQL & "           AND b.商品コード     = '" & Replace(product_cd, "'", "''") & "' "	' 2012/01/23 GV Mod (コード内にシングルクオーテーションが存在した場合の対応)
wSQL = wSQL & "      INNER JOIN Web顧客住所  c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.顧客番号       = a.顧客番号 "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        c.住所連番 = 1 "
wSQL = wSQL & "    AND a.顧客番号 = " & wUserID
' 2012/01/18 GV Mod End

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

'@@@@@response.write(wSQL)

wPrefecture = ""
wHandleName = ""

if RSv.EOF = false then
	wPrefecture = RSv("顧客都道府県")

	if IsNull(RSv("ハンドルネーム")) = false then
		wHandleName = Trim(RSv("ハンドルネーム"))
	end if

	if RSv("購入回数") > 0 AND IsNull(RSv("ID")) = true then
		wCanWriteReviewFl = "Y"
	end if
end if

RSv.Close

End Function

'========================================================================
'
'	Function	メーカー/商品HTML作成
'
'========================================================================
'
Function CreateProductHTML()

Dim vInventoryCd
Dim vInventoryImage
Dim vFreeShippingHTML				' 2011/02/18 GV Add
Dim v_price					' 2012/07/10 GV Add
Dim v_exprice					' 2012/07/10 GV Add
Dim vUrl					' 2012/07/20 GV Add
Dim RSv

wHTML = ""

'---- タイトル
'2012/07/10 GV Mod Start
'wHTML = wHTML & "<table width='188' border='0' cellspacing='0' cellpadding='0' id='Shop_right_Detail'>" & vbNewLine

'wHTML = wHTML & "  <tr>" & vbNewLine
'wHTML = wHTML & "    <td class='head'>メーカー/商品名</td>" & vbNewLine
'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & "  <tr>" & vbNewLine

'wHTML = wHTML & "    <td class='base' itemprop='offerDetails' itemscope itemtype='http://data-vocabulary.org/Offer'>" & vbNewLine   '2011/11/22 an mod
'wHTML = wHTML & "      <table width='180' border='0' cellspacing='0' cellpadding='0' class='ProductDetail'>" & vbNewLine
wHTML = wHTML & "<div id='detail_side_inner01'><div id='detail_side_inner02'>" & vbNewLine

wHTML = wHTML & "  <div id='detail_pp' itemprop='offerDetails' itemscope itemtype='http://data-vocabulary.org/Offer'>" & vbNewLine
'2012/07/10 GV End

'---- メーカー名、商品名、カテゴリー名
'2012/07/10 GV Mod Start
'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td align='left'><a href='SearchList.asp?i_type=m&s_maker_cd=" & RS("メーカーコード") & "' class='link'>" & RS("メーカー名") & " (" & RS("メーカー名カナ") & ")</a></td>" & vbNewLine
'wHTML = wHTML & "        </tr>" & vbNewLine

'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td align='left'><span itemprop='itemreviewed'>" & RS("商品名") & "</span></td>" & vbNewLine   '2011/11/22 an mod
'wHTML = wHTML & "        </tr>" & vbNewLine

'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td align='left'><a href='SearchList.asp?i_type=c&s_category_cd=" & RS("カテゴリーコード") & "' class='link'><span itemprop='category'>" & RS("カテゴリー名") & "</span></a></td>" & vbNewLine   '2011/11/22 an mod
'wHTML = wHTML & "        </tr>" & vbNewLine

' 2011/02/18 GV Add Start
'---- 送料完全無料商品の場合に挿入出力する 【送料無料キャンペーン中】 を生成
'vFreeShippingHTML = ""
'If wFreeShippingFlag = "Y" Then
'	vFreeShippingHTML = "<br/><strong class='freeshipping'>【送料無料キャンペーン中】<span>※沖縄・離島を除く</span></strong>"	'2011/06/15 if-web mod
'End If
' 2011/02/18 GV Add End
' 2012/08/01 ok Mod Start
'wHTML = wHTML & "    <h3 class='item_name'>" & RS("メーカー名") & " (" & RS("メーカー名カナ") & ")<br>" & RS("商品名") & "<br>" & RS("カテゴリー名") & "</h3>" & vbNewLine
wHTML = wHTML & "    <h3 class='item_name'><a href='SearchList.asp?i_type=m&s_maker_cd=" & RS("メーカーコード") & "'>" & RS("メーカー名") & " (" & RS("メーカー名カナ") & ")</a><br>" & RS("商品名") & "<br><a href ='SearchList.asp?i_type=c&s_category_cd=" & RS("カテゴリーコード") & "'>" & RS("カテゴリー名") & "</a></h3>" & vbNewLine
' 2012/08/01 ok Mod End
wHTML = wHTML & "    <p>商品ID:" & RS("商品ID") & "</p>" & vbNewLine
wHTML = wHTML & "    <ul class='icon_list'>" & vbNewLine
If wFreeShippingFlag = "Y" Then
	wHTML = wHTML & "    <li><img src='images/icon_free.gif' alt='送料無料'></li>" & vbNewLine
End If
'2012/07/10 GV Mod End
'2012/07/10 GV Add Start
'---- プライスダウンの場合に挿入
If isNULL(RS("前回単価変更日")) = False Then
	If DateAdd("d", 60, RS("前回単価変更日")) >= Date() AND RS("前回販売単価") > RS("販売単価") AND RS("前回販売単価") <> 0 Then
	wHTML = wHTML & "    <li><img src='images/icon_discount.gif' alt='値下げしました'></li>" & vbNewLine
	End If
End If
wHTML = wHTML & "    </ul>" & vbNewLine
'2012/07/10 GV Add End
'---- 販売単価
v_price = calcPrice(RS("販売単価"), wSalesTaxRate)
v_exprice = calcPrice(RS("前回販売単価"), wSalesTaxRate)
'1行目の表示（ASK商品ではない値下げ品の旧価格）
If RS("ASK商品フラグ") <> "Y" Then
	If RS("B品フラグ") = "Y" OR (RS("個数限定数量") > RS("個数限定受注済数量") AND RS("個数限定数量") > 0) OR ( isNULL(RS("前回単価変更日")) = False AND DateAdd("d", 60, RS("前回単価変更日")) >= Date() AND RS("前回販売単価") > RS("販売単価") AND RS("前回販売単価") <> 0) Then
		'値下げ品の旧価格を表示
		If isNULL(RS("前回単価変更日")) = False AND DateAdd("d", 60, RS("前回単価変更日")) >= Date() AND RS("前回販売単価") > RS("販売単価") Then
			wHTML = wHTML & "<p class='cancel'>" & FormatNumber(v_exprice,0) & "円</p>" & vbNewLine
		'B品、限定品は販売価格を旧価格として表示
		Else
			wHTML = wHTML & "<p class='cancel'>" & FormatNumber(v_price,0) & "円</p>" & vbNewLine
		End If
	End If
End If
'---- 販売単価
wPrice = calcPrice(RS("販売単価"), wSalesTaxRate)
wEAPriceExcTax = FormatNumber(RS("販売単価"),0)

'wHTML = wHTML & "        <tr>" & vbNewLine	'2012/07/10 GV Del

if RS("ASK商品フラグ") = "Y" then
	wTwPriceData = "ASK"
'2011/10/19 hn mod s
'	wHTML = wHTML & "          <td>特価：<a href='JavaScript:void(0);' onClick=""askWin=window.open('AskPrice.asp?MakerName=" & Server.URLEncode(RS("メーカー名")) & "&ProductName=" & Server.URLEncode(wProductName) & "&Price=" & wPrice & "' ,'ask', 'width=250 height=80 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');"" class='link'>ASK</a>" & vFreeShippingHTML & "</td>" & vbNewLine
	if RS("B品フラグ") = "Y" OR (RS("個数限定数量") > RS("個数限定受注済数量") AND RS("個数限定数量") > 0) then

		if RS("B品フラグ") = "Y" then
			wPrice = calcPrice(RS("B品単価"), wSalesTaxRate)
			wEAPriceExcTax = FormatNumber(RS("B品単価"),0)
			'2012/07/10 GV Mod Start
'			wHTML = wHTML & "          <td>わけあり品特価：<a class='tip'>ASK<span>" & FormatNumber(wPrice,0) & "円(税込)</span></a></td>" & vbNewLine
'2014/03/19 GV mod start --->
'			wHTML = wHTML & "          <p class='price'>わけあり品特価：<a class='tip'>ASK<span>" & FormatNumber(wPrice,0) & "円(税込)</span></a></p>" & vbNewLine
			wHTML = wHTML & "          <p class='price'>わけあり品特価：<a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RS("B品単価"),0) & "円(税抜)</span>"
			wHTML = wHTML & "<span class='inc-tax'>(税込&nbsp;" & FormatNumber(wPrice,0) & "円)</span></a></p>" & vbNewLine
'2014/03/19 GV mod end <-----
			'2012/07/10 GV Mod End
			wTwPriceLabel = "わけあり品特価"
		end if

		if (RS("個数限定数量") > RS("個数限定受注済数量") AND RS("個数限定数量") > 0) then
			wPrice = calcPrice(RS("個数限定単価"), wSalesTaxRate)
			wEAPriceExcTax = FormatNumber(RS("個数限定単価"),0)
			'2012/07/10 GV Mod
'			wHTML = wHTML & "          <td>限定特価：<a class='tip'>ASK<span>" & FormatNumber(wPrice,0) & "円(税込)</span></a></td>" & vbNewLine
'2014/03/19 GV mod start --->
'			wHTML = wHTML & "          <p class='price'>限定特価：<a class='tip'>ASK<span>" & FormatNumber(wPrice,0) & "円(税込)</span></a></p>" & vbNewLine
			wHTML = wHTML & "          <p class='price'>限定特価：<a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RS("個数限定単価"),0) & "円(税抜)</span>"
			wHTML = wHTML & "<span class='inc-tax'>(税込&nbsp;" & FormatNumber(wPrice,0) & "円)</span></a></p>" & vbNewLine
'2014/03/19 GV mod end <-----
			'2012/07/10 GV Mod
			wTwPriceLabel = "限定特価"
		end if
	else
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "          <td>特価：<a class='tip'>ASK<span>" & FormatNumber(wPrice,0) & "円(税込)</span></a></td>" & vbNewLine
'2014/03/19 GV start ---->
'		wHTML = wHTML & "          <p class='price'>特価：<a class='tip'>ASK<span>" & FormatNumber(wPrice,0) & "円(税込)</span></a></p>" & vbNewLine
		wHTML = wHTML & "          <p class='price'>特価：<a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RS("販売単価"),0) & "円(税抜)</span>"
		wHTML = wHTML & "<span class='inc-tax'>(税込&nbsp;" & FormatNumber(wPrice,0) & "円)</span></a></p>" & vbNewLine
'2014/03/19 GV end <-----
		'2012/07/10 GV Mod End
		wTwPriceLabel = "衝撃特価"
	end if

else

	if RS("B品フラグ") = "Y" OR (RS("個数限定数量") > RS("個数限定受注済数量") AND RS("個数限定数量") > 0) then
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "            <td><div class='price_table'><del>" & FormatNumber(wPrice,0) & "円(税込)</del><br>" & vbNewLine
'		wHTML = wHTML & "            <p class='price'><strong class='red_large'><meta itemprop='currency' content='JPY' /><span itemprop='price'>"
		wHTML = wHTML & "            <p class='price'><strong class='red_large'>"
		'2012/07/10 GV Mod End

		'---- B品特価
		if RS("B品フラグ") = "Y" then
			wPrice = calcPrice(RS("B品単価"), wSalesTaxRate)
			wEAPriceExcTax = FormatNumber(RS("B品単価"),0)

'2014/03/19 GV start ---->
'			wHTML = wHTML & FormatNumber(wPrice,0) & "円</span></strong><span>(税込)</span></p>" & vbNewLine
			wHTML = wHTML & FormatNumber(RS("B品単価"),0) & "円</strong><span>(税抜)</span></p>" & vbNewLine
			wHTML = wHTML & "<p>(税込&nbsp;<meta itemprop='currency' content='JPY' /><span itemprop='price'>" & FormatNumber(wPrice,0) & "円</span>)</p>" & vbNewLine
'2014/03/19 GV end <-----
' 2011/02/18 GV Mod Start
'			wHTML = wHTML & "            <span class='price'>" & FormatNumber(wPrice,0) & "円</span><span class='tax'>(税込)</span><br><b>わけあり品特価</b></div></td>" & vbNewLine '2010/01/26 an 修正 2010/02/06 if-web 修正 2010/02/22 st 修正
			'2012/07/10 GV Mod Start
'			wHTML = wHTML & "            <meta itemprop='currency' content='JPY' /><span class='price' itemprop='price'>" & FormatNumber(wPrice,0) & "円</span><span class='tax'>(税込)</span><br><b>わけあり品特価</b>" & vFreeShippingHTML & "</div></td>" & vbNewLine
			wHTML = wHTML & "            <p class='deals'>わけあり品特価</p>" & vbNewLine
			'2012/07/10 GV Mod End
			wTwPriceLabel = "わけあり品特価"
' 2011/02/18 GV Mod End   '2011/11/22 an mod

		else
		'---- 個数限定単価
			wPrice = calcPrice(RS("個数限定単価"), wSalesTaxRate)
			wEAPriceExcTax = FormatNumber(RS("個数限定単価"),0)

'2014/03/19 GV start ---->
'			wHTML = wHTML & FormatNumber(wPrice,0) & "円</span></strong><span>(税込)</span></p>" & vbNewLine
			wHTML = wHTML & FormatNumber(RS("個数限定単価"),0) & "円</strong><span>(税抜)</span></p>" & vbNewLine
			wHTML = wHTML & "<p>(税込&nbsp;<meta itemprop='currency' content='JPY' /><span itemprop='price'>" & FormatNumber(wPrice,0) & "円</span>)</p>" & vbNewLine
'2014/03/19 GV end <-----
' 2011/02/18 GV Mod Start
'			wHTML = wHTML & "            <span class='price'>" & FormatNumber(wPrice,0) & "円</span><span class='tax'>(税込)</span><br><b>限定特価</b></div></td>" & vbNewLine
			'2012/07/10 GV Mod Start
'			wHTML = wHTML & "            <meta itemprop='currency' content='JPY' /><span class='price' itemprop='price'>" & FormatNumber(wPrice,0) & "円</span><span class='tax'>(税込)</span><br><b>限定特価</b>" & vFreeShippingHTML & "</div></td>" & vbNewLine
			wHTML = wHTML & "           <p class='deals'>限定特価</p>" & vbNewLine
			'2012/07/10 GV Mod End
			wTwPriceLabel = "限定特価"
' 2011/02/18 GV Mod End   '2011/11/22 an mod
		end if
	else
' 2011/02/18 GV Mod Start
'		wHTML = wHTML & "            <td><div class='price_table'>特価：<span class='price'>" & FormatNumber(wPrice,0) & "円</span><span class='tax'>(税込)</span></div></td>" & vbNewLine '2010/02/06 if-web 修正
		'2012/07/12 GV Mod Start
'		wHTML = wHTML & "            <td><div class='price_table'>特価：<meta itemprop='currency' content='JPY' /><span class='price' itemprop='price'>" & FormatNumber(wPrice,0) & "円</span><span class='tax'>(税込)</span>" & vFreeShippingHTML & "</div></td>" & vbNewLine '2010/02/06 if-web 修正

'2014/03/19 GV mod start ---->
'		wHTML = wHTML & "            <p class='price'><strong class='red_large'><meta itemprop='currency' content='JPY' /><span itemprop='price'>" & FormatNumber(wPrice,0) & "円</span></strong><span>(税込)</span></p>" & vbNewLine '2010/02/06 if-web 修正
		wHTML = wHTML & "            <p class='price'><strong class='red_large'>" & FormatNumber(RS("販売単価"),0) & "円</strong><span>(税抜)</span></p>" & vbNewLine
		wHTML = wHTML & "            <p>(税込&nbsp;<meta itemprop='currency' content='JPY' /><span itemprop='price'>" & FormatNumber(wPrice,0) & "円</span>)</p>" & vbNewLine
'2014/03/19 GV mod end <----

		'2012/07/12 GV Mod End
		wTwPriceLabel = "衝撃特価"
' 2011/02/18 GV Mod End   '2011/11/22 an mod
	end if
	wTwPriceData = FormatNumber(wPrice,0) & "円(税込)"
end if

'wHTML = wHTML & "        </tr>" & vbNewLine	'2012/07/12 GV Del

if wIroKikakuSelectedFl = true then
	'----- 在庫状況
	vInventoryCd = GetInventoryStatus(RS("メーカーコード"),RS("商品コード"),RS("色"),RS("規格"),RS("引当可能数量"),RS("希少数量"),RS("セット商品フラグ"),RS("メーカー直送取寄区分"),RS("引当可能入荷予定日"),wProdTermFl)

	'---- 在庫状況、色を最終セット
	call GetInventoryStatus2(RS("引当可能数量"), RS("Web納期非表示フラグ"), RS("入荷予定未定フラグ"), RS("廃番日"), RS("B品フラグ"), RS("B品引当可能数量"), RS("個数限定数量"), RS("個数限定受注済数量"), wProdTermFl, vInventoryCd, vInventoryImage)

	'----
'2010/11/04 GV Mod Start
'	wHTML = wHTML & "        <tr>" & vbNewLine
'	wHTML = wHTML & "          <td>在庫状況：<img src='images/" & vInventoryImage & "' width='10' height='10'> " & vInventoryCd & "</td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine
	'---- 完売御礼でない場合のみ、在庫状況を表示
	if (IsNull(RS("取扱中止日")) = false) OR (IsNull(RS("完売日")) = false) OR (RS("B品フラグ") = "Y" AND RS("B品引当可能数量") <= 0) OR (IsNull(RS("廃番日")) = false AND RS("引当可能数量") <= 0 AND RS("発注数量") <= 0 AND wIroKikakuSelectedFl = true) then	'2011/06/09 hn mod
		wTwInventoryData = "完売しました"
	else
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "        <tr>" & vbNewLine
'		wHTML = wHTML & "          <td>在庫状況：<img src='images/" & vInventoryImage & "' width='10' height='10'> "    '2011/11/22 an mod s
		wHTML = wHTML & "          <p class='stock'><img src='images/" & vInventoryImage & "' alt='" & vInventoryCd & "'> "    '2011/11/22 an mod s
		'2012/07/10 GV Mod End

		if vInventoryCd = "在庫あり" OR vInventoryCd = "在庫僅少" OR vInventoryCd = "在庫限り" OR Left(vInventoryCd, 2) = "限定" then
			wHTML = wHTML & "<span itemprop='availability' content='in_stock'>" & vInventoryCd & "</span>"
		else
			wHTML = wHTML & vInventoryCd
		end if                                                                                                          '2011/11/22 an mod e

		wTwInventoryData = vInventoryCd

		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "          </td>" & vbNewLine
'		wHTML = wHTML & "        </tr>" & vbNewLine
		wHTML = wHTML & "          </p>" & vbNewLine
		'2012/07/10 GV Mod End
	end if
'2010/11/04 GV Mod End
end if


'2012/07/10 GV Del Start
'---- 送料について
'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td><a href='../guide/kaimono.asp#souryou' class='link'>送料について</a></td>" & vbNewLine
'wHTML = wHTML & "        </tr>" & vbNewLine
'2012/07/10 GV Del End

'---- カートボタン
'2012/07/10 GV Mod Start
'wHTML = wHTML & "        <tr>" & vbNewLine
wHTML = wHTML & "          <form name='f_data' method='post' action='OrderPreInsert.asp' onSubmit='return order_onClick(this);'>" & vbNewLine
'wHTML = wHTML & "          <td>" & vbNewLine
'2012/07/10 GV Mod End
wHTML = wHTML & wIroKikakuCombo

if (IsNull(RS("取扱中止日")) = false) OR (IsNull(RS("完売日")) = false) OR (RS("B品フラグ") = "Y" AND RS("B品引当可能数量") <= 0) OR (IsNull(RS("廃番日")) = false AND RS("引当可能数量") <= 0 AND RS("発注数量") <= 0 AND wIroKikakuSelectedFl = true) then	'2011/06/09 hn mod
'2012/07/10 GV Mod Start
'  wHTML = wHTML & "<img src='images/Kanbai2.jpg'>" & vbNewLine
  wHTML = wHTML & "<p class='sold'><img src='images/icon_sold.png' alt='完売しました'></p>" & vbNewLine
'2012/07/10 GV Mod End

else
	wHTML = wHTML & "            <div id='cart'>" & vbNewLine
	wHTML = wHTML & "                <span>個数<input type='text' name='qt' value='1'></span><input type='image' src='images/btn_cart_productdetail.png' alt='カートに入れる' class='opover'>" & vbNewLine
	wHTML = wHTML & "                <input type='hidden' name='Item' value='" & RS("メーカーコード") & "^" & RS("商品コード") & "^" & Trim(RS("色")) & "^" & Trim(RS("規格")) & "'>" & vbNewLine
	wHTML = wHTML & "            </div>" & vbNewLine
end if

'wHTML = wHTML & "          </td>" & vbNewLine	'2012/07/10 GV Del

'---- 一緒に購入するチェック時の追加商品登録用(メーカーコード^商品コード^色^規格 ,区切りで格納）
wHTML = wHTML & "          <input type='hidden' name='AdditionalItem' value=''>" & vbNewLine

wHTML = wHTML & "          </form>" & vbNewLine
'wHTML = wHTML & "        </tr>" & vbNewLine	'2012/07/10 GV Del

vUrl = ""
vUrl = vUrl & RS("メーカーコード") & "^" & RS("商品コード") & "^" & Trim(RS("色")) & "^" & Trim(RS("規格"))

'----- ウィッシュリスト
if wProdTermFl = "Y" OR wIroKikakuSelectedFl = false then
else
	if wIroKikakuSelectedFl = true then
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "        <tr>"
'		wHTML = wHTML & "          <td align='right'><a href='WishListAdd.asp?Item=" & Server.URLEncode(RS("メーカーコード") & "^" & RS("商品コード") & "^" & RS("色") & "^" & RS("規格")) & "' class='link'>ウィッシュリスト</a></td>" & vbNewLine
'		wHTML = wHTML & "        </tr>" & vbNewLine
	
		wHTML = wHTML & "          <p class='btn_wish'><a href='"
		if wUserID = "" Then
			wHTML = wHTML & g_HTTPS & "shop/LoginCheck.asp?RtnURL=" & g_HTTP & "shop/WishListAdd.asp?Item=" & Server.URLEncode(vUrl)
		Else
			wHTML = wHTML & "WishListAdd.asp?Item=" & Server.URLEncode(vUrl)
		End If
			wHTML = wHTML & "' ><img src='images/btn_wish.png' alt='ウィッシュリストに追加' class='opover' width='200' height='25'></a></p>" & vbNewLine
		'2012/07/10 GV Mod End
	end if
end if

'----- 商品ID
'2012/07/10 GV Del Start
'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td align='right'>商品ID:" & RS("商品ID") & "</td>" & vbNewLine
'wHTML = wHTML & "        </tr>" & vbNewLine
'2012/07/10 GV Del End

'---- Twitter,Facebook
wHTML = wHTML & "          <ul class='sns'>" & vbNewLine
wHTML = wHTML & "            <li><a href='http://twitter.com/share' class='twitter-share-button' data-via='soundhouse_jp' data-lang='ja'>ツイート</a></li>" & vbNewLine
'wHTML = wHTML & "            <li><div class='fb-like' data-href='http://www.soundhouse.co.jp/shop/ProductDetail.asp?Item=" & Server.URLEncode(maker_cd & "^" & product_cd & "^" & iro & "^" & kikaku) & "' data-send='false' data-layout='button_count' data-width='145' data-show-faces='false'></div></li>" & vbNewLine
wHTML = wHTML & "            <li><iframe src='//www.facebook.com/plugins/like.php?href=http%3A%2F%2Fwww.soundhouse.co.jp%2Fshop%2FProductDetail.asp%3FItem%3D" & Server.URLEncode(vUrl) & "&amp;send=false&amp;layout=button_count&amp;width=100&amp;show_faces=false&amp;action=like&amp;colorscheme=light&amp;font&amp;height=21&amp;appId=191447484218062' scrolling='no' frameborder='0' style='border:none; overflow:hidden; width:120px; height:21px;' allowTransparency='true'></iframe></li>" & vbNewLine
wHTML = wHTML & "          </ul>" & vbNewLine

'2012/07/10 GV Add Start
'---- 評価
wHTML = wHTML & wHyoukaHTML

wHTML = wHTML & "         <ul class='info'>" & vbNewLine
'---- この商品の問合せ
wHTML = wHTML & "            <li><a href='" & g_HTTPS & "shop/Inquiry.asp?MakerNm=" & Server.URLEncode(RS("メーカー名")) & "&ProductCd=" & Server.URLEncode(RS("商品コード")) & "&CategoryNm=" & Server.URLEncode(RS("カテゴリー名")) & "'>この商品へのお問い合わせ</a></li>" & vbNewLine

'---- 送料について
wHTML = wHTML & "            <li><a href='../guide/kaimono.asp#souryou'>送料について</a></li>" & vbNewLine
'2012/07/10 GV Add End

'---- 友達に勧める
if wUserID <> "" AND wProdTermFl <> "Y" then
	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "        <tr>" & vbNewLine
'	wHTML = wHTML & "          <td colspan='2' align='center' height='26'><a href='TellaFriend.asp?Item=" & Server.URLEncode(RS("メーカーコード") & "^" & RS("商品コード")) & "' class='link'><img src='images/TomodachiNiSusumeru.gif' border='0'></a></td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine
	wHTML = wHTML & "            <li><a href='TellaFriend.asp?Item=" & Server.URLEncode(RS("メーカーコード") & "^" & RS("商品コード")) & "'>友達にすすめる</a></li>" & vbNewLine
	'2012/07/10 GV Mod End
end if

'---- この商品の問合せ
'2012/07/10 GV Del Start
'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td colspan='2' align='center' height='26'><a href='" & g_HTTPS & "shop/Inquiry.asp?MakerNm=" & Server.URLEncode(RS("メーカー名")) & "&ProductCd=" & Server.URLEncode(RS("商品コード")) & "&CategoryNm=" & Server.URLEncode(RS("カテゴリー名")) & "' class='link'><img src='images/ShouhinNoToiawase.gif' border='0'></a></td>" & vbNewLine
'wHTML = wHTML & "        </tr>" & vbNewLine
'2012/07/10 GV Del End

'---- Twitterリンク  2010/09/27 an add s
'2012/07/10 GV Del Start
'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td colspan='2' align='center' height='26'>" & vbNewLine
'wHTML = wHTML & "            <ul class='smbtn'>" & vbNewLine
'wHTML = wHTML & "              <li><a href='http://twitter.com/share' class='twitter-share-button' data-count='horizontal' data-via='soundhouse_jp' data-lang='ja'>Tweet</a></li>" & vbNewLine
'wHTML = wHTML & "              <li><a name='fb_share'>シェアする</a></li>" & vbNewLine
'wHTML = wHTML & "            </ul>" & vbNewLine
'wHTML = wHTML & "          </td>" & vbNewLine
'wHTML = wHTML & "        </tr>" & vbNewLine    '2010/09/27 an add e

'wHTML = wHTML & "      </table>" & vbNewLine

'wHTML = wHTML & "    </td>" & vbNewLine
'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & "</table>" & vbNewLine
'2012/07/10 GV Del End

'2012/07/10 GV Add Start
wHTML = wHTML & "    </ul>" & vbNewLine
wHTML = wHTML & "  </div>" & vbNewLine
wHTML = wHTML & "</div></div>" & vbNewLine
'2012/07/10 GV Add End

wProductHTML = wHTML

'2013/05/17 GV #1505 add start
wEAInventoryData = vInventoryCd
wEAPrice         = wPrice
'2013/05/17 GV #1505 add end

End Function

'========================================================================
'
'	Function	レコメンド結果	2009/12/17
'
'========================================================================
'
Function CreateRecommendHTML()

'2013/05/17 GV #1505 add start
wRecommendJS = fEAgency_CreateRecommendJS(wEAProductDetailData, wEAIroKikakuData)
'2013/05/17 GV #1505 add end

'2013/08/07 if-web del s
'Dim RSv
'Dim vPrice
'
'
'---- レコメンド結果
'wSQL = ""
'wSQL = wSQL & "SELECT DISTINCT TOP 5"
'wSQL = wSQL & "       a.メーカーコード"
'wSQL = wSQL & "     , a.商品コード"
'wSQL = wSQL & "     , a.商品名"
'wSQL = wSQL & "     , a.商品画像ファイル名_小"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN a.B品フラグ = 'Y' THEN a.B品単価"    '2010/11/12 an add
'wSQL = wSQL & "         WHEN a.個数限定数量 > a.個数限定受注済数量 THEN a.個数限定単価"
'wSQL = wSQL & "         ELSE a.販売単価"
'wSQL = wSQL & "       END AS 実販売単価"    '2010/11/12 an mod
'wSQL = wSQL & "     , a.ASK商品フラグ"
'wSQL = wSQL & "     , b.メーカー名"
'wSQL = wSQL & "     , e.類似度"
'wSQL = wSQL & "  FROM Web商品 a WITH (NOLOCK)"
'wSQL = wSQL & "     , メーカー b WITH (NOLOCK)"
'wSQL = wSQL & "     , Web色規格別在庫 d WITH (NOLOCK)"
'
'if wUserID = "" then
'	wSQL = wSQL & "     , レコメンド結果アクセス e WITH (NOLOCK)"
'else
'	wSQL = wSQL & "     , レコメンド結果購買 e WITH (NOLOCK)"
'end if
'
'wSQL = wSQL & " WHERE a.メーカーコード = e.レコメンドメーカーコード"
'wSQL = wSQL & "   AND a.商品コード = e.レコメンド商品コード"
'wSQL = wSQL & "   AND d.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND d.商品コード = a.商品コード"
'wSQL = wSQL & "   AND b.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND a.Web商品フラグ = 'Y'"
'wSQL = wSQL & "   AND a.取扱中止日 IS NULL"
'wSQL = wSQL & "   AND ((a.廃番日 IS NULL) OR (a.廃番日 IS NOT NULL AND d.引当可能数量 > 0 AND d.発注数量 > 0))"	'2011/06/09 hn mod
'wSQL = wSQL & "   AND e.メーカーコード = '" & maker_cd & "'"
'wSQL = wSQL & "   AND e.商品コード = '" & Replace(product_cd, "'", "''") & "'"	' 2012/01/23 GV Mod (コード内にシングルクオーテーションが存在した場合の対応)
'wSQL = wSQL & " ORDER BY"
'wSQL = wSQL & "       e.類似度 DESC"
'
'@@@@@@response.write(wSQL)
'
'Set RSv = Server.CreateObject("ADODB.Recordset")
'RSv.Open wSQL, Connection, adOpenStatic
'
'wHTML = ""
'
'if RSv.EOF = true then
'	RSv.close
'	exit function
'end if
'
'----- レコメンド商品HTML編集
'2012/07/10 GV Del Start
'wHTML = wHTML & "<table id=Shop_right_relation cellSpacing=0 cellPadding=0 width=188 border=0>" & vbNewLine
'wHTML = wHTML & "  <tr>" & vbNewLine
'2012/07/10 GV Del End
'
'if wUserID = "" then
'	wHTML = wHTML & "    <td style='padding:5px; border:#999999 solid 1px;' bgcolor='#FFCC66'>このアイテムを見た人は<br>こんなアイテムも見ています。</td>" & vbNewLine	'2012/07/10 GV Del
'else
'	wHTML = wHTML & "    <td style='padding:5px; border:#999999 solid 1px;' bgcolor='#CCFF00'>このアイテムを買った人は<br>こんなアイテムも買っています。</td>" & vbNewLine
'end if
'2012/07/10 GV Add Start
'wHTML = wHTML & "<div class='detail_side_inner01'><div class='detail_side_inner02'>" & vbNewLine
'wHTML = wHTML & "  <div class='detail_side_inner_box'>" & vbNewLine
'wHTML = wHTML & "    <!--このアイテムを見た人は -->" & vbNewLine
'wHTML = wHTML & "    <h4 class='detail_sub'>このアイテムを見た人は<br>こんなアイテムも見ています。</h4>" & vbNewLine
'wHTML = wHTML & "    <ul class='check_item'>" & vbNewLine
'2012/07/10 GV Add End
'
'wHTML = wHTML & "  </tr>" & vbNewLine	'2012/07/10 GV Del
'
'Do Until RSv.EOF = true
'	vPrice = calcPrice(RSv("実販売単価"), wSalesTaxRate) '2010/11/12 an mod 販売単価→実販売単価
'
'	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "  <tr>" & vbNewLine
'	wHTML = wHTML & "    <td class=base align=middle>" & vbNewLine
'	wHTML = wHTML & "      <table id=Shop_right_product cellSpacing=0 cellPadding=0 width=180 border=0>" & vbNewLine
'	wHTML = wHTML & "        <tr>" & vbNewLine
'	wHTML = wHTML & "          <td><a href='ProductDetail.asp?Item=" & Server.URLEncode(RSv("メーカーコード") & "^" & RSv("商品コード")) & "'><img src='prod_img/" & RSv("商品画像ファイル名_小") & "' width='170' height='85' border='0'></a></td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine
'	wHTML = wHTML & "        <tr>" & vbNewLine
'	wHTML = wHTML & "          <td>" & RSv("メーカー名") & " " & RSv("商品名") & "</td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine
'	wHTML = wHTML & "        <tr>" & vbNewLine
'	wHTML = wHTML & "      <li>" & vbNewLine
'
'	wHTML = wHTML & "        <p><a href='ProductDetail.asp?Item=" & Server.URLEncode(RSv("メーカーコード") & "^" & RSv("商品コード")) & "'>"
'	If RSv("商品画像ファイル名_小") <> "" Then
'		wHTML = wHTML & "<img src='prod_img/" & RSv("商品画像ファイル名_小") & "' alt='" & Replace(RSv("メーカー名") & " / " & RSv("商品名"),"'","&#39;") & "' class='opover'>"
'	End If
'	wHTML = wHTML & RSv("メーカー名") & " / " & RSv("商品名") & "</a></p>" & vbNewLine
'	'2012/07/10 GV Mod End
'
'	if RSv("ASK商品フラグ") <> "Y" then  '2010/04/06 an changed start
'		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "          <td>" & FormatNumber(vPrice, 0) & "円(税込)</td>" & vbNewLine
'		wHTML = wHTML & "        <p>" & FormatNumber(vPrice, 0) & "円(税込)</p>" & vbNewLine
'		'2012/07/10 GV Mod End
'	else
'2011/10/19 hn mod s
'		wHTML = wHTML & "          <td><a href='JavaScript:void(0);' onClick=""askWin=window.open('AskPrice.asp?MakerName=" & Server.URLEncode(RSv("メーカー名")) & "&ProductName=" & Server.URLEncode(RSv("商品名")) & "&Price=" & vPrice & "' ,'ask', 'width=250 height=80 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');"" class='link'><strong>ASK</b></strong></td>" & vbNewLine
'
'		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "          <td><a class='tip'>ASK<span>" & FormatNumber(vPrice, 0) & "円(税込)</span></a></td>" & vbNewLine
'		wHTML = wHTML & "        <p><a class='tip'>ASK<span>" & FormatNumber(vPrice, 0) & "円(税込)</span></a></p>" & vbNewLine
'		'2012/07/10 GV Mod End
'2011/10/19 hn mod e
'
'	end if       '2010/04/06 an changed end,  2010/04/21 an changed
'
'	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "        </tr>" & vbNewLine
'	wHTML = wHTML & "      </table>" & vbNewLine
'	wHTML = wHTML & "    </td>" & vbNewLine
'	wHTML = wHTML & "  </tr>" & vbNewLine
'	wHTML = wHTML & "      </li>" & vbNewLine
'	'2012/07/10 GV Mod End
'
'	RSv.MoveNext
'Loop
'
'2012/07/10 GV Add Start
'wHTML = wHTML & "    </ul>" & vbNewLine
'wHTML = wHTML & "  </div>" & vbNewLine
'wHTML = wHTML & "</div></div>" & vbNewLine
'2012/07/10 GV Add End
'
'wHTML = wHTML & "</table>" & vbNewLine	'2012/07/10 GV Del
'
'wRecommendHTML = wHTML
'
'RSv.close
'2013/08/07 if-web del e

End function

'========================================================================
'
'	Function	レコメンド商品取得  '2012/04/10
'
'========================================================================
Function CreateRecommendBuyHTML()

'2013/05/17 GV #1505
wRecommendBuyJS = fEAgency_CreateRecommendBuyJS(wEAProductDetailData, wEAInventoryData, wEAPrice, wEAPriceExcTax, wEAIroKikakuData)

'2013/08/07 if-web del s
'Dim RSv
'Dim iCnt	'2012/07/10 GV Add
'
'
'---- レコメンド商品取得(類似度が大きい5商品)
'wSQL = ""
'
'1行に表示する件数が5件から4件に変更 2012/07/20 ok Mod
'wSQL = wSQL & "SELECT DISTINCT TOP 5"
'wSQL = wSQL & "SELECT DISTINCT TOP 4"
'wSQL = wSQL & "       a.メーカーコード"
'wSQL = wSQL & "     , a.商品コード"
'wSQL = wSQL & "     , a.商品名"
'wSQL = wSQL & "     , a.商品画像ファイル名_小"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN (a.個数限定数量 > a.個数限定受注済数量 AND a.個数限定数量 > 0) THEN a.個数限定単価"
'wSQL = wSQL & "         ELSE a.販売単価"
'wSQL = wSQL & "       END AS 販売単価"
'wSQL = wSQL & "     , a.ASK商品フラグ"
'wSQL = wSQL & "     , a.カテゴリーコード"
'wSQL = wSQL & "     , b.メーカー名"
'wSQL = wSQL & "     , e.類似度"
'wSQL = wSQL & "  FROM Web商品 a WITH (NOLOCK)"
'wSQL = wSQL & "     , メーカー b WITH (NOLOCK)"
'wSQL = wSQL & "     , Web色規格別在庫 d WITH (NOLOCK)"
'wSQL = wSQL & "     , レコメンド結果購買 e WITH (NOLOCK)"
'wSQL = wSQL & " WHERE a.メーカーコード = e.レコメンドメーカーコード"
'wSQL = wSQL & "   AND a.商品コード = e.レコメンド商品コード"
'wSQL = wSQL & "   AND d.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND d.商品コード = a.商品コード"
'wSQL = wSQL & "   AND b.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND a.Web商品フラグ = 'Y'"
'wSQL = wSQL & "   AND a.取扱中止日 IS NULL"
'wSQL = wSQL & "   AND ((a.廃番日 IS NULL) OR (a.廃番日 IS NOT NULL AND d.引当可能数量 > 0))"
'wSQL = wSQL & "   AND e.メーカーコード = '" & maker_cd & "'"
'wSQL = wSQL & "   AND e.商品コード = '" & Replace(product_cd, "'", "''") & "'"
'wSQL = wSQL & " ORDER BY"
'wSQL = wSQL & "       e.類似度 DESC"
'wSQL = wSQL & "     , a.カテゴリーコード"
'
'@@@@response.write(wSQL)
'
'Set RSv = Server.CreateObject("ADODB.Recordset")
'RSv.Open wSQL, Connection, adOpenStatic
'
'wHTML = ""
'iCnt = 0	'2012/07/10 GV Add
'
'if RSv.EOF = false then
'
'	wHTML = ""
'	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "<h2 id=""recommend_h"">このアイテムを買った人はこんなアイテムも買っています。</h2>" & vbNewLine
'	wHTML = wHTML & "<ul id=""recommend_box"">" & vbNewLine
'	wHTML = wHTML & "<h2 class='detail_title'>このアイテムを買った人はこんなアイテムも買っています</h2>" & vbNewLine
'	'2012/07/10 GV Mod End
'
'	Do Until RSv.EOF = True
'
'		'2012/07/10 GV Add Start
'		if iCnt mod 4 = 0 then
'			wHTML = wHTML & "<ul class='relation other'>" & vbNewLine
'		end if
'		'2012/07/10 GV Add End
'
'		wPrice = calcPrice(RSv("販売単価"), wSalesTaxRate)
'
'		wHTML = wHTML & "  <li>" & vbNewLine	'2012/07/10 GV Add
'		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "    <a href=""ProductDetail.asp?Item=" & RSv("メーカーコード") & "%5E" & RSv("商品コード") & """>"
'		wHTML = wHTML & "    <p><a href='ProductDetail.asp?Item=" & RSv("メーカーコード") & "%5E" & RSv("商品コード") & "'>"
'		'2012/07/10 GV Mod End
'		if RSv("商品画像ファイル名_小") <> "" then
'			'2012/07/10 GV Mod Start
'			wHTML = wHTML & "<img src=""prod_img/" & RSv("商品画像ファイル名_小") & """ alt=""" & RSv("メーカー名") & " " & RSv("商品名") & """>"
'			wHTML = wHTML & "<img src='prod_img/" & RSv("商品画像ファイル名_小") & "' alt='" & Replace(RSv("メーカー名") & " " & RSv("商品名"),"'","&#39;") & "' class='opover'></a></p>"
'			'2012/07/10 GV Mod End
'		else
'			wHTML = wHTML & "<img src=""prod_img/n/nopict-.jpg"" alt="""">"
'		end if
'		wHTML = wHTML & "</a><br>" & vbNewLine	'2012/07/10 GV Del
'		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "    " & RSv("メーカー名") & " " & RSv("商品名") & "<br>" & vbNewLine
'		wHTML = wHTML & "    <p><a href='ProductDetail.asp?Item=" & RSv("メーカーコード") & "%5E" & RSv("商品コード") & "'>"
'		wHTML = wHTML & "    " & RSv("メーカー名") & " / " & RSv("商品名") & "</a></p>" & vbNewLine
'		'2012/07/10 GV Mod End
'
'		wHTML = wHTML & "    <div class='box'>" & vbNewLine	'2012/07/10 GV Add
'		If RSv("ASK商品フラグ") <> "Y" Then
'			'2012/07/10 GV Mod Start
'			wHTML = wHTML & "    " & FormatNumber(wPrice,0) & "円(税込)" & vbNewLine
'			wHTML = wHTML & "      <p>" & FormatNumber(wPrice,0) & "円(税込)</p>" & vbNewLine
'			'2012/07/10 GV Mod End
'		Else
'			'2012/07/10 GV Mod Start
'			wHTML = wHTML & "    <a class='tip'>ASK<span>"& FormatNumber(wPrice,0) & "円(税込)</span></a>" & vbNewLine
'			wHTML = wHTML & "      <p><a class='tip'>ASK<span>"& FormatNumber(wPrice,0) & "円(税込)</span></a></p>" & vbNewLine
'			'2012/07/10 GV Mod End
'
'		End If
'		wHTML = wHTML & "    </div>" & vbNewLine	'2012/07/10 GV Add
'
'		wHTML = wHTML & "  </li>" & vbNewLine		'2012/07/10 GV Add
'
'		RSv.MoveNext
'
'		'2012/07/10 GV Add Start
'		if (iCnt mod 4 = 3) Or (RSv.RecordCount = iCnt+1) then
'			wHTML = wHTML & "</ul>" & vbNewLine
'		end if
'		iCnt = iCnt + 1
'		'2012/07/10 GV Add End
'	Loop
'
'End if
'
'RSv.Close
'
'wRecommendBuyHTML = wHTML
'2013/08/07 if-web del e

End function

'========================================================================
'
'	Function	関連シリーズ商品
'
'========================================================================
'
Function CreateSeriesHTML()

Dim RSv
Dim vPrice
Dim vRecordCount
Dim vCount

'---- 関連シリーズ商品
wSQL = ""
' 2012/01/18 GV Mod Start
'wSQL = wSQL & "SELECT DISTINCT "
'wSQL = wSQL & "       a.メーカーコード"
'wSQL = wSQL & "     , a.商品コード"
'wSQL = wSQL & "     , a.商品名"
'wSQL = wSQL & "     , a.商品画像ファイル名_小"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN a.B品フラグ = 'Y' THEN a.B品単価"   '2010/11/12 an add
'wSQL = wSQL & "         WHEN a.個数限定数量 > a.個数限定受注済数量 THEN a.個数限定単価"
'wSQL = wSQL & "         ELSE a.販売単価"
'wSQL = wSQL & "       END AS 実販売単価"   '2010/11/12 an mod
'wSQL = wSQL & "     , a.ASK商品フラグ"
'wSQL = wSQL & "     , b.メーカー名"
'wSQL = wSQL & "     , c.カテゴリー名"
'wSQL = wSQL & "     , c.表示順"
'wSQL = wSQL & "     , d.色"
'wSQL = wSQL & "     , d.規格"
'wSQL = wSQL & "  FROM Web商品 a WITH (NOLOCK)"
'wSQL = wSQL & "     , メーカー b WITH (NOLOCK)"
'wSQL = wSQL & "     , カテゴリー c WITH (NOLOCK)"
'wSQL = wSQL & "     , Web色規格別在庫 d WITH (NOLOCK)"
'wSQL = wSQL & " WHERE b.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND c.カテゴリーコード = a.カテゴリーコード"
'wSQL = wSQL & "   AND d.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND d.商品コード = a.商品コード"
'wSQL = wSQL & "   AND a.Web商品フラグ = 'Y'"
'wSQL = wSQL & "   AND a.取扱中止日 IS NULL"
'wSQL = wSQL & "   AND ((a.廃番日 IS NULL) OR (a.廃番日 IS NOT NULL AND d.引当可能数量 > 0 AND d.発注数量 > 0))"		'2011/06/09 hn mod
'wSQL = wSQL & "   AND NOT (a.メーカーコード = '" & RS("メーカーコード") & "'"
'wSQL = wSQL & "       AND a.商品コード = '" & RS("商品コード") & "')"
'wSQL = wSQL & "   AND a.シリーズコード = '" & RS("シリーズコード") & "'"
'wSQL = wSQL & " ORDER BY"
'wSQL = wSQL & "       c.表示順"
'wSQL = wSQL & "     , b.メーカー名"
'wSQL = wSQL & "     , a.商品名"
'wSQL = wSQL & "     , d.色"
'wSQL = wSQL & "     , d.規格"
wSQL = wSQL & "SELECT DISTINCT "
wSQL = wSQL & "      a.メーカーコード "
wSQL = wSQL & "    , a.商品コード "
wSQL = wSQL & "    , a.商品名 "
wSQL = wSQL & "    , a.商品画像ファイル名_小 "
wSQL = wSQL & "    , CASE "
wSQL = wSQL & "        WHEN a.B品フラグ = 'Y'                     THEN a.B品単価 "
wSQL = wSQL & "        WHEN a.個数限定数量 > a.個数限定受注済数量 THEN a.個数限定単価 "
wSQL = wSQL & "        ELSE                                            a.販売単価 "
wSQL = wSQL & "      END AS 実販売単価 "
wSQL = wSQL & "    , a.ASK商品フラグ "
wSQL = wSQL & "    , b.メーカー名 "
wSQL = wSQL & "    , c.カテゴリー名 "
wSQL = wSQL & "    , c.表示順 "
wSQL = wSQL & "    , d.色 "
wSQL = wSQL & "    , d.規格 "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    Web商品                      a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN メーカー        b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.メーカーコード   = a.メーカーコード "
wSQL = wSQL & "      INNER JOIN カテゴリー      c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.カテゴリーコード = a.カテゴリーコード "
wSQL = wSQL & "      INNER JOIN Web色規格別在庫 d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.メーカーコード   = a.メーカーコード "
wSQL = wSQL & "           AND d.商品コード       = a.商品コード "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' )   t1 "
wSQL = wSQL & "        ON     a.Web商品フラグ    = t1.ShohinWebY  "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        t1.ShohinWebY IS NOT NULL "
wSQL = wSQL & "    AND a.取扱中止日 IS NULL "
wSQL = wSQL & "    AND (    (a.廃番日 IS NULL) "
wSQL = wSQL & "         OR  (    a.廃番日 IS NOT NULL "
wSQL = wSQL & "              AND d.引当可能数量 > 0 "
wSQL = wSQL & "              AND d.発注数量 > 0)) "
wSQL = wSQL & "    AND NOT  (    a.メーカーコード = '" & RS("メーカーコード") & "' "
wSQL = wSQL & "              AND a.商品コード = '" & Replace(RS("商品コード"), "'", "''") & "') "	' 2012/01/23 GV Mod (コード内にシングルクオーテーションが存在した場合の対応)
wSQL = wSQL & "    AND a.シリーズコード = '" & RS("シリーズコード") & "' "
wSQL = wSQL & "ORDER BY "
wSQL = wSQL & "      c.表示順 "
wSQL = wSQL & "    , b.メーカー名 "
wSQL = wSQL & "    , a.商品名 "
wSQL = wSQL & "    , d.色 "
wSQL = wSQL & "    , d.規格 "
' 2012/01/18 GV Mod End

'@@@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic
vRecordCount = RSv.RecordCount

wHTML = ""
vCount = 0

if RSv.EOF = true then
	RSv.close
	exit function
end if

'2012/07/10 GV Add Start
wHTML = wHTML & "<div id='detail_side'>" & vbNewLine
wHTML = wHTML & "  <div id='detail_side_inner01'><div id='detail_side_inner02'>" & vbNewLine
wHTML = wHTML & "    <div class='detail_side_inner_box'>" & vbNewLine
wHTML = wHTML & "      <!-- 関連シリーズ -->" & vbNewLine
wHTML = wHTML & "      <h4 class='detail_sub'><a href='SearchList.asp'>CLASSIC PRO CPAHWシリーズ</a></h4>" & vbNewLine
wHTML = wHTML & "      <ul class='check_item'>" & vbNewLine
wHTML = wHTML & "        <li>" & vbNewLine
wHTML = wHTML & "          <p><a href='SearchList.asp'><img src='prod_img/f/fender_jbesquireb-.jpg' alt='PLAYTECH / PAUL REED SMITH PRS Guitar Strings' class='opover'></a>移動に便利なキャリングハンドル・キャスター付きの樹脂製ラックケース。重い機材の移動、運搬に最適です。</p>" & vbNewLine
wHTML = wHTML & "        </li>" & vbNewLine
wHTML = wHTML & "      </ul>" & vbNewLine
wHTML = wHTML & "    </div>" & vbNewLine
wHTML = wHTML & "  </div></div>" & vbNewLine
wHTML = wHTML & "</div>" & vbNewLine
'2012/07/10 GV Add End
'----- 関連シリーズ商品HTML編集
wHTML = wHTML & "<table width='188' border='0' cellspacing='0' cellpadding='0' id='Shop_right_relation'>" & vbNewLine
wHTML = wHTML & "  <tr>" & vbNewLine
wHTML = wHTML & "    <td align='left' class='head'>関連シリーズ商品</td>" & vbNewLine
wHTML = wHTML & "  </tr>" & vbNewLine

Do Until (RSv.EOF = true OR vCount >= 5)
	vPrice = calcPrice(RSv("実販売単価"), wSalesTaxRate)   '2010/11/12 an mod 販売単価→実販売単価

  wHTML = wHTML & "  <tr>" & vbNewLine
  wHTML = wHTML & "    <td align='center' class='base'>" & vbNewLine
  wHTML = wHTML & "      <table width='180' border='0' cellpadding='0' cellspacing='0' id='Shop_right_product'>" & vbNewLine
  wHTML = wHTML & "        <tr>" & vbNewLine

  wHTML = wHTML & "          <td><a href='ProductDetail.asp?Item=" & Server.URLEncode(RSv("メーカーコード") & "^" & RSv("商品コード") & "^" & RSv("色") & "^" & RSv("規格")) & "'>"
  If RSv("商品画像ファイル名_小") <> "" Then
    wHTML = wHTML & "<img src='prod_img/" & RSv("商品画像ファイル名_小") & "' width='170' height='85' border='0'>"
  End If
  wHTML = wHTML & "</a></td>" & vbNewLine

  wHTML = wHTML & "        </tr>" & vbNewLine
  wHTML = wHTML & "        <tr>" & vbNewLine
  wHTML = wHTML & "          <td>" & RSv("メーカー名") & " " & RSv("商品名") & "</td>" & vbNewLine
  wHTML = wHTML & "        </tr>" & vbNewLine
  wHTML = wHTML & "        <tr>" & vbNewLine

'2011/10/19 hn add s
	if RSv("ASK商品フラグ") <> "Y" then
'2014/03/19 GV mod start ---->
'		wHTML = wHTML & "          <td>" & FormatNumber(vPrice, 0) & "円(税込)</td>" & vbNewLine
		wHTML = wHTML & "          <td>" & FormatNumber(RSv("実販売単価"), 0) & "円(税抜)</td>" & vbNewLine
		wHTML = wHTML & "          <td><strong>(税込)" & FormatNumber(vPrice, 0) & "円</strong></td>" & vbNewLine
'2014/03/19 GV mod end <-----
	else
'2014/03/19 GV mod start ---->
'		wHTML = wHTML & "          <td><a class='tip'>ASK<span>" & FormatNumber(vPrice, 0) & "円(税込)</span></a></td>" & vbNewLine
		wHTML = wHTML & "          <td><a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RSv("実販売単価"), 0) & "円(税抜)</span>"
		wHTML = wHTML & "<span class='inc-tax'>(税込)" & FormatNumber(vPrice, 0) & "円</span></a></td>" & vbNewLine
'2014/03/19 GV mod end <-----
	end if
'2011/10/19 hn add e

  wHTML = wHTML & "        </tr>" & vbNewLine
  wHTML = wHTML & "      </table>" & vbNewLine
  wHTML = wHTML & "    </td>" & vbNewLine
  wHTML = wHTML & "  </tr>" & vbNewLine

	RSv.MoveNext
	vCount = vCount + 1
Loop

if vRecordCount > vCount then
	wHTML = wHTML & "  <tr>" & vbNewLine
	wHTML = wHTML & "    <td><a href='SearchList.asp?i_type=se&sSeriesCd=" & RS("シリーズコード") & "' class='link'>その他関連シリーズ商品>></a></td>" & vbNewLine
	wHTML = wHTML & "  </tr>" & vbNewLine
end if

wHTML = wHTML & "</table>" & vbNewLine

wSeriesHTML = wHTML

RSv.close

End function

'========================================================================
'
'	Function	最近チェックした商品に追加
'
'========================================================================
'
Function AddViewdProduct()

Dim RSv

'----
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM 最近チェックした商品"
wSQL = wSQL & " WHERE 顧客番号 = " & wUserID
wSQL = wSQL & "   AND メーカーコード = '" & maker_cd & "'"
wSQL = wSQL & "   AND 商品コード = '" & Replace(product_cd, "'", "''") & "'"	' 2012/01/23 GV Mod (コード内にシングルクオーテーションが存在した場合の対応)
wSQL = wSQL & "   AND 色 = '" & iro & "'"
wSQL = wSQL & "   AND 規格 = '" & kikaku & "'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RSv.EOF = true then
	RSv.AddNew

	RSv("顧客番号") = wUserID
	RSv("メーカーコード") = maker_cd
	RSv("商品コード") = product_cd
	RSv("色") = iro
	RSv("規格") = kikaku
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
wSQL = wSQL & " WHERE SessionID = '" & gSessionID & "'"			'2011/04/14 hn mod
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
	wSQL = wSQL & "   AND 商品コード = '" & Replace(product_cd, "'", "''") & "'"	' 2012/01/23 GV Mod (コード内にシングルクオーテーションが存在した場合の対応)
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
'	Function	レコメンド商品アクセスログ	2009/12/17
'
'========================================================================
'
Function AddRecommendAccessLog()

Dim RSv

'----
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM レコメンド商品アクセスログ"
wSQL = wSQL & " WHERE 1 = 2"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

'---- レコメンド商品アクセスログ登録
RSv.AddNew

RSv("レコメンドユーザーID") = gSessionID		'2011/04/14 hn mod
RSv("メーカーコード") = maker_cd
RSv("商品コード") = product_cd
RSv("ユーザーエージェント") = Request.ServerVariables("HTTP_USER_AGENT")
RSv("アクセス日") = Now()

RSv.Update
RSv.close

End function

'========================================================================
'
'	Function	代替商品取得メソッド GV 2012/05/01
'
'========================================================================
Function GetSubstituteItem()
    
    Dim RSv
    Dim RSvSub
    Dim vSql

    wSQL = ""
	wSQL = wSQL & "SELECT "
    wSQL = wSQL & "    a.メーカーコード, "
    wSQL = wSQL & "    a.商品コード, "
    wSQL = wSQL & "    a.商品名, "
    wSQL = wSQL & "    a.後継機種メーカーコード, "
    wSQL = wSQL & "    a.後継機種商品コード, "
    wSQL = wSQL & "    b.引当可能数量 "
	wSQL = wSQL & "FROM "
    wSQL = wSQL & "    Web商品 a WITH (NOLOCK) "
    wSQL = wSQL & "INNER JOIN  "
    wSQL = wSQL & "    Web色規格別在庫 b WITH (NOLOCK) ON  "
    wSQL = wSQL & "    a.メーカーコード = b.メーカーコード AND "
    wSQL = wSQL & "    a.商品コード = b.商品コード "
	wSQL = wSQL & "WHERE a.メーカーコード = '" & maker_cd & "'"
	wSQL = wSQL & "    AND a.商品コード = '" & Replace(product_cd, "'", "''") & "'"
	
    Set RSv = Server.CreateObject("ADODB.Recordset")
	RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic
    
	'2013/08/14 GV add start
	If RSv.EOF = True Then
		'//データが存在しない場合終了
		RSv.Close
		Exit Function
	End If
	'2013/08/14 GV add end

    '//在庫が存在しない場合で代替商品が存在する場合
	If RSv("引当可能数量") <= 0 And (RSv("後継機種メーカーコード") <> "" And RSv("後継機種商品コード") <> "") Then
        
        '//後継機商品コードの商品情報、在庫を取得する
        'calcPrice(RS("販売単価"), wSalesTaxRate)
        vSql = ""
        vSql = vSql & " SELECT "
	    vSql = vSql & " a.メーカーコード,"
	    vSql = vSql & " c.メーカー名,"
	    vSql = vSql & " a.商品コード,"
	    vSql = vSql & " a.商品名,"
		vSql = vSql & " CASE"
		vSql = vSql & "   WHEN (a.個数限定数量 > a.個数限定受注済数量 AND a.個数限定数量 > 0) THEN a.個数限定単価"
		vSql = vSql & "   WHEN (a.B品フラグ = 'Y') THEN a.B品単価"
		vSql = vSql & "   ELSE a.販売単価"
		vSql = vSql & " END AS 販売単価,"
	    vSql = vSql & " a.商品画像ファイル名_大,"
	    vSql = vSql & " a.商品画像ファイル名_小,"
		vSql = vSql & " a.ASK商品フラグ"
        vSql = vSql & " FROM "
        vSql = vSql & " 	Web商品 a WITH (NOLOCK) "
        vSql = vSql & " INNER JOIN "
        vSql = vSql & " 	Web色規格別在庫 b WITH (NOLOCK) ON "
        vSql = vSql & " 		a.メーカーコード = b.メーカーコード AND "
        vSql = vSql & " 		a.商品コード = b.商品コード "
        vSql = vSql & " INNER JOIN メーカー c WITH (NOLOCK) ON "
        vSql = vSql & " 	c.メーカーコード = a.メーカーコード "
        vSql = vSql & " WHERE  "
        vSql = vSql & " 	a.メーカーコード = '" & RSv("後継機種メーカーコード") &  "' AND "
        vSql = vSql & " 	a.商品コード = '" & RSv("後継機種商品コード") & "' AND "
        vSql = vSql & " 	a.Web商品フラグ = 'Y' AND "
        vSql = vSql & " 	((b.引当可能数量 > 0 ) OR (a.B品フラグ = 'Y') AND (b.B品引当可能数量 > 0))"
        '@@@@@@Response.Write(vSql)
        
        Set RSvSub = Server.CreateObject("ADODB.Recordset")
	    RSvSub.Open vSql, Connection, adOpenStatic, adLockOptimistic
        
        If RSvSub.EOF = True Then
            '//データが存在しない場合終了
            RSv.close
	        RSvSub.close
	        Exit Function
        End If

        wSubItemHTML = ""
	'2012/07/10 GV Mod Start
'        wSubItemHTML = wSubItemHTML & "<div id='alt-item'>" & vbNewLine
'        wSubItemHTML = wSubItemHTML & "<p class='head'>このアイテムは<br>すぐにお届けできます。</p>" & vbNewLine
'        wSubItemHTML = wSubItemHTML & "<p><a href='ProductDetail.asp?Item=" & RSvSub("メーカーコード") & "%5E" & RSvSub("商品コード") & "'>" & vbNewLine
'        wSubItemHTML = wSubItemHTML & "<img src='prod_img/" & RSvSub("商品画像ファイル名_小") & "' width='170' height='85' border='0'></a><br>" & vbNewLine
'        wSubItemHTML = wSubItemHTML & RSvSub("メーカー名") & " " & RSvSub("商品名") & "<br>" & vbNewLine
'		If RSvSub("ASK商品フラグ") <> "Y" Then
'	        wSubItemHTML = wSubItemHTML & FormatNumber(calcPrice(RSvSub("販売単価"), wSalesTaxRate),0) & "円(税込)</p>" & vbNewLine
'		Else
'	        wSubItemHTML = wSubItemHTML & "<a class='tip'>ASK<span>" & FormatNumber(calcPrice(RSvSub("販売単価"), wSalesTaxRate),0) & "円(税込)</span></a></p>" & vbNewLine
'		End If
'        wSubItemHTML = wSubItemHTML & "</div>" & vbNewLine
	wSubItemHTML = wSubItemHTML & "<div class='detail_side_inner01'><div class='detail_side_inner02'>"
	wSubItemHTML = wSubItemHTML & "<div class='detail_side_inner_box'>" & vbNewLine
	wSubItemHTML = wSubItemHTML & "  <!-- このアイテムはすぐにお届けできます -->" & vbNewLine
	wSubItemHTML = wSubItemHTML & "  <h4 class='detail_sub truck'>このアイテムは<br>すぐにお届けできます</h4>" & vbNewLine
	wSubItemHTML = wSubItemHTML & "  <ul class='check_item'>" & vbNewLine
	wSubItemHTML = wSubItemHTML & "    <li>" & vbNewLine

	wSubItemHTML = wSubItemHTML & "      <p><a href='ProductDetail.asp?Item=" & RSvSub("メーカーコード") & "%5E" & RSvSub("商品コード") & "'>"
	If RSvSub("商品画像ファイル名_小") <> "" Then
		wSubItemHTML = wSubItemHTML & "<img src='prod_img/" & RSvSub("商品画像ファイル名_小") & "' alt='" & Replace(RSvSub("メーカー名") & " / " & RSvSub("商品名"),"'","&#39;") & "' class='opover'>"
	End If		
	wSubItemHTML = wSubItemHTML & RSvSub("メーカー名") & " / " & RSvSub("商品名") & "</a></p>" & vbNewLine

	wSubItemHTML = wSubItemHTML & "    </li>" & vbNewLine
	wSubItemHTML = wSubItemHTML & "  </ul>" & vbNewLine
	wSubItemHTML = wSubItemHTML & "</div>" & vbNewLine
	wSubItemHTML = wSubItemHTML & "</div></div>" & vbNewLine
	'2012/07/10 GV Mod End

    Else
        RSv.close
        Exit Function
    End If

End Function


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
'
'	Function	仮装パスから物理パスを取得する（ファイルの存在確認有り）
'               2011/11/22 an FHからコピー
'
'========================================================================
Function GetMapPath(pTargetPath, pFileExtention)

Dim vFSO
Dim vTarget

pFileExtention = ""
GetMapPath = ""

Set vFSO = CreateObject("Scripting.FileSystemObject")
vTarget = Server.MapPath(pTargetPath)

pFileExtention = vFSO.GetExtensionName(vTarget)

' 拡張子がtxtの場合、"/shop"からの相対アドレスが指定されている
If LCase(pFileExtention) = "txt" Then
	vTarget = Server.MapPath(pTargetPath)
End If

If vFSO.FileExists(vTarget) = True Then
	GetMapPath = vTarget
End If

Set vFSO = Nothing

End Function

'2012/07/10 GV Add Start
'========================================================================
'
'	Function	最近チェックした商品一覧
'
'========================================================================
'
Function CreateViewedProductList()

Dim RSv
Dim vHTML
Dim vSQL

Dim vPrice
'Dim vCnt
Dim vName

NAVIViewedProductListHTML = ""

'---- 最近チェックした商品 取り出し
vSQL = ""
'表示件数を10件から5件に変更	'2012/07/20 ok Mod
'vSQL = vSQL & "SELECT TOP 10"
vSQL = vSQL & "SELECT TOP 5"
vSQL = vSQL & "       a.メーカーコード"
vSQL = vSQL & "     , a.商品コード"
vSQL = vSQL & "     , a.色"
vSQL = vSQL & "     , a.規格"
vSQL = vSQL & "     , b.商品画像ファイル名_小"
vSQL = vSQL & "     , b.商品名"
vSQL = vSQL & "     , c.メーカー名"
vSQL = vSQL & "  FROM 最近チェックした商品 a WITH (NOLOCK)"
vSQL = vSQL & "     , Web商品 b WITH (NOLOCK)"
vSQL = vSQL & "     , メーカー c WITH (NOLOCK)"
vSQL = vSQL & " WHERE b.メーカーコード = a.メーカーコード"
vSQL = vSQL & "   AND b.商品コード = a.商品コード"
vSQL = vSQL & "   AND c.メーカーコード = a.メーカーコード"
vSQL = vSQL & "   AND a.顧客番号 = " & wUserID
vSQL = vSQL & " ORDER BY"
vSQL = vSQL & "       a.チェック日 DESC"

'@@@@@@@@@@response.write(vSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

If RSv.EOF Then
	Exit Function
End If

	vHTML = vHTML & "    <div class='detail_side_inner01'><div class='detail_side_inner02'>" & vbNewLine
	vHTML = vHTML & "      <div class='detail_side_inner_box'>" & vbNewLine
	vHTML = vHTML & "        <h4 class='detail_sub'>最近チェックした商品</h4>" & vbNewLine
	vHTML = vHTML & "        <ul class='check_item'>" & vbNewLine

'vCnt = 0

Do Until RSv.EOF

'	vCnt = vCnt  + 1
	vName = ""
	
	
	vName = vName & RSv("メーカー名") & " / " & RSv("商品名")
	If Trim(RSv("色")) <> "" AND Trim(RSv("規格")) <> ""then
		vName = vName & " / " & Trim(RSv("色")) & " / " & Trim(RSv("規格"))
	End If

	If Trim(RSv("色")) <> "" AND Trim(RSv("規格")) = ""then
		vName = vName & " / " & Trim(RSv("色"))
	End If

	If Trim(RSv("色")) = "" AND Trim(RSv("規格")) <> ""then
		vName = vName & " / " & Trim(RSv("規格"))
	End If
		
	'---- メーカー名、商品名/色/規格
	vHTML = vHTML & "    <li>" & vbNewLIne
	vHTML = vHTML & "      <p><a href='ProductDetail.asp?Item=" & RSv("メーカーコード") & "^" & Server.URLEncode(RSv("商品コード")) & "^" & Trim(RSv("色")) & "^" & Trim(RSv("規格")) & "'>"

'	If vCnt <= 5 Then		'5個までは画像表示
		'---- 商品画像
	If RSv("商品画像ファイル名_小") <> "" Then
		vHTML = vHTML & "<img src='prod_img/" & RSv("商品画像ファイル名_小") & "' alt='" & Replace(vName,"'","&#39;")  & "' class='opover'>" 
	End If
'	End If

	vHTML = vHTML & vName & "</a></p>" & vbNewLine
	vHTML = vHTML & "    </li>" & vbNewLIne

	RSv.MoveNext

Loop

	vHTML = vHTML & "      </ul>" & vbNewLine
	vHTML = vHTML & "    </div>" & vbNewLine
	vHTML = vHTML & "  </div></div>" & vbNewLine

RSv.close

wViewHTML = vHTML

End function
'2012/07/10 GV Add End

'2012/10/30 nt Add Start
'========================================================================
'
'	Function	関連コンテンツ取得
'
'========================================================================
Function CreateContentsHTML()

Dim RSv
Dim vCnt

'---- コンテンツ取得
wSQL = ""
wSQL = wSQL & "SELECT DISTINCT TOP 3"
wSQL = wSQL & "       b.コンテンツ名"
wSQL = wSQL & "     , b.URL"
wSQL = wSQL & "     , b.説明"
wSQL = wSQL & "     , b.画像ファイル名"
wSQL = wSQL & "     , a.関連区分"
wSQL = wSQL & "     , b.優先順位"
wSQL = wSQL & "  FROM 商品コンテンツ a WITH (NOLOCK)"
wSQL = wSQL & "     , コンテンツ b WITH (NOLOCK)"
wSQL = wSQL & " WHERE a.コンテンツ番号 = b.コンテンツ番号"
'wSQL = wSQL & "   AND b.リンクフラグ = 'Y'"
wSQL = wSQL & "   AND ((a.メーカーコード = '" & maker_cd & "'"
wSQL = wSQL & "   AND a.商品コード = '" & Replace(product_cd, "'", "''") & "'"
wSQL = wSQL & "   AND a.関連区分='1')"
'wSQL = wSQL & "   OR  ('" & Replace(product_cd, "'", "''") & "' LIKE a.商品コード + '%'"
'wSQL = wSQL & "   AND a.関連区分='2')"
wSQL = wSQL & "   OR  (a.シリーズコード='" & wSeriesCd & "'"
wSQL = wSQL & "   AND a.関連区分='3')"
wSQL = wSQL & "   OR  (a.メーカーコード='" & maker_cd & "'"
wSQL = wSQL & "   AND a.関連区分='4')"
wSQL = wSQL & "   OR  (a.カテゴリーコード='" & wCategoryCode & "'"
wSQL = wSQL & "   AND a.関連区分='5')"
wSQL = wSQL & "   OR  (a.中カテゴリーコード='" & wMidCategoryCd & "'"
wSQL = wSQL & "   AND a.関連区分='6')"
wSQL = wSQL & "   OR  (a.大カテゴリーコード='" & wLargeCategoryCd & "'"
wSQL = wSQL & "   AND a.関連区分='7')"
wSQL = wSQL & "   OR  (a.カテゴリーコード='" & wCategoryCode & "'"
wSQL = wSQL & "   AND a.メーカーコード = '" & maker_cd & "'"
wSQL = wSQL & "   AND a.関連区分='8')"
wSQL = wSQL & "   OR  (a.中カテゴリーコード='" & wMidCategoryCd & "'"
wSQL = wSQL & "   AND a.メーカーコード = '" & maker_cd & "'"
wSQL = wSQL & "   AND a.関連区分='9'))"
wSQL = wSQL & " ORDER BY"
wSQL = wSQL & "       b.優先順位 ASC, a.関連区分 ASC"

'@@@@@@@ Debug
'response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

wContentsHTML = ""

If RSv.EOF = True Then
	RSv.close
	Exit Function
End If

wHTML = ""
wHTML = wHTML & "<div class='detail_side_inner01'><div class='detail_side_inner02'>" & vbNewLine
wHTML = wHTML & "<div class='detail_side_inner_box'>" & vbNewLine
wHTML = wHTML & "  <h4 class='detail_sub'>この商品に関連するセレクション</h4>" & vbNewLine
wHTML = wHTML & "  <ul class='special'>" & vbNewLine

Do Until RSv.EOF = True
	wHTML = wHTML & "    <li>" & vbNewLine

	If RSv("画像ファイル名") <> "" Then
		wHTML = wHTML & "        <p class='photo'><a href='" & RSv("URL") & "'><img src='../" & RSv("画像ファイル名") & "' alt='" & RSv("コンテンツ名") & "' class='opover' width='120' height='90' /></a></p>" & vbNewLine
	End If
	wHTML = wHTML & "          <p class='txt'><a href='" & RSv("URL") & "'>" & RSv("コンテンツ名") & "</a></p>" & vbNewLine

	wHTML = wHTML & "    </li>" & vbNewLine

	RSv.MoveNext
Loop

wHTML = wHTML & "  </ul>" & vbNewLine
wHTML = wHTML & "</div>" & vbNewLine
wHTML = wHTML & "</div></div>" & vbNewLine


wContentsHTML = wContentsHTML & wHTML

RSv.close

End Function
'2012/10/30 nt Add End

'========================================================================
%>
<!DOCTYPE html>
<html lang="ja">
<head prefix="og: http://ogp.me/ns# fb: http://ogp.me/ns/fb#">
<meta charset="Shift_JIS">
<meta name="robots" content="noindex,nofollow">
<link rel="canonical" href="http://www.soundhouse.co.jp/shop/ProductDetail.asp?Item=<%=Server.URLEncode(maker_cd & "^" & product_cd & "^" & iro & "^" & kikaku)%>">
<title><%=wMakerName%>&gt;<%=wProductName%>｜サウンドハウス</title>
<% if wTokucho <> "" then%><meta name="description" content="<%=wTokucho%>"><% end if %>
<meta name="keywords" content="<%=wLargeCategoryName%>,<%=wMidCategoryName%>,<%=wCategoryName%>,<%=wMakerName%>,<%=wProductName%>">
<meta name="twitter:card" content="product">
<meta name="twitter:url" content="http://www.soundhouse.co.jp/shop/ProductDetail.asp?Item=<%=Server.URLEncode(maker_cd & "^" & product_cd & "^" & iro & "^" & kikaku)%>">
<meta name="twitter:site" content="@soundhouse_jp">
<meta name="twitter:image:width" content="600">
<meta name="twitter:image:height" content="300">
<meta name="twitter:label1" content="<%=wTwPriceLabel%>">
<meta name="twitter:data1" content="<%=wTwPriceData%>">
<meta name="twitter:label2" content="在庫状況">
<% if wTwInventoryData <> "" Then%><meta name="twitter:data2" content="<%=wTwInventoryData%>"><% Else %><meta name="twitter:data2" content="サイトをご覧ください"><% End If %>
<meta property="og:title" content="<%=wMakerName%>&gt;<%=wProductName%>｜サウンドハウス">
<meta property="og:type" content="article">
<meta property="og:url" content="http://www.soundhouse.co.jp/shop/ProductDetail.asp?Item=<%=Server.URLEncode(maker_cd & "^" & product_cd & "^" & iro & "^" & kikaku)%>">
<meta property="og:image" content="<%=g_HTTP%>shop/prod_img/<%=wMainProdPic%>">
<% if wTokucho <> "" Then%><meta property="og:description" content="<%=wTokucho%>"><% Else %><meta property="og:description" content="<%=wMakerName%>&gt;<%=wProductName%>"><% End If %>
<meta property="og:site_name" content="サウンドハウス">
<meta property="og:locale" content="ja_JP">
<meta property="fb:app_id" content="191447484218062">
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/shop.css?20140401a" type="text/css">
<link rel="stylesheet" href="style/jquery.fancybox-1.3.4.css" type="text/css">
<link rel="stylesheet" href="Style/ProductDetail.css?20140401" type="text/css">
<link rel="stylesheet" href="style/ask.css?20140401a" type="text/css">
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

	if (pForm.IroKikaku.length > 1){
		if (pForm.IroKikaku.selectedIndex == 0){
			alert("<%=wIroKikakuSelectMsg%>");
			return false;
		}else{			//色規格を送信エリアへセット
			pForm.Item.value = pForm.Item.value + "^" + pForm.IroKikaku.options[pForm.IroKikaku.selectedIndex].value;
		}
	}

	return true;

}
//
// ====== 	Function:	BuyTogether_onClick
//
function BuyTogether_onClick(pProd){

	if (pProd.checked == true){
		document.f_data.AdditionalItem.value = document.f_data.AdditionalItem.value + "," + pProd.value;	}else{
		document.f_data.AdditionalItem.value = document.f_data.AdditionalItem.value.replace("," + pProd.value,"");
	}
}

//
// ====== 	Function:	IroKikaku_onChange
//
function IroKikaku_onChange(pForm){

	var i;

	i = pForm.IroKikaku.selectedIndex;

	document.fIroKikaku.Item.value = document.fIroKikaku.Item.value + "^" + pForm.IroKikaku.options[i].value;
	document.fIroKikaku.submit();
}

//
// ====== 	Function:	review_onSubmit
//
function review_onSubmit(pForm){

	if (pForm.Title.value == ""){
		alert("タイトルを入力してください｡");
		return false;
	}
	if (pForm.HandleName.value == ""){
		alert("お名前を入力してください｡");
		return false;
	}
	if (pForm.Review.value == ""){
		alert("レビューを入力してください｡");
		return false;
	}
	if (pForm.Review.value.length > 1000){
		alert("レビュー文字数が1000文字を超えています｡　1000文字以内でお願いします。");
		return false;
	}
	if (pForm.Review.value.indexOf("ttp://",0)  > 0){
		alert("リンクが含まれています。レビューへはリンクは登録できません。");
		return false;
	}
	return true;
}

//
//	ReviewSankou_onClick		'2010/03/08 hn add
//
function ReviewSankou_onClick(pID, pItem, pSankou){

var vAction;

		vAction = "Review" + "Sankou.asp";
		document.fReviewSankou.ID.value = pID;
		document.fReviewSankou.Item.value = pItem;
		document.fReviewSankou.Sankou.value = pSankou;
		document.fReviewSankou.action = vAction;
    	document.fReviewSankou.submit();
}

</script>

</head>
<body>
<div id="fb-root"></div>
<script>(function(d, s, id) {
  var js, fjs = d.getElementsByTagName(s)[0];
  if (d.getElementById(id)) return;
  js = d.createElement(s); js.id = id;
  js.src = "//connect.facebook.net/ja_JP/all.js#xfbml=1&appId=191447484218062";
  fjs.parentNode.insertBefore(js, fjs);
}(document, 'script', 'facebook-jssdk'));</script>
<!--#include file="../Navi/Navitop.inc"-->

<div id="globalMain">
	<span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>
	<!-- コンテンツstart -->
	<div id="globalContents" itemscope itemtype="http://data-vocabulary.org/Product">
    	<div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
			<p class="home"><a href="../"><img src="../images/icon_home.gif" alt="HOME"></a></p>
			<ul id="path">
				<!-- パン屑リスト -->
				<%=wTitleWithLink%>
			</ul>
		</div></div></div>
        
        <!-- 商品詳細 -->
        <h1 class="title"><span itemprop="brand"><%=wMakerName%></span> / <span itemprop="name"><%=wProductname%></span></h1>
		<div id="productdetail">
        
        <div id="detail">
        	
			<div id="detail_inner01"><div id="detail_inner02">
<!-- 商品画像 -->
<%=wPictureHTML%>

<!-- 関連リンク -->
<%=wKanrenLinkHTML%>

<div id="inner_box">

<!-- 特徴 -->
<%=wTokuchoHTML%>

<!-- スペック -->
<%=wSpecHTML%>

<!-- レコメンド2 -->
<%= wRecommendBuyJS %>

<!-- オプション -->
<%if wOptionHTML <> "" then %>

<%=wOptionHTML%>

<% end if %>

<!-- パーツ -->
<%if wPartsHTML <> "" then %>

<%=wPartsHTML%>

<% end if %>

<!-- 商品レビュー -->
<% if wReviewHTML <> "" then %>

<h2 id="review" class="detail_title">商品レビュー</h2>
<%=wReviewHTML%>

<% elseif wCanWriteReviewFl = "Y" and WriteReview <> "Y" then %>

<h2 class="detail_title">レビューを投稿する</h2>
<ul class="btn_review">

<% end if %>

<% if wCanWriteReviewFl = "Y" then %>
<% '2013/05/17 GV #1507
'旧レビュー編集部分を表示しないよう、if文で制御
'if WriteReview = "Y" then %>
	<% if 1=0 then %>
  </ul>
					<h2 class="detail_title">レビューを投稿する</h2>
					<div class="comment_box no_line">
						<form name="f_review" method="post" action="ReviewStore.asp" onSubmit="return review_onSubmit(this);">
						<table class="comment">
							<tr>
							 	<th><span class="pp">評価</span></th>
							 	<td>
							 	<select name="Rating">
									<option value="5">♪×5</option>
									<option value="4">♪×4</option>
									<option value="3">♪×3</option>
									<option value="2">♪×2</option>
									<option value="1">♪×1</option>
								</select>
								<span>最大5つまで</span>
							</td>
							</tr>
							<tr>
								<th><span class="pp">タイトル</span></th><td><input type="text" name="Title" id="Title" maxsize="50"></td>
							</tr>
							<tr>
		<% if wHandleName = "" then %>
								<th><span class="pp">ハンドルネーム</span></th><td><input type="text" name="HandleName" id="HandleName" maxlength="30"></td>
		<% else %>
								<th><span class="pp">ハンドルネーム</span></th><td><%=wHandleName%><input type="hidden" name="HandleName" id="HandleName" value="<%=wHandleName%>"></td>
		<% end if %>
							</tr>
							<tr>
								<th><span class="pp">住所</span></th><td class="address"><%=wPrefecture%>（会員登録住所が表示されます）</td>
							</tr>
							<tr>
								<th><span class="pp">レビュー</span><br>（1000文字まで）</th><td><textarea name="Review" rows="8" cols="70" style="width:325px;ime-mode:auto;"></textarea></td>
							</tr>
						 </table>
                         <div class="submit">
                         	<div class="review_attention">
                            	<h4>当レビューは、商品に関するコメントのみをお願いします。</h4>
                                <p>以下に該当する場合、弊社の判断にて削除、訂正等を行なう場合もございますので、あらかじめご了承ください。</p>
                                <ul>
                                	<li>商品に対しての評価とは関係の無いコメント</li>
                                    <li>他のレビューに対しての意見、コメント</li>
                                    <li>誹謗中傷や、いたずらと思われる記述</li>
                                </ul>
                            </div>
                            <input type="image" src="images/btn_review_submit.png" alt="投稿する"><p>一旦投稿されたレビューは変更できません。</p>
                            <input type="hidden" name="OrderNo" value="<%=OrderNo%>">
                            <input type="hidden" name="Item" value="<%=item%>">
                        </div>
                        </form>
					</div>

	<% else %> 
<%
'2013/05/17 GV #1507 modified start
'  <li><a href="ProductDetail.asp?Item=<%=item%'>&WriteReview=Y"><img src="images/btn_review_write.png" alt="この商品のレビューを書く" class="opover"></a></li>
'  </ul>
Dim UrlEncodeItem
UrlEncodeItem = Server.URLEncode(Item)
%>
  <li><a href="<%=g_HTTPS%>Shop/ReviewWrite.asp?Item=<%=UrlEncodeItem%>"><img src="images/btn_review_write.png" alt="この商品のレビューを書く" class="opover"></a></li>
</ul>
<%
'2013/05/17 GV #1507 modified end
%>
	<% end if %>
<% else %>
	<% if wReviewHTML <> "" Or (wCanWriteReviewFl = "Y" and WriteReview <> "Y")then %>
</ul>
	<% end if %>
<% end if %>
				 </div>
			</div></div>
		<!--/#detail --></div>
		
<!-- ここから右側 ====================================================== -->
    <div id="detail_side">
<!-- メーカー/商品名 -->
<%=wProductHTML%>

<!-- カート情報 -->
<%=wCartHTML%>
<!-- 代替商品表示 -->
<%=wSubItemHTML%>

<!-- 関連コンテンツ -->
<%=wContentsHTML%>

<!--　関連マップへのリンク　-->
<!--
    <div class="detail_side_inner01"><div class="detail_side_inner02">
      <div class="detail_side_inner_box">
        <ul class="check_item">
          <li><a href="../recommend/RecommendMap.asp?item=<%=item%>"><img src="images/btn_recommendmap.png" alt="関連アイテムをマップで表示" class="opover">関連アイテムをマップで表示</a></li>
        </ul>
      </div>
    </div></div>
-->

<!-- 関連シリーズ商品 -->
<%=wSeriesHTML%>

<!-- レコメンド結果 -->
<%=wRecommendJS%>

<!-- 最近チェックした商品 -->
<%=wViewHTML%>

<!-- 色規格選択されたときにProductDetail.aspを再呼び出し-->
<form name="fIroKikaku" method="get" action="ProductDetail.asp">
	<input type="hidden" name="Item" value="<%=maker_cd%>^<%=product_cd%>">
</form>
<!-- レビュー　はい/いいえ　時ReviewSankou.asp呼び出し用 2010/03/08 hn add -->
<form name="fReviewSankou" method="post" action="">
	<input type="hidden" name="ID" value="">
	<input type="hidden" name="Item" value="">
	<input type="hidden" name="Sankou" value="">
</form>
        </div>
        </div>
      <!--/#contents --></div>
	<div id="globalSide">
	<!--#include file="../Navi/NaviSide.inc"-->
	<!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/Navibottom.inc"-->
<div class="tooltip"><p>ASK</p></div>
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript" src="jslib/jquery.fancybox-1.3.4.pack.js"></script>
<script type="text/javascript" src="jslib/jquery.easing.1.3.js"></script>
<script type="text/javascript" src="../jslib/jquery.carouFredSel-5.5.0-packed.js"></script>
<script type="text/javascript" src="jslib/ask.js?20140401a"></script>
<script type="text/javascript" src="jslib/ProductDetail.js?20130709"></script>
<script type="text/javascript" src="http://platform.twitter.com/widgets.js" charset="utf-8"></script>
</body>
</html>
