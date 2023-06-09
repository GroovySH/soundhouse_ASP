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
'	ベストセラー一覧ページ
'
'更新履歴
'2005/03/31 色規格がある商品へのリンクは商品個別ではなく商品一覧に変更
'2006/04/05 集計単位を中カテゴリー別に変更
'2006/11/08 メーカー名+商品名を25文字でカット
'2007/05/25 シリーズ対応
'2009/04/30 エラー時にerror.aspへ移動
'2010/05/29 ランキングページリニューアル対応
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/01/19 GV データ取得 SELECT文へ LACクエリー案を適用
'2012/01/20 GV データ取得 SELECT文から売筋商品テーブルの最新月のデータのみ抽出する条件を削除
'2012/08/07 if-web リニューアルレイアウト調整
'
'========================================================================

On Error Resume Next

'----2010/05/29 st add
Dim LargeCategoryCd

Dim wLargeCategoryHTML
Dim wLargeCategoryName
Dim wNoData
'----2010/05/29

Dim wYYYYMM

Dim Connection
Dim RS

Dim wSQL
Dim wHTML
Dim w_error_msg
Dim wErrDesc   '2011/08/01 an add

'========================================================================

'---- Get input data 2010/05/29 st add
LargeCategoryCd = ReplaceInput(Trim(Request("LargeCategoryCd")))

'---- 大カテゴリーコードの指定がない場合
if LargeCategoryCd = "" then
	LargeCategoryCd = "1"
end if
'----  2010/05/29

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "BestSellerList.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
end if                                           '2011/08/01 an add e

call close_db()

'---- 2010/05/29 st mod
if wNoData = "Y" OR Err.Description <> "" then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if
'----  2010/05/29

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

'----2010/05/29 st del
'Dim vLCatHTML
'Dim vCatHTML
'Dim vProdHTML(10)
'Dim vBreakKey2
'Dim vBreakNextKey2
'Dim i
'----2010/05/29

'----2010/05/29 st add
Dim vRank
Dim vBGColor
'----2010/05/29

Dim vCatCount
Dim vProdCount
Dim vMakerProduct
Dim vYYYYMM
Dim vBreakKey1
Dim vBreakNextKey1

vYYYYMM = Left(fFormatDate(Date()), 7)


'---- 大カテゴリー一覧作成
call CreateLargeCategoryHTML()
if wNoData <> "Y" then  '想定外の大カテゴリーを指定されてNoDataの場合はエラー


'---- 売れ筋ランキング 取り出し
wSQL = ""
' 2012/01/19 GV Mod Start
'wSQL = wSQL & "SELECT"
'wSQL = wSQL & "       a.メーカーコード"
'wSQL = wSQL & "     , a.商品コード"
'wSQL = wSQL & "     , '' AS シリーズコード"
'wSQL = wSQL & "     , a.受注数量"
'wSQL = wSQL & "     , b.商品名"
'wSQL = wSQL & "     , c.メーカー名"
'wSQL = wSQL & "     , e.中カテゴリーコード"
'wSQL = wSQL & "     , e.中カテゴリー名日本語"
'wSQL = wSQL & "     , e.表示順 AS 中カテゴリー表示順"
'wSQL = wSQL & "     , f.大カテゴリーコード"
'wSQL = wSQL & "     , f.大カテゴリー名"
'wSQL = wSQL & "     , COUNT(g.商品コード) AS 色規格別在庫件数"
'wSQL = wSQL & "  FROM "
'wSQL = wSQL & "       売筋商品 a WITH (NOLOCK)"
'wSQL = wSQL & "     , Web商品 b WITH (NOLOCK)"
'wSQL = wSQL & "     , メーカー c WITH (NOLOCK)"
'wSQL = wSQL & "     , カテゴリー d WITH (NOLOCK)"
'wSQL = wSQL & "     , 中カテゴリー e WITH (NOLOCK)"
'wSQL = wSQL & "     , 大カテゴリー f WITH (NOLOCK)"
'wSQL = wSQL & "     , Web色規格別在庫 g WITH (NOLOCK)"
'wSQL = wSQL & " WHERE "
'wSQL = wSQL & "       b.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND b.商品コード = a.商品コード"
'wSQL = wSQL & "   AND b.カテゴリーコード = a.カテゴリーコード"
'wSQL = wSQL & "   AND c.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND d.カテゴリーコード = a.カテゴリーコード"
'wSQL = wSQL & "   AND e.中カテゴリーコード = d.中カテゴリーコード"
'wSQL = wSQL & "   AND f.大カテゴリーコード = e.大カテゴリーコード"
'wSQL = wSQL & "   AND g.メーカーコード = b.メーカーコード"
'wSQL = wSQL & "   AND g.商品コード = b.商品コード"
'wSQL = wSQL & "   AND b.終了日 IS NULL"
'wSQL = wSQL & "   AND g.終了日 IS NULL"
'wSQL = wSQL & "   AND b.Web商品フラグ = 'Y'"
'wSQL = wSQL & "   AND d.売れ筋ランキング表示フラグ = 'Y'"
'wSQL = wSQL & "   AND a.年月 = (SELECT MAX(年月) FROM 売筋商品)"
'wSQL = wSQL & "   AND f.大カテゴリーコード = '" & LargeCategoryCd & "'" '2010/05/29 st add
'wSQL = wSQL & " GROUP BY"
'wSQL = wSQL & "       a.メーカーコード"
'wSQL = wSQL & "     , a.商品コード"
'wSQL = wSQL & "     , a.受注数量"
'wSQL = wSQL & "     , b.商品名"
'wSQL = wSQL & "     , c.メーカー名"
'wSQL = wSQL & "     , e.中カテゴリーコード"
'wSQL = wSQL & "     , e.中カテゴリー名日本語"
'wSQL = wSQL & "     , e.表示順"
'wSQL = wSQL & "     , f.大カテゴリーコード"
'wSQL = wSQL & "     , f.大カテゴリー名"
'
'wSQL = wSQL & " UNION "
'
'wSQL = wSQL & "SELECT"
'wSQL = wSQL & "       a.メーカーコード"
'wSQL = wSQL & "     , '' AS 商品コード"
'wSQL = wSQL & "     , a.シリーズコード"
'wSQL = wSQL & "     , a.受注数量"
'wSQL = wSQL & "     , b.シリーズ名 AS 商品名"
'wSQL = wSQL & "     , c.メーカー名"
'wSQL = wSQL & "     , e.中カテゴリーコード"
'wSQL = wSQL & "     , e.中カテゴリー名日本語"
'wSQL = wSQL & "     , e.表示順 AS 中カテゴリー表示順"
'wSQL = wSQL & "     , f.大カテゴリーコード"
'wSQL = wSQL & "     , f.大カテゴリー名"
'wSQL = wSQL & "     , 2 AS 色規格別在庫件数"
'wSQL = wSQL & "  FROM "
'wSQL = wSQL & "       売筋商品 a WITH (NOLOCK)"
'wSQL = wSQL & "     , シリーズ b WITH (NOLOCK)"
'wSQL = wSQL & "     , メーカー c WITH (NOLOCK)"
'wSQL = wSQL & "     , カテゴリー d WITH (NOLOCK)"
'wSQL = wSQL & "     , 中カテゴリー e WITH (NOLOCK)"
'wSQL = wSQL & "     , 大カテゴリー f WITH (NOLOCK)"
'wSQL = wSQL & " WHERE "
'wSQL = wSQL & "       b.シリーズコード = a.シリーズコード"
'wSQL = wSQL & "   AND b.カテゴリーコード = a.カテゴリーコード"
'wSQL = wSQL & "   AND c.メーカーコード = a.メーカーコード"
'wSQL = wSQL & "   AND d.カテゴリーコード = a.カテゴリーコード"
'wSQL = wSQL & "   AND e.中カテゴリーコード = d.中カテゴリーコード"
'wSQL = wSQL & "   AND f.大カテゴリーコード = e.大カテゴリーコード"
'wSQL = wSQL & "   AND d.売れ筋ランキング表示フラグ = 'Y'"
'wSQL = wSQL & "   AND a.年月 = (SELECT MAX(年月) FROM 売筋商品)"
'wSQL = wSQL & "   AND f.大カテゴリーコード = '" & LargeCategoryCd & "'" '2010/05/29 st add
'
'wSQL = wSQL & " ORDER BY"
'wSQL = wSQL & "       f.大カテゴリーコード"
'wSQL = wSQL & "     , e.表示順"		'中カテゴリー表示順
'wSQL = wSQL & "     , a.受注数量 DESC"
wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.メーカーコード "
wSQL = wSQL & "    , a.商品コード "
wSQL = wSQL & "    , '' AS シリーズコード "
wSQL = wSQL & "    , a.受注数量 "
wSQL = wSQL & "    , b.商品名 "
wSQL = wSQL & "    , c.メーカー名 "
wSQL = wSQL & "    , e.中カテゴリーコード "
wSQL = wSQL & "    , e.中カテゴリー名日本語 "
wSQL = wSQL & "    , e.表示順 AS 中カテゴリー表示順 "
wSQL = wSQL & "    , f.大カテゴリーコード "
wSQL = wSQL & "    , f.大カテゴリー名 "
wSQL = wSQL & "    , COUNT(g.商品コード) AS 色規格別在庫件数 "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    売筋商品                     a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN Web商品         b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.メーカーコード     = a.メーカーコード "
wSQL = wSQL & "           AND b.商品コード         = a.商品コード "
wSQL = wSQL & "           AND b.カテゴリーコード   = a.カテゴリーコード "
wSQL = wSQL & "      INNER JOIN メーカー        c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.メーカーコード     = a.メーカーコード "
wSQL = wSQL & "      INNER JOIN カテゴリー      d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.カテゴリーコード   = a.カテゴリーコード "
wSQL = wSQL & "      INNER JOIN 中カテゴリー    e WITH (NOLOCK) "
wSQL = wSQL & "        ON     e.中カテゴリーコード = d.中カテゴリーコード "
wSQL = wSQL & "      INNER JOIN 大カテゴリー    f WITH (NOLOCK) "
wSQL = wSQL & "        ON     f.大カテゴリーコード = e.大カテゴリーコード "
wSQL = wSQL & "      INNER JOIN Web色規格別在庫 g WITH (NOLOCK) "
wSQL = wSQL & "        ON     g.メーカーコード     = b.メーカーコード "
wSQL = wSQL & "           AND g.商品コード         = b.商品コード "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' )   t1 "
wSQL = wSQL & "        ON     b.Web商品フラグ    = t1.ShohinWebY "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'HotSellingY' )  t2 "
wSQL = wSQL & "        ON     d.売れ筋ランキング表示フラグ = t2.HotSellingY "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        t1.ShohinWebY  IS NOT NULL "
wSQL = wSQL & "    AND t2.HotSellingY IS NOT NULL "
wSQL = wSQL & "    AND b.終了日       IS NULL "
wSQL = wSQL & "    AND g.終了日       IS NULL "
'wSQL = wSQL & "    AND a.年月 = (SELECT MAX(年月) FROM 売筋商品) "				' 2012/01/20 GV Del
wSQL = wSQL & "    AND f.大カテゴリーコード = '" & LargeCategoryCd & "' "
wSQL = wSQL & "GROUP BY "
wSQL = wSQL & "      a.メーカーコード "
wSQL = wSQL & "    , a.商品コード "
wSQL = wSQL & "    , a.受注数量 "
wSQL = wSQL & "    , b.商品名 "
wSQL = wSQL & "    , c.メーカー名 "
wSQL = wSQL & "    , e.中カテゴリーコード "
wSQL = wSQL & "    , e.中カテゴリー名日本語 "
wSQL = wSQL & "    , e.表示順 "
wSQL = wSQL & "    , f.大カテゴリーコード "
wSQL = wSQL & "    , f.大カテゴリー名 "

wSQL = wSQL & "UNION "

wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.メーカーコード "
wSQL = wSQL & "    , '' AS 商品コード "
wSQL = wSQL & "    , a.シリーズコード "
wSQL = wSQL & "    , a.受注数量 "
wSQL = wSQL & "    , b.シリーズ名 AS 商品名 "
wSQL = wSQL & "    , c.メーカー名 "
wSQL = wSQL & "    , e.中カテゴリーコード "
wSQL = wSQL & "    , e.中カテゴリー名日本語 "
wSQL = wSQL & "    , e.表示順 AS 中カテゴリー表示順 "
wSQL = wSQL & "    , f.大カテゴリーコード "
wSQL = wSQL & "    , f.大カテゴリー名 "
wSQL = wSQL & "    , 2 AS 色規格別在庫件数 "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    売筋商品                  a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN シリーズ     b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.シリーズコード     = a.シリーズコード "
wSQL = wSQL & "           AND b.カテゴリーコード   = a.カテゴリーコード "
wSQL = wSQL & "      INNER JOIN メーカー     c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.メーカーコード     = a.メーカーコード "
wSQL = wSQL & "      INNER JOIN カテゴリー   d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.カテゴリーコード   = a.カテゴリーコード "
wSQL = wSQL & "      INNER JOIN 中カテゴリー e WITH (NOLOCK) "
wSQL = wSQL & "        ON     e.中カテゴリーコード = d.中カテゴリーコード "
wSQL = wSQL & "      INNER JOIN 大カテゴリー f WITH (NOLOCK) "
wSQL = wSQL & "        ON     f.大カテゴリーコード = e.大カテゴリーコード "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'HotSellingY' )  t1  "
wSQL = wSQL & "        ON     d.売れ筋ランキング表示フラグ = t1.HotSellingY "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        t1.HotSellingY IS NOT NULL "
'wSQL = wSQL & "    AND a.年月 = (SELECT MAX(年月) FROM 売筋商品) "				' 2012/01/20 GV Del
wSQL = wSQL & "    AND f.大カテゴリーコード = '" & LargeCategoryCd & "' "

wSQL = wSQL & "ORDER BY"
wSQL = wSQL & "      f.大カテゴリーコード "
wSQL = wSQL & "    , e.表示順 "
wSQL = wSQL & "    , a.受注数量 DESC "
' 2012/01/19 GV Mod End
'@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

'----- 売れ筋ランキングHTML編集
'---- Break Key initialize
If RS.EOF = True Then
    vBreakNextKey1 = "@EOF"
'    vBreakNextKey2 = "@EOF"  '2010/05/29 st del
Else
    vBreakNextKey1 = RS("中カテゴリーコード")
'    vBreakNextKey2 = RS("大カテゴリーコード")  '2010/05/29 st del
End If

'---- 大カテゴリー見出し編集
wHTML = ""
wHTML = wHTML & "<!--  container START  -->" & vbNewLine
wHTML = wHTML & "  <div id='container'>" & vbNewLine
wHTML = wHTML & "  <h1>" & RS("大カテゴリー名") & "</h1>" & vbNewLine


Do Until (RS.EOF = True)

'---- 大カテゴリー見出し編集 2010/05/29 st del
'	if vBreakKey2 <> vBreakNextKey2 then
'		wHTML = wHTML & "  <h1>" & RS("大カテゴリー名") & "</h1>" & vbNewLine
'		vCatCount = 0
'	end if
'2010/05/29 st del

	vBreakKey1 = vBreakNextKey1
    'vBreakKey2 = vBreakNextKey2  '2010/05/29 st del

	'---- 中カテゴリー別商品HTML 作成
	If vCatCount Mod 3 = 0 Then
		'---- 順位タイトル
		wHTML = wHTML & "<!-- ranking_container START -->" & vbNewLine
		wHTML = wHTML & "    <div class='ranking_container'>" & vbNewLine
		wHTML = wHTML & "      <div class='juni'>" & vbNewLine
		wHTML = wHTML & "       <div class='juni_th'>順位</div>" & vbNewLine

		'---- 商品順位 1～10設定
		for vRank=1 to 10

			'---- 偶数と奇数で背景色を変更
			if vRank Mod 2 <> 0 then
				vBGColor = "juni_td bg_color1"
			else
				vBGColor = "juni_td bg_color2"
			end if

			wHTML = wHTML & "       <div class='" & vBGColor & "'>" & vRank & "位</div>" & vbNewLine

		Next

		wHTML = wHTML & "      </div>" & vbNewLine
	End If

	'---- カテゴリーヘッダ
	wHTML = wHTML & "      <div class='rank_cat_box'>" & vbNewLine
	wHTML = wHTML & "        <div class='rank_cat_th'>" & vbNewLine
	wHTML = wHTML & "          <a href='MidCategoryList.asp?MidCategoryCd=" & RS("中カテゴリーコード") & "'>" & RS("中カテゴリー名日本語") & "</a>" & vbNewLine
	wHTML = wHTML & "        </div>" & vbNewLine

	vCatCount = vCatCount + 1
	vRank = 0

	'---- 1～10位の商品作成　カテゴリー毎売れ筋商品
  Do Until (vBreakKey1 <> vBreakNextKey1)
    vRank = vRank + 1

		'----メーカー名+商品名セット
		vMakerProduct = RS("メーカー名") & ":" & RS("商品名")
		if Len(vMakerProduct) > 25 then
			vMakerProduct = Left(vMakerProduct, 22) & "..."
		end if

		'---- メーカー名，商品名
		'色規格なし：商品個別へリンク
		'色規格あり：商品一覧へリンク

		'---- 偶数と奇数でタグを変更
		if vRank Mod 2 <> 0 then
			vBGColor = "rank_cat_td1"
		else
			vBGColor = "rank_cat_td2"
		end if

		wHTML = wHTML & "        <div class='" & vBGColor & "'>" & vbNewLine


		if RS("色規格別在庫件数") = 1 then
    		wHTML = wHTML & "          <a href='ProductDetail.asp?Item=" & Server.URLEncode(RS("メーカーコード") & "^" & RS("商品コード") & "^^") & "'>" & vMakerProduct & "</a>" & vbNewLine '2010/05/29 st mod
		else
			'---- 色規格あり商品
			if RS("商品コード") <> "" then
	    		wHTML = wHTML & "          <a href='SearchList.asp?i_type=mp2&s_maker_cd=" & RS("メーカーコード") & "&s_product_cd=" & Server.URLEncode(RS("商品コード")) & "'>" & vMakerProduct & "</a>" & vbNewLine

			'---- シリーズ
			else
	    		wHTML = wHTML & "          <a href='SearchList.asp?i_type=se&sSeriesCd=" & RS("シリーズコード") & "'>" & vMakerProduct & "</a>" & vbNewLine
			end if
		end if

		wHTML = wHTML & "        </div>" & vbNewLine

    If vRank = 10 Then
      '----次のカテゴリーまで読み飛ばし
      Do Until (vBreakKey1 <> vBreakNextKey1)
        RS.MoveNext
        If RS.EOF = true then
          vBreakNextKey1 = "@EOF"
'          vBreakNextKey2 = "@EOF" '2010/05/29 st del
        Else
          vBreakNextKey1 = RS("中カテゴリーコード")
'          vBreakNextKey2 = RS("大カテゴリーコード") '2010/05/29 st del
        End If
      Loop
    Else
      '次のレコード
      RS.MoveNext
      If RS.EOF = true then
        vBreakNextKey1 = "@EOF"
'        vBreakNextKey2 = "@EOF" '2010/05/29 st del
      Else
        vBreakNextKey1 = RS("中カテゴリーコード")
'        vBreakNextKey2 = RS("大カテゴリーコード")  '2010/05/29 st del
      End If
    End If
  Loop

	'---- 10位までない場合、空商品をセット
	for vRank = vRank + 1 to 10

		'---- 偶数と奇数でタグを変更
		if vRank Mod 2 <> 0 then
			vBGColor = "rank_cat_td1"
		else
			vBGColor = "rank_cat_td2"
		end if

		wHTML = wHTML & "       <div class='" & vBGColor & "'>" & vbNewLine

		wHTML = wHTML & "        </div>" & vbNewLine

	next

'---- 大カテゴリーブレーク（中カテゴリータイトル、商品セット)
'	if vBreakKey2 <> vBreakNextKey2 then
'		Do until vCatCount Mod 3 = 0
'			vCatHTML = vCatHTML & "<td width='225' bgcolor='#eeeeee'>　</td>" & vbNewLine
'			for i=1 to 10
'				vProdHTML(i) = vProdHTML(i) & "<td>　</td>" & vbNewLine
'			next
'			vCatCount = vCatCount + 1
'		Loop
'	end if

'---- 3中カテゴリーブレーク（中カテゴリータイトル、商品セット)
	if vCatCount Mod 3 = 0 Then
		wHTML = wHTML & "      </div>" & vbNewLine
		wHTML = wHTML & "    </div>" & vbNewLine
		wHTML = wHTML & "<!-- ranking_container END -->" & vbNewLine
	else
		wHTML = wHTML & "      </div>" & vbNewLine
	end if
Loop


if vCatCount Mod 3 <> 0 Then
wHTML = wHTML & "    </div>" & vbNewLine
wHTML = wHTML & "<!-- ranking_container END -->" & vbNewLine
end if
wHTML = wHTML & "  </div>" & vbNewLine
wHTML = wHTML & "<!-- container END -->" & vbNewLine
wHTML = wHTML & "</div>" & vbNewLine

RS.close

end if


End function

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

	wHTML = wHTML & "<a href='BestSellerList.asp?LargeCategoryCd=" & RS("大カテゴリーコード") & "'>" & RS("大カテゴリー名") & "</a>"

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
<title>売れ筋商品｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/Ranking.css?20120921" type="text/css">
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
      <li class="now">売れ筋商品</li>
    </ul>
  </div></div></div>
  <h1 class="title">売れ筋商品</h1>
-->

<!-- Mainpage START -->
<div id="ranking_key_main_flame">

<!-- Menu START -->
  <div id="ranking_key_top_menu">
    <div class="top_menu_parts">
      <a href="BestSellerList.asp">
      <img src="images/ranking/ts_btn_on.jpg" alt="" name="Image15" width="114" height="80" />
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
      <a href="RankingAccess.asp?RankType=<%=Server.URLEncode("欲しいものリスト")%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image14','','images/ranking/wl_btn_on.jpg',1)"><img src="images/ranking/wl_btn_off.jpg" alt="" name="Image14" width="114" height="80" /></a>
    </div>
    <div class="top_menu_parts">
      <a href="RankingReview.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image16','','images/ranking/nor_btn_on.jpg',1)">
        <img src="images/ranking/nor_btn_off.jpg" alt="" name="Image16" width="114" height="80" />
      </a>
    </div>
    <div class="top_menu_parts">
      <a href="RankingReviewPoint.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image17','','images/ranking/rr_btn_on.jpg',1)">
        <img src="images/ranking/rr_btn_off.jpg" alt="" name="Image17" width="113" height="80" />
      </a>
    </div>
  </div>
<!-- Menu END -->

<!-- 大カテゴリー一覧 -->
<%=wLargeCategoryHTML%>

<!-- 売れ筋ランキング -->
<%=wHTML%>

  </div>
  <div id="globalSide">
    <!--#include file="../Navi/NaviSide.inc"-->
  </div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>