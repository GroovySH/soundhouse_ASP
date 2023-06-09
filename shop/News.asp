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
'	News(商品記事)表示ページ
'
'更新履歴
'2007/01/23 画像と記事本文を分割
'2007/04/20 新NAVI対応
'2007/04/27 大カテゴリー別のときMAX30件に、全件ボタンもつける
'2008/05/07 入力データチェック強化
'2008/05/23 入力データチェック強化
'2008/08/18 記事区分にプレスリリース追加
'2008/08/19 (変更依頼#478)商品記事.中カテゴリー廃止→商品記事中カテゴリーテーブル追加
'2009/04/30 エラー時にerror.aspへ移動
'2009/11/18 an 大カテゴリーに対して書かれた記事も表示するように修正
'2010/08/20 an NewsNo指定有時は<title>を自動表示するように修正
'2010/11/05 an 記事No指定時にmeta keyword,descriptionを追加
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/03/22 if-web ソーシャルボタン(twitter, facebook)自動表示
'2012/05/31 ok #1366 商品記事 取り出しのソート条件を記事日付から記事番号に変更
'2012/07/19 GV #1404 NEWSページデザイン変更
'2012/07/27 GV #1404 納品後の検収にて生じた変更
'2012/08/06 ok デザイン微調整
'========================================================================

On Error Resume Next

Dim NewsNo
Dim NewsDate
Dim NewsDate0		'2012/07/19 GV Add
Dim LargeCategoryCd
Dim CalenderYYYYMM
Dim CalenderYYYYMM0	'2012/07/19 GV Add
'Dim ShowAll		'2012/07/19 GV Del
Dim NewsCategory       '2008/08/18
Dim PageNo		'2012/07/19 GV Add
Dim wPageNo		'2012/07/19 GV Add
Dim wNowPage		'2012/07/19 GV Add

Dim wTitle             '2010/08/20 an add
Dim wMetaKeyword       '2010/11/05 an add
Dim wMetaDescription   '2010/11/05 an add
Dim wLargeCategoryName	'2012/07/19 GV Add
Dim iCnt		'2012/07/19 GV Add
Dim wAddParameter	'2012/07/19 GV Add
Dim wImg

Dim Connection
Dim RS

Dim wSQL
Dim wHTML
Dim xHTML		'2012/07/19 GV Add
Dim wMSG
Dim wErrDesc   '2011/08/01 an add

'========================================================================

'---- Get input data
NewsNo = ReplaceInput(Trim(Request("NewsNo")))
NewsDate = ReplaceInput(Trim(Request("NewsDate")))
LargeCategoryCd = ReplaceInput(Trim(Request("LargeCategoryCd")))
CalenderYYYYMM = ReplaceInput(Request("NaviCalenderYYYYMM"))
'ShowAll = ReplaceInput(Trim(Request("ShowAll")))	'2012/07/19 GV Del
NewsCategory = ReplaceInput(Trim(Request("NewsCategory")))  	'2008/08/18
PageNo = ReplaceInput(Trim(Request("PageNo")))	'2012/07/19 GV Add
  
if NewsNo = "" OR IsNumeric(NewsNo) = false then
	NewsNo = ""
end if

'2012/07/27 GV Mod Start
'if NewsDate = "" OR IsDate(NewsDate) = false then
if NewsDate = "" OR IsNumeric(Replace(NewsDate, "/", "")) = false then
'2012/07/27 GV Mod End
	NewsDate = ""
end if

if CalenderYYYYMM = "" OR IsNumeric(Replace(CalenderYYYYMM, "/", "")) = false then
	CalenderYYYYMM = ""
end if

'2012/07/19 GV Add Start
if PageNo = "" OR IsNumeric(PageNo) = false then
	PageNo = ""
end if
'2012/07/19 GV Add End

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "News.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

if NewsNo <> "" then
	Response.Status="301 Moved Permanently" 
	Response.AddHeader "Location", "http://www.soundhouse.co.jp/news/detail?NewsNo=" & NewsNo
end if

if LargeCategoryCd <> "" then
	Response.Status="301 Moved Permanently" 
	Response.AddHeader "Location", "http://www.soundhouse.co.jp/news/index?LargeCategoryCd=" & LargeCategoryCd
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

Dim wSELECT
Dim wFROM
Dim wWHERE
Dim wWHERE2	'2012/07/19 GV Add
Dim wORDER
Dim vDelStrings
Dim vRegEx
Dim wRecNum	'2012/07/19 GV Add
Dim RSv		'2012/07/19 GV Add
Dim RSx		'2012/07/19 GV Add
Dim vSQL	'2012/07/19 GV Add
Dim xSQL	'2012/07/19 GV Add
Dim wFromPage	'2012/07/19 GV Add
Dim wToPage	'2012/07/19 GV Add
Dim wFromRec	'2012/07/19 GV Add
Dim wToRec	'2012/07/19 GV Add
Dim wPageNum	'2012/07/19 GV Add

wSQL = ""
xSQL = ""	'2012/07/19 GV Add

Const PAGE_COUNT = 7	'2012/07/19 GV Add
Const ITEM_COUNT = 10	'2012/07/19 GV Add


'件数カウント
'2012/07/19 GV Add Start
xSQL = "SELECT COUNT(DISTINCT  a.記事番号)"	'件数カウント用SQL
xSQL = xSQL & " FROM     (商品記事 a WITH (NOLOCK) "
xSQL = xSQL & " LEFT JOIN 商品記事中カテゴリー b WITH (NOLOCK) ON a.記事番号 = b.記事番号) "
xSQL = xSQL & " LEFT JOIN 中カテゴリー c WITH (NOLOCK) on c.中カテゴリーコード = b.中カテゴリーコード "
if NewsDate <> "" then
	xSQL = xSQL & "WHERE  Year(a.記事日付) = " & Year(NewsDate) & " "
	xSQL = xSQL & "  AND Month(a.記事日付) = " & Month(NewsDate) & " "
	xSQL = xSQL & "  AND   Day(a.記事日付) = " & Day(NewsDate) & " "
end if
if LargeCategoryCd <> "" then
	xSQL = xSQL & "WHERE c.大カテゴリーコード = '" & LargeCategoryCd & "' "
	xSQL = xSQL & "   OR a.大カテゴリーコード = '" & LargeCategoryCd & "' " 
end if
if CalenderYYYYMM <> "" then
	CalenderYYYYMM = CalenderYYYYMM & "/1"
	xSQL = xSQL & "WHERE  Year(a.記事日付) = " & Year(CalenderYYYYMM) & " "
	xSQL = xSQL & "  AND Month(a.記事日付) = " & Month(CalenderYYYYMM) & " "
end if

'@@@@response.write(xSQL)
Set RSx = Server.CreateObject("ADODB.Recordset")
RSx.Open xSQL, Connection, adOpenStatic
wRecNum =  RSx("")	'該当件数
wPageNum = Round(wRecNum/10+0.5)	'全ページ数
RSx.Close
'2012/07/19 GV Add End

'---- 商品記事 取り出し
'---- select句
'2012/07/19 GV Del Start
'wSELECT = "SELECT DISTINCT a.記事番号 "

'if NewsNo = "" AND NewsDate = "" AND LargeCategoryCd = "" AND NewsCategory = "" then
'	wSELECT = "SELECT DISTINCT TOP 7 a.記事番号 " 
'end if

'if (LargeCategoryCd <> "" OR NewsCategory <> "") AND ShowAll <> "Y" then
'	wSELECT = "SELECT DISTINCT TOP 30 a.記事番号 "
'end if
'2012/07/19 GV Del End

'2012/07/19 GV Mod Start
'wSELECT = wSELECT & "             , a.記事日付 "
'wSELECT = wSELECT & "             , a.記事タイトル "
'wSELECT = wSELECT & "             , a.記事内容 "
'wSELECT = wSELECT & "             , a.メーカーコード "
'wSELECT = wSELECT & "             , a.商品コード "
'wSELECT = wSELECT & "             , a.記事画像ファイルURL "
if NewsNo <> "" then		'抽出用SQL(記事番号指定)
	wSELECT = "SELECT"
	wSELECT = wSELECT & "  h.記事番号,"
	wSELECT = wSELECT & "  h.記事日付,"
	wSELECT = wSELECT & "  h.記事タイトル,"
	wSELECT = wSELECT & "  h.記事内容,"
	wSELECT = wSELECT & "  h.メーカーコード,"
	wSELECT = wSELECT & "  h.商品コード,"
	wSELECT = wSELECT & "  h.記事画像ファイルURL ,"
	wSELECT = wSELECT & "  h.記事URL ,"
	wSELECT = wSELECT & "  h.中カテゴリー名日本語,"
	wSELECT = wSELECT & "  h.大カテゴリー名,"
	wSELECT = wSELECT & "  h.メーカー名,"
	wSELECT = wSELECT & "  h.商品名"
else				'抽出用SQL(記事番号指定以外)
	wSELECT = "SELECT"
	wSELECT = wSELECT & "  e.記事番号"
	wSELECT = wSELECT & ", e.記事日付"
	wSELECT = wSELECT & ", e.記事タイトル"
	wSELECT = wSELECT & ", e.記事内容"
	wSELECT = wSELECT & ", e.メーカーコード"
	wSELECT = wSELECT & ", e.商品コード"
	wSELECT = wSELECT & ", e.記事画像ファイルURL"
	wSELECT = wSELECT & ", e.記事URL"
end if
'2012/07/19 GV Mod End

'2012/07/19 GV Del Start
'if NewsNo <> "" then    '2010/11/05 an add s
'	wSELECT = wSELECT & "             , d.中カテゴリー名日本語"
'	wSELECT = wSELECT & "             , e.大カテゴリー名 " 
'	wSELECT = wSELECT & "             , f.メーカー名 "
'	wSELECT = wSELECT & "             , g.商品名 "
'end if                  '2010/11/05 an add e
'2012/07/19 GV Del End

'---- from句
'---- where句(サブクエリー)
'2012/07/19 GV Mod Start
'wFROM = wFROM & "         FROM (商品記事 a WITH (NOLOCK) "
'wFROM = wFROM & "            LEFT JOIN 商品記事中カテゴリー b WITH (NOLOCK) "
'wFROM = wFROM & "            ON a.記事番号 = b.記事番号) "
if NewsNo <> "" then		'FROM句(記事番号指定)
	wFROM = " FROM"
	wFROM = wFROM & "  (SELECT *,ROW_NUMBER() OVER(ORDER BY 記事日付 DESC , 記事番号 DESC) AS 行番号"
	wFROM = wFROM & "  FROM"
	wFROM = wFROM & "    (SELECT"
	wFROM = wFROM & "    DISTINCT a.記事番号 ,"
	wFROM = wFROM & "      a.記事日付 ,"
	wFROM = wFROM & "      a.記事タイトル ,"
	wFROM = wFROM & "      a.記事内容 ,"
	wFROM = wFROM & "      a.メーカーコード ,"
	wFROM = wFROM & "      a.商品コード ,"
	wFROM = wFROM & "      a.記事画像ファイルURL,"
	wFROM = wFROM & "      a.記事URL,"
	wFROM = wFROM & "      c.中カテゴリー名日本語,"
	wFROM = wFROM & "      d.大カテゴリー名,"
	wFROM = wFROM & "      e.メーカー名,"
	wFROM = wFROM & "      f.商品名"
	wFROM = wFROM & "        FROM 商品記事 a WITH (NOLOCK)"
	wFROM = wFROM & "            LEFT JOIN 商品記事中カテゴリー b WITH (NOLOCK) ON a.記事番号 = b.記事番号"
	wFROM = wFROM & "            LEFT JOIN 中カテゴリー c WITH (NOLOCK) on c.中カテゴリーコード = a.中カテゴリーコード"
	wFROM = wFROM & "            LEFT JOIN 大カテゴリー d WITH (NOLOCK) on d.大カテゴリーコード = a.大カテゴリーコード"
	wFROM = wFROM & "            LEFT JOIN メーカー e WITH (NOLOCK) on e.メーカーコード = a.メーカーコード"
	wFROM = wFROM & "            LEFT JOIN Web商品 f WITH (NOLOCK) on f.メーカーコード = a.メーカーコード AND f.商品コード = a.商品コード"
	wFROM = wFROM & "        WHERE a.記事番号 =" & NewsNo
else				'FROM句(記事番号指定以外)
	wFROM = " FROM "
	wFROM = wFROM & " (SELECT *,ROW_NUMBER() OVER(ORDER BY 記事日付 DESC , 記事番号 DESC) AS 行番号"
	wFROM = wFROM & "   FROM "
	wFROM = wFROM & "    (SELECT DISTINCT  a.記事番号 , a.記事日付 , a.記事タイトル , a.記事内容 , a.メーカーコード , a.商品コード , a.記事画像ファイルURL , a.記事URL "
	wFROM = wFROM & "      FROM (商品記事 a WITH (NOLOCK) "
	wFROM = wFROM & "        LEFT JOIN 商品記事中カテゴリー b WITH (NOLOCK) ON a.記事番号 = b.記事番号) "
end if

if NewsDate <> "" then		'WHERE句(年月日指定)
	wWHERE = wWHERE & "          WHERE  Year(a.記事日付) = " & Year(NewsDate) & " "
	wWHERE = wWHERE & "            AND Month(a.記事日付) = " & Month(NewsDate) & " "
	wWHERE = wWHERE & "            AND  Day(a.記事日付) = " & Day(NewsDate) & " "
end if

if LargeCategoryCd <> "" then	'FROM句WHERE句(大カテゴリ指定)
	wFROM = wFROM & "              LEFT JOIN 中カテゴリー c WITH (NOLOCK) on c.中カテゴリーコード = b.中カテゴリーコード "
	wWHERE = wWHERE & "          WHERE  c.大カテゴリーコード = '" & LargeCategoryCd & "' "
	wWHERE = wWHERE & "             OR  a.大カテゴリーコード = '" & LargeCategoryCd & "' "
end if

if CalenderYYYYMM <> "" then	'WHERE句(年月指定)
	wWHERE = wWHERE & "          WHERE  Year(a.記事日付) = " & Year(CalenderYYYYMM) & " "
	wWHERE = wWHERE & "            AND Month(a.記事日付) = " & Month(CalenderYYYYMM) & " "
end if

'---- 記事区分=プレスリリースの場合 2008/08/18
if NewsCategory <> "" then
	wWHERE = wWHERE & "        WHERE  a.記事区分 = '" & NewsCategory & "' "
end if

if NewsNo <> "" then		'FROM句(記事番号指定)
	wFROM = wFROM & "    ) g"
	wFROM = wFROM & "  ) h"
else				'WHERE句(記事番号指定以外)
	wWHERE = wWHERE & "  ) d "
	wWHERE = wWHERE & ") e "
end if
'2012/07/19 GV Mod End

'---- where句
'2012/07/19 GV Del Start
'---- 個別記事指定時はメタタグ作成 2010/11/05 an add s
'if NewsNo <> "" then
'	wFROM = wFROM & "            LEFT JOIN 中カテゴリー d WITH (NOLOCK) on d.中カテゴリーコード = a.中カテゴリーコード"
'	wFROM = wFROM & "           	 LEFT JOIN 大カテゴリー e WITH (NOLOCK) on e.大カテゴリーコード = a.大カテゴリーコード"
'	wFROM = wFROM & "           	 	LEFT JOIN メーカー f WITH (NOLOCK) on f.メーカーコード = a.メーカーコード"
'	wFROM = wFROM & "           	 		LEFT JOIN Web商品 g WITH (NOLOCK) on g.メーカーコード = a.メーカーコード AND g.商品コード = a.商品コード"  '2010/11/05 an add e
'	wWHERE = wWHERE & "        WHERE  a.記事番号 = " & NewsNo & " "
'end if

'if NewsDate <> "" then
'	wWHERE = wWHERE & "        WHERE  Year(a.記事日付) = " & Year(NewsDate) & " "
'	wWHERE = wWHERE & "          AND Month(a.記事日付) = " & Month(NewsDate) & " "
'	wWHERE = wWHERE & "          AND Day(a.記事日付) = " & Day(NewsDate) & " "
'end if

'if LargeCategoryCd <> "" then
'	wFROM = wFROM & "            LEFT JOIN 中カテゴリー c WITH (NOLOCK) on c.中カテゴリーコード = b.中カテゴリーコード " '2009/11/18 an 変更
'	wWHERE = wWHERE & "        WHERE  c.大カテゴリーコード = '" & LargeCategoryCd & "' "
'	wWHERE = wWHERE & "           OR  a.大カテゴリーコード = '" & LargeCategoryCd & "'" '2009/11/18 an 追加
'end if

'if CalenderYYYYMM <> "" then
'	CalenderYYYYMM = CalenderYYYYMM & "/1"
'	wWHERE = wWHERE & "        WHERE  Year(a.記事日付) = " & Year(CalenderYYYYMM) & " "
'	wWHERE = wWHERE & "          AND Month(a.記事日付) = " & Month(CalenderYYYYMM) & " "
'end if

'---- 記事区分=プレスリリースの場合 2008/08/18
'if NewsCategory <> "" then
'	wWHERE = wWHERE & "        WHERE  a.記事区分 = '" & NewsCategory & "' "
'end if
'2012/07/19 GV Del End

'2012/07/19 GV Add Start
'---- where句(主クエリー)
if NewsNo = "" then	'記事番号指定以外の場合
	if PageNo = "" then
		wFromRec = 1
		wToRec = ITEM_COUNT
	else
		wToRec = PageNo * ITEM_COUNT
		wFromRec = wToRec - ( ITEM_COUNT - 1 )
	end if
	wWHERE2 = "WHERE e.行番号 BETWEEN " & wFromRec & " AND " & wToRec & " "
end if
'2012/07/19 GV Add End

'---- order句
'2012/07/19 GV Mod Start
'wORDER = wORDER & "     ORDER BY a.記事日付 DESC "				'2012/05/31 ok #1366 Mod
'wORDER = wORDER & "   ,          a.記事番号 DESC "				'2012/05/31 ok #1366 Add
if NewsNo <> "" then	'記事番号指定の場合
	wORDER = wORDER & "     ORDER BY h.記事日付 DESC,"
	wORDER = wORDER & "              h.記事番号 DESC"
else			'記事番号指定以外の場合
	wORDER = wORDER & "     ORDER BY e.記事日付 DESC "				'2012/05/31 ok #1366 Mod
	wORDER = wORDER & "   ,          e.記事番号 DESC "				'2012/05/31 ok #1366 Add
end if
'2012/07/19 GV Mod End

'---- 結合
wSQL = wSELECT & wFROM
'2012/07/19 GV Mod Start
'if wWHERE <> "" then
if wWHERE <> "" or wWHERE2 <> "" then
'	wSQL = wSQL & wWHERE & wORDER
	wSQL = wSQL & wWHERE & wWHERE2 & wORDER
'2012/07/19 GV Mod End
Else
	wSQL = wSQL & wORDER
end if

'@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

wHTML = ""

if RS.EOF = true then
'2012/08/06 ok Mod Start
'	wHTML = wHTML & "  <tr>" & vbNewLine
'	wHTML = wHTML & "    <td class='honbun'>該当Newsが見つかりません。</td>" & vbNewLine
'	wHTML = wHTML & "  </tr>" & vbNewLine
	wHTML = wHTML & "該当Newsが見つかりません。" & vbNewLine
'2012/08/06 ok Mod End
	exit function
end if

'2012/07/19 GV Add Start
'---- Newsカテゴリー取得
vSQL = ""
vSQL = vSQL & "SELECT c.大カテゴリーコード"
vSQL = vSQL & "     , c.大カテゴリー名"
vSQL = vSQL & "     , c.表示順"
vSQL = vSQL & "  FROM 商品記事 a WITH (NOLOCK)"
vSQL = vSQL & "     , 中カテゴリー b WITH (NOLOCK)"
vSQL = vSQL & "     , 大カテゴリー c WITH (NOLOCK)"
vSQL = vSQL & "     , 商品記事中カテゴリー d WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "        d.記事番号 = a.記事番号"
vSQL = vSQL & "   AND b.中カテゴリーコード = d.中カテゴリーコード"
vSQL = vSQL & "   AND c.大カテゴリーコード = b.大カテゴリーコード"
vSQL = vSQL & "   AND c.Web大カテゴリーフラグ = 'Y'"
vSQL = vSQL & " GROUP BY"
vSQL = vSQL & "       c.大カテゴリーコード"
vSQL = vSQL & "     , c.大カテゴリー名"
vSQL = vSQL & "     , c.表示順 "
vSQL = vSQL & " ORDER BY c.表示順"

'@@@@@@@@@@response.write(vSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

if RSv.EOF = true then
	Exit Function
end if

wLargeCategoryName = "最新のニュース"
Do Until RSv.EOF = true
	if LargeCategoryCd = RSv("大カテゴリーコード") then
		wLargeCategoryName = RSv("大カテゴリー名") & "のニュース"
	end if
	RSv.MoveNext
Loop

RSv.close

'---- フッタの編集開始 ----
xHTML = ""

'---- 記事番号指定の場合(前の記事番号と次の記事番号を取得)
if NewsNo <> "" then
	xHTML = xHTML & "<ul class='newsnavi'>" & vbNewLine
	'前の記事番号取り出し
	xSQL = "SELECT TOP 1 記事番号 FROM 商品記事 WHERE (記事番号 < " & NewsNo & " AND 記事日付 = '" & RS("記事日付") & "') OR (記事日付 < '" & RS("記事日付") & "') ORDER BY 記事日付 DESC,記事番号 DESC"
	Set RSx = Server.CreateObject("ADODB.Recordset")
'	response.write(xSQL)
	RSx.Open xSQL, Connection, adOpenStatic
	if RSx.EOF = true then
		xHTML = xHTML & "  <li></li>" & vbNewLine
	else
		xHTML = xHTML & "  <li><a href='News.asp?NewsNo=" & RSx("記事番号") & "'>前のニュースへ</a></li>" & vbNewLine
	end if
	RSx.close
	'次の記事番号取り出し
	xSQL = "SELECT TOP 1 記事番号 FROM 商品記事 WHERE (記事番号 > " & NewsNo & " AND 記事日付 = '" & RS("記事日付") & "') OR (記事日付 > '" & RS("記事日付") & "') ORDER BY 記事日付,記事番号"
	Set RSx = Server.CreateObject("ADODB.Recordset")
'		response.write(xSQL)
	RSx.Open xSQL, Connection, adOpenStatic
	if RSx.EOF = true then
		xHTML = xHTML & "  <li></li>" & vbNewLine
	else
		xHTML = xHTML & "  <li><a href='News.asp?NewsNo=" & RSx("記事番号") & "'>次のニュースへ</a></li>" & vbNewLine
	end if
	RSx.close
	xHTML = xHTML & "</ul>" & vbNewLine
'---- その他の場合
else
	'１ページ表示件数を超えた場合
	if wRecNum > ITEM_COUNT then
		xHTML = xHTML & "<div class='pagenavi_box'>" & vbNewLine
		xHTML = xHTML & "  <ol class='pagenavi'>" & vbNewLine
		'ページ番号がない場合
		if PageNo = "" then
			wFromPage = 1
			wNowPage = 1
			wToPage = PAGE_COUNT
		'ページ番号が中央ページ数より小さい場合
		elseIf Int(PageNo) <= Int(PAGE_COUNT / 2) then
			wFromPage = 1
			wNowPage = PageNo
			wToPage = PAGE_COUNT
		'ページ番号が中央ページ数より大きい場合
		elseIf (wPageNum - PageNo) <= Int(PAGE_COUNT / 2) then
			wFromPage = wPageNum - PAGE_COUNT + 1
			wNowPage = PageNo
			wToPage = wPageNum
		else
			wNowPage = PageNo
			wFromPage = PageNo - Int(PAGE_COUNT / 2)
			wToPage = PageNo + Int((PAGE_COUNT - 1) / 2)
		end if
		
		'ページ開始終了補正
		if wFromPage < 1 then
			wFromPage = 1
		end if
		if wToPage > wPageNum then
			wToPage = wPageNum
		end if

'		'ToPage補正
'		if wRecNum mod c1Page > 0 then
'			if wToPage > INT( wRecNum / c1Page ) + 1 then
'				wToPage = INT( wRecNum / c1Page ) + 1
'			end if
'		else
'			if wToPage > INT( wRecNum / c1Page ) then
'				wToPage = INT( wRecNum / c1Page )
'			end if
'		end if
'		if wRecNum mod c1Page > 0 then
'			if wToPage > INT( wRecNum / c1Page ) + 1 then
'				wToPage = INT( wRecNum / c1Page ) + 1
'			end if
'		else
'			if wToPage > INT( wRecNum / c1Page ) then
'				wToPage = INT( wRecNum / c1Page )
'			end if
'		end if
		'追加するパラメータの編集
		wAddParameter = ""
		if NewsDate <> "" then
			wAddParameter = "&NewsDate=" & NewsDate
		elseif LargeCategoryCd <> "" then
			wAddParameter = "&LargeCategoryCd=" & LargeCategoryCd
		elseif CalenderYYYYMM <> "" then
			wAddParameter = "&NaviCalenderYYYYMM=" & CalenderYYYYMM
			if Right(wAddParameter,2) = "/1" then
				wAddParameter = Left(wAddParameter,Len(wAddParameter)-2)
			end if
		else
		end if
		'前へ
		if wNowPage <> 1 then
			xHTML = xHTML & "    <li class='back'><a href='News.Asp?PageNo=" & wNowPage - 1 & wAddParameter & "'>前へ</a></li>" & vbNewLine
		end if
		'フッタページインデックス作成
		for iCnt = wFromPage to wToPage
			if iCnt = INT( wNowPage ) then
				xHTML = xHTML & "    <li><span class='now'>" & iCnt & "</span></li>" & vbNewLine
			else
				xHTML = xHTML & "    <li><a href='News.Asp?PageNo=" & iCnt & wAddParameter & "'>" & iCnt & "</a></li>" & vbNewLine
			end if
		next
		'次へ
		if INT( wNowPage ) <> INT( wToPage ) then
			xHTML = xHTML & "    <li class='next'><a href='News.Asp?PageNo=" & wNowPage + 1 & wAddParameter & "'>次へ</a></li>" & vbNewLine
		end if
		xHTML = xHTML & "  </ol>" & vbNewLine
'		xHTML = xHTML & "<span class='page'>" & wPageNum & "ページ中" & wNowPage & "ページ</span>" & vbNewLine		'2012/08/06 ok Del
		xHTML = xHTML & "</div>" & vbNewLine
	end if
end if
'2012/07/19 GV Add End

'---- title,metatag用データ確保
if NewsNo <> "" then  '2010/08/20 an add s
	'---- title
	wTitle = RS("記事タイトル")
	
	'---- keyword
	wMetaKeyword = RS("大カテゴリー名")  '2010/11/05 an add s
	
	if RS("中カテゴリー名日本語") <> "" then
		wMetaKeyword = wMetaKeyword & "," & RS("中カテゴリー名日本語")
	end if
	
	if RS("メーカー名") <> "" then
		wMetaKeyword = wMetaKeyword & "," & RS("メーカー名")
	end if
	
	if RS("商品名") <> "" then
		wMetaKeyword = wMetaKeyword & "," & RS("商品名")
	end if
	'---- 余計な先頭の","があれば削除
	if Left(wMetaKeyword,1) = "," then
		wMetaKeyword =  Mid(wMetaKeyword, 2)
	end if
	
	'---- description
	if RS("記事内容") <> "" then
		wMetaDescription = fDeleteHTMLTag(RS("記事内容")) 'HTML
		wMetaDescription = replace(replace(replace(wMetaDescription, vbCr, ""), vbLf, ""), vbTab, "") '改行、Tabの削除
			
		if Len(wMetaDescription) > 100 then
			wMetaDescription = Left(wMetaDescription, 97) & "..."
		else
			wMetaDescription = Left(wMetaDescription, 100)
		end if
	end if            '2010/11/05 an add  e
	
	wImg = ""
	If RS("記事画像ファイルURL") <> "" Then
		wImg = RS("記事画像ファイルURL")
		If InStr(wImg, "http") > 0 Then
		Else
			If InStr(wImg, "../") > 0 Then
				wImg = g_HTTP & Replace(wImg, "../", "")
			Else
				wImg = g_HTTP & wImg
			End If
		End If
	End If

'2012/07/19 GV Add Start
elseif NewsDate <> "" Then
	NewsDate0 = Year( NewsDate ) & "/"
	if Len( Month( NewsDate ) ) = 1 then	'ゼロパディング
		NewsDate0 = NewsDate0 & "0" & Month( NewsDate ) & "/"
	else
		NewsDate0 = NewsDate0 & Month( NewsDate ) & "/"
	end if
	if Len( Day( NewsDate ) ) = 1 then	'ゼロパディング
		NewsDate0 = NewsDate0 & "0" & Day( NewsDate )
	else
		NewsDate0 = NewsDate0 & Day( NewsDate )
	end if
	wTitle = NewsDate0 & "のニュース"
elseif CalenderYYYYMM <> "" then
	CalenderYYYYMM0 = Year( CalenderYYYYMM ) & "/"
	if Len( Month( CalenderYYYYMM ) ) = 1 then	'ゼロパディング
		CalenderYYYYMM0 = CalenderYYYYMM0 & "0" & Month( CalenderYYYYMM )
	else
		CalenderYYYYMM0 = CalenderYYYYMM0 & Month( CalenderYYYYMM )
	end if
	wTitle = CalenderYYYYMM0 & "のニュース"
else
	wTitle = wLargeCategoryName
'2012/07/19 GV Add End
end if                '2010/08/20 an add e


'2012/07/19 GV Add Start
'パン屑リスト
wHTML = wHTML & "    <div id='path_box'><div id='path_box_inner01'><div id='path_box_inner02'>" & vbNewLine
wHTML = wHTML & "      <p class='home'><a href='../'><img src='../images/icon_home.gif' alt='HOME'></a></p>" & vbNewLine
wHTML = wHTML & "      <ul id='path'>" & vbNewLine
wHTML = wHTML & "        <li><a href='News.asp'>ニュース記事一覧</a></li>" & vbNewLine
wHTML = wHTML & "        <li class='now'>" & wTitle & "</li>" & vbNewLine
wHTML = wHTML & "      </ul>" & vbNewLine
wHTML = wHTML & "    </div></div></div>" & vbNewLine
'2012/07/19 GV Add End

'2012/07/19 GV Add Start
'h1タイトル
if NewsNo <> "" then			'記事番号指定の場合
elseif NewsDate <> "" then		'年月日指定の場合
	wHTML = wHTML & "    <h1 class='title'>" & NewsDate0 & "のニュース</h1>" & vbNewLine
elseif CalenderYYYYMM <> "" then	'年月指定の場合
	wHTML = wHTML & "    <h1 class='title'>" & CalenderYYYYMM0 & "のニュース</h1>" & vbNewLine
else					'最新のニュース、または、大カテゴリ指定の場合
	wHTML = wHTML & "    <h1 class='title'>" & wLargeCategoryName & "</h1>" & vbNewLine
end if
'2012/07/19 GV Add End

'クラス定義
wHTML = wHTML & "    <ul class='article'>" & vbNewLine	'2012/07/19 GV Add

Do until RS.EOF = true

	'2012/07/19 GV Mod Start
'	if LargeCategoryCd = "" AND NewsCategory = "" then
'		wHTML = wHTML & "  <tr>" & vbNewLine
'		wHTML = wHTML & "    <td class='honbun'>" & vbNewLine
'		wHTML = wHTML & "      <h2>" & RS("記事タイトル") & "</h2>　" & fFormatDate(RS("記事日付"))
'		wHTML = wHTML & "    </td>" & vbNewLine
'		wHTML = wHTML & "  </tr>" & vbNewLine
'		wHTML = wHTML & "  <tr>" & vbNewLine
'		wHTML = wHTML & "    <td class='honbun' style='padding:10px 0px'>" & vbNewLine
		wHTML = wHTML & "      <li>" & vbNewLine
		wHTML = wHTML & "        <h2 class='subject'><a href='News.asp?NewsNo=" & RS("記事番号") & "'>" & RS("記事タイトル") & "</a></h2>" & vbNewLine
	'2012/07/19 GV Mod End
		if RS("記事画像ファイルURL") <> "" then
			'2013/05/22 if-web mod s
			If RS("記事URL") <> "" Then
				wHTML = wHTML & "        <a href='" & RS("記事URL") & "'><img src='" & RS("記事画像ファイルURL") & "' alt='" & RS("記事タイトル") & "' class='opover'></a>" & vbNewLine
			Else
				'2012/07/19 GV Mod Start
	'			wHTML = wHTML & "<img src='" & RS("記事画像ファイルURL") & "' width='200' border='0' align='left' style='MARGIN: 0px 5px 5px 0px' alt='" & RS("記事タイトル") & "'>"
				wHTML = wHTML & "        <img src='" & RS("記事画像ファイルURL") & "' alt='" & RS("記事タイトル") & "'>" & vbNewLine
				'2012/07/19 GV Mod End
			End If
			'2013/05/22 if-web mod e
		end if
		wHTML = wHTML & "        <p class='date'>" & fFormatDate(RS("記事日付")) & "</p>" & vbNewLine	'2012/07/19 GV Add

		if ISNULL(RS("記事内容")) = false then
			'2012/07/19 GV Mod Start
'			wHTML = wHTML & Replace(RS("記事内容"), vbNewLine, "<br>") & vbNewLine
			wHTML = wHTML & "        <p>" & Replace(RS("記事内容"), vbNewLine, "<br>") & "</p>" & vbNewLine
			'2012/07/19 GV Mod End
'2012/03/22 if-web add start
			'2012/07/19 GV Mod Start
'			wHTML = wHTML & "      <ul class='news_smbtn'>" & vbNewLine
'			wHTML = wHTML & "        <li><a href='http://twitter.com/share' class='twitter-share-button' data-url='http://www.soundhouse.co.jp/shop/News.asp?NewsNo=" & RS("記事番号") & "' data-text='" & RS("記事タイトル") & "' data-count='horizontal' data-via='soundhouse_jp' data-lang='ja'>Tweet</a></li>" & vbNewLine
'			wHTML = wHTML & "        <li><a name='fb_share' share_url='http://www.soundhouse.co.jp/shop/News.asp?NewsNo=" & RS("記事番号") & "'>シェアする</a></li>" & vbNewLine
'			wHTML = wHTML & "      </ul>" & vbNewLine
			wHTML = wHTML & "        <ul class='sns'>" & vbNewLine
			wHTML = wHTML & "          <li><a href='http://twitter.com/share' class='twitter-share-button' data-url='http://www.soundhouse.co.jp/shop/News.asp?NewsNo=" & RS("記事番号") & "' data-text='" & RS("記事タイトル") & "' data-count='horizontal' data-via='soundhouse_jp' data-lang='ja'>Tweet</a></li>" & vbNewLine
			wHTML = wHTML & "          <li><iframe src='//www.facebook.com/plugins/like.php?href=http%3A%2F%2Fwww.soundhouse.co.jp%2Fshop%2FNews.asp%3FNewsNo%3D" & RS("記事番号") & "&amp;send=false&amp;layout=button_count&amp;width=100&amp;show_faces=false&amp;action=like&amp;colorscheme=light&amp;font&amp;height=21&amp;appId=191447484218062' scrolling='no' frameborder='0' style='border:none; overflow:hidden; width:110px; height:21px;' allowTransparency='true'></iframe></li>" & vbNewLine
			wHTML = wHTML & "        </ul>" & vbNewLine
			'2012/07/19 GV Mod End
'2012/03/22 if-web add end
		end if

		'2012/07/19 GV Del Start
'		wHTML = wHTML & "    </td>" & vbNewLine
'		wHTML = wHTML & "  </tr>" & vbNewLine

'		wHTML = wHTML & "  <tr>" & vbNewLine
'		wHTML = wHTML & "    <td colSpan='5' height='5'><hr size='1'></td>" & vbNewLine
'		wHTML = wHTML & "  </tr>" & vbNewLine
'		'2012/07/19 GV Del End

	'2012/07/19 GV Del Start
'	else
'		wHTML = wHTML & "  <tr>" & vbNewLine
'		wHTML = wHTML & "    <td class='honbun'>" & vbNewLine
'		wHTML = wHTML & fFormatDate(RS("記事日付")) & " <a href='News.asp?NewsNo=" & RS("記事番号") & "' class='link'>" & RS("記事タイトル") & "</a>"
'		wHTML = wHTML & "    </td>" & vbNewLine
'		wHTML = wHTML & "  </tr>" & vbNewLine
'	end if
	'2012/07/19 GV Del End
	wHTML = wHTML & "      </li>" & vbNewLine
	RS.MoveNext
Loop
wHTML = wHTML & "    </ul>" & vbNewLine	'2012/07/19 GV Add

'---- 『このカテゴリーの記事を全て表示する』のURL作成
'2012/07/19 GV Del Start
'if ShowAll <> "Y" then
	'---- 大カテゴリーコードの場合
'	if LargeCategoryCd <> "" then
'		wHTML = wHTML & "  <tr>" & vbNewLine
'		wHTML = wHTML & "    <td class='honbun'><br><a href='News.asp?LargeCategoryCd=" & LargeCategoryCd & "&ShowAll=Y' class='link'><b>このカテゴリーの記事を全て表示する</b></a></td>" & vbNewLine
'		wHTML = wHTML & "  </tr>" & vbNewLine
'		exit function
'	end if
	'---- 記事区分が一般記事、個別記事以外の場合 2008/08/18
'	if NewsCategory <> "" then
'		wHTML = wHTML & "  <tr>" & vbNewLine
'		wHTML = wHTML & "    <td class='honbun'><br><a href='News.asp?NewsCategory=" & NewsCategory & "&ShowAll=Y' class='link'><b>このカテゴリーの記事を全て表示する</b></a></td>" & vbNewLine
'		wHTML = wHTML & "  </tr>" & vbNewLine
'		exit function
'	end if
'end if
'2012/07/19 GV Del End

RS.Close

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
<html xmlns:og="http://ogp.me/ns#">
<head>
<meta charset="Shift_JIS">
<title><%=wTitle%>｜サウンドハウス</title>
<% if wMetaDescription <> "" then %>
<meta name="description" content="<%=wMetaDescription%>">
<% end if%>
<% if wMetaKeyword <> "" then %>
<meta name="keywords" content="<%=wMetaKeyword%>">
<% end if%>
<% If NewsNo <> "" Then %>
<meta name="twitter:card" content="summary">
<meta name="twitter:site" content="@soundhouse_jp">
<meta property="og:title" content="<%=wTitle%>">
<meta property="og:type" content="article">
<meta property="og:description" content="<%=wMetaDescription%>">
<% If wImg <> "" Then %>
<meta property="og:image" content="<%=wImg%>">
<% End If %>
<meta property="og:url" content="<%=g_HTTP%>shop/News.asp?NewsNo=<%=NewsNo%>">
<% End If %>
<!--#include file="../Navi/NaviStyle.inc"-->
<link href="style/news.css?20140618" rel="stylesheet" type="text/css">

</head>

<body>
<!--#include file="../Navi/Navitop.inc"-->

<div id="globalMain">
	<span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>
	<!-- コンテンツstart -->
	<div id="globalContents">
<!-- ニュース -->
<%=wHTML%>
<%=xHTML%>
	<!--/#contents --></div>
	<div id="globalSide">
<!--#include file="../Navi/NaviLeftNews.inc"-->
<!--#include file="../Navi/NaviSide.inc"-->
	<!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript" src="http://platform.twitter.com/widgets.js" charset="utf-8"></script>
</body>
</html>
<%
call close_db()
%>
