<%@ LANGUAGE="VBScript" %>
<%
'ネットハウスねっとハウスネットはうす
 Option Explicit
%>
<!--#include file="../common/ADOVBS.inc"-->
<!--#include file="../common/system_common.inc"-->
<!--#include file="../common/shop_common_functions.inc"-->
<!--#include file="../common/bfunctions1.asp"-->
<!--#include file="../common/HttpsSecurity.inc"-->
<%
'========================================================================
'
'	オーダー処理中ページ
'
'2006/07/05 カードオーソリにBlueGateを呼び出し
'2006/07/19 3D呼び出し追加
'2006/10/24 コンビニ決済時eContext呼び出し追加
'2008/05/13 クロスサイトリクエストフォジェリー対策 Keyパラメータチェック
'2008/05/14 HTTPSチェック対応
'2008/10/13 新カード入力対応　（元の認証に戻す3D+Auth）
'2010/03/15 hn カード呼び出しコメントアウト(Error.aspへ）
'2012/02/15 GV セキュリティーキーの同一チェックを共通プロシージャで実行するよう変更 (セッション変数のSkeyとの比較を止め、セッションデータテーブル内の Skey との比較)
'2013/12/04 GV バリデーションチェックを追加
'2014/08/14 GV 翌日配送条件変更対応
'
'========================================================================
'2013/12/04 GV add start ---
On Error Resume Next
Response.Expires = -1			' Do not cache
Response.buffer = true

Dim wMSG
Dim vRS
Dim wDateTime
Dim wDate

'---- 届先情報
Dim wAfter13FL							'13時以降
Dim wAfter14FL							'14時以降
Dim wAfter15FL							'15時以降
Dim wAfter16FL							'16時以降
'Dim w94HokkaidouRitouFL				'配送先が　九州・四国・北海道・離島   2011/06/29 an del
Dim w94ChugokuHokkaidouFL				'配送先が　九州・四国・中国・北海道   2011/06/29 an add
Dim wHolidayFL							'日曜祭日
Dim wAvailableDate						'指定可能日
Dim wSatFl								'土曜日
Dim wFriFL								'金曜日
Dim wRitouFl							'離島フラグ
Dim wShipPrefecture						'届先都道府県
Dim wDeliveryDt							'指定納期
'2013/12/04 GV add end -----

Dim nextURL
Dim OrderTotalAm
Dim payment_method

OrderTotalAm = Trim(ReplaceInput(Request("OrderTotalAm")))
payment_method = Trim(ReplaceInput(Request("payment_method")))

'2013/12/04 GV add start ---
wDateTime = Now()
wDate = Date()
'wDateTime = "2013-12-06 1:10:30"
'wDate = "2013-12-06"

'---- DB
Dim Connection

Const w9Shuu4KokuChugokuHokkaido = "福岡県,長崎県,佐賀県,大分県,熊本県,宮崎県,鹿児島県,香川県,徳島県,愛媛県,高知県,鳥取県,岡山県,島根県,広島県,山口県,北海道"   '2011/06/29 an add
Const cAddDaysToNyukaYoteibi = 2
'2013/12/04 GV add end -----

'---- セキュリティーキーチェック
' 2012/02/15 GV Mod Start
'If Session("SKey") <> ReplaceInput(Request("SKey")) Then
If isLegalSecureKey(ReplaceInput(Request("SKey"))) = False Then
' 2012/02/15 GV Mod End
	Response.redirect "OrderInfoEnter.asp"
End If

'2013/12/04 GV add start ---
Call connect_db()

'---- 入力データーのチェック
Call validate_data()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "OrderProcessing.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

Call close_db()

If Err.Description <> "" Then
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

'---- エラーが無いときは処理続行、エラーがあれば注文内容指定ページへ
If wMSG = "" Then
Else
	Session("msg") = wMSG
	Server.Transfer "OrderInfoEnter.asp"
End If
'2013/12/04 GV add end -----

Session("BlueGate3DReturnCode") = ""		'BlueGateリターンコードクリア

If OrderTotalAm <> "0" Then
	If payment_method = "クレジットカード" Then
	''''	nextURL = "OrderCardAuthBG.asp"		'オーソリのみ取得
		''''nextURL = "OrderCard3DSecureBG.asp"		'3D+オーソリ取得
		''''Session("受注合計金額") = OrderTotalAm
		''''nextURL = "OrderCard3DAuthSendBG.asp"		'カード入力+3D+オーソリ取得
		''''nextURL = "OrderCard3DSecureBG2.asp"		'3D+オーソリ取得 NEW
		nextURL = "Error.asp"

	Else
		If payment_method = "コンビニ支払" Then
			nextURL = "OrderEcontext.asp"
		Else
			nextURL = "OrderSubmit.asp"
		End If
	End If
Else
	nextURL = "OrderSubmit.asp"
End If

'2013/12/04 GV add start ---
'========================================================================
'
'	Function	入力データーのチェック
'	OrderInfoInsert.asp の同名関数の、配送日部分のみ実装。
'	注文時間帯変更の場合は、このファイルも変更すること。
'
'========================================================================
Function validate_data()

Dim vDateTimeMsg
Dim vSQL

'---- 仮受注Recordset取り出し
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    *"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    仮受注"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, Connection, adOpenStatic, adLockOptimistic

If vRS.EOF = True Then
	wMSG = "カート情報がありません。"
	Exit Function
End If

'届先都道府県
wShipPrefecture = vRS("届先都道府県")

'指定納期
wDeliveryDt = vRS("指定納期")

'---- 離島フラグの設定
Call setRitouFlag(vRS("届先郵便番号"))

'---- 配達日指定(翌日指定時のチェック)
If vRS("指定納期") <> "" Then

	'---- 配達日指定可能かチェック（注残品で入荷予定がないか、入荷予定+2日より前は指定NG）
	If checkNyukaYoteibi() = False Then
		wMSG = wMSG & "在庫のない商品がご注文に含まれているため、配達希望日の指定はできません。<br>"
	Else
		'---- ローンの場合、配送日指定不可
		If vRS("支払方法") = "ローン" Then
			wMSG = wMSG & "お支払い方法がローンの場合、配達希望日の指定はできません。<br>"
		Else
			wAfter13FL = False
			wAfter14FL = False
			wAfter15FL = False
			wAfter16FL = False
			'w94HokkaidouRitouFL = False    '2011/06/29 an del
			w94ChugokuHokkaidouFL = False   '2011/06/29 an add
			wHolidayFL = False
			wSatFL = False
			wFriFL = False

			'---- 13時以降かどうかチェック
			If DatePart("h", wDateTime) >= 13 Then
				wAfter13FL = True
			End If

			'---- 14時以降かどうかチェック
			If DatePart("h", wDateTime) >= 14 Then
				wAfter14FL = True
			End If

			'---- 15時以降かどうかチェック
			If DatePart("h", wDateTime) >= 15 Then
				wAfter15FL = True
			End If

			'---- 16時以降かどうかチェック
			If DatePart("h", wDateTime) >= 16 Then
				wAfter16FL = True
			End If

			'---- 今日が日曜･祭日かどうかチェック
			If (DatePart("w",wDate) = vbSunday) Or (checkHoliday(Date()) = True) Then
				wHolidayFL = True
			End If

			'---- 今日が土曜日かどうかチェック
			If DatePart("w",wDate) = vbSaturday Then
				wSatFL = True
			End If

			'---- 今日が金曜日かどうかチェック
			If DatePart("w",wDate) = vbFriday Then
				wFriFL = True
			End If

			'---- 配送先が　九州・四国・中国・北海道かどうかチェック  
			'If wRitouFl = "Y" Or Instr(w9Shuu4KokuHokkaido, wShipPrefecture) > 0 Then
			if Instr(w9Shuu4KokuChugokuHokkaido, vRS("届先都道府県")) > 0 Then
				'w94HokkaidouRitouFL = True     '2011/06/29 an del
				w94ChugokuHokkaidouFL = True    '2011/06/29 an add
			End If

			'---- 指定可能日をセット
			wAvailableDate = setAvailableDate()

			'---- 配送日チェック
			If (DateDiff("d", wDeliveryDt, wAvailableDate) > 0) Then

				'2011/02/16 hn mod s
				'---- 休日
				If wHolidayFl = True Then
					vDateTimeMsg = "休日の"
				Else
					'---- コンビニ 土曜日 13時以降
					If payment_method = "コンビニ支払" AND wSatFl = True AND wAfter13Fl = true Then
						vDateTimeMsg = "コンビニ支払の土曜日13時以降の"
					End If

					'---- コンビニ 平日 14時以降
					If payment_method = "コンビニ支払" AND wSatFl = False AND wAfter14Fl = true Then
						vDateTimeMsg = "コンビニ支払の14時以降の"
					End If

					'---- 銀行振込 14時以降
					If payment_method = "銀行振込" AND wAfter14Fl = true Then
						vDateTimeMsg = "銀行振込の14時以降の"
					End If

					'---- 銀行振込 土曜日
					If payment_method = "銀行振込" AND wSatFL = true Then
						vDateTimeMsg = "銀行振込の土曜日の"
					End If

					'---- 代引き・カード 16時以降
					If (payment_method = "代引き" OR payment_method = "クレジットカード") AND wAfter16Fl = true Then
						'vDateTimeMsg = payment_method & "の16時以降の"	'2014/08/14 GV comment out
						vDateTimeMsg = payment_method & "の15時以降の"	'2014/08/14 GV add
					End If
				End If

				'---- 九州･四国･中国・北海道
				'If w94HokkaidouRitouFL = True Then                             '2011/06/29 an del
					'wMSG = wMSG & "お届け先が九州・四国・北海道・離島で、"     '2011/06/29 an del
				If w94ChugokuHokkaidouFL = True Then                            '2011/06/29 an add
					wMSG = wMSG & "お届け先が九州・四国・中国・北海道で、"      '2011/06/29 an add
				End If
				
				'---- 離島                      '2011/06/29 an add s
				If wRitouFl = "Y" Then
					wMSG = wMSG & "お届け先が沖縄・離島で、"
				End If                          '2011/06/29 an add e

				wMSG = wMSG & vDateTimeMsg & "配送日指定は、" & wAvailableDate & "以降を指定してください。<br>"
				'2011/02/16 hn mod e

			End If

			'---- 60日以内
			'2013/02/18 GV #1525 MOD START
'			If (DateDiff("d", DateAdd("d", 60,wDate), wDeliveryDt) > 0) Then
'				wMSG = wMSG & "配送日指定は60日以内の日付を指定してください。<br>"
			If (payment_method = "代引き" AND DateDiff("d", DateAdd("d", 10,wDate), wDeliveryDt) > 0) Then
				wMSG = wMSG & "配送日指定は10日以内の日付を指定してください。<br>"
			ElseIf (payment_method <> "代引き" AND DateDiff("d", DateAdd("d", 60,wDate), wDeliveryDt) > 0) Then
				wMSG = wMSG & "配送日指定は60日以内の日付を指定してください。<br>"
			'2013/02/18 GV #1525 MOD END
			End If
		End If
	End If
End If

If wMSG <> "" Then
	wMSG = "<b>以下の入力エラーを訂正してください。</b><br><br>" & wMSG
End If

End Function

'========================================================================
'
'	Function	注残の有無と入荷予定日を確認し、配達日指定可能かチェック
'
'		return:	 配達日指定可能なら True
'                配達日指定不可なら False
'
'========================================================================
Function checkNyukaYoteibi()

Dim RSv
Dim vSQL
Dim vHikiateKanouQt
Dim vSetCount
Dim vMaxNyukaYoteibi

checkNyukaYoteibi = True

vHikiateKanouQt = ""
vSetCount = ""
vMaxNyukaYoteibi =""

'---- 仮受注明細取り出し
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "     a.メーカーコード"
vSQL = vSQL & "   , a.商品コード"
vSQL = vSQL & "   , a.色"
vSQL = vSQL & "   , a.規格"
vSQL = vSQL & "   , a.受注数量"
vSQL = vSQL & "   , b.引当可能数量"
vSQL = vSQL & "   , b.引当可能入荷予定日"
vSQL = vSQL & "   , c.セット商品フラグ"
vSQL = vSQL & " FROM"
vSQL = vSQL & "     仮受注明細 a WITH (NOLOCK)"
vSQL = vSQL & "   , Web色規格別在庫 b WITH (NOLOCK)"
vSQL = vSQL & "   , Web商品 c WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "         b.メーカーコード = a.メーカーコード"
vSQL = vSQL & "     AND b.商品コード = a.商品コード"
vSQL = vSQL & "     AND b.色 = a.色"
vSQL = vSQL & "     AND b.規格 = a.規格"
vSQL = vSQL & "     AND c.メーカーコード = a.メーカーコード"
vSQL = vSQL & "     AND c.商品コード = a.商品コード"
vSQL = vSQL & "     AND b.終了日 IS NULL"
vSQL = vSQL & "     AND a.SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

If RSv.EOF = False Then

	'---- 全仮受注明細商品をチェック
	Do Until RSv.EOF = True

		'セット品の場合はセット品全体の在庫数量、入荷予定日をチェックする
		If RSv("セット商品フラグ") = "Y" Then

			'---- セット品在庫数量、MAX引当可能予定取得
			Call GetSetCount(RSv("メーカーコード"), RSv("商品コード"), RSv("色"), RSv("規格"), vSetCount, vMaxNyukaYoteibi)

			If vMaxNyukaYoteibi <wDate Then			'入荷予定のないセット品の場合は2000/01/01が入っている
				vMaxNyukaYoteibi = ""
			Else
				vMaxNyukaYoteibi = vMaxNyukaYoteibi		'セット品全体のMAX入荷予定日で上書き
			End If

			vHikiateKanouQt = vSetCount					'セット品全体のMIN在庫数量で上書き

		Else
			vHikiateKanouQt = RSv("引当可能数量")
			vMaxNyukaYoteibi = RSv("引当可能入荷予定日")
		End If

		If RSv("受注数量")  > vHikiateKanouQt Then		'注残品の場合は入荷予定日をチェック
			'---- 入荷予定のない商品が1つでもある場合は配達日指定不可
			If IsNULL(vMaxNyukaYoteibi) = True Or vMaxNyukaYoteibi = "" Then
				checkNyukaYoteibi = False
				Exit Do
			Else
				'---- 指定可能日（入荷予定日+2以降）
				wAvailableDate = cf_FormatDate(DateAdd("d", cAddDaysToNyukaYoteibi, vMaxNyukaYoteibi), "YYYY/MM/DD")

				'---- 入荷が間に合わない商品が1つでもある場合は配達日指定不可
				If wDeliveryDt < wAvailableDate Then
					checkNyukaYoteibi = False
					Exit Do
				End If
			End If
		End If

		RSv.MoveNext
	Loop
End If

RSv.Close

End Function

'========================================================================
'
'	Function	祭日チェック
'
'		input : チェックする日(YYYY/MM/DD)
'		return:	祭日なら　True
'				祭日以外　False
'
'========================================================================
Function checkHoliday(p_date)

Dim RSv
Dim vSQL

'---- 祭日チェック
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    年月日"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    カレンダー WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "        年月日 = '" & p_date & "'"
vSQL = vSQL & "    AND 休日フラグ = 'Y'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

If RSv.EOF = True Then
	checkHoliday  = False
Else
	checkHoliday = True
End If

RSv.Close

End Function

'========================================================================
'
'	Function	配送可能日セット    2010/12/17 an mod  2011/02/16 hn mod
'
'		return:	配送可能日 (YYYY/MM/DD)
'
'   支払方法＝コンビニ
'     土曜日以外
'     14時以前：受付日＝当日
'     14時以降：受付日＝翌日

'     土曜日
'     13時以前：受付日＝当日
'     13時以降：受付日＝翌日
'
'   支払方法＝銀行振込（曜日関係なし）
'     14時以前：受付日＝当日
'     14時以降：受付日＝翌日
'
'   支払方法＝代引き（曜日関係なし）
'     16時以前：受付日＝当日
'     16時以降：受付日＝翌日
'
'   支払方法＝ローン
'     納期指定不可
'
'==============================
'   受付日が、日曜、祝日、（銀行振込の場合は土曜日も含む）の場合は次の営業日を受付日とする。
'
'   納期指定可能日は、
'   受付日+1日　（下記以外）
'   受付日+2日　（九州、四国、中国、北海道）
'   受付日+5日　（沖縄、離島）
'
'========================================================================
'
Function setAvailableDate()

Dim vOrderDate

vOrderDate = cf_FormatDate(wDate, "YYYY/MM/DD")

'---- コンビニ支払：土曜日以外　14時以降は受付日は翌日
if vRS("支払方法") = "コンビニ支払" AND wSatFl = false AND wAfter14Fl = true then
	vOrderDate = cf_FormatDate(DateAdd("d", 1,wDate), "YYYY/MM/DD")
end if

'---- コンビニ支払：土曜日　13時以降は受付日は翌日
if vRS("支払方法") = "コンビニ支払" AND wSatFl = true AND wAfter13Fl = true then
	vOrderDate = cf_FormatDate(DateAdd("d", 1,wDate), "YYYY/MM/DD")
end if

'---- 銀行振込：14時以降は受付日は翌日
if vRS("支払方法") = "銀行振込" AND wAfter14Fl = true then
	vOrderDate = cf_FormatDate(DateAdd("d", 1,wDate), "YYYY/MM/DD")
end if

'2014/08/14 GV mod start
'15時移行に変更
'---- 代引き・カード：16時以降は受付日は翌日
'If (vRS("支払方法") = "代引き" OR vRS("支払方法") = "クレジットカード") AND wAfter16Fl = true Then
If (vRS("支払方法") = "代引き" OR vRS("支払方法") = "クレジットカード") AND wAfter15FL = true Then
	vOrderDate = cf_FormatDate(DateAdd("d", 1,wDate), "YYYY/MM/DD")
end if
'2014/08/14 GV mod end

'---- 受付日が、日曜、祝日、（銀行振込の場合は土曜日も含む）の場合は次の営業日を受付日とする。
Do
	if  (DatePart("w", vOrderDate) = vbSunday) OR (checkHoliday(vOrderDate) = True) then
		vOrderDate = cf_FormatDate(DateAdd("d", 1, vOrderDate), "YYYY/MM/DD")
	else 
		if  vRS("支払方法") = "銀行振込" AND (DatePart("w", vOrderDate) = vbSaturday) then
			vOrderDate = cf_FormatDate(DateAdd("d", 1, vOrderDate), "YYYY/MM/DD")
		else 
			exit do
		end if
	end if
Loop

'---- 指定可能日
'if w94HokkaidouRitouFl = true then                '2011/06/29 an del
'---- 離島：受付日+5日                             '2011/06/29 an mod s
if wRitouFl = "Y" then
	setAvailableDate = cf_FormatDate(DateAdd("d", 5, vOrderDate), "YYYY/MM/DD")
else
	'---- 九州･四国･中国・北海道：受付日+2日
	if w94ChugokuHokkaidouFL = true then
		setAvailableDate = cf_FormatDate(DateAdd("d", 2, vOrderDate), "YYYY/MM/DD")
	'---- その他：受付日+1日                       '2011/06/29 an mod e
	else
		setAvailableDate = cf_FormatDate(DateAdd("d", 1, vOrderDate), "YYYY/MM/DD")
	end if
end if

End function

'========================================================================
'
'	Function	離島フラグの設定
'
'		parm:		配送先郵便番号
'		return:	沖縄・離島なら　wRitouFl = Y
'				離島以外　　　　wRitouFl = N
'
'========================================================================
Function setRitouFlag(p_zip)

Dim vZip
Dim RSv
Dim vSQL

vZip = Replace(p_zip, "-", "")

If vZip = "" Then
	wRitouFl  = "N"
	Exit Function
End If

'---- 離島チェック
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    郵便番号"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    離島 WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    郵便番号 = '" & vZip & "'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

If RSv.EOF = True Then
	wRitouFl = "N"
	'---- 沖縄は離島扱い               '2011/06/29 an add s
	if wShipPrefecture = "沖縄県" then
		wRitouFl = "Y"                 '2011/06/29 an add e
	end if
Else
	wRitouFl = "Y"
End If

RSv.Close

End Function

'========================================================================
'
'	Function	Connect database
'
'========================================================================
Function connect_db()

'---- Connect database
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
Function close_db()

Connection.Close
Set Connection= Nothing    '2011/08/01 an add

End Function

'========================================================================
%>
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="refresh" content="0;URL=<%=nextURL%>">
<meta charset="Shift_JIS">
<title>ご注文受付中｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/StyleOrder.css?20120629a" type="text/css">
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
      <li>ご注文内容の確認</li>
      <li class="now">ご注文受付中</li>
    </ul>
  </div></div></div>

  <h1 class="title">ご注文受付中</h1>
  <ol id="step">
    <li><img src="images/step01.gif" alt="1.ショッピングカート" width="170" height="50"></li>
    <li><img src="images/step02.gif" alt="2.お届け先、お支払方法の選択" width="170" height="50"></li>
    <li><img src="images/step03_now.gif" alt="3.ご注文内容の確認" width="170" height="50"></li>
    <li><img src="images/step04.gif" alt="4.ご注文完了" width="170" height="50"></li>
  </ol>

  <p>ご注文の受付をしています。<br>しばらくお待ちください。</p>

<!--/#contents --></div>
	<div id="globalSide">
	<!--#include file="../Navi/NaviSide.inc"-->
	<!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>
