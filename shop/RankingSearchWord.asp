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
'	アクセスランキング　(検索語)
'
'	更新履歴
'2007/10/16 前月ランキング表示に変更
'2010/05/27 an リニューアル対応（デザイン変更、矢印表示）
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/08/07 if-web リニューアルレイアウト調整
'
'========================================================================

On Error Resume Next

Dim wYYYYMM

Dim Connection
Dim RS

Dim SearchWords(50)

Dim wSQL
Dim wHTML
Dim wSearchWordRankHTML
Dim wRankCount    '前々月のランキング数
Dim wNoData
Dim wMSG
Dim wErrDesc   '2011/08/01 an add

'========================================================================

'---- 送信データーの取り出し

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "RankingSearchWord.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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
Dim vRank
Dim vBGColor
Dim vPrevMonth
Dim vRankingIcon
Dim vSearchWord

'---- 前月
vPrevMonth = DateAdd("m", -1, Date())
wYYYYMM = Year(vPrevMonth) &  Right("0" & Month(vPrevMonth), 2)

'---- ランキング取り出し
wSQL = ""
wSQL = wSQL & "SELECT TOP 50"
wSQL = wSQL & "       a.検索語"
wSQL = wSQL & "  FROM 検索語アクセス件数 a WITH (NOLOCK)"
wSQL = wSQL & " WHERE a.年月 = '" & wYYYYMM & "'"
wSQL = wSQL & " ORDER BY a.検索件数 DESC"

'@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

wHTML = ""
	
if RS.EOF = true then
	wHTML = wHTML & Left(wYYYYMM,4) & "年" & right(wYYYYMM,2) & "月のランキングはありません。"
else
	'---- 前々月の検索キーワードランキングを配列に格納
	call GetRanking2MonthAgo()

	vRank = 0

	Do Until RS.EOF = true

		vRank = vRank + 1
		
		'---- 検索語が長い場合は表示上は省略
		vSearchWord = "" 
		
		if Len(RS("検索語")) > 45 then
			vSearchWord = Left(RS("検索語"),42) & "..."
		else
			vSearchWord = RS("検索語")
		end if
		
		'---- 左ボックス用div(1位～25位)
		if vRank = 1 then
			wHTML = wHTML & "    <div id='ranking_key_list_big_box_left'>" & vbNewLine

		'---- 右ボックス用div(26位～50位)
		elseif vRank = 26 then
			wHTML = wHTML & "    <div id='ranking_key_list_big_box_right'>" & vbNewLine
		end if
		
		wHTML = wHTML & "      <div class='ranking_key_list_box'>" & vbNewLine
		
		if vRank =< 3 then
			wHTML = wHTML & "        <div class='list_box_no" & vRank & "crown_box'><img height='30' src='images/ranking/ico_no" & vRank & "crown.gif' width='41' /></div>" & vbNewLine
		else
			'---- 4位以下の奇数行
			if  vRank Mod 2 <> 0 then
				wHTML = wHTML & "        <div class='list_box_number_box_odd'>" & vRank & "</div>" & vbNewLine
			else
				wHTML = wHTML & "        <div class='list_box_number_box_even'>" & vRank & "</div>" & vbNewLine
			end if
		end if
		
		'---- 前々月とのランキング比較アイコン作成（初期値はNew Entry）
		vRankingIcon = "arrow_new.gif"
		
		For i = 1 To wRankCount
			
			'--- 前々月ランキング内に一致する単語があったら、比較する
			if RS("検索語") = SearchWords(i) then

				if i < vRank  then      'ランクダウン
					vRankingIcon = "arrow_down.gif"
				else
					if i = vRank then   '同ランク
						vRankingIcon = "arrow_right.gif"
					else
						if i > vRank then   'ランクアップ
							vRankingIcon = "arrow_up.gif"
						end if
					end if
				end if
				
				Exit For  '一致したらForを抜ける
				
			end if
		Next
		
		wHTML = wHTML & "        <div class='list_box_arrow_box'><img src='images/ranking/" & vRankingIcon & "' alt='' width='30' height='30' /></div>" & vbNewLine
		
		if  vRank Mod 2 <> 0 then
			wHTML = wHTML & "        <div class='list_box_word_box_odd'>" & vbNewLine
		else
			wHTML = wHTML & "        <div class='list_box_word_box_even'>" & vbNewLine
		end if

		wHTML = wHTML & "          <div class='list_box_word_box_text'><a href='SearchList.asp?search_all=" & Server.URLEncode(RS("検索語")) & "'>" & vSearchWord & "</a></div>" & vbNewLine
		wHTML = wHTML & "        </div>" & vbNewLine
		wHTML = wHTML & "      </div>" & vbNewLine
		
		if vRank = 25 OR vRank = 50 then
			wHTML = wHTML & "    </div>" & vbNewLine
		end if

		RS.MoveNext
	Loop
	
	'---- 1列目の途中で終わっている場合
	if vRank >= 1 AND vRank < 25 then
		wHTML = wHTML & "    </div>" & vbNewLine
	end if
	
	'---- 2列目の途中で終わっている場合
	if vRank >= 26 AND vRank < 50 then
		wHTML = wHTML & "    </div>" & vbNewLine
	end if
	
end if

RS.Close

wSearchWordRankHTML = wHTML

End function

'========================================================================
'
'	Function	前々月の検索キーワードランキング取得
'
'========================================================================
'
Function GetRanking2MonthAgo()

Dim RSv
Dim vPrevMonth
Dim vYYYYMM
Dim i


'---- 前々月
vPrevMonth = DateAdd("m", -2, Date())
vYYYYMM = Year(vPrevMonth) &  Right("0" & Month(vPrevMonth), 2)

'---- ランキング取り出し
wSQL = ""
wSQL = wSQL & "SELECT TOP 50"
wSQL = wSQL & "       a.検索語"
wSQL = wSQL & "  FROM 検索語アクセス件数 a WITH (NOLOCK)"
wSQL = wSQL & " WHERE a.年月 = '" & vYYYYMM & "'"
wSQL = wSQL & " ORDER BY a.検索件数 DESC"

'@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

if RSv.EOF <> true then

	wRankCount = RSv.RecordCount

	i = 1 
	
	Do Until RSv.EOF = true
		SearchWords(i) = RSv("検索語")
		i = i + 1
		RSv.MoveNext
	Loop
	
end if

RSv.Close

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
<title>検索キーワード｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/Ranking.css" type="text/css">
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

<div id="ranking_key_main_flame">
  <div id="shukei">（集計：<%=Left(wYYYYMM,4)%>年<%=right(wYYYYMM,2)%>月）</div>
<!-- Menu START -->
  <div id="ranking_key_top_menu">
    <div class="top_menu_parts">
      <a href="BestSellerList.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image15','','images/ranking/ts_btn_on.jpg',1)"><img src="images/ranking/ts_btn_off.jpg" alt="" name="Image15" width="114" height="80" />
      </a>
    </div>
    <div class="top_menu_parts">
      <a href="RankingSearchWord.asp">
        <img src="images/ranking/sk_btn_on.jpg" alt="" name="Image163" width="114" height="80" />
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
      <a href="RankingAccess.asp?RankType=<%=Server.URLEncode("欲しいものリスト")%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image14','','images/ranking/wl_btn_on.jpg',1)"><img src="images/ranking/wl_btn_off.jpg" alt="" name="Image14" width="114" height="80" />
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
<!-- List START -->
  <div id="ranking_key_list_flame"> 
<%=wSearchWordRankHTML%>
  
  </div>
<!-- List END -->
  </div>
<!-- container END -->

</div>
<!-- Mainpage END -->
  
  </div>
  <div id="globalSide">
    <!--#include file="../Navi/NaviSide.inc"-->
  </div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>