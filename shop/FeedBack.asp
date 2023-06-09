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
'	出荷後アンケート
'
'更新履歴
'2007/04/23 入力欄桁数チェック追加
'2010/03/17 hn ログインチェックを追加
'2010/07/16 st アンケート文字数制限500文字以内に対応
'2010/12/07 an 質問3,4の順番入れ替え→フォーム名は変更しない
'2011/08/01 an #1087 Error.aspログ出力対応
'2011/08/31 an 携帯からアクセス時はmobiへRedirectする
'2012/06/08 GV #1367 アンケートでの購入者チェック
'2012/07/25 ok 購入者チェックをEmaxDBの受注テーブルから確認するよう変更
'========================================================================

On Error Resume Next

Dim userID		'2010/03/17 hn add
Dim OrderNo

Dim q1
Dim q1Name
Dim q1Department
Dim q1Comment
Dim q2
Dim q2Comment
Dim q3         '質問4に該当
Dim q3Comment  '質問4に該当
Dim q4         '質問3に該当
Dim q4Other    '質問3に該当
Dim q5
Dim q5Comment
Dim q6
Dim q6Other
Dim q7
Dim q7Comment

Dim wSQL
Dim wMsg
Dim wErrDesc   '2011/08/01 an add

'2010/07/16 st add s
Dim Connection
Dim RS
Dim wCanWriteFl
'2010/07/16 st add e
Dim ConnectionEmax		'2012/07/25 ok Add
'========================================================================

Response.buffer = true

wMsg = ""

'---- Sessionにアンケート送信フラグがあればそれを使用		2010/07/16 st add
if Session("CanWriteFl") <> "" then
	wCanWriteFl = Session("CanWriteFl")
else
	wCanWriteFl = "N"
end if

'---- UserID 取り出し	'2010/03/17 hn add
userID = Session("userID")

'---- パラメータ取り込み
OrderNo = ReplaceInput(Request("OrderNo"))

q1 = ReplaceInput(Left(Request("q1"), 10))
q1Name = ReplaceInput(Left(Request("q1Name"), 10))
q1Department = ReplaceInput(Left(Request("q1Department"), 10))
q1Comment = ReplaceInput(Left(Request("q1Comment"), 500))
q2 = ReplaceInput(Left(Request("q2"), 10))
q2Comment = ReplaceInput(Left(Request("q2Comment"), 500))
q3 = ReplaceInput(Left(Request("q3"), 10))                  '質問4に該当
q3Comment = ReplaceInput(Left(Request("q3Comment"), 500))   '質問4に該当
q4 = ReplaceInput(Left(Request("q4"), 150))                 '質問3に該当
q4Other = ReplaceInput(Left(Request("q4Other"), 50))        '質問3に該当
q5 = ReplaceInput(Left(Request("q5"), 10))
q5Comment = ReplaceInput(Left(Request("q5Comment"), 500))
q6 = ReplaceInput(Left(Request("q6"), 50))
q6Other = ReplaceInput(Left(Request("q6Other"), 500))
q7 = ReplaceInput(Left(Request("q7"), 10))
q7Comment = ReplaceInput(Left(Request("q7Comment"), 500))

'Response.Write("UserId:" & userID & "<br>")
'Response.Write("OrderNo:" & Request("OrderNo") & "<br>")

'---- 携帯からアクセスされた場合はmobiへRedirect    2011/08/31 an add s
if gPhoneType = "NMB" then
	Response.Redirect g_HTTPmobi & "shop/FeedBack.asp?OrderNo=" & OrderNo
elseif gPhoneType = "SP" then
	Response.Redirect g_HTTPsp & "shop/FeedBack.asp?OrderNo=" & OrderNo
end if                                             '2011/08/31 an add e

'---- Sessionにオーダー番号があればそれを使用		2010/03/17 hn add
if OrderNo = "" then
	 OrderNo = Session("OrderNo")
else
	Session("OrderNo") = OrderNo
end if

'---- ログインをしているかどうかのチェック	2010/03/17 hn add
if userID = "" then
	wMsg = "ログインをしてください。" 
	Session("msg") = wMSG
	Response.Redirect g_HTTPS & "shop/Login.asp?called_from=feedback"
end if

'---- 項目未入力か文字数オーバー時の書き直し			2010/07/16 st add s
if Session("msg") <> "" then
	wMsg = "<p class='error'>" & Session("msg") & "</p>"
	Session("msg") = ""
else
	call connect_db()
	call main()
	
		'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
	if Err.Description <> "" then
		wErrDesc = "FeedBack.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
		call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
	end if                                               '2011/08/01 an add e

	call close_db()
	
	if Err.Description <> "" then                     '2011/08/01 an add s
		Response.Redirect g_HTTP & "shop/Error.asp"
	end if                                            '2011/08/01 an add e

end if

'2010/07/16 st add e

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

'2012/07/25 ok Add
Set ConnectionEmax = Server.CreateObject("ADODB.Connection")
ConnectionEmax.Open g_connectionEmax

End function
'========================================================================
'
'	Function	main proc
'
'========================================================================
'
Function main()

if OrderNo = "" then
	wMsg = "<p class='error'>受注番号がありませんのでこのアンケートは送信できません。</p>"
else
    '//Add GV #1367 Start
    If BuyerCheck = False Then
        wMsg = "<p class='error'>ご購入者様のユーザーIDと異なるためこのアンケートは送信できません。</p>"
        Exit Function
    End If
    '//Add GV #1367 End

	'---- アンケート登録有無チェック
	wSQL = ""
	wSQL = wSQL & "SELECT *"
	wSQL = wSQL & "  FROM 出荷後アンケート WITH (NOLOCK)"
	wSQL = wSQL & " WHERE 受注番号 = " & OrderNo 
		  
	Set RS = Server.CreateObject("ADODB.Recordset")
	RS.Open wSQL, Connection, adOpenStatic

	if RS.EOF = false then
		wMsg = "<p class='error'>" & "受注番号[" &  OrderNo & "] に関するアンケートは、既に登録されています。" & "</p>"
	else
		wCanWriteFL = "Y"
	end if
	RS.close
end if

End function

'Add GV #1367 START
'========================================================================
'
'	Function	BuyerCheck
'
'========================================================================
Function BuyerCheck()

    wSQL = ""
	wSQL = wSQL & "SELECT *"
'2012/07/25 ok Add Start
	wSQL = wSQL & "  FROM 受注 a WITH (NOLOCK) "
	wSQL = wSQL & "WHERE "
    wSQL = wSQL & "     a.顧客番号 = " & userID & " AND " 
    wSQL = wSQL & "     a.受注番号 = " & OrderNo

	Set RS = Server.CreateObject("ADODB.Recordset")
	RS.Open wSQL, ConnectionEmax, adOpenStatic, adLockOptimistic
'2012/07/25 ok Add End

'2012/07/25 ok Del Start
'	wSQL = wSQL & "  FROM Web顧客 a WITH (NOLOCK)"
'    wSQL = wSQL & " INNER JOIN Web受注 b ON a.顧客番号 = b.顧客番号"
'	wSQL = wSQL & " WHERE " 
'    wSQL = wSQL & "     a.顧客番号 = " & userID & " AND " 
'    wSQL = wSQL & "     b.受注番号 = " & OrderNo

'    Set RS = Server.CreateObject("ADODB.Recordset")
'	RS.Open wSQL, Connection, adOpenStatic
'2012/07/25 ok Del End

'    Response.Write RS.RecordCount
    If RS.RecordCount > 0 Then
        BuyerCheck = True
    Else
        BuyerCheck = False
    End If
    
    RS.Close

End Function
'Add GV #1367 END

'========================================================================
'
'	Function	Close database
'
'========================================================================
'
Function close_db()

Connection.close
Set Connection= Nothing    '2011/08/01 an add

'2012/07/25 ok Add
ConnectionEmax.Close
Set ConnectionEmax = Nothing

End function

'========================================================================
%>

<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="Shift_JIS">
<title>ご購入者様向けアンケート｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/feedback.css" type="text/css">

<script type="text/javascript">
//
// ====== 	Function:	FeedBack_onSubmit
//
function FeedBack_onSubmit(pForm){
	if (pForm.q1Comment.value.length > 500){
		alert("質問1に入力された文字数が500文字を超えています｡　500文字以内でお願いします。");
		return false;
	}	
	if (pForm.q2Comment.value.length > 500){
		alert("質問2に入力された文字数が500文字を超えています｡　500文字以内でお願いします。");
		return false;
	}
	if (pForm.q3Comment.value.length > 500){
		alert("質問4に入力された文字数が500文字を超えています｡　500文字以内でお願いします。");
		return false;
	}
	if (pForm.q5Comment.value.length > 500){
		alert("質問5に入力された文字数が500文字を超えています｡　500文字以内でお願いします。");
		return false;
	}
	if (pForm.q6Other.value.length > 500){
		alert("質問6に入力された文字数が500文字を超えています｡　500文字以内でお願いします。");
		return false;
	}
	if (pForm.q7Comment.value.length > 500){
		alert("質問7に入力された文字数が500文字を超えています｡　500文字以内でお願いします。");
		return false;
	}
	return true;
}

</script>

</head>
<body>

<!--#include file="../Navi/NaviTop.inc"-->

<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>
  
  <!-- コンテンツstart -->
  <div id="globalContents" class="feedback">
    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="<%=g_HTTP%>"><img src="<%=g_RelLink%>images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li class="now">ご購入者様向けアンケート</li>
      </ul>
    </div></div></div>

    <h1 class="title">ご購入者様向けアンケート</h1>
    <p>この度はサウンドハウスをご利用いただき、誠にありがとうございました。<br>商品のご注文、配達、お届けした商品の状態など十分ご満足いただけましたでしょうか。</p>
    <p>私どもサウンドハウスが一番大切にいたしております「心のこもったサービス」をモットーに、お客様に十分ご満足していただけるサービスを提供するよう努力いたしております。お気付きの点がございましたら、どんな小さな事でもご遠慮なくお知らせください。業務用オーディオ、楽器、照明の総合デパートとして今後ともお客様のご要望に可能な限りお応えしていきたいと考えていますので、ご協力をお願いいたします。</p>
    
    <%=wMsg%>
    
    <form name="fFeedBack" action="FeedBackStore.asp" method="post" onSubmit="return FeedBack_onSubmit(this)">
    
    <p>受注番号：<%=OrderNo%><input type="hidden" name="OrderNo" value="<%=OrderNo%>"></p>
    
    <ol class="form">
    	<li>
        	<h2>1. スタッフの応対</h2>
            <p>例えば、ご注文やお問い合わせの際の対応はいかがでしたでしょうか？</p>
            <ul>
            	<li><input type="radio" id="q1_5" name="q1" value="大変満足" <% if q1 = "大変満足" then %> checked <% end if %>><label for="q1_5">大変満足</label></li>
                <li><input type="radio" id="q1_4" name="q1" value="満足" <% if q1 = "満足" then %> checked <% end if %>><label for="q1_4">満足</label></li>
                <li><input type="radio" id="q1_3" name="q1" value="普通" <% if q1 = "普通" then %> checked <% end if %>><label for="q1_3">普通</label></li>
                <li><input type="radio" id="q1_2" name="q1" value="不満" <% if q1 = "不満" then %> checked <% end if %>><label for="q1_2">不満</label></li>
                <li><input type="radio" id="q1_1" name="q1" value="大変不満" <% if q1 = "大変不満" then %> checked <% end if %>><label for="q1_1">大変不満</label></li>
            </ul>
            <p>お客様への応対中、特に優秀と思われました従業員がおりましたらご記入ください。</p>
            <ul>
            	<li>名前：<input name="q1Name" type="text" size="20" value="<%=q1Name%>"></li>
                <li>部署：<input name="q1Department" type="text" size="20" value="<%=q1Department%>"></li>
            </ul>
            <p>ご意見をどうぞ(500文字まで)</p>
            <textarea name="q1Comment" rows="3"><%=q1Comment%></textarea>
        </li>
        <li>
        	<h2>2. ホームページの内容</h2>
            <p>例えば、商品の検索のしやすさや、情報の充実度、疑問点がすぐに解決できたでしょうか？</p>
            <ul>
            	<li><input type="radio" id="q2_5" name="q2" value="大変使いやすい" <% if q2 = "大変使いやすい" then %> checked <% end if %>><label for="q2_5">大変使いやすい</label></li>
                <li><input type="radio" id="q2_4" name="q2" value="使いやすい" <% if q2 = "使いやすい" then %> checked <% end if %>><label for="q2_4">使いやすい</label></li>
                <li><input type="radio" id="q2_3" name="q2" value="普通" <% if q2 = "普通" then %> checked <% end if %>><label for="q2_3">普通</label></li>
                <li><input type="radio" id="q2_2" name="q2" value="使いにくい" <% if q2 = "使いにくい" then %> checked <% end if %>><label for="q2_2">使いにくい</label></li>
                <li><input type="radio" id="q2_1" name="q2" value="大変使いにくい" <% if q2 = "大変使いにくい" then %> checked <% end if %>><label for="q2_1">大変使いにくい</label></li>
            </ul>
            <p>ご意見をどうぞ(500文字まで)</p>
            <textarea name="q2Comment" rows="3"><%=q2Comment%></textarea>
        </li>
        <li>
        	<h2>3. 購読している雑誌(複数回答可)</h2>
            <p>下記のうち、よく読むものをいくつでも結構ですのでチェックしてください。<br>該当雑誌がない場合、その他の欄にご記入ください。</p>
            <ul>
            	<li><input type="checkbox" id="q4_1" name="q4" value="サウンド＆レコーディング・マガジン" <% if InStr(q4,"サウンド&amp;レコーディング・マガジン") <> "0" then %> checked <% end if %>><label for="q4_1">サウンド&amp;レコーディング・マガジン</label></li>
                <li><input type="checkbox" id="q4_2" name="q4" value="ギター・マガジン" <% if InStr(q4,"ギター・マガジン") <> "0" then %> checked <% end if %>><label for="q4_2">ギター・マガジン</label></li>
                <li><input type="checkbox" id="q4_3" name="q4" value="ベース・マガジン" <% if InStr(q4,"ベース・マガジン") <> "0" then %> checked <% end if %>><label for="q4_3">ベース・マガジン</label></li>
                <li><input type="checkbox" id="q4_4" name="q4" value="リズム＆ドラム・マガジン" <% if InStr(q4,"リズム＆ドラム・マガジン") <> "0" then %> checked <% end if %>><label for="q4_4">リズム＆ドラム・マガジン</label></li>
                <li><input type="checkbox" id="q4_5" name="q4" value="ゲッカヨ" <% if InStr(q4,"ゲッカヨ") <> "0" then %> checked <% end if %>><label for="q4_5">ゲッカヨ</label></li>
                <li><input type="checkbox" id="q4_6" name="q4" value="アコースティック・ギター・マガジン" <% if InStr(q4,"アコースティック・ギター・マガジン") <> "0" then %> checked <% end if %>><label for="q4_6">アコースティック・ギター・マガジン</label></li>
                <li><input type="checkbox" id="q4_7" name="q4" value="GROOVE" <% if InStr(q4,"GROOVE") <> "0" then %> checked <% end if %>><label for="q4_7">GROOVE</label></li>
                <li><input type="checkbox" id="q4_8" name="q4" value="ヤングギター" <% if InStr(q4,"ヤングギター") <> "0" then %> checked <% end if %>><label for="q4_8">ヤングギター</label></li>
                <li><input type="checkbox" id="q4_9" name="q4" value="DTMマガジン" <% if InStr(q4,"DTMマガジン") <> "0" then %> checked <% end if %>><label for="q4_9">DTMマガジン</label></li>
                <li><input type="checkbox" id="q4_10" name="q4" value="ビデオサロン" <% if InStr(q4,"ビデオサロン") <> "0" then %> checked <% end if %>><label for="q4_10">ビデオサロン</label></li>
                <li><input type="checkbox" id="q4_11" name="q4" value="ビデオα" <% if InStr(q4,"ビデオα") <> "0" then %> checked <% end if %>><label for="q4_11">ビデオα</label></li>
                <li><input type="checkbox" id="q4_12" name="q4" value="カラオケファン" <% if InStr(q4,"カラオケファン") <> "0" then %> checked <% end if %>><label for="q4_12">カラオケファン</label></li>
                <li><input type="checkbox" id="q4_13" name="q4" value="歌の手帖" <% if InStr(q4,"歌の手帖") <> "0" then %> checked <% end if %>><label for="q4_13">歌の手帖</label></li>
             </ul>
             <p>その他<input name="q4Other" type="text" size="50" value="<%=q4Other%>"></p>
        </li>
        <li>
        	<h2>4. 雑誌広告の内容</h2>
            <p>サウンドハウスの雑誌広告をご覧になったお客さまにご質問いたします。<br>例えば、商品について興味を持ったり、購入の際に参考になる内容でしたでしょうか？</p>
            <ul>
            	<li><input type="radio" id="q3_5" name="q3" value="大変満足"  <% if q3 = "大変満足" then %> checked <% end if %>><label for="q3_5">大変満足</label></li>
                <li><input type="radio" id="q3_4" name="q3" value="満足" <% if q3 = "満足" then %> checked <% end if %>><label for="q3_4">満足</label></li>
                <li><input type="radio" id="q3_3" name="q3" value="普通" <% if q3 = "普通" then %> checked <% end if %>><label for="q3_3">普通</label></li>
                <li><input type="radio" id="q3_2" name="q3" value="不満" <% if q3 = "不満" then %> checked <% end if %>><label for="q3_2">不満</label></li>
                <li><input type="radio" id="q3_1" name="q3" value="大変不満" <% if q3 = "大変不満" then %> checked <% end if %>><label for="q3_1">大変不満</label></li>
            </ul>
            <p>ご意見をどうぞ(500文字まで)</p>
            <textarea name="q3Comment" rows="3"><%=q3Comment%></textarea>
        </li>
        <li>
        	<h2>5. カタログ(ホットメニュー・ホットスタッフ)の内容</h2>
            <p>例えば、商品スペックの写真、スペックの見やすさや、内容の充実度はいかがでしょうか？</p>
            <ul>
            	<li><input type="radio" id="q5_5" name="q5" value="大変満足" <% if q5 = "大変満足" then %> checked <% end if %>><label for="q5_5">大変満足</label></li>
                <li><input type="radio" id="q5_4" name="q5" value="満足" <% if q5 = "満足" then %> checked <% end if %>><label for="q5_4">満足</label></li>
                <li><input type="radio" id="q5_3" name="q5" value="普通" <% if q5 = "普通" then %> checked <% end if %>><label for="q5_3">普通</label></li>
                <li><input type="radio" id="q5_2" name="q5" value="不満" <% if q5 = "不満" then %> checked <% end if %>><label for="q5_2">不満</label></li>
                <li><input type="radio" id="q5_1" name="q5" value="大変不満" <% if q5 = "大変不満" then %> checked <% end if %>><label for="q5_1">大変不満</label></li>
            </ul>
            <p>ご意見をどうぞ(500文字まで)</p>
            <textarea name="q5Comment" rows="3"><%=q5Comment%></textarea>
        </li>
        <li>
        	<h2>6. サウンドハウスをお選びいただいた理由</h2>
            <ul>
            	<li><input type="radio" id="q6_5" name="q6" value="前に利用したときの印象が良かった" <% if q6 = "前に利用したときの印象が良かった" then %> checked <% end if %>><label for="q6_5">前に利用したときの印象が良かった</label></li>
                <li><input type="radio" id="q6_4" name="q6" value="人にすすめられて" <% if q6 = "人にすすめられて" then %> checked <% end if %>><label for="q6_4">人にすすめられて</label></li>
                <li><input type="radio" id="q6_3" name="q6" value="雑誌の広告を見て" <% if q6 = "雑誌の広告を見て" then %> checked <% end if %>><label for="q6_3">雑誌の広告を見て</label></li>
                <li><input type="radio" id="q6_2" name="q6" value="インターネット検索から" <% if q6 = "インターネット検索から" then %> checked <% end if %>><label for="q6_2">インターネット検索から</label></li>
            </ul>
            <p>ご意見をどうぞ(500文字まで)</p>
            <textarea name="q6Other" rows="3"><%=q6Other%></textarea>
        </li>
        <li>
        	<h2>7. お知り合いの方にご利用をおすすめいただけますか？</h2>
            <ul>
            	<li><input type="radio" id="q7_5" name="q7" value="はい" <% if q7 = "はい" then %> checked <% end if %>><label for="q7_5">はい</label></li>
                <li><input type="radio" id="q7_4" name="q7" value="いいえ" <% if q7 = "いいえ" then %> checked <% end if %>><label for="q7_4">いいえ</label></li>
            </ul>
            <p>ご意見をどうぞ(500文字まで)</p>
            <textarea name="q7Comment" rows="3"><%=q7Comment%></textarea>
        </li>
    </ol>
    
    <input type="hidden" name="q8" value="">
	<input type="hidden" name="q9" value="">
    
    <% if wCanWriteFl = "Y" then %>
    <p>よろしければ送信ボタンを押してください。</p>
    <p class="btnBox"><input type="submit" value="送信" class="opover"></p>
	<% end if %>

	</form>
    
</div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>
