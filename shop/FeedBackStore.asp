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
'	お客様アンケート 登録
'
'更新履歴
'2008/05/23 入力データチェック強化（LEFT他)
'2009/04/30 エラー時にerror.aspへ移動
'2010/07/16 st アンケート文字数制限500文字以内に対応
'2010/12/07 an 質問3,4の順番入れ替え。質問4は任意項目に変更→フォーム名は変更しない
'2011/02/23 hn パラメータ受取時質問の500文字で長さカットを削除
'2011/08/01 an #1087 Error.aspログ出力対応
'
'========================================================================

On Error Resume Next

Dim userID
Dim msg

Dim OrderNo
Dim q1
Dim q1Name
Dim q1Department
Dim q1Comment
Dim q2
Dim q2Comment
Dim q3           '質問4に該当
Dim q3Comment    '質問4に該当
Dim q4           '質問3に該当
Dim q4Other      '質問3に該当
Dim q5
Dim q5Comment
Dim q6
Dim q6Other
Dim q7
Dim q7Comment

'2010 07/29 st del s
'Dim q8
'Dim q9
'2010 07/29 e

Dim Connection
Dim RS

Dim w_sql
Dim w_html
Dim w_msg
Dim wErrDesc   '2011/08/01 an add

Dim wCanWriteFl

'========================================================================

Response.buffer = true

'---- UserID 取り出し
userID = Session("userID")

'---- 呼び出し元からのデータ取り出し
OrderNo = ReplaceInput(Request("OrderNo"))
q1 = ReplaceInput(Left(Request("q1"), 10))
q1Name = ReplaceInput(Left(Request("q1Name"), 10))
q1Department = ReplaceInput(Left(Request("q1Department"), 10))
q1Comment = ReplaceInput(Request("q1Comment"))	'2011/02/23 hn mod
q2 = ReplaceInput(Left(Request("q2"), 10))
q2Comment = ReplaceInput(Request("q2Comment"))		'2011/02/23 hn mod
q3 = ReplaceInput(Left(Request("q3"), 10))                   '質問4に該当
q3Comment = ReplaceInput(Request("q3Comment"))						   '質問4に該当			'2011/02/23 hn mod
q4 = ReplaceInput(Left(Request("q4"), 150))                  '質問3に該当
q4Other = ReplaceInput(Left(Request("q4Other"), 50))         '質問3に該当
q5 = ReplaceInput(Left(Request("q5"), 10))
q5Comment = ReplaceInput(Request("q5Comment"))		'2011/02/23 hn mod
q6 = ReplaceInput(Left(Request("q6"), 50))
q6Other = ReplaceInput(Request("q6Other"))			'2011/02/23 hn mod
q7 = ReplaceInput(Left(Request("q7"), 10))
q7Comment = ReplaceInput(Request("q7Comment"))		'2011/02/23 hn mod

'2010 07/16 st del s
'q8 = ReplaceInput(Left(Request("q8"), 10))
'q9 = ReplaceInput(Request("q9"))
'2010 07/16 e

if isNumeric(OrderNo) = true then
	'---- Execute main
	call connect_db()
	call main()
	
	'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
	if Err.Description <> "" then
		wErrDesc = "FeedBackStore.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
		call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
	end if                                           '2011/08/01 an add e

	call close_db()
	
	if Err.Description <> "" then	
		Response.Redirect g_HTTP & "shop/Error.asp"
	end if
	
	'2010 07/16 st ad s
	if w_msg <> "" then
		Session("msg") = w_msg
		Session("CanWriteFl") = wCanWriteFl
		Server.Transfer "FeedBack.asp"
	end if
	'2010 07/16 e
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
'	Function	main proc
'
'========================================================================
'
Function main()

call ValidateData()

if w_msg <> "" then
	wCanWriteFl = "Y" '------書き直しが必要な場合 2010/07/17
else
	'---- アンケート結果登録
	w_sql = ""
	w_sql = w_sql & "SELECT *"
	w_sql = w_sql & "  FROM 出荷後アンケート"
	w_sql = w_sql & " WHERE 受注番号 = " & OrderNo 
		  
	Set RS = Server.CreateObject("ADODB.Recordset")
	RS.Open w_sql, Connection, adOpenStatic, adLockOptimistic

	if RS.EOF = false then
		w_msg = "受注番号[" &  OrderNo & "] に関するアンケートは、既に登録されています。"
		wCanWriteFl = "N"
	else
		'---- insert アンケート
		RS.AddNew

		RS("受注番号") = OrderNo
		RS("質問1") = q1
		RS("質問1名前") = q1Name
		RS("質問1部署") = q1Department
		RS("質問1意見") = q1Comment
		RS("質問2") = q2
		RS("質問2意見") = q2Comment
		RS("質問3") =q3               '質問4に該当
		RS("質問3意見") = q3Comment   '質問4に該当
		RS("質問4") = q4              '質問3に該当
		RS("質問4その他") = q4Other   '質問3に該当
		RS("質問5") = q5
		RS("質問5意見") = q5Comment
		RS("質問6") = q6
		RS("質問6その他") = q6Other
		RS("質問7") = q7
		RS("質問7意見") = q7Comment
		
		'2010 07/16 st del s
'		RS("質問8") = q8
'		RS("質問9") = q9
		'2010 07/16 e

		RS("登録日") = now()

		RS.Update
	end if
	RS.close
end if

End function

'========================================================================
'
'	Function	入力データチェック '2010/07/16 st add
'
'========================================================================
'
Function ValidateData()

if q1 = "" then
	w_msg = w_msg & "質問1が入力されていません。<br>"
end if

if (Len(q1Comment)) > 500 then
	w_msg = w_msg & "質問1に入力された文字数が500文字を超えています｡　500文字以内でお願いします。<br>"
end if

if q2 = "" then
	w_msg = w_msg & "質問2が入力されていません。<br>"
end if

if (Len(q2Comment)) > 500 then
	w_msg = w_msg & "質問2に入力された文字数が500文字を超えています｡　500文字以内でお願いします。<br>"
end if

'if q3 = "" then          '2010/12/07 an del q3は質問4に該当。質問4は必須でなくす
'	w_msg = w_msg & "質問4が入力されていません。<br>"
'end if

if (Len(q3Comment)) > 500 then
	w_msg = w_msg & "質問4に入力された文字数が500文字を超えています｡　500文字以内でお願いします。<br>"  '2010/12/07 an mod q3は質問4に該当
end if

if q5 = "" then
	w_msg = w_msg & "質問5が入力されていません。<br>"
end if

if (Len(q5Comment)) > 500 then
	w_msg = w_msg & "質問5に入力された文字数が500文字を超えています｡　500文字以内でお願いします。<br>"
end if

if q6 = "" then
	w_msg = w_msg & "質問6が入力されていません。<br>"
end if

if (Len(q6Other)) > 500 then
	w_msg = w_msg & "質問6に入力された文字数が500文字を超えています｡　500文字以内でお願いします。<br>"
end if

if q7= "" then
	w_msg = w_msg & "質問7が入力されていません。<br>"
end if

if (Len(q7Comment)) > 500 then
	w_msg = w_msg & "質問7に入力された文字数が500文字を超えています｡　500文字以内でお願いします。<br>"
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
<html lang="ja">
<head>
<meta charset="Shift_JIS">
<title>アンケートありがとうございました｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/feedback.css" type="text/css">

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
    
    <p><strong>アンケートにご協力ありがとうございました。</strong></p>
    <p>今後ともよろしくお引き立てくださいますようお願いいたします。</p>

</div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>
