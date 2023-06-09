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
<!--#include file="../common/HttpsSecurity.inc"-->

<%
'========================================================================
'
'	問合せ内容確認ページ
'     「入力内容にエラーがある/変更をクリック」でInquiry.aspに戻る
'     「送信クリック」でInquirySend.aspへ
'
'2011/04/13 an 新規作成 #725
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/06/29 if-web リニューアルレイアウト調整
'
'========================================================================
On Error Resume Next

Dim message
Dim subject
Dim ContactCategory
Dim ContactSubCategory
Dim ContactSubCategoryFl
Dim customer_nm
Dim furigana
Dim zip
Dim prefecture
Dim address
Dim telephone
Dim fax
Dim e_mail

Dim Skey
Dim Connection

Dim wMessage
Dim wMSG
Dim wErrDesc   '2011/08/01 an add

'========================================================================
Response.buffer = true

'---- 呼び出し元からのデータ取り出し
ContactCategory = ReplaceInput_NoCRLF(Left(Request("ContactCategory"),20))
ContactSubCategory = ReplaceInput_NoCRLF(Left(Request("ContactSubCategory"),20))
ContactSubCategoryFl = ReplaceInput_NoCRLF(Left(Request("ContactSubCategoryFl"),1))
subject = ReplaceInput_NoCRLF(Left(Request("subject"),151))
message = ReplaceInput(Left(Request("message"),2001))
customer_nm = ReplaceInput_NoCRLF(Left(Request("customer_nm"),31))
furigana = ReplaceInput_NoCRLF(Left(Request("furigana"),31))
zip = ReplaceInput_NoCRLF(Left(Request("zip"),9))
prefecture = ReplaceInput_NoCRLF(Left(Request("prefecture"),9))
address = ReplaceInput_NoCRLF(Left(Request("address"),41))
telephone = ReplaceInput_NoCRLF(Left(Request("telephone"),21))
fax = ReplaceInput_NoCRLF(Left(Request("fax"),21))
e_mail = ReplaceInput_NoCRLF(Left(Request("e_mail"),61))

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "InquiryConfirm.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if wMsg <> "" then
    Server.Transfer "Inquiry.asp"
end if

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
'	Function	main
'
'========================================================================
Function main()

'---- セキュリティーキーセット 
Skey = SetSecureKey()

'---- 入力チェック
call validation()

'---- 入力チェックOKなら確認画面作成
if wMsg = "" then
	'改行コードを<br>に変換
    wMessage = Replace(message, vbNewLine, "<br>")
else
	Session("Msg") = wMsg
	exit Function
end if

End function

'========================================================================
'
'    Function    問合せ入力内容チェック
'
'========================================================================
'
Function validation()

Dim vAddress

wMsg = ""

'---- 「種別」
if ContactCategory = "" OR (ContactSubCategoryFl <> "N" AND ContactSubCategory = "" ) Then
    wMsg = wMsg & "種別を選択してください。<br>"
end if

'---- 「件名」
if subject = "" Then
    wMsg = wMsg & "件名を入力してください。<br>"
else
	if Len(subject) > 150 then
    	wMsg = wMsg & "件名は150文字以内で入力してください。<br>"
	end if
end if

'---- 「メッセージ」
if message = "" Then
    wMsg = wMsg & "メッセージを入力してください。<br>"
else
	if Len(message) > 2000 then
    	wMsg = wMsg & "メッセージは2000文字以内で入力してください。<br>"
	end if
end if

'---- 「お名前」
if customer_nm = "" Then
    wMsg = wMsg & "お名前を入力してください。<br>"
else
	if Len(customer_nm) > 30 then
    	wMsg = wMsg & "お名前は30文字以内で入力してください。<br>"
	end if
end if

'---- 「フリガナ」
if cf_checkKataKana(furigana) = false Then
    wMsg = wMsg & "フリガナは全角カナで入力してください。<br>"
else
	if Len(furigana) > 30 then
    	wMsg = wMsg & "フリガナは30文字以内で入力してください。<br>"
	end if
end if

'---- 「郵便番号」
if zip <> "" then
	if IsNumeric(Replace(zip, "-", "")) = False Or cf_checkHankaku2(zip) = False Then
		wMsg = wMsg & "郵便番号を半角数字とハイフン(－)で入力してください。<br>"
	else
		if Len(zip) > 10 then
	    	wMsg = wMsg & "郵便番号は10文字以内で入力してください。<br>"
	    else
	    	if check_zip(zip, vAddress) = False Then
				wMsg = wMsg & "郵便番号が郵便番号辞書にありません。<br>"
			else
				'都道府県が選択されている場合は不整合がないかチェック
				if prefecture <> "" then
					if InStr(vAddress, Trim(prefecture)) <= 0  Then
						wMsg = wMsg & "入力された郵便番号と都道府県が一致しません。<br>"
					end if
				end if
			end if
		end if
	end if
end if

'---- 「住所」
if Len(address) > 40 then
    wMsg = wMsg & "住所は40文字以内で入力してください。<br>"
end if

'---- 「電話番号」
if telephone = "" Then
    wMsg = wMsg & "電話番号を入力してください。<br>"
else

	if IsNumeric(Replace(telephone, "-", "")) = False Or cf_checkHankaku2(telephone) = False Then
		wMsg = wMsg & "電話番号を半角数字とハイフン(－)で入力してください。<br>"
	else
        if Len(telephone) > 20 then
            wMsg = wMsg & "電話番号は20文字以内で入力してください。<br>"
        end if
    end if
end if

'---- 「FAX番号」
if fax <> "" then
	if IsNumeric(Replace(fax, "-", "")) = False Or cf_checkHankaku2(fax) = False Then
		wMsg = wMsg & "FAX番号を半角数字とハイフン(－)で入力してください。<br>"
	else
		if Len(fax) > 20 then
	    	wMsg = wMsg & "FAX番号は20文字以内で入力してください。<br>"
	    end if
	end if
end if

 '---- 「E mail」
if e_mail = "" Then
    wMsg = wMsg & "メールアドレスを入力してください。<br>"
else
	if Len(e_mail) > 60 then
    	wMsg = wMsg & "メールアドレスは60文字以内で入力してください。<br>"
    else
		if fCheckEmail(e_mail) = false then
    		wMsg = wMsg & "メールアドレスが適切ではありません。<br>"
    	end if
    end if
end if

End function

'========================================================================
'
'	Function	郵便番号辞書検索
'
'========================================================================
Function check_zip(vZip, vAddress)

Dim RSv
Dim vSQL

'---- 郵便番号辞書検索
vSQL = ""
vSQL = vSQL & "SELECT 都道府県名漢字"
vSQL = vSQL & "       , 市区町村名漢字"
vSQL = vSQL & "       , 町域名漢字"
vSQL = vSQL & "  FROM 郵便番号辞書 WITH (NOLOCK)"
vSQL = vSQL & " WHERE 郵便番号 = '" & Replace(vZip, "-", "") & "'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

If RSv.EOF = False Then
	check_zip = True
	vAddress = Trim(RSv("都道府県名漢字")) & Trim(RSv("市区町村名漢字"))
Else
	check_zip = False
	vAddress = ""
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

Connection.close
Set Connection= Nothing    '2011/08/01 an add

End function

'========================================================================

%>

<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="Shift_JIS">
<title>お問い合わせ内容の確認｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/inquiry.css" type="text/css">
<script type="text/javascript">
// ======	Function:	送信ボタンon click
function send_onClick(){
	document.f_data.submit();
}
// ======	Function:	変更ボタン on click
function return_onClick(){
	document.f_data.action = 'Inquiry.asp';
	document.f_data.submit();
}
</script>
</head>
<body>
<!--#include file="../Navi/NaviTop.inc"-->
<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>
  
  <!-- コンテンツstart -->
  <div id="globalContents">
    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="<%=g_HTTP%>"><img src="../images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li class="now">お問い合わせ内容の確認</li>
      </ul>
    </div></div></div>

    <h1 class="title">お問い合わせ内容の確認</h1>
    <p>お問い合わせ内容を確認の上、[送信する]ボタンを押してください。</p>
<table>
  <tr>
    <th>種 別</th>
    <td><%=ContactCategory%><br><%=ContactSubCategory%></td>
  </tr>
  <tr>
    <th>件 名</th>
    <td><%=subject%></td>
  </tr>
  <tr>
    <th>メッセージ</th>
    <td><%=wMessage%></td>
  </tr>
  <tr>
    <th>お名前</th>
    <td><%=customer_nm%></td>
  </tr>
  <tr>
    <th>フリガナ</th>
    <td><%=furigana%></td>
  </tr>
  <tr>
    <th>住 所</th>
    <td><%=zip%><br><%=prefecture%><%=address%></td>
  </tr>
  <tr>
    <th>電話番号</th>
    <td><%=telephone%></td>
  </tr>
  <tr>
    <th>FAX番号</th>
    <td><%=fax%></td>
  </tr>
  <tr>
    <th>E mail</th>
    <td><%=e_mail%></td>
  </tr>
</table>

<p>&laquo; <a href="JavaScript:return_onClick();">変更する</a></p>
      <form name="f_data" method="post" action="InquirySend.asp">
        <input type="hidden" name="message"            value="<% = message %>">
        <input type="hidden" name="subject"            value="<% = subject %>">
        <input type="hidden" name="ContactCategory"    value="<% = ContactCategory %>">
        <input type="hidden" name="ContactSubCategory" value="<% = ContactSubCategory %>">
        <input type="hidden" name="ContactSubCategoryFl" value="<% = ContactSubCategoryFl %>">
        <input type="hidden" name="customer_nm"        value="<% = customer_nm %>">
        <input type="hidden" name="furigana"           value="<% = furigana %>">
        <input type="hidden" name="zip"                value="<% = zip %>">
        <input type="hidden" name="prefecture"         value="<% = prefecture %>">
        <input type="hidden" name="address"            value="<% = address %>">
        <input type="hidden" name="telephone"          value="<% = telephone %>">
        <input type="hidden" name="fax"                value="<% = fax %>">
        <input type="hidden" name="e_mail"             value="<% = e_mail %>">
        <input type="hidden" name="Skey"               value="<% = Skey %>">
        <p class="btnBox"><input type="submit" value="送信する" class="opover"></p>
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