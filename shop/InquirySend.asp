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
'	問合せ送信
'
'更新履歴
'2005/05/13 OBJ_NewMail.BodyPart.Charset = "iso-2022-jp"をセット
'2005/08/18 コンタクトカテゴリー、コンタクトサブカテゴリーを追加
'2005/09/06 コンタクト管理テスト対応
'2005/09/07 自動返信メール送信
'2005/09/08 問合せ内容をWebコンタクトテーブルへ格納するように変更
'2006/08/11 入力データチェック強化
'2008/05/12 改行コードインジェクション対策（i_toパラメータ削除）
'2008/05/13 クロスサイトリクエストフォジェリー対策 Keyパラメータチェック
'2008/05/23 入力データチェック強化（LEFT他)
'2009/04/30 エラー時にerror.aspへ移動
'2009/09/03 自動返信内容を追記
'2010/01/08 an 削除する改行コードの指定をvbCr/vbLfに変更
'2011/04/13 an #725に合わせエラーチェック強化
'2011/04/14 hn Session関連変更
'2011/08/01 an #1087 Error.aspログ出力対応, wErr→Err.Descriptionに戻した（機能していないため）
'2012/06/29 if-web リニューアルレイアウト調整
'
'========================================================================

On Error Resume Next

Dim userID
'Dim msg  '2011/04/13 an del

Dim message
Dim subject
Dim ContactCategory
Dim ContactSubCategory
Dim ContactSubCategoryFl  '2011/04/20 an add
Dim customer_nm
Dim furigana
Dim zip
Dim prefecture
Dim address
Dim telephone
Dim fax
Dim e_mail

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim Connection
Dim RS

Dim wSQL
Dim wHTML
Dim wMSG
'Dim wErr		'2011/04/13, 2011/08/01 an del
Dim wErrDesc    '2011/08/01 an add

'========================================================================

Response.buffer = true

'---- セキュリティーキーチェック
if Session("SKey") <> ReplaceInput(Request("SKey")) then
	Response.redirect "Inquiry.asp"
end if

'---- UserID 取り出し
userID = Session("userID")

'---- 呼び出し元からのデータ取り出し
message = ReplaceInput(Left(Request("message"),2000))
subject = ReplaceInput(Left(Request("subject"),150))
ContactCategory = ReplaceInput(Left(Request("ContactCategory"),20))
ContactSubCategory = ReplaceInput(Left(Request("ContactSubCategory"),20))
ContactSubCategoryFl = ReplaceInput_NoCRLF(Left(Request("ContactSubCategoryFl"),1))  '2011/04/13 an add
customer_nm = ReplaceInput(Left(Request("customer_nm"),30))
furigana = ReplaceInput(Left(Request("furigana"),30))
zip = ReplaceInput(Left(Request("zip"),8))
prefecture = ReplaceInput(Left(Request("prefecture"),8))
address = ReplaceInput(Left(Request("address"),40))
e_mail = ReplaceInput_NoCRLF(Left(Request("e_mail"),60)) '2010/01/08 an  2011/04/13 an mod
telephone = ReplaceInput(Left(Request("telephone"),20))
fax = ReplaceInput(Left(Request("fax"),20))

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	Connection.RollbackTrans
	wErrDesc = "InquirySend.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if wMsg <> "" then  '2011/04/13 an add s 通常はInquiryConfirmを経由するのでここでエラーは起きないためTransferはしない
    Response.Redirect g_HTTPS & "shop/Inquiry.asp"
end if              '2011/04/13 an add e

'if wErr <> "" then	           '2011/08/01 an del
if Err.Description <> "" then  '2011/08/01 an add
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
Dim v_body
Dim v_subject
Dim OBJ_NewMail

'---- 入力チェック      '2011/04/13 an add s
call validation()

if wMsg <> "" then
	exit function
end if

Connection.BeginTrans   '2011/04/13 an add e

'---- Webコンタクトへ格納
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM Webコンタクト"
wSQL = wSQL & " WHERE 1 = 2"
	  
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic, adLockOptimistic

RS.AddNew

RS("コンタクト日") = Now()
if userID <> "" then
	RS("顧客番号") = userID
else
	RS("顧客番号") = 0
end if

RS("入力顧客名") = customer_nm
RS("入力顧客フリガナ") = furigana
RS("入力顧客Email") = e_mail
RS("入力顧客電話番号") = telephone
RS("入力顧客Fax") = fax
RS("入力顧客郵便番号") = zip
RS("入力顧客都道府県") = prefecture
RS("入力顧客住所") = address
RS("コンタクトカテゴリー") = ContactCategory
RS("コンタクトサブカテゴリー") = ContactSubCategory
RS("宛先") = "shop@soundhouse.co.jp"

v_subject = subject & " " & customer_nm & "様"
RS("件名") = v_subject
RS("本文") = message

RS.Update

RS.close

'wErr = Err.Description    '2011/04/13, 2011/08/01 an del

if Err.Description = "" then    '2011/04/13 an add 2011/08/01 an mod
                                
	Connection.CommitTrans		'Commit   '2011/04/13 an add
                                          
	'---- 自動返信メール作成（顧客へ)     
	call getCntlMst("Web","Email","トレーラ", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
                                          
	v_body = "お問い合わせありがとうございます｡" & vbNewLine _
	       & "以下の内容にてお問い合わせを受付いたしました｡" & vbNewLine _
	       & "返答まで今しばらくお待ちください｡" & vbNewLine & vbNewLine _
	       & "弊社からの返信が24時間以内（休業日を除く）に到着しない場合には、お手数ですが" & vbNewLine _
	       & "その旨お問い合わせいただきますようお願い申し上げます。" & vbNewLine & vbNewLine

	v_body = v_body & "受付日時　　　：" & now() & vbNewLine & vbNewLine
	v_body = v_body & "件名　　　　　：" & subject & vbNewLine & vbNewLine
	v_body = v_body & "カテゴリー　　：" & ContactCategory & vbNewLine
	v_body = v_body & "サブカテゴリー：" & ContactSubCategory & vbNewLine
	v_body = v_body & "メッセージ　　：" & message & vbNewLine & vbNewLine

	v_body = v_body & wItemChar1

	Set OBJ_NewMail = Server.CreateObject("CDO.Message") 

	OBJ_NewMail.from = "shop@soundhouse.co.jp"
	OBJ_NewMail.to = e_mail

	OBJ_NewMail.subject = "お問い合わせを受付いたしました"
	OBJ_NewMail.TextBody = v_body
	OBJ_NewMail.BodyPart.Charset = "iso-2022-jp"

	OBJ_NewMail.Send

	Set OBJ_NewMail = Nothing

else                                               '2011/04/13 an add s
	Connection.RollbackTrans	'Rollback
end if

End function                                       '2011/04/13 an add e

'========================================================================
'
'    Function    問合せ入力内容チェック   '2011/04/13 an add
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
'	Function	郵便番号辞書検索   '2011/04/13 an add
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
<title>お問い合わせを送信しました｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/inquiry.css" type="text/css">
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
        <li class="now">お問い合わせを送信しました</li>
      </ul>
    </div></div></div>

    <h1 class="title">お問い合わせを送信しました</h1>
    <p>お問合せを送信しました。<br>ありがとうございました。</p>
  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>