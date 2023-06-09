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
'	問合せページ
'
'更新履歴
'2005/01/31 パラメータを渡された時は、件名へ自動表示
'2005/05/24 Submitを<a>から<input type="image" に変更
'2005/08/18 コンタクトカテゴリー、コンタクトサブカテゴリーを追加
'2005/09/29 コンタクトカテゴリー(Web-Emax)は対象外
'2006/01/09 Emalチェック強化
'2007/10/19 ハッカーセーフ対応
'2008/05/12 改行コードインジェクション対策（i_toパラメータ削除）
'2008/05/13 クロスサイトリクエストフォジェリー対策 Keyパラメータセット
'2009/04/30 エラー時にerror.aspへ移動
'2011/03/02 hn SetSecureKeyの位置変更
'2011/04/13 an #725 確認画面(InquiryConfirm)追加、エラーメッセージ表示対応
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/06/29 if-web リニューアルレイアウト調整
'
'========================================================================

On Error Resume Next

Dim userID

Dim CategoryNm
Dim MakerNm
Dim ProductCd
'Dim wSubject     '2011/04/13 an del
Dim wCategoryListHTML

Dim ContactCategory     '2011/04/13 an add
Dim ContactSubCategory  '2011/04/13 an add
Dim subject             '2011/04/13 an add
Dim message             '2011/04/13 an add
Dim customer_nm
Dim furigana
Dim zip
Dim prefecture
Dim address
Dim telephone
Dim fax
Dim e_mail

'Dim Skey   '2011/04/13 an del

Dim Connection
Dim RS

Dim w_sql
Dim w_error_msg
Dim wHTML
Dim wMsg
Dim wErrDesc   '2011/08/01 an add

'========================================================================

wMsg = ""

'---- Get Session data  2011/04/13 an add
wMsg = Session("msg")
Session("msg") = ""

'---- Get input data
CategoryNm = ReplaceInput(Trim(Request("CategoryNm")))
MakerNm = ReplaceInput(Trim(Request("MakerNm")))
ProductCd = ReplaceInput(Trim(Request("ProductCd")))

'---- 入力チェックエラー時にInquiryConfirm.aspからデータ受け取り  2011/04/13 an add s
ContactCategory = ReplaceInput(Left(Request("ContactCategory"),20))
ContactSubCategory = ReplaceInput(Left(Request("ContactSubCategory"),20))
subject = ReplaceInput(Left(Request("subject"),150))
message = ReplaceInput(Left(Request("message"),2000))
customer_nm = ReplaceInput(Left(Request("customer_nm"),30))
furigana = ReplaceInput(Left(Request("furigana"),30))
zip = ReplaceInput(Left(Request("zip"),8))
prefecture = ReplaceInput(Left(Request("prefecture"),8))
address = ReplaceInput(Left(Request("address"),40))
telephone = ReplaceInput(Left(Request("telephone"),20))
fax = ReplaceInput(Left(Request("fax"),20))
e_mail = ReplaceInput_NoCRLF(Left(Request("e_mail"),60))  '2011/04/13 an add e

'wSubject = ""  '2011/04/13 an del
if (CategoryNm <> "") OR (MakerNm <> "" ) OR (ProductCd <> "") then
	'wSubject = "(" & MakerNm & "/" & ProductCd & ")"   '2011/04/13 an del
	subject = "（" & MakerNm & "/" & ProductCd & "）"
end if

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "Inquiry.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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
'	Function	main
'
'				userIDがCookieにあれば会員情報を表示
'
'========================================================================
Function main()

'---- セキュリティーキーセット 
'Skey = SetSecureKey() '2011/04/13 an del

'---- コンタクトカテゴリー、コンタクトサブカテゴリ一覧作成
call CreateCategoryListHTML()

if wMsg = "" then  'ログインしていれば顧客情報取り出し(InquiryConfirm.aspから戻った場合は入力データを表示) 2011/04/13 an add

    userID = Session("userID")

    if userID = "" then
    	exit function
    end if

    '--------- select customer
    w_sql = ""
    w_sql = w_sql & "SELECT a.顧客番号"
    w_sql = w_sql & "     , a.顧客名"
    w_sql = w_sql & "     , a.顧客フリガナ"
    w_sql = w_sql & "     , a.顧客E_mail1"
    w_sql = w_sql & "     , b.顧客郵便番号"
    w_sql = w_sql & "     , b.顧客都道府県"
    w_sql = w_sql & "     , b.顧客住所"
    w_sql = w_sql & "     , c.顧客電話番号"
    w_sql = w_sql & "  FROM Web顧客 a WITH (NOLOCK)"
    w_sql = w_sql & "     , Web顧客住所 b WITH (NOLOCK)"
    w_sql = w_sql & "     , Web顧客住所電話番号 c WITH (NOLOCK)"
    w_sql = w_sql & " WHERE b.顧客番号 = a.顧客番号" 
    w_sql = w_sql & "   AND c.顧客番号 = b.顧客番号" 
    w_sql = w_sql & "   AND c.住所連番 = b.住所連番" 
    w_sql = w_sql & "   AND b.住所連番 = 1" 
    w_sql = w_sql & "   AND c.電話連番 = 1" 
    w_sql = w_sql & "   AND a.顧客番号 = " & userID 
    		
    '@@@@@response.write(w_sql & "<BR>")

    Set RS = Server.CreateObject("ADODB.Recordset")
    RS.Open w_sql, Connection, adOpenStatic

    '-------- Move data to work area
    if RS.EOF = true then
    	exit function
    else
    	customer_nm = RS("顧客名")
    	furigana = RS("顧客フリガナ")
    	zip = RS("顧客郵便番号")
    	prefecture = RS("顧客都道府県")
    	address = RS("顧客住所")
    	telephone = RS("顧客電話番号")
    	e_mail = RS("顧客E_mail1")
    end if

    RS.close
end if    '2011/04/13 an add

end function

'========================================================================
'
'	Function	CreateCategoryListHTML
'
'========================================================================
Function CreateCategoryListHTML()

Dim RSv
Dim vBreakKey1
Dim vBreakNextKey1
Dim vRecCount

'--------- コンタクトカテゴリー、コンタクトサブカテゴリ一取り出し
w_sql = ""
w_sql = w_sql & "SELECT a.コンタクトカテゴリー名"
w_sql = w_sql & "     , b.コンタクトサブカテゴリー名"
w_sql = w_sql & "  FROM コンタクトカテゴリー a WITH (NOLOCK)"
w_sql = w_sql & "       LEFT JOIN コンタクトサブカテゴリー b WITH (NOLOCK)"
w_sql = w_sql & "              ON b.コンタクトカテゴリーコード = a.コンタクトカテゴリーコード" 
w_sql = w_sql & " WHERE a.Web非表示フラグ != 'Y'"
w_sql = w_sql & " ORDER BY"
w_sql = w_sql & "       a.コンタクトカテゴリーコード" 
w_sql = w_sql & "     , b.コンタクトサブカテゴリーコード" 
		
'@@@@@response.write(w_sql & "<BR>")

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open w_sql, Connection, adOpenStatic

'-------- 一覧作成
if RSv.EOF = true then
	exit function
end if

vBreakNextKey1 = RSv("コンタクトカテゴリー名")
vBreakKey1 = vBreakNextKey1

wHTML = ""
vRecCount = 0

'---- Main loop
Do Until vBreakNextKey1 = "@EOF"
	vRecCount = vRecCount + 1
	'---- カテゴリーラジオボタン
	wHTML = wHTML & "          <li>" & vbNewLine
	wHTML = wHTML & "            <input type=""radio"" name=""ContactCategorySel"" value=""" & RSv("コンタクトカテゴリー名") & """"

	If (CategoryNm = "") And (vRecCount = 1) Then
		wHTML = wHTML & " checked=""checked"""
	End If

	wHTML = wHTML & " id=""type_" & vRecCount & """>"
	wHTML = wHTML & "<label for=""type_" & vRecCount & """>" & RSv("コンタクトカテゴリー名") & "</label>" & vbNewLine

	If IsNull(RSv("コンタクトサブカテゴリー名")) = False Then
		wHTML = wHTML & "            <select name=""ContactSubCategorySel"" id=""subcategory" & vRecCount - 1 & """ onChange=""SubCategory_onChange('" & RSv("コンタクトカテゴリー名") & "')"">" & vbNewLine
		wHTML = wHTML & "              <option value="""">選択してください" & vbNewLine

		vBreakKey1 = vBreakNextKey1

		Do Until vBreakKey1 <> vBreakNextKey1      'カテゴリーブレークまで
			'---- サブカテゴリーSELECT OPTIONS
			wHTML = wHTML & "              <option value=""" & RSv("コンタクトサブカテゴリー名") & """"
			If CategoryNm = RSv("コンタクトサブカテゴリー名") Then
				wHTML = wHTML & " SELECTED"
			End If
			wHTML = wHTML & ">" & RSv("コンタクトサブカテゴリー名") & vbNewLine

			RSv.MoveNext
			If RSv.EOF = True Then
				vBreakNextKey1 = "@EOF"
			Else
				vBreakNextKey1 = RSv("コンタクトカテゴリー名")
			End If
		Loop

		wHTML = wHTML & "            </select>" & vbNewLine
		wHTML = wHTML & "          </li>" & vbNewLine
	Else
		wHTML = wHTML & "          </li>" & vbNewLine
		RSv.MoveNext
		If RSv.EOF = True Then
			vBreakNextKey1 = "@EOF"
		Else
			vBreakNextKey1 = RSv("コンタクトカテゴリー名")
		End If
	End If

Loop

RSv.Close

wCategoryListHTML = wHTML

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
<title>お問い合わせ｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/inquiry.css" type="text/css">

<script type="text/javascript">
//
// ====== 	Function:	check if some data was entered other than spaces
//		Parm:		p_val		Check value
//		Return value:	If entered --> True,  Not entered --> False
//
function check_required(p_val){
	if (p_val == ""){return(false);}
	for(i=0; i<p_val.length; i++){
		if (p_val.substring(i, i+1)!=" " && p_val.substring(i, i+1)!="　"){
			return(true);
		}
	}
	return(false);
}

//=====================================================================
//	住所検索 onClick
//=====================================================================
function address_search_onClick(){

	if (document.f_data.zip.value == ""){
		alert("郵便番号を入力してください。");
		return;
	}
 
	AddrWin = window.open("../comasp/Address_search.asp?zip=" + document.f_data.zip.value +"&name_prefecture=i_selected_prefecture&name_address=address","AddrSearch","width=200,height=100");
}

//
// ====== 	Function:	next on submit
//
function next_onSubmit(){
	for (var i=0; i<document.f_data.ContactCategorySel.length; i++){
		if (document.f_data.ContactCategorySel[i].checked == true){
			document.f_data.ContactCategory.value = document.f_data.ContactCategorySel[i].value;
			
			var subcategory = document.getElementById('subcategory' + i);
			if (subcategory != null){  //サブカテゴリー有の場合
				for (var j=0; j<subcategory.options.length; j++){
					if (subcategory.options[j].selected == true){
					    document.f_data.ContactSubCategory.value = subcategory.options[j].value;
						break;
					}
				}
			}else{
				document.f_data.ContactSubCategory.value = "";
				document.f_data.ContactSubCategoryFl.value = "N";
			}
			break;
		}
	}
	return;
}

//
// ======	Function:	ラジオボタン、ドロップダウンリストを以前に選択した状態にする
//
function preset_values(){

    // 種別カテゴリー
	for (var i=0; i<document.f_data.ContactCategorySel.length; i++){
		if (document.f_data.ContactCategorySel[i].value == document.f_data.ContactCategory.value){
			document.f_data.ContactCategorySel[i].checked = true;
			//種別サブカテゴリー
			if (document.f_data.ContactSubCategory.value != "" ){
			
				var subcategory = document.getElementById('subcategory' + i);
				for (var j=0; j<subcategory.options.length; j++){
			        if (subcategory.options[j].value == document.f_data.ContactSubCategory.value){
				        subcategory.options[j].selected = true;
				        break;
			        }
		        }
		    }
		break;
		}
	}

    // 都道府県
	for (var i=0; i<document.f_data.prefecture.options.length; i++){
		if (document.f_data.prefecture.options[i].value == document.f_data.i_selected_prefecture.value){
			document.f_data.prefecture.options[i].selected = true;
			break;
		}
	}
	return;
}

//
// Function: サブカテゴリー変更時に親カテゴリーを選択
//
function SubCategory_onChange(pSubCategoryValue){

	for (var i=0; i<document.f_data.ContactCategorySel.length; i++){
		if (document.f_data.ContactCategorySel[i].value == pSubCategoryValue){
			document.f_data.ContactCategorySel[i].checked = true;
			break;
		}
	}
	return;
}

//========================================================================

</script>

</head>

<body>

<!--#include file="../Navi/NaviTop.inc"-->

<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>
  
  <!-- コンテンツstart -->
  <div id="globalContents">
    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="<%=g_HTTP%>"><img src="<%=g_RelLink%>images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li class="now">お問い合わせ</li>
      </ul>
    </div></div></div>

    <h1 class="title">お問い合わせ</h1>
    <div id="notice">
      <p>お問い合わせいただく前によくあるご質問ををご確認ください</p>
      <ul>
        <li><a href="http://guide.soundhouse.co.jp/guide/qanda7.asp">E-Mailが届かない</a></li>
        <li><a href="http://guide.soundhouse.co.jp/guide/qanda5.asp">返品・交換</a></li>
        <li><a href="http://guide.soundhouse.co.jp/guide/qanda1.asp">ご注文</a></li>
        <li><a href="http://guide.soundhouse.co.jp/guide/qanda4.asp">お届け</a></li>
        <li><a href="http://guide.soundhouse.co.jp/guide/qanda11.asp">領収書</a></li>
      </ul>
    </div>

<form name="f_data" id="inquiry" action="<%=g_HTTPS%>shop/InquiryConfirm.asp" method="post" onSubmit="return next_onSubmit();">

  <!-- エラーメッセージ -->
  <% If wMsg <> "" Then %>
  <ul class="error">
    <li><%=wMsg %></li>
  </ul>
  <% End If %>

<table>
  <tr>
    <th>種別<span>*</span></th>
    <td>
      <ul>
<% = wCategoryListHTML %>
      </ul>
    </td>
  </tr>
  <tr>
    <th>件名<span>*</span></th>
    <td><input type="text" name="subject" size="60" value="<%=subject%>"></td>
  </tr>
  <tr>
    <th>メッセージ<span>*</span></th>
    <td><textarea name="message" rows="5" cols="60"><%=message%></textarea></td>
  </tr>
  <tr>
    <th>お名前<span>*</span></th>
    <td><input type="text" name="customer_nm" size="40" maxlength="30" value="<%=customer_nm%>"></td>
  </tr>
  <tr>
    <th>フリガナ</th>
    <td><input type="text" name="furigana" size="40" maxlength="30" value="<%=furigana%>"><span>（全角カナ）</span></td>
  </tr>
  <tr>
    <th>住 所</th>
    <td>
      〒<input type="text" name="zip" size="10" maxlength="8" value="<%=zip%>">（半角数字）<a href="JavaScript:address_search_onClick();" class="tipBtn">住所検索</a>郵便番号を入力してボタンを押してください｡<br>
      <input type="hidden" name="i_selected_prefecture" value="<%=prefecture%>">
      <select name="prefecture" size="1">
        <option value="">都道府県</option>
        <option value="北海道">北海道</option>
        <option value="青森県">青森県</option>
        <option value="秋田県">秋田県</option>
        <option value="岩手県">岩手県</option>
        <option value="宮城県">宮城県</option>
        <option value="山形県">山形県</option>
        <option value="福島県">福島県</option>
        <option value="栃木県">栃木県</option>
        <option value="新潟県">新潟県</option>
        <option value="群馬県">群馬県</option>
        <option value="埼玉県">埼玉県</option>
        <option value="茨城県">茨城県</option>
        <option value="千葉県">千葉県</option>
        <option value="東京都">東京都</option>
        <option value="神奈川県">神奈川県</option>
        <option value="山梨県">山梨県</option>
        <option value="長野県">長野県</option>
        <option value="岐阜県">岐阜県</option>
        <option value="富山県">富山県</option>
        <option value="石川県">石川県</option>
        <option value="静岡県">静岡県</option>
        <option value="愛知県">愛知県</option>
        <option value="三重県">三重県</option>
        <option value="奈良県">奈良県</option>
        <option value="和歌山県">和歌山県</option>
        <option value="福井県">福井県</option>
        <option value="滋賀県">滋賀県</option>
        <option value="京都府">京都府</option>
        <option value="大阪府">大阪府</option>
        <option value="兵庫県">兵庫県</option>
        <option value="岡山県">岡山県</option>
        <option value="鳥取県">鳥取県</option>
        <option value="島根県">島根県</option>
        <option value="広島県">広島県</option>
        <option value="山口県">山口県</option>
        <option value="香川県">香川県</option>
        <option value="徳島県">徳島県</option>
        <option value="愛媛県">愛媛県</option>
        <option value="高知県">高知県</option>
        <option value="福岡県">福岡県</option>
        <option value="佐賀県">佐賀県</option>
        <option value="大分県">大分県</option>
        <option value="熊本県">熊本県</option>
        <option value="宮崎県">宮崎県</option>
        <option value="長崎県">長崎県</option>
        <option value="鹿児島県">鹿児島県</option>
        <option value="沖縄県">沖縄県</option>
      </select>
      <input type="text" name="address" size="60" maxlength="40"  value="<%=address%>">
    </td>
  </tr>
  <tr>
    <th>電話番号<span>*</span></th>
    <td><input type="text" name="telephone" size="30" maxlength="20" value="<%=telephone%>" class="validate required">（半角数字）</td>
  </tr>
  <tr>
    <th>FAX番号</th>
    <td><input type="text" name="fax" size="30" maxlength="20" value="<%=fax%>">（半角数字）</td>
  </tr>
  <tr>
    <th>E-mail<span>*</span></th>
    <td><input type="text" name="e_mail" size="30" maxlength="60" value="<%=e_mail%>" class="validate mail required">（半角英数字）</td>
  </tr>
</table>
<p>「*」のついている項目は必須入力項目です。</p>
<input type="hidden" name="ContactCategory" value="<%=ContactCategory%>">
<input type="hidden" name="ContactSubCategory" value="<%=ContactSubCategory%>">
<input type="hidden" name="ContactSubCategoryFl" value="<%=ContactSubCategoryFl%>">
<p class="btnBox"><input type="submit" value="内容を確認する" class="opover"></p>
</form>

</div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript">
	preset_values();
</script>
</body>
</html>