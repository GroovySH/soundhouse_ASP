<%@ LANGUAGE="VBScript" %>
<%
'�l�b�g�n�E�X�˂��ƃn�E�X�l�b�g�͂���
'�T�E���h�n�E�X
 Option Explicit
%>
<!--#include file="../common/ADOVBS.inc"-->
<!--#include file="../common/system_common.inc"-->
<!--#include file="../common/shop_common_functions.inc"-->
<!--#include file="../common/bfunctions1.asp"-->

<%
'========================================================================
'
'	�F�B�ɂ����߂�y�[�W
'
'�X�V����
'2009/04/30 �G���[����error.asp�ֈړ�
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2014/03/19 GV ����ő��łɔ���2�d�\���Ή�
'
'========================================================================

On Error Resume Next

Dim Item
Dim ItemCnt
Dim ItemList()
Dim MakerCd
Dim ProductCd
Dim MakerNm
Dim ProductNm
Dim FromName

Dim wPrice
Dim wSalesTaxRate
Dim wMailTrailer

Dim wProductHTML
Dim wMessage1
Dim wMessage1HTML

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
Dim wErrDesc   '2011/08/01 an add

'========================================================================

'---- Get input data
Item = ReplaceInput(Trim(Request("Item")))

if Item <> "" then
	ItemCnt = cf_unstring(Item, ItemList, "^")
	MakerCd = ItemList(0)
	ProductCd = ItemList(1)
end if

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "TellaFriend.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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
'========================================================================
Function main()

'---- ����ŗ���o��
call getCntlMst("����","����ŗ�","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'����ŗ�
wSalesTaxRate = Clng(wItemNum1)

'---- ���[���g���[�����o��
call getCntlMst("Web","Email","��ʃg���[��", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
wMailTrailer = wItemChar1
wMailTrailer = Replace(wMailTrailer, vbNewLine, "<br>")

'---- ���i�����o��
wSQL = ""
wSQL = wSQL & "SELECT a.���i�R�[�h"
wSQL = wSQL & "     , a.���i��"
wSQL = wSQL & "     , CASE"
wSQL = wSQL & "         WHEN (a.�����萔�� > a.������󒍍ϐ��� AND a.�����萔�� > 0) THEN a.������P��"
wSQL = wSQL & "         ELSE a.�̔��P��"
wSQL = wSQL & "       END AS �̔��P��"
wSQL = wSQL & "     , a.B�i�P��"
wSQL = wSQL & "     , a.B�i�t���O"
wSQL = wSQL & "     , a.���i�T��Web"
wSQL = wSQL & "     , a.���i�摜�t�@�C����_��"
wSQL = wSQL & "     , a.ASK���i�t���O"
wSQL = wSQL & "     , a.���[�J�[�R�[�h"
wSQL = wSQL & "     , b.���[�J�[��"
wSQL = wSQL & "  FROM Web���i a WITH (NOLOCK)"
wSQL = wSQL & "     , ���[�J�[ b WITH (NOLOCK)"
wSQL = wSQL & " WHERE b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
wSQL = wSQL & "   AND a.Web���i�t���O = 'Y'"
wSQL = wSQL & "   AND a.���[�J�[�R�[�h = '" & MakerCd & "'"
wSQL = wSQL & "   AND a.���i�R�[�h = '" & ProductCd & "'"
		
'@@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

'---- Move data to work area
if RS.EOF = true then
	exit function
else
	MakerNm = RS("���[�J�[��")
	ProductNm = RS("���i��")
end if

'---- ���i���ҏW
wHTML = ""
wHTML = wHTML & "<dl class='form_product'>" & vbNewLine
wHTML = wHTML & "  <dt>" & vbNewLine
wHTML = wHTML & "      <a href='ProductDetail.asp?Item=" & Server.URLEncode(Item) & "'><img src='../shop/prod_img/" & RS("���i�摜�t�@�C����_��") & "'></a>" & vbNewLine
wHTML = wHTML & "  </dt>" & vbNewLine
wHTML = wHTML & "  <dd>" & RS("���[�J�[��") & "</dd>" & vbNewLine
wHTML = wHTML & "  <dd>" & "<a href='ProductDetail.asp?Item=" & Server.URLEncode(Item) & "'>" & RS("���i��") & "</a></dd>" & vbNewLine
wHTML = wHTML & "  <dd>" & RS("���i�T��Web") & "</dd>" & vbNewLine
wHTML = wHTML & "  <dd>" & vbNewLine

wPrice = calcPrice(RS("�̔��P��"), wSalesTaxRate)

if RS("ASK���i�t���O") = "Y" then
	wHTML = wHTML & "�Ռ������F<a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RS("�̔��P��"),0) & "�~(�Ŕ�)</span>"
	wHTML = wHTML & "<span class='inc-tax'>(�ō�&nbsp;" & FormatNumber(wPrice,0) & "�~)</span></a>" & vbNewLine
else
	if RS("B�i�t���O") = "Y" then
		wHTML = wHTML & "�Ռ������F<del>" & FormatNumber(wPrice,0) & "�~(�ō�)</del><br>" & vbNewLine
		wPrice = calcPrice(RS("B�i�P��"), wSalesTaxRate)
'2014/03/19 GV mod start ---->
		wHTML = wHTML & "<strong>�킯����i�����F" & FormatNumber(RS("B�i�P��"),0) & "�~(�Ŕ�)</strong>" & vbNewLine
		wHTML = wHTML & "(�ō�&nbsp;" & FormatNumber(wPrice,0) & "�~)" & vbNewLine
'2014/03/19 GV mod end   <----
	else
'2014/03/19 GV mod start ---->
'		wHTML = wHTML & "�Ռ������F" & FormatNumber(wPrice,0) & "�~(�ō�)" & vbNewLine
		wHTML = wHTML & "�Ռ������F" & FormatNumber(RS("�̔��P��"),0) & "�~(�Ŕ�)" & vbNewLine
		wHTML = wHTML & "(�ō�&nbsp;" & FormatNumber(wPrice,0) & "�~)" & vbNewLine
'2014/03/19 GV mod end   <----
	end if
end if

wHTML = wHTML & "  </dd>" & vbNewLine
wHTML = wHTML & "</dl>" & vbNewLine

wProductHTML = wHTML

RS.close

'---- ���b�Z�[�W�w�b�_�ҏW
wHTML = ""
'wHTML = wHTML & FromName & "�@�l���" & vbNewLine
wHTML = wHTML & MakerNm & "�@" & ProductNm & "���������߂���܂����B" & vbNewLine
wHTML = wHTML & "���Ј�x�A�������������B" & vbNewLine
wHTML = wHTML & "http://www.soundhouse.co.jp/shop/ProductDetail.asp?Item=" & Server.URLEncode(Item)
wMessage1 = wHTML
wMessage1HTML = Replace(wHTML, vbNewLine, "<br>")

end function

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
<title>���F�B�ɂ����߂�b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/ask.css?20140401a" type="text/css">

<script type="text/javascript">

//
// ====== 	Function:	mail on submit
//
function mail_onSubmit(pForm){

	if (pForm.ToAddr.value == ""){
		alert("\n�������͂��Ă��������B");
		pForm.ToAddr.focus();
		return false;
 	}

		return true;
}

</script>

</head>

<body>

<!--#include file="../Navi/NaviTop.inc"-->

<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>
  
  <!-- �R���e���cstart -->
  <div id="globalContents">
    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="<%=g_HTTP%>"><img src="<%=g_RelLink%>images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li class="now">���F�B�ɂ����߂�</li>
      </ul>
    </div></div></div>

    <h1 class="title">���F�B�ɂ����߂�</h1>

<%=wProductHTML%>

<form name="fMail" type="post" action="TellaFriendSend.asp" onSubmit="return mail_onSubmit(this);">
<table class="form">
  <tr>
    <th>���M�Җ�</th>
    <td><input name="FromName" type="text" size="60"></td>
  </tr>
  <tr>
    <th>����</th>
    <td><input name="ToAddr" type="text" size="60"><div>����̃��[���A�h���X�������͂��������B</div></td>
  </tr>
  <tr>
    <th>���b�Z�[�W</th>
    <td><p><%=wMessage1HTML%></p><textarea name="Message" cols="70" rows="20"></textarea></td>
  </tr>
</table>
<p>���ē��������܂������e�ɂ��܂��ĕs���ȓ_���������܂�����A���Љc�Ƃ܂ł��A�����������܂��悤���肢�������܂��B</p>
<p>
	-------------------------------------------------------<br>
	������ЃT�E���h�n�E�X<br>
	HP�F http://www.soundhouse.co.jp/<br>
	Email �F shop@soundhouse.co.jp<br>
	TEL �F 0476-89-1111<br>
	FAX �F 0476-89-2222<br>
	�i���`���F10-19���A�y�F12-17���A���j�E�j�Փ��������j<br>
	-------------------------------------------------------
</p>
<p class="btnBox"><input type="submit" value="���M" class="opover"></p>
<input type="hidden" name="Item" value="<%=Item%>">
<input type="hidden" name="Message1" value="<%=wMessage1%>">
</form>

</div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<div class="tooltip"><p>ASK</p></div>
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript" src="jslib/ask.js?20140401a"></script>
</body>
</html>