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
'	���i��r�y�[�W ���̂܂ܔ�r�ł��邩�ǂ����`�F�b�N�@(�J�e�S���[����/5�ȏ�̎� �I����ʕ\��)
'
'	�X�V����
'2008/05/07 ��؂蕶���ύX
'2009/04/30 �G���[����error.asp�ֈړ�
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/01/20 GV �f�[�^�擾 SELECT���� LAC�N�G���[�Ă�K�p
'2012/07/18 nt ���j���[�A���p�Ƀf�[�^�擾 SELECT�������asp��ʏo�͂��C��
'2014/01/31 GV �{�Ԋ��ŃN�b�L�[���擾�ł��Ȃ����ۂ��C��
'
'========================================================================

On Error Resume Next

Dim wNaveWithLink '2012/7/19 nt add
Dim wTitleWithLink

Dim wHikaku
Dim CategoryCd()
Dim MakerCd()
Dim ProductCd()
Dim Iro()
Dim Kikaku()
Dim MakerName()
Dim ProductName()
Dim wRecCnt
Dim wGotoCompareFl
Dim wParm

Dim Connection
Dim RS

Dim i
Dim wHTML
Dim wSQL
Dim wMsg
Dim wErrDesc   '2011/08/01 an add

Dim category_cd '2012/7/19 nt add

'========================================================================

Response.Buffer = true

'---- Execute main

call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "ProductCompareCheck.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

if wGotoCompareFl = true then
	For i= 1 to wRecCnt
		wParm = wParm & "$" & CategoryCd(i) & "^" & MakerCd(i) & "^" & ProductCd(i) & "^" & Trim(Iro(i)) & "^" & Trim(kikaku(i))
	Next
	Response.redirect "ProductCompare.asp?item=" & wParm
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

Dim vMoreThanOneCategoryFl
Dim vMinData
Dim vMinSub
Dim vOldCategoryCd
Dim vProductName
Dim i
Dim j
Dim vTemp
Dim vCookieData		'2014/01/31 GV add

'---- ���M�f�[�^�[�̎��o��

'---- ��rCookie���o��
'     1����(�Y����0)�̓_�~�[�f�[�^�̂��ߖ���
'2014/01/31 GV mod start
vCookieData = Request.Cookies("compare")
'�ʏ��Response.Cookies()�Ńf�[�^���擾�ł��Ȃ������ꍇ�A
'��p�̃v���V�[�W��(Shop_common_functions.inc)���g���B
If (Trim(vCookieData) = "") Or (IsNull(Trim(vCookieData)) = true)  Then
	vCookieData = getCookieValue("compare")
End If
'wHikaku = Split(Request.Cookies("compare"), "$")
wHikaku = Split(vCookieData, "$")
'2014/01/31 GV mod end
wRecCnt = Ubound(wHikaku)

ReDim CategoryCd(wRecCnt+1)
ReDim MakerCd(wRecCnt)
ReDim ProductCd(wRecCnt)
ReDim Iro(wRecCnt)
ReDim Kikaku(wRecCnt)
ReDim MakerName(wRecCnt)
ReDim ProductName(wRecCnt)

For i=1 to wRecCnt
	vTemp = Split(wHikaku(i), "^")
	CategoryCd(i) = ReplaceInput(Trim(vTemp(0)))
	MakerCd(i) = ReplaceInput(Trim(vTemp(1)))
	ProductCd(i) = ReplaceInput(Trim(vTemp(2)))
	Iro(i) = ReplaceInput(Trim(vTemp(3)))
	Kikaku(i) = ReplaceInput(Trim(vTemp(4)))
Next

'---- �����J�e�S���[�`�F�b�N
vMoreThanOneCategoryFl = false
For i=2 to wRecCnt
	if CategoryCd(1) <> CategoryCd(i) then
		vMoreThanOneCategoryFl = true
		Exit For
	end if
Next

if vMoreThanOneCategoryFl = true OR wRecCnt > 5 then

'---- ��r���i�f�[�^���o��
	call getCompareProduct()

'---- �J�e�S���[�C���[�J�[���C���i�����Ƀ\�[�g
	call SortProduct()

'---- �J�e�S���[�ʂɔ�r���i�ꗗ�쐬
	i = 1
	wHTML = ""

'---- �i�r�Q�[�V�����Z�b�g
	call SetNavi(i)
	wHTML = wHTML & wNaveWithLink

	Do until i > wRecCnt

		'---- �J�e�S���[�^�C�g���Z�b�g (�J�e�S���[�u���[�N)
		if vOldCategoryCd <> CategoryCd(i) then
			call SetTitle(i)
			wHTML = wHTML & wTitleWithLink

			'2012/07/18 nt add
			wHTML = wHTML & "<form onSubmit='return Hikaku_onSubmit(this);'>" & vbNewLine
			wHTML = wHTML & "<dl class='productcompare'>" & vbNewLine

			'2012/07/18 nt del
			'wHTML = wHTML & "<table border='1' cellspacing='0' cellpadding='3'>" & vbNewLine
			'wHTML = wHTML & "<form onSubmit='return Hikaku_onSubmit(this);'>" & vbNewLine

			'---- �^�C�g��
			'2012/07/18 nt add
			wHTML = wHTML & "<dt>" & vbNewLine
			wHTML = wHTML & "<ul>" & vbNewLine
			wHTML = wHTML & "<li>��r</li>" & vbNewLine
			wHTML = wHTML & "<li>���i��</li>" & vbNewLine
			wHTML = wHTML & "</ul>" & vbNewLine
			wHTML = wHTML & "</dt>" & vbNewLine
			wHTML = wHTML & "<dd>" & vbNewLine
			wHTML = wHTML & "<ul>" & vbNewLine

			'2012/07/18 nt del
			'wHTML = wHTML & "  <tr bgcolor='#cccccc' class='honbun'>"
			'wHTML = wHTML & "    <td width='50' align='center' nowrap>��r</td>" & vbNewLine
			'wHTML = wHTML & "    <td width='500' align='center' nowrap>���i��</td>" & vbNewLine
			'wHTML = wHTML & "  </tr>"

			vOldCategoryCd = CategoryCd(i)
		end if

		'---- ���i��/�F/�K�i
		'2012/07/18 nt add
		vProductName = ProductName(i)
		if Trim(Iro(i)) <> "" then
			vProductName = vProductName & "/" & Trim(Iro(i))
		end if
		if Trim(Kikaku(i)) <> "" then
			vProductName = vProductName & "/" & Trim(Kikaku(i))
		end if

		'---- �I���`�F�b�N�{�b�N�X
		'2012/07/18 nt add
		wHTML = wHTML & "          <li><span><input type='checkbox' name='iItem' value='$" & CategoryCd(i) & "^" & MakerCd(i) & "^" & ProductCd(i) & "^" & Trim(Iro(i)) & "^" & Trim(kikaku(i)) & "' checked></span>" & MakerName(i) & "<a href='ProductDetail.asp?item=" & MakerCd(i) & "^" & ProductCd(i) & "^" & Trim(Iro(i)) & "^" &  Trim(kikaku(i)) & "'>" & vProductName & "</a></li>" & vbNewLine

		'2012/07/18 nt del
		'wHTML = wHTML & "  <tr>"
		'wHTML = wHTML & "    <td align='center' valign='middle' nowrap>" & vbNewLine
		'wHTML = wHTML & "      <input type='checkbox' name='iItem' value='$" & CategoryCd(i) & "^" & MakerCd(i) & "^" & ProductCd(i) & "^^" & Trim(Iro(i)) & "^" & Trim(kikaku(i)) & "' CHECKED>" & vbNewLine
		'wHTML = wHTML & "    </td>" & vbNewLine

		'2012/07/18 nt del
		'---- ���[�J�[
		'wHTML = wHTML & "    <td align='left' nowrap>" & vbNewLine
		'wHTML = wHTML & "      <span class='honbun'>" & MakerName(i) & "</span><br>" & vbNewLine

		'2012/07/18 nt del
		'---- ���i��/�F/�K�i
		'vProductName = ProductName(i)
		'if Trim(Iro(i)) <> "" then
		'	vProductName = vProductName & "/" & Trim(Iro(i))
		'end if
		'if Trim(Kikaku(i)) <> "" then
		'	vProductName = vProductName & "/" & Trim(Kikaku(i))
		'end if

		'2012/07/18 nt del
		'wHTML = wHTML & "    <a href='ProductDetail.asp?item=" & MakerCd(i) & "^" & ProductCd(i) & "^" & Iro(i) & "^" & Kikaku(i) & "' class='link'>" & vProductName & "</a>" & vbNewLine
		'wHTML = wHTML & "    </td>" & vbNewLine
		'wHTML = wHTML & "  </tr>" & vbNewLine

		'---- ���f�[�^���`�F�b�N
		i = i + 1

		'---- ��r�{�^��
		if i > wRecCnt OR vOldCategoryCd <> CategoryCd(i) then
			'2012/07/18 nt add
			wHTML = wHTML & "</ul>" & vbNewLine
			wHTML = wHTML & "</dd>" & vbNewLine
			wHTML = wHTML & "</dl>" & vbNewLine
			wHTML = wHTML & "<p class='btnBox'><input type='submit' value='��r����' class='opover'></p>"
			wHTML = wHTML & "</form>"

			'2012/07/18 nt del
			'wHTML = wHTML & "  <tr>"
			'wHTML = wHTML & "    <td align='center' valign='middle' colspan=2>" & vbNewLine
			'wHTML = wHTML & "      <input type='image' src='images/Hikaku2.gif' border=0>" & vbNewLine
			'wHTML = wHTML & "    </td>" & vbNewLine
			'wHTML = wHTML & "  </tr>"
			'wHTML = wHTML & "</form>" & vbNewLine
			'wHTML = wHTML & "</table><br>" & vbNewLine
		end if

	Loop
	wGotoCompareFl = false
else
	wGotoCompareFl = true
end if

End Function

'========================================================================
'
'	Function	�J�e�S���[�C���[�J�[���C���i�����Ƀ\�[�g
'
'========================================================================
'
Function SortProduct()

Dim i
Dim RSv

wSQL = ""
For i=1 to wRecCnt
	if i > 1 then
		wSQL = wSQL & " UNION "
	end if
	wSQL = wSQL & "SELECT '" & CategoryCd(i) & "' AS CategoryCd"
	wSQL = wSQL & "     , '" & MakerName(i) & "' AS MakerName"
	wSQL = wSQL & "     , '" & ProductName(i) & "' AS ProductName"
	wSQL = wSQL & "     , '" & MakerCd(i) & "' AS MakerCd"
	wSQL = wSQL & "     , '" & ProductCd(i) & "' AS ProductCd"
	wSQL = wSQL & "     , '" & Iro(i) & "' AS Iro"
	wSQL = wSQL & "     , '" & Kikaku(i) & "' AS Kikaku"
Next
wSQL = wSQL & " ORDER BY 1, 2, 3"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

i = 1
Do until RSv.EOF = true
	CategoryCd(i) = RSv("CategoryCd")
	MakerCd(i) = RSv("MakerCd")
	ProductCd(i) = RSv("ProductCd")
	Iro(i) = RSv("Iro")
	Kikaku(i) = RSv("Kikaku")
	MakerName(i) = RSv("MakerName")
	ProductName(i) = RSv("ProductName")
	RSv.MoveNext
	i = i + 1
Loop

RSv.Close

End function

'========================================================================
'
'	Function	��r���i�f�[�^���o��
'
'========================================================================
'
Function getCompareProduct()

Dim i
Dim RSv

for i=1 to wRecCnt
	'---- ���iRecordset�쐬
	wSQL = ""
' 2012/01/20 GV Mod Start
'	wSQL = wSQL & "SELECT b.���[�J�[��"
'	wSQL = wSQL & "     , a.���i��"
'	wSQL = wSQL & "  FROM Web���i a WITH (NOLOCK)"
'	wSQL = wSQL & "     , ���[�J�[ b WITH (NOLOCK)"
'	wSQL = wSQL & " WHERE b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'	wSQL = wSQL & "   AND a.���[�J�[�R�[�h = '" & MakerCd(i) & "'"
'	wSQL = wSQL & "   AND a.���i�R�[�h = '" & ProductCd(i) & "'"
	wSQL = wSQL & "SELECT "
	wSQL = wSQL & "      b.���[�J�[�� "
	wSQL = wSQL & "    , a.���i�� "
	wSQL = wSQL & "FROM "
	wSQL = wSQL & "    Web���i               a WITH (NOLOCK) "
	wSQL = wSQL & "      INNER JOIN ���[�J�[ b WITH (NOLOCK) "
	wSQL = wSQL & "        ON     b.���[�J�[�R�[�h = a.���[�J�[�R�[�h "
	wSQL = wSQL & "WHERE "
	wSQL = wSQL & "        a.���[�J�[�R�[�h = '" & MakerCd(i) & "' "
	wSQL = wSQL & "    AND a.���i�R�[�h     = '" & Replace(ProductCd(i), "'", "''") & "' "
' 2012/01/20 GV Mod End

	Set RSv = Server.CreateObject("ADODB.Recordset")
	RSv.Open wSQL, Connection, adOpenStatic

	MakerName(i) = RSv("���[�J�[��")
	ProductName(i) = RSv("���i��")

	RSv.MoveNext
Next

RSv.Close

End function

'========================================================================
'
'	Function	�^�C�g���Z�b�g
'
'========================================================================
'
Function SetTitle(i)

Dim RSv

'---- �^�C�g���Z�b�g
wSQL = ""
' 2012/01/20 GV Mod Start
'wSQL = wSQL & "SELECT a.��J�e�S���[�R�[�h"
'wSQL = wSQL & "     , a.��J�e�S���[��"
'wSQL = wSQL & "     , b.���J�e�S���[�R�[�h"
'wSQL = wSQL & "     , b.���J�e�S���[�����{��"
'wSQL = wSQL & "     , c.�J�e�S���[�R�[�h"
'wSQL = wSQL & "     , c.�J�e�S���[��"
'wSQL = wSQL & "     , c.�����߃J�e�S���[�t���O"
'wSQL = wSQL & "  FROM ��J�e�S���[ a"
'wSQL = wSQL & "     , ���J�e�S���[ b"
'wSQL = wSQL & "     , �J�e�S���[ c"
'wSQL = wSQL & " WHERE b.��J�e�S���[�R�[�h = a.��J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND c.���J�e�S���[�R�[�h = b.���J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND c.�J�e�S���[�R�[�h = '" & CategoryCd(i) & "'"
wSQL = wSQL & "SELECT "
wSQL = wSQL & "       a.��J�e�S���[�R�[�h "
wSQL = wSQL & "     , a.��J�e�S���[�� "
wSQL = wSQL & "     , b.���J�e�S���[�R�[�h "
wSQL = wSQL & "     , b.���J�e�S���[�����{�� "
wSQL = wSQL & "     , c.�J�e�S���[�R�[�h "
wSQL = wSQL & "     , c.�J�e�S���[�� "
wSQL = wSQL & "     , c.�����߃J�e�S���[�t���O "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    ��J�e�S���[              a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN ���J�e�S���[ b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.��J�e�S���[�R�[�h = a.��J�e�S���[�R�[�h "
wSQL = wSQL & "      INNER JOIN �J�e�S���[   c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.���J�e�S���[�R�[�h = b.���J�e�S���[�R�[�h "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "         c.�J�e�S���[�R�[�h = '" & CategoryCd(i) & "' "
' 2012/01/20 GV Mod End

'@@@@@@@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

wTitleWithLink = ""
'wTitleWithLink = wTitleWithLink & "<h2 class='title'>" & RSv("��J�e�S���[��") & " / " & RSv("���J�e�S���[�����{��") & " / " & RSv("�J�e�S���[��") & "</h2>" & vbNewLine
wTitleWithLink = wTitleWithLink & "<h2 class='title'>" & RSv("�J�e�S���[��") & "</h2>" & vbNewLine
RSv.close

End Function

'========================================================================
'
'	Function	�i�r�Q�[�V�����Z�b�g
'
'========================================================================
'
Function SetNavi(i)

Dim RSv

'---- �i�r�Q�[�V�����Z�b�g
wSQL = ""
' 2012/01/20 GV Mod Start
'wSQL = wSQL & "SELECT a.��J�e�S���[�R�[�h"
'wSQL = wSQL & "     , a.��J�e�S���[��"
'wSQL = wSQL & "     , b.���J�e�S���[�R�[�h"
'wSQL = wSQL & "     , b.���J�e�S���[�����{��"
'wSQL = wSQL & "     , c.�J�e�S���[�R�[�h"
'wSQL = wSQL & "     , c.�J�e�S���[��"
'wSQL = wSQL & "     , c.�����߃J�e�S���[�t���O"
'wSQL = wSQL & "  FROM ��J�e�S���[ a"
'wSQL = wSQL & "     , ���J�e�S���[ b"
'wSQL = wSQL & "     , �J�e�S���[ c"
'wSQL = wSQL & " WHERE b.��J�e�S���[�R�[�h = a.��J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND c.���J�e�S���[�R�[�h = b.���J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND c.�J�e�S���[�R�[�h = '" & CategoryCd(i) & "'"
wSQL = wSQL & "SELECT "
wSQL = wSQL & "       a.��J�e�S���[�R�[�h "
wSQL = wSQL & "     , a.��J�e�S���[�� "
wSQL = wSQL & "     , b.���J�e�S���[�R�[�h "
wSQL = wSQL & "     , b.���J�e�S���[�����{�� "
wSQL = wSQL & "     , c.�J�e�S���[�R�[�h "
wSQL = wSQL & "     , c.�J�e�S���[�� "
wSQL = wSQL & "     , c.�����߃J�e�S���[�t���O "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    ��J�e�S���[              a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN ���J�e�S���[ b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.��J�e�S���[�R�[�h = a.��J�e�S���[�R�[�h "
wSQL = wSQL & "      INNER JOIN �J�e�S���[   c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.���J�e�S���[�R�[�h = b.���J�e�S���[�R�[�h "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "         c.�J�e�S���[�R�[�h = '" & CategoryCd(i) & "' "
' 2012/01/20 GV Mod End

'@@@@@@@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

wNaveWithLink = ""
wNaveWithLink = wNaveWithLink & "<div id='path_box'><div id='path_box_inner01'><div id='path_box_inner02'>" & vbNewLine
wNaveWithLink = wNaveWithLink & " <p class='home'><a href='../'><img src='../images/icon_home.gif' alt='HOME'></a></p>" & vbNewLine
wNaveWithLink = wNaveWithLink & " <ul id='path'>" & vbNewLine
'wNaveWithLink = wNaveWithLink & "  <li><a href='LargeCategoryList.asp?LargeCategoryCd=" & RSv("��J�e�S���[�R�[�h") & "'>" & RSv("��J�e�S���[��") & "</a></li>" & vbNewLine
'wNaveWithLink = wNaveWithLink & "  <li><a href='MidCategoryList.asp?MidCategoryCd=" & RSv("���J�e�S���[�R�[�h") & "'>" & RSv("���J�e�S���[�����{��") & "</a></li>" & vbNewLine
'wNaveWithLink = wNaveWithLink & "  <li><a href='SearchList.asp?i_type=c&s_category_cd=" & RSv("�J�e�S���[�R�[�h") & "'>" &  RSv("�J�e�S���[��") & "</a></li>" & vbNewLine
wNaveWithLink = wNaveWithLink & "  <li class='now'>���i��r</li>" & vbNewLine
wNaveWithLink = wNaveWithLink & "  </ul>" & vbNewLine
wNaveWithLink = wNaveWithLink & "</div></div></div>" & vbNewLine
wNaveWithLink = wNaveWithLink & "<h1 class='title'>���i��r</h1>" & vbNewLine
wNaveWithLink = wNaveWithLink & "<p class='error'>5�ȏ�܂��͕����̃J�e�S���[�̏��i���I������܂����B<br>��r���鏤�i��I�����Ă��������B</p>" & vbNewLine
RSv.close

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
<title>���i��r�b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="Style/productcompare.css" type="text/css">

<script type="text/javascript">
//
//	Hikaku onSubmit
//
function Hikaku_onSubmit(pForm){

	var vParm = "";
	var vCnt = 0;

// Item count check
	for (var i=0; i<pForm.iItem.length; i++){
		if (pForm.iItem[i].checked == true){
			vCnt += 1;
			vParm += pForm.iItem[i].value;
		}
	}
	if (vCnt > 5){
		alert("5�ȏ�̏��i���I������܂����B5�ȓ��őI�����Ă��������B");
		return false;
	}
	if (vCnt < 2){
		alert("2�ȏ�̏��i��I�����Ă��������B");
		return false;
	}
	window.location = "ProductCompare.asp?item=" + vParm;
	return false;
}

</script>

</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->
<div id="globalMain">
	<span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>
	<!-- �R���e���cstart -->
	<div id="globalContents">
		<%=wHTML%>
	</div>
	<div id="globalSide">
		<!--#include file="../Navi/NaviSide.inc"-->
	</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>