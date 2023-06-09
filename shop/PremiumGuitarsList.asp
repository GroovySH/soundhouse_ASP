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
'	PremiumGuitar�ꗗ�y�[�W
'
'�X�V����
'2009/09/01 an �V�K�쐬
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/01/20 an SELECT����LAC�N�G���[�Ă�K�p
'2014/03/19 GV ����ő��łɔ���2�d�\���Ή�
'
'========================================================================

On Error Resume Next

'���[�U���������ɑI�������f�[�^
Dim MakerCd
Dim iSort
Dim iPrice
Dim iPage

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim wMakerName
Dim wMinimumPrice
Dim wMidCategoryCd
Dim wSalesTaxRate
Dim wPriceFrom
Dim wPriceTo
Dim wPageCount
Dim wRecordCount

Dim wMakerHTML
Dim wListHTML
Dim wCountHTML

Dim Connection
Dim RS

Dim wSQL
Dim wErrDesc   '2011/08/01 an add

Const cPageSize = 15 '1�y�[�W������\�����i����

'========================================================================

Response.buffer = true

'---- Get input data
MakerCd = ReplaceInput(Trim(Request("MakerCd")))
iSort = ReplaceInput(Trim(Request("iSort")))
iPrice = ReplaceInput(Trim(Request("iPrice")))
iPage = ReplaceInput(Trim(Request("iPage")))

'---- �\���y�[�W�ݒ�
if iPage = "" then
	iPage = 1
else
	iPage = Clng(iPage)
end if

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "PremiumGuitarsList.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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

Dim vArrayPrice

'---- �ΏۃJ�e�S���[�R�[�h�A�Œ�P����o��
call getCntlMst("���i","PuremiumGuitar","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
wMinimumPrice = Clng(wItemNum1)
wMidCategoryCd = wItemChar1

'---- ����ŗ���o��
call getCntlMst("����","����ŗ�","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
wSalesTaxRate = Clng(wItemNum1)

if iPrice <> "" then
	vArrayPrice = split(iPrice,"-")
	wPriceFrom = vArrayPrice(0)
	wPriceTo   = vArrayPrice(1)
end if

if (ISNumeric(wPriceFrom) = false or wPriceFrom = "") then
	wPriceFrom = wMinimumPrice
end if
if (ISNumeric(wPriceTo) = false or wPriceTo = "") then
	wPriceTo = 9999999
end if

'----- ��������HTML�쐬
call CreateSearchHTML()

'----- PremiumGuitar���X�gHTML�쐬
call CreateListHTML()

End Function

'========================================================================
'
'	Function	���[�J�[���X�g HTML�쐬
'
'========================================================================
'
Function CreateSearchHTML()

'---- �ΏۂƂȂ郁�[�J�[���o��
wSQL = ""
wSQL = wSQL & "SELECT DISTINCT a.���[�J�[��"
wSQL = wSQL & "              , a.���[�J�[�R�[�h"

'wSQL = wSQL & "  FROM ���[�J�[ a WITH (NOLOCK)"     '2012/01/20 an mod s
'wSQL = wSQL & "     , Web���i  b WITH (NOLOCK)"
'wSQL = wSQL & "     , �J�e�S���[���J�e�S���[  c WITH (NOLOCK)"
'wSQL = wSQL & " WHERE  a.���[�J�[�R�[�h = b.���[�J�[�R�[�h"
'wSQL = wSQL & " AND b.�̔��P�� >" &  wMinimumPrice
'wSQL = wSQL & " AND b.�J�e�S���[�R�[�h = c.�J�e�S���[�R�[�h"
'wSQL = wSQL & " AND c.���J�e�S���[�R�[�h IN (" & wMidCategoryCd & ")"
'wSQL = wSQL & " AND b.Web���i�t���O = 'Y'"

wSQL = wSQL & " FROM"
wSQL = wSQL & "     ���[�J�[                             a WITH (NOLOCK)"
wSQL = wSQL & "       INNER JOIN Web���i                 b WITH (NOLOCK)"
wSQL = wSQL & "         ON     b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
wSQL = wSQL & "       INNER JOIN �J�e�S���[���J�e�S���[  c WITH (NOLOCK)"
wSQL = wSQL & "         ON     c.�J�e�S���[�R�[�h = b.�J�e�S���[�R�[�h"
wSQL = wSQL & "       LEFT  JOIN ( SELECT 'Y' AS 'ShohinWebY' ) t1 "
wSQL = wSQL & "         ON     b.Web���i�t���O      = t1.ShohinWebY "
wSQL = wSQL & " WHERE"
wSQL = wSQL & "          t1.ShohinWebY   IS NOT NULL "
wSQL = wSQL & "      AND b.�̔��P�� >" &  wMinimumPrice
wSQL = wSQL & "      AND c.���J�e�S���[�R�[�h IN (" & wMidCategoryCd & ")"     '2012/01/20 an mod e
wSQL = wSQL & " ORDER BY a.���[�J�[��"

'@@@@@@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

if RS.EOF = true then
	exit function
end if

wMakerHTML = ""

'���[�J�[���X�g
wMakerHTML = wMakerHTML & "	<div class='left'>���[�J�[: " & vbNewLine
wMakerHTML = wMakerHTML & "	  <select name='MakerCd'>" & vbNewLine
wMakerHTML = wMakerHTML & "	    <option value='' "

if MakerCd = "" then
	wMakerHTML = wMakerHTML & " selected"
end if

wMakerHTML = wMakerHTML & ">ALL</option> " & vbNewLine

Do Until RS.EOF = true
	wMakerHTML = wMakerHTML & "	    <option value='" & RS("���[�J�[�R�[�h") & "'"
	wMakerHTML = wMakerHTML & ">" & RS("���[�J�[��") & "</option> " & vbNewLine
	If MakerCd = RS("���[�J�[�R�[�h") Then
		wMakerName = RS("���[�J�[��")
	End If
	RS.MoveNext
Loop

wMakerHTML = wMakerHTML & "	  </select>" & vbNewLine
wMakerHTML = wMakerHTML & "	</div>" & vbNewLine


RS.Close

End function

'========================================================================
'
'	Function	PremiumGuitar���X�gHTML�쐬
'
'========================================================================
'
Function CreateListHTML()

Dim vPrice

'---- �Y�����i ���o��
wSQL = ""
wSQL = wSQL & "SELECT DISTINCT"
wSQL = wSQL & "   a.���[�J�[��"
wSQL = wSQL & " , a.���[�J�[�R�[�h"
wSQL = wSQL & " , b.���i��"
wSQL = wSQL & " , b.���i�R�[�h"
wSQL = wSQL & " , b.���i�摜�t�@�C����_��"
wSQL = wSQL & " , b.����o�^��"
wSQL = wSQL & " , CASE"
wSQL = wSQL & "   	WHEN b.�����萔�� > b.������󒍍ϐ��� THEN b.������P��"
wSQL = wSQL & "    	ELSE b.�̔��P��"
wSQL = wSQL & "   END AS ���̔��P��"

'wSQL = wSQL & " FROM ���[�J�[ a WITH (NOLOCK)"       '2012/01/20 an mod s
'wSQL = wSQL & "    , Web���i  b WITH (NOLOCK)"
'wSQL = wSQL & "    , �J�e�S���[���J�e�S���[  c WITH (NOLOCK)"
'wSQL = wSQL & " WHERE  a.���[�J�[�R�[�h = b.���[�J�[�R�[�h"
'wSQL = wSQL & "    AND (SELECT CASE"
'wSQL = wSQL & "                   WHEN x.�����萔�� > x.������󒍍ϐ��� THEN (x.������P�� * (100 + " & wSalesTaxRate & " )/100)"
'wSQL = wSQL & "                   ELSE (x.�̔��P�� * (100 + " & wSalesTaxRate & " )/100)"
'wSQL = wSQL & "                END"
'wSQL = wSQL & "         FROM web���i x "
'wSQL = wSQL & "         WHERE x.���[�J�[�R�[�h = b.���[�J�[�R�[�h"
'wSQL = wSQL & "            AND x.���i�R�[�h = b.���i�R�[�h) BETWEEN " & wPriceFrom & " AND " & wPriceTo
'wSQL = wSQL & "    AND b.�J�e�S���[�R�[�h = c.�J�e�S���[�R�[�h"
'wSQL = wSQL & "    AND c.���J�e�S���[�R�[�h IN (" & wMidCategoryCd & ")"
'wSQL = wSQL & "    AND b.Web���i�t���O = 'Y'"

wSQL = wSQL & " FROM"
wSQL = wSQL & "     ���[�J�[                             a WITH (NOLOCK)"
wSQL = wSQL & "       INNER JOIN Web���i                 b WITH (NOLOCK)"
wSQL = wSQL & "         ON     b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
wSQL = wSQL & "       INNER JOIN �J�e�S���[���J�e�S���[  c WITH (NOLOCK)"
wSQL = wSQL & "         ON     c.�J�e�S���[�R�[�h = b.�J�e�S���[�R�[�h"
wSQL = wSQL & "       LEFT  JOIN ( SELECT 'Y' AS 'ShohinWebY' ) t1 "
wSQL = wSQL & "         ON     b.Web���i�t���O      = t1.ShohinWebY "
wSQL = wSQL & " WHERE"
wSQL = wSQL & "        t1.ShohinWebY   IS NOT NULL "
wSQL = wSQL & "    AND (SELECT CASE"
wSQL = wSQL & "                   WHEN x.�����萔�� > x.������󒍍ϐ��� THEN (x.������P�� * (100 + " & wSalesTaxRate & " )/100)"
wSQL = wSQL & "                   ELSE (x.�̔��P�� * (100 + " & wSalesTaxRate & " )/100)"
wSQL = wSQL & "                END"
wSQL = wSQL & "         FROM web���i x WITH (NOLOCK)"
wSQL = wSQL & "         WHERE x.���[�J�[�R�[�h = b.���[�J�[�R�[�h"
wSQL = wSQL & "            AND x.���i�R�[�h = b.���i�R�[�h) BETWEEN " & wPriceFrom & " AND " & wPriceTo
wSQL = wSQL & "    AND c.���J�e�S���[�R�[�h IN (" & wMidCategoryCd & ")"      '2012/01/20 an mod e


if MakerCd <> "" then
	wSQL = wSQL & " AND b.���[�J�[�R�[�h =" & MakerCd
end if

if iSort = "Update_Desc" or iSort = "" then
	wSQL = wSQL & " ORDER BY b.����o�^�� DESC"
elseif iSort = "Price_Desc" then
	wSQL = wSQL & " ORDER BY ���̔��P�� DESC"
	wSQL = wSQL & "      , b.����o�^�� DESC"
elseif iSort = "Price_Asc" then
	wSQL = wSQL & " ORDER BY ���̔��P��"
		wSQL = wSQL & "    , b.����o�^�� DESC"
elseif iSort = "MakerName" then
	wSQL = wSQL & " ORDER BY a.���[�J�[��"
	wSQL = wSQL & "        , b.����o�^�� DESC"
else
	wSQL = wSQL & " ORDER BY b.����o�^�� DESC"
end if

'@@@@@@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

if RS.EOF = true then
	exit function
end if

RS.PageSize = cPageSize
if iPage > ((RS.RecordCount + (cPageSize - 1)) / cPageSize) then		'MAX�y�[�W�𒴂���ꍇ�͍ŏI�y�[�W��
	iPage = Fix(RS.RecordCount / cPageSize)
end if

RS.AbsolutePage = iPage				' start page no

'----- ���i�ꗗHTML�쐬

wListHTML = ""
wListHTML = wListHTML & "<div id='pgProductBox'>" & vbNewLine

for i=0 to (RS.PageSize - 1)

	wListHTML = wListHTML & "  <div class='productBox'>" & vbNewLine
	wListHTML = wListHTML & "    <div class='top'></div>" & vbNewLine
    '���i�摜�\��
	wListHTML = wListHTML & "    <div class='middle'><a href='PremiumGuitarsDetail.asp?Item=" & Server.URLEncode(RS("���[�J�[�R�[�h") & "^" & RS("���i�R�[�h"))
	
	wListHTML = wListHTML & "'><img src='prod_img/"
	
	if RS("���i�摜�t�@�C����_��") <> "" then
		wListHTML = wListHTML & RS("���i�摜�t�@�C����_��")
	else
		'���i�摜���o�^����Ă��Ȃ��ꍇ�̓u�����N�摜�\��
		wListHTML = wListHTML & "n/nopict.jpg"
	end if
	wListHTML = wListHTML & "'></a></div>" & vbNewLine
	'���i���\��
	wListHTML = wListHTML & "    <div class='middletextbox'><span class='maker'>" & RS("���[�J�[��") & "</span><a href='PremiumGuitarsDetail.asp?Item=" & Server.URLEncode(RS("���[�J�[�R�[�h") & "^" & RS("���i�R�[�h")) & "'>"
	wListHTML = wListHTML & RS("���i��") & "</a></div>" & vbNewLine
	
	vPrice = calcPrice(RS("���̔��P��"), wSalesTaxRate)
	
'2014/03/19 GV mod start ---->
'	wListHTML = wListHTML & "    <div class='Pricetextbox'>���i" & FormatNumber(vPrice,0) & "�~</div>" & vbNewLine
	wListHTML = wListHTML & "    <div class='Pricetextbox'>���i" & FormatNumber(RS("���̔��P��"),0) & "�~(�Ŕ�)</div>" & vbNewLine
	wListHTML = wListHTML & "    <div class='Pricetextbox'>(�ō�&nbsp;" & FormatNumber(vPrice,0) & "�~)</div>" & vbNewLine
'2014/03/19 GV mod end   <----
	wListHTML = wListHTML & "    <div class='bottom'></div>" & vbNewLine
	wListHTML = wListHTML & "  </div>" & vbNewLine
	RS.MoveNext

	if RS.EOF Then
			exit for
	end If

Next

wListHTML = wListHTML & "</div>" & vbNewLine

'----- �����\��HTML�쐬

Dim i

wCountHTML = ""	
wCountHTML = wCountHTML & "<div class='pgPager'>" & vbNewLine
if iPage <> 1 then
	wCountHTML = wCountHTML & "  <a href='JavaScript:Page_onClick(" & iPage-1 & ");'>[�O��]</a>" & vbNewLine
end if
if iPage <> RS.PageCount then
	wCountHTML = wCountHTML & "  <a href='JavaScript:Page_onClick(" & iPage+1 & ");'>[����]</a>�@" & vbNewLine
end if
wCountHTML = wCountHTML & RS.RecordCount & "������܂����B" & vbNewLine

for i=1 to RS.PageCount
	wCountHTML = wCountHTML & "  <a href='JavaScript:Page_onClick(" & i & ");'>" & i & "</a>" & vbNewLine
next

wCountHTML = wCountHTML & "&nbsp;(����" & iPage & "�y�[�W)" & vbNewLine
wCountHTML = wCountHTML & "</div>" & vbNewLine

RS.Close
	
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
<% If MakerCd <> "" Then %>
	<title>�v���~�A���M�^�[ <%= wMakerName %> �ꗗ�b�T�E���h�n�E�X</title>
<% Else %>
	<title>�v���~�A���M�^�[�ꗗ�b�T�E���h�n�E�X</title>
<% End If %>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/PremiumGuitars.css" type="text/css">
<script type="text/javascript">
//
//	�����ݒ�
//
var isMacIE;
var isNS4;
var isIE4;
var isDynamic;
var s_maker_cd;
isMacIE = ((navigator.userAgent.indexOf("IE 4") > -1) && (navigator.userAgent.indexOf("Mac") > -1));
isNS4 = ((navigator.appName == "Netscape") && (parseInt(navigator.appVersion) >= 4));
isIE4 = ((navigator.appName == "Microsoft Internet Explorer") && (parseInt(navigator.appVersion) >= 4));
isDynamic = (isNS4 || isIE4 && !isMacIE);
//
//	Page onClick
//
function Page_onClick(pPage){
	document.f_search.iPage.value = pPage;
	document.f_search.submit();
}
//=====================================================================
//	���W�I�{�^���A�h���b�v�_�E�����X�g���ȑO�ɑI��������Ԃɂ���
//=====================================================================
function preset_values(MakerCd,iSort,iPrice){
//	MakerCD
	for (var i=0; i<document.f_search.MakerCd.options.length; i++){
		if (document.f_search.MakerCd.options[i].value == MakerCd){
			document.f_search.MakerCd.options[i].selected = true;
			break;
		}
	}
//	iSort
	for (var i=0; i<document.f_search.iSort.options.length; i++){
		if (document.f_search.iSort.options[i].value == iSort){
			document.f_search.iSort.options[i].selected = true;
			break;
		}
	}
//	iPrice
	for (var i=0; i<document.f_search.iPrice.options.length; i++){
		if (document.f_search.iPrice.options[i].value == iPrice){
			document.f_search.iPrice.options[i].selected = true;
			break;
		}
	}
}
</script>
<style type="text/css">
#globalContents ul.sns {
	overflow: hidden;
	padding: 5px;
}

#globalContents ul.sns li {
	float: right;
	width: 100px;
	height: 20px;
}
</style>
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->

<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>

<!-- �R���e���cstart -->
<div id="globalContents">
    <div id='path_box'><div id='path_box_inner01'><div id='path_box_inner02'>
    <p class='home'><a href="<%=g_HTTP%>"><img src="<%=g_RelLink%>images/icon_home.gif" alt="HOME"></a></p>
    <ul id='path'>
      <li><a href="<%=g_HTTP%>material/">SPECIAL SELECTION�ꗗ</a></li>
      <li><a href="PremiumGuitars.asp">�v���~�A���M�^�[</a></li>
<% If MakerCd <> "" Then %>
      <li class="now"><%= wMakerName %> �ꗗ</li>
<% Else %>
      <li class="now">�v���~�A���M�^�[�ꗗ</li>
<% End If %>
    </ul>
  </div></div></div>
    <ul class="sns">
          <li><a href="https://twitter.com/share" class="twitter-share-button" data-lang="ja">�c�C�[�g</a><script>!function(d,s,id){var js,fjs=d.getElementsByTagName(s)[0];if(!d.getElementById(id)){js=d.createElement(s);js.id=id;js.src="//platform.twitter.com/widgets.js";fjs.parentNode.insertBefore(js,fjs);}}(document,"script","twitter-wjs");</script></li>
          <li><iframe src="//www.facebook.com/plugins/like.php?href=http%3A%2F%2Fwww.soundhouse.co.jp%2Fshop%2FPremiumGuitars.asp&amp;send=false&amp;layout=button_count&amp;width=100&amp;show_faces=false&amp;action=like&amp;colorscheme=light&amp;font&amp;height=21&amp;appId=191447484218062" scrolling="no" frameborder="0" style="border:none; overflow:hidden; width:100px; height:21px;" allowTransparency="true"></iframe></li>
        </ul>
<!--
  <h1 class="title">�v���~�A���M�^�[</h1>
-->
  <div id="pgContainer">
<!-- �g�b�v�摜 START -->
<div id="pgHeader">
  <div class="topbox">
    <div class="left"></div>
    <div class="right"></div>
  </div>
</div>
<!-- �g�b�v�摜 END -->

<div id="pgSelectBox">
  <form name="f_search">

<%=wMakerHTML%>

    <div class="left">���בւ�:
      <select name="iSort"> 
        <option value="Update_Desc">�X�V��</option>
        <option value="Price_Asc">���i����</option>
        <option value="Price_Desc">���i����</option>
        <option value="MakerName">���[�J�[��</option>
      </select>
    </div>
    <div class="left">�v���C�X:
      <select name="iPrice">
        <option value="">ALL</option>
        <option value="<%=wMinimumPrice%>-250000"><%=wMinimumPrice%>�~ - 250,000�~</option>
        <option value="250001-400000">250,001�~ - 400,000�~</option>
        <option value="400001-600000">400,001�~ - 600,000�~</option>
        <option value="600001-1000000">600,001�~ - 1,000,000�~</option>
        <option value="1000001-9999999">1,000,001�~ -</option>
      </select>
    </div>
    <div class="right">
      <input type="hidden" name="iPage" value="<%=iPage%>">
      <input type="submit" id="bottun" value="����">
    </div>
  </form>
</div>

<!-- �����\�� -->
<%=wCountHTML%>

<!-- PremiumGuitar���X�g -->
<%=wListHTML%>

<!-- �����\�� -->
<%=wCountHTML%>

<p class="arrow"><a href="#site_title"><img src="images/PremiumGuitars/white_arrow_up.gif" alt=""></a></p>

</div>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<script type="text/javascript">
	preset_values('<%=MakerCd%>','<%=iSort%>','<%=iPrice%>');
</script>
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>