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
'	�S�J�e�S���[�ꗗ�y�[�W
'
'�X�V����
'2005/09/16 1�J�e�S���[�ɕ������J�e�S���[�Ή�
'2006/03/27 Web��J�e�S���[�t���O�Ή�
'2007/06/05 �n�b�J�[�Z�[�t�Ή�
'2009/04/30 �G���[����error.asp�ֈړ�
'2009/07/28 �f�U�C���ύX�i�J�e�S�����ɉ摜��\���j�ALargeCategoryCd=""�̍ۂ͕\�������擪�̃J�e�S����\��
'2009/08/05 �J�e�S���擾��ORDER BY�̏����ɒ��J�e�S���[�R�[�h��ǉ��i�����̒��J�e�S���[�œ����\�������w�肵�Ă���ꍇ�̑Ή��j
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/09/06 ok ���j���[�A���ɔ����V�f�U�C���ɕύX�i��J�e�S���[�͌Œ�Ƃ���j
'
'========================================================================

On Error Resume Next

Dim LargeCategoryCd
Dim LargeCategoryName		'2012/09/06 ok Add

Dim wLargeCategoryListHTML
Dim wCategoryListHTML

Dim Connection
Dim RS

Dim w_sql
Dim w_html
Dim wErrDesc   '2011/08/01 an add

'========================================================================

'---- Get input data
LargeCategoryCd = ReplaceInput(Trim(Request("LargeCategoryCd")))

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "AllCategoryList.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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

if LargeCategoryCd = "" then
	LargeCategoryCd = "1"
end if

'----- HTML�쐬
'call CreateLargeCategoryListHTML()		'��J�e�S���[�ꗗ	'2012/09/06 ok Del
call CreateCategoryListHTML()					'�J�e�S���[�ꗗ

End Function

'========================================================================
'
'	Function	��J�e�S���[�ꗗ
'		'2012/09/06 ok Del
'========================================================================
'
'Function CreateLargeCategoryListHTML()
'
''---- ��J�e�S���[ ���o��
'w_sql = ""
'w_sql = w_sql & "SELECT a.��J�e�S���[��"
'w_sql = w_sql & "     , a.��J�e�S���[�R�[�h"
'w_sql = w_sql & "     , a.��J�e�S���[�摜�t�@�C������"
'w_sql = w_sql & "  FROM ��J�e�S���[ a WITH (NOLOCK)"
'w_sql = w_sql & " WHERE Web��J�e�S���[�t���O = 'Y'"
'w_sql = w_sql & " ORDER BY"
'w_sql = w_sql & "       a.�\����"
'
''@@@@@@@@@@response.write(w_sql)
'
'Set RS = Server.CreateObject("ADODB.Recordset")
'RS.Open w_sql, Connection, adOpenStatic
'
'if RS.EOF = true then 
'	exit function
'end if
'
''----- ��J�e�S���[�ꗗHTML�ҏW
'
'w_html = ""
''2012/09/06 ok Mod Start
''w_html = w_html & "<div class='category_title'><span><h2>�S�J�e�S���[�ꗗ</h2></span></div>" & vbNewLine
''w_html = w_html & "<div id='all_cat_list'>" & vbNewLine
'w_html = w_html & "    <h1 class='title'>�S�J�e�S���[�ꗗ</h1>" & vbNewLine
'w_html = w_html & "    <ul id='allcat'>" & vbNewLine
'
'Do Until RS.EOF = true
'	if LargeCategoryCd = "" then
'		LargeCategoryCd = RS("��J�e�S���[�R�[�h")
'	end if
''	w_html = w_html & "  <div class='Large_cat' style='background-image:url(images/AllCategoryList/" & RS("��J�e�S���[�摜�t�@�C������") & ")'>" & vbNewLine
''	w_html = w_html & "    <a href='AllCategoryList.asp?LargeCategoryCd=" & RS("��J�e�S���[�R�[�h")  & "'>" & vbNewLine
''	w_html = w_html & "      <div class='Large_cat_in'><span><h3>" & RS("��J�e�S���[��") & "</h3></span></div>" & vbNewLine
''	w_html = w_html & "    </a>" & vbNewLine
''	w_html = w_html & "  </div>" & vbNewLine
'	if LargeCategoryCd = RS("��J�e�S���[�R�[�h") Then
'		LargeCategoryName = RS("��J�e�S���[��")
'		w_html = w_html & "      <li class='l"+ RS("��J�e�S���[�R�[�h") +" now'>"+ LargeCategoryName +"</li>" & vbNewLine
'	Else
'		w_html = w_html & "      <li class='l"+ RS("��J�e�S���[�R�[�h") +"'><a href='AllCategoryList.asp?LargeCategoryCd=" + RS("��J�e�S���[�R�[�h") + "'>" + LargeCategoryName + "</a></li>" & vbNewLine
'	End If
'
'	RS.MoveNext
'Loop
''w_html = w_html & "</div>" & vbNewLine
'w_html = w_html & "    </ul>" & vbNewLine
''2012/09/06 ok Mod End
'
'wLargeCategoryListHTML = w_html
'
'RS.Close
'
'End Function

'========================================================================
'
'	Function	�J�e�S���[�ꗗ
'
'========================================================================
'
Function CreateCategoryListHTML()

Dim vMidCategoryCd
'Dim vMidCategoryCount	'2012/09/06 ok Del

'---- �J�e�S���[ ���o��
w_sql = ""
w_sql = w_sql & "SELECT a.���J�e�S���[�R�[�h"
w_sql = w_sql & "     , a.���J�e�S���[�����{��"
w_sql = w_sql & "     , a.���J�e�S���[�摜�t�@�C����"
w_sql = w_sql & "     , b.�J�e�S���[�R�[�h"
w_sql = w_sql & "     , b.�J�e�S���[��"
w_sql = w_sql & "  FROM ���J�e�S���[ a WITH (NOLOCK)"
w_sql = w_sql & "     , �J�e�S���[ b WITH (NOLOCK)"
w_sql = w_sql & "     , �J�e�S���[���J�e�S���[ c WITH (NOLOCK)"
w_sql = w_sql & " WHERE c.���J�e�S���[�R�[�h = a.���J�e�S���[�R�[�h"
w_sql = w_sql & "   AND b.�J�e�S���[�R�[�h = c.�J�e�S���[�R�[�h"
w_sql = w_sql & "   AND b.Web�J�e�S���[�t���O = 'Y'"
w_sql = w_sql & "   AND a.��J�e�S���[�R�[�h = '" & LargeCategoryCd &"'"
w_sql = w_sql & " ORDER BY"
w_sql = w_sql & "       a.�\����"
w_sql = w_sql & "     , a.���J�e�S���[�R�[�h"
w_sql = w_sql & "     , c.���J�e�S���[�敪"
w_sql = w_sql & "     , b.�\����"


'@@@@@@@@@@response.write(w_sql)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic

if RS.EOF = true then
	exit function
end if

'----- �J�e�S���[�ꗗHTML�ҏW

w_html = ""
'2012/09/06 ok Mod Start
'w_html = w_html & "<div class='category_title'><span><h2>�J�e�S���[����I��</h2></span></div>" & vbNewLine
'w_html = w_html & "<div id='Large_cat_list'>" & vbNewLine
w_html = w_html & "    <h2 class='allcat_title'>" + LargeCategoryName + "</h2>" & vbNewLine
w_html = w_html & "    <ul class='cat_detail'>" & vbNewLine

'vMidCategoryCount = 0

Do Until RS.EOF = true
	vMidCategoryCd = RS("���J�e�S���[�R�[�h")
'	vMidCategoryCount = vMidCategoryCount + 1
	
'	if vMidCategoryCount Mod 3 = 1 then '���J�e�S��3���Ƃ�1�s�Ƃ��ăX�^�C���ݒ�
'		w_html = w_html & "<div class='line'>" & vbNewLine
'	end if
	
'	w_html = w_html & "  <div class='Mid_cat_list'>" & vbNewLine
'	w_html = w_html & "    <div class='border'>" & vbNewLine
'	w_html = w_html & "      <div class='cat_img'><a href='MidCategoryList.asp?MidCategoryCd=" & RS("���J�e�S���[�R�[�h") &  "'><img src='cat_img/" & RS("���J�e�S���[�摜�t�@�C����") & "' border='0' alt='" & RS("���J�e�S���[�����{��") & "'></a></div>" & vbNewLine
'	w_html = w_html & "      <div class='cat_list'>" & vbNewLine
'	w_html = w_html & "        <h4><a href='MidCategoryList.asp?MidCategoryCd=" & RS("���J�e�S���[�R�[�h") &  "'>" & RS("���J�e�S���[�����{��") & "</a></h4>" & vbNewLine
'	w_html = w_html & "        <ul class='list'>" & vbNewLine
	
	w_html = w_html & "      <li>" & vbNewLine
	w_html = w_html & "        <h3 class='allcat_subtitle'><a href='MidCategoryList.asp?MidCategoryCd=" + vMidCategoryCd + "'>" + RS("���J�e�S���[�����{��") + "</a></h3>" & vbNewLine
	w_html = w_html & "        <div class='cat_inner m" + vMidCategoryCd + "'>" & vbNewLine
	w_html = w_html & "          <ul class='s_cat_list'>" & vbNewLine

	Do While RS("���J�e�S���[�R�[�h") =  vMidCategoryCd
'		w_html = w_html & "          <li><a href='SearchList.asp?i_type=c&amp;s_category_cd=" & RS("�J�e�S���[�R�[�h") &  "'>- " & RS("�J�e�S���[��") & "</a></li>" & vbNewLine
		w_html = w_html & "            <li><a href='SearchList.asp?i_type=c&amp;s_category_cd=" + RS("�J�e�S���[�R�[�h") + "'>" & RS("�J�e�S���[��") & "</a></li>" & vbNewLine
		RS.MoveNext
		if RS.EOF = true then Exit Do
	Loop
	w_html = w_html & "          </ul>" & vbNewLine
	w_html = w_html & "        </div>" & vbNewLine
	w_html = w_html & "      </li>" & vbNewLine

'	w_html = w_html & "        </ul>" & vbNewLine
'	w_html = w_html & "      </div>" & vbNewLine
'	w_html = w_html & "    </div>" & vbNewLine
'	w_html = w_html & "  </div>" & vbNewLine
	
'	if vMidCategoryCount Mod 3 = 0 then '3�̔{���Ȃ�<div class='line'>�����
'		w_html = w_html & "</div>" & vbNewLine
'	end if
Loop

'if vMidCategoryCount Mod 3 <> 0 then '3�̔{���łȂ��ꍇ���Ō�̃f�[�^�ł����<div class='line'>�����
'	w_html = w_html & "</div>" & vbNewLine
'end if

'w_html = w_html & "</div>" & vbNewLine
w_html = w_html & "    </ul>" & vbNewLine
'2012/09/06 ok Mod End

wCategoryListHTML = w_html

RS.Close

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
<title>�S�J�e�S���[�ꗗ�b�T�E���h�n�E�X</title>
<meta name="Description" content="�y��APA�����@��ADJ�EDTM�A�Ɩ��@��A�J���I�P�@�ނ��ǂ������y���������z�ł��񋟂���T�E���h�n�E�X�̑S�J�e�S���[�ꗗ�ł��B">
<meta name="keywords" content="PA���R�[�f�B���O,�M�^�[,�x�[�X,�h����,�p�[�J�b�V����,�L�[�{�[�h,DJ,DTM,���R�[�_�[,�X�^���h,���b�N,�P�[�X,�P�[�u��,�w�b�h�z��,�C���z��">
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/categorylist.css?2014812" type="text/css">
</head>
<body>
<!--#include file="../Navi/NaviTop.inc"-->
<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>
  
  <!-- �R���e���cstart -->
  <div id="globalContents">
    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="<%=g_HTTP%>"><img src="../images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li class="now">�S�J�e�S���[�ꗗ</li>
      </ul>
    </div></div></div>

    <h1 class="title">�S�J�e�S���[�ꗗ</h1>
    <ul id="allcat">
      <li class="l1"><a href="AllCategoryList.asp?LargeCategoryCd=1">PA&amp;���R�[�f�B���O</a></li>
      <li class="l12"><a href="AllCategoryList.asp?LargeCategoryCd=12">�M�^�[</a></li>
      <li class="l13"><a href="AllCategoryList.asp?LargeCategoryCd=13">�x�[�X</a></li>
      <li class="l14"><a href="AllCategoryList.asp?LargeCategoryCd=14">�h���� &amp;<br>�p�[�J�b�V����</a></li>
      <li class="l15"><a href="AllCategoryList.asp?LargeCategoryCd=15">�L�[�{�[�h</a></li>
      <li class="l16"><a href="AllCategoryList.asp?LargeCategoryCd=16">���̑� �y��</a></li>
      <li class="l7"><a href="AllCategoryList.asp?LargeCategoryCd=7">DJ &amp; VJ</a></li>
      <li class="l8"><a href="AllCategoryList.asp?LargeCategoryCd=8">DTM / DAW</a></li>
      <li class="l3"><a href="AllCategoryList.asp?LargeCategoryCd=3">�f���@��E<br>���R�[�_�[</a></li>
      <li class="l4"><a href="AllCategoryList.asp?LargeCategoryCd=4">�Ɩ��E<br>�X�e�[�W�V�X�e��</a></li>
      <li class="l9"><a href="AllCategoryList.asp?LargeCategoryCd=9">�X�^���h�e��</a></li>
      <li class="l5"><a href="AllCategoryList.asp?LargeCategoryCd=5">���b�N�E�P�[�X</a></li>
      <li class="l10"><a href="AllCategoryList.asp?LargeCategoryCd=10">�P�[�u���e��</a></li>
      <li class="l6"><a href="AllCategoryList.asp?LargeCategoryCd=6">�w�b�h�z���E<br>�C���z��</a></li>
      <li class="l51"><a href="AllCategoryList.asp?LargeCategoryCd=51">�X�^�W�I�Ƌ�</a></li>
    </ul>

<%=wCategoryListHTML%>
  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript">
$(function(){
	$(".l<%=LargeCategoryCd%>").addClass("now");
	$(".l<%=LargeCategoryCd%> a").replaceWith("<p>" + $(".l<%=LargeCategoryCd%> a").html() + "</p>");

	$(".allcat_title").text($(".now p").text());

	$(".cat_inner").equalbox();
});</script>
</body>
</html>