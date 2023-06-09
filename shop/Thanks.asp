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
<!--#include file="../common/HttpsSecurity.inc"-->
<!--#include file="../3rdParty/EAgency.inc"-->

<%
'========================================================================
'
'	�I�[�_�[���肪�Ƃ��������܂����y�[�W
'
'2012/06/18 ok �f�U�C���ύX�̂��ߋ��ł����ɐV�K�쐬
'2012/08/24 ok �C���^�[�X�y�[�X �A�t�B���G�C�g�v���O�����p�^�O��V�łɕύX
'2013/05/20 GV #1505 ���Ԃ݂��ƁI���R�����h�Ή�
'2013/07/30 GV #1618 �A�t�B���G�C�g�d�����M�Ή�
'
'========================================================================

On Error Resume Next

Dim userID
Dim userName
Dim msg

Dim w_order_estimate
Dim payment_method
Dim loan_company
Dim order_no
Dim product_am
Dim w_thanks_msg
Dim w_shiharai_about

'---- UserID ���o��
userID = Session("userID")
userName = Session("userName")

'---- Get input data
msg = Session.contents("msg")
Session("msg") = ""

'---- Execute main
call main()

'========================================================================
'
'	Function	Main
'
'========================================================================
Function main()

'2013/07/30 GV #1618 start
'1�O��OrderSubmit.asp �Ŏd���񂾒l�����݂��Ȃ��ꍇ�A�g�b�v�y�[�W�֑J��
Dim OrderAtOnce
OrderAtOnce = Session("OrderAtOnce")
If ((OrderAtOnce = "") Or (OrderAtOnce <> "1")) Then
	Response.Redirect g_HTTP
Else
	Session.Contents.Remove("OrderAtOnce")
End If
'2013/07/30 GV #1618 end

w_order_estimate = ReplaceInput(Request("order_estimate"))
payment_method = ReplaceInput(Request("payment_method"))
loan_company = ReplaceInput(Trim(Request("loan_company")))
order_no = ReplaceInput(Request("order_no"))
product_am = ReplaceInput(Request("product_am"))

w_thanks_msg = ""
w_thanks_msg = w_thanks_msg & "�����p���肪�Ƃ��������܂��B<br>"
if payment_method = "��s�U��" or payment_method = "�R���r�j�x��" then
	w_thanks_msg = w_thanks_msg & "���˗����������܂������e�̎�t�m�F���[�����������M�������܂����B<br>"
	w_thanks_msg = w_thanks_msg & "���̌�A�ʓr�����ς�����ē��������܂��̂ŁA���e�����m�F���������B<br>"
else
	w_thanks_msg = w_thanks_msg & "���˗����������܂������e�̎�t�m�F���[���𑦎��ɑ��M�������܂��B<br>"
	w_thanks_msg = w_thanks_msg & "���̌�A�ʓr�������m�F���𑗐M�������܂��̂ł��m�F���������B<br><br>"
end if
w_thanks_msg = w_thanks_msg & "��������ЃT�E���h�n�E�X�������p���������܂��B"

w_shiharai_about = ""
if payment_method = "��s�U��" or payment_method = "�R���r�j�x��" then
	w_shiharai_about = w_shiharai_about & "  <dl class='about'>"
	w_shiharai_about = w_shiharai_about & "    <dt>���x�����ɂ���</dt>"
	w_shiharai_about = w_shiharai_about & "    <dd>"

	if payment_method = "��s�U��" then
		w_shiharai_about = w_shiharai_about & "��s�U���������p�̏ꍇ�́A�ʓr���ē��������܂� �u�T�E���h�n�E�X �����Ϗ��v���[�������m�F�̏�A�d�M�����ɂĂ��U���݂��������B<br>"
		w_shiharai_about = w_shiharai_about & "�i���������̏ꍇ�́A�������m�F�܂ł����Ԃ�������܂��B�j<br>"
		w_shiharai_about = w_shiharai_about & "���������m�F��A���i�𔭑��������܂��B"
	else
		w_shiharai_about = w_shiharai_about & "�l�b�g�o���L���O�E�䂤����E�R���r�j�����������p�̏ꍇ�́A�ʓr���ē��������܂� �u�T�E���h�n�E�X �����Ϗ��v���[�������m�F�̏�A���[�����ɋL�ڂ���Ă����pURL�ɃA�N�Z�X���Ă��������B<br>"
		w_shiharai_about = w_shiharai_about & "����]�̂��x�����@��I�����A���x������t�ԍ����͎��[�@�֔ԍ�/�m�F�ԍ��i�䂤�����s�j�����m�F�̏�A�\�L���Ă���܂������܂łɂ��x�������������B <br>"
		w_shiharai_about = w_shiharai_about & "�������m�F��A���i�𔭑��������܂��B"
	end if

	w_shiharai_about = w_shiharai_about & "    <p><a href='http://guide.soundhouse.co.jp/guide/oshiharai.asp'>���x�����ɂ���</a></p>"
	w_shiharai_about = w_shiharai_about & "    </dd>"
	w_shiharai_about = w_shiharai_about & "  </dl>"
end if

End Function

'========================================================================
%>

<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<title>���������肪�Ƃ��������܂����b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/StyleOrder.css" type="text/css">
<script type="text/javascript">
<% if loan_company = "�Z�f�B�i" then %>
	window.open("CFApply.asp?order_no=<%=order_no%>&order_estimate=<%=Server.URLEncode(w_order_estimate)%>")
<% end if %>
<% if loan_company = "�W���b�N�X" then %>
	window.open("JACCSApply.asp?order_no=<%=order_no%>")
<% end if %>
</script>
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->
<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>

<!-- �R���e���cstart -->
<div id="globalContents">

  <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
    <p class="home"><a href="../"><img src="../images/icon_home.gif" alt="HOME"></a></p>
    <ul id="path">
      <li class="now">����������</li>
    </ul>
  </div></div></div>

  <h1 class="title">����������</h1>
  <ol id="step">
    <li><img src="images/step01.gif" alt="1.�V���b�s���O�J�[�g" width="170" height="50"></li>
    <li><img src="images/step02.gif" alt="2.���͂���A���x�����@�̑I��" width="170" height="50"></li>
    <li><img src="images/step03.gif" alt="3.���������e�̊m�F" width="170" height="50"></li>
    <li><img src="images/step04_now.gif" alt="4.����������" width="170" height="50" /></li>
  </ol>

  <div id="thanks">
    <p><strong>THANK YOU!</strong></p>
    <p><%= w_thanks_msg %></p>
    <img src="images/ojigi-2.gif" alt="" height="300" width="150">
  </div>
<%
'2013/05/20 GV #1505
fEAgency_CreateRecommendOrderSubmitJS(order_no)
%>

<%= w_shiharai_about %>

  <dl class="about">
    <dt>���i�̔[���ɂ���</dt>
    <dd>
      <ul>
        <li>�E�F�u�T�C�g��A����т������₨���ς莞�_�ł��ē����Ă���܂��[���ɂ��܂��ẮA�����܂ł��\��ƂȂ��Ă���A������ɂ��ύX�ƂȂ�ꍇ���������܂��B</li>
        <li>���i�̔[���ɂ��܂��ẮA���[���₨�d�b�ł̂��₢���킹�������Ă���܂��B�w����܂łɔ[�i���K�v�Ȃ������́A�����Ȃ����O�ɂ����k���������B</li>
        <li>�Ȃ��A�[���x���ɂ���Đ��������ɂ��܂��ẮA���Ђł͈�؂̐ӂ𕉂����Ƃ��ł��܂���B ���炩���߂��������������B</li>
      </ul>
      <p><a href="http://guide.soundhouse.co.jp/guide/kaimono.asp#nissuu">���i�̔[���ɂ���</a></p>
    </dd>
  </dl>
  <dl class="about">
    <dt>���w����̃T�|�[�g�ɂ���</dt>
    <dd>
      <ul>
        <li>���i�����茳�ɓ͂��܂�����A�����ɔ[�i���Ə��i���e����ѐ��ʂ����m�F���������B</li>
        <li>���ꏤ�i���j�����Ă����蒍���ƈقȂ鏤�i�������ꍇ�́A�����ɂ��A�����������B</li>
        <li>���i������A��T�Ԉȏ�o�߂��Ă��炨�\���o���������ꍇ�A��t�ł��Ȃ����Ƃ��������܂��B</li>
      </ul>
      <p><a href="http://guide.soundhouse.co.jp/guide/support.asp">���w����̃T�|�[�g�ɂ���</a></p>
    </dd>
  </dl>
  
<!--/#contents --></div>
	<div id="globalSide">
	<!--#include file="../Navi/NaviSide.inc"-->
	<!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/Navibottom.inc"-->

<% if w_order_estimate = "������" then %>
<!-- �C���^�[�X�y�[�X �A�t�B���G�C�g�v���O�����p�^�O -->
	<img src="https://is.accesstrade.net/cgi-bin/isatV2/soundhouse/isatWeaselV2.cgi?result_id=2&verify=<%=order_no%>&value=<%=product_am%>" width="1" height="1" />
<% end if %>

<!--#include file="../Navi/NaviScript.inc"-->
</body>
<!-- SmarterJP Conversion Code -->
<!--#include file="../3rdParty/SmarterMerchantOrder.class.asp"-->
<%
Dim oSMO
Dim oRtnCode

set oSMO = new SmarterMerchantOrder

oSMO.MerchantID = "SM1201A10083"		'Merchant ID (SH)
oSMO.Key = "sh10083050206"			'the Key
oSMO.OrderNum = order_no		'Order Number
oSMO.OrderAmount = product_am	'Price Total

oRtnCode = oSMO.send

%>
</html>
