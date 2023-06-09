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
<%
'========================================================================

'	���[���y�[�W
'
'2012/06/18 ok �f�U�C���ύX�̂��ߋ��ł����ɐV�K�쐬
'
'========================================================================
On Error Resume Next
Response.Expires = -1			' Do not cache

'---- Session���
Dim wUserID
Dim wUserName
Dim wMsg

Dim wLoanDownPaymentFl
Dim wLoanDownPaymentAm
Dim wLoanTermPayment
Dim wLoanTerm
Dim wLoanAm
Dim wLoanApplyFl
Dim wLoanCompany
Dim wErrDesc

'---- DB
Dim Connection

'=======================================================================
'	�󂯓n�������o��
'=======================================================================
'---- Session�ϐ�
wUserID = Session("UserID")
wUserName = Session("userName")
wMsg = Session("msg")

'---- �󂯓n�������o��

Session("msg") = ""

'---- �Z�b�V�����؂�`�F�b�N
If wUserID = ""Then
	Response.Redirect g_HTTP
End If

'=======================================================================
'	Execute main
'=======================================================================
Call connect_db()
Call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "OrderLoan.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

Call close_db()

If Err.Description <> "" Then
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

'========================================================================
'
'	Function	Connect database
'
'========================================================================
Function connect_db()

'---- Connect database
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
Function close_db()

Connection.Close
Set Connection= Nothing    '2011/08/01 an add

End Function

'========================================================================
'
'	Function	Main
'
'========================================================================
Function main()

' ���[�����̎擾
Call GetLoanInfo()

End Function

'========================================================================
'
'	Function	���[�����擾
'
'========================================================================
Function GetLoanInfo

Dim RSv
Dim vSQL

vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    ���[����������t���O"
vSQL = vSQL & "  , ���[������"
vSQL = vSQL & "  , ��]���[����"
vSQL = vSQL & "  , ���[�����z"
vSQL = vSQL & "  , �I�����C�����[���\���t���O"
vSQL = vSQL & "  , ���[�����"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    ����"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    SessionID = '" & gSessionID & "'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

wLoanDownPaymentFl  = RSv("���[����������t���O")
wLoanDownPaymentAm = RSv("���[������")
wLoanTerm = RSv("��]���[����")
wLoanAm = RSv("���[�����z")
If wLoanAm <> 0 Then
	wLoanTermPayment = "P"
Else
	wLoanTermPayment = "T"
End If
wLoanApplyFl = RSv("�I�����C�����[���\���t���O")
wLoanCompany = RSv("���[�����")

RSv.Close

End Function

'========================================================================

%>
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<title>���[���̂��\�����݁b�T�E���h�n�E�X</title>
<link rel="stylesheet" href="style/StyleOrder.css?20120629a" type="text/css">
<!--#include file="../Navi/NaviStyle.inc"-->

<script type="text/javascript">
//=====================================================================
//	���փ{�^�� onClick
//=====================================================================
function Next_onClick() {
	document.f_data.action = "OrderLoanStore.asp";
	document.f_data.submit();
}
//=====================================================================
//	�L�����Z���{�^�� onClick
//=====================================================================
function Cancel_onClick() {
	document.f_data.action = "OrderInfoEnter.asp";
	document.f_data.submit();
}
//=====================================================================
//	���W�I�{�^���A�h���b�v�_�E�����X�g�̑I��
//=====================================================================
function preset_values(){

	// ��������^�Ȃ�
	if (document.f_data.i_loan_downpayment_fl.value == "Y"){
		document.f_data.loan_downpayment_fl[1].checked = true;
	}
	if (document.f_data.i_loan_downpayment_fl.value == "N"){
		document.f_data.loan_downpayment_fl[0].checked = true;
	}

	// �I�����C���Ő\�����ށ^�g�p���Ȃ�
	if (document.f_data.i_loan_apply_fl.value == "Y"){
		document.f_data.loan_apply_fl[0].checked = true;
		// ���[�����
		if (document.f_data.i_loan_company.value == "�W���b�N�X"){
			document.f_data.loan_company[0].checked = true;
		}
		if (document.f_data.i_loan_company.value == "�Z�f�B�i"){
			document.f_data.loan_company[1].checked = true;
		}
	}
	if (document.f_data.i_loan_apply_fl.value == "N") {
		document.f_data.loan_apply_fl[1].checked = true;
		// ��]���[���񐔁^���z�x�����z
		if (document.f_data.i_loan_term_payment.value == "T"){
			document.f_data.loan_term_payment[0].checked = true;
			for (var i=0; i<document.f_data.loan_term.length; i++){
				if (document.f_data.loan_term[i].value == document.f_data.i_loan_term.value){
					document.f_data.loan_term.options[i].selected = true;
					break;
				}
			}
		}
		if (document.f_data.i_loan_term_payment.value == "P"){
			document.f_data.loan_term_payment[1].checked = true;
		}
	}

}
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
      <li>���͂���A���x�������@�̑I��</li>
      <li class="now">���[���̂��\������</li>
    </ul>
  </div></div></div>

  <h1 class="title">���͂���A���x�������@�̑I��</h1>
  <ol id="step">
    <li><img src="images/step01.gif" alt="1.�V���b�s���O�J�[�g" width="170" height="50"></li>
    <li><img src="images/step02_now.gif" alt="2.���͂���A���x�����@�̑I��" width="170" height="50"></li>
    <li><img src="images/step03.gif" alt="3.���������e�̊m�F" width="170" height="50"></li>
    <li><img src="images/step04.gif" alt="4.����������" width="170" height="50"></li>
  </ol>

  <p class="error"><% = wMsg %></p>

  <h2 class="cart_title">���[���̂��\������</h2>        

  <form name="f_data" method="post">
    <table id="address">
      <tr>
        <td class="main">
          <ul class="loan_choice">
            <li><input id="loan_downpayment_fl_n" name="loan_downpayment_fl" type="radio" value="N"><label for="loan_downpayment_fl_n">�����Ȃ�</label></li>
            <li><input id="loan_downpayment_fl_y" name="loan_downpayment_fl" type="radio" value="Y"><label for="loan_downpayment_fl_y">��������</label></li>
          </ul>
          <ul>
            <li>����<input name="loan_downpayment_am" type="text" value="<% = wLoanDownPaymentAm %>" size="12" class="field_r">�~</li>
          </ul>
        </td>
      </tr>
      <tr>
        <td class="main">
          <span class="loan_choice"><input id="loan_apply_fl_y" name="loan_apply_fl" type="radio" value="Y"><label for="loan_apply_fl_y" class="radio_strong">�I�����C���Ń��[����\������</label></span>
          <ul>
            <li><input id="loan_company_1" name="loan_company" type="radio" value="�W���b�N�X"><label for="loan_company_1"><img src="images/jaccs.gif" alt="�W���b�N�X"></label><label for="loan_company_1">�W���b�N�X</label></li>
            <li><input id="loan_company_2" name="loan_company" type="radio" value="�Z�f�B�i"><label for="loan_company_2"><img src="images/cedyna.gif" alt="�Z�f�B�i"></label><label for="loan_company_2">�Z�f�B�i</label></li>
          </ul>
          <ul class="attention">
            <li>�I�����C�����[���̏ꍇ����\����̂��������e�̕ύX�����邱�Ƃ��ł��܂���B</li>
            <li>���������e�Ƥ�I�����C�����[���\���t�H�[���̓��e�����m�F�̏�A���������������B</li>
            <li>�W���b�N�X�ł��\�����݂̏ꍇ�́A�����Ȃ��ƂȂ�܂��B</li>
          </ul>
        </td>
      </tr>
      <tr>
        <td class="main">
          <span class="loan_choice"><input id="loan_apply_fl_n" name="loan_apply_fl" type="radio" value="N"><label for="loan_apply_fl_n" class="radio_strong">�I�����C�����g�p���Ȃ�</label>�i���[���񐔂܂��͌��z���w�肵�Ă��������j</span>
          <ul>
            <li><input id="loan_term_payment_t" name="loan_term_payment" type="radio" value="T"><label for="loan_term_payment_t">��]���[����</label>
              <select name="loan_term" size="1">
                <option value="0"></option>
                <option value="1">1</option>
                <option value="2">2</option>
                <option value="3">3</option>
                <option value="6">6</option>
                <option value="10">10</option>
                <option value="12">12</option>
                <option value="15">15</option>
                <option value="18">18</option>
                <option value="20">20</option>
                <option value="24">24</option>
                <option value="30">30</option>
                <option value="36">36</option>
                <option value="42">42</option>
                <option value="48">48</option>
                <option value="54">54</option>
                <option value="60">60</option>
              </select>
            </li>
            <li><input id="loan_term_payment_p" name="loan_term_payment" type="radio" value="P"><label for="loan_term_payment_p">���z�x�����z</label><input name="loan_am" type="text" value="<% = wLoanAm %>" size="12" class="field_r">�~</li>
          </ul>
          <ul class="attention">
            <li>���[����Ђɂ�育��]�̂��x�����񐔂��w��ł��Ȃ��ꍇ���������܂��B</li>
          </ul>
        </td>
      </tr>
    </table>
    <div id="btn_box">
      <ul class="btn">
        <li><a href="javascript:Cancel_onClick();"><img src="images/btn_back.png" alt="�߂�" class="opover"></a></li>
        <li class="last"><a href="javascript:Next_onClick();"><img src="images/btn_next.png" alt="����" class="opover"></a></li>
      </ul>
    </div>
    <input type="hidden" name="i_loan_downpayment_fl" value="<% = wLoanDownPaymentFl %>">
    <input type="hidden" name="i_loan_apply_fl" value="<% = wLoanApplyFl %>">
    <input type="hidden" name="i_loan_company" value="<% = wLoanCompany %>">
    <input type="hidden" name="i_loan_term_payment" value="<% = wLoanTermPayment %>">
    <input type="hidden" name="i_loan_term" value="<% = wLoanTerm %>">
  </form>

<!--/#contents --></div>
	<div id="globalSide">
	<!--#include file="../Navi/NaviSide.inc"-->
	<!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript">
	preset_values();
</script>
</body>
</html>