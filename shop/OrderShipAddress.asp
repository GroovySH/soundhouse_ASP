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

'	�͐�y�[�W
'�ύX����
'

'========================================================================
On Error Resume Next
Response.Expires = -1			' Do not cache

'---- Session���
Dim wUserID
Dim wUserName
Dim wMsg

'---- �󂯓n����������ϐ�
Dim ship_name
Dim ship_zip
Dim ship_prefecture
Dim ship_address
dim ship_telephone

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
ship_name = Left(ReplaceInput(Trim(Request("ship_name"))), 30)
ship_zip = Left(ReplaceInput(Trim(Request("ship_zip"))), 10)
ship_prefecture = Left(ReplaceInput(Trim(Request("ship_prefecture"))), 4)
ship_address = Left(ReplaceInput(Trim(Request("ship_address"))), 40)
ship_telephone = Left(ReplaceInput(Trim(Request("ship_telephone"))), 20)

Session("msg") = ""

'---- �Z�b�V�����؂�`�F�b�N
If wUserID = ""Then
	Response.Redirect g_HTTP
End If

'========================================================================
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<title>���͂���̓o�^�b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/StyleOrder.css?20120629a" type="text/css">
<script type="text/javascript">
//=====================================================================
//	�Z������ onClick
//=====================================================================
function address_search_onClick(){

	var addrWin;

	if (document.f_data.ship_zip.value == ""){
		alert("�X�֔ԍ�����͂��Ă��������B");
		return;
	}
 
	AddrWin = window.open("../comasp/address_search.asp?zip=" + document.f_data.ship_zip.value + "&name_prefecture=i_ship_prefecture&name_address=ship_address","AddrSearch","width=200,height=100");

}
//=====================================================================
//	���W�I�{�^���A�h���b�v�_�E�����X�g���ȑO�ɑI��������Ԃɂ���
//=====================================================================
function preset_values(){

	// �Z��������������̌Ăяo�����͓s���{���݂̂��Z�b�g
	for (var i=0; i<document.f_data.ship_prefecture.options.length; i++){
		if (document.f_data.ship_prefecture.options[i].value == document.f_data.i_ship_prefecture.value){
			document.f_data.ship_prefecture.options[i].selected = true;
			break;
		}
	}
	return;

}
//=====================================================================
//	���փ{�^�� onClick
//=====================================================================
function Next_onClick() {
	document.f_data.action = "OrderShipAddressStore.asp";
	document.f_data.submit();
}
//=====================================================================
//	�L�����Z���{�^�� onClick
//=====================================================================
function Cancel_onClick() {
	document.f_data.action = "OrderInfoEnter.asp";
	document.f_data.submit();
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
      <li>���͂���A���x�����@�̑I��</li>
      <li class="now">���͂���̓o�^</li>
    </ul>
  </div></div></div>

  <h1 class="title">���͂���A���x�����@�̑I��</h1>
  <ol id="step">
    <li><img src="images/step01.gif" alt="1.�V���b�s���O�J�[�g" width="170" height="50"></li>
    <li><img src="images/step02_now.gif" alt="2.���͂���A���x�����@�̑I��" width="170" height="50"></li>
    <li><img src="images/step03.gif" alt="3.���������e�̊m�F" width="170" height="50"></li>
    <li><img src="images/step04.gif" alt="4.����������" width="170" height="50"></li>
  </ol>

  <p class="error"><% = wMsg %></p>

  <h2 class="cart_title">���͂���̓o�^</h2>

  <form name="f_data" method="post">
    <table id="shipAddress">
      <tr>
        <th>�����O</th>
        <td><input type="text" name="ship_name" id="ship_name" size=30 maxlength=60 value="<% = ship_name %>"></td>
      </tr>
      <tr>
        <th>�Z��</th>
        <td>
          ��<input type="text" name="ship_zip" id="ship_zip" size="10" maxlength="8" value="<% = ship_zip %>"><span>�i���p�j</span>
          <a href="JavaScript:address_search_onClick();" class="tipBtn">�Z������</a><span>�X�֔ԍ�����͂��ă{�^���������Ă��������</span><br>
          <select name="ship_prefecture">
            <option value="">�s���{��</option>
            <option value="�k�C��">�k�C��</option>
            <option value="�X��">�X��</option>
            <option value="�H�c��">�H�c��</option>
            <option value="��茧">��茧</option>
            <option value="�{�錧">�{�錧</option>
            <option value="�R�`��">�R�`��</option>
            <option value="������">������</option>
            <option value="�Ȗ،�">�Ȗ،�</option>
            <option value="�V����">�V����</option>
            <option value="�Q�n��">�Q�n��</option>
            <option value="��ʌ�">��ʌ�</option>
            <option value="��錧">��錧</option>
            <option value="��t��">��t��</option>
            <option value="�����s">�����s</option>
            <option value="�_�ސ쌧">�_�ސ쌧</option>
            <option value="�R����">�R����</option>
            <option value="���쌧">���쌧</option>
            <option value="�򕌌�">�򕌌�</option>
            <option value="�x�R��">�x�R��</option>
            <option value="�ΐ쌧">�ΐ쌧</option>
            <option value="�É���">�É���</option>
            <option value="���m��">���m��</option>
            <option value="�O�d��">�O�d��</option>
            <option value="�ޗǌ�">�ޗǌ�</option>
            <option value="�a�̎R��">�a�̎R��</option>
            <option value="���䌧">���䌧</option>
            <option value="���ꌧ">���ꌧ</option>
            <option value="���s�{">���s�{</option>
            <option value="���{">���{</option>
            <option value="���Ɍ�">���Ɍ�</option>
            <option value="���R��">���R��</option>
            <option value="���挧">���挧</option>
            <option value="������">������</option>
            <option value="�L����">�L����</option>
            <option value="�R����">�R����</option>
            <option value="���쌧">���쌧</option>
            <option value="������">������</option>
            <option value="���Q��">���Q��</option>
            <option value="���m��">���m��</option>
            <option value="������">������</option>
            <option value="���ꌧ">���ꌧ</option>
            <option value="�啪��">�啪��</option>
            <option value="�F�{��">�F�{��</option>
            <option value="�{�茧">�{�茧</option>
            <option value="���茧">���茧</option>
            <option value="��������">��������</option>
            <option value="���ꌧ">���ꌧ</option>
          </select>
          <input type="text" name="ship_address" id="ship_address" size="60" maxlength="80" value="<% = ship_address %>"><br><span>��Ж��A�}���V����/�r�����A�����ԍ��A�����l���A���͖Y�ꂸ���L�����������B</span>
        </td>
      </tr>
      <tr>
        <th>�d�b�ԍ�</th>
        <td><input type="text" name="ship_telephone" id="ship_telephone" size="30" maxlength="20" value="<% = ship_telephone %>"><span>�i���p�����j</span></td>
      </tr>
    </table>

    <div id="btn_box">
      <ul class="btn">
        <li><a href="javascript:Cancel_onClick();"><img src="images/btn_back.png" alt="�߂�" width="151" height="32" class="opover"></a></li>
        <li class="last"><a href="javascript:Next_onClick();"><img src="images/btn_next.png" alt="����" width="151" height="32" class="opover"></a></li>
      </ul>
    </div>

    <input type="hidden" name="i_ship_prefecture" value="<% = ship_prefecture %>">
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