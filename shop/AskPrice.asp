<%@ LANGUAGE="VBScript" %>
<%
'�l�b�g�n�E�X�˂��ƃn�E�X�l�b�g�͂���
 Option Explicit
%>
<!--#include file="../common/ADOVBS.inc"-->
<!--#include file="../common/system_common.inc"-->
<!--#include file="../common/shop_common_functions.inc"-->
<!--#include file="../common/bfunctions1.asp"-->

<%
'========================================================================
'
'	ASK���i�\���y�[�W
'
'�X�V����
'
'2009/04/09 �����̐F�𔒂ɕύX
'
'========================================================================

Dim MakerName
Dim ProductName
Dim Price

'========================================================================

'---- Get input data
MakerName = ReplaceInput(Trim(Request("MakerName")))
ProductName = ReplaceInput(Trim(Request("ProductName")))
Price = FormatNumber(ReplaceInput(Trim(Request("Price"))),0) & "�~(�ō�)"

%>

<html>
<head>
<title>ASK���i</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">

<!--#include file="../Navi/NaviStyle.inc"-->
</head>

<body>

<div class="honbun" align="center">
<span style="color:black;">
<%=MakerName%><br>
<%=ProductName%></span><br><br>
<span style="color:red; font-weight:bold;"><%=Price%></span>
</div>

</body>
</html>
