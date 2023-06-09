<%@ LANGUAGE="VBScript" %>
<%
'ネットハウスねっとハウスネットはうす
 Option Explicit
%>
<!--#include file="../common/ADOVBS.inc"-->
<!--#include file="../common/system_common.inc"-->
<!--#include file="../common/shop_common_functions.inc"-->
<!--#include file="../common/bfunctions1.asp"-->

<%
'========================================================================
'
'	ASK価格表示ページ
'
'更新履歴
'
'2009/04/09 文字の色を白に変更
'
'========================================================================

Dim MakerName
Dim ProductName
Dim Price

'========================================================================

'---- Get input data
MakerName = ReplaceInput(Trim(Request("MakerName")))
ProductName = ReplaceInput(Trim(Request("ProductName")))
Price = FormatNumber(ReplaceInput(Trim(Request("Price"))),0) & "円(税込)"

%>

<html>
<head>
<title>ASK価格</title>
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
