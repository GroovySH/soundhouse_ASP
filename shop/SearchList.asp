<%@ LANGUAGE="VBScript" %>
<%
'�l�b�g�n�E�X�˂��ƃn�E�X�l�b�g�͂���
'�T�E���h�n�E�X
Option Explicit
'========================================================================
'
'	���i�ꗗ�y�[�W(guide.soundhouse.co.jp ��p)
'
'�X�V����
'2016.02.09 GV PHP�łփ��_�C���N�g
'
On Error Resume Next

Dim url
url = "http://www.soundhouse.co.jp/search/index?"
url = url & Request.QueryString

Response.Clear()
Response.Status = "301 Moved Permanently"
Response.AddHeader "Location", url
Response.End()
'Response.Redirect url
%>