<%@ LANGUAGE="VBScript" %>
<%
'ネットハウスねっとハウスネットはうす
'サウンドハウス
 Option Explicit
%>
<!--#include file="../common/ADOVBS.inc"-->
<!--#include file="../common/system_common.inc"-->
<!--#include file="../common/shop_common_functions.inc"-->
<!--#include file="../common/bfunctions1.asp"-->
<%
'========================================================================
'
'	オーダー明細数量変更
'
'更新履歴
'2006/06/26 廃盤商品の場合、引当可能数以上に受注しないように変更
'2006/10/19 数量に空文字を入れられたらエラー対処
'2011/04/14 hn SessionID関連変更
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/08/07 GV #1400 カートページ再計算
'
'========================================================================
'
'		オーダー明細行の数量変更を行う｡
'		数量が0の場合は明細行を削除｡
'		行番号指定のときは該当行のみの更新。(削除ボタン)
'		行番号＝allのときは全行更新。（再計算ボタン)
'
'------------------------------------------------------------------------

'	更新履歴
'2008/05/23 入力データチェック強化（LEFT, Numeric, EOF他)
'2009/04/30 エラー時にerror.aspへ移動
'
'========================================================================

On Error Resume Next

Dim userID

Dim detail_no
Dim qt(100)

Dim Connection
Dim RS

Dim w_sql
Dim w_msg
Dim w_html

Dim w_detail_cnt
Dim wErrDesc   '2011/08/01 an add

'=======================================================================

Session("msg") = ""
w_msg = ""

'---- execute main process
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "OrderChange.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

if w_msg <> "" then
	w_msg = "<font color='#ff0000'>" & w_msg & "</font>"
	Session("msg") = w_msg
end if

Response.Redirect "Order.asp"

'=======================================================================

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
Dim i
Dim v_item

'---- 送信データーの取り出し
detail_no = ReplaceInput(Trim(Request("detail_no")))

for i=1 to 100
	v_item = "qt" & i
	qt(i) = ReplaceInput(Trim(Request(v_item)))

	if qt(i) <> "" then
		if isNumeric(qt(i)) = false then
			w_msg = "<center><font color='#ff0000'>数量に数字を入力してください。</font></center>"
			exit function
		end if
		if qt(i) > 100000 then
			w_msg = "<center><font color='#ff0000'>入力された数量が大きすぎます。</font></center>"
			exit function
		end if
	end if
Next

'---- 仮受注明細情報更新

if detail_no = "all" then
	call update_all()
else
	call delete_one()
end if

End Function

'========================================================================
'
'	Function	Update 仮受注明細 全件
'
'========================================================================
'
Function update_all()

'---- 受注明細情報取出し
w_sql = ""
w_sql = w_sql & "SELECT a.受注明細番号"
w_sql = w_sql & "     , a.商品コード"
w_sql = w_sql & "     , a.受注数量"
w_sql = w_sql & "     , a.受注単価"
w_sql = w_sql & "     , a.受注金額"
w_sql = w_sql & "     , b.廃番日"
w_sql = w_sql & "     , c.引当可能数量"
w_sql = w_sql & "  FROM 仮受注明細 a"
w_sql = w_sql & "     , Web商品 b"
w_sql = w_sql & "     , Web色規格別在庫 c"
w_sql = w_sql & " WHERE b.メーカーコード = a.メーカーコード"
w_sql = w_sql & "   AND b.商品コード = a.商品コード"
w_sql = w_sql & "   AND c.メーカーコード = a.メーカーコード"
w_sql = w_sql & "   AND c.商品コード = a.商品コード"
w_sql = w_sql & "   AND c.色 = a.色"
w_sql = w_sql & "   AND c.規格 = a.規格"
w_sql = w_sql & "   AND a.SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
	  
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic, adLockOptimistic

'---- 受注数量，金額 更新
Do while RS.EOF = false
	if qt(RS("受注明細番号")) <> "" then
		if CLng(qt(RS("受注明細番号"))) > 0 then
			if isNull(RS("廃番日")) = false AND RS("引当可能数量") < CLng(qt(RS("受注明細番号"))) then
				w_msg = w_msg & RS("商品コード") & "は、在庫が" & RS("引当可能数量") & "個しかありません。　数量を変更してご注文ください。<br>"
			else
				RS("受注数量") = CLng(qt(RS("受注明細番号")))
				RS("受注金額") = Fix(RS("受注単価")) * CLng(qt(RS("受注明細番号")))

				RS.Update
			end if
		else
			'2012/08/07 GV add start #1400
 			'RS.Delete
			RS("受注数量") = 0
			RS("受注金額") = 0

			RS.Update
			'2012/08/07 GV add end   #1400 
		end if
	end if

	RS.MoveNext
Loop

Rs.Close

call delete_zero()	'2012/08/07 GV add #1400

End function

'2012/08/07 GV add start #1400
'========================================================================
'
'	Function	Delete 仮受注明細（受注数量 = 0）
'
'========================================================================
'
Function delete_zero()

Dim CMD
Set CMD = Server.CreateObject("ADODB.Command")
CMD.ActiveConnection = Connection

'---- 仮受注明細削除
w_sql = ""
w_sql = w_sql & "DELETE FROM 仮受注明細"
w_sql = w_sql & " WHERE SessionID = '" & gSessionID & "'"
w_sql = w_sql & " AND   受注数量 = 0"

CMD.CommandText = w_sql
CMD.Execute

Set CMD = Nothing

End function
'2012/08/07 GV add end   #1400 

'========================================================================
'
'	Function	Delete 仮受注明細
'
'========================================================================
'
Function delete_one()

if isNumeric(detail_no) = false then
	exit function
end if

'---- 受注明細情報取出し
w_sql = ""
w_sql = w_sql & "SELECT 受注数量"
w_sql = w_sql & "  FROM 仮受注明細"
w_sql = w_sql & " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
w_sql = w_sql & "   AND 受注明細番号 = " & detail_no
	  
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic, adLockOptimistic

'---- 受注数量削除
RS.Delete
Rs.Close

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

%>
