<%@ LANGUAGE="VBScript" %>
<%
 Option Explicit
%>
<!--#include file="../common/ADOVBS.inc"-->
<!--#include file="../common/system_common.inc"-->
<!--#include file="../common/shop_common_functions.inc"-->
<!--#include file="../common/bfunctions1.asp"-->

<%
'========================================================================
'
'	保存カートを仮受注へ移動
'
'更新履歴
'2009/04/30 エラー時にerror.aspへ移動
'2011/04/14 hn SessionID関連変更
'2011/08/01 an #1087 Error.aspログ出力対応
'
'========================================================================

On Error Resume Next

Dim userID

Dim CartName

Dim wProdTermFl

Dim Connection
Dim RS

Dim wSQL
Dim wHTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.buffer = true

'---- UserID 取り出し
userID = Session("userID")

'---- 呼び出し元からのデータ取り出し
CartName = ReplaceInput(Request("CartName"))

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "SaveCartMoveToOrder.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

if wMSG = "" then
	Response.Redirect "Order.asp"
else
	Response.Redirect "SaveCartList.asp?msg=" & wMSG
end if

'========================================================================
'
'	Function	Connect database
'
'========================================================================
'
Function connect_db()
Dim i

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

Dim RSv

'----保存カート明細データ取り出し
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM 保存カート明細 WITH (NOLOCK)"
wSQL = wSQL & " WHERE 顧客番号 = " & userID
wSQL = wSQL & "   AND カート名 = '" & CartName & "'"

'@@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

if RS.EOF = true then
	wMSG = "保存されたカート情報がありません。"
	exit function
end if

'---- 仮受注情報登録
call InsertOrderHeader()

'---- 仮受注情報登録
call InsertOrderDetail()

RS.close

End function

'========================================================================
'
'	Function	insert InsertOrderHeader
'
'========================================================================
'
Function InsertOrderHeader()
Dim i
Dim RSv

'---- 該当SessionIDで仮受注が登録されてるかチェック
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM 仮受注"
wSQL = wSQL & " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RSv.EOF = true then

	'---- insert 仮受注
	RSv.AddNew

	RSv("SessionID") = gSessionID		'2011/04/14 hn mod
	RSv("入力日") = now()
	RSv("広告コード") = Session("AdID")
	RSv("最終更新日") = now()

	RSv.update
end if

RSv.close

End function

'========================================================================
'
'	Function	insert 仮受注明細
'
'========================================================================
'
Function InsertOrderDetail()

Dim vQt
Dim RSv
Dim RSvProduct

'---- 仮受注明細があれば全件削除
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM 仮受注明細"
wSQL = wSQL & " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
	  
Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

Do until RSv.EOF = true
	RSv.Delete
	RSv.Requery
Loop

Do Until RS.EOF = true
	'---- 仮受注明細登録
	vQt = RS("受注数量")

	if vQt > 0 then
		'---- 商品情報取出し
		wSQL = ""
		wSQL = wSQL & "SELECT a.メーカーコード"
		wSQL = wSQL & "     , a.商品コード"
		wSQL = wSQL & "     , a.商品名"
		wSQL = wSQL & "     , CASE"
		wSQL = wSQL & "         WHEN (a.個数限定数量 > a.個数限定受注済数量 AND a.個数限定数量 > 0) THEN a.個数限定単価"
		wSQL = wSQL & "         ELSE a.販売単価"
		wSQL = wSQL & "       END AS 販売単価"
		wSQL = wSQL & "     , a.B品単価"	
		wSQL = wSQL & "     , a.個数限定数量"	
		wSQL = wSQL & "     , a.個数限定受注済数量"	
		wSQL = wSQL & "     , a.ASK商品フラグ"
		wSQL = wSQL & "     , a.B品フラグ"
		wSQL = wSQL & "     , a.取扱中止日"
		wSQL = wSQL & "     , a.廃番日"
		wSQL = wSQL & "     , a.完売日"
		wSQL = wSQL & "     , b.メーカー名"
		wSQL = wSQL & "     , c.引当可能数量"
		wSQL = wSQL & "     , c.B品引当可能数量"

		wSQL = wSQL & "  FROM Web商品 a WITH (NOLOCK)"
		wSQL = wSQL & "     , メーカー b WITH (NOLOCK)"
		wSQL = wSQL & "     , Web色規格別在庫 c WITH ( NOLOCK)"

		wSQL = wSQL & " WHERE b.メーカーコード = a.メーカーコード"
		wSQL = wSQL & "   AND c.メーカーコード = a.メーカーコード"
		wSQL = wSQL & "   AND c.商品コード = a.商品コード"
		wSQL = wSQL & "   AND c.色 = '" & RS("色") & "'"
		wSQL = wSQL & "   AND c.規格 = '" & RS("規格") & "'"
		wSQL = wSQL & "   AND a.メーカーコード = '" & RS("メーカーコード") & "'"
		wSQL = wSQL & "   AND a.商品コード = '" & RS("商品コード") & "'"
		wSQL = wSQL & "   AND a.Web商品フラグ = 'Y'"
		wSQL = wSQL & "   AND c.終了日 IS NULL"
			  
		Set RSvProduct = Server.CreateObject("ADODB.Recordset")
		RSvProduct.Open wSQL, Connection, adOpenStatic

		if RSvProduct.EOF = false then

			'---- 終了チェック
			wProdTermFl = "N"
			if isNull(RSvProduct("取扱中止日")) = false then		'取扱中止
				wProdTermFl = "Y"
			end if
			if isNull(RSvProduct("廃番日")) = false AND RSvProduct("引当可能数量") <= 0 then		'廃番で在庫無し
				wProdTermFl = "Y"
			end if
			if isNull(RSvProduct("完売日")) = false then		'完売商品
				wProdTermFl = "Y"
			end if

			if RSvProduct("B品フラグ") <> "Y" then
				if isNull(RSvProduct("廃番日")) = false AND RSvProduct("引当可能数量") < vQt then
					vQt = RSvProduct("引当可能数量")
				end if
			else
				if isNull(RSvProduct("廃番日")) = false AND RSvProduct("B品引当可能数量") < vQt then
					vQt = RSvProduct("B品引当可能数量")
				end if
			end if

			if wProdTermFl <> "Y" AND vQt > 0 then
				'---- insert 仮受注明細
				RSv.AddNew
				RSv("SessionID") = gSessionID		'2011/04/14 hn mod
				RSv("受注明細番号") = RS("受注明細番号")
				RSv("メーカーコード") = RS("メーカーコード")
				RSv("商品コード") = RS("商品コード")
				RSv("色") = RS("色")
				RSv("規格") = RS("規格")
				RSv("メーカー名") = RSvProduct("メーカー名")
				RSv("商品名") = RSvProduct("商品名")

				if RSvProduct("B品フラグ") <> "Y" then
					RSv("受注単価") = RSvProduct("販売単価")
				else
					RSv("受注単価") = RSvProduct("B品単価")
				end if

				RSv("受注数量") = vQt
				RSv("受注金額") = Fix(RSv("受注単価")) * RSv("受注数量")

				if RSvProduct("個数限定数量") > RSvProduct("個数限定受注済数量") then
					RSv("個数限定単価フラグ") = "Y"
				else
					RSv("個数限定単価フラグ") = ""
				end if

				RSv("B品フラグ") = RSvProduct("B品フラグ")

				RSv.Update

			end if
		end if

		RSvProduct.close
	end if

	RS.MoveNext
Loop

RSv.Close

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
