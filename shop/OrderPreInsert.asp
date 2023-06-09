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
'	オーダー仮登録
'
'------------------------------------------------------------------------
'	
'		仮受注情報の基本項目の登録
'		仮受注明細情報の登録
'
'		カートへ入れるボタンで呼び出される，カートへデーターセット後order.aspへ
'
'------------------------------------------------------------------------

'	更新履歴
'2005/01/13 商品が有効かどうかのチェック（キャッシュ画面からの登録対応)
'2005/02/16 個数限定数量単価取り出し時の条件強化　個数限定数量＞0を追加
'2006/06/26 廃盤商品の場合、引当可能数以上に受注しないように変更
'2007/03/15 パラメータに対してReplaceInputを追加
'2007/04/18 B品追加
'2007/07/05 Itemパラメータ（メーカーコード^商品コード^色^規格)取得対処
'2007/07/05 商品登録は1つのみに変更
'2007/07/13 色規格あり商品は色規格が選択されたか再チェック
'2008/05/23 入力データチェック強化（LEFT, Numeric, EOF他)
'2008/12/16 On Error Resume Next 追加
'2008/12/24 AdditionalProdを追加（複数商品同時登録を可能：一緒に購入機能)
'           B品フラグ=Yの商品のときのみB品単価使用に変更
'2011/04/14 hn SessionID関連変更
'2011/08/01 an #1087 Error.aspログ出力対応
'2011/09/16 an #1112 切り売り商品の場合は合算せずに別明細として登録
'========================================================================

On Error Resume Next		'2008/12/16

Dim userID

Dim qt
Dim maker_cd
Dim product_cd
Dim iro
Dim kikaku

Dim item
Dim item_list()
Dim item_cnt

Dim AdditionalItem()
Dim AdditionalItemCnt

Dim Connection
Dim RS_order_header
Dim RS_order_detail
Dim RS_product
Dim RS

Dim w_sql
Dim w_msg
Dim w_html

Dim w_detail_cnt
Dim wErrDesc   '2011/08/01 an add

'=======================================================================

'---- execute main process

Session.Timeout = 20


Session("msg") = ""
w_msg = ""

call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "OrderPreInsert.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then				'2008/12/16
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
qt = ReplaceInput(Trim(Request("qt")))
maker_cd = Left(ReplaceInput(Trim(Request("maker_cd"))), 8)
product_cd = Left(ReplaceInput(Trim(Request("product_cd"))), 20)
iro = Left(ReplaceInput(Trim(Request("iro"))), 20)
kikaku = Left(ReplaceInput(Trim(Request("kikaku"))), 20)

if isNumeric(qt) = false Or qt = "" then
	qt = 0
end if

'if qt > 10000 then
'	qt = 10000
'end if

item = ReplaceInput(Trim(Request("Item")))

if item <> "" then
	item_cnt = cf_unstring(item, item_list, "^")
	maker_cd = Left(ReplaceInput(Trim(item_list(0))), 8)
	product_cd = Left(ReplaceInput(Trim(item_list(1))), 20)
	if item_cnt > 2 then
		iro = Left(ReplaceInput(Trim(item_list(2))), 20)
		if item_cnt > 3 then
			kikaku = Left(ReplaceInput(Trim(item_list(3))), 20)
		end if
	end if
end if

if ReplaceInput(Trim(Request("AdditionalItem"))) <> "" then
	AdditionalItemCnt = cf_unstring(ReplaceInput(Trim(Request("AdditionalItem"))), AdditionalItem, ",")
end if

'---- 仮受注情報登録
call insert_order_header()

'---- 基本商品登録
call insert_order_detail(maker_cd, product_cd, iro, kikaku, qt)

'---- 追加商品登録
for i=1 to AdditionalItemCnt-1
	if AdditionalItem(i) <> "" then
		item_cnt = cf_unstring(AdditionalItem(i), item_list, "^")
		maker_cd = Left(ReplaceInput(Trim(item_list(0))), 8)
		product_cd = Left(ReplaceInput(Trim(item_list(1))), 20)
		if item_cnt > 2 then
			iro = Left(ReplaceInput(Trim(item_list(2))), 20)
			if item_cnt > 3 then
				kikaku = Left(ReplaceInput(Trim(item_list(3))), 20)
			end if
		end if
	end if

	call insert_order_detail(maker_cd, product_cd, iro, kikaku, qt)
Next

End Function

'========================================================================
'
'	Function	insert order_header
'
'========================================================================
'
Function insert_order_header()
Dim i

'---- 該当SessionIDで仮受注が登録されてるかチェック
w_sql = "SELECT *" _
	  & "  FROM 仮受注" _
		& " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod

Set RS_order_header = Server.CreateObject("ADODB.Recordset")
RS_order_header.CursorType = adOpenDynamic
RS_order_header.LockType = adLockOptimistic
RS_order_header.Open w_sql, Connection

if RS_order_header.EOF = true then

	'---- insert 仮受注
	RS_order_header.AddNew

	RS_order_header("SessionID") = gSessionID		'2011/04/14 hn mod
	RS_order_header("入力日") = now()
	RS_order_header("広告コード") = Session("AdID")
	RS_order_header("最終更新日") = now()

	RS_order_header.update
end if

RS_order_header.close

End function

'========================================================================
'
'	Function	insert 仮受注明細
'
'========================================================================
'
Function insert_order_detail(pMakerCd, pProductCd, pIro, pKikaku, pQt)
Dim i
Dim w_max_detail_no
Dim w_update_cnt
Dim wPrice

'---- MAX受注明細番号取出し
w_sql = "SELECT MAX(受注明細番号) AS MAX受注明細番号" _
	    & "  FROM 仮受注明細" _
	    & " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
	  
Set RS = Server.CreateObject("ADODB.Recordset")
RS.CursorType = adOpenDynamic
RS.LockType = adLockOptimistic
RS.Open w_sql, Connection

if RS.EOF = false then
	if isNULL(RS("MAX受注明細番号")) = false then
		w_max_detail_no = RS("MAX受注明細番号")
	else
		w_max_detail_no = 0
	end if
else
	w_max_detail_no = 0
end if

RS.close

w_update_cnt = 0

if pQt > 0 then
	w_update_cnt = w_update_cnt + 1

	'---- 仮受注明細Recordset作成
	w_sql = ""
	w_sql = w_sql & "SELECT *"
	w_sql = w_sql & "  FROM 仮受注明細"
	w_sql = w_sql & " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
	w_sql = w_sql & "   AND メーカーコード = '" & pMakerCd & "'"
	w_sql = w_sql & "   AND 商品コード = '" & pProductCd & "'"
	w_sql = w_sql & "   AND 色 = '" & pIro & "'"
	w_sql = w_sql & "   AND 規格 = '" & pKikaku & "'"
		  
	Set RS_order_detail = Server.CreateObject("ADODB.Recordset")
	RS_order_detail.CursorType = adOpenDynamic
	RS_order_detail.LockType = adLockOptimistic
	RS_order_detail.Open w_sql, Connection

'@@@@@		response.write(w_sql)

	'---- 商品情報取出し
	w_sql = ""
	w_sql = w_sql & "SELECT a.メーカーコード"
	w_sql = w_sql & "     , a.商品コード"
	w_sql = w_sql & "     , a.商品名"
	w_sql = w_sql & "     , CASE"
	w_sql = w_sql & "         WHEN (a.個数限定数量 > a.個数限定受注済数量 AND a.個数限定数量 > 0) THEN a.個数限定単価"
	w_sql = w_sql & "         ELSE a.販売単価"
	w_sql = w_sql & "       END AS 販売単価"
	w_sql = w_sql & "     , a.B品単価"	
	w_sql = w_sql & "     , a.個数限定数量"	
	w_sql = w_sql & "     , a.個数限定受注済数量"	
	w_sql = w_sql & "     , a.ASK商品フラグ"
	w_sql = w_sql & "     , a.B品フラグ"
	w_sql = w_sql & "     , a.廃番日"
	w_sql = w_sql & "     , a.切り売りフラグ"    '2011/09/16 an add
	w_sql = w_sql & "     , b.メーカー名"
	w_sql = w_sql & "     , c.引当可能数量"
	w_sql = w_sql & "     , c.B品引当可能数量"

	'色規格があるかどうか 2007/07/13
	w_sql = w_sql & "     , (SELECT COUNT(*)"
	w_sql = w_sql & "          FROM Web色規格別在庫 t"
	w_sql = w_sql & "         WHERE t.メーカーコード = a.メーカーコード"
	w_sql = w_sql & "           AND t.商品コード = a.商品コード"
	w_sql = w_sql & "           AND (t.色 != '' OR t.規格 != '')"
	w_sql = w_sql & "           AND t.終了日 IS NULL"
	w_sql = w_sql & "       ) AS 色規格CNT"

	w_sql = w_sql & "  FROM Web商品 a"
	w_sql = w_sql & "     , メーカー b"
	w_sql = w_sql & "     , Web色規格別在庫 c"
	w_sql = w_sql & " WHERE b.メーカーコード = a.メーカーコード"
	w_sql = w_sql & "   AND c.メーカーコード = a.メーカーコード"
	w_sql = w_sql & "   AND c.商品コード = a.商品コード"
	w_sql = w_sql & "   AND c.色 = '" & pIro & "'"
	w_sql = w_sql & "   AND c.規格 = '" & pKikaku & "'"
	w_sql = w_sql & "   AND a.メーカーコード = '" & pMakerCd & "'"
	w_sql = w_sql & "   AND a.商品コード = '" & pProductCd & "'"
	w_sql = w_sql & "   AND a.Web商品フラグ = 'Y'"
		  
	Set RS_product = Server.CreateObject("ADODB.Recordset")
	RS_product.CursorType = adOpenDynamic
	RS_product.LockType = adLockOptimistic
	RS_product.Open w_sql, Connection

	if RS_product.EOF = true then
		w_msg = w_msg & pProductCd & "は、取扱をしておりません。<br>"
	else
		if RS_product("色規格CNT") > 0 AND pIro ="" AND pKikaku = "" then
			w_msg = w_msg & "色規格を選択してください。<br>"
		else
			'---- 未登録商品、切り売りフラグYの場合は仮受注明細を追加
			if RS_order_detail.EOF = true OR RS_product("切り売りフラグ") = "Y" then    '2011/09/16 an mod
				if isNull(RS_product("廃番日")) = false AND RS_product("引当可能数量") < CLng(pQt) then
					w_msg = w_msg & pProductCd & "は、在庫が" & RS_product("引当可能数量") & "個しかありません。　数量を変更してご注文ください。<br>"
				else
					if RS_product("B品フラグ") = "Y" AND RS_product("B品引当可能数量") > 0 AND RS_product("B品引当可能数量") < CLng(pQt) then
						w_msg = w_msg & pProductCd & "は、在庫が" & RS_product("B品引当可能数量") & "個しかありません。　数量を変更してご注文ください。<br>"
					else
						'---- insert 仮受注明細
						w_max_detail_no = w_max_detail_no + 1

						if RS_product("B品フラグ") = "Y" then
							wPrice = RS_product("B品単価")
						else
							wPrice = RS_product("販売単価")
						end if

						RS_order_detail.AddNew
						RS_order_detail("SessionID") = gSessionID		'2011/04/14 hn mod
						RS_order_detail("受注明細番号") = w_max_detail_no
						RS_order_detail("メーカーコード") = pMakerCd
						RS_order_detail("商品コード") = pProductCd
						RS_order_detail("色") = pIro
						RS_order_detail("規格") = pKikaku
						RS_order_detail("メーカー名") = RS_product("メーカー名")
						RS_order_detail("商品名") = RS_product("商品名")
						RS_order_detail("受注単価") = wPrice
						RS_order_detail("受注数量") = Clng(pQt)
						RS_order_detail("受注金額") = Fix(RS_order_detail("受注単価")) * Clng(pQt)

						if RS_product("個数限定数量") > RS_product("個数限定受注済数量") then
							RS_order_detail("個数限定単価フラグ") = "Y"
						else
							RS_order_detail("個数限定単価フラグ") = ""
						end if

						RS_order_detail("B品フラグ") = RS_product("B品フラグ")

						RS_order_detail.Update
					end if
				end if
			'---- 登録済み商品は受注数量を追加する（切り売り商品以外）
			else
				if isNull(RS_product("廃番日")) = false AND RS_product("引当可能数量") < RS_order_detail("受注数量") + Clng(pQt) then
					w_msg = w_msg & pProductCd & "は、在庫が" & RS_product("引当可能数量") & "個しかありません。　数量を変更してご注文ください。<br>"
				else
					if RS_product("B品引当可能数量") > 0 AND RS_product("B品引当可能数量") < CLng(pQt) then
						w_msg = w_msg & pProductCd & "は、在庫が" & RS_product("B品引当可能数量") & "個しかありません。　数量を変更してご注文ください。<br>"
					else
						'---- update 仮受注明細
						RS_order_detail("受注数量") = RS_order_detail("受注数量") + Clng(pQt)
						RS_order_detail("受注金額") = Fix(RS_order_detail("受注単価")) * RS_order_detail("受注数量")
						RS_order_detail.Update
					end if
				end if
			end if
		end if
	end if
end if

if w_update_cnt > 0 then
	RS_order_detail.close
end if

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
