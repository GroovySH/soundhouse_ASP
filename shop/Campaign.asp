<%@ LANGUAGE="VBScript" %>

<%
Option Explicit
%>
<!--#include file="../common/ADOVBS.inc"-->
<!--#include file="../common/system_common.inc"-->
<!--#include file="../common/shop_common_functions.inc"-->
<!--#include file="../common/bfunctions1.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<title>キャンペーン一覧｜サウンドハウス</title>
<meta name="description" content="サウンドハウスで開催中のオトクなキャンペーン一覧です。">
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/news.css" type="text/css">
<link rel="stylesheet" href="style/campaign.css?20131121" type="text/css">
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->

<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>

<!-- コンテンツstart -->
<div id="globalContents">
  <div id='path_box'><div id='path_box_inner01'><div id='path_box_inner02'>
    <p class='home'><a href='../'><img src='../images/icon_home.gif' alt='HOME'></a></p>
    <ul id='path'>
      <li class="now">キャンペーン一覧</li>
    </ul>
  </div></div></div>

  <img src="images/campaign/campaign_banner.jpg" alt="キャンペーン一覧">

  <ul id="campaign">
    <li class="loading"><img src="../images/ajax-loader.gif" alt="loading..."></li>
  </ul>
  <ul class="pager"></ul>
<!--
  <p id="test"></p>
-->
</div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
<script src="jslib/campaign.js?20131224"></script>
</body>
</html>