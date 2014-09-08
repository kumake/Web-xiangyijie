<!--#include file="conn.asp"-->
<!--#include file="sp_inc/class_page.asp"-->
<!--#include file="plugIn/Setting.Config.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=config_sitename%></title>
<meta name="keywords" content="<%=config_seokeyword%>">
<meta name="description" content="<%=config_seocontent%>">
<link href="css/public.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.proLi{ width:160px; line-height:30px; border-bottom:#CCCCCC solid 1px; display:block; background:url(images/li.jpg) no-repeat 30px 50%; padding-left:50px; margin-left:32px;}
 -->
</style>
</head>
<body>
<div id="container">
<table width="1366" height="1127" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td colspan="2" width="1366" height="437">
			<!--#include file="head.asp" -->
		</td>
	</tr>
	<tr>
		<td width="449" height="585" >
			<!--#include file="left_in.asp" -->
		</td>
		<td width="917" height="585">
			<table width="917" height="585" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td>
						<img src="images/main_in_01ab.jpg" width="743" height="57" alt=""></td>
					<td rowspan="2" width="174" height="585" ></td>
				</tr>
				<tr>
					<td width="743" height="528">
						<div style="line-height:26px; margin:4px 8px 10px 10px">
						　<% singlename = VerificationUrlParam("signle","string","公司简介")
						response.Write UserSinglePage(singlename)%>
						</div></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td colspan="2"  width="1366" height="105" >
			<!--#include file="footer.asp" -->
		</td>
	</tr>
	<tr>
		<td colspan="2"width="1366"></td>
	</tr>
</table>
</div>
</body>
</html>
