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
#demo {
overflow:hidden;
height: 490px;
text-align: center;
float: left;
}
#demo img {
display: block;
}
-->
</style>
<style type="text/css">
<!--
.proLi{ width:160px; line-height:30px; border-bottom:#CCCCCC solid 1px; display:block; background:url(images/li.jpg) no-repeat 30px 50%; padding-left:50px; margin-left:32px;}
 -->
</style>
</head>
<body>
<div id="container">
<table width="1366" height="1400" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td colspan="2" width="1366" height="437">
			<!--#include file="head.asp" -->
		</td>
	</tr>
	<tr>
		<td width="449" height="858" >
			<!--#include file="left.asp" -->
		</td>
		<td width="917" height="858">
			<table width="917" height="858" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td>
						<img src="images/main_01.jpg" width="753" height="51" alt=""></td>
					<td rowspan="5" width="164" height="858" ></td>
				</tr>
				<tr>
					<td width="753" height="205">
						<div class="aa">
							<span style=" float:right"><img src="images/ppp.jpg" /></span>
							<p>南京香依界服饰贸易有限公司是南京一家专业生产企事业工作服、职业装、体恤衫、工厂制服、酒店制服等团体制服的生产厂家。公司自成立以来，不断的开拓与摸索市场，积累了丰富的经验。公司秉承“专业、创新、服务”的理念使产品畅销华东及全国各地并一致受到客户良好赞誉。现已成为具有先进市场理念，成熟运作经验，不竭创新能力和较强综合实力的职业装企业。</p>
							<p>我们始终坚持“以尽可能低的价格提供绝对合格之产品，并尽最大限度满足顾客需求”的经营理念，坚持“客户第一”的原则为广大客户提供优质的服务。提倡严谨务实的工作作风，微利经营，注重实绩，为确保我厂的产品质量在同行业中名列前茅奠定了坚实的基础。优良的售前，售后服务，使我厂的业绩连年来高速度发展，已被相关指定为标志服装定点生产厂；随时欢迎各地的“政府行政服务中心、工商、税务、监察、司法、质量技术监督、文化、广电、电力、安全生产、车站、保安、大中院校等行业单位，来人、来函、来电洽谈业务。</p>
						</div></td>
				</tr>
				<tr>
					<td>
						<img src="images/main_04.jpg" width="753" height="55" alt=""></td>
				</tr>
				<tr>
					<td background="images/main_05.jpg" width="753" height="507">
						<div>
						<div id="demo">
						<div id="demo1">
							<div style="width:700px; margin-left:32px">
								<table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top:1px; margin-bottom:10px">
								<%
								set rs = UICon.QUery("select top 5 * from user_item where categoryID=132 order by id desc ")
								if rs.recordcount<>0 then
								do while not rs.eof
								for i=1 to 1
								%>
								<tr>
								<% 
								for j=1 to 5
								%>
								<td width="120">
								<table width="120" border="0" align="center" cellpadding="0" cellspacing="0"  height="150" >
								<tr>
								<td width="120"><a href="product_in.asp?categoryID=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>" target="_blank" title="<%=rs("title")%>"><img style="border:#333333 solid 2px" src="<%=rs("Field_picture")%>" width="120" height="150"   border="0" /></a>
								</td>
								</tr>
								</table>
								</td>
								<%
								rs.movenext
								if rs.eof then exit for
								next
								%> 
								</tr>
								<%
								
								if rs.eof then exit for
								next
								loop
								end if
								%> 
								</table>
							</div>
							
							<div style="width:700px; margin-left:32px">
								<table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top:1px; margin-bottom:10px">
								<%
								set rs = UICon.QUery("select top 5 * from user_item where categoryID=131 order by id desc ")
								if rs.recordcount<>0 then
								do while not rs.eof
								for i=1 to 1
								%>
								<tr>
								<% 
								for j=1 to 5
								%>
								<td width="120">
								<table width="120" border="0" align="center" cellpadding="0" cellspacing="0"  height="150" >
								<tr>
								<td width="120"><a href="product_in.asp?categoryID=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>" target="_blank" title="<%=rs("title")%>"><img style="border:#333333 solid 2px" src="<%=rs("Field_picture")%>" width="120" height="150"   border="0" /></a>
								</td>
								</tr>
								</table>
								</td>
								<%
								rs.movenext
								if rs.eof then exit for
								next
								%> 
								</tr>
								<%
								
								if rs.eof then exit for
								next
								loop
								end if
								%> 
								</table>
							</div>
							
							<div style="width:700px; margin-left:32px">
								<table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top:1px; margin-bottom:10px">
								<%
								set rs = UICon.QUery("select top 5 * from user_item where categoryID=130 order by id desc ")
								if rs.recordcount<>0 then
								do while not rs.eof
								for i=1 to 1
								%>
								<tr>
								<% 
								for j=1 to 5
								%>
								<td width="120">
								<table width="120" border="0" align="center" cellpadding="0" cellspacing="0"  height="150" >
								<tr>
								<td width="120"><a href="product_in.asp?categoryID=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>" target="_blank" title="<%=rs("title")%>"><img style="border:#333333 solid 2px" src="<%=rs("Field_picture")%>" width="120" height="150"   border="0" /></a>
								</td>
								</tr>
								</table>
								</td>
								<%
								rs.movenext
								if rs.eof then exit for
								next
								%> 
								</tr>
								<%
								
								if rs.eof then exit for
								next
								loop
								end if
								%> 
								</table>
							</div>
							<div style="width:700px; margin-left:32px">
								<table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top:1px; margin-bottom:10px">
								<%
								set rs = UICon.QUery("select top 5 * from user_item where categoryID=129 order by id desc ")
								if rs.recordcount<>0 then
								do while not rs.eof
								for i=1 to 1
								%>
								<tr>
								<% 
								for j=1 to 5
								%>
								<td width="120">
								<table width="120" border="0" align="center" cellpadding="0" cellspacing="0"  height="150" >
								<tr>
								<td width="120"><a href="product_in.asp?categoryID=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>" target="_blank" title="<%=rs("title")%>"><img style="border:#333333 solid 2px" src="<%=rs("Field_picture")%>" width="120" height="150"   border="0" /></a>
								</td>
								</tr>
								</table>
								</td>
								<%
								rs.movenext
								if rs.eof then exit for
								next
								%> 
								</tr>
								<%
								
								if rs.eof then exit for
								next
								loop
								end if
								%> 
								</table>
							</div>
							
							<div style="width:700px; margin-left:32px">
								<table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top:1px; margin-bottom:10px">
								<%
								set rs = UICon.QUery("select top 5 * from user_item where categoryID=128 order by id desc ")
								if rs.recordcount<>0 then
								do while not rs.eof
								for i=1 to 1
								%>
								<tr>
								<% 
								for j=1 to 5
								%>
								<td width="120">
								<table width="120" border="0" align="center" cellpadding="0" cellspacing="0"  height="150" >
								<tr>
								<td width="120"><a href="product_in.asp?categoryID=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>" target="_blank" title="<%=rs("title")%>"><img style="border:#333333 solid 2px" src="<%=rs("Field_picture")%>" width="120" height="150"   border="0" /></a>
								</td>
								</tr>
								</table>
								</td>
								<%
								rs.movenext
								if rs.eof then exit for
								next
								%> 
								</tr>
								<%
								
								if rs.eof then exit for
								next
								loop
								end if
								%> 
								</table>
							</div>
							
							<div style="width:700px; margin-left:32px">
								<table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top:1px; margin-bottom:10px">
								<%
								set rs = UICon.QUery("select top 5 * from user_item where categoryID=127 order by id desc ")
								if rs.recordcount<>0 then
								do while not rs.eof
								for i=1 to 1
								%>
								<tr>
								<% 
								for j=1 to 5
								%>
								<td width="120">
								<table width="120" border="0" align="center" cellpadding="0" cellspacing="0"  height="150" >
								<tr>
								<td width="120"><a href="product_in.asp?categoryID=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>" target="_blank" title="<%=rs("title")%>"><img style="border:#333333 solid 2px" src="<%=rs("Field_picture")%>" width="120" height="150"   border="0" /></a>
								</td>
								</tr>
								</table>
								</td>
								<%
								rs.movenext
								if rs.eof then exit for
								next
								%> 
								</tr>
								<%
								
								if rs.eof then exit for
								next
								loop
								end if
								%> 
								</table>
							</div>
							
							<div style="width:700px; margin-left:32px">
								<table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top:1px; margin-bottom:10px">
								<%
								set rs = UICon.QUery("select top 5 * from user_item where categoryID=126 order by id desc ")
								if rs.recordcount<>0 then
								do while not rs.eof
								for i=1 to 1
								%>
								<tr>
								<% 
								for j=1 to 5
								%>
								<td width="120">
								<table width="120" border="0" align="center" cellpadding="0" cellspacing="0"  height="150" >
								<tr>
								<td width="120"><a href="product_in.asp?categoryID=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>" target="_blank" title="<%=rs("title")%>"><img style="border:#333333 solid 2px" src="<%=rs("Field_picture")%>" width="120" height="150"   border="0" /></a>
								</td>
								</tr>
								</table>
								</td>
								<%
								rs.movenext
								if rs.eof then exit for
								next
								%> 
								</tr>
								<%
								
								if rs.eof then exit for
								next
								loop
								end if
								%> 
								</table>
							</div>
							
							<div style="width:700px; margin-left:32px">
								<table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top:1px; margin-bottom:10px">
								<%
								set rs = UICon.QUery("select top 5 * from user_item where categoryID=125 order by id desc ")
								if rs.recordcount<>0 then
								do while not rs.eof
								for i=1 to 1
								%>
								<tr>
								<% 
								for j=1 to 5
								%>
								<td width="120">
								<table width="120" border="0" align="center" cellpadding="0" cellspacing="0"  height="150" >
								<tr>
								<td width="120"><a href="product_in.asp?categoryID=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>" target="_blank" title="<%=rs("title")%>"><img style="border:#333333 solid 2px" src="<%=rs("Field_picture")%>" width="120" height="150"   border="0" /></a>
								</td>
								</tr>
								</table>
								</td>
								<%
								rs.movenext
								if rs.eof then exit for
								next
								%> 
								</tr>
								<%
								
								if rs.eof then exit for
								next
								loop
								end if
								%> 
								</table>
							</div>
							
							<div style="width:700px; margin-left:32px">
								<table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top:1px; margin-bottom:10px">
								<%
								set rs = UICon.QUery("select top 5 * from user_item where categoryID=124 order by id desc ")
								if rs.recordcount<>0 then
								do while not rs.eof
								for i=1 to 1
								%>
								<tr>
								<% 
								for j=1 to 5
								%>
								<td width="120">
								<table width="120" border="0" align="center" cellpadding="0" cellspacing="0"  height="150" >
								<tr>
								<td width="120"><a href="product_in.asp?categoryID=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>" target="_blank" title="<%=rs("title")%>"><img style="border:#333333 solid 2px" src="<%=rs("Field_picture")%>" width="120" height="150"   border="0" /></a>
								</td>
								</tr>
								</table>
								</td>
								<%
								rs.movenext
								if rs.eof then exit for
								next
								%> 
								</tr>
								<%
								
								if rs.eof then exit for
								next
								loop
								end if
								%> 
								</table>
							</div>
							
							<div style="width:700px; margin-left:32px">
								<table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top:1px; margin-bottom:10px">
								<%
								set rs = UICon.QUery("select top 5 * from user_item where categoryID=123 order by id desc ")
								if rs.recordcount<>0 then
								do while not rs.eof
								for i=1 to 1
								%>
								<tr>
								<% 
								for j=1 to 5
								%>
								<td width="120">
								<table width="120" border="0" align="center" cellpadding="0" cellspacing="0"  height="150" >
								<tr>
								<td width="120"><a href="product_in.asp?categoryID=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>" target="_blank" title="<%=rs("title")%>"><img style="border:#333333 solid 2px" src="<%=rs("Field_picture")%>" width="120" height="150"   border="0" /></a>
								</td>
								</tr>
								</table>
								</td>
								<%
								rs.movenext
								if rs.eof then exit for
								next
								%> 
								</tr>
								<%
								
								if rs.eof then exit for
								next
								loop
								end if
								%> 
								</table>
							</div>
							</div>
<div id="demo2"></div>
<div id="demo2"></div>
</div>
<script>
<!--
var speed=16; //数字越大速度越慢
var tab=document.getElementById("demo");
var tab1=document.getElementById("demo1");
var tab2=document.getElementById("demo2");
tab2.innerHTML=tab1.innerHTML; //克隆demo1为demo2
function Marquee(){
if(tab2.offsetTop-tab.scrollTop<=0)//当滚动至demo1与demo2交界时
tab.scrollTop-=tab1.offsetHeight //demo跳到最顶端
else{
tab.scrollTop++
}
}
var MyMar=setInterval(Marquee,speed);
tab.onmouseover=function() {clearInterval(MyMar)};//鼠标移上时清除定时器达到滚动停止的目的
tab.onmouseout=function() {MyMar=setInterval(Marquee,speed)};//鼠标移开时重设定时器
-->
</script>
						</div></td>
				</tr>
				<tr>
					<td>
						<img src="images/main_06.jpg" width="753" height="40" alt=""></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td colspan="2"  width="1366" height="105" >
			<!--#include file="footer.asp" -->
		</td>
	</tr>
</table>
</div>
</body>
</html>
