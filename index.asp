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
							<p>�Ͼ����������ó�����޹�˾���Ͼ�һ��רҵ��������ҵ��������ְҵװ���������������Ʒ����Ƶ��Ʒ��������Ʒ����������ҡ���˾�Գ������������ϵĿ����������г��������˷ḻ�ľ��顣��˾���С�רҵ�����¡����񡱵�����ʹ��Ʒ����������ȫ�����ز�һ���ܵ��ͻ��������������ѳ�Ϊ�����Ƚ��г���������������飬���ߴ��������ͽ�ǿ�ۺ�ʵ����ְҵװ��ҵ��</p>
							<p>����ʼ�ռ�֡��Ծ����ܵ͵ļ۸��ṩ���Ժϸ�֮��Ʒ����������޶�����˿����󡱵ľ�Ӫ�����֡��ͻ���һ����ԭ��Ϊ���ͻ��ṩ���ʵķ����ᳫ�Ͻ���ʵ�Ĺ������磬΢����Ӫ��ע��ʵ����Ϊȷ���ҳ��Ĳ�Ʒ������ͬ��ҵ������ǰé�춨�˼�ʵ�Ļ�������������ǰ���ۺ����ʹ�ҳ���ҵ�����������ٶȷ�չ���ѱ����ָ��Ϊ��־��װ��������������ʱ��ӭ���صġ����������������ġ����̡�˰�񡢼�졢˾�������������ල���Ļ�����硢��������ȫ��������վ������������ԺУ����ҵ��λ�����ˡ�����������Ǣ̸ҵ��</p>
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
var speed=16; //����Խ���ٶ�Խ��
var tab=document.getElementById("demo");
var tab1=document.getElementById("demo1");
var tab2=document.getElementById("demo2");
tab2.innerHTML=tab1.innerHTML; //��¡demo1Ϊdemo2
function Marquee(){
if(tab2.offsetTop-tab.scrollTop<=0)//��������demo1��demo2����ʱ
tab.scrollTop-=tab1.offsetHeight //demo�������
else{
tab.scrollTop++
}
}
var MyMar=setInterval(Marquee,speed);
tab.onmouseover=function() {clearInterval(MyMar)};//�������ʱ�����ʱ���ﵽ����ֹͣ��Ŀ��
tab.onmouseout=function() {MyMar=setInterval(Marquee,speed)};//����ƿ�ʱ���趨ʱ��
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
