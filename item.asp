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
						<img src="images/main_in_01it.jpg" width="743" height="57" alt=""></td>
					<td rowspan="2" width="174" height="585" ></td>
				</tr>
				<tr>
					<td width="743" height="528">
						<div style="line-height:26px; margin:4px 8px 10px 10px">
						��	<%  
							key=trim(replace(request("key"),"'",""))
 CategoryID = VerificationUrlParam("CategoryID","int","0")%>
                <%
				''''''���ø÷�ҳ���͵ĺô���
				''�����ҳ��ʽ����һ���Խ����ݶ����ڴ�ͨ���α���ʾÿҳ��¼
				''�����ݱ��¼����10������1000��ʱ��
				''����һ֧�ʸ÷�ҳ��ʽ�ĺô�������Ҫ��������¼�ʹӱ��ж���������¼
				''���ݲ������ݼ�¼Խ��Ч��Խ���ԡ����ԣ�asp+sql2000+������100��
				'''''''
				Dim total_record   	'�ܼ�¼��
				Dim current_page	'��ǰҳ��
				Dim PCount			'ѭ��ҳ��ʾ��Ŀ
				Dim pagesize		'ÿҳ��ʾ��¼��
				Dim showpageNum:showpageNum = true		'�Ƿ���ʾ����ѭ��ҳ
				Dim showpagetotal:showpagetotal = true	'�Ƿ���ʾ��¼����
				Dim IsEnglish:IsEnglish = false			'�Ƿ���ʾӢ�ķ�ҳ��ʽ		
				Dim strwhere:strwhere = "isdel=0"		'��ѯ����
				'''��ȡ��ѯ����
				
				
				if categoryID<>0 then strwhere = strwhere &" and categoryID="&categoryID&""
				if key<>"" then strwhere = strwhere &" and title like '%"&key&"%'"
				''''������¼ֻȡһ�ν�ʡ���ݿ�ѹ��
				Dim total:total = VerificationUrlParam("total","int","0")
				if total = 0 then 
					Dim Tarray:Tarray = UICon.QueryData("select count(*) as total from user_item where "&strwhere&"")
					total_record = Tarray(0,0)
				else
					total_record = total
				end if
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6			'''��ҳѭ����ʾ��¼��
				pagesize = 9		'''ÿҳ��ʾ��¼��
				
				'���ַ�ʽΪ����IDΪ�ؼ�������
				'���з�ҳ��ʽЧ����á�����ʹ��,��������Ч���ܵ�����
				'Dim sql:sql = UICon.QueryPageByNum("categoryID,id,title,posttime","user_news", ""&strwhere&"", "ID", true,current_page,pagesize)
				'���ַ�ʽΪ����IndexID����IndexIDֵԽ��Խ��ǰ
				'									 "--------------��Ҫ���õ��ֶ�----------------","-������-","--��ѯ����---", "���ùؼ���","�����ֶ�","����true","��ǰҳ","ÿҳ��¼"
				Dim sql:sql = UICon.QueryPageByNotIn("*","user_item", ""&strwhere&"", "ID", "IndexID desc,ID",false,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rs = UICon.QUery(sql)
				if rs.recordcount<>0 then
				do while not rs.eof
				'''''''''��ô�ֶ���''''''
				''��ҳ�����DIV+css�����Է���ʵ�������ǳ����㲻��Ҫ��ҳ��
				''ֻ��Ҫ�ı�css�� procontent ��ǩ�µ�li�Ŀ�ȼ���
				''һ�� ֻҪprocontent �Ŀ�Ⱥ�li�Ŀ��һ��
				''��Ҫ���� ֻҪ��li�Ŀ������Ϊ procontent�ļ���֮��
			   %>
				<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
				  <% 
				  for i=1 to 3
				   %>
				  <tr>
					<% 
					for k=1 to 3
					%>
					<td align="center">
				
								  <table width="200" height="160" border="0" cellpadding="0" cellspacing="0">
									<tr>
									  <td height="150" align="center" valign="middle">
									  <%if rs("Field_picture")<>"" then%>
<a href="item_in.asp?categoryID=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>" ><img style="border:#666666 solid 2px" src="<%=rs("Field_picture")%>" width="140" height="200"/></a>
										<%else%>
										<a href="item_in.asp?categoryID=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>" class="pro_pic">��������ͼ</a>
										<%end if%>
										</td>
									</tr>
									<tr>
									  <td height="23"   align="center" valign="middle"><span style="line-height:150%"><a href="item_in.asp?categoryID=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>"><%=rs("Title")%></a></span></td>
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
					
					%>
				</table>

			  <div style="line-height:24px; text-align:center; width:100%;"> <%PageIndexUrl total_record,current_page,PCount,pagesize,showpageNum,showpagetotal,IsEnglish%></div>
			  				<% 
				else
				Response.Write("&nbsp;")
				end if
				 %>
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
