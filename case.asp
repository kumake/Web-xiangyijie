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
						<img src="images/main_in_01case.jpg" width="743" height="57" alt=""></td>
					<td rowspan="2" width="174" height="585" ></td>
				</tr>
				<tr>
					<td width="743" height="528">
						<div style="line-height:26px; margin:4px 8px 10px 10px">
				<%  
				CategoryID = VerificationUrlParam("CategoryID","int","0")%>
				<%
				''''''���ø÷�ҳ���͵ĺô���
				''�����ҳ��ʽ����һ���Խ����ݶ����ڴ�ͨ���α���ʾÿҳ��¼
				''�����ݱ���¼����10������1000��ʱ��
				''����һ֧�ʸ÷�ҳ��ʽ�ĺô�������Ҫ��������¼�ʹӱ��ж���������¼
				''���ݲ������ݼ�¼Խ��Ч��Խ���ԡ����ԣ�asp+sql2000+������500��
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
				''''������¼ֻȡһ�ν�ʡ���ݿ�ѹ��
				if CategoryID<>0 then 
				strwhere = strwhere & " and categoryID="&CategoryID&""
				
				end if
				Dim total:total = VerificationUrlParam("total","int","0")
				if total = 0 then 
				Dim Tarray:Tarray = UICon.QueryData("select count(*) as total from user_case where "&strwhere&"")
				total_record = Tarray(0,0)
				else
				total_record = total
				end if
				'''''
				
				current_page = VerificationUrlParam("page","int","1")
				PCount = 6			'''��ҳѭ����ʾ��¼��
				pagesize = 5		'''ÿҳ��ʾ��¼��
				'���ַ�ʽΪ����IDΪ�ؼ�������
				'���з�ҳ��ʽЧ����á�����ʹ��,��������Ч���ܵ�����
				'Dim sql:sql = UICon.QueryPageByNum("categoryID,id,title,posttime","user_case", ""&strwhere&"", "ID", true,current_page,pagesize)
				'���ַ�ʽΪ����IndexID����IndexIDֵԽ��Խ��ǰ
				Dim sql:sql = UICon.QueryPageByNotIn("*","user_case", ""&strwhere&"", "ID", "indexID desc,ID",false,current_page,pagesize)
				'response.Write sql
				'response.End()
				set rsn = UICon.Query(sql)
				if rsn.recordcount<>0 then
				%>
								
							
																 				
				<%
				do while not rsn.eof
				%>
					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-bottom:18px">
					  <tr>
						<td width="21%"><a href="case_in.asp?categoryID=<%=rsn("categoryID")%>&amp;itemid=<%=rsn("id")%>" ><img style="border:#F5E781 solid 3px" src="<%=rsn("Field_picture")%>" width="146" height="62"></a></td>
						<td width="79%"><div style="text-align:center; line-height:60px; font-size:14px"><a href="case_in.asp?categoryID=<%=rsn("categoryID")%>&amp;itemid=<%=rsn("id")%>" ><%=rsn("title")%></a></div></td>
						</tr> </table>	
							<%
				rsn.movenext
				loop
				
				%>	
					 
		
		
		
		
			
				<% end if %>
				 <div style="line-height:30px; text-align:center; width:100%;"> <%PageIndexUrl total_record,current_page,PCount,pagesize,showpageNum,showpagetotal,IsEnglish%></div>
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