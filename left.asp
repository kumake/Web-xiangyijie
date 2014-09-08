<table width="449" height="858" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td rowspan="7" width="168" height="858" ></td>
		<td>
			<img src="images/left_02.jpg" width="281" height="57" alt=""></td>
	</tr>
	<tr>
		<td background="images/left_03.jpg" width="281" height="109" >
			<div style="line-height:20px; margin-left:10px">
				<div>公司名称：南京香依界服饰贸易有限公司</div>
				<div>联 系 人：陆经理</div>
				<div>联系方式：13584092678</div>
				<div>电话：4000-5858-36</div>
				<div>传真：025-58788092</div>
				<div>地址:南京市白下区中山南路华美大厦03栋705室</div>
			</div></td>
	</tr>
	<tr>
		<td>
			<img src="images/left_04.jpg" width="281" height="135" alt=""></td>
	</tr>
	<tr>
		<td background="images/left_05.jpg" width="281" height="221">
			<div style="margin-top:2px">
			                <%
		  	set rscate = UICon.Query("select * from Sp_dictionary where  categoryID = 16 order by  IndexID")
			if rscate.recordcount<>0 then
			do while not rscate.eof
		  %>
                <a style="color:#FF0000" href="item.asp?categoryID=<%=rscate("id")%>" class="fl"><%=rscate("categoryname")%></a>
              
            <%
		  	rscate.movenext
			loop
			end if
		  %> 
			</div></td>
	</tr>
	<tr>
		<td>
			<img src="images/left_06.jpg" width="281" height="64" alt=""></td>
	</tr>
	<tr>
		<td background="images/left_07.jpg" width="281" height="215" >
			<div style="line-height:26px; margin-left:16px; margin-top:12px; margin-right:10px">
										<%
							set rsn = UICon.Query("select top 7 PostTime,title,categoryID,id from user_news order by id desc")
							if rsn.recordcount<>0 then
							for i=1 to 7
							%> 
								<li style="background-image:url(images/pp.jpg); background-position:left; background-repeat:no-repeat"><span style="float:right">[&nbsp;<%=year(rsn("PostTime"))%>-<%=month(rsn("PostTime"))%>-<%=day(rsn("PostTime"))%>&nbsp;]</span>&nbsp;&nbsp;<a  href="news_in.asp?categoryid=<%=rsn("categoryID")%>&amp;itemid=<%=rsn("id")%>" title="<%=rsn("title")%>" id="title" target="_blank" >
								<% if len(rsn("title"))>10 then %>
								<%=left(rsn("title"),10)%>...
								<% Else %>
								<%=rsn("title")%>
								<% End If %>
								</a>
								</li>
							<%
							rsn.movenext
							if rsn.eof then exit for
							next
							end if
							rsn.close
							set rsn=nothing
							%>
			</div></td>
	</tr>
	<tr>
		<td>
			<img src="images/left_08.jpg" width="281" height="57" alt=""></td>
	</tr>
</table>