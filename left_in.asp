<table width="449" height="585" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td rowspan="5" width="164" height="585"></td>
		<td>
			<img src="images/left_in_02.jpg" width="285" height="55" alt=""></td>
	</tr>
	<tr>
		<td background="images/left_in_03.jpg" width="285" height="109">
			<div style="line-height:20px; margin-left:14px">
				<div>��˾���ƣ��Ͼ����������ó�����޹�˾</div>
				<div>�� ϵ �ˣ�½����</div>
				<div>��ϵ��ʽ��13584092678</div>
				<div>�绰��4000-5858-36</div>
				<div>���棺025-58788092</div>
				<div>��ַ:�Ͼ��а�������ɽ��·��������03��705��</div>
			</div></td>
	</tr>
	<tr>
		<td>
			<img src="images/left_in_04.jpg" width="285" height="137" alt=""></td>
	</tr>
	<tr>
		<td background="images/left_in_05.jpg" width="285" height="218">
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
			<img src="images/left_in_06.jpg" width="285" height="66" alt=""></td>
	</tr>
</table>