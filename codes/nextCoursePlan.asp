<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��ѧ�ڿγ̰���</title>
<link href="css.css" rel="stylesheet" type="text/css" />
</head>

<body>
<h3 align="center">��ѯ���б��޿���Ϣ</h3>
<div align="center">
<form action="nextCoursePlan.asp" method="post" name="form" id="form">
<input type="submit" name="button_search" id="ButtonStyle" value="��ѯ" />
<input type="submit" name="button_return" id="ButtonStyle" value="����" />
</form>
</div>
<!--�ֽ���-->
<hr color="#9259B0" />
<form action="nextCPinsert.asp" method="post" name="form1" id="form1">
<%
	if (request("button_search")="����") then
	response.Redirect("nextCoursePlan.asp")
	end if
	'��ѯ�γ���Ϣ
	if (request("button_search")="��ѯ") then
	'�������ݿ�
 sql  ="select * from CourseInfo "
 set conn=server.createobject("adodb.connection")
 conn.open "provider=sqloledb;data source=localhost\SQLEXPRESS;uid=sa;pwd=11111;database=web"
 set rs=server.createobject("adodb.recordset")
 rs.open sql,conn,1,3
	response.Write("<table border=1 cellpadding=0 cellspacing=6 align=center class=TableBorder><tbody><tr><td valign=top><table width=100% border=0 align=center><tr><td class=TableName align=center>���޿���Ϣ��</td></tr></table><table width=100% border=0 align=center>")
	response.Write("<tr bgcolor=#EAE2F3>")
	response.Write("<td align=center>�γ̺�</td>")
	response.Write("<td align=center>��ʦ��</td>")
	response.Write("<td align=center>�γ���</td>")
	response.Write("<td align=center>�γ�����</td>")
	response.Write("<td align=center>ѧ��</td>")
	response.Write("<td align=center>ѧʱ</td>")
	response.Write("<td align=center>��������</td>")
	response.Write("<td align=center>�Ƿ�˫ѧλ</td>")
	response.Write("<td align=center colspan=2>�Ƿ����</td>")
	response.Write("</tr>")
	sql="select * from CourseInfo where ct in ('����')"
	set rs=conn.execute(sql)
	do while not rs.eof
		response.Write("<tr bgcolor=#CCCCCC>")
		for i=0 to rs.fields.count-1
			response.Write("<td class=TableFont align=center>"&rs(i)&"</td>")
			
		next
		response.Write("<td align=center><input name=""course_number"" type = ""checkbox"" id = ""checkbox1"" value="&rs(0)&"></td>")
		response.Write("</tr>")
		rs.movenext
	loop
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	else 
	'�������ݿ�
 sql  ="select nextCourseArrange.cno,cname,TermNum,MajorNum,MaxStu,PresentStu,week,day,hour,wtow,place from nextCourseArrange,CourseInfo where nextCourseArrange.cno=CourseInfo.cno "
 set conn=server.createobject("adodb.connection")
 conn.open "provider=sqloledb;data source=localhost\SQLEXPRESS;uid=sa;pwd=11111;database=web"
 set rs=server.createobject("adodb.recordset")
 rs.open sql,conn,1,3
'��ʾ��ѧ�ڿγ̰���
	response.Write("<table border=1 cellpadding=0 cellspacing=6 align=center class=TableBorder")
	response.Write("><tbody><tr><td valign=top><table width=100% border=0 align=center><tr><td class=TableName align=center>��ѧ�ڿγ̰��ű�</td></tr></table><table width=100% border=0 align=center>")
	response.Write("<td align=center>�γ̺�</td>")
	response.Write("<td align=center>�γ���</td>")
	response.Write("<td align=center>ѧ�ں�</td>")
	response.Write("<td align=center>רҵ��</td>")
	response.Write("<td align=center>������</td>")
	response.Write("<td align=center>��ǰ����</td>")
	response.Write("<td align=center>�Ͽ�����</td>")
	response.Write("<td align=center>�Ͽνڴ�</td>")
	response.Write("<td align=center>�Ͽ�ѧʱ</td>")
	response.Write("<td align=center>�Ͽ��ܴ�</td>")
	response.Write("<td align=center>�Ͽεص�</td>")
	response.Write("<td align=center colspan=2>����</td>")
	response.Write("</tr>")
	do while not rs.eof
		response.Write("<tr bgcolor=#CCCCCC>")
		for i=0 to rs.fields.count-1
			response.Write("<td class=TableFont align=center>"&rs(i)&"</td>")
		next
		response.Write("<td align=center><a href=""nextCPupdate.asp?course_number="&rs(0)&"&term_number="&rs(2)&""">�޸�</a></td>")
		response.Write("<td align=center><a href=""nextCPdelete.asp?course_number="&rs(0)&"&term_number="&rs(2)&""">ɾ��</a></td>")
		response.Write("</tr>")
		rs.movenext
	loop
end if
%>
</table>
</tbody>
</table>
<!--�ֽ���-->
<hr color="#9259B0" />
<div align="center">
    <table align ="center" class="TableBorder">
      <tr>
        <td align="center">��ѡ��רҵ����</td>
        <td align="center">
        <select name = "zhuanye">
        <option value="0101">�������</option>
        <option value="0102">����ý��</option>
        <option value="0103">����</option>
        <option value="0104">�����</option>
        </select>
        </td>
        <td>������ѧ�ںţ�
        <input type="text" name="xueqi" />
        </td>
        <td align="center"><input type="submit" name="button_submit" id="ButtonStyle" value="�ύ" /></td>
      </tr>
    </table>
  </form>
</div>
</body>
</html>
