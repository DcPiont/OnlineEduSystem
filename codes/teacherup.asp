<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��ʦ��Ϣά��</title>
<link href="css.css" rel="stylesheet" type="text/css" />
</head>

<body>
<p>&nbsp;</p>
<h3 align="center">������Ҫ���Ӽ�¼�������Ϣ</h3>
<br />
<div align="center">
  <form action="teacherinsert.asp" method="post" name="form1" id="form1">
   <table align = "center" class="TableBorder">



<tr>
<td>
<label style="font-size:18px">��ʦ��</label>

</td>
<td>
<input type="text" name="tno" id="TextField" placeholder="��ʦ��" />
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">����</label></td>
<td>
 <input type="text" name="tname" id="TextField" placeholder="����" />
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">�绰</label>

</td>
<td>
<input type="text" name="ttel" id="TextField" placeholder="�绰" />
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">���֤��
</label>

</td>
<td>
<input type="text" name="ID" id="TextField" placeholder="ID" />
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">�Ա�</label>

</td>
<td>
<select name = "sex">
        <option value="��">��
        <option value="Ů">Ů
      
        </select>
</td>
</tr>
<tr>
<td align="center" colspan="2">
<input  type="submit" name="button_submit" id="ButtonStyle2" value="�ύ" />

<input type="reset" name="button_reset" id="ButtonStyle" value="����" />
</td>
</tr>
</table>
  </form>
</div>
<br />
'
<%
''�������ݿ�
' dim sql
' dim course_score
' dim course_period
' sql  ="select * from CourseInfo "
' set conn=server.createobject("adodb.connection")
' conn.open "provider=sqloledb;source=(local);uid=sa;pwd=123;database=website"
' set rs=server.createobject("adodb.recordset")
' rs.open sql,conn,1,3'�������
'course_number=request.Form("course_number")
'teacher_number=request.Form("teacher_number")
'course_name=request.Form("course_name")
'course_type=request.Form("course_type")
'course_score=request.Form("course_score")
'course_period=request.Form("course_period")
'test_type=request.Form("test_type")
'double_degree=request.Form("double_degree")
'rs.addnew
'rs("cno")=trim(course_number)
'rs("tno")=trim(teacher_number)
'rs("cname")=trim(course_name)
'rs("ct")=trim(course_type)
'rs("cscore")=course_score
'rs("cperiod")=course_period
'rs("doubledegree")=trim(double_degree)
'rs.update
'%>
<!--��ʾ���-->
<table width="600" border="1 " cellpadding="0" cellspacing="6" align="center" class="TableBorder">
  <tbody>
    <tr>
      <td valign="top"><table width="100%" border="0" align="center">
        <tr>
          <td class="TableName" align="center">��ʦ������Ϣ�� </td>
        </tr>
      </table>
        <table width="100%" border="0" align="center">
<%
 sql  ="select tno,tname,sex,ttel,tid from teacher "
 set conn=server.createobject("adodb.connection")
 conn.open "provider=sqloledb;data source=localhost\SQLEXPRESS;uid=sa;pwd=11111;database=web"
 set rs=server.createobject("adodb.recordset")	
response.Write("<tr bgcolor=#EAE2F3>")
rs.open sql,conn,1,3
	response.Write("<td align=center>��ʦ��</td>")
	response.Write("<td align=center>����</td>")
	response.Write("<td align=center>�Ա�</td>")
	response.Write("<td align=center>�绰</td>")
	response.Write("<td align=center>���֤��</td>")
	response.Write("<td align=center colspan=2>�޸�</td>")
	response.Write("</tr>")
	do while not rs.eof
		response.Write("<tr bgcolor=#CCCCCC>")
		for i=0 to rs.fields.count-1
			response.Write("<td class=TableFont align=center>"&rs(i)&"</td>")
			
		next
		response.Write("<td align=center><a href=""teacherupdate.asp?tno="&rs(0)&""">�޸�</a></td>")
		response.Write("<td align=center><a href=""teacherdelete.asp?tno="&rs(0)&""">ɾ��</a></td>")
		response.Write("</tr>")
		rs.movenext
	loop
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
        </table></td>
    </tr>
  </tbody>
</table>
<p>&nbsp;</p>
</body>
</html>
