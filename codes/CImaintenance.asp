<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�γ���Ϣά��</title>
<link href="css.css" rel="stylesheet" type="text/css" />
<!--#include file="connCourseInfo.asp"-->
</head>

<body>
<h3 align="center">������Ҫ���Ӽ�¼�������Ϣ</h3>
<br />
<div align="center">
<form name="form1" method="post" action="CIinsert.asp">
<table class="TableBorder">
<tr>
<td>
<input type="text" name="course_number" id="TextField" placeholder="�γ̺�" />
</td>
<td>
<input type="text" name="teacher_number" id="TextField" placeholder="��ʦ��"/>
</td>
<td>
<input type="text" name="course_name" id="TextField" placeholder="�γ���" />
</td>
</tr>
<tr>
<td>
<input type="text" name="course_type" id="TextField" placeholder="�γ�����" />
</td>
<td>
<input type="text" name="course_score" id="TextField" placeholder="ѧ��" />
</td>
<td>
<input type="text" name="course_period" id="TextField" placeholder="ѧʱ" />
</td>
</tr>
<tr>
<td>
<input type="text" name="test_type" id="TextField" placeholder="��������" />
</td>
<td>
<input type="text" name="double_degree" id="TextField" placeholder="�Ƿ�˫ѧλ" />
</td>
<td align="center">
<input type="submit" name="button_submit" id="ButtonStyle" value="�ύ" />
<input type="reset" name="button_reset" id="ButtonStyle" value="����" />
</td>
</tr>
</table>
</form>
</div>
<br />
<!--��ʾ���-->
<div>
<table width="780" border="1 " cellpadding="0" cellspacing="6" align="center" class="TableBorder">
<tbody>
<tr>
<td valign="top">
<table width="100%" border="0" align="center">
<tr>
<td class="TableName" align="center">�γ̻�����Ϣ��
</td>
</tr>
<tr>
<td>
<table width="100%" border="0" align="center">
<%	response.Write("<tr bgcolor=#EAE2F3>")
	response.Write("<td align=center>�γ̺�</td>")
	response.Write("<td align=center>��ʦ��</td>")
	response.Write("<td align=center>�γ���</td>")
	response.Write("<td align=center>�γ�����</td>")
	response.Write("<td align=center>ѧ��</td>")
	response.Write("<td align=center>ѧʱ</td>")
	response.Write("<td align=center>��������</td>")
	response.Write("<td align=center>�Ƿ�˫ѧλ</td>")
	response.Write("<td align=center colspan=2>����</td>")
	response.Write("</tr>")
	rs.movenext
	'����γ���Ϣ������
	do while not rs.eof
		response.Write("<tr bgcolor=#CCCCCC>")
		for i=0 to rs.fields.count-1
			response.Write("<td class=TableFont align=center>"&rs(i)&"</td>")
		next
		response.Write("<td align=center><a href=""CIupdate.asp?course_number="&rs(0)&""">�޸�</a></td>")
		response.Write("<td  align=center><a href=CIdelete.asp?"&rs(0)&">ɾ��</a></td>")
		response.Write("</tr>")
		rs.movenext
	loop
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
</table>
</td>
</tr>
</table>
</td>
</tr>
</tbody>
</table>
</div>
</body>
</html>
