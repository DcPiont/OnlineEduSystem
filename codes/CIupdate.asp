<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�γ���Ϣ�����</title>
<!--#include file="connCourseInfo.asp"-->
<%
dim rs
sql="select * from CourseInfo where cno='"&request("course_number")&"'"
set rs=conn.execute(sql)
%>
</head>

<body>
<form name="form1" method="post" action="CIupdate.asp">
<table align = "center" class="TableBorder">
<tr>
<td>�γ̺ţ�</td>
<input type="hidden" name="course_number" value="<%=rs(0)%>" />
<td><%=rs(0)%></td>
</tr>
<tr>
<td>
<label style="font-size:18px">��ʦ��</label>
</td>
<td>
<input type="text" name="teacher_number" id="TextField" value="<%=rs(1)%>" />
</td>
</tr>
<tr>
<td>�γ���</td>
<td>
<input type="text" name="course_name" id="TextField" value="<%=rs(2)%>" />
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">�γ�����</label>

</td>
<td>
 <select name = "course_type">
 <%
 	if trim(rs(3))="����" then
 	response.Write("<option value=""����"" selected=""seleceted"">����")
	response.Write("<option value=""ѡ��"">ѡ��")
	else
	response.Write("<option value=""����"">����")
	response.Write("<option value=""ѡ��"" selected=""seleceted"">ѡ��")
	end if
 %>
<!--        <option value="bixiu">����
        <option value="xuanxiu">ѡ��-->
      
</select>
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">ѧ��</label>

</td>
<td>
<input type="text" name="course_score" id="TextField" value="<%=rs(4)%>" />
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">ѧʱ</label>

</td>
<td>
<input type="text" name="course_period" id="TextField" value="<%=rs(5)%>" />
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">��������</label>

</td>
<td>
<select name = "test_type">
<%
 	if trim(rs(6))="����" then
 	response.Write("<option value=""����"" selected=""seleceted"">����")
	response.Write("<option value=""����"">����")
	else
	response.Write("<option value=""����"">����")
	response.Write("<option value=""����"" selected=""seleceted"">����")
	end if
 %>
<!--        <option value="kaoshi">����
        <option value="kaocha">����-->
      
        </select>
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">�Ƿ�˫ѧλ</label>

</td>
<td>
<select name = "double_degree">
<%
 	if trim(rs(7))="��" then
 	response.Write("<option value=""��"" selected=""seleceted"">��")
	response.Write("<option value=""��"">��")
	else
	response.Write("<option value=""��"">��")
	response.Write("<option value=""��"" selected=""seleceted"">��")
	end if
 %>
<!--        <option value="shi">��
        <option value="fou">��-->
      
</select>
</td>
</tr>
<tr>
<td>
<input  type="submit" name="button_submit" id="ButtonStyle2" value="�ύ" />

<input type="reset" name="button_reset" id="ButtonStyle" value="����" />
</td>
</tr>
</table>
</form>
<%
if (request("button_submit")="�ύ") then
'��ȡ��Ϣ
teacher_number=trim(request.Form("teacher_number"))
course_name=trim(request.Form("course_name"))
course_type=request.Form("course_type")
course_period=request.Form("course_period")
test_type=request.Form("test_type")
double_degree=request.Form("double_degree")
'�ж�ѧʱ��ѧ���Ƿ�Ϊ������
course_score=request.Form("course_score")
if trim(course_score)="" then
course_score=0
else if not IsNumeric(course_score) then
response.write("<script>alert('���벻�Ǵ����֣�');history.back();</script>")
end if
end if
course_period=request.Form("course_period")
if trim(course_period)="" then
course_period=0
else if not IsNumeric(course_period) then
response.write("<script>alert('���벻�Ǵ����֣�');history.back();</script>")
end if
end if
'���±���Ϣ
sql="update CourseInfo set tno='"&teacher_number&"',cname='"&course_name&"',ct='"&course_type&"',cscore='"&course_score&"',cperiod='"&course_period&"',testtype='"&test_type&"',doubledegree='"&double_degree&"' where cno='"&rs(0)&"'"
conn.execute sql
conn.close
response.write("<script type=text/javascript>alert('�޸ĳɹ���');window.location.href='CImaintenance.asp';</script>")
end if
%>
</body>
</html>
