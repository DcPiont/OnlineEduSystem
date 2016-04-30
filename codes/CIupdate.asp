<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>课程信息表更新</title>
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
<td>课程号：</td>
<input type="hidden" name="course_number" value="<%=rs(0)%>" />
<td><%=rs(0)%></td>
</tr>
<tr>
<td>
<label style="font-size:18px">教师号</label>
</td>
<td>
<input type="text" name="teacher_number" id="TextField" value="<%=rs(1)%>" />
</td>
</tr>
<tr>
<td>课程名</td>
<td>
<input type="text" name="course_name" id="TextField" value="<%=rs(2)%>" />
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">课程类型</label>

</td>
<td>
 <select name = "course_type">
 <%
 	if trim(rs(3))="必修" then
 	response.Write("<option value=""必修"" selected=""seleceted"">必修")
	response.Write("<option value=""选修"">选修")
	else
	response.Write("<option value=""必修"">必修")
	response.Write("<option value=""选修"" selected=""seleceted"">选修")
	end if
 %>
<!--        <option value="bixiu">必修
        <option value="xuanxiu">选修-->
      
</select>
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">学分</label>

</td>
<td>
<input type="text" name="course_score" id="TextField" value="<%=rs(4)%>" />
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">学时</label>

</td>
<td>
<input type="text" name="course_period" id="TextField" value="<%=rs(5)%>" />
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">考试类型</label>

</td>
<td>
<select name = "test_type">
<%
 	if trim(rs(6))="考试" then
 	response.Write("<option value=""考试"" selected=""seleceted"">考试")
	response.Write("<option value=""考查"">考查")
	else
	response.Write("<option value=""考试"">考试")
	response.Write("<option value=""考察"" selected=""seleceted"">考查")
	end if
 %>
<!--        <option value="kaoshi">考试
        <option value="kaocha">考查-->
      
        </select>
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">是否双学位</label>

</td>
<td>
<select name = "double_degree">
<%
 	if trim(rs(7))="是" then
 	response.Write("<option value=""是"" selected=""seleceted"">是")
	response.Write("<option value=""否"">否")
	else
	response.Write("<option value=""是"">是")
	response.Write("<option value=""否"" selected=""seleceted"">否")
	end if
 %>
<!--        <option value="shi">是
        <option value="fou">否-->
      
</select>
</td>
</tr>
<tr>
<td>
<input  type="submit" name="button_submit" id="ButtonStyle2" value="提交" />

<input type="reset" name="button_reset" id="ButtonStyle" value="重置" />
</td>
</tr>
</table>
</form>
<%
if (request("button_submit")="提交") then
'获取信息
teacher_number=trim(request.Form("teacher_number"))
course_name=trim(request.Form("course_name"))
course_type=request.Form("course_type")
course_period=request.Form("course_period")
test_type=request.Form("test_type")
double_degree=request.Form("double_degree")
'判断学时、学分是否为纯数字
course_score=request.Form("course_score")
if trim(course_score)="" then
course_score=0
else if not IsNumeric(course_score) then
response.write("<script>alert('输入不是纯数字！');history.back();</script>")
end if
end if
course_period=request.Form("course_period")
if trim(course_period)="" then
course_period=0
else if not IsNumeric(course_period) then
response.write("<script>alert('输入不是纯数字！');history.back();</script>")
end if
end if
'更新表信息
sql="update CourseInfo set tno='"&teacher_number&"',cname='"&course_name&"',ct='"&course_type&"',cscore='"&course_score&"',cperiod='"&course_period&"',testtype='"&test_type&"',doubledegree='"&double_degree&"' where cno='"&rs(0)&"'"
conn.execute sql
conn.close
response.write("<script type=text/javascript>alert('修改成功！');window.location.href='CImaintenance.asp';</script>")
end if
%>
</body>
</html>
