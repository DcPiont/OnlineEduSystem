<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>历史课程安排信息修改</title>
<!--#include file="connOldCoursePlan.asp"-->
</head>
<%
dim rs
sql="select oldCourseArrange.cno,cname,TermNum,MajorNum,MaxStu,PresentStu,week,day,hour,wtow,place from oldCourseArrange,CourseInfo where oldCourseArrange.cno=CourseInfo.cno and oldCourseArrange.cno='"&request("course_number")&"' and oldCourseArrange.TermNum='"&request("term_number")&"'"
set rs=conn.execute(sql)
%>
<body>
<form name="form1" method="post" action="oldCPupdate.asp">
<table align = "center" class="TableBorder">
<tr>
<td>课程号：</td>
<input type="hidden" name="course_number" value="<%=rs(0)%>" />
<td><%=rs(0)%></td>
</tr>
<tr>
<td>
<label style="font-size:18px">课程名</label>
</td>
<input type="hidden" name="term_number" value="<%=rs(2)%>" />
<td>
<%=rs(1)%>
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">学期号</label>
</td>
<td>
<%=rs(2)%>
</td>
</tr>
<tr>
<td>专业号</td>
<td>
<input type="text" name="major_number" id="TextField" value="<%=rs(3)%>" />
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">课容量</label>
</td>
<td>
<input type="text" name="max_student" id="TextField" value="<%=rs(4)%>" />
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">已选人数</label>
</td>
<td>
<input type="text" name="present_student" id="TextField" value="<%=rs(5)%>" />
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">上课星期</label>

</td>
<td>
<input type="text" name="week" id="TextField" value="<%=rs(6)%>" />
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">上课节次</label>
</td>
<td>
<input type="text" name="jieci" id="TextField" value="<%=rs(7)%>" />
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">上课学时</label>
</td>
<td>
<input type="text" name="xueshi" id="TextField" value="<%=rs(8)%>" />
</td>
<tr>
<td>
<label style="font-size:18px">上课周次</label>
</td>
<td>
<input type="text" name="wtow" id="TextField" value="<%=rs(9)%>" />
</td>
<tr>
<td>
<label style="font-size:18px">上课地点</label>
</td>
<td>
<input type="text" name="place" id="TextField" value="<%=rs(10)%>" />
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
major_number=trim(request.Form("major_number"))
max_student=request.Form("max_student")
present_student=request.Form("present_student")
week=request.Form("week")
jieci=request.Form("jieci")
xueshi=request.Form("xueshi")
wtow=request.Form("wtow")
place=request.Form("place")
'判断人数是否为纯数字
if trim(max_student)="" then
max_student=0
else if not IsNumeric(max_student) then
response.write("<script>alert('输入不是纯数字！');history.back();</script>")
end if
end if
if trim(present_student)="" then
present_student=0
else if not IsNumeric(present_student) then
response.write("<script>alert('输入不是纯数字！');history.back();</script>")
end if
end if
'response.Write(major_number)
'response.Write(max_student)
'response.Write(present_student)
'response.Write(week)
'response.Write(jieci)
'response.Write(xueshi)
'response.Write(wtow)
'response.Write(place)
'更新表信息
sql="update oldCourseArrange set MajorNum='"&major_number&"',MaxStu='"&max_student&"',PresentStu='"&present_student&"',week='"&week&"',day='"&jieci&"',hour='"&xueshi&"',wtow='"&wtow&"',place='"&place&"' where cno='"&rs(0)&"' and TermNum='"&rs(2)&"'"
conn.execute sql
conn.close
response.write("<script type=text/javascript>alert('修改成功！');window.location.href='oldCoursePlan.asp';</script>")
end if
%>
</body>
</html>
