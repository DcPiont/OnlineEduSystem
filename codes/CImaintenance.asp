<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>课程信息维护</title>
<link href="css.css" rel="stylesheet" type="text/css" />
<!--#include file="connCourseInfo.asp"-->
</head>

<body>
<h3 align="center">请输入要增加记录的相关信息</h3>
<br />
<div align="center">
<form name="form1" method="post" action="CIinsert.asp">
<table class="TableBorder">
<tr>
<td>
<input type="text" name="course_number" id="TextField" placeholder="课程号" />
</td>
<td>
<input type="text" name="teacher_number" id="TextField" placeholder="教师号"/>
</td>
<td>
<input type="text" name="course_name" id="TextField" placeholder="课程名" />
</td>
</tr>
<tr>
<td>
<input type="text" name="course_type" id="TextField" placeholder="课程类型" />
</td>
<td>
<input type="text" name="course_score" id="TextField" placeholder="学分" />
</td>
<td>
<input type="text" name="course_period" id="TextField" placeholder="学时" />
</td>
</tr>
<tr>
<td>
<input type="text" name="test_type" id="TextField" placeholder="考试类型" />
</td>
<td>
<input type="text" name="double_degree" id="TextField" placeholder="是否双学位" />
</td>
<td align="center">
<input type="submit" name="button_submit" id="ButtonStyle" value="提交" />
<input type="reset" name="button_reset" id="ButtonStyle" value="重置" />
</td>
</tr>
</table>
</form>
</div>
<br />
<!--显示表格-->
<div>
<table width="780" border="1 " cellpadding="0" cellspacing="6" align="center" class="TableBorder">
<tbody>
<tr>
<td valign="top">
<table width="100%" border="0" align="center">
<tr>
<td class="TableName" align="center">课程基本信息表
</td>
</tr>
<tr>
<td>
<table width="100%" border="0" align="center">
<%	response.Write("<tr bgcolor=#EAE2F3>")
	response.Write("<td align=center>课程号</td>")
	response.Write("<td align=center>教师号</td>")
	response.Write("<td align=center>课程名</td>")
	response.Write("<td align=center>课程类型</td>")
	response.Write("<td align=center>学分</td>")
	response.Write("<td align=center>学时</td>")
	response.Write("<td align=center>考试类型</td>")
	response.Write("<td align=center>是否双学位</td>")
	response.Write("<td align=center colspan=2>操作</td>")
	response.Write("</tr>")
	rs.movenext
	'输出课程信息表内容
	do while not rs.eof
		response.Write("<tr bgcolor=#CCCCCC>")
		for i=0 to rs.fields.count-1
			response.Write("<td class=TableFont align=center>"&rs(i)&"</td>")
		next
		response.Write("<td align=center><a href=""CIupdate.asp?course_number="&rs(0)&""">修改</a></td>")
		response.Write("<td  align=center><a href=CIdelete.asp?"&rs(0)&">删除</a></td>")
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
