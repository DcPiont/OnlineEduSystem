<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>下学期课程安排</title>
<link href="css.css" rel="stylesheet" type="text/css" />
</head>

<body>
<h3 align="center">查询所有必修课信息</h3>
<div align="center">
<form action="nextCoursePlan.asp" method="post" name="form" id="form">
<input type="submit" name="button_search" id="ButtonStyle" value="查询" />
<input type="submit" name="button_return" id="ButtonStyle" value="返回" />
</form>
</div>
<!--分界线-->
<hr color="#9259B0" />
<form action="nextCPinsert.asp" method="post" name="form1" id="form1">
<%
	if (request("button_search")="返回") then
	response.Redirect("nextCoursePlan.asp")
	end if
	'查询课程信息
	if (request("button_search")="查询") then
	'连接数据库
 sql  ="select * from CourseInfo "
 set conn=server.createobject("adodb.connection")
 conn.open "provider=sqloledb;data source=localhost\SQLEXPRESS;uid=sa;pwd=11111;database=web"
 set rs=server.createobject("adodb.recordset")
 rs.open sql,conn,1,3
	response.Write("<table border=1 cellpadding=0 cellspacing=6 align=center class=TableBorder><tbody><tr><td valign=top><table width=100% border=0 align=center><tr><td class=TableName align=center>必修课信息表</td></tr></table><table width=100% border=0 align=center>")
	response.Write("<tr bgcolor=#EAE2F3>")
	response.Write("<td align=center>课程号</td>")
	response.Write("<td align=center>教师名</td>")
	response.Write("<td align=center>课程名</td>")
	response.Write("<td align=center>课程类型</td>")
	response.Write("<td align=center>学分</td>")
	response.Write("<td align=center>学时</td>")
	response.Write("<td align=center>考试类型</td>")
	response.Write("<td align=center>是否双学位</td>")
	response.Write("<td align=center colspan=2>是否添加</td>")
	response.Write("</tr>")
	sql="select * from CourseInfo where ct in ('必修')"
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
	'连接数据库
 sql  ="select nextCourseArrange.cno,cname,TermNum,MajorNum,MaxStu,PresentStu,week,day,hour,wtow,place from nextCourseArrange,CourseInfo where nextCourseArrange.cno=CourseInfo.cno "
 set conn=server.createobject("adodb.connection")
 conn.open "provider=sqloledb;data source=localhost\SQLEXPRESS;uid=sa;pwd=11111;database=web"
 set rs=server.createobject("adodb.recordset")
 rs.open sql,conn,1,3
'显示下学期课程安排
	response.Write("<table border=1 cellpadding=0 cellspacing=6 align=center class=TableBorder")
	response.Write("><tbody><tr><td valign=top><table width=100% border=0 align=center><tr><td class=TableName align=center>下学期课程安排表</td></tr></table><table width=100% border=0 align=center>")
	response.Write("<td align=center>课程号</td>")
	response.Write("<td align=center>课程名</td>")
	response.Write("<td align=center>学期号</td>")
	response.Write("<td align=center>专业号</td>")
	response.Write("<td align=center>课容量</td>")
	response.Write("<td align=center>当前人数</td>")
	response.Write("<td align=center>上课星期</td>")
	response.Write("<td align=center>上课节次</td>")
	response.Write("<td align=center>上课学时</td>")
	response.Write("<td align=center>上课周次</td>")
	response.Write("<td align=center>上课地点</td>")
	response.Write("<td align=center colspan=2>操作</td>")
	response.Write("</tr>")
	do while not rs.eof
		response.Write("<tr bgcolor=#CCCCCC>")
		for i=0 to rs.fields.count-1
			response.Write("<td class=TableFont align=center>"&rs(i)&"</td>")
		next
		response.Write("<td align=center><a href=""nextCPupdate.asp?course_number="&rs(0)&"&term_number="&rs(2)&""">修改</a></td>")
		response.Write("<td align=center><a href=""nextCPdelete.asp?course_number="&rs(0)&"&term_number="&rs(2)&""">删除</a></td>")
		response.Write("</tr>")
		rs.movenext
	loop
end if
%>
</table>
</tbody>
</table>
<!--分界线-->
<hr color="#9259B0" />
<div align="center">
    <table align ="center" class="TableBorder">
      <tr>
        <td align="center">请选择专业课名</td>
        <td align="center">
        <select name = "zhuanye">
        <option value="0101">软件工程</option>
        <option value="0102">数字媒体</option>
        <option value="0103">电子</option>
        <option value="0104">计算机</option>
        </select>
        </td>
        <td>请输入学期号：
        <input type="text" name="xueqi" />
        </td>
        <td align="center"><input type="submit" name="button_submit" id="ButtonStyle" value="提交" /></td>
      </tr>
    </table>
  </form>
</div>
</body>
</html>
