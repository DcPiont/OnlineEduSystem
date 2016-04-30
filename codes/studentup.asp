<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>学生信息维护</title>
<link href="css.css" rel="stylesheet" type="text/css" />
</head>

<body>
<p>&nbsp;</p>
<h3 align="center">请输入要增加记录的相关信息</h3>
<br />
<div align="center">
  <form action="studentinsert.asp" method="post" name="form1" id="form1">
   <table align = "center" class="TableBorder">



<tr>
<td>
<label style="font-size:18px">学号</label>

</td>
<td>
<input type="text" name="xuehao" id="TextField" placeholder="学生号" />
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">姓名</label></td>
<td>
 <input type="text" name="xuesname" id="TextField" placeholder="姓名" />
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">班级</label>

</td>
<td>
<input type="text" name="class" id="TextField" placeholder="班级" />
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">身份证号
</label>

</td>
<td>
<input type="text" name="ID" id="TextField" placeholder="ID" />
</td>
</tr>
<tr>
<td>
<label style="font-size:18px">性别</label>

</td>
<td>
<select name = "sex">
        <option value="男">男
        <option value="女">女
      
        </select>
</td>
</tr>
<tr>
<td>专业</td>
<td>
 <select name = "zhuaniye">
        <option value="软件工程">软件工程
        <option value="数字媒体">数字媒体
        <option value="电子">电子
        <option value="计算机">计算机
        </select></td>
</tr>
<tr>
<td align="center" colspan="2">
<input  type="submit" name="button_submit" id="ButtonStyle2" value="提交" />

<input type="reset" name="button_reset" id="ButtonStyle" value="重置" />
</td>
</tr>
</table>
  </form>
</div>
<br />
'
<%
''连接数据库
' dim sql
' dim course_score
' dim course_period
' sql  ="select * from CourseInfo "
' set conn=server.createobject("adodb.connection")
' conn.open "provider=sqloledb;source=(local);uid=sa;pwd=123;database=website"
' set rs=server.createobject("adodb.recordset")
' rs.open sql,conn,1,3'添加数据
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
<!--显示表格-->
<table width="600" border="1 " cellpadding="0" cellspacing="6" align="center" class="TableBorder">
  <tbody>
    <tr>
      <td valign="top"><table width="100%" border="0" align="center">
        <tr>
          <td class="TableName" align="center">学生基本信息表 </td>
        </tr>
      </table>
        <table width="100%" border="0" align="center">
<%
 sql  ="select sno,sname,sex,major,class from student "
 set conn=server.createobject("adodb.connection")
 conn.open "provider=sqloledb;data source=localhost\SQLEXPRESS;uid=sa;pwd=11111;database=web"
 set rs=server.createobject("adodb.recordset")	
response.Write("<tr bgcolor=#EAE2F3>")
rs.open sql,conn,1,3
	response.Write("<td align=center>学号</td>")
	response.Write("<td align=center>姓名</td>")
	response.Write("<td align=center>性别</td>")
	response.Write("<td align=center>专业</td>")
	response.Write("<td align=center>班级</td>")
	response.Write("<td align=center colspan=2>操作</td>")
	response.Write("</tr>")
	do while not rs.eof
		response.Write("<tr bgcolor=#CCCCCC>")
		for i=0 to rs.fields.count-1
			response.Write("<td class=TableFont align=center>"&rs(i)&"</td>")
			
		next
		response.Write("<td align=center><a href=""studentupdate.asp?student_number="&rs(0)&""">修改</a></td>")
		response.Write("<td align=center><a href=""studentdelete.asp?student_number="&rs(0)&""">删除</a></td>")
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
