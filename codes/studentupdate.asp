<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>学生信息修改</title>
<link href="css.css" rel="stylesheet" type="text/css" />
</head>
<%
'连接数据库
sql  ="select * from student "
 set conn=server.createobject("adodb.connection")
 conn.open "provider=sqloledb;data source=localhost\SQLEXPRESS;uid=sa;pwd=11111;database=web"
 set rs=server.createobject("adodb.recordset")
 rs.open sql,conn,1,3
 sql="select * from student where sno='"&request("student_number")&"'"
 set rs=conn.execute(sql)
%>
<body>
<form name="form1" method="post" action="studentupdate.asp">
<table align = "center" class="TableBorder">
<tr class="TableName">
<td align="center" colspan="8">学生信息修改
</td>
</tr>
<tr>
<td class="tdTittle">学号：</td>
<input type="hidden" name="student_number" value="<%=rs(0)%>" />
<td class="tdTittle"><%=rs(0)%></td>
<td class="tdTittle">
<label style="font-size:18px">姓名：</label>
</td>
<td>
<input type="text" name="student_name" value="<%=rs(1)%>" />
</td>
<td class="tdTittle">
<label style="font-size:18px">性别：</label>
</td>
<td>
<input type="text" name="sex" id="TextField" value="<%=rs(5)%>" />
</td>
<td class="tdTittle">
<label style="font-size:18px">登录密码：</label>
</td>
<td>
<input type="text" name="pwd" id="TextField" value="<%=rs(7)%>" />
</td>
</tr>
<tr>
<td class="tdTittle">
<label style="font-size:18px">学院：</label>
</td>
<td>
<input type="text" name="xueyuan" value="<%=rs(2)%>" />
</td>
<td class="tdTittle">
<label style="font-size:18px">专业：</label>
</td>
<td>
<input type="text" name="zhuanye" id="TextField" value="<%=rs(3)%>" />
</td>
<td class="tdTittle">
<label style="font-size:18px">班级：</label>
</td>
<td>
<input type="text" name="banji" id="TextField" value="<%=rs(4)%>" />
</td>
<td class="tdTittle">
<label style="font-size:18px">学生类别：</label>
</td>
<td>
<input type="text" name="student_type" id="TextField" value="<%=rs(12)%>" />
</td>
</tr>
<tr>
<td class="tdTittle">
<label style="font-size:18px">出生日期：</label>
</td>
<td>
<input type="text" name="birth" id="TextField" value="<%=rs(6)%>" />
</td>
<td class="tdTittle">
<label style="font-size:18px">出生地：</label>
</td>

<td>
<input type="text" name="place" id="TextField" value="<%=rs(9)%>" />
</td>
<td class="tdTittle">
<label style="font-size:18px">民族：</label>
</td>
<td>
<input type="text" name="minzu" id="TextField" value="<%=rs(8)%>" />
</td>
<td class="tdTittle">
<label style="font-size:18px">政治面貌：</label>
</td>
<td>
<input type="text" name="polity" id="TextField" value="<%=rs(10)%>" />
</td>
</tr>
<tr>
<td class="tdTittle">
<label style="font-size:18px">毕业高中：</label>
</td>
<td>
<input type="text" name="middleschool" id="TextField" value="<%=rs(14)%>" />
</td>
<td class="tdTittle">
<label style="font-size:18px">高考分数：</label>
</td>
<td>
<input type="text" name="tp" id="TextField" value="<%=rs(16)%>" />
</td>
<td class="tdTittle">
<label style="font-size:18px">入学时间：</label>
</td>
<td>
<input type="text" name="entertime" id="TextField" value="<%=rs(11)%>" />
</td>
<td class="tdTittle">
<label style="font-size:18px">身份证号码：</label>
</td>
<td>
<input type="text" name="id" id="TextField" value="<%=rs(15)%>" />
</td>
</tr>
<tr>
<td class="tdTittle">
<label style="font-size:18px">外语语种：</label>
</td>
<td>
<input type="text" name="language" id="TextField" value="<%=rs(13)%>" />
</td>
</tr>
<tr align="center">
<td colspan="8" align="center">
<input  type="submit" name="button_submit" id="ButtonStyle2" value="提交" />
<input type="reset" name="button_reset" id="ButtonStyle" value="重置" />
</td>
</tr>
</table>
</form>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<%
if (request("button_submit")="提交") then
'获取信息
student_name=trim(request.Form("student_name"))
sex=request.Form("sex")
pwd=request.Form("pwd")
xueyuan=request.Form("xueyuan")
zhuanye=request.Form("zhuanye")
banji=request.Form("banji")
student_type=request.Form("student_type")
birth=request.Form("birth")
place=request.Form("place")
minzu=request.Form("minzu")
polity=request.Form("polity")
middleschool=request.Form("middleschool")
tp=request.Form("tp")
if trim(tp)="" then
tp=0
else if not IsNumeric(tp) then
response.write("<script>alert('输入不是纯数字！');history.back();</script>")
end if
end if
entertime=request.Form("entertime")
id=request.Form("id")
language=request.Form("language")
'更新表信息
sql="select * from student where sno='"&request.Form("student_number")&"'"
set conn=server.createobject("adodb.connection")
 conn.open "provider=sqloledb;data source=localhost\SQLEXPRESS;uid=sa;pwd=11111;database=web"
 set rs=server.createobject("adodb.recordset")
 rs.cursorlocation = 3 '添加修改权限
 rs.open sql,conn,1,3
rs("sname")=student_name
rs("college")=xueyuan
rs("major")=zhuanye
rs("class")=banji
rs("sex")=sex
rs("birth")=birth
rs("pwd")=pwd
rs("nation")=minzu
rs("place")=place
rs("polity")=polity
rs("entertime")=entertime
rs("stype")=student_type
rs("language")=language
rs("middleschool")=middleschool
rs("id")=id
rs("tp")=tp
rs.update
response.write("<script type=text/javascript>alert('修改成功！');window.location.href='studentup.asp';</script>")
end if
%>
</body>
</html>
