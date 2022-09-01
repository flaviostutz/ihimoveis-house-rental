<!--#include file="common.inc" --> 
<html>
<head>
<title>MdbSql - Access database management - Snack</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF">
<br>
<form name="form1" method="get">
  <input type="password" name="password" value="<%=Request.QueryString("password")%>">
<br>
  <p><font face="Verdana, Arial, Helvetica, sans-serif" size="3"><b><font color="#800000">Access 
    SQL command:<br>
    </font></b></font> 
    <input type="text" name="sqlStr" size="75">
    <input type="submit" name="action" value="Executar">
  </p>
  </form>
<p><br>
  <font face="Verdana, Arial, Helvetica, sans-serif" size="3"><b><font color="#800000">SQL 
  Result:</font></b></font> </p>
<blockquote> 
  <blockquote> 
    <blockquote> 
      <blockquote>
        <p><br>
          <%
'On Error Resume Next 

if Request.QueryString("action") = "Executar" and Request.QueryString("password")="straight" then
	Set myConn = Server.CreateObject("ADODB.Connection")
	myConn.Open("DBQ=" & Server.MapPath(Application("mdbFile")) & ";Driver={Microsoft Access Driver (*.mdb)}")
	SqlStr = Request.QueryString("sqlStr")
	Set RS = myConn.Execute(SqlStr)
	counter = 0

	if inStr(1,LCase(mid(sqlStr,1,10)),"select") > 0 then
		write("<table bgcolor=#000000>")
		for c=0 to RS.Fields.count-1
			write("<td bgcolor=#000000><font face=verdana color=#ffffff><b>" & RS.Fields(c).Name & "</font></b></td>")
		next
		RS.MoveFirst
		while not RS.EOF
			write("<tr>")
				for c=0 to RS.Fields.count-1
					write("<td bgcolor=#ffffff><font face=verdana>" & RS.Fields(c) & "</font></td>")
				next
			write("</tr>")
			counter = counter + 1
			RS.moveNext
		wend
		write("</table>")
		%>
		<font face="Verdana, Arial, Helvetica, sans-serif"><i><font size="3">Foram 
          encontrados<b> <%=counter%> </b>registros</font></i></font><%
	else
		write("<font face=verdana size=3 color=#000000><b>Não foi encontrado nenhum registro para ser mostrado.</b></font><br>")
		write("<font face=verdana size=3 color=#000000><b>Instrução processada.</b></font>")
	end if
else%> <i><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><font color="#800000">	
          <font size="3"> Digite sua autoriza&ccedil;&atilde;o e a instru&ccedil;&atilde;o 
          SQL</font></font></font></i> <%end if
%> </p>
	</blockquote>
    </blockquote>
  </blockquote>
</blockquote>

<p>&nbsp;</p>
<blockquote>
  <p><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
    <%	
	Response.write(Err.description) & "<br>"
	Response.write(Err.source)
%>
    </font></b></p>
</blockquote>
<p>&nbsp; </p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>ALTER TABLE 
  table {ADD {COLUMN field type[(size)] [NOT NULL] [CONSTRAINT index] | CONSTRAINT 
  multifieldindex} | DROP {COLUMN field I CONSTRAINT indexname} }</b></font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">ex.: ALTER TABLE 
  Employees ADD COLUMN Notes TEXT(25)</font></p>
<p>&nbsp;</p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>INSERT INTO 
  target [IN externaldatabase] [(field1[, field2[, ...]])] SELECT [source.]field1[, 
  field2[, ...] FROM tableexpression</b></font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">ex.: INSERT INTO 
  Clientes SELECT * FROM NovosClientes;</font></p>
<p>&nbsp;</p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>DELETE [table.*] 
  FROM table WHERE criteria;</b></font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">ex.: DELETE * FROM 
  Employees WHERE Title = 'Trainee';</font></p>
<p>&nbsp;</p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>DROP {TABLE 
  table | INDEX index ON table}</b></font></p>
<p>ex.: DROP INDEX NewIndex ON Employees;</p>
<p>&nbsp;</p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>UPDATE table 
  SET newvalue WHERE criteria;</b></font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">ex.: UPDATE Funcionários 
  SET Supervisor = 5 WHERE Supervisor = 2;</font></p>
</body>
</html>
<%
RS.Close
myConn.close
set RS = Nothing
set myConn = Nothing

' by Flávio Stutz (flaviostutz@hotmail.com)
%>

