<%Response.buffer=true%>
<%Response.Expires=0%>
<html>
<!--#include file="common.inc" --> 
<!--#include file="myCommon.inc" --> 
<head>
<TITLE>www.ihimoveis.com.br - Atualização de Arquivos</TITLE>
<meta http-equiv="pragma" content="no-cache">
</head>
<body bgcolor="#FFFFFF">
<div align="center"> 
  <table border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td> 
        <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#993300"><b>Atualiza&ccedil;&atilde;o 
          de arquivos</b></font></div>
      </td>
    </tr>
    <tr> 
      <td> 
        <p><br>
        </p>
        <p><br>
          <br>
          <br>
          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#330033">Essa 
          verifica&ccedil;&atilde;o apaga figuras de im&oacute;veis que n&atilde;o 
          constem mais no cadastro.</font><br>
          <br>
          <br>
          <br>
          <br>
        </p>
      </td>
    </tr>
    <tr> 
      <td> <b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#993300"> 
        <%
set myConn = getMdb(Application("mdbFile"))
sqlStr = "select cod,comment1,comment2,comment3,comment4,comment5,comment6,comment7,comment8 from imoveis;"
set RS = myConn.Execute(sqlStr)
nfile = 0

'//cleans erased imov's pics
on error resume next
imovCount=0
dim cods(200)
dim comment1(200)
dim comment2(200)
dim comment3(200)
dim comment4(200)
dim comment5(200)
dim comment6(200)
dim comment7(200)
dim comment8(200)

RS.MoveFirst
while not RS.EOF
	imovCount=imovCount+1
	cods(imovCount) = RS("cod")
	comment1(imovCount) = RS("comment1")
	comment2(imovCount) = RS("comment2")
	comment3(imovCount) = RS("comment3")
	comment4(imovCount) = RS("comment4")
	comment5(imovCount) = RS("comment5")
	comment6(imovCount) = RS("comment6")
	comment7(imovCount) = RS("comment7")
	comment8(imovCount) = RS("comment8")
	
	RS.MoveNext
wend

Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.GetFolder(realDir & "\imoveis\")
Set fc = f.Files
For Each f1 in fc
	underPos = instr(1,f1.name,"_")
	pointPos = instr(1,f1.name,"p.")
	if pointPos < 1 then  pointPos = instr(1,f1.name,".")

	if underPos < 1 or pointPos < 1 then
		keep = true
	else
		keep = false
		fileCod = left(f1.name, underPos-1)
		seqImg = mid(f1.name,underPos,pointPos-underPos)
	end if
	
	for c=1 to imovCount
			if fileCod = cods(c) then
				keep = true
			end if

		next
	next
	if not keep then 
		fs.DeleteFile(f1) 'erase the pics whose imov isn't listed in the DB
		if Err.description="" then nfile = nfile + 1
	end if
Next

if Err.description="" or nfile=0 then%>
        Atualização realizada com sucesso. <font size="3"><%=nfile%></font> arquivo(s) 
        excluído(s).
<%else%>
        Opera&ccedil;&atilde;o n&atilde;o conclu&iacute;da.<br>
		<%=Err.description%>
<%end if%>

	</font></b></td>
    </tr>
  </table>
  <p align="center"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="3" color="#330033"><br>
    <br>
    </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="http://localhost/adminImov.asp">Sess&atilde;o 
    de administra&ccedil;&atilde;o</a> | <a href="http://www.ihimoveis.com.br">Ir 
    para site On-Line</a></font><font size="3" color="#330033"><br>
    </font></b></font></p>
  <p>&nbsp;</p>
</div>
</body>
</html>
<%Response.flush%>