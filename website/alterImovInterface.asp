<%Response.buffer=true%>
<%Response.Expires=0%>
<html>
<!--#include file="common.inc" --> 
<!--#include file="myCommon.inc" --> 
<head>
<meta http-equiv="pragma" content="no-cache">
<title>Imobiliária Higienópolis - Administração</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<style TYPE="text/css">
<!--
	A:link {color:"darkblue"; text-decoration: none}
	A:visited {color: "darkblue"; text-decoration: none}
	A:hover {color:"red"; text-decoration: underline}
-->
</style>

<body bgcolor="#FFFFFF" text="#000000">
<p align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#336666"><b>Altera&ccedil;&atilde;o 
  de Im&oacute;veis</b></font></p>
<p><font size="2" color="#336666" face="Verdana, Arial, Helvetica, sans-serif"><b>Selecione 
  o im&oacute;vel a alterar/apagar:</b></font> <br>
  <br>
<%
set myConn = getMdb(Application("mdbFile"))
if Request.QueryString("action") = "delete" then
on error resume next
	sqlStr = "delete * from imoveis where cod='" & Request.QueryString("cod") & "';"
	set RS = myConn.Execute(sqlStr)
	if Err.description <> "" then 
		write Err.description
	else
		'apaga somente os arquivos do OFFLINE	
		fs.DeleteFile(realDir & "\imoveis\" & Request.QueryString("cod") & "_*.*")
		'substitui o banco de dados online
		fs.CopyFile realDir & "\" & Application("mdbFile"), realDir & "\saida\" & Application("mdbFile"), true
	end if
end if

sqlStr = "select * from imoveis order by cod;"
set RS = myConn.Execute(sqlStr)

if RS.EOF and RS.BOF then
	%>
<i> <b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000"><br>
  <font size="1">Não existe nenhum imóvel cadastrado.</font></font></b> </i> 
  <%
else
	RS.MoveFirst
	%>
<table width="100%" border="0" cellspacing="0" cellpadding="3">
  <tr bgcolor="#205391"> 
    <td width="10%" bgcolor="#205391">&nbsp;</td>
    <td width="8%"><b><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">Cod</font></b></td>
    <td width="10%"><b><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">Transa&ccedil;&atilde;o</font></b></td>
    <td width="10%"><b><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">Tipo</font></b></td>
    <td width="32%" bgcolor="#205391"><b><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">Endere&ccedil;o</font></b></td>
    <td width="15%"><b><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">Foto</font></b></td>
    <td width="6%">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="7" bgcolor="#D5D9EE"> 
      <table border="0" cellspacing="1" cellpadding="4" width="100%">
        <%while not RS.EOF%>
        <%for i=1 to 8%>
        <input type="hidden" name="<%="comment"&i%>" value="<%=RS("comment"&i)%>">
          <%next%>
          <tr bgcolor="#DAE8FE"> 
			<form name="form<%=RS("cod")%>" method="get" action="cadImovInterface.asp">
            <td align="right" width="10%"><b> 
              <input type="submit" name="submit<%=RS("cod")%>" value="Alterar">
              </b></td>
            <td width="8%"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("cod")%> 
              </font></b></td>
            <td width="10%"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=transactionStr(RS("transaction"))%> 
              </font></b></td>
            <td width="10%"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=itypeStr(RS("itype"))%> 
              </font></b></td>
            <td width="32%" bgcolor="#DAE8FE"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("address")%> 
              </font></b></td>
            <td width="15%" align="center"> <b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"> 
              <%if fs.FileExists(realDir & "\imoveis\" & RS("cod") & "_1p.jpg") then%>
              <img src=<%="file:///" & realDir & "\imoveis\" & RS("cod") & "_1p.jpg"%> border="0" width="80" height="60"> 
              <%else%>
              Sem fotos 
              <%end if%>
              </font></b></td>
            <td width="6%"> 
              <div align="center"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><a href="alterImovInterface.asp?action=delete&cod=<%=RS("cod")%>" onClick="return confirm('Deseja realmente apagar esse imóvel?');">APAGAR</a></font></b></div>
            </td>
          <%
	if RS("IPTU")>=0 then
		IPTU = RS("IPTU")
	else
		IPTU = "Isento"
	end if
	%>
          <input type="hidden" name="address" value="<%=RS("address")%>">
          <input type="hidden" name="itype" value="<%=RS("itype")%>">
          <input type="hidden" name="transaction" value="<%=RS("transaction")%>">
          <input type="hidden" name="cod" value="<%=RS("cod")%>">
          <input type="hidden" name="quarter" value="<%=RS("quarter")%>">
          <input type="hidden" name="city" value="<%=RS("city")%>">
          <input type="hidden" name="building" value="<%=RS("building")%>">
          <input type="hidden" name="rooms" value="<%=RS("rooms")%>">
          <input type="hidden" name="area" value="<%=RS("area")%>">
          <input type="hidden" name="IPTU" value="<%=IPTU%>">
          <input type="hidden" name="cond" value="<%=RS("cond")%>">
          <input type="hidden" name="description" value="<%=RS("description")%>">
          <input type="hidden" name="lote" value="<%=RS("lote")%>">
          <input type="hidden" name="price" value="<%=RS("price")%>">
          <input type="hidden" name="comment1" value="<%=RS("comment1")%>">
          <input type="hidden" name="comment2" value="<%=RS("comment2")%>">
          <input type="hidden" name="comment3" value="<%=RS("comment3")%>">
          <input type="hidden" name="comment4" value="<%=RS("comment4")%>">
          <input type="hidden" name="comment5" value="<%=RS("comment5")%>">
          <input type="hidden" name="comment6" value="<%=RS("comment6")%>">
          <input type="hidden" name="comment7" value="<%=RS("comment7")%>">
          <input type="hidden" name="comment8" value="<%=RS("comment8")%>">
		</form>
          </tr>
        <%
  	RS.MoveNext
	wend
%>
      </table>
    </td>
  </tr>
</table>
<p><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
  <%
end if
%>
  <br>
  <br>
  </font></b></p>
<p align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#006699"><a href="adminImov.asp"><b>Sess&atilde;o 
  de administra&ccedil;&atilde;o</b></a></font></p>
<p>&nbsp;</p>
</body>
</html>
<%Response.flush%>