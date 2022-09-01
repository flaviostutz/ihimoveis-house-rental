<%Response.buffer=true%>
<%Response.Expires=60%>
<html>
<!--#include file="common.inc" --> 
<!--#include file="myCommon.inc" --> 
<style TYPE="text/css">
<!--
	A:link {color:"#006666"; text-decoration: none}
	A:visited {color: "#006666"; text-decoration: none}
	A:hover {color:"#009966"; text-decoration: underline}
-->
</style>
<%
On error resume next
transaction = Request.QueryString("transaction") 'whatever show available sales or rent
itype = Request.QueryString("itype") 'whatever show which type of imov
%>
<head>
<title>Imobiliária Higienópolis - www.ihimoveis.com.br</title>
</head>
<body bgcolor="#FFFFFF" text="#000000" link="#006666" vlink="#006666" alink="#009966">
<%
Set fs = CreateObject("Scripting.FileSystemObject")
set myConn = getMdb(Application("mdbFile"))

if transaction=1 and itype>0 then
	sqlStr = "select * from imoveis where transaction=" & transaction & " and itype=" & itype & " order by cod;"
elseif transaction=2 and itype>0 then
	sqlStr = "select * from imoveis where transaction=" & transaction  & " and itype=" &itype & " order by cod;"
else
'	sqlStr = "select * from imoveis order by cod;"
end if

set rs = myConn.Execute(sqlStr)
realDir = Server.MapPath(".")

if RS.EOF and RS.BOF then 
	%>
<br>
<br>
<p align="center"><br>
  <font color="#333366" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>N&atilde;o 
  dispomos d<font color="#000000">e</font><font size="3" color="#FF0033"> <%=LCase(itypeStr(itype))%>s 
  </font>par<font color="#000000">a</font><font size="3" color="#FF0033"> <%=LCase(transactionStr(transaction))%> 
  </font>hoje. </b></font></p>
<p align="center"><font color="#333366" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>Contate-nos 
  para maiores informa&ccedil;&otilde;es.</b></font><font color="#333366"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><br>
  <br>
  </font> </b> </font> <%
else
	rs.MoveFirst

	while not rs.EOF
	%> <br>
</p>
<div align="center">
  <table border="0" width="96%" cellpadding="0" cellspacing="3" bgcolor="#f0f0f0">
    <tr> 
      <td width="21%" align="center" valign="middle"> 
        <div align="center"> 
          <%if fs.FileExists(realDir & "\imoveis\" & RS("cod") & "_1.jpg") then%>
          <a href="showDetails.asp?cod=<%=RS("cod")%>" target="_self"><img border="1" src="<%="imoveis/" & RS("cod")&"_1p.jpg"%>"> 
          </a></div>
        <%else%>
        <font color="#006666" size="1" face="tahoma"><b>FOTOS N&Atilde;O<br>
        DISPONIVEIS</b></font> 
        <%end if%>
      </td>
      <td width="79%"> 
        <table border="0" width="100%" cellpadding="0" cellspacing="2">
          <tr> 
            <td colspan="4"><font face="Arial, Helvetica, sans-serif" size="1" color="#000000"><%=iTypeStr(RS("itype"))%> 
              - Código <%=RS("cod")%></font></td>
          </tr>
          <tr> 
            <td colspan="4"><font size="2" color="#000000" face="Verdana, Arial, Helvetica, sans-serif"><b><%=RS("building")%> 
              <%if RS("building")<>"" then%>
              - </b> 
              <%end if%>
              <%=RS("quarter")%> 
              <%if RS("building")<>"" then write "</b>"%>
              </font></td>
          </tr>
          <tr> 
            <td colspan="4"><font face="Arial, Helvetica, sans-serif" color="#000000"><b><font size="2">Endereço:</font></b> 
              <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("address")%></font><font size="2"> 
              - </font><font face="Verdana, Arial, Helvetica, sans-serif" color="#000000" size="2"><%=RS("city")%></font></font></td>
          </tr>
		  <%if RS("rooms") <> "" or RS("area") <> "" then%>
			  <tr> 
				<%if RS("rooms") <> "" then%>
					<td colspan="2"> 
					  <font face="Arial, Helvetica, sans-serif" color="#000000"><b><font size="2">Nº 
					  de quartos:</font></b> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("rooms")%></font></font> 
					</td>
				<%end if%>
				<%if RS("area") <> "" then%>
					
            <td colspan="2"> <font face="Arial, Helvetica, sans-serif" color="#000000"><b><font size="2">&Aacute;rea 
              total:</font></b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000000" size="2"> 
              <%=RS("area")%></font></font>m&sup2; </td>
				<%end if%>
				<%if RS("rooms") = "" or RS("area") = "" then%><td>&nbsp; </td><%end if%>
			  </tr>
			 <%end if%>
          <tr> 
            <td colspan="2"><font face="Arial, Helvetica, sans-serif" color="#000000"><b><font size="2">Valor:</font></b> 
              <font size="2" face="Verdana, Arial, Helvetica, sans-serif"> <%=FormatCurrency(RS("price"))%></font></font></td>
            <td align="right" valign="bottom" width="36%">&nbsp; </td>
            <td width="35%" align="right" valign="bottom"> 
              <%if fs.FileExists(realDir & "\imoveis\" & RS("cod") & "_1.jpg") then%>
              <a href="showDetails.asp?cod=<%=RS("cod")%>" target="_self"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="imagens/bot_fotosdetalhes.gif" width="115" height="16" border="0"></font></b></a> 
              <%else%>
              <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><b><a href="showDetails.asp?cod=<%=RS("cod")%>" target="_self">Veja 
              detalhes</a></b></font> 
              <%end if%>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>

 <%
	rs.MoveNext

wend
end if
%>
  <table width="96%" border="0" cellspacing="0" cellpadding="0" height="20">
    <tr valign="bottom"> 
      <td width="73%"><br>
        <a href="javascript:history.go(-1)" target="middle"><img src="imagens/bot_voltar.gif" border="0"></a></td>
      </tr>
  </table>
</div>
</body>
</html>
<%Response.flush%>