
<%
transaction=Request.QueryString("transaction")
itype=Request.QueryString("itype")
%>
<html>
<frameset framespacing="0" border="0" rows="76,*,33" frameborder="0" cols="*">
  <frame name="top" src="top.asp?transaction=<%=transaction%>&itype=<%=itype%>" scrolling="no" noresize>
  <frame name="principal" src="showImov.asp?transaction=<%=transaction%>&itype=<%=itype%>" scrolling="auto" noresize>
  <frame name="menu" scrolling="no" noresize src="menu.html" target="_self">
  <noframes> 
  <body topmargin="0" leftmargin="0">
  <p>Esta página usa quadros mas seu navegador não aceita quadros.</p>
  </body>
  </noframes> </frameset>
</html>
