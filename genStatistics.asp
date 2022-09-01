<%response.buffer=true%>
<!--#include file="common.inc" -->
<html>
<%'get Query
visitDay = Request.QueryString("visitDay")
visitMonth = Request.QueryString("visitMonth")
visitYear = Request.QueryString("visitYear")
visitDay2 = Request.QueryString("visitDay2")
visitMonth2 = Request.QueryString("visitMonth2")
visitYear2 = Request.QueryString("visitYear2")

visitNDay = Request.QueryString("visitNDay")-1
visitItem = Request.QueryString("visitItem")
acompType = Request.QueryString("acompType")
%>

<%'fun animation stuff
if Request.QueryString("animate")="true" and acompType = "" then
	if Request.QueryString("animating")<>1 then
		visitDay2 = visitDay
	end if
	%>
	<META HTTP-EQUIV="REFRESH" CONTENT="1; url=genStatistics.asp?action=analise&acompType=&visitDay=<%=visitDay%>&visitMonth=<%=visitMonth%>&visitYear=<%=visitYear%>&visitDay2=<%=visitDay2+1%>&visitMonth2=<%=visitMonth%>&visitYear2=<%=visitYear%>&button3=Detalhar&animate=true&animating=1">
	<%
	if visitDay2>31 then Response.redirect("showStatistics.asp")
	%>
<%end if
'end of fun
%>

<title>Estat&iacute;sticas Radiology.com.br</title>
<body bgcolor="#FFFFFF">
<div align="left">
  <p>&nbsp;</p>
  <p align="center"><font color="#660033" face="Verdana, Arial, Helvetica, sans-serif" size="5"><b>Estat&iacute;sticas</b></font> 
  </p>
  </div>
<p> 
  <%
'configuration
'multiplicadores do tamanho da barra
percTotal=2 'analise percent bar size
if Request.QueryString("visitItem") <> "*" then
	barHeight = 170
	barWidth = 20
else
	barHeight = 80
	barWidth = 10
end if

'database stuff
Set myConn = getMdb(Application("mdbStat"))

select case Request.QueryString("action")
	case "visitGraph"
		if visitItem <> "*" then visitItem = "theDate," & visitItem
		'get the data
		sqlStr = "select cod from statistics where theDate='" & Date & "' order by cod;"
		set RS = myConn.Execute(sqlStr)
        if RS.EOF and RS.BOF then%>
			</p>
			<p>&nbsp;</p>
			<p>&nbsp;</p>
			<p><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">Erro: 
			  Data inválida (<%=visitDay%>/<%=visitMonth%>/<%=visitYear%>). Tente outra.</font></b> 
		<%else
			dateCod = RS("cod")
			sqlStr = "select " & visitItem & " from statistics where cod between " & dateCod & " and " & (dateCod-visitNDay) & " order by cod;"
			set RS = myConn.Execute(sqlStr)
			'make the friendly graph
			%>
			</p>
			<p>&nbsp; </p>
			<p align="center"><br>
			</p>
					<table border="0" cellspacing="0" cellpadding="0">
			  <tr>
				<td>
				  <p> 
			<%
			fc=visitNDay
			
			'//for bar sizing stuff
			for each rsField in RS.Fields
				if rsField.Name <> "theDate" and rsField.Name <> "cod" then
					while not RS.EOF
						if (RS(rsField.Name) > topField) then topField = RS(rsField.Name)
						RS.MoveNext
					wend
				end if
			next
			'//end

			for each rsField in RS.Fields
				if rsField.Name <> "theDate" and rsField.Name <> "cod" then
					%>
					<br>
					  <br>
					<p> <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660000"><b><%=rsField.Name%></b></font>	
					  <br>
					  <br>
					
					
      <table border="0" cellspacing="0" cellpadding="4" width="100%">
        <tr> 
          <td> 
            <div align="right"><font color="#7D77CA" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>visitas</b></font></div>
          </td>
          <%
						RS.MoveFirst
						month1 = Month(RS("theDate"))
						year1 = Year(RS("theDate"))
						totalField = 0
						if topField=0 then topField = 1
						while not RS.EOF%>

          <td align="center" valign="bottom" bgcolor="#C7CBE9"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=RS(rsField.Name)%><br>
            <img src="imagens/stat.gif" width="<%=barWidth%>" height="<%=barHeight*(RS(rsField.Name)/topField)%>"></font></td>
          <%
							totalField = totalField + RS(rsField.Name)
							month2 = Month(RS("theDate"))
							year2 = Year(RS("theDate"))
							RS.moveNext
						wend
						%>
          <td bgcolor="#006699" align="center"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><%=totalField%></font></b></td>
        </tr>
        <tr> 
          <td> 
            <div align="right"><font color="#7D77CA" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>dia</b></font></div>
          </td>
          <%
					RS.MoveFirst
					while not RS.EOF
						if Weekday(RS("theDate"))=7 or Weekday(RS("theDate"))=1 then
							cellColor = "#494994"
						else
							cellColor = "#432E6D"
						end if
						%>
          <td align="center" valign="top" bgcolor="<%=cellColor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><%=day(RS("theDate"))%></font></td>
          <%
						RS.moveNext
					wend
					%>
          <td bgcolor="#333399" align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b>total</b></font></td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
          <td colspan="<%=Int(fc/2)%>" align="left"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#333366"><%=Mid(MonthName(month1),1,3)%></font></b></td>
          <td colspan="<%=fc-Int(fc/2)+1%>" align="right"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#333366"><%=Mid(MonthName(month2),1,3)%></font></b></td>
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
          <td colspan="<%=Int(fc/2)%>" align="left" valign="top"><i><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><%=year1%></font></b></i></td>
          <td colspan="<%=fc-Int(fc/2)+1%>" align="right" valign="top"><i><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><%=year2%></font></b></i></td>
          <td>&nbsp;</td>
        </tr>
      </table>
	
			<%
			end if 'if non system field do the thing
		next 'draw each session's graph
		%>
		</td>
	  </tr>
	
	  <%'draw totals if showing all
	  if visitItem = "*" then
	  %>
		<tr>
		<td> 
		<p><br>
		<br>
        <font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#3300FF"><b>Total 
        visitas<br>
        </b></font></p>
	     
      <table border="0" cellspacing="0" cellpadding="4" width="100%">
        <tr> 
          <td> 
            <div align="right"><font color="#7D77CA" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>visitas</b></font></div>
          </td>
          <%
		  'draw the totals menu
			RS.MoveFirst
			month1 = Month(RS("theDate"))
			year1 = Year(RS("theDate"))
			totalVisits = 0
			while not RS.EOF
				totalVotes = 0
				for each rsField in RS.Fields
					if rsField.Name <> "theDate" and rsField.Name <> "cod" and rsField.Name <> "visitantes" then
						totalVotes = totalVotes + RS(rsField.Name)
					end if
				next
				totalVisits = totalVisits + totalVotes
				%>
          <td align="center" valign="bottom" bgcolor="#C7CBE9"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=totalVotes%><br>
            <img src="imagens/stat.gif" width="<%=barWidth%>" height="<%=(totalVotes/topField)*barHeight%>"></font></td>
          <%
				month2 = Month(RS("theDate"))
				year2 = Year(RS("theDate"))
				RS.moveNext
			wend
			%>
          <td bgcolor="#006699" align="center"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><%=totalVisits%></font></b></td>
        </tr>
        <tr> 
          <td> 
            <div align="right"><font color="#7D77CA" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>dia</b></font></div>
          </td>
          <%
			RS.MoveFirst
			while not RS.EOF
				if Weekday(RS("theDate"))=7 or Weekday(RS("theDate"))=1 then
					cellColor = "#494994"
				else
					cellColor = "#432E6D"
				end if
				%>
          <td align="center" valign="top" bgcolor="<%=cellColor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><%=Day(RS("theDate"))%></font></td>
          <%
				RS.moveNext
			wend
			%>
          <td align="center" bgcolor="#333399"><b><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">total</font></b></td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
          <td colspan="<%=Int(fc/2)%>" align="left"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#333366"><%=Mid(MonthName(month1),1,3)%></font></b></td>
          <td colspan="<%=fc-Int(fc/2)+1%>" align="right"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#333366"><%=Mid(MonthName(month2),1,3)%></font></b></td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
          <td colspan="<%=Int(fc/2)%>" align="left" valign="top"><i><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><%=year1%></font></b></i></td>
          <td colspan="<%=fc-Int(fc/2)+1%>" align="right" valign="top"><i><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><%=year2%></font></b></i></td>
        </tr>
      </table>
			</td></tr>
		<%end if%>
  
		</table>

		
<%
		end if
		
		
		
		
	'//NUMBERS
	case "numbers"
		sqlStr = "select * from statistics where theDate='" & visitDay & "/" & visitMonth & "/" & visitYear & "' order by cod;"
		set RS = myConn.Execute(sqlStr)
        if RS.EOF and RS.BOF then%>
			<p>&nbsp;</p>
			<p><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">Erro: 
			  Data inválida (<%=visitDay%>/<%=visitMonth%>/<%=visitYear%>). Tente outra.</font></b> 
		<%else%>
			</p>
			<p>&nbsp;</p>
			
<p><font size="3" color="#660033" face="Verdana, Arial, Helvetica, sans-serif"><b>Resumo 
  de </b></font><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#660033"><b><%=visitDay%>/<%=visitMonth%>/<%=visitYear%></b></font></p>
			<p>&nbsp;</p>
			
			<table border="0" cellspacing="0" cellpadding="4" width="102">
			  <tr> 
				<td align="right" bgcolor="#6666CC"><b><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Item</font></b></td>
				<%for each rsField in RS.Fields
					if rsField.Name <> "visitantes" and rsField.Name <> "theDate" and rsField.Name <> "cod" then%>
						<td align="center" bgcolor="#6666CC"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><%=rsField.Name%></font></td>
					<%end if%>
				<%next%>
				<td align="center" bgcolor="#6666CC"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b>total</b></font></td>
			  </tr>
			  <tr bgcolor="#DEDEF3"> 
				<td align="right" bgcolor="#CCCCFF"><b><font color="#6666CC" size="1" face="Verdana, Arial, Helvetica, sans-serif">Visitas</font></b></td>
				<%for each rsField in RS.Fields
					if rsField.Name <> "visitantes" and rsField.Name <> "theDate" and rsField.Name <> "cod" then
						dayTotal = dayTotal + RS(rsField.Name)%>
						<td align="center" bgcolor="#CCCCFF"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%=RS(rsField.Name)%></font></td>
					<%end if%>
				<%next%>
				<td align="center" bgcolor="#CCCCFF"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%=dayTotal%></font></b></td>
			  </tr>
			</table>
			<p><br>
			  <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#660033">N&uacute;mero 
			  de visitantes</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#660033">: 
			  <font size="3"><b><%=RS("visitantes")%></b></font></font></p>
			<p>&nbsp; </p>
			
		
			<p> 
			  <%end if%>
			  <%	



  	'//ANALISE	
	case "analise"
		set totalFields = CreateObject("Scripting.Dictionary")
		set totalPer = CreateObject("Scripting.Dictionary")
		set indexPer = CreateObject("Scripting.Dictionary")
		
		if acompType = "" then%>
			<p><font size="3" color="#660033" face="Verdana, Arial, Helvetica, sans-serif"><b>An&aacute;lise 
			de <font size="2"><%=visitDay%>/<%=visitMonth%>/<%=visitYear%> a <%=visitDay2%>/<%=visitMonth2%>/<%=visitYear2%></font></b></font></p>
		<%else%>
			<p><font size="3" color="#660033" face="Verdana, Arial, Helvetica, sans-serif"><b>Análise dos últimos <%=acompType%> dias</b></font></p>
		
		<%end if

		if acompType = "" then
			sqlStr = "select * from statistics where cod < (select cod from statistics where theDate='" & visitDay & "/" & visitMonth & "/" & visitYear & "') order by cod;"
		else
			sqlStr = "select * from statistics where cod < ((select cod from statistics where theDate='" & Date & "')-" & acompType & ") order by cod;"
		end if
		
  		set RS = myConn.Execute(sqlStr)
		if RS.EOF and RS.BOF then
			%>
			<b><font color="#660033" size="1" face="Verdana, Arial, Helvetica, sans-serif"><i>    * Sem 
			par&acirc;metros para an&aacute;lise</i></font></b> 
			<%

  		else
			rcTotal=0
			RS.MoveFirst
			while not RS.EOF
				for each rsField in RS.Fields
					if rsField.Name <> "theDate" and rsField.Name <> "cod" then
						totalFields(rsField.Name) = totalFields(rsField.Name) + RS(rsField.Name)
					end if
				next
				RS.MoveNext
				rcTotal=rcTotal+1
			wend
		end if

		
		if acompType = "" then
			sqlStr = "select * from statistics where cod between (select cod from statistics where theDate = '" & visitDay & "/" & visitMonth & "/" & visitYear & "') and (select cod from statistics where theDate = '" & visitDay2 & "/" & visitMonth2 & "/" & visitYear2 & "') order by cod;"
		else
			sqlStr = "select cod from statistics where theDate='" & Date & "' order by cod;"
			set RS = myConn.Execute(sqlStr)
			if (RS("cod")-acompType) >= 0 then
				if acompType = 0 then
					sqlStr = "select * from statistics where cod = " & RS("cod") & " order by cod;"
				else
					sqlStr = "select * from statistics where cod >= " & RS("cod")-acompType & " and cod < " & RS("cod") & " order by cod;"
				end if
			else
				sqlStr = "select * from statistics order by cod;"
			end if
		end if

  		set RS = myConn.Execute(sqlStr)
		if RS.EOF and RS.BOF then%>
			<p>&nbsp;</p>
			<p>&nbsp;</p>
			
			<p><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">Erro: 
			  Data inválida (<%=visitDay%>/<%=visitMonth%>/<%=visitYear%> a <%=visitDay2%>/<%=visitMonth2%>/<%=visitYear2%>). 
			  Tente outra.</font></b> 
		<%else%>
			</p>
			<p>&nbsp;</p>
	
				
<table border="0" cellspacing="0" cellpadding="4">
  <tr> 
    <td align="right" bgcolor="#3939AE" height="25"><b><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Item</font></b></td>
    <%for each rsField in RS.Fields
		if rsField.Name <> "theDate" and rsField.Name <> "cod" then%>
    <td align="center" bgcolor="#4B4BC5" height="25"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"> 
      <%if rsField.Name="visitantes" then write("<i><font color=#FFFF00>")%>
      <%=rsField.Name%>
      <%if rsField.Name="visitantes" then write("</i></font>")%>
      </font></b></td>
    <%end if
					next%>
    <td align="center" bgcolor="#3939AE" height="25"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF">total</font></b></td>
  </tr>
  <tr bgcolor="#DEDEF3"> 
    <td align="right" bgcolor="#7B7BD2"><font color="#FFFFFF"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">M&eacute;dia<br>
      Visitas/dia<br>
      </font></b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">(do 
      per&iacute;odo)</font><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
      </font></b></font></td>
    <%
				rcPer=0
				RS.MoveFirst
				while not RS.EOF
					for each rsField in RS.Fields
						if rsField.Name <> "theDate" and rsField.Name <> "cod" then
							totalPer(rsField.Name) = totalPer(rsField.Name) + RS(rsField.Name)
						end if
					next
					RS.MoveNext
					rcPer=rcPer+1
				wend
				
				for each rsField in RS.Fields
					if rsField.Name <> "theDate" and rsField.Name <> "cod" then
					
						avgField = totalPer(rsField.Name)/rcPer
						%>
    <td align="center" bgcolor="#D9D9F2"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%=FormatNumber(avgField,2)%></font></td>
    <%
					if  rsField.Name <> "visitantes" then avgTotal = avgTotal + avgField
					end if
				next%>
    <td align="center" bgcolor="#7E7ED3"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><%=FormatNumber(avgTotal,2)%></font></b></td>
  </tr>
  <tr bgcolor="#DEDEF3"> 
    <%'//variation%>
    <td align="right" bgcolor="#6666CC" rowspan="2"><font color="#FFFFFF"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Varia&ccedil;&atilde;o<br>
      </font></b><font size="1" face="Verdana, Arial, Helvetica, sans-serif">(em 
      rela&ccedil;&atilde;o <br>
      ao per&iacute;odo<br>
      anterior)</font></font></td>
    <%'positive growth
					for each rsField in RS.Fields
						if rsField.Name <> "theDate" and rsField.Name <> "cod" then
							if rcTotal=0 then rcTotal = rcPer
							avgTotal = totalFields(rsField.Name)/rcTotal
							avgPer = totalPer(rsField.Name)/rcPer
							
							if avgTotal=0 then
								if avgPer = 0 then
									indexPer(rsField.Name) = 0
								else
									'dúvida: se antes era 0 e passou a x, qual o crecimento percentual?
									indexPer(rsField.Name) = 1
								end if
							else
								indexPer(rsField.Name) = ((avgPer-avgTotal))/avgTotal 'percentage of growth of the period in relation to the total average
							end if
							
							if indexPer(rsField.Name) >= 0 then
							%>
    <td align="center" bgcolor="#ECECFF" valign="bottom"><b><font size="1" color="#000000" face="Verdana, Arial, Helvetica, sans-serif"><%=FormatPercent(indexPer(rsField.Name),0)%></font></b><b><font size="1" color="#000000" face="Verdana, Arial, Helvetica, sans-serif"><br>
      <img src="imagens/stat.gif" width="<%=barWidth%>" height="<%=Int(barHeight*(indexPer(rsField.Name)/percTotal))%>"><br>
      </font></b></td>
    <%
							else
							%>
    <td bgcolor="#ECECFF">&nbsp;</td>
    <%
							end if
							'total stuff
							if rsField.Name <> "visitantes" then 
								allAvgPer = allAvgPer + avgPer
								allAvgTotal = allAvgTotal + avgTotal
							end if
						end if
					next

					if allAvgTotal=0 then
						if allAvgPer = 0 then
						else
							'dúvida: se antes era 0 e passou a x, qual o crecimento percentual?
							indexTotal = 1
						end if
					else
						indexTotal = ((allAvgPer-allAvgTotal))/allAvgTotal 'percentage of growth of all fields
					end if

 					if indexTotal >= 0 then
						%>
    <td align="center" bgcolor="#9494DA" valign="bottom"><font color="#000033" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><%=FormatPercent(indexTotal,1)%><br>
      <img src="imagens/stat.gif" width="<%=barWidth%>" height="<%=Int(barHeight*(indexTotal/percTotal))%>"></b></font></td>
    <%
					else
					%>
    <td bgcolor="#9494DA">&nbsp;</td>
    <%
					end if
					%>
  </tr>
  <tr> 
    <%'negative growth
					for each rsField in RS.Fields
						if rsField.Name <> "theDate" and rsField.Name <> "cod" then
							avgTotal = totalFields(rsField.Name)/rcTotal
							avgPer = totalPer(rsField.Name)/rcPer
							if avgTotal=0 then avgTotal=1
							if indexPer(rsField.Name) < 0 then
							%>
    <td align="center" bgcolor="#D9D9F2" valign="top"><b><font size="1" color="#000000" face="Verdana, Arial, Helvetica, sans-serif"><img src="imagens/stat.gif" width="<%=barWidth%>" height="<%=Int(barHeight*(indexPer(rsField.Name)/percTotal))*(-1)%>"><br>
      <%=FormatPercent(indexPer(rsField.Name),0)%></font></b></td>
    <%
							else
							%>
    <td bgcolor="#D9D9F2">&nbsp;</td>
    <%
							end if
						end if
					next
 					if indexTotal < 0 then%>
    <td align="center" bgcolor="#7E7ED3" valign="top"><font color="#990000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><img src="imagens/stat.gif" width="<%=barWidth%>" height="<%=Int(barHeight*(indexTotal/percTotal))*(-1)%>"><br>
      <%=FormatPercent(indexTotal,1)%></b></font></td>
    <%
					else
						%>
    <td bgcolor="#7E7ED3">&nbsp;</td>
    <%
					end if%>
  </tr>
</table>
<%end if
		
	case else
		Response.Redirect("showStatistics.asp")
end select

%>
<p>&nbsp;</p>
<p align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><i><b>An&aacute;lise 
  em <%=Date%> às <%=Time%></b></i></font></p>
</body>
</html>
<%Response.Flush%>
<%
RS.Close
myConn.close
set RS = Nothing
set myConn = Nothing

' by Flávio Stutz (flaviostutz@hotmail.com)
%>
