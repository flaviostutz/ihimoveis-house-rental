<!--#include file="common.inc" --> 
<html>
<head>
<title>www.ihimoveis.com.br - Estatísticas</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">

<p align="center">&nbsp;</p>
<p align="center"><font color="#990000"><b><font size="4"><font face="Verdana, Arial, Helvetica, sans-serif" size="5">Estat&iacute;sticas</font></font></b></font></p>
<p align="left">&nbsp;</p>
<form name="form1" method="get" action="genStatistics.asp">
  <table border="0" cellspacing="0" cellpadding="8" width="434">
    <tr> 
      <td colspan="3" bgcolor="#3A76B1"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b><font size="3">Gr&aacute;fico 
        de Visitas 
        <input type="hidden" name="action" value="visitGraph">
        </font></b></font></td>
    </tr>
    <tr bgcolor="#E9F0F8"> 
      <td width="18%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">Per&iacute;odo: 
        </font></td>
      <td colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033"> 
        <select name="visitNDay">
          <option value="7" selected>&Uacute;ltimos 7 dias</option>
          <option value="15">&Uacute;ltimos 15 dias</option>
          <option value="30">&Uacute;ltimos 30 dias</option>
          <option value="45">&Uacute;ltimos 45 dias</option>
          <option value="60">&Uacute;ltimos 60 dias</option>
          <option value="90">&Uacute;ltimos 90 dias</option>
          <option value="180">&Uacute;ltimos 180 dias</option>
          <option value="365">&Uacute;ltimos 365 dias</option>
        </select>
        </font>do item 
        <select name="visitItem">
          <option value="*" selected><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">-- 
          todos --</font></option>
          <%
			Set myConn = getMdb(Application("mdbStat"))
			sqlStr = "select * from statistics;"
			set RS = myConn.Execute(sqlStr)
			for c=0 to RS.Fields.count-1
				if RS.Fields(c).Name <> "theDate" and RS.Fields(c).Name <> "cod" then%>
          <option value="<%=RS.Fields(c).Name%>"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033"><%=RS.Fields(c).Name%></font></option>
          <%end if
			next%>
        </select>
      </td>
    </tr>
    <tr bgcolor="#E9F0F8"> 
      <td width="18%">&nbsp;</td>
      <td width="45%" bgcolor="#E9F0F8">&nbsp;</td>
      <td width="37%" bgcolor="#E9F0F8" align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033"> 
        <input type="submit" name="button1" value="Detalhar">
        </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033"> 
        </font></td>
    </tr>
  </table>
</form>
<p>&nbsp;</p>

<form name="form1" method="get" action="genStatistics.asp">
  <table border="0" cellspacing="0" cellpadding="8" width="434">
    <tr> 
      <td colspan="3" bgcolor="#3A76B1"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#FFFFFF"><b>N&uacute;meros</b></font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b><font size="3"> 
        <input type="hidden" name="action" value="numbers">
        </font></b></font></td>
    </tr>
    <tr bgcolor="#E9F0F8"> 
      <td width="101"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">Data 
        desejada:</font></td>
      <td width="174"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033"> 
        <select name="visitDay">
          <option value="01"<%=selItem(1,Day(Date))%>>01</option>
          <option value="02"<%=selItem(2,Day(Date))%>>02</option>
          <option value="03"<%=selItem(3,Day(Date))%>>03</option>
          <option value="04"<%=selItem(4,Day(Date))%>>04</option>
          <option value="05"<%=selItem(5,Day(Date))%>>05</option>
          <option value="06"<%=selItem(6,Day(Date))%>>06</option>
          <option value="07"<%=selItem(7,Day(Date))%>>07</option>
          <option value="08"<%=selItem(8,Day(Date))%>>08</option>
          <option value="09"<%=selItem(9,Day(Date))%>>09</option>
          <option value="10"<%=selItem(10,Day(Date))%>>10</option>
          <option value="11"<%=selItem(11,Day(Date))%>>11</option>
          <option value="12"<%=selItem(12,Day(Date))%>>12</option>
          <option value="13"<%=selItem(13,Day(Date))%>>13</option>
          <option value="14"<%=selItem(14,Day(Date))%>>14</option>
          <option value="15"<%=selItem(15,Day(Date))%>>15</option>
          <option value="16"<%=selItem(16,Day(Date))%>>16</option>
          <option value="17"<%=selItem(17,Day(Date))%>>17</option>
          <option value="18"<%=selItem(18,Day(Date))%>>18</option>
          <option value="19"<%=selItem(19,Day(Date))%>>19</option>
          <option value="20"<%=selItem(20,Day(Date))%>>20</option>
          <option value="21"<%=selItem(21,Day(Date))%>>21</option>
          <option value="22"<%=selItem(22,Day(Date))%>>22</option>
          <option value="23"<%=selItem(23,Day(Date))%>>23</option>
          <option value="24"<%=selItem(24,Day(Date))%>>24</option>
          <option value="25"<%=selItem(25,Day(Date))%>>25</option>
          <option value="26"<%=selItem(26,Day(Date))%>>26</option>
          <option value="27"<%=selItem(27,Day(Date))%>>27</option>
          <option value="28"<%=selItem(28,Day(Date))%>>28</option>
          <option value="29"<%=selItem(29,Day(Date))%>>29</option>
          <option value="30"<%=selItem(30,Day(Date))%>>30</option>
          <option value="31"<%=selItem(31,Day(Date))%>>31</option>
        </select>
        / 
        <select name="visitMonth">
          <option value="01"<%=selItem(1,Month(Date))%>>01</option>
          <option value="02"<%=selItem(2,Month(Date))%>>02</option>
          <option value="03"<%=selItem(3,Month(Date))%>>03</option>
          <option value="04"<%=selItem(4,Month(Date))%>>04</option>
          <option value="05"<%=selItem(5,Month(Date))%>>05</option>
          <option value="06"<%=selItem(6,Month(Date))%>>06</option>
          <option value="07"<%=selItem(7,Month(Date))%>>07</option>
          <option value="08"<%=selItem(8,Month(Date))%>>08</option>
          <option value="09"<%=selItem(9,Month(Date))%>>09</option>
          <option value="10"<%=selItem(10,Month(Date))%>>10</option>
          <option value="11"<%=selItem(11,Month(Date))%>>11</option>
          <option value="12"<%=selItem(12,Month(Date))%>>12</option>
        </select>
        / 
        <select name="visitYear">
          <option value="2001" selected>2001</option>
          <option value="2002">2002</option>
        </select>
        </font></td>
      <td width="86" align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033"> 
        <input type="submit" name="button2" value="Detalhar">
        </font></td>
    </tr>
  </table>
</form>
<p>&nbsp;</p>
<form name="form1" method="get" action="genStatistics.asp">
  <table border="0" cellspacing="0" cellpadding="8" width="434">
    <tr> 
    <td bgcolor="#3A76B1"> 
        <p><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b>An&aacute;lise 
          Crescimento</b></font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b><font size="3"> 
          <font size="3"><font size="3"><b><font size="3"> 
          <input type="hidden" name="action" value="analise">
        </font></b></font></font></font></b></font></p>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#E9F0F8"> 
        <table border="0" cellspacing="0" cellpadding="5" width="100%">
          <tr> 
            <td width="25%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">Tipo:</font><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#660033"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#660033"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#660033"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033"><b><font size="3"> 
              </font></b></font></font></font></font></td>
            <td width="44%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#660033"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#660033"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#660033"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033"> 
              <select name="acompType">
                <option value="" selected>-- per&iacute;odo entre datas --</option>
                <option value="0">Hoje</option>
                <option value="1">Ontem</option>
                <option value="3">&Uacute;ltimos 3 dias</option>
                <option value="7">&Uacute;ltimos 7 dias</option>
                <option value="15">&Uacute;ltimos 15 dias</option>
                <option value="30">&Uacute;ltimos 30 dias</option>
                <option value="45">&Uacute;ltimos 45 dias</option>
                <option value="60">&Uacute;ltimos 60 dias</option>
                <option value="90">&Uacute;ltimos 90 dias</option>
                <option value="180">&Uacute;ltimos 180 dias</option>
                <option value="365">&Uacute;ltimos 365 dias</option>
              </select>
              </font></font></font></font></td>
            <td align="right" width="31%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#660033"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#660033"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#660033"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#660033"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033"> 
              </font></font></font></font></font></td>
          </tr>
        </table>
        <br>
      </td>
  </tr>
  <tr> 
    <td bgcolor="#E9F0F8"> 
        <table border="0" cellspacing="0" cellpadding="5" width="100%">
          <tr> 
          <td><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#660033"><font size="2">Data 
            inicial:</font></font></td>
            <td> 
              <select name="visitDay">
                <option value="01"<%=selItem(1,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">01</font></option>
                <option value="02"<%=selItem(2,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">02</font></option>
                <option value="03"<%=selItem(3,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">03</font></option>
                <option value="04"<%=selItem(4,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">04</font></option>
                <option value="05"<%=selItem(5,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">05</font></option>
                <option value="06"<%=selItem(6,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">06</font></option>
                <option value="07"<%=selItem(7,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">07</font></option>
                <option value="08"<%=selItem(8,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">08</font></option>
                <option value="09"<%=selItem(9,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">09</font></option>
                <option value="10"<%=selItem(10,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">10</font></option>
                <option value="11"<%=selItem(11,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">11</font></option>
                <option value="12"<%=selItem(12,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">12</font></option>
                <option value="13"<%=selItem(13,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">13</font></option>
                <option value="14"<%=selItem(14,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">14</font></option>
                <option value="15"<%=selItem(15,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">15</font></option>
                <option value="16"<%=selItem(16,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">16</font></option>
                <option value="17"<%=selItem(17,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">17</font></option>
                <option value="18"<%=selItem(18,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">18</font></option>
                <option value="19"<%=selItem(19,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">19</font></option>
                <option value="20"<%=selItem(20,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">20</font></option>
                <option value="21"<%=selItem(21,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">21</font></option>
                <option value="22"<%=selItem(22,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">22</font></option>
                <option value="23"<%=selItem(23,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">23</font></option>
                <option value="24"<%=selItem(24,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">24</font></option>
                <option value="25"<%=selItem(25,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">25</font></option>
                <option value="26"<%=selItem(26,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">26</font></option>
                <option value="27"<%=selItem(27,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">27</font></option>
                <option value="28"<%=selItem(28,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">28</font></option>
                <option value="29"<%=selItem(29,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">29</font></option>
                <option value="30"<%=selItem(30,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">30</font></option>
                <option value="31"<%=selItem(31,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">31</font></option>
              </select>
              / 
              <select name="visitMonth">
                <option value="01"<%=selItem(1,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">01</font></option>
                <option value="02"<%=selItem(2,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">02</font></option>
                <option value="03"<%=selItem(3,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">03</font></option>
                <option value="04"<%=selItem(4,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">04</font></option>
                <option value="05"<%=selItem(5,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">05</font></option>
                <option value="06"<%=selItem(6,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">06</font></option>
                <option value="07"<%=selItem(7,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">07</font></option>
                <option value="08"<%=selItem(8,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">08</font></option>
                <option value="09"<%=selItem(9,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">09</font></option>
                <option value="10"<%=selItem(10,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">10</font></option>
                <option value="11"<%=selItem(11,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">11</font></option>
                <option value="12"<%=selItem(12,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">12</font></option>
              </select>/
              <select name="visitYear">
                <option value="2001" selected><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">2001</font></option>
                <option value="2002"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033">2002</font></option>
              </select>
            </td>
            <td><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#660033"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#660033"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#660033"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033"><b><font size="3"> 
              <font size="1"> 
              <input type="checkbox" name="animate" value="true">
              <font size="2">animar</font></font></font></b></font></font></font></font></td>
        </tr>
        <tr> 
          <td><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#660033"><font size="2">Data 
            final:</font></font></td>
            <td> 
              <select name="visitDay2">
                <option value="01"<%=selItem(1,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">01</font></option>
                <option value="02"<%=selItem(2,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">02</font></option>
                <option value="03"<%=selItem(3,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">03</font></option>
                <option value="04"<%=selItem(4,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">04</font></option>
                <option value="05"<%=selItem(5,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">05</font></option>
                <option value="06"<%=selItem(6,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">06</font></option>
                <option value="07"<%=selItem(7,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">07</font></option>
                <option value="08"<%=selItem(8,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">08</font></option>
                <option value="09"<%=selItem(9,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">09</font></option>
                <option value="10"<%=selItem(10,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">10</font></option>
                <option value="11"<%=selItem(11,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">11</font></option>
                <option value="12"<%=selItem(12,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">12</font></option>
                <option value="13"<%=selItem(13,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">13</font></option>
                <option value="14"<%=selItem(14,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">14</font></option>
                <option value="15"<%=selItem(15,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">15</font></option>
                <option value="16"<%=selItem(16,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">16</font></option>
                <option value="17"<%=selItem(17,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">17</font></option>
                <option value="18"<%=selItem(18,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">18</font></option>
                <option value="19"<%=selItem(19,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">19</font></option>
                <option value="20"<%=selItem(20,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">20</font></option>
                <option value="21"<%=selItem(21,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">21</font></option>
                <option value="22"<%=selItem(22,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">22</font></option>
                <option value="23"<%=selItem(23,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">23</font></option>
                <option value="24"<%=selItem(24,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">24</font></option>
                <option value="25"<%=selItem(25,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">25</font></option>
                <option value="26"<%=selItem(26,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">26</font></option>
                <option value="27"<%=selItem(27,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">27</font></option>
                <option value="28"<%=selItem(28,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">28</font></option>
                <option value="29"<%=selItem(29,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">29</font></option>
                <option value="30"<%=selItem(30,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">30</font></option>
                <option value="31"<%=selItem(31,Day(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">31</font></option>
              </select>
              / 
              <select name="visitMonth2">
                <option value="01"<%=selItem(1,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">01</font></option>
                <option value="02"<%=selItem(2,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">02</font></option>
                <option value="03"<%=selItem(3,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">03</font></option>
                <option value="04"<%=selItem(4,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">04</font></option>
                <option value="05"<%=selItem(5,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">05</font></option>
                <option value="06"<%=selItem(6,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">06</font></option>
                <option value="07"<%=selItem(7,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">07</font></option>
                <option value="08"<%=selItem(8,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">08</font></option>
                <option value="09"<%=selItem(9,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">09</font></option>
                <option value="10"<%=selItem(10,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">10</font></option>
                <option value="11"<%=selItem(11,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">11</font></option>
                <option value="12"<%=selItem(12,Month(Date))%>><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">12</font></option>
              </select>
              / 
              <select name="visitYear2">
                <option value="2001" selected><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">2001</font></option>
                <option value="2002"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#660033">2002</font></option>
              </select>
            </td>
          <td align="right"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#660033"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033"> 
            <input type="submit" name="button3" value="Detalhar">
            </font></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</form>
<p align="left"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#660033"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#660033"><br>
  </font></font></p>
</body>
</html>
<%
RS.Close
myConn.close
set RS = Nothing
set myConn = Nothing

' by Flávio Stutz (flaviostutz@hotmail.com)
%>
