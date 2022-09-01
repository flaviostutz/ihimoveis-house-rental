<%response.buffer=true%>
<!--#include file="common.inc" -->
<%
'item = Request.QueryString("item") - db field name
'Request.QueryString("URL") - url to redirect to 

On Error Resume Next

addStatistics Request.QueryString("item"),Application("mdbStat")

if Request.QueryString("URL") <> "" then
	Response.redirect(Request.QueryString("URL"))
end if

Response.flush
%>
<%
' by Flávio Stutz (flaviostutz@hotmail.com)
%>