<script language=JavaScript>
//opens an window for confirmation, info etc
function openBox(htmlFile,wName,wWidth,wHeight,haveScroll,posX,posY) {
	if (wName!="" && wWidth>0 && wHeight>0) return window.open(htmlFile,wName,"toolbar=no,scrollbars="+(haveScroll==true?"yes":"no")+",directories=no,status=no,menubar=no,resizable=no,screenX=" + posX + ",screenY=" + posY + "top=" + posY + ",left=" + posX + ",width=" + wWidth + ",height=" + wHeight);
}
</script>


<%
Function getMdb(mdbFile)
	Set myConn = Server.CreateObject("ADODB.Connection")
	myConn.open "DBQ=" & Server.MapPath(mdbFile) & ";Driver={Microsoft Access Driver (*.mdb)}"
	set getMdb = myConn
End Function


'the htmlTitle will be as a commentary and between '[' and ']', inside the html file (has to be among the 100 first chars
Function getTitle(htmlFile)
	set FS = Server.CreateObject("scripting.FileSystemObject") 
	set htmlStream = FS.OpenTextFile(Server.MapPath(htmlFile))
	strHtml = htmlStream.Read(100)
	beginPos = inStr(1, strHtml, "<!--[")+5
	if beginPos > 5 then getTitle = Mid(strHtml, beginPos, inStr(beginPos, strHtml, "]") - beginPos)
End Function

'just write to the user
sub write(str)
	Response.write(str)
end sub

'makes the field and the value sql string part from an html form (each pass does one field)
'fieldStr - field name
'valueStr - field value
'valueStr - value of the form object
'sqlFields and sqlValues - receives by addition the formatted sql string (eg. sqlFields="name,age" - sqlValues="'Flávio',19")
'quoted - quote the fieldValue or don't
sub separateField(fieldStr,valueStr,byref sqlFields,byref sqlValues,quoted)
	if not quoted then valueStr = ReplaceStr(valueStr,",",".") 'put the numbers in the 321.36 format (instead of 321,36)
	if valueStr <> "" then
		if sqlFields <> "" then 
			sqlFields = sqlFields & ", " & fieldStr
			if quoted then
				sqlValues = sqlValues & ", '" & valueStr & "'"
			else
				sqlValues = sqlValues & ", " & valueStr
			end if
		else
			sqlFields = fieldStr
			if quoted then
				sqlValues = "'" & valueStr & "'"
			else
				sqlValues = valueStr
			end if
		end if
	end if
end sub

'commonly used functions

Function replaceStr(whatString, whichString, toString)
	returnStr = whatString
	if whatString <> "" then
		oldLoc = 1
		returnStr = ""
		newLoc = InStr(1, whatString, whichString)
		while newLoc > 0
			returnStr = returnStr + Mid(whatString, oldLoc, (newLoc-oldLoc)) + toString
			oldLoc = newLoc + Len(whichString)
			newLoc = InStr(oldLoc, whatString, whichString)
		wend
		returnStr = returnStr + Mid(whatString, oldLoc)
	end if
	replaceStr = returnStr
end Function

'change to currency format
Function toCurStr(whichValue)
	tempDecStr = left(CStr(CInt((whichValue - Int(whichValue)) * 100)), 2)
	tempIntStr = CStr(Int(whichValue))
	intLen = Len(tempIntStr)
	tempIntStr3 = ""
	
'integer part
for i=1 to Int(intLen/3)+1
	if intLen-(i*3) >= 0 then
		tempIntStr2 = Mid(tempIntStr, intLen-(i*3)+1, 3)
	else
		tempIntStr2 = Mid(tempIntStr, 1, intLen-((i-1)*3))
	end if

	if tempIntStr3 <> "" then
		if tempIntStr2 <> "" then	
			tempIntStr3 = tempIntStr2 + "." + tempIntStr3
		else
			tempIntStr3 = tempIntStr2 + tempIntStr3
		end if
	else
		tempIntStr3 = tempIntStr2
	end if
next

'dec part
	if len(tempDecStr) = 0 then
		toCurStr = tempIntStr3 + ",00"
	elseif len(tempDecStr) = 1 then
		toCurStr = tempIntStr3 + "," + tempDecStr + "0"
	elseif len(tempDecStr) = 2 then
		toCurStr = tempIntStr3 + "," + tempDecStr
	end if
End Function

'debugging stuff
Function alert(msg)
	Response.write("<script>alert('|" + CStr(msg) + "|')</script>")
End Function


'generate random numbers between the interval upperNbr to lowerNbr
Randomize(Second(Time))
function getRnd(lowerNbr, upperNbr)
	getRnd = Int((upperNbr - lowerNbr + 1) * Rnd + lowerNbr)
end function


Function showBanner(bannerSession)
	Select Case bannerSession
		Case "default"
			nbrBanners = 3
			Dim banners(3)
			banners(0) = "imagens/enviar_cinza.gif"
			banners(1) = "imagens/banner1.gif"
			banners(2) = "imagens/banner2.gif"
			if Session("bannerToShow") = "" then
				toShow = getRnd(0,nbrBanners-1)
			else
				toShow = Session("bannerToShow")
				toShow = toShow + 1
				if toShow => nbrBanners then toShow = 0
			end if
			Session("bannerToShow") = toShow
			showBanner = banners(toShow)
		case else

	End Select
End Function


Sub addStatistics(item,mdbFile)
On Error Resume Next
	Set myConn = getMdb(mdbFile)
	if Application("lastUpdateDay") <> Day(Date) then
		sqlStr = "select max(cod) as nextCod from statistics;"
		set RS = myConn.Execute(sqlStr)
		nextCod = RS("nextCod")+1
		sqlStr = "insert into statistics (cod,theDate) values (" & nextCod & ",'" & Date & "');"
		myConn.Execute(sqlStr)
	end if
	
	sqlStr = "update statistics set " & item & "=" & item & "+1 where theDate='" & Date & "';"
	Response.write(sqlStr)
	myConn.Execute(sqlStr)
	Application("lastUpdateDay") = Day(Date)
	myConn.close
	set myConn = Nothing
End Sub

function swapDayMonth(theDate) 'swaps - dd/mm to mm/dd vice-versa
	firstStr = mid(theDate,1,instr(1,theDate,"/")-1)
	secondPos = instr(1,theDate,"/")+1
	secondStr = mid(theDate,secondPos,instr(secondPos,theDate,"/") - secondPos)
	swapDayMonth = secondStr & "/" & firstStr & "/" & Year(theDate)
end function

'returns the string " selected" for combo boxes if optValue = varValue
function selItem(optValue,varValue)
	if (optValue=CInt(varValue)) then selItem = " selected"
end function
%>
