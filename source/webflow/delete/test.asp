<%Option Explicit%>

<!--#include file="inc/adovbs.inc"-->
<!--#include file="inc/db.inc"-->
<!--#include file="inc/connUtil.inc"-->
<!--#include file="inc/flowUtil.inc"-->
<!--#include file="inc/date.inc"-->
<!--#include file="inc/flow.inc"-->
<!--#include file="inc/hisUtil.inc"-->

<%
debug = true
%>

<form action="<%=request.serverVariables("SCRIPT_NAME")%>">
Select form:
<select>
</select><br>
<input type="submit" name="a" value="Create new"><input type="submit" name="b" value="Refresh"><input type="submit" name="c" value="Clear all"><input type="button" name="d" value="Back" onClick="javascript: window.location.href = 'index.asp'"><br>
<input type="text" name="no" value="1"><br>
<input type="text" name="orgNo" value="50603"><br>
<input type="text" name="y" value="02"><br>
<input type="text" name="requester" value="92043"><br>
<input type="text" name="inputer" value="01025"><br>
<input type="text" name="remark" value="Test"><br>
</form>
<%
Dim rsOra, sqlOra

set rsOra = createObject("adodb.recordset")

connOra.beginTrans

if request.queryString("a") <> "" then
createFlow connOra, request.queryString("no"), request.queryString("orgNo"), request.queryString("y"), request.queryString("requester"),request.queryString("inputer"), request.queryString("remark")
end if

if request.queryString("c") <> "" then
	sqlOra = "delete from cntPerson"
	rsOra.open sqlOra, connOra
	
	sqlOra = "delete from eMail"
	rsOra.open sqlOra, connOra
	
	sqlOra = "delete from calendar"
	rsOra.open sqlOra, connOra
	
	sqlOra = "delete from flow"
	rsOra.open sqlOra, connOra

	sqlOra = "delete from form"
	rsOra.open sqlOra, connOra
end if

if request.queryString("d") <> "" then
	approve connOra, request.queryString("no"), request.queryString("orgNo"), request.queryString("y"), request.queryString("y2"), request.queryString("runNo"), request.queryString("apv"), request.queryString("remark")
end if

if request.queryString("e") <> "" then
	reject connOra, request.queryString("no"), request.queryString("orgNo"), request.queryString("y"), request.queryString("y2"), request.queryString("runNo"), request.queryString("apv"), request.queryString("remark")
end if

sqlOra = "select to_char(nRunNo) as nRunNo from form where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' and cYear2 = '" & year(date) & "' order by nRunNo desc"
rsOra.open sqlOra, connOra

do while not rsOra.EOF
response.write rsOra("nRunNo")' & "<br>"
showFlow connOra, request.queryString("no"), request.queryString("orgNo"), request.queryString("y"), year(date), rsOra("nRunNo"), "01026", "01025"
rsOra.moveNext
loop

connOra.commitTrans

historyAdd request.queryString("SCRIPT_NAME")
%>