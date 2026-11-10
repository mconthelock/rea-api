<%@ Language=VBScript %>
<!--#include virtual="/sys/inc/system.inc"-->
<!--#include virtual="/form/inc/system.inc"-->
<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/form/inc/connUtil.inc"-->
<!--#include virtual="/form/inc/authen.inc"-->
<!--#include virtual="/sys/inc/userUtil.inc"-->
<!--#include virtual="/form/inc/formUtil.inc"-->
<!--#include virtual="/sys/inc/date.inc"-->
<!--#include virtual="/sys/inc/date2.inc"-->
<!--#include virtual="/sys/inc/misc.inc"-->
<!--#include virtual="/sys/inc/report.inc"-->
<!--#include virtual="/sys/inc/reportData.inc"-->
<!--#include virtual="/sys/inc/search.inc"-->
<!--#include virtual="/sys/inc/oraUtil.inc"-->
<!--#include virtual="/sys/inc/pageUtil.inc"-->


<html>
<head>
<title>FIN Report</title>
<META http-equiv="Content-Type" content="text/html; charset=windows-874">
<link href="/sys/css/common.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
	dim rs
	dim sql
	dim chk,count
	dim reqDate
	count=1
	set rs=CreateObject("Adodb.recordset")

	if  ((request("txtAssNo")="") and (request("txtAssName")="") and (request("txtLocName") = "") and (request("txtPoNo") = "")  and (request("txtPrNo") = "") and (request("selGrpCode") = "") and (request("selLocCode")="")) then
		response.write "Please select condition"
		response.end
	end if
	sql = "select famst.*, grpMst.grpdesc, locName from fxa_faMst famst, fxa_locMst  locMst, fxa_grpMst  grpMst" _
	& " where famst.locCode = locMst.locCode and famst.grpCode = grpMst.grpCode"

	if request("txtAssNo") <> "" then
		sql=sql & " and assetno like '%" & request("txtAssNo") & "%'"
	end if
	if request("txtAssName") <> "" then
		sql=sql & " and Upper(assDesc) like '%" & UCase(request("txtAssName")) & "%'"
	end if
	if request("txtLocName") <> "" then
		sql=sql & " and Upper(locName) like '%" & UCase(request("txtLocName")) & "%'"
	end if
	if request("txtPoNo") <> "" then
		sql=sql & " and pono like '%" & request("txtPoNo") & "%'"
	end if
	if request("txtPrNo") <> "" then
		sql=sql & " and prno like '%" & request("txtPrNo") & "%'"
	end if
	if request("selGrpCode") <> "" then
		sql=sql & " and grpMst.GrpCode = '" & request("selGrpCode") & "'"	
	end if 
	
	if request("selLocCode") <> "" then
		sql=sql & " and locMst.LocCode = '" & request("selLocCode") & "'"	
	end if 
	
	if (request("selReqYear") <> "" and request("selReqMonth") <> "" and request("selReqDate") <> "" ) or  (request("selReqYear2") <> "" and request("selReqMonth2") <> "" and request ("selReqDate2") <> "") then
		if (request("selReqYear") & request("selReqMonth") & request("selReqDate")) = (request("selReqYear2") & request("selReqMonth2") & request ("selReqDate2")) then
			sql=sql & " and docdate = to_date('" & request("selReqDate") & "/" & request("selReqMonth") & "/" & request("selReqYear") &"','dd/mm/yyyy')"
			reqDate = request("selReqDate") & "/" & request("selReqMonth") & "/" & request("selReqYear")
		else
			if request("selReqDate") <> ""   and  request("selReqMonth") <> ""  and   request("selReqYear") <> "" then
				sql=sql & " and docdate >= to_date('" & request("selReqDate") & "/" & request("selReqMonth") & "/" & request("selReqYear") &"','dd/mm/yyyy')"
				reqDate = request("selReqDate") & "/" & request("selReqMonth") & "/" & request("selReqYear")					
			end if 
			
			if request ("selReqDate2") <> ""  and    request("selReqMonth2") <> ""   and  request("selReqYear2") <> "" then
				sql=sql & " and docdate <= to_date('" & request("selReqDate2") & "/" & request("selReqMonth2") & "/" & request("selReqYear2") &"','dd/mm/yyyy')"
				reqDate = reqDate &" and " & request("selReqDate2") & "/" & request("selReqMonth2") & "/" & request("selReqYear2")										
			end if 		
		end if
	end if
	
	sql=sql & " order by assDesc"
	'response.write sql
	'response.end
	rs.Open sql, connOra

	if not rs.EoF then	
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr valign="bottom"> 
    <td align="right"><img src="/<%=SYSTEM_FOLDER %>/imgs/border/report/tlc.gif" width="17" height="18"></td>
    <td background="/<%=SYSTEM_FOLDER %>/imgs/border/report/t.gif" width="100%"><img src="/<%=SYSTEM_FOLDER %>/imgs/border/report/t.gif" width="17" height="18"></td>
    <td align="left"><img src="/<%=SYSTEM_FOLDER %>/imgs/border/report/trc.gif" width="17" height="18"></td>
  </tr>
  <tr>
    <td align="right" background="/<%=SYSTEM_FOLDER %>/imgs/border/report/l.gif">&nbsp;</td>
    <td align="center" background="/<%=SYSTEM_FOLDER %>/imgs/border/report/bg.gif"><b><font color="#000000">Search Result</font></b></td>
    <td align="left" background="/<%=SYSTEM_FOLDER %>/imgs/border/report/r.gif">&nbsp;</td>
  </tr>
  <tr> 
    <td align="right" background="/<%=SYSTEM_FOLDER %>/imgs/border/report/l.gif"><img src="/<%=SYSTEM_FOLDER %>/imgs/border/report/l.gif" width="17" height="18"></td>
    <td bgcolor="#0000FF"> 
      <table width="100%" border="0" cellspacing="1" cellpadding="2">
        <tr align="center"> 
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">No.</font></b></td>
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">Group No.</font></b></td>
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">Description</font></b></td>
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">L/C</font></b></td>
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">Description L/C</font></b></td>
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">Asset No.</font></b></td>    
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">Description</font></b></td>	
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">Doc. Date</font></b></td>	
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">INITIAL VALUE</font></b></td>
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">INV.NO</font></b></td>
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">MODEL NO</font></b></td>
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">S/N</font></b></td>
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">PO No.</font></b></td>
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">UNIT</font></b></td>    
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">SUPPLIER NAME</font></b></td>    
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">PR NO</font></b></td>    	
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">REQUEST BY</font></b></td>    		
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">PICTURE</font></b></td>    			
  </tr>
  <% do while not rs.EOF %>
  <tr  bgcolor="#FFFFFF"> 
    <td nowrap align="center"><font color="#000000"><%=count%></font></td>
    <td nowrap align="center"><font color="#000000"><%=rs("grpCode")%>&nbsp;</font></td>
    <td nowrap align="center"><font color="#000000"><%=rs("grpdesc")%>&nbsp;</font></td>    
    <td nowrap align="center"><font color="#000000"><%=rs("locCode")%>&nbsp;</font></td>
    <td nowrap align="center"><font color="#000000"><%=rs("locName")%>&nbsp;</font></td>	
    <td nowrap align="center"><font color="#000000"><%=rs("assetno")%>&nbsp;</font></td>	
    <td nowrap align="center"><font color="#000000"><%=rs("assDesc")%>&nbsp;</font></td>		
    <td nowrap align="center"><font color="#000000"><%=rs("docdate")%>&nbsp;</font></td>			
    <td nowrap align="center"><font color="#000000"><%=rs("initialvalue")%>&nbsp;</font></td>			
    <td nowrap align="center"><font color="#000000"><%=rs("invno")%>&nbsp;</font></td>			
   <td nowrap align="center"><font color="#000000"><%=rs("modelno")%>&nbsp;</font></td>			
    <td nowrap align="center"><font color="#000000"><%=rs("snno")%>&nbsp;</font></td>			
    <td nowrap align="center"><font color="#000000"><%=rs("pono")%>&nbsp;</font></td>			
    <td nowrap align="center"><font color="#000000"><%=rs("unit")%>&nbsp;</font></td>			
    <td nowrap align="center"><font color="#000000"><%=rs("supname")%>&nbsp;</font></td>			
    <td nowrap align="center"><font color="#000000"><%=rs("prno")%>&nbsp;</font></td>			
    <td nowrap align="center"><font color="#000000"><%=rs("reqby")%>&nbsp;</font></td>			
    <td nowrap align="center"><font color="#000000">&nbsp;</font></td>					
  </tr>
  <%
				rs.MoveNext
				count=count+1
  		loop
%>
      </table>
    </td>
    <td align="left" background="/<%=SYSTEM_FOLDER %>/imgs/border/report/r.gif"><img src="/<%=SYSTEM_FOLDER %>/imgs/border/report/r.gif" width="17" height="18"></td>
  </tr>
  <tr valign="top"> 
    <td align="right" background="/<%=SYSTEM_FOLDER%>/imgs/border/report/l.gif"><img src="/<%=SYSTEM_FOLDER%>/imgs/border/report/l.gif" width="17" height="18"></td>
    <td background="/<%=SYSTEM_FOLDER %>/imgs/border/report/bg.gif" align="center">
		<input type="button" name="btnExpExc" value="Export Excel" onclick="javascript:window.location.href='finExpExcelSchFX.asp?sqlOra=<%=server.urlencode(sql)%>'"> 
		</td>
    <td align="left" background="/<%=SYSTEM_FOLDER %>/imgs/border/report/r.gif"><img src="/<%=SYSTEM_FOLDER%>/imgs/border/report/r.gif" width="17" height="18"></td>
  </tr>
  <tr valign="top"> 
    <td align="right"><img src="/<%=SYSTEM_FOLDER %>/imgs/border/report/blc.gif" width="17" height="18"></td>
    <td background="/<%=SYSTEM_FOLDER %>/imgs/border/report/b.gif"><img src="/<%=SYSTEM_FOLDER%>/imgs/border/report/b.gif" width="17" height="18"></td>
    <td align="left"><img src="/<%=SYSTEM_FOLDER %>/imgs/border/report/brc.gif" width="17" height="18"></td>
  </tr>
</table>
<%	
	else
		response.write "data not found"
	end if	
%>
</body>
</html>
