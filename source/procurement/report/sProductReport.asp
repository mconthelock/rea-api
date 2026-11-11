<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file ="../List.asp" -->
<!-- #include file="../fcommon.asp" -->

<%
	' Author: worakrit yodthonglert
	' Version: 1.01
%>
<%
	Dim i
	Dim  Rs,sSQL,sSQL2,RS2
	Dim sScriptPath,sexptype,sPoscode,sProductcode,sProdcode,sVendDesc
	sProductcode=Request("sProdcode")
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")

%>
<HTML>
<HEAD>
<TITLE>Summary By Total</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=">
<LINK REL="stylesheet" HREF="../styles.css">
<SCRIPT LANGUAGE="JavaScript">
<!--
function Form_Load(){
	document.frmName.sProductcode.focus();
}

function OpenNewWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
-->
</SCRIPT>
</HEAD>
<BODY BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		

  <table width="98%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td align="left" valign="top" width="1%" height="2"><img src="../images/ccb.gif" width="10" height="10"></td>
      <td width="99%" height="2" align="center" valign="top"><img src="../images/spacer.gif" width="8" height="8"></td>
    </tr>
    <tr> 
      <td width="1%">&nbsp;</td>
      <td width="99%" height="30" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="+1" color="#666666"><b><b><b><b><b><b><b><font color="#BF00BF">Report 
        Manager</font></b></b></b></b></b></b></b></font></td>
    </tr>
    <tr> 
      <td width="1%"><img src="../images/spacer.gif" width="1" height="1"></td>
      <td width="99%" align="center" valign="top" height="0"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="98%" height="2" align="left" valign="top" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><img src="../images/y01.jpg" width="12" height="12"></b></font></td>
          </tr>
          <tr> 
            <td width="98%" align="left" valign="top" height="0"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr align="left" valign="top" bgcolor="#FFD553"> 
                  <td colspan="2" align="center">&nbsp;</td>
                </tr>
                <tr align="center" valign="top"> 
                  <td colspan="2" bgcolor="#FFD553"> 
                    
                  <table border=0 width=90% align="Center" cellspacing="1" cellpadding="0">
                    <tr bgcolor="#CC00CC"> 
                      <td colspan="2" height="20"> 
                        <div align="center"><b><font size="2"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">Summary 
                          Report Overview By Invoice</font></font></b></div>
                      </td>
                    </tr>
                    
                    <tr bgcolor="#EEEEEE"> 
                      <td width="31%" height="10"> 
                        <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp;</font></div>
                      </td>
                      <td width="69%" height="10">&nbsp; </td>
                    </tr>
                    <Form Action="vmProductReport.asp" method="post" name=frmName>
                      <tr bgcolor="#EEEEEE"> 
                        <td width="31%" height="10" align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Product 
                          : </font></td>
                        <td width="69%" height="10"> 
                          <input type="text" name="sProductcode" value="<%=sProductcode%>">
                          <!--:  | 
						  <a href="Searchproduct.asp">Search 
                          Product </a>
						  
						  <a href="JavaScript:OpenNewWindow('Searchproduct.asp','SearchProduct','width=500,height=500')">Search 
                          Product </a> |--></td>
                      </tr>
                      <tr bgcolor="#EEEEEE"> 
                        <td width="31%" height="10"> 
                          <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">From 
                            :&nbsp;</font></div>
                        </td>
                        <td width="69%" height="10"> 
                          <select name="mFrom_month">
                            <% Create_MonthList(month(date-31))%>
                          </select>
                          <select name="mFrom_Year">
                            <% Create_YearList(Year(date))%>
                          </select>
                        </td>
                      </tr>
                      <tr bgcolor="#EEEEEE"> 
                        <td width="31%" height="10"> 
                          <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">To 
                            :&nbsp;</font></div>
                        </td>
                        <td width="69%" height="10"> 
                          <select name="mTo_Month">
                            <% Create_MonthList(month(date))%>
                          </select>
                          <select name="mTo_Year">
                            <% Create_YearList(Year(date))%>
                          </select>
                        </td>
                      </tr>
                      <tr bgcolor="#EEEEEE"> 
                        <td width="31%" height="10">&nbsp;</td>
                        <td width="69%" height="10">&nbsp; </td>
                      </tr>
                      <tr bgcolor="#EEEEEE"> 
                        <td width="31%" height="10"> 
                          <div align="right"><img src="../images/dot.gif"></div>
                        </td>
                        <td width="69%" height="10"> 
                          <input type="submit" name="Submit" value="Submit">
                          <input type="reset" name="Reset" value="Reset">
                        </td>
                      </tr>
                    </Form>
                  </table>
                  </td>
                </tr>
                <tr align="left" valign="top"> 
                  <td width="98%" bgcolor="#FFD553">&nbsp;</td>
                  <td width="2%" align="right" bgcolor="#FFD553" valign="bottom"><img src="../images/y04.jpg" width="12" height="12"></td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
        <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"></font></td>
    </tr>
  </table>
  <table width="98%" border="0" cellspacing="0" cellpadding="0">
    <tr align="center"> 
      <td width="1%"> <img src="../report/images/spacer.gif" width="1" height="1"><br>
      </td>
      <td width="99%"> 
         <!--#include file="../inc/copyRight.inc"--></td>
    </tr>
  </table>

</BODY>
</HTML>
<%
Sub Create_Yearlist(searchyear)
Dim syear
	For i=2000 to searchyear
		if i=searchyear then
			Response.Write "<option value= '"& i & "' selected>"& i &"</option>"& vbCrLf
		else
			Response.Write "<option value= '"& i & "'>"& i &"</option>"& vbCrLf
		End if
	Next
End Sub
%>