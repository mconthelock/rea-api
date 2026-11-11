<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
	
   Dim  Rs,sSQL,sSQL2,RS2,RS3,RS4,RS5
	Dim sPrno,sExptype,dPrdate,nPrcount,i,chkreceive,nPrcount2
	Dim Color,noPr,dlastappdate,Maxstep,dinvdate,Empno,chk,chkrec
	Dim Prno(),Step(),Approve(),Appdate()
	Dim Prno2(),Step2(),Approve2(),Appdate2()
	Color="#F1B3A7"
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	Set Rs5 = Server.CreateObject("ADODB.RecordSet")
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	dlastappdate=date-14

%>
<html>
<head>
<title>Work Flow Manager</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<link rel="stylesheet" href="../styles.css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#6666CC" vlink="#6666CC" alink="#6666CC" text="#6666CC">
<table width="99%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="left" valign="top" width="1%" height="2"><img src="../datamanager/images/ccb.gif" width="10" height="10"></td>
    <td width="38%" height="2" align="center" valign="top"><img src="../security/images/spacer.gif" width="8" height="8"></td>
    <td width="61%" height="2" align="center"><img src="../datamanager/images/spacer.gif" width="8" height="8"></td>
  </tr>
  <tr>
    <td width="1%">&nbsp;</td>
    <td width="38%" align="left" valign="top" height="30"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#660066" size="+1">WorkFlow 
      Manager</font></b></font></b></font></b></font></b></font></b></font> </td>
    <td width="61%" align="center" valign="top">&nbsp;</td>
  </tr>
  <tr> 
    <td width="1%"><img src="../security/images/spacer.gif" width="1" height="1"></td>
    <td width="38%" align="left" valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="2" align="left" valign="top" bgcolor="#FFD553" colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><img src="../security/images/y01.jpg" width="12" height="12"></b></font></td>
          <td align="right" valign="top" width="3%">&nbsp;</td>
        </tr>
        <tr> 
          <td align="center" valign="top" height="226" colspan="2"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr align="center" valign="top" bgcolor="#FFD553"> 
                <td colspan="2"><img src="../images/okpic003.gif" width="193" height="180"></td>
              </tr>
              <tr align="left" valign="top" bgcolor="#FFD553"> 
                <td width="98%">&nbsp;</td>
                <td width="2%" align="right">&nbsp;</td>
              </tr>
            </table>
            <br>
          </td>
          <td width="3%" align="right" height="226"><img src="../security/images/spacer.gif" width="1" height="1"></td>
        </tr>
        <%	if Mid(Session("sAuthor"),13,3) = "AED" Or Mid(Session("sAuthor"),13,3) = "VIW" Then%> 
        <tr> 
          <td width="6%" bgcolor="#FFD553" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><img src="../security/images/y01.jpg" width="12" height="12"></b></font></td>
          <td width="91%" align="right" bgcolor="#FFD553" valign="top"><img src="../datamanager/images/y02.jpg" width="13" height="12"></td>
          <td width="3%" align="right" valign="bottom">&nbsp;</td>
        </tr>
        <tr> 
          <td width="6%" bgcolor="#FFD553">&nbsp;</td>
          <td width="91%" bgcolor="#FFD553"> 
            <form method="post" action="../workflow/searchpr.asp">
              <b><font color="#0066FF">Search P/R By Person :</font></b><br>
              <input type="text" name="ssearchpr" size="13" maxlength="12">
              <input type ="Submit" name="submit" value="Search">
              <br>
              <select name="status">
			     <option value="0" selected></option>
                <option value="1" >Key</option>
                <option value="2">Requester</option>
                <option value="3">Approver</option>
              </select>
              <br>
              <font color="#3399FF">Search Between</font> <br>
              <select name="Searchdate_day">
                <%Create_daylist(day(date-7)) %> 
              </select>
              <select name="Searchdate_month">
                <% Create_MonthList(month(date-7))%> 
              </select>
              <select name="searchdate_Year">
                <% Create_YearList(Year(date-7))%> 
              </select>
              <br>
              <font color="#3399FF">And</font><br>
              <select name="Searchtodate_day">
                <%Create_daylist(day(date)) %> 
              </select>
              <select name="Searchtodate_month">
                <% Create_MonthList(month(date))%> 
              </select>
              <select name="searchtodate_Year">
                <% Create_YearList(Year(date))%> 
              </select>
            </form>
          </td>
          <td width="3%" align="right" valign="bottom">&nbsp;</td>
        </tr>

        <tr> 
          <td width="6%" bgcolor="#FFD553" valign="Bottom"><img src="../datamanager/images/y03.jpg" width="12" height="12"></td>
          <td width="91%" bgcolor="#FFD553" align="right" valign="Bottom"><img src="../datamanager/images/y04.jpg" width="12" height="12"></td>
          <td width="3%" align="right" valign="bottom"><img src="../security/images/spacer.gif" width="1" height="1"></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="3" height="16" valign="bottom">&nbsp;</td>
        </tr>

        <tr> 
          <td width="6%" bgcolor="#FFD553" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><img src="../security/images/y01.jpg" width="12" height="12"></b></font></td>
          <td width="91%" align="right" bgcolor="#FFD553" valign="top"><img src="../datamanager/images/y02.jpg" width="13" height="12"></td>
          <td width="3%" align="right" valign="bottom">&nbsp;</td>
        </tr>
        <tr> 
          <td width="6%" bgcolor="#FFD553">&nbsp;</td>
          <td width="91%" bgcolor="#FFD553"> 
            <form method="post" action="../workflow/searchprpo.asp">
              <b><font color="#0066FF">Search P/R For P/O:</font></b><br>
              <input type="text" name="ssearchpr" size="13" maxlength="12">
              <input type ="Submit" name="submit" value="Search">
              <br>

              <font color="#3399FF">Search Between</font> <br>
              <select name="Searchdate_day">
                <%Create_daylist(day(date-7)) %> 
              </select>
              <select name="Searchdate_month">
                <% Create_MonthList(month(date-7))%> 
              </select>
              <select name="searchdate_Year">
                <% Create_YearList(Year(date-7))%> 
              </select>
              <br>
              <font color="#3399FF">And</font><br>
              <select name="Searchtodate_day">
                <%Create_daylist(day(date)) %> 
              </select>
              <select name="Searchtodate_month">
                <% Create_MonthList(month(date))%> 
              </select>
              <select name="searchtodate_Year">
                <% Create_YearList(Year(date))%> 
              </select>
            </form>
          </td>
          <td width="3%" align="right" valign="bottom">&nbsp;</td>
        </tr>

        <tr> 
          <td width="6%" bgcolor="#FFD553" valign="Bottom"><img src="../datamanager/images/y03.jpg" width="12" height="12"></td>
          <td width="91%" bgcolor="#FFD553" align="right" valign="Bottom"><img src="../datamanager/images/y04.jpg" width="12" height="12"></td>
          <td width="3%" align="right" valign="bottom"><img src="../security/images/spacer.gif" width="1" height="1"></td>
        </tr>



        <tr bgcolor="#FFFFFF"> 
          <td colspan="3" height="16" valign="bottom">&nbsp;</td>
        </tr>
			<%End if%>
        <%
if Trim(Mid(Session("sAuthor"),22,3)) <> "" Then
	if Mid(Session("sAuthor"),22,3) = "AED" or Mid(Session("sAuthor"),22,3) = "VIW" Then%> 
        <tr> 
          <td width="6%" bgcolor="#FFD553" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><img src="../security/images/y01.jpg" width="12" height="12"></b></font></td>
          <td width="91%" align="right" bgcolor="#FFD553" valign="top"><img src="../datamanager/images/y02.jpg" width="13" height="12"></td>
          <td width="3%" align="right" valign="bottom">&nbsp;</td>
        </tr>
        <tr> 
          <td width="6%" bgcolor="#FFD553">&nbsp;</td>
          <td width="91%" bgcolor="#FFD553"> 
            <form method="post" action="../workflow/searchpo.asp">
              <b><font color="#0066FF">Search P/O :</font></b><br>
              <input type="text" name="ssearchpo" size="13" maxlength="12">
              <input type ="Submit" name="submit" value="Search">
 <!--             <br>
              <select name="status">
                <option value="1" selected>P/R No.</option>
                <option value="2">Invoice No.</option>
                <option value="3">Tax No.</option>
              </select> -->
              <br>
              <font color="#3399FF">Search Between</font> <br>
              <select name="Searchdate_day">
                <%Create_daylist(day(date-7)) %> 
              </select>
              <select name="Searchdate_month">
                <% Create_MonthList(month(date-7))%> 
              </select>
              <select name="searchdate_Year">
                <% Create_YearList(Year(date-7))%> 
              </select>
              <br>
              <font color="#3399FF">And</font><br>
              <select name="Searchtodate_day">
                <%Create_daylist(day(date)) %> 
              </select>
              <select name="Searchtodate_month">
                <% Create_MonthList(month(date))%> 
              </select>
              <select name="searchtodate_Year">
                <% Create_YearList(Year(date))%> 
              </select>
            </form>
          </td>
          <td width="3%" align="right" valign="bottom">&nbsp;</td>
        </tr>
        <tr> 
          <td width="6%" height="16" bgcolor="#FFD553" valign="bottom"><img src="../datamanager/images/y03.jpg" width="12" height="12" vspace="0" hspace="0"></td>
          <td width="91%" height="16" bgcolor="#FFD553" align="right" valign="bottom"><img src="../datamanager/images/y04.jpg" width="12" height="12"></td>
          <td width="3%" align="right" valign="bottom" height="16"><img src="../security/images/spacer.gif" width="1" height="1"></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="3" height="16" valign="bottom">&nbsp;</td>
        </tr>
        <%End if
		End if%> 
      </table>
      <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"></font></td>
    <td width="61%" align="center" valign="top"> 
<%	if Mid(Session("sAuthor"),13,3) <> "NOT" or Mid(Session("sAuthor"),22,3) <> "NOT" or Mid(Session("sAuthor"),31,3) <> "NOT" Then%> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor="#FFD553"> 
          <td align="left" valign="top" width="2%" height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFCC66">..</font></b></font></td>
          <td width="96%" height="2" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFCC66">..</font></b></font></td>
          <td align="right" valign="top" width="2%" height="2"><img src="../datamanager/images/y02.jpg" width="13" height="12"></td>
        </tr>
        <tr> 
          <td width="2%" bgcolor="#FFD553"><font color="#FFCC66">..</font></td>
          <td width="96%" align="left" valign="top" bgcolor="#FFCC99"> 
<%	if Mid(Session("sAuthor"),13,3) <> "NOT" Then%> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr bgcolor="#6666CC"> 
                <td width="135" ><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">P/R 
                  Approval</font></b></font></td>
                <td width="100" > 
                  <div align="right"><%
	if Mid(Session("sAuthor"),13,1) = "A" Then
%><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>[ 
                    <a href="addpr.asp"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Create 
                    PR </font></a> ]</b></font><FONT COLOR="#006699"><% End if%>
					<img src="../images/dot.gif"></FONT></div>
                </td>
              </tr>
              <tr> 
                <td height="14" colspan="2"> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <%'If nPrcount >= 1 or nPrcount2>=1  then%> 

                    <tr> 
                      <td bgcolor="#CCCCCC"  colspan=3><font color="#006699"><b>&nbsp;Waiting For Your Approval</b></font></td>
                    </tr>
                    <%work()%>
                    <tr> 
                      <td bgcolor="#FFCC99" colspan="3">&nbsp;</td>
                    </tr> 
             	   <tr> 
                      <td bgcolor="#CCCCCC" width="195"> <FONT color="#006699" face="MS Sans Serif" >&nbsp;&nbsp;<b>Representative 
                        Approve </b></FONT></td>
                      <td width="100" align="center" bgcolor="#CCCCCC">&nbsp;</td>
                      <td width="169" align="center" bgcolor="#CCCCCC"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#006699">[ 
                        <a href="../workflow/dupity.asp"><font color="#FF6666">Set Representative</font></a> ] &nbsp;</font></b></font> </td>
                    </tr>
             	   <tr> 
                      <td bgcolor="#CCCCCC" nowrap colspan="3"><FONT color="#006699" face="MS Sans Serif" >&nbsp;&nbsp;<b><%
					'########## ADD BY DERMPAN ON 22-SEP-03 ##########
					 ssql = "select count(distinct sprno) as c from pprapp where sempno = '" & session("sEmpNo") & "' and not sempnoapp is null and sempnoapp <> sempno"

					sSQL = "SELECT  count(distinct sprno) as c " _
					& "FROM vpprapp " _
					& "where (sEmpno='" & Session("sEmpno") & "'" _
					& " and not sempnoapp is null and sempnoapp <> sempno) and cApprov='1'" _
					& " and tapptime<>'-'" _
					& " and cpo <>'4' and nstep<>'0' " _
					& "ORDER BY sprno "
					'response.write sSql
					rs.open ssql, conn
					  
					if rs("c") <> "0" then
					  	response.write "<a href=""repapplst.asp"">Approve by representative</a> (" & rs("c") & ")"
					else
						response.write "&nbsp;"
					end if
					
					rs.close
					'########## END ADD BY DERMPAN ON 22-SEP-03 ##########
					  %></b></FONT></td>
                    </tr>
							<%Dupity()%>
                    <tr> 
                      <td bgcolor="#FFCC99" colspan="3">&nbsp;</td>
                    </tr> 
                    <tr> 
                      <td bgcolor="#CCCCCC" width="195" HEIGHT="2" colspan="3"> <FONT color="#006699" face="MS Sans Serif" >&nbsp;&nbsp;<a href="../workflow/finish.asp">Finished</a> 
                        </FONT>( <%LastWork()%> )</td>
                   <!--   <td width="116" align="center" bgcolor="#CCCCCC" HEIGHT="2">&nbsp;</td>
                      <td width="159" align="center" bgcolor="#CCCCCC" HEIGHT="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b><font color="#006699">&nbsp;</font></b></font> 
                      </td> -->
                    </tr>
                    <tr> 
                      <td bgcolor="#FFCC99" colspan="3">&nbsp;</td>
                    </tr>                     
                    <tr> 
                      <td bgcolor="#CCCCCC" width="195" colspan="3"> <FONT color="#006699" face="MS Sans Serif" >&nbsp;&nbsp;<b><a href="../workflow/come.asp">Coming Soon</a> </b></font>( <%wait()%> )</td>
<!--                      <td width="116" align="center" bgcolor="#CCCCCC"></td>
                      <td width="159" align="center" bgcolor="#CCCCCC"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b><font color="#006699">&nbsp;</font></b></font> 
                      </td> -->
                    </tr>
                    <tr> 
                      <td bgcolor="#FFCC99" colspan="3">&nbsp;</td>
                    </tr> 
                    <tr> 
                      <td bgcolor="#CCCCCC" width="195"><font color="#006699">&nbsp;&nbsp;<a href="../workflow/mine.asp">Mine</a> </font>( <% Req()%> )</td>
                      <td width="116" align="center" bgcolor="#CCCCCC">&nbsp;</td>
                      <td width="159" align="center" bgcolor="#CCCCCC"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b><font color="#006699">&nbsp; 
                        </font></b></font> </td>
                    </tr>
                    <tr> 
                      <td bgcolor="#FFCC99" colspan="3">&nbsp;</td>
                    </tr> 
                    <tr> 
                      <td bgcolor="#CCCCCC" width="195"><font color="#006699">&nbsp;&nbsp;<a href="../workflow/urgent.asp">Urgent Ack.</a> </font>( <% urgent()%> )</td>
                      <td width="116" align="center" bgcolor="#CCCCCC">&nbsp;</td>
                      <td width="159" align="center" bgcolor="#CCCCCC"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b><font color="#006699">&nbsp; 
                        </font></b></font> </td>
                    </tr>                          
                  </table>
                </td>
              </tr>
            </table>

            <%
	End if

	if Mid(Session("sAuthor"),22,3) = "VIW" or Mid(Session("sAuthor"),22,3) = "AED" or Mid(Session("sAuthor"),22,3) = "ADD" Then
%><br>
<%

sSQL = "SELECT /*+ORDERED NO_EXPAND * */ count(distinct(spono)) as countpr FROM pPO,ppr,pprapp where (ppo.cStatus='2' or ppo.cStatus='3') " &_
	" and ppo.srefno=ppr.sprno" &_
	" and ppr.cpo <> '1'" &_
	" and ppr.sprno=pprapp.sprno" &_
	" and (pprapp.sEmpno='" & Session("sEmpno") &_
	"' or pprapp.sEmpnoapp='" & Session("sEmpno") &_
	"' or ppr.sKey='" & Session("sEmpno") &_
	"') and (ppo.ctax <>'2' or ppo.cinvoice<>'2' Or ppo.ctax is null or ppo.cinvoice is null) " &_
	" ORDER BY dPodate DESC"
Rs2.Open sSQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
%> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr bgcolor="#6666CC"> 
                <td width="66%"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">P/O 
                  Management </font></b></font></td>
                <td width="34%" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b><font color="#006699">[ 
                  <a href="../workflow/dupity.asp"><font color="#FF6666">Set Representative</font></a> 
                  ] </font></b></font></td>
              </tr>
              <tr> 
                <td  colspan="2"> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
				<%	if Mid(Session("sAuthor"),22,3) = "AED" then%>
		            <tr> 
                      <td bgcolor="#CCEECC" width="251"> <FONT color="#006699" face="MS Sans Serif" >&nbsp;&nbsp;<b>
						<a href="../workflow/poforprint.asp">P/O For Print</a> </b></font>( <%wait_createpo()%> )</td>
                      <td width="147" align="center" bgcolor="#CCEECC">&nbsp;</td>
                      <td width="198" align="center" bgcolor="#CCEECC"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b><font color="#006699">&nbsp;</font></b></font> 
                      </td>
                    </tr>
                    <tr> 
                      <td bgcolor="#FFCC99" colspan="3">&nbsp;</td>
                    </tr> 
				<%End if%>
                    <tr> 
                      <td bgcolor="#CCEECC" width="251"> <FONT color="#006699" face="MS Sans Serif">&nbsp;&nbsp; 
                        <a href="../workflow/poprint.asp">P/O Printed</a> </font>( <%response.write RS2("countpr")%> )</td>
                      <td width="147" align="center" bgcolor="#CCEECC">&nbsp;</td>
                      <td width="198" align="center" bgcolor="#CCEECC"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b><font color="#006699">&nbsp;</font></b></font> 
                      </td>
                    </tr>
                    <tr> 
                      <td bgcolor="#FFCC99" colspan="3">&nbsp;</td>
                    </tr>
                    <tr> 
<%
Rs2.Close
sSQL = "SELECT  count(distinct(spono)) countpo FROM pPO,ppr,pprapp where ppo.cStatus='9'" &_
	" and ppo.srefno=ppr.sprno" &_
	" and ppr.sprno=pprapp.sprno" &_
	" and (pprapp.sEmpno='" & Session("sEmpno") &_
	"' or pprapp.sEmpnoapp='" & Session("sEmpno") &_
	"') and ppo.dAcceptdate >=" & to_date(date-14) 

Rs2.Open sSQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
%>
                      <td bgcolor="#CCEECC" width="251"> <FONT color="#006699" face="MS Sans Serif">&nbsp;&nbsp; 
                        <a href="../workflow/pocancel.asp">P/O Cancel</a> </font>( <%response.write RS2("countpo")%> )</td>
                      <td width="147" align="center" bgcolor="#CCEECC">&nbsp;</td>
                      <td width="198" align="center" bgcolor="#CCEECC"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b><font color="#006699">&nbsp;</font></b></font> 
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
            <% End if%> <br>
            <%
sSQL="Select count(distinct(ppo.spono)) reccount from ppr,pprapp,ppo  where  ppr.cpo='2'" &_
" and ppr.sprno=pprapp.sprno" &_
" and ppr.sprno=ppo.srefno" &_
" and ppo.cstatus<>'9'" &_
" and ppo.cstatus='2'" &_
" and (pprapp.sEmpno='" & Session("sEmpno") & "'" &_
" Or ppr.sKey='" & Session("sEmpno") & "') "



Rs5.Open sSQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
sSQL="Select count(distinct(planreceive.ID)) countrec from ppr,pprapp,ppo,planreceive  where " &_
" planreceive.sprno=ppr.sprno" &_
" and planreceive.spono=ppo.spono" &_
" and ppr.sprno=pprapp.sprno" &_
" and ppr.sprno=ppo.srefno" &_
" and ppo.cstatus<>'9'" &_
" and (pprapp.sEmpno='" & Session("sEmpno") & "'" &_
" Or ppr.sKey='" & Session("sEmpno") & "') "



Rs.Open sSQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

%>	
<%if RS5("recCount")<>"0" Or RS("countrec")<> "0" then %>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr bgcolor="#6666CC"> 
                <td><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">User Plan Receive Date</font></b></font> </td>
              </tr>
              <tr> 
                <td> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">

                    <tr> 
                      <td bgcolor="#CCCCCC" width="300" colspan="3"> <font color="#006699" face="Verdana, Arial, Helvetica, sans-serif" >
						&nbsp;&nbsp;<b>
						<a href="../workflow/waitplanrec.asp">Waiting for (your) confirmation</a> ( <%=Rs5("Reccount")%> )
                       </b></font></td>
                      </td>
                    </tr>
                    <tr> 
                      <td bgcolor="#FFCC99" width="256" colspan="3"> <font color="#006699" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;
                       </b></font></font></td>
					</tr>
<%

				
%> 
                    <tr> 
                      <td bgcolor="#CCCCCC" width="300" colspan="3"> <font color="#006699" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;
						<a href="../workflow/finishplanrec.asp">Finish confirmation</a> ( <%=RS("countrec")%> )
                        </font></font></td>
                    </tr>
                    <%
' Do While Not RS.EOF 
'							sprno=RS("sPrno")
'							chkUser()
%> 
  
                  </table>
                </td>
              </tr>
            </table>
            <%
End if
RS5.Close
RS.Close
%>		
			
			<br>
<%
			
			sSQL = "SELECT  count(distinct(sRecno)) countrec FROM pReceive,ppr,pprapp  where ppr.cpo <>'4'" &_
" and preceive.sprno=ppr.sprno" &_
" and ppr.sprno=pprapp.sprno" &_
" and (pprapp.sEmpno='" & Session("sEmpno") & "'" &_
" Or ppr.sKey='" & Session("sEmpno") & "')" &_
" ORDER BY dRecdate DESC"
'response.write sSQL
Rs.Open sSQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
'IF (Not RS5.EOF Or Not Rs.EOF) and (Mid(Session("sAuthor"),22,3) = "VIW" or Mid(Session("sAuthor"),22,3) = "AED") then

%> 
<%if RS("countrec")<> "0" then %>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr bgcolor="#6666CC"> 
                <td><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Receiving 
                  Management </font></b></font> </td>
              </tr>
              <tr> 
                <td> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
              

                    <tr> 
                      <td bgcolor="#CCCCCC" width="256" colspan="3"> <font color="#006699" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;
						<a href="../workflow/finishrec.asp">Finished Receive</a> ( <%=RS("countrec")%> )
                        </font></font></td>
                    </tr>
                    <%
' Do While Not RS.EOF 
'							sprno=RS("sPrno")
'							chkUser()
%> 
  
                  </table>
                </td>
              </tr>
            </table>
            <%
End if
RS.Close
			
			
			
			
			
			
          
if Mid(Session("sAuthor"),37,3) = "AED" then
sSQL="Select count(distinct(ppo.spono)) reccount from ppr,pprapp,ppo  where  ppr.cpo='2'" &_
" and ppr.sprno=pprapp.sprno" &_
" and ppr.sprno=ppo.srefno" &_
" and ppo.cstatus<>'9'" &_
" and ppo.cstatus='2'"
'response.write ssql

Rs5.Open sSQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
sSQL = "SELECT  count(distinct(sRecno)) countrec FROM pReceive,ppr,pprapp  where ppr.cpo <>'4'" &_
" and preceive.sprno=ppr.sprno" &_
" and ppr.sprno=pprapp.sprno" &_
" ORDER BY dRecdate DESC"
Rs.Open sSQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
'IF (Not RS5.EOF Or Not Rs.EOF) and (Mid(Session("sAuthor"),22,3) = "VIW" or Mid(Session("sAuthor"),22,3) = "AED") then

%> 

<%if RS5("recCount")<>"0" Or RS("countrec")<> "0"  then %>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr bgcolor="#6666CC"> 
                <td><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">W/H Receive</font></b></font> </td>
              </tr>
              <tr> 
                <td> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">

                    <tr> 
                      <td bgcolor="#CCCCCC" width="256" colspan="3"> <font color="#006699" face="Verdana, Arial, Helvetica, sans-serif" >
						&nbsp;&nbsp;<b>
						<a href="../workflow/waitreceive.asp">Waiting Receiving</a> ( <%=Rs5("Reccount")%> )
                       </b></font></td>
                      </td>
                    </tr>
                    <tr> 
                      <td bgcolor="#FFCC99" width="256" colspan="3"> <font color="#006699" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;
                       </b></font></font></td>
					</tr>
<%

				RS5.Close
%> 
                    <tr> 
                      <td bgcolor="#CCCCCC" width="256" colspan="3"> <font color="#006699" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;
						<a href="../workflow/finishrec.asp">Finished Receive</a> ( <%=RS("countrec")%> )
                        </font></font></td>
                    </tr>
                    <%
' Do While Not RS.EOF 
'							sprno=RS("sPrno")
'							chkUser()
%> 
  
                  </table>
                </td>
              </tr>
            </table>
            <%
End if
RS.Close
End if
'if Mid(Session("sAuthor"),31,3) = "AED" or Mid(Session("sAuthor"),31,3) = "VIW" then
if Mid(Session("sAuthor"),31,3) = "AED" then
	'sSQL = "SELECT  Count(sInvno) countInv FROM pInvoice,preceive,ppr  where  pInvoice.sRecno=Preceive.sRecno and pReceive.sPrno=ppr.sprno " &_
	'" and pInvoice.ddate >=" &to_date(date-14) &_
	'" ORDER BY ddate DESC"
sSQL="Select count(sInvno) countInv from pInvoice where ddate >="& to_date(date-14)
'	sSQL = "SELECT  Count(sInvno) countInv FROM pInvoice,preceive,ppr  where  pInvoice.sRecno=Preceive.sRecno and pReceive.sPrno=ppr.sprno " &_
'	" and pInvoice.ddate >=" &to_date(day(date-14)&"/"&month(date-14)&"/"&year(date-14)) &_
'	" ORDER BY ddate DESC"
'	Response.Write sSQL
	Rs.Open sSQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
else
	sSQL = "SELECT  Count(distinct(pInvoice.sInvno)) countInv FROM pInvoice,preceive,pprapp,ppr  where pInvoice.ddate >=" &to_date(date-14) &_
	" and pinvoice.srecno=preceive.srecno" &_
	" and preceive.sprno=pprapp.sprno" &_
	" and pprapp.sprno=ppr.sprno" &_
	" and pprapp.sEmpno='" & Session("sEmpno") & "'" &_
	" ORDER BY ddate DESC"
	Rs.Open sSQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

End if

'	if Mid(Session("sAuthor"),22,3) = "VIW" or Mid(Session("sAuthor"),22,3) = "AED" Then

'		If Session("sSeccode")="40202" Or Mid(Session("sAuthor"),1,3) = "999" Or Mid(Session("sAuthor"),31,3) = "AED" Or Mid(Session("sAuthor"),31,3) = "VIW" then
'		If Mid(Session("sAuthor"),1,3) = "999" Or Mid(Session("sAuthor"),31,3) = "AED" Or Mid(Session("sAuthor"),31,3) = "VIW" then
%><br>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr bgcolor="#6666CC"> 
                <td><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Invoice/Tax (Copy) 
                  </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#666666"><b><font color="#FFFFFF">Management</font></b></font><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF"> 
                  </font></b></font></b></font></td>
              </tr>
              <tr> 
                <td> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
<%	if Mid(Session("sAuthor"),31,3) = "AED" then %>
                    <tr> 
                      <td bgcolor="#CCCCCC" width="254" colspan="3"> <font color="#006699" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;<font size="2"><b>
						<a href="../workflow/waitinvoice.asp">Waiting For</a> </b></font></font>( <%chkInvoice()%> )</td>
                    </tr>
                    <tr> 
                      <td bgcolor="#FFCC99" colspan="3">&nbsp;</td>
                    </tr>                     
<%End if%>
                    <tr> 
                      <td bgcolor="#CCCCCC" width="254" colspan="3"> <font color="#006699" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;<font size="2">
						<a href="../workflow/finishInv.asp">Finished</a> </font></font>( <%=RS("countInv")%> )</td>
                    </tr>
                    <tr> 
                      <td width="254"> <font color="#006699" face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
                        </font></td>
                      <td width="143" align="center">&nbsp;</td>
                      <td width="189" align="center"> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b> 
                        </b></font><FONT COLOR="#006699"> </FONT> </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
            <% 
RS.Close
'End if
			%><br>
            <%
'	if Mid(Session("sAuthor"),22,3) = "VIW" or Mid(Session("sAuthor"),22,3) = "AED" Then
%><%

'if Mid(Session("sAuthor"),22,3) = "AED" then
'	sSQL = "SELECT  Count(sTaxNo) counttax FROM pTax where ddate >=" & to_date(date-14)
'	Rs.Open sSQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
'else
	sSQL = "SELECT  Count(distinct(ptax.sTaxNo)) counttax FROM pTax,preceive,pprapp,ppr where ptax.ddate >=" &to_date(date-14) &_
	" and ptax.srecno=preceive.srecno" &_
	" and preceive.sprno=pprapp.sprno" &_
	" and ppr.sprno=pprapp.sprno" &_
	" and (pprapp.Sempno='" & Session("sEmpno") & "'" &_
	" Or ppr.sKey='" & Session("sEmpno") & "')"

	Rs.Open sSQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
'End if
'response.write sSQL
%> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr bgcolor="#6666CC"> 
                <td><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Invoice/Tax (Original) 
                  Management </font></b></font></b></font></td>
              </tr>
              <tr> 
                <td> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
				<%	if Mid(Session("sAuthor"),22,3) = "AED" then%>
                    <tr> 
                      <td bgcolor="#CCCCCC" width="255" colspan="3"> <font color="#006699" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;<b>
						<a href="../workflow/waittax.asp">Waiting For</a> </b></font>( <% chkTax()%> )</font></td>
                    </tr>
                    <tr> 
                      <td bgcolor="#FFCC99" colspan="3">&nbsp;</td>
                    </tr>                    
				<%End if%>
                    <tr> 
                      <td bgcolor="#CCCCCC" width="255" colspan="3"> <font color="#006699" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;
						<a href="../workflow/finishtax.asp">Finished</a> </font></font>( <%Response.write RS("Counttax")%> )</td>
                    </tr>
                    <tr> 
                      <td width="255"> <font color="#006699" face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
                        </font></td>
                      <td width="143" align="center">&nbsp;</td>
                      <td width="188" align="center"> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b> 
                        <a href="../workflow/pTaxedt.asp"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                        </font></a> </b></font><FONT COLOR="#006699"> </FONT> 
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
            <%' End if%><font color="#006699"> </font></td>
          <td width="2%" bgcolor="#FFD553"><font color="#FFCC66">..</font></td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom"><img src="../datamanager/images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="bottom"><img src="../images/spacer.gif" width="10" height="12"></td>
          <td width="2%" align="right" valign="bottom"><img src="../datamanager/images/y04.jpg" width="12" height="12"></td>
        </tr>
      </table>
<%End if%>
    </td>
  </tr>
</table>
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr align="center"> 
    <td width="1%"> <img src="images/spacer.gif" width="10" height="10"><br>
    </td>
    <td width="99%"> 
     <!--#include file="../inc/copyRight.inc"--></td>
  </tr>
</table>

</body>
</html>
<%

Sub Wait()
		Set Rs2=Server.CreateObject("ADODB.Recordset")
	'	sSQL="Select count(sprno) countpr from pprapp where (sEmpno='" & Session("sEmpno") & "'" &_
	'	" Or sEmpno='" & Session("sEmpno") & "'" &_
	'	") and cApprov='0'"

	sSQL = "SELECT  count(distinct(ppr.sprno)) countpr  FROM pprapp,ppr  where pprapp.sprno=ppr.sprno" &_
	" and (pprapp.sEmpno='" & Session("sEmpno")  &_
	"'  Or pprapp.sEmpnoapp='" & Session("sEmpno")  &_
	"') and pprapp.cApprov='0'" &_
	" and pprapp.nstep<>ppr.nstep" &_
	" and ppr.nstep<>'0'" &_
	" and pprapp.nstep<>'35'" &_
	" ORDER BY pprapp.sprno "
		 Rs2.Open sSQL,conn,adOpenForwardOnly, adLockReadOnly, adCmdText
	if Not RS2.EOF then
		response.write Rs2("countpr")
	End if
	RS2.Close
'	Set Rs2=Nothing
End Sub

Sub work()
		Set Rs2=Server.CreateObject("ADODB.Recordset")
	Color="#F1B3A7"
'For i=0 to nPrcount-1
'	If approve(i)=0 or Approve(i)="0" then
	'31-MAY-04,Dermpan,refer to monthly working report for MAY-04 item 7.
	sSQL= "Select *  from vpprapprove  where  sEmpno='" & Session("sEmpno") & "'" &_
	" and nstep<>'35' and cApprov = '0' ORDER BY sprno" 
	Rs2.Open sSQL,conn,adOpenForwardOnly, adLockReadOnly, adCmdText
	if Not RS2.EOF then
		Do While Not Rs2.EOF %>
<tr>
	<td bgcolor="<%=Color%>" colspan="1">&nbsp;&nbsp;
		<a href='pprapp.asp?sprno=<%=RS2("sPrno")%>'>
			<%response.write RS2("sPrno")%>
		</a>
	</td>
   <td  align="Left" bgcolor="<%=Color%>" colspan=2>&nbsp;
		<% if RS2("cprtype")="1" then %><b><Blink>
			<%=RS2("sKey")%>/<%=RS2("sRequester")%></Blink></b>
		<%Else%>
		<%=RS2("sKey")%>/<%=RS2("sRequester")%><%End if%></td>
<!--		<td bgcolor="<%=Color%>"><img src="../images/dot.gif"></td>-->
</tr>
<%
		RS2.MoveNext
if Color="#F8D7D1" then
		Color="#F1B3A7"
else
		Color="#F8D7D1"
end if
		Loop
		End if
	RS2.Close
'	end if
'Next
End Sub

Sub LastWork()
		Set Rs2=Server.CreateObject("ADODB.Recordset")
		sSQL="Select count(distinct(sprno)) countpr from vpprapp where (((sEmpno='" & Session("sEmpno") & "'" &_
		" Or sComment='" & Session("sEmpno") & "'" &_
		") and cApprov='1'" &_
		" and tapptime<>'-'" &_
		" and cpo <>'4' and nstep<>'0' )"  &_
		"  or ( nstep='0'  and dFinishdate >="& to_date(date-14) & "))"

		 Rs2.Open sSQL,conn,adOpenForwardOnly, adLockReadOnly, adCmdText

	if Not RS2.EOF then
		response.write Rs2("countpr")
	End if
	RS2.Close
'	Set Rs2=Nothing

End Sub

Sub Wait_createpo()
		Set Rs3=Server.CreateObject("ADODB.Recordset")
		sSQL = "SELECT count(distinct(ppo.spono)) countpo  FROM  ppo,pprapp WHERE  ppo.sRefno=pprapp.sPrno" &_
		" and (pprapp.sEmpno='" & Session("sEmpno") &_
		"' or pprapp.sEmpnoapp='" & Session("sEmpno") &_
		"') and ppo.cStatus<>'2' and ppo.cStatus<>'3' and cStatus<>'9'"
		 Rs3.Open sSQL,conn,adOpenForwardOnly, adLockReadOnly, adCmdText
'response.write sSQL
	if Not RS3.EOF then
		response.write Rs3("CountPo")
	End if
	RS3.Close
	Set Rs3=Nothing
End Sub

Sub wait_Createpo2()
		Set Rs3 = Server.CreateObject("ADODB.RecordSet")
'For i=0 to nPrcount-1
		sSQL = "SELECT distinct (pprapp.sprno),ppo.srefno,ppo.spono  FROM  ppo,pprapp WHERE  ppo.sRefno=pprapp.sPrno" &_
		" and (pprapp.sEmpno='" & Session("sEmpno") &_
		"' or pprapp.sEmpnoapp='" & Session("sEmpno") &_
		"') and nstep<>'35' and cStatus<>'2' and cStatus<>'3'"
		Rs3.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
		IF Not Rs3.EOF then
			Do While Not RS3.EOF
			sprno=RS3("sRefno")
			chkUser()
'	If Session("sEmpno")=Mid(Empno,1,5) then
If chk="1" then
%>
<tr>
<td bgcolor="<%=Color%>" colspan="1">&nbsp;&nbsp;
<a href='ppodesc.asp?spono=<%=RS3("sPono")%>'>
<%response.write RS3("sPono") & "</a></td>"%>
</a>
	<td bgcolor="<%=Color%>" align="center"><%=RS3("sRefno")%></td>
  <td bgcolor="<%=Color%>" align="center" colspan="1"><a href='ppodesc.asp?spono=<%=RS3("sPono")%>'> 
    <%	if  Mid(Session("sAuthor"),22,3) = "AED" Then %> </a><a href=../workflow/ppo.asp?spono='<%=RS3("sPono")%>'>Edit 
    P/O </a><%End if%>
  <td></tr>
<%

End if
			RS3.MoveNext
if Color="#F8D7D1" then
		Color="#F1B3A7"
else
		Color="#F8D7D1"
end if
			Loop
		End if
			RS3.Close

End Sub
Sub ChkUser()
		Set Rs4 = Server.CreateObject("ADODB.RecordSet")
				sSQL="Select * from pprapp where  sPrno='" & sPrno & "'" &_
				" Order By nStep "
		Rs4.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	If Not Rs4.EOF then
		Do while Not Rs4.EOF
			Empno=RS4("sEmpno")
				If RS4("sEmpno")=Session("sEmpno") Or Mid(Session("sAuthor"),22,3) = "VIW" Then
					chk="1"
				end if
		Rs4.Movenext
		Loop
	End if
		RS4.Close
End Sub
Sub ChkReceiv()
		Set Rs4 = Server.CreateObject("ADODB.RecordSet")
				sSQL="Select * from ppritem where  sPrno='" & sPrno & "'" &_
				" and cStatus='2'" &_
				"  Order By nItem "
		Rs4.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	If Not Rs4.EOF then
		Do while Not Rs4.EOF
					chkreceive="0"
		Rs4.Movenext
		Loop
	Else
		chkreceive="1"
	End if
		RS4.Close
End Sub

Sub ChkRec1()
		Set Rs4 = Server.CreateObject("ADODB.RecordSet")
				sSQL="Select * from pReceive where  sPrno='" & sPrno & "'" 
		Rs4.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	If Not Rs4.EOF then
		Do while Not Rs4.EOF
					chkrec=RS4("sRecno")
		Rs4.Movenext
		Loop
	Else
		chkrec=""
	End if
		RS4.Close
End Sub

Sub Req()
	Set Rs2=Server.CreateObject("ADODB.Recordset")
	
	sSQl="Select count(sprno) countpr from ppr where (sRequester='" & Session("sEmpno") & "'" &_
	" Or skey='" & Session("sEmpno") & "'" &_
	") and ((ppr.nstep<>'0' and ppr.cPo<>'4')" &_
	" Or ppr.sprno in (Select sprno from ppr where nstep='0' and dprdate >="& to_date(date-60)  &"))" 

	sSQl="Select count(sprno) countpr from ppr where (sRequester='" & Session("sEmpno") & "'" &_
	" Or skey='" & Session("sEmpno") & "'" &_
	") and ((ppr.nstep<>'0' and ppr.cPo<>'4')" &_
	" Or ppr.sprno in (Select sprno from ppr where nstep='0' and dprdate >= to_date('"& day(date-60) & "/" & month(date-60) & "/" & year(date-60)  &"','dd/mm/yyyy')))" 


'response.write day(date-60) & "/" & month(date-60) & "/" & year(date-60)	
'response.write date-60
'	response.write sSQl
	Rs2.Open sSQL,conn,adOpenForwardOnly, adLockReadOnly, adCmdText
	if Not RS2.EOF then
		response.write Rs2("countpr")
	End if
	RS2.Close
'	Set Rs2=Nothing

End Sub


Sub Urgent()
		Set Rs2=Server.CreateObject("ADODB.Recordset")
			sSQl="Select count(distinct(ppr.sprno)) countpr from ppr,pprapp where  ppr.sprno=pprapp.sprno" &_
			" and pprapp.sEmpno='" & session("sEmpno") & "'" &_
			" and ppr.sRequester<>'" & session("sEmpno") & "'" &_
			" and ppr.cprtype='1'" &_
			" and pprapp.sstep<>'10'" &_
			" and pprapp.nstep <> '30'" &_
			" and pprapp.nstep <> '35'" &_ 
			" and pprapp.nstep<>ppr.nstep" &_ 
			" and tApptime='-'" &_
			" and ((ppr.nstep > pprapp.nstep" &_
			" and ppr.cpo <>'4')" &_
			" OR (ppr.nstep='0'  and dFinishdate>="&to_date(date-14)&"))"
			
		 Rs2.Open sSQL,conn,adOpenForwardOnly, adLockReadOnly, adCmdText
	if Not RS2.EOF then
		response.write Rs2("countpr")
	End if
	RS2.Close
'	Set Rs2=Nothing

End Sub

Sub Dupity()
		Set Rs2=Server.CreateObject("ADODB.Recordset")
	Color="#F1B3A7"
'For i=0 to nPrcount2-1
'	If approve2(i)=0 or Approve2(i)="0" then
	sSQL= "Select *  from vpprapprove  where  sEmpnoapp='" & Session("sEmpno") & "'" &_
	" and nstep<>'35' and CAPPROV = 0 ORDER BY sprno" 
	'response.write sSQL
		 Rs2.Open sSQL,conn,adOpenForwardOnly, adLockReadOnly, adCmdText
	if Not RS2.EOF then
		Do While Not Rs2.EOF %>
<tr>
	<td bgcolor="<%=Color%>" colspan="1">&nbsp;&nbsp;
		<a href='pprapp.asp?sprno=<%=RS2("sPrno")%>'>
			<%response.write RS2("sPrno")%>
		</a>
	</td>
   <td  align="Left" bgcolor="<%=Color%>" colspan=2>&nbsp;
		<% if RS2("cprtype")="1" then %><b><Blink>
			<%=RS2("sKey")%>/<%=RS2("sRequester")%></Blink></b>
		<%Else%>
		<%=RS2("sKey")%>/<%=RS2("sRequester")%><%End if%></td>
<!--		<td bgcolor="<%=Color%>"><img src="../images/dot.gif"></td>-->
</tr>
<%
		RS2.MoveNext
if Color="#F8D7D1" then
		Color="#F1B3A7"
else
		Color="#F8D7D1"
end if
		Loop
		End if
	RS2.Close
'	end if
'Next
End Sub

Sub Create_Yearlist(searchyear)
Dim syear
	For i=2000 to searchyear
		IF i=cInt(searchyear) then
			Response.Write "<option value= '"& i & "'selected>"& i &"</option>"& vbCrLf
		Else
			Response.Write "<option value= '"& i & "'>"& i &"</option>"& vbCrLf
		End IF
	Next
End Sub
Function GetMonthName(m)
	Select Case m
		Case 1
			GetMonthName = "January"
		Case 2
			GetMonthName = "Febuary"
		Case 3
			GetMonthName = "March"
		Case 4
			GetMonthName = "April"
		Case 5
			GetMonthName = "May"
		Case 6
			GetMonthName = "June"
		Case 7
			GetMonthName = "July"
		Case 8
			GetMonthName = "August"
		Case 9
			GetMonthName = "September"
		Case 10
			GetMonthName = "October"
		Case 11
			GetMonthName = "November"
		Case 12
			GetMonthName = "December"
	End Select
End Function

Sub chkInvoice1()
		Set Rs2=Server.CreateObject("ADODB.Recordset")
	Color="#F1B3A7"
		sSQL="Select distinct ppo.spono,ppr.sprno,ppo.dpodate,ppo.svendcode from ppo,ppr where ppo.srefno=ppr.sprno" &_
		" and ppr.cpo='3'" &_
		" and ppo.cInvoice='1'" &_
		" and ppo.cInvoice<>'2'"
		 Rs2.Open sSQL,conn,adOpenForwardOnly, adLockReadOnly, adCmdText
	if Not RS2.EOF then
		Do While Not Rs2.EOF %>
<tr>
	<td bgcolor="<%=Color%>" colspan="1">&nbsp;&nbsp;
		<a href='pinvoiceedt.asp?spono=<%=RS2("sPono")%>&svendcode=<%=RS2("sVendcode")%>'>
			<%response.write RS2("sPono")%>
		</a>
	</td>
	<td bgcolor="<%=Color%>" colspan="1"align="center">
			<%response.write RS2("sPrno")%>
	</td>
	<td bgcolor="<%=Color%>" colspan="1"align="center">
			<%response.write RS2("dpodate")%>
	</td>
</tr>
<%
		RS2.MoveNext
if Color="#F8D7D1" then
		Color="#F1B3A7"
else
		Color="#F8D7D1"
end if
		Loop
		End if
	RS2.Close

End Sub

Sub chkInvoice()
		Set Rs2=Server.CreateObject("ADODB.Recordset")
	Color="#F1B3A7"
		sSQL="Select count(sRecno) countrec from preceive where cstatus <> 'A' and cstatus <> 'I'"
		 Rs2.Open sSQL,conn,adOpenForwardOnly, adLockReadOnly, adCmdText
	if Not RS2.EOF then
		response.write RS2("Countrec")
	End if
End Sub


Sub chkTax()
		Set Rs2=Server.CreateObject("ADODB.Recordset")
	Color="#F1B3A7"
		sSQL="Select count(distinct(preceive.sRecno)) countrec from preceive,ppr,pprapp where preceive.sprno=ppr.sprno" &_
		" and ppr.sprno=pprapp.sprno" &_
		" and (pprapp.sEmpno='" & Session("sEmpno") & "'" &_
		" Or ppr.sKey='" & Session("sEmpno") & "')" &_
		" and preceive.cstatus <> 'A' and preceive.cstatus <> 'T'"
		 Rs2.Open sSQL,conn,adOpenForwardOnly, adLockReadOnly, adCmdText
	if Not RS2.EOF then
		response.write RS2("Countrec")
	End if
End Sub
%>