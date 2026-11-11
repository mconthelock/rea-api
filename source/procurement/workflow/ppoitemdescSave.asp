<% Option Explicit %>
<!-- #include file ="../adovbs.inc"  -->
<!-- #include file ="../Conn.asp" -->

<%    	

     dim objsmartupload,rsimagelib ,rs,sSQL	                      'for obj
     dim spath,fd,file,sql,smode,check1,check2            'other variable
     dim sprno,sfilename,nitem,sitemno   'for variable in table     
     dim sImageType,sCatID,sCurcode,sUnitcode,sAcccode,chkfilename
     Dim sprodcode,sexptype,svendcode,sprodid,stprodname,seprodname,smodel,stdesc,sdesc,nprice,sfile,dadddate,sempno,status
     Dim dDate,dDay,dMonth,dYear,dHour,dMinute,dSecond
	 
     set objsmartupload = server.createobject("aspSmartUpload.SmartUpload")'create Obj     
'	 set uploadObj = server.createObject("SoftArtisans.FileUp")

     objsmartupload.MaxFileSize = 2500000                                   'Max filesize 

     objsmartupload.Upload                                     'upload to objsmartupload

     ObjSmartUpload.DeniedFilesList = "bat,exe,com,asp,dll"    

     smode = objsmartupload.form("EVENT")
	 status=objsmartupload.form("status")
     'data from pre page      	

     if smode = "ADD" then
		sprno=Session("sprno")
		nItem=Session("nitem")
				dDate=Now
				dDay=(Day(dDate))
				If dDay < 10 then dDay="0" & Day(dDate)
						dMonth=Month(dDate)
				If dMonth < 10 then dMonth="0" & Month(dDate)
						dYear=Mid(Year(dDate),2,2)
						dHour=Hour(dDate)
				If dHour < 10 then dHour="0" & Hour(dDate)
						dMinute=Minute(dDate)
				If dMinute < 10 then dMinute="0" & Minute(dDate)
						dSecond=Second(dDate)
				If dSecond < 10 then dSecond="0" & Second(dDate)
		   sitemno=dYear & dMonth & dDay & dHour & dMinute & dSecond
		sdesc = objsmartupload.Form("sdesc")
		sfile=objsmartupload.Form("sfile")
     else
		sprno=Session("sprno")
		nItem=Session("nitem")
		sItemno=Session("sitemno")
		sdesc = objsmartupload.Form("sdesc")
		sfile=objsmartupload.Form("sfile")
     end if

    spath = request.serverVariables("APPL_PHYSICAL_PATH")      'ชื่อ ทางphysic ของ server
    'response.write spath & "<br>"
	
''    spath = spath & "\workflow\ppritemdescImages\" & year(date) & "\" & dMonth         'check path corp
    spath = spath & "workflow\ppritemdescImages\" & year(date) & "\" & dMonth   
    'spath = spath & "procurement\workflow\ppritemdescImages\" & year(date) & "\" & dMonth         'check path corp
  'response.write spath

'response.end	
	 for each file in objsmartupload.files
	 	 sfilename = file.filename             'ชื่อของ file ที่load มา
 	 next	
     'If not ObjSmartUpload.Files.IsMissing Then
     'sFileName = ObjSmartUpload.Files.FileName
     'End if                                   'samesame
	
'	 check_blank()
     check_filename() 

if Instr(sFilename," ")="0" then     
  if chkfilename=0 then
     fd = objsmartupload.save(spath)           'save file to server at spath

     set rsimagelib = server.createobject("ADODB.Recordset")

     if smode = "ADD" then 

         fd = objsmartupload.save(spath)           'save file to server at spath

    	 'create recordset & save to DB
         sql = "select * from ppritemdesc"
         rsimagelib.open sql,conn,adOpenKeySet,adLockPessimistic,adCmdtext
	     rsimagelib.addnew
			rsimagelib("sprno") = sprno
			rsimagelib("nItem") = 1
			rsimagelib("sItemno") = sItemno
			rsimagelib("sdesc") = sdesc
			rsimagelib("sfile") = year(date) & "/" & dMonth  & "/" & LCase(sfilename)
	     rsimagelib.update    	 
     else 'smode = edit
		 'update Table
         sql = "SELECT * from ppritemdesc WHERE sprno='" & sprno & "'" &_
         " and nitem='" & nitem & "'" &_
         " and sitemno='" & sitemno & "'"
         rsimagelib.open sql,conn,adOpenKeySet,adLockPessimistic,adCmdtext
          If Not rsimagelib.EOF Then
			rsimagelib("sprno") = sprno
			rsimagelib("nItem") = 1
			rsimagelib("sItemno") = sItemno
			rsimagelib("sdesc") = sdesc
		if sfilename<>"" then
			rsimagelib("sfile") = year(date) & "/" & dMonth  & "/" & LCase(sfilename)
		end if
	     rsimagelib.update            
	     End IF
     end if

	 rsimagelib.close
	 set rsimagelib = nothing
'response.write "Add Complete"
 '    response.redirect("Javascript:window.close()")
 'IF status="" then
     Response.Redirect "ppo.asp?spono=" & objsmartupload.form("spono")
'Else
 '    Response.Redirect "pprapp.asp?sprno=" & sprno
'End IF
else
%>
<br>
<center><font color="Red" size=+2 ><Blink><b>Files's name Already in Use ,Please rename Your File!!!</b></Blink></font><br>
<a href="javascript:history.back()"><< back</a>
</center>
<%
	end if
else
%>
<br>
<center><font color="Red" size=+2 ><Blink><b>Files's name not use spacebar ,Please rename Your File!!!</b></Blink></font><br>
<a href="javascript:history.back()"><< back</a>
</center>
<%
End if
 'check blank 
   'set objSmartUpload & Upload & set sMode

   Sub check_blank()
     dim str

     check1 = false
     check2 = false 
            
     if trim(sfilename)  = "" then 
        check1 = true
     end if
     if trim(stimgdesc) = "" then 
        check2 = true
     end if 

        
    if (sMode = "add") and (check1 or check2) then 
        str = "bbimagehome.asp?"         
        if not check2 then
           str = str & "stimgdesc=" & stimgdesc & "&"    
        end if
        if (trim(seimgdesc) <> "") then
           str = str & "seimgdesc=" & "seimgdesc" & "&"    
        end if
        str = str & "event=" & sMode
        response.redirect(str)        
    end if

    if (sMode = "edit") and (check2) then 
        str = "sbimagehome.asp?"         
        if not check2 then
           str = str & "stimgdesc=" & stimgdesc & "&"    
        end if
        if (trim(seimgdesc) <> "") then
           str = str & "seimgdesc=" & seimgdesc & "&"    
        end if
        str = str & "simgid=" & simgid & "&"
        str = str & "event=" & sMode
        response.redirect(str)        
    end if


   End Sub
%>

<%
   Sub check_filename()
    Dim Rs,sSql,str

    sSql = "SELECT * FROM ppritemdesc where sfile = '" & year(date) & "/" & dMonth  & "/" & lcase(sfilename) & "'"
    set Rs = Server.CreateObject("ADODB.Recordset")
    Rs.Open sSql,Conn,adOpenKeySet,adLockPessimistic,adCmdtext
   
	if not Rs.eof then
 		chkfilename=1
	else
 		chkfilename=0           	
   	end if

   End Sub

%>
