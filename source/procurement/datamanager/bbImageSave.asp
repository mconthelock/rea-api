<% Option Explicit %>
<!-- #include file ="../adovbs.inc"  -->
<!-- #include file ="../../Conn.asp" -->

<%
     dim objsmartupload,rsimagelib 	                      'for obj
     dim spath,fd,file,sql,smode,check1,check2            'other variable
     dim scorpcode,simgid,sfilename,stimgdesc,seimgdesc   'for variable in table     
	  dim sImageType    

     set objsmartupload = server.createobject("aspSmartUpload.SmartUpload")'create Obj     
     objsmartupload.MaxFileSize = 50000                                   'Max filesize 
     objsmartupload.Upload                                     'upload to objsmartupload

     ObjSmartUpload.DeniedFilesList = "bat,exe,com,asp,dll"    

     smode = objsmartupload.form("event")
     'data from pre page      	
     if smode = "add" then
  	   scorpcode  = session("gbcorpcode")
       simgid     = mid(year(Date),3,2)&month(Date)&Day(date)&Hour(Now)&Minute(Now)&Second(now)
     else
       scorpcode  = objsmartupload.form("scorpcode")
       simgid     = objsmartupload.form("simgid")
     end if
     stimgdesc  = objsmartupload.form("stimgdesc")
     seimgdesc  = objsmartupload.form("seimgdesc")
     
	 spath = request.serverVariables("APPL_PHYSICAL_PATH")      'ª×èÍ ·Ò§physic ¢Í§ server
    spath = spath &"\shop\"& scorpcode & "\images"          'check path corp
'    response.write spath
     fd = objsmartupload.save(spath)           'save file to server at spath

	 for each file in objsmartupload.files
	 	 sfilename = file.filename             'ª×èÍ¢Í§ file ·Õèload ÁÒ
 	 next	
     'If not ObjSmartUpload.Files.IsMissing Then
     'sFileName = ObjSmartUpload.Files.FileName
     'End if                                   'samesame

	 check_blank()
     check_filename() 

     set rsimagelib = server.createobject("ADODB.Recordset")
	 
     if smode = "add" then 

         fd = objsmartupload.save(spath)           'save file to server at spath

    	 'create recordset & save to DB
         sql = "select * from bimagelib"
         rsimagelib.open sql,conn,adOpenKeySet,adLockPessimistic,adCmdtext
	     rsimagelib.addnew
			rsimagelib("scorpcode") = scorpcode
			rsimagelib("simgid")    = simgid  
			rsimagelib("sfilename") = sfilename
			rsimagelib("stimgdesc") = stimgdesc
			rsimagelib("seimgdesc") = seimgdesc
			rsimagelib("sImageType") = sImageType
	     rsimagelib.update    	 
     else 'smode = edit
		 'update Table
         sql = "SELECT stimgdesc,seimgdesc FROM bimagelib WHERE scorpcode='" & scorpcode
         sql = sql & "' and simgid='" & simgid & "'"
        
         rsimagelib.open sql,conn,adOpenKeySet,adLockPessimistic,adCmdtext
          If Not rsimagelib.EOF Then
			rsimagelib("stimgdesc") = stimgdesc
			rsimagelib("seimgdesc") = seimgdesc
			'rsimagelib("sImageType") = sImageType
	     rsimagelib.update            
	     End IF
     end if

	 rsimagelib.close
	 set rsimagelib = nothing
     response.redirect("bbImageGroup.asp?sCorpCode="&sCorpCode)
%>

<% 'check blank 
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

    sSql = "SELECT * FROM simagelib where sfilename = '" & Ucase(sfilename) & "'"
    set Rs = Server.CreateObject("ADODB.Recordset")
    Rs.Open sSql,Conn,adOpenKeySet,adLockPessimistic,adCmdtext
   
	if not Rs.eof then
      str = "bbimagehome.asp?"
      if not check2 then
           str = str & "stimgdesc=" & stimgdesc & "&"    
      end if
      if (trim(seimgdesc) <> "") then
           str = str & "seimgdesc=" & seimgdesc & "&"    
      end if
      str = str & "simgid=" & simgid & "&"
      str = str & "event=" & sMode & "&"
      str = str & "Repeat=" & true      
      Response.Redirect(str)               	
   	end if

   End Sub
%>
