<% Option Explicit %>
<!-- #include file ="../adovbs.inc"  -->
<!-- #include file ="../Conn.asp" -->

<%
     dim objsmartupload,rsimagelib 	                      'for obj
     dim spath,fd,file,sql,smode,check1,check2            'other variable
     dim sprno,sfilename,nitem,sitemno   'for variable in table     
	  dim sImageType,sCatID,sCurcode,sUnitcode,sAcccode
     Dim sprodcode,sexptype,svendcode,sprodid,stprodname,seprodname,smodel,stdesc,sdesc,nprice,sfile,dadddate,sempno
     set objsmartupload = server.createobject("aspSmartUpload.SmartUpload")'create Obj     
     objsmartupload.MaxFileSize = 2500000                                   'Max filesize 
     objsmartupload.Upload                                     'upload to objsmartupload

     ObjSmartUpload.DeniedFilesList = "bat,exe,com,asp,dll"    

     smode = objsmartupload.form("EVENT")
     'data from pre page      	
     if smode = "ADD" then
		sprno=objsmartupload.Form("sprno")
		nItem=objsmartupload.Form("nItem")
		sItemno=objsmartupload.Form("sItemno")
		sdesc = objsmartupload.Form("sdesc")
		sfile=objsmartupload.Form("sfile")
     else
		sprno=objsmartupload.Form("sprno")
		nItem=objsmartupload.Form("nItem")
		sItemno=objsmartupload.Form("sItemno")
		sdesc = objsmartupload.Form("sdesc")
		sfile=objsmartupload.Form("sfile")
     end if
   
	 spath = request.serverVariables("APPL_PHYSICAL_PATH")      'ª×èÍ ·Ò§physic ¢Í§ server
    spath = spath & "procurement\workflow\ppritemdesctImages"          'check path corp

     fd = objsmartupload.save(spath)           'save file to server at spath

	 for each file in objsmartupload.files
	 	 sfilename = file.filename             'ª×èÍ¢Í§ file ·Õèload ÁÒ
 	 next	
     'If not ObjSmartUpload.Files.IsMissing Then
     'sFileName = ObjSmartUpload.Files.FileName
     'End if                                   'samesame

'	 check_blank()
'     check_filename() 

     set rsimagelib = server.createobject("ADODB.Recordset")

		 'update Table
         sql = "SELECT * from ppritem WHERE sprno='" & sprno & "'" &_
         " and nitem='" & nitem & "'" 
         rsimagelib.open sql,conn,adOpenKeySet,adLockPessimistic,adCmdtext
          If Not rsimagelib.EOF Then
		if sfilename<>"" then
			rsimagelib("simagespec") = sfilename
		end if
	     rsimagelib.update            
	     End IF

	 rsimagelib.close
	 set rsimagelib = nothing
     'response.redirect("../datamanager/pproductdesc.asp")
     'Response.Redirect "../datamanager/pproductdesc.asp?sprodcode="&sprodcode&"&sexptype="&sexptype&"&sCatid="&sCatID
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

    sSql = "SELECT * FROM pproduct where simage = '" & Ucase(sfilename) & "'"
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
