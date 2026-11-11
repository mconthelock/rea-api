<%
dim testpath
testpath = request.serverVariables("APPL_PHYSICAL_PATH")
response.write "THIS IS " & testpath & "<BR>"
testpath = testpath & "procurement\workflow\ppritemdescImages"
response.write "THIS IS " & testpath & "<BR>"
%>