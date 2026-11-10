
var counter = 0; 
var maxSize = 0;
var checkAttFile = false;
	function addUpload(seq, tmaxSize) 
	
	{ 
	
	//	var divs = document.getElementsByTagName("DIV");
		maxSize = tmaxSize;
		var fileName = "uploadFile"+seq ;
		
		var div = document.createElement('DIV'); 
		
		div.innerHTML = '<input id="'  + fileName+counter + '" name = "'  +fileName+ counter + 
	
						 '" type="file"  onchange = "LimitedSize(this)" size="60"/>' + 
	
						 '<input id="Button' + counter + '" type="button" ' + 
	
						'value="Remove" onclick = "RemoveFileUpload(this)" />'; 
		document.getElementById(fileName).appendChild(div); 
		 counter++;
		 div.onchange = "";
	
	} 

	function RemoveFileUpload(div) 
	{ 
		 div.parentNode.parentNode.removeChild(div.parentNode)
	} 

	function LimitedSize(div) 
	{ 
	var i; 
	//var oas = new ActiveXObject("Scripting.FileSystemObject");
	var y = div.value
	//var e = oas.getFile(y);
	//var f = (e.size);
	//var fname = oas.GetFileName(y);
	//f =(f/(1024*1024));
	//f =(f/(1024));
	//var n = div.name;
	//var fld = document.getElementById(n);
	var n = div.name;
	var f = div.files[0].size/1024;
	var fname = div.files[0].name;
		if (f > maxSize){
			alert ("File Size is more than " + maxSize + " KB");
			div.outerHTML = '<input name="'+ n +'" value="" type="file" onChange="LimitedSize(this)"  size="60">' ;
		}
		if(fname.search("'")>-1)
		{
			alert ("File name can not contain ' ");
			div.outerHTML = '<input name="'+ n +'" value="" type="file" onChange="LimitedSize(this)"  size="60">' ;
		
		}
		if ((n.substr(10,1)) == 1) {
			checkAttFile = true;
		} 
	} 
