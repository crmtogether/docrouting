var win=null;
function openSendDocument() {	
	var _url = crm.url("DocRouting/sendDocument.asp") + "&PopupWin=Y";			
	win = window.open(_url, 'DocRouting', 'width=1200, height=800, scrollbars=yes,resizable=yes,status=yes,titlebar=yes, top=0, left=0');
}

crm.ready(function() {
	var enableForEntity = [!ENTITIES];
	var T = crm.getArg("T",crm.url()).toLowerCase();
	if (T=="")
	{
		var RecentValue=crm.getArg('RecentValue', crm.url());	
		var _url=new String(crm.url());
		if (RecentValue!="")
		{
			if (_url.indexOf("RecentValue=200X")>0)
			{
				T="company";
			}else 
			if (_url.indexOf("RecentValue=220X")>0)
			{
				T="person";
			}else 
			if (_url.indexOf("RecentValue=200X")>0)
			{
				T="company";
			}else 
			if (_url.indexOf("RecentValue=281X")>0)
			{
				T="case";
			}else 
			if (_url.indexOf("RecentValue=260X")>0)
			{
				T="opportunity";
			}else 
			if (_url.indexOf("RecentValue=1469X")>0)
			{
				T="quote";
			}else 
			if (_url.indexOf("RecentValue=1463X")>0)
			{
				T="order";
			}else 
			if (_url.indexOf("RecentValue=1284X")>0)
			{
				T="solutions";
			}				
		}
	}
	var key = crm.getArg('Key0', crm.url());	
	if ($.inArray( T, enableForEntity) > -1) {		
		crm.addButton("Send.gif","Email " +  T + " document", "Email " +  T + " Document","javascript:openSendDocument()");
	}
});