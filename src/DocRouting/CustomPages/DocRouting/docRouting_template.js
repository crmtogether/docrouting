var win=null;
function openSendDocument() {	
	var _url = crm.url("DocRouting/sendDocument.asp") + "&PopupWin=Y";			
	win = window.open(_url, 'DocRouting', 'width=1200, height=800, scrollbars=yes,resizable=yes,status=yes,titlebar=yes, top=0, left=0');
}

crm.ready(function() {
	var enableForEntity = [!ENTITIES];
	var T = crm.getArg("T",crm.url()).toLowerCase();
	var key = crm.getArg('Key0', crm.url());	
	if ($.inArray( T, enableForEntity) > -1) {		
		crm.addButton("Send.gif","Email " +  T + " document", "Email " +  T + " Document","javascript:openSendDocument()");
	}
});