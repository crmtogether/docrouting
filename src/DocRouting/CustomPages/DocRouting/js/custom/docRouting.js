var win=null;
function openSendDocument() {	
	var _url = crm.url("DocRouting/sendDocument.asp") + "&PopupWin=Y";	
	console.log("????" + _url)
	//if (win) win.focus();
	win = window.open(_url, 'DocRouting', 'width=1200, height=800, scrollbars=yes,resizable=yes,status=yes,titlebar=yes, top=0, left=0');
}

crm.ready(function() {

	var enableForEntity = ["quote", "case", "opportunity", "company"];
	
	var keys = [1,2,6,7,8,86,58]; 
	/*
	Company=1; 
	Person=2; 
	Communication=6; 
	Opportunity=7; 
	Case=8; 
	Quote = 86;
	Custom Entity = 58
	*/	
	
	
	
	
	var T = crm.getArg("T",crm.url()).toLowerCase();
	var Key0 = parseInt(crm.getArg("Key0",crm.url()));
	
	if (($.inArray( T, enableForEntity) > -1) &&  ($.inArray( Key0, keys) > -1)  ) {		
		crm.addButton("Send.gif","Email " +  T + " Document", "Email " +  T + " Document","javascript:alert(2);openSendDocument();");
	}
});


