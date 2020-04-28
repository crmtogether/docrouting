<!-- #include file ="eWareNoHist.js" -->
<%



/*quote = CRM.FindRecord("Quotes" , "quot_orderquoteid=" + Request.QueryString("Key86"));
oppoId = Request.QueryString("oppoId");
oppo = eWare.FindRecord("Opportunity", "oppo_opportunityid=" + oppoId);
person = eWare.FindRecord("Person", "pers_personid=" + oppo.oppo_primarypersonid);
if (Defined(oppo.oppo_primarycompanyid))
	company = eWare.FindRecord("Company", "comp_companyid=" + oppo.oppo_primarycompanyid); 
else
	company = null;
	*/
	
var personId = Defined(Request.Form('personId')) ? Request.Form('personId') : null;
var companyId = Defined(Request.Form('companyId')) ? Request.Form('companyId') : null ;	
var recordId = Request.Form('recordId');
var entity = Request.Form('entity');
var T = Request.Form('T');


var sysParam = eWare.FindRecord("Custom_SysParams", "parm_name='DocTemplates'");
var LIB_PATH = sysParam.parm_value;
var docType = Request.QueryString("type");

var attachments = new String(Request.Form("attach")).split(',');
var fs = Server.CreateObject("Scripting.FileSystemObject");

try {

	var em_from = Request.Form('em_from');
	var em_to = Request.Form('em_to');
	var em_body = Request.Form('em_body');
	var em_subject = Request.Form('em_subject');
	
	var email=new Object();
	email.subject = em_subject;
	email.HTMLBody = em_body;

	
	////****************************************************************************
	////****************************** CUSTOMER SMTP SETTINGS***********************************
	////****************************************************************************
	
	var sql = "SELECT parm_value FROM Custom_SysParams WHERE parm_deleted IS NULL AND parm_name='ctos_docrouting_smtp_pwd'";
	Response.Write(sql);
	var execQuery = CRM.CreateQueryObj(sql);
	execQuery.SelectSql();
	var smtpPassword = "";
	if (execQuery.eof) {
		eWare.AddContent("SMTP Password not set. Please go to Administration->CRM Together OS->Doc Routing and set the Email SMTP Password.");
	} else {
		smtpPassword = DeCrypt(execQuery("parm_value"));
	}
	
	
	email.smtpserver=getParamValue('SMTPServer');   
	email.sendusername=getParamValue('SMTPUserName');
	email.smtpserverport= getParamValue('SMTPPort');  
	email.sendpassword=smtpPassword;
	email.EnableSsl=getParamValue('UseSecureSMTP') == "on" ?  "checked" : "";
	
	email.from=em_from;
	email.sender=em_from; 
	email.to = em_to; 

	////****************************************************************************
	////****************************************************************************
	////****************************************************************************
	
	
	

   var config = new ActiveXObject("CDO.Configuration");
   var sch = "http://schemas.microsoft.com/cdo/configuration/";   
   config.Fields.item(sch + "sendusing") = 2;// ' cdoSendUsingPort
   config.Fields.item(sch + "smtpserver") = email.smtpserver;//
   config.Fields.item(sch + "smtpserverport") = email.smtpserverport;
   config.Fields.item(sch + "smtpauthenticate") = 1;//'basic auth
   config.Fields.item(sch + "smtpusessl") = email.EnableSsl;
   config.Fields.item(sch + "sendusername") = email.sendusername;
   config.Fields.item(sch + "sendpassword") = email.sendpassword;
   config.Fields.item(sch + "smtpconnectiontimeout") = 60;
   config.Fields.item(sch + "smtpaccountname") = email.sendusername;       
   config.Fields.update();
   

	var myMail= new ActiveXObject("CDO.Message");
	myMail.configuration = config;

	   
	var hasAttach = "N";
	var docDir = '';
	
	var libRecords = [];
	if (Defined(attachments)) {
		for (var i =0; i< attachments.length; i++) {
			var lib = eWare.FindRecord("Library", "libr_libraryid=" + attachments[i]);
			libRecords.push(lib);
			f = LIB_PATH + lib.libr_filepath + "\\" + lib.libr_filename;
			if (fs.FileExists(f)) {
				myMail.AddAttachment(f);
				hasAttach = "Y";				
			}
		}
	}
	
   myMail.to = email.to;
   myMail.from = email.from;
   myMail.subject = email.subject;
   myMail.HTMLBody = email.HTMLBody;

   myMail.send();
} catch(e) {
	Response.Write("<table class='ErrorContent' width='100%'><tr><td>Error sending email:" + e.message + "</td></tr></table>");		
	Response.End();
}




	//create communication
	var cdate = new Date();
	var new_comm = eWare.CreateRecord( "Communication" );
	
	if (T != "person" && T != "company") {
		new_comm("Comm_" + T + "id")   = recordId;
	}
	
	new_comm("Comm_Type")   = "Task";
	new_comm("Comm_Action") = "EmailOut";
	new_comm("Comm_Email")  = em_body;
	new_comm("Comm_From")   = em_from
	new_comm("Comm_To")     = em_to;
	new_comm("Comm_Subject")     = em_subject;	
	new_comm("Comm_isHtml") = 'Y';
	new_comm("Comm_Status") = "Complete";	
	new_comm("Comm_datetime") = cdate.getVarDate();
	new_comm("Comm_Priority" )   = "Normal";
	new_comm("Comm_HasAttachments" ) = hasAttach; 			
	new_comm.SaveChanges();	
	
	
	//create comm link 
	var new_link = eWare.CreateRecord( "Comm_Link" );	              
	new_link("CmLi_Comm_PersonId") = personId;	
	new_link("CmLi_Comm_CompanyId") = companyId;  
	new_link("CmLi_Comm_CommunicationId" ) = new_comm.RecordId;
	new_link.CmLi_RecordId = recordId;
	new_link.SaveChanges();
	
	//create library records
	for (var i=0; i< libRecords.length; i++) {
		var libRecord = libRecords[i];
		var newlibRec = eWare.CreateRecord("Library");
		
		
		newlibRec("libr_"+ T +"id") = recordId;
		
		
		newlibRec("libr_type") = libRecord.libr_type;
		newlibRec("libr_status") = "Final";
		newlibRec("libr_category") = "Sales";
		newlibRec("libr_filepath") = libRecord.libr_filepath;
		newlibRec("libr_filename") = libRecord.libr_filename;
		newlibRec("libr_note") = libRecord.libr_note;
		newlibRec("libr_communicationId") = new_comm.RecordId;		
		newlibRec.SaveChanges();
		Response.Write(newlibRec.RecordId + "<br/>");
	}
	
	Response.Write('<script>alert("Email Sent");window.close();</script>');


function getParamValue(parmName) {
	var r = eWare.FindRecord("Custom_SysParams", "parm_name='"+parmName+"'");
	if (r.eof) return "";
	
	return r.parm_value;
}

%><script language="vbscript" runat="server">

   dim stringainput
'inputstring = "topofthemorningtoya@irishaccents.com"
'Response.Write "original string: " & inputstring &"<br>"
  
'dim cryptedstring
'cryptedstring = EnCrypt(inputstring )
'Response.Write "stringa criptata: " & cryptedstring &"<br>"
'Response.Write "stringa criptata legth: " & Len(cryptedstring) &"<br>"
  
'Response.Write "decrypted string: " & DeCrypt(cryptedstring) &"<br>"
  
Function EnCrypt(strCryptThis)
    Dim strChar, iKeyChar, iStringChar, i,g_Key, varSessionID
    varSessionID = "Ratamahatta"
g_keypos=0
for i=0 to len(strCryptThis)
g_Key=g_Key & mid(varSessionID,1,g_keypos)
g_keypos=g_keypos+1
if g_keypos>len(varSessionID) Then g_keypos=0
next
for i = 1 to Len(strCryptThis)
iKeyChar = Asc(mid(g_Key,i,1))
iStringChar = Asc(mid(strCryptThis,i,1))
iCryptChar = iKeyChar Xor iStringChar
iCryptCharHex = Hex(iCryptChar)
iCryptCharHexStr = cstr(iCryptCharHex)
if len(iCryptCharHexStr)=1 then iCryptCharHexStr = "0" & iCryptCharHexStr
strEncrypted = strEncrypted & iCryptCharHexStr
next
EnCrypt = strEncrypted
End Function
  
Function DeCrypt(strEncrypted)
Dim strChar, iKeyChar, iStringChar, i,g_Key, varSessionID
varSessionID = "Ratamahatta"
LenGKey=Len(strEncrypted)/2
g_keypos=0
For i=0 to LenGKey
g_Key=g_Key & mid(varSessionID,1,g_keypos)
g_keypos=g_keypos+1
if g_keypos>len(varSessionID) Then g_keypos=0
Next
for i = 1 to Len(strEncrypted) /2
iKeyChar = (Asc(mid(g_Key,i,1)))
iStringChar2 = mid(strEncrypted,(i*2)-1,2)
iStringChar = CLng("&H" & iStringChar2)
iDeCryptChar = iKeyChar Xor iStringChar
strDecrypted = strDecrypted & chr(iDeCryptChar)
next
DeCrypt = strDecrypted
End Function
  
'Function ReadKeyFromFile()
  
'Const strFileName = "C:\Inetpub\wwwroot\key.txt" 'Change this string
'Dim keyFile, fso, f
'set fso = Server.CreateObject("Scripting.FileSystemObject") 
'set f = fso.GetFile(strFileName) 
'set ts = f.OpenAsTextStream(1, -2)
'Do While not ts.AtEndOfStream
'keyFile = keyFile & ts.ReadLine
'Loop 
'ReadKeyFromFile = keyFile
'End Function
   
</script>