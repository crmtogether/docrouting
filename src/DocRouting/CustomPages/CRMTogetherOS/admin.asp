<!-- #include file ="../sagecrm.js" -->
<%   
var forReading = 1, forWriting = 2, forAppending = 8;
var newline="<br>";
var newlinetext="\n";

//cryptedstring = EnCrypt("abcdef" )
//Response.Write(cryptedstring);
//Response.Write(DeCrypt(cryptedstring));	


//get the setting from custom_sysparams
var sql="select * from custom_sysparams where parm_name='ctos_docrouting' and parm_deleted is null";
var sqlq=CRM.CreateQueryObj(sql);
sqlq.SelectSQL();

if (CRM.Mode==2)
{
  //save....insert or update
  if (sqlq.eof)
  {
    //insert
	var idb=CRM.CreateRecord("custom_sysparams");
	idb("parm_name") = "ctos_docrouting";
	idb("parm_value") = Request.Form("fieldstoflag");
	idb.SaveChanges();
  }else{
    //update
	var idb=CRM.FindRecord("custom_sysparams","parm_name='ctos_docrouting' and parm_deleted is null");
	idb("parm_value")=Request.Form("fieldstoflag");
	idb.SaveChanges();	
  }
  sql="select * from custom_sysparams where parm_name='ctos_docrouting' and parm_deleted is null";
  sqlq=CRM.CreateQueryObj(sql);
  sqlq.SelectSQL();
  //update the file and replace [!ENTITIES] with our code
  var fs=Server.CreateObject("Scripting.FileSystemObject");
  var TemplateFileName=Request.ServerVariables("APPL_PHYSICAL_PATH")+"\CustomPages\\DocRouting\\docRouting_template.js";
  var NewFileName=Request.ServerVariables("APPL_PHYSICAL_PATH")+"\\js\\custom\\docRouting.js";
  var lineData="";
  if (fs.FileExists(TemplateFileName)==true)
  {
		var f = fs.CopyFile(TemplateFileName, NewFileName);    
		// Open the file 
		var is = fs.OpenTextFile(NewFileName, forReading, true );
		// start and continue to read until we hit
		// the end of the file. 
		while( !is.AtEndOfStream ){
			lineData  += is.ReadLine()+newlinetext;
		}
		// Close the stream 
		is.Close();
		lineData=lineData.replace("[!ENTITIES]",sqlq("parm_value"));
		var flOutput = fs.CreateTextFile(NewFileName, true); //true for overwrite
		flOutput.Write(lineData);
		flOutput.Close();
  }
  
	//set custom smtp password 
  
	//see if we haev the pasword param 
	
	var sql="select * from custom_sysparams where parm_name='ctos_docrouting_smtp_pwd' and parm_deleted is null";
	var smtp_sqlq=CRM.CreateQueryObj(sql);
	smtp_sqlq.SelectSQL();
	if (smtp_sqlq.eof)
	{		
		var idb=CRM.CreateRecord("custom_sysparams");
		idb("parm_name") = "ctos_docrouting_smtp_pwd";		
		idb.SaveChanges();
	}

	
	
	if(Request.Form("SMTPPassword") != "") {		
		encpass = EnCrypt(Request.Form("SMTPPassword"));
	} else {
		encpass="";
	}
	
	var sql_updated_pass = "UPDATE TOP(1) Custom_SysParams set parm_value =  '"+ encpass +"' WHERE parm_deleted IS NULL AND parm_name='ctos_docrouting_smtp_pwd'";	
	var execQuery = CRM.CreateQueryObj(sql_updated_pass);
	execQuery.ExecSql();
		
  
  
  
  
} else if (CRM.Mode==3) {
  //reset
  var idb=CRM.FindRecord("custom_sysparams","parm_name='ctos_docrouting' and parm_deleted is null");
  idb.DeleteRecord = true;
  idb.SaveChanges();	
  sql="select * from custom_sysparams where parm_name='ctos_docrouting' and parm_deleted is null";
  sqlq=CRM.CreateQueryObj(sql);
  sqlq.SelectSQL();
}

var CurrentUser=CRM.GetContextInfo("user", "User_logon");
CurrentUser=CurrentUser.toLowerCase();
var CurrentUserName=CRM.GetContextInfo("user", "User_firstname");
var CurrentUserEmail=CRM.GetContextInfo("user", "User_emailaddress");

Container=CRM.GetBlock("container");
Container.DisplayButton(Button_Default) =false;
Container.DisplayForm = true;

content=eWare.GetBlock('content');
content.contents = newline+"Hi "+CurrentUserName;
content.contents +=newline+"Welcome to the CRM Together Open Source page for the <b>CRM Document Routing</b>.";
content.contents +=newline+"This will enable your system to send documents";
content.contents +=newline+"Regards";
content.contents +=newline+"The CRM Together Open Source Team"+newline+newline;

Container.AddBlock(content);

btnUpdate= eWare.Button("Update", "", "javascript:updateinfo()");
Container.AddButton(btnUpdate);
	 
btnSend = eWare.Button("Email CRM Together", "", "mailto:sagecrm@crmtogether.com?subject=CRM Together Open Source Document Routing");
Container.AddButton(btnSend);

if (!sqlq.eof)
{
	btnReset= eWare.Button("Reset", "", "javascript:resetallhtmlflags()");
	Container.AddButton(btnReset);
}

content.contents+="<textarea class=\"EDIT\" name=\"fieldstoflag\" id=\"fieldstoflag\" rows=\"20\" cols=\"90\">"
if (sqlq.eof)
{
  //put in sample code
  content.contents+="ARRAY STRUCTURE. EDIT AS NEEDED AND REMOVE THIS LINE"+newlinetext+newlinetext;
  content.contents+='["quote", "case", "opportunity"]';
}else{
  content.contents+=sqlq("parm_value");
}
content.contents+="</textarea>";


SMTPServer = getParamValue('SMTPServer');
SMTPUserName = getParamValue('SMTPUserName');
SMTPPort = getParamValue('SMTPPort');
SMTPPassword = DeCrypt(getParamValue('ctos_docrouting_smtp_pwd'));
UseSecureSMTP = getParamValue('UseSecureSMTP') == "on" ?  "checked" : "";


content.contents += "<div>" + newline + "<strong>E-mail Configuration</strong></div><p></p>";

content.contents += "<div class='InfoContent'><strong>To change the SMTP values go to the admin area in CRM. Only set your password here</strong></div><br/>";

content.contents += "<table>";
content.contents += '<tr>';
content.contents += '<td valign="TOP"><span id="_CaptSMTPServer" class="VIEWBOXCAPTION">Outgoing mail server (SMTP):</span><br>' +
		'<span id="_DataSMTPServer" class="VIEWBOX"><input type="text" readOnly class="EDIT" id="SMTPServer" name="SMTPServer" value="'+SMTPServer+'" maxlength="50" size="25"></td>';

content.contents += '<td valign="TOP"><span id="_CaptSMTPPort" class="VIEWBOXCAPTION">SMTP port:</span><br><span id="_DataSMTPPort" class="VIEWBOX">' + 
					'<input type="text"  readOnly class="EDIT" id="SMTPPort" name="SMTPPort" value="'+SMTPPort+'" maxlength="50" size="10"></span></td>';
content.contents += '<td><span class="VIEWBOXCAPTION" id="_CaptUseSecureSMTP"><label for="_IDUseSecureSMTP">Use TLS for SMTP</label>:</span>' + 
					'<br><span class="VIEWBOXCAPTION" id="_DataUseSecureSMTP"><input onclick="return false;" type="checkbox" name="UseSecureSMTP" checked="'+UseSecureSMTP+'" id="_IDUseSecureSMTP"></span></td>';
content.contents += '</tr>';
content.contents += '<tr>';
content.contents += '<td valign="TOP"><span id="_CaptSMTPUserName" class="VIEWBOXCAPTION">SMTP User Name:</span><br><span id="_DataSMTPUserName" class="VIEWBOX">';
content.contents += '<input type="text" readOnly class="EDIT" id="SMTPUserName" name="SMTPUserName" value="'+SMTPUserName+'" maxlength="50" size="25"></span></td>';
content.contents += '<td valign="TOP"><span id="_CaptSMTPPassword" class="VIEWBOXCAPTION">SMTP Password:</span><br><span id="_DataSMTPPassword" class="VIEWBOX">';
content.contents += '<input type="password" class="EDIT" id="SMTPPassword" name="SMTPPassword" value="'+SMTPPassword+'" maxlength="24" size="25" autocomplete="off"></span></td>';
content.contents += '</tr>';



content.contents += "</table>";



CRM.AddContent(Container.Execute());

Response.Write(CRM.GetPage('CRMTogetherOS'));


function getParamValue(parmName) {
	var r = eWare.FindRecord("Custom_SysParams", "parm_name='"+parmName+"'");
	if (r.eof || r.parm_value == null) return "";
	
	return r.parm_value;
}

%>

<script>
//ref: https://stackoverflow.com/questions/3710204/how-to-check-if-a-string-is-a-valid-json-string-in-javascript-without-using-try
function IsJsonString(str) {
    try {
        JSON.parse(str);
    } catch (e) {
        return false;
    }
    return true;
}

function updateinfo(){
	if (window.confirm("Do you really want to update? This cannot be undone")) { 
	  document.getElementsByName("em")[0].value="2";
	  document.forms[0].submit();
	}
}

function resetallhtmlflags()
{
	if (window.confirm("Do you really want to reset? This cannot be undone")) { 
	  document.getElementsByName("em")[0].value="3";
	  document.forms[0].submit();
	}
}





</script>

<script language="vbscript" runat="server">

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
