<!-- #include file ="eWareNoHist.js" -->
<LINK REL="stylesheet" HREF="/<%=sInstallName%>/Themes/ergonomic.css">

<style>

#modal-background {
        display: none;
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background-color: gray;
        opacity: .50;
        -webkit-opacity: .5;
        -moz-opacity: .5;
        filter: alpha(opacity=50);
        z-index: 1000;
    }
    
    #userlist {
        background-color: white;
        border-radius: 10px;
        -webkit-border-radius: 10px;
        -moz-border-radius: 10px;
        box-shadow: 0 0 20px 0 #222;
        -webkit-box-shadow: 0 0 20px 0 #222;
        -moz-box-shadow: 0 0 20px 0 #222;
        display: none;
        height: 500px;
        left: 25%;
        margin: -120px 0 0 -160px;
        padding: 10px;
        position: absolute;
        top: 43%;
        width: 600px;
        z-index: 1000;
    }
	#personlist {
        background-color: white;
        border-radius: 10px;
        -webkit-border-radius: 10px;
        -moz-border-radius: 10px;
        box-shadow: 0 0 20px 0 #222;
        -webkit-box-shadow: 0 0 20px 0 #222;
        -moz-box-shadow: 0 0 20px 0 #222;
        display: none;
        height: 500px;
        left: 25%;
        margin: -120px 0 0 -160px;
        padding: 10px;
        position: absolute;
        top: 43%;
        width: 600px;
        z-index: 1000;
    }

    #modal-background.active, #userlist.active {
        display: block;
    }
	#modal-background.active, #personlist.active {
        display: block;
    }

</style>

<script type="text/javascript" src="/<%=sInstallName%>/js/lib/jquery.min.js"></script>
<script type="text/javascript" src="/<%=sInstallName%>/js/lib/jquery-ui.min.js"></script>
<script type="text/javascript" src="/<%=sInstallName%>/ckeditor/ckeditor.js"></script>
<script type="text/javascript" language="javascript">
var loadingCompleted = false;
var sBasePath = "/<%=sInstallName%>/ckeditor/";
var sBaseHref = "http://localhost/<%=sInstallName%>/eware.dll/";
CKEDITOR.on('instanceCreated', function(e) {
	CKEDITOR.basePath	= sBasePath;
	CKEDITOR.BaseHref	= sBaseHref;
	CKEDITOR.on('onkeyup', function(e) {SageCRM.wsEmailClient.saveCaretEmailUseHTMLMail;});
	CKEDITOR.on('onselect', function(e) {SageCRM.wsEmailClient.saveCaretEmailUseHTMLMail;});
	CKEDITOR.on('onclick', function(e) {SageCRM.wsEmailClient.saveCaretEmailUseHTMLMail;});
	CKEDITOR.on('onpaste', function(e) {SageCRM.wsEmailClient.saveCaretEmailUseHTMLMail;});
	if(typeof(EditSource)!= "undefined" ){
		if (EditSource.value == "Y") {
			oEditor.Commands.GetCommand("Source").Execute();
		}
	}
	loadingCompleted = true;
	});
	
	$(document).ready( function(){
	
	togglediv();
		CKEDITOR.replace('em_body',
	{
		language : 'en',
		height : 200
	});
	
	

	});
	
	function clearInput(obj)
	{
	   $('#'+obj).val('');
	   return false;
	}
	function togglediv(obj)
	{
	//  $('#personlist').hide();
	 // $('#userlist').hide();
	  if (obj)
	  {
	     $("#"+obj+",#modal-background").toggleClass("active");
		 // $('#'+obj).toggle();		
	  }
	  return false;
	}
	function addToEmail(_sender)
	{
	  //alert(_sender.value);
	  var _curval=$('#em_to').val();
	  if (_curval!="")
	  {
	    _curval+=";"
	  }
	  $('#em_to').val(_curval+_sender.value);
	}
</script>
<style>
body {

overflow:auto;
}
</style>
<%


PRODUCTION = true;
MERGE_SUBJECT = true;
MERGE_EMAIL_BODY =  true;


EMAIL_BODY = '';
EMAIL_SUBJECT = '';
EMAIL_FROM = '';
EMAIL_TO='';

//----------------------------------
//important to leave this at the start of the script 
if (!String.prototype.trim) {
  String.prototype.trim = function () {
    return this.replace(/^\s+|\s+$/g, '');
  };
}
//----------------------------------


var key = Request.QueryString("Key0");
var T = new String(Request.QueryString("T")).toLowerCase();
var entity = T;

var _url=new String(Request.QueryString);
if ((T=="")||(T+""=="undefined"))
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

switch (T) {
	case "quote" : 
		entity = "quotes";
		break;
	case "case" : 
		entity = "cases";
		break;
	default:
		entity = T;
}

//get entity info

var tableInfo = CRM.FindRecord("Custom_Tables", "bord_name='" + entity + "'");
var idField = tableInfo.bord_idfield;
var prefix = new String(tableInfo.bord_prefix).toLowerCase();
var record = CRM.FindRecord(entity , idField + "=" + Request.QueryString("Key" + key));

//get person and company details
var recordSummary = eWare.FindRecord(entity + ",vSummary" + T, idField + "=" + record.RecordId);

if (entity == "quotes") {
	var personId = recordSummary("oppo_primarypersonid");
	var companyId= recordSummary("oppo_primarycompanyid");
} else if (entity != "person" && entity != "company") {
	var personId = recordSummary(prefix + "_primarypersonid");
	var companyId= recordSummary(prefix + "_primarycompanyid");
} else if (entity == "person") {
	var personId = CRM.GetContextInfo("Person", "pers_personid");
	var companyId = CRM.GetContextInfo("Person", "pers_companyid");
} else if (entity = "company") {
	var personId = CRM.GetContextInfo("Company", "comp_primarypersonid");
	var companyId = CRM.GetContextInfo("Company", "comp_companyid");
}

var person = CRM.FindRecord("Person", "pers_personid=" + personId);
var _persEmail = CRM.CreateQueryObj("select pers_emailaddress from vPErsonPE where pers_personid="+personId+" and pers_emailaddress is not null");
_persEmail.SelectSql();
		
var _compEmail = CRM.CreateQueryObj("select comp_emailaddress from vCompanyPE where comp_companyid="+companyId+" and comp_emailaddress is not null");
_compEmail.SelectSql();	
		


//Response.Write(record.RecordId);
//Response.End();

/*
oppoId = Defined(Request.QueryString("oppoId")) ?  Request.QueryString("oppoId") : Request.QueryString("Key7");

oppo = eWare.FindRecord("Opportunity", "oppo_opportunityid=" + oppoId);
if (oppo.oppo_primarypersonid == null) {
	Response.Write('Opportunity does not have a person associated');
	Response.End();
}
*/


var user = eWare.FindRecord("Users", "User_userid="  + eWare.GetContextInfo("user", "user_userid"));

var Container=CRM.GetBlock("container");
Container.DisplayButton(Button_Default) =false;
Container.DisplayForm = false;



var sTitle = 'Send Document';

    if (PRODUCTION)
	{
		
		if (!_persEmail.eof)
		{
			EMAIL_TO=_persEmail("pers_emailaddress");
		}
		  		
		if (!EMAIL_TO)
		{
			EMAIL_TO="No email address found for this person:"+person.pers_firstname+" "+person.pers_lastname;
		}
	} else {
		EMAIL_TO = "";
	}
	
	
	

//-------------------------------------
if (MERGE_SUBJECT  || MERGE_EMAIL_BODY) {


	emailDetails = getEmTemplate("Send " + T + " Document");
	if (emailDetails.eof)
	{
	  emailDetails.emte_comm_note='EXPECTED '+'Send ' + T + ' Document Email Template';
	}
	emailDetails.emte_comm_from = CRM.GetContextInfo("User","user_emailaddress");
	EMAIL_FROM = emailDetails.emte_comm_from;

	//------------------------------------- Merge email subject
	if (MERGE_SUBJECT ) {
		
		EMAIL_SUBJECT = (emailDetails.emte_comm_note != null) ? emailDetails.emte_comm_note : "";
		hashFields = getHashFields(EMAIL_SUBJECT);
		EMAIL_SUBJECT = doMerge(EMAIL_SUBJECT, hashFields, record, prefix + "_");
		//EMAIL_SUBJECT = doMerge(EMAIL_SUBJECT, hashFields, person, 'pers_');
		//EMAIL_SUBJECT = doMerge(EMAIL_SUBJECT, hashFields, company, 'comp_');
		EMAIL_SUBJECT = doMerge(EMAIL_SUBJECT, hashFields, user, 'user_');
		//EMAIL_SUBJECT = doMerge(EMAIL_SUBJECT, hashFields, quote, 'quot_');
	}
	//------------------------------------- Merge Email body 
	if (MERGE_EMAIL_BODY) {	
		EMAIL_BODY = Defined(emailDetails.emte_comm_email) ? emailDetails.emte_comm_email : ""; 
		hashFields = getHashFields(EMAIL_BODY);
		EMAIL_BODY = doMerge(EMAIL_BODY, hashFields, record, prefix);
		//EMAIL_BODY = doMerge(EMAIL_BODY, hashFields, person, 'pers_');
		//EMAIL_BODY = doMerge(EMAIL_BODY, hashFields, company, 'comp_');
		EMAIL_BODY = doMerge(EMAIL_BODY, hashFields, user, 'user_');
		//EMAIL_BODY = doMerge(EMAIL_BODY, hashFields, quote, 'quot_');
	}
} // End if merge
//-------------------------------------
i=0;




var libraryField = "Libr_"+ T + "Id";
var libSql = "select top 20 * FROM vLibrary  WITH (NOLOCK)  WHERE  " + 
	" COALESCE(Libr_Global, N'') <> N'Y' AND  ((COALESCE(Libr_Private, N'') = N'') OR ((COALESCE(Libr_Private, N'') <> N'') " + 
	" AND (Libr_UserId = " + eWare.GetContextInfo("User", "user_userid")+ ")))";
libSql += " AND "+libraryField+" =" + record.RecordId + " ORDER BY libr_createddate desc";
//Response.Write(libSql);	
docs = eWare.CreateQueryObj(libSql);
docs.SelectSql();




var sTable = "<strong>" + sTitle + "</strong><form name='emForm' id='emForm' action='"+eWare.URL("DocRouting/sendEmail.asp?entity=" + entity + "&recordId=" + record.RecordId) +"' method='POST'>";
sTable += getEmailForm();
sTable += "<table class='content' width='100%'>";
sTable += '<tr class="gridhead"><td>Attach</td><td>Document Name</td><td>Type</td><td>Description</td></tr>';




while (!docs.eof) {

	rowClass = (i%2 == 0) ? 'ROW1' : 'ROW2';
	sTr ='<tr class="'+rowClass+'">';
	sTr += '<td><input type="checkbox" value="'+docs("libr_libraryid")+'" name="attach"/></td>';
	sTr += '<td>' + getVal(docs("libr_filename"))+ '</td>';
	sTr += '<td>' + getVal(docs("libr_type"))+ '</td>';
	sTr += '<td>' + getVal(docs("libr_note")) + '</td>';
	sTr += '</tr>';
	
	sTable += sTr;
	i++;
	docs.NextRecord();
}

sTable += '</table>';


//

sTable += "<input type='hidden' name='personId' value='"+personId+"' />";
sTable += "<input type='hidden' name='companyId' value='"+companyId+"' />";
sTable += "<input type='hidden' name='recordId' value='"+record.RecordId+"' />";
sTable += "<input type='hidden' name='entity' value='"+entity+"' />";
sTable += "<input type='hidden' name='T' value='"+T+"' />";
sTable += '</form>';



content=eWare.GetBlock('content');
content.contents = sTable;
Container.AddBlock(content);

//MR 4th Oct 16-removed permissions on the button
//btnSend = eWare.Button("Send", "", "javascript:document.emForm.submit();", "VIEW", "Opportunity");
btnSend = eWare.Button("Send", "", "javascript:if (confirm('Send email?')) { document.emForm.submit(); }");
Container.AddButton(btnSend);

//btnSend = eWare.Button("Close", "", "javascript:window.close();", "VIEW", "Opportunity");
btnSend = eWare.Button("Close", "", "javascript:window.close();");
Container.AddButton(btnSend);


Response.Write(Container.Execute());

//----------------------------------------

function getVal(s) {
	return (Defined(s) ? s: "");
}

//----------------------------------------


function getEmailForm() {

	var s = '<div id="modal-background"></div><table width="100%">';

	s += '<tr><td>';
	s += '<br/><span class="VIEWBOXCAPTION">From:</span><br/>';
	s += '<input style="width:90%;" size="60" type="text" class="EDIT" name="em_from" id="em_from" value="' +EMAIL_FROM+ '"/>';
	s += '</td></tr>';


	s += '<tr><td>';
	s += '<br/><span class="VIEWBOXCAPTION">To:</span><br/>';
	s += '<textarea style="width:90%;height: 75px;" class="EDIT" type="text" name="em_to" id="em_to" >' +EMAIL_TO + '</textarea>';
    s += '<button style="height:20px" onclick="return clearInput(\'em_to\')" >X</button>';	
    s += '</td></tr>';
	
	//may 16 - new lists to select from
	s += '<tr><td>';
    s += '<button style="height:20px" onclick="return togglediv(\'userlist\')" >User Email List</button>&nbsp;';	
    s += '<button style="height:20px" onclick="return togglediv(\'personlist\')" >Company Email List</button>';	

	s += '<div class="modal-content" id="userlist" >';
	s+='<span class="VIEWBOXCAPTION">User List - Double click to add to the TO list &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>';
	s+='<button id="modal-close" onclick="return togglediv(\'userlist\')" >Close</button>'
	s+='<select style="width:99%;height:96%;" size="20" ondblclick ="addToEmail(this)">';
	var userql='select User_EmailAddress, User_FirstName, User_LastName from users where user_deleted is null and User_EmailAddress is not null order by User_firstName,User_LastName';
	var userq=CRM.CreateQueryObj(userql);
	userq.SelectSQL();
	
    try{
		while (!userq.eof)
		{
			var _fname="";
			if (userq('user_firstname'))
			{
				_fname=userq('user_firstname');
			}
			_fname=_fname.trim();
			var _lname=userq('user_lastname');
			_lname=_lname.trim();
			s+='<option value="'+userq('user_emailaddress')+'" >'+_fname+' ' +_lname+'('+userq('user_emailaddress')+')</option>';
			userq.NextRecord();
		}
    }catch(e) {
		Response.Write(e);
		Response.Write(userql);
		Response.End();
	}
	
	
	s += '</select></div>';	
	s += '<div  class="modal-content" id="personlist" >';
	
	s+='<span class="VIEWBOXCAPTION">Company Person List - Double click to add to the TO list &nbsp;&nbsp;&nbsp;&nbsp;&nbsp</span>';
	s+='<button id="modal-close" onclick="return togglediv(\'personlist\')" >Close</button>'
	s+='<select style="width:99%;height:96%;" size="20" ondblclick ="addToEmail(this)">';
	
	
	var persql = "SELECT pers_firstname, pers_lastname,pers_emailaddress FROM vSummaryPerson WHERE Pers_PersonId=" +  personId+" ORDER BY pers_firstname, pers_lastname";
	var personq=CRM.CreateQueryObj(persql);
	personq.SelectSQL();
	while (!personq.eof)
	{
	  var _fname=personq('pers_firstname');
	  _fname=_fname.trim();
	  var _lname=personq('pers_lastname');
	  _lname=_lname.trim();
	  s+='<option value="'+personq('pers_emailaddress')+'" >'+_fname+' ' +_lname+'('+personq('pers_emailaddress')+')</option>';
	  personq.NextRecord();
	}
	
	s += '</select></div>';	
	s += '</td></tr>';

	s += '<tr><td>';
	s += '<br/><span class="VIEWBOXCAPTION">Subject:</span><br/>';
	s += '<input style="width:90%;" size="60" class="EDIT" type="text" name="em_subject" id="em_subject" value="' +EMAIL_SUBJECT + '"/>'; 
    s += '<button style="height:20px" onclick="return clearInput(\'em_subject\')" >X</button>';	
	s += '</td></tr>';
	

	s += '<tr><td>';
	s += '<br/><span class="VIEWBOXCAPTION">Message:</span><br/>';
	s += '<textarea cols=60 class="EDIT" rows=10 name="em_body" id="em_body">' + EMAIL_BODY + '</textarea>'; 
	s += '</td></tr>';
	s+='</table>';	
	
	
	
	
	
	return s; 
}

function getEmTemplate(emteName) {
	var query = eWare.FindRecord("EmailTemplates", "emte_name='" + emteName + "'");	
	//Response.Write("emte_name='" + emteName + "'");
	return query;	
}

//does a replace of ##fields with record field value
function doMerge(str, hashFields, _record, pre) {	
	s = str;	
	for (var j =0; j< hashFields.length; j++) {
		fieldName = hashFields[j];				
		if (fieldName.indexOf(pre) == 0 ) {			
			if (hashFields[j].toLowerCase() == fieldName) {
				

				try {
					v = Defined(_record(fieldName)) ? _record(fieldName) : "";
				
				} catch (ex) {
					//field found in the template does not exist on the record
					v = "#" + fieldName + "#";
				}				
				s = s.replace("#" + hashFields[j] + "#", v);							
			}
		}
	}	
	return s;
}


function getHashFields(str) {
	retArr = [];
	var re  = /#(\w*)#/g; 
	var m; 
	while ((m = re.exec(str)) !== null) {
		if (m.index === re.lastIndex) {
			re.lastIndex++;
		}
		retArr.push(m[0].replace(/#/g, ""));	
	}
	return retArr;
}
%>







