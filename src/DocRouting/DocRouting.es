/*

Screen Objects that have been added or updated

*/

ObjectName='CRMTogetherOS';
ObjectType='TabGroup';
EntityName='system';
var CObjId10865 = AddScreenObject();

/*

Tab groups that have been added or updated

*/

var TabsId11081 = AddCustom_Tabs(0,0,11,'Admin','CRM Together OS','customfile','CRMTogetherOS/admin.asp','','waves.gif',0,'',false,0);

var TabsId11082 = AddCustom_Tabs(0,0,1,'CRMTogetherOS','Home','customfile','CRMTogetherOS/admin.asp','','',0,'',false,0);

var TabsId10937 = AddCustom_Tabs(0,0,4,'CRMTogetherOS','Document Routing','customfile','DocRouting/admin.asp','','',0,'',false,0);

//copy files
var CRMTogetherOS="CRMTogetherOS";
CreateNewDir(GetDLLDir() + '\\CustomPages\\' + CRMTogetherOS);
CopyASPTo(CRMTogetherOS+'\\admin.asp','\\CustomPages\\'+CRMTogetherOS+'\\admin.asp');
CopyASPTo(CRMTogetherOS+'\\sagecrm.js','\\CustomPages\\'+CRMTogetherOS+'\\sagecrm.js');
CopyASPTo(CRMTogetherOS+'\\sagecrmnolang.js','\\CustomPages\\'+CRMTogetherOS+'\\sagecrmnolang.js');
CopyASPTo(CRMTogetherOS+'\\workflow.js','\\CustomPages\\'+CRMTogetherOS+'\\workflow.js');

var DocRouting="DocRouting";
//CreateNewDir(GetDLLDir() + '\\CustomPages\\' + DocRouting+'\\js');
CopyASPTo(DocRouting+'\\admin.asp','\\CustomPages\\'+DocRouting+'\\admin.asp');
CopyASPTo(DocRouting+'\\docRouting_template.js','\\CustomPages\\'+DocRouting+'\\docRouting_template.js');
CopyASPTo(DocRouting+'\\sagecrm.js','\\CustomPages\\'+DocRouting+'\\sagecrm.js');
CopyASPTo(DocRouting+'\\sagecrmnolang.js','\\CustomPages\\'+DocRouting+'\\sagecrmnolang.js');
CopyASPTo(DocRouting+'\\ewareNoHist.js','\\CustomPages\\'+DocRouting+'\\ewareNoHist.js');
CopyASPTo(DocRouting+'\\sendDocument.asp','\\CustomPages\\'+DocRouting+'\\sendDocument.asp');
CopyASPTo(DocRouting+'\\sendEmail.asp','\\CustomPages\\'+DocRouting+'\\sendEmail.asp');
CopyASPTo(DocRouting+'\\js\\custom\\docRouting.js','\\js\\custom\\docRouting.js');