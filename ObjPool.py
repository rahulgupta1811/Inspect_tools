from configparser import ConfigParser

ConfigLocation = 'Asset.ini'
Config = ConfigParser()
Config.read(ConfigLocation)

ProductionSystem= Config['Element']['PrdSystem']
UserName = Config['Element']['UserIco']
VendorPara = Config['Element']['VendorParameter']
Findbut = Config['Element']['FinBin']
CompCode = Config['Element']['CompanyCode']
ExportBtn = Config['Element']['Exportbtn']
NewSAPSession = Config['Element']['NewSAPSession']
AgingReportTCode = Config['Element']['AgingReportTCode']
ExcludeBAD_DM = Config['Element']['ExcludeBADCheck']
DisplaySavedList = Config['Element']['DisplaySavedList']
WCReport = Config['Element']['WCReport']
ZeroTo999 = Config['Element']['ZeroTo999']
Above999 = Config['Element']['Above999']
ExpDeb = Config['Element']['ExpDeb']
Sec_Allow = Config['Element']['SecAllow']
SpreadOption = Config['Element']['SpreadOption']
FIBLTCode = Config['Element']['FIBLTCode']
ExpandOption = Config['Element']['ExpandOption']
CopyFromClip = Config['Element']['CopyFromClip']
ExcludeBalanced = Config['Element']['ExcludeBalanced']
CredentialIssue = Config['Element']['CredentialIssue']
OpenDialog = Config['Element']['OpenDialog']
Dialog = Config['OutControl']['Dialog']
MacroFilePath = Config['OutControl']['MacroFilePath']
SAPVersion = Config['OutControl']['SAPVersion']
WCReportFolder = Config['Element']['WCReportFolder']
SelAll = Config['Element']['SelAll']
SelectExcelFormat = Config['Element']['SelectExcelFormat']
OfficeXlsx = Config['Element']['OfficeXlsx']
RWReport = Config['Element']['RWReport']
SAPPro =Config['Element']['SAPPro']
Extend = Config['Element']['Extend']


#Credentials
ConfigLocation2 = 'Credentials.ini'
Config2 = ConfigParser()
Config2.read(ConfigLocation2)
SAP_UserName = Config2['Cred']['SAPUserName']
SAP_Password = Config2['Cred']['SAPPassword']

#CompCodes
ConfigLocation3 = 'CompCode.ini'
Config3 = ConfigParser()
Config3.read(ConfigLocation3)
VP_Code1 = Config3['VendorParameter']['CompCode1']
VP_Code2 = Config3['VendorParameter']['CompCode2']
VP_Code3 = Config3['VendorParameter']['CompCode3']
WC_Code1 = Config3['WCRawReport']['CompCode1']
WC_Code2 = Config3['WCRawReport']['CompCode2']
WC_Code3 = Config3['WCRawReport']['CompCode3']
FIBL_Code1 = Config3['FIBLReport']['CompCode1']
FIBL_Code2 = Config3['FIBLReport']['CompCode2']
FIBL_Code3 = Config3['FIBLReport']['CompCode3']