#pragma once


_MITC_BASIC_BEGIN

#define WM_PROJECTINFO_MESSAGE WM_USER + 16
#define WM_SHOW_STATEPROGRESS (WM_USER + 116)
#define WM_SET_END_STATEPROGRESS (WM_USER + 117)

MITC_BASIC_EXT CString GetSystemPath();//! "Bin"or "Debug"
MITC_BASIC_EXT CString GetProjectPath();//! "Projects"
MITC_BASIC_EXT CString GetDebugDataPath();//! "Data\\调试数据\\DumpData"
MITC_BASIC_EXT CString GetProjectBMPPath();//! "Data\\Resource\\Bmp"
MITC_BASIC_EXT CString GetProjectHelpDataPath();//! "Data\\Resource\\Bmp"

MITC_BASIC_EXT CString GetProjectTemplateMaterialPath();//! "Data\\Template\\Material"
MITC_BASIC_EXT CString GetProjectTemplateBridgePath();//! "Data\\Template\\Bridge"
MITC_BASIC_EXT CString GetProjectTemplateComponentPath();//! "Data\\Template\\Component"
MITC_BASIC_EXT CString GetProjectAccessoryEquipmentPath();//! "Data\\Template\\AccessoryEquipment"
MITC_BASIC_EXT CString GetProjectPrestressEquipmentPath();//! "Data\\Template\\PrestressEquipment"
MITC_BASIC_EXT CString GetProjectTemplateSteelWireCharacter();//! "Data\\Template\\SteelWireCharacter"
MITC_BASIC_EXT CString GetProjectTemplateBedStone();//! "Data\\Template\\BedStone"
MITC_BASIC_EXT CString GetProjectTemplateWedgeBlk();//! "Data\\Template\\WedgeBlk"

MITC_BASIC_EXT CString GetProjectCalculateBookPath();//! "Data\\Template\\CalculateBook"
MITC_BASIC_EXT CString GetProjectDesignReportPath();//! "Data\\Template\\DesignReport"
MITC_BASIC_EXT CString GetProjectAuditPath();//! "Data\\Template\\Audit"
MITC_BASIC_EXT CString GetProjectliveload();//! "Data\\Template\\liveload"
MITC_BASIC_EXT CString GetProjectDrawFramePath();//! "Data\\DrawFrames"
MITC_BASIC_EXT CString GetProjectDrawingSettingPath();//! "Data\\Options\\DrawingSetting"
MITC_BASIC_EXT CString GetProjectViewSettingPath();//! "Data\\Options\\ViewSetting"
MITC_BASIC_EXT CString GetProjectEnvironmentSettingPath();//! "Data\\Options\\EnvironmentSetting"

MITC_BASIC_EXT CString GetProjectDrawingFontPath();//! "Data\\Fonts"

MITC_BASIC_EXT CString GetProjectDebugInfoPath();//! "Bin\\DebugInfo"

MITC_BASIC_EXT void SendMessage2Main(const CString& strMassage,StateFlag enStateFlag=SF_SUCCESS,DoInfoGrade enDoInfoGrade=PDG_COMPONENT);
MITC_BASIC_EXT void SendMessage2Main(const std::vector<CString>& vctMessages,StateFlag enStateFlag=SF_SUCCESS,DoInfoGrade enDoInfoGrade=PDG_COMPONENT);

MITC_BASIC_EXT const char* WcharToUtf8(const wchar_t *pwStr);   
MITC_BASIC_EXT const wchar_t* Utf8ToWchar(const char *pStr);

MITC_BASIC_EXT bool isLeap(int y);//判断是否是闰年
MITC_BASIC_EXT int daysOfMonth(int y,int m);
MITC_BASIC_EXT UINT CaculateWeekDay(int y,int m, int d);
MITC_BASIC_EXT CString GetTimeString(int y,int m, int d);


_MITC_BASIC_END