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
MITC_BASIC_EXT CString GetTimeString(int y,int m, int d, const CString& strSplit=_T(""));
MITC_BASIC_EXT void SplitTimeString(const CString& strTime,int& y,int& m, int& d);
MITC_BASIC_EXT CString  CalcTimeString(const CString& strTime,BOOL bFront=TRUE);
MITC_BASIC_EXT BOOL  CalcTimeString(const CString& strTime,UINT uTimeContinue,UINT uTimePeriod,vector<CString>& vctTimeslist,BOOL bFront=TRUE);//0 天, 1 周 ,2 月

const double NS_MIN_VALUE=1.0e-6;
// a==b
inline bool Equal2Dbl(double val1,double val2=0,double dEps=NS_MIN_VALUE) { return fabs(val2-val1) < dEps; }
// a==b with precision
inline bool Equal2DblEx(double val1,double val2,double dEps=NS_MIN_VALUE) {	return fabs(val2-val1) < dEps; }
// a!=b
inline bool DblNotEqual(double a,double b,double dEps=NS_MIN_VALUE) { return !Equal2DblEx(a, b, dEps); }
// a<b
inline bool DblLT(double a,double b,double dEps=NS_MIN_VALUE) { return DblNotEqual(a, b,dEps) && (a<b);}
// a<=b
inline bool DblLE(double a,double b,double dEps=NS_MIN_VALUE) { return Equal2DblEx(a, b,dEps)||DblLT(a, b);}
// a>b
inline bool DblGT(double a,double b,double dEps=NS_MIN_VALUE) { return DblNotEqual(a, b,dEps) && (a>b);}
// a>=b
inline bool DblGE(double a,double b,double dEps=NS_MIN_VALUE) { return Equal2DblEx(a, b,dEps) || DblGT(a, b);}

MITC_BASIC_EXT CString GetDesktopPath();
//
//MITC_BASIC_EXT void CheckDirectory(CString sDirectory);
//
//MITC_BASIC_EXT  BOOL DeleteDirectory(const CString& strDirName) ;
//
//MITC_BASIC_EXT CString Time2Str(double dCost);


_MITC_BASIC_END