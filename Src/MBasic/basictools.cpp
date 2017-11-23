// CSPC001T01S0102HS.cpp : ʵ���ļ�
#include "stdafx.h"
#include "BasicTools.h"
#include <direct.h>
#include <shlwapi.h>


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

_MITC_BASIC_BEGIN

	const char* WcharToUtf8(const wchar_t *pwStr)  
{  
    if (pwStr == NULL)  
    {  
        return NULL;  
    }  
  
    int len = WideCharToMultiByte(CP_UTF8, 0, pwStr, -1, NULL, 0, NULL, NULL);  
    if (len <= 0)  
    {  
        return NULL;  
    }  
    char *pStr = new char[len];  
    WideCharToMultiByte(CP_UTF8, 0, pwStr, -1, pStr, len, NULL, NULL);  
    return pStr;  
}  
  
const wchar_t* Utf8ToWchar(const char *pStr)  
{  
    if (pStr == NULL)  
    {  
        return NULL;  
    }  
  
    int len = MultiByteToWideChar(CP_UTF8, 0, pStr, -1, NULL, 0);  
    if (len <= 0)  
    {  
        return NULL;  
    }  
    wchar_t *pwStr = new wchar_t[len];  
    MultiByteToWideChar(CP_UTF8, 0, pStr, -1, pwStr, len);  
    return pwStr;  
}  

bool isLeap(int y)//�ж��Ƿ�������  
{  
    return y%4==0&&y%100!=0||y%400==0;//�淵��Ϊ1����Ϊ0  
} 

int daysOfMonth(int y,int m)  
{  
    int day[12]={31,28,31,30,31,30,31,31,30,31,30,31};  
    if(m!=2)  
        return day[m-1];  
    else   
        return 28+isLeap(y);  
}  

UINT CaculateWeekDay(int y,int m, int d)
{
    if(m==1||m==2) {
        m+=12;
        y--;
    }
    int iWeek=(d+2*m+3*(m+1)/5+y+y/4-y/100+y/400)%7+1;
    //switch(iWeek)
    //{
    //case 0: cout <<"����һ"<<endl; break;
    //case 1: cout <<"���ڶ�"<<endl; break;
    //case 2: cout <<"������"<<endl; break;
    //case 3: cout <<"������"<<endl; break;
    //case 4: cout <<"������"<<endl; break;
    //case 5: cout <<"������"<<endl; break;
    //case 6: cout <<"������"<<endl; break;
    //}
	return iWeek;
}

CString GetTimeString(int y,int m, int d)
{
	CString strTmp,strName;
	strTmp.Format(_T("%d"),y);
	strName+=strTmp;
	if(m<10)
	{
		strTmp.Format(_T("0%d"),m);
	}
	else
	{
		strTmp.Format(_T("%d"),m);
	}
	strName+=strTmp;

	if(d<10)
	{
		strTmp.Format(_T("0%d"),d);
	}
	else
	{
		strTmp.Format(_T("%d"),d);
	}
	strName+=strTmp;

	return strName;
}

CString GetSystemPathByReg()
{
  // Dynamic allocation will be used.
  HKEY hKey;
  TCHAR szProductType[512];
  memset(szProductType,0,sizeof(szProductType));
  DWORD dwBufLen = 512;
  LONG lRet;
  CString strPath = _T("SOFTWARE\\MIDAS\\Smart BDS\\PATH\\");
  CString strAppName = _T("Installed Path");

  // �����Ǵ�ע���, ֻ�д򿪺�������������� 
  lRet = RegOpenKeyEx(HKEY_CURRENT_USER,  // Ҫ�򿪵ĸ��� 
    strPath, // Ҫ�򿪵����Ӽ� 
    0,        // ���һ��ҪΪ0 
    KEY_QUERY_VALUE,  //  ָ���򿪷�ʽ,��Ϊ�� 
    &hKey);    // �������ؾ�� 

  lRet = RegQueryValueEx(hKey,  // ��ע���ʱ���صľ�� 
    strAppName,  //Ҫ��ѯ������,qq��װĿ¼��¼��������� 
    NULL,   // һ��ΪNULL����0 
    NULL,   
    (LPBYTE)szProductType, // ����Ҫ�Ķ����������� 
    &dwBufLen);
  if(lRet != ERROR_SUCCESS)  // �ж��Ƿ��ѯ�ɹ� 
    return _T("");
  RegCloseKey(hKey);


  CString strFullBDSExePath(szProductType);
  return strFullBDSExePath;

}

CString GetDesktopPath()
{
	TCHAR szPath[MAX_PATH];
	SHGetSpecialFolderPath(NULL,szPath,CSIDL_DESKTOP,FALSE);
	CString strPath = szPath;
	strPath += "\\";
	return strPath;
};

CString GetSystemPath()
{
	TCHAR AppPathName[MAX_PATH];

  CString strModulePath = GetSystemPathByReg();
  CFileFind filefind;

 // if(!filefind.FindFile(strModulePath))
  {
    HINSTANCE hwnd=	AfxGetAppModuleState()->m_hCurrentInstanceHandle;
    GetModuleFileName(hwnd,AppPathName,MAX_PATH); 
    strModulePath = CString(AppPathName);
  }

 /* CString strExe = strModulePath.Right(12);
  
  if(strExe.CompareNoCase(_T("SmartBDS.exe")) != 0)
  {
    AfxMessageBox(_T("Smart BDS����δ��װ��"));
  }
 	  */
  int nBinPos=strModulePath.ReverseFind(_T('\\'));
	
	if(nBinPos!=-1)
	 strModulePath = strModulePath.Left(nBinPos);

	return strModulePath;
}

CString GetSysPath()
{
	CString strPath=GetSystemPath();
	strPath=strPath.MakeUpper();

	int nBinPos(0);
	nBinPos= strPath.Find(_T("BIN"));

	#ifdef _DEBUG
		nBinPos= strPath.Find(_T("DEBUG"));
	#endif	

	CString strProject=strPath.Left(nBinPos);

	return strProject;
}

CString GetProjectPath()
{
	CString strProject=GetSysPath();
	strProject+=_T("Projects");

	return strProject;
}

CString GetProjectBMPPath()
{
	CString strBMPPath=GetSysPath();
	strBMPPath+=_T("Data\\Resource\\Bmp");

	return strBMPPath;
}
CString GetDebugDataPath()
{
	CString strPath=GetSystemPath();
	strPath+=_T("\\��������\\DumpData");

	return strPath;
}

CString GetProjectDebugInfoPath()
{
	CString strPath=GetSystemPath();
	strPath+=_T("\\DebugInfo");

	return strPath;
}

CString GetProjectDrawFramePath()
{
	CString strDrawFramePath=GetSysPath();
	strDrawFramePath+=_T("Data\\DrawFrames");
	return strDrawFramePath;
}
CString GetProjectTemplateMaterialPath()//! "Data\\Template\\Material"
{
	CString strPath=GetSysPath();
	strPath+=_T("Data\\Template\\Material");
	return strPath;
}
CString GetProjectTemplateBridgePath()//! "Data\\Template\\Bridge"
{
	CString strPath=GetSysPath();
	strPath+=_T("Data\\Template\\Bridge");
	return strPath;
}
CString GetProjectTemplateComponentPath()//! "Data\\Template\\Component"
{
	CString strPath=GetSysPath();
	strPath+=_T("Data\\Template\\Component");
	return strPath;
}
CString GetProjectAccessoryEquipmentPath()//! "Data\\Template\\AccessoryEquipment"
{
	CString strPath=GetSysPath();
	strPath+=_T("Data\\Template\\AccessoryEquipment");
	return strPath;
}
CString GetProjectPrestressEquipmentPath()//! "Data\\Template\\PrestressEquipment"
{
	CString strPath=GetSysPath();
	strPath+=_T("Data\\Template\\PrestressEquipment");
	return strPath;
}
CString GetProjectTemplateBedStone()
{
	CString strPath=GetSysPath();
	strPath+=_T("Data\\Template\\BedStone");
	return strPath;
};//! "Data\\Template\\BedStone"
CString GetProjectTemplateWedgeBlk()
{
	CString strPath=GetSysPath();
	strPath+=_T("Data\\Template\\WedgeBlk");
	return strPath;
};//! "Data\\Template\\WedgeBlk"

CString GetProjectTemplateSteelWireCharacter()
{
	CString strPath=GetSysPath();
	strPath+=_T("Data\\Template\\SteelWireCharacter");
	return strPath;
}//! "Data\\Template\\SteelWireCharacter"
CString GetProjectliveload()
{
CString strPath=GetSysPath();
	strPath+=_T("Data\\Template\\liveload");
	return strPath;
}//! "Data\\Template\\liveload"

CString GetProjectCalculateBookPath()
{
	CString strPath=GetSysPath();
	strPath+=_T("Data\\Template\\CalculateBook");
	return strPath;
}//! "Data\\Template\\CalculateBook"

CString GetProjectDesignReportPath()
{
		CString strPath=GetSysPath();
	strPath+=_T("Data\\Template\\DesignReport");
	return strPath;
}//! "Data\\Template\\DesignReport"
CString GetProjectAuditPath()
{
	CString strPath=GetSysPath();
	strPath+=_T("Data\\Template\\Audit");
	return strPath;
}//! "Data\\Template\\Audit"

//! "Data\\Options\\DrawingSetting"
CString GetProjectDrawingSettingPath()
{
	CString strPath=GetSysPath();
	strPath+=_T("Data\\Options\\DrawingSetting");
	return strPath;
}

CString GetProjectViewSettingPath()
{
	CString strPath=GetSysPath();
	strPath+=_T("Data\\Options\\ViewSetting");

	return strPath;
}//! "Data\\Options\\ViewSetting"
CString GetProjectEnvironmentSettingPath()
{
	CString strPath=GetSysPath();
	strPath+=_T("Data\\Options\\EnvironmentSetting");

	return strPath;
}//! "Data\\Options\\EnvironmentSetting"

CString GetProjectDrawingFontPath()
{
	CString strPath=GetSysPath();
	strPath+=_T("Data\\Fonts");
	return strPath;
}//! 
void SendMessage2Main(const CString& strMassage,StateFlag enStateFlag,DoInfoGrade enDoInfoGrade)
{
	CProjectDoInfo* projectDoInfo=new CProjectDoInfo;
	projectDoInfo->m_enStateFlag=enStateFlag;
	projectDoInfo->m_enDoInfoGrade=enDoInfoGrade;
	projectDoInfo->m_vctStrings.push_back(strMassage);
  if(AfxGetMainWnd())
	  AfxGetMainWnd()->/*SendMessage*/PostMessage(WM_PROJECTINFO_MESSAGE,0,  (LPARAM) (projectDoInfo));
}
void SendMessage2Main(const std::vector<CString>& vctMessages,StateFlag enStateFlag,DoInfoGrade enDoInfoGrade)
{
	CProjectDoInfo* projectDoInfo=new CProjectDoInfo;
	projectDoInfo->m_enStateFlag=enStateFlag;
	projectDoInfo->m_enDoInfoGrade=enDoInfoGrade;
	projectDoInfo->m_vctStrings=vctMessages;
  if(AfxGetMainWnd())
	  AfxGetMainWnd()->PostMessage(WM_PROJECTINFO_MESSAGE,0,  (LPARAM) (projectDoInfo));
}

 CString GetProjectHelpDataPath()
{
  CString strPath=GetSysPath();
  strPath+=_T("Data\\HelpData");

  return strPath;
}


_MITC_BASIC_END 
