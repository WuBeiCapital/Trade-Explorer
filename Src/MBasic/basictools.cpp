// CSPC001T01S0102HS.cpp : 实现文件
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

bool isLeap(int y)//判断是否是闰年  
{  
    return y%4==0&&y%100!=0||y%400==0;//真返回为1，假为0  
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
    //case 0: cout <<"星期一"<<endl; break;
    //case 1: cout <<"星期二"<<endl; break;
    //case 2: cout <<"星期三"<<endl; break;
    //case 3: cout <<"星期四"<<endl; break;
    //case 4: cout <<"星期五"<<endl; break;
    //case 5: cout <<"星期六"<<endl; break;
    //case 6: cout <<"星期日"<<endl; break;
    //}
	return iWeek;
}
CString  CalcTimeString(const CString& strTime,BOOL bFront)
{
	int y,m, d;
	//int yOrg,mOrg, dOrg;
	CString strTmp,strName;
	strTmp=strTime;
	SplitTimeString(strTmp,y,m,d);
	UINT uW=6;
	if(bFront)//时间向前推
	{//!
		strTmp=_T("");
		strName=_T("");		
		do
		{//!		
			if(d>1)
			{//!
				d-=1;
			}
			else
			{//!
				if(m>1)
				{
					m-=1;
					d=daysOfMonth(y,m);
				}
				else
				{			
					y-=1;
					m=12;
					d=daysOfMonth(y,m);
				}
			}
			uW=CaculateWeekDay(y,m,d);
		}while(uW>5);

		strName=GetTimeString(y,m,d);
	}
	else
	{//！
		do
		{//!	
			d++;
			if(d<=daysOfMonth(y,m))
			{//!
				uW=CaculateWeekDay(y,m,d);
				if(uW<6)
				{
					strName=GetTimeString(y,m,d);
					strName=_T("A")+strName;
				}				
			}
			else
			{//!				
				if(m<12)
				{
					m+=1;					
				}
				else
				{
					m=1;
					y+=1;
				}
				d=1;
			
				uW=CaculateWeekDay(y,m,d);
				if(uW<6)
				{
					strName=GetTimeString(y,m,d);
					strName=_T("A")+strName;
				}
			}				
		}while(uW>5);
	}

	return strName;
}
BOOL  CalcTimeString(const CString& strTime,UINT uTimeContinue,UINT uTimePeriod,vector<CString>& vctTimeslist,BOOL bFront)
{
	int y,m, d;
	SYSTEMTIME sys; 
	GetLocalTime(&sys);	
	vector<CString> vctList;
	int yOrg=sys.wYear,mOrg=sys.wMonth, dOrg=sys.wDay;

	CString strTmp,strName;
	strTmp=strTime;
	SplitTimeString(strTmp,y,m,d);
	//!
	switch(uTimePeriod)
	{
		case 0:// 天
		{
			UINT uW=6;
			if(bFront)//时间向前推
			{//!
				d++;
				for(int i=0;i <uTimeContinue+1;++i)
				{//!
					strTmp=_T("");
					strName=_T("");		
					do
					{//!		
						if(d>1)
						{//!
							d-=1;
						}
						else
						{//!
							if(m>1)
							{
								m-=1;
								d=daysOfMonth(y,m);
							}
							else
							{			
								y-=1;
								m=12;
								d=daysOfMonth(y,m);
							}
						}
						uW=CaculateWeekDay(y,m,d);						
					}while(uW>5);

					strName=GetTimeString(y,m,d);
					strName=_T("A")+strName;
					vctTimeslist.push_back(strName);
				}
			}
			else
			{//！		
				d--;
				for(int i=0;i <uTimeContinue+1;++i)
				{//!
					strTmp=_T("");
					strName=_T("");		
					do
					{//!	
						d++;
						if(d<=daysOfMonth(y,m))
						{//!
							uW=CaculateWeekDay(y,m,d);
							if(uW<6)
							{
								strName=GetTimeString(y,m,d);
								strName=_T("A")+strName;
								vctTimeslist.push_back(strName);
							}				
						}
						else
						{//!				
							if(m<12)
							{
								m+=1;					
							}
							else
							{
								m=1;
								y+=1;
							}
							d=1;
			
							uW=CaculateWeekDay(y,m,d);
							if(uW<6)
							{
								strName=GetTimeString(y,m,d);
								strName=_T("A")+strName;
								vctTimeslist.push_back(strName);
							}
						}
						if((y>yOrg)||(y==yOrg && m>mOrg) ||(y==yOrg && m==mOrg && d>24/*dOrg*/))
						{						
							return FALSE;
						}
					}while(uW>5);
				}
			}
			break;
		}
		case 1://!周	
		{
			UINT uW=5;
			if(bFront)//时间向前推
			{//!
				for(int i=0;i <uTimeContinue+1;++i)
				{//!
					strTmp=_T("");
					strName=_T("");	
				
					uW=CaculateWeekDay(y,m,d);
					if(uW>=5)
					{
						//!调整到本周五
						d-=uW-5;
					}
					else
					{
						//!调整到上周五
						d-=uW+2;
					}
					//！
					if(d<0)
					{//!
						if(m>1)
						{
							m-=1;
							d=daysOfMonth(y,m)+d;
						}
						else
						{			
							y-=1;
							m=12;
							d=daysOfMonth(y,m)+d;
						}
					}					
					//uW=CaculateWeekDay(y,m,d);
					strName=GetTimeString(y,m,d);
					strName=_T("A")+strName;	
					vctTimeslist.push_back(strName);
					d-=7;
				}
			}
			else
			{
				for(int i=0;i <uTimeContinue+1;++i)
				{//!
					strTmp=_T("");
					strName=_T("");	
				
					uW=CaculateWeekDay(y,m,d);
					if(uW>=5)
					{
						//!调整到本周五
						d-=uW-5;
					}
					else
					{
						//!调整到本周五
						d-=uW-5;
					}
					//！
					if(d<0)
					{//!
						if(m>1)
						{
							m-=1;
							d=daysOfMonth(y,m)+d;
						}
						else
						{			
							y-=1;
							m=12;
							d=daysOfMonth(y,m)+d;
						}
					}

					if(d>daysOfMonth(y,m))
					{//!
						if(m==12)
						{
							d=d-daysOfMonth(y,m);
							m=1;
							y+=1;							
						}
						else
						{							
							m+=1;
							d=d-daysOfMonth(y,m);
						}
					}
					strName=GetTimeString(y,m,d);
					strName=_T("A")+strName;	
					vctTimeslist.push_back(strName);
					d+=7;
				}
			}
			break;	
		}
		default:
			break;
	}	
	return TRUE;
}
void SplitTimeString(const CString& strTime,int& y,int& m, int& d)
{
	CString strTmp,strName;
	strTmp=strTime;
	strName=strTmp.Mid(0,4);
	y=StrToInt(strName);
	strName=strTmp.Mid(4,2);
	m=StrToInt(strName);
	strName=strTmp.Mid(6,2);
	d=StrToInt(strName);
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

  // 下面是打开注册表, 只有打开后才能做其他操作 
  lRet = RegOpenKeyEx(HKEY_CURRENT_USER,  // 要打开的根键 
    strPath, // 要打开的子子键 
    0,        // 这个一定要为0 
    KEY_QUERY_VALUE,  //  指定打开方式,此为读 
    &hKey);    // 用来返回句柄 

  lRet = RegQueryValueEx(hKey,  // 打开注册表时返回的句柄 
    strAppName,  //要查询的名称,qq安装目录记录在这个保存 
    NULL,   // 一定为NULL或者0 
    NULL,   
    (LPBYTE)szProductType, // 我们要的东西放在这里 
    &dwBufLen);
  if(lRet != ERROR_SUCCESS)  // 判断是否查询成功 
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
    AfxMessageBox(_T("Smart BDS程序未安装！"));
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
	strPath+=_T("\\调试数据\\DumpData");

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

 
//void CheckDirectory(CString sDirectory)
//{
//  //检查目录是否存在
//  WIN32_FIND_DATA fd; 
//  HANDLE hFind = FindFirstFile(sDirectory, &fd); 
//  if (!((hFind != INVALID_HANDLE_VALUE) && (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)))
//  { 
//    ::CreateDirectory(sDirectory,NULL);
//  } 
//  FindClose(hFind); 
//}
//
//BOOL IsDirectory(const char *pDir)  
//{  
//    char szCurPath[500];  
//    ZeroMemory(szCurPath, 500);  
//    sprintf_s(szCurPath, 500, "%s//*", pDir);  
//    WIN32_FIND_DATAA FindFileData;        
//    ZeroMemory(&FindFileData, sizeof(WIN32_FIND_DATAA));  
//  
//    HANDLE hFile = FindFirstFileA(szCurPath, &FindFileData); /**< find first file by given path. */  
//  
//    if( hFile == INVALID_HANDLE_VALUE )   
//    {  
//        FindClose(hFile);  
//        return FALSE; /** 如果不能找到第一个文件，那么没有目录 */  
//    }else  
//    {     
//        FindClose(hFile);  
//        return TRUE;  
//    }  
//      
//}  
//BOOL DeleteDirectory(const CString& strDirName)  
//{  
////  CFileFind tempFind;     //声明一个CFileFind类变量，以用来搜索  
//    char szCurPath[MAX_PATH];       //用于定义搜索格式  
//    _snprintf(szCurPath, MAX_PATH, "%s//*.*", strDirName); //匹配格式为*.*,即该目录下的所有文件  
//    WIN32_FIND_DATAA FindFileData;        
//    ZeroMemory(&FindFileData, sizeof(WIN32_FIND_DATAA));  
//    HANDLE hFile = FindFirstFileA(szCurPath, &FindFileData);  
//    BOOL IsFinded = TRUE;  
//    while(IsFinded)  
//    {  
//        IsFinded = FindNextFileA(hFile, &FindFileData); //递归搜索其他的文件  
//        if( strcmp(FindFileData.cFileName, ".") && strcmp(FindFileData.cFileName, "..") ) //如果不是"." ".."目录  
//        {  
//            CString strFileName = "";  
//            strFileName = strFileName+ strDirName + _T("//") + FindFileData.cFileName;  
//            CString strTemp;  
//            strTemp = strFileName;  
//            if( IsDirectory(strFileName)) ) //如果是目录，则递归地调用  
//            {     
//                printf("目录为:%s/n", strFileName.c_str());  
//                DeleteDirectory(strTemp.c_str());  
//            }  
//            else  
//            {  
//                DeleteFileA(strTemp);  
//            }  
//        }  
//    }  
//    FindClose(hFile);  
//  
//    //BOOL bRet = RemoveDirectoryA(DirName);  
//    //if( bRet == 0 ) //删除目录  
//    //{  
//    //    printf("删除%s目录失败！/n", DirName);  
//    //    return FALSE;  
//    //}  
//    return TRUE;  
//}  

CString Time2Str(double dCost)
{//!
	CString strTime(_T(""));
	double dS(0);
	if(dCost < 60)//m
	{
		strTime.Format(_T("%d秒"),(int)dCost);
	}
	else if(dCost > 60 && dCost < 3600)
	{
		int    nM =int(floor(dCost/60.0));
		dS=dCost-nM*60;

		strTime.Format(_T("%d分%d秒"),nM,(int)dS);
	}
	else
	{
		int    nH = int(floor(dCost/3600));
		int nM = int(floor((dCost-nH*3600)/60));	

		dS=dCost-nH*3600-nM*60;

		strTime.Format(_T("%d小时%d分%d秒"),nH,nM,(int)dS);
	}
	//!
	return strTime;
};//ms

_MITC_BASIC_END 
