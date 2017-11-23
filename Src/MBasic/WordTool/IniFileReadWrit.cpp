//IniFileReadWrite.cpp
#include "stdafx.h"
#include "global.h"
#include "IniFileReadWrit.h"
_MITC_BASIC_BEGIN

BOOL IniFileReadWrite::WriteValue(LPCTSTR lpszSection,LPCTSTR lpszSectionKey,LPCTSTR lpszValue)
{
  return WritePrivateProfileString(lpszSection,lpszSectionKey,lpszValue,m_sPathFileName);
}

BOOL IniFileReadWrite::WriteValue(LPCTSTR lpszSection,LPCTSTR lpszSectionKey,int iValue)
{
  CString sValue;
  sValue.Format(L"%d",iValue);
  return WriteValue(lpszSection,lpszSectionKey,sValue);
}

BOOL IniFileReadWrite::WriteValue(LPCTSTR lpszSection,LPCTSTR lpszSectionKey,double dValue)
{
  CString sValue;
  sValue.Format(L"%0.2f",dValue);
  return WriteValue(lpszSection,lpszSectionKey,sValue);
}

BOOL IniFileReadWrite::ReadValue(LPCTSTR lpszSection,LPCTSTR lpszSectionKey,CString &szValue)
{
  WCHAR buff[1024];
  DWORD dwRead=GetPrivateProfileString(lpszSection,lpszSectionKey,NULL,buff,1024,m_sPathFileName);
  if(dwRead==0)
    return FALSE;
  szValue=buff;
  return TRUE;
}

BOOL IniFileReadWrite::ReadValue(CString lpszSection,LPCTSTR lpszSectionKey,int &iValue)
{
  WCHAR buff[1024];
  DWORD dwRead=GetPrivateProfileString(lpszSection,lpszSectionKey,NULL,(LPWSTR)buff,1024,m_sPathFileName);
  if(dwRead==0)
    return FALSE;
  int value=_wtoi(buff);
  iValue=value;
  return TRUE;
}

BOOL IniFileReadWrite::ReadValue(CString lpszSection,LPCTSTR lpszSectionKey,double &dValue)
{
  WCHAR buff[1024];
  DWORD dwRead=GetPrivateProfileString(lpszSection,lpszSectionKey,NULL,(LPWSTR)buff,1024,m_sPathFileName);
  if(dwRead==0)
    return FALSE;
  double value=atof((char*)buff);
  dValue=value;
  return TRUE;
}

//更改文件路径 C:a .
void IniFileReadWrite::SetPathFileName(LPCTSTR lpszPathFileName)
{
  CheckDirectory(lpszPathFileName);
  m_sPathFileName=lpszPathFileName;
}

IniFileReadWrite::IniFileReadWrite(LPCTSTR lpszPathFileName)
{
  m_sPathFileName=lpszPathFileName;
  CFileFind ff;
  if(!ff.FindFile(m_sPathFileName))
    CreateUnicodeFile(m_sPathFileName);//创建Unicode格式文件
}

IniFileReadWrite::IniFileReadWrite(void)
{
  //获取当前应用程序的路径
  CStringW strPath=GetCurrentPath();
  //文件路径
  strPath=strPath+_T("set.ini");
  CFileFind ff;
  if(!ff.FindFile(strPath))
    CreateUnicodeFile(strPath);//创建Unicode格式文件
  m_sPathFileName=strPath;
}

IniFileReadWrite::~IniFileReadWrite(void)
{

}

BOOL IniFileReadWrite::ReadAllSections( vector<CString>& vctSections )
{
	vctSections.clear();
	// TODO: Add your control notification handler code here  
	TCHAR strAppNameTemp[1024];//所有AppName的返回值  
	TCHAR strKeyNameTemp[1024];//对应每个AppName的所有KeyName的返回值  
	TCHAR strReturnTemp[1024];//返回值  
	DWORD dwKeyNameSize;//对应每个AppName的所有KeyName的总长度  
	//所有AppName的总长度  
	DWORD dwAppNameSize = GetPrivateProfileString(NULL,NULL,NULL,strAppNameTemp,1024,m_sPathFileName);  
	if(dwAppNameSize>0)  
	{  
		TCHAR *pAppName = new TCHAR[dwAppNameSize];  
		int nAppNameLen=0;  //每个AppName的长度  
		for(int i = 0;i<dwAppNameSize;i++)  
		{  
			pAppName[nAppNameLen++]=strAppNameTemp[i];  
			if(strAppNameTemp[i]=='\0')  
			{  
				CString strAppName(pAppName);
				vctSections.push_back(strAppName);
				OutputDebugString(strAppName);  
				OutputDebugString(_T("\r\n"));  
				dwKeyNameSize = GetPrivateProfileString(pAppName,NULL,NULL,strKeyNameTemp,1024,m_sPathFileName);  
				if(dwAppNameSize>0)  
				{  
					TCHAR *pKeyName = new TCHAR[dwKeyNameSize];  
					int nKeyNameLen=0;    //每个KeyName的长度  
					for(int j = 0;j<dwKeyNameSize;j++)  
					{  

						pKeyName[nKeyNameLen++]=strKeyNameTemp[j];  
						if(strKeyNameTemp[j]=='\0')  
						{  
							OutputDebugString(pKeyName);  
							OutputDebugString(_T("="));  
							if(GetPrivateProfileString(pAppName,pKeyName,NULL,strReturnTemp,1024,m_sPathFileName))  
								OutputDebugString(strReturnTemp);  
							memset(pKeyName,0,dwKeyNameSize);  
							nKeyNameLen=0;  
							OutputDebugString(_T("\r\n"));  
						}  
					}  
					
					delete[]pKeyName;  

				}  
				memset(pAppName,0,dwAppNameSize);  
				nAppNameLen=0;  
			}  
		}  
		delete[]pAppName;  
	}  
	return TRUE;
}

_MITC_BASIC_END