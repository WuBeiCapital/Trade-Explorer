#include "StdAfx.h"
#include "global.h"

_MITC_BASIC_BEGIN

CString GetCurrentPath()
{
  TCHAR exeFullPath[MAX_PATH]; // MAX_PATH��API�ж����˰ɣ�������
  GetModuleFileName(NULL,exeFullPath,MAX_PATH);
  CString sPath;
  sPath=exeFullPath;
  int nPos;
  nPos=sPath.ReverseFind ('/'); 
  sPath=sPath.Left (nPos+1); 
  return sPath;
}

void CheckDirectory(CString sDirectory)
{
  //���Ŀ¼�Ƿ����
  WIN32_FIND_DATA fd; 
  HANDLE hFind = FindFirstFile(sDirectory, &fd); 
  if (!((hFind != INVALID_HANDLE_VALUE) && (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)))
  { 
    ::CreateDirectory(sDirectory,NULL);
  } 
  FindClose(hFind); 
}

CString FormartLastError(DWORD Error)
{
  LPVOID lpMsgBuf; 
  FormatMessage( 
    FORMAT_MESSAGE_ALLOCATE_BUFFER | 
    FORMAT_MESSAGE_FROM_SYSTEM | 
    FORMAT_MESSAGE_IGNORE_INSERTS, 
    NULL, 
    GetLastError(), 
    MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language 
    (TCHAR *) &lpMsgBuf, 
    0, 
    NULL 
    ); 
  TCHAR *p=(TCHAR*)lpMsgBuf;
  CString str=p;
  LocalFree( lpMsgBuf ); 
  return str;
}
BOOL CheckDirectory(LPCTSTR lpszDirectory)
{
  //���Ŀ¼�Ƿ����
  WIN32_FIND_DATA fd; 
  HANDLE hFind = FindFirstFile(lpszDirectory, &fd); 
  if (!((hFind != INVALID_HANDLE_VALUE) && (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)))
  { 
    ::CreateDirectory(lpszDirectory,NULL);
  } 
  FindClose(hFind); 
  return TRUE;
}

HANDLE CreateUnicodeFile(CStringW strPathFile)
{
  HANDLE hFile=NULL;
  //�����ļ�
  hFile=CreateFile(strPathFile,
    GENERIC_WRITE|GENERIC_READ,
    FILE_SHARE_READ|FILE_SHARE_DELETE|FILE_SHARE_WRITE,
    NULL,
    OPEN_ALWAYS,
    FILE_ATTRIBUTE_NORMAL,
    NULL);
  if(INVALID_HANDLE_VALUE==hFile)
  {
    AfxMessageBox(FormartLastError(GetLastError()));
    return NULL;
  }
  DWORD dwValue=0;
  DWORD dwSize=0;
  dwSize = GetFileSize (hFile, NULL) ; 
  if (dwSize == 0xFFFFFFFF) 
  { 
    AfxMessageBox(FormartLastError(GetLastError()));
    CloseHandle(hFile);
    return NULL;
  } 
  if(dwSize==0)
  {
    TCHAR p=0xfeff;//UNICODE�ļ���ͷ��־
    if(!WriteFile(hFile,&p,sizeof(TCHAR),&dwValue,NULL))
    {
      AfxMessageBox(FormartLastError(GetLastError()));
      CloseHandle(hFile);
      return NULL;
    }
  }
  return hFile;
}

BOOL WriteLogFile(CStringW sLogMsg)
{
  CStringW sFileName;
  if(sFileName.IsEmpty())
  {
    sFileName=GetCurrentPath();//��ȡӦ�ó�������Ŀ¼
    sFileName=sFileName+TEXT("log/");//����log�ļ���
    CheckDirectory(sFileName);
    CString sdate;
    CTime tt=CTime::GetTickCount();
    CString strtt=tt.Format("log_%Y-%m-%d.txt");
    sFileName=sFileName+strtt;
  }
  HANDLE hFile=CreateUnicodeFile(sFileName);//����UNICODE��ʽ�ļ�
  if(NULL==hFile)
  {
    AfxMessageBox(FormartLastError(GetLastError()));
    return FALSE;
  }
  DWORD dwValue=0;
  DWORD dwSize=0;
  dwSize = GetFileSize (hFile, NULL) ; 
  if (dwSize == 0xFFFFFFFF) 
  { 
    AfxMessageBox(FormartLastError(GetLastError()));
    CloseHandle(hFile);
    return FALSE;
  }
  long logcount=0;
  int iLength=0;
  DWORD ftype=GetFileType(hFile);
  if(ftype!=FILE_TYPE_DISK)//����ļ��Ƿ�Ϊ�����ļ�
    return FALSE;
  TCHAR buff[10];
  wmemset((WCHAR*)buff,L'0',10);
  if(dwSize!=2)//����Ѿ�д��־
  {
    //�ƶ����ļ���ͷsizeof(TCHAR)��
    DWORD p=SetFilePointer(hFile,sizeof(TCHAR),NULL,FILE_CURRENT);
    if(p==0xFFFFFFFF)
      return FALSE;
    //��ȡ��־��¼��   00000000
    if(!ReadFile(hFile,buff,10*sizeof(TCHAR),&dwValue,NULL))
    {
      AfxMessageBox(FormartLastError(GetLastError()));
      CloseHandle(hFile);
      return FALSE;
    }
    logcount=wcstol((WCHAR*)buff,NULL,10);
  }
  CStringW sCount;
  logcount=logcount+1;
  sCount.Format(TEXT("%d"),logcount);
  for(int i=0;i<sCount.GetLength();i++)
    buff[8-sCount.GetLength()+i]=sCount[i];
  buff[8]=' ';
  buff[9]=' ';
  SetFilePointer(hFile,sizeof(TCHAR),NULL,FILE_BEGIN);//�ƶ����ļ���ͷ2
  //д��־��¼��
  if(!WriteFile(hFile,buff,(int)10*sizeof(TCHAR),&dwValue,NULL))
  {
    AfxMessageBox(FormartLastError(GetLastError()));
    CloseHandle(hFile);
    return FALSE;
  }
  SetFilePointer(hFile,NULL,NULL,FILE_END);//�ƶ����ļ�β��
  CTime t=CTime::GetTickCount();
  CStringW sMsg=t.Format("[%Y-%m-%d %H:%M:%S] ");
  sLogMsg=sMsg+sLogMsg+TEXT(" ");
  iLength = sLogMsg.GetLength();
  //д��־
  if(!WriteFile(hFile,sLogMsg.GetBuffer(),(int)iLength*sizeof(TCHAR), &dwValue, NULL))
  {
    AfxMessageBox(FormartLastError(GetLastError()));
    CloseHandle(hFile);
    return FALSE;
  }
  //�ر��ļ�
  CloseHandle(hFile);
  return TRUE;
}

_MITC_BASIC_END