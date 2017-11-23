/*! \file: ResourceHandle.h    ��Ȩ���� (c) 2002-2008 , ��������˹�������޹�˾ \n
* ����������Get the MoudleHandle of the Dll                    \n
* �� �� �ߣ�echo	     �� ��  �� �ڣ�2007-7-13 17:21:15 \n
* �� �� �ߣ�echo	     ����޸����ڣ� -  - \n
* ��ʷ��¼��V 00.00.20(ÿ���޸��������һ������)  \n
*/
#include "stdafx.h"
#include "ResourceHandle.h"
#include <shlwapi.h>


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


_MITC_BASIC_BEGIN


CResourceHandle::CResourceHandle(const CString& strSrcName)
{	
	_ASSERTE(strSrcName.Find(L'.') < 0);
	CString strName(_T(""));

	if(!strSrcName.IsEmpty())
	{
		#ifdef _DEBUG
			strName=strSrcName+_T("ud.dll");
		#else
			strName=strSrcName+_T("u.dll");
		#endif
	}
		
	m_hCurrentHandle=AfxGetResourceHandle();
	m_hMoudleHandle=GetModuleHandle(strName);
	_ASSERTE(m_hMoudleHandle);
	if(m_hMoudleHandle)
	{
		AfxSetResourceHandle(m_hMoudleHandle);
	}
	else
	{
		AfxMessageBox(strName + L" failed loading!");
	}
}

CResourceHandle::CResourceHandle(HINSTANCE hMoudleHandle)
{
	_ASSERTE(hMoudleHandle);
	m_hCurrentHandle=AfxGetResourceHandle();
	m_hMoudleHandle=hMoudleHandle;

	if(m_hCurrentHandle)
	{
		AfxSetResourceHandle(m_hMoudleHandle);
	}
}
CResourceHandle::~CResourceHandle(void)
{
	if(m_hCurrentHandle)
	{
		AfxSetResourceHandle(m_hCurrentHandle);
	}
}



_MITC_BASIC_END 
