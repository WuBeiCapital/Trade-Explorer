/*! \file: ResourceHandle.h    ��Ȩ���� (c) 2002-2008 , ��������˹�������޹�˾ \n
* ����������Get the MoudleHandle of the Dll                    \n
*          1. ʹ��ǰ����ȷ��exe��dll�Ѿ�����          \n
*          2. ʹ��ȫ��,��"Test.exe"��"Test.dll"      \n
* �� �� �ߣ�echo	     �� ��  �� �ڣ�2007-7-13 17:21:15 \n
* �� �� �ߣ�echo	     ����޸����ڣ� -  - \n
* ��ʷ��¼��V 00.00.20(ÿ���޸��������һ������)  \n
*/

#pragma once
#include "MBasicMacro.h"

_MITC_BASIC_BEGIN

class MITC_BASIC_EXT CResourceHandle
{
public:
	CResourceHandle(const CString& strSrcName);
	CResourceHandle(HINSTANCE hMoudleHandle);
public:
	~CResourceHandle(void);
private:
	HINSTANCE m_hMoudleHandle;
	HINSTANCE m_hCurrentHandle;
};

_MITC_BASIC_END 

