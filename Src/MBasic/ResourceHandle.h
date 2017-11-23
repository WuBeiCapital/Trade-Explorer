/*! \file: ResourceHandle.h    版权所有 (c) 2002-2008 , 北京迈达斯技术有限公司 \n
* 功能描述：Get the MoudleHandle of the Dll                    \n
*          1. 使用前，请确定exe或dll已经加载          \n
*          2. 使用全名,如"Test.exe"、"Test.dll"      \n
* 编 制 者：echo	     完 成  日 期：2007-7-13 17:21:15 \n
* 修 改 者：echo	     最后修改日期： -  - \n
* 历史记录：V 00.00.20(每次修改升级最后一个数字)  \n
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

