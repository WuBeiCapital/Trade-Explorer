
// stdafx.h : ���ڴ˘��^�n�а����˜ʵ�ϵ�y Include �n��
// ���ǽ���ʹ�Ås����׃����
// �������� Include �n��

#pragma once

#ifndef VC_EXTRALEAN
#define VC_EXTRALEAN            // �� Windows ���^�ų�����ʹ�õĳɆT
#endif

#include "targetver.h"

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS      // ���_���x���ֵ� CString ������ʽ

// �P�] MFC �[��һЩ��Ҋ��ɺ��Ծ���ӍϢ�Ĺ���
#define _AFX_ALL_WARNINGS

#include <afxwin.h>         // MFC �����c�˜�Ԫ��
#include <afxext.h>         // MFC �U�书��


#include <afxdisp.h>        // MFC Automation e



#ifndef _AFX_NO_OLE_SUPPORT
#include <afxdtctl.h>           // MFC ֧Ԯ�� Internet Explorer 4 ͨ�ÿ����
#endif
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>             // MFC ֧Ԯ�� Windows ͨ�ÿ����
#endif // _AFX_NO_AFXCMN_SUPPORT

#include <afxcontrolbars.h>     // ���܅^�Ϳ����е� MFC ֧Ԯ


#include <afx.h>
#include <algorithm>
#include <afxtempl.h>
#include <math.h>
#include <vector>
#include <list>
#include <map>
#include <crtdbg.h>
#include <afxpriv.h>

#include <cstdio>
#include <cstring>
using namespace std;

//#include <gdiplus.h>
//using namespace Gdiplus;
//#pragma comment( lib, "gdiplus.lib" )

#define _DELPTR(X) \
if((X)){delete (X);(X)=NULL;}

#define _MITC_TMP_BEGIN   namespace MBasic {
#define _MITC_TMP_END     }

//#include "../XLReport/XLControl.h"
//#include "../PlanBase/AhArray.h"

#include "../MBasic/MBasicInc.h"
#pragma comment(lib, "sqlite3.lib")
#pragma comment(lib, "libgm.lib")

using namespace MBasic;
#include <sstream>

#include <ctime>
#include <iostream>
#include <map>
#include <vector>

//!
#include <stdio.h>
#include <stdlib.h>
#include <string>
#include <iostream>
#include "gm\mdapi.h"

//!
#include "atlbase.h"
#include "atlstr.h"
#include <windows.h> 
#include <stdio.h> 

#include "XLControl.h"
#include "StockData.h"

const CString strTitckCode=_T("��Ʊ����");
const CString strTitckName=_T("��Ʊ����");
const CString strTitckNumber=_T("�ֹ�����");
const CString strTitckFactor=_T("�ֹɱ���");

const int COL_CODE=0;
const int COL_NAME=1;
const int COL_COUNT=2;
const int COL_FACTOR=3;

//LPCTSTR lpcsSystem = _T("ϵͳ����");


#ifdef _UNICODE
#if defined _M_IX86
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_X64
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif
#endif


#if !defined(_X64)
	#if defined(_DEBUG)
		#if defined(_AFXEXT)
			#define AUTOLIBNAME "../../Lib/Debug/Win32/sqlite3.lib"
		#else
			#define AUTOLIBNAME "../../Lib/Debug/Win32/sqlite3.lib"
		#endif
	#else
		#if defined(_AFXEXT)
			#define AUTOLIBNAME "../../Lib/Release/Win32/sqlite3.lib"
		#else
			#define AUTOLIBNAME "../../Lib/Release/Win32/sqlite3.lib"
		#endif
	#endif
#else
	#if defined(_DEBUG)
		#if defined(_AFXEXT)
			#define AUTOLIBNAME "../../Lib/Debug/x64/sqlite3.lib"
		#else
			#define AUTOLIBNAME "../../Lib/Debug/x64/sqlite3.lib"
		#endif
	#else
		#if defined(_AFXEXT)
			#define AUTOLIBNAME "../../Lib/Release/x64/sqlite3.lib"
		#else
			#define AUTOLIBNAME "../../Lib/Release/x64/sqlite3.lib"
		#endif
	#endif
#endif

// Perform autolink here:
#pragma message( "automatically link with (" AUTOLIBNAME ")")
#pragma comment(lib, AUTOLIBNAME)
//#undef AUTOLIBNAME