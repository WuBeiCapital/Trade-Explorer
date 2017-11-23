
// stdafx.h : 可在此標頭檔中包含標準的系統 Include 檔，
// 或是經常使用卻很少變更的
// 專案專用 Include 檔案

#pragma once

#ifndef VC_EXTRALEAN
#define VC_EXTRALEAN            // 從 Windows 標頭排除不常使用的成員
#endif

#include "targetver.h"

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS      // 明確定義部分的 CString 建構函式

// 關閉 MFC 隱藏一些常見或可忽略警告訊息的功能
#define _AFX_ALL_WARNINGS

#include <afxwin.h>         // MFC 核心與標準元件
#include <afxext.h>         // MFC 擴充功能


#include <afxdisp.h>        // MFC Automation 類別



#ifndef _AFX_NO_OLE_SUPPORT
#include <afxdtctl.h>           // MFC 支援的 Internet Explorer 4 通用控制項
#endif
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>             // MFC 支援的 Windows 通用控制項
#endif // _AFX_NO_AFXCMN_SUPPORT

#include <afxcontrolbars.h>     // 功能區和控制列的 MFC 支援


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

const CString strTitckCode=_T("股票代码");
const CString strTitckName=_T("股票名称");
const CString strTitckNumber=_T("持股数量");
const CString strTitckFactor=_T("持股比例");

const int COL_CODE=0;
const int COL_NAME=1;
const int COL_COUNT=2;
const int COL_FACTOR=3;

//LPCTSTR lpcsSystem = _T("系统数据");


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