#pragma once

//#ifdef _DEBUG
//	#pragma comment(lib,"MBasic.lib") 
//	#pragma message("Automatically linking with MBasic.dll")
//#else
//	#pragma comment(lib,"MBasic.lib") 
//	#pragma message("Automatically linking with MBasic.dll") 
//#endif

#if !defined(_X64)
	#if defined(_DEBUG)
		#if defined(_AFXEXT)
			#define AUTOLIBNAME "../../Lib/Debug/Win32/MBasic.lib"
		#else
			#define AUTOLIBNAME "../../Lib/Debug/Win32/MBasic.lib"
		#endif
	#else
		#if defined(_AFXEXT)
			#define AUTOLIBNAME "../../Lib/Release/Win32/MBasic.lib"
		#else
			#define AUTOLIBNAME "../../Lib/Release/Win32/MBasic.lib"
		#endif
	#endif
#else
	#if defined(_DEBUG)
		#if defined(_AFXEXT)
			#define AUTOLIBNAME "../../Lib/Debug/x64/MBasic.lib"
		#else
			#define AUTOLIBNAME "../../Lib/Debug/x64/MBasic.lib"
		#endif
	#else
		#if defined(_AFXEXT)
			#define AUTOLIBNAME "../../Lib/Release/x64/MBasic.lib"
		#else
			#define AUTOLIBNAME "../../Lib/Release/x64/MBasic.lib"
		#endif
	#endif
#endif

// Perform autolink here:
#pragma message( "automatically link with (" AUTOLIBNAME ")")
#pragma comment(lib, AUTOLIBNAME)


#include <complex>		//模板类complex的标准头文件
#include <valarray>	//模板类valarray的标准头文件
#include <math.h>		//数学头文件
#include <iostream>			//模板类输入输出流标准头文件
#include <map>
#include <list>
#include <set>
#include <vector>
#include <map>
#include <CString>
#include <afxTempl.h>
#include <float.h>
#include <algorithm>

#define _CRTDBG_MAP_ALLOC
#include <stdlib.h>
#include <crtdbg.h>

using namespace std;

#include "MBasicMacro.h"
#include "MBasicEnum.h"
#include "ResourceHandle.h"

#include "Clear.h"
#include "Transition.h"
#include "Serialize.h"
#include "MBasicExport.h"

#include "BasicTools.h"
#include "wordTool\MsWordToolDecorator.h"
#include "wordTool\CreatePic.h"
#include "wordTool\MsWordTool.h"
#include "wordTool\IniFileReadWrit.h"

#include "SQL\sqlite3.h"
#include "SQL\CppSQLite3.h"
//#include "CppSQLite3U.h"

#include "EnvironmentManager.h"

#include "SqlStringCreator.h"












