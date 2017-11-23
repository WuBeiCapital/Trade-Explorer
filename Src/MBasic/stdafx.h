// stdafx.h : ��׼ϵͳ�����ļ��İ����ļ���
// ���Ǿ���ʹ�õ��������ĵ�
// �ض�����Ŀ�İ����ļ�

#pragma once

#ifndef VC_EXTRALEAN
#define VC_EXTRALEAN		// �� Windows ͷ���ų�����ʹ�õ�����
#endif

// ������뽫λ������ָ��ƽ̨֮ǰ��ƽ̨��ΪĿ�꣬���޸����ж��塣
// �йز�ͬƽ̨��Ӧֵ��������Ϣ����ο� MSDN��
#ifndef WINVER				// ����ʹ���ض��� Windows XP ����߰汾�Ĺ��ܡ�
#define WINVER 0x0501		// ����ֵ����Ϊ��Ӧ��ֵ���������� Windows �������汾��
#endif

#ifndef _WIN32_WINNT		// ����ʹ���ض��� Windows XP ����߰汾�Ĺ��ܡ�
#define _WIN32_WINNT 0x0501	// ����ֵ����Ϊ��Ӧ��ֵ���������� Windows �������汾��
#endif						

#ifndef _WIN32_WINDOWS		// ����ʹ���ض��� Windows 98 ����߰汾�Ĺ��ܡ�
#define _WIN32_WINDOWS 0x0410 // ����ֵ����Ϊ�ʵ���ֵ����ָ���� Windows Me ����߰汾��ΪĿ�ꡣ
#endif

#ifndef _WIN32_IE			// ����ʹ���ض��� IE 6.0 ����߰汾�Ĺ��ܡ�
#define _WIN32_IE 0x0600	// ����ֵ����Ϊ��Ӧ��ֵ���������� IE �������汾��
#endif

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// ĳЩ CString ���캯��������ʽ��

#include <afxwin.h>         // MFC ��������ͱ�׼���
#include <afxext.h>         // MFC ��չ

#ifndef _AFX_NO_OLE_SUPPORT
#include <afxole.h>         // MFC OLE ��
#include <afxodlgs.h>       // MFC OLE �Ի�����
#include <afxdisp.h>        // MFC �Զ�����
#endif // _AFX_NO_OLE_SUPPORT

#ifndef _AFX_NO_DB_SUPPORT
#include <afxdb.h>			// MFC ODBC ���ݿ���
#endif // _AFX_NO_DB_SUPPORT

#ifndef _AFX_NO_DAO_SUPPORT
#include <afxdao.h>			// MFC DAO ���ݿ���
#endif // _AFX_NO_DAO_SUPPORT

#ifndef _AFX_NO_OLE_SUPPORT
#include <afxdtctl.h>		// MFC �� Internet Explorer 4 �����ؼ���֧��
#endif
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>			// MFC �� Windows �����ؼ���֧��
#endif // _AFX_NO_AFXCMN_SUPPORT







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



// Unreferenced parameter.
#pragma warning(disable:4100)
// Unreferenced local instance.
#pragma warning(disable:4101)
// Local instance has initialized but not referenced.
#pragma warning(disable:4189)
// The expression in the if clouse is a constant.
#pragma warning(disable:4127)
// From double to INT may lose precision.
#pragma warning(disable:4244)
// From "" to Gdiplus::ARGB
#pragma warning(disable:4245)
// From size_t to long may lose precision.
#pragma warning(disable:4267)
//warning C4996: '_swprintf': This function or variable may be unsafe. Consider using _swprintf_s instead. To disable deprecation, use _CRT_SECURE_NO_WARNINGS. See online help for details.
#pragma warning(disable:4996)

////#include "vld\vld.h"


#include "MBasicMacro.h"
#include "MBasicEnum.h"
#include "MBasicExport.h"

#include <atlbase.h>  // Ϊ�˷������ VARIANT ���ͱ�����ʹ�� CComVariant ģ����
#include "wordtool\msword9.h"
//#include "wordtool\excel8.h"

#include "wordTool\MsWordToolDecorator.h"
#include "wordTool\createpic.h"
#include "wordTool\MsWordTool.h"
#include "wordTool\IniFileReadWrit.h"


#include "MBasicMacro.h"
#include "MBasicEnum.h"
#include "ResourceHandle.h"

#include "Clear.h"
#include "Transition.h"
#include "Serialize.h"
#include "MBasicExport.h"

#include "sqlite3.h"
#pragma comment(lib, "sqlite3.lib")
