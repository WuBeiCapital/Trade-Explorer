#pragma once

#ifdef MBASIC
#define MITC_BASIC_EXT __declspec(dllexport)
#else
#define MITC_BASIC_EXT __declspec(dllimport)
#endif

#define _MITC_BASIC_BEGIN   namespace MBasic {
#define _MITC_BASIC_END     }

#define _ASSERTE_RT(expr) \
        _ASSERTE(expr);  \
      if(!(expr))  return; \

#define _ASSERTE_RT_BL(expr) \
        _ASSERTE(expr);  \
      if(!(expr))  return FALSE; \

#define _ASSERTE_RT_UI(expr) \
	_ASSERTE(expr);  \
	if(!(expr))  return 0; \

#define _ASSERTE_RT_DBL(expr) \
	_ASSERTE(expr);  \
	if(!(expr))  return 0.; \

#define GETSTR(X)   #X

const long   lEndVersion =       999;    //�����汾��

//Ԥ������
#define _PRESAVE(beginMark) \
  CFile* pFile = ar.GetFile(); \
{ \
  CString strbeginMark(beginMark); \
  ar << strbeginMark; \
} \

//��ʼ����һ���汾������
#define _BEGINESAVE(version) \
{\
  ar << version; \
  ar.Flush(); \
  ULONGLONG posBegine = pFile->GetPosition();\
  LONGLONG dLen = 3; \
  ar << dLen; \
  { \
//��������һ���汾������
#define _ENDSAVE\
  } \
  ar.Flush(); \
  ULONGLONG posEnd = pFile->GetPosition();\
  LONG posOffset = LONG(posBegine - posEnd);\
  /*ULONGLONG lngFile = */pFile->Seek(posOffset, CFile::current); \
  posOffset = - posOffset; \
  pFile->Write(&posOffset, sizeof(LONGLONG)); \
  pFile->SeekToEnd(); \
}
//�������
#define _POSTSAVE \
  ar << lEndVersion; \
  ar.Flush();\
  pFile->SeekToEnd(); \

//Ԥ�����
#define _PREOPEN(beginMarkCheck) \
  int iVersion = -1; \
  LONGLONG dLen = 0;     \
  TCHAR *char1 = NULL; \
  CString strMark(_T("")); \
  ar >> strMark; \
  _ASSERT(strMark==beginMarkCheck); \
  OutputDebugString(strMark);\
  while (ar >>iVersion,iVersion < lEndVersion) \
  { \
    ar >> dLen; \
    switch(iVersion) \
    { \
//�򿪺���
#define _POSTOPEN \
		default: \
		{ \
			char1 = new TCHAR[(size_t)dLen]; \
			LONGLONG lngReadLen = dLen - sizeof(LONGLONG); \
			if (lngReadLen > 0) \
			{ \
			  ar.Read(&char1,(UINT)lngReadLen); \
			  delete []char1;  char1 = NULL;\
			}\
		} \
		break; \
    } \
  } \
  if ( iVersion == lEndVersion) \
  { \
    int i = 0; \
    i = iVersion; \
    \
  } \

