#pragma once


_MITC_BASIC_BEGIN
MITC_BASIC_EXT CString GetDesktopPath();

//�������е��ļ�Ŀ¼����ָ���ļ�
MITC_BASIC_EXT BOOL MOpenFile(CFile& file, const CString& strFileFullPath);
//�������е��ļ�Ŀ¼
MITC_BASIC_EXT BOOL MCreateFolder(const CString& strFileFullPath);


MITC_BASIC_EXT BOOL ParseFileName(const CString &strFull, CString &strPathOnly, CString &strFileName, CString &strExtension);


///************************************************************************************************************
///************************************************************************************************************
///************************************************************************************************************
_MITC_BASIC_END