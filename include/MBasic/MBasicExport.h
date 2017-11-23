#pragma once


_MITC_BASIC_BEGIN
MITC_BASIC_EXT CString GetDesktopPath();

//创建所有的文件目录并打开指定文件
MITC_BASIC_EXT BOOL MOpenFile(CFile& file, const CString& strFileFullPath);
//创建所有的文件目录
MITC_BASIC_EXT BOOL MCreateFolder(const CString& strFileFullPath);


MITC_BASIC_EXT BOOL ParseFileName(const CString &strFull, CString &strPathOnly, CString &strFileName, CString &strExtension);


///************************************************************************************************************
///************************************************************************************************************
///************************************************************************************************************
_MITC_BASIC_END