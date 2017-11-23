#include "StdAfx.h"
#include "MBasicExport.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


_MITC_BASIC_BEGIN

BOOL MOpenFile(CFile& file, const CString& strFileFullPath)
{
	if(!MCreateFolder(strFileFullPath))
	{
		return FALSE;
	}	
	return file.Open(strFileFullPath, CFile::modeCreate|CFile::modeReadWrite);
};

//创建所有的文件目录
BOOL MCreateFolder(const CString& strFileFullPath)
{
	CString str = strFileFullPath;
	str.Replace(_T("//"), _T("\\"));
	str.Replace(_T("/"), _T("\\"));

	//CStringArray FolderArray;
	std::vector<CString>  vctFolders;

	int n = str.ReverseFind('\\');

	while ( n >= 0 )
	{
		str = str.Mid(0, n);
		//FolderArray.InsertAt(0, str);
		vctFolders.push_back(str);

		n = str.ReverseFind('\\');
	}

	for ( int i=vctFolders.size()-1; i>=0; --i )
	{
		CString strFolder = vctFolders.at(i);

		if (PathFileExists(strFolder) )
			continue;

		if ( _wmkdir(strFolder) != 0 )
			return FALSE;
	}

	return TRUE;
};

BOOL ParseFileName(const CString &strFull, CString &strPathOnly, CString &strFileName, CString &strExtension)
{
	strPathOnly = strFileName = strExtension = _T("");
	if (strFull.IsEmpty())
	{
		return TRUE;
	}
	CString strFile = strFull;
	int nSlashIndex = strFull.ReverseFind(_T('/'));
	if (nSlashIndex == -1)
	{
		nSlashIndex = strFull.ReverseFind(_T('\\'));
	}
	if (nSlashIndex != -1)
	{
		strPathOnly = strFull.Left(nSlashIndex + 1);
		strFile = strFull.Right(strFull.GetLength() - strPathOnly.GetLength());
	}
	strFileName = strFile;
	int nDotIndex = strFile.Find(_T('.'));
	if (nDotIndex != -1)
	{
		strFileName = strFile.Left(nDotIndex);
		strExtension = strFile.Right(strFile.GetLength() - strFileName.GetLength() - 1);
	}

	return TRUE;
};

_MITC_BASIC_END

