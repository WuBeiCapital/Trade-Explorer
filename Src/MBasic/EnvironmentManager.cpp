// ProgressManager.cpp : 实现文件
//

#include "stdafx.h"
#include "EnvironmentManager.h"
#include "BasicTools.h"

_MITC_BASIC_BEGIN
EnvironmentData::EnvironmentData()
{
	m_uAutoSaveCounter=20;	//!
	m_bToolTip=TRUE;
	m_bAutoSave=TRUE;
}// 标准构造函数
EnvironmentData::~EnvironmentData()
	 {
	 }

BOOL EnvironmentData::IsAutoSave() const
	{
		return m_bAutoSave;
	}
	void EnvironmentData::SetAutoSave(BOOL bAutoSave)
	{
		m_bAutoSave=bAutoSave;
	}

	
	BOOL EnvironmentData::IsToolTip() const
	{
		return m_bToolTip;
	}
	void EnvironmentData::SetToolTip(BOOL bToolTip)
	{
		m_bToolTip=bToolTip;
	}
	//!
	UINT EnvironmentData::GetAutoSaveCounter() const
	{
		return m_uAutoSaveCounter;
	}

	void EnvironmentData::SetAutoSaveCounter(UINT uAutoSaveCounter) 
	{
		m_uAutoSaveCounter=uAutoSaveCounter;
	}

void EnvironmentData::Serialize(CArchive& ar)
{
	/*
	if(ar.IsStoring())
	{
		_PRESAVE(_T("EnvironmentData"))
		{
		  _BEGINESAVE(BRIDGEUIFRAME_VERSION1)

		ar<< m_bAutoSave;
		ar<< m_uAutoSaveCounter;

		 _ENDSAVE

	   _BEGINESAVE(BRIDGEUIFRAME_VERSION2)
		ar<< m_bToolTip;
		 _ENDSAVE
		}
		_POSTSAVE
	}
	else
	{
		 _PREOPEN(_T("EnvironmentData"))
		{
		  case BRIDGEUIFRAME_VERSION1:
		{
			ar>> m_bAutoSave;
			ar>> m_uAutoSaveCounter;
	
		  break;
		}	
		  case BRIDGEUIFRAME_VERSION2:
		{
			ar>> m_bToolTip;
	
		  break;
		}	
		}	
		_POSTOPEN
	}
	*/
}

// EnvironmentManager 对话框
EnvironmentManager::EnvironmentManager()
{
	m_pEnvironmentData=new EnvironmentData;

	m_strPath=GetSystemPath() + L"\\config.ini";
	//!	
	Load(m_strPath);
}

EnvironmentManager::~EnvironmentManager()
{
	Save(m_strPath);

	DeleteObj(m_pEnvironmentData);
}

EnvironmentData* EnvironmentManager::GetEnvironmentData()
{
	return m_pEnvironmentData;
}

void EnvironmentManager::Load(const CString& strFullPath)
{
	/*
	CFile cfile;
	if(cfile.Open(strFullPath,CFile::modeRead))
	{
		CArchive ar(&cfile, CArchive::load);
		m_pEnvironmentData->Serialize(ar);
		ar.Close();
		cfile.Close();
	}
	else
	{
		_ASSERT("读取文件失败！");
		cfile.Close();
	}
	*/
int uAutoSaveCounter=20;	//!
int    bToolTip=TRUE;
int    bAutoSave=TRUE;

	CString lpszSection = L"CONFIG";
	CString lpszSectionKey = L"AutoSaveCounter";
	IniFileReadWrite iniRW(strFullPath);
	 
	iniRW.ReadValue(lpszSection,lpszSectionKey,uAutoSaveCounter);

	 lpszSectionKey = L"ToolTip";
	iniRW.ReadValue(lpszSection,lpszSectionKey,bToolTip);

	 lpszSectionKey = L"AutoSave";
	iniRW.ReadValue(lpszSection,lpszSectionKey,bAutoSave);

	m_pEnvironmentData->SetAutoSaveCounter(uAutoSaveCounter);
	m_pEnvironmentData->SetToolTip(bToolTip);
	m_pEnvironmentData->SetAutoSave(bAutoSave);

}

void EnvironmentManager::Save(const CString& strFullPath) const
{
	_ASSERTE_RT(m_pEnvironmentData);

	CString lpszSection = L"CONFIG";
	CString lpszSectionKey = L"AutoSaveCounter";
	IniFileReadWrite iniRW(strFullPath);
	iniRW.WriteValue(lpszSection,lpszSectionKey,(int)(m_pEnvironmentData->GetAutoSaveCounter()));

	lpszSectionKey = L"ToolTip";
	iniRW.WriteValue(lpszSection,lpszSectionKey,(int)(m_pEnvironmentData->IsToolTip()));

	lpszSectionKey = L"AutoSave";
	iniRW.WriteValue(lpszSection,lpszSectionKey,(int)(m_pEnvironmentData->IsAutoSave()));


	/*

	CFile cfile;	
	if (cfile.Open(strFullPath,CFile::modeCreate|CFile::modeWrite))
	{
		CArchive ar(&cfile, CArchive::store);
		m_pEnvironmentData->Serialize(ar);

		ar.Close();
		cfile.Close();
	}
	else
	{
		_ASSERT("存储文件失败！");
		cfile.Close();
	}
	*/
}

EnvironmentManager* GetEnvironmentManagerInstance()
{
	static EnvironmentManager manager;

	return &manager;
}

_MITC_BASIC_END