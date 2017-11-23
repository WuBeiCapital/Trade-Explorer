#pragma once

_MITC_BASIC_BEGIN
// EnvironmentManager �Ի���
class MITC_BASIC_EXT EnvironmentData
{
public:
	EnvironmentData();   // ��׼���캯��
	 ~EnvironmentData();

	void Serialize(CArchive& ar);

	BOOL IsAutoSave() const;
	void SetAutoSave(BOOL bAutoSave);

	UINT GetAutoSaveCounter() const;
	void SetAutoSaveCounter(UINT uAutoSaveCounter);

	BOOL IsToolTip() const;
	void SetToolTip(BOOL bToolTip);
protected:

private:
	BOOL m_bAutoSave;
	UINT m_uAutoSaveCounter;	//!
	BOOL m_bToolTip;

};


class MITC_BASIC_EXT EnvironmentManager
{
public:
	EnvironmentManager();   // ��׼���캯��
	 ~EnvironmentManager();

	void Load(const CString& strPath);
	void Save(const CString& strPath) const;

	EnvironmentData* GetEnvironmentData();

protected:

private:
	CString m_strPath;
	EnvironmentData* m_pEnvironmentData;
};

MITC_BASIC_EXT EnvironmentManager* GetEnvironmentManagerInstance();


_MITC_BASIC_END