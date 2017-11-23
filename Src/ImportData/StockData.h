#pragma once
class CStockData
{
public:
	CStockData(void);
	~CStockData(void);

	BOOL IsEmpty() const;
	void SetCode(DWORD dCode);
	DWORD GetCode() const;
	void SetName(const CString& strName);
	CString GetName() const;
	void SetCount(DWORD dCount);
	DWORD GetCount() const;
	void SetFactor(double dFactor);
	double GetFactor() const;

protected:
	DWORD m_uCode;
	CString m_strName;
	DWORD  m_dCount;
	double m_dFactor;
};


