#pragma once
class CStockData
{
public:
	CStockData(void);
	~CStockData(void);

	BOOL IsEmpty() const;
	void SetTime(const CString& strTime);
	CString GetTime() const;
	void SetCode(DWORD dCode);
	DWORD GetCode() const;
	void SetName(const CString& strName);
	CString GetName() const;
	void SetCount(DWORD dCount);
	DWORD GetCount() const;
	void SetFactor(double dFactor);
	double GetFactor() const;
	void SetValue(double dValue);
	double GetValue() const;
protected:
	CString m_strTime;
	DWORD m_uCode;
	CString m_strName;
	DWORD  m_dCount;
	double m_dFactor;
	double m_dValue;
};


