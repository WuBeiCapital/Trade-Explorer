#pragma once
class CHKStockData
{
public:
	CHKStockData(void);
	~CHKStockData(void);

	BOOL IsEmpty() const;
	void SetTime(const CString& strTime);
	CString GetTime() const;
	void SetCode(DWORD dCode);
	DWORD GetCode() const;
	void SetName(const CString& strName);
	CString GetName() const;
	void SetNumber(const CString& strNumber);
	CString GetNumber() const;

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
	CString m_strNumber;
};


