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

class FactorData
{
public:
	FactorData(void)
	{
		m_dFactorSample=0;//_T("��������"));
		m_dMaxUp=0;//_T("����Ƿ�")
		m_dMaxLow=0;//_T("������")
		m_dFactorWF=0;//_T("ӯ����")
		m_dAvg=0;//_T("���Ƿ�")	
		m_strFirstTime=_T("");
	};
	~FactorData(void){};

	double m_dFactorSample;//_T("��������"));
	CString m_strFirstTime;
	double m_dMaxUp;//_T("����Ƿ�")
	double m_dMaxLow;//_T("������")
	double m_dFactorWF;//_T("ӯ����")
	double m_dAvg;//_T("���Ƿ�")	
};


class ConditionItem
{
public:
	ConditionItem(void)
	{
		m_uTimeType=0;//0 ��,1 ��,2 ��;
		m_uTypeUporLow=0;//!��������  0 Up, 1 low;
		m_uContinueCount=0;//!������λ 
	};
	~ConditionItem(void){};

	UINT  m_uTimeType;//0 ��,1 ��,2 ��;
	UINT  m_uTypeUporLow;//!��������  0 Up, 1 low;
	UINT  m_uContinueCount;//!������λ 
};


