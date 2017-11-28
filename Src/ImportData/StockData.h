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
		m_dFactorSample=0;//_T("样本比例"));
		m_dMaxUp=0;//_T("最大涨幅")
		m_dMaxLow=0;//_T("最大跌幅")
		m_dFactorWF=0;//_T("盈亏比")
		m_dAvg=0;//_T("均涨幅")	
		m_strFirstTime=_T("");
	};
	~FactorData(void){};

	double m_dFactorSample;//_T("样本比例"));
	CString m_strFirstTime;
	double m_dMaxUp;//_T("最大涨幅")
	double m_dMaxLow;//_T("最大跌幅")
	double m_dFactorWF;//_T("盈亏比")
	double m_dAvg;//_T("均涨幅")	
};


class ConditionItem
{
public:
	ConditionItem(void)
	{
		m_uTimeType=0;//0 天,1 周,2 月;
		m_uTypeUporLow=0;//!趋势类型  0 Up, 1 low;
		m_uContinueCount=0;//!持续单位 
	};
	~ConditionItem(void){};

	UINT  m_uTimeType;//0 天,1 周,2 月;
	UINT  m_uTypeUporLow;//!趋势类型  0 Up, 1 low;
	UINT  m_uContinueCount;//!持续单位 
};


