#include "stdafx.h"
#include "StockData.h"


CHKStockData::CHKStockData(void)
{
		m_strTime=_T("");
	 m_uCode=0;
	 m_strName=_T("");
	  m_dCount=0;
	 m_dFactor=0;
	 m_dValue=0;
}

CHKStockData::~CHKStockData(void)
{
}

BOOL CHKStockData::IsEmpty() const
{
	return m_strName.IsEmpty();
}

	void CHKStockData::SetTime(const CString& strTime)
	{
		m_strTime=strTime;
	}
	CString CHKStockData::GetTime() const
	{
		return m_strTime;
	}
	void CHKStockData::SetLastTime(const CString& strTime)
		{
		m_strLastTime=strTime;
	}
	CString CHKStockData::GetLastTime() const
	{
		return m_strLastTime;
	}

	void CHKStockData::SetCode(DWORD dCode)
	{
		m_uCode=dCode;
	}
	DWORD CHKStockData::GetCode() const
		{return m_uCode;}
	void CHKStockData::SetName(const CString& strName)
		{
			m_strName=strName;
	}
	CString CHKStockData::GetName() const
		{
			return m_strName;
	}
	void CHKStockData::SetCount(DWORD dCount)
		{
			m_dCount=dCount;
	}
	DWORD CHKStockData::GetCount() const
		{
			return m_dCount;
	}
	void CHKStockData::SetFactor(double dFactor)
		{
			m_dFactor=dFactor;
	}
	double CHKStockData::GetFactor() const
		{
			return m_dFactor;
	}
	void CHKStockData::SetValue(double dValue)
	{
		m_dValue=dValue;
	}
	double CHKStockData::GetValue() const
	{
		return m_dValue;
	}
		void CHKStockData::SetNumber(const CString& strNumber)
		{
			m_strNumber=strNumber;
		}
	CString CHKStockData::GetNumber() const
	{
		return m_strNumber;
	}