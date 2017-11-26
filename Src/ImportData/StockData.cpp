#include "stdafx.h"
#include "StockData.h"


CStockData::CStockData(void)
{
		m_strTime=_T("");
	 m_uCode=0;
	 m_strName=_T("");
	  m_dCount=0;
	 m_dFactor=0;
	 m_dValue=0;
}

CStockData::~CStockData(void)
{
}

BOOL CStockData::IsEmpty() const
{
	return m_strName.IsEmpty();
}

	void CStockData::SetTime(const CString& strTime)
	{
		m_strTime=strTime;
	}
	CString CStockData::GetTime() const
	{
		return m_strTime;
	}
	void CStockData::SetCode(DWORD dCode)
	{
		m_uCode=dCode;
	}
	DWORD CStockData::GetCode() const
		{return m_uCode;}
	void CStockData::SetName(const CString& strName)
		{
			m_strName=strName;
	}
	CString CStockData::GetName() const
		{
			return m_strName;
	}
	void CStockData::SetCount(DWORD dCount)
		{
			m_dCount=dCount;
	}
	DWORD CStockData::GetCount() const
		{
			return m_dCount;
	}
	void CStockData::SetFactor(double dFactor)
		{
			m_dFactor=dFactor;
	}
	double CStockData::GetFactor() const
		{
			return m_dFactor;
	}
	void CStockData::SetValue(double dValue)
	{
		m_dValue=dValue;
	}
	double CStockData::GetValue() const
	{
		return m_dValue;
	}