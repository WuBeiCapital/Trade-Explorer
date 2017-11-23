#include "stdafx.h"
#include "StockData.h"


CStockData::CStockData(void)
{
}


CStockData::~CStockData(void)
{
}

BOOL CStockData::IsEmpty() const
{
	return m_strName.IsEmpty();
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
