#include "StdAfx.h"
#include "SqlStringCreator.h"
_MITC_BASIC_BEGIN

const CString CS_FILED_CHAR  = _T("©");
const CString CS_COND_CHAR  = _T("®");

CSqlStringCreator::CSqlStringCreator(SQLOPTYPE type,const CString& tablename)
{
	m_sqltype = type;
	m_tablename = tablename;

	if(m_sqltype == SQLOT_INSERT)
	{
		m_sqlstring.Format(L"INSERT INTO %s (%s) VALUES (%s)",tablename,CS_FILED_CHAR,CS_COND_CHAR)  ;
	}
	else if(m_sqltype == SQLOT_UPDATE)
	{
		m_sqlstring.Format(L"UPDATE %s SET %s WHERE %s",tablename,CS_FILED_CHAR,CS_COND_CHAR)  ;
	}
	else if(m_sqltype == SQLOT_DELETE)
	{
		m_sqlstring.Format(L"DELETE FROM %s WHERE %s",tablename,CS_COND_CHAR)  ;
	}
	m_FiledIndex = 0 ;
	m_ConditionIndex = 0 ;
}

CSqlStringCreator::~CSqlStringCreator(void)
{
}

void CSqlStringCreator::AddParameter( const CString& filedname,int val )
{
	int filedindex = FindFieldIndex();
	if(filedindex == -1)
		return ;
	
	CString par = L"";

	if(m_sqltype == SQLOT_INSERT)
	{
		m_sqlstring.Insert(filedindex ,filedname + L",");
		int valueindex = FindConditionIndex();
		if(valueindex == -1)
			return ;
		par.Format(L"%d,",val);
		m_sqlstring.Insert(valueindex ,par);
	}
	else if(m_sqltype == SQLOT_UPDATE)
	{
		par.Format(L"%s = %d,",filedname,val);
		m_sqlstring.Insert(filedindex,par);

	}
   
}

void CSqlStringCreator::AddParameter( const CString& filedname,double val )
{
	int filedindex = FindFieldIndex();
	if(filedindex == -1)
		return ;
	CString par = L"";

	if(m_sqltype == SQLOT_INSERT)
	{
		m_sqlstring.Insert(filedindex ,filedname + L",");
		int valueindex = FindConditionIndex();
		if(valueindex == -1)
			return ;
		par.Format(L"%g,",val);
		m_sqlstring.Insert(valueindex ,par);
	}
	else if(m_sqltype == SQLOT_UPDATE)
	{
		par.Format(L"%s = %g,",filedname,val);
		m_sqlstring.Insert(filedindex ,par);

	}

}

void CSqlStringCreator::AddParameter( const CString& filedname,const CString& val )
{
	int filedindex = FindFieldIndex();
	if(filedindex == -1)
		return ;
	CString par = L"";

	if(m_sqltype == SQLOT_INSERT)
	{
		m_sqlstring.Insert(filedindex ,filedname + L",");
		int valueindex = FindConditionIndex();
		if(valueindex == -1)
			return ;

		par.Format(L"'%s',",val);
		m_sqlstring.Insert(valueindex ,par);
	}
	else if(m_sqltype == SQLOT_UPDATE)
	{
		par.Format(L"%s = '%s',",filedname,val);
		m_sqlstring.Insert(filedindex ,par);

	}

}

void CSqlStringCreator::AddParameter( const CString& filedname,UINT val )
{
	int filedindex = FindFieldIndex();
	if(filedindex == -1)
		return ;
	CString par = L"";

	if(m_sqltype == SQLOT_INSERT)
	{
		m_sqlstring.Insert(filedindex ,filedname + L",");
		int valueindex = FindConditionIndex();
		if(valueindex == -1)
			return ;

		par.Format(L"%u,",val);
		m_sqlstring.Insert(valueindex ,par);
	}
	else if(m_sqltype == SQLOT_UPDATE)
	{
		par.Format(L"%s = %u,",filedname,val);
		m_sqlstring.Insert(filedindex ,par);

	}
}

CString CSqlStringCreator::getsqlstring()
{
	m_sqlstring.Replace(L"," + CS_FILED_CHAR,L"");
	if(m_sqltype == SQLOT_INSERT)
	{
		m_sqlstring.Replace(L"," + CS_COND_CHAR,L"");
	}
	else
	 m_sqlstring.Replace(L"AND " + CS_COND_CHAR,L"");

	if(m_sqlstring.Find(CS_FILED_CHAR) != -1)
	{
		return L"";
	}

	if(m_sqlstring.Find(CS_COND_CHAR) != -1)
		m_sqlstring.Replace(CS_COND_CHAR,L" 1 = 1");

  return m_sqlstring;
}

void CSqlStringCreator::AddCondition( const CString& filedname,int val )
{
	int valueindex = FindConditionIndex();
	if(valueindex == -1)
		return ;

	if(m_sqltype == SQLOT_UPDATE || m_sqltype == SQLOT_DELETE)
	{
		CString par = L"";
		par.Format(L"%s = %d AND ",filedname,val);
		m_sqlstring.Insert(valueindex ,par);
	}

}

void CSqlStringCreator::AddCondition( const CString& filedname,double val )
{
	int valueindex = FindConditionIndex();
	if(valueindex == -1)
		return ;

	if(m_sqltype == SQLOT_UPDATE || m_sqltype == SQLOT_DELETE)
	{
		CString par = L"";
		par.Format(L"%s = %g AND ",filedname,val);
		m_sqlstring.Insert(valueindex ,par);
	}

}

void CSqlStringCreator::AddCondition( const CString& filedname,const CString& val )
{
	int valueindex = FindConditionIndex();
	if(valueindex == -1)
		return ;

	if(m_sqltype == SQLOT_UPDATE || m_sqltype == SQLOT_DELETE)
	{
		CString par = L"";
		par.Format(L"%s = '%s' AND ",filedname,val);
		m_sqlstring.Insert(valueindex ,par);
	}

}

void CSqlStringCreator::AddCondition( const CString& filedname,UINT val )
{
	int valueindex = FindConditionIndex();
	if(valueindex == -1)
		return ;

	if(m_sqltype == SQLOT_UPDATE || m_sqltype == SQLOT_DELETE)
	{
		CString par = L"";
		par.Format(L"%s = %u AND ",filedname,val);
		m_sqlstring.Insert(valueindex ,par);
	}
}

int CSqlStringCreator::FindFieldIndex()
{
	return m_sqlstring.Find(CS_FILED_CHAR);
}

int CSqlStringCreator::FindConditionIndex()
{
    return m_sqlstring.Find(CS_COND_CHAR);
}

CString CSqlStringCreator::getMaxFieldSqlString( const CString& filedname )
{
   return L"SELECT MAX(" + filedname + L") FROM " + m_tablename;
}

CString CSqlStringCreator::getMinFieldSqlString( const CString& filedname )
{
   return L"SELECT MIN(" + filedname + L") FROM " + m_tablename;
}

_MITC_BASIC_END