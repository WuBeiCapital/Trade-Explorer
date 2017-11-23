#pragma once
_MITC_BASIC_BEGIN

class MITC_BASIC_EXT CSqlStringCreator
{
public:
	CSqlStringCreator(SQLOPTYPE type,const CString& tablename);
	~CSqlStringCreator(void);

public:
	void AddParameter(const CString& filedname,UINT val);
	void AddParameter(const CString& filedname,int val);
	void AddParameter(const CString& filedname,double val);
	void AddParameter(const CString& filedname,const CString&  val);

	void AddCondition(const CString& filedname,UINT val);
	void AddCondition(const CString& filedname,int val);
	void AddCondition(const CString& filedname,double val);
	void AddCondition(const CString& filedname,const CString&  val);


	CString getsqlstring();

	CString getMaxFieldSqlString(const CString& filedname);
	CString getMinFieldSqlString(const CString& filedname);

protected:
    CString m_sqlstring;
	SQLOPTYPE m_sqltype;
	CString m_tablename;
	int m_FiledIndex;
	int m_ConditionIndex;

	int FindFieldIndex();
	int FindConditionIndex();

	
};
_MITC_BASIC_END