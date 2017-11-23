#include "StdAfx.h"
#include "Transition.h"


_MITC_BASIC_BEGIN

CString PairDbl2Str(const std::vector<pair<double,double> >& vctPair)
{
	CString strRlt(_T("")),strT1(_T("")),strT2(_T(""));
	for(std::vector<pair<double,double> >::const_iterator p=vctPair.begin();p!=vctPair.end();++p)
	{
		strT1.Format(_T("%.2f"),(*p).first);//! 
		strT2.Format(_T("%.2f"),(*p).second);//! 
		if(!strRlt.IsEmpty())
			strRlt+=';';
			
		strRlt+=strT1+','+strT2;
	}

	return strRlt;
}

void Str2PairDbl(const CString& strSrc,std::vector<pair<double,double> >& vctPair)
{
	CString strRlt(_T("")),strT1(_T("")),strT2(_T(""));
	strRlt=strSrc;
	int nIndex=strRlt.Find(';');
	vector<CString>  vctStrSrc;
	if(nIndex!=-1)
	{
		while(nIndex!=-1)
		{
			strT1=strRlt.Left(nIndex);
			strRlt=strRlt.Right(nIndex+1);
			vctStrSrc.push_back(strT1);
			nIndex=strRlt.Find(';');
		}
	}
	else//!只有一个数据的情况
		vctStrSrc.push_back(strRlt);

	for(std::vector<CString>::const_iterator p=vctStrSrc.begin();p!=vctStrSrc.end();++p)
	{
		pair<double,double> pairValue;
		nIndex=(*p).Find(',');
		if(nIndex!=-1)
		{
			strT1=(*p).Left(nIndex);
			strT2=(*p).Right(nIndex+1);
			pairValue.first=_tstof(strT1);
			pairValue.second=_tstof(strT2);
			vctPair.push_back(pairValue);
		}
	}
}

_MITC_BASIC_END