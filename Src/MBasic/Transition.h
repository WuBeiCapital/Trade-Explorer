#pragma once

_MITC_BASIC_BEGIN
/*! \fn
* 函数功能：把vector中的值赋给list \n 
* 输入参数：const std::vector<_T>& val1,std::list<_T>& val2 \n
* 输出参数：bool型变量 \n 
* 返 回 值：TRUEorFALSE \n 
*/
template<class T> inline bool vct2lst(const std::vector<T>& vctSrc,std::list<T>& lstDec)
{
	if(vctSrc.empty())
		return false;

	lstDec.clear();
	copy(vctSrc.begin(),vctSrc.end(),back_inserter(lstDec));
	return true;
}


template<class T> inline bool vct2lst(const std::vector<T*>& vctSrc,std::list<T*>& lstDec)
{
	if(vctSrc.empty())
		return false;

	clearlst(lstDec);
	copy(vctSrc.begin(),vctSrc.end(),back_inserter(lstDec));
	return true;
}

template<class T> inline bool vct2lst(const std::vector<T*>& vctSrc,std::list<const T*>& lstDec)
{
	if(vctSrc.empty())
		return false;

	clearlst(lstDec);
	copy(vctSrc.begin(),vctSrc.end(),back_inserter(lstDec));
	return true;
}


/*! \fn
* 函数功能：把list中的值赋给vector \n  
* 输入参数：const std::list<_T>& val1,std::vector<_T>& val2 \n
* 输出参数：const std::list<_T>& val1 \n
* 返 回 值：bool型变量 返回 TRUE \n
*/
template<class T>  inline bool lst2vct(const std::list<T>& lstSrc,std::vector<T>& vctDec)
{
	if(lstSrc.empty())
		return false;

	vctDec.clear();
	copy(lstSrc.begin(),lstSrc.end(),back_inserter(vctDec));
	return true;
}

template<class T>  inline bool lst2vct(const std::list<T*>& lstSrc,std::vector<T*>& vctDec)
{
	if(lstSrc.empty())
		return false;

	clearvct(vctDec);
	copy(lstSrc.begin(),lstSrc.end(),back_inserter(vctDec));
	return true;
}

template<class T>  inline bool lst2vct(const std::list<T*>& lstSrc,std::vector<const T*>& vctDec)
{
	if(lstSrc.empty())
		return false;

	clearvct(vctDec);
	copy(lstSrc.begin(),lstSrc.end(),back_inserter(vctDec));
	return true;
}

MITC_BASIC_EXT CString PairDbl2Str(const std::vector<pair<double,double> >& vctPair);
MITC_BASIC_EXT void Str2PairDbl(const CString& strSrc,std::vector<pair<double,double> >& vctPair);


_MITC_BASIC_END