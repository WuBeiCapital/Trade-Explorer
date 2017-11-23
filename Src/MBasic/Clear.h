#pragma once

_MITC_BASIC_BEGIN

template<typename _Ty>  inline bool clearvct(std::vector<_Ty*>& vctVal)
{
	for(std::vector<_Ty*>::iterator p=vctVal.begin();p!=vctVal.end();++p)
	{
		delete static_cast<_Ty*>(*p);
		*p=NULL;
	}
	vctVal.clear();
	return true;
}

template<typename _Ty>  inline bool clearlst(std::list<_Ty*>& lstVal)
{
	for(std::list<_Ty*>::iterator p=lstVal.begin();p!=lstVal.end();++p)
	{
		delete static_cast<_Ty*>(*p);
		*p=NULL;
	}
	lstVal.clear();
	return true;
}

template<typename _Ty1,typename _Ty2>  inline bool clearmap(std::map<_Ty1,_Ty2*>& mapVal)
{
	for(std::map<_Ty1,_Ty2*>::iterator p=mapVal.begin();p!=mapVal.end();++p)
	{
		delete static_cast<_Ty2*>(p->second);
		static_cast<_Ty2*>(p->second)=NULL;
	}
	mapVal.clear();
	return true;
}

#define DeleteObj(X) \
if((X)){delete (X);(X)=NULL;}\

_MITC_BASIC_END