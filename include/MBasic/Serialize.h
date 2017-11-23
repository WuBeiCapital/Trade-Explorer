#pragma once

_MITC_BASIC_BEGIN

//! ------------------各数据类型的序列化-------------------
//!适用于枚举类型
template <typename _Ty> void Serialize(CArchive& ar, _Ty& enumValue)
{
	if(ar.IsStoring())
	{
		int temp=static_cast<int>(enumValue);
		ar<< temp;
	}
	else
	{
		int temp;
		ar>> temp;
		enumValue= static_cast<_Ty>(temp);
	}
}

//!适用于pair类型，基本类型
template <typename _Ty1,typename _Ty2> void Serialize(CArchive& ar, std::pair<_Ty1,_Ty2> &pairValue)
{
	if(ar.IsStoring())
	{
		ar<< pairValue.first;
		ar<< pairValue.second;
	}
	else
	{
		_Ty1 temp1;
		ar>> temp1;
		_Ty2 temp2;
		ar>> temp2;
		pairValue= make_pair(temp1,temp2);
	}
}

//!适用C++基本类型
template<typename _Ty>inline void Serialize(CArchive& ar,std::vector<_Ty>& vctTs)
{
	int dSize(0);
	if(ar.IsStoring())
	{
		ar<<(UINT)vctTs.size();
		for(std::vector<_Ty>::iterator p=vctTs.begin();p!=vctTs.end();++p)
		{
			ar<<*p;
		}
	}
	else
	{
		ar>>dSize;
		for(int i=0; i<dSize;++i)
		{
			_Ty tTemp;
			ar>>tTemp;
			vctTs.push_back(tTemp);
		}
	}
}
//!适用自定义扩展类型
template<typename _Ty> inline void SerializeEx(CArchive& ar,std::vector<_Ty>& vctTs)
{
	int dSize(0);
	if(ar.IsStoring())
	{
		ar<<(UINT)vctTs.size();
		for(std::vector<_Ty>::iterator p=vctTs.begin();p!=vctTs.end();++p)
		{
			(*p).Serialize(ar);
		}
	}
	else
	{
		ar>>dSize;
		for(int i=0; i<dSize;++i)
		{
			_Ty tTemp;
			tTemp.Serialize(ar);
			vctTs.push_back(tTemp);
		}
	}
}

//!适用自定义扩展类型指针，有默认构造函数
template<typename _Ty> inline void Serialize(CArchive& ar,std::vector<_Ty*>& vctTs)
{
	int dSize(0);
	if(ar.IsStoring())
	{
		ar<<(UINT)vctTs.size();
		for(std::vector<_Ty*>::iterator p=vctTs.begin();p!=vctTs.end();++p)
		{
			(*p)->Serialize(ar);
		}
	}
	else
	{
		ar>>dSize;
		for(int i=0; i<dSize;++i)
		{
			_Ty* tTemp=new _Ty;
			tTemp->Serialize(ar);
			vctTs.push_back(tTemp);
		}
	}
}
//!适用C++基本类型
template<typename _Ty> inline void Serialize(CArchive& ar,std::list<_Ty>& lstTs)
{
	int dSize(0);
	if(ar.IsStoring())
	{
		ar<<(UINT)lstTs.size();
		for(std::list<_Ty>::iterator p=lstTs.begin();p!=lstTs.end();++p)
		{
			ar<<*p;
		}
	}
	else
	{
		ar>>dSize;
		for(int i=0; i<dSize;++i)
		{
			_Ty tTemp;
			ar>>tTemp;
			lstTs.push_back(tTemp);
		}
	}
}

//!适用自定义扩展类型
template<typename _Ty> inline void SerializeEx(CArchive& ar,std::list<_Ty>& lstTs)
{
	int dSize(0);
	if(ar.IsStoring())
	{
		ar<<(UINT)lstTs.size();
		for(std::list<_Ty>::iterator p=lstTs.begin();p!=lstTs.end();++p)
		{
			(*p).Serialize(ar);
		}
	}
	else
	{
		ar>>dSize;
		for(int i=0; i<dSize;++i)
		{
			_Ty tTemp;		
			tTemp.Serialize(ar);
			lstTs.push_back(tTemp);
		}
	}
}

//!适用自定义扩展类型指针，有默认构造函数
template<typename _Ty> inline void Serialize(CArchive& ar,std::list<_Ty*>& lstTs)
{
	int dSize(0);
	if(ar.IsStoring())
	{
		ar<<(UINT)lstTs.size();
		for(std::list<_Ty*>::iterator p=lstTs.begin();p!=lstTs.end();++p)
		{
			(*p)->Serialize(ar);
		}
	}
	else
	{
		ar>>dSize;
		for(int i=0; i<dSize;++i)
		{
			_Ty* tTemp=new _Ty;		
			tTemp->Serialize(ar);
			lstTs.push_back(tTemp);
		}
	}
}

//!适用T2 C++基本类型
template<typename _Kty,class _Ty> inline void Serialize(CArchive& ar,std::map<_Kty, _Ty>& mapTs)
{
	int dSize(0);
	if(ar.IsStoring())
	{
		ar<<(UINT)mapTs.size();
		for(std::map<_Kty,_Ty>::iterator p=mapTs.begin();p!=mapTs.end();++p)
		{
			ar<<p->first;
			ar<<p->second;
		}
	}
	else
	{
		ar>>dSize;
		for(int i=0; i<dSize;++i)
		{
			_Kty tTemp;
			ar>>tTemp;
			_Ty tTemp2;
			ar>>tTemp2;
			mapTs[tTemp]=tTemp2;
		}
	}
}
//!适用T2 基本类型
template<typename _Kty,typename _Ty> inline void SerializeEx(CArchive& ar,std::map<_Kty, std::vector<_Ty> >& mapTs)
{
	int dSize(0);
	if(ar.IsStoring())
	{
		ar<<(UINT)mapTs.size();
		for(std::map<_Kty,std::vector<_Ty> >::iterator p=mapTs.begin();p!=mapTs.end();++p)
		{
			ar<<p->first;
			MBasic::Serialize(ar,p->second);
		}
	}
	else
	{
		ar>>dSize;
		for(int i=0; i<dSize;++i)
		{
			_Kty tTemp;
			ar>>tTemp;
			std::vector<_Ty> tTemp2;

			MBasic::Serialize(ar,tTemp2);
			mapTs[tTemp]=tTemp2;
		}
	}
}
//!适用自定义扩展类型指针，有默认构造函数
template<typename _Kty,typename _Ty> inline void Serialize(CArchive& ar,std::map<_Kty, _Ty*>& mapTs)
{
	int dSize(0);
	if(ar.IsStoring())
	{
		ar<<(UINT)mapTs.size();
		for(std::map<_Kty,_Ty*>::iterator p=mapTs.begin();p!=mapTs.end();++p)
		{
			ar<<p->first;
			p->second->Serialize(ar);
		}
	}
	else
	{
		ar>>dSize;
		for(int i=0; i<(int)dSize;++i)
		{
			_Kty tTemp;
			ar>>tTemp;
			_Ty* tTemp2=new _Ty;
			tTemp2->Serialize(ar);
			mapTs[tTemp]=tTemp2;
		}
	}
}

//!适用T2 基本类型_Ty 扩展类型_ty2自定义扩展类型指针，有默认构造函数
template<typename _Kty,typename _Ty,typename _ty2> inline void Serialize(CArchive& ar,std::map<_Kty, 
																		   std::map<_Ty,_ty2*> >& mapTs)
{
	int dSize(0);
	if(ar.IsStoring())
	{
		ar<<(UINT)mapTs.size();
		for(std::map<_Kty,map<_Ty,_ty2*> >::iterator p=mapTs.begin();p!=mapTs.end();++p)
		{
			ar<<p->first;
			MBasic::Serialize(ar,p->second);
		}
	}
	else
	{
		ar>>dSize;
		for(int i=0; i<dSize;++i)
		{
			_Kty tTemp;
			ar>>tTemp;

			std::map<_Ty,_ty2*> tTemp2;

			MBasic::Serialize(ar,tTemp2);
			mapTs[tTemp]=tTemp2;
		}
	}
}

//!适用自定义扩展类型指针，有默认构造函数
template<typename _Kty,typename _Ty> inline void Serialize(CArchive& ar,std::map<_Kty, std::vector<_Ty*> >& mapTs)
{
	int dSize(0);
	if(ar.IsStoring())
	{
		ar<<(UINT)mapTs.size();
		for(std::map<_Kty,std::vector<_Ty*>>::iterator p=mapTs.begin();p!=mapTs.end();++p)
		{
			ar<<p->first;
			MBasic::Serialize(ar,p->second);			
		}
	}
	else
	{
		ar>>dSize;
		for(int i=0; i<dSize;++i)
		{
			_Kty tTemp;
			ar>>tTemp;
			std::vector<_Ty*> tTemp2;
			MBasic::Serialize(ar,tTemp2);
			mapTs[tTemp]=tTemp2;
		}
	}
}

//!适用T2 自定义扩展类型
template<typename _Kty,typename _Ty> inline void SerializeEx(CArchive& ar,std::map<_Kty, _Ty>& mapTs)
{
	int dSize(0);
	if(ar.IsStoring())
	{
		ar<<(UINT)mapTs.size();
		for(std::map<_Kty,_Ty>::iterator p=mapTs.begin();p!=mapTs.end();++p)
		{
			ar<<p->first;
			p->second.Serialize(ar);
		}
	}
	else
	{
		ar>>dSize;
		for(int i=0; i<dSize;++i)
		{
			_Kty tTemp;
			ar>>tTemp;
			_Ty tTemp2;
			tTemp2.Serialize(ar);
			mapTs[tTemp]=tTemp2;
		}
	}
}

//////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////// 操作符 << 和 >> 的重载 ////////////////////////
//////////////////////////////////////////////////////////////////////////////////////

//! 适用于C++基本类型或已重载序列化的自定义类型
template <typename _Fst,typename _Scd> CArchive& operator << (CArchive &ar,const pair<_Fst,_Scd> &pairVal)
{
	ar<<pairVal.first<<pairVal.second;

	return ar;
}

template <typename _Fst,typename _Scd> CArchive& operator >> (CArchive &ar,pair<_Fst,_Scd> &pairVal)
{
	ar>>pairVal.first>>pairVal.second;
	return ar;
}

//! 适用于C++基本类型或已重载序列化的自定义类型
template <typename _Ty> CArchive& operator << (CArchive &ar,const vector<_Ty> &vctVals)
{
	size_t nValsCount = vctVals.size();
	ar<<(unsigned int)nValsCount;
	for (size_t nValIndex = 0 ; nValIndex < nValsCount ; ++nValIndex)
	{
		ar<<vctVals.at(nValIndex);
	}
	return ar;
}
template <typename _Ty> CArchive& operator >> (CArchive &ar,vector<_Ty> &vctVals)
{
	vctVals.clear();
	size_t nValsCount = 0;
	ar>>nValsCount;
	for (size_t nValIndex = 0 ; nValIndex < nValsCount ; ++nValIndex)
	{
		_Ty Val;
		ar>>Val;
		vctVals.push_back(Val);
	}

	return ar;
}
//! 适用于C++基本类型或已重载序列化的自定义类型
template <typename _Ty> CArchive& operator << (CArchive &ar,const list<_Ty> &lstVals)
{
	size_t nValsCount = lstVals.size();
	ar<<(unsigned int)nValsCount;
	for (list<_Ty>::const_iterator vit = lstVals.begin() ; vit != lstVals.end() ; ++vit)
	{
		ar<<*vit;
	}

	return ar;
}
template <typename _Ty> CArchive& operator >> (CArchive &ar,list<_Ty> &lstVals)
{
	lstVals.clear();
	size_t nValsCount = 0;
	ar>>nValsCount;
	for (size_t nValIndex = 0 ; nValIndex < nValsCount ; ++nValIndex)
	{
		_Ty Val;
		ar>>Val;
		lstVals.push_back(Val);
	}

	return ar;
}

//! 适用于C++基本类型或已重载序列化的自定义类型
template <typename _key,typename _val> CArchive& operator << (CArchive &ar,const map<_key,_val> &mapVals)
{
	size_t nValsCount = mapVals.size();
	ar<<(unsigned int)nValsCount;
	for (map<_key,_val>::const_iterator vit = mapVals.begin() ; vit != mapVals.end() ; ++vit)
	{
		ar<<vit->first;
		ar<<vit->second;
	}
	return ar;
}

template <typename _key,typename _val> CArchive& operator >> (CArchive &ar,map<_key,_val> &mapVals)
{
	mapVals.clear();
	size_t nValsCount = 0;
	ar>>nValsCount;
	for (size_t nValIndex = 0 ; nValIndex < nValsCount ; ++nValIndex)
	{
		_key key;
		ar>>key;
		_val val;
		ar>>val;
		mapVals.insert(make_pair(key,val));
	}
	return ar;
}

_MITC_BASIC_END