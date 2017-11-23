#pragma once

_MITC_BASIC_BEGIN

//! ------------------���������͵����л�-------------------
//!���û�������
template <typename _Ty> void Dump(std::ofstream &out, _Ty& tValue)
{
	out<<#tValue<< "\t"<<tValue <<"\n";
}
//!������pair���ͣ���������
template <typename _Ty1,typename _Ty2> void Dump(std::ofstream &out, std::pair<_Ty1,_Ty2> pairValue)
{
	out << pairValue.first << ","<< pairValue.first<<"\n";
}
//!����C++��������
template<typename _Ty>inline void Dump(std::ofstream &out,std::vector<_Ty>& vctTs)
{
	Dump(out,vctTs.size);
	for(int i=0;i< vctTs.size();++i)
	{
		out<< "["<<i<< "]"<<"\t"<<vctTs.at(i)<<"\n";
	}
}
//!�����Զ�����չ����
template<typename _Ty> inline void DumpEx(std::ofstream &out,std::vector<_Ty>& vctTs)
{
	Dump(out,vctTs.size);
	for(std::vector<_Ty>::iterator p=vctTs.begin();p!=vctTs.end();++p)
	{
		(*p).Dump(out);
	}	
}

//!�����Զ�����չ����ָ��
template<typename _Ty> inline void Dump(std::ofstream &out,std::vector<_Ty*>& vctTs)
{
	Dump(out,vctTs.size);
	for(std::vector<_Ty*>::iterator p=vctTs.begin();p!=vctTs.end();++p)
	{
		(*p)->Dump(out);
	}
}
//!����C++��������
template<typename _Ty> inline void Dump(std::ofstream &out,std::list<_Ty>& lstTs)
{
	Dump(out,lstTs.size);
	int i(0);
	for(std::list<_Ty>::iterator p=lstTs.begin();p!=lstTs.end();++p,++i)
	{
		out<< "["<<i<< "]"<<"\t"<<*p<<"\n";	
	}
}

//!�����Զ�����չ����
template<typename _Ty> inline void DumpEx(std::ofstream &out,std::list<_Ty>& lstTs)
{
	Dump(out,lstTs.size);
	int i(0);
	for(std::list<_Ty>::iterator p=lstTs.begin();p!=lstTs.end();++p,++i)
	{
		(*p).Dump(out);	
	}
}

//!�����Զ�����չ����ָ�룬��Ĭ�Ϲ��캯��
template<typename _Ty> inline void Dump(std::ofstream &out,std::list<_Ty*>& lstTs)
{
	Dump(out,lstTs.size);
	int i(0);
	for(std::list<_Ty>::iterator p=lstTs.begin();p!=lstTs.end();++p,++i)
	{
		(*p)->Dump(out);	
	}
}

//!����T2 C++��������
template<typename _Kty,class _Ty> inline void Dump(std::ofstream &out,std::map<_Kty, _Ty>& mapTs)
{
	Dump(out,mapTs.size);
	for(std::map<_Kty,_Ty>::iterator p=mapTs.begin();p!=mapTs.end();++p)
	{
		out<<"["<<p->first<<"]"<<"\t"<<p->second<<"\n";			
	}	
}

//!����T2 �Զ�����չ����
template<typename _Kty,typename _Ty> inline void DumpEx(std::ofstream &out,std::map<_Kty, _Ty>& mapTs)
{
	Dump(out,mapTs.size);
	for(std::map<_Kty,_Ty>::iterator p=mapTs.begin();p!=mapTs.end();++p)
	{
		out<<"["<<p->first<<"]"<<"\t"<<"\n";	
		(p->second).Dump(out);
	}	
}
//!�����Զ�����չ����ָ�룬��Ĭ�Ϲ��캯��
template<typename _Kty,typename _Ty> inline void Dump(std::ofstream &out,std::map<_Kty, _Ty*>& mapTs)
{
	Dump(out,mapTs.size);
	for(std::map<_Kty,_Ty>::iterator p=mapTs.begin();p!=mapTs.end();++p)
	{
		out<<"["<<p->first<<"]"<<"\t"<<"\n";	
		(p->second)->Dump(out);
	}
}

///************************************************************************************************************
///************************************************************************************************************
///************************************************************************************************************

_MITC_BASIC_END