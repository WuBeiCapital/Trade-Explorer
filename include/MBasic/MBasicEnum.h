/*! \file: MBasicEnum.h   ��Ȩ���� (c) 2002-2008 , ��������˹�������޹�˾ \n
* ����������                            \n
* �� �� �ߣ�echo	     �� ��  �� �ڣ�2007-7-10 11:35:32 \n
* �� �� �ߣ�echo	     ����޸����ڣ� -  - \n
* ��ʷ��¼��V 00.00.00(ÿ���޸��������һ������)  \n
*/
#pragma once

#include "MBasicMacro.h"

_MITC_BASIC_BEGIN

enum SerializeState
{
	SLS_ORG,
	SLS_ALL,
};

enum StateFlag   //!�������ڵ���ͼ����̬
{
	SF_UNFULFILMENT,//!δʵ��
	SF_SUCCESS,//!�ɹ�	
	SF_FAILING,//!ʧ��
	SF_WARNING,//!����
};

enum DoInfoGrade
{
	PDG_PROJECT,
	PDG_ROUTE,
	PDG_BRIDGE,
	PDG_COMPONENT
};

class MITC_BASIC_EXT CProjectDoInfo
{
public:	
	CProjectDoInfo()
	{
		m_enStateFlag=SF_SUCCESS;
		m_enDoInfoGrade=PDG_COMPONENT;
	};
	~CProjectDoInfo(){};

	StateFlag m_enStateFlag;
	DoInfoGrade m_enDoInfoGrade;

	std::vector<CString> m_vctStrings;
};
enum SQLOPTYPE
{
	SQLOT_INSERT,
	SQLOT_UPDATE,
	SQLOT_DELETE,
};
_MITC_BASIC_END
