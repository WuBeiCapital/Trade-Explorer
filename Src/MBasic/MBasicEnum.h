/*! \file: MBasicEnum.h   版权所有 (c) 2002-2008 , 北京迈达斯技术有限公司 \n
* 功能描述：                            \n
* 编 制 者：echo	     完 成  日 期：2007-7-10 11:35:32 \n
* 修 改 者：echo	     最后修改日期： -  - \n
* 历史记录：V 00.00.00(每次修改升级最后一个数字)  \n
*/
#pragma once

#include "MBasicMacro.h"

_MITC_BASIC_BEGIN

enum SerializeState
{
	SLS_ORG,
	SLS_ALL,
};

enum StateFlag   //!控制树节点上图标形态
{
	SF_UNFULFILMENT,//!未实现
	SF_SUCCESS,//!成功	
	SF_FAILING,//!失败
	SF_WARNING,//!警告
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
