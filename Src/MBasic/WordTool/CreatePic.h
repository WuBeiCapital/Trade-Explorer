#pragma once

#include <gdiplus.h>
using namespace Gdiplus;
#pragma comment( lib, "gdiplus.lib" )



_MITC_BASIC_BEGIN


enum ENDIJ
{
  END_I ,
  END_J
};
enum VALUETYPE
{
  ALLOW_VAL ,
  REAL_VAL
};
enum MAXMINTOPBOT  //受力最大最小
{
  F_NULL,
  F_MAXTOP,
  F_MINBOT
};

// enum TOPBOT  //受力定底面
// {
//   P_NULL,
//   P_TOP,
//   P_BOT
// };

class MITC_BASIC_EXT CGraphKey
{

public:

  CGraphKey(void)
  {
    Init();
  }
  ~CGraphKey(void){};
  void Init()
  {
    iElemID = 0 ;
    ePos = END_I;
    eXn  = F_NULL;

    eValType = REAL_VAL;
  }

  CGraphKey& operator=(const CGraphKey& src)
  {
    if(this == &src)       
      return *this;

    //添加赋值代码
    iElemID       = src.iElemID    ;
    ePos      = src.ePos   ;//!
    eXn       = src.eXn    ;

    eValType  = src.eValType    ;

    return *this;
  }

  bool operator <  (const CGraphKey& ptr) const
  {
    bool blnReturn = false;
    if (iElemID == ptr.iElemID)
    {
      if (ePos == ptr.ePos)
      {
        if (eXn == ptr.eXn)
        {

          blnReturn = eValType < ptr.eValType;
        }
        else
        {
          blnReturn = (eXn < ptr.eXn);
        }
      }
      else
      {
        blnReturn = (ePos < ptr.ePos);
      }
    }
    else
    {
      blnReturn = (iElemID < ptr.iElemID);
    }


    return blnReturn;
  }

  bool operator >  (const CGraphKey& ptr) const
  {
    bool blnReturn = false;
    if (iElemID == ptr.iElemID)
    {
      if (ePos == ptr.ePos)
      {
        if (eXn == ptr.eXn)
        {

          blnReturn = eValType > ptr.eValType;
        }
        else
        {
          blnReturn = (eXn > ptr.eXn);
        }
      }
      else
      {
        blnReturn = (ePos > ptr.ePos);
      }
    }
    else
    {
      blnReturn = (iElemID > ptr.iElemID);
    }

    return blnReturn;

  }

  bool operator ==  (const CGraphKey& ptr) const
  {
    bool blnReturn = false;

    blnReturn = (ePos == ptr.ePos) && (eXn == ptr.eXn) && 
      (iElemID == ptr.iElemID) && (eValType == ptr.eValType);//!

    return blnReturn;
  }
  bool operator !=  (const CGraphKey& ptr) const
  {
    bool blnReturn = !(this->operator ==(ptr));

    return blnReturn;
  }

public:
  UINT iElemID;
  ENDIJ ePos ;
  MAXMINTOPBOT eXn;
  VALUETYPE eValType ;


};

class MITC_BASIC_EXT CCreatePic
{
public:
  CCreatePic(void);
  ~CCreatePic(void);

  CString CreatGraph(const map<CGraphKey,double>& srcData,const CString& strX_Name,const CString& strY_Name,const map<CGraphKey,CString>& strTLS,const CString& picName);
  void Create(bool isTest = true);
protected:
  ULONG_PTR m_gdiplusToken;

  UINT m_uDrawZoom_Width;
  UINT m_uDrawZoom_Height;

  double m_dX_e;  //x轴方向的单位长度
  double m_dY_e;  //y轴方向的单位长度

  double m_dX_r;  //x轴方向的实际单位长度
  double m_dY_r;  //y轴方向的实际单位长度

  CString str_X_Name ;   //X轴标注
  CString str_Y_Name ;   //Y轴标注

  double m_dMax_X;
  double m_dMax_Y;


  map<CGraphKey,CString> m_strTL;  //图例说明

  vector<Color> m_TLColors;
  vector<DashStyle> m_TLLineStyle;


  map<CGraphKey,double> m_srcData;

  double m_dY_Less ;   //控制y轴上下预留的空间
  double m_dX_Less ;   //控制x轴右侧预留的空间
  UINT m_uYPts ;   //控制y轴上标记点的个数

  CString m_strOutPath ;

protected:
  void DrawAxis(Graphics* graphics,Gdiplus::Point O_Pt,UINT xMin,UINT xMax,double yMin,double yMax,double yMinT);
  void DrawPot(Graphics* graphics,Gdiplus::Point O_Pt,UINT xMin,UINT xMax,double yMinT);
  Gdiplus::Point ChangeGDIPlusCoord(const Gdiplus::Point& pt);
  Gdiplus::PointF ChangeGDIPlusCoord(const Gdiplus::PointF& pt);

  void GetAxisMaxMin(UINT& xMin,UINT& xMax,double& yMin,double& yMax);



};

_MITC_BASIC_END