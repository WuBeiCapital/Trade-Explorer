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
enum MAXMINTOPBOT  //���������С
{
  F_NULL,
  F_MAXTOP,
  F_MINBOT
};

// enum TOPBOT  //����������
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

    //��Ӹ�ֵ����
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

  double m_dX_e;  //x�᷽��ĵ�λ����
  double m_dY_e;  //y�᷽��ĵ�λ����

  double m_dX_r;  //x�᷽���ʵ�ʵ�λ����
  double m_dY_r;  //y�᷽���ʵ�ʵ�λ����

  CString str_X_Name ;   //X���ע
  CString str_Y_Name ;   //Y���ע

  double m_dMax_X;
  double m_dMax_Y;


  map<CGraphKey,CString> m_strTL;  //ͼ��˵��

  vector<Color> m_TLColors;
  vector<DashStyle> m_TLLineStyle;


  map<CGraphKey,double> m_srcData;

  double m_dY_Less ;   //����y������Ԥ���Ŀռ�
  double m_dX_Less ;   //����x���Ҳ�Ԥ���Ŀռ�
  UINT m_uYPts ;   //����y���ϱ�ǵ�ĸ���

  CString m_strOutPath ;

protected:
  void DrawAxis(Graphics* graphics,Gdiplus::Point O_Pt,UINT xMin,UINT xMax,double yMin,double yMax,double yMinT);
  void DrawPot(Graphics* graphics,Gdiplus::Point O_Pt,UINT xMin,UINT xMax,double yMinT);
  Gdiplus::Point ChangeGDIPlusCoord(const Gdiplus::Point& pt);
  Gdiplus::PointF ChangeGDIPlusCoord(const Gdiplus::PointF& pt);

  void GetAxisMaxMin(UINT& xMin,UINT& xMax,double& yMin,double& yMax);



};

_MITC_BASIC_END