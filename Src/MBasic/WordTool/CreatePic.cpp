#include "stdafx.h"
#include "CreatePic.h"
#include "..\basictools.h"

_MITC_BASIC_BEGIN


CCreatePic::CCreatePic(void)
{
  Gdiplus::GdiplusStartupInput gdiplusStartupInput;
  Gdiplus::GdiplusStartup(&m_gdiplusToken, &gdiplusStartupInput, NULL);

  m_uDrawZoom_Width = 1000;
  m_uDrawZoom_Height = 250;

  m_dY_Less = 30 ;
  m_dX_Less = 150 ;

  m_uYPts = 10 ;

  str_X_Name = L"";
  str_Y_Name = L"";

//   m_strTL.push_back(L"Vn");
//   m_strTL.push_back(L"dVn");


  m_TLColors.push_back(Color::Blue);
  m_TLColors.push_back(Color::Red);
  m_TLColors.push_back(Color::Green);
  m_TLColors.push_back(Color::Brown);
  m_TLColors.push_back(Color::Black);
  m_TLColors.push_back(Color::Gold);

	m_TLLineStyle.push_back(DashStyleSolid);           // 1
	m_TLLineStyle.push_back(DashStyleDash );          // 2
	m_TLLineStyle.push_back(DashStyleSolid  );      // 3
	m_TLLineStyle.push_back(DashStyleDash  );   // 4
	m_TLLineStyle.push_back(DashStyleCustom  );        // 5



}

void CCreatePic::GetAxisMaxMin(UINT& xMin,UINT& xMax,double& yMin,double& yMax)
{
  yMin =  1E+60;
  yMax = -1E+60;

  for (map<CGraphKey,double>::iterator pit = m_srcData.begin();pit!=m_srcData.end();++pit)
  {
    if(pit == m_srcData.begin())
    {
      xMin = pit->first.iElemID;
    }
    else if(pit == --m_srcData.end())
    {
      xMax = pit->first.iElemID;
    }

    if(pit->second - yMax > 1E-7)
    {
      yMax = pit->second;
    }

    if(yMin - pit->second > 1E-7)
    {
      yMin = pit->second;
    }
}
}


CCreatePic::~CCreatePic(void)
{
  Gdiplus::GdiplusShutdown(m_gdiplusToken);

}

int   GetCodecClsid(const   WCHAR*   format,   CLSID*   pClsid)   
{   
  UINT   num   =   0;     //   number   of   image   encoders   
  UINT   size   =   0;   //   size   of   the   image   encoder   array   in   bytes   

  ImageCodecInfo*   pImageCodecInfo   =   NULL;   

  GetImageEncodersSize(&num,   &size);   
  if(size   ==   0)   
    return   -1;   //   Failure   

  pImageCodecInfo   =   (ImageCodecInfo*)(malloc(size));   
  if(pImageCodecInfo   ==   NULL)   
    return   -1;   //   Failure   

  GetImageEncoders(num,   size,   pImageCodecInfo);   

  for(UINT   j   =   0;   j   <   num;   ++j)   
  {   
    if(   wcscmp(pImageCodecInfo[j].MimeType,   format)   ==   0   )   
    {   
      *pClsid   =   pImageCodecInfo[j].Clsid;   
      return   j;   //   Success   
    }   
  }   //   for   

  return   -1;   //   Failure   
}   //   GetCodecClsid   

void CCreatePic::Create(bool isTest)
{




  Bitmap _bufferImage(m_uDrawZoom_Width,m_uDrawZoom_Height);
  Graphics   graphics(&_bufferImage);

  graphics.SetTextRenderingHint(TextRenderingHintAntiAliasGridFit);
  //g.SetTextRenderingHint(TextRenderingHintSingleBitPerPixel);   //可以调节 
  graphics.SetSmoothingMode(SmoothingModeAntiAlias);

  //Matrix mat(1,0,0,-1,0,m_uDrawZoom_Height);

  //graphics.SetTransform(&mat);


  SolidBrush WhiteBrush(Color::White) ; 
  graphics.FillRectangle(&WhiteBrush, 0, 0, m_uDrawZoom_Width, m_uDrawZoom_Height);


  UINT xMin = 0;
  UINT xMax=0;
  double yMin=0.0;
  double yMax=0.0;
  Point O_Pt(80,0);
GetAxisMaxMin(xMin,xMax,yMin,yMax);

double yMaxT =yMax > 100 ? (int)(yMax/100 + 1) * 100 : yMax  ;
double yMinT =yMin < -100 ?  (int)(yMin/100 - 1) * 100 : yMin ;

double dxValidLength = m_uDrawZoom_Width - O_Pt.X - m_dX_Less ;
double dyValidLength = m_uDrawZoom_Height - 40;

m_dX_e = dxValidLength/(xMax - xMin + 1 + 0.5 );

m_dY_e = dyValidLength/(m_uYPts + 0.5);
m_dY_r = (yMaxT - yMinT) / (m_uYPts ) ;

O_Pt.Y = (-1.0 * yMinT) * m_dY_e / m_dY_r + m_dY_Less ;
if(O_Pt.Y < m_dY_Less)
  O_Pt.Y = m_dY_Less ;


DrawAxis(&graphics,O_Pt,xMin,xMax,yMin,yMax,yMinT);
DrawPot(&graphics,O_Pt,xMin,xMax,yMinT);


  CLSID   BmpCodec;  
  GetCodecClsid(L"image/jpeg",   &BmpCodec);   
  _bufferImage.Save(m_strOutPath,&BmpCodec,   NULL);  
}

void CCreatePic::DrawAxis(Graphics* graphics,Gdiplus::Point O_Pt,UINT xMin,UINT xMax,double yMin,double yMax,double yMinT)
{
  Color Axis_Color = Color::Black;
  UINT  Axis_Width = 2;

  SolidBrush fBrush(Color::Blue) ; 

  Point ptXStart(O_Pt.X-m_dX_Less,O_Pt.Y);
  Point ptXEnd(m_uDrawZoom_Width-m_dX_Less,O_Pt.Y);
  Point ptYStart(O_Pt.X,0);
  Point ptYEnd(O_Pt.X,m_uDrawZoom_Height - 20);

  int iJZjL = 0 ;
  if(m_uDrawZoom_Height - 20 < O_Pt.Y)
  {
	ptXStart.Y = m_uDrawZoom_Height - 20;
	ptXEnd.Y = m_uDrawZoom_Height - 20;

	iJZjL = O_Pt.Y - m_uDrawZoom_Height + 20 ;

  }



  ptXStart = ChangeGDIPlusCoord(ptXStart);
  ptXEnd = ChangeGDIPlusCoord(ptXEnd);
  ptYStart = ChangeGDIPlusCoord(ptYStart);
  ptYEnd = ChangeGDIPlusCoord(ptYEnd);




  Pen AxisPen(Axis_Color, Axis_Width);

  Pen AxisNetPen(Axis_Color, 1);
  AxisNetPen.SetDashStyle(DashStyleDash);


  GraphicsPath   endPath;   //   创建起点和终点路径对象 
  Point   polygonPoints[4]   =   {Point(-2,   0),   Point(2,   0),   Point(0,   10)}; 
  endPath.AddPolygon(polygonPoints,   3);//   终点箭头 
  CustomLineCap   endCap(NULL,   &endPath);   //   创建终点线帽 


  FontFamily fontfamily(L"Times New Roman");
  Gdiplus::Font font(&fontfamily,12,FontStyleRegular,UnitPixel);

  Gdiplus::Font   myFont(L"隶书",   16);   
  
  StringFormat   formatX;   
  formatX.SetAlignment(StringAlignmentCenter);   
  StringFormat   formatY;   
  formatY.SetAlignment(StringAlignmentCenter);   

  SolidBrush   blackBrush(Color(255,   0,   0,   0));   


  Pen AxisBiaoPen(Axis_Color, 1);

  AxisPen.SetCustomEndCap(&endCap);


  graphics->DrawLine(&AxisPen,ptXStart,ptXEnd);
  graphics->DrawLine(&AxisPen,ptYStart,ptYEnd);

  CString strAxisPt = L"";
  double dmaxX = 0.0 ;
  int iSkip = 1;
  iSkip = (xMax - xMin + 1)/45 + 1 ;
  for(UINT i=xMin;i<=xMax;i++)
  {
	if((i-xMin)%iSkip == 0)
	{
    strAxisPt.Format(L"%u",i);
    dmaxX = O_Pt.X + (i-xMin+1)*m_dX_e ;
    graphics->DrawString(strAxisPt,   -1,   &font,   ChangeGDIPlusCoord(PointF(dmaxX-0.5*m_dX_e,O_Pt.Y-5-iJZjL)),  &formatX ,&fBrush); 
    }
  }
  m_dMax_X = dmaxX;
  graphics->DrawString(str_X_Name,   -1,   &font, PointF(ptXEnd.X,ptXEnd.Y+10)  ,  &fBrush); 

  double dmaxY = 0.0 ;
  double dminY = 0.0 ;

  for (UINT i= 0;i<=m_uYPts;i++)
  {
      double dmaxyyy = max(fabs(yMinT),fabs(yMinT + m_uYPts*m_dY_r));
	  if(fabs(dmaxyyy) > 10)
	  {
		  strAxisPt.Format(L"%0.0f",yMinT + i*m_dY_r) ;

	  }
	  else if(fabs(dmaxyyy) > 1)
	  {
		  strAxisPt.Format(L"%0.1f",yMinT + i*m_dY_r) ;

	  }
	  else if(fabs(dmaxyyy) > 0.1)
	  {
		  strAxisPt.Format(L"%0.2f",yMinT + i*m_dY_r) ;

	  }
	  else 
	  {
		  strAxisPt.Format(L"%0.3f",yMinT + i*m_dY_r) ;

	  }

    SizeF sizeLayout(80,20);
    RectF layoutRect(ChangeGDIPlusCoord(PointF(O_Pt.X-sizeLayout.Width-5,m_dY_Less + m_dY_e * i + 0.5*sizeLayout.Height)),sizeLayout);   

    graphics->DrawString(   
      strAxisPt,   
      -1,   
      &font,   
      layoutRect,   
      &formatY,   
      &blackBrush);   
    //网格

    if(i == 0)
      dminY = ChangeGDIPlusCoord(Point(O_Pt.X,m_dY_Less + m_dY_e * (i))).Y  ;
    else if(i == m_uYPts)
      dmaxY = ChangeGDIPlusCoord(Point(O_Pt.X,m_dY_Less + m_dY_e * (i))).Y  ;

    graphics->DrawLine(&AxisNetPen,ChangeGDIPlusCoord(Point(O_Pt.X,m_dY_Less + m_dY_e * (i))),ChangeGDIPlusCoord(Point(dmaxX,m_dY_Less + m_dY_e * (i))));

  }

  graphics->DrawString(str_Y_Name,   -1,   &font, PointF(ptYEnd.X+5,ptYEnd.Y-20)  ,  &blackBrush); 



  //网格
  ///!x方向网格
  m_dMax_Y = dmaxY;
  for(UINT i=xMin;i<=xMax;i++)
  {
    Point pt1 = ChangeGDIPlusCoord(Point(O_Pt.X + (i-xMin+1)*m_dX_e,0));
    pt1.Y = dminY ;
    Point pt2 = ChangeGDIPlusCoord(Point(O_Pt.X + (i-xMin+1)*m_dX_e,0));
    pt2.Y = dmaxY ;
    graphics->DrawLine(&AxisNetPen,pt1,pt2);
  }



}
void CCreatePic::DrawPot(Graphics* graphics,Gdiplus::Point O_Pt,UINT xMin,UINT xMax,double yMinT)
{
  int SolidLineWidth = 2;

  double dx = 0.0;
  double dy = 0.0;

  Gdiplus::Point dotPt1;
  Gdiplus::Point dotPt2;
  // 
  map<CGraphKey,Gdiplus::Point> lastPoints ;
  map<ENDIJ,Gdiplus::Point> NodePos ;

  FontFamily fontfamily(L"Times New Roman");
  Gdiplus::Font font(&fontfamily,12,FontStyleRegular,UnitPixel);



  Gdiplus::Point lastPoint(ChangeGDIPlusCoord(O_Pt)) ;

  map<CGraphKey,double>::iterator pFindI = m_srcData.begin();
  map<CGraphKey,double>::iterator pFindJ = m_srcData.begin();
  CGraphKey findKey ;

  for (UINT i=F_MAXTOP ; i<=F_MINBOT;i++)
  {
    for(UINT j=ALLOW_VAL ; j<=REAL_VAL;j++)
    {
      Color POT_Color = m_TLColors[i-1+2*j];
      SolidBrush POT_Brush(POT_Color) ; 
      if(j==ALLOW_VAL)
        SolidLineWidth = 4;
      else
        SolidLineWidth = 2 ;

      Pen POT_Pen(POT_Color, SolidLineWidth);
	  POT_Pen.SetDashStyle(m_TLLineStyle[i-1+2*j]);

      for (UINT k = xMin ; k<=xMax ; k++)
      {
        dx = O_Pt.X + (k-xMin) * m_dX_e;
        NodePos[END_I].X = dx;
        NodePos[END_J].X = dx + m_dX_e;

        findKey.iElemID = k ;
        findKey.ePos = END_I ;
        findKey.eXn = (MAXMINTOPBOT)i;
        findKey.eValType = (VALUETYPE)j;

        pFindI = m_srcData.find(findKey) ;
        findKey.ePos = END_J ;
        pFindJ = m_srcData.find(findKey) ;
        if(pFindI == m_srcData.end() || pFindJ == m_srcData.end())
          continue;

        dy = m_dY_Less + m_dY_e * ((pFindI->second - yMinT)/m_dY_r) ;
        NodePos[END_I].Y = dy ;
        dotPt1 = ChangeGDIPlusCoord(NodePos[END_I]);
        graphics->FillEllipse(&POT_Brush,dotPt1.X-2,dotPt1.Y-2,4,4);

        if(k != xMin)
          graphics->DrawLine(&POT_Pen,lastPoint,dotPt1);

        dy = m_dY_Less + m_dY_e * ((pFindJ->second - yMinT)/m_dY_r) ;
        NodePos[END_J].Y = dy ;
        dotPt2 = ChangeGDIPlusCoord(NodePos[END_J]);
        graphics->FillEllipse(&POT_Brush,dotPt2.X-3,dotPt2.Y-2,4,4);

        graphics->DrawLine(&POT_Pen,dotPt1,dotPt2);

        lastPoint = dotPt2 ;

      }

      //绘制图例
      Gdiplus::Point tl_Pt ;
      tl_Pt.X = m_uDrawZoom_Width-m_dX_Less + 30 ;
      tl_Pt.Y = m_dMax_Y +10;

      if(m_TLColors.size() < m_strTL.size())
      {
        return;
      }
      findKey.iElemID = 0 ;
      findKey.ePos = END_I ;
      findKey.eXn = (MAXMINTOPBOT)i;
      findKey.eValType = (VALUETYPE)j;
      map<CGraphKey,CString>::iterator pFindTl =  m_strTL.find(findKey);
      if(pFindTl == m_strTL.end())continue;
      CString strTl = pFindTl->second;
        

        PointF sPt(tl_Pt.X,tl_Pt.Y + 20*(i-1+2*j));
        PointF ePt(tl_Pt.X + 20 ,tl_Pt.Y + 20*(i-1+2*j));
        //ePt = ChangeGDIPlusCoord(ePt);

        graphics->DrawLine(&POT_Pen,sPt,ePt);

        ePt.X += 5 ;
        ePt.Y -= 6 ;

        graphics->DrawString(strTl,   -1,   &font,  ePt  ,  &POT_Brush); 
      
    }
  }



}

Gdiplus::Point CCreatePic::ChangeGDIPlusCoord(const Gdiplus::Point& pt)
{
  Point Newpt(pt);
  Newpt.Y = m_uDrawZoom_Height - Newpt.Y;
  return Newpt;
}

Gdiplus::PointF CCreatePic::ChangeGDIPlusCoord(const Gdiplus::PointF& pt)
{
  PointF Newpt(pt);
  Newpt.Y = m_uDrawZoom_Height - Newpt.Y;
  return Newpt;
}

//////////////////////////////////////////////////////////////////////////
/*
map<UINT,vector<pair<double,double>>>& srcData 源数据 map<单元号,vector<pair<i端数据,J端数据>>>
strX_Name  X坐标轴输出名称
strY_Name  Y坐标轴输出名称
strTL      图形中的图形说明，类似于Excel图表的说明性文字,vector<pair<i端数据,J端数据>>中有几项应输出几项

return 返回一个图片路径 默认存储在C:\\下
*/
//////////////////////////////////////////////////////////////////////////
CString CCreatePic::CreatGraph(const map<CGraphKey,double>& srcData,const CString& IstrX_Name,const CString& IstrY_Name,const map<CGraphKey,CString>& strTLS,const CString& picName)
{
  CString PathOut = GetSystemPath() + L"\\" ;
  PathOut  += picName + L".jpg" ;
  m_srcData = srcData ;
  m_strTL = strTLS ;
  if(m_srcData.empty() || m_strTL.empty())
    return PathOut;

  m_strOutPath = PathOut ;

  str_X_Name = IstrX_Name ;
  str_Y_Name = IstrY_Name;
  
  Create();

  return PathOut;
  
}


_MITC_BASIC_END