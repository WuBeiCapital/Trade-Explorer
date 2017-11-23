#pragma once

//class CEttArc;
//class CEttPolyline;
//class CPlan;
//class CRect2;
//class GridCtrl;
  
//#include "Excel9.h"
#include "Excel8.h"
#include "XlControlDefine.h"

//class CMsgDlg;

_MITC_TMP_BEGIN

//#include "HeaderPre.h"
class  CXLCell : public CObject
{
public:
	CXLCell();
	CXLCell(long row, long col, CString str)
	{
		m_nRow = row;
		m_nCol = col;
		m_strData = str;
	};

	virtual ~CXLCell();

	long GetRow() const { return m_nRow; }
	void SetRow(long nRow) { m_nRow = nRow; }
	long GetCol() const { return m_nCol; }
	void SetCol(long nCol) { m_nCol =  nCol; }
	CString GetData() const { return m_strData; }
	void SetData(CString strData) { m_strData = strData; }

	void SetXL(long row, long col, CString str){
		m_nRow = row;
		m_nCol = col;
		m_strData = str;
	}
  
protected:
	long m_nRow;
	long m_nCol;
	CString m_strData;
};
  
class CGridCtrl;

class  CXLControl : public CObject
{
public:
	CXLControl(BOOL bVisible=FALSE);
	virtual ~CXLControl();
    
// Data Member
public:
	///////////////////////////////////////////////////////////////
	// * Excel * //
	///////////////////////////////////////////////////////////////
	_Application	m_App;
	Workbooks		m_Books;
	Worksheets		m_Sheets;
	_Workbook		m_Book;
	_Worksheet		m_Sheet;

	Range			m_Range;
	Shapes			m_Shapes;
	XLFont			m_Font;
	Borders			m_Borders;
	Interior		m_Interior;
	Window			m_Window;

	enum { XL_DOWN = 5, XL_UP, XL_LEFT, XL_TOP, XL_BOTTOM, XL_RIGHT, XL_INSIDE_VER, XL_INSIDE_HOR };	// LineÀÇ ¹æÇâ

protected:
	///////////////////////////////////////////////////////////////
	// * ±âÅ¸¸â¹öº¯¼ö * //
	///////////////////////////////////////////////////////////////
	COleSafeArray*	m_pSafeArr;
	long	m_nRow;
	long	m_nCol;
	
	COleVariant		m_covOptional;

	CString			m_strOpenName;
	CString			m_strSaveName;
	long			m_nSheet;		// Sheet Index : default 1
	BOOL			m_bVisible;		// Excel Visible or non-Visible
	BOOL			m_bInstance;	// Ole Init or Non Init;

	BOOL			m_bDomUseColor;
	CDWordArray		m_DeleteLine;
	float			m_fDisRow,m_fDisCol;

	long m_xlLineDirection;
	long m_xlLineStyle;
	long m_xlLineWeight;
	long m_xlLineColor;
	long m_xlFontColor;
	long m_xlCellColor;

//=============================================================
// Implementation
//=============================================================

///////////////////////////////////////////////////////////////
// * Program Start and End * //
///////////////////////////////////////////////////////////////
protected:
	CArray<int, int> m_HandleArr;
	void CheckExcelPre();

	BOOL CreateOleInstance();
	BOOL CreateExcelInstance();

public:
	BOOL NewXL();
	void GetXLApp(_Application	App, CString strFileName = _T(""));
	void GetXLApp(LPDISPATCH lpDisp);

	BOOL OpenXL(CString strFileName, CString sPassWord = _T(""));

	void Save();
	void SaveAs(CString strSaveName);
//	void SaveAs(CString strSheetName, CGridCtrl *pGrid, CMsgDlg *pDlg);


	void QuitXL();	// Open¸¸ ÇÑ °æ¿ì(¼öÁ¤ ¾øÀ½) 2001.0630
	void DeleteDummyExcelProcess();
	static BOOL CALLBACK EnumDeleteProcess(HWND hwnd, LPARAM data);
	void TerminateExcel();		//°­Á¦Á¾·á...ÂÁ...

	//void CloseXL();
	//void CloseXL(CString strSaveName=_T(""));

	void SetXLVisible(BOOL bVisible) { m_bVisible = bVisible; }
	BOOL GetXLVisible() const { return m_bVisible; }

///////////////////////////////////////////////////////////////
// * Sheet * //
///////////////////////////////////////////////////////////////
protected:
	BOOL IsUseSheet(const CStringArray& SA,const CString& sSheetName) const;
	long GetSheetNo() const { return m_nSheet; }

public:
	void AddSheet();
	void SetSheetNumbers(LONG_PTR nNum);
	long GetSheetsu();
	CString GetSheetName();
	void SheetMoveCopy(BOOL bBefore,BOOL bCopy,const CString& csNewSheetName=_T(""));
	void SheetMoveCopy(LPCTSTR lpstr,BOOL bCopy,const CString& csNewSheetName=_T(""));

	void SetActiveSheet(long nSheet);
	BOOL SetActiveSheet(const CString& sheetName);
	
	void DeleteSheetNotUsed(const CStringArray& SA);
	void SheetDelete();
	// ADD CMS
	BOOL SheetDelete(const CString& sSheetName);

	void SetSheetName(LPCTSTR lpname, long nZoomSize = 100);
	void SetSheetVisible(const CString& sSheetName, BOOL bVisible);
	void SetSheetVisible(BOOL bVisible);

	BOOL GetSheetVisible();


///////////////////////////////////////////////////////////////
// * Cell Data Setting * //
///////////////////////////////////////////////////////////////	
protected:
	void DeleteObjects(LPCTSTR lp1,LPCTSTR lp2);
	
	void SetValueXL(LPCTSTR x, LPCTSTR x1, LPCTSTR newValue, BOOL bFormula=FALSE);// Data Setting : x ¿Í y »çÀÌ¿¡ °°Àº Data¸¦ ³ÖÀ½

public:
	void CreateMatrix(long nRow, long nCol, long nType=1);
	void SetMatrix(long nRow, long nCol, CString szNewValue);	
	void SetMatrix(long nRow, long nCol, double NewValue);
	void SetMatrix(long nRow, long nCol, long lNewValue);
	void WriteMatrix(long nRow1, long nCol1, long nRow2, long nCol2);
	void WriteMatrix(long nRow, long nCol);
	void WriteMatrix(CString sttCell, CString endCell);
	void DeleteMatrix();

	void WriteMatrix(long nRow, long nCol, CTypedPtrArray<CObArray, CXLCell*>* pArray, long nSttRow, long nSttCol);
	
	// CMS void WriteMatrix(CMatrixStr &XLMat, long nSttRow, long nSttCol);

	CString GetXL(LPCTSTR lp1);
	CString GetXL(long r,long c);
	CString GetCellStr(long nRow,long nCol) const;
	double GetXLValue(long r, long c);

	void Clear(LPCTSTR lpstr1, LPCTSTR lpstr2);
	void Clear(long nRow1, long nCol1, long nRow2, long nCol2);

	void ClearContentsOnly(LPCTSTR lpstr1, LPCTSTR lpstr2);
	void ClearContentsOnly(long nRow1, long nCol1, long nRow2, long nCol2);

	void SetXL(long r,long c,LPCTSTR newValue,BOOL bFormula=FALSE);
	void SetXL(LPCTSTR x,LPCTSTR newValue,BOOL bFormula=FALSE);
	void SetXL(long r,long c, double dNewValue, BOOL bFormula=FALSE);
	void SetXL(long r,long c, int nNewValue, BOOL bFormula=FALSE);
	void SetXLDouble(long r,long c, double dNewValue, long nFloatPos = 3);
	void SetXLStr(long r,long c, CString sNewValue, CString strFormat = _T("#,##0.000"), long TA_ALIGN = TA_CENTER, BOOL bOneChar = TRUE);
	void SetXLStr(long rStt,long cStt, long rEnd,long cEnd, CString sNewValue, CString strFormat = _T("#,##0.000"),
		          long TA_ALIGN = TA_CENTER, BOOL bOneChar = TRUE, BOOL bBorders = TRUE);


///////////////////////////////////////////////////////////////
// * Cell Format °ü·Ã : Data and Cell Box Line * //
///////////////////////////////////////////////////////////////	
public:
	void InsertRowLine(long nInsertLine,long nQtyLine = 1);
	void SetNumberFormat(long nRow, long nCol, LPCTSTR strFormat);
	void SetNumberFormat(long nRow1, long nCol1, long nRow2, long nCol2, LPCTSTR strFormat);
	void SetNumberFormat(LPCTSTR strCell, LPCTSTR strFormat);
	void SetNumberFormat(LPCTSTR strCell1, LPCTSTR strCell2, LPCTSTR strFormat);

	void SetTextToColumns(CString strCell);
	void SetTextToColumns(CString strCell1, CString strCell2);
	void SetTextToColumns(long nCol, long nRow1, long nRow2);

	void SetMergeCell(long nRow1, long nCol1, long nRow2, long nCol2);
	void SetMergeCell(LPCTSTR lpstr1,LPCTSTR lpstr2);

	VARIANT GetLineStyle(LPCTSTR lpstr1, LPCTSTR lpstr2);
	void SetLineStyle(LPCTSTR lpstr1,LPCTSTR lpstr2,long newValue);
	
	void SetHoriAlign(long nRow, long nCol,long nRow2, long nCol2, long TA_ALIGN);
	void SetHoriAlign(LPCTSTR lpstr,LPCTSTR lpstr2, long TA_ALIGN = 7);
	
	void SetVerAlign(CString strCell1, long nAlign/*=1*/);
	void SetVerAlign(CString strCell1, CString strCell2, long nAlign/*=1*/);
	void SetVerAlign(long nRow1, long nCol1, long nAlign/*=1*/);
	void SetVerAlign(long nRow1, long nCol1, long nRow2, long nCol2, long nAlign/*=1*/);

	void DeleteColSell(long nSttRow,long nSttCol,long nEndRow,long nEndCol);
	void DeleteColSell(LPCTSTR lp1,LPCTSTR lp2);
	
	void DeleteRowLineDirect(long nSttRow, long nEndRow);
	void DeleteRowLine(long nSttRow,long nEndRow);
	
	void InsertCopyRowLine(long nSttRow,long nEndRow,long nDesRow);
	
	void Copy(long nSttRow,long nEndRow,long nDestinationRow);
	void DeleteRowLineEnd();

	void CopyRange(long nSourceSttRow, long nSourceSttCol, long nSourceEndRow, long nSourceEndCol, CString sTargetSheet,   long nTargetSttRow, long nTargetSttCol);

	void SetBorders(CString strCell, long nStyle=1);
	void SetBorders(CString strCell1, CString strCell2, long nStyle=1);
	void SetBorders(long nRow, long nCol, long nStyle = 1);
	void SetBorders(long nRow1, long nCol1, long nRow2, long nCol2, long nStyle = 1);

	void TextBoxValue(long nRow, long nCol, CString strValue, long nStyle = 1);
	void TextBoxValue(CString strCell, CString strValue, long nStyle=1);

	void SetFonts(CString strCell1, CString strCell2, short nSize=10, CString strFont=_T("ËÎÌå"), short nColor=1, long bBold=TRUE);
	void SetFonts(CString strCell, short nSize=10, CString strFont=_T("ËÎÌå"), short nColor=1, long bBold=TRUE);
	void SetFonts(long nRow1, long nCol1, long nRow2, long nCol2, short nSize=10, CString strFont=_T("ËÎÌå"), short nColor=1, long bBold=TRUE);
	void SetFonts(long nRow, long nCol, short nSize=10, CString strFont=_T("ËÎÌå"), short nColor=1, long bBold=TRUE);
	void SetFontCharacters(long nRow, long nCol, short nstart, short length, BOOL bSuperscript);
	void SetFontCharacters(long nRow, long nCol, CString str, BOOL bOneChar = TRUE);

	void CellLine(long nRow, long nCol, long nEdge=XL_BOTTOM, long nStyle=1, long nWeight=2);
	void CellLine(long nRow, long nCol, long nRow1, long nCol1, long nEdge=XL_BOTTOM, long nStyle=1, long nWeight=2);
	void CellLine(CString strCell, long nEdge=XL_BOTTOM, long nStyle=1, long nWeight=2);
	void CellLine(CString strCell1, CString strCell2, long nEdge=XL_BOTTOM, long nStyle=1, long nWeight=2);

	void CellOutLine(CString strCell1, CString strCell2, long nStyle=1, long nColorIndex = 1, long nWeight = 2);
	void CellOutLine(CString strCell, long nStyle=1, long nColorIndex = 1, long nWeight = 2);
	void CellOutLine(long nRow, long nCol, long nRow1, long nCol1, long nStyle=1, long nColorIndex = 1, long nWeight = 2);
	void CellOutLine(long nRow, long nCol, long nStyle=1, long nColorIndex = 1, long nWeight = 2);

	void SetCellWidth(long nCol1, long nCol2, long nLength);
	void SetCellWidth(long nCol, long nLength);
	void SetCellWidth(LPCTSTR x, LPCTSTR x1, long nLength);
	void SetCellWidth(LPCTSTR x, long nLength);

	long GetCellWidth(long nCol1, long nCol2=-1);

	void SetCellHeight(long nRow, double Height);
	void SetCellHeight(long nRow1, long nRow2, double Height);
	double GetCellHeight(long nRow1, long nRow2=-1);

	long SetLineDirection(long xlLineDirection)
	{ 
		long nOld = m_xlLineDirection;
		m_xlLineDirection = xlLineDirection;
		return nOld;
	}
	long SetLineStyle(long xlLineStyle)
	{ 
		long nOld = m_xlLineStyle;
		m_xlLineStyle = xlLineStyle;
		return nOld;
	}
	long SetLineWeight(long xlLineWeight)
	{ 
		long nOld = m_xlLineWeight;		
		m_xlLineWeight = xlLineWeight;
		return nOld;
	}
	long SetLineColor(long xlLineColor)
	{
		long nOld = m_xlLineColor;
		m_xlLineColor = xlLineColor;
		return nOld;
	}
	long SetFontColor(long xlFontColor) 
	{ 
		long nOld = m_xlFontColor;
		m_xlFontColor = xlFontColor;
		return nOld;
	}
	long SetCellColor(long xlCellColor)
	{ 
		long nOld = m_xlCellColor;
		m_xlCellColor = xlCellColor;
		return nOld;
	}

	long GetLineDirection() const { return m_xlLineDirection; }
	long GetLineStyle()  const { return m_xlLineStyle; }
	long GetLineWeight() const { return m_xlLineWeight; }
	long GetLineColor()  const { return m_xlLineColor; }

	long GetFontColor()  const { return m_xlFontColor; }
	long GetCellColor()  const { return m_xlCellColor; }

	// 2001.02.14 ktp
	void SetCellColor(CString stt, CString end, long nColNum);
	void SetCellColor(long nRow1, long nCol1, long nRow2, long nCol2, long nColNum);

///////////////////////////////////////////////////////////////
// * ¼¿°ú »ó°ü¾ø´Â ¼± ±×¸®±â * //
///////////////////////////////////////////////////////////////	
protected:

public:
	void DrawRootLine(long nRowStt, long nColStt, long nRowEnd, long nColEnd);
//	CRect2 GetXlCoordinatesXY(long nRow, long nCol);
	void DrawLine(double Sx, double Sy, double Ex, double Ey, long nWeight=1, long nColor=8, long nStyle =1);

	void DrawLine(double Sx, double Sy, double Ex, long nWeight=1, long nColor=8/*long nStyle =1*/);
	void DrawTextBox(long nSttRow, long nSttCol, long nEndRow, long nEndCol, CString strText);
	void DrawTextBox(long nSttRow, long nSttCol, long nEndRow, long nEndCol, long nHorz=2, short nSize=10,
						CString strFont=_T("ËÎÌå"), CString strText = _T(""));


///////////////////////////////////////////////////////////////
// * Domyun * //
///////////////////////////////////////////////////////////////	
/*
protected:	
	void AddDomArcSub(CEttArc* pArc,CPlan *pDomP);
	void AddDomLine(CPlan *pDomP);
	void AddDomPolyLine(CPlan *pDomP);
	void AddDomArc(CPlan *pDomP);
	void AddDomText(CPlan *pDomP,double RateTextHeight);
	void AddDomSolid(CPlan *pDomP);
	CRect2 GetMTextBorder(const CPoint2& xy,const CString& sMText,TEXTSTYLE* pStyle,double THeight) const;

public:
// CMS	void AddDomImage(CBitmap *pBitmap, double dLeft, double dTop, double dWidth = 0, double dHeight = 0);
// CMS	void AddDomImage(CPlan *pDomP, CString szFileName, double dLeft, double dTop, double dWidth, double dHeight );
	void AddDomyun(CPlan *pDomP,double x,double y,double Scale =0.01,double RateTextHeight = 2.0);
	void AddDomyunRC(CPlan *pDomP,long nRow,long nCol,double Scale =0.01,double RateTextHeight=2.0);
	void SetDomDisRC(double xDis,double yDis);
	void SetDomUseColor(BOOL bDomUseColor) { m_bDomUseColor = bDomUseColor; }
	*/

///////////////////////////////////////////////////////////////
// * ±â Å¸ * //
///////////////////////////////////////////////////////////////	
protected:
	void SetThreadUse(LPDISPATCH lpdisp);
	BYTE* GetBitInfo(BYTE* pByte32,long nSou) const;
public:
	void ExcelExecuteFile(CString strFileName);
	void SetPrintTitleRows(long nRow1, long nRow2, long nSheetNumber=-1);
	void SetPrintTitleCols(long nCol1, long nCol2, long nSheetNumber=-1);
	void SetPrintTitleCols(CString strCol1, CString strCol2, long nSheetNumber=-1);

	CString GetVertStrCol(long nCol) const;

	void SetFreezePanes(long nRow, long nCol, long nSheetNumber=1, BOOL bFreeze=TRUE);

	void SetHeader(CString strHeader, long nSheetNumber=1, long nCenter=1);
	void SetFooter(CString strFooter, long nSheetNumber=1, long nCenter=1);

	void SetPageBreak(long nRow, long nCol);
	void SetPageBreak(CString strCell);
	void SetPrintMargin(double dLeft, double dRight, double dTop, double dBottom);
	void SetPrintCenterHorizon(BOOL bCenterHor = FALSE);
	void SetPrintCenterVertical(BOOL bCenterVer = FALSE);

	void SetPringZoom(long nSize);

	CString GetPrintArea();
	void SetPrintArea(LPCTSTR lpstr1, LPCTSTR lpstr2);
	void SetOrientation(BOOL bLandScape);


	/* CMS
	void CopyPicture(CXLControl *xl, CString sPicName, CString Des);
	void CopyPicture(CXLControl &xl, CString sPicName, CString Des);
	void CopyPicture(CString sXLPicPath, long nRow1, long nCol1, long nRow2, long nCol2, long nRowDes, long nColDes, long nType); 
	void CopyPicture(CString sXLPicPath, CString sPicName, long nRowDes, long nColDes); 
	void CopyPicture(CString sXLPicPath, CString sPicName, CString Des);
	void CopyPicture(CBitmap *pBitmap, double dLeft, double dTop, double dWidth, double dHeight);
	*/
	void CopyPicture(CString sPictureSheetName, long nRowStt, long nColStt, long nRowEnd, long nColEnd, long nRowTarget, long nColTarget);
	void InsertPictureRowCol(CString sPath, long nRowStt, long nColStt, long nRowEnd=0, long nColEnd=0, BOOL bLockAspectRatio = FALSE);
	void InsertPicture(CString sPath, double dLeft, double dTop, double dWidth=0, double dHeight=0, BOOL bLockAspectRatio = FALSE);

	void SetDisplayGridLine(BOOL bGridLine);
	void SetUserControl(BOOL bUserCtrl = TRUE);
	void SetViewZoom(long nSize);


	void PrintOutDlg();
	void PageSetupDlg();
	void PrintPewView();

//	void SetCellHidden(long nRow, long nCol, long nSheetNumber/*=1*/,  BOOL bHidden=TRUE);

};
//#include "HeaderPost.h"

_MITC_TMP_END
