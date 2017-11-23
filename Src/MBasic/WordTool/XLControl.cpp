// XLControl.cpp: implementation of the CXLControl class.
//
//////////////////////////////////////////////////////////////////////
  
#include "stdafx.h"

//#include "../PlanBase/Point2.h"
//#include "../PlanBase/TypedArray.h"
//#include "../GridCtrl/GridCtrl.h"

#include "XLControl.h"
//#include "MsgDlg.h"

#include <afxdisp.h>
#include <tlhelp32.h>

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif
 
_MITC_BASIC_BEGIN
//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CXLCell::CXLCell() 
{
}

CXLCell::~CXLCell()
{
}

 
CXLControl::CXLControl(BOOL bVisible/*FALSE*/)
{
	m_bVisible = bVisible;
	m_bInstance = FALSE;

	m_covOptional = COleVariant((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	m_strOpenName = _T("");
	m_nSheet = 1;

	m_bDomUseColor = FALSE;
	m_fDisRow = 19.35466667f;
	m_fDisCol = 10.0f;
}

CXLControl::~CXLControl()
{

}

///////////////////////////////////////////////////////////////
// * Program Start and End * //
///////////////////////////////////////////////////////////////
BOOL CXLControl::CreateOleInstance()
{
	if(m_bInstance == TRUE)  // 이미 초기화 되어 있으면 빠져나간다.
		return TRUE;

	_AFX_THREAD_STATE* pState = AfxGetThreadState();
	if( pState->m_bNeedTerm )    // calling it twice?
		return TRUE;

	if(!AfxOleInit())
	{
		AfxMessageBox(_T("Could not initialize COM dll"));
		return FALSE;
	}
	m_bInstance = TRUE;
	return TRUE;
}

BOOL CXLControl::CreateExcelInstance()
{
	// Excel 초기화 모드
	if(!m_App.CreateDispatch(_T("Excel.Application")))
	{
		AfxMessageBox(_T("Couldn't start Excel."));
		return FALSE;
	}
	
	m_App.SetDisplayAlerts(FALSE); // Excel 종료시 저장 여부를 물어보지 않는다.
	m_bInstance = TRUE;
	return TRUE;
}

// new file
BOOL CXLControl::NewXL()
{
	// ToOpen Exsiting workBook and Get m_Sheet1
	try
	{
		CheckExcelPre();

		if( CreateOleInstance() == FALSE )  return FALSE;
		if( CreateExcelInstance()  == FALSE ) return FALSE;		

		/*
		// Register the Analysis ToolPak.		
		CString	strAppPath;
		strAppPath.Format("%s\\Analysis\\Analys32.xll", m_App.GetLibraryPath());
		
		if( !m_App.RegisterXLL(strAppPath))
			AfxMessageBox("Didn't register the Analys32.xll");
		*/
		
		// Get the Work Books collection.
		LPDISPATCH		lpDisp=NULL;
		lpDisp = m_App.GetWorkbooks();	// Get the IDispatch pointer;
		ASSERT(lpDisp);
		m_Books.AttachDispatch(lpDisp);
		
		lpDisp = m_Books.Add( m_covOptional );
		ASSERT(lpDisp);
		m_Book.AttachDispatch(lpDisp);

		lpDisp = m_Book.GetSheets();
		ASSERT(lpDisp);
		m_Sheets.AttachDispatch(lpDisp);

		// Get Sheet1
		lpDisp = m_Sheets.GetItem( COleVariant((short)(m_nSheet)) );	// parameter --> m_Sheet index;
		ASSERT(lpDisp);
		m_Sheet.AttachDispatch(lpDisp);

		m_App.SetVisible(m_bVisible);
		m_App.SetUserControl(TRUE);
		m_App.SetDisplayAlerts(FALSE); // Excel 종료시 저장 여부를 물어보지 않는다.
	}
	catch(COleException *e)
	{
		QuitXL();

		char buf[1024];
		sprintf(buf, _T("COleException. SCODE: %081x."), (long)e->m_sc);
		::MessageBox(NULL, buf, _T("COleException"), MB_SETFOREGROUND | MB_OK);


	}
	catch(COleDispatchException* e)
	{
		QuitXL();

		char buf[1024];
		sprintf(buf, _T("COleDispatchException. SCODE: %081x, Description: \"%s\"."), (long)e->m_wCode,
			(LPSTR)e->m_strDescription.GetBuffer(512));
		::MessageBox(NULL, buf, _T("COlePispatchException"), MB_SETFOREGROUND | MB_OK);
	}
	catch(...)
	{
		QuitXL();

		::MessageBox(NULL, _T("General Exception caught."), _T("Catch-All"), MB_SETFOREGROUND | MB_OK);
	}

	return TRUE;
}

// 기존의 파일을 열어서 사용
BOOL CXLControl::OpenXL(CString strOpenName, CString sPassWord)
{
	// ToOpen Exsiting workBook and Get m_Sheet1
	try
	{
		CheckExcelPre();

		if( CreateOleInstance() == FALSE )  return FALSE;
		if( CreateExcelInstance()  == FALSE ) return FALSE;		
		m_App.SetVisible(m_bVisible);
//		m_App.SetUserControl(TRUE);

		LPDISPATCH		lpDisp=NULL;
		lpDisp = m_App.GetWorkbooks();	// Get the IDispatch pointer;
		ASSERT(lpDisp);
		m_Books.AttachDispatch(lpDisp);

		// 기존의 File을 열어서 실행 할 경우
		m_strOpenName = strOpenName;
		COleVariant sPass = COleVariant(sPassWord);
		lpDisp = m_Books.Open(m_strOpenName, m_covOptional, m_covOptional, m_covOptional,
						sPass, m_covOptional, m_covOptional, m_covOptional, m_covOptional,
						m_covOptional, m_covOptional, m_covOptional, m_covOptional);
		m_Book.AttachDispatch(lpDisp);

		lpDisp = m_Book.GetSheets();
		ASSERT(lpDisp);
		m_Sheets.AttachDispatch(lpDisp);

		// Get Sheet1
		lpDisp = m_Sheets.GetItem( COleVariant((short)(m_nSheet)) );	// parameter --> m_Sheet index;
		ASSERT(lpDisp);
		m_Sheet.AttachDispatch(lpDisp);
		m_App.SetDisplayAlerts(FALSE); // Excel 종료시 저장 여부를 물어보지 않는다.
	}
	catch(COleException *e)
	{
		QuitXL();

		char buf[1024];
		sprintf(buf, _T("COleException. SCODE: %081x."), (long)e->m_sc);
		::MessageBox(NULL, buf, _T("COleException"), MB_SETFOREGROUND | MB_OK);
	}
	catch(COleDispatchException* e)
	{
		QuitXL();

		char buf[1024];
		sprintf(buf, _T("COleDispatchException. SCODE: %081x, Description: \"%s\"."), (long)e->m_wCode,
			(LPSTR)e->m_strDescription.GetBuffer(512));
		::MessageBox(NULL, buf, _T("COlePispatchException"), MB_SETFOREGROUND | MB_OK);
	}
	catch(...)
	{
		QuitXL();

		::MessageBox(NULL, _T("General Exception caught."), _T("Catch-All"), MB_SETFOREGROUND | MB_OK);
	}

	return TRUE;
}

void CXLControl::QuitXL()	// Open만 한 경우(수정 없음) 2001.0630
{
	//if(m_App) m_App.Quit();

//	m_Sheet.m_bAutoRelease = TRUE;
//	m_Sheet.ReleaseDispatch();
//	m_Sheet.DetachDispatch();
//
//	m_Sheets.m_bAutoRelease = TRUE;
//	m_Sheets.ReleaseDispatch();
//	m_Sheets.DetachDispatch();
//	
//	m_Book.m_bAutoRelease = TRUE;
//	m_Book.Close(COleVariant((short)FALSE), m_covOptional, m_covOptional);
//    m_Book.ReleaseDispatch();
//	m_Book.DetachDispatch();
//	
//	m_Books.m_bAutoRelease = TRUE;
//	m_Books.ReleaseDispatch();
//	m_Books.DetachDispatch();
//
//	m_Range.m_bAutoRelease = TRUE;
//	m_Range.ReleaseDispatch();
//	m_Range.DetachDispatch();
//
//	m_Shapes.m_bAutoRelease = TRUE;
//	m_Shapes.ReleaseDispatch();
//	m_Shapes.DetachDispatch();
//
//	m_Font.m_bAutoRelease = TRUE;
//	m_Font.ReleaseDispatch();
//	m_Font.DetachDispatch();
//
//	m_Borders.m_bAutoRelease = TRUE;
//	m_Borders.ReleaseDispatch();
//	m_Borders.DetachDispatch();
//
//	m_Interior.m_bAutoRelease = TRUE;
//	m_Interior.ReleaseDispatch();
//	m_Interior.DetachDispatch();
//
//	m_Window.m_bAutoRelease = TRUE;
//	m_Window.ReleaseDispatch();
//	m_Window.DetachDispatch();
//	
//	m_App.m_bAutoRelease = TRUE;
//	m_App.Quit();
//	m_App.ReleaseDispatch();
//	m_App.DetachDispatch();
////	DeleteDummyExcelProcess();

	m_App.Quit();

	TerminateExcel();

}
void CXLControl::DeleteDummyExcelProcess()
{
    EnumWindows((WNDENUMPROC)EnumDeleteProcess, NULL);
}
//---------------------------------------------------------------------------
BOOL CALLBACK CXLControl::EnumDeleteProcess(HWND hwnd, LPARAM data)
{
    int const max_len = 1024;
    char WndTitle[max_len];
    char WndClass[max_len];

	//char WndTitleChild[max_len];

    if (!hwnd)
        return false;

    GetClassName(hwnd, WndClass, max_len);
    GetWindowText(hwnd, WndTitle, max_len);

    CString fclass = WndClass;
    CString ftext =WndTitle;
    if(fclass == _T("XLMAIN"))
    {
//        int dot_pos = ftext.Pos(".");
  //      if(dot_pos > 0)
    //        ftext = ftext.SubString(1, dot_pos-1);
//		HWND hwndchild = GetWindow(hwnd, GW_CHILD);
//		GetWindowText(hwndchild, WndTitleChild, max_len);
//		CString ftextChild = WndTitleChild;

        if(ftext == _T("Microsoft Excel"))             // Dummy Process 일경우의 window title
        {
            DWORD PID;
            GetWindowThreadProcessId(hwnd, &PID);      //CreateProcess
            HANDLE process_handle = OpenProcess(PROCESS_ALL_ACCESS, false, PID);
            TerminateProcess(process_handle, 0);
            return true;
        }
    }
    return true;
}
// 현재의 이름으로 저장
void CXLControl::Save()
{
	m_Book.Save();
}

// 다름 이름으로 저장
void CXLControl::SaveAs(CString strSaveName)
{
	COleVariant  covTrue((short)TRUE);
	COleVariant  covFalse((short)FALSE);
//	m_Book.SaveCopyAs((COleVariant)strSaveName);
	m_Book.SaveAs(COleVariant(strSaveName), COleVariant((short)-4143),  //xlnormal=-4143  
                        COleVariant(""),COleVariant(""),covFalse,covTrue,
                         (long)0,covFalse, covFalse, covFalse, covFalse);
	m_Book.SetSaved(TRUE);	
}

void CXLControl::SaveAs(CString strSheetName, CGridCtrl *pGrid, CMsgDlg *pDlg)
{		
	long nRows = pGrid->GetRowCount();
	long nCols = pGrid->GetColumnCount();
	long nFixedRows = pGrid->GetFixedRowCount();
	long nFixedCols = pGrid->GetFixedColumnCount();

	long nProgress = 0;
	long TotalNum = nRows * nCols;
	if(pDlg) pDlg->m_Progress.SetRange(0, 100);
		
	SetVerAlign(0,0,nRows-1,nCols-1,2);

	SetFonts(0, 0, nRows-1, nCols-1, 10, _T("굴림체"), 1, FALSE);	// 스타일

	for(long row=0; row < nFixedRows; row++)
	{
		SetCellColor(row, 0, row, nCols-1, 35);
		SetFonts(row, 0, row, nCols-1, 10, _T("굴림체"), 1, TRUE);	// 스타일
	}
	if(nFixedRows > 0)	SetBorders(0, 0, nFixedRows-1, nCols-1);

	for(long col=0; col < nFixedCols; col++)
	{
		SetCellColor(0, col, nRows-1, col, 35);
		SetFonts(0, 0, nRows-1, col, 10, _T("굴림체"), 1, TRUE);	// 스타일
	}
	if(nFixedCols > 0)	SetBorders(0, 0, nRows-1, nFixedCols-1);

	for(row=0;row<nRows;row++)
	{
		SetCellHeight(row, pGrid->GetRowHeight(row) * 0.8);							
		for(col=0;col<nCols;col++)
		{
			if(row==0) SetCellWidth(col,pGrid->GetColumnWidth(col)/9);
			if(col < pGrid->GetFixedColumnCount())
				SetCellColor(row, col, row, col, 35);			

			UINT nFormat = pGrid->GetItemFormat(row,col);			
			long TA_ALIGN = TA_LEFT;
			if(nFormat & DT_RIGHT) TA_ALIGN = TA_RIGHT;
			else if(nFormat & DT_CENTER) TA_ALIGN = TA_CENTER;

			if(TA_ALIGN >= 0) 
			{
				CCellRange Range = pGrid->GetCell(row,col)->GetMergeRange();
				int nMinRow = Range.GetMinRow();
				int nMaxRow = Range.GetMaxRow();
				int nMinCol = Range.GetMinCol();
				int nMaxCol = Range.GetMaxCol();
				/*
				if(Range.GetMinRow() == -1 || Range.GetMaxRow() == -1 || Range.GetMinCol() == -1 || Range.GetMaxCol() == -1)
				{
				}
				else if(Range.GetMinRow() != Range.GetMaxRow() || Range.GetMinCol() != Range.GetMaxCol())
					SetMergeCell(Range.GetMinRow(),Range.GetMinCol(),Range.GetMaxRow(),Range.GetMaxCol());
					*/
				
				//SetHoriAlign(Range.GetMinRow(),Range.GetMinCol(),Range.GetMaxRow(),Range.GetMaxCol(),TA_ALIGN);
			}
			CString szText = pGrid->GetItemText(row,col);			
			if(!szText.IsEmpty())
				SetXL(row,col,szText[0]=='\n' ? szText.Mid(1) : szText);
			
			if(pDlg) 
			{
				if(pDlg->IsAbort()) return;				
				pDlg->m_Progress.SetPos((int)((double)++nProgress/TotalNum*100.0));
			}
		}
	}				
	SetSheetName(strSheetName);	
}


/*
// Open으로 시작하는 경우는 Close로 닫는다. => 저장하지 않고 닫기
void CXLControl::CloseXL()
{
	COleVariant  covTrue((short)TRUE);
	COleVariant  covFalse((short)FALSE);
	COleVariant	 covOptional((long)DISP_E_PARAMNOTFOUND , VT_ERROR);

	m_Book.Close(covFalse, COleVariant(m_strOpenName), covOptional);
	m_Books.Close();
}
*/
///////////////////////////////////////////////////////////////
// * Sheet * //
///////////////////////////////////////////////////////////////
void CXLControl::AddSheet()
{
	long nCount = m_Sheets.GetCount();

	VARIANT vNotPassed;

    V_VT(&vNotPassed) = VT_ERROR;
    V_ERROR(&vNotPassed) = DISP_E_PARAMNOTFOUND;


	m_Sheets.Add(vNotPassed,
				vNotPassed,	// 의미 확인 안됨
			COleVariant((short)(1)),	// 추가할 Sheet 수
			COleVariant((short)1));	// 1 Sheet, 2 Chart, 3 Macro

}

void CXLControl::SetSheetNumbers(LONG_PTR nNum)
{
  if (nNum < 1)
  {
    return;
  }

  long nCount = m_Sheets.GetCount();
  if (nCount <= nNum)
  {
    for (int nldx = nCount+1 ; nldx <= nNum ;  nldx++)
    {
      AddSheet();
    }
  }
  else
  {
    for (int nldx = nCount ; nldx > nNum ;  nldx--)
    {
      SetActiveSheet(nCount -1);
      SheetDelete();
    }
  }
}

BOOL CXLControl::SetActiveSheet(const CString &sheetName)
{
	LPDISPATCH	lpDisp = m_Sheets.GetItem( COleVariant((CString)(sheetName)) );	// parameter --> m_Sheet index;
	ASSERT(lpDisp);
	m_Sheet.AttachDispatch(lpDisp);
	m_nSheet = m_Sheet.GetIndex(); 
	m_Sheet.Activate();
	return TRUE;
//
//	// 현재의 deleteline정보 초기화
//	m_DeleteLine.RemoveAll();
//
//	long nShtsu = GetSheetsu();
//	for(long n = 1; n <= nShtsu; n++)
//	{
//		SetActiveSheet(n);
//		if(m_Sheet.GetName() == sheetName) 
//			return TRUE;
//	}
//
////	ASSERT(FALSE);
//	return FALSE;
}

void CXLControl::DeleteSheetNotUsed(const CStringArray &SA)
{

	for(long n = 1; n <= GetSheetsu(); n++)
	{
		SetActiveSheet(n);
		if( !IsUseSheet(SA,m_Sheet.GetName()) )
		{
			SheetDelete();
			n--;
		}
	}
}

BOOL CXLControl::IsUseSheet(const CStringArray &SA, const CString &sSheetName) const
{
	for(long n = 0; n < SA.GetSize(); n++)
		if(SA[n] == sSheetName) return TRUE;
	return FALSE;
}

void CXLControl::SetActiveSheet(long nSheet)
{
	LPDISPATCH	lpDisp = m_Sheets.GetItem( COleVariant((short)(nSheet)) );	// parameter --> m_Sheet index;
	ASSERT(lpDisp);
	m_Sheet.AttachDispatch(lpDisp);
	m_nSheet = nSheet; 
	m_Sheet.Activate();
}	

long CXLControl::GetSheetsu()
{
	return m_Sheets.GetCount();
}

CString CXLControl::GetSheetName()
{
	return m_Sheet.GetName();
}

void CXLControl::SetSheetName(LPCTSTR lpname, long nZoomSize)
{
	m_Sheet.SetName(lpname);
	if(nZoomSize != 100)	SetViewZoom(nZoomSize);
}

void CXLControl::SetSheetVisible(const CString& sSheetName, BOOL bVisible)
{
	if(SetActiveSheet(sSheetName))
		m_Sheet.SetVisible(bVisible);
}

BOOL CXLControl::GetSheetVisible()
{
	return m_Sheet.GetVisible();
}

void CXLControl::SetSheetVisible(BOOL bVisible)
{
	m_Sheet.SetVisible(bVisible);
}

void CXLControl::SheetDelete()
{
	m_Sheet.Delete();
}

BOOL CXLControl::SheetDelete(const CString& sSheetName)
{
//	CString strTemp = GetSheetName();
	BOOL bResult = SetActiveSheet(sSheetName);
	m_Sheet.Delete();
	//BOOL bResult = SetActiveSheet(sSheetName);
	//if (bResult)
	//	m_Sheet.Delete();
//	SetActiveSheet(strTemp);
	return bResult;

/*
	if(SetActiveSheet(sSheetName))
	{
		m_Sheet.Delete();
		return TRUE;
	}
	else
		return FALSE;
*/
}


void CXLControl::SheetMoveCopy(BOOL bBefore,BOOL bCopy,const CString& csNewSheetName)
{
	long nShtsu = GetSheetsu();
	LPDISPATCH	lpDisp;
	_Worksheet sheet;

	if(bBefore)	// 맨 앞로 move
	{
		lpDisp = m_Sheets.GetItem( COleVariant((short)(1)) );	// parameter --> m_Sheet index;

		ASSERT(lpDisp);
		sheet.AttachDispatch(lpDisp);		

		VARIANT v;
		v.vt = VT_DISPATCH;
		v.pdispVal = lpDisp;
		if(bCopy)	m_Sheet.Copy( v, m_covOptional );
		else		m_Sheet.Move( v, m_covOptional );
		sheet.ReleaseDispatch();
		sheet.DetachDispatch();

		lpDisp = m_Sheets.GetItem( COleVariant((short)(1)) );

	}
	else	// 맨 뒤로 move
	{
		lpDisp = m_Sheets.GetItem( COleVariant((short)(nShtsu)) );	// parameter --> m_Sheet index;

		ASSERT(lpDisp);
		sheet.AttachDispatch(lpDisp);		

		VARIANT v;
		v.vt = VT_DISPATCH;
		v.pdispVal = lpDisp;
		if(bCopy)	m_Sheet.Copy( m_covOptional, v );
		else		m_Sheet.Move( m_covOptional, v );
		sheet.ReleaseDispatch();
		sheet.DetachDispatch();

		lpDisp = m_Sheets.GetItem( COleVariant((short)(m_Sheets.GetCount())) );
		
	}


	// 새로운 시트 이름 
	if(csNewSheetName.GetLength() > 0)
	{
		sheet.AttachDispatch(lpDisp);
		sheet.SetName(csNewSheetName);
		sheet.ReleaseDispatch();
		sheet.DetachDispatch();
	}

}

// 현재 활성화된 시트를 lpstr이름 시트의 앞으로 이동
void CXLControl::SheetMoveCopy(LPCTSTR lpstr,BOOL bCopy,const CString& csNewSheetName)
{
	long nShtsu = GetSheetsu();
	LPDISPATCH	lpDisp;
	_Worksheet sheet;

	for(long n = 1; n <= nShtsu; n++)
	{
		lpDisp = m_Sheets.GetItem( COleVariant((short)(n)) );	// parameter --> m_Sheet index;

		ASSERT(lpDisp);
		sheet.AttachDispatch(lpDisp);		

		CString cs( sheet.GetName() );
		if( cs == lpstr) 	break;

		sheet.ReleaseDispatch();
		sheet.DetachDispatch();
	}
	if(GetSheetsu() < n) return;	// 찾을수 없음

	VARIANT v;
	v.vt = VT_DISPATCH;
	v.pdispVal = lpDisp;

	if(bCopy)	m_Sheet.Copy( v,m_covOptional );
	else		m_Sheet.Move( v,m_covOptional );
	sheet.ReleaseDispatch();
	sheet.DetachDispatch();


	// 새로운 시트 이름 
	if(csNewSheetName.GetLength() > 0)
	{
		lpDisp = m_Sheets.GetItem( COleVariant((short)(n)) );
		sheet.AttachDispatch(lpDisp);
		sheet.SetName(csNewSheetName);
		sheet.ReleaseDispatch();
		sheet.DetachDispatch();
	}
}

///////////////////////////////////////////////////////////////
// * Cell Data Setting * //
///////////////////////////////////////////////////////////////	

//	void Clear();
//	void ClearContents();
//	void ClearFormats();
//	void ClearNotes();
//	void ClearOutline();

void CXLControl::Clear(LPCTSTR lpstr1, LPCTSTR lpstr2)
{
	LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(lpstr1), COleVariant(lpstr2));
	ASSERT(lpDisp);
	m_Range.AttachDispatch(lpDisp);

	m_Range.Clear();

	m_Range.ReleaseDispatch();
	m_Range.DetachDispatch();

}

void CXLControl::Clear(long nRow1, long nCol1, long nRow2, long nCol2)
{
	CString lpstr1 = GetCellStr(nRow1, nCol1);
	CString lpstr2 = GetCellStr(nRow2, nCol2);

	Clear(lpstr1, lpstr2);
}

void CXLControl::ClearContentsOnly(LPCTSTR lpstr1, LPCTSTR lpstr2)
{
	LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(lpstr1), COleVariant(lpstr2));
	ASSERT(lpDisp);
	m_Range.AttachDispatch(lpDisp);

	m_Range.ClearContents();

	m_Range.ReleaseDispatch();
	m_Range.DetachDispatch();

}

void CXLControl::ClearContentsOnly(long nRow1, long nCol1, long nRow2, long nCol2)
{
	CString lpstr1 = GetCellStr(nRow1, nCol1);
	CString lpstr2 = GetCellStr(nRow2, nCol2);

	ClearContentsOnly(lpstr1, lpstr2);
}

void CXLControl::DeleteObjects(LPCTSTR lp1, LPCTSTR lp2)
{

//	LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(lpstr1), COleVariant(lpstr2));
	LPDISPATCH lpDisp;

	lpDisp = m_Sheet.OLEObjects( COleVariant( short(1) ) );
	OLEObjects OleObj(lpDisp);
	OleObj.Delete();
}

CString CXLControl::GetXL(long r, long c)
{
	ASSERT(r >= 0 && r <= 65535);
	ASSERT(c >= 0 && c <= 255);

	CString sCell = GetCellStr(r,c);
	return GetXL(sCell);

}
CString CXLControl::GetXL(LPCTSTR lp1)
{
	LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(lp1), COleVariant(lp1));
	
	ASSERT(lpDisp);
	m_Range.AttachDispatch(lpDisp);

	VARIANT v = m_Range.GetValue();
	m_Range.ReleaseDispatch();
	m_Range.DetachDispatch();

	
	if(v.vt == VT_BSTR)
		return (CString)v.bstrVal;
	else if(v.vt == VT_R8)
	{
		CString sFmt;
		sFmt.Format(_T("%g"),v.dblVal);
		return sFmt;
	}
	else if(v.vt == VT_DATE)
	{
		CString strDate;
		strDate.Format(_T("%g"),  v.date);
		return strDate;
	}
	
	// else if(v.vt == VT_EMPTY)

	return _T("");
}

// nRow, nCol은 0부터 시작하는 Index 이다.
// nCol = 0 => A 열
// nRow = 1 => 2 행
CString CXLControl::GetCellStr(long nRow,long nCol) const
{
	if(nRow < 0 || nCol < 0) return _T("");

	CString sRow(_T("")), sCol(_T(""));
	sRow.Format(_T("%ld"), nRow + 1);
	if(nCol < 26)
		sCol.Format(_T("%c"), 'A' + nCol);
	else if( nCol < 255 )	//else if(nCol / 26 < 26)
	{
		long h,f;
		h = nCol / 26 - 1;
		f = nCol % 26;
		sCol.Format(_T("%c%c"), 'A' + h, 'A' + f);
	}
	else sCol = _T("IV");

//	return sCol + sRow; // Release Error

	// 수정
	sCol += sRow;
	return sCol;
}

double CXLControl::GetXLValue(long r, long c)
{
	return atof(GetXL(r, c));
}


// 셀(r, c)에 값을 넣는다
// XL.SetXL( 0, 0, "우리나라" );
// XL.SetXL( 0, 0, "=SUM(A1:B2), TRUE );
void CXLControl::SetXL(long r,long c,LPCTSTR newValue,BOOL bFormula)
{
	ASSERT(r >= 0 && r <= 65535);
	ASSERT(c >= 0 && c <= 255);

	if(newValue == _T("0"))		return;
	CString sCell = GetCellStr(r,c);
	SetXL(sCell, newValue,bFormula);
}


void CXLControl::SetXL(LPCTSTR x,LPCTSTR newValue,BOOL bFormula)
{
	if(newValue == _T("0")	)	return;
	SetValueXL(x,x, newValue ,bFormula);
}

void CXLControl::SetXL(long r,long c, double dNewValue, BOOL bFormula)
{
	CString sBurf;
	sBurf.Format(_T("%.3f"), dNewValue);

	SetXL(r,c,sBurf,bFormula);
}

void CXLControl::SetXL(long r,long c, int nNewValue, BOOL bFormula)
{
	CString sBurf;
	sBurf.Format(_T("%d"), nNewValue);

	SetXL(r,c,sBurf,bFormula);
}
void CXLControl::SetXLStr(long r,long c, CString sNewValue, CString strFormat, long TA_ALIGN, BOOL bOneChar)
{
	SetHoriAlign(r,c,r,c,TA_ALIGN);
	SetNumberFormat(r,c,strFormat);
//	SetXL(r,c,sNewValue);
	SetFontCharacters(r,c,sNewValue, bOneChar);
}
void CXLControl::SetXLStr(long rStt,long cStt, long rEnd,long cEnd, CString sNewValue, CString strFormat, long TA_ALIGN, BOOL bOneChar, BOOL bBorders)
{
	if (rStt != rEnd)
	{
		SetMergeCell(rStt, cStt, rEnd, cEnd);
		SetXLStr(rStt, cStt, sNewValue, strFormat, TA_ALIGN, bOneChar);
	}
	else
	{
		SetHoriAlign(rStt,cStt,rEnd,cEnd,TA_ALIGN);
		SetNumberFormat(rStt,cStt,strFormat);
		//	SetXL(r,c,sNewValue);
		SetFontCharacters(rStt,cStt,sNewValue, bOneChar);
	}

	if(bBorders)
		SetBorders(rStt, cStt, rEnd, cEnd, 1);

}

void CXLControl::SetXLDouble(long r,long c, double dNewValue, long nFloatPos)
{
	CString sBurf;
	sBurf.Format(_T("%.*f"), nFloatPos, dNewValue);

	SetXL(r,c,sBurf);
}

void CXLControl::SetValueXL(LPCTSTR x, LPCTSTR x1, LPCTSTR newValue, BOOL bFormula/*=FALSE*/)
{
	try
	{
		LPDISPATCH		lpDisp;
		lpDisp = m_Sheet.GetRange(COleVariant(x), COleVariant(x1));

		ASSERT(lpDisp);

		m_Range.AttachDispatch(lpDisp);

		if(bFormula)
			m_Range.SetFormula(COleVariant(newValue));		// Ex : "=SUM(x,x1)"
		else
			m_Range.SetValue(COleVariant(newValue));			// Ex : "123"

		m_Range.ReleaseDispatch();
		m_Range.DetachDispatch();
	}
	catch(COleException *e)
	{
		char buf[1024];
		sprintf(buf, _T("COleException. SCODE: %081x."), (long)e->m_sc);
		::MessageBox(NULL, buf, _T("COleException"), MB_SETFOREGROUND | MB_OK);
	}
	catch(COleDispatchException* e)
	{
		char buf[1024];
		sprintf(buf, _T("COleDispatchException. SCODE: %081x, Description: \"%s\"."), (long)e->m_wCode,
			(LPSTR)e->m_strDescription.GetBuffer(512));
		::MessageBox(NULL, buf, _T("COlePispatchException"), MB_SETFOREGROUND | MB_OK);
	}
	catch(...)
	{
		::MessageBox(NULL, _T("General Exception caught."), _T("Catch-All"), MB_SETFOREGROUND | MB_OK);
	}	
}

// 일괄 Data Setting
// ① CreateMatrix(row, col, nType) : (row, col)인 nType형의 이차원 배열을 만든다 (nType 1: 문자 2: long, 3: double형이다)
// ② SetMatrix(row, col, newValue) : Data Setting newValue는 배열을 만들때 사용한 Data 형이다.
// ③ WriteMatrix("A1", "B4") : 배열에 있는 Data를 XL에 써 넣는다.
// ④ DeleteMatrix() : COleSafeArray라는 Data형을 사용하므로 지운다. 
// 5   SetTextToColumns();
void CXLControl::CreateMatrix(long nRow, long nCol, long nType/*=1*/)
{
	m_pSafeArr = new COleSafeArray;
	DWORD numElements[]={nRow, nCol};
	DWORD dwDims = 2;
	
	m_nRow = nRow;
	m_nCol = nCol;

	switch(nType)
	{
	case 1:
		m_pSafeArr->Create(VT_BSTR, dwDims, numElements);
		break;
	case 2:
		m_pSafeArr->Create(VT_I4, dwDims, numElements);	// long 32bit longeger
		break;
	case 3:
		m_pSafeArr->Create(VT_R8, dwDims, numElements);	// double 64bit floating-polong
		break;
	}
}

void CXLControl::SetMatrix(long nRow, long nCol, CString szNewValue)
{
	OLECHAR FAR* sz = szNewValue.AllocSysString();

	VARIANT v;
	long index[2];
	index[0] = nRow;
    index[1] = nCol;
	VariantInit(&v);
	
	// String
	v.vt = VT_BSTR;
	v.bstrVal = SysAllocString(sz);
	
	m_pSafeArr->PutElement(index, v.bstrVal);

	SysFreeString(v.bstrVal);
    VariantClear(&v);
}

void CXLControl::SetMatrix(long nRow, long nCol, double NewValue)
{
	long index[2];
	index[0] = nRow;
    index[1] = nCol;

	// double
	m_pSafeArr->PutElement(index, &NewValue);
}

void CXLControl::SetMatrix(long nRow, long nCol, long lNewValue)
{
	long index[2];
	index[0] = nRow;
    index[1] = nCol;

	// long
	m_pSafeArr->PutElement(index, &lNewValue);
}

void CXLControl::WriteMatrix(long nRow, long nCol)
{
	CString sttCell = GetCellStr(nRow, nCol);
	CString endCell = GetCellStr(nRow+m_nRow-1, nCol+m_nCol-1);
	WriteMatrix(sttCell, endCell);
}

void CXLControl::WriteMatrix(long nRow1, long nCol1, long nRow2, long nCol2)
{
	CString sttCell = GetCellStr(nRow1, nCol1);
	CString endCell = GetCellStr(nRow2, nCol2);
	WriteMatrix(sttCell, endCell);
}

void CXLControl::WriteMatrix(CString sttCell, CString endCell)
{
	m_Range = m_Sheet.GetRange(COleVariant(sttCell), COleVariant(endCell));
    m_Range.SetValue(COleVariant(m_pSafeArr));	
}

void CXLControl::DeleteMatrix()
{
	m_pSafeArr->Detach();

	_DELPTR(m_pSafeArr);
}

void CXLControl::WriteMatrix(long nRow, long nCol, CTypedPtrArray<CObArray, CXLCell*>* pArray, long nSttRow, long nSttCol)
{
	CreateMatrix(nRow, nCol, 1);

	CXLCell* pStg;
	long nSize = (int)pArray->GetSize();
	for( long n = 0; n < nSize; n++ )
	{
		pStg = pArray->GetAt(n);
		CString str = pStg->GetData();
		SetMatrix(pStg->GetRow(), pStg->GetCol(), str);
	}

	CString strSttCell = GetCellStr(nSttRow, nSttCol);
	CString strEndCell = GetCellStr(nSttRow+nRow-1, nSttCol+nCol-1);

	WriteMatrix(strSttCell, strEndCell);
	DeleteMatrix();

	//for( long c = 0; c < nCol; c++ )
	//	SetTextToColumns(nSttCol++, nSttRow, nSttRow+nRow-1);
}

/* CMS
void CXLControl::WriteMatrix(CMatrixStr &XLMat, long nSttRow, long nSttCol)
{
	long nRow = XLMat.GetRows();
	long nCol = XLMat.GetCols();

	CreateMatrix(nRow, nCol, 1);
	CString str;

	for(long r=0; r<nRow; r++)
	{
		for(long c=0; c< nCol; c++)
		{
			str = XLMat.GetData(r,c);
			SetMatrix(r, c, str);
		}
	}

	CString strSttCell = GetCellStr(nSttRow, nSttCol);
	CString strEndCell = GetCellStr(nSttRow+nRow-1, nSttCol+nCol-1);

	WriteMatrix(strSttCell, strEndCell);
	DeleteMatrix();

	//for( long c = 0; c < nCol; c++ )
	//	SetTextToColumns(nSttCol++, nSttRow, nSttRow+nRow-1);
}
*/

///////////////////////////////////////////////////////////////
// * Cell Format 관련 : Data and Cell Box Line * //
///////////////////////////////////////////////////////////////	
// row,col max range : 65536, IV(256)
// Array에 저장하지 않고 즉시 삭제
void CXLControl::DeleteRowLineDirect(long nSttRow, long nEndRow)
{
	CDWordArray DArDeleteLine;
	long nStt, nEnd;

	if(nSttRow > nEndRow) { long a=nSttRow; nSttRow=nEndRow;nEndRow=a;}
	nSttRow--;
	nEndRow--;
	for(long v = nSttRow; v <= nEndRow; v++)
	{
		BOOL bInsert = FALSE;
		for(long i = 0; i < DArDeleteLine.GetSize(); i++)
		{
			if(v < (long)DArDeleteLine[i])
			{
				DArDeleteLine.InsertAt(i,v);
				bInsert = TRUE;
				break;
			}
		}
		if(bInsert==FALSE)
			DArDeleteLine.Add(v);
	}

	for(long n = (int)DArDeleteLine.GetUpperBound(); n >= 0; n--)
	{
		nStt = nEnd = n;
		while(n > 0 &&	(long)DArDeleteLine[n] == 1 + (long)DArDeleteLine[n-1])
		{
			nStt--; n--;
		}

		long nSttRow = (long)DArDeleteLine[nStt];
		long nEndRow = (long)DArDeleteLine[nEnd];

		ASSERT( nSttRow >= 0 && nSttRow <= 65535);
		ASSERT( nEndRow >= 0 && nEndRow <= 65535);
		ASSERT( nSttRow <= nEndRow );


		CString cs1 = GetCellStr(nSttRow ,0);
		CString cs2 = GetCellStr(nEndRow,255);
		LPCTSTR lpx1 = (LPCTSTR)cs1;
		LPCTSTR lpx2 = (LPCTSTR)cs2;

		LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(lpx1), COleVariant(lpx2));
		ASSERT(lpDisp);
		m_Range.AttachDispatch(lpDisp);

		m_Range.Delete( COleVariant((long)1) ) ;

		m_Range.ReleaseDispatch();
		m_Range.DetachDispatch();
	}


}

void CXLControl::DeleteRowLine(long nSttRow, long nEndRow)
{
	if(nSttRow > nEndRow) { long a=nSttRow; nSttRow=nEndRow;nEndRow=a;}
	nSttRow--;
	nEndRow--;
	for(long v = nSttRow; v <= nEndRow; v++)
	{
		BOOL bInsert = FALSE;
		for(long i = 0; i < m_DeleteLine.GetSize(); i++)
		{
			if(v < (long)m_DeleteLine[i])
			{
				m_DeleteLine.InsertAt(i,v);
				bInsert = TRUE;
				break;
			}
		}
		if(bInsert==FALSE)
			m_DeleteLine.Add(v);
	}

}

void CXLControl::DeleteRowLineEnd()
{
	long nStt, nEnd;
	for(long n = (int)m_DeleteLine.GetUpperBound(); n >= 0; n--)
	{
		nStt = nEnd = n;
		while(n > 0 &&	(long)m_DeleteLine[n] == 1 + (long)m_DeleteLine[n-1])
		{
			nStt--; n--;
		}

		long nSttRow = (long)m_DeleteLine[nStt];
		long nEndRow = (long)m_DeleteLine[nEnd];

		ASSERT( nSttRow >= 0 && nSttRow <= 65535);
		ASSERT( nEndRow >= 0 && nEndRow <= 65535);
		ASSERT( nSttRow <= nEndRow );


		CString cs1 = GetCellStr(nSttRow ,0);
		CString cs2 = GetCellStr(nEndRow,255);
		LPCTSTR lpx1 = (LPCTSTR)cs1;
		LPCTSTR lpx2 = (LPCTSTR)cs2;

		LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(lpx1), COleVariant(lpx2));
		ASSERT(lpDisp);
		m_Range.AttachDispatch(lpDisp);

		m_Range.Delete( COleVariant((long)1) ) ;

		m_Range.ReleaseDispatch();
		m_Range.DetachDispatch();
	}

	// 현재의 deleteline정보 초기화
	m_DeleteLine.RemoveAll();
}

void CXLControl::InsertRowLine(long nInsertLine, long nQtyLine/*=1*/)
{
	nInsertLine--;

	CString cs1 = GetCellStr(nInsertLine,0);
	CString cs2 = GetCellStr(nInsertLine + nQtyLine - 1,255);


	LPCTSTR lpx1 = (LPCTSTR)cs1;
	LPCTSTR lpx2 = (LPCTSTR)cs2;

	LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(lpx1), COleVariant(lpx2));
	ASSERT(lpDisp);
	m_Range.AttachDispatch(lpDisp);

	m_Range.Insert(m_covOptional);

	m_Range.ReleaseDispatch();
	m_Range.DetachDispatch();

}

void CXLControl::Copy(long nSttRow, long nEndRow, long nDestinationRow)
{
	nSttRow--;
	nEndRow--;
	nDestinationRow--;

	CString cs1 = GetCellStr(nSttRow,0);
	CString cs2 = GetCellStr(nEndRow,255);
	CString cs3 = GetCellStr(nDestinationRow,0);
	CString cs4 = GetCellStr(nDestinationRow,255);

	LPCTSTR lpx1 = (LPCTSTR)cs1;
	LPCTSTR lpx2 = (LPCTSTR)cs2;
	LPCTSTR lpxDes1 = (LPCTSTR)cs3;
	LPCTSTR lpxDes2 = (LPCTSTR)cs4;

	LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(lpx1), COleVariant(lpx2));
	LPDISPATCH lpDes = m_Sheet.GetRange(COleVariant(lpxDes1), COleVariant(lpxDes2));
	
	m_Range.AttachDispatch(lpDisp);

    VARIANT v;
	v.vt = VT_DISPATCH;
	v.pdispVal = lpDes;
	m_Range.Copy(v);

	//xlPasteAllExceptBorders
//	m_Range.PasteSpecial(-4163,-4142,covOpt,covOpt);


	m_Range.ReleaseDispatch();
	m_Range.DetachDispatch();
}

void CXLControl::CopyRange(long nSourceSttRow, long nSourceSttCol, long nSourceEndRow, long nSourceEndCol,
						   CString sTargetSheet,   long nTargetSttRow, long nTargetSttCol)
{
	CString cs1 = GetCellStr(nSourceSttRow, nSourceSttCol);
	CString cs2 = GetCellStr(nSourceEndRow, nSourceEndCol);
	CString cs3 = GetCellStr(nTargetSttRow, nTargetSttCol);
	CString cs4 = GetCellStr(nTargetSttRow + (nSourceEndRow - nSourceSttRow), nTargetSttCol + (nSourceEndCol - nSourceSttCol));

	LPCTSTR lpx1 = (LPCTSTR)cs1;
	LPCTSTR lpx2 = (LPCTSTR)cs2;
	LPCTSTR lpxTarget1 = (LPCTSTR)cs3;
	LPCTSTR lpxTarget2 = (LPCTSTR)cs4;

	LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(lpx1), COleVariant(lpx2));
	m_Range.AttachDispatch(lpDisp);

	SetActiveSheet(sTargetSheet);
	LPDISPATCH lpTarget = m_Sheet.GetRange(COleVariant(lpxTarget1), COleVariant(lpxTarget2));
	
	VARIANT v;
	v.vt = VT_DISPATCH;
	v.pdispVal = lpTarget;
	m_Range.Copy(v);

	//xlPasteAllExceptBorders
//	m_Range.PasteSpecial(-4163,-4142,covOpt,covOpt);

	//this->m_Sheet.Paste(v, COleVariant((short)FALSE));

	m_Range.ReleaseDispatch();
	m_Range.DetachDispatch();
}

void CXLControl::InsertCopyRowLine(long nSttRow, long nEndRow, long nDesRow)
{
	InsertRowLine(nDesRow, nEndRow - nSttRow + 1);
	Copy(nSttRow,nEndRow,nDesRow);
}

void CXLControl::DeleteColSell(LPCTSTR lp1,LPCTSTR lp2)
{
	LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(lp1), COleVariant(lp2));
	ASSERT(lpDisp);
	m_Range.AttachDispatch(lpDisp);

	m_Range.Delete( COleVariant((long)1) );

	m_Range.ReleaseDispatch();
	m_Range.DetachDispatch();
}

void CXLControl::DeleteColSell(long nSttRow, long nSttCol, long nEndRow, long nEndCol)
{
	ASSERT(nSttRow >= 1 && nSttRow <= 65536);
	ASSERT(nEndRow >= 1 && nEndRow <= 65536);
	ASSERT(nSttCol >= 1 && nSttCol <= 256);
	ASSERT(nEndCol >= 1 && nEndCol <= 256);


	nSttRow--;nSttCol--;nEndRow--;nEndCol--;
	CString cs1 = GetCellStr(nSttRow,nSttCol);
	CString cs2 = GetCellStr(nEndRow,nEndCol);
	DeleteColSell(cs1,cs2);
}


// 수평정렬
// TA_LEFT or TA_RIGHT or TA_CENTER
void CXLControl::SetHoriAlign(LPCTSTR lpstr,LPCTSTR lpstr2, long TA_ALIGN)
{
	VARIANT v;
	v.vt = VT_I4;
	v.lVal = 4;

	if(TA_ALIGN == TA_LEFT)
		v.lVal = xlLeft;
	else if(TA_ALIGN == TA_CENTER)
		v.lVal = xlCenterAcrossSelection;	//xlCenter;
	else if(TA_ALIGN == TA_RIGHT)
		v.lVal = xlRight;
//	else
//		v.lVal = xlCenterAcrossSelection;

	LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(lpstr), COleVariant(lpstr2));
	ASSERT(lpDisp);
	m_Range.AttachDispatch(lpDisp);

	m_Range.SetHorizontalAlignment( v );

	m_Range.ReleaseDispatch();
	m_Range.DetachDispatch();
}

void CXLControl::SetHoriAlign(long nRow, long nCol,long nRow2, long nCol2, long TA_ALIGN)
{
	ASSERT(nRow >= 0 && nRow <= 65535);
	ASSERT(nCol >= 0 && nCol <= 255);
	ASSERT(nRow2 >= 0 && nRow2 <= 65535);
	ASSERT(nCol2 >= 0 && nCol2 <= 255);


	CString cs1 = GetCellStr( nRow, nCol);
	CString cs2 = GetCellStr( nRow2, nCol2);

	SetHoriAlign(cs1,cs2,TA_ALIGN);
}

// 수직정렬
// nAlign :=> top = 1, center = 2, bottom = 3;
void CXLControl::SetVerAlign(CString strCell1, long nAlign/*=1*/)
{
	SetVerAlign(strCell1, strCell1, nAlign);
}

void CXLControl::SetVerAlign(CString strCell1, CString strCell2, long nAlign/*=1*/)
{
	// Cell 범위선택
	LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(strCell1), COleVariant(strCell2));
	ASSERT(lpDisp);
	m_Range.AttachDispatch(lpDisp);
	m_Range.SetVerticalAlignment(COleVariant(nAlign));
}

void CXLControl::SetVerAlign(long nRow1, long nCol1, long nAlign/*=1*/)
{
	CString strCell1 = GetCellStr(nRow1, nCol1);
	SetVerAlign(strCell1, strCell1, nAlign);
}

void CXLControl::SetVerAlign(long nRow1, long nCol1, long nRow2, long nCol2, long nAlign/*=1*/)
{
	// 첫번째, 두번째  Cell선택
	CString strCell1 = GetCellStr(nRow1, nCol1);
	CString strCell2 = GetCellStr(nRow2, nCol2);

	SetVerAlign(strCell1, strCell2, nAlign);
}


/*/////////////////
	XL.SetNumberFormat(1,1, 2,2,	"0 ");
	XL.SetNumberFormat("A1",		"#,##0 ");
	XL.SetNumberFormat("A1", "C1",	"0.00 ");
	XL.SetNumberFormat("A1", "A10", "0.00E+00");	// 지수형식
/*/////////
void CXLControl::SetNumberFormat(long nRow, long nCol,LPCTSTR strFormat)
{
	CString strCell = GetCellStr(nRow, nCol);
	SetNumberFormat(strCell, strCell, strFormat);
}

void CXLControl::SetNumberFormat(long nRow1, long nCol1, long nRow2, long nCol2, LPCTSTR strFormat)
{
	CString strCell1 = GetCellStr(nRow1, nCol1);
	CString strCell2 = GetCellStr(nRow2, nCol2);
	SetNumberFormat(strCell1, strCell2, strFormat);
}

void CXLControl::SetNumberFormat(LPCTSTR strCell, LPCTSTR strFormat)
{
	SetNumberFormat(strCell, strCell, strFormat);
}

void CXLControl::SetNumberFormat(LPCTSTR strCell1, LPCTSTR strCell2, LPCTSTR strFormat)
{
	// Set Value
	LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(strCell1), COleVariant(strCell2));
	ASSERT(lpDisp);
	m_Range.AttachDispatch(lpDisp);

	m_Range.SetNumberFormat(COleVariant(strFormat));
	m_Range.ReleaseDispatch();
	m_Range.DetachDispatch();
}

// SetTextToColumns => Text를 숫자로 바꾸는 함수
// 예 "123" --> 숫자 123으로 바꾼다.
// 한번에 한 칼럼만을 바꿀 수 있다. 그러나 행은 여러 행이어도 상관 없다.
// SetTextToColumns("A2");	// A2만을 바꿈
// SetTextToColumns(nCol, nRow1, nRow2);	// (nRow1, nCol), (nRow2, nCol) 까지 바꿈
// SetTextToColumns("C2", "C20");	// 여기서 ("C1", D20")을 하면 에러를 유발 한다. 열이 같아야 한다.
void CXLControl::SetTextToColumns(CString strCell)
{
	SetTextToColumns(strCell, strCell);
}

void CXLControl::SetTextToColumns(long nCol, long nRow1, long nRow2)
{
	CString strCell1, strCell2;
	strCell1 = GetCellStr(nRow1 ,nCol);
	strCell2 = GetCellStr(nRow2 ,nCol);

	SetTextToColumns( strCell1, strCell2);
}

void CXLControl::SetTextToColumns(CString strCell1, CString strCell2)
{
	LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(strCell1), COleVariant(strCell2));
	ASSERT(lpDisp);
	m_Range.AttachDispatch(lpDisp);

	COleSafeArray* Array = new COleSafeArray;
	DWORD numElements[]={1, 2};
	DWORD dwDims = 2;
	Array->Create(VT_I2, dwDims, numElements);

	long lNewValue = 1;
	long lNewValue1 = 1;
	long index[2];
	index[0] = 0;	// row
	index[1] = 0;	// col
	Array->PutElement(index, &lNewValue);
    index[0] = 0;	// row
    index[1] = 1;	// col
	Array->PutElement(index, &lNewValue1);

	COleVariant  covTrue((short)TRUE);
	COleVariant  covFalse((short)FALSE);

	long i = 1;
	const VARIANT& Destination = COleVariant(m_Range.Get_Default(COleVariant(i),COleVariant(i)));
	long DataType = 1;
	long TextQualifier = 1;
	const VARIANT& ConsecutiveDelimiter = COleVariant((short)(0));
	const VARIANT& Tab = COleVariant((short)(1));
	const VARIANT& Semicolon = COleVariant((short)(0));
	const VARIANT& Comma = COleVariant((short)(0));
	const VARIANT& Space = COleVariant((short)(0));
	const VARIANT& Other = COleVariant((short)(0));
	const VARIANT& OtherChar = COleVariant((short)(0));
	const VARIANT& FieldInfo = COleVariant(Array);

	m_Range.TextToColumns(Destination, DataType, TextQualifier, ConsecutiveDelimiter, Tab,
		Semicolon, Comma, Space, Other, OtherChar,  FieldInfo);

	Array->Detach();
	_DELPTR(Array);
}

void CXLControl::SetLineStyle(LPCTSTR lpstr1, LPCTSTR lpstr2, long newValue)
{
	LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(lpstr1), COleVariant(lpstr2));
	ASSERT(lpDisp);
	m_Range.AttachDispatch(lpDisp);

	Borders bds( m_Range.GetBorders() );

	long nCount = bds.GetCount();
	Border bd( bds.GetItem( 1 ) );

	VARIANT v;// = bd.GetLineStyle();
	v.vt = VT_I4;
	v.lVal = newValue;
	bd.SetLineStyle(v);

	m_Range.ReleaseDispatch();
	m_Range.DetachDispatch();

}
VARIANT CXLControl::GetLineStyle(LPCTSTR lpstr1, LPCTSTR lpstr2)
{
	LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(lpstr1), COleVariant(lpstr2));
	ASSERT(lpDisp);
	m_Range.AttachDispatch(lpDisp);

	Borders bds( m_Range.GetBorders() );

	long nCount = bds.GetCount();
	Border bd( bds.GetItem( 1 ) );

	VARIANT v = bd.GetLineStyle();

	m_Range.ReleaseDispatch();
	m_Range.DetachDispatch();
	
	return v;
}

// 표시한 두 범위를 합친다.
// SetMergeCell(2,2, 3,3);
void CXLControl::SetMergeCell(long nRow1, long nCol1, long nRow2, long nCol2)
{
	CString strCell1 = GetCellStr(nRow1, nCol1);
	CString strCell2 = GetCellStr(nRow2, nCol2);
	SetMergeCell(strCell1, strCell2);
}

// SetMergeCell("A1", "B2");
void CXLControl::SetMergeCell(LPCTSTR lpstr1, LPCTSTR lpstr2)
{
	COleVariant covTrue = COleVariant((short)TRUE);

	LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(lpstr1), COleVariant(lpstr2));
	ASSERT(lpDisp);
	m_Range.AttachDispatch(lpDisp);
	m_Range.SetMergeCells( covTrue );	// default : FALSE

	m_Range.ReleaseDispatch();
	m_Range.DetachDispatch();
}

// 여러셀 테두리 : 모든 셀에 테두리를 친다.
void CXLControl::SetBorders(long nRow, long nCol, long nStyle/*=1*/)
{
	CString strCell = GetCellStr(nRow, nCol);
	SetBorders(strCell, strCell, nStyle);
}

void CXLControl::SetBorders(long nRow1, long nCol1, long nRow2, long nCol2, long nStyle/*=1*/)
{
	// 첫번째, 두번째  Cell선택
	CString strCell1 = GetCellStr(nRow1, nCol1);
	CString strCell2 = GetCellStr(nRow2, nCol2);
	
	SetBorders(strCell1, strCell2, nStyle);
}

void CXLControl::SetBorders(CString strCell, long nStyle/*=1*/)
{
	SetBorders(strCell, strCell, nStyle);
}

void CXLControl::SetBorders(CString strCell1, CString strCell2, long nStyle/*=1*/)
{
	// Cell 범위선택
	LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(strCell1), COleVariant(strCell2));
	ASSERT(lpDisp);
	m_Range.AttachDispatch(lpDisp);

	m_Borders = m_Range.GetBorders();
	m_Borders.SetLineStyle(COleVariant((short)nStyle));
}

// 한 셀 테두리와 값을 입력
void CXLControl::TextBoxValue(long nRow, long nCol, CString strValue, long nStyle/*=1*/)
{
	// Cell선택
	CString strCell = GetCellStr(nRow, nCol);
	TextBoxValue(strCell, strValue, nStyle);
}

void CXLControl::TextBoxValue(CString strCell, CString strValue, long nStyle/*=1*/)
{
	// Set Value
	LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(strCell), COleVariant(strCell));
	ASSERT(lpDisp);
	m_Range.AttachDispatch(lpDisp);

	m_Range.SetValue(COleVariant(strValue));
	
	m_Borders = m_Range.GetBorders();
	m_Borders.SetLineStyle(COleVariant((short)nStyle));
}

// Font Style Change
// XL.SetFonts(nRow1, nCol1, nRow2, nCol2, 18, "굴림체", TRUE); // TRUE 문자를 두껍게
void CXLControl::SetFonts(long nRow1, long nCol1, short nSize/*=10*/, CString strFont/*=바탕체*/, short nColor/*=1*/, long bBold/*=TRUE*/)
{
	CString strCell = GetCellStr(nRow1, nCol1);
	SetFonts(strCell, strCell, nSize, strFont, nColor, bBold);
}

void CXLControl::SetFonts(long nRow1, long nCol1, long nRow2, long nCol2,
											short nSize/*=10*/, CString strFont/*=바탕체*/, short nColor/*=1*/, long bBold/*=TRUE*/)
{

	// 첫번째, 두번째  Cell선택
	CString strCell1 = GetCellStr(nRow1, nCol1);
	CString strCell2 = GetCellStr(nRow2, nCol2);

	SetFonts(strCell1, strCell2, nSize, strFont, nColor, bBold);
}

void CXLControl::SetFonts(CString strCell, short nSize/*=10*/, CString strFont/*=바탕체*/, short nColor/*=1*/, long bBold/*=TRUE*/)
{
	SetFonts(strCell, strCell, nSize, strFont, nColor, bBold);
}


void CXLControl::SetFonts(CString strCell1, CString strCell2,
											short nSize/*=10*/, CString strFont/*=바탕체*/, short nColor/*=1*/, long bBold/*=TRUE*/)
{
	// Cell 범위선택
	LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(strCell1), COleVariant(strCell2));
	ASSERT(lpDisp);
	m_Range.AttachDispatch(lpDisp);
	
	// FONT지정
	m_Font = m_Range.GetFont();
	m_Font.SetSize(COleVariant(nSize));
	m_Font.SetName(COleVariant(strFont));
	m_Font.SetBold(COleVariant(bBold));
	if(nColor != 0)
		m_Font.SetColorIndex(COleVariant(nColor));
}
void CXLControl::SetFontCharacters(long nRow, long nCol, CString str, BOOL bOneChar)
{
	CPoint2ObjArray scriptArr;
	int nLength = str.GetLength();
	int nExcelIdx = 0;
	int nCount2Byte = 0;
	BOOL bRemoveFirst = FALSE;
	if (nLength > 0)
	{
		CString strTemp = str.Left(2);
		if (strTemp == _T("'="))
			bRemoveFirst = TRUE;
	}
	for (int idx = 0; idx < nLength; idx++)
	{
		char cChar = str.GetAt(idx);
		int nChar = str.GetAt(idx);
		if (nChar < 0)
			nCount2Byte++;
		if (cChar == '^' || cChar == '_')
		{
			int nCount = 0;
			if (bOneChar)
				nCount = 1;
			else
			{
				CString strTemp;
				for (int i = idx+1; i < nLength; i++)
				{
					char cChartemp = str.GetAt(i);
					if (cChartemp == ' ')
						break;
					else
						strTemp.AppendChar(cChartemp);
				}

				nCount = strTemp.GetLength();
			}
			str.Delete(idx);

			nExcelIdx = idx+1;
			nExcelIdx -= int(nCount2Byte / 2.0);
			if (bRemoveFirst)
				nExcelIdx--;

			CPoint2 po(nExcelIdx, nCount);
			BOOL bSuperscript = cChar == '^' ? TRUE : FALSE;
			scriptArr.AddPoint2Obj(po, bSuperscript ? 1.0 : -1.0);

			nLength = str.GetLength();
		}
	}

	SetXL(nRow, nCol, str);

	for (int idx = 0; idx < scriptArr.GetCount(); idx++)
		SetFontCharacters(nRow, nCol, (short)scriptArr.GetPoint2(idx).x, (short)scriptArr.GetPoint2(idx).y, scriptArr.GetObjDouble(idx) > 0 ? TRUE : FALSE);

}
void CXLControl::SetFontCharacters(long nRow, long nCol, short nstart, short length, BOOL bSuperscript)
{
	// Cell 범위선택
	CString strCell = GetCellStr(nRow, nCol);
	LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(strCell), COleVariant(strCell));
	ASSERT(lpDisp);
	m_Range.AttachDispatch(lpDisp);
	LPDISPATCH lpDispchar = m_Range.GetCharacters(COleVariant(nstart), COleVariant(length));
	Characters characters;
	characters.AttachDispatch(lpDispchar);

	LPDISPATCH lpDispFont = characters.GetFont();
	XLFont fontchar;
	fontchar.AttachDispatch(lpDispFont);

	if (bSuperscript)
		fontchar.SetSuperscript(COleVariant((short)TRUE));	// 윗첨자
	else
		fontchar.SetSubscript(COleVariant((short)TRUE));	// 아래첨자

	fontchar.DetachDispatch();
	characters.DetachDispatch();
	m_Range.DetachDispatch();
}
 
// Cell에 위, 아래, 사선등을 긋는다.
void CXLControl::CellLine(long nRow, long nCol, long nEdge/*=BOTTOM*/, long nStyle/*=1*/, long nWeight/*=2*/)
{
	CString strCell1 = GetCellStr(nRow, nCol);
	CellLine(strCell1, strCell1, nEdge, nStyle, nWeight);
}

void CXLControl::CellLine(long nRow, long nCol, long nRow1, long nCol1, long nEdge/*=BOTTOM*/, long nStyle/*=1*/, long nWeight/*=2*/)
{
	CString strCell1 = GetCellStr(nRow, nCol);
	CString strCell2 = GetCellStr(nRow1, nCol1);
	CellLine(strCell1, strCell2, nEdge, nStyle, nWeight);
}

void CXLControl::CellLine(CString strCell, long nEdge/*=BOTTOM*/, long nStyle/*=1*/, long nWeight/*=2*/)
{
	// Cell선택
	CellLine(strCell, strCell, nEdge, nStyle, nWeight);
}
 
void CXLControl::CellLine(CString strCell1, CString strCell2, long nEdge/*=BOTTOM*/, long nStyle/*=1*/, long nWeight/*=2*/)
{
	// Set Value
	LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(strCell1), COleVariant(strCell2));
	ASSERT(lpDisp);
	m_Range.AttachDispatch(lpDisp);

	m_Borders = m_Range.GetBorders();
	lpDisp = m_Borders.GetItem(nEdge);
	m_Borders.AttachDispatch(lpDisp);

	m_Borders.SetLineStyle(COleVariant((short)nStyle));
	m_Borders.SetWeight(COleVariant((short)nWeight));
}

#if 0
	// 참고 Index 모음

	// Borders : 방향
	// xlDiagonalDown = 5;  // 위에서 아래로 역 슬래위
	// xlDiagonalUp = 6;	// 아래에서 위로 /
	// xlEdgeLeft = 7;
	// xlEdgeTop = 8;
    // xlEdgeBottom = 9;	
	// xlEdgeRight = 10; 

	// LineStyle
	// xlContinuous = 1
	// xlDash = -4115
	// xlDashDot = 4
	// xlDashDotDot = 5
	// xlDot = -4118
	// xlDouble = -4119
	// xlLineStyleNont = -4142
	// xlSlantDashDot = 13

	// XlBorderWeight
	// xlHairLine = 1;
	// xlMedium = -4138;
	// xlThick = 4;
	// xlThin = 2;
#endif

// 셀 테두리를 그음 : 지정한 셀의 범위에서 바깥쪽에만 Line이 나타남
void CXLControl::CellOutLine(long nRow, long nCol, long nStyle/*=1*/, long nColorIndex/* = 1*/, long nWeight/* = 2*/)
{
	CString strCell = GetCellStr(nRow, nCol);
	CellOutLine(strCell, strCell, nStyle, nColorIndex, nWeight);
}
void CXLControl::CellOutLine(long nRow, long nCol, long nRow1, long nCol1, long nStyle/*=1*/, long nColorIndex/* = 1*/, long nWeight/* = 2*/)
{
	CString strCell1 = GetCellStr(nRow, nCol);
	CString strCell2 = GetCellStr(nRow1, nCol1);

	CellOutLine(strCell1, strCell2, nStyle, nColorIndex, nWeight);
}

void CXLControl::CellOutLine(CString strCell, long nStyle/*=1*/, long nColorIndex/* = 1*/, long nWeight/* = 2*/)
{
	CellOutLine(strCell, strCell, nStyle, nColorIndex, nWeight);
}

void CXLControl::CellOutLine(CString strCell1, CString strCell2, long nStyle/*=1*/, long nColorIndex/* = 1*/, long nWeight/* = 2*/)
{
	LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(strCell1), COleVariant(strCell2));
	ASSERT(lpDisp);
	m_Range.AttachDispatch(lpDisp);

	VARIANT Color(COleVariant((short)16));
	m_Range.BorderAround(COleVariant((short)nStyle), nWeight, nColorIndex, Color);
}

// 셀의 넓이를 지정
void CXLControl::SetCellWidth(long nCol1, long nCol2, long nLength)
{
	CString strCell1 = GetCellStr(1, nCol1);
	CString strCell2 = GetCellStr(1, nCol2);
	SetCellWidth(strCell1, strCell2, nLength);
}

void CXLControl::SetCellWidth(long nCol, long nLength)
{
	CString strCell1 = GetCellStr(1, nCol);
	SetCellWidth(strCell1, strCell1, nLength);
}

void CXLControl::SetCellWidth(LPCTSTR x, long nLength)
{
	SetCellWidth(x, x, nLength);
}

void CXLControl::SetCellWidth(LPCTSTR x, LPCTSTR x1, long nLength)
{
	LPDISPATCH	lpDisp = m_Sheet.GetRange(COleVariant(x), COleVariant(x1));
	ASSERT(lpDisp);
	m_Range.AttachDispatch(lpDisp);

	VARIANT width = m_Range.GetColumnWidth();
	width.dblVal = nLength;

	m_Range.SetColumnWidth(width);
	m_Range.ReleaseDispatch();
	m_Range.DetachDispatch();
}

//  셀의 넓이 구하기
long CXLControl::GetCellWidth(long nCol1, long nCol2/*=-1*/)
{
	if( nCol2 == -1 ) nCol2 = nCol1;

	CString strCell1 = GetCellStr(1, nCol1);
	CString strCell2 = GetCellStr(1, nCol2);
	LPDISPATCH	lpDisp = m_Sheet.GetRange(COleVariant(strCell1), COleVariant(strCell2));
	ASSERT(lpDisp);
	m_Range.AttachDispatch(lpDisp);

	VARIANT width = m_Range.GetColumnWidth();
	m_Range.ReleaseDispatch();
	m_Range.DetachDispatch();
	return long(width.dblVal);
}

// 셀의 높이 구하기
double CXLControl::GetCellHeight(long nRow1, long nRow2/*=-1*/)
{
	if( nRow2 == -1 ) nRow2 = nRow1;

	CString strCell1 = GetCellStr(nRow1, 1);
	CString strCell2 = GetCellStr(nRow2, 1);
	LPDISPATCH	lpDisp = m_Sheet.GetRange(COleVariant(strCell1), COleVariant(strCell2));
	ASSERT(lpDisp);
	m_Range.AttachDispatch(lpDisp);

	VARIANT height = m_Range.GetRowHeight();
	m_Range.ReleaseDispatch();
	m_Range.DetachDispatch();
	return height.dblVal;
}

void CXLControl::SetCellHeight(long nRow, double Height)
{
	SetCellHeight(nRow, nRow, Height);
}

void CXLControl::SetCellHeight(long nRow1, long nRow2, double Height)
{
	CString strCell1 = GetCellStr(nRow1, 1);
	CString strCell2 = GetCellStr(nRow2, 1);
	LPDISPATCH	lpDisp = m_Sheet.GetRange(COleVariant(strCell1), COleVariant(strCell2));
	ASSERT(lpDisp);
	m_Range.AttachDispatch(lpDisp);

	m_Range.SetRowHeight(COleVariant(Height));
	m_Range.ReleaseDispatch();
	m_Range.DetachDispatch();
}
///////////////////////////////////////////////////////////////
// * 셀과 상관없는 선 그리기 * //
///////////////////////////////////////////////////////////////	
void CXLControl::DrawRootLine(long nRowStt, long nColStt, long nRowEnd, long nColEnd)
{
	CRect2 rectStt = GetXlCoordinatesXY(nRowStt, nColStt);
	CRect2 rectEnd = GetXlCoordinatesXY(nRowEnd, nColEnd);

	double dCellSttW = fabs(rectStt.right - rectStt.left);
	double dCellSttH = fabs(rectStt.top - rectStt.bottom);

	double dTermW = dCellSttW / 15.0;
	double dTermH = dCellSttH / 15.0;
	
	CPoint2 po1(rectStt.left+dTermW, rectStt.bottom-dTermH);
	CPoint2 po2(rectStt.left+dTermW+dCellSttW*2.0/10.0, rectStt.top+dTermH);
	DrawLine(po1.x, po1.y, po2.x, po2.y, (long)0.25);

	po1 = CPoint2(rectEnd.right-dTermW, rectEnd.top+dTermH);
	DrawLine(po1.x, po1.y, po2.x, po2.y, (long)0.25);

	po1 = CPoint2(rectStt.left+dTermW, rectStt.bottom-dTermH);
	po2 = CPoint2(rectStt.left+dTermW-dCellSttW*2.0/10.0, rectStt.bottom-dCellSttH*2.0/10.0);
	DrawLine(po1.x, po1.y, po2.x, po2.y, (long)0.25);
}
CRect2 CXLControl::GetXlCoordinatesXY(long nRow, long nCol)
{
	CRect2 rect(0,0,0,0);
	CString range = GetCellStr(nRow, nCol);

	Range oRange;
	oRange = m_Sheet.GetRange(COleVariant(range), COleVariant(range));

	rect.left = oRange.GetLeft().dblVal;
	rect.top = oRange.GetTop().dblVal;
	rect.right = oRange.GetLeft().dblVal + oRange.GetWidth().dblVal;
	rect.bottom = oRange.GetTop().dblVal + oRange.GetHeight().dblVal;

	return rect;
}
void CXLControl::DrawLine(double Sx, double Sy, double Ex, double Ey, long nWeight/*=1*/, long nColor/*=8*/, long nStyle/*=1*/)
{
	LPDISPATCH lpDisp = m_Sheet.GetShapes();
	m_Shapes.AttachDispatch(lpDisp);
		
	lpDisp = m_Shapes.AddLine((float)Sx,(float)Sy,(float)Ex,(float)Ey);
	Shape shp(lpDisp);
	LineFormat Line(shp.GetLine());
	Line.SetWeight((float)nWeight);
	Line.SetStyle((long)nStyle);
	
	ColorFormat Color(Line.GetForeColor());
	Color.SetSchemeColor(nColor);

	// black = 8
	// red = 10
	// blue = 12

	// style
	// 1 = solid
	// 2 = line sqaredot
	// 3 = 점선
	// 4 = line dash
	// 5 = ...
	//m_Shapes.DetachDispatch();
}
void CXLControl::DrawLine(double Sx, double Sy, double Ex, long nWeight/*=1*/, long nColor/*=8*/)
{
	LPDISPATCH lpDisp = m_Sheet.GetShapes();
	m_Shapes.AttachDispatch(lpDisp);
		
	lpDisp = m_Shapes.AddLine((float)Sx,(float)Sy,(float)Ex,(float)Sy);

	Shape shp(lpDisp);
	LineFormat Line(shp.GetLine());
	Line.SetWeight((float)nWeight);
	// Line.SetStyle(long nStyle);
	
	ColorFormat Color(Line.GetForeColor());
	Color.SetSchemeColor(nColor);
	// black = 8
	// red = 10
	// blue = 12

	// style
	// 1 = solid
	// 2 = line sqaredot
	// 3 = 점선
	// 4 = line dash
	// 5 = ...
	//m_Shapes.DetachDispatch();
}

void CXLControl::DrawTextBox(long nSttRow, long nSttCol, long nEndRow, long nEndCol, CString strText)
{
	CString rangeStt = GetCellStr(nSttRow, nSttCol);
	CString rangeEnd = GetCellStr(nEndRow, nEndCol);
	if (nEndRow==0 ||  nEndCol == 0)
		rangeEnd = GetCellStr(nSttRow, nSttCol);

	Range oRange;
	oRange = m_Sheet.GetRange(COleVariant(rangeStt), COleVariant(rangeEnd));

	COleVariant VLeft, VTop, VWidth, VHeight;
	VLeft = oRange.GetLeft();
	VTop = oRange.GetTop();
	VWidth = oRange.GetWidth();
	VHeight = oRange.GetHeight();

	LPDISPATCH lpDisp = m_Sheet.GetShapes();
	m_Shapes.AttachDispatch(lpDisp);
	Shape shp;
	lpDisp = m_Shapes.AddTextbox(1, (float)VLeft.dblVal, (float)VTop.dblVal, (float)VWidth.dblVal, (float)VHeight.dblVal);

	shp.AttachDispatch(lpDisp);
	lpDisp = shp.GetTextFrame();

	// TextFrame
	TextFrame vTextframe(lpDisp);
	lpDisp = vTextframe.Characters(COleVariant((short)0),COleVariant((short)0));

	long nHorz = 2; //left : 1 center : 2 right : 3
	vTextframe.SetVerticalAlignment( nHorz);
	vTextframe.SetHorizontalAlignment( nHorz);
	vTextframe.SetOrientation(1);
	
	// Characters 
	Characters vCharacters(lpDisp);
	vCharacters.SetText(strText);

	XLFont vFont(vCharacters.GetFont());
	COleVariant  covTrue((short)TRUE);
	//COleVariant  covFalse((short)FALSE);

	vFont.SetSize(COleVariant((short)10));
	vFont.SetName(COleVariant(_T("굴림체")));
	vFont.SetBold(covTrue);

	LineFormat vLineformat(shp.GetLine());
	vLineformat.SetVisible(FALSE);

	FillFormat vFillformat(shp.GetFill());
	vFillformat.SetVisible(FALSE);

	shp.DetachDispatch();

	//// Font
	//Font vFont(vCharacters.GetFont());
	//if(m_bDomUseColor)	// Color 사용
	//	vFont.SetColor( COleVariant( (long)pDomP->GetRGBFromCADColor(pDomP->GetDomColor(pDom->m_nColor,pDom->GetsLayer()))) );
	//if (pDom->m_TextStyle.Height * RateTextHeight >= 1.0)
	//	vFont.SetSize( COleVariant(pDom->m_TextStyle.Height * RateTextHeight) );
	//else
	//	vFont.SetSize( COleVariant(1.0) );
	//// LineFormat
	//LineFormat vLineformat(shp.GetLine());
	//vLineformat.SetVisible(FALSE);
	//// FillFormat
	//FillFormat vFillformat(shp.GetFill());
	//vFillformat.SetVisible(FALSE);
	//shp.DetachDispatch();
}

void CXLControl::DrawTextBox(long nSttRow, long nSttCol, long nEndRow, long nEndCol, long nHorz/*= 2*/, short nSize/*= 10*/, CString strFont/*=바탕체*/, CString strText)
{
	CString rangeStt = GetCellStr(nSttRow, nSttCol);
	CString rangeEnd = GetCellStr(nEndRow, nEndCol);
	if (nEndRow==0 ||  nEndCol == 0)
		rangeEnd = GetCellStr(nSttRow, nSttCol);

	Range oRange;
	oRange = m_Sheet.GetRange(COleVariant(rangeStt), COleVariant(rangeEnd));

	COleVariant VLeft, VTop, VWidth, VHeight;
	VLeft = oRange.GetLeft();
	VTop = oRange.GetTop();
	VWidth = oRange.GetWidth();
	VHeight = oRange.GetHeight();

	LPDISPATCH lpDisp = m_Sheet.GetShapes();
	m_Shapes.AttachDispatch(lpDisp);
	Shape shp;
	lpDisp = m_Shapes.AddTextbox(1, (float)VLeft.dblVal, (float)VTop.dblVal, (float)VWidth.dblVal, (float)VHeight.dblVal);

	shp.AttachDispatch(lpDisp);
	lpDisp = shp.GetTextFrame();

	// TextFrame
	TextFrame vTextframe(lpDisp);
	lpDisp = vTextframe.Characters(COleVariant((short)0),COleVariant((short)0));

	//nHorz = 2; //left : 1 center : 2 right : 3
	vTextframe.SetVerticalAlignment(nHorz);
	vTextframe.SetHorizontalAlignment(nHorz);
	vTextframe.SetOrientation(1);
	
	// Characters 
	Characters vCharacters(lpDisp);
 	vCharacters.SetText(strText);

	XLFont vFont(vCharacters.GetFont());
	//COleVariant  covTrue((short)TRUE);
	COleVariant  covFalse((short)FALSE);

	vFont.SetSize(COleVariant(nSize));
	vFont.SetName(COleVariant(strFont));
	//vFont.SetBold(covTrue);

	LineFormat vLineformat(shp.GetLine());
	vLineformat.SetVisible(FALSE);

	FillFormat vFillformat(shp.GetFill());
	vFillformat.SetVisible(FALSE);

	shp.DetachDispatch();

	//// Font
	//Font vFont(vCharacters.GetFont());
	//if(m_bDomUseColor)	// Color 사용
	//	vFont.SetColor( COleVariant( (long)pDomP->GetRGBFromCADColor(pDomP->GetDomColor(pDom->m_nColor,pDom->GetsLayer()))) );
	//if (pDom->m_TextStyle.Height * RateTextHeight >= 1.0)
	//	vFont.SetSize( COleVariant(pDom->m_TextStyle.Height * RateTextHeight) );
	//else
	//	vFont.SetSize( COleVariant(1.0) );
	//// LineFormat
	//LineFormat vLineformat(shp.GetLine());
	//vLineformat.SetVisible(FALSE);
	//// FillFormat
	//FillFormat vFillformat(shp.GetFill());
	//vFillformat.SetVisible(FALSE);
	//shp.DetachDispatch();
}

///////////////////////////////////////////////////////////////
// * Plan * //
///////////////////////////////////////////////////////////////	
/*
void CXLControl::AddDomLine(CPlan *pDomP)
{
	LPDISPATCH lpDisp;
	POSITION pos;
	pos = pDomP->m_LineArr.GetHeadPosition();
	while(pos)
	{
		CObLine *pDom = (CObLine*)pDomP->m_LineArr.GetNext(pos);
		CString szLayer = pDom->GetsLayer();
		if(szLayer.IsEmpty()) continue;

		lpDisp = m_Shapes.AddLine( (float)pDom->m_SttPoint.x,(float)pDom->m_SttPoint.y,
							(float)pDom->m_EndPoint.x,(float)pDom->m_EndPoint.y);
		
		CObLayer* pLayer = NULL;
		pDomP->m_LayerArr.Lookup(szLayer,(void*&)pLayer);
		if(pLayer && pLayer->m_LineTypeName != _T("CONTINUOUS")
		{
			Shape shp(lpDisp);
			LineFormat vLineformat(shp.GetLine());					

			if(pLayer->m_LineTypeName == _T("DASH")
			{
				vLineformat.SetDashStyle(4);
			}
			else if(pLayer->m_LineTypeName == _T("DOT")
			{
				vLineformat.SetDashStyle(3);
			}
			else if(pLayer->m_LineTypeName == _T("LDASHDOT")
			{
				vLineformat.SetDashStyle(7);
			}
			else if(pLayer->m_LineTypeName == _T("LDASHDOTDOT")
			{
				vLineformat.SetDashStyle(8);
			}

		}
		if(m_bDomUseColor)	// Color 사용
		{						
			Shape shp(lpDisp);
			LineFormat vLineformat(shp.GetLine());		
			ColorFormat vColorformat(vLineformat.GetForeColor());
			vColorformat.SetRgb( (long)pDomP->GetRGBFromCADColor( pDomP->GetDomColor(pDom->m_nColor,pDom->GetsLayer()) ));
		}
	}
}
*/

/* // CMS
void CXLControl::AddDomImage(CBitmap *pBitmap, double dLeft, double dTop, double dWidth, double dHeight )
{
	CImage Image;
	CString szFileName = _T("C:\\_instemp.jpg";
	
	// 이미지 작성
	if(!Image.CopyFromBmp(pBitmap))
	{
		AfxMessageBox("지원할수 없는 형식입니다 !");		
		return;
	}	
	Image.SaveImage(szFileName.GetBuffer(szFileName.GetLength()));		

	LPDISPATCH lpDisp = m_Sheet.GetShapes();
	m_Shapes.AttachDispatch(lpDisp);	
	lpDisp = m_Shapes.AddPicture(szFileName,TRUE, TRUE,(float)dLeft, (float)dTop, (float) dWidth, (float) dHeight);			
	Shape shp(lpDisp);
	COleVariant dDumy = (long)0;	
	if(dWidth==0)  shp.ScaleWidth(1.0,TRUE,dDumy);
	if(dHeight==0) shp.ScaleHeight(1.0,TRUE,dDumy);	
	m_Shapes.DetachDispatch();
	CFile::Remove(szFileName);
}

void CXLControl::AddDomImage(CPlan *pDomP, CString szFileName, double dLeft, double dTop, double dWidth, double dHeight )
{
	if(pDomP) pDomP->SaveAsImage((long)dWidth,(long)dHeight,szFileName);
	LPDISPATCH lpDisp = m_Sheet.GetShapes();
	m_Shapes.AttachDispatch(lpDisp);	
	lpDisp = m_Shapes.AddPicture(szFileName,TRUE, TRUE,(float)dLeft, (float)dTop, (float) dWidth, (float) dHeight);			
	Shape shp(lpDisp);
	COleVariant dDumy = (long)0;	
	shp.ScaleWidth(1.0,TRUE,dDumy);
	shp.ScaleHeight(1.0,TRUE,dDumy);	
	m_Shapes.DetachDispatch();
	CFile::Remove(szFileName);
}
*/
 
/*
void CXLControl::AddDomPolyLine(CPlan *pDomP)
{
	COleSafeArray SaRetArr;
	LPDISPATCH lpDisp;
//	CObLine* pLine = NULL;

	POSITION pos;
	pos = pDomP->m_PolyLineArr.GetHeadPosition();
	while(pos)
	{
		CObPolyLine* pObPolyLine = (CObPolyLine*)pDomP->m_PolyLineArr.GetNext(pos);
		if(pObPolyLine->GetsLayer().IsEmpty()) continue;

		DWORD numElements[] = { pObPolyLine->m_vLineArr.GetSize(), 2};
		if(pObPolyLine->m_bPolygon) numElements[0]++;

		SaRetArr.Create(VT_R4, 2, numElements);

		for(long m = 0; m < pObPolyLine->m_vLineArr.GetSize(); m++)
		{
			CVector v = pObPolyLine->m_vLineArr[m];

			long index[2];
			float fe;
			index[0] = m;	index[1] = 0;
			fe = (float)v.x;
			SaRetArr.PutElement(index, &fe);


			index[0] = m;	index[1] = 1;
			fe = (float)v.y;
			SaRetArr.PutElement(index, &fe);
		}

		if(pObPolyLine->m_bPolygon)
		{
			long index[2];
			float fe;

			CVector v = pObPolyLine->m_vLineArr[0];

			index[0] = m;	index[1] = 0;
			fe = (float)v.x;
			SaRetArr.PutElement(index, &fe);

			index[0] = m;	index[1] = 1;
			fe = (float)v.y;
			SaRetArr.PutElement(index, &fe);
		}

		lpDisp = m_Shapes.AddPolyline(SaRetArr);
		
		Shape shape(lpDisp);
		FillFormat vFillformat(shape.GetFill());
		vFillformat.SetVisible(FALSE);	
		shape.DetachDispatch();

		CObLayer* pPrevLayer = pDomP->m_pCurLayer;	
		if(pPrevLayer->m_LineTypeName != _T("CONTINUOUS")
		{				
			LineFormat vLineformat(shape.GetLine());					

			if(pPrevLayer->m_LineTypeName == _T("DASH")
			{
				vLineformat.SetDashStyle(3);
			}
			else if(pPrevLayer->m_LineTypeName == _T("DOT")
			{
				vLineformat.SetDashStyle(2);
			}
			else if(pPrevLayer->m_LineTypeName == _T("LDASHDOT")
			{
				vLineformat.SetDashStyle(4);
			}
			else if(pPrevLayer->m_LineTypeName == _T("LDASHDOTDOT")
			{
				vLineformat.SetDashStyle(7);
			}

		}
		if(m_bDomUseColor)	// Color 사용
		{
			Shape shp(lpDisp);
			LineFormat vLineformat(shp.GetLine());
			ColorFormat vColorformat(vLineformat.GetForeColor());
			vColorformat.SetRgb( (long)pDomP->GetRGBFromCADColor(pDomP->GetDomColor(pObPolyLine->m_nColor,pObPolyLine->GetsLayer())));
		}
		SaRetArr.Detach();
		SaRetArr.Destroy();
	} 
}
*/

/*
void CXLControl::AddDomArc(CPlan *pDomP)
{
	POSITION pos;
	pos = pDomP->m_ArcArr.GetHeadPosition();
	while(pos)
	{
		CObArc* pArc = (CObArc*)pDomP->m_ArcArr.GetNext(pos);
		if(pArc->GetsLayer().IsEmpty()) continue;		
		
		if( pArc->m_bFillstyle==TRUE)	// 채원진 원
		{
			LPDISPATCH lpDisp;
			try {
				lpDisp = m_Shapes.AddShape(9,(float)0,
											 (float)0,
											 (float) pArc->m_Radius * 2, 
											 (float) pArc->m_Radius * 2);							
			}
			catch(...)
			{
			}

			Shape shape(lpDisp);			
			FillFormat vFillformat(shape.GetFill());
			ColorFormat vColorformat(vFillformat.GetForeColor());
			vColorformat.SetRgb(0);		
			vFillformat.SetForeColor(vColorformat);
			vFillformat.SetBackColor(vColorformat);
			shape.DetachDispatch();
			
			Window window(lpDisp);
			window.SetLeft((float)pArc->m_SttPoint.x - pArc->m_Radius);
			window.SetTop((float)pArc->m_SttPoint.y - pArc->m_Radius);
		}
		else  if( pArc->m_bCircle==TRUE )	// 빈 원
		{
			LPDISPATCH lpDisp;
			lpDisp = m_Shapes.AddShape(9,(float)0,
											 (float)0,
											 (float) pArc->m_Radius * 2, 
											 (float) pArc->m_Radius * 2);
			Window window(lpDisp);
			window.SetLeft((float)pArc->m_SttPoint.x - pArc->m_Radius);
			window.SetTop((float)pArc->m_SttPoint.y - pArc->m_Radius);

			Shape shape(lpDisp);
			FillFormat vFillformat(shape.GetFill());
			vFillformat.SetVisible(FALSE);	
			shape.DetachDispatch();
			//pArc->m_Ang1 = 0;
			//pArc->m_Ang2 = 360;
			//AddDomArcSub(pArc,pDomP);
		}
		else									// Arc
			AddDomArcSub(pArc,pDomP);
	}
}
*/

/*
void CXLControl::AddDomText(CPlan *pDomP,double RateTextHeight)
{
	LPDISPATCH lpDisp;
	Shape shp;
	POSITION pos;
	pos = pDomP->m_TextArr.GetHeadPosition();
	while(pos)
	{
		CObText*pDom = (CObText*)pDomP->m_TextArr.GetNext(pos);
		if(pDom->GetsLayer().IsEmpty()) continue;

		CRect2 BR = GetMTextBorder((CPoint2)pDom->m_SttPoint,
			pDom->m_TextString,&(pDom->m_TextStyle),
			pDom->m_TextStyle.Height*RateTextHeight);

		lpDisp = m_Shapes.AddTextbox(1,(float)BR.left,(float)BR.top,
										(float)BR.Width(),(float)BR.Height());
		shp.AttachDispatch(lpDisp);
		lpDisp = shp.GetTextFrame();
		// TextFrame
		TextFrame vTextframe(lpDisp);
		lpDisp = vTextframe.Characters(COleVariant((short)0),COleVariant((short)0));
//		vTextframe.SetOrientation(2);
//		if(pDom->m_TextStyle.Horizontal == TA_CENTER)
		long nHorz;
		if(pDom->m_TextStyle.Horizontal == TA_LEFT)			nHorz = 1;
		else if(pDom->m_TextStyle.Horizontal == TA_CENTER)		nHorz = 2;
		else nHorz = 3; // if(pDom->m_TextStyle.Horizontal == TA_RIGHT)	

		if(pDom->m_TextStyle.Angle==90 || pDom->m_TextStyle.Angle==270)
		{
			vTextframe.SetVerticalAlignment( nHorz);
			vTextframe.SetOrientation(2);
		}
		else
		{
			vTextframe.SetHorizontalAlignment( nHorz);
		}

		// Characters 
		Characters vCharacters(lpDisp);
		vCharacters.SetText(pDom->m_TextString);
		// Font
		Font vFont(vCharacters.GetFont());
		if(m_bDomUseColor)	// Color 사용
			vFont.SetColor( COleVariant( (long)pDomP->GetRGBFromCADColor(pDomP->GetDomColor(pDom->m_nColor,pDom->GetsLayer()))) );
		if (pDom->m_TextStyle.Height * RateTextHeight >= 1.0)
			vFont.SetSize( COleVariant(pDom->m_TextStyle.Height * RateTextHeight) );
		else
			vFont.SetSize( COleVariant(1.0) );
		// LineFormat
		LineFormat vLineformat(shp.GetLine());
		vLineformat.SetVisible(FALSE);
		// FillFormat
		FillFormat vFillformat(shp.GetFill());
		vFillformat.SetVisible(FALSE);
		shp.DetachDispatch();
	}
}
*/

/*
void CXLControl::AddDomArcSub(CObArc* pArc,CPlan *pDomP)
{ 
	double X = pArc->m_SttPoint.x;
	double Y = pArc->m_SttPoint.y;
	double dwRadius = pArc->m_Radius;
	double fStartDegrees = pArc->m_Ang1;
	double fSweepDegrees = pArc->m_Ang2;

     double f;         // Current angle in radians
     double fStepAngle = 0.02f;  // The sweep increment value in radians
     double fStartRadians;       // Start angle in radians
     double fEndRadians;         // End angle in radians
     double ix, iy;                // Current uXy on arc
//     double fTwoPi = 2.0f * 3.1415926535897932385f;
     double fTwoPi = 6.283185307196f;

     // Get the starting and ending angle in radians 
	if (fSweepDegrees > 0.0f) 
	{
		fStartRadians = ((fStartDegrees / 360.0f) * fTwoPi);
		fEndRadians = (((fStartDegrees + fSweepDegrees) / 360.0f) *  fTwoPi);
	} 
	else 
	{
		fStartRadians = (((fStartDegrees + fSweepDegrees)  / 360.0f) * fTwoPi);
		fEndRadians =  ((fStartDegrees / 360.0f) * fTwoPi);
	}
	//fStepAngle = 10.0 / dwRadius;
//	fStepAngle = ConstPi / 36;

	// Array 수량 구하기
	BOOL bEndDraw = FALSE;
	long nArraysu = 1;	// 1개가 많이 존재하므로
	for (f = fStartRadians; ; f += fStepAngle) 
	{
		if(f > fEndRadians) { f = fEndRadians; bEndDraw = TRUE; }
		nArraysu++;
		if(bEndDraw) break;
	}

	// 
	COleSafeArray SaRetArr;
	DWORD numElements[] = { nArraysu, 2};
	SaRetArr.Create(VT_R4, 2, numElements);

	ix = X + (double)dwRadius * (double)cos(fStartRadians);
	iy = Y - (double)dwRadius * (double)sin(fStartRadians);
	CPoint2 p1(ix,iy), p2;
	bEndDraw = FALSE;
	for (f = fStartRadians,nArraysu = 0; ; f += fStepAngle) 
	{
		if(f > fEndRadians) { f = fEndRadians; bEndDraw = TRUE; }

		ix = X + (double)dwRadius * (double)cos(f);
		iy = Y - (double)dwRadius * (double)sin(f);

		p2 = CPoint2(ix,iy);

		long index[2];
		float fe;
		if(f == fStartRadians)
		{
			index[0] = nArraysu;	index[1] = 0;
			fe = (float)p1.x;
			SaRetArr.PutElement(index, &fe);
			index[0] = nArraysu;	index[1] = 1;
			fe = (float)p1.y;
			SaRetArr.PutElement(index, &fe);
			nArraysu++;
			if(p1.x < 0) p1.x = 0;
			if(p1.y < 0) p1.y = 0;
//			ASSERT(p1.x >=0 && p1.y >=0);
		}

		index[0] = nArraysu;	index[1] = 0;
		fe = (float)p2.x;
		SaRetArr.PutElement(index, &fe);
		index[0] = nArraysu;	index[1] = 1;
		fe = (float)p2.y;
		SaRetArr.PutElement(index, &fe);
		nArraysu++;
		if(p2.x < 0) p2.x = 0;
		if(p2.y < 0) p2.y = 0;
//		ASSERT(p2.x >=0 && p2.y >=0);

		p1 = p2;		
		if(bEndDraw) break;
     }
	
	LPDISPATCH lpDisp;
	lpDisp = m_Shapes.AddPolyline(SaRetArr);
//	lpDisp = m_Shapes.AddCurve(SaRetArr);
//	m_pShapes->AddCurve(SaRetArr);

	if(m_bDomUseColor)	// Color 사용
	{
		Shape shp(lpDisp);
		LineFormat vLineformat(shp.GetLine());
		ColorFormat vColorformat(vLineformat.GetForeColor());
		vColorformat.SetRgb( (long)pDomP->GetRGBFromCADColor(pDomP->GetDomColor(pArc->m_nColor,pArc->GetsLayer())));
	}

	SaRetArr.Detach();
	SaRetArr.Destroy();	

}
*/

/*
void CXLControl::AddDomSolid(CPlan *pDomP)
{
	COleSafeArray SaRetArr;
	LPDISPATCH lpDisp;

	POSITION pos;
	pos = pDomP->m_SolidArr.GetHeadPosition();
	while(pos)
	{
		CObSolid* pObSolid = (CObSolid*)pDomP->m_SolidArr.GetNext(pos);
		if(pObSolid->GetsLayer().IsEmpty()) continue;
		
		if(pObSolid->m_Point[0] != pObSolid->m_Point[3]) continue;

		DWORD numElements[] = {4, 2};
		SaRetArr.Create(VT_R4, 2, numElements);
		
		for(long m = 0; m < 4; m++)
		{
			CPoint2 Pt;
			if(m==0)		Pt = (CPoint2)pObSolid->m_Point[0];
			else if(m==1)	Pt = (CPoint2)pObSolid->m_Point[1];
			else if(m==2)	Pt = (CPoint2)pObSolid->m_Point[2];
			else			Pt = (CPoint2)pObSolid->m_Point[3];

			long index[2];
			float fe;
			index[0] = m;	index[1] = 0;
			fe = (float)Pt.x;
			SaRetArr.PutElement(index, &fe);


			index[0] = m;	index[1] = 1;
			fe = (float)Pt.y;
			SaRetArr.PutElement(index, &fe);

		}

		lpDisp = m_Shapes.AddPolyline(SaRetArr);

		if(m_bDomUseColor)	// Color 사용
		{
			Shape shp(lpDisp);
			LineFormat vLineformat(shp.GetLine());
			ColorFormat vColorformat(vLineformat.GetForeColor());
			vColorformat.SetRgb( (long)pDomP->GetRGBFromCADColor(pDomP->GetDomColor(pObSolid->m_nColor,pObSolid->GetsLayer())));

			FillFormat vFillformat( shp.GetFill() );
			lpDisp = vFillformat.GetForeColor();
			ColorFormat vCf(lpDisp);
			vCf.SetRgb( (long)pDomP->GetRGBFromCADColor(pDomP->GetDomColor(pObSolid->m_nColor,pObSolid->GetsLayer())));
		
		}
		else
		{
			Shape shp(lpDisp);
			FillFormat vFillformat( shp.GetFill() );
			ColorFormat vColorformat(vFillformat.GetForeColor());
			vColorformat.SetRgb(0);		
			vFillformat.SetForeColor(vColorformat);
			vFillformat.SetBackColor(vColorformat);
		}

		SaRetArr.Detach();
		SaRetArr.Destroy();

	}

}
*/

/*
void CXLControl::SetDomDisRC(double xDis,double yDis)
{
	m_fDisRow = (float)yDis;
	m_fDisCol = (float)xDis;
}
*/

//void CXLControl::AddDomyunRC(CPlan *pDomP,long nRow,long nCol,double Scale /*=0.01*/,double RateTextHeight/*=2.0*/)
//{
//	double y  = m_fDisRow * nRow;
//	double x  = m_fDisCol * nCol;

//	AddDomyun(pDomP,x,y,Scale,RateTextHeight);
//}


//void CXLControl::AddDomyun(CPlan *pDomP,double x,double y,double Scale /*=0.01*/,double RateTextHeight/*=2.0*/)
//{
/*
	CPlan* pDom = pDomP;
	pDom->RedrawByScale(Scale);
	pDom->SetPosition(x,y);

	LPDISPATCH lpDisp = m_Sheet.GetShapes();
	m_Shapes.AttachDispatch(lpDisp);
	
	AddDomSolid(pDom);
	AddDomLine(pDom);
	AddDomPolyLine(pDom);
	AddDomArc(pDom);
	AddDomText(pDom,RateTextHeight);

	m_Shapes.DetachDispatch();
}
*/

/*
CRect2 CXLControl::GetMTextBorder(const CPoint2& xy,const CString& sMText,TEXTSTYLE* pStyle,double THeight) const
{
	long nHoriAlign = pStyle->Horizontal;
	// MText Line 수 구하기
	CStringArray StrArr;
	CString cs = sMText;
	int pos;
	while( (pos=cs.Find("\r\n")) >= 0 )
	{
		StrArr.Add(cs.Left(pos));
		cs = cs.Mid(pos+2);
	}
	StrArr.Add(cs);	// Last String	
	long nMLinesu = StrArr.GetSize();

	// Max Length 수 구하기
	long nStrLen = 0;
	for(long n = 0; n < nMLinesu;n++) 
		if(nStrLen < (long) strlen(StrArr[n])) nStrLen = strlen(StrArr[n]); 

	CRect2 rResult;
	double Height = THeight * 1.4 * nMLinesu;
	double Width = THeight * nStrLen * 1.2;

	if( pStyle->Angle==90 || pStyle->Angle==270 )
		Height *= 1.33;
	if( nStrLen <= 1 )
		Width *= 1.33;

	if(nHoriAlign==TA_LEFT)
	{
		rResult.left = xy.x;
		rResult.top = xy.y - Height;
		rResult.right = xy.x + Width;
		rResult.bottom = xy.y;		
		if( pStyle->Angle==90 || pStyle->Angle==270 )
		{
			rResult.left = xy.x - Height*3/4;
			rResult.top = xy.y;
			rResult.right = xy.x + Height/2;
			rResult.bottom = xy.y + Width;
		}
	}
	else if(nHoriAlign==TA_CENTER)
	{
		rResult.left = xy.x - Width/2;
		rResult.top = xy.y - Height;
		rResult.right = xy.x + Width/2;
		rResult.bottom = xy.y;
		if( pStyle->Angle==90 || pStyle->Angle==270 )
		{
			rResult.left = xy.x - Height*3/4;
			rResult.top = xy.y - Width/2;
			rResult.right = xy.x + Height/2;
			rResult.bottom = xy.y + Width/2;
		}
	}
	else // if(nHoriAlign==TA_RIGHT)
	{
		rResult.left = xy.x - Width;
		rResult.top = xy.y - Height;
		rResult.right = xy.x;
		rResult.bottom = xy.y;
		if( pStyle->Angle==90 || pStyle->Angle==270 )
		{
			rResult.left = xy.x - Height*3/4;
			rResult.top = xy.y - Width;
			rResult.right = xy.x + Height/2;
			rResult.bottom = xy.y;
		}
	}

	return rResult;
}
*/

///////////////////////////////////////////////////////////////
// * 기 타 * //
///////////////////////////////////////////////////////////////	
void CXLControl::SetThreadUse(LPDISPATCH lpdisp)
{
	m_App.AttachDispatch(lpdisp);
	lpdisp = m_App.GetWorkbooks();	// Get the IDispatch pointer;
	ASSERT(lpdisp);

	m_Books.AttachDispatch(lpdisp);	// Attach the IDispatch pointer to the Books object

	lpdisp = m_Books.GetItem( COleVariant( (short)1) );
	ASSERT(lpdisp);

	m_Book.AttachDispatch(lpdisp);
	lpdisp = m_Book.GetSheets();
	ASSERT(lpdisp);

	m_Sheets.AttachDispatch(lpdisp);
	ASSERT(lpdisp);
	lpdisp = m_Sheets.GetItem( COleVariant((short)(m_nSheet)) );	// parameter --> m_Sheet index;

	ASSERT(lpdisp);
	m_Sheet.AttachDispatch(lpdisp);
}

BYTE* CXLControl::GetBitInfo(BYTE *pByte32, long nSou) const
{
	unsigned long bt = 1;

	for(long n = 0; n < 32; n++)
	{
		pByte32[n] = nSou & bt ? 1 : 0;
		bt <<= 1;
	}
	
	return pByte32;
}

// 다른 엑셀 파일 실행
void CXLControl::ExcelExecuteFile(CString strFileName)
{
	HWND hWnd = AfxGetMainWnd()->m_hWnd;
	ShellExecute(hWnd,"open",strFileName,NULL,NULL,SW_SHOWNORMAL);
}

// HeadTitle 실행
// Row만 설정하는 경우
void CXLControl::SetPrintTitleRows(long nRow1, long nRow2, long nSheetNumber/*=1*/)
{
	// Get Sheets
	LPDISPATCH lpDisp = m_Book.GetSheets();
	ASSERT(lpDisp);
	m_Sheets.AttachDispatch(lpDisp);
	// Get Sheet
	if (nSheetNumber <= 0)
		lpDisp = m_Sheets.GetItem(COleVariant((short)m_nSheet));
	else
		lpDisp = m_Sheets.GetItem(COleVariant((short)nSheetNumber));

	ASSERT(lpDisp);
	m_Sheet.AttachDispatch(lpDisp);
	m_Sheet.Activate();

	lpDisp = m_Sheet.GetPageSetup();

	CString strRows;
	strRows.Format(_T("$%d:$%d"), nRow1, nRow2);

	PageSetup page;
	page.AttachDispatch(lpDisp);
	page.SetPrintTitleRows(strRows);
	page.ReleaseDispatch();
	page.DetachDispatch();
}


CString CXLControl::GetVertStrCol(long nCol) const
{
	ASSERT(nCol > 0 && nCol <= 255);
	
	CString sCol(_T(""));
	if(nCol <= 26)
		sCol.Format(_T("%c"),'A' + nCol-1);
	//else if(nCol / 26 < 26)
	else if( nCol < 255 )
	{
		long h,f;
		h = nCol / 26 - 1;
		f = nCol % 26 - 1;
		sCol.Format(_T("%c%c"),'A'+h,'A' + f);
	}

	return sCol;
}

// Col로써 범위를 설정하는 경우
void CXLControl::SetPrintTitleCols(long nCol1, long nCol2, long nSheetNumber/*=-1*/)
{
	SetPrintTitleCols(GetVertStrCol(nCol1), GetVertStrCol(nCol1), nSheetNumber);
}

void CXLControl::SetPrintTitleCols(CString strCol1, CString strCol2, long nSheetNumber/*=-1*/)
{
	if( strCol1 == _T("")) return;

	// Get Sheets
	LPDISPATCH lpDisp = m_Book.GetSheets();
	ASSERT(lpDisp);
	m_Sheets.AttachDispatch(lpDisp);
	// Get Sheet
	if (nSheetNumber <= 0)
		lpDisp = m_Sheets.GetItem(COleVariant((short)m_nSheet));
	else
		lpDisp = m_Sheets.GetItem(COleVariant((short)nSheetNumber));
	ASSERT(lpDisp);
	m_Sheet.AttachDispatch(lpDisp);
	m_Sheet.Activate();
	
	lpDisp = m_Sheet.GetPageSetup();

	CString strCols;
	strCols.Format(_T("$%s:$%s"), strCol1, strCol2);

	PageSetup page;
	page.AttachDispatch(lpDisp);
	page.SetPrintTitleColumns(strCols);
	page.ReleaseDispatch();
	page.DetachDispatch();
}

// 틀고정
// bFreeze == FALSE 는 해제
// Sheet1에 3번째 행과 두번째 열까지 틀을 고정하고 싶을 때 
// XL.SetFreezePanes(3, 2, 1, TRUE);
void CXLControl::SetFreezePanes(long nRow, long nCol, long nSheetNumber/*=1*/, BOOL bFreeze/*=TRUE*/)
{
	// Get Sheets
	LPDISPATCH lpDisp = m_Book.GetSheets();
	ASSERT(lpDisp);
	m_Sheets.AttachDispatch(lpDisp);
	// Get Sheet
	lpDisp = m_Sheets.GetItem(COleVariant((short)nSheetNumber));
	ASSERT(lpDisp);
	m_Sheet.AttachDispatch(lpDisp);
	m_Sheet.Activate();

	CString strCell = GetCellStr(nRow, nCol);
	lpDisp = m_Sheet.GetRange(COleVariant(strCell), COleVariant(strCell));
	ASSERT(lpDisp);
	m_Range.AttachDispatch(lpDisp);
	m_Range.Select();

	lpDisp = m_App.GetActiveWindow();
	Window wind(lpDisp);
	wind.SetFreezePanes(bFreeze);
	wind.ReleaseDispatch();
	wind.DetachDispatch();
}

// 머리말 꼬리말 설정하기
// P : 쪽 번호
// N : 전체 쪽수
// D : 날짜
// T : 시간
// F : 파일이름
// A : 시트이름
// \n : 행 구분
// "예) 바닥글로 현재 쪽 번호, 전체 쪽수 그리고 한길정보통신을 다음칸에 인쇄"
// nCenter => Left = 0, center = 1, right = 2;
// nCheetNumber => 설정할 시트의 번호
// strHeader = _T("&P/&N \n한길정보통신"); bHeader=FASLE, nCenter=1;
// XL.SetHeader(strHeader, nCheetNumber, nCenter);
void CXLControl::SetHeader(CString strHeader, long nSheetNumber/*=1*/, long nCenter/*=1*/)
{
	// Get Sheets
	LPDISPATCH lpDisp = m_Book.GetSheets();
	ASSERT(lpDisp);
	m_Sheets.AttachDispatch(lpDisp);
	// Get Sheet
	lpDisp = m_Sheets.GetItem(COleVariant((short)nSheetNumber));
	ASSERT(lpDisp);
	m_Sheet.AttachDispatch(lpDisp);
	m_Sheet.Activate();
	
	lpDisp = m_Sheet.GetPageSetup();

	PageSetup page;
	page.AttachDispatch(lpDisp);
	switch(nCenter)
	{
	case 0:
		page.SetLeftHeader(strHeader);
		break;
	case 1:
		page.SetCenterHeader(strHeader);
		break;
	case 2:
		page.SetRightHeader(strHeader);
		break;
	}
	page.ReleaseDispatch();
	page.DetachDispatch();
}

// XL.SetFoorter(strHeader, nCheetNumber, nCenter);
void CXLControl::SetFooter(CString strFooter, long nSheetNumber/*=1*/, long nCenter/*=1*/)
{
	// Get Sheets
	LPDISPATCH lpDisp = m_Book.GetSheets();
	ASSERT(lpDisp);
	m_Sheets.AttachDispatch(lpDisp);
	// Get Sheet
	lpDisp = m_Sheets.GetItem(COleVariant((short)nSheetNumber));
	ASSERT(lpDisp);
	m_Sheet.AttachDispatch(lpDisp);
	m_Sheet.Activate();
	
	lpDisp = m_Sheet.GetPageSetup();

	PageSetup page;
	page.AttachDispatch(lpDisp);
	switch(nCenter)
	{
	case 0:
		page.SetLeftFooter(strFooter);
		break;
	case 1:
		page.SetCenterFooter(strFooter);
		break;
	case 2:
		page.SetRightFooter(strFooter);
		break;
	}
	page.ReleaseDispatch();
	page.DetachDispatch();
}
 
// Page에 쪽을 설정하기
// ⓐ SetActiveSheet(nSheet);
// ⓑ SetPageBreak(5,5) or SetPageBreak("F6");
// 쪽 구분선이 'F6'를 기준으로 바로 위에 나타난다.
void CXLControl::SetPageBreak(long nRow, long nCol)
{
	CString strCell = GetCellStr(nRow, nCol);
	SetPageBreak(strCell);
}

void CXLControl::SetPageBreak(CString strCell)
{
	m_Sheet.Activate();
	LPDISPATCH lpDisp = m_Sheet.GetRange(COleVariant(strCell), COleVariant(strCell));
	ASSERT(lpDisp);

	HPageBreaks hp(m_Sheet.GetHPageBreaks());
	hp.Add(lpDisp);

	VPageBreaks vp(m_Sheet.GetVPageBreaks());
	vp.Add(lpDisp);
}
////////////////////////////////////////////////////////////
void CXLControl::SetCellColor(CString stt, CString end, long nColNum)
{
	Range resizedrange;
	resizedrange = m_Sheet.GetRange(COleVariant(stt), COleVariant(end));
	m_Interior = resizedrange.GetInterior();
	m_Interior.SetColorIndex(COleVariant((short)nColNum));
}

void CXLControl::SetCellColor(long nRow1, long nCol1, long nRow2, long nCol2, long nColNum)
{
	CString sStt = GetCellStr(nRow1,nCol1);
	CString sEnd = GetCellStr(nRow2,nCol2);

	SetCellColor(sStt, sEnd, nColNum);
}

/* CMS
void CXLControl::CopyPicture(CString sXLPicPath, CString sPicName, long nRowDes, long nColDes)
{
	CXLControl xl(FALSE);
	xl.OpenXL(sXLPicPath);
	xl.SetActiveSheet("Pic1");

	// 임시값
	long nRow1 = 7;
	long nRow2 = 14;
	long nCol1 = 0;
	long nCol2 = 8;

	CString stt = GetCellStr(nRow1, nCol1);
	CString end = GetCellStr(nRow2, nCol2);
	CString Des = GetCellStr(nRowDes, nColDes);

	Range orgRange;
	orgRange = xl.m_Sheet.GetRange(COleVariant(stt), COleVariant(end));

	long xlPrinter = 2;
	long xlPicture = -4147;
	orgRange.CopyPicture(xlPrinter, xlPicture);

	LPDISPATCH lpDes = m_Sheet.GetRange(COleVariant(Des), COleVariant(Des));

    VARIANT v;
	v.vt = VT_DISPATCH;
	v.pdispVal = lpDes; 

	m_Sheet.Paste(v, COleVariant((short)FALSE));
	xl.QuitXL();
	
}

void CXLControl::CopyPicture(CXLControl *xl, CString sPicName, CString Des)
{
//	CXLControl xl(FALSE);
//	xl.OpenXL(sXLPicPath);
//	xl.SetActiveSheet("Pic1");

	if(!xl->m_App.m_lpDispatch) return;

	VARIANT range;
	range = xl->m_Sheet.Evaluate(COleVariant(sPicName));

	LPDISPATCH lpDisp = range.pdispVal;
	Range orgRange(lpDisp);

	long xlPrinter = 2;
	long xlPicture = -4147;
	orgRange.CopyPicture(xlPrinter, xlPicture);

	LPDISPATCH lpDes = m_Sheet.GetRange(COleVariant(Des), COleVariant(Des));

    VARIANT v;
	v.vt = VT_DISPATCH;
	v.pdispVal = lpDes; 

	m_Sheet.Paste(v, COleVariant((short)FALSE));
//	xl.QuitXL();

}

void CXLControl::CopyPicture(CXLControl &xl, CString sPicName, CString Des)
{
//	CXLControl xl(FALSE);
//	xl.OpenXL(sXLPicPath);
//	xl.SetActiveSheet("Pic1");

	if(!xl.m_App.m_lpDispatch) return;

	VARIANT range;
	range = xl.m_Sheet.Evaluate(COleVariant(sPicName));

	LPDISPATCH lpDisp = range.pdispVal;
	Range orgRange(lpDisp);

	long xlPrinter = 2;
	long xlPicture = -4147;
	orgRange.CopyPicture(xlPrinter, xlPicture);

	LPDISPATCH lpDes = m_Sheet.GetRange(COleVariant(Des), COleVariant(Des));

    VARIANT v;
	v.vt = VT_DISPATCH;
	v.pdispVal = lpDes; 

	m_Sheet.Paste(v, COleVariant((short)FALSE));
//	xl.QuitXL();

}
*/

/*
void CXLControl::CopyPicture(CString sXLPicPath, CString sPicName, CString Des)
{
	CXLControl xl(FALSE);
	xl.OpenXL(sXLPicPath);
//	xl.SetActiveSheet("Pic1");

	VARIANT range;
	range = xl.m_Sheet.Evaluate(COleVariant(sPicName));

	LPDISPATCH lpDisp = range.pdispVal;
	Range orgRange(lpDisp);

	long xlPrinter = 2;
	long xlPicture = -4147;
	orgRange.CopyPicture(xlPrinter, xlPicture);

	LPDISPATCH lpDes = m_Sheet.GetRange(COleVariant(Des), COleVariant(Des));

    VARIANT v;
	v.vt = VT_DISPATCH;
	v.pdispVal = lpDes; 

	m_Sheet.Paste(v, COleVariant((short)FALSE));
	xl.QuitXL();

}


void CXLControl::CopyPicture(CString sXLPicPath, long nRow1, long nCol1, long nRow2, long nCol2, long nRowDes, long nColDes, long nType) 
{
	CXLControl xl(FALSE);
	xl.OpenXL(sXLPicPath);
	xl.SetActiveSheet("Pic1");

	CString stt = GetCellStr(nRow1, nCol1);
	CString end = GetCellStr(nRow2, nCol2);
	CString Des = GetCellStr(nRowDes, nColDes);

	Range orgRange;
	orgRange = xl.m_Sheet.GetRange(COleVariant(stt), COleVariant(end));

	long xlPrinter = 2;
	long xlPicture = -4147;
	orgRange.CopyPicture(xlPrinter, xlPicture);

	LPDISPATCH lpDes = m_Sheet.GetRange(COleVariant(Des), COleVariant(Des));

    VARIANT v;
	v.vt = VT_DISPATCH;
	v.pdispVal = lpDes; 

	m_Sheet.Paste(v, COleVariant((short)FALSE));
	xl.QuitXL();
}

void CXLControl::CopyPicture(CBitmap *pBitmap, double dLeft, double dTop, double dWidth, double dHeight )
{
	CImage Image;
	CString szFileName = _T("C:\\_instemp.jpg";
	
	// 이미지 작성
	if(!Image.CopyFromBmp(pBitmap))
	{
		AfxMessageBox("지원할수 없는 형식입니다 !");		
		return;
	}	
	Image.SaveImage(szFileName.GetBuffer(szFileName.GetLength()));		

	LPDISPATCH lpDisp = m_Sheet.GetShapes();
	m_Shapes.AttachDispatch(lpDisp);	
	lpDisp = m_Shapes.AddPicture(szFileName,TRUE, TRUE,(float)dLeft, (float)dTop, (float) dWidth, (float) dHeight);			
	Shape shp(lpDisp);
	COleVariant dDumy = (long)0;	
	if(dWidth==0)  shp.ScaleWidth(1.0,TRUE,dDumy);
	if(dHeight==0) shp.ScaleHeight(1.0,TRUE,dDumy);	
	m_Shapes.DetachDispatch();
	CFile::Remove(szFileName);
}
*/

void CXLControl::CopyPicture(CString sPictureSheetName, long nRowStt, long nColStt, long nRowEnd, long nColEnd, long nRowTarget, long nColTarget)
{
	CString sheetname = this->m_Sheet.GetName();
	this->SetActiveSheet(sPictureSheetName);

	CString sttCell = GetCellStr(nRowStt, nColStt);
	CString endCell = GetCellStr(nRowEnd, nColEnd);

	Range sourceRange;
	sourceRange = this->m_Sheet.GetRange(COleVariant(sttCell), COleVariant(endCell));
//	long xlPrinter = 2;
//	long xlPicture = -4147;
	sourceRange.CopyPicture(xlPrinter, xlPicture);

	this->SetActiveSheet(sheetname);
	CString target = GetCellStr(nRowTarget, nColTarget);
	LPDISPATCH lpTarget = m_Sheet.GetRange(COleVariant(target), COleVariant(target));

    VARIANT vtarget;
	vtarget.vt = VT_DISPATCH;
	vtarget.pdispVal = lpTarget; 

	this->m_Sheet.Paste(vtarget, COleVariant((short)FALSE));
}

void CXLControl::InsertPicture(CString sPath, double dLeft, double dTop, double dWidth, double dHeight, BOOL bLockAspectRatio)
{
	LPDISPATCH lpDisp = m_Sheet.GetShapes();
	m_Shapes.AttachDispatch(lpDisp);
	
	if (sPath != _T(""))
	{
		CFileFind finder;
        BOOL bResult = finder.FindFile(sPath);
		if (bResult)
		{
			lpDisp = m_Shapes.AddPicture(sPath, FALSE, TRUE, (float)dLeft,(float)dTop,(float)dWidth,(float)dHeight);

			Shape shp(lpDisp);
			COleVariant dDumy = (long)0;

			shp.ScaleWidth(1.0, TRUE, dDumy);
			shp.ScaleHeight(1.0, TRUE, dDumy);
			if (dWidth == 0 || dHeight == 0)
			{
				m_Shapes.ReleaseDispatch();
				m_Shapes.DetachDispatch();
				return;
			}

			double w = shp.GetWidth();
			double h = shp.GetHeight();

//			shp.SetLockAspectRatio((bLockAspectRatio ? (long)-1 : (long)0));  
			shp.SetLockAspectRatio(FALSE);

			if (bLockAspectRatio)
			{
				float dScaleW = float(dWidth/w);
				float dScaleH = float(dHeight/h);
				float dScale = dScaleW < dScaleH ? dScaleW : dScaleH;

				shp.ScaleWidth((float)(dScale),TRUE,dDumy);
				shp.ScaleHeight((float)(dScale),TRUE,dDumy);

				w = shp.GetWidth();
				h = shp.GetHeight();

				float dLeftMove = float(dLeft + (dWidth - w) / 2.0);
				float dTopMove = float(dTop + (dHeight - h) / 2.0);

				shp.SetLeft(dLeftMove);
				shp.SetTop(dTopMove);
			}
			else
			{
				if(w > dWidth)
					shp.ScaleWidth((float)(dWidth/w),TRUE,dDumy);
				if(h > dHeight )
					shp.ScaleHeight((float)(dHeight/h),TRUE,dDumy);
			}
		}
		else
		{
			CString str;
			str.Format(_T("\"%s\"\n파일을 찾을 수 없습니다."), sPath);
			AfxMessageBox(str, MB_OK);
		}
	}
//	else
		//AfxMessageBox("파일을 찾을 수 없습니다 1.", MB_OK);

	m_Shapes.ReleaseDispatch();
	m_Shapes.DetachDispatch();
}

void CXLControl::InsertPictureRowCol(CString sPath, long nRowStt, long nColStt, long nRowEnd, long nColEnd, BOOL bLockAspectRatio)
{
	CString rangeStt = GetCellStr(nRowStt, nColStt);
	CString rangeEnd = GetCellStr(nRowEnd, nColEnd);
	if (nRowEnd==0 ||  nColEnd == 0)
		rangeEnd = GetCellStr(nRowStt, nColStt);

	Range oRange;
	oRange = m_Sheet.GetRange(COleVariant(rangeStt), COleVariant(rangeEnd));

	COleVariant VLeft, VTop, VWidth, VHeight;
	VLeft = oRange.GetLeft();
	VTop = oRange.GetTop();
	VWidth = oRange.GetWidth();
	VHeight = oRange.GetHeight();

	if (nRowEnd==0 ||  nColEnd == 0)
	{
		InsertPicture(sPath, VLeft.dblVal, VTop.dblVal, 0, 0, bLockAspectRatio);
	}
	else
	{
		InsertPicture(sPath, VLeft.dblVal, VTop.dblVal, VWidth.dblVal, VHeight.dblVal, bLockAspectRatio);
	}
}


void CXLControl::SetDisplayGridLine(BOOL bGridLine)
{
	Window wd;
	LPDISPATCH lpDisp = m_App.GetActiveWindow();
	wd.AttachDispatch(lpDisp);
	wd.SetDisplayGridlines(bGridLine);
	wd.ReleaseDispatch();
	wd.DetachDispatch();
}

void CXLControl::SetUserControl(BOOL bUserCtrl)
{
	m_App.SetUserControl(bUserCtrl);
}

// 화면 배율 
void CXLControl::SetViewZoom(long nSize)
{
	LPDISPATCH lpDisp = m_App.GetActiveWindow();
	Window wnd(lpDisp);
	COleVariant ZoomSize = nSize;
	wnd.SetZoom(ZoomSize);
	wnd.ReleaseDispatch();
	wnd.DetachDispatch();
}

// 페이지 여백 주기
void CXLControl::SetPrintMargin(double dLeft, double dRight, double dTop, double dBottom)
{
	LPDISPATCH lpDisp = m_Sheet.GetPageSetup();
	PageSetup page;

	page.AttachDispatch(lpDisp);

	page.SetLeftMargin(m_App.InchesToPoints(dLeft*0.393700787401575));
	page.SetRightMargin(m_App.InchesToPoints(dRight*0.393700787401575));
	page.SetTopMargin(m_App.InchesToPoints(dTop*0.393700787401575));
	page.SetBottomMargin(m_App.InchesToPoints(dBottom*0.393700787401575));

	page.ReleaseDispatch();
	page.DetachDispatch();
}

// 페이지의 가운데로
void CXLControl::SetPrintCenterHorizon(BOOL bCenterHor)
{
	LPDISPATCH lpDisp = m_Sheet.GetPageSetup();
	PageSetup page;

	page.AttachDispatch(lpDisp);

	page.SetCenterHorizontally(bCenterHor);

	page.ReleaseDispatch();
	page.DetachDispatch();
}

void CXLControl::SetPrintCenterVertical(BOOL bCenterVer)
{
	LPDISPATCH lpDisp = m_Sheet.GetPageSetup();
	PageSetup page;

	page.AttachDispatch(lpDisp);

	page.SetCenterVertically(bCenterVer);

	page.ReleaseDispatch();
	page.DetachDispatch();
}

// 페이지 배율
void CXLControl::SetPringZoom(long nSize)
{
	LPDISPATCH lpDisp = m_Sheet.GetPageSetup();
	PageSetup page(lpDisp);
	COleVariant ZoomSize = nSize;

//	page.AttachDispatch(lpDisp);

	page.SetZoom(ZoomSize);

	page.ReleaseDispatch();
	page.DetachDispatch();
}


// 인쇄 영역 설정
void CXLControl::SetPrintArea(LPCTSTR lpstr1, LPCTSTR lpstr2)
{
	LPDISPATCH lpDisp = m_Sheet.GetPageSetup();
	PageSetup page;

	page.AttachDispatch(lpDisp);

	CString str;
	str.Format(_T("$%s:$%s"), lpstr1, lpstr2);
	
	page.SetPrintArea((LPCTSTR)str);

	page.ReleaseDispatch();
	page.DetachDispatch();
}

// 인쇄 영역 가져오기
CString CXLControl::GetPrintArea()
{
	LPDISPATCH lpDisp = m_Sheet.GetPageSetup();
	PageSetup page;

	CString str;

	page.AttachDispatch(lpDisp);
	str = page.GetPrintArea();

	page.ReleaseDispatch();
	page.DetachDispatch();

	return str;
}

void CXLControl::SetOrientation(BOOL bLandScape)
{
	LPDISPATCH lpDisp = m_Sheet.GetPageSetup();
	PageSetup page;

	page.AttachDispatch(lpDisp);

	//long xlLandscape = 2;
	//long xlPortrait = 1;

	if (bLandScape)
		page.SetOrientation(xlLandscape);
	else
		page.SetOrientation(xlPortrait);

	page.ReleaseDispatch();
	page.DetachDispatch();
}

/*
    Range("Z24:AA24").Select
    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
*/
void CXLControl::GetXLApp(LPDISPATCH lpDisp)
{
	m_Book.AttachDispatch(lpDisp);
	lpDisp = m_Book.GetApplication();

	m_App.AttachDispatch(lpDisp);

	lpDisp = m_App.GetWorkbooks();
	m_Books.AttachDispatch(lpDisp);

	lpDisp = m_Book.GetSheets();
	ASSERT(lpDisp);
	m_Sheets.AttachDispatch(lpDisp);
}
void CXLControl::GetXLApp(_Application	App, CString strFileName)
{
	m_App = App;
	
	LPDISPATCH lpDisp = m_App.GetWorkbooks();	// Get the IDispatch pointer;
	ASSERT(lpDisp);
	m_Books.AttachDispatch(lpDisp);

	CString strBook = _T("");
	for (int idx = strFileName.GetLength()-1; idx >= 0; idx--)
	{
		char ch = strFileName.GetAt(idx);
		if (ch == '\\' )
			break;
		strBook.Insert(0, ch);
	}
	
	int nCount = m_Books.GetCount();
	for (int idx = 1; idx <= nCount; idx++)
	{
		lpDisp = m_Books.GetItem( COleVariant((short)(idx)) );
		ASSERT(lpDisp);
		m_Book.AttachDispatch(lpDisp);
		CString strBookName = m_Book.GetName();
		if (strBook == strBookName)
			break;
	}

	lpDisp = m_Book.GetSheets();
	ASSERT(lpDisp);
	m_Sheets.AttachDispatch(lpDisp);
//	m_App.SetVisible(FALSE);
//	m_App.SetUserControl(TRUE);
//	m_App.SetDisplayAlerts(FALSE); // Excel 종료시 저장 여부를 물어보지 않는다.
}

void CXLControl::CheckExcelPre()
{
	HANDLE         hProcessSnap = NULL; 
	DWORD          Return       = FALSE; 
	PROCESSENTRY32 pe32         = {0};
	hProcessSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0); 
	if (hProcessSnap != INVALID_HANDLE_VALUE) 
	{
		pe32.dwSize = sizeof(PROCESSENTRY32);
		if (Process32First(hProcessSnap, &pe32))
		{ 
			DWORD Code = 0;
			DWORD dwPriorityClass; 
			do 
			{ 
				Sleep(10);
				HANDLE hProcess = NULL; 
				hProcess = OpenProcess (PROCESS_ALL_ACCESS,  FALSE, pe32.th32ProcessID); 
				dwPriorityClass = GetPriorityClass (hProcess); 
				char * Temp = strupr(pe32.szExeFile);
				if (strstr(Temp, _T("EXCEL.EXE")) != 0 )
					m_HandleArr.Add(pe32.th32ProcessID);
				CloseHandle (hProcess);
			} 
			while (Process32Next(hProcessSnap, &pe32));
		} 
		CloseHandle(hProcessSnap);
	}
}

void CXLControl::TerminateExcel()
{
	HANDLE         hProcessSnap = NULL; 
	DWORD          Return       = FALSE; 
	PROCESSENTRY32 pe32         = {0};
	hProcessSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0); 
	if (hProcessSnap != INVALID_HANDLE_VALUE) 
	{
		pe32.dwSize = sizeof(PROCESSENTRY32);
		if (Process32First(hProcessSnap, &pe32))
		{ 
			DWORD Code = 0;
			DWORD dwPriorityClass; 
			do 
			{ 
				Sleep(10);
				HANDLE hProcess = NULL; 
				hProcess = OpenProcess (PROCESS_ALL_ACCESS,  FALSE, pe32.th32ProcessID); 
				dwPriorityClass = GetPriorityClass (hProcess); 
				char * Temp = strupr(pe32.szExeFile);
				if (strstr(Temp, _T("EXCEL.EXE")) != 0 )
				{
					int nCount = (int)m_HandleArr.GetCount();
					BOOL bKillProcess = TRUE;
					for (int i = 0; i < nCount; i++)
					{
						int hProcessOld = m_HandleArr.GetAt(i);
						if (hProcessOld == pe32.th32ProcessID)
						{	
							bKillProcess = FALSE;
							break;
						}
					}
					if (bKillProcess)
					{
						GetExitCodeProcess(hProcess, &Code);
						if (TerminateProcess(hProcess, Code))
						{
							Code =  WaitForSingleObject(hProcess, INFINITE);
							Return = TRUE;
						}
						else
						{
							break;
						}
					}
				}
				CloseHandle (hProcess);
			} 
			while (Process32Next(hProcessSnap, &pe32));
		} 
		CloseHandle(hProcessSnap);
	}
}
//0 ?
//1 ?
//2 ? 
//3 ?
//4 ?
//5 SaveAs
//6 ?
//page Setup 7
//print 8
//9 ?
//10 
//11
//12
//13
//14

void CXLControl::PrintOutDlg()
{
	Dialogs dialogs; 
	LPDISPATCH lpDisp = lpDisp = m_App.GetDialogs();
	dialogs.AttachDispatch(lpDisp);

	lpDisp = dialogs.GetItem(8);

	Dialog dlg;
	dlg.AttachDispatch(lpDisp);
	dlg.Show(m_covOptional ,m_covOptional,m_covOptional,m_covOptional,m_covOptional,
		m_covOptional,m_covOptional,m_covOptional,m_covOptional,m_covOptional,
		m_covOptional,m_covOptional,m_covOptional,m_covOptional,m_covOptional,
		m_covOptional,m_covOptional,m_covOptional,m_covOptional,m_covOptional,
		m_covOptional,m_covOptional,m_covOptional,m_covOptional,m_covOptional,
		m_covOptional,m_covOptional,m_covOptional,m_covOptional,m_covOptional);
}

void CXLControl::PrintPewView()
{
	COleVariant  covTrue((short)TRUE);

	LPDISPATCH lpDisp = m_App.GetSheets();
	m_Sheets.AttachDispatch(lpDisp); 

	m_Sheet.PrintPreview(covTrue); 
}
void CXLControl::PageSetupDlg()
{
	Dialogs dialogs; 
	LPDISPATCH lpDisp = m_App.GetDialogs();
	dialogs.AttachDispatch(lpDisp);

	lpDisp = dialogs.GetItem(7);

	VARIANT v;
	v.vt = VT_DISPATCH;
	v.pdispVal = lpDisp;

	Dialog dlg;
	dlg.AttachDispatch(lpDisp);
	dlg.Show(m_covOptional ,m_covOptional,m_covOptional,m_covOptional,m_covOptional,
		m_covOptional,m_covOptional,m_covOptional,m_covOptional,m_covOptional,
		m_covOptional,m_covOptional,m_covOptional,m_covOptional,m_covOptional,
		m_covOptional,m_covOptional,m_covOptional,m_covOptional,m_covOptional,
		m_covOptional,m_covOptional,m_covOptional,m_covOptional,m_covOptional,
		m_covOptional,m_covOptional,m_covOptional,m_covOptional,m_covOptional);
}

_MITC_BASIC_END