#include "StdAfx.h"
#include "MsWordTool.h"
#include "math.h"

_MITC_BASIC_BEGIN

CTextFormat::CTextFormat()
{
  strContent = L"";
  bBold = false;
  nStep = 1;
  nFontSize = 10.5;
  FontName = L"宋体";
  FontColor = RGB(0,0,0);
  IsHaveFisrtSigle = true;
  LeftIndent = 0;
  iOutLineLevel = 10;

  IsInsertPic = FALSE;     //是否为插入图片
  strPicPath = L"";    // IsInsertPic = TRUE 插入图片的路径
  strPicExpress = L"";// IsInsertPic = TRUE 插入图片的描述，通常写在图片下方

}

CTextFormat::~CTextFormat()
{

}
CMsWordTool::CMsWordTool(void)
{
}

CMsWordTool::~CMsWordTool(void)
{
	ExitWordApp();
}

CMsWordTool::CMsWordTool(CMsWordTooldecorator* pMsWordTooldecorator,CString strTemplatePath)
{
	_ASSERTE(pMsWordTooldecorator);
	m_pMsWordTooldecorator=pMsWordTooldecorator;
	COleVariant vTrue((short)TRUE),vFalse((short)FALSE),vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	 _Application* pApplication=pMsWordTooldecorator->GetApp();
  Documents m_pDocs = (Documents)pMsWordTooldecorator->GetApp()->GetDocuments();
  pMsWordTooldecorator->SetDocuments(&m_pDocs);
	_Document myDoc= m_pDocs.Add(COleVariant(strTemplatePath),vOpt,vOpt,vOpt);	
  
	m_pDoc = new _Document(myDoc);
	InitializeWordApp();

  m_pDocSection = new Selection;
  m_pDocSection->m_bAutoRelease = TRUE;

  m_pCreatePic = new CCreatePic;

}

void CMsWordTool::InitializeWordApp()
{
   //初始化

  Window wordWindow = m_pDoc->GetActiveWindow();
  wordWindow.SetDocumentMap(TRUE);
  Pane wordPane = (Pane)wordWindow.GetActivePane();
  View wordView = (View)wordPane.GetView();
  Zoom wordZoom = (Zoom)wordView.GetZoom();
  wordZoom.SetPercentage(100);

  //利用替换功能
  m_pRange=new Range(m_pDoc->GetContent()); 
  m_pFndInDoc=new Find(m_pRange->GetFind()); 
  m_pFndInDoc->ClearFormatting(); 	
  m_pRpInDoc=new Replacement(m_pFndInDoc->GetReplacement()); 
  m_pRpInDoc->ClearFormatting(); 
}

void CMsWordTool::ReplaceParameter(CString src,CString data)
{
  CString replaceStr = src;//被替换
  CString replaceStrWith = data;//替换

  COleVariant Text(replaceStr); //被替换
  COleVariant MatchCase((short)FALSE); 
  COleVariant MatchWholeWord((short)FALSE); 
  COleVariant MatchWildcards((short)FALSE); 
  COleVariant MatchSoundsLike((short)FALSE); 
  COleVariant MatchAllWordForms((short)FALSE); 
  COleVariant Forward((short)TRUE); 
  COleVariant Wrap((short)1);//用msgbox(wdFindContinue)得到 
  COleVariant format((short)FALSE); 
  COleVariant ReplaceWith=(replaceStrWith);//替换 
  COleVariant Replace((short)2);//用msgbox(wdReplaceAll)得到   //替换所有
  COleVariant MatchKashida=((short)FALSE); //以下四个参数默认false
  COleVariant MatchDiacritics=((short)FALSE); 
  COleVariant MatchAlefHamza=((short)FALSE); 
  COleVariant MatchControl=((short)FALSE);

  m_pFndInDoc->Execute(&Text, &MatchCase, &MatchWholeWord, &MatchWildcards, 
    &MatchSoundsLike, &MatchAllWordForms, &Forward, &Wrap, 
    &format, &ReplaceWith, &Replace, &MatchKashida, 
    &MatchDiacritics, &MatchAlefHamza, &MatchControl);
}

void CMsWordTool::InsertPicture(CString strbookmark,CString path)
{

  CFileStatus   status; 
  if(!CFile::GetStatus(path,status)) 
  { 
    CString msg = L"";
    msg.Format( L"\n%s   does   not   exist! \n",path); 
    OutputDebugString(msg)   ; 
    return ;
  } 
  if(!GetSelection(strbookmark))
    return;


  InlineShapes shaps = m_pDocSection->GetInlineShapes();
  

  COleVariant link((short)FALSE);
  COleVariant savew((short)TRUE);
  COleVariant  vOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

  shaps.AddPicture(path,&link,&savew,&vOptional);
}
void CMsWordTool::InsertPicture(CString path)
{
  CFileStatus   status; 
  if(!CFile::GetStatus(path,status)) 
  { 
    CString msg = L"";
    msg.Format( L"\n%s   does   not   exist! \n",path); 
    OutputDebugString(msg)   ; 
  } 
  else
  {
    InlineShapes shaps = m_pDocSection->GetInlineShapes();

    COleVariant link((short)FALSE);
    COleVariant savew((short)TRUE);
    COleVariant  vOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

    shaps.AddPicture(path,&link,&savew,&vOptional);

  }
}
void CMsWordTool::FillvctTable(CString strBookMark,const std::vector<CString>& ContentList,int nColumns)
{
  if(ContentList.size()<=0)
  {
    return;
  }
  COleVariant vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
  Selection wdSelect;
  wdSelect.AttachDispatch(m_pMsWordTooldecorator->GetApp()->GetSelection());
  COleVariant vName = COleVariant(strBookMark, VT_BSTR);

  try
  {
    wdSelect.GoTo(COleVariant((short)-1),vOpt,vOpt,vName);
  }
  catch(CException* e)
  {
    return ;
  }
  

  int nRows = ContentList.size()/nColumns;

//   for (int i=0;i<nRows-1;i++)
//   {
//     wdSelect.MoveRight(COleVariant((short)1),COleVariant((short)nColumns),COleVariant((short)0));
//     wdSelect.InsertRowsBelow(COleVariant((short)1));
//     wdSelect.Collapse(COleVariant((short)1));
//   }
//  wdSelect.MoveUp(COleVariant((short)5),COleVariant((short)(nRows-1)),COleVariant((short)0));
  for (int i=0;i<(int)ContentList.size();i++)
  {
         if(i>0 && !((i)%(nColumns) == 0))
         {
           wdSelect.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)0));
         }

    wdSelect.TypeText(ContentList.at(i));
//     if(i != ContentList.size() -1)
//     {
//       wdSelect.MoveRight(COleVariant((short)12),COleVariant((short)1),COleVariant((short)0));
//     }

         if((i+1)%(nColumns) == 0)
         {
           if(i+1 != (int)ContentList.size())
           {
             wdSelect.InsertRowsBelow(COleVariant((short)1));
             wdSelect.Collapse(COleVariant((short)1));
           }
         }
  }

}
void CMsWordTool::FillTable(CString strBookMark,CArray<CString,CString>& ContentList,int nColumns)
{
  if(ContentList.GetSize()<=0)
  {
    return;
  }


  COleVariant vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
  Selection wdSelect;
  wdSelect.AttachDispatch(m_pMsWordTooldecorator->GetApp()->GetSelection());
  COleVariant vName = COleVariant(strBookMark, VT_BSTR);

  try
  {
    wdSelect.GoTo(COleVariant((short)-1),vOpt,vOpt,vName);
  }
  catch(CException* e)
  {
    return ;
  }


  int nRows = ContentList.GetSize()/nColumns;

  for (int i=0;i<nRows-1;i++)
  {
    wdSelect.MoveRight(COleVariant((short)1),COleVariant((short)nColumns),COleVariant((short)0));
    wdSelect.InsertRowsBelow(COleVariant((short)1));
    wdSelect.Collapse(COleVariant((short)1));
  }
  wdSelect.MoveUp(COleVariant((short)5),COleVariant((short)(nRows-1)),COleVariant((short)0));
  for (int i=0;i<ContentList.GetSize();i++)
  {
//     if(i>0 && !((i)%(nColumns) == 0))
//     {
//       wdSelect.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)0));
//     }

    wdSelect.TypeText(ContentList.GetAt(i));
    wdSelect.MoveRight(COleVariant((short)12),COleVariant((short)1),COleVariant((short)0));

//     if((i+1)%(nColumns) == 0)
//     {
//       if(i+1 != ContentList.GetSize())
//       {
//         wdSelect.InsertRowsBelow(COleVariant((short)1));
//         wdSelect.Collapse(COleVariant((short)1));
//       }
//     }
  }
}
void CMsWordTool::UpdateAllDomainCode()
{

  m_pDocSection->WholeStory();
//   _Font   font;   
//   font   =   m_pDocSection->GetFont();   
// 
//   font.SetNameFarEast(_T("宋体"));   
//   font.SetNameAscii(L"Times New Roman");   
//   font.SetNameOther(L"Times New Roman");  


  Fields fs;
  fs.AttachDispatch(m_pDocSection->GetFields());
  fs.Update();
}

void CMsWordTool::SaveAs(CString SaveAsPath /*= NULL*/,CString SaveAsNewDocName/* = NULL*/)
{
  if(SaveAsNewDocName.IsEmpty() || SaveAsPath.IsEmpty())
  {
    return ;
  }
  COleVariant vName = COleVariant(SaveAsNewDocName, VT_BSTR);
 
  m_pMsWordTooldecorator->GetApp()->ChangeFileOpenDirectory(SaveAsPath);
  COleVariant FileFormat((short)0);
  COleVariant LockComments((short)FALSE);
  CString strpassword = _T("");
  COleVariant Password(strpassword);
  COleVariant AddToRecentFiles((short)TRUE);
  COleVariant WritePassword(strpassword);
  COleVariant ReadOnlyRecommended((short)FALSE);
  COleVariant EmbedTrueTypeFonts((short)FALSE);
  COleVariant SaveNativePictureFormat((short)FALSE);
  COleVariant SaveFormsData((short)FALSE);
  COleVariant SaveAsAOCELetter((short)FALSE);
  try
  {
    m_pDoc->SaveAs(&vName, &FileFormat, &LockComments, &Password, &AddToRecentFiles, &WritePassword, &ReadOnlyRecommended, &EmbedTrueTypeFonts, &SaveNativePictureFormat, &SaveFormsData, &SaveAsAOCELetter);
  }
  catch (CException* e)
  {
    AfxMessageBox(L"请检查该报告文件是否已打开！",MB_ICONERROR);
    return;
  }
}

void CMsWordTool::OpenDoc(LPCTSTR SaveAsPath /*= NULL*/,LPCTSTR SaveAsNewDocName /*= NULL*/)
{
  m_pMsWordTooldecorator->GetApp()->ChangeFileOpenDirectory(SaveAsPath);
  m_pMsWordTooldecorator->GetApp()->SetVisible(TRUE);
  
  COleVariant ConfirmConversions((short)FALSE);
  COleVariant ReadOnly((short)FALSE);
  COleVariant AddToRecentFiles((short)FALSE);
  CString strNull = _T("");
  COleVariant PasswordDocument(strNull);
  COleVariant PasswordTemplate(strNull);
  COleVariant Revert((short)FALSE);

  COleVariant WritePasswordDocument(strNull);
  COleVariant WritePasswordTemplate(strNull);
  COleVariant Format((short)0);
  COleVariant XMLTransform(strNull);
  COleVariant Visiable((short)FALSE);
  COleVariant vName = COleVariant(SaveAsNewDocName, VT_BSTR);
  Documents m_pDocs = (Documents)m_pMsWordTooldecorator->GetApp()->GetDocuments();
  m_pDocs.Open(&vName,&ConfirmConversions,&ReadOnly,&AddToRecentFiles,&PasswordDocument,&PasswordTemplate,&Revert,&WritePasswordDocument,&WritePasswordTemplate,&Format,&XMLTransform,&Visiable);
}
void CMsWordTool::Delete4BookMarks(CString bm1,CString bm2)
{
  Selection wdSel;
  if(!GetSelection4bm(wdSel,bm1,bm2))
    return ;
  //内部参数
  COleVariant un((short)1);
  COleVariant co((short)1);
  wdSel.Delete(&un,&co);
}
void CMsWordTool::ExitWordApp()
{
  //数据清空
  m_pRange->ReleaseDispatch();
  m_pFndInDoc->ReleaseDispatch();
  m_pRpInDoc->ReleaseDispatch();
  m_pDoc->ReleaseDispatch();

  delete   m_pRange;
  m_pRange=NULL;
    delete   m_pFndInDoc;
  m_pFndInDoc=NULL;
    delete   m_pRpInDoc;
  m_pRpInDoc=NULL;
      delete   m_pDoc;
  m_pDoc=NULL; 

}

//tool
CString CMsWordTool::DoubleToString(double src,int nAfterDot)
{
  CString OutRst = _T("");
  if(nAfterDot == 0)
  {
    OutRst.Format(_T("%0.0f"),src);
  }
  else if(nAfterDot == 1)
  {
    OutRst.Format(_T("%0.1f"),src);
  }
  else if(nAfterDot == 2)
  {
    OutRst.Format(_T("%0.2f"),src);
  }
  else if(nAfterDot == 3)
  {
    OutRst.Format(_T("%0.3f"),src);
  }
  else if(nAfterDot == 4)
  {
    OutRst.Format(_T("%0.4f"),src);
  }
  else if(nAfterDot == 5)
  {
    OutRst.Format(_T("%0.5f"),src);
  }
  else
    OutRst.Format(_T("%f"),src);

  return OutRst;
}

CString CMsWordTool::IntToString(int src)
{
  CString OutRst = _T("");
  OutRst.Format(_T("%d"),src);
  return OutRst;
}

void CMsWordTool::loopSelection(CString bm1,CString bm2,CArray<CString,CString>& ReplaceList,CArray<CString,CString>& ContentList)
{
  Selection wdSel;
  if(!GetSelection4bm(wdSel,bm1,bm2))
    return ;

  wdSel.Copy();
  int iReplace = ReplaceList.GetSize();
  int iConts = ContentList.GetSize();
  if(iReplace ==0 || iConts ==0){ASSERT(0);return;}
  int iLines = iConts/iReplace;
  int iCounts = 0;
  COleVariant un((short)5);
  COleVariant voPaste((short)0);
  COleVariant  vOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
  for (int i=0;i<iLines;i++)
  {
    for (int j=0;j<iReplace;j++)
    {
      iCounts = i*iReplace+j;
      ReplaceParameter(ReplaceList[j],ContentList[iCounts]);
    }
    if(i < iLines-1)
    {
      wdSel.EndKey(&un,&vOptional);
      wdSel.TypeParagraph();
      wdSel.Paste();
    }
  }

}
bool CMsWordTool::GetSelection4bm(Selection& wdSel,CString bm1,CString bm2)
{
  COleVariant cbookstart(bm1);
  COleVariant cbookend(bm2);

  COleVariant vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);


  wdSel.AttachDispatch(m_pMsWordTooldecorator->GetApp()->GetSelection());

  Selection wdSelect;
  Selection wdSelect2;
  wdSelect.AttachDispatch(m_pMsWordTooldecorator->GetApp()->GetSelection());
  Range m_pRange1;
  Range m_pRange2;

  try
  {
    m_pRange1  = (Range)wdSelect.GoTo(COleVariant((short)-1),vOpt,vOpt,cbookstart);
  }
  catch(CException* e)
  {
     return false;
  }

  try
  {
    m_pRange2  = (Range)wdSelect.GoTo(COleVariant((short)-1),vOpt,vOpt,cbookend);
  }
  catch(CException* e)
  {
    return false;
  }


  wdSel.SetStart(m_pRange1.GetStart()) ;
  wdSel.SetEnd(m_pRange2.GetEnd());

  return true;

}
void CMsWordTool::FillMutiTable(CString strBookMark,CArray<CString,CString>& ContentList,int nRows,int nColumns,int iMutiStartColumn,int iMutiEndColumn,int InnerLines)
{
  COleVariant vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
  Selection s;
  s.AttachDispatch(m_pMsWordTooldecorator->GetApp()->GetSelection());
  COleVariant vName = COleVariant(strBookMark, VT_BSTR);
  s.GoTo(COleVariant((short)-1),vOpt,vOpt,vName);


  int iCurColumID = 0;
  //绘制表格
  for (int i=0;i<nRows * nColumns ;i++)
  {

    if(i>0 && !((i)%(nColumns) == 0))
    {
      s.MoveRight(COleVariant((short)12),COleVariant((short)1),COleVariant((short)0));
    }

    //s.TypeText(ContentList.GetAt(i));
    if((i+1)%(nColumns) == 0)
    {
      if(i+1 != nRows * nColumns)
      {
        s.InsertRowsBelow(COleVariant((short)1));
        s.Collapse(COleVariant((short)1));
      }

    }
  }
  //填充数据
  //!1.恢复到表格开始位置
  s.MoveLeft(COleVariant((short)12),COleVariant((short)(nColumns-1)),COleVariant((short)0));
  s.MoveUp(COleVariant((short)5),COleVariant((short)(nRows - 1)),COleVariant((short)0));

  int iCount =0;
  for (int j=0;j<nRows;j++)
  {
    for(int i=0;i<nColumns;i++)
    {
      if(i+1 >=iMutiStartColumn && i+1 <= iMutiEndColumn)
      {
        Cells cs = s.GetCells();
        cs.Split(COleVariant((short)2),COleVariant((short)1),COleVariant((short)FALSE));
        s.TypeText(ContentList.GetAt(iCount));
        iCount++;

        s.MoveDown(COleVariant((short)5),COleVariant((short)1),COleVariant((short)0));
        s.TypeText(ContentList.GetAt(iCount));
        s.MoveUp(COleVariant((short)5),COleVariant((short)1),COleVariant((short)0));
        s.MoveRight(COleVariant((short)12),COleVariant((short)1),COleVariant((short)0));
        iCount++;

      }
      else
      {
        s.TypeText(ContentList.GetAt(iCount));
        s.MoveRight(COleVariant((short)12),COleVariant((short)1),COleVariant((short)0));
        iCount++;
      }
    }

    s.MoveDown(COleVariant((short)5),COleVariant((short)1),COleVariant((short)0));
    //    s.MoveLeft(COleVariant((short)12),COleVariant((short)(nColumns-1)),COleVariant((short)0));


  }
}

void CMsWordTool::WriteText(CString strBookMark,CTextFormat& TextFormat,int nAlignment)
{
//   COleVariant vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
//   
//   m_pDocSection->AttachDispatch(m_pMsWordTooldecorator->GetApp()->GetSelection());
//   COleVariant vName = COleVariant(strBookMark, VT_BSTR);
//   try
//   {
//     m_pDocSection->GoTo(COleVariant((short)-1),vOpt,vOpt,vName);
//   }
//   catch(CException* e)
//   {
//     return;
//   }

  if(!GetSelection(strBookMark))
    return;

  _Font   font;   
  font   =   m_pDocSection->GetFont();   
  font.SetNameFarEast(TextFormat.FontName);   
  //font.SetNameFarEast(_T("宋体"));   
  font.SetNameAscii(L"Times New Roman");   
  font.SetNameOther(L"Times New Roman");  

  font.SetSize(TextFormat.nFontSize);   
  font.SetBold(TextFormat.bBold ? -1 : 0);
  font.SetColor(TextFormat.FontColor);


  //插入图片操作
  if(TextFormat.IsInsertPic)
  {
    SetParaphFormat(1,0);
    InsertPicture(TextFormat.strPicPath);
    m_pDocSection->TypeParagraph();
    font.SetUnderline(1);
	WriteTextHaveEquation(TextFormat.strPicExpress);
    //m_pDocSection->TypeText(TextFormat.strPicExpress);
    font.SetUnderline(0);
    m_pDocSection->TypeParagraph();

  }
  else
  {
	SetOutLineLevel(TextFormat.iOutLineLevel);
    SetParaphFormat(nAlignment,TextFormat.LeftIndent);
	WriteTextHaveEquation(TextFormat.strContent);
    //m_pDocSection->TypeText(TextFormat.strContent);
  }

  //m_pDocSection->TypeParagraph();


}
bool CMsWordTool::GetSelection(CString strBookMark)
{
  COleVariant vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

  m_pDocSection->AttachDispatch(m_pMsWordTooldecorator->GetApp()->GetSelection());
  COleVariant vName = COleVariant(strBookMark, VT_BSTR);
  try
  {
    m_pDocSection->GoTo(COleVariant((short)-1),vOpt,vOpt,vName);
    //_ParagraphFormat p = m_pDocSection->GetParagraphFormat();
    //p.SetOutlineLevel(10);
    //缩进4cm
    //p.SetLeftIndent(m_pMsWordTooldecorator->GetApp()->CentimetersToPoints(-4));
    //p.SetFirstLineIndent(m_pMsWordTooldecorator->GetApp()->CentimetersToPoints(-4));
    
    //m_pDocSection->SetParagraphFormat(p);

    
  }
  catch(CException* e)
  {
    return false;
  }

  return true;

}

void CMsWordTool::WriteText(CString strBookMark,std::vector<CTextFormat*>& strContent,int nAlignment)
{
//   COleVariant vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
//   
//   m_pDocSection->AttachDispatch(m_pMsWordTooldecorator->GetApp()->GetSelection());
//   COleVariant vName = COleVariant(strBookMark, VT_BSTR);
//   try
//   {
//     m_pDocSection->GoTo(COleVariant((short)-1),vOpt,vOpt,vName);
//   }
//   catch(CException* e)
//   {
//     return ;
//   }
  if(!GetSelection(strBookMark))
    return;

  _Font   font;   
  font   =   m_pDocSection->GetFont();   
  CString strfront = L"";
  CString strQZ = L"";



  
  //p.SetOutlineLevel(10);
  //缩进4cm
  //double LeftIndent = 0.0;
  //p.SetLeftIndent(LeftIndent);
  //p.SetFirstLineIndent(LeftIndent);


  int iIndex = 0 ;
  for (std::vector<CTextFormat*>::iterator pIt = strContent.begin();pIt!=strContent.end();++pIt,++iIndex)
  {
    //font.SetName((*pIt)->FontName); 
	font.SetNameFarEast((*pIt)->FontName);   
	//font.SetNameFarEast(_T("宋体"));   
	font.SetNameAscii(L"Times New Roman");   
	font.SetNameOther(L"Times New Roman");  

    font.SetSize((*pIt)->nFontSize);   
    font.SetBold((*pIt)->bBold ? -1 : 0);
    font.SetColor((*pIt)->FontColor);
    //插入图片操作
    if((*pIt)->IsInsertPic)
    {
      SetParaphFormat(1,0);
      InsertPicture((*pIt)->strPicPath);
      m_pDocSection->TypeParagraph();
      font.SetUnderline(1);
      //m_pDocSection->TypeText((*pIt)->strPicExpress);
	  WriteTextHaveEquation((*pIt)->strPicExpress);
      font.SetUnderline(0);
      m_pDocSection->TypeParagraph();
    }
    else
    {

        SetOutLineLevel((*pIt)->iOutLineLevel);
		SetParaphFormat(nAlignment,((*pIt)->nStep-1) * 2 ,(*pIt)->LeftIndent);

      if((*pIt)->nStep == 2)
      {
        //strQZ = ((*pIt)->IsHaveFisrtSigle)? L"\t" : L"\t";
		//SetParaphFormat(nAlignment,2);

        strfront = strQZ + (*pIt)->strContent;
      }
      else if((*pIt)->nStep == 3)
      {
        //strQZ =((*pIt)->IsHaveFisrtSigle)? L"\t\t" : L"\t\t" ;
		//  SetParaphFormat(nAlignment,4);
        strfront =strQZ + (*pIt)->strContent;
      }
      else if((*pIt)->nStep == 4)
      {
        //strQZ =((*pIt)->IsHaveFisrtSigle)? L"\t\t\t" : L"\t\t\t";
		//    SetParaphFormat(nAlignment,6);
        strfront = strQZ + (*pIt)->strContent;
      }
      else
        strfront = (*pIt)->strContent;



      WriteTextHaveEquation(strfront);

	  if(iIndex < strContent.size()-1 )
		  m_pDocSection->TypeParagraph();

    }


   

  }

}



void CMsWordTool::SetEnter()
{
  m_pDocSection->TypeParagraph();
}
void CMsWordTool::WriteTextBySection(CTextFormat& TextFormat,int nAlignment/* = 0*/)
{
  if (!m_pDocSection->m_lpDispatch) 
  {
    //AfxMessageBox("Select为空，字体设置失败!", MB_OK|MB_ICONWARNING);
    return;
  }

  _Font   font;   
  font   =   m_pDocSection->GetFont();   
  //font.SetName(TextFormat.FontName);   
  font.SetNameFarEast(TextFormat.FontName);   
  //font.SetNameFarEast(_T("宋体"));   
  font.SetNameAscii(L"Times New Roman");   
  font.SetNameOther(L"Times New Roman");  

  font.SetSize(TextFormat.nFontSize);   
  font.SetBold(TextFormat.bBold ? -1 : 0);
  font.SetColor(TextFormat.FontColor);

  SetOutLineLevel(TextFormat.iOutLineLevel);
  SetParaphFormat(nAlignment,TextFormat.LeftIndent);

  WriteTextHaveEquation(TextFormat.strContent);

  //m_pDocSection->TypeText(TextFormat.strContent);
  //m_pDocSection->TypeParagraph();

  
}
void CMsWordTool::WriteTextBySection(std::vector<CTextFormat*>& strContent,int nAlignment/* = 0*/)
{
  if (!m_pDocSection->m_lpDispatch) 
  {
    //AfxMessageBox("Select为空，字体设置失败!", MB_OK|MB_ICONWARNING);
    return;
  }

  _Font   font;   
  font   =   m_pDocSection->GetFont();   
  CString strfront = L"";
  CString strQZ = L"";


int iIndex = 0 ;
  for (std::vector<CTextFormat*>::iterator pIt = strContent.begin();pIt!=strContent.end();++pIt,++iIndex)
  {
    //font.SetName((*pIt)->FontName);   

	font.SetNameFarEast((*pIt)->FontName);   
	//font.SetNameFarEast(_T("宋体"));   
	font.SetNameAscii(L"Times New Roman");   
	font.SetNameOther(L"Times New Roman");  

    font.SetSize((*pIt)->nFontSize);   
    font.SetBold((*pIt)->bBold ? -1 : 0);
    font.SetColor((*pIt)->FontColor);

	SetOutLineLevel((*pIt)->iOutLineLevel);
    SetParaphFormat(nAlignment,((*pIt)->nStep-1) * 2 ,(*pIt)->LeftIndent);


    if((*pIt)->nStep == 2)
    {
      //strQZ = ((*pIt)->IsHaveFisrtSigle)? L"\t" : L"\t";
	  //SetParaphFormat(nAlignment,2);
      strfront = strQZ + (*pIt)->strContent;
    }
    else if((*pIt)->nStep == 3)
    {
      //strQZ =((*pIt)->IsHaveFisrtSigle)? L"\t\t" : L"\t\t" ;
		//SetParaphFormat(nAlignment,4);
      strfront =strQZ + (*pIt)->strContent;
    }
    else if((*pIt)->nStep == 4)
    {
      //strQZ =((*pIt)->IsHaveFisrtSigle)? L"\t\t\t" : L"\t\t\t";
		//SetParaphFormat(nAlignment,6);
      strfront = strQZ + (*pIt)->strContent;
    }
    else
      strfront = (*pIt)->strContent;


   WriteTextHaveEquation(strfront);
   // m_pDocSection->TypeText(strfront);
   if(iIndex < strContent.size() -1)
		m_pDocSection->TypeParagraph();

  }
}

//!0-left 1-center 2-right
void CMsWordTool::SetParaphFormat(int nAlignment/* = 1*/,int nLeftIndent /*= 0*/,int nFirstLineIndent /*= 0*/)
{
  if (!m_pDocSection->m_lpDispatch) 
  {
    //AfxMessageBox("Select为空，字体设置失败!", MB_OK|MB_ICONWARNING);
    return;
  }

  _ParagraphFormat p = m_pDocSection->GetParagraphFormat();
  p.SetAlignment(nAlignment);
  p.SetCharacterUnitLeftIndent(nLeftIndent);
  p.SetCharacterUnitFirstLineIndent(nFirstLineIndent);
  m_pDocSection->SetParagraphFormat(p);

}
void CMsWordTool::SetFontStyle(double nFontSize,COLORREF FontColor/*=RGB(0,0,0)*/,CString FontName /*= _T("宋体")*/,BOOL bBold /*= FALSE*/, BOOL bItalic /*= FALSE*/ , BOOL bUnderLine /*= FALSE*/)
{
  if (m_pDocSection->m_lpDispatch) 
  {
    //AfxMessageBox("Select为空，字体设置失败!", MB_OK|MB_ICONWARNING);
    return;
  }
  _Font   font;   
  font   =   m_pDocSection->GetFont();   

  
  font.SetItalic(bItalic);
  font.SetUnderline(bUnderLine);

  font.SetNameFarEast(FontName);   
  font.SetNameAscii(L"Times New Roman");   
  font.SetNameOther(L"Times New Roman");  

  font.SetSize(nFontSize);   
  font.SetBold(bBold ? -1 : 0);
  font.SetColor(FontColor);

  m_pDocSection->SetFont(font);

}

//返回一个正文的Selection
void CMsWordTool::WriteTitle(CString strTitle,int nStep/* = 1*/)
{
  if (m_pDocSection->m_lpDispatch) 
  {
    //AfxMessageBox("Select为空，字体设置失败!", MB_OK|MB_ICONWARNING);
    return;
  }
  long nVBAStep = -1*nStep -1;
  COleVariant   vStyle(nVBAStep);     //   -2为一级标题，...，-10为九级标题 
  _ParagraphFormat   pf   =   m_pDocSection->GetParagraphFormat(); 
  pf.SetStyle(&vStyle); 
  

  m_pDocSection->TypeText(strTitle);
  m_pDocSection->TypeParagraph();

  vStyle.lVal = -67;
  pf.SetStyle(&vStyle); 
 
}

bool CMsWordTool::AddTable(int nRows,int nColumns,vector<CString>& vctContent,vector<double>& vctColumWidth,int nAlignment/* = 1*/,double FontSize/* = 9*/)
{

  //m_wdTb = tbs.Add(m_wdSel.GetRange(), nRow, nColumn, &vtDefault, &vtAuto);	
  //m_wdTb = tbs.Item(1);


  if (!m_pDocSection->m_lpDispatch) 
  {
    //AfxMessageBox("Select为空，生成表格失败!", MB_OK|MB_ICONWARNING);
    return false;
  }

  if((int)vctContent.size() < nRows*nColumns)
  {
    //AfxMessageBox("填充内容不符合表格要求!", MB_OK|MB_ICONWARNING);
    for (int i=0;i<nRows*nColumns-(int)vctContent.size();i++)
    {
      vctContent.push_back(L"");
    }
  }

  Tables tables=m_pDoc->GetTables();
//   COleVariant defaultBehavior((short)1);
//   COleVariant AutoFitBehavior((short)0);

  
  //SetParaphFormat(nAlignment);

  Range rang = m_pDocSection->GetRange() ;

  _ParagraphFormat p = rang.GetParagraphFormat();
  p.SetAlignment(0);
  rang.SetParagraphFormat(p);

  VARIANT vtDefault, vtAuto;
  vtDefault.vt = VT_I4;
  vtAuto.vt = VT_I4;
  vtDefault.intVal = 1;
  vtAuto.intVal = 0;

  Table table = tables.Add(rang,nRows,nColumns,&vtDefault, &vtAuto);
  table.Select();
  p = m_pDocSection->GetParagraphFormat();
  p.SetAlignment(1);
  m_pDocSection->SetParagraphFormat(p);
  

//   table.Select();
//   SetParaphFormat(1);
//   Cells cs=m_pDocSection->GetCells();
//   cs.SetVerticalAlignment(1);


  
  for (int i=1;i<=nRows;i++)
  {
    for (int j=1;j<=nColumns;j++)
    {
      Cell c1=table.Cell(i,j);
      c1.Select();

      _ParagraphFormat pCell = m_pDocSection->GetParagraphFormat();
      pCell.SetAlignment(nAlignment);
      m_pDocSection->SetParagraphFormat(pCell);

      
      //cs.SetVerticalAlignment(1);
      _Font   font;   
      font   =   m_pDocSection->GetFont();   
      font.SetItalic(0);
      font.SetBold(0);
      m_pDocSection->SetFont(font);
      pCell.ReleaseDispatch();
      font.ReleaseDispatch();

      Cells cs=m_pDocSection->GetCells();
      cs.SetVerticalAlignment(1);



      WriteTextHaveEquation(vctContent[(i-1)*nColumns+j-1]);
      cs.ReleaseDispatch();
      c1.ReleaseDispatch();

    }
  }


  Columns oColumns = table. GetColumns();
  for (size_t i=0;i<vctColumWidth.size();i++)
  {
    Column oCol =  oColumns.Item(i+1);
    oCol.SetPreferredWidth(vctColumWidth.at(i));
  }

  Rows rs =  table.GetRows();
  rs.SetAlignment(1);

  Borders bds = table.GetBorders();
  int wdBorderLeft = -2 ;
  int wdBorderRight = -4 ;
  int wdBorderTop = -1 ;
  int wdBorderBottom = -3 ;
  int wdLineWidth150pt = 12 ;



  Border bd =  bds.Item(wdBorderLeft) ;
  bd.SetLineWidth(wdLineWidth150pt);

  bd =  bds.Item(wdBorderRight) ;
  bd.SetLineWidth(wdLineWidth150pt);

  bd =  bds.Item(wdBorderTop) ;
  bd.SetLineWidth(wdLineWidth150pt);

  bd =  bds.Item(wdBorderBottom) ;
  bd.SetLineWidth(wdLineWidth150pt);


  bd.ReleaseDispatch();
  bds.ReleaseDispatch();
  table.ReleaseDispatch();
  tables.ReleaseDispatch();

  //   wdSel.MoveRight(COleVariant((short)1),COleVariant(short(1)),COleVariant(short(0)));
  //   wdSel.TypeParagraph();

  COleVariant firstParameter((short)1);
  COleVariant secondParameter((short)2);
  COleVariant thirdParameter((short)0);

  m_pDocSection->MoveRight(&firstParameter, &secondParameter, &thirdParameter);
  m_pDocSection->TypeParagraph();

//   rang = m_pDocSection->GetRange() ;
//   p = rang.GetParagraphFormat();
//   p.SetAlignment(0);
//   rang.SetParagraphFormat(p);


  return true;

}

CString CMsWordTool::CreateCalBookGraph(const map<CGraphKey,double>& srcData,const CString& strX_Name,const CString& strY_Name,const map<CGraphKey,CString>& strTLS,const CString& picName)
{
  return m_pCreatePic->CreatGraph(srcData,strX_Name,strY_Name,strTLS,picName);
}
void CMsWordTool::SetHeader(const CString& strHeader)
{
  Sections oSecs = m_pDoc->GetSections();
  Section sSec = oSecs.GetFirst();

  //页眉
  HeadersFooters pHFs = sSec.GetHeaders(); 
  HeaderFooter pHFP = pHFs.Item(1); 
  Range pRange = pHFP.GetRange(); 
  _ParagraphFormat p = pRange.GetParagraphFormat();
  p.SetAlignment(2);
  pRange.SetParagraphFormat(p);
  pRange.SetText(strHeader); 
}
void CMsWordTool::SetFooter(const CString& strFooter)
{
  Sections oSecs = m_pDoc->GetSections();
  Section sSec = oSecs.GetFirst();
  //页脚
  HeadersFooters pHFs = sSec.GetFooters(); 
  HeaderFooter pHFP = pHFs.Item(1); 
  Range pRange = pHFP.GetRange(); 
  _ParagraphFormat p = pRange.GetParagraphFormat();
  p.SetAlignment(2);
  pRange.SetParagraphFormat(p);
  pRange.SetText(strFooter); 
}

void CMsWordTool::SetOutLineLevel( int nlevel /*= 10*/ )
{
	if (!m_pDocSection->m_lpDispatch) 
	{
		//AfxMessageBox("Select为空，字体设置失败!", MB_OK|MB_ICONWARNING);
		return;
	}

	_ParagraphFormat p = m_pDocSection->GetParagraphFormat();
	p.SetOutlineLevel(nlevel);
	m_pDocSection->SetParagraphFormat(p);
}


//////////////////////////////////////////////////////////////////////////

void CMsWordTool::WriteSuperScript(LPCTSTR lpszText, LPCTSTR lpszSuper)
{
	
	
	{
		_Font oFont = m_pDocSection->GetFont();
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(0));
		m_pDocSection->TypeText(lpszText);
		oFont.SetSuperscript(TRUE);

		m_pDocSection->TypeText(lpszSuper);

		oFont.SetSuperscript(FALSE);
	}
}


void CMsWordTool::WriteSuperScript(long ascii, LPCTSTR lpszSuper)
{
	
	
	{
		_Font oFont = m_pDocSection->GetFont();
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(0));
		m_pDocSection->InsertSymbol(ascii,L"Symbol");

		oFont.SetSuperscript(TRUE);

		m_pDocSection->TypeText(lpszSuper);

		oFont.SetSuperscript(FALSE);
	}
}

void CMsWordTool::WriteSuperScript(long ascii, long sup)
{
	
	
	{
		_Font oFont = m_pDocSection->GetFont();
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(0));
		m_pDocSection->InsertSymbol(ascii,L"Symbol");

		oFont.SetSuperscript(TRUE);

		m_pDocSection->InsertSymbol(sup,L"Symbol");

		oFont.SetSuperscript(FALSE);
	}
}

void CMsWordTool::WriteSuperScript(LPCTSTR lpszText,long ascii)
{
	
	
	{
		_Font oFont = m_pDocSection->GetFont();
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(0));
		m_pDocSection->TypeText(lpszText);
		oFont.SetSuperscript(TRUE);
		m_pDocSection->InsertSymbol(ascii,L"Symbol");
		oFont.SetSuperscript(FALSE);
	}

}


void CMsWordTool::WriteSuperScript1(LPCTSTR lpszText, LPCTSTR lpszSuper)
{
	
	
	{
		_Font oFont = m_pDocSection->GetFont();
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(1));
		m_pDocSection->TypeText(lpszText);
		oFont.SetItalic(long(0));
		oFont.SetSuperscript(TRUE);
		m_pDocSection->TypeText(lpszSuper);
		oFont.SetSuperscript(FALSE);
	}
}


void CMsWordTool::WriteSuperScript1(long ascii, LPCTSTR lpszSuper)
{
	
	
	{
		_Font oFont = m_pDocSection->GetFont();
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(1));
		m_pDocSection->InsertSymbol(ascii,L"Symbol");
		oFont.SetItalic(long(0));
		oFont.SetSuperscript(TRUE);
		m_pDocSection->TypeText(lpszSuper);
		oFont.SetSuperscript(FALSE);
	}
}

void CMsWordTool::WriteSuperScript1(long ascii, long sup)
{
	
	
	{
		_Font oFont = m_pDocSection->GetFont();
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(1));
		m_pDocSection->InsertSymbol(ascii,L"Symbol");
		oFont.SetItalic(long(0));
		oFont.SetSuperscript(TRUE);
		m_pDocSection->InsertSymbol(sup,L"Symbol");
		oFont.SetSuperscript(FALSE);
	}
}

void CMsWordTool::WriteSuperScript1(LPCTSTR lpszText,long ascii)
{
	
	
	{
		_Font oFont = m_pDocSection->GetFont();
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(1));
		m_pDocSection->TypeText(lpszText);
		oFont.SetItalic(long(0));
		oFont.SetSuperscript(TRUE);
		m_pDocSection->InsertSymbol(ascii,L"Symbol");
		oFont.SetSuperscript(FALSE);
	}

}


void CMsWordTool::WriteSubScript(LPCTSTR lpszText, LPCTSTR lpszSub)
{
	
	
	{
		_Font oFont = m_pDocSection->GetFont();
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(0));
		m_pDocSection->TypeText(lpszText);
		oFont.SetSubscript(TRUE);
		m_pDocSection->TypeText(lpszSub);
		oFont.SetSubscript(FALSE);
	}
}

void CMsWordTool::WriteSubScript(long ascii, LPCTSTR lpszSub)
{
	
	
	{
		_Font oFont = m_pDocSection->GetFont();
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(0));
		m_pDocSection->InsertSymbol(ascii,L"Symbol");
		oFont.SetSubscript(TRUE);
		m_pDocSection->TypeText(lpszSub);
		oFont.SetSubscript(FALSE);
	}
}

void CMsWordTool::WriteSubScript(long ascii, long sub)
{
	
	
	{
		_Font oFont = m_pDocSection->GetFont();
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(0));
		m_pDocSection->InsertSymbol(ascii,L"Symbol");
		oFont.SetSubscript(TRUE);
		m_pDocSection->InsertSymbol(sub,L"Symbol");
		oFont.SetSubscript(FALSE);
	}
}

void CMsWordTool::WriteSubScript(LPCTSTR lpszText,long ascii)
{
	
	
	{
		_Font oFont = m_pDocSection->GetFont();
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(0));
		m_pDocSection->TypeText(lpszText);
		oFont.SetSubscript(TRUE);
		m_pDocSection->InsertSymbol(ascii,L"Symbol");
		oFont.SetSubscript(FALSE);
	}
}

void CMsWordTool::WriteSubScript1(LPCTSTR lpszText, LPCTSTR lpszSub)
{
	
	
	{
		_Font oFont = m_pDocSection->GetFont();
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(1));
		m_pDocSection->TypeText(lpszText);
		oFont.SetItalic(long(0));
		oFont.SetSubscript(TRUE);
		m_pDocSection->TypeText(lpszSub);
		oFont.SetSubscript(FALSE);
	}
}


void CMsWordTool::WriteSubScript1(long ascii, long sub)
{
	
	
	{
		_Font oFont = m_pDocSection->GetFont();
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(1));
		m_pDocSection->InsertSymbol(ascii,L"Symbol");
		oFont.SetItalic(long(0));
		oFont.SetSubscript(TRUE);
		m_pDocSection->InsertSymbol(sub,L"Symbol");
		oFont.SetSubscript(FALSE);
	}
}

void CMsWordTool::WriteSubScript1(long ascii, LPCTSTR lpszSub)
{
	
	
	{
		_Font oFont = m_pDocSection->GetFont();
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(1));
		m_pDocSection->InsertSymbol(ascii,L"Symbol");
		oFont.SetItalic(long(0));
		oFont.SetSubscript(TRUE);
		m_pDocSection->TypeText(lpszSub);
		oFont.SetSubscript(FALSE);
	}
}

void CMsWordTool::WriteSubScript1(LPCTSTR lpszText,long ascii)
{
	
	
	{
		_Font oFont = m_pDocSection->GetFont();
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(1));
		m_pDocSection->TypeText(lpszText);
		oFont.SetItalic(long(0));
		oFont.SetSubscript(TRUE);
		m_pDocSection->InsertSymbol(ascii,L"Symbol");
		oFont.SetSubscript(FALSE);
	}
}

void CMsWordTool::WriteSubScript2(long ascii, long sub)
{
	
	
	{
		_Font oFont = m_pDocSection->GetFont();
		CString string;
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(0));
		string.Format(L"%d",ascii);
		m_pDocSection->TypeText(string);
		oFont.SetSubscript(TRUE);
		m_pDocSection->InsertSymbol(sub,L"Symbol");
		oFont.SetSubscript(FALSE);
	}
}

void CMsWordTool::WriteSubScript2(long ascii, LPCTSTR lpszSub)
{
	
	
	{
		_Font oFont = m_pDocSection->GetFont();
		CString string;
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(0));
		string.Format(L"%d",ascii);
		m_pDocSection->TypeText(string);
		oFont.SetSubscript(TRUE);
		m_pDocSection->TypeText(lpszSub);
		oFont.SetSubscript(FALSE);
	}
}

void CMsWordTool::WriteSubScript2(LPCTSTR lpszText, LPCTSTR lpszSub)
{
	
	
	{
		_Font oFont = m_pDocSection->GetFont();
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(0));
		m_pDocSection->TypeText(lpszText);
		oFont.SetSubscript(TRUE);
		m_pDocSection->TypeText(lpszSub);
		oFont.SetSubscript(FALSE);
	}
}

void CMsWordTool::WriteSubScript2(LPCTSTR lpszText,long ascii)
{
	
	
	{
		_Font oFont = m_pDocSection->GetFont();
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(0));
		m_pDocSection->TypeText(lpszText);
		oFont.SetSubscript(TRUE);
		m_pDocSection->InsertSymbol(ascii,L"Symbol");
		oFont.SetSubscript(FALSE);
	}
}

void CMsWordTool::WriteSuperScript2(long ascii, long sub)
{
	
	
	{
		_Font oFont = m_pDocSection->GetFont();
		CString string;
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(0));
		string.Format(L"%d",ascii);
		m_pDocSection->TypeText(string);
		oFont.SetSuperscript(TRUE);
		m_pDocSection->InsertSymbol(sub,L"Symbol");
		oFont.SetSuperscript(FALSE);
	}
}

void CMsWordTool::WriteSuperScript2(long ascii, LPCTSTR lpszSub)
{
	
	
	{
		_Font oFont = m_pDocSection->GetFont();
		CString string;
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(0));
		string.Format(L"%d",ascii);
		m_pDocSection->TypeText(string);
		oFont.SetSuperscript(TRUE);
		m_pDocSection->TypeText(lpszSub);
		oFont.SetSuperscript(FALSE);
	}
}

void CMsWordTool::WriteSuperScript2(LPCTSTR lpszText, LPCTSTR lpszSuper)
{
	
	
	{
		_Font oFont = m_pDocSection->GetFont();
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(0));
		m_pDocSection->TypeText(lpszText);
		oFont.SetSuperscript(TRUE);
		m_pDocSection->TypeText(lpszSuper);
		oFont.SetSuperscript(FALSE);
	}
}

void CMsWordTool::WriteSuperScript2(LPCTSTR lpszText,long ascii)
{

	
	{
		_Font oFont = m_pDocSection->GetFont();
		oFont.SetSuperscript(FALSE);
		oFont.SetSubscript(FALSE);
		oFont.SetItalic(long(0));
		m_pDocSection->TypeText(lpszText);
		oFont.SetSuperscript(TRUE);
		m_pDocSection->InsertSymbol(ascii,L"Symbol");
		oFont.SetSuperscript(FALSE);
	}
}

void CMsWordTool::WriteEquation(LPCTSTR strTxt)
{

	CString lineText = strTxt;
	WriteEquation(lineText);
}

void CMsWordTool::WriteEquation(CString strTxt)
{

	//WriteSubScript(LPCTSTR lpszText, LPCTSTR lpszSub);
	//WriteSubScript(long ascii, LPCTSTR lpszSub);
	//WriteSubScript1(LPCTSTR lpszText, LPCTSTR lpszSub);
	//WriteSubScript1(long ascii, LPCTSTR lpszSub);
	//WriteSuperScript(LPCTSTR lpszText, LPCTSTR lpszSuper);
	//m_pDocSection->InsertSymbol(long)


	long nSubSup = strTxt.Find(L"SubSup");
	while(nSubSup >= 0)
	{
		CString LeftStr,RightStr,SubStr,SupStr,SubSupStr;
		long n,n1,n2;
		CString str = L"",str1 = L"";
		LeftStr = strTxt.Left(nSubSup);
		RightStr = strTxt.Mid(nSubSup);
		n1 = RightStr.Find(L"(");
		n2 = RightStr.Find(L")");
		SubSupStr = RightStr.Mid(n1+1,n2-n1-1);
		RightStr = RightStr.Mid(n2+1);

		n = SubSupStr.Find(L",");
		if(n < 0)
		{
			SubStr = SubSupStr;
			SubSupStr = "Sub(,";
			SubSupStr += SubStr;
			SubSupStr += ")";
		}
		else
		{
			SubStr = SubSupStr.Left(n);
			SupStr = SubSupStr.Mid(n + 1);
			SubSupStr = "\\o(Sub(,";
			SubSupStr += SubStr;
			SubSupStr += "),";
			SubSupStr += "Sup(,";
			SubSupStr += SupStr;
			SubSupStr += "))";
		}
		strTxt = LeftStr + SubSupStr + RightStr;
		nSubSup = strTxt.Find(L"SubSup");
	}

	/*
	strTxt.Replace(L"SubSub","RealSjqySSuubb");
	strTxt.Replace(L"SupSup","RealSjqySSuupp");
	strTxt.Replace(L"SymSym","RealSjqySSyymm");
	strTxt.Replace(L"Symbol","RealSjqySSyymmbBOOLl");
	strTxt.Replace(L"Sub","WriteSjqySSuubbScript");
	strTxt.Replace(L"Sup","WriteSjqySSuuppeerrScript");
	strTxt.Replace(L"Sym","WriteSjqySSyymmbBOOLl");
	strTxt.Replace(L"RealSjqySSuubb","Sub");
	strTxt.Replace(L"RealSjqySSuupp","Sup");
	strTxt.Replace(L"RealSjqySSyymmbBOOLl",L"Symbol");
	strTxt.Replace(L"RealSjqySSyymm","Sym");
	*/
	strTxt.Replace(L"SubSub",L"RealSjqySSuubb");
	strTxt.Replace(L"SupSup",L"RealSjqySSuupp");
	strTxt.Replace(L"SymSym",L"RealSjqySSyymm");
	strTxt.Replace(L"Symbol",L"RealSjqySSyymmbBOOLl");
	strTxt.Replace(L"Sub(",L"WriteSjqySSuubbScript(");
	strTxt.Replace(L"Sup(",L"WriteSjqySSuuppeerrScript(");
	strTxt.Replace(L"Sub1(",L"WriteSjqySSuubbScript1(");
	strTxt.Replace(L"Sub2(",L"WriteSjqySSuubbScript2(");
	strTxt.Replace(L"Sup1(",L"WriteSjqySSuuppeerrScript1(");
	strTxt.Replace(L"Sup2(",L"WriteSjqySSuuppeerrScript2(");
	strTxt.Replace(L"Sym(",L"WriteSjqySSyymmbBOOLl(");
	strTxt.Replace(L"RealSjqySSuubb",L"Sub");
	strTxt.Replace(L"RealSjqySSuupp",L"Sup");
	strTxt.Replace(L"RealSjqySSyymmbBOOLl",L"Symbol");
	strTxt.Replace(L"RealSjqySSyymm",L"Sym");
	strTxt.Replace(L"Under(",L"WriteSjqyUUnnddeerrLine(");

	
	
	{
		Fields fields;
		Field field;
		COleVariant var1(long(-1),VT_I4);
		COleVariant var2(strTxt);
		COleVariant var3((short)0, VT_BOOL);
		_Font oFont = m_pDocSection->GetFont();
		fields = m_pDocSection->GetFields();
		Range range = m_pDocSection->GetRange();
		field = fields.Add(range,var1,var2,var3);
		int len = int(strTxt.GetLength());
		//int mbslen = int(_mbslen((const unsigned char *)(LPWSTR)(LPCTSTR)strTxt));
        int mbslen = strTxt.GetLength();
		m_pDocSection->MoveLeft(wdCharacter,mbslen + 2);
		//	m_pDocSection->MoveLeft(wdCharacter,len + 2);

		CString LeftStr;
		CString LeftmbsStr;
		long n = strTxt.Find(L"WriteSjqy");
		long mbsn;
		if(n > -1)
		{
			LeftmbsStr = strTxt.Left(n);
			mbsn = long(LeftmbsStr.GetLength());
		}
		while(n >= 0)
		{
			m_pDocSection->MoveRight(wdCharacter,mbsn);
			//		m_pDocSection->MoveRight(wdCharacter,n);
			strTxt = strTxt.Mid(n);
			int ascii = 0;
			int sym = 0;
			COleVariant vCount(long(1),VT_I4);
			COleVariant vUint(long(1),VT_I4);
			CString str = L"",str1 = L"";
			if( strTxt.Left(23) == "WriteSjqyUUnnddeerrLine")
			{
				n = strTxt.Find(L"(") + 1;
				//			m_pDocSection->MoveRight(wdCharacter,n,wdExtend);
				vCount.lVal = n;
				m_pDocSection->Delete(vUint, vCount);
				strTxt = strTxt.Mid(n);

				n = strTxt.Find(L")");
				//			m_pDocSection->MoveRight(wdCharacter,n+1,wdExtend);
				//				vCount.lVal = n + 1;
				LeftmbsStr = strTxt.Left(n);
				mbsn = int(LeftmbsStr.GetLength());
				vCount.lVal = mbsn + 1;

				m_pDocSection->Delete(vUint, vCount);
				LeftStr = strTxt.Left(n);
				strTxt = strTxt.Mid(n + 1);
				WriteUnderLine(LeftStr);
			}
			else if( strTxt.Left(22) == "WriteSjqySSuubbScript1")
			{
				n = strTxt.Find(L"(") + 1;
				//			m_pDocSection->MoveRight(wdCharacter,n,wdExtend);
				vCount.lVal = n;
				m_pDocSection->Delete(vUint, vCount);
				strTxt = strTxt.Mid(n);
				if( strTxt.Left(21) == "WriteSjqySSyymmbBOOLl")
				{
					oFont.SetItalic(long(1));
					n = strTxt.Find(L"(") + 1;
					//			m_pDocSection->MoveRight(wdCharacter,n,wdExtend);
					vCount.lVal = n;
					m_pDocSection->Delete(vUint, vCount);
					strTxt = strTxt.Mid(n);
					n = strTxt.Find(L")");
					LeftStr = strTxt.Left(n);
					strTxt = strTxt.Mid(n + 1);
					ascii = _ttoi(LeftStr);
					//			m_pDocSection->MoveRight(wdCharacter,n+1,wdExtend);
					//				vCount.lVal = n + 1;
					LeftmbsStr = strTxt.Left(n);
					mbsn = int(LeftmbsStr.GetLength());
					vCount.lVal = mbsn + 1;
					m_pDocSection->Delete(vUint, vCount);

					n = LeftStr.Find(L",");
					if(n < 0)
					{
						ascii = _ttoi(LeftStr);
						m_pDocSection->InsertSymbol(ascii,L"Symbol");
					}
					else
					{
						str = LeftStr.Left(n);
						ascii = _ttoi(str);
						str = LeftStr.Mid(n + 1);
						n = str.Find(L",");
						if(n < 0)
						{
							m_pDocSection->InsertSymbol(ascii, str);
						}
						else
						{
							str = str.Left(n);
							m_pDocSection->InsertSymbol(ascii, str);
						}
					}
					oFont.SetItalic(long(0));
				}

				n = strTxt.Find(L")");
				//			m_pDocSection->MoveRight(wdCharacter,n+1,wdExtend);
				//				vCount.lVal = n + 1;
				LeftmbsStr = strTxt.Left(n);
				mbsn = int(LeftmbsStr.GetLength());
				vCount.lVal = mbsn + 1;

				m_pDocSection->Delete(vUint, vCount);
				LeftStr = strTxt.Left(n);
				strTxt = strTxt.Mid(n + 1);
				n = LeftStr.Find(L",");
				if(n < 0)
					str = LeftStr;
				else
				{
					str = LeftStr.Left(n);
					str1 = LeftStr.Mid(n + 1);
				}
				ascii = _ttoi(str);
				sym = _ttoi(str1);
				if(ascii > 1000 || ascii < -1000)
				{
					if(sym < 1000 && sym > -1000 )
						WriteSubScript1(ascii,str1);
					else
						WriteSubScript1(ascii,sym);
				}
				else
				{
					if(sym < 1000 && sym > -1000 )
						WriteSubScript1(str,str1);
					else
						WriteSubScript1(str,sym);
				}
			}
			else if( strTxt.Left(22) == "WriteSjqySSuubbScript2")
			{
				n = strTxt.Find(L"(") + 1;
				//			m_pDocSection->MoveRight(wdCharacter,n,wdExtend);
				vCount.lVal = n;
				m_pDocSection->Delete(vUint, vCount);
				strTxt = strTxt.Mid(n);
				if( strTxt.Left(21) == "WriteSjqySSyymmbBOOLl")
				{
					oFont.SetItalic(long(1));
					n = strTxt.Find(L"(") + 1;
					//			m_pDocSection->MoveRight(wdCharacter,n,wdExtend);
					vCount.lVal = n;
					m_pDocSection->Delete(vUint, vCount);
					strTxt = strTxt.Mid(n);
					n = strTxt.Find(L")");
					LeftStr = strTxt.Left(n);
					strTxt = strTxt.Mid(n + 1);
					ascii = _ttoi(LeftStr);
					//			m_pDocSection->MoveRight(wdCharacter,n+1,wdExtend);
					//				vCount.lVal = n + 1;
					LeftmbsStr = strTxt.Left(n);
					mbsn = int(LeftmbsStr.GetLength());
					vCount.lVal = mbsn + 1;
					m_pDocSection->Delete(vUint, vCount);

					n = LeftStr.Find(L",");
					if(n < 0)
					{
						ascii = _ttoi(LeftStr);
						m_pDocSection->InsertSymbol(ascii,L"Symbol");
					}
					else
					{
						str = LeftStr.Left(n);
						ascii = _ttoi(str);
						str = LeftStr.Mid(n + 1);
						n = str.Find(L",");
						if(n < 0)
						{
							m_pDocSection->InsertSymbol(ascii, str);
						}
						else
						{
							str = str.Left(n);
							m_pDocSection->InsertSymbol(ascii, str);
						}
					}
					oFont.SetItalic(long(0));
				}

				n = strTxt.Find(L")");
				//			m_pDocSection->MoveRight(wdCharacter,n+1,wdExtend);
				//				vCount.lVal = n + 1;
				LeftmbsStr = strTxt.Left(n);
				mbsn = int(LeftmbsStr.GetLength());
				vCount.lVal = mbsn + 1;
				m_pDocSection->Delete(vUint, vCount);
				LeftStr = strTxt.Left(n);
				strTxt = strTxt.Mid(n + 1);
				n = LeftStr.Find(L",");
				if(n < 0)
					str = LeftStr;
				else
				{
					str = LeftStr.Left(n);
					str1 = LeftStr.Mid(n + 1);
				}
				//ascii = _ttoi(str);
				sym = _ttoi(str1);

				if(sym < 1000 && sym > -1000 )
					WriteSubScript2(str,str1);
				else
					WriteSubScript2(str,sym);

			}
			else if( strTxt.Left(21) == "WriteSjqySSuubbScript")
			{
				n = strTxt.Find(L"(") + 1;
				//			m_pDocSection->MoveRight(wdCharacter,n,wdExtend);
				vCount.lVal = n;
				m_pDocSection->Delete(vUint, vCount);

				strTxt = strTxt.Mid(n);
				if( strTxt.Left(21) == "WriteSjqySSyymmbBOOLl")
				{
					n = strTxt.Find(L"(") + 1;
					//			m_pDocSection->MoveRight(wdCharacter,n,wdExtend);
					vCount.lVal = n;
					m_pDocSection->Delete(vUint, vCount);
					strTxt = strTxt.Mid(n);
					n = strTxt.Find(L")");
					LeftStr = strTxt.Left(n);
					strTxt = strTxt.Mid(n + 1);
					ascii = _ttoi(LeftStr);
					//			m_pDocSection->MoveRight(wdCharacter,n+1,wdExtend);
					//				vCount.lVal = n + 1;
					LeftmbsStr = strTxt.Left(n);
					mbsn = int(LeftmbsStr.GetLength());
					vCount.lVal = mbsn + 1;
					m_pDocSection->Delete(vUint, vCount);

					n = LeftStr.Find(L",");
					if(n < 0)
					{
						ascii = _ttoi(LeftStr);
						m_pDocSection->InsertSymbol(ascii,L"Symbol");
					}
					else
					{
						str = LeftStr.Left(n);
						ascii = _ttoi(str);
						str = LeftStr.Mid(n + 1);
						n = str.Find(L",");
						if(n < 0)
						{
							m_pDocSection->InsertSymbol(ascii,L"Symbol");
						}
						else
						{
							str = str.Left(n);
							m_pDocSection->InsertSymbol(ascii, str);
						}
					}
				}

				n = strTxt.Find(L")");
				//			m_pDocSection->MoveRight(wdCharacter,n+1,wdExtend);
				//				vCount.lVal = n + 1;
				LeftmbsStr = strTxt.Left(n);
				mbsn = int(LeftmbsStr.GetLength());
				vCount.lVal = mbsn + 1;
				m_pDocSection->Delete(vUint, vCount);
				LeftStr = strTxt.Left(n);
				strTxt = strTxt.Mid(n + 1);
				n = LeftStr.Find(L",");
				if(n < 0)
					str = LeftStr;
				else
				{
					str = LeftStr.Left(n);
					str1 = LeftStr.Mid(n + 1);
				}
				ascii = _ttoi(str);
				sym = _ttoi(str1);
				if(ascii > 1000 || ascii < -1000)
				{
					if(sym < 1000 && sym > -1000 )
						WriteSubScript(ascii,str1);
					else
						WriteSubScript(ascii,sym);
				}
				else
				{
					if(sym < 1000 && sym > -1000 )
						WriteSubScript(str,str1);
					else
						WriteSubScript(str,sym);
				}
			}
			else if( strTxt.Left(26) == "WriteSjqySSuuppeerrScript1")
			{
				n = strTxt.Find(L"(") + 1;
				//			m_pDocSection->MoveRight(wdCharacter,n,wdExtend);
				vCount.lVal = n;
				m_pDocSection->Delete(vUint, vCount);

				strTxt = strTxt.Mid(n);
				if( strTxt.Left(21) == "WriteSjqySSyymmbBOOLl")
				{
					oFont.SetItalic(long(1));
					n = strTxt.Find(L"(") + 1;
					//			m_pDocSection->MoveRight(wdCharacter,n,wdExtend);
					vCount.lVal = n;
					m_pDocSection->Delete(vUint, vCount);
					strTxt = strTxt.Mid(n);
					n = strTxt.Find(L")");
					LeftStr = strTxt.Left(n);
					strTxt = strTxt.Mid(n + 1);
					ascii = _ttoi(LeftStr);
					//			m_pDocSection->MoveRight(wdCharacter,n+1,wdExtend);
					//				vCount.lVal = n + 1;
					LeftmbsStr = strTxt.Left(n);
					mbsn = int(LeftmbsStr.GetLength());
					vCount.lVal = mbsn + 1;
					m_pDocSection->Delete(vUint, vCount);

					n = LeftStr.Find(L",");
					if(n < 0)
					{
						ascii = _ttoi(LeftStr);
						m_pDocSection->InsertSymbol(ascii,L"Symbol");
					}
					else
					{
						str = LeftStr.Left(n);
						ascii = _ttoi(str);
						str = LeftStr.Mid(n + 1);
						n = str.Find(L",");
						if(n < 0)
						{
							m_pDocSection->InsertSymbol(ascii,L"Symbol");
						}
						else
						{
							str = str.Left(n);
							m_pDocSection->InsertSymbol(ascii, str);
						}
					}
					oFont.SetItalic(long(0));
				}

				n = strTxt.Find(L")");
				//			m_pDocSection->MoveRight(wdCharacter,n+1,wdExtend);
				//				vCount.lVal = n + 1;
				LeftmbsStr = strTxt.Left(n);
				mbsn = int(LeftmbsStr.GetLength());
				vCount.lVal = mbsn + 1;
				m_pDocSection->Delete(vUint, vCount);
				LeftStr = strTxt.Left(n);
				strTxt = strTxt.Mid(n + 1);
				n = LeftStr.Find(L",");
				if(n < 0)
					str = LeftStr;
				else
				{
					str = LeftStr.Left(n);
					str1 = LeftStr.Mid(n + 1);
				}
				ascii = _ttoi(str);
				sym = _ttoi(str1);
				if(ascii > 1000 || ascii < -1000)
				{
					if(sym < 1000 && sym > -1000 )
						WriteSuperScript1(ascii,str1);
					else
						WriteSuperScript1(ascii,sym);
				}
				else
				{
					if(sym < 1000 && sym > -1000 )
						WriteSuperScript1(str,str1);
					else
						WriteSuperScript1(str,sym);
				}
			}
			else if( strTxt.Left(26) == "WriteSjqySSuuppeerrScript2")
			{
				n = strTxt.Find(L"(") + 1;
				//			m_pDocSection->MoveRight(wdCharacter,n,wdExtend);
				vCount.lVal = n;
				m_pDocSection->Delete(vUint, vCount);

				strTxt = strTxt.Mid(n);
				if( strTxt.Left(21) == "WriteSjqySSyymmbBOOLl")
				{
					oFont.SetItalic(long(1));
					n = strTxt.Find(L"(") + 1;
					//			m_pDocSection->MoveRight(wdCharacter,n,wdExtend);
					vCount.lVal = n;
					m_pDocSection->Delete(vUint, vCount);
					strTxt = strTxt.Mid(n);
					n = strTxt.Find(L")");
					LeftStr = strTxt.Left(n);
					strTxt = strTxt.Mid(n + 1);
					ascii = _ttoi(LeftStr);
					//			m_pDocSection->MoveRight(wdCharacter,n+1,wdExtend);
					//				vCount.lVal = n + 1;
					LeftmbsStr = strTxt.Left(n);
					mbsn = int(LeftmbsStr.GetLength());
					vCount.lVal = mbsn + 1;
					m_pDocSection->Delete(vUint, vCount);

					n = LeftStr.Find(L",");
					if(n < 0)
					{
						ascii = _ttoi(LeftStr);
						m_pDocSection->InsertSymbol(ascii,L"Symbol");
					}
					else
					{
						str = LeftStr.Left(n);
						ascii = _ttoi(str);
						str = LeftStr.Mid(n + 1);
						n = str.Find(L",");
						if(n < 0)
						{
							m_pDocSection->InsertSymbol(ascii,L"Symbol");
						}
						else
						{
							str = str.Left(n);
							m_pDocSection->InsertSymbol(ascii, str);
						}
					}
					oFont.SetItalic(long(0));
				}

				n = strTxt.Find(L")");
				//			m_pDocSection->MoveRight(wdCharacter,n+1,wdExtend);
				//				vCount.lVal = n + 1;
				LeftmbsStr = strTxt.Left(n);
				mbsn = int(LeftmbsStr.GetLength());
				vCount.lVal = mbsn + 1;
				m_pDocSection->Delete(vUint, vCount);
				LeftStr = strTxt.Left(n);
				strTxt = strTxt.Mid(n + 1);
				n = LeftStr.Find(L",");
				if(n < 0)
					str = LeftStr;
				else
				{
					str = LeftStr.Left(n);
					str1 = LeftStr.Mid(n + 1);
				}
				//ascii = _ttoi(str);
				sym = _ttoi(str1);
				//if(ascii > 1000 || ascii < -1000)
				//{
				if(sym < 1000 && sym > -1000 )
					WriteSuperScript2(str,str1);
				else
					WriteSuperScript2(str,sym);
				//}
				//else
				//{
				//if(sym < 1000 && sym > -1000 )
				//	WriteSuperScript2(str,str1);
				//else
				//WriteSuperScript2(str,sym);
				//}
			}
			else if( strTxt.Left(25) == "WriteSjqySSuuppeerrScript")
			{
				n = strTxt.Find(L"(") + 1;
				//			m_pDocSection->MoveRight(wdCharacter,n,wdExtend);
				vCount.lVal = n;
				m_pDocSection->Delete(vUint, vCount);

				strTxt = strTxt.Mid(n);
				if( strTxt.Left(21) == "WriteSjqySSyymmbBOOLl")
				{
					n = strTxt.Find(L"(") + 1;
					//			m_pDocSection->MoveRight(wdCharacter,n,wdExtend);
					vCount.lVal = n;
					m_pDocSection->Delete(vUint, vCount);
					strTxt = strTxt.Mid(n);
					n = strTxt.Find(L")");
					LeftStr = strTxt.Left(n);
					strTxt = strTxt.Mid(n + 1);
					ascii = _ttoi(LeftStr);
					//			m_pDocSection->MoveRight(wdCharacter,n+1,wdExtend);
					//				vCount.lVal = n + 1;
					LeftmbsStr = strTxt.Left(n);
					mbsn = int(LeftmbsStr.GetLength());
					vCount.lVal = mbsn + 1;
					m_pDocSection->Delete(vUint, vCount);

					n = LeftStr.Find(L",");
					if(n < 0)
					{
						ascii = _ttoi(LeftStr);
						m_pDocSection->InsertSymbol(ascii,L"Symbol");
					}
					else
					{
						str = LeftStr.Left(n);
						ascii = _ttoi(str);
						str = LeftStr.Mid(n + 1);
						n = str.Find(L",");
						if(n < 0)
						{
							m_pDocSection->InsertSymbol(ascii,L"Symbol");
						}
						else
						{
							str = str.Left(n);
							m_pDocSection->InsertSymbol(ascii, str);
						}
					}
				}
				n = strTxt.Find(L")");
				//			m_pDocSection->MoveRight(wdCharacter,n+1,wdExtend);
				//				vCount.lVal = n + 1;
				LeftmbsStr = strTxt.Left(n);
				mbsn = int(LeftmbsStr.GetLength());
				vCount.lVal = mbsn + 1;
				m_pDocSection->Delete(vUint, vCount);
				LeftStr = strTxt.Left(n);
				strTxt = strTxt.Mid(n + 1);
				n = LeftStr.Find(L",");
				if(n < 0)
					str = LeftStr;
				else
				{
					str = LeftStr.Left(n);
					str1 = LeftStr.Mid(n + 1);
				}
				ascii = _ttoi(str);
				sym = _ttoi(str1);
				if(ascii > 1000 || ascii < -1000)
				{
					if(sym < 1000 && sym > -1000 )
						WriteSuperScript(ascii,str1);
					else
						WriteSuperScript(ascii,sym);
				}
				else
				{
					if(sym < 1000 && sym > -1000 )
						WriteSuperScript(str,str1);
					else
						WriteSuperScript(str,sym);
				}
			}
			else if( strTxt.Left(21) == "WriteSjqySSyymmbBOOLl")
			{
				n = strTxt.Find(L"(") + 1;
				//			m_pDocSection->MoveRight(wdCharacter,n,wdExtend);
				vCount.lVal = n;
				m_pDocSection->Delete(vUint, vCount);
				strTxt = strTxt.Mid(n);
				n = strTxt.Find(L")");
				LeftStr = strTxt.Left(n);
				strTxt = strTxt.Mid(n + 1);
				ascii = _ttoi(LeftStr);
				//			m_pDocSection->MoveRight(wdCharacter,n+1,wdExtend);
				//				vCount.lVal = n + 1;
				LeftmbsStr = strTxt.Left(n);
				mbsn = int(LeftmbsStr.GetLength());
				vCount.lVal = mbsn + 1;
				m_pDocSection->Delete(vUint, vCount);

				n = LeftStr.Find(L",");
				oFont.SetSuperscript(FALSE);
				oFont.SetSubscript(FALSE);
				if(n < 0)
				{
					ascii = _ttoi(LeftStr);
					m_pDocSection->InsertSymbol(ascii,L"Symbol");
				}
				else
				{
					str = LeftStr.Left(n);
					ascii = _ttoi(str);
					str = LeftStr.Mid(n + 1);
					n = str.Find(L",");
					if(n < 0)
					{
						m_pDocSection->InsertSymbol(ascii,L"Symbol");
					}
					else
					{
						str = str.Left(n);
						m_pDocSection->InsertSymbol(ascii, str);
					}
				}
			}
			n = strTxt.Find(L"WriteSjqy");
			if(n > -1)
			{
				LeftmbsStr = strTxt.Left(n);
				mbsn = int(LeftmbsStr.GetLength());
			}
		}
		//  [12/31/2006 gyj]
		m_pDocSection->MoveDown(wdParagraph, 1, wdExtend);
		m_pDocSection->EndKey(wdLine);

		//2005.8.11 改善word显示速度
		/*
		oFont.SetBold(long(0));
		oFont.SetUnderline(long(0));
		oFont.SetSize(float(10.5));
		oFont.SetSubscript(long(0));
		oFont.SetSuperscript(long(0));
		oFont.SetColor(-16777216);
		*/
	}
}

void CMsWordTool::WriteTextHaveEquation(LPCTSTR strTxt)
{

	CString lineText = strTxt;
	WriteTextHaveEquation(lineText);
}

void CMsWordTool::WriteTextHaveEquation(CString strTxt)
{



	strTxt.Replace(L"SubSub",L"RealSjqySSuubb");
	strTxt.Replace(L"SupSup",L"RealSjqySSuupp");
	strTxt.Replace(L"SymSym",L"RealSjqySSyymm");
	strTxt.Replace(L"Symbol",L"RealSjqySSyymmbBOOLl");
	strTxt.Replace(L"Sub(",L"WriteSjqySSuubbScript(");
	strTxt.Replace(L"Sup(",L"WriteSjqySSuuppeerrScript(");
	strTxt.Replace(L"Sub1(",L"WriteSjqySSuubbScript1(");
	strTxt.Replace(L"Sub2(",L"WriteSjqySSuubbScript2(");
	strTxt.Replace(L"Sup1(",L"WriteSjqySSuuppeerrScript1(");
	strTxt.Replace(L"Sup2(",L"WriteSjqySSuuppeerrScript2(");
	strTxt.Replace(L"Sym(",L"WriteSjqySSyymmbBOOLl(");
	strTxt.Replace(L"RealSjqySSuubb",L"Sub");
	strTxt.Replace(L"RealSjqySSuupp",L"Sup");
	strTxt.Replace(L"RealSjqySSyymmbBOOLl",L"Symbol");
	strTxt.Replace(L"RealSjqySSyymm",L"Sym");
	strTxt.Replace(L"Under(",L"WriteSjqyUUnnddeerrLine(");

	CString LeftStr;
	int n = strTxt.Find(L"WriteSjqy");
	
	{
		_Font oFont = m_pDocSection->GetFont();

		while(n >= 0)
		{
			LeftStr = strTxt.Left(n);
			m_pDocSection->TypeText(LeftStr);
			strTxt = strTxt.Mid(n);
			int ascii = 0, sym = 0;
			CString str = L"",str1 = L"";
			if( strTxt.Left(23) == "WriteSjqyUUnnddeerrLine")
			{
				n = strTxt.Find(L"(") + 1;
				strTxt = strTxt.Mid(n);
				n = strTxt.Find(L")");
				LeftStr = strTxt.Left(n);
				strTxt = strTxt.Mid(n + 1);
				WriteUnderLine(LeftStr);
			}
			else if( strTxt.Left(22) == "WriteSjqySSuubbScript1")
			{
				n = strTxt.Find(L"(") + 1;
				strTxt = strTxt.Mid(n);
				if( strTxt.Left(21) == "WriteSjqySSyymmbBOOLl")
				{
					oFont.SetItalic(long(1));
					n = strTxt.Find(L"(") + 1;
					strTxt = strTxt.Mid(n);
					n = strTxt.Find(L")");
					LeftStr = strTxt.Left(n);
					strTxt = strTxt.Mid(n + 1);
					n = LeftStr.Find(L",");
					if(n < 0)
					{
						ascii = _ttoi(LeftStr);
						m_pDocSection->InsertSymbol(ascii,L"Symbol");
					}
					else
					{
						str = LeftStr.Left(n);
						ascii = _ttoi(str);
						str = LeftStr.Mid(n + 1);
						n = str.Find(L",");
						if(n < 0)
						{
							m_pDocSection->InsertSymbol(ascii,L"Symbol");
						}
						else
						{
							str = str.Left(n);
							m_pDocSection->InsertSymbol(ascii, str);
						}
					}
					oFont.SetItalic(long(0));
				}

				n = strTxt.Find(L")");
				LeftStr = strTxt.Left(n);
				strTxt = strTxt.Mid(n + 1);
				n = LeftStr.Find(L",");
				if(n < 0)
					str = LeftStr;
				else
				{
					str = LeftStr.Left(n);
					str1 = LeftStr.Mid(n + 1);
				}
				ascii = _ttoi(str);
				sym = _ttoi(str1);
				if(ascii > 1000 || ascii < -1000)
				{
					if(sym < 1000 && sym > -1000 )
						WriteSubScript1(ascii,str1);
					else
						WriteSubScript1(ascii,sym);
				}
				else
				{
					if(sym < 1000 && sym > -1000 )
						WriteSubScript1(str,str1);
					else
						WriteSubScript1(str,sym);
				}
			}

			else if( strTxt.Left(22) == "WriteSjqySSuubbScript2")
			{
				n = strTxt.Find(L"(") + 1;
				strTxt = strTxt.Mid(n);
				if( strTxt.Left(21) == "WriteSjqySSyymmbBOOLl")
				{
					oFont.SetItalic(long(1));
					n = strTxt.Find(L"(") + 1;
					strTxt = strTxt.Mid(n);
					n = strTxt.Find(L")");
					LeftStr = strTxt.Left(n);
					strTxt = strTxt.Mid(n + 1);
					n = LeftStr.Find(L",");
					if(n < 0)
					{
						ascii = _ttoi(LeftStr);
						m_pDocSection->InsertSymbol(ascii,L"Symbol");
					}
					else
					{
						str = LeftStr.Left(n);
						ascii = _ttoi(str);
						str = LeftStr.Mid(n + 1);
						n = str.Find(L",");
						if(n < 0)
						{
							m_pDocSection->InsertSymbol(ascii,L"Symbol");
						}
						else
						{
							str = str.Left(n);
							m_pDocSection->InsertSymbol(ascii, str);
						}
					}
					oFont.SetItalic(long(0));
				}

				n = strTxt.Find(L")");
				LeftStr = strTxt.Left(n);
				strTxt = strTxt.Mid(n + 1);
				n = LeftStr.Find(L",");
				if(n < 0)
					str = LeftStr;
				else
				{
					str = LeftStr.Left(n);
					str1 = LeftStr.Mid(n + 1);
				}
				sym = _ttoi(str1);

				if(sym < 1000 && sym > -1000 )
					WriteSubScript2(str,str1);
				else
					WriteSubScript2(str,sym);


			}

			else if( strTxt.Left(21) == "WriteSjqySSuubbScript")
			{
				n = strTxt.Find(L"(") + 1;
				strTxt = strTxt.Mid(n);
				if( strTxt.Left(21) == "WriteSjqySSyymmbBOOLl")
				{
					n = strTxt.Find(L"(") + 1;
					strTxt = strTxt.Mid(n);
					n = strTxt.Find(L")");
					LeftStr = strTxt.Left(n);
					strTxt = strTxt.Mid(n + 1);
					n = LeftStr.Find(L",");
					if(n < 0)
					{
						ascii = _ttoi(LeftStr);
						m_pDocSection->InsertSymbol(ascii,L"Symbol");
					}
					else
					{
						str = LeftStr.Left(n);
						ascii = _ttoi(str);
						str = LeftStr.Mid(n + 1);
						n = str.Find(L",");
						if(n < 0)
						{
							m_pDocSection->InsertSymbol(ascii,L"Symbol");
						}
						else
						{
							str = str.Left(n);
							m_pDocSection->InsertSymbol(ascii, str);
						}
					}
				}
				n = strTxt.Find(L")");
				LeftStr = strTxt.Left(n);
				strTxt = strTxt.Mid(n + 1);
				n = LeftStr.Find(L",");
				if(n < 0)
					str = LeftStr;
				else
				{
					str = LeftStr.Left(n);
					str1 = LeftStr.Mid(n + 1);
				}
				ascii = _ttoi(str);
				sym = _ttoi(str1);
				if(ascii > 1000 || ascii < -1000)
				{
					if(sym < 1000 && sym > -1000 )
						WriteSubScript(ascii,str1);
					else
						WriteSubScript(ascii,sym);
				}
				else
				{
					if(sym < 1000 && sym > -1000 )
						WriteSubScript(str,str1);
					else
						WriteSubScript(str,sym);
				}
			}
			else if( strTxt.Left(26) == "WriteSjqySSuuppeerrScript1")
			{
				n = strTxt.Find(L"(") + 1;
				strTxt = strTxt.Mid(n);
				if( strTxt.Left(21) == "WriteSjqySSyymmbBOOLl")
				{
					oFont.SetItalic(long(1));
					n = strTxt.Find(L"(") + 1;
					strTxt = strTxt.Mid(n);
					n = strTxt.Find(L")");
					LeftStr = strTxt.Left(n);
					strTxt = strTxt.Mid(n + 1);
					n = LeftStr.Find(L",");
					if(n < 0)
					{
						ascii = _ttoi(LeftStr);
						m_pDocSection->InsertSymbol(ascii,L"Symbol");
					}
					else
					{
						str = LeftStr.Left(n);
						ascii = _ttoi(str);
						str = LeftStr.Mid(n + 1);
						n = str.Find(L",");
						if(n < 0)
						{
							m_pDocSection->InsertSymbol(ascii,L"Symbol");
						}
						else
						{
							str = str.Left(n);
							m_pDocSection->InsertSymbol(ascii, str);
						}
					}
					oFont.SetItalic(long(0));
				}

				n = strTxt.Find(L")");
				LeftStr = strTxt.Left(n);
				strTxt = strTxt.Mid(n + 1);
				n = LeftStr.Find(L",");
				if(n < 0)
					str = LeftStr;
				else
				{
					str = LeftStr.Left(n);
					str1 = LeftStr.Mid(n + 1);
				}
				ascii = _ttoi(str);
				sym = _ttoi(str1);
				if(ascii > 1000 || ascii < -1000)
				{
					if(sym < 1000 && sym > -1000 )
						WriteSuperScript1(ascii,str1);
					else
						WriteSuperScript1(ascii,sym);
				}
				else
				{
					if(sym < 1000 && sym > -1000 )
						WriteSuperScript1(str,str1);
					else
						WriteSuperScript1(str,sym);
				}
			}
			else if( strTxt.Left(26) == "WriteSjqySSuuppeerrScript2")
			{
				n = strTxt.Find(L"(") + 1;
				strTxt = strTxt.Mid(n);
				if( strTxt.Left(21) == "WriteSjqySSyymmbBOOLl")
				{
					oFont.SetItalic(long(1));
					n = strTxt.Find(L"(") + 1;
					strTxt = strTxt.Mid(n);
					n = strTxt.Find(L")");
					LeftStr = strTxt.Left(n);
					strTxt = strTxt.Mid(n + 1);
					n = LeftStr.Find(L",");
					if(n < 0)
					{
						ascii = _ttoi(LeftStr);
						m_pDocSection->InsertSymbol(ascii,L"Symbol");
					}
					else
					{
						str = LeftStr.Left(n);
						ascii = _ttoi(str);
						str = LeftStr.Mid(n + 1);
						n = str.Find(L",");
						if(n < 0)
						{
							m_pDocSection->InsertSymbol(ascii,L"Symbol");
						}
						else
						{
							str = str.Left(n);
							m_pDocSection->InsertSymbol(ascii, str);
						}
					}
					oFont.SetItalic(long(0));
				}

				n = strTxt.Find(L")");
				LeftStr = strTxt.Left(n);
				strTxt = strTxt.Mid(n + 1);
				n = LeftStr.Find(L",");
				if(n < 0)
					str = LeftStr;
				else
				{
					str = LeftStr.Left(n);
					str1 = LeftStr.Mid(n + 1);
				}
				sym = _ttoi(str1);
				//if(ascii > 1000 || ascii < -1000)
				//{
				if(sym < 1000 && sym > -1000 )
					WriteSuperScript2(str,str1);
				else
					WriteSuperScript2(str,sym);


			}
			else if( strTxt.Left(25) == "WriteSjqySSuuppeerrScript")
			{
				n = strTxt.Find(L"(") + 1;
				strTxt = strTxt.Mid(n);
				if( strTxt.Left(21) == "WriteSjqySSyymmbBOOLl")
				{
					n = strTxt.Find(L"(") + 1;
					strTxt = strTxt.Mid(n);
					n = strTxt.Find(L")");
					LeftStr = strTxt.Left(n);
					strTxt = strTxt.Mid(n + 1);
					n = LeftStr.Find(L",");
					if(n < 0)
					{
						ascii = _ttoi(LeftStr);
						m_pDocSection->InsertSymbol(ascii,L"Symbol");
					}
					else
					{
						str = LeftStr.Left(n);
						ascii = _ttoi(str);
						str = LeftStr.Mid(n + 1);
						n = str.Find(L",");
						if(n < 0)
						{
							m_pDocSection->InsertSymbol(ascii,L"Symbol");
						}
						else
						{
							str = str.Left(n);
							m_pDocSection->InsertSymbol(ascii, str);
						}
					}
				}

				n = strTxt.Find(L")");
				LeftStr = strTxt.Left(n);
				strTxt = strTxt.Mid(n + 1);
				n = LeftStr.Find(L",");
				if(n < 0)
					str = LeftStr;
				else
				{
					str = LeftStr.Left(n);
					str1 = LeftStr.Mid(n + 1);
				}
				ascii = _ttoi(str);
				sym = _ttoi(str1);
				if(ascii > 1000 || ascii < -1000)
				{
					if(sym < 1000 && sym > -1000 )
						WriteSuperScript(ascii,str1);
					else
						WriteSuperScript(ascii,sym);
				}
				else
				{
					if(sym < 1000 && sym > -1000 )
						WriteSuperScript(str,str1);
					else
						WriteSuperScript(str,sym);
				}
			}
			else if( strTxt.Left(21) == "WriteSjqySSyymmbBOOLl")
			{
				n = strTxt.Find(L"(") + 1;
				strTxt = strTxt.Mid(n);
				n = strTxt.Find(L")");
				LeftStr = strTxt.Left(n);
				strTxt = strTxt.Mid(n + 1);
				n = LeftStr.Find(L",");
				oFont.SetSuperscript(FALSE);
				oFont.SetSubscript(FALSE);
				if(n < 0)
				{
					ascii = _ttoi(LeftStr);
					m_pDocSection->InsertSymbol(ascii,L"Symbol");
				}
				else
				{
					str = LeftStr.Left(n);
					ascii = _ttoi(str);
					str = LeftStr.Mid(n + 1);
					n = str.Find(L",");
					if(n < 0)
					{
						m_pDocSection->InsertSymbol(ascii, str);
					}
					else
					{
						LeftStr = str.Left(n);
						str = str.Mid(n + 1);
						long Bias = _ttol(str);
						m_pDocSection->InsertSymbol(ascii, LeftStr,Bias);
					}
				}
			}
			else
				m_pDocSection->TypeText(strTxt);
			n = strTxt.Find(L"WriteSjqy");
		}
		m_pDocSection->TypeText(strTxt);
		
	}

}

void CMsWordTool::WriteUnderLine( LPCTSTR lpszText )
{
	_Font oFont = m_pDocSection->GetFont();
	oFont.SetBold(long(1));
	m_pDocSection->TypeText(lpszText);
	oFont.SetBold(long(0));

}

_MITC_BASIC_END
