#pragma once

class   _Document ;
class  Range ;
 class Find ;   
 class Replacement ;
 class Selection;
 class _Font;
 class CCreatePic;

_MITC_BASIC_BEGIN
class MITC_BASIC_EXT CTextFormat
{
public:
  CTextFormat();
  
  ~CTextFormat();

  CString strContent;
  bool bBold;
  int nStep;
  double nFontSize;
  CString FontName;
  COLORREF FontColor;
  bool IsHaveFisrtSigle;
  int  LeftIndent;       //左端悬挂缩进字符数
  int  iOutLineLevel;    //大纲级别，0-正文 1-9

  bool IsInsertPic ;     //是否为插入图片
  CString strPicPath;    // IsInsertPic = TRUE 插入图片的路径
  CString strPicExpress ;// IsInsertPic = TRUE 插入图片的描述，通常写在图片下方
protected:
private:
};

class CMsWordTooldecorator;

class  MITC_BASIC_EXT CMsWordTool
{
public: 
  CMsWordTool(CMsWordTooldecorator* pMsWordTooldecorator,CString strTemplatePath);
public:
  ~CMsWordTool(void); 
protected:
  CMsWordTool(void);
private:
  CString m_pDocPath;
  _Document* m_pDoc;
  Range* m_pRange;
  Find* m_pFndInDoc;   
  Replacement* m_pRpInDoc; 

  //这个变量为临时存储,如果重新赋值，直接调用GetSelection
  Selection* m_pDocSection;

public:
  void SetDocName(CString docName);
  void ReplaceParameter(CString src,CString data);
  void Delete4BookMarks(CString bm1,CString bm2);
  void InsertPicture(CString strbookmark,CString path);
  void InsertPicture(CString path);

  void FillTable(CString strBookMark,CArray<CString,CString>& ContentList,int nColumns);
  void FillvctTable(CString strBookMark,const std::vector<CString>& ContentList,int nColumns);

  void FillMutiTable(CString strBookMark,CArray<CString,CString>& ContentList,int nRows,int nColumns,int iMutiStartColumn,int iMutiEndColumn,int InnerLines);

  void loopSelection(CString bm1,CString bm2,CArray<CString,CString>& ReplaceList,CArray<CString,CString>& ContentList);
  void UpdateAllDomainCode();
  void SaveAs(CString SaveAsPath =_T(""),CString SaveAsNewDocName = _T(""));
  void ExitWordApp();
  void OpenDoc(LPCTSTR SaveAsPath = _T(""),LPCTSTR SaveAsNewDocName = _T(""));
  //tool
  static CString DoubleToString(double src,int nAfterDot);   //nAfterDot 小数点后几位
  static CString IntToString(int src);
  void WriteText(CString strBookMark,CTextFormat& TextFormat,int nAlignment = 0);
  void WriteText(CString strBookMark,std::vector<CTextFormat*>& strContent,int nAlignment = 0);
  void WriteTextHaveEquation(LPCTSTR strTxt);
  void WriteTextHaveEquation(CString strTxt);

  void WriteTextBySection(CTextFormat& TextFormat,int nAlignment = 0);
  void WriteTextBySection(std::vector<CTextFormat*>& strContent,int nAlignment = 0);

  //!0-left 1-center 2-right
  void SetParaphFormat(int nAlignment = 1,int nLeftIndent = 0,int nFirstLineIndent = 0);
  void SetOutLineLevel(int nlevel = 10);
  void SetFontStyle(double nFontSize,COLORREF FontColor=RGB(0,0,0),CString FontName = _T("宋体"),BOOL bBold = FALSE, BOOL bItalic = FALSE , BOOL bUnderLine = FALSE);

  //返回一个正文的Selection
  void WriteTitle(CString strTitle,int nStep = 1);

  bool AddTable(int nRows,int nColumns,vector<CString>& vctContent,vector<double>& vctColumWidth,int nAlignment = 1,double FontSize = 10.5);

  bool GetSelection(CString strBookMark);

  CString CreateCalBookGraph(const map<CGraphKey,double>& srcData,const CString& strX_Name,const CString& strY_Name,const map<CGraphKey,CString>& strTLS,const CString& picName);

  void SetHeader(const CString& strHeader);
  void SetFooter(const CString& strFooter);

  void SetEnter();

  void WriteSubScript(LPCTSTR lpszText, LPCTSTR lpszSub);
  void WriteSubScript(long ascii, LPCTSTR lpszSub);
  void WriteSubScript(long ascii, long sub);
  void WriteSubScript(LPCTSTR lpszText,long ascii);
  void WriteSubScript1(LPCTSTR lpszText, LPCTSTR lpszSub);
  void WriteSubScript1(long ascii, LPCTSTR lpszSub);
  void WriteSubScript1(long ascii, long sub);
  void WriteSubScript1(LPCTSTR lpszText,long ascii);
  void WriteSubScript2(LPCTSTR lpszText, LPCTSTR lpszSub);
  void WriteSubScript2(long ascii, LPCTSTR lpszSub);
  void WriteSubScript2(long ascii, long sub);
  void WriteSubScript2(LPCTSTR lpszText,long ascii);

  void WriteSuperScript(LPCTSTR lpszText, LPCTSTR lpszSuper);
  void WriteSuperScript(LPCTSTR lpszText, long ascii);
  void WriteSuperScript(long ascii, LPCTSTR lpszSuper);
  void WriteSuperScript(long ascii, long sup);
  void WriteSuperScript1(LPCTSTR lpszText, LPCTSTR lpszSuper);
  void WriteSuperScript1(LPCTSTR lpszText, long ascii);
  void WriteSuperScript1(long ascii, LPCTSTR lpszSuper);
  void WriteSuperScript1(long ascii, long sup);
  void WriteSuperScript2(LPCTSTR lpszText, LPCTSTR lpszSuper);
  void WriteSuperScript2(LPCTSTR lpszText, long ascii);
  void WriteSuperScript2(long ascii, LPCTSTR lpszSuper);
  void WriteSuperScript2(long ascii, long sup);

  void WriteUnderLine(LPCTSTR lpszText);

  void WriteEquation(LPCTSTR strTxt);
  void WriteEquation(CString strTxt);

private:
  void InitializeWordApp();
  bool GetSelection4bm(Selection& decSelection,CString bm1,CString bm2);
  CMsWordTooldecorator* m_pMsWordTooldecorator;
  CCreatePic* m_pCreatePic;
};

_MITC_BASIC_END
