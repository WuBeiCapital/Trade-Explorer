#pragma once

class _Document;
class _Application;
class Documents;


_MITC_BASIC_BEGIN

class CMsWordTool;

class MITC_BASIC_EXT CMsWordTooldecorator
{
public:
  CMsWordTooldecorator(void);
  _Application* GetApp();
  Documents* GetDocuments();
  CMsWordTool* CreateWordTool(LPCTSTR docPath);
  void SetDocuments(Documents* pDocument);
  void CloseWordApplication();
public:
  ~CMsWordTooldecorator(void);
private: 
  Documents* m_pDocs; 
  _Application* m_pApp; 
public:
private:
};

_MITC_BASIC_END