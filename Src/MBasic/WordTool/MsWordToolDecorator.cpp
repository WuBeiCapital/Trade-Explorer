#include "StdAfx.h"
#include "MsWordTooldecorator.h"


_MITC_BASIC_BEGIN

CMsWordTooldecorator::CMsWordTooldecorator(void)
{	
 
  m_pApp=new _Application;

  COleException* exc = new COleException;

  if(!m_pApp->CreateDispatch(_T("Word.Application"), exc))
  {
    AfxMessageBox(_T("创建Word.Application服务失败!")); 
    throw exc;

  }
  else
    m_pApp->SetVisible(FALSE);	
}

CMsWordTooldecorator::~CMsWordTooldecorator(void)
{
	if(m_pApp)
	{
		CloseWordApplication();
	}
}

void CMsWordTooldecorator::CloseWordApplication()
{
	CComVariant SaveChanges(false),OriginalFormat,RouteDocument;  
  if(m_pApp)
  {
	  m_pApp->Quit(&SaveChanges,&OriginalFormat,&RouteDocument);	
	  m_pApp->ReleaseDispatch();
  }
}

CMsWordTool* CMsWordTooldecorator::CreateWordTool(LPCTSTR docPath)
{
	return new CMsWordTool(this,docPath);
}

_Application* CMsWordTooldecorator::GetApp()
{
	return m_pApp;
}

Documents* CMsWordTooldecorator::GetDocuments()
{
	return m_pDocs;
}
void CMsWordTooldecorator::SetDocuments(Documents* pDocument)
{
  m_pDocs = pDocument;
}

_MITC_BASIC_END