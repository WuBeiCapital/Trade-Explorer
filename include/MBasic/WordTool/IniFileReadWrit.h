#pragma once
//ini�ļ���д��
_MITC_BASIC_BEGIN
class MITC_BASIC_EXT IniFileReadWrite
{
public:
	BOOL ReadAllSections(vector<CString>& vctSections);
  //дini�ļ�,Ĭ���ļ��ڵ�ǰ��������Ŀ¼��,�ļ���Ϊset.ini
  BOOL WriteValue(LPCTSTR lpszSection,LPCTSTR lpszSectionKey,LPCTSTR lpszValue);
  BOOL WriteValue(LPCTSTR lpszSection,LPCTSTR lpszSectionKey,int iValue);
  BOOL WriteValue(LPCTSTR lpszSection,LPCTSTR lpszSectionKey,double dValue);
  //��ini�ļ�,Ĭ���ļ��ڵ�ǰ��������Ŀ¼��,�ļ���Ϊset.ini
  BOOL ReadValue(LPCTSTR lpszSection,LPCTSTR lpszSectionKey,CString &szValue);
  BOOL ReadValue(CString lpszSection,LPCTSTR lpszSectionKey,int &iValue);
  BOOL ReadValue(CString lpszSection,LPCTSTR lpszSectionKey,double &dValue);
  //�����ļ�·��
  void SetPathFileName(LPCTSTR lpszPathFileName);
  //���캯��
  IniFileReadWrite(LPCTSTR lpszPathFileName);
  IniFileReadWrite(void);
  ~IniFileReadWrite(void);
private:
  //�ļ�·�����ļ�
  CString m_sPathFileName;
};
_MITC_BASIC_END