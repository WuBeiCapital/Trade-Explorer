#pragma once
//ini文件读写类
_MITC_BASIC_BEGIN
class MITC_BASIC_EXT IniFileReadWrite
{
public:
	BOOL ReadAllSections(vector<CString>& vctSections);
  //写ini文件,默认文件在当前程序运行目录下,文件名为set.ini
  BOOL WriteValue(LPCTSTR lpszSection,LPCTSTR lpszSectionKey,LPCTSTR lpszValue);
  BOOL WriteValue(LPCTSTR lpszSection,LPCTSTR lpszSectionKey,int iValue);
  BOOL WriteValue(LPCTSTR lpszSection,LPCTSTR lpszSectionKey,double dValue);
  //读ini文件,默认文件在当前程序运行目录下,文件名为set.ini
  BOOL ReadValue(LPCTSTR lpszSection,LPCTSTR lpszSectionKey,CString &szValue);
  BOOL ReadValue(CString lpszSection,LPCTSTR lpszSectionKey,int &iValue);
  BOOL ReadValue(CString lpszSection,LPCTSTR lpszSectionKey,double &dValue);
  //更改文件路径
  void SetPathFileName(LPCTSTR lpszPathFileName);
  //构造函数
  IniFileReadWrite(LPCTSTR lpszPathFileName);
  IniFileReadWrite(void);
  ~IniFileReadWrite(void);
private:
  //文件路径及文件
  CString m_sPathFileName;
};
_MITC_BASIC_END