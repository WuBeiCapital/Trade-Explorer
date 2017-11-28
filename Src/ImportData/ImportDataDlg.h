
// ImportDataDlg.h : ���^�n
//

#pragma once


// CImportDataDlg ��Ԓ���K
class CImportDataDlg : public CDialogEx
{
// ����
public:
	CImportDataDlg(CWnd* pParent = NULL);	// �˜ʽ�����ʽ

// ��Ԓ���K�Y��
	enum { IDD = IDD_IMPORTDATA_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧Ԯ

	// ��ʽ�a����

	CComboBox m_cmbState;
	CComboBox m_cmbTimePro;
	CEdit m_edtTime;

	UINT m_uTimePro;
	UINT m_uState;
	UINT m_uTime;

	void Dlg2Data();
protected:
	HICON m_hIcon;

	CString GetDataPath();
	CString GetDataHistoryPath();
	BOOL Anasylis_factor();
		
	// �a����ӍϢ������ʽ
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
	
	map<UINT,CString> m_mapHK2Alists;
public:
	BOOL GetDataA2HK();
	BOOL Select();
	BOOL CreateA2HKList();
	
	//!Query
	CString GetNumberByID(UINT uID) const;//HK2A
	UINT GetIDByNumber(const CString& strNumber) const;//A2HK
	//!mapSrcData ��ѯ���ɣ�A�ɱ��롢��ѯ���ʱ�䣩��ĳһʱ�����״γ���ָ��������ʱ��ڵ�;���û�����������ľ�û�У�
	BOOL ExcuteQueryTime(const map<CString,CString>& mapSrcData,const ConditionItem& cdItem,map<CString,CString>& mapDecData);

	afx_msg void OnBnClickedBtnImport();
	afx_msg void OnBnUpdateA2HKData();
	afx_msg void OnBnCreateA2HKList();
	afx_msg void OnCbnSelchangeCmbState();
	afx_msg void OnBnClickedBtnAnasy();
};
