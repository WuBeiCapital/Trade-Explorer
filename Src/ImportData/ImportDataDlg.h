
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

	BOOL IsFindTagFromExcel(CXLControl* pCXLControl,const CString strTag,int& nReturnRow,int& nReturnCol);//��excel�ļ��в��ұ��Force
	BOOL IsGetRowDataFromExcel(CXLControl* pCXLControl,CArray<CStockData>& arRowData,int nRow,int nCol1,int nCol2);
	CString GetInsertSqlString(CArray<CStockData>& arRowData,const CString& strName);
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
	
	// �a����ӍϢ������ʽ
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	BOOL GetDataA2HK();
	BOOL Select();
	BOOL CreateA2HKList();
	afx_msg void OnBnClickedBtnImport();
	afx_msg void OnBnUpdateA2HKData();
	afx_msg void OnBnCreateA2HKList();
	afx_msg void OnBnClickedButton2();
	afx_msg void OnCbnSelchangeCmbState();
};
