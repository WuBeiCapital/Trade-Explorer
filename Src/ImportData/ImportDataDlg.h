
// ImportDataDlg.h : 標頭檔
//

#pragma once


// CImportDataDlg 對話方塊
class CImportDataDlg : public CDialogEx
{
// 建構
public:
	CImportDataDlg(CWnd* pParent = NULL);	// 標準建構函式

// 對話方塊資料
	enum { IDD = IDD_IMPORTDATA_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支援

	// 程式碼實作

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
		
	// 產生的訊息對應函式
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
	afx_msg void OnCbnSelchangeCmbState();
	afx_msg void OnBnClickedBtnAnasy();
};
