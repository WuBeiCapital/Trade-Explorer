
// ImportDataDlg.cpp : 實作檔
//

#include "stdafx.h"
#include "ImportData.h"
#include "ImportDataDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


LPCTSTR lpcsCode = _T("股票代码");//股票代码
LPCTSTR lpcsName = _T("股票名称");//股票名称
LPCTSTR lpcsCount = _T("持股数量");//持股数量
LPCTSTR lpcsFactor = _T("持股比例");//持股比例
LPCTSTR lpcsChange = _T("变化比例");//

// 對 App About 使用 CAboutDlg 對話方塊

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 對話方塊資料
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支援

// 程式碼實作
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CImportDataDlg 對話方塊



CImportDataDlg::CImportDataDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CImportDataDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);

	CString strPath=GetDataHistoryPath();
	SQLite sqlite; 
    // 打开或创建数据库
    //******************************************************  
    if(!sqlite.Open(strPath))  
    {   
		strPath+=_T("_文件打开错误!");
		AfxMessageBox(strPath);
        return ;  
    } 
	//!获取港股通列表
	double dFactor,dFactorTmp;
	CString strSql;
	strSql=_T("select * from A2HK");	
	SQLiteDataReader Reader = sqlite.ExcuteQuery(strSql); 
	CString strNumber;
	UINT uID;
	while(Reader.Read()) 
    { 	
		uID=Reader.GetInt64Value(0);
		strNumber=Reader.GetStringValue(2);	
		m_mapHK2Alists[uID]=strNumber;
	}  
    Reader.Close();
}

void CImportDataDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);

  DDX_Control(pDX, IDC_CMB_STATE,m_cmbState);
  DDX_Control(pDX, IDC_CMB_TIME,m_cmbTimePro);
  DDX_Control(pDX, IDC_EDT_TIME,m_edtTime);
}

BEGIN_MESSAGE_MAP(CImportDataDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BTN_Import, &CImportDataDlg::OnBnClickedBtnImport)
	ON_BN_CLICKED(IDC_BTN_Export, &CImportDataDlg::OnBnUpdateA2HKData)
	ON_BN_CLICKED(IDC_BUTTON1, &CImportDataDlg::OnBnCreateA2HKList)
	ON_CBN_SELCHANGE(IDC_CMB_STATE, &CImportDataDlg::OnCbnSelchangeCmbState)
	ON_BN_CLICKED(IDC_BTN_ANASY, &CImportDataDlg::OnBnClickedBtnAnasy)
END_MESSAGE_MAP()


// CImportDataDlg 訊息處理常式

BOOL CImportDataDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 將 [關於...] 功能表加入系統功能表。

	// IDM_ABOUTBOX 必須在系統命令範圍之中。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 設定此對話方塊的圖示。當應用程式的主視窗不是對話方塊時，
	// 框架會自動從事此作業
	SetIcon(m_hIcon, TRUE);			// 設定大圖示
	SetIcon(m_hIcon, FALSE);		// 設定小圖示

	// TODO: 在此加入額外的初始設定
	m_cmbState.ResetContent();
	m_cmbState.AddString(_T("连续增仓"));
	m_cmbState.AddString(_T("连续减仓"));
	m_cmbState.SetCurSel(0);

	m_cmbTimePro.ResetContent();
	m_cmbTimePro.AddString(_T("日"));
	m_cmbTimePro.AddString(_T("周"));
	m_cmbTimePro.SetCurSel(0);

	m_edtTime.SetWindowTextW(_T("3"));

	return TRUE;  // 傳回 TRUE，除非您對控制項設定焦點
}

void CImportDataDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果將最小化按鈕加入您的對話方塊，您需要下列的程式碼，
// 以便繪製圖示。對於使用文件/檢視模式的 MFC 應用程式，
// 框架會自動完成此作業。

void CImportDataDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 繪製的裝置內容

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 將圖示置中於用戶端矩形
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 描繪圖示
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// 當使用者拖曳最小化視窗時，
// 系統呼叫這個功能取得游標顯示。
HCURSOR CImportDataDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CImportDataDlg::OnBnClickedBtnImport()
{
	// TODO: Add your control notification handler code here	
	Select();//!	
}

void CImportDataDlg::OnBnUpdateA2HKData()
{
	// TODO: Add your control notification handler code here
	GetDataA2HK();
}

void CImportDataDlg::OnBnCreateA2HKList()
{
	// TODO: Add your control notification handler code here
	//!
	CreateA2HKList();	
}
CString CImportDataDlg::GetNumberByID(UINT uID) const//HK2A
{
	CString strNumber(_T(""));
	map<UINT,CString>::const_iterator p=m_mapHK2Alists.find(uID);
	if(p!=m_mapHK2Alists.end())
	{
		strNumber=(*p).second;
	}
	return strNumber;
}
UINT CImportDataDlg::GetIDByNumber(const CString& strNumber) const//A2HK
{
	UINT uID=0;
	for(map<UINT,CString>::const_iterator p=m_mapHK2Alists.begin();p!=m_mapHK2Alists.end();++p)
	{//!
		if((*p).second==strNumber)
			uID=(*p).first;
	}

	return uID;
}
//!mapSrcData 查询个股（A股编码、查询起点时间）自某一时间点后，首次出现指定条件的时间节点;如果没有满足条件的就没有；
BOOL CImportDataDlg::ExcuteQueryTime(const map<CString,CString>& mapSrcData,const ConditionItem& cdItem,map<CString,CString>& mapDecData)
{//!
	//!
	CString strPath=GetDataPath();
    SQLite sqlite;  
    // 打开或创建数据库
    //******************************************************  
    if(!sqlite.Open(strPath))  
    {  
       // _tprintf(_T("%s\n"),sqlite.GetLastErrorMsg());  
		strPath+=_T("_文件打开错误!");
		AfxMessageBox(strPath);
        return 0;  
    } 
	//!获取列表
	TCHAR sql[512] = {0}; 
    memset(sql,0,sizeof(sql));  
	/////////////////////////////////////////////////////////////////////////////////////////
	//!策略//////////////////////////////////////////////////////////////////////////////////
	/////////////////////////////////////////////////////////////////////////////////////////
	//!
	UINT uTimeContinue=cdItem.m_uContinueCount;//!连续
	BOOL bStateType=!cdItem.m_uTypeUporLow;//!TRUE，增仓；FALSE，减仓
	UINT bTimeType=cdItem.m_uTimeType;//!0、天；1、周；2、月
	//!
	uTimeContinue=m_uTime;

	switch(bStateType)
	{
		case 0:
			bStateType=FALSE;
			break;
		case 1:
			bStateType=TRUE;
			break;
		default:
			break;
	}
	//!
	switch(bTimeType)
	{
		case 0:
			bTimeType=0;
			break;
		case 1:
			bTimeType=1;
			break;
		case 2:
			bTimeType=2;
			break;
		default:
			break;
	}
	//!
	/////////////////////////////////////////////////////////////////////////////////////////
	//分解策略////////////////////////////////////////////////////////////////////////////////
	/////////////////////////////////////////////////////////////////////////////////////////
	//!1、解读时间；2、需要搜索的表名
	SYSTEMTIME sys; 
	GetLocalTime(&sys);	
	vector<CString> vctList;
	int y=sys.wYear,m=sys.wMonth, d=sys.wDay;

	UINT uW=6;
	CString strTmp,strName;
	//switch(bTimeType)
	//{
	//	case 0://day
	//		for(int i=0;i <uTimeContinue+1;++i)
	//		{//!
	//			strTmp=_T("");
	//			strName=_T("");		
	//			do
	//			{//!		
	//				if(d>1)
	//				{//!
	//					d-=1;
	//				}
	//				else
	//				{//!
	//					if(m>1)
	//					{
	//						m-=1;
	//						d=daysOfMonth(y,m);
	//					}
	//					else
	//					{			
	//						y-=1;
	//						m=12;
	//						d=daysOfMonth(y,m);
	//					}
	//				}
	//				uW=CaculateWeekDay(y,m,d);
	//			}while(uW>5);

	//			strName=GetTimeString(y,m,d);
	//			strName=_T("A")+strName;
	//			vctList.push_back(strName);
	//		}
	//		break;
 //       case 1://week
	//		//！
	//		for(int i=0;i <uTimeContinue+1;++i)
	//		{//!
	//			strTmp=_T("");
	//			strName=_T("");		
	//			
	//			uW=CaculateWeekDay(y,m,d);
	//			if(uW>=5)
	//			{
	//				//!调整到本周五
	//				d-=uW-5;
	//			}
	//			else
	//			{
	//				//!调整到上周五
	//				d-=uW+2;
	//			}
	//			//！
	//			if(d<0)
	//			{//!
	//				if(m>1)
	//				{
	//					m-=1;
	//					d=daysOfMonth(y,m)+d;
	//				}
	//				else
	//				{			
	//					y-=1;
	//					m=12;
	//					d=daysOfMonth(y,m)+d;
	//				}
	//			}					
	//			//uW=CaculateWeekDay(y,m,d);
	//			strName=GetTimeString(y,m,d);
	//			strName=_T("A")+strName;
	//			vctList.push_back(strName);
	//			d-=7;
	//		}
	//		break;
	//	default:
	//		break;
	//}	
	//_ASSERTE(vctList.size());
	////！用最近一份数据建立比较样本
	double dFactor,dFactorTmp;
	CString strSql;
	////strTmp.Format(_T("%d"),uID);
	//strSql=_T("select * from ");
	//strSql+=vctList.at(0);
	////strSql+=_T(" where id = ")+strTmp;
	//SQLiteDataReader Reader = sqlite.ExcuteQuery(strSql); 
	//map<UINT,double> mapList;//id  factor
	UINT uID,uNumber;
	//while(Reader.Read()) 
 //   {  
	//	uID=Reader.GetInt64Value(0);	
	//	dFactor=Reader.GetFloatValue(3);	
	//	mapList[uID]=dFactor;
 //   }  
 //   Reader.Close();
	//！用比较样本 去遍历符合条件的表，记录符合策略的ID；
	UINT uCount=0;
	vector<UINT> vctIDs;	
	//const map<CString,CString>& mapSrcData,const ConditionItem& cdItem,map<CString,CString>& mapDecData
	//for(map<UINT,double>::iterator p=mapList.begin();p!=mapList.end();++p)
	//{//!		
		//uID=(*p).first;
		//dFactor=(*p).second;
		for(map<CString,CString>::const_iterator pIt=mapSrcData.begin(); pIt!=mapSrcData.end();++pIt)
		{
			uID=GetIDByNumber((*pIt).first);
			strTmp.Format(_T("%d"),uID);
			strSql=_T("select * from ");
			strSql+=_T("A")+(*pIt).second;
			strSql+=_T(" where id = ")+strTmp;
			//!
			SQLiteDataReader Reader = sqlite.ExcuteQuery(strSql);	
			while(Reader.Read()) 
			{  	
				dFactorTmp=Reader.GetFloatValue(3);
				if(bStateType)//!
				{
					if(dFactorTmp < dFactor)
						uCount++;					
				}
				else
				{
					if(dFactorTmp > dFactor)
						uCount++;			
				}
				dFactor=dFactorTmp;
			} 
			Reader.Close();
		}
		if(uCount==uTimeContinue)
			vctIDs.push_back(uID);
	//}
	//!	
	//！用符合策略的ID，去获取最新数据，建立数据集；
	map<UINT,vector<CHKStockData>> mapDatas;
	for(vector<UINT>::iterator pIt=vctIDs.begin(); pIt!=vctIDs.end();++pIt)
	{
		uID=(*pIt);
		vector<CHKStockData> vctstockDatas;
		for(vector<CString>::iterator pIt=vctList.begin(); pIt!=vctList.end();++pIt)
		{
			strTmp.Format(_T("%d"),uID);
			strSql=_T("select * from ");
			strSql+=(*pIt);
			strSql+=_T(" where id = ")+strTmp;
			//!
			SQLiteDataReader Reader = sqlite.ExcuteQuery(strSql);	
			while(Reader.Read()) 
			{  		
				CHKStockData tmp;
				tmp.SetTime((*pIt));
				tmp.SetCode(Reader.GetInt64Value(0));
				tmp.SetName(Reader.GetStringValue(1));
				tmp.SetCount(Reader.GetInt64Value(2));
				tmp.SetFactor(Reader.GetFloatValue(3));		
				vctstockDatas.push_back(tmp);
			} 
			Reader.Close();
		}
		mapDatas[uID]=vctstockDatas;
	}  
	// 关闭数据库  
    sqlite.Close(); 
	//!获取行情数据


	return TRUE;
}

BOOL CImportDataDlg::CreateA2HKList()
{//!
	//!
	TCHAR *szDbPath =_T("D:\\AHistory2011.db");// _T("D:\\test.db");CString strPath=_T("D:\\AHistory2011.db");
	//::DeleteFile(szDbPath);   
    SQLite sqlite;  
  
    // 打开或创建数据库  
    //******************************************************  
    if(!sqlite.Open(szDbPath))  
    {  
        _tprintf(_T("%s\n"),sqlite.GetLastErrorMsg());  
        return 0;  
    }  
    ////******************************************************    
    //// 创建数据库表  
    ////******************************************************  
    TCHAR sql[512] = {0};  
    
    //// 插入数据【普通方式】  
    DWORD dwBeginTick = GetTickCount();   
  
    // 查询  
    dwBeginTick = GetTickCount();  
    //******************************************************  
    memset(sql,0,sizeof(sql));  
    _stprintf_s(sql,_T("%s"),_T("select * from Ａ20171112"));  
  
	int index = 0;  
    int len = 0;  
	CString strName,strValueCode,strValueName;
	__int64 uValue;
	double dValue;

    SQLiteDataReader Reader = sqlite.ExcuteQuery(sql); 
	map<CString,pair<CString,__int64>> mapList;
	pair<CString,__int64> pairTmp;
	while(Reader.Read()) 
    {  
        _tprintf( _T("***************【第%d条记录】***************\n"),++index);
		strName=Reader.GetName(0);
		strValueCode=Reader.GetStringValue(0);

		strName=Reader.GetName(1);
		strValueName=Reader.GetStringValue(1);
		pairTmp.first=strValueCode;
		pairTmp.second=0;
		strValueName.Replace(_T(" "),_T(""));
		mapList[strValueName]=pairTmp;
    }  
    Reader.Close();
	_stprintf_s(sql,_T("%s"),_T("select * from Book")); // order by 1
	//_stprintf(sql,_T("%s"),_T("select * from Book where name = '海康威视'")); 
	SQLiteDataReader Reader2 = sqlite.ExcuteQuery(sql);	
    while(Reader2.Read()) 
    {  
        _tprintf( _T("***************【第%d条记录】***************\n"),++index);
		strName=Reader2.GetName(0);
		uValue=Reader2.GetInt64Value(0);

		strName=Reader2.GetName(1);
		strValueName=Reader2.GetStringValue(1);

		//strName=Reader.GetName(2);
		//uValue=Reader.GetInt64Value(2);

		//strName=Reader.GetName(3);
		//dValue=Reader.GetFloatValue(3);
		strValueName.Replace(_T(" "),_T(""));
		if(mapList.find(strValueName)!=mapList.end())
		{//!			
			pairTmp=mapList[strValueName];
			pairTmp.second=uValue;
			mapList[strValueName]=pairTmp;
		}
    }  
    Reader2.Close();

	memset(sql,0,sizeof(sql));
	 _stprintf_s(sql,_T("%s"),  
        _T("CREATE TABLE [List] (")  
        _T("[id] INTEGER NOT NULL PRIMARY KEY, ")  
        _T("[name] NVARCHAR(20), ")  
        _T("[number] NVARCHAR(20)); ") 
        );  
    if(!sqlite.ExcuteNonQuery(sql))  
    {  
        printf("Create database table failed...\n");  
    }
	else
	{//!
  	  //// 当一次性插入多条记录时候，采用事务的方式，提高效率  
		sqlite.BeginTransaction();  
		memset(sql,0,sizeof(sql));  
		_stprintf_s(sql,_T("insert into List(id,name,number) values(?,?,?)"));  
		SQLiteCommand cmd(&sqlite,sql);  
		// 批量插入数据  
		for(map<CString,pair<CString,__int64>>::iterator p=mapList.begin();p!=mapList.end();++p)  
		{  
			if((*p).second.second>0)
			{
				TCHAR strValue[16] = {0};  
				_stprintf_s(strValue,_T("%d"),(*p).second.second);  
				// 绑定第一个参数（id字段值）  
				cmd.BindParam(1,strValue);  
				// 绑定第二个参数（name字段值）  
				cmd.BindParam(2,((*p).first));  		
				// 绑定第三个参数（number字段值） 
				//_stprintf_s(strValue,_T("%d"),(*p).second.first); 
				cmd.BindParam(3,(*p).second.first);  
				if(!sqlite.ExcuteNonQuery(&cmd))  
				{  
					_tprintf(_T("%s\n"),sqlite.GetLastErrorMsg());  
					break;  
				} 
			}		   
		}  
		// 清空cmd  
		cmd.Clear();  
		// 提交事务  
		sqlite.CommitTransaction();  
	}
    //printf("Query Take %dMS...\n",GetTickCount()-dwBeginTick);  
    //******************************************************  
  
    // 关闭数据库  
    sqlite.Close(); 

	return 0;
}
void CImportDataDlg::Dlg2Data()
{//!
	UpdateData(FALSE);
	//!
	m_uState=m_cmbState.GetCurSel();
	m_uTimePro=m_cmbTimePro.GetCurSel();
	CString strText;
	m_edtTime.GetWindowTextW(strText);
	//!
	m_uTime=_tstoi(strText);
}
CString CImportDataDlg::GetDataPath()
{
	CString strPath=GetSystemPath();
	int index=strPath.Find(_T("Bin"));
	if(index!=-1)
	{//!		
		strPath=strPath.Left(index);
		strPath+=_T("Data\\A2HK2017.db");
	}
	else
	{
		index=strPath.Find(_T("bin"));
		if(index!=-1)
		{
			strPath=strPath.Left(index);
			strPath+=_T("Data\\A2HK2017.db");
		}
	}
	return strPath;
}
CString CImportDataDlg::GetDataHistoryPath()
{
	CString strPath=GetSystemPath();
	int index=strPath.Find(_T("Bin"));
	if(index!=-1)
	{//!		
		strPath=strPath.Left(index);
		strPath+=_T("Data\\AHistory2017.db");
	}
	else
	{
		index=strPath.Find(_T("bin"));
		if(index!=-1)
		{
			strPath=strPath.Left(index);
			strPath+=_T("Data\\AHistory2017.db");
		}
	}

	return strPath;
}

BOOL CImportDataDlg::GetDataA2HK()//
{//!
	int ret;
    ret = gm_login("13480922739", "a7612006");
	if (ret != 0)
	{
		printf("login fail");
		return ret;
	}

	CString strPath=GetDataHistoryPath();//_T("D:\\AHistory2017.db");
	SQLite sqlite; 
    // 打开或创建数据库
    //******************************************************  
    if(!sqlite.Open(strPath))  
    {  
       // _tprintf(_T("%s\n"),sqlite.GetLastErrorMsg());  
		strPath+=_T("_文件打开错误!");
		AfxMessageBox(strPath);
        return 0;  
    } 
	//!获取港股通列表
	double dFactor,dFactorTmp;
	CString strSql;
	//strTmp.Format(_T("%d"),uID);
	strSql=_T("select * from A2HK");	
	//strSql+=_T(" where id = ")+strTmp;
	SQLiteDataReader Reader = sqlite.ExcuteQuery(strSql); 
	vector<CString> vctA2HKlist;
	CString strNumber;
	while(Reader.Read()) 
    { 		
		strNumber=Reader.GetStringValue(2);	
		vctA2HKlist.push_back(strNumber);
    }  
    Reader.Close();
	for(vector<CString>::iterator p=vctA2HKlist.begin();p!=vctA2HKlist.end();++p)
	{
		//!名称转换
		CString strName=*p;
		int index =strName.Find(_T("6"));
		if(index==0)
		{
			strName=_T("SHSE.")+strName;
		}
		else
			strName=_T("SZSE.")+strName;
		// 先得到要转换为字符的长度
		const size_t strsize=(strName.GetLength()+1)*2; // 宽字符的长度;
		char* pstr= new char[strsize]; //分配空间;
		size_t sz=0;
		wcstombs_s(&sz,pstr,strsize,strName,_TRUNCATE);
		//!获取行情数据
		DailyBar *dbar = nullptr;
		int count = 0;//,SZSE.000001	
		ret = gm_md_get_dailybars((const char*)pstr,"2017-01-01", "2017-11-24",&dbar, &count);
		delete pstr;
		pstr=NULL;

		DailyBar dbData;
		vector<DailyBar> vctDailyBars;
		for(int i=0;i<count;++i,dbar++)
		{//!
			dbData=*dbar;
			vctDailyBars.push_back(dbData);
		}//!
		//!创建表
		//!名称转换
		index=strName.Find(_T("."));
		strName=strName.Right(strName.GetLength()-1-index);
		strName=_T("A")+strName;
		//!先查找
		CString sql = _T("CREATE TABLE [")+strName+_T("] (")+
			 _T("[data] NVARCHAR(20),")+
			 _T("[open] REAL,")+
			 _T("[close] REAL,")+
			 _T("[high] REAL,")+
			 _T("[low] REAL,")+
			 _T("[volume] REAL,")+
			 _T("[amount] REAL,")+
			 _T("[adj_factor] REAL);");
		if(!sqlite.ExcuteNonQuery(sql))  
		{  
			printf("Create database table failed...\n");  
		}
		else
		{//!
  			//// 当一次性插入多条记录时候，采用事务的方式，提高效率  
			sqlite.BeginTransaction();  
			//memset(sqll,0,sizeof(sqll));  
			sql=_T("insert into ")+strName+_T("(data,open,close,high,low,volume,amount,adj_factor) values(?,?,?,?,?,?,?,?)");
			//_stprintf_s(sqll,_T("insert into SHSE.600000(data,open,close,high,low,avgprice) values(?,?,?,?,?,?)"));  
			SQLiteCommand cmd(&sqlite,sql);	
			// 批量插入数据
			for(vector<DailyBar>::iterator p=vctDailyBars.begin();p!=vctDailyBars.end();++p)
			{//!			
				DailyBar ItemTest=(*p);
				//TCHAR strValue[16] = {0};  
				//_stprintf_s(strValue,_T("%d"),(*p).second.second);  
				// 绑定第一个参数（id字段值） data
				CString strTime;
				strTime.Format(_T("%s"),CA2T(ItemTest.strtime));
				int index=strTime.Find(_T("T"));
				if(index!=-1)
				{//!
					strTime=strTime.Left(index);			
				}
				strTime.Replace(_T("-"),_T(""));
				LPCTSTR str = (LPCTSTR)(strTime);//data
				cmd.BindParam(1,str);  
				// 绑定第二个参数（name字段值）  ,open
				double dTmp=ItemTest.open;
				cmd.BindParam(2,dTmp);  		
				// 绑定第三个参数  ,close
				dTmp=ItemTest.close;
				cmd.BindParam(3,dTmp);  
				dTmp=ItemTest.high;//,high
				cmd.BindParam(4,dTmp); 
				dTmp=ItemTest.low;//,low,
				cmd.BindParam(5,dTmp); 
				dTmp=ItemTest.volume;//volume
				cmd.BindParam(6,dTmp); 
				dTmp=ItemTest.amount;//amount,
				cmd.BindParam(7,dTmp); 
				dTmp=ItemTest.adj_factor;//adj_factor
				cmd.BindParam(8,dTmp); 
				if(!sqlite.ExcuteNonQuery(&cmd)) 
				{  
					_tprintf(_T("%s\n"),sqlite.GetLastErrorMsg());  
					break;  
				}
			}  
			// 清空cmd  
			cmd.Clear();  
			// 提交事务  
			sqlite.CommitTransaction();  
		}	
	}	
	// 关闭数据库  
    sqlite.Close(); 
 //   // 设置事件回调函数
	//gm_md_set_tick_callback(on_tick);
	//gm_md_set_bar_callback(on_bar);
	//gm_md_set_error_callback(on_error);
	//gm_md_set_login_callback(on_login);

	//ret = gm_md_init(MD_MODE_LIVE, "SHSE.*.bar.*");

    //初始化失败，退出。
	if (ret)
    {
        printf("gm_md_login return: %d\n", ret); 
        return ret;
	}

    // waiting...
    gm_run();

	return TRUE;
}

BOOL CImportDataDlg::Select()//! 
{//!
	//!
	CString strPath=GetDataPath();
    SQLite sqlite;  
    // 打开或创建数据库
    //******************************************************  
    if(!sqlite.Open(strPath))  
    {  
       // _tprintf(_T("%s\n"),sqlite.GetLastErrorMsg());  
		strPath+=_T("_文件打开错误!");
		AfxMessageBox(strPath);
        return 0;  
    } 
	//!获取列表
	TCHAR sql[512] = {0}; 
    memset(sql,0,sizeof(sql));  
	/////////////////////////////////////////////////////////////////////////////////////////
	//!策略//////////////////////////////////////////////////////////////////////////////////
	/////////////////////////////////////////////////////////////////////////////////////////
	//!
	UINT uTimeContinue=3;//!连续
	BOOL bStateType=FALSE;//!TRUE，增仓；FALSE，减仓
	UINT bTimeType=0;//!0、天；1、周；2、月

	Dlg2Data();	
	//!
	uTimeContinue=m_uTime;

	switch(m_uState)
	{
		case 0:
			bStateType=FALSE;
			break;
		case 1:
			bStateType=TRUE;
			break;
		default:
			break;
	}
	//!
	switch(m_uTimePro)
	{
		case 0:
			bTimeType=0;
			break;
		case 1:
			bTimeType=1;
			break;
		case 2:
			bTimeType=2;
			break;
		default:
			break;
	}
	//!
	/////////////////////////////////////////////////////////////////////////////////////////
	//分解策略////////////////////////////////////////////////////////////////////////////////
	/////////////////////////////////////////////////////////////////////////////////////////
	//!1、解读时间；2、需要搜索的表名
	SYSTEMTIME sys; 
	GetLocalTime(&sys);	
	vector<CString> vctList;
	int y=sys.wYear,m=sys.wMonth, d=sys.wDay;

	UINT uW=6;
	CString strTmp,strName;
	switch(bTimeType)
	{
		case 0://day
			for(int i=0;i <uTimeContinue+1;++i)
			{//!
				strTmp=_T("");
				strName=_T("");		
				do
				{//!		
					if(d>1)
					{//!
						d-=1;
					}
					else
					{//!
						if(m>1)
						{
							m-=1;
							d=daysOfMonth(y,m);
						}
						else
						{			
							y-=1;
							m=12;
							d=daysOfMonth(y,m);
						}
					}
					uW=CaculateWeekDay(y,m,d);
				}while(uW>5);

				strName=GetTimeString(y,m,d);
				strName=_T("A")+strName;
				vctList.push_back(strName);
			}
			break;
        case 1://week
			//！
			for(int i=0;i <uTimeContinue+1;++i)
			{//!
				strTmp=_T("");
				strName=_T("");		
				
				uW=CaculateWeekDay(y,m,d);
				if(uW>=5)
				{
					//!调整到本周五
					d-=uW-5;
				}
				else
				{
					//!调整到上周五
					d-=uW+2;
				}
				//！
				if(d<0)
				{//!
					if(m>1)
					{
						m-=1;
						d=daysOfMonth(y,m)+d;
					}
					else
					{			
						y-=1;
						m=12;
						d=daysOfMonth(y,m)+d;
					}
				}					
				//uW=CaculateWeekDay(y,m,d);
				strName=GetTimeString(y,m,d);
				strName=_T("A")+strName;
				vctList.push_back(strName);
				d-=7;
			}
			break;
		default:
			break;
	}	
	_ASSERTE(vctList.size());
	//！用最近一份数据建立比较样本
	double dFactor,dFactorTmp;
	CString strSql;
	//strTmp.Format(_T("%d"),uID);
	strSql=_T("select * from ");
	strSql+=vctList.at(0);
	//strSql+=_T(" where id = ")+strTmp;
	SQLiteDataReader Reader = sqlite.ExcuteQuery(strSql); 
	map<UINT,double> mapList;//id  factor
	UINT uID,uNumber;
	while(Reader.Read()) 
    {  
		uID=Reader.GetInt64Value(0);	
		dFactor=Reader.GetFloatValue(3);	
		mapList[uID]=dFactor;
    }  
    Reader.Close();
	//！用比较样本 去遍历符合条件的表，记录符合策略的ID；
	UINT uCount=0;
	vector<UINT> vctIDs;	
	for(map<UINT,double>::iterator p=mapList.begin();p!=mapList.end();++p)
	{//!
		uCount=0;
		uID=(*p).first;
		dFactor=(*p).second;
		for(vector<CString>::iterator pIt=vctList.begin(); pIt!=vctList.end();++pIt)
		{
			strTmp.Format(_T("%d"),uID);
			strSql=_T("select * from ");
			strSql+=(*pIt);
			strSql+=_T(" where id = ")+strTmp;
			//!
			SQLiteDataReader Reader = sqlite.ExcuteQuery(strSql);	
			while(Reader.Read()) 
			{  	
				dFactorTmp=Reader.GetFloatValue(3);
				if(bStateType)//!
				{
					if(dFactorTmp < dFactor)
						uCount++;					
				}
				else
				{
					if(dFactorTmp > dFactor)
						uCount++;			
				}
				dFactor=dFactorTmp;
			} 
			Reader.Close();
		}
		if(uCount==uTimeContinue)
			vctIDs.push_back(uID);
	}
	//!	
	//！用符合策略的ID，去获取最新数据，建立数据集；
	map<UINT,vector<CHKStockData>> mapDatas;
	for(vector<UINT>::iterator pIt=vctIDs.begin(); pIt!=vctIDs.end();++pIt)
	{
		uID=(*pIt);
		vector<CHKStockData> vctstockDatas;
		for(vector<CString>::iterator pIt=vctList.begin(); pIt!=vctList.end();++pIt)
		{
			strTmp.Format(_T("%d"),uID);
			strSql=_T("select * from ");
			strSql+=(*pIt);
			strSql+=_T(" where id = ")+strTmp;
			//!
			SQLiteDataReader Reader = sqlite.ExcuteQuery(strSql);	
			while(Reader.Read()) 
			{  		
				CHKStockData tmp;
				tmp.SetTime((*pIt));
				tmp.SetCode(Reader.GetInt64Value(0));
				tmp.SetName(Reader.GetStringValue(1));
				tmp.SetCount(Reader.GetInt64Value(2));
				tmp.SetFactor(Reader.GetFloatValue(3));		
				vctstockDatas.push_back(tmp);
			} 
			Reader.Close();
		}
		mapDatas[uID]=vctstockDatas;
	}  
	// 关闭数据库  
    sqlite.Close(); 
	//!获取行情数据
	
	//!输出符合策略的数据集；
	strName=GetTimeString(sys.wYear,sys.wMonth,sys.wDay);	
	switch(bTimeType)
	{
		case 0:
			strTmp.Format(_T("北上资金连续%d日"),uTimeContinue);
			break;
		case 1:
			strTmp.Format(_T("北上资金连续%d周"),uTimeContinue);
			break;
		case 2:
			strTmp.Format(_T("北上资金连续%d月"),uTimeContinue);
			break;
		default:
			break;
	}		
	if(bStateType)
	{//!
		strTmp+=_T("增仓");
	}
	else
	{
		strTmp+=_T("减仓");
	}
	strName=strTmp+_T("_")+strName;
	//!
	CString strFileName=strName;
	CFileDialog dlgsave(FALSE,NULL,strFileName,OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,_T("Excel Files(*.xlsx)|*.xlsx|Excel Files(*.xls)|*.xls|All Files (*.*)|*.*||"));  
	if(dlgsave.DoModal() == IDOK)  
	{  
		CXLControl* pCXLControl = new CXLControl; 

		pCXLControl->NewXL();//创建表格
		strFileName =  dlgsave.GetPathName(); 

		int nExcelCol = 0;
		int nExcelRow = 1;

		pCXLControl->SetXL(nExcelRow,0,lpcsCode);
		pCXLControl->SetXL(nExcelRow,1,lpcsName);
		pCXLControl->SetXL(nExcelRow,2,lpcsCount);
		pCXLControl->SetXL(nExcelRow,3,lpcsFactor);
		pCXLControl->SetXL(nExcelRow++,4,lpcsChange);
		for(map<UINT,vector<CHKStockData>>::iterator p=mapDatas.begin();p!=mapDatas.end();++p,++nExcelRow)  
		{ 	
			vector<CHKStockData> vctstockDatas=(*p).second;
			CHKStockData tmp;
			if(bStateType)
			{
				tmp=vctstockDatas.at(0);
			}
			else
			{
				tmp=vctstockDatas.at(vctstockDatas.size()-1);
			}		
			//名称
			strTmp.Format(_T("%d"),(long)tmp.GetCode());
			pCXLControl->SetXL(nExcelRow,0,strTmp);

			pCXLControl->SetXL(nExcelRow,1,tmp.GetName());

			strTmp.Format(_T("%d"),(long)tmp.GetCount());
			pCXLControl->SetXL(nExcelRow,2,strTmp);

			double dFactor=tmp.GetFactor()*100;
			strTmp.Format(_T("%.2f"),dFactor);
			strTmp+=_T("%");
			pCXLControl->SetXL(nExcelRow,3,strTmp);

			double dChange=(vctstockDatas.at(0).GetFactor()-vctstockDatas.at(vctstockDatas.size()-1).GetFactor());
			strTmp.Format(_T("%.2f"),dChange*100);
			strTmp+=_T("%");
			pCXLControl->SetXL(nExcelRow,4,strTmp);		
		}		
		pCXLControl->SetCellWidth(0,4,10);//设置列宽
		pCXLControl->SetHoriAlign(1,0,nExcelRow,4,1);//设置单元格水平对齐方式
		pCXLControl->SetFonts(2,0,nExcelRow,4,12,NULL,1,FALSE);//设置字体
		pCXLControl->SetFonts(0,0,0,4,16,NULL,1,FALSE);//设置字体
		pCXLControl->SetFonts(1,0,1,4,12,NULL,1,FALSE);//设置字体
		if(bStateType)
			pCXLControl->SetCellColor(1,0,1,4,3);//设置单位颜色为绿色 0无 1 黑 2白3红 4绿
		else
			pCXLControl->SetCellColor(1,0,1,4,4);//设置单位颜色为绿色 0无 1 黑 2白3红 4绿

		pCXLControl->SetFonts(1,0,1,4,10,NULL,1,TRUE);//设置表头文字加粗
		pCXLControl->SetBorders(1,0,nExcelRow-1,4,1);//设置显示网格线

		strTmp.Format(_T(":总计%d只个股"),mapDatas.size());
		strTmp=strName+strTmp;
		pCXLControl->SetXL(0,0,strTmp);
		pCXLControl->SetXL(++nExcelRow,0,_T("附注:"));
		pCXLControl->SetXL(++nExcelRow,0,_T("1、连续多日指连续多个交易日;"));
		pCXLControl->SetXL(++nExcelRow,0,_T("2、连续多周、月均指连续多个自然周、月;"));
		pCXLControl->SetXL(++nExcelRow,0,_T("3、统计结果仅供参考，不作为操作依据。"));

		pCXLControl->SetSheetName(strName);//设置sheet表名称
		pCXLControl->SaveAs(strFileName);//另存为

		pCXLControl->TerminateExcel();//终止excel
		_DELPTR(pCXLControl);
		AfxMessageBox(_T("数据导出完毕!"));
	} 
	//!
	//memset(sql,0,sizeof(sql));
	// _stprintf_s(sql,_T("%s"),  
 //       _T("CREATE TABLE [Tmp] (")  
 //       _T("[id] INTEGER NOT NULL PRIMARY KEY, ")  
 //       _T("[name] NVARCHAR(20), ")  
	//	_T("[count] INTEGER, ") 
	//	_T("[factor] REAL, ") 
 //       _T("[change] REAL); ") 
 //       );  
 //   if(!sqlite.ExcuteNonQuery(sql))
 //   {  
 //       printf("Create database table failed...\n");  
 //   }
	//else
	//{//!
 // 	  //// 当一次性插入多条记录时候，采用事务的方式，提高效率
	//	sqlite.BeginTransaction();  
	//	memset(sql,0,sizeof(sql));  
	//	_stprintf_s(sql,_T("insert into Tmp(id,name,count,factor,change) values(?,?,?,?,?)"));  
	//	//strSql=_T("insert into ")+strName+_T("(id,name,count,factor) values(?,?,?,?)");
	//	SQLiteCommand cmd(&sqlite,sql);  
	//	// 批量插入数据  
	//	for(map<UINT,vector<CHKStockData>>::iterator p=mapDatas.begin();p!=mapDatas.end();++p)  
	//	{ 	
	//		TCHAR strValue[16] = {0};  
	//		vector<CHKStockData> vctstockDatas=(*p).second;
	//		CHKStockData tmp;
	//		if(bStateType)
	//		{
	//			tmp=vctstockDatas.at(0);
	//		}
	//		else
	//		{
	//			tmp=vctstockDatas.at(vctstockDatas.size()-1);
	//		}
	//		_stprintf_s(strValue,_T("%d"),(*p).first);  
	//		// 绑定第一个参数（id字段值）  
	//		cmd.BindParam(1,(int)((*p).first) );// 
	//		// 绑定第二个参数（name字段值）  
	//		cmd.BindParam(2,(tmp.GetName())); 		
	//		// 绑定第三个参数（count字段值）  
	//		_stprintf_s(strValue,_T("%d"),(int)(tmp.GetCount())); 
	//		cmd.BindParam(3,(int)(tmp.GetCount()));//
	//		// 绑定第三个参数（factor字段值）
	//		_stprintf_s(strValue,_T("%.4f"),tmp.GetFactor()); 
	//		cmd.BindParam(4,(double)(tmp.GetFactor()));//  
	//		double dChange=(vctstockDatas.at(0).GetFactor()-vctstockDatas.at(vctstockDatas.size()-1).GetFactor());
	//		cmd.BindParam(5,dChange);//  
	//		if(!sqlite.ExcuteNonQuery(&cmd)) 
	//		{  
	//			_tprintf(_T("%s\n"),sqlite.GetLastErrorMsg());  
	//			break;  
	//		} 					   
	//	}  
	//	// 清空cmd  
	//	cmd.Clear();  
	//	// 提交事务  
	//	sqlite.CommitTransaction();  
	//}	
    // 关闭数据库  
   // sqlite.Close(); 
	//！
	return 0;
}

void CImportDataDlg::OnCbnSelchangeCmbState()
{
	// TODO: 在此添加控件通知处理程序代码
}

BOOL CImportDataDlg::Anasylis_factor()//! 
{//!
	//!
	CString strPath=GetDataPath();
    SQLite sqlite;  
    // 打开或创建数据库
    //******************************************************  
    if(!sqlite.Open(strPath))  
    {  
       // _tprintf(_T("%s\n"),sqlite.GetLastErrorMsg());  
		strPath+=_T("_文件打开错误!");
		AfxMessageBox(strPath);
        return 0;  
    } 
	//!获取列表
	TCHAR sql[512] = {0}; 
    memset(sql,0,sizeof(sql));  
	/////////////////////////////////////////////////////////////////////////////////////////
	//!策略//////////////////////////////////////////////////////////////////////////////////
	/////////////////////////////////////////////////////////////////////////////////////////
	//!
	UINT uTimeContinue=3;//!连续
	BOOL bStateType=FALSE;//!TRUE，增仓；FALSE，减仓
	UINT bTimeType=0;//!0、天；1、周；2、月

	Dlg2Data();	
	//!
	//！读取列表
	double dFactor,dFactorTmp;
	CString strSql,strNumber;
	strSql=_T("select * from A2HK");
	SQLiteDataReader Reader = sqlite.ExcuteQuery(strSql); 
	map<UINT,CString> mapList;//id  number
	UINT uID,uNumber;
	while(Reader.Read()) 
    {  
		uID=Reader.GetInt64Value(0);	
		strNumber=Reader.GetStringValue(2);	
		mapList[uID]=strNumber;
    }  
	//!策略 1\dFactor>=0.01; 首次出现指定比例个股，至今
	dFactor=0.01;
	double dStep=0.01;	
	//分解策略////////////////////////////////////////////////////////////////////////////////
	/////////////////////////////////////////////////////////////////////////////////////////
	//!1、解读时间；2、需要搜索的表名
	SYSTEMTIME sys; 
	GetLocalTime(&sys);	
	int yOrg=sys.wYear,mOrg=sys.wMonth, dOrg=sys.wDay;
	int y=2017,m=3, d=17,uW=0;
	vector<CString> vctList;
	CString strTmp=_T(""),strName=_T("");			
	do
	{//!		
		if(d<=daysOfMonth(y,m))
		{//!
			uW=CaculateWeekDay(y,m,d);
			if(uW<6)
			{
				strName=GetTimeString(y,m,d);
				strName=_T("A")+strName;
				vctList.push_back(strName);
			}				
		}
		else
		{//!				
			if(y<yOrg)
			{
				if(m<12)
				{
					m+=1;					
				}
				else
				{
					m=1;
					y+=1;
				}
				d=1;
			}
			else
			{			
				if(y>yOrg)
					break;

				if(m<mOrg)
				{
					m+=1;						
				}
				else
				{
					break;
				}
				d=1;
			}
			uW=CaculateWeekDay(y,m,d);
			if(uW<6)
			{
				strName=GetTimeString(y,m,d);
				strName=_T("A")+strName;
				vctList.push_back(strName);
			}
		}
		d++;
		if(y>= yOrg && m>=mOrg)
		{//！
			if(d>dOrg)
				break;
		}
	}while(true);

	//!循环搜索符合策略的个股;
	dFactor=0;
	map<UINT,map<UINT,CHKStockData>> mapFactor2ID2Data;	
	map<UINT,CString>::iterator p; //id  number
	for(int i=0;i<6;++i)
	{//!
		dFactor+=dStep;
		map<UINT,CHKStockData> mapID2Data;
		for(vector<CString>::iterator pIt=vctList.begin(); pIt!=vctList.end();++pIt)
		{			
			//strTmp.Format(_T("%d"),uID);
			strSql=_T("select * from ");
			strSql+=(*pIt);
			//strSql+=_T(" where id = ")+strTmp;
			//!
			SQLiteDataReader Reader = sqlite.ExcuteQuery(strSql);	
			while(Reader.Read()) 
			{  		
				uID=Reader.GetInt64Value(0);
				dFactorTmp=Reader.GetFloatValue(3);	
				if(DblGE(dFactorTmp,dFactor))//!首次
				{//！
					BOOL bExist=FALSE;
					if(mapID2Data.size()>0)
					{
						if(mapID2Data.find(uID)!=mapID2Data.end())
							bExist=TRUE;
					}	
					else
						bExist=FALSE;

					if(!bExist)
					{
						CHKStockData tmp;
						strTmp=(*pIt);
						strTmp.Replace(_T("A"),_T(""));
						tmp.SetTime(strTmp);
						tmp.SetCode(Reader.GetInt64Value(0));
						tmp.SetName(Reader.GetStringValue(1));
						tmp.SetCount(Reader.GetInt64Value(2));
						tmp.SetFactor(Reader.GetFloatValue(3));	
						p=mapList.find(Reader.GetInt64Value(0));
						if(p!=mapList.end())
						{
							tmp.SetNumber((*p).second);
							mapID2Data[uID]=tmp;
						}	
					}									
				}
			} 
			Reader.Close();
		}
		mapFactor2ID2Data[dFactor*100]=mapID2Data;
	}
	/////////////////////////////////////////////////////////////////////////////////////////	
	//// 关闭数据库  
    sqlite.Close(); 	
	//!打开新数据
	strPath=GetDataHistoryPath();//_T("D:\\AHistory2017.db");
	SQLite sqlite2; 
    // 打开或创建数据库
    //******************************************************  
    if(!sqlite2.Open(strPath))  
    {  
       // _tprintf(_T("%s\n"),sqlite.GetLastErrorMsg());  
		strPath+=_T("_文件打开错误!");
		AfxMessageBox(strPath);
        return 0;  
    }
	////!获取行情数据
	CString strLastTime=_T("20171124");
	//map<UINT,map<UINT,CHKStockData>> mapFactor2ID2Data;	
	for(map<UINT,map<UINT,CHKStockData>>::iterator pIt=mapFactor2ID2Data.begin(); pIt!=mapFactor2ID2Data.end();++pIt)
	{//!	
		//!
		map<UINT,CHKStockData>  piterator=(*pIt).second;
		for(map<UINT,CHKStockData>::iterator p=(piterator).begin(); p!=(piterator).end();++p)
		{		
			CHKStockData tmp=(*p).second;			
								
			double dPreValue=0,dPreFactor=0;
			BOOL bFind=FALSE;
			CString strTime=tmp.GetTime();			
			do{				
				strTmp=_T("A")+tmp.GetNumber();
				strSql=_T("select * from ")+strTmp;
				strSql+=_T(" where data = ")+strTime;
				SQLiteDataReader Reader = sqlite2.ExcuteQuery(strSql);
				while(Reader.Read()) 
				{  		
					bFind=TRUE;
					strTmp=Reader.GetStringValue(0);
					double dTmp=Reader.GetFloatValue(6);
					double dTmp2=Reader.GetFloatValue(5);
					dPreValue=dTmp/dTmp2;
					dPreFactor=Reader.GetFloatValue(7);			
				} 
				if(!bFind)//!如果未取到数据
				{//！向上回溯
					strTime=CalcTimeString(strTime);
				}
				Reader.Close();
			}while(!bFind);			
			
			bFind=FALSE;
			double dCurValue=0,dCurFactor=0;
			strTime=strLastTime;
			do{
				strTmp=_T("A")+tmp.GetNumber();
				strSql=_T("select * from ")+strTmp;
				strSql+=_T(" where data = ")+strTime;
		
				SQLiteDataReader Reader2 = sqlite2.ExcuteQuery(strSql);	
				
				while(Reader2.Read()) 
				{  	
					bFind=TRUE;
					strTmp=Reader2.GetStringValue(0);
					double dTmp=Reader2.GetFloatValue(6);
					double dTmp2=Reader2.GetFloatValue(5);
					dCurValue=dTmp/dTmp2;
					dCurFactor=Reader2.GetFloatValue(7);
				}
				if(!bFind)//!如果未取到数据
				{//！向上回溯
					strTime=CalcTimeString(strTime);
				}
				Reader2.Close();
			}while(!bFind);	
		
			dPreValue=dPreValue*dPreFactor/dCurFactor;
			dFactorTmp=(dCurValue-dPreValue)/dPreValue;		
			tmp.SetValue(dFactorTmp);
			piterator[(*p).first]=tmp;
		}
		mapFactor2ID2Data[(*pIt).first]=piterator;
	}	
	//!输出符合策略的数据集；	
	strName=GetTimeString(sys.wYear,sys.wMonth,sys.wDay);	
	strName=_T("北上资金持股数据统计_比例")+strLastTime;
	////!
	CString strFileName=strName;
	CFileDialog dlgsave(FALSE,NULL,strFileName,OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,_T("Excel Files(*.xlsx)|*.xlsx|Excel Files(*.xls)|*.xls|All Files (*.*)|*.*||"));  
	if(dlgsave.DoModal() == IDOK)  
	{  
		CXLControl* pCXLControl = new CXLControl; 

		pCXLControl->NewXL();//创建表格
		strFileName =  dlgsave.GetPathName(); 

		int nExcelCol = 0,nCol = 0;
		int nExcelRow = 1,nRow = 1;

		pCXLControl->SetXL(nExcelRow,0,_T("样本比例"));
		pCXLControl->SetXL(nExcelRow,1,_T("编号"));
		pCXLControl->SetXL(nExcelRow,2,_T("名称"));
		pCXLControl->SetXL(nExcelRow,3,_T("达到样本比例日期"));
		pCXLControl->SetXL(nExcelRow,4,_T("实际比例"));
		pCXLControl->SetXL(nExcelRow++,5,_T("涨幅"));

		for(map<UINT,map<UINT,CHKStockData>>::iterator p=mapFactor2ID2Data.begin(); p!=mapFactor2ID2Data.end();++p)
		{//!			
			map<UINT,CHKStockData> mapStockDatas=(*p).second;
			for(map<UINT,CHKStockData>::iterator pIt=mapStockDatas.begin();pIt!=mapStockDatas.end();++pIt)
			{
				nExcelCol=nCol;
				CHKStockData tmp=((*pIt).second);	
				//样本比例
				strTmp.Format(_T("%d"),(long)(*p).first);
				strTmp+=_T("%");
				pCXLControl->SetXL(nExcelRow,0,strTmp);
				//!编号
				pCXLControl->SetXL(nExcelRow,1,tmp.GetNumber());
				//!名称
				pCXLControl->SetXL(nExcelRow,2,tmp.GetName());
				//!达到样本比例日期
				pCXLControl->SetXL(nExcelRow,3,tmp.GetTime());
				//!实际比例
				strTmp.Format(_T("%.2f"),tmp.GetFactor()*100);
				strTmp+=_T("%");
				pCXLControl->SetXL(nExcelRow,4,strTmp);
				//!涨幅
				strTmp.Format(_T("%.2f"),tmp.GetValue()*100);
				pCXLControl->SetXL(nExcelRow++,5,strTmp);		
			}		
		}		
		pCXLControl->SetCellWidth(0,5,10);//设置列宽
		pCXLControl->SetHoriAlign(1,0,5,4,1);//设置单元格水平对齐方式
		pCXLControl->SetFonts(2,0,nExcelRow,5,12,NULL,1,FALSE);//设置字体
		pCXLControl->SetFonts(0,0,0,5,16,NULL,1,FALSE);//设置字体
		pCXLControl->SetFonts(1,0,1,5,12,NULL,1,FALSE);//设置字体
		//if(bStateType)
		//	pCXLControl->SetCellColor(1,0,1,4,3);//设置单位颜色为绿色 0无 1 黑 2白3红 4绿
		//else
		//	pCXLControl->SetCellColor(1,0,1,4,4);//设置单位颜色为绿色 0无 1 黑 2白3红 4绿

		pCXLControl->SetFonts(1,0,1,nExcelCol,10,NULL,1,TRUE);//设置表头文字加粗
		pCXLControl->SetBorders(1,0,nExcelRow-1,nExcelCol,1);//设置显示网格线

		//strTmp.Format(_T(":总计%d只个股"),mapDatas.size());
		//strTmp=strName+strTmp;
		//pCXLControl->SetXL(0,0,strTmp);
		//pCXLControl->SetXL(++nExcelRow,0,_T("附注:"));
		//pCXLControl->SetXL(++nExcelRow,0,_T("1、连续多日指连续多个交易日;"));
		//pCXLControl->SetXL(++nExcelRow,0,_T("2、连续多周、月均指连续多个自然周、月;"));
		//pCXLControl->SetXL(++nExcelRow,0,_T("3、统计结果仅供参考，不作为操作依据。"));

		pCXLControl->SetSheetName(strName);//设置sheet表名称
		pCXLControl->SaveAs(strFileName);//另存为

		pCXLControl->TerminateExcel();//终止excel
		_DELPTR(pCXLControl);
	
		map<UINT,FactorData>  mapFactorData;
		for(map<UINT,map<UINT,CHKStockData>>::iterator p=mapFactor2ID2Data.begin(); p!=mapFactor2ID2Data.end();++p)
		{//!			
			FactorData tmp;
			tmp.m_dFactorSample=(*p).first;			
			map<UINT,CHKStockData> mapStockDatas=(*p).second;
			for(map<UINT,CHKStockData>::iterator pIt=mapStockDatas.begin();pIt!=mapStockDatas.end();++pIt)
			{	
				if(tmp.m_dMaxUp<(*pIt).second.GetValue())
					tmp.m_dMaxUp=(*pIt).second.GetValue();

				if(tmp.m_dMaxLow>(*pIt).second.GetValue())
					tmp.m_dMaxLow=(*pIt).second.GetValue();	

				if((*pIt).second.GetValue()>0)
					tmp.m_dFactorWF++;

				tmp.m_dAvg+=(*pIt).second.GetValue();
				tmp.m_strFirstTime=(*pIt).second.GetTime();
			}
			tmp.m_dAvg=tmp.m_dAvg/mapStockDatas.size();
			tmp.m_dFactorWF=tmp.m_dFactorWF/(mapStockDatas.size());
			mapFactorData[(*p).first]=tmp;
		}//！	
		
		//！result
		pCXLControl = new CXLControl; 
		pCXLControl->NewXL();//创建表格

		nExcelCol = 0;
		nCol = 0;
		nExcelRow = 1;
		nRow = 1;

		strTmp=_T("自20170317以来到")+strLastTime+_T("收盘的统计");
		pCXLControl->SetXL(0,0,strTmp);

		pCXLControl->SetXL(nExcelRow,0,_T("持股比例"));
		pCXLControl->SetXL(nExcelRow,1,_T("首次达到比例日期"));
		pCXLControl->SetXL(nExcelRow,2,_T("最大涨幅"));
		pCXLControl->SetXL(nExcelRow,3,_T("最大跌幅"));	
		pCXLControl->SetXL(nExcelRow,4,_T("胜率"));
		pCXLControl->SetXL(nExcelRow++,5,_T("均涨幅"));
		CString strTmp;
		for(map<UINT,FactorData>::iterator p=mapFactorData.begin(); p!=mapFactorData.end();++p)
		{//!			
			FactorData tmp=((*p).second);	
			strTmp.Format(_T("%d"),(int)(tmp.m_dFactorSample));//
			strTmp+=_T("%");
			pCXLControl->SetXL(nExcelRow,0,strTmp);//_T("样本比例"));

			pCXLControl->SetXL(nExcelRow,1,tmp.m_strFirstTime);//_T("样本比例"));

			strTmp.Format(_T("%.2f"),tmp.m_dMaxUp*100);//
			strTmp+=_T("%");
			pCXLControl->SetXL(nExcelRow,2,strTmp);//_T("最大涨幅")
	
			strTmp.Format(_T("%.2f"),tmp.m_dMaxLow*100);//
			strTmp+=_T("%");
			pCXLControl->SetXL(nExcelRow,3,strTmp);//_T("最大跌幅")
			//!实际比例
			strTmp.Format(_T("%.2f"),tmp.m_dFactorWF*100);//_T("胜率")
			strTmp+=_T("%");
			pCXLControl->SetXL(nExcelRow,4,strTmp);
			//!涨幅
			strTmp.Format(_T("%.2f"),tmp.m_dAvg*100);
			strTmp+=_T("%");
			pCXLControl->SetXL(nExcelRow++,5,strTmp);		
		}
		pCXLControl->SetCellWidth(0,5,10);//设置列宽ui 
		pCXLControl->SetHoriAlign(1,0,nExcelRow,5,1);//设置单元格水平对齐方式h
		pCXLControl->SetFonts(2,0,nExcelRow,5,12,NULL,1,FALSE);//设置字体
		pCXLControl->SetFonts(0,0,0,5,16,NULL,1,FALSE);//设置字体
		pCXLControl->SetFonts(1,0,1,5,12,NULL,1,FALSE);//设置字体

		pCXLControl->SetFonts(1,0,1,nExcelCol,10,NULL,1,TRUE);//设置表头文字加粗
		pCXLControl->SetBorders(1,0,nExcelRow-1,5,1);//设置显示网格线

		strName=_T("统计数据");
		strFileName+=strName;
		pCXLControl->SetSheetName(strName);//设置sheet表名称
		pCXLControl->SaveAs(strFileName);//另存为

		pCXLControl->TerminateExcel();//终止excel
		_DELPTR(pCXLControl);

		AfxMessageBox(_T("数据导出完毕!"));
	} 
	//!
	//memset(sql,0,sizeof(sql));
	// _stprintf_s(sql,_T("%s"),  
 //       _T("CREATE TABLE [Tmp] (")  
 //       _T("[id] INTEGER NOT NULL PRIMARY KEY, ")  
 //       _T("[name] NVARCHAR(20), ")  
	//	_T("[count] INTEGER, ") 
	//	_T("[factor] REAL, ") 
 //       _T("[change] REAL); ") 
 //       );  
 //   if(!sqlite.ExcuteNonQuery(sql))
 //   {  
 //       printf("Create database table failed...\n");  
 //   }
	//else
	//{//!
 // 	  //// 当一次性插入多条记录时候，采用事务的方式，提高效率
	//	sqlite.BeginTransaction();  
	//	memset(sql,0,sizeof(sql));  
	//	_stprintf_s(sql,_T("insert into Tmp(id,name,count,factor,change) values(?,?,?,?,?)"));  
	//	//strSql=_T("insert into ")+strName+_T("(id,name,count,factor) values(?,?,?,?)");
	//	SQLiteCommand cmd(&sqlite,sql);  
	//	// 批量插入数据  
	//	for(map<UINT,vector<CHKStockData>>::iterator p=mapDatas.begin();p!=mapDatas.end();++p)  
	//	{ 	
	//		TCHAR strValue[16] = {0};  
	//		vector<CHKStockData> vctstockDatas=(*p).second;
	//		CHKStockData tmp;
	//		if(bStateType)
	//		{
	//			tmp=vctstockDatas.at(0);
	//		}
	//		else
	//		{
	//			tmp=vctstockDatas.at(vctstockDatas.size()-1);
	//		}
	//		_stprintf_s(strValue,_T("%d"),(*p).first);  
	//		// 绑定第一个参数（id字段值）  
	//		cmd.BindParam(1,(int)((*p).first) );// 
	//		// 绑定第二个参数（name字段值）  
	//		cmd.BindParam(2,(tmp.GetName())); 		
	//		// 绑定第三个参数（count字段值）  
	//		_stprintf_s(strValue,_T("%d"),(int)(tmp.GetCount())); 
	//		cmd.BindParam(3,(int)(tmp.GetCount()));//
	//		// 绑定第三个参数（factor字段值）
	//		_stprintf_s(strValue,_T("%.4f"),tmp.GetFactor()); 
	//		cmd.BindParam(4,(double)(tmp.GetFactor()));//  
	//		double dChange=(vctstockDatas.at(0).GetFactor()-vctstockDatas.at(vctstockDatas.size()-1).GetFactor());
	//		cmd.BindParam(5,dChange);//  
	//		if(!sqlite.ExcuteNonQuery(&cmd)) 
	//		{  
	//			_tprintf(_T("%s\n"),sqlite.GetLastErrorMsg());  
	//			break;  
	//		} 					   
	//	}  
	//	// 清空cmd  
	//	cmd.Clear();  
	//	// 提交事务  
	//	sqlite.CommitTransaction();  
	//}	
    // 关闭数据库  
   // sqlite.Close(); 
	//！
	return 0;
}
void CImportDataDlg::OnBnClickedBtnAnasy()
{
	// TODO: 在此添加控件通知处理程序代码
	Anasylis_factor();
}
