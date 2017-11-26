
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
			 _T("[avgprice] REAL);");
		if(!sqlite.ExcuteNonQuery(sql))  
		{  
			printf("Create database table failed...\n");  
		}
		else
		{//!
  			//// 当一次性插入多条记录时候，采用事务的方式，提高效率  
			sqlite.BeginTransaction();  
			//memset(sqll,0,sizeof(sqll));  
			sql=_T("insert into ")+strName+_T("(data,open,close,high,low,avgprice) values(?,?,?,?,?,?)");
			//_stprintf_s(sqll,_T("insert into SHSE.600000(data,open,close,high,low,avgprice) values(?,?,?,?,?,?)"));  
			SQLiteCommand cmd(&sqlite,sql);	
			// 批量插入数据  
			for(vector<DailyBar>::iterator p=vctDailyBars.begin();p!=vctDailyBars.end();++p)
			{//!			
				DailyBar ItemTest=(*p);
				//TCHAR strValue[16] = {0};  
				//_stprintf_s(strValue,_T("%d"),(*p).second.second);  
				// 绑定第一个参数（id字段值）
				CString strTime;
				strTime.Format(_T("%s"),CA2T(ItemTest.strtime));
				int index=strTime.Find(_T("T"));
				if(index!=-1)
				{//!
					strTime=strTime.Left(index);			
				}
				strTime.Replace(_T("-"),_T(""));
				LPCTSTR str = (LPCTSTR)(strTime);//ItemTest.strtime;
				cmd.BindParam(1,str);  
				// 绑定第二个参数（name字段值）  
				double dTmp=ItemTest.open/**ItemTest.adj_factor*/;
				cmd.BindParam(2,dTmp);  		
				// 绑定第三个参数（number字段值）  
				dTmp=ItemTest.close/**ItemTest.adj_factor*/;
				cmd.BindParam(3,dTmp);  
				dTmp=ItemTest.high/**ItemTest.adj_factor*/;
				cmd.BindParam(4,dTmp); 
				dTmp=ItemTest.low/**ItemTest.adj_factor*/;
				cmd.BindParam(5,dTmp); 
				dTmp=ItemTest.amount/ItemTest.volume/**ItemTest.adj_factor*/;
				cmd.BindParam(6,dTmp); 
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
	map<UINT,double> mapList;
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
	map<UINT,vector<CStockData>> mapDatas;
	for(vector<UINT>::iterator pIt=vctIDs.begin(); pIt!=vctIDs.end();++pIt)
	{
		uID=(*pIt);
		vector<CStockData> vctstockDatas;
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
				CStockData tmp;
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
		for(map<UINT,vector<CStockData>>::iterator p=mapDatas.begin();p!=mapDatas.end();++p,++nExcelRow)  
		{ 	
			vector<CStockData> vctstockDatas=(*p).second;
			CStockData tmp;
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
	//	for(map<UINT,vector<CStockData>>::iterator p=mapDatas.begin();p!=mapDatas.end();++p)  
	//	{ 	
	//		TCHAR strValue[16] = {0};  
	//		vector<CStockData> vctstockDatas=(*p).second;
	//		CStockData tmp;
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


void CImportDataDlg::OnBnClickedBtnAnasy()
{
	// TODO: 在此添加控件通知处理程序代码
}
