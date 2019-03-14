
// FullAppDlg.cpp: 实现文件
//

#include "stdafx.h"
#include "FullApp.h"
#include "FullAppDlg.h"
#include "afxdialogex.h"
#include <stdio.h>

#include <vector>
#ifdef _DEBUG
#define new DEBUG_NEW
#endif
using namespace std;

// CFullAppDlg 对话框



CFullAppDlg::CFullAppDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_FULLAPP_DIALOG, pParent)
	, countnum(_T(""))
	, name(_T(""))
	, price(_T(""))
	, unit(_T(""))
	, unitprice(_T(""))
	, siglenum(_T(""))
	, note(_T(""))
	, path(_T(""))
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CFullAppDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_CbNote, m_note);
	DDX_Control(pDX, IDC_EdCountNum, m_countNum);
	DDX_Control(pDX, IDC_EdName, m_name);
	DDX_Control(pDX, IDC_EdPath, m_path);
	DDX_Control(pDX, IDC_EdPrice, m_price);
	DDX_Control(pDX, IDC_EdSigleNum, m_siglenume);
	DDX_Control(pDX, IDC_EdTatol, m_tatol);
	DDX_Control(pDX, IDC_EdUnit, m_unit);
	DDX_Control(pDX, IDC_EdUnitPrice, m_unitprice);
	DDX_Text(pDX, IDC_EdCountNum, countnum);
	DDX_Text(pDX, IDC_EdName, name);
	DDX_Text(pDX, IDC_EdPrice, price);
	DDX_Text(pDX, IDC_EdUnit, unit);
	DDX_Text(pDX, IDC_EdUnitPrice, unitprice);
	DDX_Text(pDX, IDC_EdSigleNum, siglenum);
	DDX_CBString(pDX, IDC_CbNote, note);
	DDX_Control(pDX, IDC_rdcheck, m_check);
	DDX_Text(pDX, IDC_EdPath, path);
}

BEGIN_MESSAGE_MAP(CFullAppDlg, CDialogEx)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BtSure, &CFullAppDlg::OnBnClickedBtsure)
	ON_BN_CLICKED(IDC_BtInsert, &CFullAppDlg::OnBnClickedBtinsert)
END_MESSAGE_MAP()


// CFullAppDlg 消息处理程序
BOOL CFullAppDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	ShowWindow(SW_MINIMIZE);
	// TODO: 在此添加额外的初始化代码
	CHAR m_ip[MAX_PATH]; 
	CHAR m_port[MAX_PATH];
	CHAR m_databasename[MAX_PATH];
	CHAR m_driver[MAX_PATH];
	CHAR m_uid[MAX_PATH];
	CHAR m_password[MAX_PATH];
	CHAR m_filename[MAX_PATH];
	CHAR szModuleName[MAX_PATH];
	CHAR szConfigFileName[MAX_PATH];
	GetModuleFileNameA(NULL, szModuleName, MAX_PATH);
	_snprintf_s(szConfigFileName, MAX_PATH, "%s\\..\\%s", szModuleName, DT_PATH);

	CHAR szBuffer[1024];
	GetPrivateProfileStringA(DT_NAME, "IP", NULL, szBuffer, 1024, szConfigFileName);
	strcpy_s(m_ip, szBuffer);

	GetPrivateProfileStringA(DT_NAME, "Port", NULL, szBuffer, 1024, szConfigFileName);
	strcpy_s(m_port, szBuffer);

	GetPrivateProfileStringA(DT_NAME, "DtName", NULL, szBuffer, 1024, szConfigFileName);
	strcpy_s(m_databasename, szBuffer);

	GetPrivateProfileStringA(DT_NAME, "Driver", NULL, szBuffer, 1024, szConfigFileName);
	strcpy_s(m_driver, szBuffer);

	GetPrivateProfileStringA(DT_NAME, "Uid", NULL, szBuffer, 1024, szConfigFileName);
	strcpy_s(m_uid, szBuffer);

	GetPrivateProfileStringA(DT_NAME, "Password", NULL, szBuffer, 1024, szConfigFileName);
	strcpy_s(m_password, szBuffer);

	GetPrivateProfileStringA(DT_NAME, "FileName", NULL, szBuffer, 1024, szConfigFileName);
	strcpy_s(m_filename, szBuffer);

	string timeName;
	timeName += m_filename;
	timeName += "\\";
	timeName += GetCurrentTime();
	//连接到MS SQL Server
	//初始化指针
	::CoInitialize(NULL);
	HRESULT hr = pMyConnect.CreateInstance(__uuidof(Connection));
	if (FAILED(hr))
		return FALSE;
	//初始化链接参数
	char Str[MAX_PATH];
	std::string strbuf = "Provider=%s;Server=%s,%s;Database=%s;uid=%s;pwd=%s;";
	sprintf_s(Str, strbuf.c_str(), m_driver, m_ip, m_port, m_databasename, m_uid, m_password);
	//执行连接
	try
	{
		pMyConnect->Open(Str, "", "", adModeUnknown);
		m_countNum.SetWindowTextW(_T("1"));
		m_unit.SetWindowTextW(_T("㎡"));
		m_unitprice.SetWindowTextW(_T("0"));
		m_price.SetWindowTextW(_T("0"));
		m_siglenume.SetWindowTextW(_T("1"));
		CString strTimeName(timeName.c_str());
		//strTimeName.Format(_T("%s"), timeName);
		m_path.SetWindowTextW(strTimeName);
		excelOperation = new CExcelOperation();
	}
	catch (_com_error &e)
	{
		MessageBox(e.Description(), _T("警告"), MB_OK | MB_ICONINFORMATION);
	}//发生链接错误
	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CFullAppDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CFullAppDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CFullAppDlg::sqloperate()
{
	_RecordsetPtr pRst(__uuidof(Recordset));//定义记录集对象并实例化对象  
	CString sqlstr = _T("select Note from ExcelModle where Name like"), textstr;
	m_name.GetWindowTextW(textstr);
	sqlstr += "'%";
	sqlstr += textstr;
	sqlstr += "%'";
	try 
	{
		pRst = pMyConnect->Execute((_bstr_t)sqlstr, NULL, adCmdText);//执行SQL语句
		if (!pRst->BOF)
		{
			pRst->MoveFirst();
		}
		else
		{
			return;
		}
		vector<_bstr_t> column_name;

		/*存储表的所有列名，显示表的列名*/
		for (int i = 0; i < pRst->Fields->GetCount(); i++)
		{
			column_name.push_back(pRst->Fields->GetItem(_variant_t((long)i))->Name);
		}
		int conutnum = 0;
		while (!pRst->adoEOF)
		{
			if (pRst->GetCollect(2).vt != VT_NULL)
			{
				CString temp = (TCHAR *)(_bstr_t)pRst->GetFields()->GetItem
				("name")->Value;
				m_note.InsertString(conutnum, temp);
				//m_note.SetItemData(conutnum, conutnum);
			}
			else
			{
				return;
			}
			conutnum++;
			pRst->MoveNext();
		}
		m_note.SetCurSel(0);
	}
	catch (_com_error &e)
	{
		return;
	}
}

void CFullAppDlg::exlceloperate()
{
	if (m_check.GetCheck())
	{
		initString(indexcount);
		initTitle();
		indexcount++;
	}
	if (m_name.GetWindowTextLengthW() != 0)
	{
		initString(indexcount);
		insertValue();
		indexcount++;
	}
	excelOperation->SaveWorkbookAs(path);
}



void CFullAppDlg::OnBnClickedBtsure()
{
	// TODO: 在此添加控件通知处理程序代码
	//执行python接口 创建excel表 
	if (m_tatol.GetWindowTextLengthW() != 0) 
	{
		excelOperation->Create();
		CString title,maxtitle;
		m_tatol.GetWindowTextW(title);
		excelOperation->SetCellMerge(_T("A1"), _T("G1"), title);
		indexcount = 2;
		initString(indexcount);
		initTitle();
		indexcount++;

	}
	else 
	{
		return;
	}
}



void CFullAppDlg::OnBnClickedBtinsert()
{
	// TODO: 在此添加控件通知处理程序代码
	exlceloperate();
}


BOOL CFullAppDlg::PreTranslateMessage(MSG* pMsg)
{
	if (pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_RETURN)
	{
			if (GetDlgItem(IDC_EdName) == GetFocus()) {
				// 你想做的事，如果按下回车时焦点在你想要的控件上
				sqloperate();
			}
		return TRUE;
	}
	if (pMsg->message == 0x4d)
	{//屏蔽F1帮助文档功能  并执行自己的代码
		exlceloperate();
		return TRUE;
	}
	return CDialogEx::PreTranslateMessage(pMsg);
}

std::string CFullAppDlg::GetCurrentTime()
{
	struct tm t;   //tm结构指针
	time_t now;  //声明time_t类型变量
	time(&now);      //获取系统日期和时间
	localtime_s(&t, &now);   //获取当地日期和时间
	char temp[50] = { 0 };
	sprintf_s(temp, "%d-%d.xls",t.tm_mon + 1, t.tm_mday);
	std::string  pTemp = temp;
	return pTemp;
}

void CFullAppDlg::insertValue()
{
	UpdateData(TRUE);
	price.Format(_T("%.2f"), _ttof(countnum)*_ttof(unitprice));
	excelOperation->setCellValue(strTitle[0], siglenum);
	excelOperation->setCellValue(strTitle[1], name);
	excelOperation->setCellValue(strTitle[2], unit);
	excelOperation->setCellValue(strTitle[3], countnum);
	excelOperation->setCellValue(strTitle[4], unitprice);
	excelOperation->setCellValue(strTitle[5], price);
	excelOperation->setCellValue(strTitle[6], note);
}

void CFullAppDlg::initTitle()
{
	UpdateData(TRUE);
	excelOperation->setCellValue(strTitle[0], _T("序号"));
	excelOperation->setCellValue(strTitle[1], _T("项目名称"));
	excelOperation->setCellValue(strTitle[2], _T("单位"));
	excelOperation->setCellValue(strTitle[3], _T("数量"));
	excelOperation->setCellValue(strTitle[4], _T("单价"));
	excelOperation->setCellValue(strTitle[5], _T("小计"));
	excelOperation->setCellValue(strTitle[6], _T("备注"));
}

void CFullAppDlg::initString(int indexxx)
{
	for (int i = 0; i < 7; i++) {
		strTitle[i] += to_string(indexxx);
	}
}