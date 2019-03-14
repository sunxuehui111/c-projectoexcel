
// FullAppDlg.h: 头文件
//

#pragma once
#include "CExcelOperation.h"
#include <string>
#include <vector>
#import "c:\Program Files\Common Files\System\ado\msado15.dll" no_namespace rename("EOF", "adoEOF") 
#define DT_PATH "ServerInit.ini"
#define DT_NAME "MSSQL"
// CFullAppDlg 对话框
class CFullAppDlg : public CDialogEx
{
// 构造
public:
	CFullAppDlg(CWnd* pParent = nullptr);	// 标准构造函数

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_FULLAPP_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;
	CExcelOperation* excelOperation;
	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
	void sqloperate();
	void exlceloperate();
	int indexcount = 1;
	std::vector<std::string> strTitle = { "A","B","C","D","E","F","G" };
	void initString(int indexxx);
	void initTitle();
	void insertValue();
	static std::string GetCurrentTime();
public:
	afx_msg void OnBnClickedBtsure();
	CComboBox m_note;
	CEdit m_countNum;
	CEdit m_maxTatel;
	CEdit m_name;
	CEdit m_path;
	CEdit m_price;
	CEdit m_siglenume;
	CEdit m_tatol;
	CEdit m_unit;
	CEdit m_unitprice;
	_ConnectionPtr pMyConnect;//定义连接对象并实例化对象 
	afx_msg void OnBnClickedBtinsert();
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	CString countnum;
	CString name;
	CString price;
	CString unit;
	CString unitprice;
	CString siglenum;
	CString note;
	CButton m_check;
	CString path;
	afx_msg void OnBnClickedBtexit();
};
