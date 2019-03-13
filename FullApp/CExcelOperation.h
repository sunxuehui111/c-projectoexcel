#pragma once
#include "CRange.h"
#include "CApplication.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CWorksheet.h"
#include "CWorksheets.h"
#include <string>
class CExcelOperation
{
private:
	CApplication m_ecApp;
	CWorkbooks m_ecBooks;
	CWorkbook m_ecBook;
	CWorksheets m_ecSheets;
	CWorksheet m_ecSheet;
	CRange m_ecRange;
	VARIANT ret;//���浥Ԫ���ֵ  
public:
	CExcelOperation();
	virtual ~CExcelOperation();
public:
	//����  
	//**********************������EXCEL*******************************************  
	BOOL CreateApp();
	BOOL CreateWorkbooks();                //����һ���µ�EXCEL����������  
	BOOL CreateWorkbook();                //����һ���µ�EXCEL������  
	BOOL CreateWorksheets();                //����һ���µ�EXCEL��������  
	BOOL CreateWorksheet(short index);                //����һ���µ�EXCEL������  
	BOOL CreateSheet(short index);
	BOOL Create(short index = 1);                         //�����µ�EXCELӦ�ó��򲢴���һ���¹������͹�����  
	void ShowApp();                        //��ʾEXCEL�ĵ�  
	void HideApp();                        //����EXCEL�ĵ�  
//**********************���ĵ�*********************************************  
	BOOL OpenWorkbook(CString fileName, short index = 1);
	BOOL Open(CString fileName);        //�����µ�EXCELӦ�ó��򲢴�һ���Ѿ����ڵ��ĵ���  
	BOOL SetActiveWorkbook(short i);    //���õ�ǰ������ĵ���  

	//**********************�����ĵ�*********************************************  
	BOOL SaveWorkbook();                //Excel���Դ���ʽ�����档  
	BOOL SaveWorkbookAs(CString fileName);//Excel�Դ�����ʽ�����档  
	BOOL CloseWorkbook();
	void CloseApp();
	//**********************����Ϣ********************************  
	BOOL GetRangeAndValue(CString begin, CString end);//�õ�begin��end��Range����֮���ֵ���õ�ret��  
	void GetRowsAndCols(long &rows, long &cols);//�õ�ret���У�����  
	BOOL GetTheValue(int rows, int cols, CString &dest);//���ص�rows��cols�е�ֵ��ע��ֻ�����ı����͵ģ���dest��  
	BOOL SetTextFormat(CString &beginS, CString &endS);//��beginS��endS����Ϊ�ı���ʽ(���ֵĻ�Ҫ������ķ�����תһ��)  
	BOOL SetRowToTextFormat(CString &beginS, CString &endS);
	void setCellValue(std::string ccellIndexChar, CString valueChar);
	void SetCellMerge(CString index1, CString index2, CString value);
	//��beginS��endS(������������)����Ϊ�ı���ʽ  

};

