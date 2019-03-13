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
	VARIANT ret;//保存单元格的值  
public:
	CExcelOperation();
	virtual ~CExcelOperation();
public:
	//操作  
	//**********************创建新EXCEL*******************************************  
	BOOL CreateApp();
	BOOL CreateWorkbooks();                //创建一个新的EXCEL工作簿集合  
	BOOL CreateWorkbook();                //创建一个新的EXCEL工作簿  
	BOOL CreateWorksheets();                //创建一个新的EXCEL工作表集合  
	BOOL CreateWorksheet(short index);                //创建一个新的EXCEL工作表  
	BOOL CreateSheet(short index);
	BOOL Create(short index = 1);                         //创建新的EXCEL应用程序并创建一个新工作簿和工作表  
	void ShowApp();                        //显示EXCEL文档  
	void HideApp();                        //隐藏EXCEL文档  
//**********************打开文档*********************************************  
	BOOL OpenWorkbook(CString fileName, short index = 1);
	BOOL Open(CString fileName);        //创建新的EXCEL应用程序并打开一个已经存在的文档。  
	BOOL SetActiveWorkbook(short i);    //设置当前激活的文档。  

	//**********************保存文档*********************************************  
	BOOL SaveWorkbook();                //Excel是以打开形式，保存。  
	BOOL SaveWorkbookAs(CString fileName);//Excel以创建形式，保存。  
	BOOL CloseWorkbook();
	void CloseApp();
	//**********************读信息********************************  
	BOOL GetRangeAndValue(CString begin, CString end);//得到begin到end的Range并将之间的值设置到ret中  
	void GetRowsAndCols(long &rows, long &cols);//得到ret的行，列数  
	BOOL GetTheValue(int rows, int cols, CString &dest);//返回第rows，cols列的值，注意只返回文本类型的，到dest中  
	BOOL SetTextFormat(CString &beginS, CString &endS);//将beginS到endS设置为文本格式(数字的还要用下面的方法再转一次)  
	BOOL SetRowToTextFormat(CString &beginS, CString &endS);
	void setCellValue(std::string ccellIndexChar, CString valueChar);
	void SetCellMerge(CString index1, CString index2, CString value);
	//将beginS到endS(包括数字类型)设置为文本格式  

};

