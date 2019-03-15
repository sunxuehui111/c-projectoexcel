#include "stdafx.h"
#include "CExcelOperation.h"

CExcelOperation::CExcelOperation()
{

}

CExcelOperation::~CExcelOperation()
{

}

BOOL CExcelOperation::CreateApp()
{
	//if (FALSE == m_wdApp.CreateDispatch("Word.Application"))  
	COleException pe;
	if (!m_ecApp.CreateDispatch(_T("Excel.Application"), &pe))
	{
		AfxMessageBox(_T("Application创建失败，请确保安装了wps或以上版本!"), MB_OK | MB_ICONWARNING);
		pe.ReportError();
		throw &pe;
		return FALSE;
	}
	return TRUE;
}

BOOL CExcelOperation::CreateWorkbooks()               //创建一个新的EXCEL工作簿集合  
{
	if (FALSE == CreateApp())
	{
		return FALSE;
	}
	m_ecBooks = m_ecApp.get_Workbooks();
	if (!m_ecBooks.m_lpDispatch)
	{
		AfxMessageBox(_T("WorkBooks创建失败!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	return TRUE;
}

BOOL CExcelOperation::CreateWorkbook()               //创建一个新的EXCEL工作簿  
{
	if (!m_ecBooks.m_lpDispatch)
	{
		AfxMessageBox(_T("WorkBooks为空!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}

	COleVariant vOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	m_ecBook = m_ecBooks.Add(vOptional);
	if (!m_ecBook.m_lpDispatch)
	{
		AfxMessageBox(_T("WorkBook为空!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	/*
		//得到document变量
		m_wdDoc = m_wdApp.GetActiveDocument();
		if (!m_wdDoc.m_lpDispatch)
		{
			AfxMessageBox("Document获取失败!", MB_OK|MB_ICONWARNING);
			return FALSE;
		}
		//得到selection变量
		m_wdSel = m_wdApp.GetSelection();
		if (!m_wdSel.m_lpDispatch)
		{
			AfxMessageBox("Select获取失败!", MB_OK|MB_ICONWARNING);
			return FALSE;
		}
		//得到Range变量
		m_wdRange = m_wdDoc.Range(vOptional,vOptional);
		if(!m_wdRange.m_lpDispatch)
		{
			AfxMessageBox("Range获取失败!", MB_OK|MB_ICONWARNING);
			return FALSE;
		}
	*/
	return TRUE;
}

BOOL CExcelOperation::CreateWorksheets()                //创建一个新的EXCEL工作表集合  
{
	if (!m_ecBook.m_lpDispatch)
	{
		AfxMessageBox(_T("WorkBook为空!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	m_ecSheets = m_ecBook.get_Sheets();
	if (!m_ecSheets.m_lpDispatch)
	{
		AfxMessageBox(_T("WorkSheets为空!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	return TRUE;
}

BOOL CExcelOperation::CreateWorksheet(short index)                //创建一个新的EXCEL工作表  
{
	if (!m_ecSheets.m_lpDispatch)
	{
		AfxMessageBox(_T("WorkSheets为空!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	m_ecSheet = m_ecSheets.get_Item(COleVariant(index));
	if (!m_ecSheet.m_lpDispatch)
	{
		AfxMessageBox(_T("WorkSheet为空!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	return TRUE;
}

BOOL CExcelOperation::CreateSheet(short index)
{
	if (CreateWorksheets() == FALSE)
	{
		return FALSE;
	}
	if (CreateWorksheet(index) == FALSE)
	{
		return FALSE;
	}
	return TRUE;
}

BOOL CExcelOperation::Create(short index)                        //创建新的EXCEL应用程序并创建一个新工作簿和工作表  
{
	if (CreateWorkbooks() == FALSE)
	{
		return FALSE;
	}
	if (CreateWorkbook() == FALSE)
	{
		return FALSE;
	}
	if (CreateSheet(index) == FALSE)
	{
		return FALSE;
	}
	return TRUE;
}

void CExcelOperation::ShowApp()                        //显示WORD文档  
{
	m_ecApp.put_Visible(TRUE);
}

void CExcelOperation::HideApp()                       //隐藏word文档  
{
	m_ecApp.put_Visible(FALSE);
}

//**********************打开文档*********************************************  
BOOL CExcelOperation::OpenWorkbook(CString fileName, short index)
{
	if (!m_ecBooks.m_lpDispatch)
	{
		AfxMessageBox(_T("WorkSheets为空!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	//COleVariant vFileName(_T(fileName));  
	COleVariant VOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	m_ecBook = m_ecBooks.Open(fileName, VOptional, VOptional, VOptional, VOptional, VOptional, VOptional, VOptional, VOptional, VOptional, VOptional, VOptional, VOptional, VOptional, VOptional);
	if (!m_ecBook.m_lpDispatch)
	{
		AfxMessageBox(_T("WorkSheet获取失败!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	if (CreateSheet(index) == FALSE)
	{
		return FALSE;
	}
	return TRUE;
}

BOOL CExcelOperation::Open(CString fileName)        //创建新的EXCEL应用程序并打开一个已经存在的文档。  
{
	if (CreateWorkbooks() == FALSE)
	{
		return FALSE;
	}
	return OpenWorkbook(fileName);
}

/*BOOL CExcelOperation::SetActiveWorkbook(short i)    //设置当前激活的文档。
{
}*/

//**********************保存文档*********************************************  
BOOL CExcelOperation::SaveWorkbook()                //文档是以打开形式，保存。  
{
	if (!m_ecBook.m_lpDispatch)
	{
		AfxMessageBox(_T("Book获取失败!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	m_ecBook.Save();
	return TRUE;
}
BOOL CExcelOperation::SaveWorkbookAs(CString fileName)//文档以创建形式，保存。  
{
	if (!m_ecBook.m_lpDispatch)
	{
		AfxMessageBox(_T("Book获取失败!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	COleVariant vTrue((short)TRUE),
		vFalse((short)FALSE),
		vOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	COleVariant vFileName(fileName);

	m_ecBook.SaveAs(
		vFileName,    //VARIANT* FileName  
		vOptional,    //VARIANT* FileFormat  
		vOptional,    //VARIANT* LockComments  
		vOptional,    //VARIANT* Password  
		vOptional,    //VARIANT* AddToRecentFiles  
		vOptional,    //VARIANT* WritePassword  
		0,    //VARIANT* ReadOnlyRecommended  
		vOptional,    //VARIANT* EmbedTrueTypeFonts  
		vOptional,    //VARIANT* SaveNativePictureFormat  
		vOptional,    //VARIANT* SaveFormsData  
		vOptional,    //VARIANT* SaveAsAOCELetter  
		vOptional    //VARIANT* ReadOnlyRecommended  
/*                vOptional,    //VARIANT* EmbedTrueTypeFonts
				vOptional,    //VARIANT* SaveNativePictureFormat
				vOptional,    //VARIANT* SaveFormsData
				vOptional    //VARIANT* SaveAsAOCELetter*/
	);
	return    TRUE;
}
BOOL CExcelOperation::CloseWorkbook()
{
	COleVariant vTrue((short)TRUE),
		vFalse((short)FALSE),
		vOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	m_ecBook.Close(vFalse,    // SaveChanges.  
		vTrue,            // OriginalFormat.  
		vFalse            // RouteDocument.  
	);
	m_ecBook = m_ecApp.get_ActiveWorkbook();
	if (!m_ecBook.m_lpDispatch)
	{
		AfxMessageBox(_T("Book获取失败!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	if (CreateSheet(1) == FALSE)
	{
		return FALSE;
	}
	return TRUE;
}
void CExcelOperation::CloseApp()
{
	SaveWorkbook();
	COleVariant vTrue((short)TRUE),
		vFalse((short)FALSE),
		vOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	//m_ecDoc.Save();  
	m_ecBook.Close(vOptional, vOptional, vOptional);
	m_ecBooks.Close();
	m_ecApp.Quit();
	if (m_ecRange.m_lpDispatch)
		m_ecRange.ReleaseDispatch();
	if (m_ecSheet.m_lpDispatch)
		m_ecSheet.ReleaseDispatch();
	if (m_ecSheets.m_lpDispatch)
		m_ecSheets.ReleaseDispatch();
	if (m_ecBook.m_lpDispatch)
		m_ecBook.ReleaseDispatch();
	if (m_ecBooks.m_lpDispatch)
		m_ecBooks.ReleaseDispatch();
	if (m_ecApp.m_lpDispatch)
		m_ecApp.ReleaseDispatch();
	
}

BOOL CExcelOperation::GetRangeAndValue(CString begin, CString end)
{
	if (!m_ecSheet.m_lpDispatch)
	{
		AfxMessageBox(_T("Sheet获取失败!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	m_ecRange = m_ecSheet.get_Range(COleVariant(begin), COleVariant(end));
	if (!m_ecRange.m_lpDispatch)
	{
		AfxMessageBox(_T("Range获取失败!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	ret = m_ecRange.get_Value2();//得到表格中的值  
	return TRUE;
}

void CExcelOperation::GetRowsAndCols(long &rows, long &cols)
{
	COleSafeArray sa(ret);
	sa.GetUBound(1, &rows);
	sa.GetUBound(2, &cols);
}

//只返回CString类型的，其他类型概视为错误  
BOOL CExcelOperation::GetTheValue(int rows, int cols, CString &dest)
{
	long rRows, rCols;
	long index[2];
	VARIANT val;
	COleSafeArray sa(ret);
	sa.GetUBound(1, &rRows);
	sa.GetUBound(2, &rCols);
	if (rows < 1 || cols < 1 || rRows < rows || rCols < cols)
	{
		AfxMessageBox(_T("出错点1"));
		return FALSE;
	}
	index[0] = rows;
	index[1] = cols;
	sa.GetElement(index, &val);
	if (val.vt != VT_BSTR)
	{
		CString str;
		str.Format(_T("出错点2, %d"), val.vt);
		AfxMessageBox(str);
		return FALSE;
	}
	dest = val.bstrVal;
	return TRUE;
}

//将beginS到endS之间设置为文本格式  
BOOL CExcelOperation::SetTextFormat(CString &beginS, CString &endS)
{
	if (GetRangeAndValue(beginS, endS))
	{
		m_ecRange.Select();
		m_ecRange.put_NumberFormatLocal(COleVariant(_T("@")));
		return TRUE;
	}
	return FALSE;
}

//将beginS到endS之间(必须是一列)设置为真正的文本格式  
BOOL CExcelOperation::SetRowToTextFormat(CString &beginS, CString &endS)
{
	if (GetRangeAndValue(beginS, endS))
	{
		m_ecRange.Select();
		CRange m_tempRange = m_ecSheet.get_Range(COleVariant(beginS), COleVariant(beginS));
		if (!m_tempRange.m_lpDispatch) return FALSE;
		COleVariant vTrue((short)TRUE),
			vFalse((short)FALSE);
		//int tempArray[2] = {1, 2};  
		COleSafeArray saRet;
		DWORD numElements = { 2 };
		saRet.Create(VT_I4, 1, &numElements);
		long index = 0;
		int val = 1;
		saRet.PutElement(&index, &val);
		index++;
		val = 2;
		saRet.PutElement(&index, &val);
		//m_tempRange.GetItem(COleVariant((short)5),COleVariant("A"));  
		m_ecRange.TextToColumns(m_tempRange.get_Item(COleVariant((short)1), COleVariant((short)1)), 1, 1, vFalse, vTrue, vFalse, vFalse, vFalse, vFalse, vFalse, saRet, vFalse, vFalse, vTrue);
		m_tempRange.ReleaseDispatch();
		return TRUE;
	}
	return FALSE;
}

void ConstCharConver(const char* pFileName, CString &pWideChar)
{ //计算char *数组大小，以字节为单位，一个汉字占两个字节 
	int charLen = strlen(pFileName);
	//计算多字节字符的大小，按字符计算。
	int len = MultiByteToWideChar(CP_ACP, 0, pFileName, charLen, NULL, 0);
	//为宽字节字符数组申请空间，数组大小为按字节计算的多字节字符大小 
	TCHAR *buf = new TCHAR[len + 1];
	//多字节编码转换成宽字节编码 
	MultiByteToWideChar(CP_ACP, 0, pFileName, charLen, buf, len);
	buf[len] = '\0'; //添加字符串结尾，注意不是len+1 
	//将TCHAR数组转换为CString
	pWideChar.Append(buf);
}

void CExcelOperation::setCellValue(std::string ccellIndexChar, CString valueChar)
{ 
	//const char* ccellIndexChar, const char* valueChar
	//CString cellIndex, value; 
	//ConstCharConver(ccellIndexChar, cellIndex); 
	//ConstCharConver(valueChar, value);
	m_ecRange = m_ecSheet.get_Range(_variant_t(ccellIndexChar.c_str()), _variant_t(ccellIndexChar.c_str()));
	m_ecRange.put_Value2(_variant_t(valueChar));
}

void CExcelOperation::SetCellMerge(CString index1, CString index2,CString value)
{ 
	//获取区域
	m_ecRange = m_ecSheet.get_Range(_variant_t(index1), _variant_t(index2));
	//合并单元格
	m_ecRange.Merge(_variant_t((long)1));
	m_ecRange.put_Value2(_variant_t(value));
}