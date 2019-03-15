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
		AfxMessageBox(_T("Application����ʧ�ܣ���ȷ����װ��wps�����ϰ汾!"), MB_OK | MB_ICONWARNING);
		pe.ReportError();
		throw &pe;
		return FALSE;
	}
	return TRUE;
}

BOOL CExcelOperation::CreateWorkbooks()               //����һ���µ�EXCEL����������  
{
	if (FALSE == CreateApp())
	{
		return FALSE;
	}
	m_ecBooks = m_ecApp.get_Workbooks();
	if (!m_ecBooks.m_lpDispatch)
	{
		AfxMessageBox(_T("WorkBooks����ʧ��!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	return TRUE;
}

BOOL CExcelOperation::CreateWorkbook()               //����һ���µ�EXCEL������  
{
	if (!m_ecBooks.m_lpDispatch)
	{
		AfxMessageBox(_T("WorkBooksΪ��!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}

	COleVariant vOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	m_ecBook = m_ecBooks.Add(vOptional);
	if (!m_ecBook.m_lpDispatch)
	{
		AfxMessageBox(_T("WorkBookΪ��!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	/*
		//�õ�document����
		m_wdDoc = m_wdApp.GetActiveDocument();
		if (!m_wdDoc.m_lpDispatch)
		{
			AfxMessageBox("Document��ȡʧ��!", MB_OK|MB_ICONWARNING);
			return FALSE;
		}
		//�õ�selection����
		m_wdSel = m_wdApp.GetSelection();
		if (!m_wdSel.m_lpDispatch)
		{
			AfxMessageBox("Select��ȡʧ��!", MB_OK|MB_ICONWARNING);
			return FALSE;
		}
		//�õ�Range����
		m_wdRange = m_wdDoc.Range(vOptional,vOptional);
		if(!m_wdRange.m_lpDispatch)
		{
			AfxMessageBox("Range��ȡʧ��!", MB_OK|MB_ICONWARNING);
			return FALSE;
		}
	*/
	return TRUE;
}

BOOL CExcelOperation::CreateWorksheets()                //����һ���µ�EXCEL��������  
{
	if (!m_ecBook.m_lpDispatch)
	{
		AfxMessageBox(_T("WorkBookΪ��!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	m_ecSheets = m_ecBook.get_Sheets();
	if (!m_ecSheets.m_lpDispatch)
	{
		AfxMessageBox(_T("WorkSheetsΪ��!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	return TRUE;
}

BOOL CExcelOperation::CreateWorksheet(short index)                //����һ���µ�EXCEL������  
{
	if (!m_ecSheets.m_lpDispatch)
	{
		AfxMessageBox(_T("WorkSheetsΪ��!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	m_ecSheet = m_ecSheets.get_Item(COleVariant(index));
	if (!m_ecSheet.m_lpDispatch)
	{
		AfxMessageBox(_T("WorkSheetΪ��!"), MB_OK | MB_ICONWARNING);
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

BOOL CExcelOperation::Create(short index)                        //�����µ�EXCELӦ�ó��򲢴���һ���¹������͹�����  
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

void CExcelOperation::ShowApp()                        //��ʾWORD�ĵ�  
{
	m_ecApp.put_Visible(TRUE);
}

void CExcelOperation::HideApp()                       //����word�ĵ�  
{
	m_ecApp.put_Visible(FALSE);
}

//**********************���ĵ�*********************************************  
BOOL CExcelOperation::OpenWorkbook(CString fileName, short index)
{
	if (!m_ecBooks.m_lpDispatch)
	{
		AfxMessageBox(_T("WorkSheetsΪ��!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	//COleVariant vFileName(_T(fileName));  
	COleVariant VOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	m_ecBook = m_ecBooks.Open(fileName, VOptional, VOptional, VOptional, VOptional, VOptional, VOptional, VOptional, VOptional, VOptional, VOptional, VOptional, VOptional, VOptional, VOptional);
	if (!m_ecBook.m_lpDispatch)
	{
		AfxMessageBox(_T("WorkSheet��ȡʧ��!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	if (CreateSheet(index) == FALSE)
	{
		return FALSE;
	}
	return TRUE;
}

BOOL CExcelOperation::Open(CString fileName)        //�����µ�EXCELӦ�ó��򲢴�һ���Ѿ����ڵ��ĵ���  
{
	if (CreateWorkbooks() == FALSE)
	{
		return FALSE;
	}
	return OpenWorkbook(fileName);
}

/*BOOL CExcelOperation::SetActiveWorkbook(short i)    //���õ�ǰ������ĵ���
{
}*/

//**********************�����ĵ�*********************************************  
BOOL CExcelOperation::SaveWorkbook()                //�ĵ����Դ���ʽ�����档  
{
	if (!m_ecBook.m_lpDispatch)
	{
		AfxMessageBox(_T("Book��ȡʧ��!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	m_ecBook.Save();
	return TRUE;
}
BOOL CExcelOperation::SaveWorkbookAs(CString fileName)//�ĵ��Դ�����ʽ�����档  
{
	if (!m_ecBook.m_lpDispatch)
	{
		AfxMessageBox(_T("Book��ȡʧ��!"), MB_OK | MB_ICONWARNING);
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
		AfxMessageBox(_T("Book��ȡʧ��!"), MB_OK | MB_ICONWARNING);
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
		AfxMessageBox(_T("Sheet��ȡʧ��!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	m_ecRange = m_ecSheet.get_Range(COleVariant(begin), COleVariant(end));
	if (!m_ecRange.m_lpDispatch)
	{
		AfxMessageBox(_T("Range��ȡʧ��!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	ret = m_ecRange.get_Value2();//�õ�����е�ֵ  
	return TRUE;
}

void CExcelOperation::GetRowsAndCols(long &rows, long &cols)
{
	COleSafeArray sa(ret);
	sa.GetUBound(1, &rows);
	sa.GetUBound(2, &cols);
}

//ֻ����CString���͵ģ��������͸���Ϊ����  
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
		AfxMessageBox(_T("�����1"));
		return FALSE;
	}
	index[0] = rows;
	index[1] = cols;
	sa.GetElement(index, &val);
	if (val.vt != VT_BSTR)
	{
		CString str;
		str.Format(_T("�����2, %d"), val.vt);
		AfxMessageBox(str);
		return FALSE;
	}
	dest = val.bstrVal;
	return TRUE;
}

//��beginS��endS֮������Ϊ�ı���ʽ  
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

//��beginS��endS֮��(������һ��)����Ϊ�������ı���ʽ  
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
{ //����char *�����С�����ֽ�Ϊ��λ��һ������ռ�����ֽ� 
	int charLen = strlen(pFileName);
	//������ֽ��ַ��Ĵ�С�����ַ����㡣
	int len = MultiByteToWideChar(CP_ACP, 0, pFileName, charLen, NULL, 0);
	//Ϊ���ֽ��ַ���������ռ䣬�����СΪ���ֽڼ���Ķ��ֽ��ַ���С 
	TCHAR *buf = new TCHAR[len + 1];
	//���ֽڱ���ת���ɿ��ֽڱ��� 
	MultiByteToWideChar(CP_ACP, 0, pFileName, charLen, buf, len);
	buf[len] = '\0'; //����ַ�����β��ע�ⲻ��len+1 
	//��TCHAR����ת��ΪCString
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
	//��ȡ����
	m_ecRange = m_ecSheet.get_Range(_variant_t(index1), _variant_t(index2));
	//�ϲ���Ԫ��
	m_ecRange.Merge(_variant_t((long)1));
	m_ecRange.put_Value2(_variant_t(value));
}