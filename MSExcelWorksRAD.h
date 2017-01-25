//---------------------------------------------------------------------------
#ifndef MSEXCELWORKS
#define MSEXCELWORKS

/*******************************************************************************
	����� ��� ����� � OLE-������� Excel.Application
    ������ �� 10.11.2014


    �������� ���������� ������ � �����������:
    1. ������� ������ ������ MSExcelWorks � ������� ����� � ���� ��� ������ ������
        MSExcelWorks msexcel;
        workbook = msexcel.OpenExcel();
        Variant worksheet1 = msexcel.GetSheet(workbook, 1);
    2. ���� ���������� �������� ������ ������ � �������, �� ������� ������ AnsiString:
        std::vector<AnsiString> format_body;
    3. ���� ���������� �������� ���� ������, ������, ������������ � �������, �� ������� ���������� Variant
        � ������ CELLFORMAT:
        Variant region_body;
        CELLFORMAT cf_body;
        cf_body.BorderStyle = CELLFORMAT::xlContinuous;
        cf_body.FontStyle = cfHead.FontStyle << CELLFORMAT::fsBold;
        cf_body.bSetFontColor = true;
        cf_body.FontColor = clRed;
        cf_body.bWrapText = false;
    4. ������� ����������-������ Variant ��� ������:
        Variant data_body;
        data_body = CreateVariantArray(1, FieldCount);
    5. ��������� ����������-������ � �������
        data_body.PutElement("Value", i, j);
    6. ��������� ������
        format_body.push_back("@");
        format_body.push_back("0");
        format_body.push_back("��.��.����");
        format_body.push_back("��:��:��");
    7. ������� ������ � ���� Excel
        region_body = msexcel.WriteTable(worksheet1, ArrayDataBody, 4 <����� ������>, 1 <����� �������>, format_body);
    8. ������ ������ �����:
        msexcel.SetRangeFormat(region_body, cf_body);
    9. ���������� ��������:
        msexcel.SetVisibleExcel(true, true);
    10. ��������� ������� ������:
        cf_body.clear();
        VarClear(ArrayDataHead);
        ArrayDataHead = NULL;

    ---
    CopyRange(worksheet, range_body, int Row, int Column, bool flag);
    ��� Row - ������ �� ������ ������ range_body
    Column - ������ �� ������� ������� range_body
    flag - ������� ���������� �� ������ � ��������������� ��������� ������

*******************************************************************************/

#include "system.hpp"
#include <utilcls.h>
#include "Comobj.hpp"
#include "sysutils.hpp"
#include "TaskutilsRAD.h"

class CellFormat {
public:
    __fastcall CellFormat();
    enum TDataAlignment {daDefault = 0, daTop = 2, daBottom = 4, daLeft = 2, daCenter = 3, daRight=4};
    enum TFontStyle {fsDefault = 0, fsNormal, fsBold, fsItalic, fsUnderline, fsStrikeOut};
    enum TBorderStyle {bsDefault = -1, bsNone = 0, xlContinuous = 1, bsBold = 7, bsDash1 = 2, bsDash2 = 3, bsDash3 = 4};
    enum TBorderLine {xlEdgeLeft = 7, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideHorizontal, xlInsideVertical};
    AnsiString DataFormat;
    TDataAlignment HorizontalAlignment;
    TDataAlignment VerticalAlignment;
    TBorderStyle BorderStyle;
    int FontSize;
    bool bSetFontColor;
    unsigned long FontColor;
    bool bSetBorderColor;
    unsigned long BorderColor;
    bool bSetFillColor;
    unsigned long FillColor;
    Set <char, 0, 9> FontStyle;
    Set <char, 0, 12> BordeLine;
    char ShrinkToFit;
    char bWrapText;

protected:

private:

};

__fastcall CellFormat::CellFormat()
{
    BorderColor = RGB(0,0,0);
    FontColor = BorderColor;
    FontSize = 0;
    DataFormat = "";
    HorizontalAlignment = daDefault;
    VerticalAlignment = daDefault;
    FontStyle.Clear();
    //FontStyle = FontStyle << fsNormal;
    BordeLine = BordeLine << xlEdgeTop << xlEdgeLeft<< xlEdgeBottom << xlEdgeRight << xlInsideHorizontal << xlInsideVertical;
    BorderStyle = bsDefault;
    bSetFontColor = false;
    bSetBorderColor = false;
    bSetFillColor = false;
    ShrinkToFit = -1;
    bWrapText = -1;
}

typedef CellFormat CELLFORMAT;
typedef std::vector<String> DATAFORMAT;

class MSExcelWorks
{
private:

public:
    //enum TDirection {Down = 1, Up, Left, Right};
    //enum xlBooleane {xlDefault = -1, xlFalse, xlTrue};
    Variant __fastcall OpenExcel(String TemplateName = "");     //������ Excel
    Variant __fastcall OpenWorksheetFromFile(String& FileName);
	void __fastcall CloseExcel(HWND hExW);                  // ������� �������� Excel
	void __fastcall CloseExcel();                           // ������� �������� Excel
    void __fastcall CloseWorkbook(Variant Book);
    void __fastcall SaveAsDocument(Variant workbook, String& FileName);
    Variant __fastcall AddSheet(Variant Book, String& SheetName, int SheetIndex = -1);
    Variant __fastcall AddBook(String TemplateName = "");
    void __fastcall SetActiveWorkbook(Variant Workbook);
    void __fastcall SetActiveWorksheet(Variant Worksheet);
    void __fastcall SetActiveRange(Variant Worksheet, int firstRow, int firstCol, int lastRow = 0, int lastCol = 0);
	Variant __fastcall GetActiveSheet();
	Variant __fastcall GetSheet(Variant Workbook, int SheetIndex = 1);
    Variant __fastcall GetRange(Variant Worksheet, int firstRow, int firstCol, int countRow = 1, int countCol = 1);
    Variant __fastcall GetRangeByName(Variant Worksheet, String& RangeName);
    Variant __fastcall GetRangeFromRange(Variant range, int firstRow, int firstCol, int countRow = 1, int countCol = 1);
    int __fastcall GetRangeRowsCount(Variant range);
    int __fastcall GetRangeColumnsCount(Variant range);
    String __fastcall GetRangeFormat(Variant range);

    void __fastcall InsertRows(Variant worksheet, int RowIndex, int RowsCount);
    Variant __fastcall WriteTable(Variant worksheet, const Variant &ArrayData,  int firstRow, int firstCol, std::vector<String> *DataFormat = NULL);

    Variant __fastcall WriteTable(Variant worksheet, const Variant &ArrayData, String CellName, std::vector<String> *DataFormat = NULL);
	Variant __fastcall WriteTableToRange(Variant range, const Variant &ArrayData,  int firstRow, int firstCol, std::vector<String> *DataFormat = NULL);


	Variant __fastcall WriteToRange(const String& txt, Variant range, String format = "");
	//Variant __fastcall WriteToRange(const String& txt, const String& sRangeName, String format = "");
	Variant __fastcall WriteToCell(Variant worksheet, const String& txt, int Row, int Col, String format = "");
    Variant __fastcall WriteToCell(Variant worksheet, const String& txt, String CellName, String format = "");
    Variant __fastcall WriteFormulaToCell(Variant wst, const String& txt, int Row, int Col, bool fBold = false);
    Variant __fastcall WriteFormula(Variant worksheet, const String& txt, int Row, int Col, int countRow = 1, int countCol = 1,  bool fBold = false);
    Variant __fastcall MergeCells(Variant worksheet, int firstRow, int firstCol, int lastRow, int lastCol);

    void __fastcall SetRangeFormat(Variant range, const CellFormat& cf, int firstRow, int firstCol, int countRow = 1, int countCol = 1);
    void __fastcall SetRangeFormat(Variant range, const CellFormat& cf);
	void __fastcall SetRangeDataFormat(Variant range, String format);
	void __fastcall RangeFormat(Variant wst, int firstRow, int CountRow, int firstCol, int lastCol, int Size, int Font_Color, int Inter_Color, bool Bold); // ������������� �����
    void __fastcall ClearFormats(Variant range);
	void __fastcall DrawBorders(Variant range, bool r7 = true, bool r8 = true, bool r9 = true, bool r10 = true, bool r11 = true, bool r12 = true); // ������������� ������� ������� ������� ������ ����� ��������� ���������
    void __fastcall SetRangeColumnsFormat(Variant range, const std::vector<String> &cf);
	void __fastcall RangeShtrich(Variant wst, int firstRow, int CountRow, int firstCol, int lastCol, int Shtrich);
	void __fastcall SetColumnsAutofit(Variant range);
    void __fastcall SetAutoFilter(Variant range);
	void __fastcall SetRowsAutofit(Variant range);

    void __fastcall SetRowHeight(Variant range, int Height);
	//void __fastcall SetRowHeight(Variant worksheet, int ColumnIndex, int width);
    void __fastcall SetColumnWidth(Variant worksheet, int ColumnIndex, int width);
    void __fastcall SetColumnWidth(Variant range, int Width);

	void __fastcall SetVisibleExcel(bool fVisible = true, bool fForeground = true); // ���������� �������� Excel
	void __fastcall DateTimeCreateDoc(Variant wst, int Row, int Col);

    Variant __fastcall ReadRange(Variant worksheet, int firstRow, int firstCol, int lastRow, int lastCol);
    String __fastcall ReadCell(Variant worksheet, int Row, int Col); // ������ ������ �� ������ ������� Excell
    Variant __fastcall ReadCellFormula(Variant worksheet, int Row, int Col); // ������ ������ �� ������ ������� Excell

	std::vector<String> __fastcall GetDataFormat(Variant ArrayData, int RowIndex);//, std::vector<CELLFORMAT> *formats);
    Variant CreateVariantArray(int RowCount, int ColCount);
    void RedimVariantArray(Variant ArrayData, int RowCount);
    void __fastcall CopyArray(const Variant SrcArrayData, Variant* ArrayData,  int srcFirstRow, int srcFirstCol, int srcLastRow, int srcLastCol, int dstFirstRow, int dstFirstCol);

    inline int GetRangeFirstRow(Variant range);
    inline int GetRangeFirstColumn(Variant range);

    Variant __fastcall CopyRange(Variant worksheet, Variant range, int Row = 0, int Column = 0, bool fCopyData = true);
    Variant __fastcall CopyRange(Variant worksheet, String sRangeName, int Row = 0, int Column = 0, bool fCopyData = true);
	Variant __fastcall CopyRangeTo(Variant worksheet, Variant range, int Row = 1, int Column = 1, bool fCopyData = true);

    Variant ExcelApp;
    Variant WorkBooks;
};

//----------------------------------------------------------------------------
// ��������� ����� � ����
void __fastcall MSExcelWorks::SaveAsDocument(Variant workbook, String& FileName)
{
    try {
        // ���������� ��������� � ����
		workbook.OleProcedure("SaveAs", WideString(FileName));
	} catch (...) {
        throw Exception("�� ������� ��������� �������� � ������ " + FileName);
    }
}

//----------------------------------------------------------------------------
// ��������/��������� ����������
void __fastcall MSExcelWorks::SetAutoFilter(Variant range)
{
    // expression .AutoFilter(Field, Criteria1, Operator, Criteria2, VisibleDropDown)
    range.OleFunction("AutoFilter");
}

//----------------------------------------------------------------------------
// ��������� ������� ������� �� �����������
void __fastcall MSExcelWorks::SetColumnsAutofit(Variant range)
{
    range.OlePropertyGet("Columns").OleProcedure("AutoFit");
}

//----------------------------------------------------------------------------
//
void __fastcall MSExcelWorks::SetColumnWidth(Variant worksheet, int ColumnIndex, int width)
{
    worksheet.OlePropertyGet("Columns").OlePropertyGet("Item", ColumnIndex).OlePropertySet("ColumnWidth", width);
}

//----------------------------------------------------------------------------
// ��������� ������ ������� �� �����������
void __fastcall MSExcelWorks::SetRowsAutofit(Variant range)
{
    range.OlePropertyGet("Rows").OleProcedure("AutoFit");
}

//----------------------------------------------------------------------------
//
void __fastcall MSExcelWorks::SetRowHeight(Variant range, int Height)
{
    //range.OlePropertyGet("Rows").OlePropertySet("RowWidth", Width);
    range.OlePropertySet("RowHeight", Height);
}

//----------------------------------------------------------------------------
//
void __fastcall MSExcelWorks::SetColumnWidth(Variant range, int Width)
{
    range.OlePropertySet("ColumnWidth", Width);
}


//----------------------------------------------------------------------------
// ��������� ���������� ����� � Range
int __fastcall MSExcelWorks::GetRangeRowsCount(Variant range)
{
	return range.OlePropertyGet("Rows").OlePropertyGet("Count");
}

//----------------------------------------------------------------------------
// ��������� ���������� �������� � Range
int __fastcall MSExcelWorks::GetRangeColumnsCount(Variant range)
{
    return range.OlePropertyGet("Columns").OlePropertyGet("Count");
}

//----------------------------------------------------------------------------
// ���������� ������ Range
String __fastcall MSExcelWorks::GetRangeFormat(Variant range)
{
    String result;
    result = range.OlePropertyGet("NumberFormat");
    return result;
}

//----------------------------------------------------------------------------
// ������� ������ Range
void __fastcall MSExcelWorks::ClearFormats(Variant range)
{
    range.OleProcedure("ClearFormats");
}

//----------------------------------------------------------------------------
// ��������� Range
Variant __fastcall MSExcelWorks::GetRange(Variant Worksheet, int firstRow, int firstCol, int countRow, int countCol)
{
    Variant sell_left_top = Worksheet.OlePropertyGet("Cells", firstRow, firstCol);
	Variant sell_right_bottom = Worksheet.OlePropertyGet("Cells", firstRow+countRow-1, firstCol+countCol-1);
	Variant range = Worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);
    return range;
}

//----------------------------------------------------------------------------
// ��������� Range �� �����
Variant __fastcall MSExcelWorks::GetRangeByName(Variant Worksheet, String& RangeName)
{
    try {
        //Variant range = Worksheet.OlePropertyGet("Cells", RangeName);       // ������� ����� ��������� �� �����
        Variant range = Worksheet.OlePropertyGet("Range", RangeName);       // ������� ����� ��������� �� �����
        return range;
    } catch (EOleSysError &e) {
        throw Exception("�� ������� ���������� Range �� ����� " + RangeName);
        //return Unassigned;
    }

/*    Variant Workbook = Worksheet.OlePropertyGet("Parent");

    Variant vNames = Workbook.OlePropertyGet("Names");
    String strName;
    //String strRefersTo, strSheetName;
    //String srCellName("");
    int nNameCount = vNames.OlePropertyGet("Count");

    for(int i=1; i < nNameCount + 1; i++) {
        strName = vNames.OleFunction("Item", i).OlePropertyGet("Name");
        int n = strName.Pos("!");
        strName = strName.SubString(n+1, strName.Length() - n);

         if(strName == RangeName) {
            return vNames.OleFunction("Item", i).OlePropertyGet("RefersToRange");

            //int nRow = vNames.OleFunction("Item", i).OlePropertyGet("RefersToRange").OlePropertyGet("Row");
            //int nCol = vNames.OleFunction("Item", i).OlePropertyGet("RefersToRange").OlePropertyGet("Column");
            //strRefersTo = vNames.OleFunction("Item", i).OlePropertyGet("RefersTo");
            //strSheetName = strRefersTo.SubString(2, strRefersTo.Pos("!") - 2);
            //return WriteTable(worksheet, ArrayData,  nRow, nCol, DataFormat);
        }
    }
    //VarClear(
    return Unassigned; */
}

//----------------------------------------------------------------------------
// ��������� Range ������ ��������� Range
Variant __fastcall MSExcelWorks::GetRangeFromRange(Variant range, int firstRow, int firstCol, int countRow, int countCol)
{
    Variant Cells = range.OlePropertyGet("Cells");
    Variant sell_left_top = Cells.OlePropertyGet("Item", firstRow, firstCol);
	Variant sell_right_bottom = Cells.OlePropertyGet("Item", firstRow+countRow-1, firstCol+countCol-1);
    Variant Worksheet = range.OlePropertyGet("Worksheet");
    return Worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);
}

//------------------------------------------------------------------------------
// ��������� ������ ����� ���������� �� Range                   !!! �� ���������
Variant __fastcall MSExcelWorks::ReadRange(Variant worksheet, int firstRow, int firstCol, int lastRow, int lastCol)
{

    Variant ArrayData;
    Variant sell_left_top = worksheet.OlePropertyGet("Cells", firstRow, firstCol);
	Variant sell_right_bottom = worksheet.OlePropertyGet("Cells", lastRow, lastCol);
	Variant range = worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);
    ArrayData = range.OlePropertyGet("Value");

    return ArrayData;
}

//------------------------------------------------------------------------------
// �������� ����� ������� � ������ ������                       !!! �� ��������
void __fastcall MSExcelWorks::CopyArray(const Variant SrcArrayData, Variant* ArrayData,  int srcFirstRow, int srcFirstCol, int srcLastRow, int srcLastCol, int dstFirstRow, int dstFirstCol)
{
    for (int i = srcFirstRow; i <= srcLastRow; i ++)
        for (int j = srcFirstCol; j <= srcLastCol; j++) {
            ArrayData->PutElement(SrcArrayData.GetElement(i,j), i, j);   // ����������!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        }
}

//----------------------------------------------------------------------------
// ������ ������ ����� � Range
void __fastcall MSExcelWorks::SetRangeDataFormat(Variant range, String format)
{
	if (format != "") {   // "m/d/yyyy" "@" "0.00" "General"
		range.OlePropertySet("NumberFormat", WideString(format));
	}
}

//----------------------------------------------------------------------------
// ������ ������ ����� � Range �������� ������������
void __fastcall MSExcelWorks::SetRangeFormat(Variant range,  const CellFormat& cf, int firstRow, int firstCol, int countRow, int countCol)
{
	// ������������ ��������� range
    //Variant cell_left_top = range.OlePropertyGet("Cells", firstRow, firstCol);
	//Variant cell_right_bottom = range.OlePropertyGet("Cells", firstRow + countRow - 1, firstCol + countCol - 1);
	//Variant range_tmp = range.OlePropertyGet("Range", cell_left_top, cell_right_bottom);

    Variant range_tmp = GetRangeFromRange(range, firstRow, firstCol, countRow, countCol);
 	SetRangeFormat(range_tmp, cf);
}

//----------------------------------------------------------------------------
// ������ ������ ����� � Range
void __fastcall MSExcelWorks::SetRangeFormat(Variant range, const CellFormat& cf)
{
    if (cf.DataFormat != "") {   // "m/d/yyyy" "@" "0.00" "General"
        try {
            range.OlePropertySet("NumberFormat", cf.DataFormat);
        }
        catch (...) {
        }
    }

    // ��� 2 - �� ������ ����, 3 - �� ������, 4 - �� �������)
    if (cf.HorizontalAlignment)
        range.OlePropertySet ("HorizontalAlignment", cf.HorizontalAlignment);

    if (cf.VerticalAlignment)
        range.OlePropertySet ("VerticalAlignment", cf.VerticalAlignment - 1);

    if (cf.ShrinkToFit > -1) {  // ���������� ������� ������ �� ������ ������
        range.OlePropertySet("ShrinkToFit", cf.ShrinkToFit);
    }

    if (cf.bWrapText > -1) // ���������� �� ������
        range.OlePropertySet("WrapText", cf.bWrapText);


    Variant font = range.OlePropertyGet("Font");
    if (cf.FontStyle.Contains(CellFormat::fsNormal)) {
        font.OlePropertySet("Bold", false);
        font.OlePropertySet("Italic", false);
        font.OlePropertySet("Underline", false);
    }
    else {
        if (cf.FontStyle.Contains(CellFormat::fsBold))
            font.OlePropertySet("Bold", true);
        if (cf.FontStyle.Contains(CellFormat::fsItalic))
            font.OlePropertySet("Italic", true);
        if (cf.FontStyle.Contains(CellFormat::fsUnderline))
            font.OlePropertySet("Underline", true);
    }

    if (cf.FontSize > 0)
        font.OlePropertySet("Size", cf.FontSize);

    if (cf.bSetFontColor)     // ���� ������ � ������
        font.OlePropertySet("Color", cf.FontColor);

    if (cf.bSetFillColor)
        range.OlePropertyGet("Interior").OlePropertySet("Color", cf.FillColor);

    //if (cf.bWrapText)






    // ��������!!!!!
    if (cf.BorderStyle >= 0) {
        Variant borders = range.OlePropertyGet("Borders");
        //borders.OlePropertySet("LineStyle", cf.BorderStyle);
        for (int i = CellFormat::xlEdgeLeft; i <= CellFormat::xlInsideVertical; i++)
        {
            if (cf.BordeLine.Contains(i))
      	        //try {range.OlePropertyGet("Borders", i).OlePropertySet("LineStyle", cf.BorderStyle);} catch(...) {};
      	        range.OlePropertyGet("Borders", i).OlePropertySet("LineStyle", cf.BorderStyle);
        }
    }

}
//----------------------------------------------------------------------------
// ������ ������ �������� � Range
void __fastcall MSExcelWorks::SetRangeColumnsFormat(Variant range, const std::vector<String> &df)
{
	int lastCol = GetRangeColumnsCount(range);
	Variant Columns = range.OlePropertyGet("Columns");
    int SizeDF = df.size();
	try {
		for (int i = 0; i < lastCol && i < SizeDF; i++) {         // ��� ������� ������� �� ������ ������
			//AnsiString s = "@";
			//SetRangeDataFormat(Columns.OlePropertyGet("Item", i+1), s);
			//SetRangeDataFormat(Columns.OlePropertyGet("Item", i+1), L"0");
			//SetRangeDataFormat(Columns.OlePropertyGet("Item", i+1), "@");
			SetRangeDataFormat(Columns.OlePropertyGet("Item", i+1), df[i]);
		}
	} catch (...) {
	}
}

//----------------------------------------------------------------------------
// ������� ������ �����
void __fastcall MSExcelWorks::InsertRows(Variant worksheet, int RowIndex, int RowsCount)
{
    Variant Rows = worksheet.OlePropertyGet("Rows");
    //Variant Row = Rows.OlePropertyGet("Item", RowIndex, 5);
    Variant Row = Rows.OlePropertyGet("Range", IntToStr(RowIndex) + ":" + IntToStr(RowIndex+RowsCount-1));
    Row.OleProcedure("Insert", 0xFFFFEFE7, 0);

    // OleProcedure("Insert", xlDown, xlFormatFromLeftOrAbove);
    // xlDown = 0xFFFFEFE7,
    // xlToLeft = 0xFFFFEFC1,
    // xlToRight = 0xFFFFEFBF,
    // xlUp = 0xFFFFEFBE
    // xlFormatFromLeftOrAbove = 0
    // xlFormatFromRightOrBelow = 1
}

//----------------------------------------------------------------------------
// ������� ���������� ������� � ��������� �� ����� �������
Variant __fastcall MSExcelWorks::WriteTableToRange(Variant range, const Variant &ArrayData,  int firstRow, int firstCol, std::vector<String> *DataFormat)
{
    if (DataFormat != NULL) {
		SetRangeColumnsFormat(range, *DataFormat);
    }

	range.OlePropertySet("Value", ArrayData);		// ����� ������ � ��������. ��� ����� ���� ����� ��� ������� ���������� ������
    return range;
}

//----------------------------------------------------------------------------
// ������� ���������� ������� � ��������� �� ����� �������
Variant __fastcall MSExcelWorks::WriteTable(Variant worksheet, const Variant &ArrayData,  int firstRow, int firstCol, std::vector<String> *DataFormat)
{
    Variant ArrayRowsCount = VarArrayHighBound(ArrayData, 1) - VarArrayLowBound(ArrayData, 1)+1;
    Variant ArrayColsCount = VarArrayHighBound(ArrayData, 2) - VarArrayLowBound(ArrayData, 2)+1;

	int lastRow = firstRow + ArrayRowsCount - 1; // firstRow .. lastRow
	int lastCol = firstCol + ArrayColsCount - 1; // firstCol  .. lastCol

    Variant sell_left_top = worksheet.OlePropertyGet("Cells", firstRow, firstCol);
	Variant sell_right_bottom = worksheet.OlePropertyGet("Cells", lastRow, lastCol);
	Variant range = worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);

	return WriteTableToRange(range, ArrayData, firstRow, firstCol, DataFormat);

/*    if (DataFormat != NULL) {
        SetRangeColumnsFormat(range, *DataFormat);
    }

    range.OlePropertySet("Value", ArrayData);		// ����� ������ � ��������. ��� ����� ���� ����� ��� ������� ���������� ������
    return range; */
}

//----------------------------------------------------------------------------
// ������� ���������� ������� � ��������� (�����������) �� ����� �������
Variant __fastcall MSExcelWorks::WriteTable(Variant worksheet, const Variant &ArrayData, String CellName, std::vector<String> *DataFormat)
{
    Variant vNames = worksheet.OlePropertyGet("Names");

    //int nNameCount = vNames.OlePropertyGet("Count");

    String strName;

    Variant range = GetRangeByName(worksheet, CellName);

    int nRow = range.OlePropertyGet("Row");
    int nCol = range.OlePropertyGet("Column");

    return WriteTable(worksheet, ArrayData,  nRow, nCol, DataFormat);
}

//------------------------------------------------------------------------------
// ���������� � �������� ������ ������� ���� � �����
void __fastcall MSExcelWorks::DateTimeCreateDoc(Variant wst, int Row, int Col)
{
	String txt = "���� �������� ���������: "+DateTimeToStr(Now());
	WriteToCell(wst, txt, Row, Col);
}

//------------------------------------------------------------------------------
// ������� � ������ ������ � �������� �������
Variant __fastcall MSExcelWorks::WriteToRange(const String &txt, Variant range, String format)
{
    if (range.IsEmpty())
        return range;
	range.OlePropertySet("Value", WideString(txt));
	if (format != "")
		range.OlePropertySet("NumberFormat", WideString(format));   // ������������� ������ ������ ��� ������
	return range;
}

/*//------------------------------------------------------------------------------
// ������� � ������ ������ � �������� �������   2
Variant __fastcall MSExcelWorks::WriteToRange(const String& txt, const String& sRangeName, String format)
{
    Variant range = GetRangeByName(sRangeName);
    return WriteToRange(txt, range, format);
}   */

//------------------------------------------------------------------------------
// ������� � ������ ������ � �������� �������
Variant __fastcall MSExcelWorks::WriteToCell(Variant worksheet, const String &txt, int Row, int Col, String format)
{
	Variant range = worksheet.OlePropertyGet("Cells").OlePropertyGet("Item", Row, Col);
    return WriteToRange(txt, range, format);
}

//------------------------------------------------------------------------------
// ������� � ������ � �������� ������ ������ � �������� �������
Variant __fastcall MSExcelWorks::WriteToCell(Variant worksheet, const String &txt, String CellName, String format)
{
    Variant range = GetRangeByName(worksheet, CellName);

    return WriteToRange(txt, range, format);
}

//------------------------------------------------------------------------------
// ������� ������� � ������
Variant __fastcall MSExcelWorks::WriteFormulaToCell(Variant wst, const String &txt, int Row, int Col, bool fBold)
{
	Variant range = wst.OlePropertyGet("Cells").OlePropertyGet("Item", Row, Col);
	range.OlePropertySet("FormulaR1C1", WideString(txt));
    range.OlePropertyGet("Font").OlePropertySet("Bold",fBold);

	return range;
}

//------------------------------------------------------------------------------
// ������� ������� � ������
Variant __fastcall MSExcelWorks::WriteFormula(Variant worksheet, const String &txt, int Row, int Col, int countRow, int countCol,  bool fBold)
{
    Variant sell_left_top = worksheet.OlePropertyGet("Cells", Row, Col);
	Variant sell_right_bottom = worksheet.OlePropertyGet("Cells", Row+countRow-1, Col+countCol-1);
	Variant range = worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);
	//Variant range = GetRange(wst, Row, Col, Row+countRow-1, Col+countCol-1);
    // wst.OlePropertyGet("Cells").OlePropertyGet("Item", Row, Col);
	range.OlePropertySet("FormulaR1C1", WideString(txt));
    range.OlePropertyGet("Font").OlePropertySet("Bold",fBold);

	return range;
}

/*//------------------------------------------------------------------------------
// ������� ������� � ������
Variant __fastcall MSExcelWorks::WriteFormula(Variant wst, int firstRow, int firstCol, int lastRow = 0, int lastCol = 0);
*/

//---------------------------------------------------------------------------
// ��������� ������ �� ������ ������� Excel
String __fastcall MSExcelWorks::ReadCell(Variant worksheet, int Row, int Col)
{
    return worksheet.OlePropertyGet("Cells").OlePropertyGet("Item", Row, Col);
}

//---------------------------------------------------------------------------
// ��������� ������ �� ������ ������� Excel
Variant __fastcall MSExcelWorks::ReadCellFormula(Variant worksheet, int Row, int Col)
{
    return worksheet.OlePropertyGet("Cells").OlePropertyGet("Item", Row, Col).OlePropertyGet("Formula");
}

//------------------------------------------------------------------------------
// ������� ������ ���� varVariant
Variant MSExcelWorks::CreateVariantArray(int RowCount, int ColCount)
{
    int Bounds[4] = {1, RowCount, 1, ColCount};
//    return VarArrayCreate(Bounds, varString);
    return VarArrayCreate(Bounds, 3, varVariant);
}

//------------------------------------------------------------------------------
// Increase the length of the variant array
void MSExcelWorks::RedimVariantArray(Variant ArrayData, int RowCount)
{
    VarArrayRedim(ArrayData, RowCount);
}

//------------------------------------------------------------------------------
// ���������� ������ ������ � �������
std::vector<String> __fastcall MSExcelWorks::GetDataFormat(Variant ArrayData, int RowIndex)
{
    std::vector<String> formats;

    int firstCol = VarArrayLowBound(ArrayData, 2);
    int lastCol = VarArrayHighBound(ArrayData, 2);

    formats.reserve(lastCol - firstCol + 1);
    for (int i = firstCol; i <= lastCol; i++) {         // ��� ������� ������� �� ������ ������
        String s = ArrayData.GetElement(RowIndex, i);
        String format;

        // ������������� ������ ����� ��� ������
        if ( IsDate(s.c_str()) )
            format = "��.��.����"; //"m/d/yyyy";    //"��.��.����";
        else if ( IsFloat(s.c_str()) )
            format = "0.00";
        else if ( IsInt(s.c_str()) )   // �������, ��� ��� ������� ����� ����� �������� �� 0
             format = "0";
        else
            format = "@";   // "General"

        formats.push_back(format);
    }
    return formats;
}

//------------------------------------------------------------------------------
// ������������� �������� ����� Excel
void __fastcall MSExcelWorks::SetActiveWorkbook(Variant Workbook)
{
    Workbook.OleProcedure("Activate");
}

//------------------------------------------------------------------------------
// ������������� �������� ���� � ����� Excel
void __fastcall MSExcelWorks::SetActiveWorksheet(Variant worksheet)
{
    worksheet.OleProcedure("Activate");
}

//------------------------------------------------------------------------------
// ������������� ������ �� ������ ������ Excel
void __fastcall MSExcelWorks::SetActiveRange(Variant Worksheet, int firstRow, int firstCol, int lastRow, int lastCol)
{
    GetRange(Worksheet, firstRow, firstCol, lastRow, lastCol).OleProcedure("Select");
}

//------------------------------------------------------------------------------
// ����������� ��������� Excel
// � ���������� ������ �� ������
void __fastcall MSExcelWorks::SetVisibleExcel(bool fVisible, bool fForeground)
{
    String ExcelCaption = ExcelApp.OlePropertyGet("Caption");
	HWND hExcelWindow = FindWindow(L"XLMAIN", ExcelCaption.c_str());

	if(!ExcelApp.IsEmpty()) {
    	ExcelApp.OlePropertySet("Visible", fVisible);
        //SetWindowPos(hExcelWindow, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE || SWP_NOSIZE);
    	// ExcelApp.OlePropertySet("Activate", true);
        ExcelApp.OlePropertySet("DisplayAlerts", true);         // ���������� ��������������

        if (fVisible && fForeground) {
            // ���� �� �������� ����
            ExcelApp.OlePropertySet("UserControl", true);
            //SetForegroundWindow(hExcelWindow);
            //SetWindowPos(hExcelWindow, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE || SWP_NOSIZE);
            //SetWindowPos(hExcelWindow, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE || SWP_NOSIZE);
            SendMessage(hExcelWindow,WM_SYSCOMMAND,SC_MAXIMIZE,0);
        }
    }
}

//---------------------------------------------------------------------------
//������ Excel � ����������� ������� �����
Variant __fastcall MSExcelWorks::OpenExcel(String TemplateName)
{
    if(ExcelApp.IsEmpty())
    {
        try {
    	    ExcelApp = CreateOleObject("Excel.Application");
            ExcelApp.OlePropertySet("DisplayAlerts", false);
            WorkBooks = ExcelApp.OlePropertyGet("Workbooks");
            ExcelApp.OlePropertySet("Visible", false);
        } catch (Exception &exception) {
            //Application->MessageBox("�� ������� ������� �������� Excel","��������",MB_ICONSTOP + MB_OK);
            VariantClear(ExcelApp);
            throw Exception("�� ������� ������� �������� Excel");
            //throw Exception(exception);
            //return Unassigned;
        }
  	}
 	return AddBook(TemplateName);
}

//---------------------------------------------------------------------------
// ���������� (������ ��� �������� �� �������) ������� �����
Variant __fastcall MSExcelWorks::AddBook(String TemplateName)
{
    Variant Book;
    if (TemplateName != "") {   // ������� �� ������ �������
	    try {
            Book = WorkBooks.OlePropertyGet("Open", TemplateName);
  	    } catch (Exception &exception)	{
    	    //MessageBoxStop(("������ �������� ������� ������ " + TemplateName).c_str());
            //throw new Exception("������ �������� ������� ������: " + TemplateName);
            throw Exception("������ �������� ������� ������: " + TemplateName);
            //return NULL;
            //return Unassigned;
        }
    }
    else {                      // ������� ������ ��������
        try {
            Book = WorkBooks.OleFunction("Add");
        } catch (Exception &exception) {
            VariantClear(ExcelApp);
    	    //MessageBoxStop("������ �������� ��������� Excel");
            throw Exception("������ �������� ��������� Excel");
        }
    }
  	return Book;
}

//---------------------------------------------------------------------------
// ���������� �����
Variant __fastcall MSExcelWorks::AddSheet(Variant Book, String& SheetName, int SheetIndex)
{
    Variant Sheets = Book.OlePropertyGet("Worksheets");
    Variant Sheet;

    Variant position;
    if (SheetIndex <= 0) {  // ���������� � ����� ����� (�� ���������)
        position = Sheets.OlePropertyGet("Count");
        Variant After = Sheets.OlePropertyGet("Item", position);

        Sheet = Sheets.OleFunction("Add", VT_EMPTY, After);
        //Sheet = Sheets.OleFunction("Add", EmptyParam, After);
    } else if (SheetIndex == 1){    // ���������� � ������ �����
        Sheet = Sheets.OleFunction("Add");
    } else {                        // ���������� � ������� SheetIndex
        position = SheetIndex - 1;
        Variant After = Sheets.OlePropertyGet("Item", position);
        Sheet = Sheets.OleFunction("Add", VT_EMPTY, After);
        //Sheet = Sheets.OleFunction("Add", EmptyParam, After);
    }

    Sheet.OlePropertySet("Name", WideString(SheetName));
    return Sheet;
}

//---------------------------------------------------------------------------
// ���������� �������� �������� Excel
Variant __fastcall MSExcelWorks::GetActiveSheet()
{
    if (!VarIsNull(ExcelApp)) {  // VarIsEmpty
        return ExcelApp.OlePropertyGet("ActiveSheet");
    } else {
        return NULL;
    }
}

//---------------------------------------------------------------------------
// �������� �� ������                               // ����������, ����� ��������� ����� �����
Variant __fastcall MSExcelWorks::GetSheet(Variant Workbook, int SheetIndex)
{
    if (!VarIsNull(ExcelApp)) {  // VarIsEmpty
        Variant Worksheet = Workbook.OlePropertyGet("Worksheets", SheetIndex);
        return  Worksheet;
        //return  WorkSheets.OlePropertyGet("Item", SheetIndex);
    } else {
        return NULL;
    }
}

//----------------------------------------------------------------------------
// ������� �������� Excel
void __fastcall MSExcelWorks::CloseExcel()
{
    if(!ExcelApp.IsEmpty())
    {
/*        try{
            int WorkbooksCount = WorkBooks.OlePropertyGet("Count");
            for (int i = 1; i <= WorkbooksCount; i++)
                ExcelApp.OlePropertyGet("WorkBooks", i).OleProcedure("Close");
                //WorkBooks.OlePropertyGet("WorkBooks", i).OleProcedure("Close");
        } catch(...){ }   */
        

        WorkBooks.OleProcedure("Close");     // ��� ���������� ���������
        ExcelApp.OleProcedure("Quit");

        /*String ExcelCaption = ExcelApp.OlePropertyGet("Caption");
        HWND hExcelWindow = FindWindow("XLMAIN", ExcelCaption.c_str());
        if ( hExcelWindow != NULL)
            SendMessage(hExcelWindow, WM_DESTROY, 0, 0);  */

        VariantClear(WorkBooks);
        VariantClear(ExcelApp);
    }
}

//----------------------------------------------------------------------------
// ������� �������� ������� ����� Excel
void __fastcall MSExcelWorks::CloseWorkbook(Variant Book)
{
    if(!Book.IsEmpty())
    {
        Book.OleProcedure("Close", false);
        VariantClear(Book);
    }
}

//------------------------------------------------------------------------------
// ������������� ������� ������� ������� ������ ����� ��������� ���������
void __fastcall MSExcelWorks::DrawBorders(Variant range, bool r7, bool r8, bool r9, bool r10, bool r11, bool r12)
{
	// r7  - ������� �����
  	// r8  - ������� ������
  	// r9  - ������� �����
  	// r10 - ������� ������
  	// r11 - ����� ������
  	// r12 - ����� �������

  	if (r7) range.OlePropertyGet("Borders",7).OlePropertySet("LineStyle", 1);
  	if (r8) range.OlePropertyGet("Borders",8).OlePropertySet("LineStyle", 1);
  	if (r9) range.OlePropertyGet("Borders",9).OlePropertySet("LineStyle", 1);
  	if (r10) range.OlePropertyGet("Borders",10).OlePropertySet("LineStyle", 1);
  	if (r11) try { range.OlePropertyGet("Borders",11).OlePropertySet("LineStyle", 1); } catch(...) {}
  	if (r12) try { range.OlePropertyGet("Borders",12).OlePropertySet("LineStyle", 1);} catch(...) {}
}

//------------------------------------------------------------------------------
// ���������
void __fastcall MSExcelWorks::RangeShtrich(Variant wst, int firstRow, int firstCol, int CountRow, int lastCol, int Shtrich)
{
  int lastRow = firstRow + CountRow - 1; // ����� �������� ������
  Variant sell_left_top = wst.OlePropertyGet("Cells", firstRow, firstCol),
          sell_right_bottom = wst.OlePropertyGet("Cells", lastRow, lastCol), diap;
  diap = wst.OlePropertyGet("Range", sell_left_top, sell_right_bottom);
  diap.OlePropertyGet("Interior").OlePropertySet("Pattern", Shtrich);
}

//------------------------------------------------------------------------------
// ����������� �����
Variant __fastcall MSExcelWorks::MergeCells(Variant worksheet, int firstRow, int firstCol, int lastRow, int lastCol)
{
    //int lastRow = firstRow + CountRow - 1; // ����� �������� ������
    Variant sell_left_top = worksheet.OlePropertyGet("Cells", firstRow, firstCol);
    Variant sell_right_bottom = worksheet.OlePropertyGet("Cells", lastRow, lastCol);
    Variant range = worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);
    range.OlePropertySet("MergeCells", true);
    // Across - True, ����� ���������� ������ � ������ ������ ���������� ��������� ��� ��������� �������. �������� �� ��������� False.

    return range;
}

//------------------------------------------------------------------------------
//
Variant __fastcall MSExcelWorks::OpenWorksheetFromFile(String& FileName)
{
        Variant workbook;
        Variant worksheet;
        workbook = OpenExcel(FileName);
        worksheet = GetSheet(workbook, 1);
        return worksheet;
}

//------------------------------------------------------------------------------
//
inline int MSExcelWorks::GetRangeFirstRow(Variant range)
{
    return range.OlePropertyGet("Row");
}

//------------------------------------------------------------------------------
//
inline int MSExcelWorks::GetRangeFirstColumn(Variant range)
{
    return range.OlePropertyGet("Column");
}

//---------------------------------------------------------------------------
//
Variant __fastcall MSExcelWorks::CopyRange(Variant worksheet, Variant range, int Row, int Column, bool fCopyData)
{
    //Variant worksheet = range.OlePropertyGet("Worksheet");

    int rowsCount = GetRangeRowsCount(range);
    int colsCount = GetRangeColumnsCount(range);

    int rangeFirstRow = GetRangeFirstRow(range);
    int rangeFirstColumn = GetRangeFirstColumn(range);
    int firstRow = rangeFirstRow + rowsCount + Row;
    int firstCol = rangeFirstColumn + Column;
//    int firstRow = rangeFirstRow + rowsCount + Row;
//   int firstCol = GetRangeFirstColumn(range) + Column;

    int lastRow = firstRow + rowsCount-1;
    int lastCol = firstCol + colsCount-1;


    Variant sell_left_top = worksheet.OlePropertyGet("Cells", firstRow, firstCol);
	Variant sell_right_bottom = worksheet.OlePropertyGet("Cells", lastRow, lastCol);
	Variant range_new = worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);

    // ��������! ��������� ������ ����������� ���-�� ����� = Row
    //if (firstRow != rangeFirstRow)
    //if (Row != 0)
	range_new.OleProcedure("Insert", -4121);  // xlDown

    sell_left_top = worksheet.OlePropertyGet("Cells", firstRow, firstCol);
	sell_right_bottom = worksheet.OlePropertyGet("Cells", lastRow, lastCol);
	range_new = worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);

	range.OleProcedure(L"Copy", range_new);

    if (!fCopyData)
        range_new.OleProcedure("ClearContents");

    return range_new;
}

//---------------------------------------------------------------------------
//
Variant __fastcall MSExcelWorks::CopyRangeTo(Variant worksheet, const Variant range, int Row, int Column, bool fCopyData)
{
    //Variant worksheet = range.OlePropertyGet("Worksheet");

    int rowsCount = GetRangeRowsCount(range);
    int colsCount = GetRangeColumnsCount(range);

    //int rangeFirstRow = GetRangeFirstRow(range);
    //int rangeFirstColumn = GetRangeFirstColumn(range);
    int firstRow = Row;
    int firstCol = Column;
    //int firstRow = rangeFirstRow + rowsCount + Row;
    //int firstCol = GetRangeFirstColumn(range) + Column;

    int lastRow = firstRow + rowsCount-1;
    int lastCol = firstCol + colsCount-1;


    Variant sell_left_top = worksheet.OlePropertyGet("Cells", firstRow, firstCol);
	Variant sell_right_bottom = worksheet.OlePropertyGet("Cells", lastRow, lastCol);
	Variant range_new = worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);

    // ��������! ��������� ������ ����������� ���-�� ����� = Row
    //if (firstRow != rangeFirstRow)
    //if (Row != 0)
        range_new.OleProcedure("Insert", -4121);  // xlDown

    sell_left_top = worksheet.OlePropertyGet("Cells", firstRow, firstCol);
	sell_right_bottom = worksheet.OlePropertyGet("Cells", lastRow, lastCol);
	range_new = worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);

    range.OleProcedure("Copy", range_new);

    if (!fCopyData)
        range_new.OleProcedure("ClearContents");

    return range_new;
}

//---------------------------------------------------------------------------
//
Variant __fastcall MSExcelWorks::CopyRange(Variant worksheet, String sRangeName, int Row, int Column, bool fCopyData)
{
    Variant range = GetRangeByName(worksheet, sRangeName);
    return CopyRange(worksheet, range, Row, Column, fCopyData);
}

/*
//---------------------------------------------------------------------------
// ������ ������ �� ������ ������� Excell
String __fastcall ReadCell(Variant wst, int Row, int Col)
{
  return(Trim(wst.OlePropertyGet("Cells").OlePropertyGet("Item", Row, Col)));
}

//---------------------------------------------------------------------------
// �������� ��������
Variant __fastcall actPage(Variant appl)
{
  return(appl.OlePropertyGet("ActiveSheet"));
}

*/

//---------------------------------------------------------------------------
#endif
