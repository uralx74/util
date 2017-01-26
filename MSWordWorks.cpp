#include "MSWordWorks.h"
#include "JsonObject.h"
#include <cassert>

//---------------------------------------------------------------------------
//
MergeTable::MergeTable()
{
    FieldsCount = 0;
    CurrentRecordIndex = 1;
    RecCount = 0;
    PagePerDocument = 500;
}

//---------------------------------------------------------------------------
// �������� (�� ����� ���� ���� ������ �������� RecCount)
void __fastcall MergeTable::ShrinkRecords(int RecCount)
{
    if (RecCount <0)
    {
        this->RecCount = CurrentRecordIndex - 1;
    }
    else
    {
        this->RecCount = RecCount;
    }
}

//---------------------------------------------------------------------------
// ���������� (��������) ������� ������
void __fastcall MergeTable::PrepareFields(int ColCount)
{
    VariantClear(head);
    FieldsCount = ColCount;
    head = CreateVariantArray(1, ColCount);
}

//---------------------------------------------------------------------------
//
void __fastcall MergeTable::PrepareRecords(int RowCount)
{
    VariantClear(data);
    RecCount = RowCount;
    data = CreateVariantArray(RowCount, FieldsCount);
}

//---------------------------------------------------------------------------
//
void __fastcall MergeTable::AddField(int FieldIndex, const AnsiString &FieldName)
{
     head.PutElement(FieldName, 1, FieldIndex);
}

//---------------------------------------------------------------------------
//
void __fastcall MergeTable::PutRecord(const AnsiString &Value, int RecordIndex, int FieldIndex)
{
    data.PutElement(Value, RecordIndex, FieldIndex);

    if (RecCount < CurrentRecordIndex)
        RecCount = CurrentRecordIndex;  // �������� ����� ������� ����� ��� ������������� ������� �������� ������

}

//---------------------------------------------------------------------------
// ������� �������� � ������� ������
void __fastcall MergeTable::PutRecord(int FieldIndex, const AnsiString &Value)
{
    data.PutElement(Value, CurrentRecordIndex, FieldIndex);

}

//---------------------------------------------------------------------------
// ������� �������
void __fastcall MergeTable::Free()
{
    VariantClear(data);
    VariantClear(head);

}

//---------------------------------------------------------------------------
//
void __fastcall MergeTable::Next()
{
    CurrentRecordIndex++;
    // ��� �������������� ����������� ������ �� ������� �������
    //if (CurrentRecordIndex > VarArrayHighBound(data,1) {
    //      RedimVariantArray(data, RecCount, fields.size()+100);
    //}
}

//---------------------------------------------------------------------------
//
void __fastcall MergeTable::First()
{
    CurrentRecordIndex = 1;
}

//---------------------------------------------------------------------------
//
void __fastcall MSWordWorks::SetDisplayAlerts(bool flg)
{
    WordApp.OlePropertySet("DisplayAlerts", flg);
}


//---------------------------------------------------------------------------
// ������ �������� Word
Variant __fastcall MSWordWorks::OpenWord()
{
    Variant Document;
	try
    {
        // �������� ������� Word.Application
        WordApp = CreateOleObject("Word.Application.8");

        // ���� handle ����
        randomize();
        String OldTitle = WordApp.OlePropertyGet("Caption");
        String TempTitle = "Temp - " + IntToStr(random(1000000));
        WordApp.OlePropertySet("Caption", TempTitle.c_str());
        //Handle = FindWindow(NULL, TempTitle.c_str());
        Handle = FindWindow("OpusApp", TempTitle.c_str());
        WordApp.OlePropertySet("Caption", OldTitle.c_str());

        // ��������� ����� ������ ��������������.
        WordApp.OlePropertySet("DisplayAlerts", false);

        // ��������� �������� ���������� ��� ��������� ������
		WordApp.OlePropertyGet("Options").OlePropertySet("CheckSpellingAsYouType", false);
        WordApp.OlePropertyGet("Options").OlePropertySet("CheckGrammarAsYouType", false);
        WordApp.OlePropertyGet("Options").OlePropertySet("CheckGrammarWithSpelling", false);

        // ����������� Word.Application
        WordApp.OlePropertySet("Visible", false);

        // ������������� ������ �� ���������
	  	Documents = WordApp.OlePropertyGet("Documents");
    }
    catch (Exception &e)
    {
       try
       {
           CloseApplication();
        }
        catch(...)
        {
        }
        throw Exception(e);
    }
}

/*//---------------------------------------------------------------------------
// ������ �������� Word - ��������� ��� ������������� �� ������� ��������
Variant __fastcall MSWordWorks::OpenWord(const String &DocumentFileName, bool fAsTemplate)
{
    Variant Document;
	try
    {
        // �������� ������� Word.Application
        WordApp = CreateOleObject("Word.Application.8");

        randomize();

        // ���� handle ����
        WideString OldTitle = WordApp.OlePropertyGet("Caption");
        WideString TempTitle = "Temp - " + IntToStr(random(1000000));
        WordApp.OlePropertySet("Caption", TempTitle.c_str());
        //Handle = FindWindow(NULL, TempTitle.c_str());
        Handle = FindWindow("OpusApp", TempTitle.c_str());
        WordApp.OlePropertySet("Caption", OldTitle);


        // ��������� ����� ������ ��������������.
        WordApp.OlePropertySet("DisplayAlerts", false);

        // ��������� �������� ���������� ��� ��������� ������
		WordApp.OlePropertyGet("Options").OlePropertySet("CheckSpellingAsYouType", false);
        WordApp.OlePropertyGet("Options").OlePropertySet("CheckGrammarAsYouType", false);
        WordApp.OlePropertyGet("Options").OlePropertySet("CheckGrammarWithSpelling", false);

        // ����������� Word.Application
        WordApp.OlePropertySet("Visible", false);

        // ������������� ������ �� ���������
	  	Documents = WordApp.OlePropertyGet("Documents");

        // ����� ��������� �� �����
        if (fAsTemplate)
        {
        	// ������ ����� ����������� �������� �������� ������ Item = 1
        	Documents.OleProcedure("Add", DocumentFileName, false,0);
        	Document = Documents.OleFunction("Item",1); // ������ � ���������
        }
        else
        {
        	Document = Documents.OleFunction("Open", DocumentFileName);
        }
        return Document;


    }
    catch (Exception &e)
    {
        try
        {
            if (!VarIsEmpty(Document))     // ���� ������ ������
            {
                CloseDocument(Document);
            }
        }
        catch(...)
        {}
        try
        {
           CloseApplication();
        }
        catch(...)
        {
        }
        throw Exception(e);
    }
} */

//---------------------------------------------------------------------------
// ������ �������� Word �� �����
Variant __fastcall MSWordWorks::OpenDocument(const String &DocumentFileName, bool fAsTemplate)
{
    // ����� ��������� �� �����
    // � Ole-��������� Open ���������� ��������� �������������� ����������
    Variant document;
    if ( fAsTemplate )
    {
        // ����������� ������������ ��������
    	document = Documents.OleFunction("Open", DocumentFileName);
    }
    else
    {
        // ��������� ����� ��������, ���� ����������� �������� - ������
    	// ������ ����� ����������� �������� �������� ������ Item = 1
    	document = Documents.OleFunction("Add", DocumentFileName.c_str(), false, 0);
    }
    return document;
}

//---------------------------------------------------------------------------
// ��������� ��������� �� �������
Variant __fastcall MSWordWorks::GetDocument(int DocIndex)
{
	try
    {
        if (DocIndex >= 0)
        {
            //int k = WordApp.OlePropertyGet("Documents").OlePropertyGet("Count");
	        return WordApp.OlePropertyGet("Documents").OleFunction("Item", DocIndex);
        }
        else
        {
	        return WordApp.OlePropertyGet("ActiveDocument");
        }
    }
    catch (...)
    {
    	return NULL;
    }
}

//---------------------------------------------------------------------------
//
Variant __fastcall MSWordWorks::GetPage(Variant Document, int PageIndex)
{
   	return Document.OlePropertyGet("Item", wdPropertyPages).OlePropertyGet("Value");
   	//return Document.OlePropertyGet("BuiltInDocumentProperties").OlePropertyGet("Item", wdPropertyPages).OlePropertyGet("Value");
}

//---------------------------------------------------------------------------
//
int MSWordWorks::GetCurrentPageNumber(Variant Document)
{
    int page = Document.OlePropertyGet("Selection").OlePropertyGet("Information", wdActiveEndPageNumber);
    return page;
}

//---------------------------------------------------------------------------
// ������ ��� �������� ��������
void MSWordWorks::SetVisible(bool fVisible)
{
	// ����������� Word.Application
	WordApp.OlePropertySet("Visible", fVisible);
}

//---------------------------------------------------------------------------
// ������� ������ � �������� (�������� ���������� ������������)
void __fastcall MSWordWorks::SetTextToBookmark(Variant Document, String BookmarkName, WideString Text)
{
    Variant Bookmark = Document.OlePropertyGet("Bookmarks").OleFunction("Item", (OleVariant)BookmarkName);
	Bookmark.OlePropertyGet("Range").OlePropertySet("Text", Text);
}

//---------------------------------------------------------------------------
// ������� ����������� � �������� (����� ���������� ������������)
bool MSWordWorks::SetPictureToBookmark(Variant Document, String BookmarkName, String PictureFileName, int Width, int Height)
{
    Variant Bookmarks=Document.OlePropertyGet("Bookmarks");
    Variant Bookmark = Bookmarks.OleFunction("Item", (OleVariant)BookmarkName);
    Bookmark.OlePropertyGet("Range").OlePropertyGet("InlineShapes").OleProcedure("AddPicture", PictureFileName, false, true);

    // �� ���������!
    Bookmark.OleProcedure("Delete");


    //vBookmark.OlePropertyGet("Range").OlePropertySet("Text","12");
    /*// ����� Bookmark �� �����
    Variant Bookmark = Bookmarks.OleFunction("Item", (OleVariant)BookmarkName);
    // �������� ����������� �� �����
    Bookmark.OlePropertyGet("Range").OlePropertyGet("InlineShapes").OleProcedure("AddPicture", PictureFileName, false, true);
	//������ �������
	// ����������� ���������� ������� � ���������
    //Variant Selection = WordApp.OlePropertyGet("Selection");
    //Variant Selection = Document.OlePropertyGet("Selection");

	// �������� ����� �������!!!!
    /*Variant InlineShape, PictureShape;
    InlineShape =  Document.OlePropertyGet("InlineShapes").OleFunction("Item", 1);
    //InlineShape = Field.OlePropertyGet("Range").OlePropertyGet("InlineShapes").OleFunction("Item", i);
    */

}

//---------------------------------------------------------------------------
// ���������� ������ � ������� ����� FormFields
std::vector<String> __fastcall MSWordWorks::GetFormFields(Variant Document)
{
    // ������� �������, �� ��� �������� �� ������������� ���� � ������ FieldName
    //Variant Field = Document.OlePropertyGet("FormFields").OleFunction("Item", (OleVariant)FieldName);
    //Field.OlePropertyGet("Range").OlePropertySet("Text", Text);

    Variant FormFields = Document.OlePropertyGet("FormFields");
    Variant Field = Unassigned;
    int n = FormFields.OlePropertyGet("Count");

    std::vector<String> vFormFields;
    vFormFields.reserve(n);


    for (int i = 1; i <= n; i++)   // ��������� ������ � ������� �����
    {
        String Name = UpperCase(FormFields.OleFunction("Item", i).OlePropertyGet("Result"));
        vFormFields.push_back(Name);
        /*if (FieldName == Name) {
            Field = Document.OlePropertyGet("FormFields").OleFunction("Item", (OleVariant)FieldName);
        }*/
    }

    return vFormFields;
}

//---------------------------------------------------------------------------
// ������� ������ � ���� (���� ���������� �������) - ������� �������, ��� �������� ������������� ����
void __fastcall MSWordWorks::SetTextToFieldF(Variant Document, String FieldName, WideString Text)
{
    // ������� �������, �� ��� �������� �� ������������� ���� � ������ FieldName
    Variant Field = Document.OlePropertyGet("FormFields").OleFunction("Item", (OleVariant)FieldName);
    Field.OlePropertyGet("Range").OlePropertySet("Text", Text);
}

//---------------------------------------------------------------------------
// ������� ������ � ���� (���� ���������� �������) - ������� �������, ��� �������� ������������� ����
void __fastcall MSWordWorks::SetTextToFieldF(Variant Document, int fieldIndex, WideString Text)
{
    // ������� �������, �� ��� �������� �� ������������� ���� � ������ FieldName
    Variant Field = Document.OlePropertyGet("FormFields").OleFunction("Item", fieldIndex);
    Field.OlePropertyGet("Range").OlePropertySet("Text", Text);
}


//---------------------------------------------------------------------------
// ������� ������ � ���� (���� ���������� �������)
void __fastcall MSWordWorks::SetTextToField(Variant Document, String FieldName, WideString Text)
{
    Variant FormFields = Document.OlePropertyGet("FormFields");
    Variant Field = Unassigned;
    FieldName = UpperCase(FieldName);
    int n = FormFields.OlePropertyGet("Count");
    for (int i = 1; i <= n; i++)   // ���� �� ������� ����. ��������� ������� ���� FieldName
    {
        String Name = UpperCase(FormFields.OleFunction("Item", i).OlePropertyGet("Result"));
        if (FieldName == Name)
        {
            Field = Document.OlePropertyGet("FormFields").OleFunction("Item", (OleVariant)FieldName);
        }
    }

    if (!Field.IsEmpty())
    {
        Field.OlePropertyGet("Range").OlePropertySet("Text", Text);
    }



/*  Variant Fields = Document.OlePropertyGet("MailMerge").OlePropertyGet("Fields");
    int FieldsCount = Fields.OlePropertyGet("Count");
    Variant Field = Fields.OleFunction("Item", 1);
    Variant Code = Field.OlePropertyGet("Code");
    String Text = code.OlePropertyGet("Text");
    String Type = Field.OlePropertyGet("Type");
*/

}

/* ����������� InlineShape � ��������� Shape
   zOrder - ��������� ������������ ������
   4 - ����� �������
   5 - �� �������
*/
Variant __fastcall MSWordWorks::ConverInlineShapeToShape(Variant inlineShape, int zOrder)
{
    // ������������ � Shape
    Variant shape = inlineShape.OleFunction("ConvertToShape");

    // ������������ ����������� ����� �������
    shape.OleFunction("ZOrder", zOrder);
}

/* ������ ������� ���������� shape */
void __fastcall MSWordWorks::SetShapeSize(Variant shape, int width, int height)
{
    // ������������� ������ �����������
    //if (Width != 0 && Height != 0) {
    	shape.OlePropertySet("Width", width);
    	shape.OlePropertySet("Height", height);
    //}
}

/* ������ ��������� ���������� Shape
   �� ���������!
*/
void __fastcall MSWordWorks::SetShapePos(Variant shape, int x, int y)
{
    // ������������� ������������ �� �����
    shape.OleFunction("IncrementLeft", x);
    shape.OleFunction("IncrementTop", y);
}


Variant __fastcall MSWordWorks::SetPictureToField(Variant Document, Variant Field, String PictureFileName, int Width, int Height)
{
    try
    {
        Variant InlineShapes = Field.OlePropertyGet("Range").OlePropertyGet("InlineShapes");
        Variant InlineShape = InlineShapes.OleFunction("AddPicture", PictureFileName.c_str(), false, true);

        // �� ���������!
        Field.OleProcedure("Delete");

        if (Width != 0 || Height != 0)
        {
            InlineShape.OlePropertySet("Width", Width);
            InlineShape.OlePropertySet("Height", Height);
        }

        return InlineShape;
    }
    catch (Exception &e)
    {
        throw(Exception("Exception has occurred.\nFile \"" + PictureFileName + "\" not found."));
    }
}


//---------------------------------------------------------------------------
// ������� ������� � ���� (���� ��������)
Variant __fastcall MSWordWorks::SetPictureToField(Variant Document, int fieldIndex, String PictureFileName, int Width, int Height)
{
    Variant Field = Document.OlePropertyGet("FormFields").OleFunction("Item", fieldIndex);
    SetPictureToField(Document, Field, PictureFileName, Width, Height);
}

//---------------------------------------------------------------------------
// ������� ������� � ���� (���� ��������)
Variant __fastcall MSWordWorks::SetPictureToField(Variant Document, String FieldName, String PictureFileName, int Width, int Height)
{
    Variant Field = Document.OlePropertyGet("FormFields").OleFunction("Item", (OleVariant)FieldName);
    SetPictureToField(Document, Field, PictureFileName, Width, Height);
}

//---------------------------------------------------------------------------
// ����� ������ � �������
void __fastcall MSWordWorks::FindTextForReplace(Variant document, String Text, String ReplaceText, bool fReg)
{
        document.OleProcedure("Activate");
        // ����������� ���������� ������� � ���������
        //Variant Selection = Document.OlePropertyGet("Selection");
        Variant Selection = WordApp.OlePropertyGet("Selection");

        // ����� ������ �� ����� ���������
		Variant Find = Selection.OlePropertyGet("Find");
        Find.OleProcedure("Execute", Text.c_str()/*�����, ������� ����� ������*/, fReg/*��������� �������e*/,
        	false/*������ ������ �����*/,false/*��������� ������������� �������*/,false/*������ ������������ ���*/,
        	false/*������ ��� ����������*/,true/*������ ������*/,1/*��������� ������ ����� �����*/,
        	false/* ������� ������� */, ReplaceText.c_str()/*�� ��� ��������*/,2/*�������� ���*/);   // ���� ������� ��������

        /*
        Find.OleProcedure("ClearFormatting");                                         // ���� ������� �� ��������, ���� �����������
        Find.OlePropertyGet("Replacement").OleProcedure("ClearFormatting");
        Find.OlePropertySet("Text",Text);
        Find.OlePropertyGet("Replacement").OlePropertySet(Text,ReplaceText);
        Find.OlePropertySet("Forward",True);
        Find.OlePropertySet("Wrap",1);
        Find.OlePropertySet("Format",False);
        Find.OlePropertySet("MatchCase",False);
        Find.OlePropertySet("MatchWholeWord",False);
        Find.OlePropertySet("MatchWildcards",False);
        Find.OlePropertySet("MatchSoundsLike",False);
        Find.OlePropertySet("MatchAllWordForms",False);
        Find.OleProcedure("Execute",2);   /**/
}

//---------------------------------------------------------------------------
// (����������!) ����������� �������� �� ������   (����������!)(����������!)(����������!)(����������!)(����������!)
Variant MSWordWorks::CopyPage(int PageNumber)
{
	//Variant Selection = Document.OlePropertyGet("Selection");
	// ����������� ������
    WordApp.OlePropertyGet("Selection").OleFunction("WholeStory");
    //CurrentSelection.OleProcedure("MoveUp", wdLine, 1, wdExtend);
    //Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
    return WordApp.OlePropertyGet("Selection").OleFunction("Copy");
}

//---------------------------------------------------------------------------
//  ������� �������� �� ������(����������!)(����������!)(����������!)(����������!)(����������!)(����������!)(����������!)
void __fastcall MSWordWorks::PastePage(Variant Document, int PageNumber)
{
    Variant Selection = WordApp.OlePropertyGet("Selection");
    //Variant Selection = Document.OlePropertyGet("Selection");
	// ����������� ������ � ����� ���������
    Selection.OleProcedure("EndKey", wdStory);
	//CurrentSelection.OleProcedure("InsertNewPage");
    Selection.OleFunction("Paste");
    //Selection.OleFunction("PasteAndFormat", 0);
    //Selection.OleFunction("PasteAndFormat", Page);
}

//---------------------------------------------------------------------------
//
void __fastcall MSWordWorks::InsertFile(Variant Document, AnsiString FileName)
{
    // ����������� ������ � ����� ���������
    Variant Selection = Document.OlePropertyGet("Selection");
    Selection.OleProcedure("EndKey", wdStory);
    Selection.OleProcedure("InsertFile", FileName.c_str(), "", false, false);
}

//---------------------------------------------------------------------------
// ����������� ������ � ������ ��������� (����������!)
void __fastcall  MSWordWorks::MoveUpCursor(Variant Document)
{
	//Variant Selection = Document.OlePropertyGet("Selection");
	// ����������� ������
    Variant Selection = Document.OlePropertyGet("Selection");
	Selection.OleProcedure("MoveUp", 7, 1);
}

//---------------------------------------------------------------------------
// ��������� � ����
void __fastcall MSWordWorks::SaveAsDocument(Variant Document, String FileName/*, bool fAddToRecentFiles*/)
{
    // ���������� ��������� � ����
    // � ������ Ole-��������� ��������� �������������� ����������
    // ����� ���������� ������ ���� ��������� - SaveAs2
    Document.OleProcedure("SaveAs", FileName);
}

//---------------------------------------------------------------------------
//
void __fastcall MSWordWorks::SetActiveDocument(Variant Document)
{
	// ��������� ���������
    Document.OleProcedure("Activate");
}

//---------------------------------------------------------------------------
// ������� �������
Variant MSWordWorks::CreateTable(Variant Document, int nCols, int nRows)
{
    Variant range = Document.OleFunction("Range");
	// ������� ������� � ������� Range
    Document.OlePropertyGet("Tables").OleProcedure("Add", range, (OleVariant) nCols, (OleVariant) nRows);

	// ����� ������������ �������
	//Table = Tables.OleFunction("Item", 1);
	//RowCount = Table.OlePropertyGet("Rows").OlePropertyGet("Count");
	//ColCount = Table.OlePropertyGet("Columns").OlePropertyGet("Count");
}

//---------------------------------------------------------------------------
// ������� � ��������
void __fastcall MSWordWorks::GoToBookmark(Variant Document, String BookmarkName)
{
    Document.OleFunction("Range").OleProcedure("GoTo", wdGoToBookmark, 0, 0, WideString(BookmarkName));
//    Document.OlePropertyGet("Selection").OleProcedure("GoTo",(int)-1, 0, 0, WideString(BookmarkName));
}

//---------------------------------------------------------------------------
// ������� � ������
void __fastcall MSWordWorks::GoToText(Variant Document, String Text, bool fReg, bool fWord)
{
    // ����������� ���������� ������� � ���������
    //Variant Selection = WordApp.OlePropertyGet("Selection");
    //CurrentSelection = WordApp.OlePropertyGet("Selection");

    Variant Selection = Document.OleFunction("Range");

    // ����� ������ �� ����� ���������
	Variant Find = Selection.OlePropertyGet("Find");
    Find.OleProcedure("Execute", Text/*�����, ������� ����� ������*/, fReg/*��������� �������e*/,
        fWord/*������ ������ �����*/,false/*��������� ������������� �������*/,false/*������ ������������ ���*/,
        false/*������ ��� ����������*/,true/*������ ������*/,1/*��������� ������ ����� �����*/,
        false/* ������� ������� */, 0/*�� ��� ��������*/,0/*�������� ���*/);   // ���� ������� ��������

    //�urrentSelection.OleProcedure("GoTo",(int)-1, 0, 0, WideString(BookmarkName));
}

//---------------------------------------------------------------------------
// �������� �����������
void __fastcall MSWordWorks::InsertPicture(Variant Document, String PictureFileName, int Width, int Height)
{
     // �������� ����������� �� ����� � ������� CurrentSelection
     Variant Selection = Document.OleFunction("Range");
     Selection.OlePropertyGet("InlineShapes").OleProcedure("AddPicture", "C:\\_project\\InsertPicToWord\\tmp\\podpis.bmp", false, true);
}

//---------------------------------------------------------------------------
// �������� �����
void __fastcall MSWordWorks::InsertText(Variant Document, WideString Text)
{
     // �������� ����� � ������� CurrentSelection
     Variant Selection = Document.OleFunction("Range");
     Selection.OleProcedure("TypeText", Text);
}

//---------------------------------------------------------------------------
// �������� �� Clipboard (���������� ������� �� ���������� Clipboard)
bool MSWordWorks::PasteFromClipboard()
{
     // �������� �� �������
	WordApp.OlePropertyGet("Selection").OleFunction("Paste");
    return true;
}

//---------------------------------------------------------------------------
//  �������� ���������� Word � ��������� ���� �������� ����������
void __fastcall MSWordWorks::CloseApplication()
{
	// �������� ���������� Word (� ������� �� ���������� ���������)
    if (!WordApp.IsEmpty())
    {
        Variant document;
        while (Documents.OlePropertyGet("Count") > 0)
        {
            document = WordApp.OlePropertyGet("ActiveDocument");
            document.OleFunction("Close", false);
        }
	    WordApp.OleProcedure("Quit");
    }
}

//---------------------------------------------------------------------------
// �������� ���������
void __fastcall MSWordWorks::CloseDocument(Variant Document, bool fCloseAppIfNoDoc)
{
	Document.OleFunction("Close", false);

    if (fCloseAppIfNoDoc && Documents.OlePropertyGet("Count") == 0)
    {
        CloseApplication();
    }
}

//---------------------------------------------------------------------------
// ����������� ���������� �������
int MSWordWorks::GetPagesCount(Variant Document)
{
	return Document.OlePropertyGet("BuiltInDocumentProperties").OlePropertyGet("Item", wdPropertyPages).OlePropertyGet("Value");
}

//---------------------------------------------------------------------------
// ������� � �������� MERGETABLE � ����
std::vector<String> __fastcall MSWordWorks::MergeDocumentToFiles(Variant TemplateDocument, MERGETABLE &md)
{
    //md.TemplateDocument = OpenWord(md.TemplateFileName, true);

    int nFiles;
    if (md.PagePerDocument <= 0)          // ����������� ���-�� �������������� ������
    {
        md.PagePerDocument = md.RecCount;
        nFiles = 1;
    }
    else
    {
        nFiles = ceil((double)md.RecCount/(double)md.PagePerDocument);
    }

    int nPad = IntToStr(nFiles).Length();  // ���-�� ������ � ������� � ����� ������

    std::vector<String> vFiles;
    vFiles.reserve(nFiles);

    int FileIndex = 0;
    for (int i = 1; i <= md.RecCount; i = i + md.PagePerDocument)
    {
        FileIndex++;
        //AnsiString filename = md.ResultFileNamePrefix + str_pad(IntToStr(FileIndex), nPad, "0", STR_PAD_LEFT) + ".doc";
        AnsiString counterStr = StrPadL(IntToStr(FileIndex), nPad, "0");
        TReplaceFlags replaceflags = TReplaceFlags() << rfReplaceAll << rfIgnoreCase;
        AnsiString filename = StringReplace(md.resultFilename, "[:counter]", counterStr, replaceflags);

        Variant ResultDocument = MergeDocument(TemplateDocument, md, i);

        // ��������� ��������, ��� ��������� ��� � ������ ��������� ������ (AddToRecentFiles = false - 6� ��������)
        try
        {
            ResultDocument.OleProcedure("SaveAs", filename.c_str(), 0, false, "", false);
        }
        catch (Exception &e)
        {
            throw Exception("������ ��� ���������� � ����\n" + filename);
        }
        CloseDocument(ResultDocument);

        vFiles.push_back(filename);
    }
    //CloseDocument(TemplateDocument);

    return vFiles;
    //return FileIndex;
}

//---------------------------------------------------------------------------
// ������� � �������� MERGETABLE � ������ Word - Document
Variant __fastcall MSWordWorks::MergeDocument(Variant TemplateDocument, MERGETABLE &md, int StartIndex)
{
    randomize();
    AnsiString TmpFileName = "ds" + IntToStr(random(100000000)) + ".html";

    TmpFileName = GetTempPath() + TmpFileName;

    int ArrayRowsCount = md.RecCount;

    //int ArrayRowsCount = VarArrayHighBound(ArrayData.data, 1) - VarArrayLowBound(ArrayData.data, 1)+1;
    //int ArrayColsCount = VarArrayHighBound(md.data, 2) - VarArrayLowBound(md.data, 2)+1;

    //int ArrayRowsCount = ArrayData.data2[0].size();
    //int ArrayColsCount1 = md.FieldsCount;

    //int startCol = VarArrayLowBound(ArrayData.data, 2);
    int startCol = 1;
    int LastRecordIndex;

    if (StartIndex <= 0 )
    {
        StartIndex = 1;
    }
    int PagesCount = StartIndex + md.PagePerDocument-1;
    if (PagesCount > md.RecCount)
    {
        PagesCount = md.RecCount;
    }

     ofstream out(TmpFileName.c_str());

    // ��������� HTML
    out<<"<html>\n";
    out<<"<head><META http-equiv=""content-type"" content=""text/html; charset=windows-1251""></head>\n";
    out<<"<body>\n<table>";

    // ���������
    out<<"<tr>";
    for (int j = startCol; j <= md.FieldsCount; j++)
    {
        AnsiString s ="<td>"  + md.head.GetElement(1, j) + "</td>";
        out<< s.c_str();
    }
    out<<"</tr>";
    out<<"\n";

    // ����
    for (int i = StartIndex; i <= PagesCount; i++)
    {
        out<<"<tr>";
        for (int j = startCol; j <= md.FieldsCount; j++)
        {
            //AnsiString s ="<td>#@#sep_"  + md.data.GetElement(i, j) + "</td>";  // ���������������� 2016-07-21
            AnsiString s ="<td>"  + md.data.GetElement(i, j) + "</td>";
            out<< s.c_str();
        }
        out<<"</tr>";
        out<<"\n";
    }
    out<<"</table></body></html>";

    out.close();

    Variant Document = MergeDocumentFromFile(TemplateDocument, TmpFileName, 1, PagesCount);

    //FindTextForReplace("#@#sep_", ""); // ���������������� 2016-07-21
    // �������� ������ ������� �� ������ ��������
    // ��� ����������, ����� ������������ ����� �������� �������� �������.
    // ����� � ��������� ���������� ��������� ������ � ���� �������� 
    FindTextForReplace(Document, "^b", "^m");

    for (int i = 0; i < 5; i ++)       // Delete temporary file
    {
        if (remove(TmpFileName.c_str()) == 0)
        {
            break;
        }
        Sleep(500);
    }

    return Document;
}

/*
//---------------------------------------------------------------------------
// ������� � �������� MERGETABLE � ������ Word - Document
Variant __fastcall MSWordWorks::MergeDocument(Variant TemplateDocument, MERGETABLE &md, int FirstRecordIndex)
{
    randomize();
    AnsiString TmpFileName = "ds" + IntToStr(random(100000000)) + ".html";
    TmpFileName = GetTempPath() + TmpFileName;

    int ArrayRowsCount = md.RecCount;

    //int ArrayRowsCount = VarArrayHighBound(ArrayData.data, 1) - VarArrayLowBound(ArrayData.data, 1)+1;
    int ArrayColsCount = VarArrayHighBound(md.data, 2) - VarArrayLowBound(md.data, 2)+1;

    //int ArrayRowsCount = ArrayData.data2[0].size();
    int ArrayColsCount1 = md.fields.size();

    //int ArrayRowsCount = VarArrayHighBound(ArrayData.data, 1) - VarArrayLowBound(ArrayData.data, 1)+1;
    //int ArrayColsCount = VarArrayHighBound(ArrayData.data, 2) - VarArrayLowBound(ArrayData.data, 2)+1;

    //int startCol = VarArrayLowBound(ArrayData.data, 2);
    int startCol = 1;
    int LastRecordIndex;

    if (FirstRecordIndex <= 0 )
        FirstRecordIndex = 1;

    //int delta = ArrayRowsCount - FirstRecordIndex + 1;
    int PagesCount;

    if (md.PagePerDocument <= 0)
    {
        //PagePerDocument = ArrayRowsCount - FirstRecordIndex + 1;
        PagesCount = md.RecCount;
    } else if (md.PagePerDocument >= delta) {
        //PagePerDocument = delta;
        PagesCount = md.PagePerDocument;
    }
    LastRecordIndex = FirstRecordIndex + PagePerDocument - 1;

    int FieldsCount_Document = TemplateDocument.OlePropertyGet("Fields").OlePropertyGet("Count");

    ofstream out(TmpFileName.c_str());

    out<<"<html>\n";
    out<<"<head><META http-equiv=""content-type"" content=""text/html; charset=windows-1251""></head>\n";
    out<<"<body>\n<table>";


    // ���������
    out<<"<tr>";
    for (std::map<AnsiString, int>::iterator field = ArrayData.fields.begin(); field != ArrayData.fields.end(); ++field)
    {
        AnsiString s ="<td>"  + field->first + "</td>";
        out<< s.c_str();
    }
    out<<"</tr>";
    out<<"\n";

    // ����
    for (int i = FirstRecordIndex; i <= LastRecordIndex; i++)
    {
        out<<"<tr>";
        for (int j = startCol; j <= ArrayColsCount; j++)
        {

            //AnsiString s ="<td>"  + ArrayData.data2[i][j] + "</td>";
            AnsiString s ="<td>"  + ArrayData.data.GetElement(i, j) + "</td>";
            out<< s.c_str();
        }
        out<<"</tr>";
        out<<"\n";
    }
    out<<"</table></body></html>";

    out.close();


    Variant Document = MergeDocumentFromFile(TemplateDocument, TmpFileName, 1, PagePerDocument);

    for (int i = 0; i < 5; i ++) {      // Delete temporary file
        if (remove(TmpFileName.c_str()) == 0)
            break;
        Sleep(500);
    }

    return Document;
}  */

//---------------------------------------------------------------------------
//
Variant __fastcall MSWordWorks::MergeDocument(Variant TemplateDocument, const Variant &ArrayData, int FirstRecordIndex, int PagePerDocument, int titleRowIndex)
{
    randomize();
    AnsiString TmpFileName = "ds" + IntToStr(random(100000000)) + ".html";
    TmpFileName = GetTempPath() + TmpFileName;


    int ArrayRowsCount = VarArrayHighBound(ArrayData, 1) - VarArrayLowBound(ArrayData, 1)+1;
    int ArrayColsCount = VarArrayHighBound(ArrayData, 2) - VarArrayLowBound(ArrayData, 2)+1;

    //int startRow = VarArrayLowBound(*ArrayData, 1);
    int startCol = VarArrayLowBound(ArrayData, 2);
    int LastRecordIndex;


    if (titleRowIndex <= 0)
    {
        titleRowIndex = VarArrayLowBound(ArrayData, 1);
    }

    if (FirstRecordIndex <= 0 )
    {
        FirstRecordIndex = titleRowIndex + 1;
    }

    int delta = ArrayRowsCount - FirstRecordIndex + 1;
    if (PagePerDocument <= 0)
    {
        PagePerDocument = ArrayRowsCount - FirstRecordIndex + 1;
    } else if (PagePerDocument >= delta) {
        PagePerDocument = delta;
    }

    LastRecordIndex = FirstRecordIndex + PagePerDocument - 1;

/*    if (RecordCount <= 0)
    {
        LastRecordIndex = VarArrayHighBound(ArrayData, 1);
    }
    else
    {
        LastRecordIndex = FirstRecordIndex + RecordCount;
        if (LastRecordIndex > ArrayRowsCount) {
            LastRecordIndex = ArrayRowsCount;
            RecordCount =
        }
    }     */

    //int FieldsCount_Document = TemplateDocument.OlePropertyGet("Fields").OlePropertyGet("Count");

    //int FieldsCount = FieldsCount_Document < ArrayColsCount? FieldsCount_Document : ArrayColsCount;

    ofstream out(TmpFileName.c_str());

    out<<"<html>\n";
    out<<"<head><META http-equiv=""content-type"" content=""text/html; charset=windows-1251""></head>\n";
    out<<"<body>\n<table>";

    // ���������
    out<<"<tr>";
    for (int j = startCol; j <= ArrayColsCount; j++)
    {
        AnsiString s ="<td>"  + ArrayData.GetElement(titleRowIndex, j) + "</td>";
        out<< s.c_str();
    }
    out<<"</tr>";
    out<<"\n";


    for (int i = FirstRecordIndex; i <= LastRecordIndex; i++)
    {
        out<<"<tr>";
        for (int j = startCol; j <= ArrayColsCount; j++)
        {
            // AnsiString s ="<td>#@#sep_"  + ArrayData.GetElement(i, j) + "</td>";   // ����������������� 2016-07-21
            AnsiString s ="<td>#@#sep_"  + ArrayData.GetElement(i, j) + "</td>";
            out<< s.c_str();
        }
        out<<"</tr>";
        out<<"\n";
    }
    out<<"</table></body></html>";

    out.close();

    Variant Document = MergeDocumentFromFile(TemplateDocument, TmpFileName, 1, PagePerDocument);
    //FindTextForReplace("#@#sep_", "");  // ���������������� 2016-07-21

    for (int i = 0; i < 5; i ++)       // Delete temporary file
    {
        if (remove(TmpFileName.c_str()) == 0)
        {
            break;
        }
        Sleep(500);
    }

    return Document;
}

//---------------------------------------------------------------------------
// ������� �� �������� ����� � ������� (html)
Variant __fastcall MSWordWorks::MergeDocumentFromFile(Variant TemplateDocument, AnsiString DatasetFileName, int FirstRecordIndex, int PagePerDocument)
{
    Variant ResultDocument;
    Variant MailMerge;
    Variant PasswordDocument, PasswordTemplate, WritePasswordDocument, WritePasswordTemplate, SQLStatement, SQLStatement1;
    PasswordDocument = "";
    PasswordTemplate = "";
    WritePasswordDocument = "";
    WritePasswordTemplate = "";
    SQLStatement = "";
    //SQLStatement = "SELECT * FROM [����1$]";
    SQLStatement = "SELECT * FROM `Table`";
    SQLStatement1 = "SELECT * FROM `Table`";
    String chemin, texte;
    texte = "test";

    MailMerge = TemplateDocument.OlePropertyGet("MailMerge");
    MailMerge.OlePropertySet("MainDocumentType", 0);    // wdFormLetters = 0
    //MailMerge.OleProcedure("OpenDataSource", chemin.c_str(), 1, true, true, false, false, PasswordDocument, PasswordTemplate, false, WritePasswordDocument, WritePasswordTemplate, texte.c_str(), SQLStatement, SQLStatement1, false);
    MailMerge.OleProcedure("OpenDataSource", DatasetFileName.c_str(), 0, false, false, true, false, PasswordDocument, PasswordTemplate, false, WritePasswordDocument, WritePasswordTemplate, texte.c_str(), SQLStatement, SQLStatement1, false);
    MailMerge.OlePropertySet("Destination", 0);
    MailMerge.OlePropertySet("SuppressBlankLines", 0);

    int LastRecordIndex;
    if (FirstRecordIndex <= 0)
    {
        FirstRecordIndex = 1;           //wdDefaultFirstRecord = 1
    }

    if (PagePerDocument  <= 0)
    {
        LastRecordIndex = 0xFFFFFFF0;   //wdDefaultLastRecord = 0xFFFFFFF0
    }
    else
    {
        LastRecordIndex = FirstRecordIndex + PagePerDocument - 1;
    }

    MailMerge.OlePropertyGet("DataSource").OlePropertySet("FirstRecord", FirstRecordIndex);
    MailMerge.OlePropertyGet("DataSource").OlePropertySet("LastRecord", LastRecordIndex);

    MailMerge.OleProcedure("Execute", false);

    // Free datasource
    MailMerge.OlePropertySet("MainDocumentType", 0xFFFFFFFF);    // wdNotAMergeDocument = 0xFFFFFFFF



    // Return the new document
    WordApp = MailMerge.OlePropertyGet("Application");
    return WordApp.OlePropertyGet("Documents").OleFunction("Item", 1);


    //return GetDocument(1);

    /*MailMerge.ExecFunction("OpenDataSource") <<XLSFileName               // Name
                                    <<0                         // Format
                                    <<false                     // ConfirmConversions
                                    <<false                     // ReadOnly
                                    <<true                      // LinkToSource
                                    <<false                     // AddToRecentFiles
                                    <<EmptStr                   // PasswordDocument
                                    <<EmptStr                   // PasswordTemplate
                                    <<false                     // Revert
                                    <<EmptStr                   // WritePasswordDocument
                                    <<EmptStr                   // WritePasswordTemplate
                                    <<"Entire Spreadsheet"      // Connection
                                    <<EmptStr                   // SQLStatement
                                    <<EmptStr                   // SQLStatement1
                                    <<false                     // OpenExclusive
                                    <<8                         // SubType
         );*/

}

//---------------------------------------------------------------------------
// ������� �� �������� ����� � ������� (html)
std::vector<String> MSWordWorks::ExportToWordFields(TDataSet* QTable, Variant Document, const String& resultPath, int PagePerDocument)
{

    try
    {
        Variant vFields = Document.OlePropertyGet("MailMerge").OlePropertyGet("Fields");
        int FieldCount = vFields.OlePropertyGet("Count");

        //String Type = Field.OlePropertyGet("Type");
        int RecordCount = QTable->RecordCount;

        // �������� �������-������� ��� �������
        MERGETABLE mergetable;
        mergetable.resultFilename = resultPath;
        mergetable.PrepareFields(FieldCount);
        mergetable.PrepareRecords(RecordCount);
        mergetable.PagePerDocument = PagePerDocument;
        QTable->First();


        // ���� �� ��������� MergeField
        // ��������� ����� ������� ������� ����� (������ ���� = ������ ����� � Query)
        for (int i = 1; i <= FieldCount; i++)
        {
            Variant vField = vFields.OleFunction("Item", i);
            Variant vCode = vField.OlePropertyGet("Code");

            // �������� ������� ����������, ��� ��� ���� ������������ �������
            // �������� ��� ���� � �������� �� ���� ��� ����
            // Mergefield ��� �����
            String Text = vCode.OlePropertyGet("Text");
            Text = Text.Trim();
            AnsiString sepMergeField = "MERGEFIELD ";
            int SepPos = Text.Pos(sepMergeField) + sepMergeField.Length()-1;
            String FieldName = Text.SubString(SepPos+1, Text.Length()-SepPos).Trim();
            SepPos = FieldName.Pos(" ");
            if (SepPos == 0)
            {
                SepPos = FieldName.Length();
            }
            FieldName = FieldName.SubString(1,SepPos).Trim();

            // ������� ���� MergeField �� ���������, ���� ������������ ���� ��� � ���������
            if (QTable->FindField(FieldName) == NULL)
            {
                //vCode.OlePropertySet("Text", "��������");     // ��������, �� ������� ����
                //vField.OleProcedure("Delete");                // ��������, �� ������� ����

                // �������� ����������� � ��������� ���� �� ������ ���� NA_���_����
                // �������� �������� �� ����������, ��� ��� ������ � ����
                // �������������� ����� WordApp.OlePropertyGet("Selection")
                // �������� ���������� ���� � ������ ���������� ���������� Word
                // ����� ����������� ����������� ��������.
                vField.OleFunction("Select");
                Variant vSelection = WordApp.OlePropertyGet("Selection");
                Variant vrange = vSelection.OlePropertyGet("Range");
                //vSelection.OleProcedure("TypeText","hello");
                vrange.OlePropertySet("Text", ("NA_" + FieldName).c_str());
                vrange.OlePropertySet("HighlightColorIndex", 7); // wdYellow = 7 ;

                // ���������� ��������� ���
                // ��� MERGEFIELD ����� ��������� ������ ���� ������ MERGEFIELD
                //int delta = FieldCount - vFields.OlePropertyGet("Count");
                FieldCount = vFields.OlePropertyGet("Count");
                i--;
            }
            else
            {
                mergetable.AddField(i, FieldName);
            }  
        }

        // ���������� ������� ������
        for (int i = 1; i <= RecordCount; i++)
        {
            for (int j = 1; j <= FieldCount; j++)
            {
                String FieldName = mergetable.head.GetElement(1, j);

                mergetable.PutRecord(j, QTable->FieldByName(FieldName)->AsString);

                // ������ �������
                // �������� ����������� � ��������� ���� �� ������ ���� NA_���_����
                // �������� �� ����������, ��� ��� ���������� ��� ��������� ��� ������ ������
                // �� ��������� ������. �������� ������� ���� ������� �����.
                //TField* Field = QTable->FindField(FieldName);
                //if (Field) {
                //    mergetable.PutRecord(j, QTable->FieldByName(FieldName)->AsString);
                //} else {    // Field does not exist in souce Query
                //    mergetable.PutRecord(j, "NA_" + FieldName);
                //}

            }

            QTable->Next();
            mergetable.Next();

        }

        MSWordWorks msword;
        return msword.MergeDocumentToFiles(Document, mergetable);
    }
    catch (Exception &e)
    {
        throw Exception(e); // ��������� 2016-03-25. ���������!
    }

    //return std::vector<String> ();
}

/* �������� ���� FormFields, ��������� �������� �� ������� ������ dataSet
*/
void MSWordWorks::ReplaceFormFields(Variant Document, TDataSet* dataSet)
{
    if (!dataSet->Active || dataSet->Eof)
    {
        return;
    }

    Variant FormFields = Document.OlePropertyGet("FormFields");
    Variant Field = Unassigned;
    int n = FormFields.OlePropertyGet("Count");

    for (int i = n; i > 0; i--)
    {
        Variant formFieldsItem = FormFields.OleFunction("Item", i);
        String fieldNameCode = UpperCase(formFieldsItem.OlePropertyGet("Result"));
        String fieldName = "";
        bool isImg = false;


        // ����� �� ��������� �� ���� ����������� ��� json
        //String str = "\"img\":{\"name\"=\"visa\",\"zorder\"=5, \"width\"=\"100\", \"height\" = \"50\"}";

        TJsonObject json(fieldNameCode);
        json.parse();
        TJsonNode* rootNode = json.getRootNode();
        TJsonNode* imgSubNode = NULL;
        if (rootNode != NULL)
        {
            imgSubNode = rootNode->getSubNode("img");
        }


        if (imgSubNode != NULL)
        {
            Variant s = imgSubNode->getParam("zorder", 0);
            isImg = true;
        }
        else
        {
            fieldName = fieldNameCode;
        }

        TField* Field = dataSet->Fields->FindField(fieldName);
        if (Field != NULL) // ���� ����� ����
        {
            if (isImg)
            {
                String imgPath = Field->AsString;
                if ( FileExists(imgPath) )
                {
                    SetPictureToField(Document, i, imgPath);
                }
                else
                {
                    SetTextToFieldF(Document, i, "���� ����������� �� ������! (" + imgPath + ")");
                }
            }
            else
            {
                SetTextToFieldF(Document, i, Field->AsString.c_str());
            }
        }
    }
}

/*
SaveAs2 ������ SaveAs. ����� ���� ���������� ��������� ������� (��. �������� CompatibilityMode)

ActiveDocument.SaveAs2 FileName
        FileFormat:=wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=14
        


wdFormatDocument                    =  0
wdFormatDocument97                  =  0
wdFormatDocumentDefault             = 16
wdFormatDOSText                     =  4
wdFormatDOSTextLineBreaks           =  5
wdFormatEncodedText                 =  7
wdFormatFilteredHTML                = 10
wdFormatFlatXML                     = 19
wdFormatFlatXMLMacroEnabled         = 20
wdFormatFlatXMLTemplate             = 21
wdFormatFlatXMLTemplateMacroEnabled = 22
wdFormatHTML                        =  8
wdFormatPDF                         = 17
wdFormatRTF                         =  6
wdFormatTemplate                    =  1
wdFormatTemplate97                  =  1
wdFormatText                        =  2
wdFormatTextLineBreaks              =  3
wdFormatUnicodeText                 =  7
wdFormatWebArchive                  =  9
wdFormatXML                         = 11
wdFormatXMLDocument                 = 12
wdFormatXMLDocumentMacroEnabled     = 13
wdFormatXMLTemplate                 = 14
wdFormatXMLTemplateMacroEnabled     = 15
wdFormatXPS                         = 18
wdFormatOfficeDocumentTemplate      = 23
wdFormatMediaWiki                   = 24



//Documents.OleFunction("Open", FileName, ConfirmConversions, ReadOnly, AddToRecentFiles,
//    PasswordDocument, PasswordTemplate, Revert,
//    WritePasswordDocument, WritePasswordTemplate, Format   )

*/



