#include "DocumentWriter.h"


/* */
void __fastcall TDocumentWriterResult::clear()
{
    resultFiles.clear();
}

void __fastcall TDocumentWriterResult::addResultFile(String filename)
{
    resultFiles.push_back(filename);
}

void __fastcall TDocumentWriterResult::appendResultFiles(std::vector<String> filenames)
{
    resultFiles.insert(resultFiles.end(), filenames.begin(), filenames.end());
}


/*
 ���������� ������� MS Word
 QueryMerge - �������� ������, ������������ � �������� ��������� ������ ��� �������
 QueryFormFields - ��������������� ������, ������������ � �������� ��������� ������
 ��� ������ ����� FormFields � ������� MS Word. ����� ���� NULL.
 ���� QueryFormFields == NULL, �� ����������� ������ �������.
 ���� � ���������� wordExportParams �� ������ �����, �� �� QueryFormFields ������������ ������ ������� ������.
 �����:
   1. ������� ����� �������� ��������� ������� � ������������ DataSet.
   2. ������� ����� �������� �������� Filter � ������������ DataSet.
*/
void __fastcall TDocumentWriter::ExportToWordTemplate(const TWordExportParams* wordExportParams, TDataSet *QueryMerge, TDataSet *QueryFormFields)
{
    CoInitialize(NULL);
    result.clear();


    //String TemplateFullName = AppPath + param_word.template_name; // ���������� ���� � �����-�������
    //String SavePath = ExtractFilePath(wordExportParams->resultFileDirectory);         // ���� ��� ���������� �����������
    //String ResultFileNamePrefix = ExtractFileName(DstFileName);     // ������� ����� �����-����������

    //std::vector<String> formFields;    // ������ � ������� ������ - �����������

    if (QueryMerge->RecordCount == 0)
    {
        return;
    }


    MSWordWorks msword;// = new MSWordWorks();
    Variant Document;   // ������

    try
    {
        msword.OpenWord();

        #ifndef NDEBUG
        msword.SetVisible(true);
        msword.SetDisplayAlerts(true);
        #endif

        Document =  msword.OpenDocument(wordExportParams->templateFilename, false);
    }
    catch (Exception &e)
    {
        /*
        switch (e)
        {
        case 1:
        }
        String msg = "��������� ������� ��������� ���������� Microsoft Word."
            "\n����������, ���������� � ���������� ��������������.\n" + e.Message;

        msword.CloseApplication();
        VarClear(Document);

        String msg = "��������� ������� ������ " + wordExportParams->templateFilename +
            "\n����������, ���������� � ���������� ��������������.\n" + e.Message;
        throw Exception(msg);*/
        return;

    }



    // ����� �� ��������� QueryFormFields->RecordCount ?
    bool bFilterExist = wordExportParams->filter_main_field != "" && wordExportParams->filter_sec_field != "";    // ���� � ���������� ����� ������, �� �������, ��� ���������� ������


    if (QueryFormFields == NULL)
    {
        // ���� ����� ���� ������, �� ������ ������ �������
        // ������� ��������� Word � ��������
        if (QueryMerge->RecordCount > 0)
        {

            std::vector<AnsiString> vNew;
            vNew = msword.ExportToWordFields(QueryMerge, Document, wordExportParams->resultFilename, wordExportParams->pagePerDocument);
            result.appendResultFiles(vNew);
        }
    }
    else
    {
        // ���� ������ ��� �������, ��:
        // 1. ���� ����� ������ � ����� ������ ������ ��������� �������
        // 2. ����������� �������� � FormFields-���� � �������
        // 3. ������ �������
        //int n_doc = 0;  // ���������� ����� ��������� ������� (������������ � ����� ������ �����������)
        //int nPadLength = IntToStr(QueryFormFields->RecordCount).Length();

        String oldFilter = QueryMerge->Filter;

        while ( !QueryFormFields->Eof )
        {

            if ( VarIsEmpty(Document) )           // ���� ������ �� ������, ��������� ��� (��������� �� ������ ���� �����)
            {
                Document = msword.OpenDocument(wordExportParams->templateFilename, false);
            }

            if ( bFilterExist )  // ���� �� ��������������� ������� ������ 1 ������, �� ��������� ������
            {
                try
                {
                    String sFilter = wordExportParams->filter_main_field + "='" + QueryFormFields->FieldByName(wordExportParams->filter_sec_field)->AsString + "'";
                    if (oldFilter != "")
                    {
                        sFilter = " AND " + sFilter;
                    }

                    //QueryMerge->Filtered = false;
                    QueryMerge->Filter = oldFilter + sFilter;
                    QueryMerge->Filtered = true;
                }
                catch ( Exception &e )
                {
                    QueryMerge->Filtered = false;
                    //String msg = "��������� ������������ ���������� ������� � ���������� �������� ��� ���������� � ���������� ��������������.\n" + e.Message;
                    //throw Exception(msg);
                    break;
                }

            }

            if (QueryMerge->RecordCount != 0)         // ���� ��� �������, �� ��������� ��� �����
            {

                //������ ����� FormFields
                msword.ReplaceFormFields(Document, QueryFormFields);

                // �������
                std::vector<String> vNew;   // ���������� ��� ���������� ����������� �������

                try
                {
                    vNew = msword.ExportToWordFields(QueryMerge, Document, wordExportParams->resultFilename, wordExportParams->pagePerDocument);
                }
                catch (Exception &e)
                {
                    /*_threadStatus = WM_THREAD_ERROR_IN_PROCESS;
                    _threadMessage = "� �������� ������� ��������� � ���������� ������ ��������� ������."
                        "\n���������� � ���������� ��������������."
                        "\n" + e.Message; */
                    break;
                }

                result.appendResultFiles(vNew);
                vNew.clear();
            }



            #ifndef DEBUG
            msword.CloseDocument(Document);
            VarClear(Document);
            #endif

            if ( bFilterExist )
            {
                QueryFormFields->Next();
            }
            else
            {
                // ���� ������ �� ����������, ����� ������� �� �����
                break;
            }
        }
    }

    if (!VarIsEmpty(Document))      // ���� ������ ������
    {
        #ifndef DEBUG
        msword.CloseDocument(Document);
        VarClear(Document);
        #endif
    }
    msword.CloseApplication();

    CoUninitialize();
}

/*
//---------------------------------------------------------------------------
// ������������ ������ MS EXCEL
void __fastcall TDocumentWriter::ExportToExcel(TOraQuery *OraQuery)
{
    CoInitialize(NULL);

    bool fDone = false;

    // ���������� ���������� �������
    OraQuery->Last();
	int RecCount = OraQuery->RecordCount;

    // ���������� ���������� �����
    int FieldCount = OraQuery->FieldCount;

    Variant data_body;
    Variant data_head;
    DATAFORMAT df_body;
    df_body.reserve(FieldCount);

    int ExcelFieldCount = param_excel.Fields.size();

    try {     // ����������� ������ �����, ������������ ����� �������, ����������� ���� ������
        //data_body = CreateVariantArray(RecCount, FieldCount);  // ������� ������ ��� �������
        //data_head = CreateVariantArray(1, FieldCount);  // ����� �������

        if (ExcelFieldCount >= FieldCount)   // ���������� ���� ���� ���� � ExcelFields
        {
            data_body = CreateVariantArray(RecCount, ExcelFieldCount);     // ������� ������ ��� �������
            data_head = CreateVariantArray(1, ExcelFieldCount);            // ����� �������

            for (unsigned int j = 0; j < ExcelFieldCount; j++)
            {
                data_head.PutElement(param_excel.Fields[j].name, 1, j+1);
                df_body.push_back(param_excel.Fields[j].format);
            }
        }
        else
        {
            data_body = CreateVariantArray(RecCount, FieldCount);  // ������� ������ ��� �������
            data_head = CreateVariantArray(1, FieldCount);         // ����� �������

            // ��������� ����� �������
            for (int j = 1; j <= FieldCount; j++ )  		// ���������� ��� ����
            {
                TField* field = OraQuery->Fields->FieldByNumber(j);
                // ������ ������ �������� � ������� Excel
                AnsiString sCellFormat;

                data_head.PutElement(field->DisplayName, 1, j);
                switch (field->DataType) {  // ����� ������������ � ��������� (�������� ������� � ��.)
                case ftString:
                    sCellFormat = "@";
                    break;
                case ftTime:
                    sCellFormat = "��:��:��";
                    break;
                case ftDate:
                    sCellFormat = "��.��.����";
                    break;
                case ftDateTime:
                    sCellFormat = "��.��.����";
                    break;
                case ftCurrency: case ftFloat:
                    sCellFormat = "0.00";
                    break;
                case ftSmallint: case ftInteger: case ftLargeint:
                    sCellFormat = "0";
                    break;
                default:
                    sCellFormat = "@";
                }
                df_body.push_back(sCellFormat);
	        }
        }
    }
    catch (Exception &e)
    {
        VarClear(data_head);
        VarClear(data_body);
        CoUninitialize();
        _threadMessage = e.Message;
        throw Exception(_threadMessage);
        //fDone = true;
    }

    MSExcelWorks msexcel;
    Variant Workbook;
    Variant Worksheet1;

    if (!fDone)           // ���������� ������� ������
    {
        AnsiString s = "";
        OraQuery->First();	// ��������� � ������ ������ (�� ������ ������)
        VarArrayLock(data_body);
        int i = 1;          // ���������� ����� �������
	    while (!OraQuery->Eof)
        {
		    for (int j = 1; j <= FieldCount; j++ )
            {
        	    s = OraQuery->Fields->FieldByNumber(j)->AsString;
                data_body.PutElement(s, i, j);
            }
            OraQuery->Next();  // ��������� � ��������� ������
            i++;
	    }
        VarArrayUnlock(data_body);
       try
       {
            msexcel.OpenApplication();
            Workbook = msexcel.OpenDocument();
        }
        catch (Exception &e)
        {
            VarClear(data_head);
            VarClear(data_body);
            CoUninitialize();
            _threadMessage = e.Message;
            throw Exception(_threadMessage);
        }
        Worksheet1 = msexcel.GetSheet(Workbook, 1);
    }


    if (!fDone && !VarIsEmpty(Worksheet1))  // ��������� �������� Excel
    {
    //if (!VarIsEmpty(Worksheet1)) { // ��������� �������� Excel
        TDateTime DateTime = TDateTime::CurrentDateTime();

        CELLFORMAT cf_body;
        CELLFORMAT cf_head;
        CELLFORMAT cf_title;
        CELLFORMAT cf_createtime;
        CELLFORMAT cf_sql;

        cf_body.BorderStyle = CELLFORMAT::xlContinuous;
        cf_head.BorderStyle = CELLFORMAT::xlContinuous;
        cf_head.FontStyle = cf_head.FontStyle << CELLFORMAT::fsBold;

        cf_head.bWrapText = false;

        cf_title.FontStyle = cf_title.FontStyle << CELLFORMAT::fsBold;
        cf_createtime.bSetFontColor = true;
        cf_createtime.FontColor = clRed;
        cf_sql.bWrapText = false;

        // ���������� ������ ������
        //std::vector<MSExcelWorks::CELLFORMAT> formats;
        //formats = msexcel.GetDataFormat(ArrayDataBody, 1);
        //std::vector<AnsiString> DataFormat;
        //DataFormat = msexcel.GetDataFormat(ArrayDataBody, 1);
        //for (int i=0; i  < QueryParams.size(); i++) {
            //QueryParams[i].
        //}


        // ��������� ������, �� ��������� ����������, ��������� �������������
        // �������� � ������� ������� ������������� ��������� � ����� "separator",
        Variant data_parameters;
        int param_count = UserParams.size();  // ��������� ������
        int visible_param_count = 0;
        for (int i=0; i <= param_count-1; i++)    // ������������ ���-�� ������������ ����������
        {
            if ( UserParams[i]->isVisible() )
            {
                visible_param_count++;
            }
        }
        if (param_count > 0) {    // ������ ���������� ��� ������ � Excel
            data_parameters = CreateVariantArray(visible_param_count, 1);
            for (int i=0; i <= param_count-1; i++)
            {
                if ( !UserParams[i]->isVisible() )
                {
                    continue;
                }

                if (UserParams[i]->type != "separator")
                {
                    data_parameters.PutElement(UserParams[i]->getCaption() + ": " + UserParams[i]->getDisplay(), i+1, 1);
                }
                else
                {
                    data_parameters.PutElement("[" + UserParams[i]->getCaption() + "]", i+1, 1);
                }
            }
        }

        // ����� ������ �� ���� Excel
        Variant range_title = msexcel.WriteToCell(Worksheet1, param_excel.title_label , 1, 1);
        Variant range_createtime = msexcel.WriteToCell(Worksheet1, "�� ��������� ��: " + DateTime.DateTimeString(), 2, 1);
        Variant range_parameters;
        if (param_count > 0)
        {
            range_parameters = msexcel.WriteTable(Worksheet1, data_parameters, 3, 1);
        }

        Variant range_tablehead = msexcel.WriteTable(Worksheet1, data_head, 3 + visible_param_count, 1);
        Variant range_tablebody = msexcel.WriteTable(Worksheet1, data_body, 4 + visible_param_count, 1, &df_body);

        msexcel.SetRangeFormat(range_tablehead, cf_head);
        msexcel.SetRangeFormat(range_tablebody, cf_body);
        msexcel.SetRangeFormat(range_title, cf_title);
        msexcel.SetRangeFormat(range_createtime, cf_createtime);
        if (param_count > 0)
        {
            msexcel.SetRangeFormat(range_parameters, cf_createtime);
        }


        Variant range_all = msexcel.GetRangeFromRange(range_tablehead, 1, 1, msexcel.GetRangeRowsCount(range_tablebody)+1, msexcel.GetRangeColumnsCount(range_tablebody));


        if (this->param_excel.title_height > 0)
        {
            msexcel.SetRowHeight(range_tablehead, this->param_excel.title_height);    // ������ ������ ��������� �������
        }

        msexcel.SetAutoFilter(range_all);   // �������� ����������
        msexcel.SetColumnsAutofit(range_all);  // ������ ����� �� �����������


        // ��������� ��������� (������ � ��)
        for (int i=0; i < ExcelFieldCount; i++)   // ���������� ���� ���� ���� � ExcelFields
        {    //CELLFORMAT cf_cell;
            //cf_cell.bSetFontColor = true;
            //cf_cell.FontColor = clGreen;

            if (param_excel.Fields[i].bwraptext >= 0)
            {
                CELLFORMAT cf_cell;
                cf_cell.bWrapText = param_excel.Fields[i].bwraptext;
                msexcel.SetRangeFormat(range_tablehead, cf_cell, 1, i+1);
            }

            if (param_excel.Fields[i].width >= 0)
            {
                msexcel.SetColumnWidth(range_tablehead, i+1, param_excel.Fields[i].width);    // ������ ������ ��������� �������
                //msexcel.SetColumnWidth(range_tablehead, 1);    // ������ ������ ��������� �������
            }
        }

        //msexcel.SetRowsAutofit(range_tablehead);


        // ������� ����� sql-������� �� ������ ����
        Variant Worksheet2 = msexcel.GetSheet(Workbook, 2);


        Variant range_sqltext;
        int PartMaxLength = 4000;  // 8 192  - ������������ ����� ������ � ������ EXCEL
        int n = ceil( (float) _mainQueryText.Length() / PartMaxLength);
        for (int i = 1; i <= n; i++)
        {
            AnsiString sQueryPart = _mainQueryText.SubString(((i-1) * PartMaxLength) + 1, PartMaxLength);
            range_sqltext = msexcel.WriteToCell(Worksheet2, sQueryPart, i, 1);
            msexcel.SetRangeFormat(range_sqltext, cf_sql);
        }

        df_body.clear();

        if (DstFileName == "")
        {
            msexcel.SetVisible(Workbook);
        }
        else
        {
            msexcel.SaveDocument(Workbook, DstFileName);
            VarClear(Workbook);
            VarClear(Worksheet1);
            VarClear(Worksheet2);
            msexcel.CloseApplication();
            _resultFiles.push_back(DstFileName);
        }


        //if (ExportMode == EM_EXCEL_FILE) {
        //    msexcel.SaveAsDocument(Workbook, DstFileName);
        //    msexcel.CloseExcel();
        //} else {
            //Workbook.OlePropertySet("Name", "blabla");
        //    msexcel.SetVisibleExcel(true, true);
        //}
    }

    // ������������ ������
    VarClear(data_head);
    data_head = NULL;

    VarClear(data_body);
    data_body = NULL;

    CoUninitialize();

    if (fDone)
    {
        throw Exception("����������.");
    }

}
*/

//---------------------------------------------------------------------------
// ���������� Excel ����� � �������������� ������� xlt
void __fastcall TDocumentWriter::ExportToExcelTemplate(const TExcelExportParams* excelExportParams, TDataSet* QueryTable, TDataSet* QueryFields)
{
    CoInitialize(NULL);

    String TemplateFullName = excelExportParams->templateFilename; // ���������� ���� � �����-�������

    // ��������� ������ MS Excel
    MSExcelWorks msexcel;
    Variant Workbook;
    Variant Worksheet;

    try
    {
        msexcel.OpenApplication();
        Workbook = msexcel.OpenDocument(TemplateFullName);
        Worksheet = msexcel.GetSheet(Workbook, 1);
    }
    catch (Exception &e)
    {
        try
        {
            msexcel.CloseApplication();
        }
        catch (...)
        {
        }
        CoUninitialize();
        String msg = "������ ��� �������� �����-������� " + TemplateFullName + ".\n���������� � ���������� ��������������.";
        throw Exception(msg);
    }

    // ������� ������ ������ �����
    try
    {
        if (QueryFields != NULL)
        {
            msexcel.ExportToExcelFields(QueryFields, Worksheet);
        }
    }
    catch (Exception &e)
    {
        msexcel.CloseApplication();
        CoUninitialize();
        //String msg = e.Message;
        throw Exception(e);
    }

    // ����� ��������� ��������� �����
    try
    {
        if (QueryTable != NULL && excelExportParams->table_range_name != "") // ������ ���� ������ ��� ��������� �������� �����
        {
            msexcel.ExportToExcelTable(QueryTable, Worksheet, excelExportParams->table_range_name, excelExportParams->fUnbounded);
        }
    }
    catch (Exception &e)
    {
        try {
            msexcel.CloseApplication();
        }
        catch (...)
        {
        }
        CoUninitialize();
        //_threadMessage = e.Message;
        throw Exception(e);
    }

    if (excelExportParams->resultFilename == "")         // ������ ��������� ��������, ���� ��� �����-���������� �� ������
    {
        msexcel.SetVisible(Workbook);
    }
    else
    {                        // ����� ��������� � ����
        try
        {
            msexcel.SaveDocument(Workbook, excelExportParams->resultFilename);
            msexcel.CloseApplication();
            result.addResultFile(excelExportParams->resultFilename);
        }
        catch (Exception &e)
        {
            try
            {
                msexcel.CloseApplication();
            }
            catch (...)
            {
            }
            CoUninitialize();
            String msg = "������ ��� ���������� ���������� � ���� " + excelExportParams->resultFilename + ".\n" + e.Message;
            throw Exception(msg);
        }
    }

    // � ���������� ������� ���������� �������� � MS Word
    // ����������� ���� ������ QueryFields � QueryTable

    CoUninitialize();
}



/*
//---------------------------------------------------------------------------
// ���������� DBF-�����
// ���������� ��� ������� � ������������� ���������� TDbf
void __fastcall TDocumentWriter::ExportToDBF(TOraQuery *OraQuery)
{
    //TStringList* ListFields;
    //int n = this->param_dbase.Fields.size();
    //if (n > 0)    // ��������� ������ ����� ��� �������� � DBF ("���;���;�����;����� ������� �����")
    //{
    //    ListFields = new TStringList();
    //    for (int i = 0; i < n; i++)
    //    {
    //        ListFields->Add(param_dbase.Fields[i].name + ";" + param_dbase.Fields[i].type + ";"+ param_dbase.Fields[i].length + ";" + param_dbase.Fields[i].decimals);
    //    }
    //}
    //else
   // {
    //    _threadMessage = "�� ����� ������ ����� � ���������� ��������."
    //        "\n����������, ���������� � ���������� ��������������.";
    //    throw Exception(_threadMessage);
    //}

    // ��� ������� ������, � ����� � ���, ��� ��������� ���� ����� ���������� �������
    if (param_dbase.Fields.size() > OraQuery->FieldCount && !param_dbase.fAllowUnassignedFields)
    {
        _threadMessage = "���������� ��������� ����� ��������� ���������� ����� � ��������� ������."
            "\n����������, ���������� � ���������� ��������������.";
        throw Exception(_threadMessage);
    }

    if (param_dbase.Fields.size() == 0) {
        _threadMessage = "�� ����� ������ ����� � ���������� ��������."
            "\n����������, ���������� � ���������� ��������������.";
        throw Exception(_threadMessage);
    }

    // ������� dbf-���� ����������
    TDbf* pTable = new TDbf(NULL);

    //pTableDst->TableLevel = 7; // required for AutoInc field
    pTable->TableLevel = 4;
    pTable->LanguageID = DbfLangId_RUS_866;

    pTable->TableName = ExtractFileName(DstFileName);
    pTable->FilePathFull = ExtractFilePath(DstFileName);


    // ������� ����������� ����� ������� �� ����������
    TDbfFieldDefs* TempFieldDefs = new TDbfFieldDefs(NULL);

    if (TempFieldDefs == NULL) {
        _threadMessage = "Can't create storage.";
        throw Exception(_threadMessage);
    }

    for(std::vector<DBASEFIELD>::iterator it = param_dbase.Fields.begin(); it < param_dbase.Fields.end(); it++ )
    {
        TDbfFieldDef* TempFieldDef = TempFieldDefs->AddFieldDef();
        TempFieldDef->FieldName = it->name;
        //TempFieldDef->Required = true;
        //TempFieldDef->FieldType = Field->type;    // Use FieldType if Field->Type is TFieldType else use NativeFieldType
        TempFieldDef->NativeFieldType = it->type[1];
        TempFieldDef->Size = it->length;
        TempFieldDef->Precision = it->decimals;
    }

    if (TempFieldDefs->Count == 0)
    {
        delete pTable;
        _threadMessage = "�� ������� ��������� �������� �����.";
        throw Exception(_threadMessage);
    }

    pTable->CreateTableEx(TempFieldDefs);
    pTable->Exclusive = true;
    try
    {
        pTable->Open();
    }
    catch (Exception &e)
    {
        _threadMessage = e.Message;
    }

    // ������ ������ � �������
    try
    {
	    while ( !OraQuery->Eof )
        {
            pTable->Append();
		    for (int j = 1; j <= OraQuery->FieldCount; j++ )
            {
                pTable->Fields->FieldByNumber(j)->Value = OraQuery->Fields->FieldByNumber(j)->Value;
            }
            OraQuery->Next();  // ��������� � ��������� ������
	    }
        pTable->Post();
        pTable->Close();

        _resultFiles.push_back(DstFileName);

    }
    catch(Exception &e)
    {
        pTable->Close();

        delete TempFieldDefs;
        delete pTable;

        _threadMessage = e.Message;
        throw Exception(e);
    }

    delete TempFieldDefs;
    delete pTable;
}       */
