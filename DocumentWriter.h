#ifndef DOCUMENTWRITER_H
#define DOCUMENTWRITER_H

#include "Ora.hpp"
#include "OraDataTypeMap.hpp"
//#include "odacutils.h"
//#include "math.h"
#include "MSWordWorks.h"
#include "MSExcelWorks.h"
#include <vector>


class TDocumentWriterResult
{
public:
    std::vector<String> resultFiles;
    void __fastcall addResultFile(String filename);
    void __fastcall appendResultFiles(std::vector<String> filenames);
    void __fastcall clear();
};


// ��������� ��� �������� ���������� ���� (�������) DBASE
typedef struct {    // ��� �������� ��������� dbf-�����
    String type;    // ��� fieldtype is a single character [C,D,F,L,M,N]
    String name;    // ��� ���� (�� 10 ��������).
    int length;         // ����� ����
    int decimals;       // ����� ���������� �����
    // Character 	 1-255
    // Date	  8
    // Logical	  1
    // Memo	  10
    // Numeric	1-30
    // Decimals is 0 for non-numeric, number of decimals for numeric.
} DBASEFIELD;


/* Excel parameters */
// ��������� ��� �������� ���������� ���� (�������) MS Excel
typedef struct {    // ��� �������� ������� ����� � Excel
    AnsiString format;      // ������ ������ � Excel
    AnsiString name;        // ��� ����
    //int title_rows;       // ������ ��������� � �������
    int width;              // ������ �������
    int bwraptext;          // ���� ������� �� ������
} EXCELFIELD;

// ��������� ��� �������� ���������� �������� � MS Excel
class TExcelExportParams  {
public:
    String id;
    String label;
    //bool fDefault;
    String templateFilename;       // ��� ����� ������� Excel
    String resultFilename;
    AnsiString title_label;         // ������ - ��������� � �������� ��������� � ������ Excel (��������� � ��������� ���������)
    int title_height;               // ������ ��������� � �������  (��������� � ��������� ���������)
    std::vector<EXCELFIELD> Fields;     // ������ ����� ��� �������� � ���� MS Excel
    String table_range_name;        // ��� ��������� ��������� ����� (��� ������ � ������)
    bool fUnbounded;                    // ���� ����, ��� �������� table_range_name ����� ��������, � ������������ � ����������� ������� � ��������� ������
};

/* Word */
// ��������� ��� �������� ���������� �������� � MS Word
class TWordExportParams
{
public:
    String templateFilename;   // ������ ��� ����� ������� MS Word
    String resultFilename;   // ������ ��� ����� ������� MS Word
    String imageFilesDirectory;   // ������� � ������� ������������, ����������� � ���� [img]
    int pagePerDocument;           // ���������� ������� �� �������� MS Word

    /* DataSets links*/
    String filter_main_field;      // ��� ���� �� ��������� ������� ��� ��������� �� ��������� ���� word_filter_sec_field
    String filter_sec_field;       // ��� ���� �� ���������������� ������� (��. word_filter_main_field)
    //String filter_infix_sec_field; // ��� ���� �� ���������������� �������, �������� �������� ����� ������������ � ����� ��������������� �����
};

// ��������� ��� �������� ���������� �������� � DBF
typedef struct {    // ��� �������� ������� ����� � Excel
    String id;
    String label;
    //bool fDefault;
    bool fAllowUnassignedFields;
    std::vector<DBASEFIELD> Fields;    // ������ ����� ��� ������� � ���� DBF
} EXPORT_PARAMS_DBASE;


class TDocumentWriter
{
private:

public:
    TDocumentWriterResult result;

    void __fastcall ExportToWordTemplate(const TWordExportParams* wordExportParams, TDataSet *QueryMerge, TDataSet *QueryFormFields);  // ���������� ������ Word �� ���� �������
    void __fastcall ExportToExcelTemplate(const TExcelExportParams* excelExportParams, TDataSet* QueryTable, TDataSet* QueryFields);

    //void __fastcall ExportToExcel(TOraQuery *OraQuery); // ���������� ������ Excel
    //void __fastcall ExportToExcelTemplate(TOraQuery *QueryTable, TOraQuery *QueryFields);
    //void __fastcall ExportToDBF(TOraQuery *OraQuery);   // ���������� DBF-�����
};

#endif

