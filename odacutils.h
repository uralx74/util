/*******************************************************************************
    ������������ ������ taskutil.h
    �������� ��������������� �������

    ������ ����� �� 08.10.2014

    // ������� � ��������� ������
    vector<string> Explode(string& const str, string separator, bool addEmpty = true)
    string Implode(const vector<string> &pieces, const string &glue = "")
    vector<string> ExplodeByBackslash(string str, string separatorstart, string separatorend, vector<bool>& backslash, bool addEmpty = true)

    string& ReplaceAll(string& context, const string& from, const string& to)
    void trim(string& s)

    // ������� �����
    bool IsDate(string str)
    bool IsTime(string str)
    bool IsDataTime(string str)
    bool IsFloat(string str)
    bool IsInt(string str)
*******************************************************************************/

#ifndef ODACUTILS_H
#define ODACUTILS_H

#include "Ora.hpp"


//------------------------------------------------------------------------------
//  ������� ���������� �������
int GetRecCount(TOraQuery *OraQuery)
{   // ������� ��� �������� ���������� ������� � OraQuery

    TOraQuery *OraQueryCount = new TOraQuery(NULL);//OraQuery->Last();
    OraQueryCount->Session = OraQuery->Session;
    OraQuery->SQL->Add( "select count(*) N from (" + OraQuery->FinalSQL + ")" );
    OraQueryCount->Open();
    int RecCount = OraQueryCount->FieldByName("N")->AsInteger;

    OraQueryCount->Close();
    delete OraQueryCount;
    OraQueryCount = NULL;

    return RecCount;
}

//------------------------------------------------------------------------------
// �������� � ���������� OraQuery
TOraQuery* OpenOraQuery(TOraSession* OraSession, AnsiString StrQuery, bool FetchAll = true)
{
    TOraQuery* OraQuery = new TOraQuery(NULL);
    OraQuery->FetchAll = FetchAll;
    OraQuery->Session = OraSession;

    //OraQuery->SQL->Clear();
    OraQuery->SQL->Add(StrQuery);

    try
    {
        if (OraQuery->Active)
        {
            OraQuery->Refresh();
        }
        else
        {
            OraQuery->Open();
        }
    }
    catch(Exception &e)
    {
        delete OraQuery;
        OraQuery = NULL;
        //Application->ShowException(&exception);
        throw Exception(e);   // ��������� 2016-03-22
    }
    return OraQuery;
}


// ������ nvl Oracle
String ora_nvl(TField* field, String val1)
{
    return field->IsNull ? val1 : field->AsString;
}

// ������ nvl2 Oracle
String ora_nvl2(TField* field, String val1, String val2)
{
    return field->IsNull ? val2 : val1;
}

#endif ODACUTILS_H
