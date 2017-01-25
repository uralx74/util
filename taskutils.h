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

#ifndef TASKUTILS_H
#define TASKUTILS_H

#include <vector.h>
#include <MemDS.hpp>
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ComCtrls.hpp>

#include "Messages.h"


using namespace std;

enum str_pad_type {STR_PAD_LEFT = 0, STR_PAD_RIGHT};

typedef struct { //  ��������� ��� ������������� � �������� ExplodeByBackslash
    bool fBacksleshed;   // ������� ����, ��� ��������� ��������� � ������ � ����� ���������������� �����������
    AnsiString text;    // ���������� ���������
} EXPLODESTRING;

typedef struct { //  ��������� ��� ������������� � �������� ExplodeByBackslash
    bool fBacksleshed;   // ������� ����, ��� ��������� ��������� � ������ � ����� ���������������� �����������
    AnsiString text;    // ���������� ���������
    int startpos; // ������� � ������ ����������� ���������
    int endpos;   // ������� � ������ ����������� ���������
    String startsep;
    String endsep;
} EXPLODESTRING2;


/*
//------------------------------------------------------------------------------
//
inline int MessageBoxInf(String msg, String title, unsigned short flags = MB_ICONINFORMATION + MB_OK + MB_SYSTEMMODAL + MB_SETFOREGROUND + MB_TOPMOST)
{
    return(Application->MessageBox(msg.c_str(), title.c_str(), flags));
}

//------------------------------------------------------------------------------
// ��������� MB_ICONINFORMATION
inline int MessageBoxInf(String msg, unsigned short flags = MB_ICONINFORMATION + MB_OK + MB_SYSTEMMODAL + MB_SETFOREGROUND + MB_TOPMOST)
{
    return(Application->MessageBox(msg.c_str(), Application->Title.c_str(), flags));
}

//------------------------------------------------------------------------------
// ��������� MB_ICONQUESTION
inline int MessageBoxQuestion(String msg, unsigned short flags = MB_ICONQUESTION + MB_YESNO + MB_SYSTEMMODAL + MB_SETFOREGROUND + MB_TOPMOST)
{
    return(Application->MessageBox(msg.c_str(), Application->Title.c_str(), flags));
}

//------------------------------------------------------------------------------
// ��������� MB_ICONSTOP
inline int MessageBoxStop(String msg, unsigned short flags = MB_ICONSTOP + MB_OK + MB_SYSTEMMODAL + MB_SETFOREGROUND + MB_TOPMOST)
{
    return(Application->MessageBox(msg.c_str(), Application->Title.c_str(), flags));
}*/

//------------------------------------------------------------------------------
// ��� �� TDate
Word YearOf(const TDate dt)
{
    unsigned short dd, mm, yyyy;
    dt.DecodeDate(&yyyy, &mm, &dd);
    return yyyy;
}

//------------------------------------------------------------------------------
// ����� �� TDate
Word MonthOf(const TDate dt)
{
    unsigned short dd, mm, yyyy;
    dt.DecodeDate(&yyyy, &mm, &dd);
    return mm;
}

//------------------------------------------------------------------------------
// ���� �� TDate
Word DayOf(const TDate dt)
{
    unsigned short dd, mm, yyyy;
    dt.DecodeDate(&yyyy, &mm, &dd);
    return dd;
}

//------------------------------------------------------------------------------
// ���������� ���� � ������
Word DaysInAMonth(const TDate dt)
{
    unsigned short dd, mm, yyyy;
    dt.DecodeDate(&yyyy, &mm, &dd);
    int leap = IsLeapYear(yyyy) ? 1 : 0;
    return Sysutils::MonthDays[leap][mm - 1];
}

//---------------------------------------------------------------------------
// ��������� ������ ������
String StrPadR(String Str, int Length, String Symb)
{
    while (Str.Length() < Length)
    {
        Str += Symb;
    }
    return Str;
}

//---------------------------------------------------------------------------
// ��������� ������ �����
String StrPadL(String Str, int Length, String Symb)
{
    while (Str.Length() < Length)
    {
        Str = Symb + Str;
    }
    return Str;
}


String getStrParamValue(const String& str, const String& blockName, const String& paramName)
{
           // String blockName = "IMG";
            //String paramName = "ZORDER";
            String paramValue = "";
            int p0 = str.Pos("[" + blockName);

            if (p0 >= 0)
            {
                int p1 = PosEx(paramName, str, PosEx(paramName, str, p0));
                if (p1 > 0)
                {
                    p1 = PosEx("=", str, p1 + paramName.Length());
                    p1 = PosEx("\"", str, p1 + 1);
                }
                if (p1 > 0)
                {
                    p1 = p1 + 1;
                    //int offset = p1+1;
                    int p2 = PosEx("\"", str, p1);
                    paramValue = str.SubString(p1, p2-p1);
                }
            }
	return paramValue;

}
//------------------------------------------------------------------------------
//
void ExploreDirectory(HWND Handle, AnsiString Path)
{
	ShellExecute(Handle, "OPEN", Path.c_str(), NULL, NULL, SW_SHOWNORMAL);
}

void ExploreFile(HWND Handle, AnsiString Path)
{
    ShellExecute(Handle, "OPEN", "EXPLORER", ("/select, " + Path).c_str(), NULL ,SW_NORMAL);
}

//------------------------------------------------------------------------------
// �������� ������� ���� varVariant
Variant CreateVariantArray(int RowCount, int ColCount)
{
    int Bounds[4] = {1, RowCount, 1, ColCount};
    return VarArrayCreate(Bounds, 3,  varVariant);
}

//------------------------------------------------------------------------------
// �������� ������� ���� varVariant
Variant CreateVariantArray(int RowCount)
{
    int Bounds[2] = {1, RowCount};
    return VarArrayCreate(Bounds, 3,  varVariant);
}

TDateTime __fastcall ReplaceDay(const TDateTime& dt, unsigned short day)
{
    unsigned short dd, mm, yyyy;
    dt.DecodeDate(&yyyy, &mm, &dd);
    return EncodeDate(yyyy, mm, day);
}

//------------------------------------------------------------------------------
// �������� ������� ���� varVariant
void RedimVariantArray(Variant *DataArray, int RowCount, int ColCount)
{

    int nSrcRows = VarArrayHighBound(*DataArray, 1);
    int nSrcCols = VarArrayHighBound(*DataArray, 2);

    Variant ResultArray = CreateVariantArray(RowCount, ColCount);
    if (RowCount > nSrcRows) RowCount = nSrcRows;
    if (ColCount > nSrcCols) ColCount = nSrcCols;
    for (int i = 1; i <= RowCount; i++)
        for (int j = 1; j <= ColCount; j++)
            ResultArray.PutElement(DataArray->GetElement(i, j), i, j);
            //ResultArray[i,j] = DataArray[i,j];

    VarClear(*DataArray);
    *DataArray = ResultArray;
}

/*//------------------------------------------------------------------------------
// ���������� ������ �������� � ���� ������ ��������� �����������
string Implode(const vector<string> &pieces, const string &glue = "")
{
	string a;
	int leng=pieces.size();
 	for(int i=0; i<leng; i++)
 	{
 		a+= pieces[i];
 		if (  i < (leng-1) )
 			a+= glue;
 	}
 	return a;
}     */

//------------------------------------------------------------------------------
// ���������� ������ �������� � ���� ������ ��������� �����������
AnsiString Implode(const vector<AnsiString> &pieces, const AnsiString &glue = "")
{
	AnsiString a;
	int leng=pieces.size();
 	for(int i=0; i<leng; i++)
 	{
 		a+= pieces[i];
 		if (  i < (leng-1) )
 			a+= glue;
 	}
 	return a;
}

//------------------------------------------------------------------------------
// ���������� ������ �������� � ���� ������ ��������� �����������
// ����������� ������ �������� �� ������ ������ ������
AnsiString ImplodeNvl(const vector<AnsiString> &pieces, const AnsiString &glue = "")
{
	AnsiString result;
	int leng=pieces.size();
 	for(int i=0; i<leng; i++)
 	{
        if (pieces[i] == "")
            continue;
        if (result != "")
            result += glue;
 		result+= pieces[i];
 	}
 	return result;
}

//------------------------------------------------------------------------------
// ��������� � ������ ������ �� ������ ��������
void PushBackNvl(vector<AnsiString> &v, AnsiString value)
{
    if (value != "")
        v.push_back(value);
}
//------------------------------------------------------------------------------
// ���������� ������ �������� � ���� ������ ��������� �����������
AnsiString Implode(const vector<EXPLODESTRING> &pieces, const AnsiString &glue = "")
{
	AnsiString a;
	int leng = pieces.size();
 	for(int i=0; i < leng; i++)
 	{
 		a += pieces[i].text;
 		if (  i < (leng-1) )
        {
 			a += glue;
        }
 	}
 	return a;
}


/*//------------------------------------------------------------------------------
// �������� � ������ ��� ��������� ��������� �� ������ ���������
AnsiString& ReplaceAll(AnsiString& context, const AnsiString& from, const AnsiString& to)
{
	size_t lookHere = 0;
	size_t foundHere;
	while((foundHere = context.find(from, lookHere)) != string::npos)
	{

 		context.replace(foundHere, from.size(), to);
 		lookHere = foundHere + to.size();
 	}
 	return context;
}
    */
/*
//------------------------------------------------------------------------------
// �������� � ������ ��� ��������� ��������� �� ������ ���������
string& ReplaceAll(string& context, const string& from, const string& to)
{
	size_t lookHere = 0;
	size_t foundHere;
	while((foundHere = context.find(from, lookHere)) != string::npos)
	{
 		context.replace(foundHere, from.size(), to);
 		lookHere = foundHere + to.size();
 	}
 	return context;
}
    */
/**/

//------------------------------------------------------------------------------
// ��������� ������ �� ������ ����� ��������� ��� ��������� ��������� ���������
std::vector<AnsiString> Explode(AnsiString str, AnsiString separator, bool addEmpty = true)
{   // ������� ������ �� ��������� � ������
    // addEmpty - ��������� � ��������� ������ ������
	std::vector<AnsiString> results;
 	unsigned int found1 = 1;
 	unsigned int found2;
    int nLengthSep = separator.Length();

 	found2 = PosEx(separator,str, 1);
 	while(found2 > 0)
    {
 		if(found2 > 0 || addEmpty)
        {
            if (found2 > 0)
            {
 			    results.push_back(str.SubString(found1, found2-found1));
                 //ss =  str.SubString(found1, found2-found1);
            }
		}
        found1 = found2 + nLengthSep;
        if (found2 !=0)
        {
 		    found2 = PosEx(separator, str, found1);
        }
 	}
 	if(str.Length() > 0 || addEmpty)
    {
        results.push_back(str.SubString(found1, str.Length()-found1+1));
        //ss =  str.SubString(found1, str.Length()-found1+1);
 		// results.push_back(str);   // ����������������� ����� �������� 30.06.2015! ����� ������������� ��������������
	}
	return results;

}


/*
//------------------------------------------------------------------------------
// ��������� ������ �� ������ ����� ��������� ��� ��������� ��������� ���������
vector<AnsiString> Explode(AnsiString &const str, AnsiString separator, bool addEmpty = true)
{   // ������� ������ �� ��������� � ������
    // addEmpty - ��������� � ��������� ������ ������

	vector<AnsiString> results;
 	unsigned int found;
    //found = str.find_first_of(separator);
 	found = str.Pos(separator);
 	while(found != string::npos){
 		if(found > 0 || addEmpty){
 			results.push_back(str.SubString(1, found));
 		}
 		str = str.SubString(1, found+1);
 		//found = str.find_first_of(separator);
 		found = str.Pos(separator);
 	}
 	if(str.Length() > 0 || addEmpty){
 		results.push_back(str);
	}
	return results;
} */


/*//------------------------------------------------------------------------------
// ��������� ������ �� ������ ����� ��������� ��� ��������� ��������� ���������-������� ������ � �����
vector<string> ExplodeByBackslash(string str, string separatorstart, string separatorend, vector<bool>& backslash, bool addEmpty = true)
{
	vector<string> results;
 	unsigned int found_start, found_end;
    //found = str.find_first_of(separator);
    int seplength_start = separatorstart.length();
    int seplength_end = separatorend.length();

 	found_start = str.find(separatorstart);
    found_end = str.find(separatorend, found_start+seplength_start);

    // ���� addEmpty ��������� ��������� �� ������ �������� ���������
   /*if(addEmpty && found_start > 0){
 		results.push_back(str.substr(0, found_start));
       	backslash.push_back(false);
	}*/

/* 	while(found_start != string::npos && found_end != string::npos && found_start < found_end){
        // ���� addEmpty � ����� ����������� � ����� ����������� ���� �����, ��������� ���� �����

        if (addEmpty && found_start > 0) {      // ���� addEmpty ��������� ��������� �� ������ �������� ���������
 			results.push_back(str.substr(/*found_end_old*//*0, found_start));
/*        	backslash.push_back(false);
        }

        // ��������� ������ � �������
        results.push_back(str.substr(found_start, found_end+seplength_end-found_start));
        backslash.push_back(true);

        // ����������� ������
 		str = str.substr(found_end+seplength_end);

        //found_end_old = found_end+seplength_end;
 		found_start = str.find(separatorstart, 0);
    	found_end = str.find(separatorend, found_start+seplength_start);
 	}

    // ���� addEmpty, ��������� ��������� ������
 	if(addEmpty && str.length()){
 		results.push_back(str);
       	backslash.push_back(false);
	}
	return results;
}    */



//------------------------------------------------------------------------------
// ��������� ������ �� ������ ����� ��������� ��� ��������� ��������� ���������-������� ������ � �����
// ����������� ������ ������� ExplodeByBackslash
vector<EXPLODESTRING2> ExplodeByBackslash2(AnsiString str, AnsiString separatorstart, AnsiString separatorend, bool addEmpty = true)
{
    //str = "_date(0,(0),0,0,'mm.yyyy')";

	vector<EXPLODESTRING2> result;
 	unsigned int found_start, found_end;

    int seplength_start = separatorstart.Length();
    int seplength_end = separatorend.Length();

 	found_start = str.Pos(separatorstart);
    found_end = 1;

    EXPLODESTRING2 item;

 	while(found_start != 0 && found_end != 0 /*&& found_start < found_end*/){
        if (addEmpty && found_start > 0) { // �������� �� ��������
            item.text = str.SubString(found_end, found_start-found_end);
            item.fBacksleshed = false;
            item.startpos = found_end;    // ���� �����������!!!!!!!!!!!!!!!!!!
            item.endpos = found_end;      // ���� �����������!!!!!!!!!!!!!!!!!!
            result.push_back(item);
        }
        found_end = PosEx(separatorend, str, found_start+seplength_start)+seplength_end;

        // �������� � �������
        item.text = str.SubString(found_start+seplength_start, found_end-found_start-seplength_start-seplength_end);
        item.fBacksleshed = true;
        item.startpos = found_end;    // ���� �����������!!!!!!!!!!!!!!!!!!
        item.endpos = found_end;      // ���� �����������!!!!!!!!!!!!!!!!!!
        item.startsep = separatorstart;    // ���� �����������!!!!!!!!!!!!!!!!!!
        item.endsep = separatorend;      // ���� �����������!!!!!!!!!!!!!!!!!!

        result.push_back(item);

 		found_start = PosEx(separatorstart, str, found_end);
 	}

    // �������� �� �������� � ����� ������
 	if(addEmpty && (found_end < str.Length())){
        item.text = str.SubString(found_end, str.Length()-found_end+1);
        item.fBacksleshed = false;
        item.startpos = found_end;    // ���� �����������!!!!!!!!!!!!!!!!!!
        item.endpos = found_end;      // ���� �����������!!!!!!!!!!!!!!!!!!
 		result.push_back(item);
	}
	return result;
}


//------------------------------------------------------------------------------
// ��������� ������ �� ������ ����� ��������� ��� ��������� ��������� ���������-������� ������ � �����
vector<EXPLODESTRING> ExplodeByBackslash(AnsiString str, AnsiString separatorstart, AnsiString separatorend, bool addEmpty = true)
{
	vector<EXPLODESTRING> result;
 	unsigned int found_start, found_end;

    int seplength_start = separatorstart.Length();
    int seplength_end = separatorend.Length();

 	found_start = str.Pos(separatorstart);
    found_end = 1;

    EXPLODESTRING item;

 	while(found_start != 0 && found_end != 0 /*&& found_start < found_end*/)
    {
        if (addEmpty && found_start > 0)  // �������� �� ��������
        {
            item.text = str.SubString(found_end, found_start-found_end);
            item.fBacksleshed = false;
            result.push_back(item);
        }
        // ������� ���������� ����������� �������
        found_end = PosEx(separatorend, str, found_start+seplength_start) + seplength_end;

        // �������� � �������
        item.text = str.SubString(found_start, found_end - found_start);
        item.fBacksleshed = true;

        result.push_back(item);

 		found_start = PosEx(separatorstart, str, found_end);
 	}

    // �������� �� �������� � ����� ������
 	if( addEmpty && (found_end <= str.Length()) )
    {
        item.text = str.SubString(found_end, str.Length() - found_end + 1);
        item.fBacksleshed = false;
 		result.push_back(item);
	}
	return result;
}

/*//------------------------------------------------------------------------------
// ������� ������� � ������ ����� � ������
void trim(string& s)
{
	size_t p = s.find_first_not_of(" \t");
	s.erase(0, p);
	p = s.find_last_not_of(" \t");
	if (string::npos != p)
	s.erase(p+1);
}  */

/*
//------------------------------------------------------------------------------
// ��������� ������ ������ ������� �������� ������
string str_pad(string input, int pad_length, string pad_string, int pad_type)
{
    int n = input.length();
    int npad = pad_string.length();
    string spad = "";
    while (n < pad_length)
    {
        spad = spad + pad_string;
        n = n + npad;
    }

    return pad_type == STR_PAD_LEFT? spad + input : input + spad ;
} */


//------------------------------------------------------------------------------
// ��������� ������ ������ ������� �������� ������
/*AnsiString str_pad(const AnsiString &input, int pad_length, const AnsiString &pad_string, int pad_type)
{
    int n = input.Length();
    int npad = pad_string.Length();
    AnsiString spad = "";
    while (n < pad_length)
    {
        spad = spad + pad_string;
        n = n + npad;
    }

    return pad_type == STR_PAD_LEFT? spad + input : input + spad ;
}*/


/*
//------------------------------------------------------------------------------
// ��������� ������ ������ ������� �������� ������
AnsiString str_pad(const AnsiString &input, const AnsiString &pad_string, int pad_length, int pad_type)
{
    int n = input.Length();
    int npad = pad_string.Length();
    AnsiString spad = "";
    while (n < pad_length)
    {
        spad = spad + pad_string;
        n = n + npad;
    }

    return pad_type == STR_PAD_LEFT? spad + input : input + spad ;
} */

//------------------------------------------------------------------------------
// ��������� �������� �� ������ Date dd.mm.yyyy
bool IsDate(String str)
{
    int l = str.Length();

    if ( l <= 0)
    {
        return 0;
    }

    if (l == 10)
    {
            if (str[3] == str[6] && (str[3] == '.' || str[3] =='/' || str[3] =='-' ) &&
                isdigit(str[1]) && isdigit(str[2]) && isdigit(str[4]) && isdigit(str[5]) && isdigit(str[7]) && isdigit(str[8]) && isdigit(str[9])&& isdigit(str[10]))
            {
                return true;
            }
     }
     return false;
}

//------------------------------------------------------------------------------
// ��������� �������� �� ������ Time hh:mm:ss
bool IsTime(String str)
{   
    int l = str.Length();

    if ( l <= 0)
    {
        return 0;
    }

    if (l == 8)
    {
            if (str[3] == str[6] && (str[3] == ':') &&
                isdigit(str[1]) && isdigit(str[2]) && isdigit(str[4]) && isdigit(str[5]) && isdigit(str[7]) && isdigit(str[8]))
            {
                return true;
            }
     }
     return false;
}

//------------------------------------------------------------------------------
//
bool IsDataTime(String str)
{   // ������� �������� �������� �� ������ DataTime dd.mm.yyyy hh:mm:ss
    int l = str.Length();

    if ( l <= 0)
    {
        return 0;
    }

    if (l != 18)
    {
        return false;
    }

    return IsDate(str.SubString(0,10)) && IsTime(str.SubString(10,17));
}


//------------------------------------------------------------------------------
// ��������� �������� �� ������ ������ � ��������� ������
bool IsFloat(String str)
{   // ������� �������� �������� �� ������ Float 9,9999999
    int l = str.Length();

    if ( l <= 0 )
    {
        return false;
    }

    bool oneSep = false;

    for (int i = 1; i <= l; i++)
    {
        char a = str[i];
        if (a == '1' || a== '2' || a == '3' || a == '4' || a== '5' ||
            a== '6' || a == '7' || a == '8' || a== '9' || a== '0') {
            continue;
        };

        if (a== '-' || a =='+')      // ���� + ��� - �� � ������ ������
        {
            if (i > 1)
            {
                return false;
            }
            else
            {
                continue;
            }
        }

        if  (a=='.' || a== ',')
        {
            if (i == 1 || oneSep)   // ���� ������ ������ ����������� ������� �����
            {
                return false;
            }
            else
            {
                oneSep = true;
                continue;
            }
        }
        return false;
    }

    return oneSep;
}

//------------------------------------------------------------------------------
// ��������� �������� �� ������ ����� ������
bool IsInt(String str)
{   // ������� �������� �������� �� ������ Int 9999999
    int l = str.Length();

    if ( l <= 0)
    {
        return 0;
    }

    for (int i = 1; i <= l; i++)
    {
        char a = str[i];
        if (a == '1' || a== '2' || a == '3' || a == '4' || a== '5' ||
            a== '6' || a == '7' || a == '8' || a== '9' || a== '0')
        {
            continue;
        };

/*        if (a == '0') {             // ���� ���� � ������ ������
            if (i == 1 && l > 1)
                return false;
            continue;
        } */

        if (a== '-' || a =='+')      // ���� + ��� - �� � ������ ������
        {
            if (i > 1)
            {
                return false;
            }
            else
            {
                continue;
            }
        }
        return false;
    }
    return true;
}

//----------------------------------------------------------------------------
// ��������� ����� ������� �� ��� PID'�
bool __fastcall KillProcess(DWORD PID)
{
    bool ReturnCode = false;
    HANDLE hProcess = OpenProcess(PROCESS_TERMINATE, false, PID);
    if (hProcess != NULL || hProcess != INVALID_HANDLE_VALUE)
    {
        if (TerminateProcess(hProcess, -1))
        {
            ReturnCode = true;
        }
        CloseHandle( hProcess );
    }
    return ReturnCode;
}

//----------------------------------------------------------------------------
// ������� ���� � ���������� ����������
AnsiString __fastcall CreateWorkDir(AnsiString work_dir)
{
    AnsiString tek_kat = ExtractFilePath(Application->ExeName);
    if (! SetCurrentDir(work_dir))
    {
        if (! CreateDir(work_dir))
        {
            Application->MessageBox("������ �� ����� C:\\ !","�������� ��������",MB_ICONSTOP + MB_OK);
        }
    }
    SetCurrentDir(tek_kat);
    return (tek_kat);
}

//------------------------------------------------------------------------------
// ��������� ������ ���� � ���������� �������� ������������ Windows
AnsiString __fastcall GetTempPath()
{
    const DWORD size = 512;
    char TempDirectory[size];
    DWORD Er = GetTempPath(size, TempDirectory);

    if(Er > size || Er == 0)
    {
        Er=GetLastError();
		MessageBoxStop("Error: " + IntToStr(Er));
        return NULL;
    }
    else
    {
        return (AnsiString)TempDirectory;
    }

}

//------------------------------------------------------------------------------
//
AnsiString AddWhere(AnsiString whereblock, AnsiString condition, bool addif)
{
    if (!addif)
    {
        return whereblock;
    }
        
    condition = condition.Trim();
    if (condition.Length() == 0)
    {
        return whereblock;
    }

    whereblock = whereblock.Trim();
    int w_length = whereblock.Length();

    if (w_length > 5)
    {
        whereblock = whereblock + " AND " + condition;
    }
    else
    {
        whereblock = "WHERE " + condition;
    }

    return whereblock;
}

//---------------------------------------------------------------------------
// ��������� Caption-Value
typedef struct {
    AnsiString Caption;
    AnsiString Value;
} TValueListItem;

//---------------------------------------------------------------------------
// ����� ��� �������� �������� Caption-Value
class TValueList {
public:
    __fastcall TValueList();
    __fastcall ~TValueList();
    void Free();
    TValueListItem* GetItem(int ItemIndex);
    void AddItem(AnsiString Caption, AnsiString Value);
    int Size();
private:
    //TList
    TList* pItems;
protected:
    int size;

};

//---------------------------------------------------------------------------
// ValueList
__fastcall TValueList::TValueList()
{
    //Items = new TList;
    pItems = new TList;
    size = 0;
}

//---------------------------------------------------------------------------
//
__fastcall TValueList::~TValueList()
{
    //Clear();
    Free();
    delete pItems;
    pItems = NULL;
}

//---------------------------------------------------------------------------
//
int TValueList::Size()
{
    return size;
}

//---------------------------------------------------------------------------
// �������� Item
TValueListItem* TValueList::GetItem(int ItemIndex)
{
    TValueListItem *Item = (TValueListItem*) pItems->Items[ItemIndex];
    return Item;
}
 /*
//---------------------------------------------------------------------------
// �������� Value
AnsiString ValueList::GetValue(int ItemIndex)
{
    TValueListItem *Item = (ValueListItem*) Items->Items[ItemIndex];
    return Item->Value;
}

//---------------------------------------------------------------------------
// �������� Caption
AnsiString ValueList::GetCaption(int ItemIndex)
{
    TValueListItem *Item = (ValueListItem*) Items->Items[ItemIndex];
    return Item->Caption;
}
   */
//---------------------------------------------------------------------------
// �������� � ������ Caption-Value
void TValueList::AddItem(AnsiString Caption, AnsiString Value)
{
    TValueListItem* Item = new TValueListItem;
    Item->Caption = Caption;
    Item->Value = Value;
    pItems->Add(Item);
    size++;
}

//---------------------------------------------------------------------------
// �������� ������ ��������
void TValueList::Free()
{
    for (int i = 0; i < size; i++)
    {
        delete (TValueListItem*)(pItems->Items[i]);
    }
    pItems->Clear();
    size = 0;
}

//---------------------------------------------------------------------------
// ���������, �������� �� ������ ������ 
bool __fastcall IsNumber(String Value, bool bFloat, bool bSign)
{
    int n = Value.Length();
    bool bSignExist = false;
    int iStart = 1;

    if (bSign && n > 0)    // ���� + ��� - ����� ���� ������ �������
    {
        bSignExist = Value[1] == '+' || Value[1] == '-';
    }

    if (bSignExist)    // ���� ��� ���� + ��� -, ������ �������� �� ������� �������
    {
        iStart = 2;
    }

    bool nDotExist = false;
    for (int i = iStart; i <= n; i++)   // ���� �� ��������
    {
        if (!isdigit(Value[i]) )
        {
            if (bFloat && Value[i] == DecimalSeparator)
            {
                if (nDotExist)       // ���� ����� ��� ���� .
                {
                    return false;
                }
                else
                {
                    nDotExist = true;
                }
            }
            else
            {
                return false;
            }
        }
    }
    return true;
}

//---------------------------------------------------------------------------
// ���������� ������ ������ �� ����������, ������� != ""
String Nvl(String s1, String s2)
{
    if (Trim(s1) == "")
    {
        return s2;
    }
    else
    {
        return s1;
    }
}

//---------------------------------------------------------------------------
// ���������� s2, ���� ������� s1 != "", ����� s3
String Nvl2(String s1, String s2, String s3="")
{
    if (Trim(s1) != "")
    {
        return s2;
    }
    else
    {
        return s3;
    }
}

//---------------------------------------------------------------------------
// ���������� val_true, ���� val = true, ����� val_false
String Iif(bool val, String val_true, String val_false)
{
    return val ? val_true : val_false;
}


/*
String TimeBetween()
{
    int TotalSec = SecsPerDay * (StopTime - StartTime);
    int dd = TotalSec/SecsPerDay;
    int hh = (TotalSec - dd * 24 * 3600) / 3600;
    int mm = (TotalSec / 60) % 60;
    int ss = TotalSec % 60;
    sTotalTime = s
}*/

#endif TASKUTILS_H
