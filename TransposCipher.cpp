//---------------------------------------------------------------------------

#include "TransposCipher.h"


TTransposCipher::TTransposCipher()
{
}

TTransposCipher::~TTransposCipher()
{
}

//---------------------------------------------------------------------------
// ����������
AnsiString __fastcall TTransposCipher::Encrypt(AnsiString SrcStr, const int* _pRouteRow, const int* _pRouteCol, AnsiString Extension)
{
    char* Src = SrcStr.c_str();
    int n_src = strlen(Src);        // ����� �������� ������

    int size = ceil(sqrt(n_src));   // ������ ���������� ����������� �������
    int n_dst = pow(size,2);        // ������ ������-����������

    int* pRouteRow = new int[size];
    memcpy(pRouteRow, _pRouteRow, size * sizeof(_pRouteRow));

    int* pRouteCol = new int[size];
    memcpy(pRouteCol, _pRouteCol, size * sizeof(_pRouteCol));


    // ��������� ���������� ������ ��� ����� ��������� ������ - �������� ���.��������
    char* Ext;
    int n_ext;
    if (Extension != "") {
        Ext = Extension.c_str();
        n_ext = strlen(Ext);
    } else {
        Ext = Src;
        n_ext = n_src;
    }

    // ����� ����� ������ ������ ���� ����� �������� �����
    char* SrcExt = new char[n_dst+1];
    SrcExt[n_ext] = '\0';
    for (int i = 0; i < n_src; i++) {
        SrcExt[i] = Src[i];
    }
    for (int i= n_src; i < n_dst; i++) {
        SrcExt[i] = Ext[i % n_ext];
    }


    // ��������� ����������� ������� ������������ �� �������
    int* RouteRow = new int[size];
    for (int i=0; i<size; i++) {
        int minindex = i;
        for (int j = 0; j < size; j++) {    // ����� ������������ ��������
            if (pRouteRow[j] < pRouteRow[minindex]) {
                minindex = j;
            }
        }
        RouteRow[minindex] = i;
        pRouteRow[minindex] = 9999;
    }

    // ��������� ����������� ������� ������������ �� ��������
    int* RouteCol = new int[size];
    for (int i=0; i < size; i++) {
        int minindex = i;
        for (int j = 0; j < size; j++) {    // ����� ������������ ��������
            if (pRouteCol[j] < pRouteCol[minindex]) {
                minindex = j;
            }
        }
        RouteCol[minindex] = i;
        pRouteCol[minindex] = 9999;
    }

    // ������������
    char** TmpArr = new char*[size];    // ������������� ������, ��� �������� ����� ��������������

    // �������������� ����������� ������� � ��������� ������
    int k = 0;
    for(int i=0; i< size; i++) {
        TmpArr[i] = new char[size];
        for(int j=0; j < size; j++) {
            TmpArr[i][j] = SrcExt[k++];
        }
    }

    // ������� ����� ������ ��� �������� �����������
    char** DstArr = new char*[size];

    // ������������ �����
    for(int i=0; i < size; i++) {
        DstArr[i] = new char[size];
        for(int j=0; j < size; j++) {
            DstArr[i][j] = TmpArr[RouteRow[i]][j];
        }
    }

    // ������ ������� ��������� ������ �
    char** pTmp = DstArr;
    DstArr = TmpArr;
    TmpArr = pTmp;

    // ������������ ��������
    for(int j=0; j < size; j++) {
        for(int i=0; i < size; i++) {
            DstArr[i][j] = TmpArr[i][RouteCol[j]];
        }
    }

    // �������������� �� ����������� ������� � ��������
    char* Dst = new char[n_dst+1];
    Dst[n_dst] = '\0';
    k = 0;
    for(int i=0; i < size; i++) {
        for(int j=0; j < size; j++) {
            Dst[k++] = DstArr[i][j];
        }
    }

    AnsiString result = AnsiString(Dst);

    // ������� ������
    for(int i=0; i < size; i++) {
        delete []TmpArr[i];
        delete []DstArr[i];
    }

    delete []TmpArr;
    delete []DstArr;    // ��� ����� ���� ��� ������� �������
    delete Dst;
    delete SrcExt;
    delete []pRouteRow;
    delete []pRouteCol;

    return result;
}

//---------------------------------------------------------------------------
// �����������
AnsiString __fastcall TTransposCipher::Decrypt(AnsiString SrcStr, const int* _pRouteRow, const int* _pRouteCol)
{
    char* SrcExt = SrcStr.c_str();
    int n_dst = strlen(SrcExt);        // ����� �������� ������

    //int n_ext = pow(size,2);        // ����� ������ ������
    double tmp_size = sqrt(n_dst);

    int ceil_tmp_size = ceil(tmp_size);
    if (tmp_size != ceil_tmp_size)
        throw Exception("The length of the string is not equal to the square of any number.");

    int size = (int)tmp_size;   // ������ ���������� ����������� �������

    int* pRouteRow = new int[size];
    memcpy(pRouteRow, _pRouteRow, size * sizeof(_pRouteCol));
    int* pRouteCol = new int[size];
    memcpy(pRouteCol, _pRouteCol, size * sizeof(_pRouteCol));

    // ��������� ����������� ������� ������������ �� �������
    int* RouteRow = new int[size];
    for (int i=0; i < size; i++) {
        //int min = pRouteRow[i];
        int minindex = i;
        for (int j = 0; j < size; j++) {    // ����� ������������ ��������
            if (pRouteRow[j] < pRouteRow[minindex]) {
                //min = pRouteRow[j];
                minindex = j;
            }
        }
        RouteRow[minindex] = i;
        pRouteRow[minindex] = 9999;
    }

    // ��������� ����������� ������� ������������ �� ��������
    int* RouteCol = new int[size];
    for (int i=0; i < size; i++) {
        int minindex = i;
        for (int j = 0; j < size; j++) {    // ����� ������������ ��������
            if (pRouteCol[j] < pRouteCol[minindex]) {
                minindex = j;
            }
        }
        RouteCol[minindex] = i;
        pRouteCol[minindex] = 9999;
    }

    // ������������
    char** TmpArr = new char*[size];    // ������������� ������, ��� �������� ����� ��������������

    // �������������� ����������� ������� � ��������� ������
    int k = 0;
    for(int i=0; i < size; i++) {
        TmpArr[i] = new char[size];
        for(int j=0; j < size; j++) {
            TmpArr[i][j] = SrcExt[k++];
        }
    }


    // ������� ����� ������ ��� �������� �����������
    char** DstArr = new char*[size];
    for(int i=0; i < size; i++)
        DstArr[i] = new char[size];


    // ������������ �����
    for(int i=0; i < size; i++) {
        for(int j=0; j < size; j++) {
            DstArr[RouteRow[i]][j] = TmpArr[i][j];
        }
    }


    // ������ ������� ��������� ������ �
    char** pTmp = DstArr;
    DstArr = TmpArr;
    TmpArr = pTmp;

    // ������������ ��������
    for(int j=0; j < size; j++) {
        for(int i=0; i < size; i++) {
            DstArr[i][RouteCol[j]] = TmpArr[i][j];
        }
    }


    // �������������� �� ����������� ������� � ��������
    char* Dst = new char[n_dst+1];
    Dst[n_dst] = '\0';

    k = 0;
    for(int i=0; i < size; i++) {
        for(int j=0; j < size; j++) {
            Dst[k++] = DstArr[i][j];
        }
    }

    AnsiString result = AnsiString(Dst);

    // ������� ������
    for(int i=0; i < size; i++) {
        delete []TmpArr[i];
        delete []DstArr[i];
    }

    delete []TmpArr;
    delete []DstArr;
    delete Dst;
    delete []pRouteRow;
    delete []pRouteCol;

    return result;
}

