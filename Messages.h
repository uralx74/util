/*******************************************************************************
    ������������ ������ Messages.h
    �������� ��������������� �������

    ������ ����� �� 13.05.2016

*******************************************************************************/

#ifndef MESSAGES_H
#define MESSAGES_H

#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ComCtrls.hpp>


using namespace std;

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
}


#endif MESSAGES_H
