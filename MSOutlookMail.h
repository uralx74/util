//---------------------------------------------------------------------------
#ifndef MSOUTLOOKMAIL
#define MSOUTLOOKMAIL

/*******************************************************************************
	����� ��� ����� � OLE-������� Outlook.Application
    ������ �� 17.06.2015


    �������� ���������� ������ � �����������:
    1. ������� ������ ������ MSExcelWorks
        MSOutlookMail msoutlook;
    2. ������� ������ ������
        Variant MailItem = msoutlook.CreateMailItem();
    3. ������ ��������� ������
        msoutlook.MailItemSetTo(MailItem, "V.Ovchinnikov@cf.esbt.ru");
        msoutlook.MailItemSetSubject(MailItem, "��������");
        msoutlook.MailItemSetBody(MailItem, "����� ������");
        msoutlook.MailItemAddAttachments(MailItem, "c:\\filename.xlsx")
    4. ��������� �������� ������
        msoutlook.SendMail(MailItem);
    5. ������� ������
        msoutlook.Close();

*******************************************************************************/

#include "system.hpp"
#include <utilcls.h>
#include "Comobj.hpp"

class MSOutlookMail
{
private:
	bool bSelfCreate;

public:
    MSOutlookMail();
	Variant __fastcall GetApplication();
	Variant __fastcall CreateMailItem();
	void __fastcall MailItemSetTo(Variant MailItem, String address);
	void __fastcall MailItemSetSubject(Variant MailItem, String subject);
	void __fastcall MailItemSetBody(Variant MailItem, WideString body);
	Variant __fastcall MailItemAddAttachments(Variant MailItem, WideString filename);
	Variant __fastcall SendMail(Variant MailItem);
	Variant __fastcall Close(bool bCloseForce = false);
   	Variant OutlookApp;
    Variant NameSpaceMapi;
};

//---------------------------------------------------------------------------
//
MSOutlookMail::MSOutlookMail()
{
	try {
    	OutlookApp = Variant::GetActiveObject("Outlook.Application.14");
        NameSpaceMapi = OutlookApp.OleFunction("GetNameSpace", "MAPI");
        bSelfCreate = false;
	}
	catch ( ... ) {
        OutlookApp = Variant::CreateObject("Outlook.Application.14");
        NameSpaceMapi = OutlookApp.OleFunction("GetNameSpace", "MAPI");
    	NameSpaceMapi.OleFunction("Logon", "", "", true, true);
	    bSelfCreate = true;
	}
}

//---------------------------------------------------------------------------
//
Variant __fastcall MSOutlookMail::GetApplication()
{
	return OutlookApp;
}

//---------------------------------------------------------------------------
//
Variant __fastcall MSOutlookMail::CreateMailItem()
{
	try
	{
	    Variant MailItem = OutlookApp.OleFunction("CreateItem", 0);
	    return MailItem;
	} catch (...) {
		return Unassigned;
	}
}

//---------------------------------------------------------------------------
//
void __fastcall MSOutlookMail::MailItemSetTo(Variant MailItem, String address)
{
	MailItem.OlePropertySet("To", address);
}

void __fastcall MSOutlookMail::MailItemSetSubject(Variant MailItem, String subject)
{
	MailItem.OlePropertySet("Subject", subject);
}

//---------------------------------------------------------------------------
//
void __fastcall MSOutlookMail::MailItemSetBody(Variant MailItem, WideString body)
{
	MailItem.OlePropertySet("Body", body);
}

//---------------------------------------------------------------------------
//
void __fastcall MSOutlookMail::MailItemAddAttachments(Variant MailItem, WideString filename)
{
	Variant Attachments = MailItem.OlePropertyGet("Attachments");
	Attachments.OleFunction("Add", filename);
}

//---------------------------------------------------------------------------
//
void __fastcall MSOutlookMail::SendMail(Variant MailItem)
{
	MailItem.OleFunction("Send");
}

//---------------------------------------------------------------------------
//
void __fastcall MSOutlookMail::Close(bool bCloseForce)
{
    if (bCloseForce || bSelfCreate) {
        //Variant NameSpaceMapi = OutlookApp.OleFunction("GetNameSpace", "MAPI");
	    NameSpaceMapi.OleFunction("Logoff");
	    OutlookApp.OleFunction("Quit");
    }
}




/*
Variant __fastcall MSOutlookMail::SendMail()
{
	try
	{
        //myItem.OleFunction("Display");
        //myItem.OlePropertyGet("Recipients").OleFunction("Add", "V.Ovchinnikov@cf.esbt.ru");
		//myItem.OlePropertySet("Body", "blablabla");
		//Variant myAttachments = myItem.OlePropertyGet("Attachments");
		//myAttachments.OleFunction("Add", WideString("F:\\fam.txt"));
        //myItem.OleProcedure("Display");
		//myItem.OleProcedure("Send");
        vEspaceDeNom.OleFunction("Logoff");
		MSOApp.OleFunction("Quit");
	}
	catch(...)
	{
		ShowMessage("Error");
	}

}*/


//---------------------------------------------------------------------------
#endif
