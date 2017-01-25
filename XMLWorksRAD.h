//---------------------------------------------------------------------------
#ifndef XMLWORKS
#define XMLWORKS

/*******************************************************************************
	����� ��� ����� � OLE-������� MSXml.Application
    ������ �� 08.10.2014


*******************************************************************************/

/*
#include "system.hpp"
#include <utilcls.h>
#include "Comobj.hpp"
#include <fstream.h>
#include "taskutils.h"
/**/

class MSXMLWorks
{

private:

public:
    MSXMLWorks();
    void __fastcall LoadXMLFile(UnicodeString XMLFileName);
    void __fastcall LoadXMLText(UnicodeString XMLText);

    Variant __fastcall GetRootNode();
    UnicodeString __fastcall GetNodeName(Variant Node);
    Variant __fastcall GetFirstNode(Variant Node);
    Variant __fastcall GetNextNode(Variant Node);
	UnicodeString __fastcall GetAttributeValue(Variant Node, int AttributeIndex);
	UnicodeString __fastcall GetAttributeValue(Variant Node, UnicodeString AttributeName);
	Variant GetAttribute(Variant Node, UnicodeString AttributeName);

	//AnsiString __fastcall GetValueAttribute(Variant Attribute);
    int __fastcall GetAttributesCount(Variant Node);

	UnicodeString __fastcall GetParseError();

    Variant xmlDoc;
};

//---------------------------------------------------------------------------
//
MSXMLWorks::MSXMLWorks()
{
    //Variant xmlObj = CreateOleObject("Microsoft.XMLDOM");
    //Variant xmlDoc = CreateOleObject("MSXML.DOMDocument");
    xmlDoc = CreateOleObject("Msxml2.DOMDocument.3.0");
    xmlDoc.OlePropertySet("Async", false);
}

//---------------------------------------------------------------------------
//
void __fastcall MSXMLWorks::LoadXMLFile(UnicodeString XMLFileName)
{
    xmlDoc.OlePropertyGet("Load", XMLFileName.c_str());
}

//---------------------------------------------------------------------------
//
void __fastcall MSXMLWorks::LoadXMLText(UnicodeString XMLText)
{
	//StringToOleStr(XMLText);
	/*String s = "<?xml version=""1.0""?>"
	"<parameters>"
	"<parameter>";
		"</parameter>"
		"</parameters>";
		/*"<parameter type="""" name=""date"">
		"</parameter>"
		"</parameters>";*/

	xmlDoc.OlePropertyGet("LoadXML", StringToOleStr(XMLText) );
	//xmlDoc.OlePropertyGet("LoadXML", XMLText.c_str());
}

//---------------------------------------------------------------------------
// ���������, ���������� �� �������
Variant MSXMLWorks::GetAttribute(Variant Node, UnicodeString AttributeName)
{
    return Node.OlePropertyGet("attributes").OleFunction("getNamedItem", AttributeName);
    //return attribute.IsEmpty();
}

//---------------------------------------------------------------------------
// ���������� ���������� ��������� ����
UnicodeString __fastcall MSXMLWorks::GetAttributeValue(Variant Node, int AttributeIndex)
{
    return Node.OlePropertyGet("attributes").OlePropertyGet("item",AttributeIndex).OlePropertyGet("Value");
}

//---------------------------------------------------------------------------
// ���������� �������� �������� �� �����
UnicodeString __fastcall MSXMLWorks::GetAttributeValue(Variant Node, UnicodeString AttributeName)
{
    Variant attribute = Node.OlePropertyGet("attributes").OleFunction("getNamedItem", StringToOleStr(AttributeName));
	if (!attribute.IsEmpty())
		return attribute.OlePropertyGet("text");
    else
        return "";

	// ������ ������
    //return Node.OleFunction("GetAttribute", StringToOleStr(AttributeName));
}

/*//---------------------------------------------------------------------------
// ���������� �������� ��������
AnsiString MSXMLWorks::GetValueAttribute(Variant Attribute)
{
    return Attribute.OlePropertyGet("Value");
}  */

//---------------------------------------------------------------------------
// ���������� ���������� ��������� ����
int __fastcall MSXMLWorks::GetAttributesCount(Variant Node)
{
    return Node.OlePropertyGet("attributes").OlePropertyGet("length");
}

//---------------------------------------------------------------------------
//
Variant __fastcall MSXMLWorks::GetRootNode()
{
    return xmlDoc.OlePropertyGet("DocumentElement");
}

//---------------------------------------------------------------------------
//
UnicodeString __fastcall MSXMLWorks::GetNodeName(Variant Node)
{
    return Node.OlePropertyGet("NodeName");
}

//---------------------------------------------------------------------------
// ��������� ������ �������� ����
Variant __fastcall MSXMLWorks::GetFirstNode(Variant Node)
{
    return Node.OlePropertyGet("firstChild");
}

//---------------------------------------------------------------------------
// ���������� ��������� ���� �� ����������
Variant __fastcall MSXMLWorks::GetNextNode(Variant Node)
{
    return Node.OlePropertyGet("nextSibling");
}

//---------------------------------------------------------------------------
// ��������� ������� ������ ������� XML
UnicodeString __fastcall MSXMLWorks::GetParseError()
{
    if( xmlDoc.OlePropertyGet("parseError").OlePropertyGet("errorCode")!=0 )
	{
		return xmlDoc.OlePropertyGet("parseError").OlePropertyGet("reason");
	} else {
		return "";
    }
}

//Variant oItems = rootNode.OleFunction("SelectNodes", "//item");


/*
//---------------------------------------------------------------------------
//
void __fastcall MSXMLWorks::ParseXML(AnsiString XMLText)
{

}

//---------------------------------------------------------------------------
//
XMLNode* XMLWorks::GetNode(XMLNode *node, int number)
{


}

*/
//---------------------------------------------------------------------------
#endif
