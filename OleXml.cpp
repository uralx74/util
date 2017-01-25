#include "OleXml.h"

OleXml::OleXml(bool validateOnParse)
{
    xmlDoc = CreateOleObject("Msxml2.DOMDocument.3.0");
    xmlDoc.OlePropertySet("Async", false);
    xmlDoc.OlePropertySet("validateOnParse", validateOnParse);
    xmlDoc.OlePropertySet("resolveExternals", false);
}

OleXml::~OleXml()
{
    xmlDoc = Unassigned;   // ������������ VarClear(xmlDoc)
}

/* ������� �������� ����
 */
Variant __fastcall OleXml::CreateRootNode(const AnsiString& nodeName)
{
    Variant processingInstructions = xmlDoc.OleFunction("createProcessingInstruction", "xml", "version='1.0' encoding='windows-1251'");
    xmlDoc.OleFunction("appendChild", processingInstructions);

    Variant node = xmlDoc.OleFunction("CreateElement", nodeName);
    xmlDoc.OlePropertySet("DocumentElement", node);

    return node;
}

/* ������� �������� ����
 */
Variant __fastcall OleXml::CreateNode(Variant parentNode, const AnsiString& nodeName)
{
    //node.OlePropertySet("DataType", "bin.base64");

    Variant node = xmlDoc.OleFunction("CreateElement", nodeName);
    return node;
}

/* ��������� �������� ����
 * ���������� ��������, ��� ��� �� ����������� ���� � ��� �� ���� ��������� ���
 */
Variant __fastcall OleXml::AddChildNode(Variant parentNode, Variant childNode)
{
    Variant node = parentNode.OleFunction("appendChild", childNode);
    return node;
}

/* ��������� ����
   bool deep - ��������� ���������� ��������� ���� ��� ���
 */
Variant __fastcall OleXml::CloneNode(Variant node, bool deep)
{
    return node.OleFunction("cloneNode", deep);
}

/* ��������� �������
 */
Variant __fastcall OleXml::AddAttributeNode(Variant node, const AnsiString& attributeName, const AnsiString& attributeValue)
{
    Variant attribute = xmlDoc.OleFunction("CreateAttribute", attributeName);
    attribute.OlePropertySet("Value", attributeValue);
    node.OleProcedure("setAttributeNode", attribute);

    return attribute;
}

/* ��������� ����� � ����
   ����� ����� ���� ��������� ��������� � ��������� � ���� �������� ���� ������
 */
void __fastcall OleXml::SetTextNode(Variant node, const AnsiString& nodeText)
{
    //Variant node = AddNode(parentNode, nodeText);
    node.OlePropertySet("text", nodeText);
}

/* ������������� �������� ��������
*/
void __fastcall OleXml::SetAttributeValue(Variant node, const AnsiString& attributeName, const AnsiString& value)
{
    GetAttribute(node, attributeName).OlePropertySet("text", value);
}

/*
   ���������� ��������������
 */
bool __fastcall OleXml::IsTextElement(Variant node) const
{
    return node.OlePropertyGet("IsTextElement");
}

/* ���������, ���� �� � ��������� ���� �������� ����
 */
bool __fastcall OleXml::HasChildNodes(Variant node) const
{
    return node.OleFunction("hasChildNodes");
}

/* ��������� xml-�������� � ����
 */
void __fastcall OleXml::Save(const String& filename) const
{
    xmlDoc.OleProcedure("Save", filename);
}

/* ��������� xml �� ����� ����
 */
void __fastcall OleXml::LoadXMLFile(const AnsiString& XMLFileName)
{
    xmlDoc.OlePropertyGet("Load", XMLFileName.c_str());
    //rootNode = XmlDoc.OlePropertyGet("documentElement");
}

/* ��������� xml �� ������
 */
void __fastcall OleXml::LoadXMLText(const AnsiString& XMLText)
{
    xmlDoc.OlePropertyGet("LoadXML", XMLText.c_str());
}

/* ���������� ���������� ��������� ����
 */
int __fastcall OleXml::GetAttributesCount(Variant node) const
{
    return node.OlePropertyGet("attributes").OlePropertyGet("length");
}

AnsiString __fastcall OleXml::GetNodeText(Variant node) const
{
    return node.OlePropertyGet("text");
}

/* ���������� �������� ����
 */
Variant __fastcall OleXml::GetRootNode() const
{
    return xmlDoc.OlePropertyGet("DocumentElement");
}

/* ���������� ��� ����
 */
AnsiString __fastcall OleXml::GetNodeName(Variant node) const
{
    return node.OlePropertyGet("NodeName");
}

/* ��������� ������ �������� ����
 */
Variant __fastcall OleXml::GetFirstNode(Variant node) const
{
    return node.OlePropertyGet("firstChild");
}

/* ���������� ��������� ���� �� ����������
 */
Variant __fastcall OleXml::GetNextNode(Variant node) const
{
    return node.OlePropertyGet("nextSibling");
}

/* ���������� ���� �� ���� xpath
   ������������� �� ��������� �������
 */
Variant __fastcall OleXml::SelectSingleNode(const AnsiString& xpath) const
{
    return xmlDoc.OleFunction("selectSingleNode", xpath); // selectSingleNode
}

/* ���������� ���� �� ���� xpath
 */
Variant __fastcall OleXml::SelectSingleNode(Variant node, const AnsiString& xpath) const
{
    return node.OleFunction("selectSingleNode", xpath); // selectSingleNode
}


/* ���������, ���������� �� �������
 */
Variant OleXml::GetAttribute(Variant node, const AnsiString& attributeName) const
{
    // test
    //int n = attributeName.Length() + 1;
    // wchar_t* s = new wchar_t[n];
    //StringToWideChar(attributeName, s, n);

    //IXMLDOMNode* pIXMLDOMNode;

    //IXMLDOMNamedNodeMap* attr;
    //attr->getNamedItem()
    //pIXMLDOMNode->get
    //BSTR bstrAtrName = attributeName.WideChar();
    //Variant a = node.OlePropertyGet("attributes");

    //wstring s5(L"type");
    //delete s;
    //Variant result = a.OleFunction("getNamedItem", attributeName.c_str());


    return node.OlePropertyGet("attributes").OleFunction("getNamedItem", attributeName.c_str());
    //return attribute.IsEmpty();
}

/* ���������� ������� �� �������
*/
AnsiString __fastcall OleXml::GetAttributeValue(Variant node, int attributeIndex) const
{
    return node.OlePropertyGet("attributes").OlePropertyGet("item", attributeIndex).OlePropertyGet("Value");
}

/* ���������� �������� �������� �� �����
*/
AnsiString __fastcall OleXml::GetAttributeText(Variant node, const AnsiString& attributeName) const
{
    Variant attribute = GetAttribute(node, attributeName);
    if (!attribute.IsEmpty())
    {
        return attribute.OlePropertyGet("text");
    }
    else
    {
        throw Exception("Attribute " + attributeName + "is empty.");
    }

    // ������ ������
    //return Node.OleFunction("GetAttribute", StringToOleStr(AttributeName));
}

/* ���������� �������� ��������,
   ���� ������� �����������, �� ���������� �������� DefaultValue
 */
AnsiString __fastcall OleXml::GetAttributeValue(Variant node, const AnsiString& attributeName, AnsiString DefaultValue) const
{
    Variant attribute = GetAttribute(node, attributeName);
    return (VarIsClear(attribute)) ? DefaultValue : VarToStr( attribute.OlePropertyGet("text") );

    // commented 2016-11-17
    //return (attribute.IsEmpty()) ? DefaultValue : VarToStr( attribute.OlePropertyGet("text") );



    /*
    AnsiString attribute = Trim( GetAttributeValue( Node, StringToOleStr(AttributeName) ) );

    if (attribute != "")
    {
        return attribute;
    } else
    {
        return DefaultValue;
    }*/
}

/* ���������� �������� ��������,
   ���� ������� �����������, �� ���������� �������� DefaultValue
 */
int __fastcall OleXml::GetAttributeValue(Variant node, const AnsiString& attributeName, int defaultValue) const
{

    Variant attribute = GetAttribute(node, attributeName);
    try
    {
        // ���������������� 2016-11-17
        //return (attribute.IsEmpty()) ? defaultValue : attribute.OlePropertyGet("Value");
        return (VarIsClear(attribute)) ? defaultValue : attribute.OlePropertyGet("Value");
    }
    catch (...)  // in case if attribute value is not integer value
    {
        return defaultValue;
    }

/*    AnsiString attribute = Trim(GetAttributeValue(Node, AttributeName));  // ������ �������
    if (attribute != "")
    {
        try
        {
            return StrToInt(attribute);
        }
        catch (...)
        {
            return DefaultValue;
        }
    }
    else
    {
        return DefaultValue;
    }*/
}

/* ���������� �������� ��������,
   ���� ������� �����������, �� ���������� �������� DefaultValue
 */
bool __fastcall OleXml::GetAttributeValue(Variant node, const AnsiString& attributeName, bool defaultValue) const
{
    Variant attribute = GetAttribute(node, attributeName);

    //if (attribute.IsEmpty()) {
    if ( VarIsClear(attribute) ) {
        return defaultValue;
    } else {
        String textValue = attribute.OlePropertyGet("Value");
        if (textValue == TRUE_STR_VALUE)
        {
            return true;
        }
        else if (textValue == FALSE_STR_VALUE)
        {
            return false;
        }
        else
        {
            return defaultValue;
        }
    }
}

/* ��������� ������� ������ ������� XML
 */
AnsiString __fastcall OleXml::GetParseError() const
{
    if( xmlDoc.OlePropertyGet("parseError").OlePropertyGet("errorCode")!=0 )
    {
        return AnsiString(xmlDoc.OlePropertyGet("parseError").OlePropertyGet("reason"));
    }
    else
    {
        return "";
    }
}








