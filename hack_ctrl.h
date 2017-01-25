#ifndef HACK_CTRL
#define HACK_CTRL

#include <Classes.hpp>
#include <Controls.hpp>

/* ������ ���� Enabled �������� ��������� ����������
*/
void switchEnabledGroupBox(TGroupBox* groupBox)
{
    bool isEnabled = groupBox->Enabled;

    for (int i = 0; i < groupBox->ControlCount; i++)
    {
        groupBox->Controls[i]->Enabled = isEnabled;
    }
}

#endif
