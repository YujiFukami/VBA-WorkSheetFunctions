Attribute VB_Name = "ModWSFunction"
Option Explicit
'���[�N�V�[�g�֐�
Function SheetName$()
'���̓Z���̃V�[�g�����o��
'20210726
    Application.Volatile '�����Čv�Z��L���ɂ���
    SheetName = Application.ThisCell.Parent.Name
End Function

Function MojiKugiri$(Target As Range, KugiriMoji$, OutputNum%)
'������𕪊����Ďw��ԍ��̕������o�͂���
'20210726
    
'Target�E�E�E�w��Z��
'KugiriMoji�E�E�E������̋�؂蕶��
'OutputNum�E�E�E�������ďo�͂��镶���̔ԍ�
    
    Application.Volatile '�����Čv�Z��L���ɂ���
    Dim TargetStr$
    TargetStr = Target.Value
    MojiKugiri = Split(TargetStr, KugiriMoji)(OutputNum - 1)

End Function
