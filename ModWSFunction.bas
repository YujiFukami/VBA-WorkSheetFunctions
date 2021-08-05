Attribute VB_Name = "ModWSFunction"
Option Explicit
'ワークシート関数
Function SheetName$()
'入力セルのシート名を出力
'20210726
    Application.Volatile '自動再計算を有効にする
    SheetName = Application.ThisCell.Parent.Name
End Function

Function MojiKugiri$(Target As Range, KugiriMoji$, OutputNum%)
'文字列を分割して指定番号の文字を出力する
'20210726
    
'Target・・・指定セル
'KugiriMoji・・・分割基準の区切り文字
'OutputNum・・・分割して出力する文字の番号
    
    Application.Volatile '自動再計算を有効にする
    Dim TargetStr$
    TargetStr = Target.Value
    MojiKugiri = Split(TargetStr, KugiriMoji)(OutputNum - 1)

End Function
