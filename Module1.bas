Attribute VB_Name = "Module1"
Sub Btn_add()
add_ui.Show

End Sub
Sub Btn_admin()
admin_ui.Show
End Sub

Sub Btn_print()
Attribute Btn_print.VB_ProcData.VB_Invoke_Func = " \n14"
Application.ScreenUpdating = False

ThisWorkbook.Sheets("calculator").Activate
    Columns("A:L").Select
    With ActiveSheet.PageSetup
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    Selection.PrintOut Copies:=1, Collate:=True
    Range("D3:I3").Select

'Call Btn_clear

Application.ScreenUpdating = True
End Sub
Sub Btn_clear()
Dim a As Variant
Dim b As Variant
Set a = Worksheets("calculator")
Set b = Worksheets("tmp")

    'calculator sheet kiürítése
    a.Range("D3:I3").ClearContents
    'a.Range("D4:I4").ClearContents
    a.Range("D8:D9").ClearContents
    a.Range("D8:D8").Value = "100"
    a.Range("D9:D9").Value = "10"
    a.Range("B17:D67").ClearContents
    
    'temp sheet kiürítése
    b.Range("B3:L3").ClearContents
    'b.Range("B4:L4").ClearContents
    b.Range("D8:D9").ClearContents
    b.Range("B17:D67").ClearContents
    
End Sub

Sub Btn_save()
'a calculatorrol áttölteni minden adatot a tmp-re és elmenteni új sheetként
Dim a As Variant
Dim b As Variant

Set a = Worksheets("calculator")
Set b = Worksheets("tmp")

Application.ScreenUpdating = False

'mentés előtt mennyiség szerint csükkenőbe minden
Call subfunc_sort

    'copy calculator data to tmp sheet
b.Range("B3").Value = a.Range("D3").Value   'recept neve
'b.Range("B4").Value = a.Range("D4").Value   'cég neve
b.Range("D8").Value = a.Range("D8").Value   'egy szelet tömege
b.Range("D9").Value = a.Range("D9").Value   'anyagveszteség
b.Range("C17:D67").Value = a.Range("C17:D67").Value 'összetevők és mennyiségük

   
    'copy tmp sheet, new name: D3:I3 value
    b.Visible = True
    b.Select
    b.Copy After:=a
    Sheets(a.Index + 1).Name = a.Range("D3")
    b.Visible = False
'meghívni a törlés gomb fv-ét, hogy ki legyen tisztítva minden.
Call Btn_clear
Application.ScreenUpdating = True

End Sub

Sub subfunc_sort()
Dim a As Variant

Set a = Worksheets("calculator")
Application.ScreenUpdating = False
    a.Sort.SortFields.Clear
    a.Sort.SortFields.Add Key:=Range("D17:D67"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With a.Sort
        .SetRange Range("C17:D67")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
Application.ScreenUpdating = True
End Sub
Sub subfunc_ingred_sort()
Dim a As Variant

Set a = Worksheets("ingredient")
Application.ScreenUpdating = False
    a.Sort.SortFields.Clear
    a.Sort.SortFields.Add Key:=Range("A2:A6000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With a.Sort
        .SetRange Range("A2:I6000")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
Application.ScreenUpdating = True
End Sub

Sub uf_clear()

Set a = add_ui
    a.new_data_head.Caption = ""
    a.output_label_kj_new.Caption = ""
    a.output_label_kcal_new.Caption = ""
    a.output_label_protein_new.Caption = ""
    a.output_label_fat_new.Caption = ""
    a.output_label_carbohydrate_new.Caption = ""
    a.output_label_sugar_new.Caption = ""
    a.output_label_salt_new.Caption = ""
    a.output_label_kj.Caption = ""
    a.output_label_kcal.Caption = ""
    a.output_label_protein.Caption = ""
    a.output_label_fat.Caption = fat
    a.output_label_carbohydrate.Caption = ""
    a.output_label_sugar.Caption = ""
    a.output_label_salt.Caption = ""

End Sub


