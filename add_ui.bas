Private Sub UserForm_Initialize()
'nothing
add_ui.input_lb_ingred.Visible = False

End Sub

Private Sub input_lb_ingred_Change()
'combobox change esemény lekezelése, találat alapján adatok betöltése

Dim lookFor As String
Dim data_range As Range
Dim cnl As Integer
Dim original_quantity, kj, kcal, protein, fat, carbohydrate, sugar, salt As Variant

Set a = add_ui

'megszámolni hány tétel van a Nutrition List-en
cnl = WorksheetFunction.CountA(Sheets("ingredient").Columns(1))

If a.input_lb_ingred.Value <> "" Then
    'VLookUp tömbje
    lookFor = a.input_lb_ingred.Value
    drange = "A2:I" & cnl
    Set data_range = Sheets("ingredient").Range(drange)
    
    On Error Resume Next
    'rohadt sok VLookUp az adatok kikereséséhez
    original_quantity = WorksheetFunction.VLookup(lookFor, data_range, 2, 0)
    carbohydrate = WorksheetFunction.VLookup(lookFor, data_range, 3, 0)
    sugar = WorksheetFunction.VLookup(lookFor, data_range, 4, 0)
    protein = WorksheetFunction.VLookup(lookFor, data_range, 5, 0)
    fat = WorksheetFunction.VLookup(lookFor, data_range, 6, 0)
    kj = WorksheetFunction.VLookup(lookFor, data_range, 7, 0)
    kcal = WorksheetFunction.VLookup(lookFor, data_range, 8, 0)
    salt = WorksheetFunction.VLookup(lookFor, data_range, 9, 0)
    
    'mértékegység és mennyiség frissítése
    a.original_data_head = "Per " & original_quantity & " g"
    
    'adatok betöltése UserForm-ra
    a.s_label_ounit.Caption = original_quantity 'csak rejtett mezobe felrakva!
    a.output_label_kj.Caption = kj
    a.output_label_kcal.Caption = kcal
    a.output_label_protein.Caption = protein
    a.output_label_fat.Caption = fat
    a.output_label_carbohydrate.Caption = carbohydrate
    a.output_label_sugar.Caption = sugar
    a.output_label_salt.Caption = salt

    'ha input_unit fel van töltve akkor input_unit_Change-el frissíteni a kalkulált mezőket
    If a.input_unit.Value <> "" Then
        Call input_unit_Change
    End If

End If
End Sub

Private Sub input_tx_preingred_Change()
Dim search As Variant
Dim ingredlist As Range
Dim result(200)
Dim cnl As Integer
Dim drange As Variant
Dim i As Integer


Set a = add_ui
search = a.input_tx_preingred.Value

add_ui.input_lb_ingred.Visible = True

'listbox kiürítése minden karakter után
Do While a.input_lb_ingred.ListCount > 0
    a.input_lb_ingred.RemoveItem (0)
Loop


'megszámolni hány tétel van a ingredient list-en és range-t összeállítani
cnl = WorksheetFunction.CountA(Sheets("ingredient").Columns(1))
drange = "A2:I" & cnl
Set ingredlist = Sheets("ingredient").Range(drange)

'ingredient list-ben keresni a megadott hozzávaló alapján és feltölteni input_lb_ingred listbox-ot
With ingredlist
   
    Set whatfind = .find(search, lookAt:=xlPart)
        
    If Not whatfind Is Nothing Then
        FirstAddress = whatfind.Address
        
        Do
            On Error Resume Next
            result(i) = whatfind.Address
            a.input_lb_ingred.AddItem (whatfind)
            i = i + 1
                        
        Set whatfind = .FindNext(whatfind)
            
        Loop While Not whatfind Is Nothing And whatfind.Address <> FirstAddress
    Else
    add_ui.input_lb_ingred.Visible = False
    End If
        
End With

End Sub
Private Sub input_lb_ingred_Click()
Dim data As Variant

Set a = add_ui
data = a.input_lb_ingred.Value

For i = 0 To a.input_lb_ingred.ListCount - 1
    a.input_lb_ingred.Selected(i) = False
Next i
a.input_tx_preingred.Value = data

a.input_lb_ingred.Visible = False

End Sub


Private Sub input_unit_Change()
'unit kitöltésének lekezelése, számítások

Dim b As Integer
Dim kj As Integer

Set a = add_ui

'eltérő mértékegység miatt b változóba az osztó
    b = a.s_label_ounit.Caption * 1

'input_unit-ba csak szám mehet a kalkuláció miatt, másképp hiba
If IsNumeric(a.input_unit.Value) Then
    'new_data_head frissítése a megadott input_unit és label_unit alapján
    a.new_data_head = "Per " & a.input_unit.Value & " " & a.output_label_unit
    'UserForm-ra betöltöt adatok alapján számítás és _new label-k frissítése
    a.output_label_kj_new.Caption = a.output_label_kj.Caption * (a.input_unit.Value / b)
    a.output_label_kcal_new.Caption = a.output_label_kcal.Caption * (a.input_unit.Value / b)
    a.output_label_protein_new.Caption = a.output_label_protein.Caption * (a.input_unit.Value / b)
    a.output_label_fat_new.Caption = a.output_label_fat.Caption * (a.input_unit.Value / b)
    a.output_label_carbohydrate_new.Caption = a.output_label_carbohydrate.Caption * (a.input_unit.Value / b)
    a.output_label_sugar_new.Caption = a.output_label_sugar.Caption * (a.input_unit.Value / b)
    a.output_label_salt_new.Caption = a.output_label_salt.Caption * (a.input_unit.Value / b)
    
Else
    If a.input_tx_preingred.Value <> "" Then
    'MsgBox-al hibát dobni és tartalmat törölni, ha input_unit-ba nem num kerül
    MsgBox ("Numbers only!")
    a.input_unit.Value = ""
    End If
    
End If

End Sub

Public Sub btn_addtolist_Click()
Dim fcc As Integer
Dim s As Worksheet
Dim ingred, quantity As Variant

Set s = Worksheets("calculator")
Set a = add_ui

'feloldani a sheet zárolását
ActiveSheet.Unprotect Password:="12345"

'csak akkor másolódjon át az adat, ha input_unit ki van töltve és input_cb_ingred ki van választva.
If IsNumeric(a.input_unit.Value) And a.input_tx_preingred.Value <> "" Then

    'az Ingredient lista első szabad helye
    fcc = Range("C17:C" & Rows.Count).Cells.SpecialCells(xlCellTypeBlanks).Row
    
    'valtozoba minden adattal
    ingred = a.input_tx_preingred.Value
    quantity = a.input_unit.Value
    
    'UserForm-ról minden _new kiírása a tábla első üres sorába. caption miatt *1 (!)
    s.Range("B" & fcc) = (fcc - 16)
    s.Range("C" & fcc) = ingred
    s.Range("D" & fcc) = quantity * 1
    
    'realtime sorbarendezni menyniség szerint beszúrás után
    Call Module1.subfunc_sort
    
    'zárolást visszarakni jelszóval
    ActiveSheet.Protect Password:="12345", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
    
    a.input_tx_preingred.Value = ""
    a.input_unit.Value = ""
    Call Module1.uf_clear

Else
    'hibaüzenet ha input_unit nincs töltve vagy input_cb_ingred nincs kiválasztva
    MsgBox ("Choose ingredients and add the quantity!")
End If

End Sub

Private Sub btn_exit_Click()
'exit btn click esemény lekezelése, form elrejtése, unload
add_ui.Hide
Unload Me
End Sub

