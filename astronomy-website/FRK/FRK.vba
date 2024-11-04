Option Explicit

Public Const Array_Size = 40

Public Const Table_Start_Line As Long = 3
Public Const Table_Row_Offset As Long = 1

' In Class Module named clsCabinet
Option Explicit

Public Typical As Worksheet

Public File_Names As String
Public Typical_Names As String
Public Sie_Werk_Nr As String

Public Empty_Cabinet As Boolean
Public Cabinet_Processing As Boolean
Public Wide_Cabinet As Boolean

Public Cabinet_Multipiler As Long

Public Total_Pages As Long

Public First_Page_BB As Long
Public Last_Page_BB As Long

Public First_Page_BT As Long
Public Last_Page_BT As Long

Public First_Page_NB As Long
Public Last_Page_NB As Long

Public First_Page_NT As Long
Public Last_Page_NT As Long

Public First_Row_BB As Long
Public Last_Row_BB As Long
Public Total_Row_BB As Long
Public Wire_Amount_BB As Long

Type Cabinets
    
    Typical As Worksheet
    
    File_Names As String
    Typical_Names As String
    Sie_Werk_Nr As String
    
    Empty_Cabinet As Boolean
    Cabinet_Processing As Boolean
    Wide_Cabinet As Boolean
    
    Cabinet_Multipiler  As Integer
    
    Total_Pages As Integer
    
    First_Page_BB As Integer
    Last_Page_BB As Integer
    
    First_Page_BT As Integer
    Last_Page_BT As Integer
    
    First_Page_NB As Integer
    Last_Page_NB As Integer
    
    First_Page_NT As Integer
    Last_Page_NT As Integer
    
    First_Row_BB As Integer
    Last_Row_BB As Integer
    Total_Row_BB As Integer
    Wire_Amount_BB As Integer
    
    First_Row_BT As Integer
    Last_Row_BT As Integer
    Total_Row_BT As Integer
    Wire_Amount_BT As Integer
    
    First_Row_NB As Integer
    Last_Row_NB As Integer
    Total_Row_NB As Integer
    Wire_Amount_NB As Integer
    
    First_Row_NT As Integer
    Last_Row_NT As Integer
    Total_Row_NT As Integer
    Wire_Amount_NT As Integer
    
    Wire_Amount_UEG As Integer
    
    Total_TRIP_Labels As Integer
    Total_POTENTIAL_Labels As Integer
    Total_WireSize_Two_Dot_Five As Integer
    Total_WireSize_Over_Two_Dot_Five As Integer

End Type

' Globale Variables
Public ST_Files_Count As Integer
Public DataStruct(0 To Array_Size) As Cabinets
Public Without_Interconnection_Table As Boolean
Public PDF_Import_Done As Boolean
Public Table1_Page_Name As Integer
Public Cabinet_Name_New As String
Public Cabinet_Name_Old As String
Public Text_Line1() As String
Public Text_Line2(0 To 35) As String
Public Line_Number As Integer
Public Table_BB As Boolean
Public Table_BT As Boolean
Public Table_NB As Boolean
Public Table_NT As Boolean
Public Table_FF As Boolean
Public Table_InterLoops As Boolean
Public Table As Range
Public Table_Buffer() As Variant
Public Table_Row_Index As Long

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Sub Storing_Variables(Cabinet_Amount As Integer)
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Dim i As Integer
Dim Variable_Array_Length As Integer
Dim Variable_Array_First_Item_Pos As Integer
Dim Variable_Pointer As Integer

Variable_Array_Length = Variables.Range("C2").Value
Variable_Array_First_Item_Pos = Variables.Range("C3").Value
Variable_Pointer = Variables.Range("C3").Value

    
    Variables.Range("C5").Value = ST_Files_Count
    Variables.Range("E2").Value = Heat_Shrink_Tubes
    Variables.Range("C8").Value = Main_WireType_Name
    
For i = 0 To Cabinet_Amount Step 1
    
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).File_Names
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Typical_Names
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Sie_Werk_Nr
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Empty_Cabinet
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Cabinet_Processing
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Wide_Cabinet
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Cabinet_Multipiler
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Total_Pages
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).First_Page_BB
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Last_Page_BB
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).First_Page_BT
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Last_Page_BT
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).First_Page_NB
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Last_Page_NB
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).First_Page_NT
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Last_Page_NT
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).First_Row_BB
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Last_Row_BB
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Total_Row_BB
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Wire_Amount_BB
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).First_Row_BT
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Last_Row_BT
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Total_Row_BT
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Wire_Amount_BT
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).First_Row_NB
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Last_Row_NB
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Total_Row_NB
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Wire_Amount_NB
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).First_Row_NT
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Last_Row_NT
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Total_Row_NT
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Wire_Amount_NT
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Wire_Amount_UEG
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Total_TRIP_Labels
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Total_POTENTIAL_Labels
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Total_WireSize_Two_Dot_Five
    Variable_Pointer = Variable_Pointer + 1
    Variables.Cells(Variable_Pointer, 3).Value = DataStruct(i).Total_WireSize_Over_Two_Dot_Five
    Variable_Pointer = Variable_Pointer + 1

Next i
    
    ThisWorkbook.Save

End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Sub Loading_Variables()
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Dim i As Integer, k As Integer
Dim File_Ctn As Integer
Dim Variable_Array_Length As Integer
Dim Variable_Array_First_Item_Pos As Integer
Dim Variable_Pointer As Integer


Variable_Array_Length = Variables.Range("C2").Value
Variable_Array_First_Item_Pos = Variables.Range("C3").Value
File_Ctn = Variables.Range("C5").Value
Variable_Pointer = Variable_Array_First_Item_Pos

ST_Files_Count = Variables.Range("C5").Value
Heat_Shrink_Tubes = Variables.Range("E2").Value
Call Refresh_Heat_Shrink_Tubes_Display_State

    '----------------------------
    
        If (Variables.Range("D5").Value = False) Then       'Without_Interconnection_Table = False
            k = 0
        ElseIf (Variables.Range("D5").Value = True) Then    'Without_Interconnection_Table = True
            k = 1
        End If
        
    '----------------------------
    
        For i = 4 To Worksheets.Count Step 1
            Set DataStruct(k).Typical = Worksheets(i)
            k = k + 1
        Next i
    
    '----------------------------

For i = 0 To File_Ctn Step 1

    DataStruct(i).File_Names = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Typical_Names = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Sie_Werk_Nr = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Empty_Cabinet = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Cabinet_Processing = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Wide_Cabinet = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Cabinet_Multipiler = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Total_Pages = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).First_Page_BB = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Last_Page_BB = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).First_Page_BT = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Last_Page_BT = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).First_Page_NB = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Last_Page_NB = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).First_Page_NT = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Last_Page_NT = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).First_Row_BB = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Last_Row_BB = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Total_Row_BB = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Wire_Amount_BB = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).First_Row_BT = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Last_Row_BT = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Total_Row_BT = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Wire_Amount_BT = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).First_Row_NB = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Last_Row_NB = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Total_Row_NB = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Wire_Amount_NB = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).First_Row_NT = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Last_Row_NT = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Total_Row_NT = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Wire_Amount_NT = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Wire_Amount_UEG = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Total_TRIP_Labels = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Total_POTENTIAL_Labels = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Total_WireSize_Two_Dot_Five = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
    DataStruct(i).Total_WireSize_Over_Two_Dot_Five = Variables.Cells(Variable_Pointer, 3).Value
        Variable_Pointer = Variable_Pointer + 1
Next i
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Sub Init_Variables(Cabinet_Amount As Integer)
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Dim i As Integer

        For i = 0 To Cabinet_Amount
    
        DataStruct(i).File_Names = ""
        DataStruct(i).Typical_Names = ""
        DataStruct(i).Sie_Werk_Nr = ""
        DataStruct(i).Empty_Cabinet = False
        DataStruct(i).Cabinet_Processing = True
        DataStruct(i).Wide_Cabinet = False
        
        DataStruct(i).Cabinet_Multipiler = 0
        
        DataStruct(i).First_Page_BB = 0
        DataStruct(i).Last_Page_BB = 0
        DataStruct(i).First_Page_BT = 0
        DataStruct(i).Last_Page_BT = 0
        DataStruct(i).First_Page_NB = 0
        DataStruct(i).Last_Page_NB = 0
        DataStruct(i).First_Page_NT = 0
        DataStruct(i).Last_Page_NT = 0
        DataStruct(i).Total_Pages = 0
        
        DataStruct(i).First_Row_BB = 0
        DataStruct(i).Last_Row_BB = 0
        DataStruct(i).Total_Row_BB = 0
        DataStruct(i).Wire_Amount_BB = 0
        
        DataStruct(i).First_Row_BT = 0
        DataStruct(i).Last_Row_BT = 0
        DataStruct(i).Total_Row_BT = 0
        DataStruct(i).Wire_Amount_BT = 0
        
        DataStruct(i).First_Row_NB = 0
        DataStruct(i).Last_Row_NB = 0
        DataStruct(i).Total_Row_NB = 0
        DataStruct(i).Wire_Amount_NB = 0
        
        DataStruct(i).First_Row_NT = 0
        DataStruct(i).Last_Row_NT = 0
        DataStruct(i).Total_Row_NT = 0
        DataStruct(i).Wire_Amount_NT = 0
        DataStruct(i).Wire_Amount_UEG = 0
        
        DataStruct(i).Total_TRIP_Labels = 0
        DataStruct(i).Total_POTENTIAL_Labels = 0
        DataStruct(i).Total_WireSize_Two_Dot_Five = 0
        DataStruct(i).Total_WireSize_Over_Two_Dot_Five = 0
        
        Next i
        
        Variables.Range("C6").Value = 0
        Variables.Range("C7").Value = 0
        Variables.Range("C8").Value = ""
        PDF_Import_Done = False
        Call Storing_Variables(Cabinet_Amount)
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Sub Refresh_Cabinet_List()
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Dim Clear_All_Data As Boolean

Clear_All_Data = True
Init_Variables (ST_Files_Count)
Read_Names_Of_PDF_Documents (Clear_All_Data)

End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Sub Read_Names_Of_PDF_Documents(Clear_All_Data As Boolean)
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Dim offset As Integer

Application.ScreenUpdating = False

    offset = ST_Files_Count + 1

    If (Clear_All_Data = True) Then
        Call Delete_Check_Boxses
        Main_Sheet.Range("B11", "V" & offset + 11).Cells.Clear
    Else
        Main_Sheet.Range("B11", "Q" & offset + 11).Cells.Clear
    End If

Call Delete_Sheets
Call LoopThroughFiles
Call Loading_Value_Of_CheckBoxses_To_Memory
Call Storing_Variables(ST_Files_Count)
Call Data_To_Display


End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'PDF Fileok helyének kivállasztása
Sub GetPDFFolder()
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Dim i As Long
Dim PDFFldr As FileDialog
Dim Clear_All_Data As Boolean
Dim Error_Status As Byte
Dim File_Path_Temp As String
        
        Set PDFFldr = Application.FileDialog(msoFileDialogFolderPicker)
        With PDFFldr
            .Title = "Select a Folder"
            If .Show <> -1 Then GoTo End_Of_GetPDFFolder:
            File_Path_Temp = .SelectedItems(1)
            Main_Sheet.Range("D3").Value = File_Path_Temp
NoSel:
        End With
        
        Call Init_Variables(Array_Size)
       
        Clear_All_Data = True
      
        Distribution_Of_Looping_Wires_Is_OK = False
        Distribution_Cabinet_Wires_Is_OK = False
        
        PDF_Import_Done = False
        Call Read_Names_Of_PDF_Documents(Clear_All_Data)
        Error_Status = Loading_Processing_Parameters(File_Path_Temp, 1)
        
        If (Error_Status <> 0) Then
                
            'Looad Default settings
            
                Wire_Type_Name = 1
                Call Refresh_Wire_Typ_Value
                Heat_Shrink_Tubes = 1
                Call Refresh_Heat_Shrink_Tubes_Display_State
                Value_Of_Cabinet_Height = 1
                Call Refresh_Cabinet_Height_Value                           'Výška skrine
               
                Main_Sheet.Range("M7").Value = ""                                   'Èíslo zákazky : FEAG SLK
                Main_Sheet.Range("I5").Value = 0
                Main_Sheet.Range("Q5").Value = 1                                    'Vyhotovenie zákazky
                Variables.Range("C6").Value = 0                                     'Celkový poèet Vodièov :
                Variables.Range("C7").Value = 0                                     'Celkový poèet Prepojov :
                
                Selection_Sheet.Range("N1").Value = 28                              'Max. poèet Kontaktov: TRIPPING RELAY
                Selection_Sheet.Range("D7").Value = False                           'Set "Prinúti vodièov"
                Selection_Sheet.Range("D9").Value = 2                               'Set "Typ vodièov"
                
                Selection_Sheet.Range("D13").Value = 2                              'Set "Triedenie Vodièov"
                
                Selection_Sheet.Range("D31").Value = False                          'Set "Spoèítanie "Potenciáové" Štítkov"
                Selection_Sheet.Range("F31").Value = False                          'Set "Potenciáové" Štítky v Kommentare"
                
                Selection_Sheet.Range("C33").Value = False                          'Set "Štítky špeciálné"
                Selection_Sheet.Range("F33").Value = True                           'Set "Iba jednostranné Oznaèenie"
                Selection_Sheet.Range("C34").Value = ""                             'Set " Pozície prístrojov -->"
                
                Selection_Sheet.Range("C36").Value = False                          'Set "Nezapojené prístroje "
                Selection_Sheet.Range("C37").Value = ""                             'Set "Pozície nezap.prístrojov -->"
                
                Selection_Sheet.Range("B40").Value = 1                              'Set "Tabu¾ky"
                Selection_Sheet.Range("D44").Value = 1                              'Set "Smer zapojenia UEG :"
                
                Wide_Cabinet_Exists = False
                For i = 1 To ST_Files_Count Step 1
                
                    DataStruct(i).Wide_Cabinet = False
                    If (DataStruct(i).Empty_Cabinet = False) Then
                        DataStruct(i).Cabinet_Processing = True
                    ElseIf (DataStruct(i).Empty_Cabinet = True) Then
                        DataStruct(i).Cabinet_Processing = False
                    End If
                
                Next i
           
        End If
        
        '-------------------------------
                If (Without_Interconnection_Table = False) Then
                    DataStruct(0).Cabinet_Processing = True
                ElseIf (Without_Interconnection_Table = True) Then
                    DataStruct(0).Cabinet_Processing = False
                End If
        '-------------------------------
            
            Storing_Variables (ST_Files_Count)
            Main_Sheet.Activate
            Call Data_To_Display
End_Of_GetPDFFolder:
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Sub LoopThroughFiles()
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Dim oFSO As Object
Dim oFolder As Object
Dim oFile As Object
Dim i As Long
Dim Text1_Pos As Long
Dim Text2_Pos As Long
Dim Text3_Pos As Long
Dim Text4_Pos As Long

Dim Temp_File_Name As String
Dim Temp_File_Path As String
Dim Typical_Name_Temp As String
Dim FilePath As String


Dim MyFSO As FileSystemObject


    Set MyFSO = New FileSystemObject
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    FilePath = (Main_Sheet.Range("D3") & "\výrobná dokumentácia\19_Fertigung")
    Set oFolder = oFSO.GetFolder(FilePath)
    
    Variables.Range("D5").Value = True
    Without_Interconnection_Table = True
    ST_Files_Count = 0
    i = 1
    
    
    For Each oFile In oFolder.Files
    
        If (InStr(1, UCase(oFile.Name), "00_VERB", vbBinaryCompare) > 0) Then
            Without_Interconnection_Table = False
            Variables.Range("D5").Value = False
        End If
    
    Next oFile
    
    
    For Each oFile In oFolder.Files
    
    If ((InStr(1, UCase(oFile.Name), "_ST", vbBinaryCompare) = 0) And (InStr(1, UCase(oFile.Name), "00_VERB", vbBinaryCompare) = 0)) Then
    GoTo Skip_File
    End If
    
    Typical_Name_Temp = oFile.Name
    
    If (InStr(1, Typical_Name_Temp, "00_VERB", [0]) > 0) Then
    Text1_Pos = InStr(1, Typical_Name_Temp, "_VERB", [0])
    DataStruct(0).Typical_Names = Left(Typical_Name_Temp, Text1_Pos - 1)
    DataStruct(0).File_Names = DataStruct(0).Typical_Names & "_VERB"
    End If
    
    If (InStr(1, Typical_Name_Temp, "_ST", [0]) > 0) Then
    Text1_Pos = InStr(1, Typical_Name_Temp, "_ST", [0])
    DataStruct(i).Typical_Names = Left(Typical_Name_Temp, Text1_Pos - 1)
    i = i + 1
    ST_Files_Count = ST_Files_Count + 1
    End If

Skip_File:
Next oFile

Call Sorting_CabinetNames

For i = 1 To ST_Files_Count
    
    Temp_File_Name = DataStruct(i).Typical_Names & "_VERB.pdf"
    Temp_File_Path = (FilePath & "\" & Temp_File_Name)
    
If MyFSO.FileExists(Temp_File_Path) Then
        DataStruct(i).Empty_Cabinet = False
        DataStruct(i).File_Names = DataStruct(i).Typical_Names & "_VERB"
    If (Main_Sheet.Range("D4") = True) Then
        Text1_Pos = InStr(1, DataStruct(i).Typical_Names, "_=", [0])
        If (Text1_Pos > 0) Then
        DataStruct(i).Typical_Names = Left(DataStruct(i).Typical_Names, Text1_Pos - 1)
        End If
    End If
Else
    DataStruct(i).Empty_Cabinet = True
    DataStruct(i).File_Names = " ----- "
    If (Main_Sheet.Range("D4") = True) Then
        Text1_Pos = InStr(1, DataStruct(i).Typical_Names, "_=", [0])
        If (Text1_Pos > 0) Then
        DataStruct(i).Typical_Names = Left(DataStruct(i).Typical_Names, Text1_Pos - 1)
        End If
    End If
End If
Next i
    
    If (Main_Sheet.Range("D4") = True) Then
    Call Grouping_CabinetNames
    End If
    
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Sub Sorting_CabinetNames()
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Dim Loop1 As Long
Dim Loop2 As Long
Dim Result As Integer
Dim Str_Temp1 As String
Dim Str_Temp2 As String
    
            'Returns
            'If string1 is equal to string2, the StrComp function will return 0.
            'If string1 is less than string2, the StrComp function will return -1.
            'If string1 is greater than string2, the StrComp function will return 1.
            'If either string1 or string2 is NULL, the StrComp function will return NULL.
          
            'VBA Constant    Value   Explanation
            'vbUseCompareOption  -1  Uses option compare
            'vbBinaryCompare 0   Binary comparison
            'vbTextCompare   1   Textual comparison
    
    
    For Loop1 = 1 To ST_Files_Count
        For Loop2 = Loop1 To ST_Files_Count
            Result = StrComp(UCase(DataStruct(Loop2).Typical_Names), UCase(DataStruct(Loop1).Typical_Names), vbTextCompare)
            If Result < 0 Then
                Str_Temp1 = DataStruct(Loop1).Typical_Names
                Str_Temp2 = DataStruct(Loop2).Typical_Names
                DataStruct(Loop1).Typical_Names = Str_Temp2
                DataStruct(Loop2).Typical_Names = Str_Temp1
            End If
        Next Loop2
    Next Loop1
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Sub Grouping_CabinetNames()
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Dim DataStruct_Temp(0 To Array_Size) As Cabinets
Dim Typical_Name_Temp As String
Dim Text_Pos As Integer
Dim i As Integer
Dim Ctn As Integer
Dim Result As Integer

Ctn = 0
For i = Ctn To ST_Files_Count
DataStruct_Temp(i) = DataStruct(i)
Next i
    Erase DataStruct()
    DataStruct(0) = DataStruct_Temp(0)
Ctn = 1
    For i = Ctn To ST_Files_Count

        Result = StrComp(UCase(DataStruct_Temp(i).Typical_Names), UCase(DataStruct_Temp(i - 1).Typical_Names), vbTextCompare)
        If Result <> 0 Then
        DataStruct(Ctn) = DataStruct_Temp(i)
        Ctn = Ctn + 1
    End If
Next i
ST_Files_Count = Ctn - 1
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Sub Get_Name_Of_Page(Page_Buffer As Variant, File_Index As Integer, Page_Ctn As Integer)
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Dim BBA, BBD, BTA, BTD, NBA, NBD, NTA, NTD, FBA, FFA, NFA, Z00 As Integer
Dim i As Integer
Dim j As Integer
Dim Z00_String As String
Dim Text_Temp As Integer
Dim Text_Temp2() As String

' Oldal név duplikációk kiszûrésére kell írni programrészt !!!!!!!

 
BBA = 0
BBD = 0
BTA = 0
BTD = 0
NBA = 0
NBD = 0
NTA = 0
NTD = 0
FBA = 0
FFA = 0
NFA = 0
Z00 = 0
    
    'vbBinaryCompare
    For i = 0 To UBound(Page_Buffer)
    
        
        If (InStr(1, DataStruct(File_Index).Typical_Names, "Z00", vbBinaryCompare)) < 1 Then
        
            Text_Temp = InStr(1, Page_Buffer(i), "VDT", vbBinaryCompare)     ' Temporary
            If (Text_Temp > 0) Then GoTo Skip_To_Next_Line
            
                BBA = InStr(1, Page_Buffer(i), "BBA", vbBinaryCompare)
                BBD = InStr(1, Page_Buffer(i), "BBD", vbBinaryCompare)
                BTA = InStr(1, Page_Buffer(i), "BTA", vbBinaryCompare)
                BTD = InStr(1, Page_Buffer(i), "BTD", vbBinaryCompare)
                NBA = InStr(1, Page_Buffer(i), "NBA", vbBinaryCompare)
                NBD = InStr(1, Page_Buffer(i), "NBD", vbBinaryCompare)
                NTA = InStr(1, Page_Buffer(i), "NTA", vbBinaryCompare)
                NTD = InStr(1, Page_Buffer(i), "NTD", vbBinaryCompare)
                FBA = InStr(1, Page_Buffer(i), "FBA", vbBinaryCompare)
                FFA = InStr(1, Page_Buffer(i), "FFA", vbBinaryCompare)
                NFA = InStr(1, Page_Buffer(i), "NFA", vbBinaryCompare)
        
        Else
'            Z00_String = "A" & CStr(Page_Ctn)
            Z00_String = "A"
            Z00 = InStr(1, Page_Buffer(i), Z00_String, vbBinaryCompare)
        End If
            
            
            If ((BBA > 0) Or (BBD > 0) Or _
                (BTA > 0) Or (BTD > 0) Or _
                (NBA > 0) Or (NBD > 0) Or _
                (NTA > 0) Or (NTD > 0) Or _
                (FBA > 0) Or (FFA > 0) Or _
                (NFA > 0) Or (Z00 > 0)) Then
            
                Exit For
            
            End If
    
Skip_To_Next_Line:
        
        Next i
        
         Table1_Page_Name = 0
        
        If (BBA > 0) Then
            Table1_Page_Name = 1
        ElseIf (BBD > 0) Then
            Table1_Page_Name = 2
        ElseIf (BTA > 0) Then
            Table1_Page_Name = 3
        ElseIf (BTD > 0) Then
            Table1_Page_Name = 4
        ElseIf (NBA > 0) Then
            Table1_Page_Name = 5
        ElseIf (NBD > 0) Then
            Table1_Page_Name = 6
        ElseIf (NTA > 0) Then
            Table1_Page_Name = 7
        ElseIf (NTD > 0) Then
            Table1_Page_Name = 8
        ElseIf (FBA > 0) Then
            Table1_Page_Name = 9
        ElseIf (FFA > 0) Then
            Table1_Page_Name = 10
        ElseIf (NFA > 0) Then
            Table1_Page_Name = 11
        ElseIf (Z00 > 0) Then
            Table1_Page_Name = 12
        Else
            Table1_Page_Name = 0
        End If
        
        If ((Table1_Page_Name = 1) Or (Table1_Page_Name = 2)) Then
            Table_BB = True
            Table_BT = False
            Table_NB = False
            Table_NT = False
            Table_FF = False
            Table_InterLoops = False
        ElseIf ((Table1_Page_Name = 3) Or (Table1_Page_Name = 4)) Then
            Table_BB = False
            Table_BT = True
            Table_NB = False
            Table_NT = False
            Table_FF = False
            Table_InterLoops = False
        ElseIf ((Table1_Page_Name = 5) Or (Table1_Page_Name = 6)) Then
            Table_BB = False
            Table_BT = False
            Table_NB = True
            Table_NT = False
            Table_FF = False
            Table_InterLoops = False
        ElseIf ((Table1_Page_Name = 7) Or (Table1_Page_Name = 8)) Then
            Table_BB = False
            Table_BT = False
            Table_NB = False
            Table_NT = True
            Table_FF = False
            Table_InterLoops = False
        ElseIf ((Table1_Page_Name = 9) Or (Table1_Page_Name = 10) Or (Table1_Page_Name = 11)) Then
            Table_BB = False
            Table_BT = False
            Table_NB = False
            Table_NT = False
            Table_FF = True
            Table_InterLoops = False
        ElseIf (Table1_Page_Name = 12) Then
            Table_BB = False
            Table_BT = False
            Table_NB = False
            Table_NT = False
            Table_FF = False
            Table_InterLoops = True
        Else
            Table_BB = False
            Table_BT = False
            Table_NB = False
            Table_NT = False
            Table_FF = False
            Table_InterLoops = False
        End If
                    
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Sub Get_Name_Of_Page_v2(Page_Buffer As Variant, File_Index As Integer, Page_Ctn As Integer)
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Dim BBA, BBD, BTA, BTD, NBA, NBD, NTA, NTD, FBA, FFA, NFA, Z00 As Integer
Dim i As Integer
Dim j As Integer
Dim Z00_String As String
Dim Text_Temp As Integer
Dim Text_Temp2() As String

' Oldal név duplikációk kiszûrésére kell írni programrészt !!!!!!!

 
BBA = 0
BBD = 0
BTA = 0
BTD = 0
NBA = 0
NBD = 0
NTA = 0
NTD = 0
FBA = 0
FFA = 0
NFA = 0
Z00 = 0
    
    'vbBinaryCompare
    For i = 0 To UBound(Page_Buffer)
    
        
'        If (InStr(1, DataStruct(File_Index).Typical_Names, "Z00", vbBinaryCompare)) < 1 Then
        If (InStr(1, DataStruct(File_Index).Typical_Names, "=J00", vbBinaryCompare)) < 1 Then
'        If (InStr(1, DataStruct(File_Index).Typical_Names, "_VERB", vbBinaryCompare)) < 1 Then
        
            Text_Temp = InStr(1, Page_Buffer(i), "VDT", vbBinaryCompare)     ' Temporary
            If (Text_Temp > 0) Then GoTo Skip_To_Next_Line
            
                BBA = InStr(1, Page_Buffer(i), "BB", vbBinaryCompare)
'                BBD = InStr(1, Page_Buffer(i), "BBD", vbBinaryCompare)
                BTA = InStr(1, Page_Buffer(i), "TB_BT_TT", vbBinaryCompare)
'                BTD = InStr(1, Page_Buffer(i), "BTD", vbBinaryCompare)
'                NBA = InStr(1, Page_Buffer(i), "NBA", vbBinaryCompare)
'                NBD = InStr(1, Page_Buffer(i), "NBD", vbBinaryCompare)
'                NTA = InStr(1, Page_Buffer(i), "NTA", vbBinaryCompare)
'                NTD = InStr(1, Page_Buffer(i), "NTD", vbBinaryCompare)
                FBA = InStr(1, Page_Buffer(i), "BF_FB", vbBinaryCompare)
'                FFA = InStr(1, Page_Buffer(i), "FFA", vbBinaryCompare)
'                NFA = InStr(1, Page_Buffer(i), "NFA", vbBinaryCompare)
        
        Else
'            Z00_String = "A" & CStr(Page_Ctn)
            Z00_String = "A"
            Z00 = InStr(1, Page_Buffer(i), Z00_String, vbBinaryCompare)
        End If
            
            
            If ((BBA > 0) Or (BBD > 0) Or _
                (BTA > 0) Or (BTD > 0) Or _
                (NBA > 0) Or (NBD > 0) Or _
                (NTA > 0) Or (NTD > 0) Or _
                (FBA > 0) Or (FFA > 0) Or _
                (NFA > 0) Or (Z00 > 0)) Then
            
                Exit For
            
            End If
    
Skip_To_Next_Line:
        
        Next i
        
         Table1_Page_Name = 0
        
        If (BBA > 0) Then
            Table1_Page_Name = 1
        ElseIf (BBD > 0) Then
            Table1_Page_Name = 2
        ElseIf (BTA > 0) Then
            Table1_Page_Name = 3
        ElseIf (BTD > 0) Then
            Table1_Page_Name = 4
        ElseIf (NBA > 0) Then
            Table1_Page_Name = 5
        ElseIf (NBD > 0) Then
            Table1_Page_Name = 6
        ElseIf (NTA > 0) Then
            Table1_Page_Name = 7
        ElseIf (NTD > 0) Then
            Table1_Page_Name = 8
        ElseIf (FBA > 0) Then
            Table1_Page_Name = 9
        ElseIf (FFA > 0) Then
            Table1_Page_Name = 10
        ElseIf (NFA > 0) Then
            Table1_Page_Name = 11
        ElseIf (Z00 > 0) Then
            Table1_Page_Name = 12
        Else
            Table1_Page_Name = 0
        End If
        
        If ((Table1_Page_Name = 1) Or (Table1_Page_Name = 2)) Then
            Table_BB = True
            Table_BT = False
            Table_NB = False
            Table_NT = False
            Table_FF = False
            Table_InterLoops = False
        ElseIf ((Table1_Page_Name = 3) Or (Table1_Page_Name = 4)) Then
            Table_BB = False
            Table_BT = True
            Table_NB = False
            Table_NT = False
            Table_FF = False
            Table_InterLoops = False
        ElseIf ((Table1_Page_Name = 5) Or (Table1_Page_Name = 6)) Then
            Table_BB = False
            Table_BT = False
            Table_NB = True
            Table_NT = False
            Table_FF = False
            Table_InterLoops = False
        ElseIf ((Table1_Page_Name = 7) Or (Table1_Page_Name = 8)) Then
            Table_BB = False
            Table_BT = False
            Table_NB = False
            Table_NT = True
            Table_FF = False
            Table_InterLoops = False
        ElseIf ((Table1_Page_Name = 9) Or (Table1_Page_Name = 10) Or (Table1_Page_Name = 11)) Then
            Table_BB = False
            Table_BT = False
            Table_NB = False
            Table_NT = False
            Table_FF = True
            Table_InterLoops = False
        ElseIf (Table1_Page_Name = 12) Then
            Table_BB = False
            Table_BT = False
            Table_NB = False
            Table_NT = False
            Table_FF = False
            Table_InterLoops = True
        Else
            Table_BB = False
            Table_BT = False
            Table_NB = False
            Table_NT = False
            Table_FF = False
            Table_InterLoops = False
        End If
                    
End Sub



'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Function Get_Project_Number(Page_Buffer As Variant, Check_String As String) As String
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Dim Project_Number As String
Dim i As Integer
Dim Werk As Integer
    
    'SWF 952876
    For i = 0 To UBound(Page_Buffer)
        
        If (Check_String = "Werk- Nr.") Then
            Werk = InStr(1, Page_Buffer(i), Check_String, [0])
            If (Werk >= 1) Then
            Project_Number = Right(Page_Buffer(i), (Len(Page_Buffer(i)) - InStr(1, Page_Buffer(i), ". ", [0])) - 1)
            Exit For
            Else
            Project_Number = " ??? "
            End If
        End If
        


        If (Check_String = "IfdNr") Then
            Werk = InStr(1, Page_Buffer(i), Check_String, [0])
            If (Werk >= 1) Then
'            Project_Number = Right(Page_Buffer(i), (Len(Page_Buffer(i)) - InStr(1, Page_Buffer(i), ": ", [0])))
            Project_Number = Page_Buffer(i - 1)
            Exit For
            Else
            Project_Number = " ??? "
            End If
        End If
              
        
        If (Check_String = "SWF ") Then
            Werk = InStr(1, Page_Buffer(i), Check_String, [0])
            If (Werk >= 1) Then
            Project_Number = Right(Page_Buffer(i), (Len(Page_Buffer(i)) - InStr(1, Page_Buffer(i), " ", [0])) - 1)
            Exit For
            Else
            Project_Number = " ??? "
            End If
        End If
    Next i
            Get_Project_Number = Project_Number
    
End Function

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Sub Read_Contents_Of_PDF_Sheets()
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Dim offset As Integer
Dim Clear_All_Data As Boolean

    Clear_All_Data = False
    Call Read_Names_Of_PDF_Documents(Clear_All_Data)

    Application.Wait (Now + TimeValue("0:00:1"))
    Call Import_PDF_Documents
    
    PDF_Import_Done = True
    Distribution_Of_Looping_Wires_Is_OK = False
    Distribution_Cabinet_Wires_Is_OK = False


End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Sub Import_PDF_Documents()
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Dim File_Path As String
Dim File_Path_Temp As String
Dim i As Integer
Dim Exists As Boolean
Dim WS_Temp_Sheet As Worksheet
Dim myRange As Range

    Application.ScreenUpdating = False

'    Exists = False
'    For i = 1 To Worksheets.Count
'            If Worksheets(i).Name = "Temp_Sheet" Then
'            Exists = True
'            End If
'            Next i
'    If Not Exists Then
'    Set WS_Temp_Sheet = Worksheets.Add(, Worksheets(Sheets.Count))
'        WS_Temp_Sheet.Name = "Temp_Sheet"
'        WS_Temp_Sheet.Visible = False
'    End If
        
    
    If (Main_Sheet.Range("M7").Value = "") Then
    MsgBox "Zadaj èíslo zakazku !!!", vbExclamation
    Main_Sheet.Range("M7").Select
    GoTo Cancel_Import_PDF_Documents
    End If
    
    
    
    File_Path_Temp = (Main_Sheet.Range("D3") & "\výrobná dokumentácia\19_Fertigung")
    File_Path = File_Path_Temp & "\" & DataStruct(0).File_Names & ".pdf"
    
    
        If (Without_Interconnection_Table = False) Then
        
            Cabinet_Name_New = DataStruct(0).Typical_Names
            Call Create_Sheets(0)
            Cabinet_Name_Old = Cabinet_Name_New
    
        
            Worksheets(DataStruct(0).Typical_Names).Activate
                With ActiveSheet
                .Rows("1:1").Select
                Selection.RowHeight = 25
                    With ActiveWindow
                    .SplitColumn = 0
                    .SplitRow = 2
                    .FreezePanes = True
                    End With
                    .Range("A1", "O1").Interior.Color = RGB(197, 217, 241)
                    .Range("A1", "O1").VerticalAlignment = xlCenter
                    .Range("A1", "O1").HorizontalAlignment = xlCenter
                    .Range("A1", "O1").Font.Name = "Arial Narrow"
                    .Range("A1", "O1").Font.FontStyle = "Bold Italic"
                    .Range("A1", "O1").Font.Size = 14
                    .Range("A1", "A1").ColumnWidth = 5
                    .Range("B1", "B1").ColumnWidth = 10
                    .Range("C1", "C1").ColumnWidth = 12
                    .Range("D1", "D1").ColumnWidth = 10
                    .Range("E1", "E1").ColumnWidth = 20
                    .Range("F1", "F1").ColumnWidth = 10
                    .Range("G1", "G1").ColumnWidth = 12
                    .Range("H1", "H1").ColumnWidth = 10
                    .Range("I1", "I1").ColumnWidth = 20
                    .Range("J1", "J1").ColumnWidth = 12
                    .Range("K1", "K1").ColumnWidth = 20
                    .Range("L1", "L1").ColumnWidth = 10
                    .Range("M1", "M1").ColumnWidth = 10
                    .Range("N1").ColumnWidth = 15
                    
                    
                    
                    .Cells(1, 1) = "Nr."
                    .Cells(1, 2) = "Ort 1 "
                    .Cells(1, 3) = "Anlage 1"
                    .Cells(1, 4) = "Info 1"
                    .Cells(1, 5) = "BMK 1"
                    .Cells(1, 6) = "Ort 2"
                    .Cells(1, 7) = "Anlage 2"
                    .Cells(1, 8) = "Info 2"
                    .Cells(1, 9) = "BMK 2"
                    .Cells(1, 10) = "Knoten"
                    .Cells(1, 11) = "Kabeltype"
                    .Cells(1, 12) = "Farbe"
                    .Cells(1, 13) = "Quer."
                    .Cells(1, 14) = "Verweis"
                    .Cells(1, 15) = "Kommentar"
                    
                    Cells(1, 1).Select
                    
                
                
                Set Table = DataStruct(0).Typical.Range("A2", "O10000")
                Table_Buffer = Table.Value
                Table_Row_Index = Table_Start_Line
                
                
                Call Get_PDF_Data(File_Path, 0)
                
                Worksheets(DataStruct(0).Typical_Names).Activate
'            With ActiveSheet
                
                .Range("A2", "O" & UBound(Table_Buffer)).VerticalAlignment = xlCenter
                .Range("A2", "O" & UBound(Table_Buffer)).HorizontalAlignment = xlCenter
                .Range("A2", "O" & UBound(Table_Buffer)).Font.Name = "Arial Narrow"
                .Range("A2", "O" & UBound(Table_Buffer)).Font.FontStyle = "Bold Italic"
                .Range("A2", "O" & UBound(Table_Buffer)).Font.Size = 10
                .Range("A2", "O" & UBound(Table_Buffer)).NumberFormat = "@"
                
                Set Table = DataStruct(0).Typical.Range("A2", "O" & UBound(Table_Buffer))
                Table.Value = Table_Buffer
                .Columns("O").EntireColumn.AutoFit
               
              
                DataStruct(0).Total_Row_BB = DataStruct(0).Last_Row_BB - DataStruct(0).First_Row_BB
                If (DataStruct(0).Total_Row_BB > 0) Then
                    DataStruct(0).Total_Row_BB = DataStruct(0).Total_Row_BB + 1
                
                    .Range(Cells(DataStruct(0).First_Row_BB - 1, 2), Cells(DataStruct(0).First_Row_BB - 1, 15)).Merge
                    .Cells(DataStruct(0).First_Row_BB - 1, 2).Font.Size = 16
                    .Cells(DataStruct(0).First_Row_BB - 1, 1).Interior.Color = RGB(255, 192, 0)
                    .Cells(DataStruct(0).First_Row_BB - 1, 2).Interior.Color = RGB(255, 192, 0)
                    .Cells(DataStruct(0).First_Row_BB - 1, 2).Font.FontStyle = "Bold"
                    .Cells(DataStruct(0).First_Row_BB - 1, 2).Value = " Zapojovacia tabu¾ka prepojov [ UEG ]"
                End If
            End With
        
        End If
    
 '----------------------------------------------------------------------------------------------------------------------------------------------
    
    For i = 1 To ST_Files_Count
    
        If DataStruct(i).Empty_Cabinet = False Then
            File_Path = File_Path_Temp & "\" & DataStruct(i).File_Names & ".pdf"
            Cabinet_Name_New = DataStruct(i).Typical_Names
            Call Create_Sheets(i)
            Cabinet_Name_Old = Cabinet_Name_New
        
            Worksheets(DataStruct(i).Typical_Names).Activate
            
            With ActiveSheet
        
            .Rows("1:1").Select
             Selection.RowHeight = 25
            With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 2
            .FreezePanes = True
            End With
            .Range("A1", "O1").Interior.Color = RGB(197, 217, 241)
            .Range("A1", "O1").VerticalAlignment = xlCenter
            .Range("A1", "O1").HorizontalAlignment = xlCenter
            .Range("A1", "O1").Font.Name = "Arial Narrow"
            .Range("A1", "O1").Font.FontStyle = "Bold Italic"
            .Range("A1", "O1").Font.Size = 14
            .Range("A1", "A1").ColumnWidth = 5
            .Range("B1", "B1").ColumnWidth = 10
            .Range("C1", "C1").ColumnWidth = 12
            .Range("D1", "D1").ColumnWidth = 10
            .Range("E1", "E1").ColumnWidth = 20
            .Range("F1", "F1").ColumnWidth = 10
            .Range("G1", "G1").ColumnWidth = 12
            .Range("H1", "H1").ColumnWidth = 10
            .Range("I1", "I1").ColumnWidth = 20
            .Range("J1", "J1").ColumnWidth = 12
            .Range("K1", "K1").ColumnWidth = 20
            .Range("L1", "L1").ColumnWidth = 10
            .Range("M1", "M1").ColumnWidth = 10
            .Range("N1").ColumnWidth = 15
           
            .Cells(1, 1) = "Nr."
            .Cells(1, 2) = "Ort 1 "
            .Cells(1, 3) = "Anlage 1"
            .Cells(1, 4) = "Info 1"
            .Cells(1, 5) = "BMK 1"
            .Cells(1, 6) = "Ort 2"
            .Cells(1, 7) = "Anlage 2"
            .Cells(1, 8) = "Info 2"
            .Cells(1, 9) = "BMK 2"
            .Cells(1, 10) = "Knoten"
            .Cells(1, 11) = "Kabeltype"
            .Cells(1, 12) = "Farbe"
            .Cells(1, 13) = "Quer."
            .Cells(1, 14) = "Verweis"
            .Cells(1, 15) = "Kommentar"
            
            Cells(1, 1).Select
            
        End With
        
        
'        Public Table As Range
'        Public Table_Buffer() As Variant
'        Public Table_Row_Index As Long
        
        
        
        Set Table = DataStruct(i).Typical.Range("A2", "O10000")
        Table_Buffer = Table.Value
        Table_Row_Index = Table_Start_Line
        
        
        Call Get_PDF_Data(File_Path, i)
        
        Worksheets(DataStruct(i).Typical_Names).Activate
        With ActiveSheet
        .Range("A2", "O" & UBound(Table_Buffer)).VerticalAlignment = xlCenter
        .Range("A2", "O" & UBound(Table_Buffer)).HorizontalAlignment = xlCenter
        .Range("A2", "O" & UBound(Table_Buffer)).Font.Name = "Arial Narrow"
        .Range("A2", "O" & UBound(Table_Buffer)).Font.FontStyle = "Bold Italic"
        .Range("A2", "O" & UBound(Table_Buffer)).Font.Size = 10
        .Range("A2", "O" & UBound(Table_Buffer)).NumberFormat = "@"

        Set Table = DataStruct(i).Typical.Range("A2", "O" & UBound(Table_Buffer))
        Table.Value = Table_Buffer
        
        .Columns("O").EntireColumn.AutoFit
        
    '-----------------------------------------------------------------------------------------------------------
    '------------------------------------------------------ BB -------------------------------------------------
    '-----------------------------------------------------------------------------------------------------------
        
        If ((DataStruct(i).Last_Row_BB <> 0) And _
            (DataStruct(i).Last_Row_BB <> DataStruct(i).First_Row_BB)) Then
            DataStruct(i).Total_Row_BB = DataStruct(i).Last_Row_BB - DataStruct(i).First_Row_BB
        ElseIf ((DataStruct(i).Last_Row_BB <> 0) And _
            (DataStruct(i).Last_Row_BB = DataStruct(i).First_Row_BB)) Then
            DataStruct(i).Total_Row_BB = 1
        End If
                
    '--------------------------

        If ((DataStruct(i).Total_Row_BB > 0) And _
            (DataStruct(i).Last_Row_BB <> DataStruct(i).First_Row_BB)) Then
                DataStruct(i).Total_Row_BB = DataStruct(i).Total_Row_BB + 1
        End If
        
    '--------------------------
        If ((DataStruct(i).Total_Row_BB > 0) And _
            (DataStruct(i).Last_Row_BB >= DataStruct(i).First_Row_BB)) Then
            
            .Range(Cells(DataStruct(i).First_Row_BB - 1, 2), Cells(DataStruct(i).First_Row_BB - 1, 15)).Merge
            .Cells(DataStruct(i).First_Row_BB - 1, 2).Font.Size = 16
            .Cells(DataStruct(i).First_Row_BB - 1, 1).Interior.Color = RGB(255, 192, 0)
            .Cells(DataStruct(i).First_Row_BB - 1, 2).Interior.Color = RGB(255, 192, 0)
            .Cells(DataStruct(i).First_Row_BB - 1, 2).Font.FontStyle = "Bold"
            .Cells(DataStruct(i).First_Row_BB - 1, 2).Value = " Zapojovacia tabu¾ka [ BB ]"
        
        End If
        
    '-----------------------------------------------------------------------------------------------------------
    '------------------------------------------------------ BT -------------------------------------------------
    '-----------------------------------------------------------------------------------------------------------

        If ((DataStruct(i).Last_Row_BT <> 0) And _
            (DataStruct(i).Last_Row_BT <> DataStruct(i).First_Row_BT)) Then
                DataStruct(i).Total_Row_BT = DataStruct(i).Last_Row_BT - DataStruct(i).First_Row_BT
        ElseIf ((DataStruct(i).Last_Row_BT <> 0) And _
            (DataStruct(i).Last_Row_BT = DataStruct(i).First_Row_BT)) Then
                DataStruct(i).Total_Row_BT = 1
        End If
          
    '--------------------------
        
        If ((DataStruct(i).Total_Row_BT > 0) And _
            (DataStruct(i).Last_Row_BT <> DataStruct(i).First_Row_BT)) Then
                DataStruct(i).Total_Row_BT = DataStruct(i).Total_Row_BT + 1
        End If
        
      '--------------------------
        If ((DataStruct(i).Total_Row_BT > 0) And _
            (DataStruct(i).Last_Row_BT >= DataStruct(i).First_Row_BT)) Then
        
          .Range(Cells(DataStruct(i).First_Row_BT - 1, 2), Cells(DataStruct(i).First_Row_BT - 1, 15)).Merge
          .Cells(DataStruct(i).First_Row_BT - 1, 2).Font.Size = 16
          .Cells(DataStruct(i).First_Row_BT - 1, 1).Interior.Color = RGB(255, 192, 0)
          .Cells(DataStruct(i).First_Row_BT - 1, 2).Interior.Color = RGB(255, 192, 0)
          .Cells(DataStruct(i).First_Row_BT - 1, 2).Font.FontStyle = "Bold"
          .Cells(DataStruct(i).First_Row_BT - 1, 2).Value = " Zapojovacia tabu¾ka [ BT ]"
            
        End If
        
    '-----------------------------------------------------------------------------------------------------------
    '------------------------------------------------------ NB -------------------------------------------------
    '-----------------------------------------------------------------------------------------------------------

        If ((DataStruct(i).Last_Row_NB <> 0) And _
            (DataStruct(i).Last_Row_NB <> DataStruct(i).First_Row_NB)) Then
                DataStruct(i).Total_Row_NB = DataStruct(i).Last_Row_NB - DataStruct(i).First_Row_NB
        ElseIf ((DataStruct(i).Last_Row_NB <> 0) And _
            (DataStruct(i).Last_Row_NB = DataStruct(i).First_Row_NB)) Then
                DataStruct(i).Total_Row_NB = 1
        End If
        
    '--------------------------
        If ((DataStruct(i).Total_Row_NB > 0) And _
            (DataStruct(i).Last_Row_NB <> DataStruct(i).First_Row_NB)) Then
                DataStruct(i).Total_Row_NB = DataStruct(i).Total_Row_NB + 1
        End If
    '--------------------------
        If ((DataStruct(i).Total_Row_NB > 0) And _
            (DataStruct(i).Last_Row_NB >= DataStruct(i).First_Row_NB)) Then
        
            .Range(Cells(DataStruct(i).First_Row_NB - 1, 2), Cells(DataStruct(i).First_Row_NB - 1, 15)).Merge
            .Cells(DataStruct(i).First_Row_NB - 1, 2).Font.Size = 16
            .Cells(DataStruct(i).First_Row_NB - 1, 1).Interior.Color = RGB(255, 192, 0)
            .Cells(DataStruct(i).First_Row_NB - 1, 2).Interior.Color = RGB(255, 192, 0)
            .Cells(DataStruct(i).First_Row_NB - 1, 2).Font.FontStyle = "Bold"
            .Cells(DataStruct(i).First_Row_NB - 1, 2).Value = " Zapojovacia tabu¾ka [ NB ]"
        End If
        
    '-----------------------------------------------------------------------------------------------------------
    '------------------------------------------------------ NT -------------------------------------------------
    '-----------------------------------------------------------------------------------------------------------
        
        If ((DataStruct(i).Last_Row_NT <> 0) And _
            (DataStruct(i).Last_Row_NT <> DataStruct(i).First_Row_NT)) Then
                DataStruct(i).Total_Row_NT = DataStruct(i).Last_Row_NT - DataStruct(i).First_Row_NT
        ElseIf ((DataStruct(i).Last_Row_NT <> 0) And _
            (DataStruct(i).Last_Row_NT = DataStruct(i).First_Row_NT)) Then
                DataStruct(i).Total_Row_NT = 1
        End If
        
        '-----------------------
            If ((DataStruct(i).Total_Row_NT > 0) And _
                (DataStruct(i).Last_Row_NT <> DataStruct(i).First_Row_NT)) Then
                    DataStruct(i).Total_Row_NT = DataStruct(i).Total_Row_NT + 1
            End If
       '-----------------------
            If ((DataStruct(i).Total_Row_NT > 0) And _
            (DataStruct(i).Last_Row_NT >= DataStruct(i).First_Row_NT)) Then
  
            .Range(Cells(DataStruct(i).First_Row_NT - 1, 2), Cells(DataStruct(i).First_Row_NT - 1, 15)).Merge
            .Cells(DataStruct(i).First_Row_NT - 1, 2).Font.Size = 16
            .Cells(DataStruct(i).First_Row_NT - 1, 1).Interior.Color = RGB(255, 192, 0)
            .Cells(DataStruct(i).First_Row_NT - 1, 2).Interior.Color = RGB(255, 192, 0)
            .Cells(DataStruct(i).First_Row_NT - 1, 1).Font.FontStyle = "Bold"
            .Cells(DataStruct(i).First_Row_NT - 1, 2).Value = " Zapojovacia tabu¾ka [ NT ]"
        End If
'--------------------------------------------------------------------------------------------
    End With
'--------------------------------------------------------------------------------------------
    
    End If
Next i

    Set Table = Nothing
    Worksheets("Main_Sheet").Activate
    Call Storing_Variables(ST_Files_Count)
    Call Data_To_Display
Cancel_Import_PDF_Documents:
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Sub Get_PDF_Data(PDF_File As String, File_Index As Integer)
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Dim AC_PD As Acrobat.AcroPDDoc                          'access pdf file
Dim AC_Hi As Acrobat.AcroHiliteList                     'set selection word count
Dim AC_PG As Acrobat.AcroPDPage                         'get the particular page
Dim AC_PGTxt As Acrobat.AcroPDTextSelect                'get the text of selection area

Dim WS_PDF As Worksheet
Dim Row_Num As Long                                       'row count
Dim Col_Num As Long                                  'column count
Dim Li_Row As Long                                      'Maximum rows limit for one column
Dim First_Page As Boolean                               'to identify beginning of page

Li_Row = Rows.Count

Dim NumberOfPage As Long                                'count pages in pdf file
Dim Page_Ctn    As Integer
Dim i As Long, j As Long, k As Long                     'looping variables
Dim T_Str As String

Dim Hld_Txt As Variant                                  'get PDF total text into array
Dim Txt_Temp As String
Dim Page_Buffer() As Variant


Row_Num = 0                                               'set the intial value
Col_Num = 1                                             'set the intial value


Set AC_PD = New Acrobat.AcroPDDoc
Set AC_Hi = New Acrobat.AcroHiliteList

AC_Hi.Add 0, 32767                                      'set maximum selection area of PDF page

With AC_PD
    .Open PDF_File                                      'open PDF file
    NumberOfPage = .GetNumPages                         'get the number of pages of PDF file
    DataStruct(File_Index).Total_Pages = NumberOfPage
    If NumberOfPage = -1 Then                           'if get pages is failed exit sub
        MsgBox "Pages Cannot determine in PDF file '" & PDF_File & "'"
        .Close
        GoTo h_end
    End If

    Page_Ctn = 1
    First_Page = True
    
    
    For i = 1 To NumberOfPage                            'looping through sheets
   

        Row_Num = 0
        Erase Page_Buffer()
        T_Str = ""
        
        Set AC_PG = .AcquirePage(i - 1)                 'get the page
        Set AC_PGTxt = AC_PG.CreateWordHilite(AC_Hi)    'get the full page selection
        
        If Not AC_PGTxt Is Nothing Then                 'if text selected successfully get the all the text into T_Str string
        
            With AC_PGTxt
                
                For j = 0 To .GetNumText - 1
                    T_Str = T_Str & .GetText(j)
                Next j
            End With
        End If                                     'transfer PDF data into sheet
                                                   'get the pdf data into single sheet
            
    If T_Str <> "" Then
            Hld_Txt = Split(T_Str, vbCrLf)
            
                For k = 0 To UBound(Hld_Txt)
                    
                    Row_Num = Row_Num + 1
                    ReDim Preserve Page_Buffer(Row_Num)
                    
                    T_Str = CStr(Hld_Txt(k))
                    
                    If Left(T_Str, 1) = "=" Then
                    T_Str = "'" & T_Str
                    End If
                    
                    Page_Buffer(Row_Num) = T_Str
                    
                Next k
                    
                    Else
                        Row_Num = Row_Num + 1
                        ReDim Preserve Page_Buffer(Row_Num)
                        Page_Buffer(Row_Num) = "No text found in page " & i
                        Row_Num = Row_Num + 1
                        ReDim Preserve Page_Buffer(Row_Num)
        End If
'----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------
                    If (Selection_Sheet.Range("B40").Value = 1) Then
                        Call Get_Name_Of_Page(Page_Buffer, File_Index, Page_Ctn)
                    ElseIf (Selection_Sheet.Range("B40").Value = 2) Then
                        Call Get_Name_Of_Page_v2(Page_Buffer, File_Index, Page_Ctn)
                    ElseIf (Selection_Sheet.Range("B40").Value = 3) Then
                    End If
'----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------
                    
                    If First_Page = True Then
                        
                        DataStruct(File_Index).Sie_Werk_Nr = Get_Project_Number(Page_Buffer, "Werk- Nr.")
                        
                        If (DataStruct(File_Index).Sie_Werk_Nr = " ??? ") Then
                            DataStruct(File_Index).Sie_Werk_Nr = Get_Project_Number(Page_Buffer, "IfdNr")
                        End If

'                        If (DataStruct(File_Index).Sie_Werk_Nr = " ??? ") Then
'                        DataStruct(File_Index).Sie_Werk_Nr = Get_Project_Number(Page_Buffer, "SWF ")
'                        End If
                        First_Page = False
                    End If
                        Page_Ctn = Page_Ctn + 1
        
'/////////////////////////////////////////////////////// [ BB ] ////////////////////////////////////////////////////////////
        
        If ((Table_BB = True) And (DataStruct(File_Index).First_Page_BB = 0)) Then
            DataStruct(File_Index).First_Page_BB = i
            DataStruct(File_Index).Last_Page_BB = i - 1
        End If
        If ((Table_BB = True) And (DataStruct(File_Index).First_Page_BB > 0)) Then
             DataStruct(File_Index).Last_Page_BB = DataStruct(File_Index).Last_Page_BB + 1
        End If
  
'/////////////////////////////////////////////////////// [ BT ] ////////////////////////////////////////////////////////////
        
        If ((Table_BT = True) And (DataStruct(File_Index).First_Page_BT = 0)) Then
            DataStruct(File_Index).First_Page_BT = i
            DataStruct(File_Index).Last_Page_BT = i - 1
        End If
        If ((Table_BT = True) And (DataStruct(File_Index).First_Page_BT > 0)) Then
             DataStruct(File_Index).Last_Page_BT = DataStruct(File_Index).Last_Page_BT + 1
        End If
        
'/////////////////////////////////////////////////////// [ NB ] ////////////////////////////////////////////////////////////
        
        If ((Table_NB = True) And (DataStruct(File_Index).First_Page_NB = 0)) Then
            DataStruct(File_Index).First_Page_NB = i
            DataStruct(File_Index).Last_Page_NB = i - 1
        End If
        If ((Table_NB = True) And (DataStruct(File_Index).First_Page_NB > 0)) Then
             DataStruct(File_Index).Last_Page_NB = DataStruct(File_Index).Last_Page_NB + 1
        End If
        
'/////////////////////////////////////////////////////// [ NT ] ////////////////////////////////////////////////////////////
        
        If ((Table_NT = True) And (DataStruct(File_Index).First_Page_NT = 0)) Then
            DataStruct(File_Index).First_Page_NT = i
            DataStruct(File_Index).Last_Page_NT = i - 1
        End If
        If ((Table_NT = True) And (DataStruct(File_Index).First_Page_NT > 0)) Then
             DataStruct(File_Index).Last_Page_NT = DataStruct(File_Index).Last_Page_NT + 1
        End If

        Call Make_Table_From_Raw_Data(File_Index, Page_Buffer)
       
Next i
    .Close
End With
        
h_end:

Set AC_PGTxt = Nothing
Set AC_PG = Nothing
Set AC_Hi = Nothing
Set AC_PD = Nothing
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Sub Create_Sheets(File_Index As Integer)
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Dim WS_Cabinet As Worksheet
Dim In_UsE As Boolean
Dim i As Integer

    If Cabinet_Name_Old <> Cabinet_Name_New Then
        In_UsE = False
        For i = 1 To Worksheets.Count
            If Worksheets(i).Name = "Cabinet_Name_Old" Then
                    In_UsE = True
            End If
                    Next i
                If In_UsE = False Then
                    Set WS_Cabinet = Worksheets.Add(, Worksheets(Sheets.Count))
                    WS_Cabinet.Name = Cabinet_Name_New
                    Set DataStruct(File_Index).Typical = WS_Cabinet
                End If
    End If
    Set WS_Cabinet = Nothing
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Sub Delete_Sheets()
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Dim j As Long
Dim k As Long
    
    Application.DisplayAlerts = False
    
' RESET Array Values

    For j = 0 To UBound(DataStruct)
    Set DataStruct(j).Typical = Nothing
    Next j
   
    Erase DataStruct()
    
    j = Worksheets.Count
    For k = j To 4 Step -1
        Sheets(k).Delete
    Next k
    Application.DisplayAlerts = True
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Sub Data_To_Display()
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Dim i, j As Integer
Dim Result As Integer
Dim offset As Integer
Dim Disp_Data_Row As Long
Dim Disp_Data_Column As Long
Dim myRange As Range
Dim Condition1 As FormatCondition
Dim Condition2 As FormatCondition
Dim CompareErr As Boolean
Dim Start_Pos As Integer
Dim cb1 As CheckBox
Dim cb2 As CheckBox

    offset = ST_Files_Count + 1
    
    Application.ScreenUpdating = False
    Main_Sheet.Range("B11", "V" & offset + 15).HorizontalAlignment = xlCenter
    Main_Sheet.Range("B11", "V" & offset + 15).VerticalAlignment = xlCenter
    Main_Sheet.Range("B11", "V" & offset + 15).Font.Name = "Arial"
    Main_Sheet.Range("B11", "V" & offset + 15).Font.FontStyle = "Bold Italic"
    Main_Sheet.Range("B11", "V" & offset + 15).Font.Size = 16
    
    Main_Sheet.Range("B11", "B" & offset + 11).NumberFormat = "0"" ."""
    Main_Sheet.Range("C11", "C" & offset + 11).NumberFormat = "@"
'    Main_Sheet.Range("R11", "R" & Offset + 11).NumberFormat = "0"" ks"""
    Main_Sheet.Range("S13", "S" & offset + 11).NumberFormat = "@"
    Main_Sheet.Range("V11").NumberFormat = "=  0""  ks"""
    Main_Sheet.Range("V11").Font.Size = 20
    Main_Sheet.Range("V11").Font.Color = vbRed
    Main_Sheet.Range("V13", "V" & offset + 11).NumberFormat = "x  0""  ks"""
    
    
    Disp_Data_Row = 11
    Disp_Data_Column = 0

    Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 2) = 1
    Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 3) = DataStruct(0).File_Names
    Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 4) = DataStruct(0).Total_Pages
    
    For j = 5 To 19
        Main_Sheet.Cells(Disp_Data_Row, Disp_Data_Column + j) = " ----- "
    Next j
    
'         Main_Sheet.Cells(Disp_Data_Row, Disp_Data_Column + 7) = DataStruct(0).Total_Row_BB
         
         
    CompareErr = False
    If (Without_Interconnection_Table = False) Then
        Start_Pos = 0
    ElseIf (Without_Interconnection_Table = True) Then
        Start_Pos = 1
    End If
    
    For i = Start_Pos To ST_Files_Count
        For j = i To ST_Files_Count

            If ((DataStruct(i).Empty_Cabinet <> True) And (DataStruct(j).Empty_Cabinet <> True)) Then
                Result = StrComp(UCase(DataStruct(i).Sie_Werk_Nr), UCase(DataStruct(j).Sie_Werk_Nr), vbTextCompare)
                If Result = 0 Then
                Main_Sheet.Range("G7").Value = DataStruct(1).Sie_Werk_Nr
                Else
                    If ((CompareErr <> True) And (Result <> 0)) Then
                        CompareErr = True
                    End If
                End If
            End If
        Next j
    Next i
        
        If (CompareErr = True) Then Main_Sheet.Range("G7").Value = " Chyba !!! "
    
    If (Main_Sheet.Range("E4").Value = "1") Then
        Main_Sheet.Range("B5").Value = "Zobrazi poèet radov :"
        Main_Sheet.Range("G10").Value = "Poèet radov :"
        Main_Sheet.Range("J10").Value = "Poèet radov :"
        Main_Sheet.Range("M10").Value = "Poèet radov :"
        Main_Sheet.Range("P10").Value = "Poèet radov :"
    ElseIf (Main_Sheet.Range("E4").Value = "2") Then
        Main_Sheet.Range("B5").Value = "Zobrazi poèet vodièov :"
        Main_Sheet.Range("G10").Value = "Poèet vodièov :"
        Main_Sheet.Range("J10").Value = "Poèet vodièov :"
        Main_Sheet.Range("M10").Value = "Poèet vodièov :"
        Main_Sheet.Range("P10").Value = "Poèet vodièov :"
    End If
    
'=========================================================== [Loop Wires] ==========================================================
    Disp_Data_Row = 11
    If (Main_Sheet.Range("D5") = True) Then
        If (Main_Sheet.Range("E4").Value = "1") Then
            Main_Sheet.Cells(Disp_Data_Row, Disp_Data_Column + 7) = DataStruct(0).Total_Row_BB
        ElseIf (Main_Sheet.Range("E4").Value = "2") Then
            Main_Sheet.Cells(Disp_Data_Row, Disp_Data_Column + 7) = DataStruct(0).Wire_Amount_BB
        End If
    Else
        Main_Sheet.Cells(Disp_Data_Row, Disp_Data_Column + 7).ClearContents
    End If
    
    If (DataStruct(0).Wire_Amount_BB < 0) Then
        If (DataStruct(0).Wire_Amount_BB = Wire_Amount_Failure) Then
            Main_Sheet.Cells(Disp_Data_Row, Disp_Data_Column + 7) = "Wire Fail"
        ElseIf (DataStruct(0).Wire_Amount_BB = Empty_Line_Found) Then
            Main_Sheet.Cells(Disp_Data_Row, Disp_Data_Column + 7) = "Empty Line !!!"
        End If
    End If
    
    '------------------------------------------------------------------------------------
    ' Put Checkbox for Looping Wires
    '------------------------------------------------------------------------------------
    
            Set cb2 = ActiveSheet.CheckBoxes.Add(Main_Sheet.Range("U11").Left, _
                                                    Main_Sheet.Range("U11").Top, _
                                                    Main_Sheet.Range("U11").Width, _
                                                    Main_Sheet.Range("U11").Height)
            With cb2
                    .Caption = ""
                    .LinkedCell = Main_Sheet.Range("U11").Address
                    .Display3DShading = True
                    
                    If (DataStruct(0).Cabinet_Processing = True) Then
                        Main_Sheet.Range("U11").Value = True
                    ElseIf (DataStruct(0).Cabinet_Processing = True) Then
                        Main_Sheet.Range("U11").Value = False
                    End If
                
            End With
                    
                    Main_Sheet.Range("U11").Font.Color = vbWhite
    
    Disp_Data_Row = 12
    For i = 1 To ST_Files_Count

    Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 2) = i
    Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 3) = DataStruct(i).File_Names
    Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 4) = DataStruct(i).Total_Pages
    Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 5) = DataStruct(i).First_Page_BB
    Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 6) = DataStruct(i).Last_Page_BB
    
    If (DataStruct(i).Cabinet_Multipiler < 1) Then
    ElseIf (DataStruct(i).Cabinet_Multipiler > 0) Then
        Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 22) = DataStruct(i).Cabinet_Multipiler
    End If

    Call Refresh_Cabinet_Height_Value
    Call Refresh_Wire_Typ_Value
    Call Refresh_Heat_Shrink_Tubes_Display_State
    
    '------------------------------------------------------------------------------------
    ' Put Checkboxes CB1 for Typicals ( Wide Cabinets )
    '------------------------------------------------------------------------------------
                            
            Set cb1 = ActiveSheet.CheckBoxes.Add(Main_Sheet.Range("T" & (i + 13) - 1).Left, _
                                                    Main_Sheet.Range("T" & (i + 13) - 1).Top, _
                                                    Main_Sheet.Range("T" & (i + 13) - 1).Width, _
                                                    Main_Sheet.Range("T" & (i + 13) - 1).Height)
            With cb1
                    .Caption = ""
                    .LinkedCell = Main_Sheet.Range("T" & (i + 13) - 1).Address
                    .Display3DShading = True
                    Main_Sheet.Range("T" & (i + 13) - 1).Value = DataStruct(i).Wide_Cabinet
            End With
                    Main_Sheet.Range("T" & (i + 13) - 1).Font.Color = vbWhite

    '------------------------------------------------------------------------------------
    ' Put Checkboxes CB2 for Typicals
    '------------------------------------------------------------------------------------
    
            Set cb2 = ActiveSheet.CheckBoxes.Add(Main_Sheet.Range("U" & (i + 13) - 1).Left, _
                                                    Main_Sheet.Range("U" & (i + 13) - 1).Top, _
                                                    Main_Sheet.Range("U" & (i + 13) - 1).Width, _
                                                    Main_Sheet.Range("U" & (i + 13) - 1).Height)
            With cb2
                    .Caption = ""
                    .LinkedCell = Main_Sheet.Range("U" & (i + 13) - 1).Address
                    .Display3DShading = True
                    Main_Sheet.Range("U" & (i + 13) - 1).Value = DataStruct(i).Cabinet_Processing
            End With
                    Main_Sheet.Range("U" & (i + 13) - 1).Font.Color = vbWhite

'=========================================================== [ BB ] ==========================================================
    
    If (Main_Sheet.Range("D5") = True) Then
        If (Main_Sheet.Range("E4").Value = "1") Then
            Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 7) = DataStruct(i).Total_Row_BB
        ElseIf (Main_Sheet.Range("E4").Value = "2") Then
            
            If ((DataStruct(i).Wire_Amount_BB = 0) And (DataStruct(i).Total_Row_BB > 0)) Then
                Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 7) = "N/A"
             Else
                Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 7) = DataStruct(i).Wire_Amount_BB
            End If
        
        End If
    Else
    Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 7).ClearContents
    End If
    
    If (DataStruct(i).Wire_Amount_BB < 0) Then
        If (DataStruct(i).Wire_Amount_BB = Wire_Amount_Failure) Then
            Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 7) = "Wire Fail"
        ElseIf (DataStruct(i).Wire_Amount_BB = Empty_Line_Found) Then
            Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 7) = "Empty Line !!!"
        End If
    End If
    
    Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 8) = DataStruct(i).First_Page_BT
    Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 9) = DataStruct(i).Last_Page_BT
    
'=========================================================== [ BT ] ==========================================================
    
    If (Main_Sheet.Range("D5") = True) Then
        If (Main_Sheet.Range("E4").Value = "1") Then
            Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 10) = DataStruct(i).Total_Row_BT
        ElseIf (Main_Sheet.Range("E4").Value = "2") Then
            
            If ((DataStruct(i).Wire_Amount_BT = 0) And (DataStruct(i).Total_Row_BT > 0)) Then
                Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 10) = "N/A"
            Else
                Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 10) = DataStruct(i).Wire_Amount_BT
            End If
        
        End If
    Else
    Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 10).ClearContents
    End If
    
    If (DataStruct(i).Wire_Amount_BT < 0) Then
        If (DataStruct(i).Wire_Amount_BT = Wire_Amount_Failure) Then
            Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 10) = "Wire Fail"
        ElseIf (DataStruct(i).Wire_Amount_BT = Empty_Line_Found) Then
            Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 10) = "Empty Line !!!"
        End If
    End If
    
    Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 11) = DataStruct(i).First_Page_NB
    Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 12) = DataStruct(i).Last_Page_NB
    
   '=========================================================== [ NB ] ==========================================================
    
    If (Main_Sheet.Range("D5") = True) Then
        If (Main_Sheet.Range("E4").Value = "1") Then
            Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 13) = DataStruct(i).Total_Row_NB
        ElseIf (Main_Sheet.Range("E4").Value = "2") Then
            
            If ((DataStruct(i).Wire_Amount_NB = 0) And (DataStruct(i).Total_Row_NB > 0)) Then
                Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 13) = "N/A"
             Else
                Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 13) = DataStruct(i).Wire_Amount_NB
            End If

        End If
    Else
    Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 13).ClearContents
    End If
    
    If (DataStruct(i).Wire_Amount_NB < 0) Then
        If (DataStruct(i).Wire_Amount_NB = Wire_Amount_Failure) Then
            Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 13) = "Wire Fail"
        ElseIf (DataStruct(i).Wire_Amount_NB = Empty_Line_Found) Then
            Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 13) = "Empty Line !!!"
        End If
    End If
    
    Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 14) = DataStruct(i).First_Page_NT
    Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 15) = DataStruct(i).Last_Page_NT
    
'=========================================================== [ NT ] ==========================================================
    
    If (Main_Sheet.Range("D5") = True) Then
    If (Main_Sheet.Range("E4").Value = "1") Then
            Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 16) = DataStruct(i).Total_Row_NT
        ElseIf (Main_Sheet.Range("E4").Value = "2") Then
            
            If ((DataStruct(i).Wire_Amount_NT = 0) And (DataStruct(i).Total_Row_NT > 0)) Then
                Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 16) = "N/A"
             Else
                Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 16) = DataStruct(i).Wire_Amount_NT
            End If
        
        End If
    Else
    Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 16).ClearContents
    End If
    
    If (DataStruct(i).Wire_Amount_NT < 0) Then
        If (DataStruct(i).Wire_Amount_NT = Wire_Amount_Failure) Then
            Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 16) = "Wire Fail"
        ElseIf (DataStruct(i).Wire_Amount_NT = Empty_Line_Found) Then
            Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 16) = "Empty Line !!!"
        End If
    End If
    
    Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 19) = DataStruct(i).Typical_Names

If (DataStruct(i).Empty_Cabinet) = False Then
    Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 17).Font.Color = vbBlack
    Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 17) = "Nie"
Else
    Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 17).Font.Color = vbRed
    Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 17) = "!!! Áno !!!"
    
    
'    Main_Sheet.Range("B11", "S" & Offset + 15).Font.Name = "Arial"
'    Main_Sheet.Range("B11", "S" & Offset + 15).Font.FontStyle = "Bold Italic"
'    Main_Sheet.Range("B11", "S" & Offset + 15).Font.Size = 16
    
    For j = 4 To 16
      Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + j) = " ----- "
    Next j
End If

    
    If (DataStruct(i).Wire_Amount_UEG > 0) Then
        Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 18) = DataStruct(i).Wire_Amount_UEG
    Else
        Main_Sheet.Cells(Disp_Data_Row + i, Disp_Data_Column + 18).ClearContents
    End If

Next i
     
        Main_Sheet.Cells(9, 18) = "Poèet [ UEG ]" & vbCrLf & "Prepojov"
     
     If (Main_Sheet.Range("D4") = True) Then
        Main_Sheet.Cells(9, 19) = "Názov" & vbCrLf & "Typicalu:"
    Else
        Main_Sheet.Cells(9, 19) = "Názov" & vbCrLf & "Skrine:"
    End If
    
    
    Main_Sheet.Range("B11", "V" & offset + 11).Borders.LineStyle = xlDouble
    
    Set myRange = Range("E13", "F" & offset + 11)
    myRange.FormatConditions.Delete
    Set Condition1 = myRange.FormatConditions.Add(xlCellValue, xlLessEqual, "=0")
    Set Condition2 = myRange.FormatConditions.Add(xlCellValue, xlGreater, "=0")

    With Condition1
    .Interior.Color = RGB(197, 217, 241)
    .Font.Color = RGB(197, 217, 241)
    .Font.Bold = True
    End With

    With Condition2
    .Interior.Color = RGB(197, 217, 241)
    .Font.Color = RGB(0, 0, 0)
    .Font.Bold = True
    End With
    
    Set myRange = Range("H13", "I" & offset + 11)
    myRange.FormatConditions.Delete
    Set Condition1 = myRange.FormatConditions.Add(xlCellValue, xlLessEqual, "=0")
    Set Condition2 = myRange.FormatConditions.Add(xlCellValue, xlGreater, "=0")
    
    With Condition1
    .Interior.Color = RGB(197, 217, 241)
    .Font.Color = RGB(197, 217, 241)
    .Font.Bold = True
    End With

    With Condition2
    .Interior.Color = RGB(197, 217, 241)
    .Font.Color = RGB(0, 0, 0)
    .Font.Bold = True
    End With
    
    Set myRange = Range("K13", "L" & offset + 11)
    myRange.FormatConditions.Delete
    Set Condition1 = myRange.FormatConditions.Add(xlCellValue, xlLessEqual, "=0")
    Set Condition2 = myRange.FormatConditions.Add(xlCellValue, xlGreater, "=0")
    
    With Condition1
    .Interior.Color = RGB(197, 217, 241)
    .Font.Color = RGB(197, 217, 241)
    .Font.Bold = True
    End With

    With Condition2
    .Interior.Color = RGB(197, 217, 241)
    .Font.Color = RGB(0, 0, 0)
    .Font.Bold = True
    End With

    Set myRange = Range("N13", "O" & offset + 11)
    myRange.FormatConditions.Delete
    Set Condition1 = myRange.FormatConditions.Add(xlCellValue, xlLessEqual, "=0")
    Set Condition2 = myRange.FormatConditions.Add(xlCellValue, xlGreater, "=0")
    
    With Condition1
    .Interior.Color = RGB(197, 217, 241)
    .Font.Color = RGB(197, 217, 241)
    .Font.Bold = True
    End With

    With Condition2
    .Interior.Color = RGB(197, 217, 241)
    .Font.Color = RGB(0, 0, 0)
    .Font.Bold = True
    End With

    Set myRange = Range("D11", "D" & offset + 11)
    myRange.FormatConditions.Delete
    Set Condition1 = myRange.FormatConditions.Add(xlCellValue, xlLessEqual, "=0")
    Set Condition2 = myRange.FormatConditions.Add(xlCellValue, xlGreater, "=0")
    
    With Condition1
    .Font.Color = RGB(255, 255, 255)
    .Font.Bold = True
    End With

    With Condition2
    .Font.Color = RGB(0, 0, 0)
    .Font.Bold = True
    End With
    
    
    Set myRange = Range("G11", "G" & offset + 11)
    myRange.FormatConditions.Delete
    Set Condition1 = myRange.FormatConditions.Add(xlCellValue, xlLessEqual, "=0")
    Set Condition2 = myRange.FormatConditions.Add(xlCellValue, xlGreater, "=0")
    
    With Condition1
    .Font.Color = RGB(255, 255, 255)
    .Font.Bold = True
    End With
    
    Set myRange = Range("G13", "G" & offset + 11)
    myRange.FormatConditions.Delete
    Set Condition1 = myRange.FormatConditions.Add(xlCellValue, xlLessEqual, "=0")
    Set Condition2 = myRange.FormatConditions.Add(xlCellValue, xlGreater, "=0")
    
    With Condition1
    .Font.Color = RGB(255, 255, 255)
    .Font.Bold = True
    End With

    With Condition2
    .Font.Color = RGB(0, 0, 0)
    .Font.Bold = True
    End With
    Set myRange = Range("J13", "J" & offset + 11)
    myRange.FormatConditions.Delete
    Set Condition1 = myRange.FormatConditions.Add(xlCellValue, xlLessEqual, "=0")
    Set Condition2 = myRange.FormatConditions.Add(xlCellValue, xlGreater, "=0")
    
    With Condition1
    .Font.Color = RGB(255, 255, 255)
    .Font.Bold = True
    End With

    With Condition2
    .Font.Color = RGB(0, 0, 0)
    .Font.Bold = True
    End With
    
    Set myRange = Range("M13", "M" & offset + 11)
    myRange.FormatConditions.Delete
    Set Condition1 = myRange.FormatConditions.Add(xlCellValue, xlLessEqual, "=0")
    Set Condition2 = myRange.FormatConditions.Add(xlCellValue, xlGreater, "=0")
    
    With Condition1
    .Font.Color = RGB(255, 255, 255)
    .Font.Bold = True
    End With

    With Condition2
    .Font.Color = RGB(0, 0, 0)
    .Font.Bold = True
    End With
    
    Set myRange = Range("P13", "P" & offset + 11)
    myRange.FormatConditions.Delete
    Set Condition1 = myRange.FormatConditions.Add(xlCellValue, xlLessEqual, "=0")
    Set Condition2 = myRange.FormatConditions.Add(xlCellValue, xlGreater, "=0")
    
    With Condition1
    .Font.Color = RGB(255, 255, 255)
    .Font.Bold = True
    End With

    With Condition2
    .Font.Color = RGB(0, 0, 0)
    .Font.Bold = True
    End With
    
    Set myRange = Range("B8", "B8")
    
    'Worksheets("Main_Sheet").Activate
    Application.ScreenUpdating = True


End Sub
