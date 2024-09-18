Sub SegmentaçãoPDF()
Application.ScreenUpdating = False


Dim Rede As SlicerCache
Dim Lojas As SlicerCache

Dim Rede1 As SlicerItem
Dim Rede2 As SlicerItem
Dim Lojas1 As SlicerItem
Dim Lojas2 As SlicerItem

Set Rede = ThisWorkbook.SlicerCaches("SegmentaçãodeDados_Rede")
Set Lojas = ThisWorkbook.SlicerCaches("SegmentaçãodeDados_Lojas")

    For Each Rede1 In Rede.SlicerItems
    Rede.ClearAllFilters
    
    For Each Rede2 In Rede.SlicerItems
    If Rede1.Name <> Rede2.Name Then
    Rede2.Selected = False
    
        End If
    Next Rede2
    
    With ActiveSheet.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        
        End With
    
   ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=ThisWorkbook.Path & "\" & "Rede " & Rede1.Name & ".pdf", _
            Quality:=xlQualityStandard, IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, OpenAfterPublish:=False

Next Rede1

    Rede.ClearAllFilters
    MsgBox ("Ok!")
    
End Sub