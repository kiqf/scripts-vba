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
            
    For Each Lojas1 In Lojas.SlicerItems
        Lojas.ClearAllFilters
        For Each Lojas2 In Lojas.SlicerItems
            If Lojas1.Name <> Lojas2.Name Then
                Lojas2.Selected = False
            
            End If
            
        Next
        
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=ThisWorkbook.Path & "\" & Lojas1.Name & " - " & Rede1.Name & ".pdf", _
            Quality:=xlQualityStandard, IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, OpenAfterPublish:=False
        
    Next Lojas1
    

Next Rede1

    Rede.ClearAllFilters
    Lojas.ClearAllFilters
    Application.ScreenUpdating = True
    MsgBox ("Ok!")
    
End Sub