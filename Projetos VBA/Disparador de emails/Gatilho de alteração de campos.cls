VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim Celula As Range
    
    Set Celula = Range("C7")
    
    'Application.EnableEvents = True
    
    
    If Not Application.Intersect(Celula, Range(Target.Address)) Is Nothing Then
     
     Call AtualizaCorpo
     
    End If
    

End Sub

