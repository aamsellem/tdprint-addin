Attribute VB_Name = "ModMain"
Option Explicit

' ============================================================================
' Module: ModMain
' Description: Point d'entrée pour l'add-in TD Print Variables
' ============================================================================

' --------------------------------------------------------------------------
' Afficher le panneau des variables TD Print
' --------------------------------------------------------------------------
Public Sub ShowTDPrintVariables()
    FrmTDPrintVariables.Show vbModeless
End Sub

' --------------------------------------------------------------------------
' Raccourci pour afficher le panneau (peut être assigné à un bouton)
' --------------------------------------------------------------------------
Public Sub TDPrint()
    ShowTDPrintVariables
End Sub
