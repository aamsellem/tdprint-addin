VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmTDPrintVariables
   Caption         =   "TD Print Variables"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   OleObjectBlob   =   "FrmTDPrintVariables.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmTDPrintVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' UserForm: FrmTDPrintVariables
' Description: Interface pour sélectionner et insérer les variables TD Print
' ============================================================================

Private Const FAVORITES_KEY As String = "TDPrintFavorites"
Private Favorites As Collection
Private AllVariablesFlat As Collection
Private CurrentFilter As String

' --------------------------------------------------------------------------
' Initialisation du formulaire
' --------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    ' Initialiser les données
    InitializeVariables

    ' Initialiser les favoris
    Set Favorites = New Collection
    LoadFavorites

    ' Créer la liste plate des variables
    CreateFlatVariableList

    ' Configurer l'interface
    SetupUI

    ' Afficher les catégories
    PopulateCategories

    ' Afficher les favoris
    UpdateFavoritesDisplay
End Sub

' --------------------------------------------------------------------------
' Configuration de l'interface utilisateur
' --------------------------------------------------------------------------
Private Sub SetupUI()
    ' Configuration de la ListBox des catégories
    With lstCategories
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "150;40"
    End With

    ' Configuration de la ListBox des variables
    With lstVariables
        .Clear
        .ColumnCount = 3
        .ColumnWidths = "180;25;200"
    End With

    ' Configuration de la ListBox des favoris
    With lstFavorites
        .Clear
        .ColumnCount = 3
        .ColumnWidths = "180;25;150"
    End With

    CurrentFilter = ""
End Sub

' --------------------------------------------------------------------------
' Remplir la liste des catégories
' --------------------------------------------------------------------------
Private Sub PopulateCategories()
    Dim i As Long

    lstCategories.Clear
    For i = 1 To 14
        lstCategories.AddItem Categories(i).Name
        lstCategories.List(lstCategories.ListCount - 1, 1) = Categories(i).Count
    Next i

    ' Sélectionner la première catégorie
    If lstCategories.ListCount > 0 Then
        lstCategories.ListIndex = 0
    End If
End Sub

' --------------------------------------------------------------------------
' Afficher les variables de la catégorie sélectionnée
' --------------------------------------------------------------------------
Private Sub lstCategories_Click()
    Dim catIndex As Long
    Dim i As Long
    Dim v As TDVariable

    If lstCategories.ListIndex < 0 Then Exit Sub

    catIndex = lstCategories.ListIndex + 1

    lstVariables.Clear

    For i = 1 To Categories(catIndex).Count
        v = Categories(catIndex).Variables(i)

        ' Appliquer le filtre si actif
        If CurrentFilter = "" Or MatchesFilter(v, CurrentFilter, Categories(catIndex).Name) Then
            lstVariables.AddItem v.Placeholder
            lstVariables.List(lstVariables.ListCount - 1, 1) = v.VarType
            lstVariables.List(lstVariables.ListCount - 1, 2) = v.Description
        End If
    Next i

    lblVariablesCount.Caption = lstVariables.ListCount & " variable(s)"
End Sub

' --------------------------------------------------------------------------
' Recherche en temps réel
' --------------------------------------------------------------------------
Private Sub txtSearch_Change()
    CurrentFilter = Trim(txtSearch.Text)

    ' Rafraîchir l'affichage
    If lstCategories.ListIndex >= 0 Then
        lstCategories_Click
    End If

    ' Rafraîchir les favoris
    UpdateFavoritesDisplay
End Sub

' --------------------------------------------------------------------------
' Vérifier si une variable correspond au filtre
' --------------------------------------------------------------------------
Private Function MatchesFilter(v As TDVariable, filter As String, catName As String) As Boolean
    Dim lowerFilter As String
    lowerFilter = LCase(filter)

    MatchesFilter = (InStr(1, LCase(v.Placeholder), lowerFilter) > 0) Or _
                    (InStr(1, LCase(v.Description), lowerFilter) > 0) Or _
                    (InStr(1, LCase(catName), lowerFilter) > 0)
End Function

' --------------------------------------------------------------------------
' Créer la liste plate de toutes les variables
' --------------------------------------------------------------------------
Private Sub CreateFlatVariableList()
    Dim i As Long, j As Long

    Set AllVariablesFlat = New Collection

    For i = 1 To 14
        For j = 1 To Categories(i).Count
            AllVariablesFlat.Add Categories(i).Variables(j).Placeholder
        Next j
    Next i
End Sub

' --------------------------------------------------------------------------
' Double-clic sur une variable pour l'insérer
' --------------------------------------------------------------------------
Private Sub lstVariables_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    InsertSelectedVariable
End Sub

' --------------------------------------------------------------------------
' Double-clic sur un favori pour l'insérer
' --------------------------------------------------------------------------
Private Sub lstFavorites_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    InsertSelectedFavorite
End Sub

' --------------------------------------------------------------------------
' Bouton Insérer
' --------------------------------------------------------------------------
Private Sub btnInsert_Click()
    ' Vérifier si on insère depuis les variables ou les favoris
    If lstVariables.ListIndex >= 0 Then
        InsertSelectedVariable
    ElseIf lstFavorites.ListIndex >= 0 Then
        InsertSelectedFavorite
    Else
        MsgBox "Veuillez sélectionner une variable.", vbInformation
    End If
End Sub

' --------------------------------------------------------------------------
' Insérer la variable sélectionnée
' --------------------------------------------------------------------------
Private Sub InsertSelectedVariable()
    If lstVariables.ListIndex < 0 Then
        MsgBox "Veuillez sélectionner une variable.", vbInformation
        Exit Sub
    End If

    Dim placeholder As String
    placeholder = lstVariables.List(lstVariables.ListIndex, 0)

    InsertTextAtCursor placeholder
End Sub

' --------------------------------------------------------------------------
' Insérer le favori sélectionné
' --------------------------------------------------------------------------
Private Sub InsertSelectedFavorite()
    If lstFavorites.ListIndex < 0 Then
        MsgBox "Veuillez sélectionner un favori.", vbInformation
        Exit Sub
    End If

    Dim placeholder As String
    placeholder = lstFavorites.List(lstFavorites.ListIndex, 0)

    InsertTextAtCursor placeholder
End Sub

' --------------------------------------------------------------------------
' Insérer le texte à la position du curseur
' --------------------------------------------------------------------------
Private Sub InsertTextAtCursor(text As String)
    On Error Resume Next
    Selection.TypeText text:=text

    If Err.Number <> 0 Then
        MsgBox "Erreur lors de l'insertion: " & Err.Description, vbExclamation
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' --------------------------------------------------------------------------
' Ajouter aux favoris
' --------------------------------------------------------------------------
Private Sub btnAddFavorite_Click()
    If lstVariables.ListIndex < 0 Then
        MsgBox "Veuillez sélectionner une variable.", vbInformation
        Exit Sub
    End If

    Dim placeholder As String
    placeholder = lstVariables.List(lstVariables.ListIndex, 0)

    ' Vérifier si déjà en favori
    If IsFavorite(placeholder) Then
        MsgBox "Cette variable est déjà dans vos favoris.", vbInformation
        Exit Sub
    End If

    ' Ajouter aux favoris
    Favorites.Add placeholder, placeholder
    SaveFavorites
    UpdateFavoritesDisplay

    MsgBox "Ajouté aux favoris!", vbInformation
End Sub

' --------------------------------------------------------------------------
' Retirer des favoris
' --------------------------------------------------------------------------
Private Sub btnRemoveFavorite_Click()
    If lstFavorites.ListIndex < 0 Then
        MsgBox "Veuillez sélectionner un favori à retirer.", vbInformation
        Exit Sub
    End If

    Dim placeholder As String
    placeholder = lstFavorites.List(lstFavorites.ListIndex, 0)

    ' Retirer des favoris
    On Error Resume Next
    Favorites.Remove placeholder
    On Error GoTo 0

    SaveFavorites
    UpdateFavoritesDisplay
End Sub

' --------------------------------------------------------------------------
' Vérifier si une variable est en favori
' --------------------------------------------------------------------------
Private Function IsFavorite(placeholder As String) As Boolean
    Dim item As Variant
    On Error Resume Next
    item = Favorites(placeholder)
    IsFavorite = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' --------------------------------------------------------------------------
' Mettre à jour l'affichage des favoris
' --------------------------------------------------------------------------
Private Sub UpdateFavoritesDisplay()
    Dim i As Long, j As Long
    Dim placeholder As String
    Dim v As TDVariable
    Dim found As Boolean

    lstFavorites.Clear

    On Error Resume Next
    For Each placeholder In Favorites
        ' Appliquer le filtre
        If CurrentFilter <> "" Then
            found = False
            ' Chercher la variable pour avoir sa description
            For i = 1 To 14
                For j = 1 To Categories(i).Count
                    If Categories(i).Variables(j).Placeholder = placeholder Then
                        v = Categories(i).Variables(j)
                        If MatchesFilter(v, CurrentFilter, Categories(i).Name) Then
                            found = True
                        End If
                        Exit For
                    End If
                Next j
                If found Then Exit For
            Next i

            If Not found Then GoTo NextFav
        End If

        ' Chercher la variable pour avoir son type et description
        For i = 1 To 14
            For j = 1 To Categories(i).Count
                If Categories(i).Variables(j).Placeholder = placeholder Then
                    v = Categories(i).Variables(j)
                    lstFavorites.AddItem placeholder
                    lstFavorites.List(lstFavorites.ListCount - 1, 1) = v.VarType
                    lstFavorites.List(lstFavorites.ListCount - 1, 2) = v.Description
                    Exit For
                End If
            Next j
        Next i
NextFav:
    Next placeholder
    On Error GoTo 0

    lblFavoritesCount.Caption = lstFavorites.ListCount & " favori(s)"

    ' Afficher/masquer la section favoris
    lblFavorites.Visible = (lstFavorites.ListCount > 0)
    lstFavorites.Visible = (lstFavorites.ListCount > 0)
    lblFavoritesCount.Visible = (lstFavorites.ListCount > 0)
    btnRemoveFavorite.Visible = (lstFavorites.ListCount > 0)
End Sub

' --------------------------------------------------------------------------
' Sauvegarder les favoris dans le registre
' --------------------------------------------------------------------------
Private Sub SaveFavorites()
    Dim favStr As String
    Dim placeholder As Variant

    favStr = ""
    On Error Resume Next
    For Each placeholder In Favorites
        If favStr <> "" Then favStr = favStr & "|"
        favStr = favStr & placeholder
    Next placeholder
    On Error GoTo 0

    SaveSetting "TDPrint", "Settings", FAVORITES_KEY, favStr
End Sub

' --------------------------------------------------------------------------
' Charger les favoris depuis le registre
' --------------------------------------------------------------------------
Private Sub LoadFavorites()
    Dim favStr As String
    Dim favArray() As String
    Dim i As Long

    On Error Resume Next
    favStr = GetSetting("TDPrint", "Settings", FAVORITES_KEY, "")
    On Error GoTo 0

    Set Favorites = New Collection

    If favStr <> "" Then
        favArray = Split(favStr, "|")
        For i = LBound(favArray) To UBound(favArray)
            If favArray(i) <> "" Then
                On Error Resume Next
                Favorites.Add favArray(i), favArray(i)
                On Error GoTo 0
            End If
        Next i
    End If
End Sub

' --------------------------------------------------------------------------
' Effacer la recherche
' --------------------------------------------------------------------------
Private Sub btnClearSearch_Click()
    txtSearch.Text = ""
    txtSearch.SetFocus
End Sub

' --------------------------------------------------------------------------
' Fermer le formulaire
' --------------------------------------------------------------------------
Private Sub btnClose_Click()
    Unload Me
End Sub

' --------------------------------------------------------------------------
' Désélectionner quand on clique sur l'autre liste
' --------------------------------------------------------------------------
Private Sub lstVariables_Click()
    lstFavorites.ListIndex = -1
End Sub

Private Sub lstFavorites_Click()
    lstVariables.ListIndex = -1
End Sub
