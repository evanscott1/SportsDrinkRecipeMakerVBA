VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmCreateRecipe 
   Caption         =   "Create Recipe"
   ClientHeight    =   8090
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "FrmCreateRecipe.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmCreateRecipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=============================================================================
' UserForm Module: FrmCreateRecipe
' Author:          Evan Scott
' Date:            April 28, 2025
' Purpose:         Provides a user interface for defining nutrient targets
'                  and selecting products to exclude, then generating a recipe
'                  using modRecipeGenerator and displaying the results.
'=============================================================================

' --- Form-Level Variables ---
' Store loaded products to avoid reloading repeatedly
Private mAvailableProducts As Collection

' --- Form Events ---

Private Sub UserForm_Initialize()
    ' Runs when the form loads, before it's shown.
    ' Populates Nutrient combo box, Available Products listbox, and clears targets.
    
    ' Clear controls
    cboSelectNutrient.Clear
    lstTargetNutrients.Clear
    lstAvailableProducts.Clear
    txtTargetAmount.Text = ""
    
    ' --- Populate Nutrient ComboBox ---
    Dim allNutrients As Collection
    Dim nutr As nutrient
    On Error Resume Next ' Handle error if repository not initialized
    Set allNutrients = ModNutrientRepository.GetAllNutrients
    On Error GoTo 0 ' Restore error handling
    
    If allNutrients Is Nothing Then
        MsgBox "Error: Nutrient repository not initialized. Cannot populate nutrient list.", vbCritical, "Initialization Error"
        Exit Sub
    End If
    
    If allNutrients.Count > 0 Then
        cboSelectNutrient.AddItem "(Select Nutrient)"
        cboSelectNutrient.List(0, 1) = 0 ' Placeholder ID
        For Each nutr In allNutrients
            cboSelectNutrient.AddItem nutr.Name
            cboSelectNutrient.List(cboSelectNutrient.ListCount - 1, 1) = nutr.id
        Next nutr
        cboSelectNutrient.ColumnCount = 2
        cboSelectNutrient.BoundColumn = 2 ' Makes .Value return ID
        cboSelectNutrient.listIndex = 0
    Else
        MsgBox "Warning: No nutrients found in the repository.", vbExclamation
    End If
    Set allNutrients = Nothing
    Set nutr = Nothing

    ' --- Populate Available Products ListBox ---
    ' Requires a function like GetAllProducts in modProductData
    On Error Resume Next ' Handle errors during product loading
    ' *** YOU NEED TO IMPLEMENT GetAllProducts in modProductData ***
    Set mAvailableProducts = ModProductData.GetAllProducts(ThisWorkbook.Worksheets(PRODUCT_DATA_SHEET_NAME))
    On Error GoTo 0 ' Restore error handling
    
    If mAvailableProducts Is Nothing Then
         MsgBox "Warning: Could not load available products. Please ensure products exist on the '" & PRODUCT_DATA_SHEET_NAME & "' sheet.", vbExclamation
         Set mAvailableProducts = New Collection ' Create empty collection to avoid errors later
    ElseIf mAvailableProducts.Count = 0 Then
         MsgBox "Warning: No products found on the '" & PRODUCT_DATA_SHEET_NAME & "' sheet.", vbExclamation
    Else
        Dim prod As Product
        For Each prod In mAvailableProducts
            With lstAvailableProducts
                .AddItem prod.ProductName ' Column 0
                .List(.ListCount - 1, 1) = prod.id ' Column 1 (hidden) - Store ID
            End With
        Next prod
    End If
    ' Configure ListBox
    With lstAvailableProducts
        .ColumnCount = 2
        .ColumnWidths = "150;0" ' Hide ID column
        .MultiSelect = fmMultiSelectMulti ' Allow multiple selections for exclusion
    End With
    Set prod = Nothing

    ' --- Configure Target Nutrients ListBox ---
    With lstTargetNutrients
        .ColumnCount = 3 ' Name, Target Amount, NutrientID (hidden)
        .ColumnWidths = "120;70;0" ' Hide the 3rd column
    End With
    
    cboSelectNutrient.SetFocus ' Initial focus
End Sub

' --- Control Events ---

Private Sub btnAddTarget_Click()
    ' Adds the selected nutrient and target amount to the lstTargetNutrients listbox.
    Dim nutrientID As Long
    Dim nutrientName As String
    Dim targetAmountKg As Double
    Dim listIndex As Long
    Dim i As Long
    Dim alreadyExists As Boolean
    
    ' --- Validation ---
    If cboSelectNutrient.listIndex <= 0 Then
        MsgBox "Please select a nutrient.", vbExclamation, "Input Required": cboSelectNutrient.SetFocus: Exit Sub
    End If
    If Not IsNumeric(txtTargetAmount.Text) Then
        MsgBox "Please enter a valid numeric Target Amount (kg).", vbExclamation, "Invalid Input": txtTargetAmount.SetFocus: Exit Sub
    End If
    targetAmountKg = CDbl(txtTargetAmount.Text)
    If targetAmountKg <= 0 Then
        MsgBox "Target Amount must be a positive value.", vbExclamation, "Invalid Input": txtTargetAmount.SetFocus: Exit Sub
    End If
    
    ' Get selected nutrient details
    listIndex = cboSelectNutrient.listIndex
    nutrientName = cboSelectNutrient.List(listIndex, 0)
    nutrientID = CLng(cboSelectNutrient.List(listIndex, 1))
    
    ' --- Check if target for this nutrient already exists ---
    alreadyExists = False
    For i = 0 To lstTargetNutrients.ListCount - 1
        If CLng(lstTargetNutrients.List(i, 2)) = nutrientID Then
            alreadyExists = True
            Exit For
        End If
    Next i
    
    If alreadyExists Then
        MsgBox "A target for '" & nutrientName & "' already exists. Remove the existing target first if you want to change it.", vbInformation, "Target Exists"
        Exit Sub
    End If
    
    ' --- Add to ListBox ---
    With lstTargetNutrients
        .AddItem nutrientName ' Column 0
        .List(.ListCount - 1, 1) = Format(targetAmountKg, "0.000000") ' Column 1
        .List(.ListCount - 1, 2) = nutrientID ' Column 2 (hidden)
    End With
    
    ' --- Clear inputs ---
    cboSelectNutrient.listIndex = 0
    txtTargetAmount.Text = ""
    cboSelectNutrient.SetFocus
End Sub

Private Sub btnRemoveTarget_Click()
    ' Removes the selected target nutrient from the listbox.
    Dim selectedIndex As Long
    selectedIndex = lstTargetNutrients.listIndex

    If selectedIndex < 0 Then
        MsgBox "Please select a target nutrient from the list to remove.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    lstTargetNutrients.RemoveItem selectedIndex
End Sub


Private Sub btnGenerateRecipe_Click()
    ' Gathers inputs, calls the generator, and displays the results.
    Dim targets As Scripting.Dictionary
    Dim excludedIDs As Scripting.Dictionary
    Dim generatedRecipe As recipe
    Dim wsOutput As Worksheet
    Dim i As Long
    Dim nextRow As Long
    Dim ri As RecipeIngredient
    Dim totalCost As Currency
    
    On Error GoTo GenerateErrorHandler
    
    ' --- 1. Gather Target Nutrients ---
    Set targets = New Scripting.Dictionary
    If lstTargetNutrients.ListCount = 0 Then
        MsgBox "Please add at least one nutrient target.", vbExclamation, "Input Required"
        Exit Sub
    End If
    For i = 0 To lstTargetNutrients.ListCount - 1
        targets.Add key:=CLng(lstTargetNutrients.List(i, 2)), Item:=CDbl(lstTargetNutrients.List(i, 1))
    Next i
    
    ' --- 2. Gather Excluded Product IDs ---
    Set excludedIDs = New Scripting.Dictionary
    If lstAvailableProducts.ListCount > 0 Then
        For i = 0 To lstAvailableProducts.ListCount - 1
            If lstAvailableProducts.Selected(i) Then ' Check if item is selected
                excludedIDs.Add key:=CLng(lstAvailableProducts.List(i, 1)), Item:=True ' Add ID of selected item
            End If
        Next i
    End If
    
    ' --- 3. Gather Available Products ---
    ' We loaded this into mAvailableProducts during Initialize
    If mAvailableProducts Is Nothing Or mAvailableProducts.Count = 0 Then
         MsgBox "No available products loaded. Cannot generate recipe.", vbCritical, "Error"
         Exit Sub
    End If

    ' --- 4. Call Generator ---
    Application.Cursor = xlWait ' Indicate processing
    Set generatedRecipe = ModRecipeGenerator.GenerateRecipe(targets, mAvailableProducts, excludedIDs)
    Application.Cursor = xlDefault

    ' --- 5. Handle Output ---
    If generatedRecipe Is Nothing Then
        MsgBox "Could not generate a recipe to meet the specified targets with the available products.", vbExclamation, "Recipe Generation Failed"
    ElseIf generatedRecipe.Ingredients.Count = 0 Then
        MsgBox "Recipe generated, but contains no ingredients (targets might have been zero or already met).", vbInformation, "Empty Recipe"
    Else
        ' Recipe generated successfully, display it
        Set wsOutput = ThisWorkbook.Worksheets(RECIPE_OUTPUT_SHEET_NAME) ' Assumes constant defined
        
        wsOutput.Cells.Clear ' Clear previous results
        wsOutput.Activate
        nextRow = 1
        totalCost = 0
        
        ' Add Title
        wsOutput.Cells(nextRow, 1).value = "Generated Recipe (Single Serving)"
        wsOutput.Cells(nextRow, 1).Font.Bold = True
        wsOutput.Cells(nextRow, 1).Font.Size = 14
        nextRow = nextRow + 2
        
        ' Add Headers
        wsOutput.Cells(nextRow, 1).value = "Ingredient"
        wsOutput.Cells(nextRow, 2).value = "Servings"
        wsOutput.Cells(nextRow, 3).value = "Amount (kg)"
        wsOutput.Cells(nextRow, 4).value = "Cost"
        wsOutput.Rows(nextRow).Font.Bold = True
        nextRow = nextRow + 1
        
        ' Add Ingredient Rows
        For Each ri In generatedRecipe.Ingredients
            wsOutput.Cells(nextRow, 1).value = ri.Product.ProductName
            wsOutput.Cells(nextRow, 2).value = ri.AmountServings
            wsOutput.Cells(nextRow, 3).value = ri.AmountKg
            wsOutput.Cells(nextRow, 4).value = ri.Cost
            wsOutput.Cells(nextRow, 4).NumberFormat = "$#,##0.00" ' Format cost
            totalCost = totalCost + ri.Cost
            nextRow = nextRow + 1
        Next ri
        
        ' Add Total Cost
        nextRow = nextRow + 1
        wsOutput.Cells(nextRow, 3).value = "Total Cost (1 Serving):"
        wsOutput.Cells(nextRow, 3).Font.Bold = True
        wsOutput.Cells(nextRow, 4).value = totalCost
        wsOutput.Cells(nextRow, 4).NumberFormat = "$#,##0.00"
        wsOutput.Cells(nextRow, 4).Font.Bold = True
        
        ' Add Multi-Serving Info
        nextRow = nextRow + 2
        wsOutput.Cells(nextRow, 1).value = "To calculate for multiple servings:"
        nextRow = nextRow + 1
        wsOutput.Cells(nextRow, 1).value = "Enter desired servings:"
        wsOutput.Cells(nextRow, 2).NumberFormat = "0" ' Format for whole number input
        wsOutput.Cells(nextRow, 2).Interior.Color = vbYellow ' Highlight input cell
        Dim multiplierCellAddress As String: multiplierCellAddress = wsOutput.Cells(nextRow, 2).Address
        nextRow = nextRow + 1
        wsOutput.Cells(nextRow, 1).value = "Total Cost (Multiple Servings):"
        wsOutput.Cells(nextRow, 2).Formula = "=" & wsOutput.Cells(nextRow - 4, 4).Address & "*" & multiplierCellAddress
        wsOutput.Cells(nextRow, 2).NumberFormat = "$#,##0.00"
        wsOutput.Cells(nextRow, 2).Font.Bold = True
        
        ' AutoFit Columns
        wsOutput.Columns("A:D").AutoFit
        
        MsgBox "Recipe generated successfully and displayed on the '" & RECIPE_OUTPUT_SHEET_NAME & "' sheet.", vbInformation, "Recipe Generated"
        Unload Me ' Close the form
    End If

    GoTo Cleanup

GenerateErrorHandler:
    Application.Cursor = xlDefault
    MsgBox "An error occurred during recipe generation:" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical, "Generation Error"

Cleanup:
    Set targets = Nothing
    Set excludedIDs = Nothing
    Set generatedRecipe = Nothing
    Set wsOutput = Nothing
    Set ri = Nothing
    
End Sub


Private Sub btnCancel_Click()
    ' Unload the form without generating a recipe.
    Unload Me
End Sub



