VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmAddProduct 
   Caption         =   "Add Product"
   ClientHeight    =   8660.001
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "FrmAddProduct.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmAddProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=============================================================================
' UserForm Module: frmAddProduct
' Author:          Evan Scott
' Date:            April 28, 2025
' Purpose:         Provides a user interface for entering details of a new
'                  product, including its base information and a list of
'                  nutrient quantities per serving. Handles validation and
'                  saving the new product data via modProductData.
'=============================================================================

' --- Form-Level Variables ---
' Optional: Use a collection or dictionary to manage the nutrient quantities added to the listbox
' Private mAddedNutrients As Collection ' Example

' --- Form Events ---

Private Sub UserForm_Initialize()
    ' Runs when the form loads, before it's shown.
    ' Populates controls like the Nutrient ComboBox and sets initial state.
    
    ' Clear existing items (if any)
    cboAddNutrient.Clear
    lstNutrients.Clear
    
    ' Populate Nutrient ComboBox
    Dim allNutrients As Collection
    Dim nutr As nutrient ' Use the correct class name
    
    On Error Resume Next ' Handle error if repository not initialized
    Set allNutrients = ModNutrientRepository.GetAllNutrients
    On Error GoTo 0 ' Restore error handling
    
    If allNutrients Is Nothing Then
        MsgBox "Error: Nutrient repository not initialized. Cannot populate nutrient list.", vbCritical, "Initialization Error"
        ' Consider disabling nutrient controls or unloading form
        Exit Sub
    End If
    
    If allNutrients.Count > 0 Then
        ' Add "Select..." as the first item
        cboAddNutrient.AddItem "(Select Nutrient)"
        cboAddNutrient.List(0, 1) = 0 ' Store 0 as ID for the placeholder
        
        ' Loop through nutrients and add Name to ComboBox, store ID
        For Each nutr In allNutrients
            cboAddNutrient.AddItem nutr.Name ' Column 0 (visible)
            ' Store the Nutrient ID in the second column (hidden by default if ColumnCount=1)
            ' We need to ensure ColumnCount is at least 2 if we want to retrieve this later
            cboAddNutrient.List(cboAddNutrient.ListCount - 1, 1) = nutr.id
        Next nutr
        cboAddNutrient.ColumnCount = 2 ' Make sure we can access the ID later
        cboAddNutrient.BoundColumn = 2 ' Optional: Makes .Value return the ID (Column 2)
        cboAddNutrient.listIndex = 0 ' Select the "(Select Nutrient)" item
    Else
        MsgBox "Warning: No nutrients found in the repository.", vbExclamation
        ' Disable nutrient controls?
    End If

    ' Initialize internal collection if using one
    ' Set mAddedNutrients = New Collection
    
    ' Set initial focus (optional)
    txtProductID.SetFocus
    
    ' Configure ListBox columns
    With lstNutrients
        .ColumnCount = 3 ' e.g., Name, Mass/Serving, NutrientID (hidden)
        .ColumnWidths = "120;60;0" ' Hide the 3rd column (NutrientID)
        ' Add headers if desired (requires Labels positioned above ListBox or other techniques)
    End With

    Set allNutrients = Nothing
    Set nutr = Nothing
End Sub


' --- Control Events ---

Private Sub btnAddNutrientToList_Click()
    ' Handles adding the selected nutrient and its mass/serving to the listbox.
    Dim nutrientID As Long
    Dim nutrientName As String
    Dim massKg As Double
    Dim listIndex As Long
    
    ' --- Validation ---
    If cboAddNutrient.listIndex <= 0 Then ' Check if a real nutrient is selected (index 0 is placeholder)
        MsgBox "Please select a nutrient from the list.", vbExclamation, "Input Required"
        cboAddNutrient.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtAddMassPerServing.Text) Then
        MsgBox "Please enter a valid numeric value for Mass (kg).", vbExclamation, "Invalid Input"
        txtAddMassPerServing.SetFocus
        Exit Sub
    End If
    
    massKg = CDbl(txtAddMassPerServing.Text)
    
    If massKg <= 0 Then
        MsgBox "Mass per serving must be a positive value.", vbExclamation, "Invalid Input"
        txtAddMassPerServing.SetFocus
        Exit Sub
    End If
    
    ' Get selected nutrient details
    listIndex = cboAddNutrient.listIndex
    nutrientName = cboAddNutrient.List(listIndex, 0)
    nutrientID = CLng(cboAddNutrient.List(listIndex, 1))
    
    ' TODO: Check if this nutrientID is already in lstNutrients? Prevent duplicates? Or allow update?
    ' For now, we'll just add it.

    ' --- Add to ListBox ---
    With lstNutrients
        .AddItem nutrientName ' Add item, creates new row
        .List(.ListCount - 1, 1) = Format(massKg, "0.000000") ' Column index 1 (0-based) - Format for display consistency
        .List(.ListCount - 1, 2) = nutrientID ' Column index 2 (hidden) - Store the ID
    End With
    
    ' TODO: Add to internal collection/dictionary if using one (for easier management than reading ListBox)

    ' --- Clear inputs for next entry ---
    cboAddNutrient.listIndex = 0
    txtAddMassPerServing.Text = ""
    cboAddNutrient.SetFocus ' Set focus back to combo box
    
End Sub

Private Sub btnRemoveNutrient_Click()
    ' Handles removing the selected nutrient from the listbox.
    Dim selectedIndex As Long
    selectedIndex = lstNutrients.listIndex ' Get index of selected item (-1 if none)

    If selectedIndex < 0 Then
        MsgBox "Please select a nutrient from the list to remove.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    ' TODO: Remove from internal collection/dictionary if using one before removing from ListBox

    ' Remove from ListBox
    lstNutrients.RemoveItem selectedIndex
    
End Sub


Private Sub btnSave_Click()
    ' Handles validation of all inputs, creation of Product and NutrientQuantity
    ' objects, and calling the SaveProduct subroutine.
    Dim newProd As Product
    Dim nq As NutrientQuantity
    Dim i As Long
    Dim isValid As Boolean
    
    ' --- Validate Product Details ---
    isValid = True ' Assume valid initially
    
    ' Example Validation (Add more checks as needed - e.g., check if Product ID already exists)
    If Not IsNumeric(txtProductID.Text) Or CLng(txtProductID.Text) <= 0 Then
        MsgBox "Please enter a valid positive Product ID.", vbExclamation, "Invalid Input": isValid = False: txtProductID.SetFocus
    ElseIf Not IsNumeric(txtPrice.Text) Or CCur(txtPrice.Text) < 0 Then
        MsgBox "Please enter a valid non-negative Price.", vbExclamation, "Invalid Input": isValid = False: txtPrice.SetFocus
    ElseIf Not IsNumeric(txtMass.Text) Or CDbl(txtMass.Text) <= 0 Then
        MsgBox "Please enter a valid positive Total Mass (kg).", vbExclamation, "Invalid Input": isValid = False: txtMass.SetFocus
    ElseIf Not IsNumeric(txtServings.Text) Or CLng(txtServings.Text) <= 0 Then
        MsgBox "Please enter a valid positive number of Servings.", vbExclamation, "Invalid Input": isValid = False: txtServings.SetFocus
    ElseIf Trim$(txtProductName.Text) = "" Then
        MsgBox "Please enter a Product Name.", vbExclamation, "Invalid Input": isValid = False: txtProductName.SetFocus
    ElseIf lstNutrients.ListCount = 0 Then
         MsgBox "Please add at least one nutrient quantity.", vbExclamation, "Invalid Input": isValid = False: cboAddNutrient.SetFocus
    End If
    
    If Not isValid Then Exit Sub ' Stop if validation failed
    
    ' --- Create and Populate Product Object ---
    On Error GoTo SaveErrorHandler
    
    Set newProd = New Product
    newProd.id = CLng(txtProductID.Text)
    newProd.ProductName = Trim$(txtProductName.Text)
    newProd.price = CCur(txtPrice.Text)
    newProd.mass = CDbl(txtMass.Text)
    newProd.servings = CLng(txtServings.Text)
    
    ' --- Add Nutrient Quantities ---
    If lstNutrients.ListCount > 0 Then
        For i = 0 To lstNutrients.ListCount - 1 ' ListBox is 0-based index
            Set nq = New NutrientQuantity
            nq.nutrientID = CLng(lstNutrients.List(i, 2)) ' Get ID from hidden column 2
            nq.MassPerServing = CDbl(lstNutrients.List(i, 1)) ' Get mass from column 1
            newProd.NutrientQuantities.Add nq
            Set nq = Nothing ' Release reference
        Next i
    End If
    
    ' --- Call Save Function ---
    ' Assumes PRODUCT_DATA_SHEET_NAME constant exists or use literal name
    ' Ensure the sheet name constant/literal matches the one used in modSetup and modProductData
    ModProductData.SaveProduct newProd, ThisWorkbook.Worksheets(PRODUCT_DATA_SHEET_NAME)
    
    ' --- Success ---
    MsgBox "Product '" & newProd.ProductName & "' saved successfully!", vbInformation, "Save Successful"
    Unload Me ' Close the form
    
    Set newProd = Nothing
    Exit Sub

SaveErrorHandler:
    MsgBox "An error occurred while saving the product:" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical, "Save Error"
    Set newProd = Nothing
    Set nq = Nothing
    
End Sub

Private Sub btnCancel_Click()
    ' Unload the form without saving changes.
    Unload Me
End Sub



