VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmUpdateProduct 
   Caption         =   "UserForm1"
   ClientHeight    =   10470
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "FrmUpdateProduct.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmUpdateProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=============================================================================
' UserForm Module: frmUpdateProduct
' Author:          Evan Scott
' Date:            April 28, 2025
' Purpose:         Provides a user interface for loading, viewing, editing,
'                  and saving changes to an existing product's details and
'                  nutrient quantities.
'=============================================================================

' --- Form-Level Variables ---
' Store the currently loaded product ID to prevent accidental changes
Private mLoadedProductID As Long
Private mIsLoaded As Boolean ' Flag to track if a product is currently loaded

' --- Form Events ---

Private Sub UserForm_Initialize()
    ' Runs when the form loads, before it's shown.
    ' Sets the initial state (most controls disabled) and populates the nutrient combo box.
    
    mLoadedProductID = 0 ' Reset loaded ID
    mIsLoaded = False
    
    ' Set initial control states
    Call EnableEditingControls(False) ' Disable editing sections initially
    txtLoadID.Enabled = True
    btnLoadProduct.Enabled = True
    btnCancel.Enabled = True ' Cancel should always be enabled
    
    ' Clear lists/fields
    txtLoadID.Text = ""
    txtProductID.Text = ""
    txtProductName.Text = ""
    txtPrice.Text = ""
    txtMass.Text = ""
    txtServings.Text = ""
    cboAddNutrient.Clear
    lstNutrients.Clear
    txtAddMassPerServing.Text = ""
    
    ' Lock the Product ID field on this form as we don't allow changing the ID
    txtProductID.Locked = True
    
    ' Populate Nutrient ComboBox
    Dim allNutrients As Collection
    Dim nutr As nutrient ' Use the correct class name
    
    On Error Resume Next ' Handle error if repository not initialized
    Set allNutrients = ModNutrientRepository.GetAllNutrients
    On Error GoTo 0 ' Restore error handling
    
    If allNutrients Is Nothing Then
        MsgBox "Error: Nutrient repository not initialized. Cannot populate nutrient list.", vbCritical, "Initialization Error"
        Exit Sub ' Or disable nutrient controls
    End If
    
    If allNutrients.Count > 0 Then
        cboAddNutrient.AddItem "(Select Nutrient)"
        cboAddNutrient.List(0, 1) = 0 ' Placeholder ID
        For Each nutr In allNutrients
            cboAddNutrient.AddItem nutr.Name
            cboAddNutrient.List(cboAddNutrient.ListCount - 1, 1) = nutr.id
        Next nutr
        cboAddNutrient.ColumnCount = 2
        cboAddNutrient.BoundColumn = 2 ' Makes .Value return ID
        cboAddNutrient.listIndex = 0
    Else
        MsgBox "Warning: No nutrients found in the repository.", vbExclamation
    End If

    ' Configure ListBox columns
    With lstNutrients
        .ColumnCount = 3 ' Name, Mass/Serving, NutrientID (hidden)
        .ColumnWidths = "120;60;0" ' Hide the 3rd column
    End With

    Set allNutrients = Nothing
    Set nutr = Nothing
    
    txtLoadID.SetFocus ' Set focus to the ID input field
End Sub

' --- Helper Sub to Enable/Disable Editing Controls ---
Private Sub EnableEditingControls(enableState As Boolean)
    ' Product Details
    'txtProductID.Enabled = enableState ' Keep ID disabled/locked
    txtProductName.Enabled = enableState
    txtPrice.Enabled = enableState
    txtMass.Enabled = enableState
    txtServings.Enabled = enableState
    
    ' Nutrient Section
    lstNutrients.Enabled = enableState
    cboAddNutrient.Enabled = enableState
    txtAddMassPerServing.Enabled = enableState
    btnAddNutrientToList.Enabled = enableState
    btnRemoveNutrient.Enabled = enableState
    
    ' Save Button
    btnSaveChanges.Enabled = enableState
End Sub

' --- Control Events ---

Private Sub btnLoadProduct_Click()
    Dim productIDToLoad As Long
    Dim loadedProd As Product
    Dim nq As NutrientQuantity
    Dim i As Long
    
    ' Validate ID input
    If Not IsNumeric(txtLoadID.Text) Or CLng(txtLoadID.Text) <= 0 Then
        MsgBox "Please enter a valid positive Product ID to load.", vbExclamation, "Invalid ID"
        txtLoadID.SetFocus
        Exit Sub
    End If
    
    productIDToLoad = CLng(txtLoadID.Text)
    
    ' Attempt to load the product
    On Error Resume Next ' Handle potential errors during load
    Set loadedProd = ModProductData.LoadProduct(productIDToLoad, ThisWorkbook.Worksheets(PRODUCT_DATA_SHEET_NAME))
    On Error GoTo 0 ' Restore error handling
    
    If loadedProd Is Nothing Then
        MsgBox "Product ID " & productIDToLoad & " not found.", vbExclamation, "Load Failed"
        Call EnableEditingControls(False) ' Keep controls disabled
        ' Clear potentially stale data if a previous load was done
        txtProductID.Text = ""
        txtProductName.Text = ""
        txtPrice.Text = ""
        txtMass.Text = ""
        txtServings.Text = ""
        lstNutrients.Clear
        mIsLoaded = False
        mLoadedProductID = 0
        txtLoadID.SetFocus
    Else
        ' Product loaded successfully - Populate the form
        mLoadedProductID = loadedProd.id ' Store the loaded ID
        mIsLoaded = True
        
        txtProductID.Text = CStr(loadedProd.id)
        txtProductName.Text = loadedProd.ProductName
        txtPrice.Text = Format(loadedProd.price, "0.00") ' Format currency
        txtMass.Text = CStr(loadedProd.mass)
        txtServings.Text = CStr(loadedProd.servings)
        
        ' Populate Nutrient ListBox
        lstNutrients.Clear
        If Not loadedProd.NutrientQuantities Is Nothing Then
            If loadedProd.NutrientQuantities.Count > 0 Then
                Dim nutrientName As String
                For Each nq In loadedProd.NutrientQuantities
                    ' Look up nutrient name - requires GetNutrientByID
                    Dim tempNutrient As nutrient
                    Set tempNutrient = ModNutrientRepository.GetNutrientByID(nq.nutrientID)
                    If Not tempNutrient Is Nothing Then
                        nutrientName = tempNutrient.Name
                    Else
                        nutrientName = "ID: " & nq.nutrientID & " (Name N/A)" ' Fallback if name lookup fails
                    End If
                    Set tempNutrient = Nothing
                    
                    With lstNutrients
                        .AddItem nutrientName
                        .List(.ListCount - 1, 1) = Format(nq.MassPerServing, "0.000000")
                        .List(.ListCount - 1, 2) = nq.nutrientID
                    End With
                Next nq
            End If
        End If
        
        ' Enable editing controls
        Call EnableEditingControls(True)
        txtProductName.SetFocus ' Set focus to first editable field
        
        MsgBox "Product ID " & productIDToLoad & " loaded successfully.", vbInformation, "Load Successful"
    End If
    
    Set loadedProd = Nothing
    Set nq = Nothing
End Sub


Private Sub btnAddNutrientToList_Click()
    ' Handles adding/updating a nutrient in the listbox for the loaded product.
    ' Similar logic to frmAddProduct, but might need adjustments if updating is allowed.
    ' For simplicity, let's assume it just adds new ones for now.
    Dim nutrientID As Long
    Dim nutrientName As String
    Dim massKg As Double
    Dim listIndex As Long
    
    ' --- Validation ---
    If Not mIsLoaded Then Exit Sub ' Should not be clickable if not loaded, but safety check
    
    If cboAddNutrient.listIndex <= 0 Then
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
    
    ' --- Check for Duplicates ---
    Dim i As Long
    Dim alreadyExists As Boolean: alreadyExists = False
    For i = 0 To lstNutrients.ListCount - 1
        If CLng(lstNutrients.List(i, 2)) = nutrientID Then
            alreadyExists = True
            Exit For
        End If
    Next i
    
    If alreadyExists Then
        MsgBox "Nutrient '" & nutrientName & "' is already in the list. Remove it first if you want to change the amount.", vbInformation, "Duplicate Nutrient"
        Exit Sub ' Prevent adding duplicates for now
    End If

    ' --- Add to ListBox ---
    With lstNutrients
        .AddItem nutrientName
        .List(.ListCount - 1, 1) = Format(massKg, "0.000000")
        .List(.ListCount - 1, 2) = nutrientID
    End With
    
    ' --- Clear inputs ---
    cboAddNutrient.listIndex = 0
    txtAddMassPerServing.Text = ""
    cboAddNutrient.SetFocus
    
End Sub

Private Sub btnRemoveNutrient_Click()
    ' Handles removing the selected nutrient from the listbox.
    Dim selectedIndex As Long
    selectedIndex = lstNutrients.listIndex

    If Not mIsLoaded Then Exit Sub ' Safety check

    If selectedIndex < 0 Then
        MsgBox "Please select a nutrient from the list to remove.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    lstNutrients.RemoveItem selectedIndex
End Sub


Private Sub btnSaveChanges_Click()
    ' Handles validation, deleting the old product data, and saving the updated product data.
    Dim updatedProd As Product
    Dim nq As NutrientQuantity
    Dim i As Long
    Dim isValid As Boolean
    
    If Not mIsLoaded Or mLoadedProductID <= 0 Then
        MsgBox "No product is currently loaded for saving.", vbCritical, "Save Error"
        Exit Sub
    End If
    
    ' --- Validate Product Details ---
    isValid = True ' Assume valid initially
    
    ' Note: ID (mLoadedProductID) is already validated during load and is read-only
    If Not IsNumeric(txtPrice.Text) Or CCur(txtPrice.Text) < 0 Then
        MsgBox "Please enter a valid non-negative Price.", vbExclamation, "Invalid Input": isValid = False: txtPrice.SetFocus
    ElseIf Not IsNumeric(txtMass.Text) Or CDbl(txtMass.Text) <= 0 Then
        MsgBox "Please enter a valid positive Total Mass (kg).", vbExclamation, "Invalid Input": isValid = False: txtMass.SetFocus
    ElseIf Not IsNumeric(txtServings.Text) Or CLng(txtServings.Text) <= 0 Then
        MsgBox "Please enter a valid positive number of Servings.", vbExclamation, "Invalid Input": isValid = False: txtServings.SetFocus
    ElseIf Trim$(txtProductName.Text) = "" Then
        MsgBox "Please enter a Product Name.", vbExclamation, "Invalid Input": isValid = False: txtProductName.SetFocus
    ElseIf lstNutrients.ListCount = 0 Then
         MsgBox "Product must have at least one nutrient quantity.", vbExclamation, "Invalid Input": isValid = False: cboAddNutrient.SetFocus
    End If
    
    If Not isValid Then Exit Sub ' Stop if validation failed
    
    ' --- Create and Populate Updated Product Object ---
    On Error GoTo SaveErrorHandler
    
    Set updatedProd = New Product
    updatedProd.id = mLoadedProductID ' Use the ID that was loaded
    updatedProd.ProductName = Trim$(txtProductName.Text)
    updatedProd.price = CCur(txtPrice.Text)
    updatedProd.mass = CDbl(txtMass.Text)
    updatedProd.servings = CLng(txtServings.Text)
    
    ' --- Add Nutrient Quantities from ListBox ---
    If lstNutrients.ListCount > 0 Then
        For i = 0 To lstNutrients.ListCount - 1
            Set nq = New NutrientQuantity
            nq.nutrientID = CLng(lstNutrients.List(i, 2)) ' Get ID from hidden column 2
            nq.MassPerServing = CDbl(lstNutrients.List(i, 1)) ' Get mass from column 1
            updatedProd.NutrientQuantities.Add nq
            Set nq = Nothing
        Next i
    End If
    
    ' --- Call Delete and Save Functions (Update Strategy) ---
    Dim dataSheet As Worksheet
    Set dataSheet = ThisWorkbook.Worksheets(PRODUCT_DATA_SHEET_NAME)
    
    ModProductData.DeleteProductData mLoadedProductID, dataSheet ' Delete old rows
    ModProductData.SaveProduct updatedProd, dataSheet ' Save new rows
    
    ' --- Success ---
    MsgBox "Product ID " & mLoadedProductID & " updated successfully!", vbInformation, "Update Successful"
    Unload Me ' Close the form
    
    Set updatedProd = Nothing
    Set dataSheet = Nothing
    Exit Sub

SaveErrorHandler:
    MsgBox "An error occurred while saving the product changes:" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical, "Save Error"
    Set updatedProd = Nothing
    Set nq = Nothing
    Set dataSheet = Nothing
    
End Sub

Private Sub btnCancel_Click()
    ' Unload the form without saving changes.
    Unload Me
End Sub



