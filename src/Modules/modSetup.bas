Attribute VB_Name = "modSetup"
Option Explicit

'=============================================================================
' Module: ModSetup
' Author: Evan Scott
' Date:   April 28, 2025
' Purpose: Contains routines to initialize the workbook environment for the
'          Recipe Generator application, intended to be run once after
'          importing code modules into a blank workbook.
'=============================================================================



'-----------------------------------------------------------------------------
' SetupWorkbookEnvironment
'-----------------------------------------------------------------------------
' Purpose: Creates necessary worksheets, adds basic controls to a dashboard,
'          and initializes data repositories. Run this ONCE after importing
'          all project code modules into a new workbook.
'-----------------------------------------------------------------------------
Public Sub SetupWorkbookEnvironment()
    Dim wsDash As Worksheet
    Dim wsData As Worksheet
    Dim wsRecipe As Worksheet
    Dim btn As OLEObject ' For ActiveX Controls
    Dim nextTop As Double ' To position controls vertically
    Dim leftMargin As Double
    
    On Error GoTo SetupErrorHandler

    Application.ScreenUpdating = False ' Turn off screen updates for faster execution

    ' --- 1. Ensure Worksheets Exist ---
    Set wsDash = GetOrCreateWorksheet(DASHBOARD_SHEET_NAME)
    Set wsData = GetOrCreateWorksheet(PRODUCT_DATA_SHEET_NAME)
    Set wsRecipe = GetOrCreateWorksheet(RECIPE_OUTPUT_SHEET_NAME)

    ' --- 2. Initialize Nutrient Repository ---
    ' Assumes modNutrientRepository and InitializeNutrientRepository sub exist
    ModNutrientRepository.InitializeNutrientRepository
    Debug.Print "Nutrient Repository Initialized."

    ' --- 3. Setup ProductData Sheet Headers ---
    ' Assumes modConstants and ProductDataColumns Enum exist
    With wsData
        If .Cells(1, ProductDataColumns.colProdId).value = "" Then ' Only add headers if row 1 is empty
            .Cells(1, ProductDataColumns.colProdId).value = "ProductID"
            .Cells(1, ProductDataColumns.colProdName).value = "ProductName"
            .Cells(1, ProductDataColumns.colProdPrice).value = "Price"
            .Cells(1, ProductDataColumns.colProdMass).value = "TotalMass"
            .Cells(1, ProductDataColumns.colProdServings).value = "Servings"
            .Cells(1, ProductDataColumns.colNutrientId).value = "NutrientID"
            .Cells(1, ProductDataColumns.colMassPerServing).value = "MassPerServing"
            .Rows(1).Font.Bold = True
            .Columns.AutoFit ' Adjust column widths
            Debug.Print "ProductData sheet headers added."
        End If
    End With
    
    ' --- 4. Setup Dashboard Sheet ---
    wsDash.Activate ' Bring dashboard to front
    wsDash.Cells.Clear ' Clear any previous content/controls
    
    ' Set initial positions/margins
    nextTop = 10
    leftMargin = 10

    ' Add Title
    With wsDash.Cells(1, 1) ' A1
        .value = "Recipe Generator Control Panel"
        .Font.Bold = True
        .Font.Size = 16
        .EntireRow.RowHeight = 25
    End With
    nextTop = wsDash.Cells(1, 1).Top + wsDash.Cells(1, 1).Height + 10 ' Position below title

    ' Add Description
    With wsDash.Range("A2:D3") ' Merge some cells for description
         .Merge
         .value = "Use the buttons below to manage products and generate recipes. Product data is stored on the '" & PRODUCT_DATA_SHEET_NAME & "' sheet. Generated recipes appear on the '" & RECIPE_OUTPUT_SHEET_NAME & "' sheet."
         .WrapText = True
         .VerticalAlignment = xlTop
         .Font.Size = 10
         .EntireRow.AutoFit
    End With
    nextTop = wsDash.Range("A2").Top + wsDash.Range("A2").Height + 15 ' Position below description

    ' Add Buttons using ActiveX Controls
    ' Note: Placeholder macro names are assigned. These subs need to be created elsewhere (e.g., in this module or dedicated UI modules).
    
    ' Add Product Button
    Set btn = AddActiveXButton(wsDash, "btnAddProduct", "Add New Product", leftMargin, nextTop, 120, 24)
    AssignMacroToActiveXButton btn, "ShowAddProductForm" ' Placeholder macro name
    nextTop = nextTop + btn.Height + 6 ' Increment position for next button

    ' Update Product Button (Placeholder - needs selection mechanism later)
    Set btn = AddActiveXButton(wsDash, "btnUpdateProduct", "Update Product (by ID)", leftMargin, nextTop, 120, 24)
    AssignMacroToActiveXButton btn, "PromptAndUpdateProduct" ' Placeholder macro name
    nextTop = nextTop + btn.Height + 6

    ' Delete Product Button (Placeholder - needs selection mechanism later)
    Set btn = AddActiveXButton(wsDash, "btnDeleteProduct", "Delete Product (by ID)", leftMargin, nextTop, 120, 24)
    AssignMacroToActiveXButton btn, "PromptAndDeleteProduct" ' Placeholder macro name
    nextTop = nextTop + btn.Height + 15 ' Add extra space before next section

    ' Create Recipe Button
    Set btn = AddActiveXButton(wsDash, "btnCreateRecipe", "Create Recipe", leftMargin, nextTop, 120, 24)
    AssignMacroToActiveXButton btn, "ShowCreateRecipeForm" ' Placeholder macro name
    nextTop = nextTop + btn.Height + 6

    ' Apply some basic formatting
    wsDash.Columns("A").ColumnWidth = 20 ' Adjust width for buttons
    wsDash.Columns("B:D").ColumnWidth = 15
    
    ' Protect sheet structure but allow interaction with controls (optional)
    ' wsDash.Protect UserInterfaceOnly:=True ' Protects cells but allows macro interaction

    Application.ScreenUpdating = True ' Turn screen updates back on
    MsgBox "Workbook environment setup complete!", vbInformation
    
Exit Sub

SetupErrorHandler:
    Application.ScreenUpdating = True ' Ensure screen updating is back on
    MsgBox "An error occurred during setup:" & vbCrLf & Err.Number & " - " & Err.Description, vbCritical, "Setup Error"
End Sub

'-----------------------------------------------------------------------------
' GetOrCreateWorksheet (Helper Function)
'-----------------------------------------------------------------------------
' Purpose: Checks if a worksheet exists, creates it if not, and returns
'          a reference to the worksheet object.
' Arguments:
'   sheetName (String): The desired name of the worksheet.
' Returns:
'   Worksheet: The worksheet object.
'-----------------------------------------------------------------------------
Private Function GetOrCreateWorksheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next ' Temporarily ignore error if sheet doesn't exist
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0 ' Turn error handling back on

    If ws Is Nothing Then ' Sheet does not exist, so create it
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
        Debug.Print "Worksheet '" & sheetName & "' created."
    Else
        Debug.Print "Worksheet '" & sheetName & "' already exists."
    End If
    
    Set GetOrCreateWorksheet = ws
End Function

'-----------------------------------------------------------------------------
' AddActiveXButton (Helper Function)
'-----------------------------------------------------------------------------
' Purpose: Adds an ActiveX Command Button to a specified worksheet.
' Arguments:
'   ws        (Worksheet): The target worksheet.
'   btnName   (String): The programmatic name for the button.
'   btnCaption(String): The text displayed on the button.
'   leftPos   (Double): The left position of the button.
'   topPos    (Double): The top position of the button.
'   btnWidth  (Double): The width of the button.
'   btnHeight (Double): The height of the button.
' Returns:
'   OLEObject: The created button object.
'-----------------------------------------------------------------------------
Private Function AddActiveXButton(ws As Worksheet, btnName As String, btnCaption As String, leftPos As Double, topPos As Double, btnWidth As Double, btnHeight As Double) As OLEObject
    Dim btn As OLEObject
    
    ' Add the button
    Set btn = ws.OLEObjects.Add(ClassType:="Forms.CommandButton.1", Link:=False, _
        DisplayAsIcon:=False, Left:=leftPos, Top:=topPos, Width:=btnWidth, Height:=btnHeight)
        
    ' Set properties
    btn.Name = btnName ' Programmatic name
    btn.Object.Caption = btnCaption ' Display text
    ' Optional: Set other properties like font
    ' btn.Object.Font.Size = 9
    
    Set AddActiveXButton = btn
End Function

'-----------------------------------------------------------------------------
' AssignMacroToActiveXButton (Helper Function)
'-----------------------------------------------------------------------------
' Purpose: Assigns a macro (subroutine name) to the Click event of an
'          ActiveX Command Button programmatically.
' Arguments:
'   btn       (OLEObject): The ActiveX button object.
'   macroName (String): The name of the Sub to run on click.
'-----------------------------------------------------------------------------
Private Sub AssignMacroToActiveXButton(btn As OLEObject, macroName As String)
    Dim codeMod As Object ' VBIDE.CodeModule - Requires reference to 'Microsoft Visual Basic for Applications Extensibility 5.3' OR late binding
    Dim lineNum As Long
    Dim procCode As String
    
    On Error Resume Next ' Handle errors, e.g., if button doesn't exist or VBE access is denied
    
    ' Define the code for the Click event handler
    procCode = "Private Sub " & btn.Name & "_Click()" & vbCrLf & _
               "    On Error Resume Next ' Basic error handling within generated code" & vbCrLf & _
               "    Application.Run """ & macroName & """" & vbCrLf & _
               "    If Err.Number <> 0 Then MsgBox ""Error running macro: "" & Err.Description, vbCritical" & vbCrLf & _
               "    On Error GoTo 0" & vbCrLf & _
               "End Sub"

    ' Access the code module of the worksheet where the button resides
    ' Note: This requires trusting access to the VBA project object model in Excel Options
    ' OR using late binding for VBIDE objects if the reference isn't added.
    Set codeMod = ThisWorkbook.VBProject.VBComponents(btn.Parent.CodeName).CodeModule

    If Not codeMod Is Nothing Then
        ' Add the event procedure code to the worksheet's code module
        lineNum = codeMod.CountOfLines + 1
        codeMod.InsertLines lineNum, procCode
        Debug.Print "Assigned macro '" & macroName & "' to button '" & btn.Name & "'."
    Else
        Debug.Print "Error: Could not access code module for sheet '" & btn.Parent.Name & "' to assign macro to '" & btn.Name & "'. Ensure VBE access is enabled."
        MsgBox "Could not assign macro to button '" & btn.Name & "'." & vbCrLf & "Please ensure 'Trust access to the VBA project object model' is enabled in Excel's Trust Center settings.", vbExclamation
    End If
    
    On Error GoTo 0
    Set codeMod = Nothing
End Sub


' --- Subs for Button Clicks ---

Public Sub ShowAddProductForm()
    ' Load and show the UserForm modally
    Load FrmAddProduct ' Optional: Loads the form into memory without showing
    FrmAddProduct.Show vbModal ' Shows the form and pauses code until it's closed
    ' Execution resumes here after the form is unloaded
End Sub

Public Sub PromptAndUpdateProduct()
    ' Load and show the UserForm modally
    Load FrmUpdateProduct ' Optional: Loads the form into memory without showing
    FrmUpdateProduct.Show vbModal ' Shows the form and pauses code until it's closed
    ' Execution resumes here after the form is unloaded
End Sub

Public Sub PromptAndDeleteProduct()
     Dim productIDString As String ' Use String to capture InputBox result
     Dim productID As Long
     Dim confirm As VbMsgBoxResult
     Dim dataSheet As Worksheet
     
     productIDString = InputBox("Enter the ID of the product to delete:", "Delete Product")
     
     ' Check if the user cancelled or entered non-numeric/invalid ID
     If productIDString = "" Then Exit Sub ' User cancelled or entered nothing
     
     If IsNumeric(productIDString) Then
         productID = CLng(productIDString)
         If productID > 0 Then
             ' Confirm deletion
             confirm = MsgBox("Are you sure you want to delete all data for Product ID " & productID & "? This cannot be undone.", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Delete")
             
             If confirm = vbYes Then
                 On Error GoTo DeleteErrorHandler
                 
                 ' Get the data sheet
                 Set dataSheet = ThisWorkbook.Worksheets(PRODUCT_DATA_SHEET_NAME) ' Assumes constant is defined in modConstants
                 
                 ' --- Replace Placeholder with Actual Call ---
                 ModProductData.DeleteProductData productID, dataSheet
                 ' --- End of Replacement ---
                 
                 ' Show success message
                 MsgBox "All data for Product ID " & productID & " has been deleted.", vbInformation, "Deletion Successful"
                 
                 Set dataSheet = Nothing
                 On Error GoTo 0 ' Turn off specific error handling
             End If
         Else
             MsgBox "Product ID must be a positive number.", vbExclamation
         End If
     Else
         MsgBox "Invalid Product ID entered. Please enter a number.", vbExclamation
     End If
     
     Exit Sub ' Ensure normal exit

DeleteErrorHandler:
    MsgBox "An error occurred while deleting product data:" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical, "Deletion Error"
    Set dataSheet = Nothing
    On Error GoTo 0 ' Turn off error handling before exiting sub
    
End Sub

Public Sub ShowCreateRecipeForm()
    ' Load and show the UserForm modally
    Load FrmCreateRecipe ' Optional: Loads the form into memory without showing
    FrmCreateRecipe.Show vbModal ' Shows the form and pauses code until it's closed
    ' Execution resumes here after the form is unloaded
End Sub


