Attribute VB_Name = "ModProductData"
Option Explicit

'=============================================================================
' Module: ModProductData
' Author: Evan Scott
' Date:   4/26/2025
' Purpose: Handles saving and loading of Product data to/from worksheets.
'=============================================================================

' --- Dependencies ---
' Requires access to the 'ProductDataColumns' Enum (defined in ModConstants)
' Requires Class Modules: Product, NutrientQuantity

' --- Constants ---
Private Const HEADER_ROW As Long = 1 ' Assuming headers are on row 1

'-----------------------------------------------------------------------------
' LoadProduct
'-----------------------------------------------------------------------------
' Purpose: Loads a Product object and its associated NutrientQuantities
'          from a specified worksheet based on the Product ID.
' Arguments:
'   productIDToLoad (Long): The ID of the Product to load.
'   ws (Worksheet): The worksheet containing the product data.
' Returns:
'   Product: A fully populated Product object if found, otherwise Nothing.
'-----------------------------------------------------------------------------
Public Function LoadProduct(productIDToLoad As Long, ws As Worksheet) As Product
    Dim prod As Product             ' The Product object to be returned
    Dim nq As NutrientQuantity      ' A NutrientQuantity object for each nutrient row
    Dim lastRow As Long             ' Last used row in the data sheet
    Dim i As Long                   ' Loop counter for rows
    Dim productDataLoaded As Boolean ' Flag to load product details only once

    ' Initialize flag and return object
    productDataLoaded = False
    Set prod = Nothing ' Ensure function returns Nothing if ID not found

    On Error GoTo ErrorHandler

    ' --- Find Data Rows for the specified Product ID ---
    ' Method 1: Looping (Simpler for this structure)
    If ws Is Nothing Then GoTo ExitFunction ' Exit if worksheet is invalid

    lastRow = ws.Cells(ws.Rows.Count, ProductDataColumns.colProdId).End(xlUp).Row

    ' Check if there's any data below the header row
    If lastRow <= HEADER_ROW Then GoTo ExitFunction

    ' Loop through potential data rows
    For i = HEADER_ROW + 1 To lastRow
        ' Check if the Product ID in the current row matches
        ' Use CLng to handle potential text values in cells, though ID should be numeric
        If CLng(ws.Cells(i, ProductDataColumns.colProdId).value) = productIDToLoad Then

            ' --- Load Product Header Data (only once) ---
            If Not productDataLoaded Then
                Set prod = New Product ' Create the product object on first match
                
                ' Check if object creation failed (unlikely but possible)
                If prod Is Nothing Then
                    Err.Raise vbObjectError + 513, "LoadProduct", "Failed to create Product object."
                    GoTo ErrorHandler ' Or Exit Function
                End If

                ' Populate Product properties from the first matching row
                prod.id = CLng(ws.Cells(i, ProductDataColumns.colProdId).value)
                prod.ProductName = CStr(ws.Cells(i, ProductDataColumns.colProdName).value)
                prod.price = CCur(ws.Cells(i, ProductDataColumns.colProdPrice).value)
                prod.mass = CDbl(ws.Cells(i, ProductDataColumns.colProdMass).value)
                prod.servings = CLng(ws.Cells(i, ProductDataColumns.colProdServings).value)
                ' Note: Assumes Product class initializes its NutrientQuantities collection

                productDataLoaded = True ' Set flag so we don't reload product details
            End If

            ' --- Load Nutrient Quantity Data (for every matching row) ---
            ' Ensure the product object exists before adding nutrients
            If Not prod Is Nothing Then
                Set nq = New NutrientQuantity ' Create a new NQ object for this row
                
                ' Check if object creation failed
                If nq Is Nothing Then
                     Err.Raise vbObjectError + 514, "LoadProduct", "Failed to create NutrientQuantity object."
                     GoTo ErrorHandler ' Or Exit Function
                End If

                ' Populate NutrientQuantity properties from the current row
                nq.nutrientID = CLng(ws.Cells(i, ProductDataColumns.colNutrientId).value)
                nq.MassPerServing = CDbl(ws.Cells(i, ProductDataColumns.colMassPerServing).value)

                ' Add the populated NutrientQuantity object to the Product's collection
                prod.NutrientQuantities.Add nq
                Set nq = Nothing ' Release reference for the next loop iteration
            Else
                ' This case should ideally not happen if productDataLoaded logic is correct,
                ' but adding a safeguard.
                Debug.Print "Warning: Found nutrient row for ID " & productIDToLoad & " but Product object wasn't initialized."
            End If

        End If ' End ID match check
    Next i ' Move to the next row

    ' --- End Method 1 ---

ExitFunction:
    ' Assign the populated product object (or Nothing if not found) to the function result
    Set LoadProduct = prod
    ' Clean up object variables used within the function
    Set nq = Nothing
    Exit Function

ErrorHandler:
    MsgBox "An error occurred in LoadProduct:" & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical, "Load Error"
    ' Ensure function returns Nothing on error
    Set prod = Nothing
    Resume ExitFunction ' Go to cleanup and exit

End Function


'-----------------------------------------------------------------------------
' SaveProduct
'-----------------------------------------------------------------------------
' Purpose: Saves a Product object and its associated NutrientQuantities
'          to the next available rows on a specified worksheet.
'          NOTE: This implementation simply appends. It does NOT check for
'                or update existing entries for the same Product ID.
' Arguments:
'   prod (Product): The Product object to save.
'   ws (Worksheet): The worksheet to save the product data to.
'-----------------------------------------------------------------------------
Public Sub SaveProduct(prod As Product, ws As Worksheet)
    Dim nq As NutrientQuantity ' Loop variable for NutrientQuantities
    Dim nextRow As Long        ' Next empty row to write data
    Dim currentColumn As ProductDataColumns ' Loop variable for columns (optional)

    On Error GoTo ErrorHandler

    ' --- Validate Inputs ---
    If prod Is Nothing Then
        Err.Raise vbObjectError + 515, "SaveProduct", "Product object provided is Nothing."
        GoTo ExitSub
    End If
    If ws Is Nothing Then
        Err.Raise vbObjectError + 516, "SaveProduct", "Worksheet object provided is Nothing."
        GoTo ExitSub
    End If
    If prod.NutrientQuantities Is Nothing Then
         Err.Raise vbObjectError + 517, "SaveProduct", "Product's NutrientQuantities collection is Nothing."
         GoTo ExitSub
    End If
    If prod.NutrientQuantities.Count = 0 Then
        Debug.Print "Warning: Product ID " & prod.id & " has no NutrientQuantities to save."
        GoTo ExitSub ' Or handle as needed - maybe save product details anyway?
    End If


    ' --- Find Next Empty Row ---
    ' Find the last row in the Product ID column and add 1
    nextRow = ws.Cells(ws.Rows.Count, ProductDataColumns.colProdId).End(xlUp).Row + 1
    ' If sheet is empty (except maybe header), start at row 2
    If nextRow <= HEADER_ROW Then nextRow = HEADER_ROW + 1


    ' --- Loop through NutrientQuantities and Write Data ---
    For Each nq In prod.NutrientQuantities
        ' Write Product details (repeated for each nutrient row)
        ws.Cells(nextRow, ProductDataColumns.colProdId).value = prod.id
        ws.Cells(nextRow, ProductDataColumns.colProdName).value = prod.ProductName
        ws.Cells(nextRow, ProductDataColumns.colProdPrice).value = prod.price
        ws.Cells(nextRow, ProductDataColumns.colProdMass).value = prod.mass
        ws.Cells(nextRow, ProductDataColumns.colProdServings).value = prod.servings

        ' Write NutrientQuantity details
        ws.Cells(nextRow, ProductDataColumns.colNutrientId).value = nq.nutrientID
        ws.Cells(nextRow, ProductDataColumns.colMassPerServing).value = nq.MassPerServing

        ' Move to the next row for the next nutrient
        nextRow = nextRow + 1
    Next nq

ExitSub:
    ' Clean up object variables
    Set nq = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in SaveProduct:" & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical, "Save Error"
    Resume ExitSub ' Go to cleanup and exit

End Sub

'-----------------------------------------------------------------------------
' DeleteProductData
'-----------------------------------------------------------------------------
' Purpose: Deletes all rows associated with a specific Product ID from
'          the specified worksheet.
' Arguments:
'   productIDToDelete (Long): The ID of the Product whose rows should be deleted.
'   ws (Worksheet): The worksheet containing the product data.
'-----------------------------------------------------------------------------
Public Sub DeleteProductData(productIDToDelete As Long, ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long

    On Error GoTo ErrorHandler

    ' --- Validate Input ---
    If ws Is Nothing Then
        Err.Raise vbObjectError + 518, "DeleteProductData", "Worksheet object provided is Nothing."
        GoTo ExitSub
    End If

    ' --- Find Last Row ---
    lastRow = ws.Cells(ws.Rows.Count, ProductDataColumns.colProdId).End(xlUp).Row

    ' Check if there's any data to potentially delete
    If lastRow <= HEADER_ROW Then GoTo ExitSub ' Nothing to delete

    ' --- Loop Backwards and Delete Matching Rows ---
    ' Looping backwards is crucial when deleting rows to avoid skipping rows
    For i = lastRow To HEADER_ROW + 1 Step -1
        ' Check if the cell is not empty and if the Product ID matches
        If Not IsEmpty(ws.Cells(i, ProductDataColumns.colProdId).value) Then
            If CLng(ws.Cells(i, ProductDataColumns.colProdId).value) = productIDToDelete Then
                ' Delete the entire row
                ws.Rows(i).Delete
            End If
        End If
    Next i

ExitSub:
    Exit Sub

ErrorHandler:
     MsgBox "An error occurred in DeleteProductData:" & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical, "Delete Error"
    Resume ExitSub ' Go to cleanup and exit

End Sub

'-----------------------------------------------------------------------------
' GetAllProducts
'-----------------------------------------------------------------------------
' Purpose: Loads ALL unique Product objects and their associated
'          NutrientQuantities from the specified worksheet.
' Arguments:
'   ws (Worksheet): The worksheet containing the product data.
' Returns:
'   Collection: A collection of all unique Product objects found on the sheet.
'               Returns Nothing on error or if the sheet is invalid/empty.
' Requires: Reference to "Microsoft Scripting Runtime" for Dictionary object.
'-----------------------------------------------------------------------------
Public Function GetAllProducts(ws As Worksheet) As Collection
    Dim allProductsDict As Scripting.Dictionary ' Temp storage, Key=ProductID, Item=Product Object
    Dim outputCollection As Collection          ' Collection to return
    Dim prod As Product                         ' Current product object being processed
    Dim nq As NutrientQuantity                  ' Current nutrient quantity object
    Dim lastRow As Long
    Dim i As Long
    Dim currentProductID As Long
    Dim productKey As String                    ' Dictionary key (string version of ID)

    On Error GoTo ErrorHandler

    ' --- Validate Input ---
    If ws Is Nothing Then
        Debug.Print "GetAllProducts Error: Worksheet object provided is Nothing."
        GoTo FunctionFailed
    End If

    ' --- Initialization ---
    Set allProductsDict = New Scripting.Dictionary
    Set outputCollection = New Collection

    ' --- Find Last Row ---
    lastRow = ws.Cells(ws.Rows.Count, ProductDataColumns.colProdId).End(xlUp).Row

    ' Check if there's any data below the header row
    If lastRow <= HEADER_ROW Then
        Set GetAllProducts = outputCollection ' Return empty collection if no data
        Exit Function
    End If

    ' --- Loop Through Data Rows ---
    For i = HEADER_ROW + 1 To lastRow
        ' Skip empty rows or rows without a valid numeric Product ID
        If Not IsEmpty(ws.Cells(i, ProductDataColumns.colProdId).value) And _
           IsNumeric(ws.Cells(i, ProductDataColumns.colProdId).value) Then
            
            currentProductID = CLng(ws.Cells(i, ProductDataColumns.colProdId).value)
            productKey = CStr(currentProductID) ' Use string key for dictionary

            ' --- Check if Product already loaded ---
            If Not allProductsDict.Exists(productKey) Then
                ' Product not found, create and populate it
                Set prod = New Product
                If prod Is Nothing Then Err.Raise vbObjectError + 520, "GetAllProducts", "Failed to create Product object."
                
                prod.id = currentProductID
                prod.ProductName = CStr(ws.Cells(i, ProductDataColumns.colProdName).value)
                prod.price = CCur(ws.Cells(i, ProductDataColumns.colProdPrice).value)
                prod.mass = CDbl(ws.Cells(i, ProductDataColumns.colProdMass).value)
                prod.servings = CLng(ws.Cells(i, ProductDataColumns.colProdServings).value)
                
                ' Add the new product to the dictionary
                allProductsDict.Add key:=productKey, Item:=prod
                Set prod = Nothing ' Release local variable, dictionary holds reference
            End If
            
            ' --- Get Product from Dictionary ---
            Set prod = allProductsDict.Item(productKey)
            
            ' --- Create and Add Nutrient Quantity ---
            If Not prod Is Nothing Then ' Should always exist if logic is correct
                Set nq = New NutrientQuantity
                If nq Is Nothing Then Err.Raise vbObjectError + 521, "GetAllProducts", "Failed to create NutrientQuantity object."
                
                nq.nutrientID = CLng(ws.Cells(i, ProductDataColumns.colNutrientId).value)
                nq.MassPerServing = CDbl(ws.Cells(i, ProductDataColumns.colMassPerServing).value)
                
                ' Add the NQ to the Product's collection
                prod.NutrientQuantities.Add nq
                Set nq = Nothing ' Release local variable
            End If
            
        End If ' End check for valid Product ID cell
    Next i

    ' --- Transfer Products from Dictionary to Output Collection ---
    Dim dictKey As Variant
    For Each dictKey In allProductsDict.Keys
        outputCollection.Add allProductsDict.Item(dictKey)
    Next dictKey

    ' --- Set Return Value ---
    Set GetAllProducts = outputCollection
    GoTo ExitFunction

FunctionFailed:
    Set GetAllProducts = Nothing ' Return Nothing on failure

ExitFunction:
    ' Clean up
    Set allProductsDict = Nothing
    Set outputCollection = Nothing ' Caller gets the returned collection reference
    Set prod = Nothing
    Set nq = Nothing
    Exit Function

ErrorHandler:
    MsgBox "An error occurred in GetAllProducts:" & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical, "Get All Products Error"
    Resume FunctionFailed ' Go to failure cleanup on error

End Function

