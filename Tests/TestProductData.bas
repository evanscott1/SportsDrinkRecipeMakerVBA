Attribute VB_Name = "TestProductData"
'@TestModule
'@Folder("Tests")
Option Explicit
Option Private Module 'Ensures Assert/Fakes aren't globally exposed if Private

'=============================================================================
' Module: TestProductData
' Author: Evan Scott
' Date:   4/28/2025
' Purpose: Contains integration tests for the modProductData module,
'          verifying the saving, loading, deleting, and updating of
'          Product data to/from an Excel worksheet using the
'          Rubberduck VBA testing framework.
'=============================================================================

' --- Conditional Compilation for Binding ---
' Allows switching between early binding (for development) and late binding (for distribution)
' Set to False for development to enable IntelliSense and compile-time checks for Assert/Fakes.
' Set to True before distributing to avoid requiring end-users to have Rubberduck referenced.
#Const LateBind = False ' Or True depending on context

#If LateBind Then
    Private Assert As Object ' Generic Object for late binding
    Private Fakes As Object  ' Generic Object for late binding (if used)
#Else
    Private Assert As Rubberduck.AssertClass ' Specific type for early binding
    Private Fakes As Rubberduck.FakesProvider ' Specific type for early binding (if used)
#End If

' --- Test Constants ---
Private Const TEST_SHEET_NAME As String = "TestProductData" ' Name of the temporary worksheet used for testing
Private Const HEADER_ROW As Long = 1                       ' Row number containing the headers on the test sheet

' --- Test State ---
' Module-level variables to hold objects needed across multiple tests or setup/teardown phases.
Private testWs As Worksheet         ' Represents the temporary test worksheet object
Private testProducts As Collection  ' Stores Product objects created during tests for potential cleanup

' --- Test Fixture Setup / Teardown ---

'@ModuleInitialize
Private Sub ModuleInitialize()
    ' Runs once before any tests in this module execute.
    ' Sets up the Assert/Fakes objects and creates the test worksheet.
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.AssertClass")
        ' Set Fakes = CreateObject("Rubberduck.FakesProvider") ' Uncomment if Fakes are needed
    #Else
        Set Assert = New Rubberduck.AssertClass
        ' Set Fakes = New Rubberduck.FakesProvider ' Uncomment if Fakes are needed
    #End If

    On Error Resume Next ' Ignore error if sheet already exists
    Set testWs = ThisWorkbook.Worksheets(TEST_SHEET_NAME)
    If testWs Is Nothing Then ' If sheet doesn't exist, create it
        Set testWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        testWs.Name = TEST_SHEET_NAME
    End If
    testWs.Visible = xlSheetVisible ' Keep visible during testing, can hide later
    On Error GoTo 0 ' Resume normal error handling
    Debug.Print "Test Module Initialized: TestProductData"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    ' Runs once after all tests in this module have executed.
    ' Cleans up module-level objects and deletes the test worksheet.
    Set Assert = Nothing
    Set Fakes = Nothing

    On Error Resume Next ' Ignore error if sheet doesn't exist
    Application.DisplayAlerts = False ' Prevent confirmation dialogs
    If Not testWs Is Nothing Then
        If WorksheetExists(TEST_SHEET_NAME) Then
             testWs.Delete ' Delete the test sheet
        End If
    End If
    Application.DisplayAlerts = True ' Restore alerts
    Set testWs = Nothing
    Set testProducts = Nothing
    On Error GoTo 0 ' Resume normal error handling
    Debug.Print "Test Module Cleaned Up: TestProductData"
End Sub

'@TestInitialize
Private Sub TestInitialize()
    ' Runs before every individual test method in this module.
    ' Ensures a clean test worksheet with headers and resets the testProducts collection.
    If Not testWs Is Nothing Then
        testWs.Cells.Clear ' Clear all previous data
        ' Set up headers using the Enum values (Enum defined in modConstants)
        With testWs
            .Cells(HEADER_ROW, ProductDataColumns.colProdId).value = "ProductID"
            .Cells(HEADER_ROW, ProductDataColumns.colProdName).value = "ProductName"
            .Cells(HEADER_ROW, ProductDataColumns.colProdPrice).value = "Price"
            .Cells(HEADER_ROW, ProductDataColumns.colProdMass).value = "TotalMass"
            .Cells(HEADER_ROW, ProductDataColumns.colProdServings).value = "Servings"
            .Cells(HEADER_ROW, ProductDataColumns.colNutrientId).value = "NutrientID"
            .Cells(HEADER_ROW, ProductDataColumns.colMassPerServing).value = "MassPerServing"
        End With
    Else
        ' Fail initialization if sheet wasn't created in ModuleInitialize
        Err.Raise vbObjectError + 999, "TestInitialize", "Test worksheet was not properly initialized."
    End If
    Set testProducts = New Collection ' Initialize a new collection for this test
End Sub

'@TestCleanup
Private Sub TestCleanup()
    ' Runs after every individual test method in this module.
    ' Cleans up objects created specifically for the test that just ran.
    Set testProducts = Nothing ' Clear any test-specific product objects
End Sub

' --- Helper Function ---

'-----------------------------------------------------------------------------
' WorksheetExists
'-----------------------------------------------------------------------------
' Purpose: Checks if a worksheet with the specified name exists in the
'          current workbook.
' Arguments:
'   shtName (String): The name of the worksheet to check for.
' Returns:
'   Boolean: True if the worksheet exists, False otherwise.
'-----------------------------------------------------------------------------
Private Function WorksheetExists(shtName As String) As Boolean
    Dim sht As Object ' Use Object to avoid error if sheet doesn't exist
    On Error Resume Next ' Prevent error if sheet is not found
    Set sht = ThisWorkbook.Worksheets(shtName)
    On Error GoTo 0 ' Restore error handling
    WorksheetExists = Not sht Is Nothing ' Returns True if Set was successful
End Function

' --- Test Methods ---

'@TestMethod
Public Sub TestProductData_SaveProduct_SingleNutrient()
    ' Tests if SaveProduct correctly writes data for a product with one nutrient quantity.
    ' Arrange: Create a product with one nutrient using the REAL Product class
    Dim prod As Product         ' Use actual Product class
    Dim nq As NutrientQuantity  ' Use actual NutrientQuantity class
    Dim dataRow As Long
    dataRow = HEADER_ROW + 1
    
    ' Create Product instance
    Set prod = New Product
    prod.id = 1
    prod.ProductName = "Potassium Chloride"
    prod.price = 20
    prod.mass = 0.3
    prod.servings = 5000
    
    ' Create NutrientQuantity instance
    Set nq = New NutrientQuantity
    nq.nutrientID = 101 ' Store the ID directly
    nq.MassPerServing = 0.35
    
    ' Add nutrient quantity to product's actual collection
    prod.NutrientQuantities.Add nq
    
    testProducts.Add prod ' Keep track for cleanup if needed
    
    ' Act: Save the product to the test worksheet
    ModProductData.SaveProduct prod, testWs
    
    ' Assert: Verify the data written to the worksheet cells using Product properties
    With testWs
        ' Use CLng for Long comparisons to handle potential type differences from worksheet reads
        Assert.AreEqual prod.id, CLng(.Cells(dataRow, ProductDataColumns.colProdId).value), "Product ID mismatch"
        Assert.AreEqual prod.ProductName, .Cells(dataRow, ProductDataColumns.colProdName).value, "Product Name mismatch"
        Assert.AreEqual prod.price, .Cells(dataRow, ProductDataColumns.colProdPrice).value, "Product Price mismatch"
        Assert.AreEqual prod.mass, .Cells(dataRow, ProductDataColumns.colProdMass).value, "Product Mass mismatch"
        Assert.AreEqual prod.servings, CLng(.Cells(dataRow, ProductDataColumns.colProdServings).value), "Product Servings mismatch"
        ' Assert Nutrient Details using NutrientQuantity properties
        Assert.AreEqual nq.nutrientID, CLng(.Cells(dataRow, ProductDataColumns.colNutrientId).value), "Nutrient ID mismatch"
        Assert.AreEqual nq.MassPerServing, .Cells(dataRow, ProductDataColumns.colMassPerServing).value, "Nutrient MassPerServing mismatch"
    End With
End Sub

'@TestMethod
Public Sub TestProductData_SaveProduct_MultipleNutrients()
    ' Tests if SaveProduct correctly writes multiple rows for a product with multiple nutrient quantities.
    ' Arrange: Create a product with two nutrients using REAL classes
    Dim prod As Product
    Dim nq1 As NutrientQuantity, nq2 As NutrientQuantity
    Dim dataRow1 As Long, dataRow2 As Long
    dataRow1 = HEADER_ROW + 1
    dataRow2 = HEADER_ROW + 2
        
    ' Create Product instance
    Set prod = New Product
    prod.id = 2
    prod.ProductName = "Salt Substitute"
    prod.price = 15
    prod.mass = 0.5
    prod.servings = 1000
    
    ' Create NutrientQuantity instances
    Set nq1 = New NutrientQuantity: nq1.nutrientID = 101: nq1.MassPerServing = 0.4
    Set nq2 = New NutrientQuantity: nq2.nutrientID = 102: nq2.MassPerServing = 0.1
    
    ' Add NutrientQuantities to Product's collection
    prod.NutrientQuantities.Add nq1
    prod.NutrientQuantities.Add nq2
    
    testProducts.Add prod
    
    ' Act: Save the product
    ModProductData.SaveProduct prod, testWs
    
    ' Assert: Verify data for both nutrient rows using actual properties
    ' Row 1 (Nutrient 1)
    With testWs
        ' Use CLng for Long comparisons
        Assert.AreEqual prod.id, CLng(.Cells(dataRow1, ProductDataColumns.colProdId).value), "Row1: Product ID mismatch"
        Assert.AreEqual prod.ProductName, .Cells(dataRow1, ProductDataColumns.colProdName).value, "Row1: Product Name mismatch"
        Assert.AreEqual nq1.nutrientID, CLng(.Cells(dataRow1, ProductDataColumns.colNutrientId).value), "Row1: Nutrient ID mismatch"
        Assert.AreEqual nq1.MassPerServing, .Cells(dataRow1, ProductDataColumns.colMassPerServing).value, "Row1: Nutrient MassPerServing mismatch"
    End With
    
    ' Row 2 (Nutrient 2)
    With testWs
        ' Use CLng for Long comparisons
        Assert.AreEqual prod.id, CLng(.Cells(dataRow2, ProductDataColumns.colProdId).value), "Row2: Product ID mismatch" ' Product details repeated
        Assert.AreEqual prod.ProductName, .Cells(dataRow2, ProductDataColumns.colProdName).value, "Row2: Product Name mismatch"
        Assert.AreEqual nq2.nutrientID, CLng(.Cells(dataRow2, ProductDataColumns.colNutrientId).value), "Row2: Nutrient ID mismatch"
        Assert.AreEqual nq2.MassPerServing, .Cells(dataRow2, ProductDataColumns.colMassPerServing).value, "Row2: Nutrient MassPerServing mismatch"
    End With
End Sub

'@TestMethod
Public Sub TestProductData_LoadProduct_SingleNutrient()
    ' Tests if LoadProduct correctly reconstructs a Product object with one nutrient quantity.
    ' Arrange: Manually write data for one product/nutrient to the sheet.
    Dim productIDToLoad As Long: productIDToLoad = 3
    Dim expectedProdName As String: expectedProdName = "Calcium Carbonate"
    Dim expectedPrice As Currency: expectedPrice = 12.5
    Dim expectedMass As Double: expectedMass = 0.25
    Dim expectedServings As Long: expectedServings = 2000
    Dim expectedNutrientID As Long: expectedNutrientID = 201 ' The ID stored in NutrientQuantity
    Dim expectedMassPerServing As Double: expectedMassPerServing = 0.45
    Dim dataRow As Long: dataRow = HEADER_ROW + 1
    
    With testWs
        .Cells(dataRow, ProductDataColumns.colProdId).value = productIDToLoad
        .Cells(dataRow, ProductDataColumns.colProdName).value = expectedProdName
        .Cells(dataRow, ProductDataColumns.colProdPrice).value = expectedPrice
        .Cells(dataRow, ProductDataColumns.colProdMass).value = expectedMass
        .Cells(dataRow, ProductDataColumns.colProdServings).value = expectedServings
        .Cells(dataRow, ProductDataColumns.colNutrientId).value = expectedNutrientID
        .Cells(dataRow, ProductDataColumns.colMassPerServing).value = expectedMassPerServing
    End With
    
    ' Act: Load the product
    Dim loadedProd As Product ' Expect LoadProduct to return a Product object
    Set loadedProd = ModProductData.LoadProduct(productIDToLoad, testWs)
    
    ' Assert: Verify the loaded product object exists
    Assert.IsNotNothing loadedProd, "Loaded product object should not be Nothing"
    If loadedProd Is Nothing Then Exit Sub ' Guard clause for further assertions
    
    ' Assert: Verify Product properties match the sheet data
    Assert.AreEqual productIDToLoad, loadedProd.id, "Loaded Product ID mismatch"
    Assert.AreEqual expectedProdName, loadedProd.ProductName, "Loaded Product Name mismatch"
    Assert.AreEqual expectedPrice, loadedProd.price, "Loaded Product Price mismatch"
    Assert.AreEqual expectedMass, loadedProd.mass, "Loaded Product Mass mismatch"
    Assert.AreEqual expectedServings, loadedProd.servings, "Loaded Product Servings mismatch"
    
    ' Assert: Verify NutrientQuantities collection contains the correct number of items
    Assert.AreEqual CLng(1), loadedProd.NutrientQuantities.Count, "Expected 1 nutrient quantity"
    
    ' Assert: Get the NutrientQuantity object and verify it exists
    Dim loadedNq As NutrientQuantity ' Use the specific class type
    Set loadedNq = loadedProd.NutrientQuantities(1) ' Assuming 1-based collection
    Assert.IsNotNothing loadedNq, "Loaded nutrient quantity object should not be Nothing"
    If loadedNq Is Nothing Then Exit Sub ' Guard clause
    
    ' Assert: Verify NutrientQuantity properties match the sheet data
    Assert.AreEqual expectedNutrientID, loadedNq.nutrientID, "Loaded Nutrient ID mismatch"
    Assert.AreEqual expectedMassPerServing, loadedNq.MassPerServing, "Loaded MassPerServing mismatch"
    
End Sub

'@TestMethod
Public Sub TestProductData_LoadProduct_MultipleNutrients()
    ' Tests if LoadProduct correctly reconstructs a Product object with multiple nutrient quantities.
    ' Arrange: Manually write data for one product with two nutrients to the sheet.
    Dim productIDToLoad As Long: productIDToLoad = 4
    Dim prodName As String: prodName = "MultiMineral"
    Dim price As Currency: price = 30
    Dim mass As Double: mass = 0.1
    Dim servings As Long: servings = 100
    
    Dim nutrient1ID As Long: nutrient1ID = 301
    Dim massPerServing1 As Double: massPerServing1 = 0.015
    
    Dim nutrient2ID As Long: nutrient2ID = 302
    Dim massPerServing2 As Double: massPerServing2 = 0.1
    
    Dim dataRow1 As Long: dataRow1 = HEADER_ROW + 1
    Dim dataRow2 As Long: dataRow2 = HEADER_ROW + 2
    
    ' Write Row 1 (Nutrient 1)
    With testWs
        .Cells(dataRow1, ProductDataColumns.colProdId).value = productIDToLoad
        .Cells(dataRow1, ProductDataColumns.colProdName).value = prodName
        .Cells(dataRow1, ProductDataColumns.colProdPrice).value = price
        .Cells(dataRow1, ProductDataColumns.colProdMass).value = mass
        .Cells(dataRow1, ProductDataColumns.colProdServings).value = servings
        .Cells(dataRow1, ProductDataColumns.colNutrientId).value = nutrient1ID
        .Cells(dataRow1, ProductDataColumns.colMassPerServing).value = massPerServing1
    End With
    
    ' Write Row 2 (Nutrient 2)
    With testWs
        .Cells(dataRow2, ProductDataColumns.colProdId).value = productIDToLoad
        .Cells(dataRow2, ProductDataColumns.colProdName).value = prodName
        .Cells(dataRow2, ProductDataColumns.colProdPrice).value = price
        .Cells(dataRow2, ProductDataColumns.colProdMass).value = mass
        .Cells(dataRow2, ProductDataColumns.colProdServings).value = servings
        .Cells(dataRow2, ProductDataColumns.colNutrientId).value = nutrient2ID
        .Cells(dataRow2, ProductDataColumns.colMassPerServing).value = massPerServing2
    End With
    
    ' Act: Load the product
    Dim loadedProd As Product ' Expect LoadProduct to return a Product object
    Set loadedProd = ModProductData.LoadProduct(productIDToLoad, testWs)
    
    ' Assert: Verify the loaded product object exists
    Assert.IsNotNothing loadedProd, "Loaded product object should not be Nothing"
    If loadedProd Is Nothing Then Exit Sub ' Guard clause
    
    ' Assert: Verify product properties
    Assert.AreEqual productIDToLoad, loadedProd.id, "Loaded Product ID mismatch"
    Assert.AreEqual prodName, loadedProd.ProductName, "Loaded Product Name mismatch"
    
    ' Assert: Verify NutrientQuantities collection count
    Assert.AreEqual CLng(2), loadedProd.NutrientQuantities.Count, "Expected 2 nutrient quantities"
    
    ' Assert: Verify Nutrient Quantity 1 (nq1)
    Dim nq1 As NutrientQuantity ' Use specific class type
    Set nq1 = loadedProd.NutrientQuantities(1) ' Assuming 1-based collection
    Assert.IsNotNothing nq1, "NQ1 object should not be Nothing"
    If Not nq1 Is Nothing Then ' Safety check for nq1 object
        Assert.AreEqual nutrient1ID, nq1.nutrientID, "NQ1 Nutrient ID mismatch"
        Assert.AreEqual massPerServing1, nq1.MassPerServing, "NQ1 MassPerServing mismatch"
    End If
    
    ' Assert: Verify Nutrient Quantity 2 (nq2)
    Dim nq2 As NutrientQuantity ' Use specific class type
    Set nq2 = loadedProd.NutrientQuantities(2) ' Assuming 1-based collection
    Assert.IsNotNothing nq2, "NQ2 object should not be Nothing"
     If Not nq2 Is Nothing Then ' Safety check for nq2 object
        Assert.AreEqual nutrient2ID, nq2.nutrientID, "NQ2 Nutrient ID mismatch"
        Assert.AreEqual massPerServing2, nq2.MassPerServing, "NQ2 MassPerServing mismatch"
    End If
    
End Sub

'@TestMethod
Public Sub TestProductData_LoadProduct_NonExistentID()
    ' Tests if LoadProduct correctly returns Nothing for a Product ID that does not exist on the sheet.
    ' Arrange: Ensure the sheet is clean or contains other data, define an ID known not to exist.
    TestInitialize ' Ensure clean sheet with headers
        
    ' Add some other data (optional, makes test clearer)
    With testWs.Cells(HEADER_ROW + 1, ProductDataColumns.colProdId)
        .value = 999 ' Some other product ID
        .Offset(0, ProductDataColumns.colProdName - ProductDataColumns.colProdId).value = "Other Product"
        .Offset(0, ProductDataColumns.colNutrientId - ProductDataColumns.colProdId).value = 901
        .Offset(0, ProductDataColumns.colMassPerServing - ProductDataColumns.colProdId).value = 0.99
    End With
            
    Dim nonExistentID As Long: nonExistentID = 500
    
    ' Act: Attempt to load the non-existent product
    Dim loadedProd As Product ' Expect LoadProduct to return a Product object (or Nothing)
    Set loadedProd = ModProductData.LoadProduct(nonExistentID, testWs)
    
    ' Assert: Verify that the function returned Nothing
    Assert.IsNothing loadedProd, "Expected Nothing for a non-existent Product ID"
End Sub


'@TestMethod
Public Sub TestProductData_DeleteProduct_RemovesAllRows()
    ' Tests if DeleteProductData correctly removes all rows for a specific Product ID,
    ' leaving other products untouched.
    ' Arrange: Save product to delete and another product to keep
    Dim prodToDelete As Product
    Dim nqDelete1 As NutrientQuantity, nqDelete2 As NutrientQuantity
    Dim productIDToDelete As Long: productIDToDelete = 6
    
    Set prodToDelete = New Product
    prodToDelete.id = productIDToDelete
    prodToDelete.ProductName = "Product To Delete"
    prodToDelete.price = 10: prodToDelete.mass = 0.1: prodToDelete.servings = 100
    Set nqDelete1 = New NutrientQuantity: nqDelete1.nutrientID = 501: nqDelete1.MassPerServing = 0.1
    Set nqDelete2 = New NutrientQuantity: nqDelete2.nutrientID = 502: nqDelete2.MassPerServing = 0.2
    prodToDelete.NutrientQuantities.Add nqDelete1
    prodToDelete.NutrientQuantities.Add nqDelete2
    ModProductData.SaveProduct prodToDelete, testWs ' Save product to delete (2 rows)
    
    Dim prodToKeep As Product
    Dim nqKeep As NutrientQuantity
    Dim productIDToKeep As Long: productIDToKeep = 7
    
    Set prodToKeep = New Product
    prodToKeep.id = productIDToKeep
    prodToKeep.ProductName = "Product To Keep"
    prodToKeep.price = 50: prodToKeep.mass = 1: prodToKeep.servings = 50
    Set nqKeep = New NutrientQuantity: nqKeep.nutrientID = 601: nqKeep.MassPerServing = 0.5
    prodToKeep.NutrientQuantities.Add nqKeep
    ModProductData.SaveProduct prodToKeep, testWs ' Save product to keep (1 row)
    
    testProducts.Add prodToDelete ' Add to collection for potential cleanup
    testProducts.Add prodToKeep
    
    ' Act: Delete the target product's data
    ModProductData.DeleteProductData productIDToDelete, testWs ' *** Assumes DeleteProductData is implemented ***
    
    ' Assert: Verify the product is gone by trying to load it
    Dim loadedProdAfterDelete As Product
    Set loadedProdAfterDelete = ModProductData.LoadProduct(productIDToDelete, testWs)
    Assert.IsNothing loadedProdAfterDelete, "Product ID " & productIDToDelete & " should not be loadable after delete"
    
    ' Assert: Verify the other product is still present and correct
    Dim loadedProdToKeep As Product
    Set loadedProdToKeep = ModProductData.LoadProduct(productIDToKeep, testWs)
    Assert.IsNotNothing loadedProdToKeep, "Product ID " & productIDToKeep & " should still be loadable"
    If Not loadedProdToKeep Is Nothing Then ' Further checks if load succeeded
        Assert.AreEqual productIDToKeep, loadedProdToKeep.id, "Kept Product ID mismatch"
        Assert.AreEqual prodToKeep.ProductName, loadedProdToKeep.ProductName, "Kept Product Name mismatch"
        Assert.AreEqual CLng(1), loadedProdToKeep.NutrientQuantities.Count, "Kept product should still have 1 nutrient quantity"
    End If
    
    ' Assert: Verify no rows remain on the sheet for the deleted product ID by direct count
    Dim rowCountDeleted As Long
    Dim i As Long
    Dim lastRow As Long
    rowCountDeleted = 0
    lastRow = testWs.Cells(testWs.Rows.Count, ProductDataColumns.colProdId).End(xlUp).Row
    ' Loop through remaining rows to count occurrences of the deleted ID
    For i = HEADER_ROW + 1 To lastRow
        If Not IsEmpty(testWs.Cells(i, ProductDataColumns.colProdId).value) Then
            If CLng(testWs.Cells(i, ProductDataColumns.colProdId).value) = productIDToDelete Then
                rowCountDeleted = rowCountDeleted + 1
            End If
        End If
    Next i
    Assert.AreEqual CLng(0), rowCountDeleted, "Delete: Expected exactly 0 rows on sheet for Product ID " & productIDToDelete
    
End Sub


'@TestMethod
Public Sub TestProductData_UpdateProduct_OverwritesCorrectly()
    ' Tests if updating an existing product's data using a Delete+Save strategy works correctly.
    ' Arrange: Save an initial version of the product
    Dim prodInitial As Product
    Dim nqInitial1 As NutrientQuantity, nqInitial2 As NutrientQuantity
    Dim productIDToUpdate As Long: productIDToUpdate = 5
    
    Set prodInitial = New Product
    prodInitial.id = productIDToUpdate
    prodInitial.ProductName = "Magnesium Glycinate (Old)"
    prodInitial.price = 25
    prodInitial.mass = 0.2
    prodInitial.servings = 120
    
    Set nqInitial1 = New NutrientQuantity: nqInitial1.nutrientID = 401: nqInitial1.MassPerServing = 0.1 ' 100mg Magnesium
    Set nqInitial2 = New NutrientQuantity: nqInitial2.nutrientID = 402: nqInitial2.MassPerServing = 0.5 ' 500mg Glycine
    prodInitial.NutrientQuantities.Add nqInitial1
    prodInitial.NutrientQuantities.Add nqInitial2
    
    ModProductData.SaveProduct prodInitial, testWs ' Save initial version (2 rows)
    
    ' Arrange: Create the updated version of the product (same ID, different data)
    Dim prodUpdated As Product
    Dim nqUpdated As NutrientQuantity
    
    Set prodUpdated = New Product
    prodUpdated.id = productIDToUpdate ' Same ID
    prodUpdated.ProductName = "Magnesium Glycinate (New Price)" ' Updated Name
    prodUpdated.price = 28.5 ' Updated Price
    prodUpdated.mass = 0.2 ' Mass same
    prodUpdated.servings = 120 ' Servings same
    
    ' Only include one nutrient in the updated version with updated value
    Set nqUpdated = New NutrientQuantity: nqUpdated.nutrientID = 401: nqUpdated.MassPerServing = 0.11 ' Updated MassPerServing
    prodUpdated.NutrientQuantities.Add nqUpdated
    
    testProducts.Add prodInitial ' Add both to collection for potential cleanup
    testProducts.Add prodUpdated
    
    ' Act: Perform the update (Delete existing, then Save updated)
    ModProductData.DeleteProductData productIDToUpdate, testWs ' *** Assumes DeleteProductData is implemented ***
    ModProductData.SaveProduct prodUpdated, testWs ' Save the updated version (1 row)
    
    ' Assert: Load the product back and verify it matches the updated version
    Dim loadedProdAfterUpdate As Product
    Set loadedProdAfterUpdate = ModProductData.LoadProduct(productIDToUpdate, testWs)
    
    Assert.IsNotNothing loadedProdAfterUpdate, "Product should exist after update"
    If loadedProdAfterUpdate Is Nothing Then Exit Sub ' Guard clause
    
    ' Check updated product properties
    Assert.AreEqual prodUpdated.id, loadedProdAfterUpdate.id, "Update: Product ID mismatch"
    Assert.AreEqual prodUpdated.ProductName, loadedProdAfterUpdate.ProductName, "Update: Product Name mismatch"
    Assert.AreEqual prodUpdated.price, loadedProdAfterUpdate.price, "Update: Product Price mismatch"
    Assert.AreEqual prodUpdated.mass, loadedProdAfterUpdate.mass, "Update: Product Mass mismatch"
    Assert.AreEqual prodUpdated.servings, loadedProdAfterUpdate.servings, "Update: Product Servings mismatch"
    
    ' Check updated nutrient collection (should only have the one from prodUpdated)
    Assert.AreEqual CLng(1), loadedProdAfterUpdate.NutrientQuantities.Count, "Update: Expected 1 nutrient quantity after update"
    
    ' Check the details of the single nutrient quantity
    If loadedProdAfterUpdate.NutrientQuantities.Count = 1 Then
        Dim loadedNqAfterUpdate As NutrientQuantity
        Set loadedNqAfterUpdate = loadedProdAfterUpdate.NutrientQuantities(1)
        Assert.IsNotNothing loadedNqAfterUpdate, "Update: Loaded NQ object should not be Nothing"
        If Not loadedNqAfterUpdate Is Nothing Then
            Assert.AreEqual nqUpdated.nutrientID, loadedNqAfterUpdate.nutrientID, "Update: Nutrient ID mismatch"
            Assert.AreEqual nqUpdated.MassPerServing, loadedNqAfterUpdate.MassPerServing, "Update: MassPerServing mismatch"
        End If
    End If
    
    ' Assert: Verify the correct number of rows exist on the sheet for this ID after update
    Dim rowCountUpdate As Long
    Dim iUpdate As Long
    Dim lastRowUpdate As Long
    rowCountUpdate = 0
    lastRowUpdate = testWs.Cells(testWs.Rows.Count, ProductDataColumns.colProdId).End(xlUp).Row
    ' Loop through remaining rows to count occurrences of the updated ID
    For iUpdate = HEADER_ROW + 1 To lastRowUpdate
        If Not IsEmpty(testWs.Cells(iUpdate, ProductDataColumns.colProdId).value) Then
            If CLng(testWs.Cells(iUpdate, ProductDataColumns.colProdId).value) = productIDToUpdate Then
                rowCountUpdate = rowCountUpdate + 1
            End If
        End If
    Next iUpdate
    Assert.AreEqual CLng(1), rowCountUpdate, "Update: Expected exactly 1 row on sheet for Product ID " & productIDToUpdate & " after update"
    
End Sub


