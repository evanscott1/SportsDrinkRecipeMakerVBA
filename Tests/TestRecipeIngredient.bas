Attribute VB_Name = "TestRecipeIngredient"
'@TestModule
'@Folder("Tests")
Option Explicit
Option Private Module 'Ensures Assert/Fakes aren't globally exposed if Private

'=============================================================================
' Module: TestRecipeIngredient
' Author: Evan Scott
' Date:   April 28, 2025
' Purpose: Contains unit tests for the RecipeIngredient class module using the
'          Rubberduck VBA testing framework.
'=============================================================================

' --- Conditional Compilation for Binding ---
#Const LateBind = False ' Or True depending on context

#If LateBind Then
    Private Assert As Object ' Generic Object for late binding
    Private Fakes As Object  ' Generic Object for late binding (if used)
#Else
    Private Assert As Rubberduck.AssertClass ' Specific type for early binding
    Private Fakes As Rubberduck.FakesProvider ' Specific type for early binding (if used)
#End If

' --- Module-Level Test Variables ---
Private testIngredient As RecipeIngredient ' The instance of the RecipeIngredient class being tested
Private testProduct As Product           ' A mock Product object for testing

' --- Test Fixture Setup / Teardown ---

'@ModuleInitialize
Private Sub ModuleInitialize()
    ' Runs once before any tests in this module execute.
    ' Used here to instantiate the Assert and Fakes objects.
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.AssertClass")
        ' Set Fakes = CreateObject("Rubberduck.FakesProvider") ' Uncomment if Fakes are needed
    #Else
        Set Assert = New Rubberduck.AssertClass
        ' Set Fakes = New Rubberduck.FakesProvider ' Uncomment if Fakes are needed
    #End If
    Debug.Print "Test Module Initialized: TestRecipeIngredient"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    ' Runs once after all tests in this module have executed.
    ' Used here to clean up the Assert and Fakes objects.
    Set Assert = Nothing
    Set Fakes = Nothing
    Debug.Print "Test Module Cleaned Up: TestRecipeIngredient"
End Sub


'@TestInitialize
Private Sub TestInitialize()
    ' Runs before every individual test method in this module.
    ' Creates a fresh instance of the RecipeIngredient and a standard mock Product.
    Set testIngredient = New RecipeIngredient
    
    ' Create a standard valid Product for most tests
    Set testProduct = New Product
    testProduct.id = 10
    testProduct.ProductName = "Test Protein Powder"
    testProduct.price = 30
    testProduct.mass = 1#  ' 1 kg total mass
    testProduct.servings = 30 ' 30 servings total
    ' Note: NutrientQuantities not strictly needed for RecipeIngredient tests,
    ' but could be added if methods depending on them were tested.
    
End Sub

'@TestCleanup
Private Sub TestCleanup()
    ' Runs after every individual test method in this module.
    ' Destroys the test objects.
    Set testIngredient = Nothing
    Set testProduct = Nothing
End Sub

' --- Test Methods ---

'@TestMethod
Public Sub TestRecipeIngredient_Init_AssignsProductCorrectly()
    ' Tests if the Init method correctly assigns the Product object.
    ' Arrange
    Dim servingsNeeded As Double: servingsNeeded = 2.5
    
    ' Act
    testIngredient.Init testProduct, servingsNeeded
    
    ' Assert
    Assert.AreSame testProduct, testIngredient.Product, "Init should assign the correct Product object reference."
End Sub

'@TestMethod
Public Sub TestRecipeIngredient_Init_StoresAmountServingsCorrectly()
    ' Tests if the Init method correctly stores the AmountServings value.
    ' Arrange
    Dim expectedServings As Double: expectedServings = 1.75
    
    ' Act
    testIngredient.Init testProduct, expectedServings
    
    ' Assert
    Assert.AreEqual expectedServings, testIngredient.AmountServings, "Init should store the correct AmountServings."
End Sub

'@TestMethod
Public Sub TestRecipeIngredient_Init_CalculatesAmountKgCorrectly()
    ' Tests if the Init method correctly calculates AmountKg.
    ' Arrange
    Dim servingsNeeded As Double: servingsNeeded = 2#
    ' Expected calculation: 2.0 servings * (1.0 kg / 30 servings) = 0.0666... kg
    Dim expectedKg As Double: expectedKg = 2# * (testProduct.mass / testProduct.servings)
    Dim tolerance As Double: tolerance = 0.000001 ' Define the acceptable tolerance
    
    ' Act
    testIngredient.Init testProduct, servingsNeeded
    Dim actualKg As Double: actualKg = testIngredient.AmountKg
    
    ' Assert
    ' Manual check for floating-point equality within tolerance
    Dim difference As Double
    difference = Abs(expectedKg - actualKg) ' Calculate absolute difference
    Assert.IsTrue difference <= tolerance, "Init should calculate AmountKg correctly within tolerance " & tolerance & " (Actual: " & actualKg & ", Diff: " & difference & ")"
End Sub

'@TestMethod
Public Sub TestRecipeIngredient_Init_HandlesZeroServingsNeeded()
    ' Tests if Init handles being called with zero servings.
    ' Arrange
    Dim servingsNeeded As Double: servingsNeeded = 0
    Dim expectedKg As Double: expectedKg = 0
    
    ' Act
    testIngredient.Init testProduct, servingsNeeded
    
    ' Assert
    Assert.AreEqual servingsNeeded, testIngredient.AmountServings, "AmountServings should be 0 when 0 is passed to Init."
    Assert.AreEqual expectedKg, testIngredient.AmountKg, "AmountKg should be 0 when 0 servings are needed."
    Assert.AreSame testProduct, testIngredient.Product, "Product should still be assigned even with 0 servings."
End Sub

'@TestMethod
Public Sub TestRecipeIngredient_Init_RaisesErrorIfProductIsNull()
    ' Tests if Init raises the correct error when passed Nothing for the Product.
    ' Arrange
    Dim servingsNeeded As Double: servingsNeeded = 1#
    Dim expectedError As Long: expectedError = vbObjectError + 601
    Set testProduct = Nothing ' Intentionally set product to Nothing
    
    ' Act
    On Error Resume Next ' Turn on error trapping
    Err.Clear            ' Clear any previous errors
    testIngredient.Init testProduct, servingsNeeded ' This line should raise error
    Dim actualError As Long: actualError = Err.Number ' Capture the error number
    On Error GoTo 0      ' Turn off error trapping
    
    ' Assert
    Assert.AreEqual expectedError, actualError, "Expected error " & expectedError & " when Product is Nothing."
    If actualError = 0 Then Assert.Fail "Error was expected but did not occur." ' Fail if no error was raised
End Sub

'@TestMethod
Public Sub TestRecipeIngredient_Init_RaisesErrorIfServingsNegative()
    ' Tests if Init raises the correct error when passed a negative servingsNeeded value.
    ' Arrange
    Dim servingsNeeded As Double: servingsNeeded = -1#
    Dim expectedError As Long: expectedError = vbObjectError + 602
    
    ' Act
    On Error Resume Next
    Err.Clear
    testIngredient.Init testProduct, servingsNeeded ' This line should raise error
    Dim actualError As Long: actualError = Err.Number
    On Error GoTo 0
    
    ' Assert
    Assert.AreEqual expectedError, actualError, "Expected error " & expectedError & " when servingsNeeded is negative."
    If actualError = 0 Then Assert.Fail "Error was expected but did not occur."
End Sub

'@TestMethod
Public Sub TestRecipeIngredient_Init_RaisesErrorIfProductServingsIsZero()
    ' Tests if Init raises the correct error if the Product's servings property is zero.
    ' Arrange
    Dim servingsNeeded As Double: servingsNeeded = 1#
    Dim expectedError As Long: expectedError = vbObjectError + 603
    testProduct.servings = 0 ' Set invalid product state
    
    ' Act
    On Error Resume Next
    Err.Clear
    testIngredient.Init testProduct, servingsNeeded ' This line should raise error
    Dim actualError As Long: actualError = Err.Number
    On Error GoTo 0
    
    ' Assert
    Assert.AreEqual expectedError, actualError, "Expected error " & expectedError & " when Product.servings is zero."
    If actualError = 0 Then Assert.Fail "Error was expected but did not occur."
End Sub

'@TestMethod
Public Sub TestRecipeIngredient_Init_RaisesErrorIfProductServingsIsNegative()
    ' Tests if Init raises the correct error if the Product's servings property is negative.
    ' Arrange
    Dim servingsNeeded As Double: servingsNeeded = 1#
    Dim expectedError As Long: expectedError = vbObjectError + 603 ' Assuming same error code as zero
    testProduct.servings = -10 ' Set invalid product state
    
    ' Act
    On Error Resume Next
    Err.Clear
    testIngredient.Init testProduct, servingsNeeded ' This line should raise error
    Dim actualError As Long: actualError = Err.Number
    On Error GoTo 0
    
    ' Assert
    Assert.AreEqual expectedError, actualError, "Expected error " & expectedError & " when Product.servings is negative."
    If actualError = 0 Then Assert.Fail "Error was expected but did not occur."
End Sub

'@TestMethod
Public Sub TestRecipeIngredient_Init_RaisesErrorIfProductMassIsZero()
    ' Tests if Init raises the correct error if the Product's mass property is zero.
    ' Arrange
    Dim servingsNeeded As Double: servingsNeeded = 1#
    Dim expectedError As Long: expectedError = vbObjectError + 604
    testProduct.mass = 0 ' Set invalid product state
    
    ' Act
    On Error Resume Next
    Err.Clear
    testIngredient.Init testProduct, servingsNeeded ' This line should raise error
    Dim actualError As Long: actualError = Err.Number
    On Error GoTo 0
    
    ' Assert
    Assert.AreEqual expectedError, actualError, "Expected error " & expectedError & " when Product.mass is zero."
    If actualError = 0 Then Assert.Fail "Error was expected but did not occur."
End Sub

'@TestMethod
Public Sub TestRecipeIngredient_Init_RaisesErrorIfProductMassIsNegative()
    ' Tests if Init raises the correct error if the Product's mass property is negative.
    ' Arrange
    Dim servingsNeeded As Double: servingsNeeded = 1#
    Dim expectedError As Long: expectedError = vbObjectError + 604 ' Assuming same error code as zero
    testProduct.mass = -1#  ' Set invalid product state
    
    ' Act
    On Error Resume Next
    Err.Clear
    testIngredient.Init testProduct, servingsNeeded ' This line should raise error
    Dim actualError As Long: actualError = Err.Number
    On Error GoTo 0
    
    ' Assert
    Assert.AreEqual expectedError, actualError, "Expected error " & expectedError & " when Product.mass is negative."
    If actualError = 0 Then Assert.Fail "Error was expected but did not occur."
End Sub

'@TestMethod
Public Sub TestRecipeIngredient_Init_RaisesErrorIfProductPriceIsZero()
    ' Tests if Init raises the correct error if the Product's mass property is zero.
    ' Arrange
    Dim servingsNeeded As Double: servingsNeeded = 1#
    Dim expectedError As Long: expectedError = vbObjectError + 605
    testProduct.price = 0 ' Set invalid product state
    
    ' Act
    On Error Resume Next
    Err.Clear
    testIngredient.Init testProduct, servingsNeeded ' This line should raise error
    Dim actualError As Long: actualError = Err.Number
    On Error GoTo 0
    
    ' Assert
    Assert.AreEqual expectedError, actualError, "Expected error " & expectedError & " when Product.mass is zero."
    If actualError = 0 Then Assert.Fail "Error was expected but did not occur."
End Sub

'@TestMethod
Public Sub TestRecipeIngredient_Init_RaisesErrorIfProductPriceIsNegative()
    ' Tests if Init raises the correct error if the Product's mass property is negative.
    ' Arrange
    Dim servingsNeeded As Double: servingsNeeded = 1#
    Dim expectedError As Long: expectedError = vbObjectError + 605 ' Assuming same error code as zero
    testProduct.price = -1#  ' Set invalid product state
    
    ' Act
    On Error Resume Next
    Err.Clear
    testIngredient.Init testProduct, servingsNeeded ' This line should raise error
    Dim actualError As Long: actualError = Err.Number
    On Error GoTo 0
    
    ' Assert
    Assert.AreEqual expectedError, actualError, "Expected error " & expectedError & " when Product.mass is negative."
    If actualError = 0 Then Assert.Fail "Error was expected but did not occur."
End Sub
