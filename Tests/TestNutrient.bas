Attribute VB_Name = "TestNutrient"
'@TestModule
'@Folder("Tests")
Option Explicit
Option Private Module ' Ensures Assert/Fakes aren't globally exposed if Private

'=============================================================================
' Module: TestNutrient
' Author: Evan Scott
' Date:   4/28/2025
' Purpose: Contains unit tests for the Nutrient class module using the
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

' --- Module-Level Test Variables ---
Private TestNutrient As nutrient ' The instance of the Nutrient class being tested

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
    Debug.Print "Test Module Initialized: TestNutrient"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    ' Runs once after all tests in this module have executed.
    ' Used here to clean up the Assert and Fakes objects.
    Set Assert = Nothing
    Set Fakes = Nothing
    Debug.Print "Test Module Cleaned Up: TestNutrient"
End Sub


'@TestInitialize
Private Sub TestInitialize()
    ' Runs before every individual test method in this module.
    ' Used here to create a fresh instance of the Nutrient class for each test.
    Set TestNutrient = New nutrient
End Sub

'@TestCleanup
Private Sub TestCleanup()
    ' Runs after every individual test method in this module.
    ' Used here to destroy the Nutrient class instance, ensuring test isolation.
    Set TestNutrient = Nothing
End Sub

' --- Test Methods ---

'@TestMethod
Public Sub TestNutrient_CanSetAndGetID()
    ' Tests if the ID property correctly stores and retrieves a value.
    ' Arrange: Define the expected ID value.
    Dim expectedID As Long
    expectedID = 12345
    
    ' Act: Set the ID property and then get it back.
    TestNutrient.id = expectedID
    Dim actualId As Long
    actualId = TestNutrient.id
    
    ' Assert: Verify the retrieved value matches the expected value.
    Assert.AreEqual expectedID, actualId, "ID property should return the value that was set."
End Sub

'@TestMethod
Public Sub TestNutrient_CanSetAndGetName()
    ' Tests if the Name property correctly stores and retrieves a value.
    ' Arrange: Define the expected Name value.
    Dim expectedName As String
    expectedName = "Nutrient123"
    
    ' Act: Set the Name property and then get it back.
    TestNutrient.Name = expectedName
    Dim actualName As String
    actualName = TestNutrient.Name
    
    ' Assert: Verify the retrieved value matches the expected value.
    Assert.AreEqual expectedName, actualName, "Name property should return the value that was set."
End Sub

'@TestMethod
Public Sub TestNutrient_CanSetAndGetDescription()
    ' Tests if the Description property correctly stores and retrieves a value.
    ' Arrange: Define the expected Description value.
    Dim expectedDescription As String
    expectedDescription = "A description"
    
    ' Act: Set the Description property and then get it back.
    TestNutrient.Description = expectedDescription
    Dim actualDescription As String
    actualDescription = TestNutrient.Description
    
    ' Assert: Verify the retrieved value matches the expected value.
    Assert.AreEqual expectedDescription, actualDescription, "Description property should return the value that was set."
End Sub

'@TestMethod
Public Sub TestNutrient_DefaultIDIsZero()
    ' Tests if the default value of the ID property is 0 upon object creation.
    ' Arrange: Define the expected default ID.
    Dim expectedDefaultId As Long
    expectedDefaultId = 0 ' Default value for Long
    
    ' Act: Get the ID property from the newly created object (in TestInitialize).
    Dim actualDefaultId As Long
    actualDefaultId = TestNutrient.id
    
    ' Assert: Verify the retrieved value matches the expected default.
    Assert.AreEqual expectedDefaultId, actualDefaultId, "Default ID should be 0."
End Sub

'@TestMethod
Public Sub TestNutrient_DefaultNameIsEmptyString()
    ' Tests if the default value of the Name property is an empty string upon object creation.
    ' Arrange: Define the expected default Name.
    Dim expectedDefaultName As String
    expectedDefaultName = vbNullString ' Default value for String
    
    ' Act: Get the Name property from the newly created object.
    Dim actualDefaultName As String
    actualDefaultName = TestNutrient.Name
    
    ' Assert: Verify the retrieved value matches the expected default.
    Assert.AreEqual expectedDefaultName, actualDefaultName, "Default Name should be an empty string."
End Sub

'@TestMethod
Public Sub TestNutrient_DefaultDescriptionIsEmptyString()
    ' Tests if the default value of the Description property is an empty string upon object creation.
    ' Arrange: Define the expected default Description.
    Dim expectedDefaultDescription As String
    expectedDefaultDescription = vbNullString ' Default value for String
    
    ' Act: Get the Description property from the newly created object.
    Dim actualDefaultDescription As String
    actualDefaultDescription = TestNutrient.Description
    
    ' Assert: Verify the retrieved value matches the expected default.
    Assert.AreEqual expectedDefaultDescription, actualDefaultDescription, "Default Description should be an empty string."
End Sub

