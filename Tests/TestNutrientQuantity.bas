Attribute VB_Name = "TestNutrientQuantity"
'@TestModule
'@Folder("Tests")
Option Explicit
Option Private Module ' Ensures Assert/Fakes aren't globally exposed if Private

'=============================================================================
' Module: TestNutrientQuantity
' Author: Evan Scott
' Date:   4/28/2025
' Purpose: Contains unit tests for the NutrientQuantity class module using the
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
Private testQuantity As NutrientQuantity ' The instance of the NutrientQuantity class being tested

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
    Debug.Print "Test Module Initialized: TestNutrientQuantity"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    ' Runs once after all tests in this module have executed.
    ' Used here to clean up the Assert and Fakes objects.
    Set Assert = Nothing
    Set Fakes = Nothing
    Debug.Print "Test Module Cleaned Up: TestNutrientQuantity"
End Sub

'@TestInitialize
Private Sub TestInitialize()
    ' Runs before every individual test method in this module.
    ' Used here to create a fresh instance of the NutrientQuantity class for each test.
    Set testQuantity = New NutrientQuantity
End Sub

'@TestCleanup
Private Sub TestCleanup()
    ' Runs after every individual test method in this module.
    ' Used here to destroy the NutrientQuantity class instance, ensuring test isolation.
    Set testQuantity = Nothing
End Sub


' --- Test Methods ---

'@TestMethod
Public Sub TestNutrientQuantity_StoresNutrientID()
    ' Tests if the nutrientID property correctly stores and retrieves a value.
    ' Arrange: Define the expected ID value.
    Dim expectedID As Long
    expectedID = 123
    
    ' Act: Set the nutrientID property and then get it back.
    testQuantity.nutrientID = expectedID
    
    ' Assert: Verify the retrieved value matches the expected value.
    Assert.AreEqual expectedID, testQuantity.nutrientID, "NutrientQuantity should store the correct Nutrient ID."
End Sub

'@TestMethod
Public Sub TestNutrientQuantity_StoresMassPerServing()
    ' Tests if the MassPerServing property correctly stores and retrieves a value.
    ' Arrange: Define the expected MassPerServing value.
    Dim expectedMass As Double
    expectedMass = 0.000456 ' Example value (e.g., 0.456 mg)

    ' Act: Set the MassPerServing property and then get it back.
    testQuantity.MassPerServing = expectedMass

    ' Assert: Verify the retrieved value matches the expected value.
    Assert.AreEqual expectedMass, testQuantity.MassPerServing, "NutrientQuantity should store the correct MassPerServing."
End Sub

'@TestMethod
Public Sub TestNutrientQuantity_DefaultNutrientIDIsZero()
    ' Tests if the default value of the nutrientID property is 0 upon object creation.
    ' Arrange: Define the expected default ID.
    Dim expectedDefaultId As Long
    expectedDefaultId = 0 ' Default value for Long
    
    ' Act: Get the nutrientID property from the newly created object (in TestInitialize).
    Dim actualDefaultId As Long
    actualDefaultId = testQuantity.nutrientID
    
    ' Assert: Verify the retrieved value matches the expected default.
    Assert.AreEqual expectedDefaultId, actualDefaultId, "Default NutrientID should be 0."
End Sub

'@TestMethod
Public Sub TestNutrientQuantity_DefaultMassPerServingIsZero()
    ' Tests if the default value of the MassPerServing property is 0 upon object creation.
    ' Arrange: Define the expected default mass.
    Dim expectedDefaultMass As Double
    expectedDefaultMass = 0 ' Default value for Double
    
    ' Act: Get the MassPerServing property from the newly created object.
    Dim actualDefaultMass As Double
    actualDefaultMass = testQuantity.MassPerServing
    
    ' Assert: Verify the retrieved value matches the expected default.
    Assert.AreEqual expectedDefaultMass, actualDefaultMass, "Default MassPerServing should be 0."
End Sub

