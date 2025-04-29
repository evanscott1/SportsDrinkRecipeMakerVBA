Attribute VB_Name = "TestRecipe"
'@TestModule
'@Folder("Tests")
Option Explicit
Option Private Module 'Ensures Assert/Fakes aren't globally exposed if Private

'=============================================================================
' Module: TestRecipe
' Author: Evan Scott
' Date:   April 28, 2025
' Purpose: Contains unit tests for the Recipe class module using the
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
Private testRecipe As recipe ' The instance of the Recipe class being tested

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
    Debug.Print "Test Module Initialized: TestRecipe"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    ' Runs once after all tests in this module have executed.
    ' Used here to clean up the Assert and Fakes objects.
    Set Assert = Nothing
    Set Fakes = Nothing
    Debug.Print "Test Module Cleaned Up: TestRecipe"
End Sub


'@TestInitialize
Private Sub TestInitialize()
    ' Runs before every individual test method in this module.
    ' Creates a fresh instance of the Recipe class for each test.
    Set testRecipe = New recipe
End Sub

'@TestCleanup
Private Sub TestCleanup()
    ' Runs after every individual test method in this module.
    ' Destroys the Recipe class instance, ensuring test isolation.
    Set testRecipe = Nothing
End Sub

' --- Helper Function to Create a Mock RecipeIngredient ---
' Avoids needing a full Product setup for simple Recipe tests
Private Function CreateMockRecipeIngredient() As RecipeIngredient
    Dim mockRI As RecipeIngredient
    Dim mockProd As Product
    
    ' Create minimal mock product needed for RI initialization
    Set mockProd = New Product
    mockProd.mass = 1 ' kg
    mockProd.servings = 10 ' servings
    mockProd.price = 5 'dollars
    
    ' Create and initialize RI
    Set mockRI = New RecipeIngredient
    mockRI.Init mockProd, 1#  ' Initialize with 1 serving
    
    Set CreateMockRecipeIngredient = mockRI
    ' Clean up local mock product
    Set mockProd = Nothing
End Function


' --- Test Methods ---

'@TestMethod
Public Sub TestRecipe_Initialize_CreatesIngredientsCollection()
    ' Tests if the Ingredients collection is created upon Recipe object instantiation.
    ' Arrange (Done in TestInitialize)
    
    ' Act
    Dim ingredientsCol As Collection
    Set ingredientsCol = testRecipe.Ingredients ' Access the property
    
    ' Assert
    Assert.IsNotNothing ingredientsCol, "Ingredients collection should be initialized (not Nothing)."
    If Not ingredientsCol Is Nothing Then
        Assert.AreEqual CLng(0), ingredientsCol.Count, "Newly initialized Ingredients collection should be empty."
    End If
End Sub

'@TestMethod
Public Sub TestRecipe_AddIngredient_AddsToCollection()
    ' Tests if AddIngredient successfully adds a valid RecipeIngredient object.
    ' Arrange
    Dim ingredientToAdd As RecipeIngredient
    Set ingredientToAdd = CreateMockRecipeIngredient() ' Use helper to create a valid RI
    
    ' Act
    testRecipe.AddIngredient ingredientToAdd
    
    ' Assert
    Assert.AreEqual CLng(1), testRecipe.Ingredients.Count, "Ingredients collection should have 1 item after adding."
    ' Optionally, verify the added item is the same object
    Dim addedItem As RecipeIngredient
    Set addedItem = testRecipe.Ingredients(1) ' Collection is 1-based
    Assert.AreSame ingredientToAdd, addedItem, "The item added should be the same object passed to AddIngredient."
    
    ' Clean up mock
    Set ingredientToAdd = Nothing
    Set addedItem = Nothing
End Sub

'@TestMethod
Public Sub TestRecipe_AddIngredient_HandlesMultipleAdds()
    ' Tests if multiple ingredients can be added correctly.
    ' Arrange
    Dim ingredient1 As RecipeIngredient: Set ingredient1 = CreateMockRecipeIngredient()
    Dim ingredient2 As RecipeIngredient: Set ingredient2 = CreateMockRecipeIngredient()
    
    ' Act
    testRecipe.AddIngredient ingredient1
    testRecipe.AddIngredient ingredient2
    
    ' Assert
    Assert.AreEqual CLng(2), testRecipe.Ingredients.Count, "Ingredients collection should have 2 items after adding two."
    
    ' Clean up mocks
    Set ingredient1 = Nothing
    Set ingredient2 = Nothing
End Sub

'@TestMethod
Public Sub TestRecipe_AddIngredient_HandlesNothingGracefully()
    ' Tests if AddIngredient handles being passed Nothing without erroring or adding.
    ' Arrange
    Dim initialCount As Long
    initialCount = testRecipe.Ingredients.Count ' Should be 0
    
    ' Act
    ' No error handling needed here, the method should just exit sub
    testRecipe.AddIngredient Nothing
    
    ' Assert
    Assert.AreEqual initialCount, testRecipe.Ingredients.Count, "Ingredients count should remain unchanged after attempting to add Nothing."
End Sub

'@TestMethod
Public Sub TestRecipe_IngredientsProperty_ReturnsCollection()
    ' Tests if the Ingredients property getter returns a Collection object.
    ' Arrange (Done in TestInitialize)
    
    ' Act
    Dim returnedValue As Variant ' Use Variant to check type
    Set returnedValue = testRecipe.Ingredients
    
    ' Assert
    Assert.IsTrue TypeOf returnedValue Is Collection, "Ingredients property should return an object of type Collection."
    Assert.IsNotNothing returnedValue, "Ingredients property should not return Nothing."
End Sub


