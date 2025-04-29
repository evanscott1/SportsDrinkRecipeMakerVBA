Attribute VB_Name = "TestRecipeGenerator"
'@TestModule
'@Folder("Tests")
Option Explicit
Option Private Module 'Ensures Assert/Fakes aren't globally exposed if Private

'=============================================================================
' Module: TestRecipeGenerator
' Author: Evan Scott
' Date:   April 28, 2025
' Purpose: Contains integration tests for the modRecipeGenerator module,
'          specifically the GenerateRecipe function, using the
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
' Dictionaries/Collections used as input for GenerateRecipe
Private testTargetNutrients As Scripting.Dictionary
Private testAvailableProducts As Collection
Private testExcludedProductIDs As Scripting.Dictionary

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
    Debug.Print "Test Module Initialized: TestRecipeGenerator"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    ' Runs once after all tests in this module have executed.
    ' Used here to clean up the Assert and Fakes objects.
    Set Assert = Nothing
    Set Fakes = Nothing
    Debug.Print "Test Module Cleaned Up: TestRecipeGenerator"
End Sub


'@TestInitialize
Private Sub TestInitialize()
    ' Runs before every individual test method in this module.
    ' Creates fresh instances of the input collections/dictionaries.
    Set testTargetNutrients = New Scripting.Dictionary
    Set testAvailableProducts = New Collection
    Set testExcludedProductIDs = New Scripting.Dictionary
End Sub

'@TestCleanup
Private Sub TestCleanup()
    ' Runs after every individual test method in this module.
    ' Destroys the test input objects.
    Set testTargetNutrients = Nothing
    Set testAvailableProducts = Nothing
    Set testExcludedProductIDs = Nothing
End Sub

' --- Helper Functions for Creating Test Data ---

' Creates a Product object with specified details and nutrient quantities
Private Function CreateTestProduct(id As Long, Name As String, massKg As Double, servings As Long, price As Currency, ParamArray nutrientData() As Variant) As Product
    Dim prod As Product
    Dim nq As NutrientQuantity
    Dim i As Integer
    
    Set prod = New Product
    prod.id = id
    prod.ProductName = Name
    prod.mass = massKg
    prod.servings = servings
    prod.price = price
    
    ' nutrientData should be pairs of (NutrientID As Long, MassPerServingKg As Double)
    If LBound(nutrientData) <= UBound(nutrientData) Then ' Check if array has elements
        If (UBound(nutrientData) - LBound(nutrientData) + 1) Mod 2 <> 0 Then
             Err.Raise vbObjectError + 1001, "CreateTestProduct", "Nutrient data must be provided in pairs (ID, MassPerServing)."
        End If
        For i = LBound(nutrientData) To UBound(nutrientData) Step 2
            Set nq = New NutrientQuantity
            nq.nutrientID = CLng(nutrientData(i))
            nq.MassPerServing = CDbl(nutrientData(i + 1))
            prod.NutrientQuantities.Add nq
        Next i
    End If
    
    Set CreateTestProduct = prod
    Set prod = Nothing ' Let caller manage lifetime
    Set nq = Nothing
End Function

' Helper to find a RecipeIngredient in a Recipe by Product ID
Private Function FindIngredientByProductID(recipe As recipe, productID As Long) As RecipeIngredient
    Dim ri As RecipeIngredient
    Set FindIngredientByProductID = Nothing ' Default if not found
    
    If recipe Is Nothing Then Exit Function
    If recipe.Ingredients Is Nothing Then Exit Function
    
    For Each ri In recipe.Ingredients
        If Not ri.Product Is Nothing Then
            If ri.Product.id = productID Then
                Set FindIngredientByProductID = ri
                Exit Function ' Found it
            End If
        End If
    Next ri
End Function

' --- Test Methods ---

'@TestMethod
Public Sub GenerateRecipe_BasicCase_SingleNutrientSingleProduct()
    ' Test: Target one nutrient, only one suitable product available.
    ' Arrange
    ' Target: 0.025 kg (25g) of Nutrient 101
    testTargetNutrients.Add key:=101, Item:=0.025
    
    ' Available: Product 1 has 0.010 kg/serving of Nutrient 101 (1kg, 100 servings)
    testAvailableProducts.Add CreateTestProduct(1, "Protein A", 1#, 100, 5, 101, 0.01)
    
    ' Expected servings: 0.025 kg needed / 0.010 kg/serving = 2.5 servings
    Dim expectedServings As Double: expectedServings = 2.5
    Dim tolerance As Double: tolerance = 0.000001 ' Tolerance for floating point comparison
    
    ' Act
    Dim resultRecipe As recipe
    Set resultRecipe = ModRecipeGenerator.GenerateRecipe(testTargetNutrients, testAvailableProducts, testExcludedProductIDs)
    
    ' Assert
    Assert.IsNotNothing resultRecipe, "Recipe object should be generated."
    If resultRecipe Is Nothing Then Exit Sub
    
    Assert.AreEqual CLng(1), resultRecipe.Ingredients.Count, "Recipe should contain exactly 1 ingredient."
    If resultRecipe.Ingredients.Count <> 1 Then Exit Sub
    
    Dim ingredient As RecipeIngredient
    Set ingredient = resultRecipe.Ingredients(1)
    Assert.AreEqual CLng(1), ingredient.Product.id, "Ingredient should be Product ID 1."
    
    ' --- UPDATED: Manual comparison for double ---
    Dim actualServings As Double: actualServings = ingredient.AmountServings
    Dim difference As Double: difference = Abs(expectedServings - actualServings)
    Assert.IsTrue difference <= tolerance, "Ingredient servings calculation is incorrect. (Expected: " & expectedServings & ", Actual: " & actualServings & ", Diff: " & difference & ")"
    ' Assert.AreEqual expectedServings, ingredient.AmountServings, tolerance, "Ingredient servings calculation is incorrect." ' Original
End Sub

'@TestMethod
Public Sub GenerateRecipe_MultipleNutrients_SingleProductLimiting()
    ' Test: Target two nutrients, both present in one product, amount determined by limiting nutrient.
    ' Arrange
    ' Target: 0.020 kg (20g) of Nutrient 101, 0.005 kg (5g) of Nutrient 102
    testTargetNutrients.Add key:=101, Item:=0.02
    testTargetNutrients.Add key:=102, Item:=0.005
    
    ' Available: Product 1 has 0.010 kg/serving of N101 and 0.002 kg/serving of N102 (1kg, 100 servings)
    testAvailableProducts.Add CreateTestProduct(1, "MultiNutrient A", 1#, 100, 5, 101, 0.01, 102, 0.002)
    
    ' Servings needed for N101: 0.020 / 0.010 = 2.0 servings
    ' Servings needed for N102: 0.005 / 0.002 = 2.5 servings -> N102 is limiting
    Dim expectedServings As Double: expectedServings = 2.5
    Dim tolerance As Double: tolerance = 0.000001
    
    ' Act
    Dim resultRecipe As recipe
    Set resultRecipe = ModRecipeGenerator.GenerateRecipe(testTargetNutrients, testAvailableProducts, testExcludedProductIDs)
    
    ' Assert
    Assert.IsNotNothing resultRecipe, "Recipe object should be generated."
    If resultRecipe Is Nothing Then Exit Sub
    
    Assert.AreEqual CLng(1), resultRecipe.Ingredients.Count, "Recipe should contain exactly 1 ingredient."
    If resultRecipe.Ingredients.Count <> 1 Then Exit Sub
    
    Dim ingredient As RecipeIngredient
    Set ingredient = resultRecipe.Ingredients(1)
    Assert.AreEqual CLng(1), ingredient.Product.id, "Ingredient should be Product ID 1."
    
    ' --- UPDATED: Manual comparison for double ---
    Dim actualServings As Double: actualServings = ingredient.AmountServings
    Dim difference As Double: difference = Abs(expectedServings - actualServings)
    Assert.IsTrue difference <= tolerance, "Ingredient servings should be based on limiting nutrient (N102). (Expected: " & expectedServings & ", Actual: " & actualServings & ", Diff: " & difference & ")"
    ' Assert.AreEqual expectedServings, ingredient.AmountServings, tolerance, "Ingredient servings should be based on limiting nutrient (N102)." ' Original
End Sub

'@TestMethod
Public Sub GenerateRecipe_MultipleProducts_SelectsHighestConcentration()
    ' Test: Target one nutrient, multiple products have it, selects the one with highest MassPerServing.
    ' Arrange
    ' Target: 0.015 kg (15g) of Nutrient 101
    testTargetNutrients.Add key:=101, Item:=0.015
    
    ' Available:
    ' Product 1: 0.010 kg/serving of N101 (Concentration = 0.010)
    ' Product 2: 0.005 kg/serving of N101 (Concentration = 0.005)
    ' Product 3: 0.012 kg/serving of N101 (Concentration = 0.012) <- Highest concentration
    testAvailableProducts.Add CreateTestProduct(1, "Low Conc", 1#, 100, 5, 101, 0.01)
    testAvailableProducts.Add CreateTestProduct(2, "Lowest Conc", 1#, 100, 5, 101, 0.005)
    testAvailableProducts.Add CreateTestProduct(3, "High Conc", 1#, 100, 5, 101, 0.012)
    
    ' Expected servings (using Product 3): 0.015 kg needed / 0.012 kg/serving = 1.25 servings
    Dim expectedServings As Double: expectedServings = 1.25
    Dim expectedProductID As Long: expectedProductID = 3
    Dim tolerance As Double: tolerance = 0.000001
    
    ' Act
    Dim resultRecipe As recipe
    Set resultRecipe = ModRecipeGenerator.GenerateRecipe(testTargetNutrients, testAvailableProducts, testExcludedProductIDs)
    
    ' Assert
    Assert.IsNotNothing resultRecipe, "Recipe object should be generated."
    If resultRecipe Is Nothing Then Exit Sub
    
    Assert.AreEqual CLng(1), resultRecipe.Ingredients.Count, "Recipe should contain exactly 1 ingredient."
    If resultRecipe.Ingredients.Count <> 1 Then Exit Sub
    
    Dim ingredient As RecipeIngredient
    Set ingredient = resultRecipe.Ingredients(1)
    Assert.AreEqual expectedProductID, ingredient.Product.id, "Ingredient should be Product ID " & expectedProductID & " (highest concentration)."
    
    ' --- UPDATED: Manual comparison for double ---
    Dim actualServings As Double: actualServings = ingredient.AmountServings
    Dim difference As Double: difference = Abs(expectedServings - actualServings)
    Assert.IsTrue difference <= tolerance, "Ingredient servings calculation is incorrect. (Expected: " & expectedServings & ", Actual: " & actualServings & ", Diff: " & difference & ")"
    ' Assert.AreEqual expectedServings, ingredient.AmountServings, tolerance, "Ingredient servings calculation is incorrect." ' Original
End Sub

'@TestMethod
Public Sub GenerateRecipe_CombinedNutrients_CalculatesCorrectly()
    ' Test: Target two nutrients, requiring two products, where the first product contributes to the second nutrient.
    ' Arrange
    ' Target: 0.020 kg (20g) of Nutrient 101, 0.010 kg (10g) of Nutrient 102
    testTargetNutrients.Add key:=101, Item:=0.02
    testTargetNutrients.Add key:=102, Item:=0.01
    
    ' Available:
    ' Product 1: N101=0.010 kg/serv, N102=0.002 kg/serv (Best for N101)
    ' Product 2: N102=0.005 kg/serv (Only has N102)
    testAvailableProducts.Add CreateTestProduct(1, "Prod X", 1#, 100, 5, 101, 0.01, 102, 0.002)
    testAvailableProducts.Add CreateTestProduct(2, "Prod Y", 1#, 100, 5, 102, 0.005)
    
    ' Expected Logic:
    ' 1. Meet N101: Need 0.020 / 0.010 = 2.0 servings of Product 1.
    ' 2. Product 1 provides: 2.0 serv * 0.002 kg/serv = 0.004 kg of N102.
    ' 3. Remaining N102 needed: 0.010 kg - 0.004 kg = 0.006 kg.
    ' 4. Meet remaining N102 with Product 2: Need 0.006 kg / 0.005 kg/serv = 1.2 servings of Product 2.
    Dim expectedServingsP1 As Double: expectedServingsP1 = 2#
    Dim expectedServingsP2 As Double: expectedServingsP2 = 1.2
    Dim tolerance As Double: tolerance = 0.000001
    
    ' Act
    Dim resultRecipe As recipe
    Set resultRecipe = ModRecipeGenerator.GenerateRecipe(testTargetNutrients, testAvailableProducts, testExcludedProductIDs)
    
    ' Assert
    Assert.IsNotNothing resultRecipe, "Recipe object should be generated."
    If resultRecipe Is Nothing Then Exit Sub
    
    Assert.AreEqual CLng(2), resultRecipe.Ingredients.Count, "Recipe should contain exactly 2 ingredients."
    If resultRecipe.Ingredients.Count <> 2 Then Exit Sub
    
    ' Check Ingredient 1 (Product 1)
    Dim ingredient1 As RecipeIngredient: Set ingredient1 = FindIngredientByProductID(resultRecipe, 1)
    Assert.IsNotNothing ingredient1, "Recipe should contain Product ID 1."
    If Not ingredient1 Is Nothing Then
        ' --- UPDATED: Manual comparison for double ---
        Dim actualServingsP1 As Double: actualServingsP1 = ingredient1.AmountServings
        Dim differenceP1 As Double: differenceP1 = Abs(expectedServingsP1 - actualServingsP1)
        Assert.IsTrue differenceP1 <= tolerance, "Product 1 servings calculation is incorrect. (Expected: " & expectedServingsP1 & ", Actual: " & actualServingsP1 & ", Diff: " & differenceP1 & ")"
        ' Assert.AreEqual expectedServingsP1, ingredient1.AmountServings, tolerance, "Product 1 servings calculation is incorrect." ' Original
    End If
    
    ' Check Ingredient 2 (Product 2)
    Dim ingredient2 As RecipeIngredient: Set ingredient2 = FindIngredientByProductID(resultRecipe, 2)
    Assert.IsNotNothing ingredient2, "Recipe should contain Product ID 2."
     If Not ingredient2 Is Nothing Then
        ' --- UPDATED: Manual comparison for double ---
        Dim actualServingsP2 As Double: actualServingsP2 = ingredient2.AmountServings
        Dim differenceP2 As Double: differenceP2 = Abs(expectedServingsP2 - actualServingsP2)
        Assert.IsTrue differenceP2 <= tolerance, "Product 2 servings calculation is incorrect. (Expected: " & expectedServingsP2 & ", Actual: " & actualServingsP2 & ", Diff: " & differenceP2 & ")"
        ' Assert.AreEqual expectedServingsP2, ingredient2.AmountServings, tolerance, "Product 2 servings calculation is incorrect." ' Original
    End If
    
End Sub

'@TestMethod
Public Sub GenerateRecipe_Exclusion_SelectsNextBestProduct()
    ' Test: Target one nutrient, exclude the best product, ensure the next best is chosen.
    ' Arrange
    ' Target: 0.015 kg (15g) of Nutrient 101
    testTargetNutrients.Add key:=101, Item:=0.015
    
    ' Available:
    ' Product 1: 0.010 kg/serving of N101 (Next Best)
    ' Product 2: 0.005 kg/serving of N101
    ' Product 3: 0.012 kg/serving of N101 (Best, but will be excluded)
    testAvailableProducts.Add CreateTestProduct(1, "Next Best", 1#, 100, 5, 101, 0.01)
    testAvailableProducts.Add CreateTestProduct(2, "Lowest Conc", 1#, 100, 5, 101, 0.005)
    testAvailableProducts.Add CreateTestProduct(3, "Best Conc", 1#, 100, 101, 5, 0.012)
    
    ' Exclude Product 3
    testExcludedProductIDs.Add key:=3, Item:=True
    
    ' Expected servings (using Product 1): 0.015 kg needed / 0.010 kg/serving = 1.5 servings
    Dim expectedServings As Double: expectedServings = 1.5
    Dim expectedProductID As Long: expectedProductID = 1
    Dim tolerance As Double: tolerance = 0.000001
    
    ' Act
    Dim resultRecipe As recipe
    Set resultRecipe = ModRecipeGenerator.GenerateRecipe(testTargetNutrients, testAvailableProducts, testExcludedProductIDs)
    
    ' Assert
    Assert.IsNotNothing resultRecipe, "Recipe object should be generated."
    If resultRecipe Is Nothing Then Exit Sub
    
    Assert.AreEqual CLng(1), resultRecipe.Ingredients.Count, "Recipe should contain exactly 1 ingredient."
    If resultRecipe.Ingredients.Count <> 1 Then Exit Sub
    
    Dim ingredient As RecipeIngredient
    Set ingredient = resultRecipe.Ingredients(1)
    Assert.AreEqual expectedProductID, ingredient.Product.id, "Ingredient should be Product ID " & expectedProductID & " (next best after exclusion)."
    
    ' --- UPDATED: Manual comparison for double ---
    Dim actualServings As Double: actualServings = ingredient.AmountServings
    Dim difference As Double: difference = Abs(expectedServings - actualServings)
    Assert.IsTrue difference <= tolerance, "Ingredient servings calculation is incorrect. (Expected: " & expectedServings & ", Actual: " & actualServings & ", Diff: " & difference & ")"
    ' Assert.AreEqual expectedServings, ingredient.AmountServings, tolerance, "Ingredient servings calculation is incorrect." ' Original
End Sub

'@TestMethod
Public Sub GenerateRecipe_TargetAlreadyMet_DoesNotAddIngredient()
    ' Test: Target a nutrient that is already fulfilled by previously added ingredients.
    ' Arrange
    ' Target: 0.010 kg (10g) of Nutrient 101, 0.001 kg (1g) of Nutrient 102
    testTargetNutrients.Add key:=101, Item:=0.01
    testTargetNutrients.Add key:=102, Item:=0.001 ' Low target for N102
    
    ' Available:
    ' Product 1: N101=0.010 kg/serv, N102=0.002 kg/serv (Best for N101)
    ' Product 2: N102=0.005 kg/serv (Only has N102)
    testAvailableProducts.Add CreateTestProduct(1, "Prod X", 1#, 100, 5, 101, 0.01, 102, 0.002)
    testAvailableProducts.Add CreateTestProduct(2, "Prod Y", 1#, 100, 5, 102, 0.005)
    
    ' Expected Logic:
    ' 1. Meet N101: Need 0.010 / 0.010 = 1.0 serving of Product 1.
    ' 2. Product 1 provides: 1.0 serv * 0.002 kg/serv = 0.002 kg of N102.
    ' 3. N102 needed: 0.001 kg. Already have 0.002 kg. Target is met/exceeded.
    ' 4. No need to add Product 2.
    Dim expectedServingsP1 As Double: expectedServingsP1 = 1#
    Dim tolerance As Double: tolerance = 0.000001
    
    ' Act
    Dim resultRecipe As recipe
    Set resultRecipe = ModRecipeGenerator.GenerateRecipe(testTargetNutrients, testAvailableProducts, testExcludedProductIDs)
    
    ' Assert
    Assert.IsNotNothing resultRecipe, "Recipe object should be generated."
    If resultRecipe Is Nothing Then Exit Sub
    
    Assert.AreEqual CLng(1), resultRecipe.Ingredients.Count, "Recipe should contain only 1 ingredient (Product 1)."
    If resultRecipe.Ingredients.Count <> 1 Then Exit Sub
    
    ' Check Ingredient 1 (Product 1)
    Dim ingredient1 As RecipeIngredient: Set ingredient1 = FindIngredientByProductID(resultRecipe, 1)
    Assert.IsNotNothing ingredient1, "Recipe should contain Product ID 1."
    If Not ingredient1 Is Nothing Then
        ' --- UPDATED: Manual comparison for double ---
        Dim actualServingsP1 As Double: actualServingsP1 = ingredient1.AmountServings
        Dim differenceP1 As Double: differenceP1 = Abs(expectedServingsP1 - actualServingsP1)
        Assert.IsTrue differenceP1 <= tolerance, "Product 1 servings calculation is incorrect. (Expected: " & expectedServingsP1 & ", Actual: " & actualServingsP1 & ", Diff: " & differenceP1 & ")"
        ' Assert.AreEqual expectedServingsP1, ingredient1.AmountServings, tolerance, "Product 1 servings calculation is incorrect." ' Original
    End If
    
    ' Check Ingredient 2 (Product 2) is NOT present
    Dim ingredient2 As RecipeIngredient: Set ingredient2 = FindIngredientByProductID(resultRecipe, 2)
    Assert.IsNothing ingredient2, "Recipe should NOT contain Product ID 2 as target N102 was already met."

End Sub

'@TestMethod
Public Sub GenerateRecipe_ImpossibleRecipe_ReturnsNothing()
    ' Test: Target a nutrient for which no available product exists.
    ' Arrange
    ' Target: 0.025 kg (25g) of Nutrient 999 (non-existent in products)
    testTargetNutrients.Add key:=999, Item:=0.025
    
    ' Available: Product 1 has Nutrient 101
    testAvailableProducts.Add CreateTestProduct(1, "Protein A", 1#, 100, 5, 101, 0.01)
    
    ' Act
    Dim resultRecipe As recipe
    Set resultRecipe = ModRecipeGenerator.GenerateRecipe(testTargetNutrients, testAvailableProducts, testExcludedProductIDs)
    
    ' Assert
    Assert.IsNothing resultRecipe, "Recipe object should be Nothing when a target cannot be met."
End Sub

'@TestMethod
Public Sub GenerateRecipe_ZeroTarget_DoesNotAddIngredient()
    ' Test: Target 0 kg of a nutrient, ensure no ingredient is added specifically for it.
    ' Arrange
    ' Target: 0.0 kg of Nutrient 101, 0.01 kg of Nutrient 102
    testTargetNutrients.Add key:=101, Item:=0#
    testTargetNutrients.Add key:=102, Item:=0.01
    
    ' Available:
    ' Product 1: N101=0.010 kg/serv
    ' Product 2: N102=0.005 kg/serv
    testAvailableProducts.Add CreateTestProduct(1, "Prod N101", 1#, 100, 5, 101, 0.01)
    testAvailableProducts.Add CreateTestProduct(2, "Prod N102", 1#, 100, 5, 102, 0.005)
    
    ' Expected Logic:
    ' 1. Target N101 is 0, skip.
    ' 2. Target N102 needs 0.01 / 0.005 = 2.0 servings of Product 2.
    Dim expectedServingsP2 As Double: expectedServingsP2 = 2#
    Dim tolerance As Double: tolerance = 0.000001
    
    ' Act
    Dim resultRecipe As recipe
    Set resultRecipe = ModRecipeGenerator.GenerateRecipe(testTargetNutrients, testAvailableProducts, testExcludedProductIDs)
    
    ' Assert
    Assert.IsNotNothing resultRecipe, "Recipe object should be generated."
    If resultRecipe Is Nothing Then Exit Sub
    
    Assert.AreEqual CLng(1), resultRecipe.Ingredients.Count, "Recipe should contain only 1 ingredient (Product 2)."
    If resultRecipe.Ingredients.Count <> 1 Then Exit Sub
    
    ' Check Ingredient 2 (Product 2)
    Dim ingredient2 As RecipeIngredient: Set ingredient2 = FindIngredientByProductID(resultRecipe, 2)
    Assert.IsNotNothing ingredient2, "Recipe should contain Product ID 2."
     If Not ingredient2 Is Nothing Then
        ' --- UPDATED: Manual comparison for double ---
        Dim actualServingsP2 As Double: actualServingsP2 = ingredient2.AmountServings
        Dim differenceP2 As Double: differenceP2 = Abs(expectedServingsP2 - actualServingsP2)
        Assert.IsTrue differenceP2 <= tolerance, "Product 2 servings calculation is incorrect. (Expected: " & expectedServingsP2 & ", Actual: " & actualServingsP2 & ", Diff: " & differenceP2 & ")"
        ' Assert.AreEqual expectedServingsP2, ingredient2.AmountServings, tolerance, "Product 2 servings calculation is incorrect." ' Original
    End If
    
    ' Check Ingredient 1 (Product 1) is NOT present
    Dim ingredient1 As RecipeIngredient: Set ingredient1 = FindIngredientByProductID(resultRecipe, 1)
    Assert.IsNothing ingredient1, "Recipe should NOT contain Product ID 1 as its target was zero."
End Sub


