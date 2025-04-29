Attribute VB_Name = "ModRecipeGenerator"
Option Explicit

'=============================================================================
' Module: ModRecipeGenerator
' Author: Evan Scott
' Date:   April 28, 2025
' Purpose: Contains functions related to generating recipe compositions
'          based on target nutrient profiles and available products.
'=============================================================================

' --- Dependencies ---
' Requires Class Modules: Recipe, RecipeIngredient, Product, NutrientQuantity
' Requires access to Scripting.Dictionary (Microsoft Scripting Runtime)
' Requires access to ProductDataColumns Enum (defined in modConstants)

'-----------------------------------------------------------------------------
' GenerateRecipe
'-----------------------------------------------------------------------------
' Purpose: Attempts to generate a single-serving recipe that meets specified
'          nutrient targets using a collection of available products,
'          excluding any specified products. Uses a greedy algorithm based
'          on highest nutrient concentration (MassPerServing) first.
' Arguments:
'   targetNutrients    (Scripting.Dictionary): Keys are Nutrient IDs (Long),
'                                              Items are target amounts (Double, in kg).
'   availableProducts  (Collection): Collection of Product objects to select from.
'   excludedProductIDs (Scripting.Dictionary): Keys are Product IDs (Long) to exclude,
'                                              Items can be anything (e.g., True) for quick lookup.
' Returns:
'   Recipe: A populated Recipe object if a valid recipe could be generated,
'           otherwise returns Nothing.
'-----------------------------------------------------------------------------
Public Function GenerateRecipe( _
    targetNutrients As Scripting.Dictionary, _
    availableProducts As Collection, _
    excludedProductIDs As Scripting.Dictionary _
) As recipe

    Dim generatedRecipe As recipe
    Dim currentNutrientTotals As Scripting.Dictionary ' Key=NutrientID (Long), Item=CurrentAmount (Double, in kg)
    Dim targetNutrientID As Variant ' Key from targetNutrients dictionary
    Dim targetAmountKg As Double
    Dim neededAmountKg As Double
    
    Dim prod As Product ' Loop variable for available products
    Dim bestProduct As Product ' Product selected for current target nutrient
    Dim nq As NutrientQuantity ' Loop variable for nutrient quantities within a product
    Dim bestNQ As NutrientQuantity ' The specific NQ object from bestProduct for the target nutrient
    Dim maxConcentration As Double ' Highest MassPerServing found for the target nutrient
    
    Dim servingsToAdd As Double
    Dim ri As RecipeIngredient ' RecipeIngredient object to add or update
    Dim existingRI As RecipeIngredient ' Existing RI if product already in recipe
    Dim newTotalServings As Double
    
    Dim amountAddedKg As Double ' Temp variable for calculation
    Dim nutrientKey As Long     ' Temp variable for nutrient ID
    
    On Error GoTo ErrorHandler

    ' --- Validate Inputs ---
    If targetNutrients Is Nothing Or targetNutrients.Count = 0 Then
        Debug.Print "GenerateRecipe Error: No target nutrients specified."
        GoTo GenerationFailed
    End If
    If availableProducts Is Nothing Or availableProducts.Count = 0 Then
        Debug.Print "GenerateRecipe Error: No available products provided."
        GoTo GenerationFailed
    End If
    If excludedProductIDs Is Nothing Then ' Allow empty exclusion list
        Set excludedProductIDs = New Scripting.Dictionary ' Ensure it's a valid dictionary
    End If

    ' --- Initialization ---
    Set generatedRecipe = New recipe
    Set currentNutrientTotals = New Scripting.Dictionary
    
    ' Initialize current totals for all target nutrients to zero
    For Each targetNutrientID In targetNutrients.Keys
        If Not currentNutrientTotals.Exists(targetNutrientID) Then
            currentNutrientTotals.Add key:=targetNutrientID, Item:=0#
        End If
    Next targetNutrientID

    ' --- Main Loop: Iterate through Target Nutrients ---
    ' Note: The order of processing targets might matter in a greedy algorithm.
    ' Consider sorting targets later if needed (e.g., by importance or amount).
    For Each targetNutrientID In targetNutrients.Keys
        targetAmountKg = targetNutrients.Item(targetNutrientID)
        
        ' Calculate amount still needed for this nutrient
        If Not currentNutrientTotals.Exists(targetNutrientID) Then
             currentNutrientTotals.Add targetNutrientID, 0#  ' Should have been added above, but safety check
        End If
        neededAmountKg = targetAmountKg - currentNutrientTotals.Item(targetNutrientID)

        ' If target is already met or exceeded (within a tiny tolerance for floating point), move to the next nutrient
        If neededAmountKg <= 0.00000001 Then GoTo NextTargetNutrient

        ' --- Find Best Product for Current Target ---
        Set bestProduct = Nothing
        Set bestNQ = Nothing
        maxConcentration = -1 ' Use -1 to ensure any positive concentration is initially better

        For Each prod In availableProducts
            ' Skip if product is excluded
            If excludedProductIDs.Exists(prod.id) Then GoTo NextProductLoop
            
            ' Find the nutrient quantity for the target nutrient within this product
            If Not prod.NutrientQuantities Is Nothing Then
                For Each nq In prod.NutrientQuantities
                    If nq.nutrientID = targetNutrientID Then
                        ' Check if this product has a higher concentration (MassPerServing)
                        ' And ensure MassPerServing is positive to avoid division by zero later
                        If nq.MassPerServing > maxConcentration And nq.MassPerServing > 0 Then
                            maxConcentration = nq.MassPerServing
                            Set bestProduct = prod ' Store reference to the best product found so far
                            Set bestNQ = nq      ' Store reference to the corresponding NQ
                        End If
                        Exit For ' Found the target nutrient in this product, move to next product
                    End If
                Next nq
            End If
NextProductLoop:
        Next prod

        ' --- Check if a suitable product was found ---
        If bestProduct Is Nothing Then ' maxConcentration check implicitly handled by initialization to -1
            ' Cannot meet the target for this nutrient with available, non-excluded products
            Debug.Print "GenerateRecipe Error: Cannot meet target for Nutrient ID " & targetNutrientID & ". No suitable product found with positive concentration."
            GoTo GenerationFailed ' Recipe generation fails
        End If

        ' --- Calculate Amount and Add/Update Ingredient ---
        ' Calculate servings needed based on the best product found
        servingsToAdd = neededAmountKg / bestNQ.MassPerServing
        
        ' --- UPDATED LOGIC: Check if product already exists in recipe ---
        Set existingRI = FindIngredientByProductID(generatedRecipe, bestProduct.id)
        
        If Not existingRI Is Nothing Then
            ' Product already exists - Update the existing ingredient
            newTotalServings = existingRI.AmountServings + servingsToAdd
            
            ' Create a NEW RecipeIngredient with the updated total
            Set ri = New RecipeIngredient
            ri.Init bestProduct, newTotalServings ' Re-initialize with new total servings
            
            ' Remove the OLD ingredient (requires finding its index or key)
            Dim i As Long
            For i = 1 To generatedRecipe.Ingredients.Count
                If generatedRecipe.Ingredients(i) Is existingRI Then
                    generatedRecipe.Ingredients.Remove i
                    Exit For
                End If
            Next i
            
            ' Add the NEW, updated ingredient
            generatedRecipe.AddIngredient ri
            
        Else
            ' Product is new - Add it as a new ingredient
            Set ri = New RecipeIngredient
            ri.Init bestProduct, servingsToAdd ' Init calculates AmountKg internally
            generatedRecipe.AddIngredient ri
        End If
        
        Set ri = Nothing ' Clean up local reference
        Set existingRI = Nothing

        ' --- Update Current Nutrient Totals for ALL nutrients in the added product ---
        ' This needs to reflect the contribution of the *servingsToAdd* amount,
        ' regardless of whether it was a new ingredient or added to an existing one.
        If Not bestProduct.NutrientQuantities Is Nothing Then
            For Each nq In bestProduct.NutrientQuantities
                nutrientKey = nq.nutrientID
                amountAddedKg = servingsToAdd * nq.MassPerServing ' Calculate mass of this nutrient portion added
                
                ' Update the total, adding the key if it wasn't an original target
                If Not currentNutrientTotals.Exists(nutrientKey) Then
                    currentNutrientTotals.Add key:=nutrientKey, Item:=amountAddedKg
                Else
                    currentNutrientTotals.Item(nutrientKey) = currentNutrientTotals.Item(nutrientKey) + amountAddedKg
                End If
            Next nq
        End If
        
        ' Clean up references for the next loop
        Set bestProduct = Nothing
        Set bestNQ = Nothing

NextTargetNutrient:
    Next targetNutrientID

    ' --- Generation Successful ---
    Set GenerateRecipe = generatedRecipe
    GoTo ExitFunction

GenerationFailed:
    ' If we reach here, generation failed, return Nothing
    Set GenerateRecipe = Nothing
    ' Clean up recipe object if partially created
    Set generatedRecipe = Nothing

ExitFunction:
    ' Clean up dictionaries and object variables
    Set currentNutrientTotals = Nothing
    Set prod = Nothing
    Set bestProduct = Nothing
    Set nq = Nothing
    Set bestNQ = Nothing
    Set ri = Nothing
    Set existingRI = Nothing
    Exit Function

ErrorHandler:
    MsgBox "An error occurred in GenerateRecipe:" & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical, "Recipe Generation Error"
    Resume GenerationFailed ' Go to failure cleanup on error

End Function


'-----------------------------------------------------------------------------
' FindIngredientByProductID (Helper Function)
'-----------------------------------------------------------------------------
' Purpose: Searches within a Recipe object's Ingredients collection for a
'          RecipeIngredient corresponding to a specific Product ID.
' Arguments:
'   recipe     (Recipe): The Recipe object to search within.
'   productID  (Long): The Product ID to search for.
' Returns:
'   RecipeIngredient: The matching RecipeIngredient object if found,
'                     otherwise returns Nothing.
'-----------------------------------------------------------------------------
Private Function FindIngredientByProductID(recipe As recipe, productID As Long) As RecipeIngredient
    Dim ri As RecipeIngredient
    Set FindIngredientByProductID = Nothing ' Default if not found
    
    ' Validate inputs
    If recipe Is Nothing Then Exit Function
    If recipe.Ingredients Is Nothing Then Exit Function
    
    ' Loop through ingredients
    For Each ri In recipe.Ingredients
        If Not ri.Product Is Nothing Then
            If ri.Product.id = productID Then
                Set FindIngredientByProductID = ri ' Found it
                Exit Function
            End If
        End If
    Next ri
    
    ' Not found if loop completes
    
End Function


' --- Other Helper Functions for Recipe Generation Might Go Here ---



