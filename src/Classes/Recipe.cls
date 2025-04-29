VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Recipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================================================================
' Class Module: Recipe
' Author:       Evan Scott
' Date:         April 28, 2025
' Purpose:      Represents a generated recipe, containing a collection of
'               RecipeIngredient objects that make up the recipe for a
'               single serving.
'=============================================================================

' --- Private Member Variables ---
Private mIngredients As Collection ' Holds the collection of RecipeIngredient objects

' --- Class Events ---

Private Sub Class_Initialize()
    ' Runs when a new Recipe object is created.
    ' Initializes the internal collection to hold ingredients.
    Set mIngredients = New Collection
    Debug.Print "Recipe object initialized."
End Sub

Private Sub Class_Terminate()
    ' Runs when the Recipe object is destroyed.
    ' Cleans up the internal collection object.
    Set mIngredients = Nothing
    Debug.Print "Recipe object terminated."
End Sub

' --- Public Properties ---

' Ingredients Property (Read-Only Collection)
' Provides access to the collection of RecipeIngredient objects.
' Users can iterate through this collection but cannot replace the collection itself.
Public Property Get Ingredients() As Collection
    Set Ingredients = mIngredients
End Property

' --- Public Methods ---

'-----------------------------------------------------------------------------
' AddIngredient
'-----------------------------------------------------------------------------
' Purpose: Adds a RecipeIngredient object to the recipe's internal collection.
' Arguments:
'   ingredient (RecipeIngredient): The ingredient object to add.
'-----------------------------------------------------------------------------
Public Sub AddIngredient(ingredient As RecipeIngredient)
    ' Validate input
    If ingredient Is Nothing Then
        Debug.Print "Warning (Recipe.AddIngredient): Attempted to add Nothing ingredient."
        Exit Sub
    End If
    
    ' Add to the internal collection
    If Not mIngredients Is Nothing Then
        mIngredients.Add ingredient
    Else
        ' This should ideally not happen due to Class_Initialize, but added as a safeguard
        Debug.Print "Error (Recipe.AddIngredient): Ingredients collection not initialized."
        ' Optionally raise an error here
        Err.Raise vbObjectError + 701, "Recipe.AddIngredient", "Ingredients collection not initialized."
    End If
End Sub

' --- Optional Future Methods ---

' Example: Method to calculate total nutrients (more complex)
' Public Function GetTotalNutrients() As Scripting.Dictionary
'     Dim totals As Scripting.Dictionary
'     Dim ri As RecipeIngredient
'     Dim nq As NutrientQuantity
'     Dim nutrientKey As String
'
'     Set totals = New Scripting.Dictionary
'
'     If mIngredients Is Nothing Then GoTo ExitFunction ' Safety check
'     If mIngredients.Count = 0 Then GoTo ExitFunction  ' Nothing to calculate
'
'     For Each ri In mIngredients
'         If Not ri.Product Is Nothing Then
'             If Not ri.Product.NutrientQuantities Is Nothing Then
'                 For Each nq In ri.Product.NutrientQuantities
'                     nutrientKey = CStr(nq.nutrientID)
'                     Dim currentTotal As Double
'                     If totals.Exists(nutrientKey) Then
'                         currentTotal = totals(nutrientKey)
'                     Else
'                         currentTotal = 0
'                     End If
'                     ' Add mass contribution from this ingredient: servings * mass_per_serving (both in kg)
'                     totals(nutrientKey) = currentTotal + (ri.AmountServings * nq.MassPerServing)
'                 Next nq
'             End If
'         End If
'     Next ri
'
' ExitFunction:
'     Set GetTotalNutrients = totals
'     Set totals = Nothing
'     Set ri = Nothing
'     Set nq = Nothing
' End Function


