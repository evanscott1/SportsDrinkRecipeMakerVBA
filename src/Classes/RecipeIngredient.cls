VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecipeIngredient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================================================================
' Class Module: RecipeIngredient
' Author:       Evan Scott
' Date:         April 28, 2025
' Purpose:      Represents a single ingredient within a generated Recipe,
'               linking a Product object with the calculated amount needed
'               for a single serving of the final recipe.
'=============================================================================

' --- Private Member Variables ---
' Store the state of the recipe ingredient.
Private mProduct As Product ' Reference to the Product object used as an ingredient
Private mAmountServings As Double ' The number of servings of mProduct needed
Private mAmountKg As Double       ' The calculated mass (in kg) of mProduct needed
Private mCost As Currency         ' The calculated cost of of mProduct needed

' --- Public Properties (Read-Only) ---
' Provide read-only access to the ingredient's details after initialization.

' Product Property (Read-Only)
' Returns the Product object associated with this ingredient line.
Public Property Get Product() As Product
    Set Product = mProduct
End Property

' AmountServings Property (Read-Only)
' Returns the calculated number of servings of the Product required.
Public Property Get AmountServings() As Double
    AmountServings = mAmountServings
End Property

' AmountKg Property (Read-Only)
' Returns the calculated mass (in kilograms) of the Product required.
Public Property Get AmountKg() As Double
    AmountKg = mAmountKg
End Property

'Cost Property (Read-Only)
'Returns the calculated cost of the Product required.
Public Property Get Cost() As Currency
    Cost = mCost
End Property

' --- Public Methods ---

'-----------------------------------------------------------------------------
' Init
'-----------------------------------------------------------------------------
' Purpose: Initializes the RecipeIngredient object with its associated
'          Product and the required number of servings. Calculates the
'          corresponding mass in kilograms.
' Arguments:
'   prod          (Product): The Product object being used.
'   servingsNeeded (Double): The number of servings of the Product required.
'-----------------------------------------------------------------------------
Public Sub Init(prod As Product, servingsNeeded As Double)
    On Error GoTo ErrorHandler

    ' --- Validate Inputs ---
    If prod Is Nothing Then
        Err.Raise vbObjectError + 601, "RecipeIngredient.Init", "Product object cannot be Nothing."
        GoTo ExitSub ' Exit if validation fails
    End If
    If servingsNeeded < 0 Then
        Err.Raise vbObjectError + 602, "RecipeIngredient.Init", "Servings needed cannot be negative."
        GoTo ExitSub ' Exit if validation fails
    End If
    If prod.servings <= 0 Then
         Err.Raise vbObjectError + 603, "RecipeIngredient.Init", "Product's total servings must be greater than zero to calculate mass per serving."
         GoTo ExitSub ' Exit if validation fails
    End If
     If prod.mass <= 0 Then
         Err.Raise vbObjectError + 604, "RecipeIngredient.Init", "Product's total mass must be greater than zero to calculate mass per serving."
         GoTo ExitSub ' Exit if validation fails
    End If
     If prod.price <= 0 Then
         Err.Raise vbObjectError + 605, "RecipeIngredient.Init", "Product's total price must be greater than zero."
         GoTo ExitSub ' Exit if validation fails
    End If


    ' --- Assign Values ---
    Set mProduct = prod
    mAmountServings = servingsNeeded

    ' --- Calculate Mass in Kilograms ---
    ' Formula: kg_needed = servings_needed * (total_product_kg / total_product_servings)
    mAmountKg = mAmountServings * (mProduct.mass / mProduct.servings)

    ' --- Calculate Cost ---
    ' Cost = (Servings Needed for Recipe) * (Total Product Price / Total Product Servings)
    mCost = mAmountServings * (mProduct.price / mProduct.servings)

ExitSub:
    Exit Sub

ErrorHandler:
    ' Log error or re-raise, depending on desired handling
    Debug.Print "Error in RecipeIngredient.Init: " & Err.Number & " - " & Err.Description
    ' Optionally re-raise the error:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume ExitSub ' Go to cleanup on error

End Sub


' --- Optional: Class Events ---
' Private Sub Class_Initialize()
'     ' Code to run when a new RecipeIngredient object is created
' End Sub

' Private Sub Class_Terminate()
'     ' Code to run when a RecipeIngredient object is destroyed
'     ' Release object references if necessary (though VBA usually handles this)
'     Set mProduct = Nothing
' End Sub

