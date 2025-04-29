VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NutrientQuantity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================================================================
' Class Module: NutrientQuantity
' Author:       Evan Scott
' Date:         4/28/2025
' Purpose:      Represents the quantity of a specific nutrient associated
'               with a Product, storing the nutrient's ID and the mass
'               of that nutrient per serving of the product.
'=============================================================================

' --- Private Member Variables ---
' These variables store the actual data for the nutrient quantity object.
' They are declared Private to enforce encapsulation.
Private mNutrientID As Long      ' Foreign key referencing the ID of the associated Nutrient
Private mMassPerServing As Double ' The mass (e.g., in grams, mg) of this nutrient per product serving

' --- Public Properties ---
' Property procedures provide controlled access to the private member variables.

' NutrientID Property (Read/Write)
' Allows getting and setting the ID of the nutrient this quantity refers to.
Public Property Get nutrientID() As Long
    nutrientID = mNutrientID
End Property
Public Property Let nutrientID(value As Long)
    ' Optional: Add validation here if needed (e.g., ID must be > 0)
    mNutrientID = value
End Property

' MassPerServing Property (Read/Write)
' Allows getting and setting the mass of this nutrient per product serving.
Public Property Get MassPerServing() As Double
    MassPerServing = mMassPerServing
End Property
Public Property Let MassPerServing(value As Double)
    ' Optional: Add validation here if needed (e.g., MassPerServing must be >= 0)
    mMassPerServing = value
End Property

' --- Optional: Class Events ---
' Private Sub Class_Initialize()
'     ' Code to run when a new NutrientQuantity object is created
'     ' e.g., Set default values
' End Sub

' Private Sub Class_Terminate()
'     ' Code to run when a NutrientQuantity object is destroyed
' End Sub


