VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================================================================
' Class Module: Product
' Author:       Evan Scott
' Date:         4/26/2025
' Purpose:      Represents a purchasable product, holding its details
'               and a collection of its nutrient quantities.
'=============================================================================

' --- Private Member Variables ---
Private mID As Long
Private mProductName As String
Private mPrice As Currency
Private mMass As Double ' Total mass of the product (e.g., in grams or kg)
Private mServings As Long
Private mNutrientQuantities As Collection ' Collection to hold NutrientQuantity objects

' --- Class Events ---

Private Sub Class_Initialize()
    ' Create a new collection object when the Product object is instantiated
    Set mNutrientQuantities = New Collection
    Debug.Print "Product object initialized."
End Sub

Private Sub Class_Terminate()
    ' Clean up the collection object when the Product object is destroyed
    Set mNutrientQuantities = Nothing
    Debug.Print "Product object terminated."
End Sub

' --- Public Properties ---

' ID Property (Read/Write)
Public Property Get id() As Long
    id = mID
End Property
Public Property Let id(value As Long)
    ' Optional: Add validation if needed (e.g., ID must be positive)
    mID = value
End Property

' ProductName Property (Read/Write)
Public Property Get ProductName() As String
    ProductName = mProductName
End Property
Public Property Let ProductName(value As String)
    ' Optional: Add validation (e.g., cannot be empty string)
    mProductName = value
End Property

' Price Property (Read/Write)
Public Property Get price() As Currency
    price = mPrice
End Property
Public Property Let price(value As Currency)
    ' Optional: Add validation (e.g., price cannot be negative)
    mPrice = value
End Property

' Mass Property (Read/Write) - Represents total product mass
Public Property Get mass() As Double
    mass = mMass
End Property
Public Property Let mass(value As Double)
    ' Optional: Add validation (e.g., mass must be positive)
    mMass = value
End Property

' Servings Property (Read/Write)
Public Property Get servings() As Long
    servings = mServings
End Property
Public Property Let servings(value As Long)
    ' Optional: Add validation (e.g., servings must be positive)
    mServings = value
End Property

' NutrientQuantities Property (Read-Only Collection)
' Provides access to the collection of NutrientQuantity objects.
' Users can Add/Remove items via the collection itself, but cannot replace the collection object.
Public Property Get NutrientQuantities() As Collection
    Set NutrientQuantities = mNutrientQuantities
End Property

