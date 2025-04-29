VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Nutrient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================================================================
' Class Module: Nutrient
' Author:       Evan Scott
' Date:         4/28/2025
' Purpose:      Represents a nutrient, holding its identifier, name,
'               and description.
'=============================================================================

' --- Private Member Variables ---
' These variables store the actual data for the nutrient object.
' They are declared Private to enforce encapsulation.
Private mID As Long          ' Unique identifier for the nutrient
Private mName As String      ' Common name of the nutrient (e.g., "Potassium")
Private mDescription As String ' Brief description of the nutrient's function/role

' --- Public Properties ---
' Property procedures provide controlled access to the private member variables.

' ID Property (Read/Write)
' Allows getting and setting the nutrient's ID.
Public Property Get id() As Long
    id = mID
End Property
Public Property Let id(value As Long)
    ' Optional: Add validation here if needed (e.g., ID must be > 0)
    mID = value
End Property

' Name Property (Read/Write)
' Allows getting and setting the nutrient's name.
Public Property Get Name() As String
    Name = mName
End Property
Public Property Let Name(value As String)
    ' Optional: Add validation here if needed (e.g., Name cannot be empty)
    mName = value
End Property

' Description Property (Read/Write)
' Allows getting and setting the nutrient's description.
Public Property Get Description() As String
    Description = mDescription
End Property
Public Property Let Description(value As String)
    mDescription = value
End Property

' --- Optional: Class Events ---
' Private Sub Class_Initialize()
'     ' Code to run when a new Nutrient object is created
'     ' e.g., Set default values
' End Sub

' Private Sub Class_Terminate()
'     ' Code to run when a Nutrient object is destroyed
'     ' e.g., Clean up other objects if necessary
' End Sub

