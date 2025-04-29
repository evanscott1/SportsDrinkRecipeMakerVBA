Attribute VB_Name = "ModConstants"
Option Explicit

'=============================================================================
' Module: ModConstants
' Author: Evan Scott
' Date:   4/28/2025
' Purpose: Holds globally accessible constants and Enums for the project.
'=============================================================================

' --- Column Definitions using Enum ---
' Using an Enum makes column assignments automatic and easier to maintain.
' Explicitly setting the first item ensures 1-based indexing matching Excel columns.
Public Enum ProductDataColumns
    colProdId = 1 ' Start explicitly at column 1
    colProdName
    colProdPrice
    colProdMass
    colProdServings
    colNutrientId  ' ID of the nutrient associated with this row/quantity
    colMassPerServing ' Mass of this nutrient per product serving
End Enum

' --- Constants for Sheet Names ---
Public Const DASHBOARD_SHEET_NAME As String = "Dashboard"
Public Const PRODUCT_DATA_SHEET_NAME As String = "ProductData" ' Should match usage in modProductData
Public Const RECIPE_OUTPUT_SHEET_NAME As String = "RecipeOutput"

' --- Other Public Constants or Enums can be added here ---
