Attribute VB_Name = "ModNutrientRepository"
Option Explicit
Option Private Module

' --- Module Level Variables ---
' These dictionaries hold the nutrient data.
' s_NutrientDictionary stores Nutrient objects, keyed by their string ID.
' s_InfoDictionary is temporary, used during initialization.

#If LateBind Then
    Private s_NutrientDictionary As Object ' Scripting.Dictionary
    Private s_InfoDictionary As Object     ' Scripting.Dictionary
#Else
    Private s_NutrientDictionary As Scripting.Dictionary
    Private s_InfoDictionary As Scripting.Dictionary
#End If

'=============================================================================
' InitializeNutrientRepository
'-----------------------------------------------------------------------------
' Purpose: Populates the nutrient repository (s_NutrientDictionary) with
'          a predefined list of common nutrients and their descriptions.
'          This acts as the initial data load for the nutrient "database".
'=============================================================================
Public Sub InitializeNutrientRepository()
    
    ' Ensure dictionaries are created if they don't exist
    If s_NutrientDictionary Is Nothing Then
        #If LateBind Then
            Set s_NutrientDictionary = CreateObject("Scripting.Dictionary")
        #Else
            Set s_NutrientDictionary = New Scripting.Dictionary
        #End If
    Else
        s_NutrientDictionary.RemoveAll ' Clear existing data if re-initializing
    End If
    
    ' Use a temporary dictionary to easily define names and descriptions
    #If LateBind Then
        Set s_InfoDictionary = CreateObject("Scripting.Dictionary")
    #Else
        Set s_InfoDictionary = New Scripting.Dictionary
    #End If

    Dim nutrient As nutrient ' Use the correct class name 'Nutrient'
    Dim currentID As Long
    currentID = 0 ' Start IDs from 1

    ' --- Populate the Info Dictionary ---
    ' Format: s_InfoDictionary.Add "Nutrient Name", "Brief Description"
    
    ' ** Electrolytes **
    s_InfoDictionary.Add "Potassium", "Electrolyte, muscle function, nerve signalling"
    s_InfoDictionary.Add "Sodium", "Electrolyte, fluid balance, nerve signalling"
    s_InfoDictionary.Add "Chloride", "Electrolyte, fluid balance, stomach acid production"
    s_InfoDictionary.Add "Calcium", "Electrolyte, bone health, muscle contraction"
    s_InfoDictionary.Add "Magnesium", "Electrolyte, muscle & nerve function, energy production"
    s_InfoDictionary.Add "Phosphate", "Electrolyte, energy metabolism, bone health"
    
    ' ** Macronutrients & Energy **
    s_InfoDictionary.Add "Calories", "Unit of energy provided by food/drink"
    s_InfoDictionary.Add "Carbohydrates", "Primary energy source"
    s_InfoDictionary.Add "Sugars", "Simple carbohydrates for quick energy"
    s_InfoDictionary.Add "Dietary Fiber", "Complex carbohydrate, aids digestion"
    s_InfoDictionary.Add "Protein", "Muscle repair and growth, various bodily functions"
    s_InfoDictionary.Add "Total Fat", "Energy source, hormone production"
    s_InfoDictionary.Add "Saturated Fat", "Type of fat, intake often monitored"
    s_InfoDictionary.Add "Cholesterol", "Lipid used in cell membranes and hormone synthesis"

    ' ** Common Drink Additives / Stimulants **
    s_InfoDictionary.Add "Caffeine", "Stimulant, increases alertness and energy"
    s_InfoDictionary.Add "Taurine", "Amino acid involved in various metabolic processes"
    s_InfoDictionary.Add "L-Carnitine", "Involved in energy production from fats"
    s_InfoDictionary.Add "Glucuronolactone", "Metabolite, sometimes found in energy drinks"

    ' ** B Vitamins (Common in Energy Drinks) **
    s_InfoDictionary.Add "Vitamin B3 (Niacin)", "Energy metabolism, nervous system function"
    s_InfoDictionary.Add "Vitamin B5 (Pantothenic Acid)", "Energy metabolism, hormone synthesis"
    s_InfoDictionary.Add "Vitamin B6 (Pyridoxine)", "Protein metabolism, neurotransmitter synthesis"
    s_InfoDictionary.Add "Vitamin B12 (Cobalamin)", "Red blood cell formation, neurological function"

    ' ** Other Common Vitamins & Minerals **
    s_InfoDictionary.Add "Vitamin C (Ascorbic Acid)", "Antioxidant, immune function"
    s_InfoDictionary.Add "Vitamin D", "Calcium absorption, bone health, immune function"
    s_InfoDictionary.Add "Zinc", "Immune function, wound healing, enzyme activity"
    
    ' --- Loop through Info Dictionary to create Nutrient objects ---
    Dim key As Variant
    For Each key In s_InfoDictionary.Keys
        currentID = currentID + 1 ' Increment ID for each nutrient
        Set nutrient = New nutrient ' Create a new instance of the Nutrient class
        
        ' Populate the Nutrient object's properties
        With nutrient
            .id = currentID
            .Name = CStr(key) ' Ensure name is stored as string
            .Description = s_InfoDictionary.Item(key)
        End With
        
        ' Add the Nutrient object to the main dictionary, keyed by its string ID
        s_NutrientDictionary.Add CStr(nutrient.id), nutrient
    Next key
    
    ' Clean up temporary dictionary and loop variable
    Set nutrient = Nothing
    Set s_InfoDictionary = Nothing
    
    ' --- Optional Debug Output ---
'    Debug.Print ("Nutrient Repository Initialized. Contents:" & vbNewLine)
'    Dim nutrientKey As Variant
'    For Each nutrientKey In s_NutrientDictionary.Keys
'        Set nutrient = s_NutrientDictionary.Item(nutrientKey)
'        Debug.Print "ID: " & nutrient.id & ", Name: " & nutrient.Name & ", Desc: " & nutrient.Description
'    Next nutrientKey
'    Set nutrient = Nothing
    ' --- End Debug Output ---
    
End Sub

'=============================================================================
' GetNutrientByID
'-----------------------------------------------------------------------------
' Purpose: Retrieves a Nutrient object from the repository based on its ID.
' Arguments:
'   nutrientID (Long): The ID of the nutrient to retrieve.
' Returns:
'   Nutrient: The Nutrient object if found, otherwise Nothing.
'=============================================================================
Public Function GetNutrientByID(nutrientID As Long) As nutrient
    Dim keyString As String
    keyString = CStr(nutrientID) ' Dictionary key is stored as String

    ' Check if the repository has been initialized
    If s_NutrientDictionary Is Nothing Then
        Debug.Print "Error: Nutrient repository not initialized. Call InitializeNutrientRepository first."
        Set GetNutrientByID = Nothing
        Exit Function
    End If
    
    ' Check if the key exists
    If s_NutrientDictionary.Exists(keyString) Then
        Set GetNutrientByID = s_NutrientDictionary.Item(keyString)
    Else
        Set GetNutrientByID = Nothing ' Return Nothing if ID not found
        Debug.Print "Warning: Nutrient ID " & nutrientID & " not found in repository."
    End If
End Function

'=============================================================================
' GetAllNutrients
'-----------------------------------------------------------------------------
' Purpose: Returns a collection of all Nutrient objects in the repository.
'          Useful for populating lists or combo boxes.
' Returns:
'   Collection: A collection containing all Nutrient objects. Returns Nothing
'               if the repository is not initialized.
'=============================================================================
Public Function GetAllNutrients() As Collection
    ' Check if the repository has been initialized
    If s_NutrientDictionary Is Nothing Then
        Debug.Print "Error: Nutrient repository not initialized. Call InitializeNutrientRepository first."
        Set GetAllNutrients = Nothing
        Exit Function
    End If

    ' Create a new collection to return (safer than returning the internal dictionary)
    Dim nutrientsCollection As Collection
    Set nutrientsCollection = New Collection
    
    Dim key As Variant
    Dim nutrient As nutrient
    
    ' Loop through the dictionary and add each Nutrient object to the collection
    For Each key In s_NutrientDictionary.Keys
        Set nutrient = s_NutrientDictionary.Item(key)
        nutrientsCollection.Add nutrient ' Add the object itself
    Next key
    
    Set GetAllNutrients = nutrientsCollection
    
    ' Clean up
    Set nutrient = Nothing
    Set nutrientsCollection = Nothing ' The caller now owns the returned collection reference
    
End Function

' --- Add other functions as needed, e.g., GetNutrientByName ---


