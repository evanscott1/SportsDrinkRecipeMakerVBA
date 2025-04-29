# Excel Recipe Generator VBA Project

**Author:** Evan Scott
**Date:** April 28, 2025 
**Version:** 1.0.0 (Product Persistence & Recipe Generation Backend)

## Overview

This Excel VBA project aims to create a tool for generating custom drink recipes (e.g., energy, endurance, electrolyte, protein drinks) based on user-defined nutritional goals and available product ingredients.

Version 1.0.0 establishes the core data structures, persistence mechanisms for products/nutrients, the backend logic for recipe generation using a greedy algorithm, associated unit/integration tests, and the basic UserForm interfaces for managing products and creating recipes.

## Setup & Initialization

### Standard User

For a standard user opening a pre-configured `.xlsm` file:
1.  Enable macros when prompted.
2.  The `Workbook_Open` event automatically initializes the Nutrient Repository.
3.  Use the controls on the "Dashboard" sheet to manage products and create recipes via the UserForms.

### Developer Setup (from code files)

If setting up the project from individual code files (`.bas`, `.cls`, `.frm`, `.frx`, `.doccls`) into a blank workbook:

1.  **Import Files:** Import all `.bas`, `.cls`, and `.frm` files into the VBE. Ensure the corresponding `.frx` files are in the same directory during `.frm` import.
2.  **Add References:** Go to `Tools > References...` and ensure "Microsoft Scripting Runtime" and "Microsoft Forms 2.0 Object Library" are checked. If developing/running tests, also check "Rubberduck".
3.  **Enable VBE Access (Required for Setup):** Go to `File > Options > Trust Center > Trust Center Settings... > Macro Settings` and check **"Trust access to the VBA project object model"**. Required *only* for the setup routine to assign macros to dashboard buttons. Can be disabled after setup.
4.  **Run Setup Routine:** In the VBE Immediate Window (`Ctrl+G`), type `ModSetup.SetupWorkbookEnvironment` and press Enter. This creates sheets, adds dashboard controls+macros, and initializes the nutrient repository. Or enter `ModSetup` and place cursure in the `SetupWorkbookEnironment` and run (`F5`).
5.  **Add Workbook_Open Code:** Double-click `ThisWorkbook` in the Project Explorer. Open the exported `ThisWorkbook.doccls` file in a text editor. Copy its contents and paste them into the `ThisWorkbook` code window in the VBE.
6.  **Save:** Save the workbook as Macro-Enabled (`.xlsm` or `.xlsb`).

## Project Structure

The project is organized into several VBA components:

### Core Classes

* **`Product.cls`**: Represents a purchasable product/ingredient. Holds details like ID, Name, Price, Total Mass, Servings, and a collection of its nutrient composition.
* **`NutrientQuantity.cls`**: Represents the amount of a *specific* nutrient within a product. Stores the `NutrientID` and the `MassPerServing` (in kg).
* **`Nutrient.cls`**: Represents a specific nutrient. Holds an `ID`, `Name`, and `Description`.
* **`RecipeIngredient.cls`**: Represents a single ingredient line item in a generated recipe, linking a `Product` to the required amount (in servings and kg) and calculated cost.
* **`Recipe.cls`**: Represents a complete single-serving recipe, holding a collection of `RecipeIngredient` objects.

### Standard Modules

* **`ModProductData.bas`**: Handles saving, loading, and deleting product data from the "ProductData" worksheet.
* **`ModNutrientRepository.bas`**: Manages the predefined list of known `Nutrient` objects in memory. Includes initialization and retrieval functions.
* **`ModRecipeGenerator.bas`**: Contains the core `GenerateRecipe` function which creates a `Recipe` object based on targets and available products using a greedy algorithm. Also contains helpers like `DisplayRecipeOnSheet`.
* **`ModConstants.bas`**: Contains shared public constants (`PRODUCT_DATA_SHEET_NAME`, etc.) and enumerations (`ProductDataColumns`).
* **`ModSetup.bas`**: Contains the `SetupWorkbookEnvironment` subroutine for initial workbook configuration (sheets, dashboard controls) and placeholder subs for button actions.

### UserForm Modules

* **`frmAddProduct.frm/.frx`**: User interface for adding new products and their nutrient quantities.
* **`frmUpdateProduct.frm/.frx`**: User interface for loading, editing, and saving existing products.
* **`frmCreateRecipe.frm/.frx`**: User interface for defining nutrient targets, selecting product exclusions, and triggering recipe generation.

### Testing Modules (Requires Rubberduck VBA)

* **`TestProductData.bas`**: Integration tests for `ModProductData`.
* **`TestNutrient.bas`**: Unit tests for the `Nutrient` class.
* **`TestNutrientQuantity.bas`**: Unit tests for the `NutrientQuantity` class.
* **`TestRecipeIngredient.bas`**: Unit tests for the `RecipeIngredient` class.
* **`TestRecipe.bas`**: Unit tests for the `Recipe` class.
* **`TestRecipeGenerator.bas`**: Integration tests for the `GenerateRecipe` function.

### Document Class Modules

* **`ThisWorkbook.doccls`**: Contains the `Workbook_Open` event handler to initialize the `ModNutrientRepository`.

## Data Storage

* **Product Data:** Stored on the `"ProductData"` worksheet. Structure defined by `ProductDataColumns` Enum in `ModConstants`.
* **Nutrient Repository:** Stored in memory (`Scripting.Dictionary`) within `ModNutrientRepository` after initialization. Data is hardcoded in `InitializeNutrientRepository`.
* **Generated Recipes:** Displayed on the `"RecipeOutput"` worksheet when generated via `frmCreateRecipe`. Not persistently stored by default.

## Dependencies

* Microsoft Excel (with VBA enabled)
* **Microsoft Scripting Runtime:** Required for `Scripting.Dictionary`. Check/add via `Tools > References...`.
* **Microsoft Forms 2.0 Object Library:** Required for UserForms and their controls. Usually referenced by default when UserForms are added.
* **Rubberduck VBA Add-In:** Required *only* for running the tests. Not required for core application functionality.

## Development Notes

### Early vs. Late Binding

Test modules and `ModNutrientRepository` use `#Const LateBind`. Set to `False` for development (with references added) and `True` for distribution.

### Running Tests

Requires Rubberduck VBA add-in and Test Explorer.

## Future Development / Refinements

* Complete implementation of UserForm logic (validation, data flow).
* Implement recipe cost calculation and display.
* Implement "Export Recipe" functionality.
* Refine the `GenerateRecipe` algorithm (e.g., handle overshooting targets, consider costs, allow nutrient ratios).
* Add more robust error handling throughout.
* Consider alternative data storage if worksheet performance becomes an issue (e.g., ADO with Access or SQLite).
* Improve UI/UX based on user feedback.

