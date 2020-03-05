# Salesforce-Risk

## Overview 

Security teams require a consistent and accurate measurement of the risk of each product or project (system) in their portfolio. This repository contains a fully functional implementation of a risk measurement and analysis framework implemented in Google Sheets. 

This tooling provides a data driven configurable application to capture, measure and manage risk. 

The level of risk of a system is calculated by measuring the inherent risk of the system and adjusting the risk level by the effectiveness of any external mitigating factors to arrive at a residual risk measurement. 

Inherent risk is defined as a measure of the risk of a system without taking any external controls into account. 

The solution is also able to track changing risk rating values over time as systems evolve.

Inherent risk score is based on the intersection of Likelihood and Impact, with the option for a specific option to pin the Likelihood or Impact score to a specific value. 

The application allows the definition of an arbitrary number of inherent risk factors for both Likelihood and Impact and an arbitrary number of options for each factor. Every option has a specific quantitative value to represent the level of risk on a scale of 1 (Very Low) to 5 (Very High).

The application also allows the defintion of an arbitrary number of residual risk factors and an arbitrary number of options for each factor. Every option has a specific quantitative value to represent the amount of risk that is mitigated. Option values can be positive, indicating a reduction of overall risk. Option values can also be negative, which indicates an increase in the overall risk.

For both inherent and residual risk calculations we are attempting to condense multiple dimensions of risk down to a single scalar value that approximates how “risky” a given system is.


## Implementation

This repository contains Google Script code that must be contained in a single Google Sheets Script project. Place the contents of each file into an appropriate file in the script project.

## Initial Setup

Once all of the scripts and HTML assets are in place in the Google Script project, run the **Initialize** function in the setup.gs module. This function will provision all of the required spreadsheet assets and result in an application ready for configuration.

### Configuration

1. Run the **Initialize** function in setup.gs
1. Add Likelihood factors and options. Each option within a factor requires an appropriately formatted row in the **Likelihood** sheet
    1. In the **Likelihood** sheet enter the text of the factor in Column A
    1. Enter the text of the option in Column B
    1. Enter the hover/help text value for the option in Column C
    1. Copy the concatenation formula from another rows' Column D to the Column D for the new option
    1. Enter the option risk value (1-5) in Column E
    1. Optional: to "pin" an option risk value copy the concatenation formula from Column D to Column F
    
    *Example:*
    
    | A | B | C | D | E |
    |---|---|---|---|---|
    | Who has access to the system? | Restricted Internal Only | Only a limited set of internal personnel have access to the system | Who has access to the system?~Restricted Internal Only | 1 |
    | Who has access to the system? | Internal only | Only internal personnel have access to the system | Who has access to the system?~Internal only | 2 |
    | Who has access to the system? | Partner users | Internal personnel and a limited set of external entities have access to the system (third party, contractors, submitters, etc.) | Who has access to the system?~Partner users | 3 |
    | Who has access to the system? | Authenticated customers | Identified entities that have an authorized use of the system | Who has access to the system?~Authenticated customers | 4 |
    | Who has access to the system? | Anonymous Internet users | Open to unauthenticated access | Who has access to the system?~Anonymous Internet users | 5 |
1. Add Impact factors and options. Each option within a factor requires an appropriately formatted row in the **Impact** sheet
    1. In the **Impact** sheet enter the text of the factor in Column A
    1. Enter the text of the option in Column B
    1. Enter the hover/help text value for the option in Column C
    1. Copy the concatenation formula from another rows' Column D to the Column D for the new option
    1. Enter the option risk value (1-5) in Column E
    1. Optional: to "pin" an option risk value copy the concatenation formula from Column D to Column F
    
    *Example:*
    
    | A | B | C | D | E |
    |---|---|---|---|---|
    | What types of information could be disclosed? | Public data | Data that is generally available to non-authenticated users | What types of information could be disclosed?~Public data | 1 |
    | What types of information could be disclosed? | User/partner data or metadata | Data used by authenticated non-adminstrative users to perform business requirements | What types of information could be disclosed?~User/partner data or metadata | 3 |
    | What types of information could be disclosed? | Administrative data or metadata | Data used by authenticated adminstrative users to perform administrative tasks | What types of information could be disclosed?~Administrative data or metadata | 4 |
    | What types of information could be disclosed? | Authentication secrets | Sensitive and compartmantilized data used for resticted purposes | What types of information could be disclosed?~Authentication secrets | 5 |
    | What types of information could be disclosed? | Compliance Data | Sensitive data with restricted access used to perform business requirements, such as HIPAA, PCI, etc. | What types of information could be disclosed?~Compliance Data | 5 |
1. Every Likelihood and Impact factor must have a matching column in the data storage sheet.
    1. For each Likelihood and Impact factor copy the text value of the factor to an open cell in the first row of the **Raw Inherent Data** sheet
    1. To set the hover/help text for the factor enter the help text as a note for the factor text cell

    *Example*

    | A | B | C | D | E | F |
    |---|---|---|---|---|---|
    | RowId | Release | Timestamp | System Identifier | Who has access to the system? | What types of information could be disclosed? |
  
### Sample Data

The **setup.gs** script also can populate a set of example data. To populate sample data in the spreadsheet run the *populateSampleData* function in the **setup.gs** script.

## Usage

### New Systems

To track the risk of a new system select *Risk Management* -> *New Inherent Risk System* and answer Yes to the "Create new risk entry?" prompt.

Fill out the form, including the system name, version and all Likelihood and Impact values. Click the **Save** button to create a new system entity.

### Capturing New Versions

To update the risk of an existing system that has released a new version, which retains the history of all previous versions, place the cursor on the row for that system and press the **New Release** buttton. Answer Yes to the prompt, then enter a version number that is higher than the previous most recent version. The application uses native Javascript alphanumeric sorting, so ensure the new version number value will sort higher than the previous version number. 

Set the appropriate Likelihood and Impact factor options and click the **Save** button to create a new release for the selected system entity.

### Modifying Existing Values

To edit the options of an existing system entity click the **Load** button and answer Yes to the prompt.

Update the appropriate values and click the **Save** button to overwrite the values for the selected release of the selected system.

### Residual Risk

Once a system has at least one inherent risk rating the residual risk can be calculated. To onboard a new system for residual risk calculation select  the *Risk Management* -> *New Residual Risk System* menu option, select the inherent risk souce, click the **Create** button, fill out the form including version and all factors. Click the **Save** button to record the data and calculate the residual risk of the sytsem.

### Deleting Systems or Versions

All inherent risk data for systems is stored in the **Raw Inherent Data** sheet. Each row stores the values for a given system/version combination. Raw data can be modified or rows can be deleted if necessary in this sheet.

### Mass Recalculation

Occasionally it may be necessary to force a refresh and recalculation of all systems. To refresh all inherent risk calculation data to the most recent version for each system select the *Risk Management* -> *Recalculate Inherent Risk* menu option and answer Yes to the prompt.

## Implementation Details

## Data Storage
The inherent risk data is stored in separate tabs, referred to as the raw data repositories. The raw data repositories contain the factor text and option answers entered by the user. No processing of data is performed in the raw data repositories. There are separate raw data tabs for recording inherent and residual risk factors and options.

## Metadata

## Calculations

### Inherent Risk Implementation Details 

The **Likelihood** and **Impact** sheets contain the text of each risk factor and option as well as the numeric value of each option. The data model supports the use of value concatenation and lookups functions that drive the calculations. 

Each row in the **Likelihood** and **Impact** sheets is uniquely identified by the concatenation of the question (factor) text, the tilde character (~) and the option text. This is the key value that the calculations use to perform lookups to determine the score of every factor/option value for the systems.

For each system, the *Likelihood* value is calculated by taking the rounded average of all of the *Likelihood* factor option values. The *Likelihood* value is then converted to a qualitative representation by the formula: (0 - 1] - Very Low, (1 - 2] - Low, (2 - 3] - Moderate, (3 - 4] - High, (4 - 5] - Very High and is displayed as the *Likelihood Rating*.

The *Impact* value is calculated by taking the average of all of the *Impact* factor option values. The *Impact* value is then converted to a qualitative representation by the formula: (0 - 1] - Very Low, (1 - 2] - Low, (2 - 3] - Moderate, (3 - 4] - High, (4 - 5] - Very High and is displayed as the *Impact Rating*.

Both *Likelihood* and *Impact* factors can also have one or more “pinned” options implemented. If a factor has a pinned option then either the *Likelihood* or *Impact* value will be the greater of the pinned option value or the average of the category used as the value of that category, regardless of the values of all other options in that category.

The overall severity level of the system is calculated using the matrix in the **RatingLookup** sheet, by cross referencing the *Likelihood Rating* and *Impact Rating* values.

### Residual Risk Implementation Details

The residual risk functionality is primarily implemented on the Residual Risk sheet, some script and form code to onboard a new system. The residual risk calculations also depend on data and named ranges.

The numeric value of each control is generally determined by multiplying the effectiveness of each mitigating control by the weightage value for that control, dividing by the configured Denominator value and multiplying by 5.

The numeric values for each control are then summed and the total amount is subtracted from the inherent risk score.

The residual risk level is then converted to a qualitative scale via a lookup of the residual risk score.

