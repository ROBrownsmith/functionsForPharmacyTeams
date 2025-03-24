# functionsForPharmacyTeams

This is an Excel(R) add-on written in VBA containing functions that may be helpful for pharmacy teams, especially those working in Primary Care working on reports derived from clinical systems.

Disclaimer. These functions should be used with professional discretion. They have not been extensively tested and may fail in unforseen edge cases. They should be useful for large data sets but calculations for individaul patients should be checked manually.

Please feel free to fork and improve the code.

## Add-on instructions.
[Download](https://github.com/ROBrownsmith/functionsForPharmacyTeams/raw/refs/heads/main/Functions%20for%20Pharmacy%20teams.xlam) the add-in file and place in a local folder.

If Developer tab is not available in the ribbon then:

File/options/tick developer under Main Tabs and OK.

then

Excel Add-ins/browse/select add-in file. 

Restart Excel.

You may need to unblock the add in if you get an "this file type is not supported in protected view" error: Follow the instructions from microsoft to unblock [here.](https://answers.microsoft.com/en-us/msoffice/forum/all/excel-error-this-file-type-is-not-supported-in/f8fc839e-a4fa-4625-8f1c-5fa9f68c98a0)

## There are some simple user-defined functions in the module, and some classes.

**cockroftAndGault**(age As Range, gender As Range, weightInKg As Range, sCrumolL As Range) As Double

derived https://bnf.nice.org.uk/medicines-guidance/prescribing-in-renal-impairment/
Calculates Creatinine Clearance in mL/min using the eponymous formula.


**eGFR2009CKDEPI**(age As Range, sex As Range, sCrumolL As Range) As Double

derived https://bnf.nice.org.uk/medicines-guidance/prescribing-in-renal-impairment/
Calculates eGFR using the eponymous formula.


**bodyMassIndex**(weightInKg As Range, heightInMetres As Range)

Throws an error if person less than, arbitrarily, a metre high


**idealWeightInKg**(gender As Range, heightInM As Range) As Double


**removeUnits**(rngValueWithUnits As Range)

Removes the units and returns a number. e.g. 75kg returns 75.


**AECScore**(formDescCells As Range) As Integer

https://kclpure.kcl.ac.uk/portal/files/94138490/Anticholinergic_effect_on_cognition_Publishedonline9June2016_GREEN_AAM.pdf
Calculates the Anticholinergic Effect on Cognition (AEC) score based on the formulation descriptions in the input cells. If more than one cell is selected the score is added. Beware double counting multiple strengths of the same preparation.

    
**addACBScore**(formDescCells1 As Range) As Integer

List from https://www.acbcalc.com/medicines

Augmented with brand synonyms from the [BNF](www.medicinescomplete.com).

note - no score for dosulepin, lofepramine, prochlorperazine, promazine, amiodarone, sertindole

note - different scores for levomepromazine and methotrimeprazine

Calculates the Anticholinergic Burden (ACB) score based on the formulation descriptions in the input cells. If more than one cell is selected the score is added. Beware double counting multiple strengths of the same preparation.

Sub **reindexHorizontalData**(indexColNum As Long, dataColNum As Long, dataRepeats As Long, worksheetName As String)


Sub **reindexHorizontalDataPrompted**()

Macro to call the reindexHorizontalData Sub.

The reindexHorizontalData Sub can be called from within other routines but this sub allows it to be called directly from the button in the custom FfPT ribbon, and prompts the user for the inputs for the active sheet.

indexColNum is how many columns are to be contained in the index.

dataColNum is how many columns are contained for each data item.

dataRepeats is how many times the data repeats.

Manipulating data this way can be necessary to create a lookup when the clinical system outputs horizontally. e.g. SystmOne's configured, or adhoc output.

example:

| NHS Number      | Name            | Drug1          | Dose1          | Drug2           | Dose2 |   |   |   |   |
|-----------------|-----------------|----------------|----------------|-----------------|-------|---|---|---|---|
| 1               | Foo Bar         | A              | B              | C               | D     |   |   |   |   |
| 2               | Bar Foo         | E              | F              | G               | H     |   |   |   |   |
|                 |                 |                |                |                 |       |   |   |   |   |
|                 |                 |                |                |                 |       |   |   |   |   |
|                 |                 |                |                |                 |       |   |   |   |   |
|                 |                 |                |                |                 |       |   |   |   |   |
| (indexColNum 1) | (indexColNum 2) | (dataColNum 1) | (dataColNum 2) | (datarepeats 2) |       |   |   |   |   |
|                 |                 |                |                |                 |       |   |   |   |   |

Becomes:

| NHS Number | Name    | Drug1 | Dose1 | Drug2 | Dose2 |   |   |   |   |
|------------|---------|-------|-------|-------|-------|---|---|---|---|
| 1          | Foo Bar | A     | B     |       |       |   |   |   |   |
| 2          | Bar Foo | E     | F     |       |       |   |   |   |   |
| 1          | Foo Bar | C     | D     |       |       |   |   |   |   |
| 2          | Bar Foo | G     | H     |       |       |   |   |   |   |
|            |         |       |       |       |       |   |   |   |   |
|            |         |       |       |       |       |   |   |   |   |
|            |         |       |       |       |       |   |   |   |   |
|            |         |       |       |       |       |   |   |   |   |


**FHIRDosage**(doseInstruction As Range)

Uses clsDosage to return a [FHIR STU3](https://hl7.org/fhir/STU3/dosage.html) compatible JSON. This is a standard identified by [NHS Digital](https://nhsconnect.github.io/Dose-Syntax-Implementation/index.html).


**scriptDurationInDays**(doseInstruction As Range, scriptQuantity As Range)

Uses clsDoseProcessor method to return how many days the prescription ought to last given the dose instructions and the quantity supplied. It could be anticipated this could be useful in the analysis of antibiotic prescribing.

Doses with a Sequence property value > 1 will not be calculated correctly.

Doses with a repeatPeriod in hours may not be calculated correctly.

**morphineEquivalentFrom**(doseInstruction As Range, formDescRng As Range)

Uses clsDoseProcessor method to return the daily morphine dose equivalent in milligrams given the formulation's description and the dose instructions.

Doses with a Sequence property value > 2 will not be calculated correctly.

**clsDosage**

Aims to convert a dose-instruction string to a FHIR STU3 compatible JSON

https://hl7.org/fhir/STU3/dosage.html

This interprets/parses the dose instruction string and returns the concepts as structured data. The Route, Site, and asNeededCodeableConcept properties are not fully implemented because they contain over 1000 possible snomed codes. It may need an API call in order to implement properly. Could possibly be useful to create data to train an AI. Duration.start and duration.end are not yet implemented either. The function can successfully convert just over 80% of the dose strings in the test document.

<ins> Class Method(s): </ins>

FHIR3JSONConvertedFrom(doseString As String) As String


**clsDoseProcessor**

A suite utilising clsDosage and clsParseJSON to    

-work out a prescription duration, given dose and quantity.
              
-work out the morphine dose equivalent per day, given the dose and product description of an opioid. Methadone is not supported and will result in an error.

<ins> Class Method(s): </ins>

scriptDurationIs(doseStr As String, scriptQuant As Double) As Double

morphineDoseEquivalentOf(doseStr As String, formDesc As String)


**clsParseJSON**

A class created from code by Daniel Ferry to parse JSON in VBA.

https://medium.com/swlh/excel-vba-parse-json-easily-c2213f4d8e7a

<ins> Class Method(s): </ins>

Public Function ParseJSON(json As String, Optional key As String = "obj") As Object

Public Sub ListPaths()

Public Function GetFilteredValues(match As String) As Variant

Public Function GetFilteredTable(cols As Variant) As Variant
