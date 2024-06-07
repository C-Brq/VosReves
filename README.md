
# 1 Project Overview
This project has for objective to allow users to add printers to a list with the consummables, the main objective behind the project is the use of the coding language VBA.

There is a few tests described in [# 3 Description of the Program](#3-description-of-the-program).

The code is available inside the Spreadsheet.

# 2 Making of the program
The program is using simple VBA trough EXCEL, the program is entirely made by hand and might need some adjustements, it's not perfect in any way but it works and is a proof of concept.

Before making the program we need to first design the User Form that will be completed:
Here is the first draw of the forms:

![Modèle de formulaire d'ajout d'imprimante](https://github.com/C-Brq/VosReves/assets/156824818/3b47a3f6-fc1f-4985-ac2b-e6933b76866f)

There is 2 main user forms, the first one includes consummables and their caracteristics such as Reference, Type of Printer, Brand, Price, Color and others...
Here is a quick view of the form:

<img width="218" alt="image" src="https://github.com/C-Brq/VosReves/assets/156824818/98e6e3d8-6ce7-4796-82db-dfe9db7b0d5e">

The second userform is purely about the Printer, it includes caracteristics such as the Brand,Reference , Price , Type of Printer and the references for the colors.
Here is a quick view of the form:

<img width="341" alt="image" src="https://github.com/C-Brq/VosReves/assets/156824818/97d13ccf-4fca-4c96-bcbf-e938332dbc5b">

# 3 Description of the Program
The Program is made to limit the user error to the maximum by preserving the errors from the inputs.

The program checks the type of printer and such parameters like the different references, The Spreadsheet has 3 pages:

The first spreadsheet is the printers

The second spreadsheet is the consummables and such.

And the third spreadsheet is the brands that are accepted and the ID linked to the brand and also includes the brand of the consummables, in addition the type of printer is error checked, by using data : a laser printer can't be linked with a ink consummable.

On the second spreadsheet there is an autoconcatenation for the references, that way the user doesn't have to type the brand of the consummable and just has to type the new reference to add, this also allows no errors linked to similar references while being different brands.
