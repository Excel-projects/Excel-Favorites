<img align="left" src="Images/ReadMe/App.png" width="64px" >

# Microsoft Excel Favorites Ribbon
This is an Excel 2010 VSTO Addin written in Visual Studio Community 2017 VB.Net. It gives the user a custom favorites ribbon.


<h1 align="left">
  <img src="Images/ReadMe/MicrosoftExcelFavoritesRibbon.png" alt="Ribbon" />
</h1>

## Table of Contents
- <a href="#dependencies">Dependencies</a>
- <a href="#glossary-of-terms">Glossary of Terms</a>
- <a href="#functionality">Functionality</a>
    - <a href="#worksheet">Worksheet</a>
    - <a href="#edit">Edit</a>
    - <a href="#print">Print</a>   
    - <a href="#program">Program</a>   
    - <a href="#options">Options</a>   
    - <a href="#about">About</a>
        - <a href="#how-to">How To...</a>  
        - <a href="#api-doc">API Doc...</a>  
        - <a href="#description">Add-in Name</a>
        - <a href="#install-date">Install Date</a>  
        - <a href="#copyright">Copyright</a>  

<a id="user-content-dependencies" class="anchor" href="#dependencies" aria-hidden="true"> </a>
## Dependencies
|Software                                   |Dependency                 |
|:------------------------------------------|:--------------------------|
|[Microsoft Visual Studio Community 2017](https://www.visualstudio.com/vs/whatsnew/)|Solution|
|[Microsoft Excel 2010](https://www.microsoft.com/en-au/software-download/office)|Project|

<a id="user-content-glossary-of-terms" class="anchor" href="#glossary-of-terms" aria-hidden="true"> </a>
## Glossary of Terms

| Term                      | Meaning                                                                                  |
|:--------------------------|:-----------------------------------------------------------------------------------------|
| COM |Component Object Model (COM) is a binary-interface standard for software components introduced by Microsoft in 1993. It is used to enable inter-process communication and dynamic object creation in a large range of programming languages. COM is the basis for several other Microsoft technologies and frameworks, including OLE, OLE Automation, ActiveX, COM+, DCOM, the Windows shell, DirectX, UMDF and Windows Runtime.  |
|VSTO |Visual Studio Tools for Office (VSTO) is a set of development tools available in the form of a Visual Studio add-in (project templates) and a runtime that allows Microsoft Office 2003 and later versions of Office applications to host the .NET Framework Common Language Runtime (CLR) to expose their functionality via .NET.|
|XML|Extensible Markup Language (XML) is a markup language that defines a set of rules for encoding documents in a format that is both human-readable and machine-readable.The design goals of XML emphasize simplicity, generality, and usability across the Internet. It is a textual data format with strong support via Unicode for different human languages. Although the design of XML focuses on documents, the language is widely used for the representation of arbitrary data structures such as those used in web services.|

<a id="user-content-functionality" class="anchor" href="#functionality" aria-hidden="true"> </a>
## Functionality
This Excel ribbon named “Favorites” is inserted after the “Home” tab when Excel opens.

<a id="user-content-worksheet" class="anchor" href="#worksheet" aria-hidden="true"> </a>
### Worksheet (Group)
#### Save (Button)

    Save (Ctrl + S)

#### Save As (Button)

    Save As (F12)

<a id="user-content-edit" class="anchor" href="#edit" aria-hidden="true"> </a>
### Edit (Group)
#### Undo (Button)

    Undo (Ctrl + Z)

#### Copy (Button)

    Copy (Ctrl + C)

#### Cut (Button)

    Cut (Ctrl + X)

#### Paste (Button)

    Paste (Ctrl + V)

#### Spelling (Button)

    Spelling (F7)

<a id="user-content-print" class="anchor" href="#print" aria-hidden="true"> </a>
### Print (Group)
#### Setup (Button)

    Show the Sheet tab of the page setup dialog box

#### Preview (Button)

    Preview (Ctrl + F2)

#### Print (Button)

    Print (Ctrl + P)

<a id="user-content-program" class="anchor" href="#program" aria-hidden="true"> </a>
### Program (Group)
#### New (Button)

    New file

#### Open (Button)

    Open (Ctrl + O)

#### Close (Button)

    Close file

#### Properties (Button)

    Open the properties of the file

#### Options (Button)

    Open the options dialog box

#### Exit (Button)

    Exit the application

<a id="user-content-options" class="anchor" href="#options" aria-hidden="true"> </a>
### Options (Group)
#### Settings (Button)

Types of Settings

Application Settings

    These settings can only be changed in the project and need to be redeployed
    They will appear disabled in the form

User Settings

    These settings can be changed by the end-user
    They will appear enabled in the form

#### COM Addins (Button)

    Manage the available COM Add-ins

<a id="user-content-about" class="anchor" href="#about" aria-hidden="true"> </a>
### About (Group)

<a id="user-content-how-to" class="anchor" href="#how-to" aria-hidden="true"> </a>
#### How To… (Button)
* How to use this Excel Addin

<a id="user-content-api-doc" class="anchor" href="#api-doc" aria-hidden="true"> </a>
#### API Doc.. (Button)
* View API documentation for this product

<a id="user-content-description" class="anchor" href="#description" aria-hidden="true"> </a>
#### Description (Label)
* The application name with the version

<a id="user-content-install-date" class="anchor" href="#install-date" aria-hidden="true"> </a>
#### Install Date (Label)
* The install date of the application

<a id="user-content-copyright" class="anchor" href="#copyright" aria-hidden="true"> </a>
#### Copyright (Label)
* The author’s name
