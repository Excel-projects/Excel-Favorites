<img align="left" src="Images/ReadMe/App.png" width="64px" >

# Microsoft Excel Favorites Ribbon <span class="Application_Version">3.0.0.0</span> 
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE "MIT License Copyright © 2017 Anthony Duguid")
![current_build Office_2013](https://img.shields.io/badge/current_build-Office_2013-red.svg)

This is an Excel Add-In written in Visual Studio Community 2017 VB.NET and [VBA](https://github.com/aduguid/MicrosoftExcelFavorites/raw/master/VBA/Favorites.xlam?raw=true "Download the VBA Add-In"). It gives the user a custom favorites ribbon.
<!---
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE "MIT License Copyright © 2017 Anthony Duguid")
[![star this repo](http://githubbadges.com/star.svg?user=aduguid&repo=MicrosoftExcelFavorites&style=flat&color=fff&background=007ec6)](http://github.com/aduguid/MicrosoftExcelFavorites)
[![fork this repo](http://githubbadges.com/fork.svg?user=aduguid&repo=MicrosoftExcelFavorites&style=flat&color=fff&background=007ec6)](http://github.com/aduguid/MicrosoftExcelFavorites/fork)
--->
<h1 align="left">
  <img src="Images/ReadMe/vsto.excel.favorites.png" alt="Ribbon" />
</h1>

## Table of Contents
- <a href="#dependencies">Dependencies</a>
- <a href="#glossary-of-terms">Glossary of Terms</a>
- <a href="#functionality">Functionality</a>
    - <a href="#worksheet">Worksheet</a>
         - <a href="#save">Save</a> 
         - <a href="#save-as">Save As</a> 
    - <a href="#edit">Edit</a>
         - <a href="#undo">Undo</a> 
         - <a href="#copy">Copy</a> 
         - <a href="#cut">Cut</a> 
         - <a href="#paste">Paste</a> 
         - <a href="#spelling">Spelling</a> 
    - <a href="#print-group">Print</a>   
         - <a href="#setup">Setup</a> 
         - <a href="#preview">Preview</a> 
         - <a href="#print">Print</a> 
    - <a href="#program">Program</a>  
         - <a href="#new">New</a> 
         - <a href="#open">Open</a> 
         - <a href="#close">Close</a> 
         - <a href="#properties">Properties</a> 
         - <a href="#options">Options</a> 
         - <a href="#exit">Exit</a> 
    - <a href="#calculator-group">Calculator</a>  
         - <a href="#calculator">Windows Calculator</a> 
         - <a href="#calculate-now">Calculate Now</a> 
    - <a href="#annotation-group">Annotation Tools</a>  
         - <a href="#camera">Excel Camera</a> 
         - <a href="#snip">Snipping Tool</a> 
         - <a href="#psr">Problem Step Recorder</a> 
    - <a href="#options-group">Options</a>  
         - <a href="#settings">Add-In Settings</a> 
    - <a href="#help">Help</a>
        - <a href="#how-to">How To...</a>  
        - <a href="#report-issue">Report Issue</a>  
    - <a href="#about">About</a>
        - <a href="#description">Add-in Name</a>
        - <a href="#install-date">Release Date</a>  
        - <a href="#copyright">Copyright</a>  

<a id="user-content-dependencies" class="anchor" href="#dependencies" aria-hidden="true"> </a>
## Dependencies
|Software                                   |Dependency                 |Project                    |
|:------------------------------------------|:--------------------------|:--------------------------|
|[Microsoft Visual Studio Community 2017](https://www.visualstudio.com/vs/whatsnew/)|Solution|VSTO|
|[Microsoft Office Developer Tools](https://blogs.msdn.microsoft.com/visualstudio/2015/11/23/latest-microsoft-office-developer-tools-for-visual-studio-2015/)|Solution|VSTO|
|[Microsoft Excel 2010 (or later)](https://www.microsoft.com/en-au/software-download/office)|Project|VBA, VSTO|
|[Visual Basic for Applications](https://msdn.microsoft.com/en-us/vba/vba-language-reference)|Code|VBA|
|[Extensible Markup Language (XML)](https://www.rondebruin.nl/win/s2/win001.htm)|Ribbon|VBA, VSTO|
|[Snagit](http://discover.techsmith.com/snagit-non-brand-desktop/?gclid=CNzQiOTO09UCFVoFKgod9EIB3g)|Read Me|VBA, VSTO|
|Badges ([Library](https://shields.io/), [Custom](https://rozaxe.github.io/factory/), [Star/Fork](http://githubbadges.com))|Read Me|VBA, VSTO|

<a id="user-content-glossary-of-terms" class="anchor" href="#glossary-of-terms" aria-hidden="true"> </a>
## Glossary of Terms

| Term                      | Meaning                                                                                  |
|:--------------------------|:-----------------------------------------------------------------------------------------|
| COM |Component Object Model (COM) is a binary-interface standard for software components introduced by Microsoft in 1993. It is used to enable inter-process communication and dynamic object creation in a large range of programming languages. COM is the basis for several other Microsoft technologies and frameworks, including OLE, OLE Automation, ActiveX, COM+, DCOM, the Windows shell, DirectX, UMDF and Windows Runtime.  |
| VBA |Visual Basic for Applications (VBA) is an implementation of Microsoft's event-driven programming language Visual Basic 6 and uses the Visual Basic Runtime Library. However, VBA code normally can only run within a host application, rather than as a standalone program. VBA can, however, control one application from another using OLE Automation. VBA can use, but not create, ActiveX/COM DLLs, and later versions add support for class modules.|
| VSTO |Visual Studio Tools for Office (VSTO) is a set of development tools available in the form of a Visual Studio add-in (project templates) and a runtime that allows Microsoft Office 2003 and later versions of Office applications to host the .NET Framework Common Language Runtime (CLR) to expose their functionality via .NET.|
| XML|Extensible Markup Language (XML) is a markup language that defines a set of rules for encoding documents in a format that is both human-readable and machine-readable.The design goals of XML emphasize simplicity, generality, and usability across the Internet. It is a textual data format with strong support via Unicode for different human languages. Although the design of XML focuses on documents, the language is widely used for the representation of arbitrary data structures such as those used in web services.|

<a id="user-content-functionality" class="anchor" href="#functionality" aria-hidden="true"> </a>
## Functionality
This Excel ribbon named “Favorites” is inserted after the “Home” tab when Excel opens.

<a id="user-content-worksheet" class="anchor" href="#worksheet" aria-hidden="true"> </a>
### Worksheet (Group)
<a id="user-content-save" class="anchor" href="#save" aria-hidden="true"> </a>
#### Save (Button)
* Save (Ctrl + S)

<a id="user-content-save-as" class="anchor" href="#save-as" aria-hidden="true"> </a>
#### Save As (Button)
* Save As (F12)

<a id="user-content-edit" class="anchor" href="#edit" aria-hidden="true"> </a>
### Edit (Group)
<a id="user-content-undo" class="anchor" href="#undo" aria-hidden="true"> </a>
#### Undo (Button)
* Undo (Ctrl + Z)

<a id="user-content-copy" class="anchor" href="#copy" aria-hidden="true"> </a>
#### Copy (Button)
* Copy (Ctrl + C)

<a id="user-content-cut" class="anchor" href="#cut" aria-hidden="true"> </a>
#### Cut (Button)
* Cut (Ctrl + X)

<a id="user-content-paste" class="anchor" href="#paste" aria-hidden="true"> </a>
#### Paste (Button)
* Paste (Ctrl + V)

<a id="user-content-spelling" class="anchor" href="#spelling" aria-hidden="true"> </a>
#### Spelling (Button)
* Spelling (F7)

<a id="user-content-print-group" class="anchor" href="#print-group" aria-hidden="true"> </a>
### Print (Group)
<a id="user-content-setup" class="anchor" href="#setup" aria-hidden="true"> </a>
#### Setup (Button)
* Show the Sheet tab of the page setup dialog box

<a id="user-content-preview" class="anchor" href="#preview" aria-hidden="true"> </a>
#### Preview (Button)
* Preview (Ctrl + F2)

<a id="user-content-print" class="anchor" href="#print" aria-hidden="true"> </a>
#### Print (Button)
* Print (Ctrl + P)

<a id="user-content-program" class="anchor" href="#program" aria-hidden="true"> </a>
### Program (Group)
<a id="user-content-new" class="anchor" href="#new" aria-hidden="true"> </a>
#### New (Button)
* New file

<a id="user-content-open" class="anchor" href="#open" aria-hidden="true"> </a>
#### Open (Button)
* Open (Ctrl + O)

<a id="user-content-close" class="anchor" href="#close" aria-hidden="true"> </a>
#### Close (Button)
* Close file

<a id="user-content-properties" class="anchor" href="#properties" aria-hidden="true"> </a>
#### Properties (Button)
* Open the properties of the file

<a id="user-content-options" class="anchor" href="#options" aria-hidden="true"> </a>
#### Options (Button)
* Open the options dialog box

<a id="user-content-exit" class="anchor" href="#exit" aria-hidden="true"> </a>
#### Exit (Button)
* Exit the application

<a id="user-content-options-group" class="anchor" href="#options-group" aria-hidden="true"> </a>
### Options (Group)
<a id="user-content-settings" class="anchor" href="#settings" aria-hidden="true"> </a>
#### Settings (Button)
- Types of Settings
  - Application Settings
    - These settings can only be changed in the project and need to be redeployed
    - They will appear disabled in the form
  - User Settings
    - These settings can be changed by the end-user
    - They will appear enabled in the form

<a id="user-content-help" class="anchor" href="#help" aria-hidden="true"> </a>
### Help (Group)

<a id="user-content-how-to" class="anchor" href="#how-to" aria-hidden="true"> </a>
#### How To… (Button)
* How to use this Excel Addin

<a id="user-content-api-doc" class="anchor" href="#report-issue" aria-hidden="true"> </a>
#### Report Issue (Button)
* Create a new issue on the project page

<a id="user-content-about" class="anchor" href="#about" aria-hidden="true"> </a>
### About (Group)

<a id="user-content-description" class="anchor" href="#description" aria-hidden="true"> </a>
#### Description (Label)
* The application name with the version

<a id="user-content-install-date" class="anchor" href="#install-date" aria-hidden="true"> </a>
#### Install Date (Label)
* The install date of the application

<a id="user-content-copyright" class="anchor" href="#copyright" aria-hidden="true"> </a>
#### Copyright (Label)
* The author’s name
