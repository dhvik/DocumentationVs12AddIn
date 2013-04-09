#Background
This addin is a conversion for the macro project I have been using for years <http://www.codeproject.com/Articles/9819/Automating-the-code-writing-process-using-macros>.
Since VisualStudio 2012 dropped the support for macros I needed to convert them in some kind of other packaging. 
This addin is a port of the VB macros into a C# visual studio addin.
Please feel free to reuse the code as you want. 

#Installation
To install this addin, Compile the project and copy the files from the output folder to any configuredaddin folder of your visual studio 2012.

To see which folders your VisualStudio 2012 monitors for addin:s, go to Tools->Options->Environment->Add-in security

If a default installation of VS 2012 is present you can copy all files to the following folder

> C:\Users\[user]\Documents\Visual Studio 2012\Addins

The following files must be copied:

* DocumentationVs12Addin.Addin
* DocumentationVs12Addin.dll
* log4net.dll
* log4net.config [optional, only to enable logging]

[Read more on msdn about addin registering](http://msdn.microsoft.com/en-us/library/vstudio/19dax6cz.aspx)

#Commands
More about these commands can be found at <http://www.codeproject.com/Articles/9819/Automating-the-code-writing-process-using-macros>

##DocumentThis (Text Editor::Ctrl+Alt+d)
Adds/updates documentation for the current member

##DocumentAndRegionizeThis (Text Editor::Ctrl+d)
Adds documentation (like DocumentThis) and surrounds/updates the member with a region directive

##EnterCodeRemark (Text Editor::Ctrl+Shift+Alt+c)
Adds a code remark for the current member

##KillLine ("Text Editor::Ctrl+Shift+k")
Kills the current line

##CreateDocumentationHeading ("Text Editor::Ctrl+Shift+Alt+h")
Creates a documentation heading

##PasteSeeParameterXmlDocTag ("Text Editor::Ctrl+Shift+x, Ctrl+Shift+d")
Pastes a see xml documentation tag in the current xml node (matching the correct type)

##EnterCodeComment ("Text Editor::Ctrl+Shift+c")
Enters a comment in code with the user/date prefilled

##FormatCurrentDocument ("Text Editor::Ctrl+Shift+d")
Formats the current document

##ActOnTab ("Text Editor::Ctrl+<")
Navigates forward in the xml documentation node

##ActOnShiftTab ("Text Editor::Ctrl+Shift+>")
Navigates backwards in the xml documentation node

##PasteTemplate ("Text Editor::Ctrl+Shift+j")
Inspects the current code and pastes a template if a match is found

##ExpandAllRegions ("Text Editor::Ctrl+Shift++")
Expands all Regions 

##CollapseAllRegions ("Text Editor::Ctrl+Shift+-")
Collapses all regions (warning, can be slow on large documents)

##ToggleParentRegion ("Text Editor::Ctrl+m,Ctrl+m")
Toggles the parent region

##SortLinesAscending ("Text Editor::Ctrl+Alt+s")
Sorts the marked lines A-Z

##SortLinesDescending ("Text Editor::Ctrl+Shift+Alt+s")
Sorts the marked lines Z-A