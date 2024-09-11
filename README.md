# VBA Syntax Highlight

Custom-built Syntax Highlight tool for Word Documents

# Demo

https://github.com/user-attachments/assets/2ad64c02-3493-4cb9-bcd5-372a5db566e9

# Install

Download this repository in your local machine and extract the files.

Enable the Developer tools in Microsoft Word and add the files from the src folder there. You can either import the codes one by one. Visual Basic will automatically assign them to the Forms, Modules or Class Modules folders inside your project.

If you want this to be available for all your documents, import to "Normal", otherwise, import to the filename you want to work with.

Alternatively, you can zip all contents of the build folder and change the extension to .docm (just don't zip the build folder itself).

# Usage

Select the text you want to highlight and apply the corresponding highlight feature. The template uploaded to this repository contain custom ribbons. If you wish to update your ribbons, follow the "Custom Ribbons" instructions

# Custom Ribbons

source: https://www.anirdesh.com/ribbon/manual.php

## Enabling a Custom Ribbon

1. Create a folder called 'Ribbon' anywhere on your hard disk.
2. Open Office Word. Do not modify the new blank document. Click on 'Save As' to save the document as customUI.docm file in the Ribbon folder. This is a macro-enabled file. Make sure you change the "Save As Type” to docm.
3. Unzip the Word file
 
```sh
unzip customUI.docm
```

4. In the _rels folder, open the .rels file with a text editor such as Notepad. (in MacOS, press `Command+Shift+Dot` to view hidden files), Replace the content of the file with the following code below. This only sets the relationship allowing you to modify the Ribbon interface as specified by the file customUI.xml. We will modify customUI.xml file in the next step.

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="customUIRelID" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="customUI/customUI.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
```

5. In the Ribbon folder, create a folder called customUI. In the customUI folder, create a text file called customUI.xml. The customUI.xml file contains the code to your Word Ribbon. The features you want in the Ribbon will determine the code of this file. The following code is provided so you can see the structure of the complete code. In the rest of the document, I will only show partial code.

```xml
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <ribbon>
    <tabs>
      <tab id="CustomRibbon" label="Custom Ribbon" insertBeforeMso="TabInsert">
        <group id="CustomGroup" label="CustomGroup">
          <button id="CustomButton" visible="true" size="large" label="Custom Button" screentip="This is a custom button" onAction="Macro1" imageMso="QueryBuilder"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

6. Zip everything back into a Word document

```sh
zip customUI.docm * -r
```

## Customizing the Ribbon

All ImageMSO list: https://bert-toolkit.com/imagemso-list.html
Control signatures: https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/aa722523(v=office.12)

Sample VBA code:

```vb
Sub Macro1(control As IRibbonControl)
	MsgBox "it works"
End Sub
```

## Adding a custom button image

In the unzipped .docm, add a folder named `_rels` and another named `images` in the CustomUI folder.

Inside the `_rels` folder, add a document named `CustomUI.xml.rels` with the following content

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships
	xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<Relationship Id="Action1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="images/image1.png"/>
	<Relationship Id="Action2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="images/image2.png"/>
	<Relationship Id="Action3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="images/image3.png"/>
</Relationships>
```

Inside the `images` folder, add the images that must be either 16x16 PNGs for small logos or 32x32 PNGs for large logos.

You can call these images in the buttons xml (customUI.xml) using the `image` attribute like so:

```xml
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
	<ribbon>
		<tabs>
			<tab id="CustomRibbon" label="Custom Ribbon" insertBeforeMso="TabInsert">
				<group id="CustomGroup" label="Custom Group 1">
					<button id="btnAction1" visible="true" size="large" label="Action 1" screentip="Execute action 1" onAction="Macro1" image="Action1" />
					<button id="btnAction2" visible="true" size="large" label="Action 2" screentip="Execute action 2" onAction="Macro2" image="Action2"/>
					<button id="btnAction3" visible="true" size="large" label="Action 3" screentip="Execute action 3" onAction="Macro3" image="Action3"/>
				</group>
			</tab>
		</tabs>
	</ribbon>
</customUI>		
```
