# Docx-To-Html Converter

Docx-To-Html Converter based on [this article](https://www.codeproject.com/Articles/1162184/Csharp-Docx-to-HTML-to-Docx).  

##Requirements
Microsoft .NET 4.5.  

**MS Word is not required!**

##How to run demo program
Copy all dll files from \bin folder into demo folder.

##How to build
You must have free [EasyCOM2INC](http://www.ingasoftplus.com/ProductDetail.php?ProductID=24) installed to use base COM classes.
YourApp.exe.manifest file must contain following child element inside parent `<dependency>` element:  
    `<dependentAssembly>`  
      `<assemblyIdentity `   
        `type="win32" `  
        `name="DocxToHtml" `  
        `version="1.0.0.0"`  
      `\>`  
    `</dependentAssembly>`  

Also, Graham Dawson kindly provided a way without using COM stuff:  
`
		!in EVENT:OpenWindow  
		myOLE = create(0,CREATE:ole)  
		myOLE{PROP:ReportException} = True  
		myOLE{prop:create} = 'mikeduglas.DocxToHtml.DocxToHtml'  
		!call SaveHtml()  
		myOLE{'SaveHtml(' & longpath(sDocxFile) & ',' & longpath(sHtmlFile) & ')'}  
`  
Pay attention this method requires a WINDOW to be opened.