  PROGRAM

  PRAGMA('project(#pragma define(_SVDllMode_ => 0))')
  PRAGMA('project(#pragma define(_SVLinkMode_ => 1))')
  !- link manifest
  PRAGMA('project(#pragma link(Test.EXE.manifest))')

  INCLUDE('DocxToHtml.inc'), ONCE

  MAP
  END


sDocxFile                     STRING(256)
sHtmlFile                     STRING(256)

Window                        WINDOW('DocxToHtml Converter'),AT(,,397,286),GRAY,SYSTEM
                                PROMPT('Select Docx file:'),AT(10,16),USE(?PROMPT1)
                                ENTRY(@s255),AT(67,15,300),USE(sDocxFile)
                                BUTTON('...'),AT(372,14,14,14),USE(?bSelectFile)
                                PROMPT('Html file name:'),AT(10,40),USE(?PROMPT2)
                                ENTRY(@s255),AT(67,38,263),USE(sHtmlFile)
                                BUTTON('Convert'),AT(336,36,50),USE(?bConvert)
                                OLE,AT(10,74,376,206),USE(?OLE1),COMPATIBILITY(020H),CREATE('Shell.Explorer.2')
                                END
                                PROMPT('HTML contents in Internet Explorer:'),AT(10,60),USE(?PROMPT3)
                              END

converter                     TDocxToHtml

  CODE
  
  sDocxFile = 'NewCIT_TPCL Spec_1.40_20140811.docx'
  sHtmlFile = 'test.html'
  
  OPEN(Window)
  ACCEPT
    CASE EVENT()
    OF EVENT:OpenWindow
      ?OLE1{'Silent'} = 1 !-- block script error messages 
    END

    CASE ACCEPTED()
    OF ?bSelectFile
      FILEDIALOG(, sDocxFile, 'Docx files|*.docx', FILE:LongName + FILE:KeepDir)
      DISPLAY(?sDocxFile)
      
    OF ?bConvert
      IF sDocxFile AND sHtmlFile
        SETCURSOR(CURSOR:Wait)
        IF converter.SaveHtml(sDocxFile, sHtmlFile)
          ?OLE1{'Navigate("file://'& LONGPATH(sHtmlFile) &'")'}
        END
        SETCURSOR()
      ELSE
        MESSAGE('Enter both file names.', 'DocxToHtml Converter', ICON:Exclamation)
      END
    END
  END
  