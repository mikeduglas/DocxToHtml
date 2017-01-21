!Generated .INC file (by the Easy COM generator v 1.14)

  INCLUDE('ecom2inc.def'),ONCE
  INCLUDE('DocxToHtml.int'),ONCE
  INCLUDE('ecombase.inc'),ONCE

IDocxToHtmlClass              CLASS(EasyCOMClass),TYPE,MODULE('DocxToHtml.clw'),LINK('DocxToHtml.clw')
IDocxToHtmlObj                  &IDocxToHtml,PROTECTED
Construct                       PROCEDURE()
Destruct                        PROCEDURE(),DERIVED
Init                            PROCEDURE(),HRESULT,PROC,DERIVED
Init                            PROCEDURE(LONG lpInterface),HRESULT,PROC,DERIVED
Kill                            PROCEDURE(),HRESULT,PROC,DERIVED
GetInterfaceObject              PROCEDURE(),*IDocxToHtml
GetInterfaceAddr                PROCEDURE(),LONG
GetLibLocation                  PROCEDURE(),STRING,DERIVED
!Interface methods
GetHtml                         PROCEDURE(BSTRING docxfile,*BSTRING pRetVal),HRESULT,PROC,VIRTUAL,PROTECTED
SaveHtml                        PROCEDURE(BSTRING docxfile,BSTRING htmlfile),HRESULT,PROC,VIRTUAL,PROTECTED
                              END

TDocxToHtml                   CLASS(IDocxToHtmlClass),TYPE,MODULE('DocxToHtml.clw'),LINK('DocxToHtml.clw')
m_COMIniter                     &CCOMIniter, PRIVATE

Construct                       PROCEDURE()
Destruct                        PROCEDURE(), DERIVED
Init                            PROCEDURE(), HRESULT, PROC, DERIVED
Kill                            PROCEDURE(), HRESULT, PROC, DERIVED

GetHtml                         PROCEDURE(STRING docxfile), STRING
SaveHtml                        PROCEDURE(STRING docxfile, STRING htmlfile), BOOL, PROC
                              END
