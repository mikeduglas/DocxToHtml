!Generated .CLW file (by the Easy COM generator v 1.14)

  MEMBER
  INCLUDE('svcomdef.inc'),ONCE
  MAP
    MODULE('WinAPI')
      ecg_DispGetParam(LONG pdispparams,LONG dwPosition,VARTYPE vtTarg,*VARIANT pvarResult,*SIGNED uArgErr),HRESULT,RAW,PASCAL,NAME('DispGetParam')
      ecg_VariantInit(*VARIANT pvarg),HRESULT,PASCAL,PROC,NAME('VariantInit')
      ecg_VariantClear(*VARIANT pvarg),HRESULT,PASCAL,PROC,NAME('VariantClear')
      ecg_VariantCopy(*VARIANT vargDest,*VARIANT vargSrc),HRESULT,PASCAL,PROC,NAME('VariantCopy')
      memcpy(LONG lpDest,LONG lpSource,LONG nCount),LONG,PROC,NAME('_memcpy')
      GetErrorInfo(ULONG dwReserved,LONG pperrinfo),HRESULT,PASCAL,PROC
    END
    INCLUDE('svapifnc.inc')
    Dec2Hex(ULONG),STRING
    INCLUDE('getvartype.inc', 'DECLARATIONS')
  END
  INCLUDE('DocxToHtml.inc')

Dec2Hex                       PROCEDURE(ULONG pDec)
locHex                          STRING(30)
  CODE
  LOOP UNTIL(~pDec)
    locHex = SUB('0123456789ABCDEF',1+pDec % 16,1) & CLIP(locHex)
    pDec = INT(pDec / 16)
  END
  RETURN CLIP(locHex)

  INCLUDE('getvartype.inc', 'CODE')
!========================================================!
! IDocxToHtmlClass implementation                        !
!========================================================!
IDocxToHtmlClass.Construct    PROCEDURE()
  CODE
  SELF.debug=true

IDocxToHtmlClass.Destruct     PROCEDURE()
  CODE
  IF SELF.IsInitialized=true THEN SELF.Kill() END

IDocxToHtmlClass.Init         PROCEDURE()
loc:lpInterface                 LONG
  CODE
  SELF.HR=CoCreateInstance(ADDRESS(IID_DocxToHtml),0,SELF.dwClsContext,ADDRESS(IID_IDocxToHtml),loc:lpInterface)
  IF SELF.HR=S_OK
    RETURN SELF.Init(loc:lpInterface)
  ELSE
    SELF.IsInitialized=false
    SELF._ShowErrorMessage('IDocxToHtmlClass.Init: CoCreateInstance',SELF.HR)
  END
  RETURN SELF.HR

IDocxToHtmlClass.Init         PROCEDURE(LONG lpInterface)
  CODE
  IF PARENT.Init(lpInterface) = S_OK
    SELF.IDocxToHtmlObj &= (lpInterface)
  END
  RETURN SELF.HR

IDocxToHtmlClass.Kill         PROCEDURE()
  CODE
  IF PARENT.Kill() = S_OK
    SELF.IDocxToHtmlObj &= NULL
  END
  RETURN SELF.HR

IDocxToHtmlClass.GetInterfaceObject   PROCEDURE()
  CODE
  RETURN SELF.IDocxToHtmlObj

IDocxToHtmlClass.GetInterfaceAddr PROCEDURE()
  CODE
  RETURN ADDRESS(SELF.IDocxToHtmlObj)
  !RETURN INSTANCE(SELF.IDocxToHtmlObj, 0)

IDocxToHtmlClass.GetLibLocation   PROCEDURE()
  CODE
  RETURN GETREG(REG_CLASSES_ROOT,'CLSID\{{C942DA77-EAF6-43B0-812D-F6DAD755AC41}\InprocServer32')

IDocxToHtmlClass.GetHtml      PROCEDURE(BSTRING pdocxfile,*BSTRING ppRetVal)
HR                              HRESULT(S_OK)

  CODE
  IF SELF.IsInitialized=false THEN SELF.HR=E_FAIL;RETURN(SELF.HR) END
  HR=SELF.IDocxToHtmlObj.GetHtml(pdocxfile,ppRetVal)
  SELF.HR=HR
  IF HR < S_OK
    SELF._ShowErrorMessage('IDocxToHtmlClass.GetHtml',HR)
  END
  RETURN HR

IDocxToHtmlClass.SaveHtml     PROCEDURE(BSTRING pdocxfile,BSTRING htmlfile)
HR                              HRESULT(S_OK)

  CODE
  IF SELF.IsInitialized=false THEN SELF.HR=E_FAIL;RETURN(SELF.HR) END
  HR=SELF.IDocxToHtmlObj.SaveHtml(pdocxfile,htmlfile)
  SELF.HR=HR
  IF HR < S_OK
    SELF._ShowErrorMessage('IDocxToHtmlClass.SaveHtml',HR)
  END
  RETURN HR

!========================================================!
! TDocxToHtml implementation                             !
!========================================================!
TDocxToHtml.Construct          PROCEDURE()
  CODE
  SELF.m_COMIniter &= NEW CCOMIniter
  SELF.Init()

TDocxToHtml.Destruct           PROCEDURE()
  CODE
  PARENT.Destruct()
  DISPOSE(SELF.m_COMIniter)

TDocxToHtml.Init               PROCEDURE()
hr                              HRESULT, AUTO

  CODE
  hr = PARENT.Init()
  IF hr <> S_OK
!    cnv::DebugInfo('PARENT.Init error '& SELF.Dec2Hex(hr))
    RETURN hr
  END

  RETURN S_OK
  
TDocxToHtml.Kill               PROCEDURE()
  CODE
  RETURN PARENT.Kill()

TDocxToHtml.GetHtml           PROCEDURE(STRING docxfile)
bVal                            BSTRING
  CODE
  PARENT.GetHtml(LONGPATH(docxfile), bVal)
  RETURN bVal

TDocxToHtml.SaveHtml          PROCEDURE(STRING docxfile, STRING htmlfile)
  CODE
  RETURN CHOOSE(PARENT.SaveHtml(LONGPATH(docxfile), LONGPATH(htmlfile)) = S_OK)
  
