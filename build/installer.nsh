; Custom NSIS include for electron-builder
; Goal: default install location in the current user's Documents folder.

; Verrouiller l'installation en "Utilisateur courant" (CurrentUser)
; Ce hook est appelé par multiUserUi.nsh avant la page de choix.
!macro customInstallMode
  StrCpy $isForceCurrentInstall "1"
  StrCpy $isForceMachineInstall "0"
!macroend

!macro customInit
  !define _cid ${__LINE__}

  ; Only for per-user installs (requested: C:\Users\%username%\Documents)
  StrCmp $installMode "CurrentUser" 0 customInitDone_${_cid}

  ; If a previous per-user install exists, keep that location.
  ; (Variable comes from multiUser.nsh)
  StrCmp $perUserInstallationFolder "" 0 customInitDone_${_cid}

  ; Set default to Documents\<ProductName>
  StrCpy $INSTDIR "$DOCUMENTS\\${PRODUCT_NAME}"

customInitDone_${_cid}:
  !undef _cid
!macroend
