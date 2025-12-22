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

  ; Si une ancienne installation "CurrentUser" existe, conserver son emplacement.
  StrCmp $perUserInstallationFolder "" 0 customInitDone_${_cid}

  ; Sinon, forcer le dossier par défaut dans Documents (pas AppData\Local\Programs).
  ; On utilise APP_FILENAME pour rester cohérent avec la logique NSIS (et éviter les doubles sous-dossiers).
  StrCpy $INSTDIR "$DOCUMENTS\\${APP_FILENAME}"

customInitDone_${_cid}:
  !undef _cid
!macroend
