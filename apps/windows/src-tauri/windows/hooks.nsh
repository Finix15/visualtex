; VisualTeX Windows installer prerequisite check and per-user native Office choice.
; The production path installs the per-user Ribbon COM add-ins and ATL OLE
; LocalServer. Legacy Office.js Trusted Catalog resources are not installed.

!define VISUALTEX_INSTALLER_VERSION "1.2.7"

Var VisualTeXOfficeChoice
Var VisualTeXOfficeOnlyRadio
Var VisualTeXOfficeNativeRadio
Var VisualTeXOcrChoice
Var VisualTeXOcrCheckbox
Var VisualTeXOcrResourcePrefix
Var VisualTeXAcceptanceMode

; The generated Tauri PageReinstall function is patched after bundling so the
; same-version maintenance page defaults to "Uninstall VisualTeX" directly at
; control creation time. Do not use a GUI timer here: it races the generated
; page and does not reliably change the checked radio button.

Function VisualTeXRepairMainUninstallRegistration
  ; Older 1.2.3 builds could leave the remembered install directory while the
  ; standard Add/Remove Programs key was missing. Reconstruct the key before
  ; Tauri's maintenance page tries to launch uninstall.exe.
  ReadRegStr $0 HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\VisualTeX" "UninstallString"
  ${If} $0 != ""
    Return
  ${EndIf}

  ReadRegStr $1 HKCU "Software\visualtex\VisualTeX" ""
  ${If} $1 == ""
    IfFileExists "$LOCALAPPDATA\VisualTeX\uninstall.exe" 0 +2
      StrCpy $1 "$LOCALAPPDATA\VisualTeX"
  ${EndIf}
  ${If} $1 == ""
    IfFileExists "$PROFILE\AppData\VisualTeX\uninstall.exe" 0 +2
      StrCpy $1 "$PROFILE\AppData\VisualTeX"
  ${EndIf}
  ${If} $1 == ""
    Return
  ${EndIf}
  IfFileExists "$1\uninstall.exe" 0 visualtex_repair_uninstall_done

  WriteRegStr HKCU "Software\visualtex\VisualTeX" "" "$1"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\VisualTeX" "DisplayName" "VisualTeX"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\VisualTeX" "DisplayIcon" '$\"$1\visualtex.exe$\"'
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\VisualTeX" "DisplayVersion" "${VISUALTEX_INSTALLER_VERSION}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\VisualTeX" "Publisher" "visualtex"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\VisualTeX" "InstallLocation" '$\"$1$\"'
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\VisualTeX" "UninstallString" '$\"$1\uninstall.exe$\"'
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\VisualTeX" "MainBinaryName" "visualtex.exe"
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\VisualTeX" "NoModify" 1
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\VisualTeX" "NoRepair" 1

visualtex_repair_uninstall_done:
FunctionEnd

Function VisualTeXOfficePageCreate
  Call VisualTeXRepairMainUninstallRegistration
  nsDialogs::Create 1018
  Pop $0
  ${If} $0 == error
    Abort
  ${EndIf}

  !insertmacro MUI_HEADER_TEXT "Tích hợp Office" "Chọn có cài tích hợp Word / PowerPoint native của VisualTeX hay không"
  ${NSD_CreateLabel} 0 0 100% 24u "Chọn chế độ tích hợp Office native trên Windows / Choose Windows native Office integration"
  Pop $0

  ${NSD_CreateRadioButton} 0 34u 100% 16u "Chỉ VisualTeX (không cài phần bổ trợ Office) / VisualTeX only"
  Pop $VisualTeXOfficeOnlyRadio

  ${NSD_CreateRadioButton} 0 58u 100% 16u "VisualTeX + tích hợp Office native (khuyên dùng)"
  Pop $VisualTeXOfficeNativeRadio
  ${NSD_Check} $VisualTeXOfficeNativeRadio

  ${NSD_CreateLabel} 0 88u 100% 44u "Chế độ native dùng phần bổ trợ Ribbon COM cho Word/PowerPoint và ATL OLE LocalServer. Trình cài đặt không khởi chạy Office và sẽ dọn phần Office.js Trusted Catalog cũ."
  Pop $0

  nsDialogs::Show
FunctionEnd

Function VisualTeXOfficePageLeave
  ${NSD_GetState} $VisualTeXOfficeNativeRadio $0
  ${If} $0 == ${BST_CHECKED}
    nsExec::ExecToStack `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -Command "if (Get-Process WINWORD,POWERPNT,EXCEL,OUTLOOK,ONENOTE,MSACCESS,MSPUB,VISIO,MSPROJECT -ErrorAction SilentlyContinue) { exit 1 }; exit 0"`
    Pop $1
    Pop $2
    ${If} $1 != "0"
      MessageBox MB_ICONEXCLAMATION|MB_YESNO "Microsoft Office vẫn đang chạy. Buộc đóng sẽ kết thúc ngay Word, PowerPoint, Excel, Outlook, OneNote, Access, Publisher, Visio và Project; tài liệu Office chưa lưu có thể bị mất.$\r$\n$\r$\nBuộc đóng tất cả tiến trình Office và tiếp tục cài đặt? Chọn $\"Không$\" để quay lại trang trước, tự lưu và đóng Office.$\r$\n$\r$\nMicrosoft Office is still running. Force closing will terminate all common Office apps immediately and may discard unsaved work.$\r$\n$\r$\nForce close all Office processes and continue? Choose No to go back and close Office yourself." IDYES visualtex_force_close_office IDNO visualtex_office_close_declined

visualtex_force_close_office:
      nsExec::ExecToStack `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -Command "Get-Process WINWORD,POWERPNT,EXCEL,OUTLOOK,ONENOTE,MSACCESS,MSPUB,VISIO,MSPROJECT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue; Start-Sleep -Milliseconds 800; if (Get-Process WINWORD,POWERPNT,EXCEL,OUTLOOK,ONENOTE,MSACCESS,MSPUB,VISIO,MSPROJECT -ErrorAction SilentlyContinue) { exit 1 }; exit 0"`
      Pop $1
      Pop $2
      ${If} $1 != "0"
        MessageBox MB_ICONSTOP "Không thể đóng hết các tiến trình Office. Hãy lưu công việc, đóng các tiến trình Office còn lại trong Task Manager rồi thử lại.$\r$\n$\r$\nThe installer could not close every Office process. Save your work, close the remaining Office processes in Task Manager, and try again."
        Abort
      ${EndIf}
      Goto visualtex_office_process_check_done

visualtex_office_close_declined:
      Abort

visualtex_office_process_check_done:
    ${EndIf}
    StrCpy $VisualTeXOfficeChoice "native"
    Return
  ${EndIf}
  StrCpy $VisualTeXOfficeChoice "none"
FunctionEnd

Function VisualTeXOcrPageCreate
  nsDialogs::Create 1018
  Pop $0
  ${If} $0 == error
    Abort
  ${EndIf}

  !insertmacro MUI_HEADER_TEXT "Tài nguyên OCR ngoại tuyến" "Chọn có cài tài nguyên OCR cục bộ cùng VisualTeX hay không"
  ${NSD_CreateLabel} 0 0 100% 28u "OCR là thành phần tùy chọn. Sau khi cài mặc định, bạn có thể cấu hình OCR trong ứng dụng; nếu chưa cần OCR, bỏ chọn bên dưới để tiết kiệm dung lượng."
  Pop $0

  ${NSD_CreateCheckbox} 0 38u 100% 18u "Cài tài nguyên OCR ngoại tuyến (khuyên dùng) / Install offline OCR resources"
  Pop $VisualTeXOcrCheckbox
  ${If} $VisualTeXOcrChoice != "none"
    ${NSD_Check} $VisualTeXOcrCheckbox
  ${EndIf}

  ${NSD_CreateLabel} 0 66u 100% 48u "Gồm Python 3.12.10 riêng của VisualTeX, các wheel ngoại tuyến, OCR worker và danh mục mô hình. Mô hình OCR S/M/L vẫn được tải riêng khi cần."
  Pop $0

  ${NSD_CreateLabel} 0 116u 100% 20u "Nếu bỏ chọn, các tệp OCR sẽ không được cài; bạn có thể chạy lại bộ cài và chọn mục này sau."
  Pop $0

  nsDialogs::Show
FunctionEnd

Function VisualTeXOcrPageLeave
  ${NSD_GetState} $VisualTeXOcrCheckbox $0
  ${If} $0 == ${BST_CHECKED}
    StrCpy $VisualTeXOcrChoice "install"
  ${Else}
    StrCpy $VisualTeXOcrChoice "none"
  ${EndIf}
FunctionEnd

; OCR uses the bundled private Python 3.12.10 x64 runtime and a fixed local
; wheelhouse. The installer must never probe or depend on system Python.
;
; Tauri normally writes every configured resource unconditionally. The custom
; NSIS template routes every resource directory/file through these macros so
; the user's OCR choice is applied before extraction. OCR resources remain
; embedded in the installer so the default checked mode stays fully offline,
; but choosing "none" never writes ocr*, wheel or private-Python payloads into
; the installation directory.
!macro VisualTeXCreateBundledResourceDirectory DESTINATION
  StrCpy $VisualTeXOcrResourcePrefix "${DESTINATION}" 3
  ${If} $VisualTeXOcrChoice == "none"
  ${AndIf} $VisualTeXOcrResourcePrefix == "ocr"
    DetailPrint "Skipping optional OCR resource directory: ${DESTINATION}"
  ${Else}
    CreateDirectory "$INSTDIR\\${DESTINATION}"
  ${EndIf}
!macroend

!macro VisualTeXInstallBundledResource DESTINATION SOURCE
  StrCpy $VisualTeXOcrResourcePrefix "${DESTINATION}" 3
  ${If} $VisualTeXOcrChoice == "none"
  ${AndIf} $VisualTeXOcrResourcePrefix == "ocr"
    DetailPrint "Skipping optional OCR resource: ${DESTINATION}"
  ${Else}
    File /a "/oname=${DESTINATION}" "${SOURCE}"
  ${EndIf}
!macroend

!macro NSIS_HOOK_PREINSTALL
  ; Normalize the two legacy 1.2.3 install locations back to Tauri's canonical
  ; current-user directory. Preserve genuinely custom directories.
  ${If} $INSTDIR == "$PROFILE\AppData\VisualTeX"
    StrCpy $INSTDIR "$LOCALAPPDATA\VisualTeX"
  ${ElseIf} $INSTDIR == "$APPDATA\VisualTeX"
    StrCpy $INSTDIR "$LOCALAPPDATA\VisualTeX"
  ${EndIf}

  ; Custom pages are skipped by NSIS /S. A release acceptance install may use
  ; /VISUALTEXOFFICE=skip to leave the machine's existing Office integration
  ; untouched while testing the exact installed desktop executable. Interactive
  ; installs retain the page choice; ordinary unattended installs default to
  ; the recommended native Office mode.
  ${GetParameters} $0
  ClearErrors
  ${GetOptions} $0 "/VISUALTEXOFFICE=" $1
  ${IfNot} ${Errors}
    ${If} $1 == "native"
      StrCpy $VisualTeXOfficeChoice "native"
    ${ElseIf} $1 == "none"
      StrCpy $VisualTeXOfficeChoice "none"
    ${ElseIf} $1 == "skip"
      StrCpy $VisualTeXOfficeChoice "skip"
      StrCpy $VisualTeXAcceptanceMode "1"
    ${Else}
      Abort "Unsupported /VISUALTEXOFFICE value: $1"
    ${EndIf}
  ${EndIf}
  ${If} $VisualTeXOfficeChoice == ""
    StrCpy $VisualTeXOfficeChoice "native"
  ${EndIf}

  ClearErrors
  ${GetOptions} $0 "/VISUALTEXOCR=" $1
  ${IfNot} ${Errors}
    ${If} $1 == "install"
      StrCpy $VisualTeXOcrChoice "install"
    ${ElseIf} $1 == "none"
      StrCpy $VisualTeXOcrChoice "none"
    ${Else}
      Abort "Unsupported /VISUALTEXOCR value: $1"
    ${EndIf}
  ${EndIf}
  ${If} $VisualTeXOcrChoice == ""
    StrCpy $VisualTeXOcrChoice "install"
  ${EndIf}

  ${If} $VisualTeXAcceptanceMode == "1"
    DetailPrint "Installed-release acceptance mode: preserving existing Office integration and skipping machine prerequisite prompts."
  ${EndIf}

  ${If} $VisualTeXOcrChoice == "install"
    DetailPrint "Installing VisualTeX OCR offline resources: private Python 3.12.10 x64 runtime, fixed wheelhouse, worker and model catalog."
  ${Else}
    DetailPrint "OCR offline resources were disabled by the user; no ocr*, wheel or private-Python resources will be written to the installation directory."
  ${EndIf}
!macroend

!macro NSIS_HOOK_POSTINSTALL
  DetailPrint "Applying the selected VisualTeX Office integration mode: $VisualTeXOfficeChoice"
  ${If} $VisualTeXOfficeChoice == "native"
    DetailPrint "Installing the machine-wide VisualTeX Ribbon add-ins and native Formula OLE LocalServer. A UAC prompt may appear."
    IfFileExists "$INSTDIR\${MAINBINARYNAME}.exe" 0 visualtex_main_binary_missing
    IfFileExists "$INSTDIR\scripts\ensure_windows_office_certificate.ps1" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\scripts\install_windows_vsto.ps1" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\scripts\install_windows_vsto_runtime.ps1" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\scripts\test_windows_office_runtime.ps1" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\windows-office\VisualTeX-WindowsOffice-VSTO-x64.msi" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\windows-office\VisualTeX-WindowsOffice-VSTO-x64.sha256.json" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\windows-office\VisualTeX-WindowsOffice-VSTO-x86.msi" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\windows-office\VisualTeX-WindowsOffice-VSTO-x86.sha256.json" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\windows-office\vstor_redist.exe" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\windows-office\vstor_redist.sha256.json" 0 visualtex_office_missing

    DetailPrint "Checking Microsoft Visual Studio Tools for Office Runtime..."
    nsExec::ExecToStack `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\install_windows_vsto_runtime.ps1" -RuntimeInstallerPath "$INSTDIR\windows-office\vstor_redist.exe" -ManifestPath "$INSTDIR\windows-office\vstor_redist.sha256.json" -CheckOnly`
    Pop $1
    Pop $2
    StrCmp $1 "0" visualtex_vsto_runtime_ready 0
    DetailPrint "Microsoft VSTO Runtime is missing. Detection output: $2"
    IfSilent visualtex_vsto_runtime_install 0
    MessageBox MB_ICONQUESTION|MB_YESNO "Máy này chưa có Microsoft Visual Studio Tools for Office Runtime, thành phần bắt buộc cho phần bổ trợ Ribbon native của VisualTeX.$\r$\n$\r$\nBộ cài đã kèm VSTO Runtime chính thức của Microsoft với chữ ký số hợp lệ. Cài ngay và tiếp tục cấu hình Office? Windows sẽ yêu cầu quyền quản trị qua UAC.$\r$\n$\r$\nChọn Không sẽ giữ ứng dụng VisualTeX nhưng bỏ qua phần bổ trợ Office." IDYES visualtex_vsto_runtime_install IDNO visualtex_vsto_runtime_declined

visualtex_vsto_runtime_declined:
    DetailPrint "The user declined Microsoft VSTO Runtime installation. VisualTeX Office integration was skipped."
    Goto visualtex_office_done

visualtex_vsto_runtime_install:
    DetailPrint "Installing the bundled Microsoft VSTO Runtime. A UAC elevation prompt may appear."
    nsExec::ExecToLog `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\install_windows_vsto_runtime.ps1" -RuntimeInstallerPath "$INSTDIR\windows-office\vstor_redist.exe" -ManifestPath "$INSTDIR\windows-office\vstor_redist.sha256.json"`
    Pop $1
    StrCmp $1 "0" visualtex_vsto_runtime_ready visualtex_vsto_runtime_failed

visualtex_vsto_runtime_failed:
    DetailPrint "Microsoft VSTO Runtime installation failed or the UAC prompt was cancelled. Native Office integration was skipped."
    IfSilent visualtex_office_done 0
    MessageBox MB_ICONEXCLAMATION "Cài Microsoft VSTO Runtime thất bại hoặc yêu cầu quyền quản trị đã bị hủy. VisualTeX đã được cài nhưng phần bổ trợ Word/PowerPoint sẽ bị bỏ qua.$\r$\n$\r$\nXem nhật ký vsto-runtime mới nhất trong %LOCALAPPDATA%\VisualTeX\office\install-logs."
    Goto visualtex_office_done

visualtex_vsto_runtime_ready:
    DetailPrint "Microsoft VSTO Runtime is installed and verified."
    DetailPrint "VisualTeX will trust one self-signed certificate for HTTPS communication restricted to localhost/127.0.0.1."

    nsExec::ExecToLog `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\ensure_windows_office_certificate.ps1" -VisualTeXPath "$INSTDIR\${MAINBINARYNAME}.exe"`
    Pop $0
    StrCmp $0 "0" 0 visualtex_office_failed

    nsExec::ExecToLog `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\install_windows_vsto.ps1" -PackageDirectory "$INSTDIR\windows-office" -VisualTeXPath "$INSTDIR\${MAINBINARYNAME}.exe"`
    Pop $0
    StrCmp $0 "0" visualtex_office_static_installed visualtex_office_failed

visualtex_office_static_installed:
    DetailPrint "Machine-wide Office files and registrations passed. Office bootstrap completed without leaving a resident VisualTeX process."
    WriteRegDWORD HKCU "Software\VisualTeX\OfficeIntegration" "RuntimeVerificationPending" 1
    DetailPrint "Companion and Word/PowerPoint connection verification are deferred until VisualTeX is launched from Finish or by the user."
    IfSilent visualtex_office_done 0
    MessageBox MB_ICONINFORMATION "Tệp tích hợp Office, đăng ký hệ thống, chứng chỉ, lớp COM và dịch vụ OLE đã được cài và kiểm tra tĩnh.$\r$\n$\r$\nChứng chỉ tự ký VisualTeX Local Office Companion chỉ dùng cho HTTPS cục bộ tại localhost/127.0.0.1 và được gỡ theo đúng dấu vân tay khi gỡ VisualTeX.$\r$\n$\r$\nBạn có thể kiểm tra kết nối Word và PowerPoint sau trong Cài đặt → Tích hợp Office."
    Goto visualtex_office_done

visualtex_office_failed:
    SetDetailsView show
    DetailPrint "VisualTeX main application installed, but the machine-wide Office files, registry entries, COM classes or OLE server failed static installation verification. See the newest vsto-bootstrap and vsto-diagnostic reports under %LOCALAPPDATA%\VisualTeX\office\install-logs."
    IfSilent visualtex_office_done 0
    MessageBox MB_ICONEXCLAMATION "VisualTeX đã được cài, nhưng tệp phần bổ trợ Office, đăng ký hệ thống, lớp COM hoặc dịch vụ OLE không qua kiểm tra tĩnh. Xem chi tiết cài đặt và báo cáo vsto-bootstrap, vsto-diagnostic mới nhất trong %LOCALAPPDATA%\VisualTeX\office\install-logs."
    Goto visualtex_office_done
  ${ElseIf} $VisualTeXOfficeChoice == "none"
    IfFileExists "$INSTDIR\scripts\uninstall_windows_vsto.ps1" 0 visualtex_office_done
    nsExec::ExecToLog `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\uninstall_windows_vsto.ps1"`
    Pop $0
    Goto visualtex_office_done
  ${Else}
    DetailPrint "Skipping Office integration changes for installed-release acceptance. Existing Office files, registrations, certificates and companion configuration are untouched."
    Goto visualtex_office_done
  ${EndIf}

visualtex_main_binary_missing:
  DetailPrint "The VisualTeX main executable is missing after installation. Windows Security or another antivirus may have quarantined it. Office integration was skipped."
  IfSilent visualtex_office_done 0
  MessageBox MB_ICONEXCLAMATION "Ứng dụng VisualTeX biến mất ngay sau khi cài; Windows Security hoặc phần mềm bảo mật khác có thể đã cách ly visualtex.exe. Phần bổ trợ Office đã bị bỏ qua.$\r$\n$\r$\nMở Windows Security → Virus & threat protection → Protection history và kiểm tra các mục như Behavior:Win32/Persistence.A!ml."
  Goto visualtex_office_done

visualtex_office_missing:
  DetailPrint "Windows native Office installation resources are missing. The VisualTeX main application was installed without Office integration."
  IfSilent visualtex_office_done 0
  MessageBox MB_ICONEXCLAMATION "Thiếu tài nguyên cài đặt Office native cho Windows. Ứng dụng VisualTeX chính đã được cài bình thường."
  Goto visualtex_office_done

visualtex_office_done:
  ${If} $VisualTeXAcceptanceMode == "1"
    DetailPrint "Installed-release acceptance mode: legacy install roots are untouched."
    Goto visualtex_postinstall_cleanup_done
  ${EndIf}

  ; Remove only recognized legacy installation roots after the canonical
  ; installation has completed. User data lives under the application bundle
  ; directories, not these installer roots.
  ${If} $INSTDIR != "$PROFILE\AppData\VisualTeX"
    IfFileExists "$PROFILE\AppData\VisualTeX\uninstall.exe" visualtex_remove_direct_appdata 0
    IfFileExists "$PROFILE\AppData\VisualTeX\visualtex.exe" visualtex_remove_direct_appdata visualtex_direct_appdata_done
visualtex_remove_direct_appdata:
    RMDir /r "$PROFILE\AppData\VisualTeX"
visualtex_direct_appdata_done:
  ${EndIf}
  ${If} $INSTDIR != "$APPDATA\VisualTeX"
    IfFileExists "$APPDATA\VisualTeX\uninstall.exe" visualtex_remove_roaming_legacy 0
    IfFileExists "$APPDATA\VisualTeX\visualtex.exe" visualtex_remove_roaming_legacy visualtex_roaming_legacy_done
visualtex_remove_roaming_legacy:
    ; Remove only known legacy application payloads. Preserve
    ; %APPDATA%\VisualTeX\ocr-storage.json, OfficeSessions, logs, and any
    ; unknown user data so a later VisualTeX installation can reuse the
    ; independently stored OCR environment without reinstalling it.
    Delete "$APPDATA\VisualTeX\visualtex.exe"
    Delete "$APPDATA\VisualTeX\VisualTeX.exe"
    Delete "$APPDATA\VisualTeX\uninstall.exe"
    Delete "$APPDATA\VisualTeX\visualtex-windows-office-bridge.exe"
    RMDir /r "$APPDATA\VisualTeX\ocr"
    RMDir /r "$APPDATA\VisualTeX\ocr-models"
    RMDir /r "$APPDATA\VisualTeX\ocr-python"
    RMDir /r "$APPDATA\VisualTeX\office"
    RMDir /r "$APPDATA\VisualTeX\scripts"
    RMDir /r "$APPDATA\VisualTeX\windows-office"
    RMDir "$APPDATA\VisualTeX"
visualtex_roaming_legacy_done:
  ${EndIf}
visualtex_postinstall_cleanup_done:
!macroend

!macro NSIS_HOOK_PREUNINSTALL
  IfFileExists "$INSTDIR\scripts\uninstall_windows_vsto.ps1" 0 visualtex_preuninstall_done
  nsExec::ExecToStack `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\uninstall_windows_vsto.ps1"`
  Pop $0
  Pop $1
  ${If} $0 != "0"
    SetDetailsView show
    DetailPrint "VisualTeX Office integration uninstall failed. ExitCode=$0 Output=$1"
    MessageBox MB_ICONSTOP "Không thể gỡ tích hợp Office hoặc chứng chỉ HTTPS của VisualTeX. Ứng dụng chính chưa bị xóa. Hãy xem nguyên nhân cụ thể trong nhật ký rồi thử lại.$\r$\n$\r$\nNếu Word hoặc PowerPoint thực sự đang chạy, hãy lưu tài liệu và đóng chúng. Xem các nhật ký vsto-uninstall-bootstrap và certificate-remove mới nhất trong %LOCALAPPDATA%\VisualTeX\office\install-logs."
    SetErrorLevel 1
    Quit
  ${EndIf}
visualtex_preuninstall_done:
!macroend

!macro NSIS_HOOK_POSTUNINSTALL
  ; The generated maintenance flow runs uninstall.exe directly from $INSTDIR
  ; with _?=$INSTDIR. The running uninstaller therefore cannot delete itself.
  ; Remove everything that is currently deletable, then launch a detached
  ; cleanup process that waits for this uninstaller PID to exit before deleting
  ; the final uninstall.exe and empty installation root.
  ;
  ; Preserve %APPDATA%\VisualTeX because it stores OfficeSessions user data,
  ; and preserve %APPDATA%\com.visualtex.studio.
  DeleteRegKey HKCU "Software\visualtex\VisualTeX"
  RMDir /r "$PROFILE\AppData\VisualTeX"
  RMDir /r "$INSTDIR"

  System::Call 'kernel32::GetCurrentProcessId() i .r0'
  Exec `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -Command "Wait-Process -Id $0 -ErrorAction SilentlyContinue; Start-Sleep -Milliseconds 300; Remove-Item -LiteralPath '$INSTDIR' -Recurse -Force -ErrorAction SilentlyContinue"`
!macroend
