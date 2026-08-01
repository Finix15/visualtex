; VisualTeX Windows installer prerequisite check and per-user native Office choice.
; The production path installs the per-user Ribbon COM add-ins and ATL OLE
; LocalServer. Legacy Office.js Trusted Catalog resources are not installed.

!define VISUALTEX_INSTALLER_VERSION "1.2.3"

Var VisualTeXOfficeChoice
Var VisualTeXOfficeOnlyRadio
Var VisualTeXOfficeNativeRadio
Var VisualTeXMaintenanceDefaultApplied
Var VisualTeXMaintenanceTimerTicks

; On Tauri's same-version maintenance page, default to the second radio:
; "Uninstall VisualTeX". Upgrades use different button text and are untouched.
!define MUI_CUSTOMFUNCTION_GUIINIT VisualTeXOnGuiInit
Function VisualTeXOnGuiInit
  StrCpy $VisualTeXMaintenanceDefaultApplied "0"
  StrCpy $VisualTeXMaintenanceTimerTicks "0"
  ${NSD_CreateTimer} VisualTeXDefaultMaintenanceUninstall 50
FunctionEnd

Function VisualTeXDefaultMaintenanceUninstall
  ${If} $VisualTeXMaintenanceDefaultApplied == "1"
    ${NSD_KillTimer} VisualTeXDefaultMaintenanceUninstall
    Return
  ${EndIf}

  IntOp $VisualTeXMaintenanceTimerTicks $VisualTeXMaintenanceTimerTicks + 1
  ${If} $VisualTeXMaintenanceTimerTicks > 200
    ${NSD_KillTimer} VisualTeXDefaultMaintenanceUninstall
    Return
  ${EndIf}

  System::Call 'user32::IsWindow(p $R2)i.r0'
  ${If} $0 == 0
    Return
  ${EndIf}
  ${NSD_GetText} $R2 $0
  ${If} $0 != "$(addOrReinstall)"
    Return
  ${EndIf}

  SendMessage $R3 ${BM_CLICK} 0 0
  ${NSD_SetFocus} $R3
  StrCpy $VisualTeXMaintenanceDefaultApplied "1"
  ${NSD_KillTimer} VisualTeXDefaultMaintenanceUninstall
FunctionEnd

Page custom VisualTeXOfficePageCreate VisualTeXOfficePageLeave

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

  ${NSD_CreateLabel} 0 0 100% 30u "请选择是否启用 Windows 原生 Office 集成 / Choose Windows native Office integration"
  Pop $0

  ${NSD_CreateRadioButton} 0 38u 100% 16u "仅 VisualTeX（不安装 Office 插件） / VisualTeX only"
  Pop $VisualTeXOfficeOnlyRadio

  ${NSD_CreateRadioButton} 0 62u 100% 16u "VisualTeX + 原生 Office 集成（推荐）"
  Pop $VisualTeXOfficeNativeRadio
  ${NSD_Check} $VisualTeXOfficeNativeRadio

  ${NSD_CreateLabel} 0 92u 100% 42u "原生模式统一使用 Word/PowerPoint Ribbon COM 加载项与 ATL OLE LocalServer。安装过程不会启动 Office，并会清理旧 Office.js Trusted Catalog 残留。"
  Pop $0

  nsDialogs::Show
FunctionEnd

Function VisualTeXOfficePageLeave
  ${NSD_GetState} $VisualTeXOfficeNativeRadio $0
  ${If} $0 == ${BST_CHECKED}
    nsExec::ExecToStack `powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -Command "if (Get-Process WINWORD,POWERPNT,EXCEL,OUTLOOK,ONENOTE,MSACCESS,MSPUB,VISIO,MSPROJECT -ErrorAction SilentlyContinue) { exit 1 }; exit 0"`
    Pop $1
    Pop $2
    ${If} $1 != "0"
      MessageBox MB_ICONEXCLAMATION|MB_YESNO "检测到 Microsoft Office 仍在运行。强制关闭会立即结束 Word、PowerPoint、Excel、Outlook、OneNote、Access、Publisher、Visio 和 Project；未保存的 Office 文档可能丢失。$\r$\n$\r$\n是否强制关闭所有这些 Office 进程并继续安装？选择“否”将返回上一页，由您自行保存并关闭 Office。$\r$\n$\r$\nMicrosoft Office is still running. Force closing will terminate all common Office apps immediately and may discard unsaved work.$\r$\n$\r$\nForce close all Office processes and continue? Choose No to go back and close Office yourself." IDYES visualtex_force_close_office IDNO visualtex_office_close_declined

visualtex_force_close_office:
      nsExec::ExecToStack `powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -Command "Get-Process WINWORD,POWERPNT,EXCEL,OUTLOOK,ONENOTE,MSACCESS,MSPUB,VISIO,MSPROJECT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue; Start-Sleep -Milliseconds 800; if (Get-Process WINWORD,POWERPNT,EXCEL,OUTLOOK,ONENOTE,MSACCESS,MSPUB,VISIO,MSPROJECT -ErrorAction SilentlyContinue) { exit 1 }; exit 0"`
      Pop $1
      Pop $2
      ${If} $1 != "0"
        MessageBox MB_ICONSTOP "无法完全关闭所有 Office 进程。请保存工作并在任务管理器中关闭残留的 Office 进程后重试。$\r$\n$\r$\nThe installer could not close every Office process. Save your work, close the remaining Office processes in Task Manager, and try again."
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

; The editor itself works without Python, so an incompatible environment warns
; the user instead of silently failing later or blocking installation outright.

!macro VisualTeXProbeLauncher SELECTOR
  nsExec::ExecToStack `"py.exe" ${SELECTOR} -c "import platform,sys;sys.exit(0 if (3,9) <= sys.version_info[:2] <= (3,13) and platform.machine().lower() in ('amd64','x86_64','x64') else 1)"`
  Pop $0
  Pop $1
  StrCmp $0 "0" visualtex_python_ok
!macroend

!macro VisualTeXProbeCommand PROGRAM
  nsExec::ExecToStack `"${PROGRAM}" -c "import platform,sys;sys.exit(0 if (3,9) <= sys.version_info[:2] <= (3,13) and platform.machine().lower() in ('amd64','x86_64','x64') else 1)"`
  Pop $0
  Pop $1
  StrCmp $0 "0" visualtex_python_ok
!macroend

!macro NSIS_HOOK_PREINSTALL
  ; Normalize the two legacy 1.2.3 install locations back to Tauri's canonical
  ; current-user directory. Preserve genuinely custom directories.
  ${If} $INSTDIR == "$PROFILE\AppData\VisualTeX"
    StrCpy $INSTDIR "$LOCALAPPDATA\VisualTeX"
  ${ElseIf} $INSTDIR == "$APPDATA\VisualTeX"
    StrCpy $INSTDIR "$LOCALAPPDATA\VisualTeX"
  ${EndIf}

  ; Custom pages are skipped by NSIS /S. Preserve an explicit interactive
  ; choice, but default unattended installs to the recommended Office mode.
  ${If} $VisualTeXOfficeChoice == ""
    StrCpy $VisualTeXOfficeChoice "native"
  ${EndIf}

  DetailPrint "Checking the Python environment required by VisualTeX OCR..."

  ; Probe every supported runtime through both the new Python Install Manager
  ; selector and the legacy py launcher selector. A default Python 3.14 must not
  ; hide a compatible side-by-side installation.
  !insertmacro VisualTeXProbeLauncher "-V:3.13"
  !insertmacro VisualTeXProbeLauncher "-3.13"
  !insertmacro VisualTeXProbeLauncher "-V:3.12"
  !insertmacro VisualTeXProbeLauncher "-3.12"
  !insertmacro VisualTeXProbeLauncher "-V:3.11"
  !insertmacro VisualTeXProbeLauncher "-3.11"
  !insertmacro VisualTeXProbeLauncher "-V:3.10"
  !insertmacro VisualTeXProbeLauncher "-3.10"
  !insertmacro VisualTeXProbeLauncher "-V:3.9"
  !insertmacro VisualTeXProbeLauncher "-3.9"

  ; Fall back to interpreters exposed directly on PATH.
  !insertmacro VisualTeXProbeCommand "python.exe"
  !insertmacro VisualTeXProbeCommand "python"

  MessageBox MB_ICONEXCLAMATION|MB_YESNO "未检测到可用于 OCR 的 64 位 Python 3.9–3.13。$\r$\n$\r$\nVisualTeX 编辑器仍可正常安装和使用，但图片公式 OCR 将不可用。请安装 x64 Python 3.13，并启用 Python Launcher。仅安装默认 Python 3.14 不兼容当前 OCR 运行环境。$\r$\n$\r$\nNo compatible 64-bit Python 3.9–3.13 installation was detected. The editor can still be installed, but formula OCR will remain unavailable until a compatible Python runtime is installed.$\r$\n$\r$\n是否继续安装？ / Continue installation?" IDYES visualtex_python_continue

  Abort "VisualTeX installation cancelled because the OCR Python prerequisite is missing."

visualtex_python_continue:
  DetailPrint "Continuing without a compatible OCR Python runtime."
  Goto visualtex_python_check_done

visualtex_python_ok:
  DetailPrint "Compatible Python 3.9–3.13 x64 runtime detected."

visualtex_python_check_done:
!macroend

!macro NSIS_HOOK_POSTINSTALL
  DetailPrint "Applying the selected VisualTeX Office integration mode: $VisualTeXOfficeChoice"
  ${If} $VisualTeXOfficeChoice == "native"
    DetailPrint "Installing the machine-wide VisualTeX Ribbon add-ins and native Formula OLE LocalServer. A UAC prompt may appear."
    IfFileExists "$INSTDIR\VisualTeX.exe" 0 visualtex_office_missing
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
    nsExec::ExecToStack `powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\install_windows_vsto_runtime.ps1" -RuntimeInstallerPath "$INSTDIR\windows-office\vstor_redist.exe" -ManifestPath "$INSTDIR\windows-office\vstor_redist.sha256.json" -CheckOnly`
    Pop $1
    Pop $2
    StrCmp $1 "0" visualtex_vsto_runtime_ready 0
    DetailPrint "Microsoft VSTO Runtime is missing. Detection output: $2"
    IfSilent visualtex_vsto_runtime_install 0
    MessageBox MB_ICONQUESTION|MB_YESNO "此电脑尚未安装 Microsoft Visual Studio Tools for Office Runtime。VisualTeX 的 Word/PowerPoint 原生 Ribbon 插件必须依赖该微软组件。$\r$\n$\r$\n安装包已内置微软官方、数字签名有效的 VSTO Runtime。是否现在安装并继续配置 Office 集成？安装过程中会出现 Windows UAC 管理员权限确认。$\r$\n$\r$\n选择“否”仍会保留 VisualTeX 主程序，但会跳过 Office 插件安装。" IDYES visualtex_vsto_runtime_install IDNO visualtex_vsto_runtime_declined

visualtex_vsto_runtime_declined:
    DetailPrint "The user declined Microsoft VSTO Runtime installation. VisualTeX Office integration was skipped."
    Goto visualtex_office_done

visualtex_vsto_runtime_install:
    DetailPrint "Installing the bundled Microsoft VSTO Runtime. A UAC elevation prompt may appear."
    nsExec::ExecToLog `powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\install_windows_vsto_runtime.ps1" -RuntimeInstallerPath "$INSTDIR\windows-office\vstor_redist.exe" -ManifestPath "$INSTDIR\windows-office\vstor_redist.sha256.json"`
    Pop $1
    StrCmp $1 "0" visualtex_vsto_runtime_ready visualtex_vsto_runtime_failed

visualtex_vsto_runtime_failed:
    DetailPrint "Microsoft VSTO Runtime installation failed or the UAC prompt was cancelled. Native Office integration was skipped."
    IfSilent visualtex_office_done 0
    MessageBox MB_ICONEXCLAMATION "Microsoft VSTO Runtime 安装失败，或管理员权限确认被取消。VisualTeX 主程序已经安装，但本次将跳过 Word/PowerPoint 原生插件。$\r$\n$\r$\n请查看 %LOCALAPPDATA%\VisualTeX\office\install-logs 中最新的 vsto-runtime 日志。"
    Goto visualtex_office_done

visualtex_vsto_runtime_ready:
    DetailPrint "Microsoft VSTO Runtime is installed and verified."

    nsExec::ExecToLog `powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\ensure_windows_office_certificate.ps1" -VisualTeXPath "$INSTDIR\VisualTeX.exe"`
    Pop $0
    StrCmp $0 "0" 0 visualtex_office_failed

    nsExec::ExecToLog `powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\install_windows_vsto.ps1" -PackageDirectory "$INSTDIR\windows-office" -VisualTeXPath "$INSTDIR\VisualTeX.exe"`
    Pop $0
    StrCmp $0 "0" visualtex_office_static_installed visualtex_office_failed

visualtex_office_static_installed:
    DetailPrint "Machine-wide Office files and registrations passed. Verifying the companion in the normal user session."
    WriteRegDWORD HKCU "Software\VisualTeX\OfficeIntegration" "RuntimeVerificationPending" 1
    nsExec::ExecToStack `powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\test_windows_office_runtime.ps1" -VisualTeXPath "$INSTDIR\VisualTeX.exe" -CompanionOnly`
    Pop $0
    Pop $1
    ${If} $0 != "0"
      DetailPrint "Companion runtime verification is not ready yet. The machine-wide Office installation remains valid. ExitCode=$0 Output=$1"
      Goto visualtex_office_runtime_pending
    ${EndIf}
    WriteRegDWORD HKCU "Software\VisualTeX\OfficeIntegration" "RuntimeVerificationPending" 0

visualtex_office_static_runtime_verified:
    DetailPrint "VisualTeX native Office static installation and non-elevated companion runtime verification passed."
    IfSilent visualtex_office_done 0
    MessageBox MB_ICONQUESTION|MB_YESNO "Office 插件和本地服务已经安装完成。要确认 Word 与 PowerPoint 的加载项是否真正连接成功，需要临时启动这两个 Office 应用进行验证。$\r$\n$\r$\n请先保存文档并关闭所有正在运行的 Word、PowerPoint 和其他 Office 窗口。是否现在开始验证？$\r$\n$\r$\n选择“否”不会影响 VisualTeX 主程序和插件安装，之后仍可在 VisualTeX 设置中点击“验证 Office 连接”。" IDYES visualtex_office_verify_connections IDNO visualtex_office_verification_deferred

visualtex_office_verify_connections:
    DetailPrint "Launching Word and PowerPoint to verify VisualTeX COMAddIn.Connect."
    nsExec::ExecToLog `powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\test_windows_office_runtime.ps1" -VisualTeXPath "$INSTDIR\VisualTeX.exe"`
    Pop $0
    StrCmp $0 "0" visualtex_office_fully_verified visualtex_office_connection_verification_failed

visualtex_office_fully_verified:
    DetailPrint "Word and PowerPoint COMAddIn.Connect verification passed."
    MessageBox MB_ICONINFORMATION "Word 和 PowerPoint 的 VisualTeX 加载项连接验证成功。打开 VisualTeX 设置时，Office 集成状态将直接显示为可正常使用。"
    Goto visualtex_office_done

visualtex_office_connection_verification_failed:
    DetailPrint "Office connection verification did not complete. The installed integration remains available for in-app retry."
    MessageBox MB_ICONEXCLAMATION "Office 插件已经安装，但本次连接验证没有完成。请确认 Word 和 PowerPoint 已全部关闭，然后进入 VisualTeX 设置点击“验证 Office 连接”；如仍无法关闭，可在提示窗口中选择强制关闭 Office 后验证。"
    Goto visualtex_office_done

visualtex_office_verification_deferred:
    DetailPrint "The user deferred Word and PowerPoint connection verification."
    Goto visualtex_office_done

visualtex_office_runtime_pending:
    DetailPrint "VisualTeX Office files, registry entries, COM classes and OLE server are installed. Companion runtime verification will be retried from VisualTeX settings."
    IfSilent visualtex_office_done 0
    MessageBox MB_ICONINFORMATION "Office 集成已经安装完成，但本地伴侣服务首次启动尚未在安装器等待时间内完成。$\r$\n$\r$\n这不代表插件安装失败。启动 VisualTeX 后会继续初始化，也可以稍后在“设置 → Office 集成”中点击“验证 Office 连接”重新检查。"
    Goto visualtex_office_done

visualtex_office_failed:
    SetDetailsView show
    DetailPrint "VisualTeX main application installed, but the machine-wide Office files, registry entries, COM classes or OLE server failed static installation verification. See the newest vsto-bootstrap and vsto-diagnostic reports under %LOCALAPPDATA%\VisualTeX\office\install-logs."
    IfSilent visualtex_office_done 0
    MessageBox MB_ICONEXCLAMATION "VisualTeX 主程序已安装，但 Office 插件的文件、注册信息、COM 类或 OLE 服务未通过静态安装验证。请查看安装详情，以及 %LOCALAPPDATA%\VisualTeX\office\install-logs 中最新的 vsto-bootstrap 和 vsto-diagnostic 报告。"
    Goto visualtex_office_done
  ${Else}
    IfFileExists "$INSTDIR\scripts\uninstall_windows_vsto.ps1" 0 visualtex_office_done
    nsExec::ExecToLog `powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\uninstall_windows_vsto.ps1"`
    Pop $0
    Goto visualtex_office_done
  ${EndIf}

visualtex_office_missing:
  DetailPrint "Windows native Office installation resources are missing. The VisualTeX main application was installed without Office integration."
  IfSilent visualtex_office_done 0
  MessageBox MB_ICONEXCLAMATION "Windows 原生 Office 安装资源缺失。VisualTeX 主程序已正常安装。"
  Goto visualtex_office_done

visualtex_office_done:
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
    RMDir /r "$APPDATA\VisualTeX"
visualtex_roaming_legacy_done:
  ${EndIf}
!macroend

!macro NSIS_HOOK_PREUNINSTALL
  IfFileExists "$INSTDIR\scripts\uninstall_windows_vsto.ps1" 0 visualtex_preuninstall_done
  nsExec::ExecToStack `powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\uninstall_windows_vsto.ps1"`
  Pop $0
  Pop $1
  ${If} $0 != "0"
    SetDetailsView show
    DetailPrint "VisualTeX Office integration uninstall failed. ExitCode=$0 Output=$1"
    MessageBox MB_ICONSTOP "无法卸载 VisualTeX Office 集成或 HTTPS 证书。主程序尚未删除，您可以关闭 Word 和 PowerPoint 后重试。$\r$\n$\r$\n请查看 %LOCALAPPDATA%\VisualTeX\office\install-logs 中最新的 vsto-uninstall-bootstrap 和 certificate-remove 日志。"
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
