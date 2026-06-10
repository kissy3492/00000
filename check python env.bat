@echo off
rem ============================================================================
rem  Python(embeddable版) 導入可否チェッカー  Phase A  v1.0
rem  - Python不要・管理者権限不要・外部通信なし・レジストリは読取のみ
rem  - 判定結果はこのbatと同じ場所の python_env_check_result.txt に保存
rem ============================================================================
setlocal enabledelayedexpansion
set "REPORT=%~dp0python_env_check_result.txt"
set "PSEXE=powershell -NoProfile -ExecutionPolicy Bypass -Command"
set CRIT_OK=0
set CRIT_NG=0

echo ============================================================ > "%REPORT%"
echo  Python導入可否チェック結果 (Phase A) >> "%REPORT%"
echo  実行日時: %date% %time% >> "%REPORT%"
echo ============================================================ >> "%REPORT%"

echo.
echo ===== Python(embeddable版) 導入可否チェック Phase A =====
echo  結果は %REPORT% にも保存されます
echo.

rem ----------------------------------------------------------------------------
echo [1] OS・PowerShell環境
echo. >> "%REPORT%"
echo [1] OS・PowerShell環境 >> "%REPORT%"
ver >> "%REPORT%"
for /f "usebackq delims=" %%A in (`%PSEXE% "$PSVersionTable.PSVersion.ToString()"`) do set "PSVER=%%A"
echo   PowerShell: !PSVER!
echo   PowerShell: !PSVER! >> "%REPORT%"
for /f "usebackq delims=" %%A in (`%PSEXE% "(Get-CimInstance Win32_OperatingSystem).Caption"`) do set "OSNAME=%%A"
echo   OS: !OSNAME!
echo   OS: !OSNAME! >> "%REPORT%"

rem ----------------------------------------------------------------------------
echo.
echo [2] スクリプト実行ポリシー（参考情報）
echo. >> "%REPORT%"
echo [2] スクリプト実行ポリシー（参考: batとインラインPSは影響を受けない） >> "%REPORT%"
%PSEXE% "Get-ExecutionPolicy -List" >> "%REPORT%" 2>&1
for /f "usebackq delims=" %%A in (`%PSEXE% "Get-ExecutionPolicy"`) do set "EPOL=%%A"
echo   実効ポリシー: !EPOL!
echo   実効ポリシー: !EPOL! >> "%REPORT%"

rem ----------------------------------------------------------------------------
echo.
echo [3] アプリ実行制御ポリシーの有無（AppLocker / SRP / WDAC）
echo. >> "%REPORT%"
echo [3] アプリ実行制御ポリシーの有無 >> "%REPORT%"

reg query "HKLM\SOFTWARE\Policies\Microsoft\Windows\SrpV2" >nul 2>&1
if !errorlevel! equ 0 (
  set "APPLOCKER=設定あり"
) else (
  set "APPLOCKER=設定なし"
)
echo   AppLocker(SrpV2): !APPLOCKER!
echo   AppLocker(SrpV2): !APPLOCKER! >> "%REPORT%"

for /f "usebackq delims=" %%A in (`%PSEXE% "try{$p=Get-AppLockerPolicy -Effective -ErrorAction Stop; $n=0; foreach($rc in $p.RuleCollections){$n=$n+$rc.Count}; 'RULES='+$n}catch{'UNAVAILABLE'}"`) do set "ALRULES=%%A"
echo   AppLocker実効ルール: !ALRULES!
echo   AppLocker実効ルール: !ALRULES! >> "%REPORT%"

reg query "HKLM\SOFTWARE\Policies\Microsoft\Windows\Safer\CodeIdentifiers" >nul 2>&1
if !errorlevel! equ 0 (
  set "SRP=設定あり"
) else (
  set "SRP=設定なし"
)
echo   ソフトウェア制限ポリシー(SRP): !SRP!
echo   ソフトウェア制限ポリシー(SRP): !SRP! >> "%REPORT%"

for /f "usebackq delims=" %%A in (`%PSEXE% "try{$d=Get-CimInstance -Namespace root\Microsoft\Windows\DeviceGuard -ClassName Win32_DeviceGuard -ErrorAction Stop; 'WDAC_ENFORCE='+$d.CodeIntegrityPolicyEnforcementStatus}catch{'WDAC_UNKNOWN'}"`) do set "WDAC=%%A"
echo   WDAC(コード整合性): !WDAC!  (2=強制 1=監査 0=無効)
echo   WDAC(コード整合性): !WDAC!  (2=強制 1=監査 0=無効) >> "%REPORT%"

rem ----------------------------------------------------------------------------
echo.
echo [4] 実地試験: ユーザー領域に置いたexeの起動可否（最重要）
echo. >> "%REPORT%"
echo [4] 実地試験: ユーザー領域に置いたexeの起動可否（最重要） >> "%REPORT%"
echo   ※署名済みWindows標準exe(where.exe)を複製して試行。 >> "%REPORT%"
echo   ※ハッシュ/発行元ベースの規則がある環境では本番(python.exe)で再確認要。 >> "%REPORT%"

set EXEC_OK_ANY=0
for %%D in ("%USERPROFILE%\Documents" "%LOCALAPPDATA%" "%USERPROFILE%\Desktop") do (
  set "TD=%%~D\__pychk__"
  mkdir "!TD!" 2>nul
  copy /y "%SystemRoot%\System32\where.exe" "!TD!\probe_check.exe" >nul 2>&1
  if exist "!TD!\probe_check.exe" (
    "!TD!\probe_check.exe" cmd.exe >nul 2>&1
    if !errorlevel! equ 0 (
      echo   [OK] %%~D : exe起動 可
      echo   [OK] %%~D : exe起動 可 >> "%REPORT%"
      set EXEC_OK_ANY=1
    ) else (
      echo   [NG] %%~D : exe起動 不可 ^(code=!errorlevel!^)
      echo   [NG] %%~D : exe起動 不可 ^(code=!errorlevel!^) >> "%REPORT%"
    )
    del /q "!TD!\probe_check.exe" 2>nul
  ) else (
    echo   [NG] %%~D : 書込またはexe複製が不可
    echo   [NG] %%~D : 書込またはexe複製が不可 >> "%REPORT%"
  )
  rmdir "!TD!" 2>nul
)
if !EXEC_OK_ANY! equ 1 (set /a CRIT_OK+=1) else (set /a CRIT_NG+=1)

rem ----------------------------------------------------------------------------
echo.
echo [5] zip展開手段
echo. >> "%REPORT%"
echo [5] zip展開手段 >> "%REPORT%"
where tar >nul 2>&1
if !errorlevel! equ 0 (
  echo   [OK] tar コマンドあり（zip展開可）
  echo   [OK] tar コマンドあり >> "%REPORT%"
) else (
  echo   [--] tar なし
  echo   [--] tar なし >> "%REPORT%"
)
for /f "usebackq delims=" %%A in (`%PSEXE% "if(Get-Command Expand-Archive -ErrorAction SilentlyContinue){'OK'}else{'NG'}"`) do set "EXPA=%%A"
echo   Expand-Archive: !EXPA!
echo   Expand-Archive: !EXPA! >> "%REPORT%"

rem ----------------------------------------------------------------------------
echo.
echo [6] localhostソケット待受（レビューUI用）
echo. >> "%REPORT%"
echo [6] localhostソケット待受（127.0.0.1のみ・外部送信なし） >> "%REPORT%"
for /f "usebackq delims=" %%A in (`%PSEXE% "try{$l=New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback,0);$l.Start();$p=([Net.IPEndPoint]$l.LocalEndpoint).Port;$c=New-Object Net.Sockets.TcpClient;$c.Connect('127.0.0.1',$p);$ok=$c.Connected;$c.Close();$l.Stop();if($ok){'OK port='+$p}else{'NG'}}catch{'NG '+$_.Exception.Message}"`) do set "SOCK=%%A"
echo   待受試験: !SOCK!
echo   待受試験: !SOCK! >> "%REPORT%"
echo !SOCK! | findstr /b "OK" >nul && (set /a CRIT_OK+=1) || (set /a CRIT_NG+=1)

rem ----------------------------------------------------------------------------
echo.
echo [7] マシン資源（並列OCRの目安）
echo. >> "%REPORT%"
echo [7] マシン資源 >> "%REPORT%"
for /f "usebackq delims=" %%A in (`%PSEXE% "(Get-CimInstance Win32_ComputerSystem).NumberOfLogicalProcessors"`) do set "CORES=%%A"
for /f "usebackq delims=" %%A in (`%PSEXE% "[math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory/1GB,1)"`) do set "RAM=%%A"
for /f "usebackq delims=" %%A in (`%PSEXE% "[math]::Round((Get-PSDrive -Name $env:SystemDrive.Substring(0,1)).Free/1GB,1)"`) do set "FREEGB=%%A"
echo   論理コア: !CORES! / メモリ: !RAM! GB / システムドライブ空き: !FREEGB! GB
echo   論理コア: !CORES! / メモリ: !RAM! GB / 空き: !FREEGB! GB >> "%REPORT%"
echo   （目安: 空き4GB以上を推奨。Python+PaddleOCR一式で2～3GB） >> "%REPORT%"

rem ----------------------------------------------------------------------------
echo.
echo [8] セキュリティ製品（参考情報）
echo. >> "%REPORT%"
echo [8] セキュリティ製品（参考情報） >> "%REPORT%"
for /f "usebackq delims=" %%A in (`%PSEXE% "try{$a=Get-CimInstance -Namespace root\SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction Stop; ($a.displayName -join ' / ')}catch{'取得不可(ドメイン管理環境の可能性)'}"`) do set "AV=%%A"
echo   検出: !AV!
echo   検出: !AV! >> "%REPORT%"

rem ----------------------------------------------------------------------------
echo.
echo ============================================================
echo  総合判定
echo ============================================================
echo. >> "%REPORT%"
echo ============================================================ >> "%REPORT%"
echo  総合判定 >> "%REPORT%"
echo ============================================================ >> "%REPORT%"

if !EXEC_OK_ANY! equ 1 (
  echo   [4]ユーザー領域exe起動: 可  →  技術的にはembeddable版Pythonが動く見込み
  echo   [4]ユーザー領域exe起動: 可 → embeddable版Pythonが動く見込み >> "%REPORT%"
  echo.
  echo   次の手順: Phase B へ
  echo     1. 自宅等で python embeddable zip を入手し、フォルダ展開して持込
  echo     2. smoke_test_python.bat を同じ場所に置いて実行
  echo   次の手順: Phase B（smoke_test_python.bat） >> "%REPORT%"
) else (
  echo   [4]ユーザー領域exe起動: 不可  →  実行制御により持込exeはブロックされます
  echo   [4]ユーザー領域exe起動: 不可 → 持込exeはブロック >> "%REPORT%"
  echo.
  echo   この場合のPython路線は「情報部門への許可申請」が前提になります。
  echo   申請が通らない場合はブラウザ版^(Oculus^)路線の継続が正解です。
  echo   → 申請前提。不可ならブラウザ版路線継続。 >> "%REPORT%"
)
echo.
echo   注意: 本チェックは技術的可否のみを判定します。
echo         規程・運用ルール上の可否は別途確認してください。
echo   注意: 規程上の可否は別途確認のこと。 >> "%REPORT%"
echo.
echo  結果ファイル: %REPORT%
echo.
pause
endlocal
