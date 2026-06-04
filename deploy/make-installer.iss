; =====================================================================
; Excel Utilities Suite — Inno Setup installer
; Based on VstoAddinInstaller (Apache 2.0)
; Run through Inno Setup Compiler to produce the installer .exe
; =====================================================================

; ── Version ───────────────────────────────────────────────────────────
; Create VERSION.TXT with two lines: semantic version, then 4-number version.
; e.g.:   1.0.0
;         1.0.0.0
#define VERSIONFILE "VERSION.TXT"
#define SEMANTIC_VERSION    "1.1.0"
#define FOUR_NUMBER_VERSION "1.1.0.0"
#define PUB_YEARS           "2024-2026"

; ── Add-in identity ───────────────────────────────────────────────────
#define TARGET_HOST      "excel"
; Generate once via InnoSetup Tools → Generate GUID (keep double-brace prefix):
#define APP_GUID         "{{A3F2C1D0-8B4E-4F7A-9C2E-1D5B6E3F8A02}"
#define ADDIN_NAME       "Excel Utilities Suite"
#define ADDIN_SHORT_NAME "ExcelUtilitiesSuite"
#define COMPANY          "Excel Utilities"
#define DESCRIPTION      "100+ productivity tools for Microsoft Excel."
#define HOMEPAGE         "https://example.com/excel-utilities"
#define HOMEPAGE_SUPPORT "https://example.com/excel-utilities/support"
#define HOMEPAGE_UPDATES "https://example.com/excel-utilities/updates"

; ── Paths ─────────────────────────────────────────────────────────────
#define SOURCEDIR "..\Utilities\bin\Release\"
#define VSTOFILE  "utilities.vsto"
#define OUTPUTDIR "releases\"
#define LOGFILE   "INST-LOG.TXT"

; ── Optional assets ───────────────────────────────────────────────────
; #define LICENSE_FILE          "setup-files\license.rtf"
; #define INSTALLER_ICO         "setup-files\icon.ico"
; #define INSTALLER_IMAGE_LARGE "setup-files\banner-large.bmp"
; #define INSTALLER_IMAGE_SMALL "setup-files\banner-small.bmp"

; ── Optional: build before packaging ─────────────────────────────────
; #define DEVENV "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\devenv.com"
; #define CSPROJ "..\Utilities\utilities.sln"
; #expr Exec(DEVENV, '"' + CSPROJ + '" /Build Release')

; ── VstoAddinInstaller engine ─────────────────────────────────────────
#include "VstoAddinInstaller\vsto-installer.iss"
