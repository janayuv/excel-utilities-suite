; idp.iss — compile-only stub for VstoAddinInstaller
;
; The real Inno Download Plugin (IDP) is only needed to download .NET or the
; VSTO runtime on machines that don't already have Office installed. If your
; target machines have Office, none of these functions run at runtime.
;
; For real download support, replace with the full IDP from:
;   https://github.com/Octanox/inno-download-plugin/releases
;
; ── HOW TO INSTALL THIS STUB ─────────────────────────────────────────────────
; Copy THIS FILE to:
;   C:\Program Files (x86)\Inno Setup 6\idp.iss
; Then recompile make-installer.iss — the error will be gone.

[Code]
type
  TIdpPage = record
    ID: Integer;
  end;
  TIdpForm = record
    Page: TIdpPage;
  end;

var
  IDPForm: TIdpForm;

procedure idpAddFileSize(Url, FileName: String; Size: Int64);
begin
  { No-op: IDP stub — runtime download not available }
end;

procedure idpDownloadAfter(PageId: Integer);
begin
  { No-op: IDP stub }
end;

function idpFilesCount: Integer;
begin
  Result := 0; { Always 0 → all download wizard pages are skipped }
end;

{ wizard-pages.pas uses this to get total download size }
function idpGetFilesSize(var Size: Int64): Boolean;
begin
  Size   := 0;
  Result := False;
end;
