program PandoraBoxEx;

uses
  Vcl.Forms,
  uMain in 'uMain.pas' {frmMain},
  Vcl.Themes,
  Vcl.Styles,
  uExcel in 'uExcel.pas',
  uAddFile in 'uAddFile.pas' {frmAddFile},
  uUtils in 'uUtils.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  TStyleManager.TrySetStyle('Tablet Dark');
  Application.CreateForm(TfrmMain, frmMain);
  Application.CreateForm(TfrmAddFile, frmAddFile);
  Application.Run;
end.
