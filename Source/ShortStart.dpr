program ShortStart;

uses
  Vcl.Forms,
  ShortStart.Forms.Main in 'ShortStart.Forms.Main.pas' {MainForm},
  ShortStart.Forms.About in 'ShortStart.Forms.About.pas' {AboutBox},
  ShortStart.Types in 'ShortStart.Types.pas',
  ShortStart.Forms.Command in 'ShortStart.Forms.Command.pas' {CommandForm};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'Short Start';
  Application.ShowMainForm := False;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
