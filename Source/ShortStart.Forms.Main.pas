unit ShortStart.Forms.Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  System.UITypes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls,
  Vcl.StdCtrls, Vcl.ComCtrls, System.Actions, Vcl.ActnList, System.ImageList,
  Vcl.ImgList, Vcl.ToolWin, Vcl.Menus,
  ShortStart.Types, System.Notification;

type
  TMainForm = class(TForm)
    Tray: TTrayIcon;
    ToolBar: TToolBar;
    ActionImageList: TImageList;
    ToolButton1: TToolButton;
    ActionList: TActionList;
    View: TListView;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    FileExitAction: TAction;
    AddAction: TAction;
    EditAction: TAction;
    DeleteAction: TAction;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    FileSaveAction: TAction;
    FileRevertAction: TAction;
    MainMenu: TMainMenu;
    File1: TMenuItem;
    Shorts1: TMenuItem;
    Help1: TMenuItem;
    Save1: TMenuItem;
    Revert1: TMenuItem;
    N1: TMenuItem;
    Exit1: TMenuItem;
    Add1: TMenuItem;
    Edit1: TMenuItem;
    Edit2: TMenuItem;
    AboutAction: TAction;
    About1: TMenuItem;
    TrayImageList: TImageList;
    TrayMenu: TPopupMenu;
    Exit2: TMenuItem;
    N2: TMenuItem;
    Show1: TMenuItem;
    TrayDefaultAction: TAction;
    ToolButton9: TToolButton;
    ToolButton10: TToolButton;
    RunAction: TAction;
    EnableAction: TAction;
    DisableAction: TAction;
    ToolButton11: TToolButton;
    ToolButton12: TToolButton;
    ToolButton13: TToolButton;
    N3: TMenuItem;
    Enable1: TMenuItem;
    Disable1: TMenuItem;
    NotificationCenter: TNotificationCenter;
    procedure FormCreate(Sender: TObject);
    procedure TrayClick(Sender: TObject);
    procedure About(Sender: TObject);
    procedure FileExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FileReload(Sender: TObject);
    procedure FileSave(Sender: TObject);
    procedure AddExecute(Sender: TObject);
    procedure EditExecute(Sender: TObject);
    procedure DeleteExecute(Sender: TObject);
    procedure SelectedActionUpdate(Sender: TObject);
    procedure ViewDblClick(Sender: TObject);
    procedure RunExecute(Sender: TObject);
    procedure EnableActionExecute(Sender: TObject);
    procedure DisableActionExecute(Sender: TObject);
    procedure EnableActionUpdate(Sender: TObject);
    procedure DisableActionUpdate(Sender: TObject);
    procedure NotificationCenterReceiveLocalNotification(Sender: TObject;
      ANotification: TNotification);
  private
    HomePath: string;
    SettingsFileName: string;
    Modified: Boolean;
    procedure UpdateView;
    procedure UpdateItem(Item: TListItem; Command: TCommand);
    procedure CheckModified;
  protected
    procedure CreateParams(var Params: TCreateParams); override;
  public
    procedure HandleError(const E: Exception);
    procedure Add;
    procedure Edit;
    procedure Delete;
    procedure Reload;
    procedure Save;
  end;

var
  MainForm: TMainForm;

implementation

uses System.IOUtils, ShortStart.Forms.About, ShortStart.Forms.Command;

resourcestring
  SSaveChanges = 'Save configuration changes?';
  SConfirmDeletion = 'Do you really want to delete this item?';

{$R *.dfm}

{ TMainForm }

procedure TMainForm.CreateParams(var Params: TCreateParams);
begin
  inherited;
  Params.ExStyle := Params.ExStyle and not WS_EX_APPWINDOW;
  Params.WndParent := Application.Handle;
end;

procedure TMainForm.FormCreate(Sender: TObject);
begin
  Caption := Application.Title;

  HomePath := IncludeTrailingPathDelimiter(TPath.GetHomePath) + 'ShortStart' + TPath.DirectorySeparatorChar;
  ForceDirectories(HomePath);
  SettingsFileName := HomePath + 'Application.config';

  Reload;
end;

procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  CheckModified;

  Action := caNone;
  Hide;
end;

procedure TMainForm.HandleError(const E: Exception);
var
  Notification: TNotification;
begin
  Notification := NotificationCenter.CreateNotification('Error', E.Message, Now);
  if Notification <> nil then
  begin
    Notification.Title := Application.Title;
    NotificationCenter.PresentNotification(Notification);
  end;
end;

procedure TMainForm.NotificationCenterReceiveLocalNotification(Sender: TObject; ANotification: TNotification);
begin
  Show;
end;

procedure TMainForm.UpdateItem(Item: TListItem; Command: TCommand);
begin
  Item.Caption := Command.FileName;
  Item.SubItems.Clear;
  Item.SubItems.Add(ShortCutToText(Command.ShortCut));
  Item.SubItems.Add(Command.Operation);
  Item.SubItems.Add(Command.Parameters);
  Item.SubItems.Add(Command.WorkingDirectory);
  Item.Data := Command;
end;

procedure TMainForm.UpdateView;
var
  I: Integer;
begin
  View.Items.BeginUpdate;
  try
    View.Items.Clear;

    for I := 0 to Commands.Count - 1 do
      UpdateItem(View.Items.Add, Commands[I]);
  finally
    View.Items.EndUpdate;
  end;
end;

procedure TMainForm.Reload;
begin
  if FileExists(SettingsFileName) then
  begin
    Commands.LoadFromFile(SettingsFileName);
    Modified := False;
    UpdateView;
  end;
end;

procedure TMainForm.Save;
begin
  Commands.SaveToFile(SettingsFileName);
  Modified := False;
end;

procedure TMainForm.CheckModified;
begin
  if not Modified then
    Exit;

  case MessageDlg(SSaveChanges, mtConfirmation, [mbYes, mbNo, mbCancel], 0) of
    mrYes: Save;
    mrCancel: Abort;
  end;
end;

procedure TMainForm.Add;
var
  Command: TCommand;
begin
  Command := TCommand.Create;
  if TCommandForm.Edit(Command) then
  begin
    Commands.Add(Command);
    Modified := True;
    UpdateItem(View.Items.Add, Command);
  end;
end;

procedure TMainForm.Edit;
var
  Command: TCommand;
begin
  if View.Selected = nil then
    Exit;

  Command := TCommand(View.Selected.Data);
  if TCommandForm.Edit(Command) then
  begin
    Modified := True;
    UpdateItem(View.Selected, Command);
  end;
end;

procedure TMainForm.Delete;
var
  Command: TCommand;
begin
  if View.Selected = nil then
    Exit;

  Command := TCommand(View.Selected.Data);

  if MessageDlg(SConfirmDeletion, mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
    Exit;

  Commands.Remove(Command);
  Command.Free;
  Modified := True;

  View.Selected.Delete;
end;

procedure TMainForm.TrayClick(Sender: TObject);
begin
  Show;
end;

procedure TMainForm.FileExit(Sender: TObject);
begin
  CheckModified;
  Application.Terminate;
end;

procedure TMainForm.FileReload(Sender: TObject);
begin
  Reload;
end;

procedure TMainForm.FileSave(Sender: TObject);
begin
  Save;
end;

procedure TMainForm.AddExecute(Sender: TObject);
begin
  Add;
end;

procedure TMainForm.EditExecute(Sender: TObject);
begin
  Edit;
end;

procedure TMainForm.DeleteExecute(Sender: TObject);
begin
  Delete;
end;

procedure TMainForm.EnableActionExecute(Sender: TObject);
begin
  Commands.Enabled := True;
end;

procedure TMainForm.EnableActionUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := not Commands.Enabled;
  if Commands.Enabled then
    Tray.IconIndex := 0
  else
    Tray.IconIndex := 1;
end;

procedure TMainForm.DisableActionExecute(Sender: TObject);
begin
  Commands.Enabled := False;
end;

procedure TMainForm.DisableActionUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := Commands.Enabled;
end;

procedure TMainForm.ViewDblClick(Sender: TObject);
begin
  Edit;
end;

procedure TMainForm.SelectedActionUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := View.Selected <> nil;
end;

procedure TMainForm.RunExecute(Sender: TObject);
begin
  TCommand(View.Selected.Data).Execute;
end;

procedure TMainForm.About(Sender: TObject);
var
  Dlg: TAboutBox;
begin
  Dlg := TAboutBox.Create(Self);
  try
    Dlg.ShowModal;
  finally
    Dlg.Free;
  end;
end;

end.
