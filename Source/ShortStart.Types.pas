unit ShortStart.Types;

interface

uses
 Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
 System.Generics.Collections, Vcl.Forms, Vcl.Menus;

const
  CM_COMMAND = WM_USER + 1;

type
  TCommand = class(TPersistent)
  private
    FFileName: string;
    FWorkingDirectory: string;
    FParameters: string;
    FOperation: string;
    FShortCut: TShortCut;
    FShowCommand: Integer;
  public
    constructor Create;
    procedure Execute;
  published
    property  ShortCut: TShortCut read FShortCut write FShortCut;
    property  FileName: string read FFileName write FFileName;
    property  Operation: string read FOperation write FOperation;
    property  Parameters: string read FParameters write FParameters;
    property  WorkingDirectory: string read FWorkingDirectory write FWorkingDirectory;
    property  ShowCommand: Integer read FShowCommand write FShowCommand default SW_SHOWNORMAL;
  end;

  TCommandList = class(TList<TCommand>)
  private
    FEnabled: Boolean;
    FHandle: THandle;
    procedure SetEnabled(const Value: Boolean);
  protected
    procedure WndProc(var Msg: TMessage);
  public
    constructor Create;
    destructor Destroy; override;
    function  Add(const ShortCut: TShortCut; const FileName: string; const Operation: string = ''; const Parameters: string = ''; const WorkingDirectory: string = ''): TCommand; overload;
    function  Add(const Key: Word; const ShiftState: TShiftState; const FileName: string; const Operation: string = ''; const Parameters: string = ''; const WorkingDirectory: string = ''): TCommand; overload;
    function  Find(const ShortCut: TShortCut): TCommand;
    procedure LoadFromStream(Stream: TStream);
    procedure SaveToStream(Stream: TStream);
    procedure LoadFromFile(const FileName: string);
    procedure SaveToFile(const FileName: string);
    property  Enabled: Boolean read FEnabled write SetEnabled;
    property  Handle: THandle read FHandle;
  end;

var
  Commands: TCommandList;

implementation

uses WinAPI.ActiveX, WinAPI.ShellAPI, XML.XmlDoc, XML.XMLIntf,
  ShortStart.Forms.Main;

procedure RaiseLastOSError(LastError: Integer);
var
  Error: EOSError;
begin
  Error := EOSError.Create(SysErrorMessage(LastError));
  raise Error at ReturnAddress;
end;

procedure CheckOSError(Flag: BOOL);
begin
  if not Flag then
    RaiseLastOSError(GetLastError);
end;

var
  ProcWnd: HWND;

function EnumProc(hwnd: HWND; ProcessId: LPARAM): BOOL; stdcall;
var
  TestId: DWORD;
begin
  GetWindowThreadProcessId(hwnd, TestId);
  if (TestId = DWORD(ProcessId)) and (IsWindowVisible(hwnd)) then
  begin
    ProcWnd := hwnd;
    Result := False;
  end
  else
    Result := True;
end;

function FindProcessWnd(hProcess: THandle): HWND;
begin
  ProcWnd := 0;
  EnumWindows(@EnumProc, GetProcessId(hProcess));
  Result := ProcWnd;
end;

{ TCommand }

constructor TCommand.Create;
begin
  inherited Create;
  FShowCommand := SW_SHOWNORMAL;
end;

procedure TCommand.Execute;
var
  ExecInfo: TShellExecuteInfo;
  NeedUninitialize: Boolean;
  Wnd: HWND;
  Count: Integer;
begin
  // This
  //   1: Run application in foreground
  //   2: Prevent keyboard focus from currently active applicaion
  SetForegroundWindow(Application.MainForm.Handle);

  NeedUninitialize := Succeeded(CoInitializeEx(nil, COINIT_APARTMENTTHREADED or COINIT_DISABLE_OLE1DDE));
  try
    FillChar(ExecInfo, SizeOf(ExecInfo), 0);
    ExecInfo.cbSize := SizeOf(ExecInfo);

    ExecInfo.Wnd := Application.MainForm.Handle;
    ExecInfo.lpVerb := Pointer(Operation);
    ExecInfo.lpFile := PChar(FileName);
    ExecInfo.lpParameters := Pointer(Parameters);
    ExecInfo.lpDirectory := Pointer(WorkingDirectory);
    ExecInfo.nShow := ShowCommand;
    ExecInfo.fMask := SEE_MASK_NOCLOSEPROCESS or SEE_MASK_NOASYNC or SEE_MASK_FLAG_NO_UI or SEE_MASK_UNICODE;

    {$WARN SYMBOL_PLATFORM OFF}
    CheckOSError(ShellExecuteEx(@ExecInfo));
    {$WARN SYMBOL_PLATFORM ON}

    // Try to set focus in opened application
    if ExecInfo.hProcess <> 0 then
    begin
      AllowSetForegroundWindow(GetProcessId(ExecInfo.hProcess));
      Count := 0;
      Wnd := FindProcessWnd(ExecInfo.hProcess);
      while (Wnd = 0) and (Count < 20) do
      begin
        Sleep(100);
        Wnd := FindProcessWnd(ExecInfo.hProcess);
        Inc(Count);
      end;

      if Wnd <> 0 then
      begin
        SetForegroundWindow(Wnd);
        SetActiveWindow(Wnd);
        SetFocus(Wnd);
      end;
    end;
  finally
    if NeedUninitialize then
      CoUninitialize;
  end;
end;

{ TCommandList }

constructor TCommandList.Create;
begin
  inherited Create;

  FEnabled := True;

  Add(Ord('N'), [ssAlt, ssCtrl], 'notepad.exe');
  Add(Ord('C'), [ssAlt, ssCtrl], 'calc.exe');
  Add(Ord('P'), [ssAlt, ssCtrl], 'mspaint.exe');
  Add(Ord('D'), [ssAlt, ssCtrl], 'explorer.exe', '', '%DEFAULTUSERPROFILE%');

  FHandle := AllocateHWnd(WndProc);
end;

destructor TCommandList.Destroy;
begin
  DeallocateHWnd(FHandle);
  FHandle := 0;

  inherited;
end;

procedure TCommandList.WndProc(var Msg: TMessage);
var
  Command: TCommand;
begin
  if Msg.Msg = CM_COMMAND then
  begin
    try
      Command := Find(Msg.WParam);
      if Command <> nil then
        Command.Execute;
    except
      on E: Exception do
        MainForm.HandleError(E);
    end;
    Msg.Result := 0;
  end
  else
    Msg.Result := DefWindowProc(FHandle, Msg.Msg, Msg.wParam, Msg.lParam);
end;

function TCommandList.Add(const Key: Word; const ShiftState: TShiftState; const FileName, Operation, Parameters, WorkingDirectory: string): TCommand;
begin
  Result := Add(ShortCut(Key, ShiftState), FileName, Operation, Parameters, WorkingDirectory);
end;

function TCommandList.Add(const ShortCut: TShortCut; const FileName, Operation, Parameters, WorkingDirectory: string): TCommand;
begin
  Result := TCommand.Create;
  Result.ShortCut := ShortCut;
  Result.FileName := FileName;
  Result.Operation := Operation;
  Result.Parameters := Parameters;
  Result.WorkingDirectory := WorkingDirectory;

  Add(Result);
end;

function TCommandList.Find(const ShortCut: TShortCut): TCommand;
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
    if Items[I].ShortCut = ShortCut then
    begin
      Result := Items[I];
      Exit;
    end;
  Result := nil;
end;

procedure TCommandList.LoadFromFile(const FileName: string);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(FileName, fmOpenRead or fmShareDenyWrite);
  try
    LoadFromStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TCommandList.SaveToFile(const FileName: string);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(FileName, fmCreate);
  try
    SaveToStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TCommandList.LoadFromStream(Stream: TStream);
var
  Doc: IXMLDocument;
  Root, Node: IXMLNode;
  Nodes: IXMLNodeList;
  I, J: Integer;
begin
  Doc := TXMLDocument.Create(nil);
  Doc.LoadFromStream(Stream);

  Clear;

  Root := Doc.DocumentElement;
  for I := 0 to Root.ChildNodes.Count - 1 do
  begin
    Node := Root.ChildNodes[I];
    if Node.NodeName = 'Commands' then
    begin
      Nodes := Node.ChildNodes;
      for J := 0 to Nodes.Count - 1 do
      begin
        Node := Nodes[J];
        Add(
          TextToShortCut(Node.Attributes['ShortCut']),
          VarToStr(Node.Attributes['FileName']),
          VarToStr(Node.Attributes['Operation']),
          VarToStr(Node.Attributes['Parameters']),
          VarToStr(Node.Attributes['WorkingDirectory'])
        );
      end;
    end;
  end;
end;

procedure TCommandList.SaveToStream(Stream: TStream);
var
  Doc: IXMLDocument;
  OptionsNode, CommandsNode, Node: IXMLNode;
  I: Integer;
  Command: TCommand;
begin
  Doc := TXMLDocument.Create(nil);
  Doc.Active := True;
  Doc.DocumentElement := Doc.CreateNode('Application', ntElement);
  OptionsNode := Doc.DocumentElement.AddChild('Options'); //Reserved for application options
  CommandsNode := Doc.DocumentElement.AddChild('Commands');
  for I := 0 to Count - 1 do
  begin
    Command := Items[I];

    Node := CommandsNode.AddChild('Command');
    Node.Attributes['ShortCut'] := ShortCutToText(Command.ShortCut);
    Node.Attributes['FileName'] := Command.FileName;
    Node.Attributes['Operation'] := Command.Operation;
    Node.Attributes['Parameters'] := Command.Parameters;
    Node.Attributes['WorkingDirectory'] := Command.WorkingDirectory;
  end;

  Doc.SaveToStream(Stream);
end;

procedure TCommandList.SetEnabled(const Value: Boolean);
begin
  FEnabled := Value;
end;

type
  KBDLLHOOKSTRUCT = record
    vkCode: DWORD;
    scanCode: DWORD;
    flags: DWORD;
    time: DWORD;
    dwExtraInfo: DWORD;
  end;
  PBDLLHOOKSTRUCT = ^KBDLLHOOKSTRUCT;
var
  hKeyboardHook: HHOOK;

function KeyboardEvent(Code: Integer; wParam: WPARAM; lParam: LPARAM): LRESULT stdcall;
var
  ShortCut: TShortCut;
begin
  try
    if Commands.Enabled and (Code = HC_ACTION) and (wParam = WM_KEYDOWN) then
    begin
      ShortCut := PBDLLHOOKSTRUCT(lParam).vkCode;
      if GetAsyncKeyState(VK_SHIFT) < 0 then
        Inc(ShortCut, scShift);
      if GetAsyncKeyState(VK_CONTROL) < 0 then
        Inc(ShortCut, scCtrl);
      if GetAsyncKeyState(VK_MENU) < 0 then
        Inc(ShortCut, scAlt);

      // Send over message queue
      PostMessage(Commands.Handle, CM_COMMAND, ShortCut, 0);
    end;
  except
  end;
  Result := CallNextHookEx(hKeyboardHook, Code, wParam, lParam);
end;

initialization
  Commands := TCommandList.Create;

  hKeyboardHook := SetWindowsHookEx(WH_KEYBOARD_LL, @KeyboardEvent, hInstance,  0);

finalization
  UnhookWindowsHookEx(hKeyboardHook);
  FreeAndNil(Commands);
end.
