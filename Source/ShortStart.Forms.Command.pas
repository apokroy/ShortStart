unit ShortStart.Forms.Command;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  System.Win.ComObj, WinAPI.ActiveX, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.StdCtrls, ShortStart.Types, Vcl.ComCtrls;

const
  IID_IAutoComplete: TGUID = '{00bb2762-6a77-11d0-a535-00c04fd7d062}';
  IID_IAutoComplete2: TGUID = '{EAC04BC0-3791-11d2-BB95-0060977B464C}';
  CLSID_IAutoComplete: TGUID = '{00BB2763-6A77-11D0-A535-00C04FD7D062}';
  IID_IACList: TGUID = '{77A130B0-94FD-11D0-A544-00C04FD7d062}';
  IID_IACList2: TGUID = '{470141a0-5186-11d2-bbb6-0060977b464c}';
  CLSID_ACLHistory: TGUID = '{00BB2764-6A77-11D0-A535-00C04FD7D062}';
  CLSID_ACListISF: TGUID = '{03C036F1-A186-11D0-824A-00AA005B4383}';
  CLSID_ACLMRU: TGUID = '{6756a641-de71-11d0-831b-00aa005b4383}';

type
  IACList = interface(IUnknown)
    ['{77A130B0-94FD-11D0-A544-00C04FD7d062}']
    function Expand(pszExpand: POLESTR): HResult; stdcall;
  end;

const
  { IACList2 }
  ACLO_NONE = 0; {don't enumerate anything}
  ACLO_CURRENTDIR = 1; {enumerate current directory}
  ACLO_MYCOMPUTER = 2; {enumerate MyComputer}
  ACLO_DESKTOP = 4; {enumerate Desktop Folder}
  ACLO_FAVORITES = 8; {enumerate Favorites Folder}
  ACLO_FILESYSONLY = 16; {enumerate only the file system}

type
  IACList2 = interface(IACList)
    ['{470141a0-5186-11d2-bbb6-0060977b464c}']
    function SetOptions(dwFlag: DWORD): HResult; stdcall;
    function GetOptions(var pdwFlag: DWORD): HResult; stdcall;
  end;
  IAutoComplete = interface(IUnknown)
    ['{00bb2762-6a77-11d0-a535-00c04fd7d062}']
    function Init(hwndEdit: HWND; const punkACL: IUnknown; pwszRegKeyPath, pwszQuickComplete: POLESTR): HResult; stdcall;
    function Enable(fEnable: BOOL): HResult; stdcall;
  end;

const
  { IAutoComplete2 }
  ACO_NONE = 0;
  ACO_AUTOSUGGEST = $1;
  ACO_AUTOAPPEND = $2;
  ACO_SEARCH = $4;
  ACO_FILTERPREFIXES = $8;
  ACO_USETAB = $10;
  ACO_UPDOWNKEYDROPSLIST = $20;
  ACO_RTLREADING = $40;

type
  IAutoComplete2 = interface(IAutoComplete)
    ['{EAC04BC0-3791-11d2-BB95-0060977B464C}']
    function SetOptions(dwFlag: DWORD): HResult; stdcall;
    function GetOptions(out pdwFlag: DWORD): HResult; stdcall;
  end;

  TEnumString = class(TInterfacedObject, IEnumString)
  private
    FStrings: TStringList;
    FCurrIndex: integer;
  public
    {IEnumString}
    function Next(celt: Longint; out elt; pceltFetched: PLongint): HResult; stdcall;
    function Skip(celt: Longint): HResult; stdcall;
    function Reset: HResult; stdcall;
    function Clone(out enm: IEnumString): HResult; stdcall;
    {VCL}
    constructor Create;
    destructor Destroy; override;
  end;

  TAutoCompleteOption  = (acAutoAppend, acAutoSuggest, acUseArrowKey);
  TAutoCompleteOptions = set of TAutoCompleteOption;
  TAutoCompleteSource  = (acsList, acsHistory, acsMRU, acsShell);

const
  DefaultAutoCompleteOptions = [acAutoAppend, acAutoSuggest, acUseArrowKey];
  DefaultAutoCompleteEnabled = True;
  DefaultAutoCompleteSource  = acsList;

type
  TAutoCompleteEvent = procedure(Sender: TObject; Index: Integer) of object;

  TEdit = class(Vcl.StdCtrls.TEdit)
  private
    FAutoComplete: IAutoComplete;
    FAutoCompleteEnabled: Boolean;
    FAutoCompleteList: TEnumString;
    FAutoCompleteOptions: TAutoCompleteOptions;
    FAutoCompleteSource: TAutoCompleteSource;
    FOnAutoComplete: TAutoCompleteEvent;
    function  GetAutoCompleteStrings: TStrings;
    procedure SetAutoCompleteEnabled(const Value: Boolean);
    procedure SetAutoCompleteOptions(const Value: TAutoCompleteOptions);
    procedure SetAutoCompleteSource(const Value: TAutoCompleteSource);
    procedure SetAutoCompleteStrings(const Value: TStrings);
  protected
    procedure CreateWnd; override;
    procedure DestroyWnd; override;
    procedure Change; override;
    procedure DoAutoComplete(Index: Integer); virtual;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    property  AutoCompleteEnabled: Boolean read FAutoCompleteEnabled write SetAutoCompleteEnabled default DefaultAutoCompleteEnabled;
    property  AutoCompleteOptions: TAutoCompleteOptions read FAutoCompleteOptions write SetAutoCompleteOptions default DefaultAutoCompleteOptions;
    property  AutoCompleteSource: TAutoCompleteSource read FAutoCompleteSource write SetAutoCompleteSource default DefaultAutoCompleteSource;
    property  AutoCompleteStrings: TStrings read GetAutoCompleteStrings write SetAutoCompleteStrings;
    property  OnAutoComplete: TAutoCompleteEvent read FOnAutoComplete write FOnAutoComplete;
  end;

type
  TCommandForm = class(TForm)
    FileNameEdit: TEdit;
    OperationEdit: TComboBox;
    ParametersEdit: TEdit;
    WorkingDirectoryEdit: TEdit;
    FileOpenDialog: TOpenDialog;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    ShortCutEdit: THotKey;
    Label5: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  public
    class function Edit(const Command: TCommand): Boolean;
  end;

implementation

uses CommCtrl;

{$R *.dfm}

{ TEnumString }

constructor TEnumString.Create;
begin
  inherited Create;
  FStrings := TStringList.Create;
  FCurrIndex := 0;
end;

destructor TEnumString.Destroy;
begin
  FStrings.Free;
  inherited;
end;

function TEnumString.Clone(out enm: IEnumString): HResult;
var
  List: TEnumString;
begin
  Result := S_OK;

  List := TEnumString.Create;
  List.FStrings.Assign(FStrings);
  enm := List;
end;

function TEnumString.Next(celt: Integer; out elt; pceltFetched: PLongint): HResult;
var
  I: Integer;
  wStr: WideString;
begin
  I := 0;
  while (I < celt) and (FCurrIndex < FStrings.Count) do
  begin
    wStr := FStrings[FCurrIndex];
    TPointerList(elt)[I] := CoTaskMemAlloc(2 * (Length(wStr) + 1));
    StringToWideChar(wStr, TPointerList(elt)[I], 2 * (Length(wStr) + 1));
    Inc(I);
    Inc(FCurrIndex);
  end;
  if pceltFetched <> nil then
    pceltFetched^ := I;
  if I = celt then
    Result := S_OK
  else
    Result := S_FALSE;
end;

function TEnumString.Reset: HResult;
begin
  FCurrIndex := 0;
  Result := S_OK;
end;

function TEnumString.Skip(celt: Integer): HResult;
begin
  if (FCurrIndex + celt) <= FStrings.Count then
  begin
    Inc(FCurrIndex, celt);
    Result := S_OK;
  end
  else
  begin
    FCurrIndex := FStrings.Count;
    Result := S_FALSE;
  end;
end;

{ TEdit }

constructor TEdit.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  FAutoCompleteList := TEnumString.Create;
  FAutoCompleteList._AddRef;
  FAutoCompleteEnabled := DefaultAutoCompleteEnabled;
  FAutoCompleteOptions := DefaultAutoCompleteOptions;
  FAutoCompleteSource  := DefaultAutoCompleteSource;
end;

destructor TEdit.Destroy;
begin
  FAutoCompleteList._Release;
  inherited;
end;

procedure TEdit.CreateWnd;
var
  Dummy: IUnknown;
  Strings: IEnumString;
begin
  inherited CreateWnd;

  try
    Dummy := CreateComObject(CLSID_IAutoComplete);
    if (Dummy <> nil) and (Dummy.QueryInterface(IID_IAutoComplete, FAutoComplete) = S_OK) then
    begin
      case FAutoCompleteSource of
        acsHistory:
          Strings := CreateComObject(CLSID_ACLHistory) as IEnumString;
        acsMRU:
          Strings := CreateComObject(CLSID_ACLMRU) as IEnumString;
        acsShell:
          Strings := CreateComObject(CLSID_ACListISF) as IEnumString;
      else
        Strings := FAutoCompleteList as IEnumString;
      end;

      if S_OK = FAutoComplete.Init(Handle, Strings, nil, nil) then
      begin
        SetAutoCompleteEnabled(FAutoCompleteEnabled);
        SetAutoCompleteOptions(FAutoCompleteOptions);
      end;
    end;
  except
    {CLSID_IAutoComplete is not available}
  end;
end;

procedure TEdit.DestroyWnd;
begin
  if (FAutoComplete <> nil) then
  begin
    FAutoComplete.Enable(False);
    FAutoComplete := nil;
  end;

  inherited DestroyWnd;
end;

function TEdit.GetAutoCompleteStrings: TStrings;
begin
  Result := FAutoCompleteList.FStrings;
end;

procedure TEdit.SetAutoCompleteStrings(const Value: TStrings);
begin
  if Value = nil then
    FAutoCompleteList.FStrings.Clear
  else
    FAutoCompleteList.FStrings.Assign(Value);
end;

procedure TEdit.SetAutoCompleteEnabled(const Value: Boolean);
begin
  if (FAutoComplete <> nil) then
  begin
    FAutoComplete.Enable(FAutoCompleteEnabled);
  end;
  FAutoCompleteEnabled := Value;
end;

procedure TEdit.SetAutoCompleteOptions(const Value: TAutoCompleteOptions);
const
  Options: array[TAutoCompleteOption] of integer = (ACO_AUTOAPPEND, ACO_AUTOSUGGEST,
    ACO_UPDOWNKEYDROPSLIST);
var
  Option: TAutoCompleteOption;
  Opt: DWORD;
  AC2: IAutoComplete2;
begin
  if (FAutoComplete <> nil) then
  begin
    if S_OK = FAutoComplete.QueryInterface(IID_IAutoComplete2, AC2) then
    begin
      Opt := ACO_NONE;
      for Option := Low(Options) to High(Options) do
      begin
        if (Option in FAutoCompleteOptions) then
          Opt := Opt or DWORD(Options[Option]);
      end;
      AC2.SetOptions(Opt);
    end;
  end;
  FAutoCompleteOptions := Value;
end;

procedure TEdit.SetAutoCompleteSource(const Value: TAutoCompleteSource);
begin
  if FAutoCompleteSource <> Value then
  begin
    FAutoCompleteSource := Value;
    RecreateWnd;
  end;
end;

procedure TEdit.Change;
var
  I: Integer;
begin
  if AutoCompleteEnabled and (AutoCompleteSource = acsList) then
  begin
    I := AutoCompleteStrings.IndexOf(Text);
    if I >= 0 then
      DoAutoComplete(I);
  end;

  inherited Change;
end;

procedure TEdit.DoAutoComplete(Index: Integer);
begin
  if Assigned(OnAutoComplete) then
    OnAutoComplete(Self, Index);
end;

{ TCommandForm }

class function TCommandForm.Edit(const Command: TCommand): Boolean;
var
  Dlg: TCommandForm;
begin
  Dlg := TCommandForm.Create(Application);
  try
    Dlg.FileNameEdit.Text := Command.FileName;
    Dlg.ShortCutEdit.HotKey := Command.ShortCut;
    Dlg.OperationEdit.Text := Command.Operation;
    Dlg.ParametersEdit.Text := Command.Parameters;
    Dlg.WorkingDirectoryEdit.Text := Command.WorkingDirectory;

    Result := Dlg.ShowModal = mrOk;
    if Result then
    begin
      Command.FileName := Dlg.FileNameEdit.Text;
      Command.ShortCut := Dlg.ShortCutEdit.HotKey;
      Command.Operation := Dlg.OperationEdit.Text;
      Command.Parameters := Dlg.ParametersEdit.Text;
      Command.WorkingDirectory := Dlg.WorkingDirectoryEdit.Text;
    end;
  finally
    Dlg.Free;
  end;
end;

procedure TCommandForm.Button1Click(Sender: TObject);
begin
  FileOpenDialog.FileName := FileNameEdit.Text;
  if FileOpenDialog.Execute then
    FileNameEdit.Text := FileOpenDialog.FileName;
end;

procedure TCommandForm.FormCreate(Sender: TObject);
begin
  FileNameEdit.AutoCompleteSource := acsShell;
  FileNameEdit.AutoCompleteEnabled := True;
  WorkingDirectoryEdit.AutoCompleteSource := acsShell;
  WorkingDirectoryEdit.AutoCompleteEnabled := True;
end;

end.
