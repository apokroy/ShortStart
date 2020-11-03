unit ShortStart.Forms.About;

interface

uses WinApi.Windows, System.SysUtils, System.Classes, Vcl.Graphics,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtCtrls;

type
  TAboutBox = class(TForm)
    ScrollBox1: TScrollBox;
    Comments: TLabel;
    Copyright: TLabel;
    ProgramIcon: TImage;
    Version: TLabel;
    ProductName: TLabel;
    OKButton: TButton;
    Label1: TLabel;
    procedure FormCreate(Sender: TObject);
  private
  public
  end;

implementation

{$R *.dfm}

procedure TAboutBox.FormCreate(Sender: TObject);
begin
  ProgramIcon.Picture.Icon := Application.Icon;
end;

end.
 
