unit U_Certificados;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Grids;

type
  TForm_certificados = class(TForm)
    grid_certificados: TStringGrid;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form_certificados: TForm_certificados;

implementation

{$R *.dfm}

end.
