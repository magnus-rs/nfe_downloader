unit U_Principal;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Menus, Vcl.Grids, Vcl.ExtCtrls,
  Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ToolWin, ACBrBase, ACBrDFe, ACBrNFe, ACBrDFeSSL,
  ACBrNFeConfiguracoes, ACBrDFeConfiguracoes;

type
  TForm_Principal = class(TForm)
    MainMenu1: TMainMenu;
    Arquivo1: TMenuItem;
    Sair1: TMenuItem;
    Sair2: TMenuItem;
    Configuracoes1: TMenuItem;
    Certificados1: TMenuItem;
    ToolBar1: TToolBar;
    StatusBar1: TStatusBar;
    ACBrNFe1: TACBrNFe;
    Panel1: TPanel;
    ToolBar2: TToolBar;
    TreeView1: TTreeView;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    ToolButton9: TToolButton;
    ToolButton10: TToolButton;
    Panel2: TPanel;
    Panel3: TPanel;
    TreeView2: TTreeView;
    Panel4: TPanel;
    StringGrid1: TStringGrid;
    procedure FormCreate(Sender: TObject);
    procedure ConfigurarACBr;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form_Principal: TForm_Principal;

implementation

{$R *.dfm}

procedure TForm_Principal.FormCreate(Sender: TObject);
begin
  ConfigurarACBr;
end;

procedure TForm_Principal.ConfigurarACBr;
begin
  SetDllDirectory(PChar(ExtractFilePath(Application.ExeName) + 'openssl'));

  with ACBrNFe1.Configuracoes do
  begin
    Geral.SSLLib := libOpenSSL;
    Geral.SSLCryptLib := cryOpenSSL;
    Geral.SSLHttpLib := httpOpenSSL;
    Geral.SSLXmlSignLib := xsLibXml2;

    Arquivos.PathSchemas :=
      ExtractFilePath(Application.ExeName) + 'Schemas\NFe\';

    Certificados.ArquivoPFX := 'D:\certificados\1007702659.pfx';
    Certificados.Senha := 'guanabara';

    WebServices.UF := 'RS';
  end;

end;


end.
