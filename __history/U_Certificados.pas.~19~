unit U_Certificados;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Grids, Vcl.StdCtrls, Vcl.ExtCtrls,
  Vcl.Mask, ACBrNFe, ACBrDFeSSL, System.DateUtils, uEmpresa, uEmpresaService,
  ACBrBase, ACBrDFe;

type
  TForm_certificados = class(TForm)
    EditCNPJ: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    EditNome: TEdit;
    Bevel1: TBevel;
    Label3: TLabel;
    EditArquivo: TEdit;
    BtnImportar: TButton;
    OpenDialog1: TOpenDialog;
    Label4: TLabel;
    EditValidade: TEdit;
    Label5: TLabel;
    EditSenha: TMaskEdit;
    ACBrNFe1: TACBrNFe;
    procedure BtnImportarClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    procedure CarregarCertificado(const Caminho: string);
    procedure PreencherDados;
    procedure SalvarEmpresa;
    procedure ConfigurarACBr;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form_certificados: TForm_certificados;

implementation

{$R *.dfm}

procedure TForm_Certificados.ConfigurarACBr;
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

  end;

end;

procedure TForm_certificados.FormCreate(Sender: TObject);
begin
   ConfigurarACBr;
end;

procedure TForm_certificados.BtnImportarClick(Sender: TObject);
begin
  OpenDialog1.Filter := 'Certificado (*.pfx)|*.pfx';

  if OpenDialog1.Execute then
  begin
    EditArquivo.Text := OpenDialog1.FileName;

    CarregarCertificado(OpenDialog1.FileName);

    SalvarEmpresa;
  end;
end;

procedure TForm_certificados.CarregarCertificado(const Caminho: string);
var
  Senha: string;
begin
  // Usa senha digitada ou solicita
  Senha := EditSenha.Text;

  if Senha = '' then
    Senha := InputBox('Certificado', 'Digite a senha do certificado:', '');

  if Senha = '' then
    Exit;

  try
    Screen.Cursor := crHourGlass;

    ACBrNFe1.Configuracoes.Certificados.ArquivoPFX := Caminho;
    ACBrNFe1.Configuracoes.Certificados.Senha := Senha;

    ACBrNFe1.SSL.CarregarCertificado;

    EditSenha.Text := Senha;

    PreencherDados;

  except
    on E: Exception do
      ShowMessage('Erro ao carregar certificado: ' + E.Message);
  end;

  Screen.Cursor := crDefault;
end;

procedure TForm_certificados.PreencherDados;
var
  DataFim: TDateTime;
begin
  // CNPJ e Nome
  EditCNPJ.Text := ACBrNFe1.SSL.CertCNPJ;
  EditNome.Text := ACBrNFe1.SSL.CertRazaoSocial;

  // Validade (normalmente o ACBr retorna vencimento)
  DataFim := ACBrNFe1.SSL.CertDataVenc;

  EditValidade.Text := DateToStr(DataFim);
end;

procedure TForm_certificados.SalvarEmpresa;
var
  Service: TEmpresaService;
  Empresa: TEmpresa;
begin
  if EditCNPJ.Text = '' then
  begin
    ShowMessage('CNPJ não informado.');
    Exit;
  end;

  Service := TEmpresaService.Create;
  try
    Service.Carregar;

    if Assigned(Service.Buscar(EditCNPJ.Text)) then
    begin
      ShowMessage('Empresa já cadastrada.');
      Exit;
    end;

    Empresa := TEmpresa.Create;
    Empresa.CNPJ := EditCNPJ.Text;
    Empresa.RazaoSocial := EditNome.Text;
    Empresa.NomeFantasia := EditNome.Text;
    Empresa.CertificadoPath := EditArquivo.Text;
    Empresa.CertificadoSenha := EditSenha.Text;
    Empresa.ValidadeInicio := ACBrNFe1.SSL.DadosCertificado.DataInicioValidade; // se não tiver início, pode ajustar depois
    Empresa.ValidadeFim := StrToDateDef(EditValidade.Text, Now);

    Service.Adicionar(Empresa);
    Service.Salvar;

    ShowMessage('Empresa salva com sucesso!');

  finally
    Service.Free;
  end;
end;

end.
