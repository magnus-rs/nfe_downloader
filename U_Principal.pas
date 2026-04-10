unit U_Principal;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Menus, Vcl.Grids, Vcl.ExtCtrls,
  Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ToolWin, System.DateUtils, System.UITypes,
  ACBrBase,
  ACBrDFe,
  ACBrNFe,
  ACBrDFeSSL,
  ACBrNFeConfiguracoes,
  ACBrDFeConfiguracoes,
  pcnConversaoNFe,
  pcnConversao,
  ACBrNFeNotasFiscais,
  ACBrNFeWebServices,
  uEmpresa,
  uEmpresaService,
  System.Generics.Collections, Data.DB, Vcl.Buttons, Vcl.DBGrids,
  Datasnap.DBClient, FileCtrl;

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
    Panel5: TPanel;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    Panel6: TPanel;
    StringGrid1: TStringGrid;
    Label1: TLabel;
    BTN_Importar_Certificado: TButton;
    BTN_Remover_Certificado: TButton;
    BTN_Atualizar_Lista: TButton;
    CDataSet_NFE_Entrada: TClientDataSet;
    DataSource_NFE_Entrada: TDataSource;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    Label2: TLabel;
    Label3: TLabel;
    Button_Buscar: TSpeedButton;
    Edit_Pasta: TEdit;
    Button_Selecionar: TSpeedButton;
    Label4: TLabel;
    Label5: TLabel;
    CDataSet_NFE_EntradaNFeNúmero: TStringField;
    CDataSet_NFE_EntradaNFeSérie: TStringField;
    CDataSet_NFE_EntradaNFeCTeTipo: TStringField;
    CDataSet_NFE_EntradaEmissão: TStringField;
    CDataSet_NFE_EntradaValor: TCurrencyField;
    CDataSet_NFE_EntradaVencimento: TDateField;
    CDataSet_NFE_EntradaEmitente: TStringField;
    CDataSet_NFE_EntradaCFOP: TStringField;
    CDataSet_NFE_EntradaNatureza: TStringField;
    CDataSet_NFE_EntradaNFeChave: TStringField;
    StatusBar2: TStatusBar;
    DBGrid_NFE_Entrada: TDBGrid;
    procedure FormCreate(Sender: TObject);
    procedure Certificados1Click(Sender: TObject);
    procedure BTN_Importar_CertificadoClick(Sender: TObject);
    procedure BTN_Remover_CertificadoClick(Sender: TObject);
    procedure BTN_Atualizar_ListaClick(Sender: TObject);
    procedure TreeView1Click(Sender: TObject);
    procedure Button_SelecionarClick(Sender: TObject);
    procedure Button_BuscarClick(Sender: TObject);
    procedure BuscarDistribuicao(Emp: TEmpresa);
    procedure ManifestarNFe(Emp: TEmpresa; Chave: string);
    procedure ConfigurarACBr(Emp: TEmpresa);
    function BaixarXMLCompleto(Emp: TEmpresa; const Chave: string; out Caminho: string): Boolean;
    function GetEmpresaSelecionada: TEmpresa;

  private
    FListaEmpresas: TObjectList<TEmpresa>;
    procedure CarregarTreeViewEmpresas;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form_Principal: TForm_Principal;

implementation

{$R *.dfm}

uses U_Certificados;

procedure TForm_Principal.FormCreate(Sender: TObject);
var
  Service: TEmpresaService;
  Emp: TEmpresa;
begin
  Service := TEmpresaService.Create;
  try
    Service.Carregar;
    FListaEmpresas := TObjectList<TEmpresa>.Create(True);
    // 🔥 cópia manual (correta)
    for Emp in Service.Lista do
      FListaEmpresas.Add(Emp);
  finally
    Service.Free;
  end;

  //ConfigurarACBr;

  StringGrid1.Cells[0,0] := 'Titular';
  StringGrid1.Cells[1,0] := 'email';
  StringGrid1.Cells[2,0] := 'CNPJ';
  StringGrid1.Cells[3,0] := 'Tipo';
  StringGrid1.Cells[4,0] := 'Vencimento';
  StringGrid1.Cells[5,0] := 'Status';
  StringGrid1.Cells[6,0] := 'Tempo Restante';
  StringGrid1.Cells[7,0] := 'Local';
  StringGrid1.Cells[8,0] := 'Ações';
  CarregarTreeViewEmpresas;
  Treeview1.FullCollapse;
  DateTimePicker1.DateTime := (Today-15);
  DateTimePicker2.DateTime := Today;
end;

procedure TForm_Principal.TreeView1Click(Sender: TObject);
var
  Node: TTreeNode;
  Emp: TEmpresa;
  Razao: string;
begin
  Node := TreeView1.Selected;

  if not Assigned(Node) then Exit;

  // sobe até o nó principal (empresa)
  while Assigned(Node.Parent) do
    Node := Node.Parent;

  if Assigned(Node.Data) then
  begin
    Emp := TEmpresa(Node.Data);
    //Emp := TEmpresa(TreeView1.Selected.Data);
    Razao := Emp.RazaoSocial;
    Edit_Pasta.Text := Emp.PastaXML;
    if ACBrNFe1.Configuracoes.Certificados.ArquivoPFX <> Emp.CertificadoPath then
      ConfigurarACBr(Emp);  end;
end;

procedure TForm_Principal.CarregarTreeViewEmpresas;
var
  Service: TEmpresaService;
  Lista: TObjectList<TEmpresa>;
  Emp: TEmpresa;
  I: Integer;
  NodeEmpresa, NodeCert, NodeNFe: TTreeNode;
  ultimaconsulta: string;
begin
  TreeView1.Items.Clear;

  Service := TEmpresaService.Create;
  try
    Service.Carregar;
    Lista := Service.Listar;

    for I := 0 to Lista.Count - 1 do
    begin
      Emp := Lista[I];
      //showmessage(emp.RazaoSocial + ' - ' + emp.PastaXML);

      //  NÍVEL 1 → Empresa
      NodeEmpresa := TreeView1.Items.Add(nil,
        Format('%.4d : %s', [I + 1, Emp.RazaoSocial]));

      NodeEmpresa.Data := Emp;

      //  CNPJ
      TreeView1.Items.AddChild(NodeEmpresa,
        'CNPJ: ' + Emp.CNPJ);

      //  CERTIFICADO A1
      NodeCert := TreeView1.Items.AddChild(NodeEmpresa,
        'CERTIFICADO A1');

      TreeView1.Items.AddChild(NodeCert,
        'data de início: ' + DateToStr(Emp.ValidadeInicio));

      TreeView1.Items.AddChild(NodeCert,
        'data de fim: ' + DateToStr(Emp.ValidadeFim));

      //  DADOS DE NFE
      NodeNFe := TreeView1.Items.AddChild(NodeEmpresa,
        'Dados de NFe');

      if Emp.UltimaConsulta > 0 then
        ultimaconsulta := DateTimeToStr(Emp.UltimaConsulta)
      else
        ultimaconsulta :=  'última procura: (nunca)'     ;

      TreeView1.Items.AddChild(NodeNFe,
        'última procura: ' + ultimaconsulta ) ;

      TreeView1.Items.AddChild(NodeNFe,
        'último nsu: ' + emp.UltimoNSU );

      TreeView1.Items.AddChild(NodeNFe,
        'pasta XML: ' + Emp.PastaXML);
    end;

  finally
    Service.Free;
  end;

  TreeView1.FullCollapse;
end;

procedure TForm_Principal.BTN_Atualizar_ListaClick(Sender: TObject);
var
  Service: TEmpresaService;
  Lista: TObjectList<TEmpresa>;
  Emp: TEmpresa;
  I: Integer;
  DiasRestantes: Integer;
  Status: string;
begin
  Service := TEmpresaService.Create;
  try
    Service.Carregar;
    Lista := Service.Listar;

    // limpa mantendo cabeçalho
    StringGrid1.RowCount := 1;

    for I := 0 to Lista.Count - 1 do
    begin
      Emp := Lista[I];

      StringGrid1.RowCount := StringGrid1.RowCount + 1;

      DiasRestantes := DaysBetween(Date, Emp.ValidadeFim);

      if Emp.ValidadeFim >= Date then
        Status := 'Ativo'
      else
        Status := 'Vencido';

      StringGrid1.Cells[0, I + 1] := Emp.RazaoSocial;
      StringGrid1.Cells[1, I + 1] := Emp.Email; // email
      StringGrid1.Cells[2, I + 1] := Emp.CNPJ;
      StringGrid1.Cells[3, I + 1] := Emp.Tipo;
      StringGrid1.Cells[4, I + 1] := DateToStr(Emp.ValidadeFim);
      StringGrid1.Cells[5, I + 1] := Status;
      StringGrid1.Cells[6, I + 1] := IntToStr(DiasRestantes) + ' dias';
      StringGrid1.Cells[7, I + 1] := '';
      StringGrid1.Cells[8, I + 1] := '';
    end;

  finally
    Service.Free;
  end;
end;

procedure TForm_Principal.BTN_Importar_CertificadoClick(Sender: TObject);
begin
    Form_Certificados.ShowModal;
    BTN_Atualizar_ListaClick(Sender);
    CarregarTreeViewEmpresas;
    Treeview1.FullCollapse;
end;

procedure TForm_Principal.BTN_Remover_CertificadoClick(Sender: TObject);
var
  Service: TEmpresaService;
  CNPJ, Senha: string;
  Empresa: TEmpresa;
  Linha: Integer;
begin
  Linha := StringGrid1.Row;

  if Linha <= 0 then
  begin
    ShowMessage('Selecione um certificado.');
    Exit;
  end;

  if not Assigned(FListaEmpresas) then
  begin
    ShowMessage('Lista de empresas não carregada.');
    Exit;
  end;

  CNPJ := StringGrid1.Cells[2, Linha];

  Empresa := nil;

  for var Emp in FListaEmpresas do
  begin
    if Emp.CNPJ = CNPJ then
    begin
      Empresa := Emp;
      Break;
    end;
  end;

  if not Assigned(Empresa) then
  begin
    ShowMessage('Certificado não encontrado.');
    Exit;
  end;

  Senha := InputBox('Confirmação', 'Digite a senha do certificado:', '');

  if Senha <> Empresa.CertificadoSenha then
  begin
    ShowMessage('Senha incorreta.');
    Exit;
  end;

  if MessageDlg('Deseja realmente remover este certificado?',
    mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    FListaEmpresas.Remove(Empresa);

    Service := TEmpresaService.Create;
    try
      Service.Salvar(FListaEmpresas);
    finally
      Service.Free;
    end;

    ShowMessage('Certificado removido com sucesso.');

    BTN_Atualizar_ListaClick(nil);
    CarregarTreeViewEmpresas;
    TreeView1.FullCollapse;
  end;
end;

procedure TForm_Principal.Button_BuscarClick(Sender: TObject);
var
  Emp: TEmpresa;
  Node: TTreeNode;
  Agora: TDateTime;
begin
  Node := TreeView1.Selected;

  while Assigned(Node.Parent) do
    Node := Node.Parent;

  if not Assigned(Node.Data) then Exit;

  Emp := TEmpresa(Node.Data);

  Agora := Now;

  // 🚨 REGRA DE 1 HORA
  if (Emp.UltimaConsulta > 0) and
     (MinutesBetween(Agora, Emp.UltimaConsulta) < 60) then
  begin
    ShowMessage('Aguarde pelo menos 1 hora para nova consulta. Última consulta: '+TimeToStr(Emp.UltimaConsulta));
    Exit;
  end;

  BuscarDistribuicao(Emp);
end;

procedure TForm_Principal.BuscarDistribuicao(Emp: TEmpresa);
var
  I: Integer;
  Chave: string;
  Caminho: string;
  Service: TEmpresaService;
begin
  // 🔒 Segurança básica
  if not Assigned(Emp) then Exit;

  // 🔹 Validação UF
  if Emp.UF <= 0 then
  begin
    ShowMessage('UF não configurada para a empresa.');
    Exit;
  end;

  if (Now - Emp.UltimaConsulta) < (1/24) then
  begin
    ShowMessage('Aguarde 1 hora para nova consulta. Última consulta feita: ' + TimeToStr(Emp.UltimaConsulta) );
    Exit;
  end;

  // 🔹 NSU inicial
  if Emp.UltimoNSU = '' then
    Emp.UltimoNSU := '000000000000000';

  // 🔹 Consulta por NSU
  ACBrNFe1.DistribuicaoDFePorUltNSU(Emp.UF, Emp.CNPJ, Emp.UltimoNSU);

  if ACBrNFe1.WebServices.DistribuicaoDFe.retDistDFeInt.docZip.Count = 0 then
  begin
      // 🔥 Atualiza NSU
      Emp.UltimoNSU :=
        ACBrNFe1.WebServices.DistribuicaoDFe.retDistDFeInt.ultNSU;

      // 🔥 Atualiza data da última consulta
      Emp.UltimaConsulta := Now;

      // 🔥 Salvar alterações no JSON
      Service := TEmpresaService.Create;
      try
        Service.Salvar(FListaEmpresas);
      finally
        Service.Free;
      end;

      StatusBar2.SimpleText := 'Consulta finalizada em ' + TimeToStr(Now);

     Showmessage('Nenhum novo documento encontrado!!!');
     exit;
  end;

  // 🔹 Percorre os documentos retornados
  for I := 0 to ACBrNFe1.WebServices.DistribuicaoDFe.retDistDFeInt.docZip.Count - 1 do
  begin
    with ACBrNFe1.WebServices.DistribuicaoDFe.retDistDFeInt.docZip.Items[I] do
    begin
      // 🔥 Compatível com qualquer versão (evita enum)
      if Pos('<resNFe', XML) > 0 then
      begin
        Chave := ResDFe.chDFe;

        Caminho := IncludeTrailingPathDelimiter(Emp.PastaXML) + Chave + '.xml';

        // 🔁 Evita duplicidade
        if FileExists(Caminho) then
          Continue;

        // 🔥 1. Manifestar
        ACBrNFe1.EventoNFe.Evento.Clear;

        with ACBrNFe1.EventoNFe.Evento.New do
        begin
          InfEvento.chNFe := Chave;
          InfEvento.CNPJ := Emp.CNPJ;
          InfEvento.tpEvento := teManifDestCiencia; // teCienciaOperacao;
          InfEvento.dhEvento := Now;
        end;

        ACBrNFe1.EnviarEvento(1);

        // 🔥 Aguarda SEFAZ liberar XML
        Sleep(2000);

        // 🔥 2. Baixar XML completo (função separada)
        if BaixarXMLCompleto(Emp, Chave, Caminho) then
        begin
          StatusBar2.SimpleText := 'XML baixado: ' + Chave;
        end;
      end;
    end;
  end;

  // 🔥 Atualiza NSU
  Emp.UltimoNSU :=
    ACBrNFe1.WebServices.DistribuicaoDFe.retDistDFeInt.ultNSU;

  // 🔥 Atualiza data da última consulta
  Emp.UltimaConsulta := Now;

  // 🔥 Salvar alterações no JSON
  Service := TEmpresaService.Create;
  try
    Service.Salvar(FListaEmpresas);
  finally
    Service.Free;
  end;

  StatusBar2.SimpleText := 'Consulta finalizada em ' + TimeToStr(Now);
end;

procedure TForm_Principal.ManifestarNFe(Emp: TEmpresa; Chave: string);
begin
  ACBrNFe1.EventoNFe.Evento.Clear;

  with ACBrNFe1.EventoNFe.Evento.New do
  begin
    InfEvento.chNFe := Chave;
    InfEvento.CNPJ := Emp.CNPJ;
    InfEvento.tpEvento := teManifDestCiencia; //teCienciaOperacao;
    InfEvento.dhEvento := Now;
  end;

  ACBrNFe1.EnviarEvento(1);
end;

function TForm_Principal.BaixarXMLCompleto(Emp: TEmpresa; const Chave: string; out Caminho: string): Boolean;
begin
  Result := False;

  // 🔒 Validação básica
  if not Assigned(Emp) then Exit;
  if Emp.UF <= 0 then Exit;

  // 🔥 Monta caminho final
  Caminho := IncludeTrailingPathDelimiter(Emp.PastaXML) + Chave + '.xml';

  // 🔁 Evita baixar novamente
  if FileExists(Caminho) then
  begin
    Result := True;
    Exit;
  end;

  // 🔹 Consulta por chave (AGORA COM UF + CNPJ)
  ACBrNFe1.DistribuicaoDFePorChaveNFe(Emp.UF, Emp.CNPJ, Chave);

  // 🔹 Se retornou XML completo
  if ACBrNFe1.NotasFiscais.Count > 0 then
  begin
    ForceDirectories(Emp.PastaXML);

    ACBrNFe1.NotasFiscais.Items[0].GravarXML(Caminho);

    Result := True;
  end;
end;

procedure TForm_Principal.Button_SelecionarClick(Sender: TObject);
var
  Dir: string;
  Node: TTreeNode;
  Emp: TEmpresa;
  Service: TEmpresaService;
  EmpTree, EmpLista: TEmpresa;
begin
  Dir := '';

  if not SelectDirectory('Selecione a pasta dos XMLs', '', Dir) then
    Exit;

  Node := TreeView1.Selected;

  if not Assigned(Node) then Exit;

  // 🔹 sobe até o nó raiz (empresa)
  while Assigned(Node.Parent) do
    Node := Node.Parent;

  if not Assigned(Node.Data) then Exit;

  //Emp := TEmpresa(Node.Data);

  // 🔥 Atualiza diretamente o objeto em memória
  //Emp.PastaXML := Dir;

  EmpTree := TEmpresa(Node.Data);

  EmpLista := nil;
  for var E in FListaEmpresas do
  begin
    if E.CNPJ = EmpTree.CNPJ then
    begin
      EmpLista := E;
      Break;
    end;
  end;
  if not Assigned(EmpLista) then
  begin
    ShowMessage('Empresa não encontrada na lista.');
    Exit;
  end;
  EmpLista.PastaXML := Dir;

  ShowMessage(EmpLista.PastaXML);

  // 🔹 Cria pasta se não existir
  if not DirectoryExists(Dir) then
    ForceDirectories(Dir);

  // 🔥 Persistência correta
  Service := TEmpresaService.Create;
  try
    Service.Salvar(FListaEmpresas);
  finally
    Service.Free;
  end;

  // 🔄 Atualiza UI
  Edit_Pasta.Text := Dir;

  // 🔄 Atualiza TreeView
  CarregarTreeViewEmpresas;

  ShowMessage('Pasta atualizada com sucesso.');
end;

procedure TForm_Principal.Certificados1Click(Sender: TObject);
begin
    if tabsheet3.TabVisible then begin
       tabsheet3.TabVisible := False;
       pagecontrol1.TabIndex :=0;
    end
    else begin
       tabsheet3.TabVisible := True;
       pagecontrol1.TabIndex := 2;
    end;
    BTN_Atualizar_ListaClick(Sender);

end;

function TForm_Principal.GetEmpresaSelecionada: TEmpresa;
var
  Node: TTreeNode;
begin
  Result := nil;

  Node := TreeView1.Selected;

  if not Assigned(Node) then Exit;

  // sobe até o nó raiz
  while Assigned(Node.Parent) do
    Node := Node.Parent;

  if Assigned(Node.Data) then
    Result := TEmpresa(Node.Data);
end;

procedure TForm_Principal.ConfigurarACBr(Emp: TEmpresa);
begin
  if not Assigned(Emp) then
  begin
    ShowMessage('Nenhuma empresa selecionada.');
    Exit;
  end;

  if not FileExists(Emp.CertificadoPath) then
  begin
    ShowMessage('Arquivo de certificado não encontrado.');
    Exit;
  end;

  SetDllDirectory(PChar(ExtractFilePath(Application.ExeName) + 'openssl'));

  with ACBrNFe1.Configuracoes do
  begin
    Geral.SSLLib := libOpenSSL;
    Geral.SSLCryptLib := cryOpenSSL;
    Geral.SSLHttpLib := httpOpenSSL;
    Geral.SSLXmlSignLib := xsLibXml2;

    Arquivos.PathSchemas :=
      ExtractFilePath(Application.ExeName) + 'Schemas\NFe\';

    // 🔥 DINÂMICO
    Certificados.ArquivoPFX := Emp.CertificadoPath;
    Certificados.Senha := Emp.CertificadoSenha;

    // 🔥 DINÂMICO
    WebServices.UF := CUFToUF(Emp.UF);
  end;

  // 🔐 Força carregar o certificado (importante!)
  try
    ACBrNFe1.SSL.CarregarCertificado;
  except
    on E: Exception do
    begin
      ShowMessage('Erro ao carregar certificado: ' + E.Message);
      Exit;
    end;
  end;
end;


end.
