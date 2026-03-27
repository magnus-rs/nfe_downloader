unit uEmpresaService;

interface

uses
  System.SysUtils,
  System.Classes,
  System.Generics.Collections,
  System.JSON,
  System.IOUtils,
  System.DateUtils,
  uEmpresa;

type
  TEmpresaService = class
  private
    FLista: TObjectList<TEmpresa>;
    FArquivo: string;

    function EmpresaToJSON(Emp: TEmpresa): TJSONObject;
    function JSONToEmpresa(Obj: TJSONObject): TEmpresa;

  public
    constructor Create(const AArquivo: string = 'empresas.json');
    destructor Destroy; override;

    procedure Adicionar(Emp: TEmpresa);
    procedure Remover(CNPJ: string);
    function Buscar(CNPJ: string): TEmpresa;
    function Listar: TObjectList<TEmpresa>;

    procedure Salvar;
    procedure Carregar;
  end;

implementation

{ TEmpresaService }

constructor TEmpresaService.Create(const AArquivo: string);
begin
  FLista := TObjectList<TEmpresa>.Create(True); // True = gerencia memória
  FArquivo := AArquivo;
end;

destructor TEmpresaService.Destroy;
begin
  FLista.Free;
  inherited;
end;

procedure TEmpresaService.Adicionar(Emp: TEmpresa);
begin
  if Assigned(Buscar(Emp.CNPJ)) then
    raise Exception.Create('Empresa já cadastrada.');

  FLista.Add(Emp);
end;

procedure TEmpresaService.Remover(CNPJ: string);
var
  Emp: TEmpresa;
begin
  Emp := Buscar(CNPJ);
  if Assigned(Emp) then
    FLista.Remove(Emp);
end;

function TEmpresaService.Buscar(CNPJ: string): TEmpresa;
var
  Emp: TEmpresa;
begin
  Result := nil;

  for Emp in FLista do
  begin
    if Emp.CNPJ = CNPJ then
      Exit(Emp);
  end;
end;

function TEmpresaService.Listar: TObjectList<TEmpresa>;
begin
  Result := FLista;
end;

procedure TEmpresaService.Salvar;
var
  JSONArray: TJSONArray;
  Emp: TEmpresa;
begin
  JSONArray := TJSONArray.Create;
  try
    for Emp in FLista do
      JSONArray.AddElement(EmpresaToJSON(Emp));

    TFile.WriteAllText(FArquivo, JSONArray.ToString);
  finally
    JSONArray.Free;
  end;
end;

procedure TEmpresaService.Carregar;
var
  JSONStr: string;
  JSONArray: TJSONArray;
  Obj: TJSONObject;
  I: Integer;
begin
  FLista.Clear;

  if not FileExists(FArquivo) then
    Exit;

  JSONStr := TFile.ReadAllText(FArquivo);
  JSONArray := TJSONObject.ParseJSONValue(JSONStr) as TJSONArray;

  try
    for I := 0 to JSONArray.Count - 1 do
    begin
      Obj := JSONArray.Items[I] as TJSONObject;
      FLista.Add(JSONToEmpresa(Obj));
    end;
  finally
    JSONArray.Free;
  end;
end;

function TEmpresaService.EmpresaToJSON(Emp: TEmpresa): TJSONObject;
begin
  Result := TJSONObject.Create;
  Result.AddPair('CNPJ', Emp.CNPJ);
  Result.AddPair('RazaoSocial', Emp.RazaoSocial);
  Result.AddPair('NomeFantasia', Emp.NomeFantasia);
  Result.AddPair('CertificadoPath', Emp.CertificadoPath);
  Result.AddPair('CertificadoSenha', Emp.CertificadoSenha);
  Result.AddPair('ValidadeInicio', DateToISO8601(Emp.ValidadeInicio));
  Result.AddPair('ValidadeFim', DateToISO8601(Emp.ValidadeFim));
end;

function TEmpresaService.JSONToEmpresa(Obj: TJSONObject): TEmpresa;
var
  StrData: string;
begin
  Result := TEmpresa.Create;

  Result.CNPJ := Obj.GetValue<string>('CNPJ');
  Result.RazaoSocial := Obj.GetValue<string>('RazaoSocial');
  Result.NomeFantasia := Obj.GetValue<string>('NomeFantasia');
  Result.CertificadoPath := Obj.GetValue<string>('CertificadoPath');
  Result.CertificadoSenha := Obj.GetValue<string>('CertificadoSenha');
  //Result.ValidadeInicio := ISO8601ToDate(Obj.GetValue<string>('ValidadeInicio'));

  StrData := Obj.GetValue<string>('ValidadeInicio', '');
  if StrData <> '' then
    Result.ValidadeInicio := ISO8601ToDate(StrData);

  //Result.ValidadeFim := ISO8601ToDate(Obj.GetValue<string>('ValidadeFim'));

  StrData := Obj.GetValue<string>('ValidadeFim', '');
  if StrData <> '' then
    Result.ValidadeFim := ISO8601ToDate(StrData);

end;

end.
