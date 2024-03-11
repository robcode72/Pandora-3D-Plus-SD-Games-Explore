Unit uExcel;

Interface

Uses 
  System.Variants, System.Win.ComObj, System.SysUtils,
   Vcl.Graphics, Winapi.ActiveX, Vcl.Forms, Data.DB,
   Winapi.Windows;
  
Type

  TTiposFormatoCondicional = (xlCellValue = 1, xlExpression = 2);

  TTiposOperadoresFormatoCondicional = (xlBetween = 1, xlEqual = 3,
    xlGreater = 5, xlGreaterEqual = 7, xlLess = 6, xlLessEqual = 8,
    xlNotBetween = 2, xlNotEqual = 4, xlNone = 0);

  TTipoAlineacion = (xlCenter = -4108, xlDistributed = -4117, xlJustify = -4130,
    xlLeft = -4131, xlRight = -4152);

  TExcel = Class(TObject)
    Excel: variant;
    numsheets: integer;
    activesheet: integer;
  public
    Constructor Crear(Avisible: Boolean = true);
    Constructor Abrir(p_ruta: String; Avisible: Boolean = true);
    Destructor Destroy; override;
    //Procedimientos ordenados alfabeticamente para realizar merge de todas las
    //unidades UExcel.pas

    procedure AjustarAlturaFila(p_altura: integer);
    procedure AjustarAltura(p_altura: integer); //DEPRECATED
    procedure AjustarTamPagina(p_tam: integer);
    procedure AjustarTextoCeldaIzq;
    Procedure AlineacionHorizontalCelda(P_fil, P_col: Integer;
              alineacion: TTipoAlineacion);
    Procedure AlinearCelda (P_fila, P_columna: Integer; P_alineacion: String); overload;
    procedure AlinearCelda (p_fil,p_col: integer; iTipoAlineacion : integer = 0); overload;
    Procedure AlinearTexto(p_orientacion: String); //I izquierda, D derecha, C centro, centro por defecto.
    Procedure AlturaLineas(p_altura:Integer);
    Procedure AnchoColumna(p_col: integer; ancho: integer);
    Procedure AnchoFila(p_fil: integer; ancho: integer);
    procedure Anyade_Comentario_Celda(p_fil, p_col: integer; p_comentario: string);
    Procedure AutoFiltroSeleccion;
    Procedure AutoFit;
    procedure AutoFit_Vert;
    Procedure CambiarSheet(p_sheet: integer);
    Procedure CeldaNegrita(p_fil, p_col: integer);
    Procedure CeldaSubrayada(P_fil, P_col: Integer);
    Procedure CenterFooter(p_text: String);
    Procedure CenterHeader(p_text: String);
    Procedure CentrarCelda(p_fil, p_col: integer);
    Procedure CentrarCeldas;
    Procedure CentrarFila(p_fila: integer);
    Procedure CentrarPagina(p_modo: integer = 1);
    Procedure CentrarRango(p_rango: String);
    procedure Cierro;
    Procedure ColorFuenteCelda(P_fil, P_col: Integer; Color: string); overload;
    Procedure ColorFuenteCelda(p_fil, p_col: integer; indice: integer); overload;
    Procedure ColorFuenteSeleccion(indice: integer);
    procedure ColorRellenoCelda(P_fil, P_col: Integer; Color: string); overload;
    Procedure ColorRellenoCelda(p_fil, p_col: integer; indice: integer); overload;
    procedure ColorRellenoSeleccion(Color: string); overload;
    Procedure ColorRellenoSeleccion(indice: integer); overload;
    Procedure Color2RellenoSeleccion(color: integer);
    function  ContenidoCelda(p_fil,p_col:integer):string;
    Procedure CombinarSeleccion;
    Procedure CopyChart(p_grafic_ori, p_grafic_des, p_rangedes: String);
    procedure Cursiva;
    Procedure DelGraf(p_grafic: String);
    Function DevNumSheets: integer;
    Procedure DistanciaPagina(p_distancia: real = 0.78740157480315);
    Procedure FechaCelda(p_fil, p_col: integer; p_fecha: TDateTime);
    procedure FechaCeldaConFormato(p_fil, p_col: integer; p_fecha: TDateTime; p_formato: string);
    Procedure FechaFormatoCelda(p_fil, p_col: integer; p_text: string);
    Procedure FijarCabecera(p_rango: String); overload;
    Procedure FijarCabecera(p_fila: integer); overload;
    Procedure FijarCabecera2 (p_fila: integer); //DEPRECATED
    Procedure FijarCabeceraImpresion(p_fil, p_col: String);
    Procedure FijarPanel(P_fil, P_col: Integer); overload;
    Procedure FijarPanel; overload;
    Procedure FilaCursiva (p_fila:integer);
    Procedure FilaNegrita(p_fila: integer);
    Procedure FormatoCelda(p_fil, p_col: integer);
    Procedure FormatoNumeroArea(p_rango, p_formato : string);
    Procedure FormatoNumeroSeleccion;
    Procedure FormatoNumeroSeleccionCelda(P_fil, P_col: Integer);
    Procedure FormatoTextoSeleccion;
    procedure Guardar(p_fichero : string);
    procedure GuardarComoCSV(p_ruta:string);
    Procedure LeftFooter(p_text: String);
    Procedure LeftHeader(p_text: String);
    Procedure LetrasColor(P_fil, P_col, P_ini, P_lon: Integer; p_Color: string);
    Procedure LetrasNegrita(p_fil, p_col, p_ini, p_lon: integer);
    Procedure MargenInferior(p_margen: real = 1.5748031496063);
    Procedure MargenSuperior(p_margen: real = 1.96850393700787);
    Procedure MesCelda(p_fil, p_col: integer; p_text: String);
    Procedure Mostrar(p_mostrar: boolean);
    Procedure MostrarLineasDivision;
    Procedure Negrita;
    Procedure NewSheet;
    Procedure NombreSheet(p_nombre: String);
    Procedure NumeroCelda(p_fil, p_col: integer; p_text: real);
    Procedure PaginaApaisada;
    Procedure PaginaVertical;
    Procedure PaletaColPersonal(color: String; indice: integer);
    Procedure PaletaColPersonal2(color: TColor; indice: integer);
    Procedure Poner2BordesSeleccion(interior: boolean = true; ancho: integer = 1;  p_inthor: boolean = FALSE);
    procedure PonerBordeInferior;
    procedure PonerBordeIzquierdo;
    Procedure PonerBordeDer (P_fil, P_col: Integer);
    Procedure PonerBordeIzq (P_fil, P_col: Integer);
    Procedure PonerBordeSup (P_fil, P_col: Integer);
    Procedure PonerBordesCelda(p_fil, p_col: integer);
    Procedure PonerBordesSeleccion(interior: boolean = true; ancho: integer = 1);
    Procedure PonerFechaPiePagina;
    Procedure PonerFormatoCondicional(TipoCondicion: TTiposFormatoCondicional;
                 Condicion, Color: string;
                 operadorCondicion: TTiposOperadoresFormatoCondicional = xlNone;
                 Condicion2: string = '');
    Procedure PonerNegritaSeleccion;
    Procedure PonerImagen(img: string;p_left, p_top: integer);
    Procedure PonerPaginacion;
    procedure PonerTextoVerticalSeleccion;
    Procedure QuitarSeleccion;
    Procedure RepetirColumnasTitulo(p_rango: String = '$1:$1');
    Procedure RightFooter(p_text: String);
    Procedure RightHeader(p_text: String);
    Procedure SeleccionarCelda(p_fil, p_col: integer);
    Procedure SeleccionarFila(p_fila: integer);
    Procedure SeleccionarRango(p_rango: String);
    Procedure SeleccionarRangoNum(p_fil1, p_col1, p_fil2, p_col2: integer);
    Procedure SetSourceData(p_grafic, p_range, p_sheet: String);
    procedure Sombra(p_color: integer);
    procedure TamanoLetraCelda(p_fil, p_col, p_tamano: integer);
    Procedure TamanoLetraFila(p_fila, p_tamano: integer);
    Procedure TextoCelda(p_fil, p_col: integer; p_text: String);
    Procedure VariantCelda(p_fil, p_col: integer; p_valor: Variant);
    Procedure TextoHorasCelda(p_fil, p_col: integer; p_text: String);
    Procedure TextoMilesoCelda(p_fil, p_col: integer; p_text: real; const numdec: integer = 2);
    Procedure WrapColumna(p_col: integer);
    Procedure WrapTexto(col, width: integer);
    procedure ImpTabla(Excel : TExcel; ivrow : Integer; P_Tabla : TDataSet);

  End;




  //constante para transformar un rango de numeros en un rango con letras
Const c_Letras: Array[1..78] Of String = ('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U',
    'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM',
    'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
    'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ');



  K_COLIXAUTO = $FFFFEFF7;  // Índice de columna automático

  // ---- Alineación, Valores para compatibilidad con XE5
  K_XALCENTER = -4108;  // alineación centrada
  K_XALLEFT   = -4131;  // alineación izquierda
  K_XALRIGHT  = -4152;  // alineación derecha

  // -------Informativo, valores para versiones anteriores a XE5
  // K_XALCENTER = $FFFFEFF4;  // alineación centrada
  // K_XALLEFT   = $FFFFEFDD;  // alineación izquierda
  // K_XALRIGHT  = $FFFFEFC8;  // alineación derecha



Procedure ExportaExcel (P_Tabla: TDataSet);
Procedure ExportaExcelAutCursos(P_Tabla1, P_Tabla2, P_Tabla3, P_Tabla4: TDataSet);


var
  NomHoja : String; // Nombre de la hoja de Excel
  V_row : Integer;
  SheetListName: Array of string;
  //NomSheet : ArrayOfString;


Implementation

// Crea el excel con un libro y una hoja.

Constructor TExcel.Crear(Avisible: Boolean = true);
Begin
  Excel := UnAssigned;
  Excel := CreateOleObject('Excel.Application');
  Excel.Visible := Avisible;
  Excel.DisplayAlerts := false;
  Excel.WorkBooks.add(EmptyParam);
  Excel.Application.SheetsInNewWorkbook := 1;

  numsheets := 1;
  activesheet := 1;
End;

// Abre un excel con la ruta que se le pase

Constructor TExcel.Abrir(p_ruta: String; Avisible: Boolean = true);
Begin
  Excel := UnAssigned;
  Excel := CreateOleObject('Excel.Application');
  Excel.Visible := Avisible;
  Excel.DisplayAlerts := false;

  Excel.Workbooks.Add(p_ruta);
  Excel.Application.SheetsInNewWorkbook := 1;
  numsheets := 1;
  activesheet := 1;
End;

// ----------------------------------------------------------
// ---- Destruimos el objeto al acabar el trabajo Feb-2014
// ----------------------------------------------------------
Destructor TExcel.Destroy;
begin
  // Necesitamos liberar nuestra parte del contrato con el objeto OLE
  // Si no lo hacemos, se mantiene un contador de referencias a 1
  // y por tanto dejamos una instancia Excel huérfana en memoria
  try
  // Descomentar ésta línea para cerrar nuestra referencia a aplicación Excel
  // Excel.Quit;  // ordenar el cierre del objeto OLE (y la ventana)
    Excel := Unassigned; // en Delphi anterior a los Variants era := Nil;
 finally
   inherited Destroy;
  end;
end;

Procedure TExcel.PonerFechaPiePagina;
Begin
  PaginaApaisada;
  Excel.ActiveWorkBook.WorkSheets[activesheet].PageSetup.CenterFooter := DateToStr(date);
End;

procedure TExcel.PonerFormatoCondicional(TipoCondicion
  : TTiposFormatoCondicional; Condicion, Color: string;
  operadorCondicion: TTiposOperadoresFormatoCondicional = xlNone;
  Condicion2: string = '');
begin
  if operadorCondicion <> xlNone then
    Excel.Selection.FormatConditions.Add(TipoCondicion, operadorCondicion,
      Condicion, Condicion2)
  else
    Excel.Selection.FormatConditions.Add(TipoCondicion, null, Condicion,
      Condicion2);

  Excel.Selection.FormatConditions.Item(Excel.Selection.FormatConditions.Count)
    .SetFirstPriority;

  Excel.Selection.FormatConditions.Item(1).Interior.PatternColorIndex := -4105;
  Excel.Selection.FormatConditions.Item(1).Interior.Color :=
    RGB(StrToInt('$' + Copy(Color, 1, 2)), StrToInt('$' + Copy(Color, 3, 2)),
    StrToInt('$' + Copy(Color, 5, 2)));
  Excel.Selection.FormatConditions.Item(1).Interior.TintAndShade := 0;

  Excel.Selection.FormatConditions.Item(1).StopIfTrue := False;
end;

procedure TExcel.PonerNegritaSeleccion;
begin
  Excel.Selection.Font.Bold := true;
end;


Procedure TExcel.Mostrar(p_mostrar: boolean);
var
  hwnd : Integer;
Begin
  Excel.Visible := p_mostrar;
  Application.processMessages;

  if p_mostrar then
  begin
    hwnd := FindWindow('XLMAIN', nil);
    ShowWindow(hwnd, SW_MAXIMIZE); //SW_RESTORE);
    BringWindowToTop(hwnd);
  end;

End;

Procedure TExcel.NewSheet;
Begin
  Excel.ActiveWorkBook.Worksheets.add;
  inc(numsheets);
  activesheet := 1;
End;

Function TExcel.DevNumSheets: integer;
Begin
  DevNumSheets := numsheets;
End;

Procedure TExcel.CambiarSheet(p_sheet: integer);
Begin
  activesheet := p_sheet;
End;

Procedure TExcel.NombreSheet(p_nombre: String);
var
  aux: String;
Begin
//  // 01/04/2014, el nombre no puede superar 30 chars
//  If p_nombre <> '' Then
//    Excel.ActiveWorkBook.WorkSheets[activesheet].name := copy(p_nombre, 1, 30)
//  Else
//    Excel.ActiveWorkBook.WorkSheets[activesheet].name := ' ';

  If p_nombre <> '' Then
    aux := copy(p_nombre, 1, 30)
  Else
    aux  := ' ';

 // 12/05/2021  No se permiten caracteres especiales
 aux := StringReplace(aux, '*', '-', [rfReplaceAll]);
 aux := StringReplace(aux, ':', '-', [rfReplaceAll]);
 aux := StringReplace(aux, '/', '-', [rfReplaceAll]);
 aux := StringReplace(aux, '\', '-', [rfReplaceAll]);
 aux := StringReplace(aux, '[', '-', [rfReplaceAll]);
 aux := StringReplace(aux, ']', '-', [rfReplaceAll]);

 Excel.ActiveWorkBook.WorkSheets[activesheet].name := aux;

End;

Procedure TExcel.CentrarCeldas;
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells.HorizontalAlignment := K_XALCENTER; //xlCenter
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells.VerticalAlignment := K_XALCENTER; //xlCenter
End;

procedure TExcel.CentrarPagina(p_modo: integer = 1);
begin
  // 1 -->horizontal y vertical
  // 2 -->horizontal
  // 3 -->vertical
  if p_modo = 1 then
  begin
    Excel.ActiveWorkBook.Worksheets[activesheet]
      .PageSetup.CenterHorizontally := true;
    Excel.ActiveWorkBook.Worksheets[activesheet]
      .PageSetup.CenterVertically := true;
  end
  else if p_modo = 2 then
    Excel.ActiveWorkBook.Worksheets[activesheet]
      .PageSetup.CenterHorizontally := true
  else if p_modo = 3 then
    Excel.ActiveWorkBook.Worksheets[activesheet]
      .PageSetup.CenterVertically := true;
end;

Procedure TExcel.DistanciaPagina(p_distancia: real = 0.78740157480315);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].PageSetup.Application.InchesToPoints(p_distancia);
End;

Procedure TExcel.LeftHeader(p_text: String);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].PageSetup.LeftHeader := p_text;
End;

Procedure TExcel.LetrasColor(P_fil, P_col, P_ini, P_lon: Integer;
  p_Color: string);
Begin
  Excel.ActiveWorkBook.Worksheets[Activesheet].Cells[P_fil, P_col].Characters
    [P_ini, P_lon].Font.Color := RGB(StrToInt('$' + Copy(p_Color, 1, 2)),
    StrToInt('$' + Copy(p_Color, 3, 2)), StrToInt('$' + Copy(p_Color, 5, 2)));

End;

Procedure TExcel.CenterHeader(p_text: String);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].PageSetup.CenterHeader := p_text;
End;

Procedure TExcel.RightHeader(p_text: String);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].PageSetup.RightHeader := p_text;
End;

Procedure TExcel.LeftFooter(p_text: String);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].PageSetup.LeftFooter := p_text;
End;

Procedure TExcel.CenterFooter(p_text: String);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].PageSetup.CenterFooter := p_text;
End;

Procedure TExcel.RightFooter(p_text: String);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].PageSetup.RightFooter := p_text;
End;

Procedure TExcel.RepetirColumnasTitulo(p_rango: String = '$1:$1');
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].PageSetup.PrintTitleRows := p_rango;
End;

Procedure TExcel.MargenSuperior(p_margen: real = 1.96850393700787);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].PageSetup.TopMargin := Excel.Application.InchesToPoints(p_margen);
End;

Procedure TExcel.MesCelda(p_fil, p_col: integer; p_text: string);
Begin
  Excel.ActiveWorkBook.Worksheets[activesheet].Cells[p_fil, p_col].NumberFormat := 'mmm-aa';
  Excel.ActiveWorkBook.Worksheets[activesheet].Cells[p_fil, p_col].FormulaR1C1 := strtodate(p_text);
End;

Procedure TExcel.MargenInferior(p_margen: real = 1.5748031496063);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].PageSetup.BottomMargin := Excel.Application.InchesToPoints(p_margen);
End;

procedure TExcel.FilaCursiva (p_fila:integer);
begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].Rows[p_fila].Font.Italic := True;
end;

Procedure TExcel.FilaNegrita(p_fila: integer);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].Rows[p_fila].Font.Bold := True;
End;

Procedure TExcel.CeldaNegrita(p_fil, p_col: integer);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col].Font.Bold := True;
End;

Procedure TExcel.CeldaSubrayada(P_fil, P_col: Integer);
Begin
  Excel.ActiveWorkBook.Worksheets[Activesheet].Cells[P_fil, P_col]
    .Font.Underline := true;
End;

Procedure TExcel.LetrasNegrita(p_fil, p_col, p_ini, p_lon: integer);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col].Characters[p_ini, p_lon].Font.Bold := True;
End;

procedure TExcel.TamanoLetraCelda(p_fil, p_col, p_tamano: integer);
begin
  Excel.ActiveWorkBook.Worksheets[activesheet].Cells[p_fil, p_col].Font.Size := p_tamano;
end;

Procedure TExcel.TamanoLetraFila(p_fila, p_tamano: integer);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].Rows[p_fila].Font.Size := p_tamano;
End;

Procedure TExcel.TextoCelda(p_fil, p_col: integer; p_text: String);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col] := p_text;
End;

procedure TExcel.VariantCelda(p_fil, p_col: integer; p_valor: Variant);
begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col] := p_valor;
end;

Procedure TExcel.TextoHorasCelda(p_fil, p_col: integer; p_text: string);
Begin
  Excel.ActiveWorkBook.Worksheets[activesheet].Cells[p_fil, p_col].NumberFormat := 'h:mm';
  Excel.ActiveWorkBook.Worksheets[activesheet].Cells[p_fil, p_col].FormulaR1C1 := p_text;
End;

Procedure TExcel.TextoMilesoCelda(p_fil, p_col: integer; p_text: real; const numdec: integer = 2);
var
  sFormato, sAux: string;
  iLen: integer;
Begin
  //sFormato := '#.##0,';
  sFormato := '#' + FormatSettings.ThousandSeparator + '##0' +
    FormatSettings.DecimalSeparator;
  iLen := Length('0');
  if (iLen > numdec) then
    sAux := Copy('0', 1, numdec)
  else
    sAux := StringOfChar('0', numdec - iLen) + '0';
  sFormato := sFormato + sAux;
  Excel.ActiveWorkBook.Worksheets[activesheet].Cells[p_fil, p_col].NumberFormat
    //:= '#.##0,00';
    := sFormato;
  Excel.ActiveWorkBook.Worksheets[activesheet].Cells[p_fil, p_col].FormulaR1C1 := p_text;
End;

Procedure TExcel.SeleccionarCelda(p_fil, p_col: integer);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col].select;
End;

Procedure TExcel.AutoFit;
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells.Columns.AutoFit;
End;

Procedure TExcel.SeleccionarRango(p_rango: String);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].range[p_rango].select;
End;

procedure TExcel.sombra (p_color: integer);
begin
  Excel.Selection.interior.ColorIndex := p_color;
end;

Procedure TExcel.SeleccionarRangoNum(p_fil1, p_col1, p_fil2, p_col2: integer);
Var v_rango: String;
Begin
  v_rango := C_Letras[p_col1] + inttostr(p_fil1) + ':' + C_Letras[p_col2] + inttostr(p_fil2);
  Excel.ActiveWorkBook.WorkSheets[activesheet].range[v_rango].select;
End;

Procedure TExcel.CombinarSeleccion;
Begin
  Excel.selection.merge;
End;

Procedure TExcel.PonerBordesSeleccion(interior: boolean = true; ancho: integer = 1);
Begin
  // Borde izquierdo
  Excel.selection.borders[7].LineStyle := 1;
  Excel.selection.borders[7].Weight := ancho;
  Excel.selection.borders[7].ColorIndex := $FFFFEFF7;

  // Borde superior
  Excel.selection.borders[8].LineStyle := 1;
  Excel.selection.borders[8].Weight := ancho;
  Excel.selection.borders[8].ColorIndex := $FFFFEFF7;

  // Borde inferior
  Excel.selection.borders[9].LineStyle := 1;
  Excel.selection.borders[9].Weight := ancho;
  Excel.selection.borders[9].ColorIndex := $FFFFEFF7;

  // Borde derecho
  Excel.selection.borders[10].LineStyle := 1;
  Excel.selection.borders[10].Weight := ancho;
  Excel.selection.borders[10].ColorIndex := $FFFFEFF7;

  If interior Then Begin
    // Borde interior
    Excel.selection.borders[11].LineStyle := 1;
    Excel.selection.borders[11].Weight := ancho;
    Excel.selection.borders[11].ColorIndex := $FFFFEFF7;
  End;

End;

Procedure TExcel.Poner2BordesSeleccion(interior: boolean = true; ancho: integer = 1;
                                       p_inthor: boolean = FALSE);
Begin
  // Borde izquierdo
  Excel.selection.borders[7].LineStyle := 1;
  Excel.selection.borders[7].Weight := ancho;
  Excel.selection.borders[7].ColorIndex := K_COLIXAUTO;

  // Borde superior
  Excel.selection.borders[8].LineStyle := 1;
  Excel.selection.borders[8].Weight := ancho;
  Excel.selection.borders[8].ColorIndex := K_COLIXAUTO;

  // Borde inferior
  Excel.selection.borders[9].LineStyle := 1;
  Excel.selection.borders[9].Weight := ancho;
  Excel.selection.borders[9].ColorIndex := K_COLIXAUTO;

  // Borde derecho
  Excel.selection.borders[10].LineStyle := 1;
  Excel.selection.borders[10].Weight := ancho;
  Excel.selection.borders[10].ColorIndex := K_COLIXAUTO;

  If interior Then Begin
    // Borde interior
    Excel.selection.borders[11].LineStyle := 1;
    Excel.selection.borders[11].Weight := ancho;
    Excel.selection.borders[11].ColorIndex := K_COLIXAUTO;
  End;
  { Borde interior horizontal }
  if p_inthor then
  begin
    Excel.selection.borders[12].LineStyle := 1;
    Excel.selection.borders[12].Weight := ancho;
    Excel.selection.borders[12].ColorIndex := K_COLIXAUTO;
  end;
End;

Procedure TExcel.AutoFiltroSeleccion;
Begin
  Excel.selection.autofilter;
End;

procedure TExcel.PonerTextoVerticalSeleccion;
begin
  Excel.selection.Orientation := 90;
end;

Procedure TExcel.QuitarSeleccion;
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].range['A1'].select;
End;

Procedure TExcel.MostrarLineasDivision;
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].PageSetup.PrintGridlines := true;
End;

Procedure TExcel.PaginaApaisada;
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].PageSetup.Orientation := 2;
End;

Procedure TExcel.PaginaVertical;
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].PageSetup.Orientation := 1;
  ;
End;

Procedure TExcel.FormatoTextoSeleccion;
Begin
  Excel.selection.NumberFormat := '@';
  // Excel.selection.NumberFormat := '#,##0';
End;

procedure TExcel.Guardar(p_fichero : string);
begin
   Excel.ActiveWorkBook.SaveAs(p_fichero);
end;

procedure TExcel.GuardarComoCSV(p_ruta: string);
begin
  Excel.ActiveWorkbook.SaveAs(p_ruta,6);
  Excel.ActiveWorkbook.Close(true);
end;

Procedure TExcel.FormatoNumeroSeleccion;
Begin
  //Excel.selection.NumberFormat := '"General"';
  Excel.selection.NumberFormat := '0,00';
  //    Excel.selection.NumberFormat := '0';
     // Excel.selection.NumberFormat := '#,##0';
End;

Procedure TExcel.FormatoNumeroSeleccionCelda(P_fil, P_col: Integer);
Begin
  Excel.ActiveWorkBook.Worksheets[Activesheet].Cells[P_fil, P_col].NumberFormat
    := '0.00%';
End;

Procedure TExcel.SeleccionarFila(p_fila: integer);
Var
  v_fila: String;
Begin
  v_fila := IntToStr(p_fila) + ':' + IntToStr(p_fila);
  Excel.ActiveWorkBook.WorkSheets[activesheet].Rows[v_fila].Select;
End;

Procedure TExcel.PaletaColPersonal2(color: TColor; indice: integer);
Begin
  Excel.ActiveWorkbook.Colors(indice) := ColorToRgb(color);
End;

Procedure TExcel.PaletaColPersonal(color: String; indice: integer);
Begin
  Excel.ActiveWorkbook.Colors(indice) := RGB(strtoint('$' + Copy(color, 1, 2)),
    strtoint('$' + Copy(color, 3, 2)),
    strtoint('$' + Copy(color, 5, 2)));
End;

procedure TExcel.ColorRellenoSeleccion(Color: string);
begin
  Excel.Selection.Interior.Color := RGB(StrToInt('$' + Copy(Color, 1, 2)),
    StrToInt('$' + Copy(Color, 3, 2)), StrToInt('$' + Copy(Color, 5, 2)));
end;


Procedure TExcel.ColorRellenoSeleccion(indice: integer);
Begin
  Excel.Selection.Interior.ColorIndex := indice;
End;

function TExcel.ContenidoCelda(p_fil,p_col:integer):string;
begin
  result := Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil,p_col].value;
end;

Procedure TExcel.ColorRellenoCelda(p_fil, p_col: integer; indice: integer);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col].Interior.ColorIndex := indice;
End;

procedure TExcel.ColorRellenoCelda(P_fil, P_col: Integer; Color: string);
begin
  Excel.ActiveWorkBook.Worksheets[Activesheet].Cells[P_fil, P_col]
    .Interior.Color := RGB(StrToInt('$' + Copy(Color, 1, 2)),
    StrToInt('$' + Copy(Color, 3, 2)), StrToInt('$' + Copy(Color, 5, 2)));
end;

procedure TExcel.Cierro;
begin
  Excel.ActiveWorkBook.Close(0);
  // Cierro el excel, no unicamente el documento
  Destroy;
end;
Procedure TExcel.ColorFuenteCelda(P_fil, P_col: Integer; Color: string);
Begin
  Excel.ActiveWorkBook.Worksheets[Activesheet].Cells[P_fil, P_col].Font.Color :=
    RGB(StrToInt('$' + Copy(Color, 1, 2)), StrToInt('$' + Copy(Color, 3, 2)),
    StrToInt('$' + Copy(Color, 5, 2)));
End;


procedure TExcel.Color2RellenoSeleccion(color: integer);
begin
  Excel.Selection.Interior.Color := color;
end;

Procedure TExcel.ColorFuenteCelda(p_fil, p_col: integer; indice: integer);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col].Font.ColorIndex := indice;
End;

Procedure TExcel.ColorFuenteSeleccion(indice: integer);
Begin
  Excel.Selection.Font.ColorIndex := indice;
End;

procedure TExcel.PonerBordeIzquierdo;
begin
  Excel.selection.borders[7].LineStyle := 1;
  Excel.selection.borders[7].Weight := 2;
  Excel.selection.borders[7].ColorIndex := $FFFFEFF7;
end;

procedure TExcel.PonerBordeInferior;
begin
  Excel.selection.borders[9].LineStyle := 1;
  Excel.selection.borders[9].Weight := 2;
  Excel.selection.borders[9].ColorIndex := $FFFFEFF7;
end;

Procedure TExcel.PonerBordeSup (P_fil, P_col: Integer);
Begin
  Excel.ActiveWorkBook.WorkSheets[Activesheet].Cells[P_fil, P_col].Borders[8].LineStyle  := 1;
  Excel.ActiveWorkBook.WorkSheets[Activesheet].Cells[P_fil, P_col].Borders[8].Weight     := 2;
  Excel.ActiveWorkBook.WorkSheets[Activesheet].Cells[P_fil, P_col].Borders[8].ColorIndex :=
  K_COLIXAUTO;
end;

Procedure TExcel.PonerBordeDer (P_fil, P_col: Integer);
Begin
  Excel.ActiveWorkBook.WorkSheets[Activesheet].Cells[P_fil, P_col].Borders[10].LineStyle  := 1;
  Excel.ActiveWorkBook.WorkSheets[Activesheet].Cells[P_fil, P_col].Borders[10].Weight     := 2;
  Excel.ActiveWorkBook.WorkSheets[Activesheet].Cells[P_fil, P_col].Borders[10].ColorIndex :=
    K_COLIXAUTO;
end;

Procedure TExcel.PonerBordeIzq (P_fil, P_col: Integer);
Begin
  Excel.ActiveWorkBook.WorkSheets[Activesheet].Cells[P_fil, P_col].Borders[7].LineStyle  := 1;
  Excel.ActiveWorkBook.WorkSheets[Activesheet].Cells[P_fil, P_col].Borders[7].Weight     := 2;
  Excel.ActiveWorkBook.WorkSheets[Activesheet].Cells[P_fil, P_col].Borders[7].ColorIndex :=
    K_COLIXAUTO;
end;

Procedure TExcel.PonerBordesCelda(p_fil, p_col: integer);
Begin
  // Borde izquierdo
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col].borders[7].LineStyle := 1;
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col].borders[7].Weight := 2;
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col].borders[7].ColorIndex := $FFFFEFF7;

  // Borde superior
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col].borders[8].LineStyle := 1;
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col].borders[8].Weight := 2;
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col].borders[8].ColorIndex := $FFFFEFF7;

  // Borde inferior
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col].borders[9].LineStyle := 1;
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col].borders[9].Weight := 2;
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col].borders[9].ColorIndex := $FFFFEFF7;

  // Borde derecho
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col].borders[10].LineStyle := 1;
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col].borders[10].Weight := 2;
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col].borders[10].ColorIndex := $FFFFEFF7;
End;

Procedure TExcel.FijarCabeceraImpresion(p_fil, p_col: String);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].PageSetup.PrintTitleRows := '$' + p_fil + ':$' + p_col;
End;

Procedure TExcel.FijarCabecera(p_rango: String);
Begin
  SeleccionarRango(p_rango);
  Excel.ActiveWindow.FreezePanes := True;
End;

Procedure TExcel.FijarCabecera (p_fila: integer);
Begin
  SeleccionarRango ('A' + IntTostr(p_fila));
  Excel.ActiveWindow.SplitColumn := 0;
  Excel.ActiveWindow.SplitRow := p_fila - 1;
  Excel.ActiveWindow.FreezePanes := True;
End;

Procedure TExcel.FijarCabecera2 (p_fila: integer);
Begin
  FijarCabecera(p_fila);
End;


procedure TExcel.FijarPanel(P_fil, P_col: Integer);
begin
  Excel.Activesheet.Cells[P_fil, P_col].Select;
  Excel.ActiveWindow.FreezePanes := true;
end;

Procedure TExcel.FijarPanel;
Begin
  Excel.Activesheet.Range['A3', 'A3'].Select;
  Excel.ActiveWindow.FreezePanes := true; // Congela la parte superior
End;

Procedure TExcel.AnchoColumna(p_col: integer; ancho: integer);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].Columns[p_col].ColumnWidth := ancho;
End;

Procedure TExcel.WrapColumna(p_col: integer);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].Columns[p_col].WrapText := True;
End;

Procedure TExcel.CentrarCelda(p_fil, p_col: integer);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col].HorizontalAlignment := K_XALCENTER; //xlCenter
End;

procedure TExcel.AjustarAlturaFila(p_altura: integer);
begin
  Excel.selection.RowHeight := p_altura;
end;

procedure TExcel.AjustarAltura(p_altura: integer);
begin
  AjustarAlturaFila(p_altura);
end;

procedure TExcel.AjustarTamPagina(p_tam: integer);
begin
  Excel.ActiveWorkBook.Worksheets[activesheet].PageSetup.zoom := p_tam;
end;

procedure TExcel.AjustarTextoCeldaIzq;
begin
  Excel.Selection.HorizontalAlignment := K_XALLEFT;
end;

Procedure TExcel.AlinearCelda (P_fila, P_columna: Integer; P_alineacion: String);
Begin
  If P_alineacion = 'R' Then
    Excel.WorkBooks[1].WorkSheets[Activesheet].Cells[P_fila, P_columna].HorizontalAlignment
      := K_XALRIGHT
  Else If P_alineacion = 'L' Then
    Excel.WorkBooks[1].WorkSheets[Activesheet].Cells[P_fila, P_columna].HorizontalAlignment
      := K_XALLEFT
  Else
    Excel.WorkBooks[1].WorkSheets[Activesheet].Cells[P_fila, P_columna].HorizontalAlignment
      := K_XALCENTER;
End;

procedure TExcel.AlinearCelda(p_fil,p_col: integer; iTipoAlineacion : integer = 0);
begin
  case iTipoAlineacion of
    //--------------------------------------------------------------------------
    //                          alineacion izquierda
    //--------------------------------------------------------------------------
    0: Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil,p_col].HorizontalAlignment := K_XALLEFT ;
    //--------------------------------------------------------------------------
    //                          alineacion centrada
    //--------------------------------------------------------------------------
    1: Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil,p_col].HorizontalAlignment := K_XALCENTER;
    //--------------------------------------------------------------------------
    //                          alineacion derecha
    //--------------------------------------------------------------------------
    2: Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil,p_col].HorizontalAlignment := K_XALRIGHT;
  end;
end;

procedure TExcel.AlineacionHorizontalCelda(P_fil, P_col: Integer;
  alineacion: TTipoAlineacion);
begin
  Excel.ActiveWorkBook.Worksheets[Activesheet].Cells[P_fil, P_col]
    .HorizontalAlignment := alineacion;
end;

Procedure TExcel.AnchoFila(p_fil: integer; ancho: integer);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].Rows[p_fil].RowHeight := ancho;
End;

Procedure TExcel.FormatoCelda(p_fil, p_col: integer);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col].NumberFormat := 'Texto';
End;

Procedure TExcel.FormatoNumeroArea(p_rango, p_formato : string);
begin
  Excel.ActiveWorkBook.WorkSheets[Activesheet].Range[P_rango].NumberFormat := p_formato;
end;

Procedure TExcel.NumeroCelda(p_fil, p_col: integer; p_text: real);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col].FormulaR1C1 := p_text;
End;

Procedure TExcel.FechaCelda(p_fil, p_col: integer; p_fecha: TDateTime);
Begin
  If p_fecha <> 0 Then
    Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col].FormulaR1C1 := p_fecha;
End;

procedure TExcel.FechaCeldaConFormato(p_fil, p_col: integer; p_fecha: TDateTime;
                                      p_formato: string);
begin
  if p_fecha <> 0 then
  begin
    try
      Excel.ActiveWorkBook.Worksheets[activesheet].Cells[p_fil, p_col].NumberFormat
        := p_formato;
      Excel.ActiveWorkBook.Worksheets[activesheet].Cells[p_fil, p_col].FormulaR1C1
        := p_fecha;
    except
      on E: Exception do
        // Si se trata de una fecha incorrecta, tal y como 20/10/003
        if pos('OLE error 800A03EC', E.Message) > 0 then
          Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col] :=
            FormatDateTime('dd/mm/yyyy', p_fecha)
        else
          Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col]
            := p_fecha;
    end;
  end;
end;

Procedure TExcel.FechaFormatoCelda(p_fil, p_col: integer; p_text: string);
Begin
  Excel.ActiveWorkBook.Worksheets[activesheet].Cells[p_fil, p_col].NumberFormat := 'dd-mm-aa';
  Excel.ActiveWorkBook.Worksheets[activesheet].Cells[p_fil, p_col].FormulaR1C1 := strtodate(p_text);
End;

Procedure TExcel.CentrarRango(p_rango: String);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].range[p_rango].HorizontalAlignment := K_XALCENTER; //xlCenter
  Excel.ActiveWorkBook.WorkSheets[activesheet].range[p_rango].VerticalAlignment := K_XALCENTER; //xlCenter

End;

Procedure TExcel.CentrarFila(p_fila: integer);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].Rows[p_fila].HorizontalAlignment := K_XALCENTER; //xlCenter
  Excel.ActiveWorkBook.WorkSheets[activesheet].Rows[p_fila].VerticalAlignment := K_XALCENTER; //xlCenter
End;

Procedure TExcel.SetSourceData(p_grafic, p_range, p_sheet: String);
Begin
  Excel.ActiveSheet.ChartObjects(p_grafic).Activate;
  Excel.ActiveChart.SetSourceData(Source := Excel.ActiveWorkBook.WorkSheets[p_sheet].Range[p_range]); //,PlotBy:=3);
End;

Procedure TExcel.CopyChart(p_grafic_ori, p_grafic_des, p_rangedes: String);
Begin
  Excel.ActiveSheet.ChartObjects(1).Activate;
  Excel.ActiveChart.ChartArea.Select;
  Excel.ActiveChart.ChartArea.Copy;
  Excel.ActiveWorkBook.WorkSheets[activesheet].range[p_rangedes].select;
  Excel.ActiveSheet.Paste;
  Excel.ActiveChart.ChartTitle.Text := p_grafic_des;
  Excel.ActiveSheet.ChartObjects(Excel.ActiveSheet.ChartObjects.Count).name := p_grafic_des;
  Excel.ActiveSheet.ChartObjects(Excel.ActiveSheet.ChartObjects.Count).Activate;
  Excel.ActiveChart.ChartArea.Select;
  Excel.ActiveChart.ChartArea.ClearContents;
End;

procedure TExcel.Cursiva();
begin
  Excel.selection.Font.Italic := true;
end;

procedure TExcel.Negrita();
begin
  Excel.selection.Font.Bold := true;
end;


Procedure TExcel.DelGraf(p_grafic: String);
Begin
  Excel.ActiveSheet.ChartObjects(p_grafic).Activate;
  Excel.ActiveChart.Parent.Delete;
End;

procedure TExcel.AlturaLineas(p_altura: Integer);
begin
  Excel.Cells.Select;
  Excel.Selection.RowHeight := 9.75;
end;

Procedure TExcel.AlinearTexto(p_orientacion: String);
Begin
  if p_orientacion = 'I' then
    Excel.Selection.HorizontalAlignment := K_XALLEFT
  else if p_orientacion = 'D' then
    Excel.Selection.HorizontalAlignment := K_XALRIGHT
  else if p_orientacion = 'C' then
    Excel.Selection.HorizontalAlignment := K_XALCENTER
  else
    Excel.Selection.HorizontalAlignment := K_XALCENTER;
End;

Procedure TExcel.PonerImagen(img: string;p_left, p_top: integer);
Begin
  //Excel.ActiveWorkBook.WorkSheets[activesheet].Pictures.Insert(img).select;
  excel.ActiveWorkBook.WorkSheets[activesheet].shapes.addpicture(img,false, true, p_left, p_top,-1, -1);
End;

procedure TExcel.WrapTexto(col, width: integer);
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].Columns[col].ColumnWidth
    := width;
  Excel.ActiveWorkBook.WorkSheets[activesheet].Columns[col].WrapText := true;
End;

procedure TExcel.PonerPaginacion;
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].PageSetup.CenterFooter :=
    'Página &P';
  Excel.ActiveWorkBook.WorkSheets[activesheet].PageSetup.Zoom := 100;
  Excel.ActiveWorkBook.WorkSheets[activesheet].PageSetup.FooterMargin :=
    Excel.Application.InchesToPoints(0);
  Excel.ActiveWorkBook.WorkSheets[activesheet].PageSetup.BottomMargin :=
    Excel.Application.InchesToPoints(1);
End;

procedure TExcel.Anyade_Comentario_Celda(p_fil, p_col: integer;
                                         p_comentario: string);
begin
  { Añadir un comentario a una celda }
  { Modificado y adaptado a partir del código original de Luismi }
  //Excel.ActiveWorkBook.WorkSheets[ActiveSheet].Range[p_rango].Select;
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells[p_fil, p_col].Select;
  Excel.Selection.AddComment(p_comentario);
  Excel.Selection.Comment.Visible := FALSE;
end;

Procedure TExcel.AutoFit_Vert;
Begin
  Excel.ActiveWorkBook.WorkSheets[activesheet].Cells.Rows.AutoFit;
End;

// ******************************************************************************
//Exporta una tabla a Excel
//Utiliza la información de los TFields definidos para saber si exporta
//un determinado campo y el titulo que le pone a la columna.
// ******************************************************************************
Procedure ExportaExcel (P_Tabla: TDataSet);
Var
  Excel:      TExcel;
  Ind, iField : Integer;
  hwnd : Integer;


Begin
  Excel := TExcel.Crear;
  Excel.Mostrar (False);
  Excel.CambiarSheet (1);

  V_row := 2;
  Excel.Filanegrita (V_row);
  Excel.CentrarFila (V_row);

  // COlumnas
  For Ind := 0 To P_Tabla.FieldCount - 1 Do
  Begin
    // No mostrar Columnas TField si visible es false
    if P_Tabla.Fields[Ind].Visible then
      iField := iField + 1;
    Excel.TextoCelda (V_row, Ind + 1, P_Tabla.Fields[Ind].DisplayLabel);
  End;

  P_Tabla.DisableControls;
  P_Tabla.First;

    // Pulsar Esc salir
    While Not P_Tabla.Eof Do
    Begin
      Application.ProcessMessages;
      If GetKeyState (VK_Escape) And 128 = 128 Then
      Begin
        If Application.MessageBox ('¿Desea interrumpir el proceso?', 'Listado',
          MB_OKCANCEL + MB_ICONINFORMATION + MB_DEFBUTTON2 + MB_TOPMOST)
          = IDOK Then
        Begin
          Excel.Mostrar (True);
          Break;
        End;
      End;

      V_row := V_row + 1;

      Excel.TamanoLetraFila (V_row, 12);
      // Imprimir rows
      For Ind := 0 To P_Tabla.FieldCount - 1 Do
      Begin
        // Imprimir row si es visible= true
        //if P_Tabla.Fields[Ind].Visible then
        //begin
          If P_Tabla.Fields[Ind].DataType = FtString Then
          Begin
            If Not P_Tabla.Fields[Ind].IsNull Then
              Excel.TextoCelda (V_row, Ind + 1, P_Tabla.Fields[Ind].AsString)
            Else
              Excel.TextoCelda (V_row, Ind + 1, '');
          End
          Else
          If P_Tabla.Fields[Ind].DataType = ftWideMemo Then
          Begin
            {
            If Not P_Tabla.Fields[Ind].IsNull Then
              Excel.TextoCelda (V_row, Ind + 1, P_Tabla.Fields[Ind].AsWideString)
            Else
            }
              Excel.TextoCelda (V_row, Ind + 1, '');
          End
          else
          Begin
            If (P_Tabla.Fields[Ind].DataType = FtDateTime) Or
              (P_Tabla.Fields[Ind].DataType = FtDate) Then
            Begin
              If Not P_Tabla.Fields[Ind].IsNull Then
                Excel.FechaCelda (V_row, Ind + 1, P_Tabla.Fields[Ind].AsDateTime)
              Else
                Excel.TextoCelda (V_row, Ind + 1, '');
            End
            Else
            Begin
              If Not P_Tabla.Fields[Ind].IsNull Then
                Excel.NumeroCelda (V_row, Ind + 1, P_Tabla.Fields[Ind].AsFloat)
              Else
                Excel.TextoCelda (V_row, Ind + 1, '');
            End;
          End;
        //End;
      End;
      P_Tabla.Next;

    End;
  P_Tabla.First;
  P_Tabla.EnableControls;
  Excel.NombreSheet (NomHoja);
  Excel.AutoFit;
  Excel.Mostrar (True);
  Excel.Free;
  hwnd := FindWindow('XLMAIN', nil);
  ShowWindow(hwnd, SW_MAXIMIZE); //SW_RESTORE);
  BringWindowToTop(hwnd);


End;


Procedure ExportaExcelAutCursos(P_Tabla1, P_Tabla2, P_Tabla3, P_Tabla4: TDataSet );
Var
  Excel:      TExcel;
  Ind, iCol : Integer;

Begin
  Excel := TExcel.Crear;
  Excel.Mostrar (False);
  Excel.CambiarSheet (1);
  //Excel.NombreSheet ('Periodos no asignados');
  Excel.NombreSheet(SheetListName[0]);

  V_row := 2;
  Excel.Filanegrita (V_row);
  Excel.CentrarFila (V_row);
  Excel.ImpTabla(Excel, V_row, P_Tabla1);

  P_Tabla1.First;
  P_Tabla1.EnableControls;


  Excel.NewSheet;
  Excel.NombreSheet (NomHoja);
  V_row := 2;
  Excel.ImpTabla(Excel, V_row, P_Tabla2);
  V_row := V_row + 2;
  Excel.ImpTabla(Excel, V_row, P_Tabla3);
  V_row := V_row + 4;
  Excel.ImpTabla(Excel, V_row, P_Tabla4);
  Excel.NombreSheet(SheetListName[1]);

  Excel.AutoFit;
  Excel.Mostrar (True);
  Excel.Free;
End;

procedure TExcel.ImpTabla(Excel : TExcel; ivrow : Integer; P_Tabla : TDataSet);
var
  Ind : Integer;
begin
  P_Tabla.DisableControls;
  P_Tabla.First;
  Excel.Filanegrita(ivrow);
  Excel.CentrarFila (ivrow);

  // COlumnas Tabla 1
  For Ind := 0 To P_Tabla.FieldCount - 1 Do
  Begin
    Excel.TextoCelda (ivrow, Ind + 1, P_Tabla.Fields[Ind].DisplayLabel);
  End;
    While Not P_Tabla.Eof Do
    Begin
      Application.ProcessMessages;
      If GetKeyState (VK_Escape) And 128 = 128 Then
      Begin
        If Application.MessageBox ('¿Desea interrumpir el proceso?', 'Listado',
          MB_OKCANCEL + MB_ICONINFORMATION + MB_DEFBUTTON2 + MB_TOPMOST)
          = IDOK Then
        Begin
          Excel.Mostrar (True);
          Break;
        End;
      End;

      ivrow := ivrow + 1;
      Excel.TamanoLetraFila (ivrow, 12);

      // Imprimir rows
      For Ind := 0 To P_Tabla.FieldCount - 1 Do
      Begin
        // Imprimir row si es visible= true
          If P_Tabla.Fields[Ind].DataType = FtString Then
          Begin
            If Not P_Tabla.Fields[Ind].IsNull Then
              Excel.TextoCelda (ivrow, Ind + 1, P_Tabla.Fields[Ind].AsString)
            Else
              Excel.TextoCelda (ivrow, Ind + 1, '');
          End
          Else
          Begin
            If (P_Tabla.Fields[Ind].DataType = FtDateTime) Or
              (P_Tabla.Fields[Ind].DataType = FtDate) Then
            Begin
              If Not P_Tabla.Fields[Ind].IsNull Then
                Excel.FechaCelda (ivrow, Ind + 1, P_Tabla.Fields[Ind].AsDateTime)
              Else
                Excel.TextoCelda (ivrow, Ind + 1, '');
            End
            Else
            Begin
              If Not P_Tabla.Fields[Ind].IsNull Then
                Excel.NumeroCelda (ivrow, Ind + 1, P_Tabla.Fields[Ind].AsFloat)
              Else
                Excel.TextoCelda (ivrow, Ind + 1, '');
            End;
          End;
      End;
      P_Tabla.Next;
    End;
  P_Tabla.First;
  P_Tabla.EnableControls;
end;

End.
