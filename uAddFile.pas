unit uAddFile;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Buttons,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, Data.DB, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client,
  // Imported
  System.IOUtils;

type
  Toper = (opNew, opEdit, opDelete);
  TGame = record
    oper    : Toper;
    SD_DRIVE : String;
    // -- info --
    CREATE_ID  : Integer; //Auto Generated (Shows games in this order)
    ROM_NAME : string; //name of the rom file
    DEFAULT_DISPLAY_NAME  : String; // Name of the game
    EMULATOR : Integer; //The emulator use to run the game
    GENRE : Integer;  //used to searching by category
    LOAD_TIME : Integer;  // 0-50000
    COIN_DURATION : Integer; // 0-250
    DEFAULT_CONFIG : String; //(NULL,4_.zip,4_2.zip,4_3.zip,_.zip,_2.zip,_3.zip)
    SUPPORTED_FEATURES : Integer; // (null,4[exit only],5,20[exit only],21,23[exit only],28[savestate],29,30,31[savestate])
    COIN_INSERT_MODE : Integer; // Default 0, ¿10000?
    FLAG1 : String; // (NULL, 0,1[n64])
    FLAG2 : String; // (NULL)
    FLAG3 : String; // (NULL)  Real field name EXTERNAL
    EXTERNAL_FILE : Integer; // (0,1)
    FILE_NAME : String; // Real file name + path (INTERNO ONLY)

    procedure Init;
    function Title: String;
  end;

  procedure ShowForm(Data : TGame);



type
  TfrmAddFile = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    edtDEFAULT_DISPLAY_NAME: TLabeledEdit;
    cbEMULATOR: TComboBox;
    Label1: TLabel;
    edtCOIN_DURATION: TLabeledEdit;
    cbGENRE: TComboBox;
    lblGenre: TLabel;
    edtLoadTime: TLabeledEdit;
    edtROM_NAME: TLabeledEdit;
    cbDEFAULT_CONFIG: TComboBox;
    Label2: TLabel;
    cbSUPPORTED_FEATURES: TComboBox;
    Label3: TLabel;
    edtCOIN_INSERT_MODE: TLabeledEdit;
    cbFLAG1: TComboBox;
    Label4: TLabel;
    cbFLAG2: TComboBox;
    Label5: TLabel;
    cbFLAG3: TComboBox;
    Label6: TLabel;
    chkENTERNO: TCheckBox;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    OpenDialog: TOpenDialog;
    qGames_Insert: TFDQuery;
    qGames_Update: TFDQuery;
    GroupBox1: TGroupBox;
    edtVideoFile: TLabeledEdit;
    SpeedButton4: TSpeedButton;
    SpeedButton3: TSpeedButton;
    edtFILE_NAME: TLabeledEdit;
    qGet_Rec_Num: TFDQuery;
    qGet_Rec_NumCREATE_ID: TIntegerField;
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
  private
    { Private declarations }
    Game : TGame;
    procedure New(Data : TGame);
    procedure Edit(Data : TGame);
    function FindRomPath(emu : Integer) : String;
    procedure FillComboBox;
  public
    { Public declarations }

  end;

var
  frmAddFile: TfrmAddFile;
  const GAMESDIR : String = '\games\data\';

implementation

{$R *.dfm}

uses uUtils, uMain;

var
  Emu : Array[0..17] of TEmuladores;

procedure ShowForm(Data : TGame);
begin
  //--
  if Data.oper = opNew then
    frmAddFile.New(Data);
  if Data.oper = opEdit then
  begin
    frmAddFile.Edit(Data);
  end;

end;

procedure TGame.Init;
begin
    CREATE_ID := 0;
    ROM_NAME := '';
    DEFAULT_DISPLAY_NAME := '';
    FILE_NAME := '';
    EMULATOR := 0;
    GENRE := 0;
    LOAD_TIME := 0;
    COIN_DURATION := 0;
    DEFAULT_CONFIG :='';
    SUPPORTED_FEATURES :=0;
    COIN_INSERT_MODE :=0;
    FLAG1 := '';
    FLAG2 := '';
    FLAG3 := '';
    EXTERNAL_FILE := 0;
    FILE_NAME := '';
end;

function TGame.Title: String;
begin
//--
end;

procedure TfrmAddFile.New(Data : TGame);
begin
  with Game do
  begin
    Init;
    SD_DRIVE := Data.SD_DRIVE;
    oper := data.oper;
  end;

  FillComboBox;
  edtROM_NAME.Clear;
  edtDEFAULT_DISPLAY_NAME.Clear;
  edtFILE_NAME.Clear;
  edtVideoFile.Clear;
  edtCOIN_INSERT_MODE.Text := '0';
  edtCOIN_DURATION.Text := '100';
  edtLoadTime.Text := '0';
  cbSUPPORTED_FEATURES.ItemIndex := 5;
  cbGENRE.ItemIndex := 0;
  cbDEFAULT_CONFIG.ItemIndex := 0;
  chkENTERNO.Checked := False;
  Show;
end;

procedure TfrmAddFile.SpeedButton1Click(Sender: TObject);
var
  Emu : Integer;
  ROM_DIR, ROM_FILE, MP4_FILE : String;
begin
  // Check required fields First;
  if edtFILE_NAME.Text = '' then
  begin
    ShowMessage('Please select a ROM file');
    exit;
  end;

  if edtROM_NAME.Text = '' then
  begin
    ShowMessage('Please insert a ROM name');
    edtROM_NAME.SetFocus;
    exit;
  end;

  if edtDEFAULT_DISPLAY_NAME.Text = '' then
  begin
    ShowMessage('Please insert a Display Name');
    edtDEFAULT_DISPLAY_NAME.SetFocus;
    exit;
  end;

  Emu :=Integer(cbEmulator.Items.Objects[cbEmulator.ItemIndex]);



  if Game.oper = opEdit then
  begin

    // ** Parameters
    {  ROMNAME : romname <-- PRI_KEY
       DEFAULT_DISPLAY_NAME =:displayname,
       EMULATOR = :emu,
       GENRE = :gen,
       LOAD_TIME = :loadtime,
       COIN_DURATION = :coinduration,
       DEFAULT_CONFIG = :defaultconfig,
       SUPPORT_FEATURES = :supportfeatures,
       COIN_INSERT_MODE = :coininsertmode,
       FLAG1 = :flag1,
       FLAG2 = :flag2,
       FLAG3 = :flag3,
       EXTERNO = :externo}

    qGames_Update.ParamByName('romname').AsString := edtROM_NAME.Text;
    qGames_Update.ParamByName('displayname').AsString := edtDEFAULT_DISPLAY_NAME.Text;
    qGames_Update.ParamByName('emu').AsInteger := Emu;
    qGames_Update.ParamByName('gen').AsInteger := cbGENRE.ItemIndex;
    if edtLoadTime.Text <> '' then
      qGames_Update.ParamByName('loadtime').AsInteger := StrToInt(edtLoadTime.Text)
    else
      qGames_Update.ParamByName('loadtime').Clear;

   if edtCOIN_DURATION.Text <> '' then
      qGames_Update.ParamByName('coinduration').AsInteger :=StrToInt(edtCOIN_DURATION.Text)
    else
      qGames_Update.ParamByName('coinduration').Clear;


    qGames_Update.ParamByName('defaultconfig').AsString := '';
    qGames_Update.ParamByName('supportfeatures').AsInteger := 28;
    if edtCOIN_INSERT_MODE.Text <> '' then
      qGames_Update.ParamByName('coininsertmode').AsInteger := StrToInt(edtCOIN_INSERT_MODE.Text)
    else
      qGames_Update.ParamByName('coininsertmode').Clear;


    qGames_Update.ParamByName('flag1').Clear;
    qGames_Update.ParamByName('flag2').Clear;
    qGames_Update.ParamByName('flag3').Clear;
    qGames_Update.ExecSQL;

  end;

  if Game.oper = opNew then
  begin
    qGet_Rec_Num.Open;
    Game.CREATE_ID := qGet_Rec_NumCREATE_ID.AsInteger + 1;
    qGet_Rec_Num.Close;

    ROM_FILE := ExtractFileName(edtFILE_NAME.Text);
    ROM_DIR := game.SD_DRIVE + FindRomPath(Emu) + TPath.GetFileNameWithoutExtension(ROM_FILE);

    if not DirectoryExists(ROM_DIR) then  // Create Directory structure
      MkDir(ROM_DIR); // -- Find ROM PATH

    // Create ROM File
    if not FileExists(ROM_DIR + '\' + ROM_FILE) then
      CopyFile(Pchar(edtFILE_NAME.Text), Pchar(ROM_DIR + '\' + ROM_FILE), false);

    // Copy MP4
    MP4_FILE := ExtractFileName(edtVideoFile.Text);
    if not FileExists(ROM_DIR + '\' + MP4_FILE) then
      CopyFile(Pchar(edtVideoFile.Text), Pchar(ROM_DIR + '\' + MP4_FILE), false);

{ -- Parameters For Insert --
      romname,
      def_display_name
      emulator,
      genre,
      load_time,
      coin_duration,
      coin_insert_mode,
      default_conf,
      suport_f,
      flag1,
      flag2,
      flag3,
      externo;

   --}

    qGames_Insert.ParamByName('createid').AsInteger := Game.CREATE_ID;

    if edtROM_NAME.Text <> '' then
      qGames_Insert.ParamByName('romname').AsString := edtROM_NAME.Text
    else
      qGames_Insert.ParamByName('romname').Clear;

    if edtDEFAULT_DISPLAY_NAME.Text <> '' then
      qGames_Insert.ParamByName('def_display_name').AsString := edtDEFAULT_DISPLAY_NAME.Text
    else
      qGames_Insert.ParamByName('def_display_name').Clear;

    qGames_Insert.ParamByName('emulator').AsInteger := Emu;
    qGames_Insert.ParamByName('genre').AsInteger := cbGENRE.ItemIndex;

    if edtLoadTime.Text <> '' then
      qGames_Insert.ParamByName('load_time').AsInteger := StrToInt(edtLoadTime.Text)
    else
      qGames_Insert.ParamByName('load_time').Clear;

    if edtCOIN_DURATION.Text <> '' then
      qGames_Insert.ParamByName('coin_duration').AsInteger := StrToInt(edtCOIN_DURATION.Text)
    else
      qGames_Insert.ParamByName('coin_duration').Clear;

    if edtCOIN_INSERT_MODE.Text <> '' then
      qGames_Insert.ParamByName('coin_insert_mode').AsInteger := StrToInt(edtCOIN_INSERT_MODE.Text)
    else
      qGames_Insert.ParamByName('coin_insert_mode').Clear;

    qGames_Insert.ParamByName('default_conf').AsString := '';
    qGames_Insert.ParamByName('suport_f').AsInteger := 28;
    qGames_Insert.ParamByName('flag1').AsString := '';
    qGames_Insert.ParamByName('flag2').AsString := '';
    qGames_Insert.ParamByName('flag3').AsString := '';
    qGames_Insert.ParamByName('externo').AsInteger := 0;
    qGames_Insert.ExecSQL;
  end;
  Close;

end;

function TfrmAddFile.FindRomPath(emu : Integer): String;
begin
  result:= '';
   case emu of
     0..1 : result:= ROMPATH_FBA42;
     2: result:= ROMPATH_MAME37;
     3: result:= ROMPATH_MAME139;
     4: result:= ROMPATH_MAME78;
     5: result:= ROMPATH_FBA42;
     11 : result:= ROMPATH_GBA;
     12 : result:= ROMPATH_MD;
     13 : result:= ROMPATH_PCE;
     14 : result:= ROMPATH_DC;
     15 : result:= ROMPATH_MAME19;
     // MISSING FOLDERS PUT THEM WHEN YOU CAN

   end;
end;

procedure TfrmAddFile.SpeedButton2Click(Sender: TObject);
begin
  Close;
end;

procedure TfrmAddFile.SpeedButton3Click(Sender: TObject);
begin
  OpenDialog.Filter := 'ZIP files (*.zip)|*.zip';
  if OpenDialog.Execute() then
  begin
    if FileExists(OpenDialog.FileName) then
    begin
        edtFILE_NAME.Text := OpenDialog.FileName;
        edtROM_NAME.Text := TPath.GetFileNameWithoutExtension(OpenDialog.FileName);
        edtDEFAULT_DISPLAY_NAME.Text := edtROM_NAME.Text;

    end;

  end;
end;

procedure TfrmAddFile.SpeedButton4Click(Sender: TObject);
begin
 OpenDialog.Filter := 'MPEG-4 files (*.mp4)|*.mp4';
 OpenDialog.DefaultExt := 'mp4';
 if OpenDialog.Execute() then
  begin
    if FileExists(OpenDialog.FileName) then
    begin
        edtVideoFile.Text := OpenDialog.FileName;
    end;

  end;
end;

procedure TfrmAddFile.FillComboBox;
var
  i : Integer;
begin
  for I := 0 to 17 do
  begin
    Emu[i].Name := Emuladores[i];
    Emu[i].Id := EmuCod[i];
  end;

  cbGenre.Items.Clear;
  for i := 0 to Length(Genero)-1 do
  begin
    cbGenre.AddItem(genero[i], TObject(Integer(i)));
  end;

    //Emuladores
  cbEMULATOR.Items.Clear;
  for i := 0 to 17 do
  begin
    cbEMULATOR.AddItem(Emu[i].Name, TObject(Integer(Emu[i].Id)));
  end;

end;

procedure TfrmAddFile.Edit(Data : TGame);
var
  emu_dir : String;
  i : Integer;
begin

  game.Init;
  game := Data;
  FillComboBox;
  // -- Fill Form
  with game do
  begin
    //  CREATE_ID := 0;
    edtROM_NAME.Text := ROM_NAME;
    edtDEFAULT_DISPLAY_NAME.Text := DEFAULT_DISPLAY_NAME;
    //FILE_NAME := '';

    // --- Emulator
    for i := 0 to 17 do
    begin
      if emu[i].Id = EMULATOR then
        cbEMULATOR.ItemIndex :=  i;
    end;


    emu_dir := cbEMULATOR.Text;
    if (emu_dir = 'PS') or (emu_dir = 'N64')
       or (emu_dir = 'PPSSPP') then
      emu_dir := 'family';

    if (emu_dir = 'FBA42_FMY') then
      emu_dir := 'fba42';


    edtFILE_NAME.Text := game.SD_DRIVE + GAMESDIR + emu_dir + '\' + ROM_NAME + '\';
    edtVideoFile.Text := game.SD_DRIVE + GAMESDIR + emu_dir + '\' + ROM_NAME + '\' +
                         ROM_NAME + '.mp4';


    if (emu_dir <> 'family') then
    begin
      edtFILE_NAME.Text := edtFILE_NAME.Text + ROM_NAME +'.zip';
    end;

    //  GENRE := 0;
    cbGenre.ItemIndex := GENRE;
    edtLoadTime.Text := IntToStr(LOAD_TIME);
    edtCOIN_DURATION.Text  := IntToStr(COIN_DURATION);
    cbDEFAULT_CONFIG.Text := DEFAULT_CONFIG;
    cbSUPPORTED_FEATURES.Text := IntToStr(SUPPORTED_FEATURES);
    edtCOIN_INSERT_MODE.Text := IntToStr(COIN_INSERT_MODE);
    cbFLAG1.Text := FLAG1;
    cbFLAG2.Text := FLAG2;
    cbFLAG3.Text := FLAG3;
    if EXTERNAL_FILE <> Null then
      chkENTERNO.Checked := EXTERNAL_FILE > 0 ;
    //FILE_NAME := '';
    Show;
  end;
end;

end.
