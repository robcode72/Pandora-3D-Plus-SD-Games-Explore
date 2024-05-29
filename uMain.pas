unit uMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, FireDAC.Stan.Intf,
  FireDAC.Stan.Option, FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf,
  FireDAC.Stan.Def, FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys,
  FireDAC.Phys.SQLite, FireDAC.Phys.SQLiteDef, FireDAC.Stan.ExprFuncs,
  FireDAC.Phys.SQLiteWrapper.Stat, FireDAC.VCLUI.Wait, FireDAC.Stan.Param,
  FireDAC.DatS, FireDAC.DApt.Intf, FireDAC.DApt, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client, Vcl.Grids, Vcl.DBGrids, Vcl.MPlayer, Vcl.ExtCtrls,
  Vcl.StdCtrls, Vcl.DBCtrls, Vcl.Mask, Vcl.Buttons, Vcl.FileCtrl, Vcl.OleCtrls,
  WMPLib_TLB, Vcl.Imaging.jpeg, Vcl.Imaging.pngimage, Vcl.ComCtrls,
  System.ImageList, Vcl.ImgList, System.Rtti, System.Bindings.Outputs,
  Vcl.Bind.Editors, Data.Bind.EngExt, Vcl.Bind.DBEngExt, Data.Bind.Components,
  Data.Bind.DBScope, Data.SqlExpr, FireDAC.Moni.Base, FireDAC.Moni.RemoteClient,
  Vcl.Menus, SHellApi, System.Actions, Vcl.ActnList, Vcl.StdActns;


type
  TfrmMain = class(TForm)
    dsGames: TDataSource;
    FDConnection1: TFDConnection;
    qGames: TFDQuery;
    qGamesROM_NAME: TStringField;
    qGamesDEFAULT_DISPLAY_NAME: TStringField;
    qGamesEMULATOR: TIntegerField;
    qGamesGENRE: TIntegerField;
    qGamesLOAD_TIME: TIntegerField;
    qGamesCOIN_DURATION: TIntegerField;
    qGamesDEFAULT_CONFIG: TStringField;
    qGamesSUPPORT_FEATURES: TIntegerField;
    qGamesCOIN_INSERT_MODE: TIntegerField;
    qGamesFLAG1: TStringField;
    qGamesFLAG2: TStringField;
    qGamesFLAG3: TStringField;
    qGamesEXTERNAL: TIntegerField;
    qGamesROM_NAME_1: TStringField;
    qGamesHIDE: TIntegerField;
    qGamesTOP_ID: TIntegerField;
    qGamesCURRENT_CONFIG: TStringField;
    qGamesTARGET_CONFIG: TStringField;
    qGamesPLAY_ID: TIntegerField;
    qGamesKEY_COMBINATION: TIntegerField;
    qGamesCREATE_ID: TIntegerField;
    StatusBar1: TStatusBar;
    pToolbar: TPanel;
    sbOpen: TSpeedButton;
    pWorkArea: TPanel;
    pGrid: TPanel;
    DBGrid1: TDBGrid;
    pDatos: TPanel;
    DBEdit4: TDBEdit;
    DBEdit1: TDBEdit;
    DBEdit3: TDBEdit;
    DBEdit2: TDBEdit;
    DBEdit5: TDBEdit;
    pVideo: TPanel;
    Image2: TImage;
    Panel1: TPanel;
    Image1: TImage;
    IlComunes: TImageList;
    Image3: TImage;
    Bevel1: TBevel;
    lbledtFind: TLabeledEdit;
    sbFilter: TSpeedButton;
    cbGenre: TComboBox;
    Image4: TImage;
    Image5: TImage;
    lblDriveName: TLabel;
    ImageList32x32: TImageList;
    cbEmulator: TComboBox;
    Image6: TImage;
    Image7: TImage;
    Image8: TImage;
    Image9: TImage;
    Image10: TImage;
    Label1: TLabel;
    qFavorites: TFDQuery;
    BindSourceDB1: TBindSourceDB;
    BindingsList1: TBindingsList;
    qGames_Update: TFDQuery;
    SpeedButton2: TSpeedButton;
    sbUp: TSpeedButton;
    sbDown: TSpeedButton;
    Label2: TLabel;
    Label3: TLabel;
    qFavoritesTOP_ID: TIntegerField;
    LinkFillControlToField1: TLinkFillControlToField;
    lvFavorites: TListView;
    qFavoritesLOCATE_DISPLAY_NAME: TStringField;
    qFavoritesROM_NAME: TStringField;
    LinkListControlToField1: TLinkListControlToField;
    cbDrivers: TComboBox;
    FDMoniRemoteClientLink1: TFDMoniRemoteClientLink;
    WindowsMediaPlayer1: TWindowsMediaPlayer;
    sbExport: TSpeedButton;
    qGamesLANGUAGE: TStringField;
    qGamesLOCATE_DISPLAY_NAME: TStringField;
    sbAddGame: TSpeedButton;
    TaskDialog: TTaskDialog;
    sbEditGame: TSpeedButton;
    qFavoritesDEFAULT_DISPLAY_NAME: TStringField;
    qFavoritesLANGUAGE: TStringField;
    dsFavoritos: TDataSource;
    qFavoritesEMULATOR: TIntegerField;
    qFavoritesGENRE: TIntegerField;
    sbUnmount: TSpeedButton;
    SpeedButton1: TSpeedButton;
    mnMain: TMainMenu;
    File1: TMenuItem;
    ools1: TMenuItem;
    YouTubeDownload1: TMenuItem;
    Close1: TMenuItem;
    N1: TMenuItem;
    Exportto1: TMenuItem;
    ImportFavorites1: TMenuItem;
    SaveDialog: TSaveDialog;
    ActionList1: TActionList;
    actSaveDB: TAction;
    actSaveFav: TAction;
    SaveFavoritesto1: TMenuItem;
    N2: TMenuItem;
    OpenDialog: TOpenDialog;
    actImport: TAction;
    FDMemTable: TFDMemTable;
    Clearfavoritelist1: TMenuItem;
    N3: TMenuItem;

    procedure qGamesEMULATORGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure sbFilterClick(Sender: TObject);
    procedure FDConnection1Error(ASender, AInitiator: TObject;
      var AException: Exception);
    procedure sbOpenClick(Sender: TObject);
    procedure qGamesGENREGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure dsGamesDataChange(Sender: TObject; Field: TField);
    procedure DBGrid1DrawDataCell(Sender: TObject; const Rect: TRect;
      Field: TField; State: TGridDrawState);
    procedure DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGrid1DblClick(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure cbDriversClick(Sender: TObject);
    procedure qGamesAfterOpen(DataSet: TDataSet);
    procedure lvFavoritesClick(Sender: TObject);
    procedure WindowsMediaPlayer1StatusChange(Sender: TObject);
    procedure lbledtFindChange(Sender: TObject);
    procedure sbDownClick(Sender: TObject);
    procedure sbUpClick(Sender: TObject);
    procedure sbAddGameClick(Sender: TObject);
    procedure TaskDialogButtonClicked(Sender: TObject;
      ModalResult: TModalResult; var CanClose: Boolean);
    procedure TaskDialogTimer(Sender: TObject; TickCount: Cardinal;
      var Reset: Boolean);
    procedure sbEditGameClick(Sender: TObject);
    procedure lvFavoritesDragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure lvFavoritesDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure qFavoritesEMULATORGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure qFavoritesGENREGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure FormActivate(Sender: TObject);
    procedure sbUnmountClick(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure Close1Click(Sender: TObject);
    procedure YouTubeDownload1Click(Sender: TObject);
    procedure DBGrid1TitleClick(Column: TColumn);
    procedure Exportto1Click(Sender: TObject);
    procedure FileSaveAsBeforeExecute(Sender: TObject);
    procedure FileSaveAsSaveDialogClose(Sender: TObject);
    procedure actSaveDBExecute(Sender: TObject);
    procedure actSaveFavExecute(Sender: TObject);
    procedure actImportExecute(Sender: TObject);
    procedure Clearfavoritelist1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    SDDrive: Char;
    procedure FindRom(RomName : String; Genre, emu : Integer);
    procedure GetDrives;
    procedure ShowRecNum;


  end;

var
  frmMain: TfrmMain;




implementation

{$R *.dfm}

uses uExcel, uAddFile, uUtils;

var
  Emu : Array[0..17] of TEmuladores;
  Game : TGame;

procedure TfrmMain.GetDrives;
var
  Drive : Char;
begin
  cbDrivers.Items.Clear;
  for Drive := 'A' to 'Z' do
  begin
    if GetDriveType(PChar(Drive + ':\')) in [DRIVE_REMOVABLE] then   //, DRIVE_CDROM, DRIVE_RAMDISK
    begin
      cbDrivers.Items.add(Drive + ':');
      if DirectoryExists(Drive + ':\games') then
      begin
        cbDrivers.ItemIndex := cbDrivers.GetCount-1;
        SDDrive := Drive ;
      end;
    end;
  end;
end;

procedure TfrmMain.lbledtFindChange(Sender: TObject);
begin
  sbFilter.OnClick(Self);
end;

procedure TfrmMain.lvFavoritesClick(Sender: TObject);
begin
  ShowRecNum;
end;

procedure TfrmMain.lvFavoritesDragDrop(Sender, Source: TObject; X, Y: Integer);
var
  MyItem : TListItem;
  Id1, Id2 : Integer;
  RomName1, RomName2 : String;
begin
  if Source = lvFavorites then
  with Source as TListView do
  begin
    if GetItemAt(X,Y)<>nil then
    begin
      Id1 := BindSourceDB1.DataSource.DataSet.FieldByName('TOP_ID').AsInteger;
      RomName1 := BindSourceDB1.DataSource.DataSet.FieldByName('ROM_NAME').AsString;
      MyItem:=Items.Insert(GetItemAt(X,Y).Index);
      MyItem.Assign(Selected);
      MyItem.Selected := True;
      Id2 := BindSourceDB1.DataSource.DataSet.FieldByName('TOP_ID').AsInteger;
      RomName2 := BindSourceDB1.DataSource.DataSet.FieldByName('ROM_NAME').AsString;
      Selected.Delete;

      // First Update Rom That Moved Up
      qGames_Update.ParamByName('num').AsInteger := Id2;
      qGames_Update.ParamByName('rom_name').AsString := RomName1;
      qGames_Update.ExecSQL;

      // Rom That Moved Down
      qGames_Update.ParamByName('num').AsInteger := Id1;
      qGames_Update.ParamByName('rom_name').AsString := RomName2;
      qGames_Update.ExecSQL;

      qGames.Refresh;

      BindSourceDB1.DataSet.Refresh;
    end;
  end;
end;

procedure TfrmMain.lvFavoritesDragOver(Sender, Source: TObject; X, Y: Integer;
  State: TDragState; var Accept: Boolean);
begin
  If Source = lvFavorites then
    begin
      Accept := True;
    end;
end;

procedure TfrmMain.qFavoritesEMULATORGetText(Sender: TField; var Text: string;
  DisplayText: Boolean);
var
 i : Integer;

begin
  for i := 0 to 17 do
  begin
    if emu[i].Id = Sender.AsInteger  then
      Text :=  emu[i].Name;
  end;

end;

procedure TfrmMain.qFavoritesGENREGetText(Sender: TField; var Text: string;
  DisplayText: Boolean);
begin
  if DisplayText then
    if Sender.AsInteger >= 0 then
      Text :=  Genero[Sender.AsInteger]; // + ' ' + Sender.AsString;

end;

procedure TfrmMain.qGamesAfterOpen(DataSet: TDataSet);
begin
  ShowRecNum;
end;

procedure TfrmMain.ShowRecNum;
begin
  StatusBar1.Panels[4].Text := IntToStr(qGames.RecNo) + '/' + IntToStr(qGames.RecordCount);
  StatusBar1.Panels[1].Text := IntToStr(lvFavorites.ItemIndex+1) + '/' +
                               IntToStr(qFavorites.RecordCount);
end;

procedure TfrmMain.qGamesEMULATORGetText(Sender: TField; var Text: string;
  DisplayText: Boolean);
var
 i : Integer;

begin
  for i := 0 to 17 do
  begin
    if emu[i].Id = Sender.AsInteger  then
      Text :=  emu[i].Name;
  end;
end;

procedure TfrmMain.qGamesGENREGetText(Sender: TField; var Text: string;
  DisplayText: Boolean);
begin
   if DisplayText then
    if Sender.AsInteger >= 0 then
      Text :=  Genero[Sender.AsInteger]; // + ' ' + Sender.AsString;

end;

procedure TfrmMain.sbOpenClick(Sender: TObject);
var
  sDir : String;
begin
  //TaskDialog.OnTimer := TaskDialogTimer;
  //TaskDialog.Execute;

  sDir := cbDrivers.Text + '\games\data\';

  if DirectoryExists(sDir) then
  try
    FDConnection1.Params.Database := sDir + '\games.db';
    FDConnection1.Connected := True;
    FindRom('', -1, -1);
    qFavorites.Open();
    ShowRecNum;
    sbUnmount.Enabled := True;
  except
      on E : EDataBaseError do
        begin
          raise;
        end;
  end;
end;

procedure TfrmMain.sbFilterClick(Sender: TObject);
var
  Gen, Emu : Integer;
begin
   Gen := Integer(cbGenre.Items.Objects[cbGenre.ItemIndex]);
   Emu := Integer(cbEmulator.Items.Objects[cbEmulator.ItemIndex]);
   FindRom(lbledtFind.Text, Gen, Emu);
end;

procedure TfrmMain.sbAddGameClick(Sender: TObject);
begin
   with Game do
  begin
    oper := opNew;
    SD_DRIVE := SDDrive + ':';
    uAddFile.ShowForm(Game);
  end;
  qGames.Refresh;
end;

procedure TfrmMain.sbUnmountClick(Sender: TObject);
begin
 try
    WindowsMediaPlayer1.close;
    FDConnection1.Close;

    EjectVolume(SDDrive);
    sbUnmount.Enabled := False;
    GetDrives;
  finally
    ShowMessage('You can now safely remove SD card');
  end;
end;

procedure TfrmMain.SpeedButton1Click(Sender: TObject);
begin
  GetDrives;
end;

procedure TfrmMain.SpeedButton2Click(Sender: TObject);
var
  i : Integer;
begin
  if lvFavorites.Items.Count = 0 then exit;

  qGames_Update.ParamByName('num').AsInteger := 0;
  qGames_Update.ParamByName('rom_name').AsString := BindSourceDB1.DataSource.DataSet.FieldByName('ROM_NAME').AsString;
  qGames_Update.ExecSQL;
  qGames.Close;
  qGames.Open;

  BindSourceDB1.DataSet.Refresh;
end;

procedure TfrmMain.TaskDialogButtonClicked(Sender: TObject;
  ModalResult: TModalResult; var CanClose: Boolean);
begin
  TaskDialog.OnTimer := nil;
  CanClose := true;
end;

procedure TfrmMain.TaskDialogTimer(Sender: TObject; TickCount: Cardinal;
  var Reset: Boolean);
var
  i : Integer;
begin
  //TaskDialog.ProgressBar.Position :=


end;

procedure TfrmMain.sbUpClick(Sender: TObject);
var
  Index, Id1, Id2: Integer;
  temp : TListItem;
  RomName1, RomName2 : String;
  BM : TBookmark;
begin
    if lvFavorites.SelCount>0 then
    begin
      Index := lvFavorites.Selected.Index;
      Id1 := BindSourceDB1.DataSource.DataSet.FieldByName('TOP_ID').AsInteger;
      RomName1 := BindSourceDB1.DataSource.DataSet.FieldByName('ROM_NAME').AsString;
      if Index>0 then
      begin
        temp := lvFavorites.Items.Insert(Index-1);
        temp.Assign(lvFavorites.Items.Item[Index+1]);
        lvFavorites.Items.Delete(Index+1);
        // fix display so moved item is selected/focused
        lvFavorites.Selected := temp;
        lvFavorites.ItemFocused := temp;
        Id2 := BindSourceDB1.DataSource.DataSet.FieldByName('TOP_ID').AsInteger;
        RomName2 := BindSourceDB1.DataSource.DataSet.FieldByName('ROM_NAME').AsString;

        // First Update Rom That Moved Up
        qGames_Update.ParamByName('num').AsInteger := Id1-1;
        qGames_Update.ParamByName('rom_name').AsString := RomName1;
        qGames_Update.ExecSQL;

        // Rom That Moved Down
        qGames_Update.ParamByName('num').AsInteger := Id2+1;
        qGames_Update.ParamByName('rom_name').AsString := RomName2;
        qGames_Update.ExecSQL;

        qGames.Refresh;

        BindSourceDB1.DataSet.Refresh;
        //LinkListControlToField1.Active:=False;
        //LinkListControlToField1.Active:=true
      end;
    end;
end;

procedure TfrmMain.sbDownClick(Sender: TObject);
var
  Index,  Id1, Id2: integer;
  temp : TListItem;
  RomName1, RomName2 : String;
  BM : TBookmark;
begin
  // use a button that cannot get focus, such as TSpeedButton

  //if lvFavorites.Focused then
    if lvFavorites.SelCount>0 then
    begin
      Index := lvFavorites.Selected.Index;
      Id1 := BindSourceDB1.DataSource.DataSet.FieldByName('TOP_ID').AsInteger;
      RomName1 := BindSourceDB1.DataSource.DataSet.FieldByName('ROM_NAME').AsString;
      if Index<lvFavorites.Items.Count then
      begin
        temp := lvFavorites.Items.Insert(Index+2);
        temp.Assign(lvFavorites.Items.Item[Index]);
        lvFavorites.Items.Delete(Index);
        // fix display so moved item is selected/focused
        lvFavorites.Selected := temp;
        lvFavorites.ItemFocused := temp;
        Id2 := BindSourceDB1.DataSource.DataSet.FieldByName('TOP_ID').AsInteger;
        RomName2 := BindSourceDB1.DataSource.DataSet.FieldByName('ROM_NAME').AsString;

        // First Update Rom That Moved down
        qGames_Update.ParamByName('num').AsInteger := Id1+1;
        qGames_Update.ParamByName('rom_name').AsString := RomName1;
        qGames_Update.ExecSQL;
        qGames.Close;
        qGames.Open;

        // Rom That Moved down
        qGames_Update.ParamByName('num').AsInteger := Id2-1;
        qGames_Update.ParamByName('rom_name').AsString := RomName2;
        qGames_Update.ExecSQL;
        qGames.Close;
        qGames.Open;
        BindSourceDB1.DataSet.Refresh;

      end;
    end;
end;

procedure TfrmMain.sbEditGameClick(Sender: TObject);
begin
  with Game do
  begin
    oper := opEdit;
    SD_DRIVE := SDDrive;
    CREATE_ID := qGamesCREATE_ID.AsInteger;
    ROM_NAME := qGamesROM_NAME.AsString;
    DEFAULT_DISPLAY_NAME := qGamesDEFAULT_DISPLAY_NAME.AsString;
    EMULATOR := qGamesEMULATOR.AsInteger;
    GENRE := qGamesGENRE.AsInteger;
    LOAD_TIME := qGamesLOAD_TIME.AsInteger;
    COIN_DURATION := qGamesCOIN_DURATION.AsInteger;
    DEFAULT_CONFIG := qGamesDEFAULT_CONFIG.AsString;
    SUPPORTED_FEATURES := qGamesSUPPORT_FEATURES.AsInteger;
    COIN_INSERT_MODE := qGamesCOIN_INSERT_MODE.AsInteger;
    FLAG1 := qGamesFLAG1.AsString;
    FLAG2 := qGamesFLAG2.AsString;
    FLAG3 := qGamesFLAG3.AsString;
    EXTERNAL_FILE := qGamesEXTERNAL.AsInteger;
    //FILE_NAME
    uAddFile.ShowForm(Game);
  end;

end;

procedure TfrmMain.WindowsMediaPlayer1StatusChange(Sender: TObject);
begin
  if WindowsMediaPlayer1.status = 'Finalizado' then
  begin
    WindowsMediaPlayer1.controls.fastReverse;
    WindowsMediaPlayer1.controls.play;
  end;
  if WindowsMediaPlayer1.status = 'Detenido' then
  begin
    WindowsMediaPlayer1.controls.play;
  end;
end;

procedure TfrmMain.YouTubeDownload1Click(Sender: TObject);
var
  pathfile : PAnsiChar;
  AnsiStr: AnsiString;
begin
  pathfile := PAnsiChar(AnsiString(ExtractFileDir(Application.ExeName) + '\tools\ytd.exe'));
  WinExec(pathfile, CmdShow);
end;

procedure TfrmMain.actImportExecute(Sender: TObject);
begin
  // Delete Favorites

  if OpenDialog.Execute then
  begin
    if OpenDialog.FileName <> '' then
    begin
      if FileExists(OpenDialog.FileName) then
        Case OpenDialog.FilterIndex of
          1 : begin
                ImportCSVToDataset(true, FDMemTable, OpenDialog.FileName);  // CSV
                if FDMemTable.Active then
                  ImportFavorites(FDMemTable, qGames_Update, False); // Import data to the DataBase
              end;
          2 : ;  // Excel
        end;
    end;
  end;
  qGames.Refresh;
  qFavorites.Refresh;
end;

procedure TfrmMain.actSaveDBExecute(Sender: TObject);
begin
  SaveDialog.FileName := 'Games';
  if SaveDialog.Execute then
    if SaveDialog.FileName <> '' then
    begin
     case SaveDialog.FilterIndex of
        1 : SaveDatasetToCSV(True, qGames,SaveDialog.FileName);
        2 : ExportaExcel(qGames);
      end;
    end;

end;

procedure TfrmMain.actSaveFavExecute(Sender: TObject);
begin
  SaveDialog.FileName := 'Favorites';
  if SaveDialog.Execute then
    if SaveDialog.FileName <> '' then
    begin
     case SaveDialog.FilterIndex of
        1 : SaveDatasetToCSV(True, qFavorites, SaveDialog.FileName);
        2 : ExportaExcel(qFavorites);
      end;
    end;
end;

procedure TfrmMain.cbDriversClick(Sender: TObject);
begin
  qGames.Close;
  qFavorites.Close;
  sbOpen.OnClick(Self);
end;

procedure TfrmMain.Clearfavoritelist1Click(Sender: TObject);
begin
// ## Reset Favorite List ##
   qGames_Update.SQL.Text := 'UPDATE option SET TOP_ID =0 WHERE TOP_ID > 0 ';
   qGames_Update.ExecSQL;
   qGames_Update.SQL.Text := 'UPDATE option SET TOP_ID =:num WHERE ROM_NAME =:rom_name';
   qFavorites.Refresh;
   qGames.Refresh;
end;

procedure TfrmMain.Close1Click(Sender: TObject);
begin
  Application.Terminate;
end;

procedure TfrmMain.DBGrid1DblClick(Sender: TObject);
var
  i : Integer;
begin
  // Assign a game as your Favorite
  i := qFavorites.RecordCount +1;
  qGames_Update.ParamByName('num').AsInteger := i;
  qGames_Update.ParamByName('rom_name').AsString := qGamesROM_NAME.AsString;
  qGames_Update.ExecSQL;
  qGames.Close;
  qGames.Open;

  BindSourceDB1.DataSet.Refresh;
end;

procedure TfrmMain.DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
  LRect : TRect;
  LGrid : TDBGrid;
  LText : String;
  LPerc : Extended;
  LTextWidth : TSize;
  LSavedPenColor , LSavedBrushColor : Integer;
  LSavedPenStyle : TPenStyle;
  LSavedBrushStyle : TBrushStyle;
  LFavorte : Extended;
  LNeedOwnerDraw : Boolean;
begin
  LGrid := TDBGrid(Sender);
  if [gdSelected, gdFocused] * State <> [] then
    LGrid.Canvas.Brush.Color := clHighlight;

  LNeedOwnerDraw := (Column.FieldName.Equals('TOP_ID'));
  if LNeedOwnerDraw then
  begin
    LRect := Rect;
    LSavedPenColor := LGrid.Canvas.Pen.Color;
    LSavedBrushColor := LGrid.Canvas.Brush.Color;
    LSavedPenStyle := LGrid.Canvas.Pen.Style;
    LSavedBrushStyle := LGrid.Canvas.Brush.Style;
  end;

  if Column.FieldName.Equals('TOP_ID') then
  begin
     LFavorte := Column.Field.AsInteger;
     if LFavorte > 0 then
     begin
        LGrid.Canvas.Brush.Color := LSavedBrushColor;
        LGrid.Canvas.Brush.Style := bsSolid;
        LGrid.Canvas.pen.Style := psClear;
        LGrid.Canvas.FillRect(LRect);
        IlComunes.Draw(LGrid.Canvas, LRect.CenterPoint.X - (IlComunes.Width div 2),
                       LRect.CenterPoint.Y - (IlComunes.Height div 2),
                       0);
     end
     else
      LGrid.Canvas.FillRect(LRect);
  end
  else
    LGrid.DefaultDrawColumnCell(Rect, DataCol, Column, State);

 if LNeedOwnerDraw then
  begin
    LGrid.Canvas.Pen.Color := LSavedPenColor;
    LGrid.Canvas.Brush.Color := LSavedBrushColor;
    LGrid.Canvas.Pen.Style := LSavedPenStyle;
    LGrid.Canvas.Brush.Style := LSavedBrushStyle;
  end;
end;

procedure TfrmMain.DBGrid1DrawDataCell(Sender: TObject; const Rect: TRect;
  Field: TField; State: TGridDrawState);
var
  Bmp: TBitmap;
begin
  if Field.Name ='TOP_ID' then
  begin
    try
      Bmp :=  TBitmap.Create; //TPngImage.Create;
      IlComunes.GetBitmap(0,Bmp);
      DBGrid1.Canvas.StretchDraw(Rect, Bmp);
    finally
      Bmp.Free;
    end
  end
  else
    DBGrid1.DefaultDrawDataCell(Rect,field,State);
end;

procedure TfrmMain.DBGrid1TitleClick(Column: TColumn);
var
  i : Integer;
begin
  for i := 0 to DBGrid1.Columns.Count -1 do
  begin
    DBGrid1.Columns.Items[i].Title.Font.Style := [];
    DBGrid1.Columns.Items[i].Font.Color := clBlack;
    DBGrid1.Columns.Items[i].Font.Style := [];
  end;
  Column.Font.Color := clBlack;
  Column.Font.Style := [fsBold];
  qGames.IndexFieldNames := Column.FieldName;
  Column.Title.Font.Style := [fsBold];
end;

procedure TfrmMain.dsGamesDataChange(Sender: TObject; Field: TField);
var
  i : Integer;
  EmuName : String;
begin
if dsGames.State in [dsBrowse] then
  begin
    ShowRecNum;
    for i := 0 to 17 do
    begin
        if Emu[i].Id = qGames.FieldByName('EMULATOR').AsInteger then
          if (Emu[i].Name = 'PPSSPP') or
            (Emu[i].Name = 'PS') or
            (Emu[i].Name = 'N64') then
            EmuName := 'family'
          else
            EmuName :=   Emu[i].Name;


    end;

    WindowsMediaPlayer1.URL := cbDrivers.Text +  '\games\data' +
                              '\' + EmuName +
                              '\' + qGames.FieldByName('ROM_NAME').AsString +
                              '\' + qGames.FieldByName('ROM_NAME').AsString + '.mp4';

  end;
end;

procedure TfrmMain.Exportto1Click(Sender: TObject);
begin
  ExportaExcel(qFavorites);
end;

procedure TfrmMain.FDConnection1Error(ASender, AInitiator: TObject;
  var AException: Exception);
begin
  raise Exception.Create(AException.Message);
end;

procedure TfrmMain.FileSaveAsBeforeExecute(Sender: TObject);
begin
  SaveDialog.FilterIndex := 1;
end;

procedure TfrmMain.FileSaveAsSaveDialogClose(Sender: TObject);
begin
   if SaveDialog.FileName <> '' then
   case SaveDialog.FilterIndex of
      1 : SaveDatasetToCSV(True, qGames,SaveDialog.FileName + '.csv');
      2 : ExportaExcel(qFavorites);
    end;
end;

procedure TfrmMain.FindRom(RomName : String; Genre, emu : Integer);
var
  i : Integer;

begin

  with qGames do
  begin
    Close;
    SQL.Clear;
    Params.Clear;
    SQL.Add('SELECT * FROM info' +
                  ' JOIN option on info.ROM_NAME=option.ROM_NAME' +
                  ' JOIN display_name on info.ROM_NAME=display_name.ROM_NAME' +
                  ' WHERE UPPER(DEFAULT_DISPLAY_NAME) LIKE UPPER(:ROMNAME)'+
                  ' AND LANGUAGE='+QuotedStr('EN'));

    if Genre > -1 then
    begin
      SQL.Add(' AND GENRE=:gen');
      Params.Items[1].DataType := TFieldType.ftInteger;
      ParamByName('gen').AsInteger := Genre
    end;
    if Emu > -1 then
    begin
      SQL.Add(' AND EMULATOR=:emu');
      Params.Items[Params.Count-1].DataType := TFieldType.ftInteger;
      ParamByName('emu').AsInteger := emu
    end;

    Params.Items[0].DataType := TFieldType.ftString;
    ParamByName('ROMNAME').Value := (RomName + '%');

    SQL.Add(' group by info.DEFAULT_DISPLAY_NAME' +
            ' order by CREATE_ID');

    try
      if Assigned(qGames) then
        Open;

    except
      on E : EDataBaseError do
        begin
          raise;
        end;
  end;
  end;

end;
procedure TfrmMain.FormActivate(Sender: TObject);
begin
  if qGames.Active then
    qGames.Refresh;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
  i : Integer;

begin
  sbExport.Caption := '';

  for I := 0 to 17 do
  begin
    Emu[i].Name := Emuladores[i];
    Emu[i].Id := EmuCod[i];
  end;

  GetDrives;
  sbOpenClick(Self);

  cbGenre.Items.Clear;
  cbGenre.AddItem('All', TObject(Integer(-1)));

  for i := 0 to Length(Genero)-1 do
  begin
    cbGenre.AddItem(genero[i], TObject(Integer(i)));
  end;
  cbGenre.ItemIndex := 0;

  //Emuladores
  cbEmulator.Items.Clear;
  cbEmulator.AddItem('All', TObject(Integer(-1)));

  for i := 0 to 17 do
  begin
    cbEmulator.AddItem(Emu[i].Name, TObject(Integer(Emu[i].Id)));
  end;
  cbEmulator.ItemIndex := 0;
  if SDDrive = #0 then
  begin
      sbUnmount.Enabled := False;
      //sbUnmount.Caption := 'Mount Micro SD';
  end;
end;

end.
