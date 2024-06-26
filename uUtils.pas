unit uUtils;

interface

uses Winapi.Windows, System.SysUtils,
     FireDAC.Comp.BatchMove,
     FireDAC.Comp.BatchMove.Text,
     FireDAC.Comp.BatchMove.DataSet,
     Data.DB, FireDAC.Comp.Client,
     VCL.Dialogs;

type TEmuladores = record
      Name : String;
      Id : Integer;
      Ext : String;
end;

const
  Emuladores: TArray<String> = ['FBA08', 'MAME37', 'MAME139', 'MAME78', 'FBA42',
    'PPSSPP', 'PS', 'N64', 'FBA42_FMY', 'FC', 'SFC', 'GBA', 'GBC','MD', 'PCE',
    'DC','MAME19','MAME199'];

  Genero: TArray<String> = ['Other', 'Fighting', 'Action', 'Shotting', 'Sport', 'Puzzle', 'Racing'];

  EmuCod : TArray<Integer> = [0,2,3,4,5,6,7,8,10,11,12,13,14,15,17,18,19,21]; // Emulator Internal Code

  Emu_FileExt: TArray<String> = ['.zip', '.iso', '', '', '', '', '']; // Emulator File Extension

  ROMPATH_DC : String = '\games\data\DC\';
  ROMPATH_FAMILY : String = '\games\data\family\';
  ROMPATH_FBA42 : String = '\games\data\fba42\';
  ROMPATH_FC : String = '\games\data\FC';
  ROMPATH_GBA : String = '\games\data\GBA\';
  ROMPATH_MAME19 : String = '\games\data\mame19\';
  ROMPATH_MAME37 : String = '\games\data\mame37\';
  ROMPATH_MAME78 : String = '\games\data\mame78\';
  ROMPATH_MAME139 : String = '\games\data\mame139\';
  ROMPATH_MAME199 : String = '\games\data\mame199\';
  ROMPATH_MD : String = '\games\data\MD\';
  ROMPATH_PCE : String = '\games\data\PCE\';
  ROMPATH_SFC : String = '\games\data\SFC\';
  ROMPATH_WSC : String = '\games\data\WSC\';

  function EjectVolume(ADrive: char): boolean;
  procedure SaveDatasetToCSV(IncludeFieldNames : Boolean; Dataset : TDataset; Filename : String);
  procedure ImportCSVToDataset(IncludeFieldNames : Boolean; dDataset : TDataset; Filename : String);
  procedure ImportFavorites(RDataSet : TDataSet; WDataSet : TFDQuery; DelFav : Boolean);

implementation

function OpenVolume(ADrive: char): THandle;
var
  RootName, VolumeName: string;
  AccessFlags: DWORD;
begin
  RootName := ADrive + ':\'; (* '\'' // keep SO syntax highlighting working *)
  case GetDriveType(PChar(RootName)) of
    DRIVE_REMOVABLE:
      AccessFlags := GENERIC_READ or GENERIC_WRITE;
    DRIVE_CDROM:
      AccessFlags := GENERIC_READ;
  else
    Result := INVALID_HANDLE_VALUE;
    exit;
  end;
  VolumeName := Format('\\.\%s:', [ADrive]);
  Result := CreateFile(PChar(VolumeName), AccessFlags,
    FILE_SHARE_READ or FILE_SHARE_WRITE, nil, OPEN_EXISTING, 0, 0);
  if Result = INVALID_HANDLE_VALUE then
    RaiseLastWin32Error;
end;

function LockVolume(AVolumeHandle: THandle): boolean;
const
  LOCK_TIMEOUT = 10 * 1000; // 10 Seconds
  LOCK_RETRIES = 20;
  LOCK_SLEEP = LOCK_TIMEOUT div LOCK_RETRIES;

// #define FSCTL_LOCK_VOLUME CTL_CODE(FILE_DEVICE_FILE_SYSTEM, 6, METHOD_BUFFERED, FILE_ANY_ACCESS)
  FSCTL_LOCK_VOLUME = (9 shl 16) or (0 shl 14) or (6 shl 2) or 0;
var
  Retries: integer;
  BytesReturned: Cardinal;
begin
  for Retries := 1 to LOCK_RETRIES do begin
    Result := DeviceIoControl(AVolumeHandle, FSCTL_LOCK_VOLUME, nil, 0,
      nil, 0, BytesReturned, nil);
    if Result then
      break;
    Sleep(LOCK_SLEEP);
  end;
end;

function DismountVolume(AVolumeHandle: THandle): boolean;
const
// #define FSCTL_DISMOUNT_VOLUME CTL_CODE(FILE_DEVICE_FILE_SYSTEM, 8, METHOD_BUFFERED, FILE_ANY_ACCESS)
  FSCTL_DISMOUNT_VOLUME = (9 shl 16) or (0 shl 14) or (8 shl 2) or 0;
var
  BytesReturned: Cardinal;
begin
  Result := DeviceIoControl(AVolumeHandle, FSCTL_DISMOUNT_VOLUME, nil, 0,
    nil, 0, BytesReturned, nil);
  if not Result then
    RaiseLastWin32Error;
end;

function PreventRemovalOfVolume(AVolumeHandle: THandle;
  APreventRemoval: boolean): boolean;
const
// #define IOCTL_STORAGE_MEDIA_REMOVAL CTL_CODE(IOCTL_STORAGE_BASE, 0x0201, METHOD_BUFFERED, FILE_READ_ACCESS)
  IOCTL_STORAGE_MEDIA_REMOVAL = ($2d shl 16) or (1 shl 14) or ($201 shl 2) or 0;
type
  TPreventMediaRemoval = record
    PreventMediaRemoval: BOOL;
  end;
var
  BytesReturned: Cardinal;
  PMRBuffer: TPreventMediaRemoval;
begin
  PMRBuffer.PreventMediaRemoval := APreventRemoval;
  Result := DeviceIoControl(AVolumeHandle, IOCTL_STORAGE_MEDIA_REMOVAL,
    @PMRBuffer, SizeOf(TPreventMediaRemoval), nil, 0, BytesReturned, nil);
  if not Result then
    RaiseLastWin32Error;
end;

function AutoEjectVolume(AVolumeHandle: THandle): boolean;
const
// #define IOCTL_STORAGE_EJECT_MEDIA CTL_CODE(IOCTL_STORAGE_BASE, 0x0202, METHOD_BUFFERED, FILE_READ_ACCESS)
  IOCTL_STORAGE_EJECT_MEDIA = ($2d shl 16) or (1 shl 14) or ($202 shl 2) or 0;
var
  BytesReturned: Cardinal;
begin
  Result := DeviceIoControl(AVolumeHandle, IOCTL_STORAGE_EJECT_MEDIA, nil, 0,
    nil, 0, BytesReturned, nil);
  if not Result then
    RaiseLastWin32Error;
end;

function EjectVolume(ADrive: char): boolean;
var
  VolumeHandle: THandle;
begin
  Result := FALSE;
  // Open the volume
  VolumeHandle := OpenVolume(ADrive);
  if VolumeHandle = INVALID_HANDLE_VALUE then
    exit;
  try
    // Lock and dismount the volume
    if LockVolume(VolumeHandle) and DismountVolume(VolumeHandle) then begin
      // Set prevent removal to false and eject the volume
      if PreventRemovalOfVolume(VolumeHandle, FALSE) then
        AutoEjectVolume(VolumeHandle);
    end;
  finally
    // Close the volume so other processes can use the drive
    CloseHandle(VolumeHandle);
  end;
end;

procedure SaveDatasetToCSV(IncludeFieldNames : Boolean; Dataset : TDataset; Filename : String);
var
  TextWriter : TFDBatchMoveTextWriter;
  DataSetReader : TFDBatchMoveDataSetReader;
  BatchMove : TFDBatchMove;
begin
  BatchMove := nil;
  try
    BatchMove := TFDBatchMove.Create(nil);
    TextWriter := TFDBatchMoveTextWriter.Create(BatchMove);
    DataSetReader := TFDBatchMoveDataSetReader.Create(BatchMove);

    DataSetReader.DataSet := Dataset;

    TextWriter.FileName := Filename;
    TextWriter.DataDef.WithFieldNames := IncludeFieldNames;

    BatchMove.Reader := DataSetReader;
    BatchMove.Writer := TextWriter;
    BatchMove.Options := BatchMove.Options + [poClearDest]; //Overrides file if it already exists
    BatchMove.Execute;
  finally
    BatchMove.Free;
    ShowMessage('The file has been exported')
  end;
end;

procedure ImportCSVToDataset(IncludeFieldNames : Boolean; dDataset : TDataset; Filename : String);
var
  TextReader : TFDBatchMoveTextReader;
  DataSetReader : TFDBatchMoveDataSetReader;
  BatchMove : TFDBatchMove;
  // TFDBatchMoveTextReader

begin
  BatchMove := nil;
  try
    BatchMove := TFDBatchMove.Create(nil);
    TextReader := TFDBatchMoveTextReader.Create(BatchMove);
    //DataSetReader := TFDBatchMoveDataSetReader.Create(BatchMove);

    //DataSetReader.DataSet := Dataset;

    TextReader.FileName := Filename;
    TextReader.DataDef.WithFieldNames := IncludeFieldNames;
    TextReader.DataDef.Separator := ';';

    with TFDBatchMoveDataSetWriter.Create(BatchMove) do begin
       // Set destination dataset
       DataSet := dDataset;
       // Do not set Optimise to True, if dataset is attached to UI
       Optimise := False;
    end;

    BatchMove.GuessFormat;
    BatchMove.Execute;
    dDataset.Open;
  finally
    BatchMove.Free;
  end;
end;

procedure ImportFavorites(RDataSet : TDataSet; WDataSet : TFDQuery; DelFav : Boolean);
var
  i : Integer;
begin
// ## Reset Favorite List ##
   WDataSet.SQL.Text := 'UPDATE option SET TOP_ID =0 WHERE TOP_ID > 0 ';
   WDataSet.ExecSQL;

   RDataSet.First;
   WDataSet.SQL.Text := 'UPDATE option SET TOP_ID =:num WHERE ROM_NAME =:rom_name';

   while not RDataSet.Eof do // Read the MemTable with the Favorites Roms
   begin
    //i := WDataSet.RecordCount +1;
    WDataSet.ParamByName('num').AsInteger := RDataSet.FieldByName('TOP_ID').AsInteger; // Mark game Rom as Favorites
    WDataSet.ParamByName('rom_name').AsString := RDataSet.FieldByName('ROM_NAME').AsString;
    WDataSet.ExecSQL;
    RDataSet.Next;
   end;
   ShowMessage('The file has been Imported');
end;

end.


