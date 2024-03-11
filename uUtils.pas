unit uUtils;

interface

type TEmuladores = record
      Name : String;
      Id : Integer;
end;

const
  Emuladores: TArray<String> = ['FBA08', 'MAME37', 'MAME139', 'MAME78', 'FBA42',
    'PPSSPP', 'PS', 'N64', 'FBA42_FMY', 'FC', 'SFC', 'GBA', 'GBC','MD', 'PCE',
    'DC','MAME19','MAME199'];

  Genero: TArray<String> = ['Other', 'Fighting', 'Action', 'Shotting', 'Sport', 'Puzzle', 'Racing'];

  EmuCod : TArray<Integer> = [0,2,3,4,5,6,7,8,10,11,12,13,14,15,17,18,19,21];

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

implementation


end.
