unit UModulTemp;

interface

uses
  Windows, Messages, SysUtils, Classes, Controls, Forms, Dialogs, KDaoTable,
  KDaoDataBase, Db;

type
  TdmTemp = class(TDataModule)
    QAll: TKADaoTable;
    QMax: TKADaoTable;
    DBMaster: TKADaoDatabase;
    DBTemp: TKADaoDatabase;
    QStatus: TKADaoTable;
  private
    { Private declarations }
  public
    { Public declarations }
    Function GantiDataMaster( AParams1,
                                  AParams2,
                                  AParams3,
                                  AParams4,
                                  AParams5,
                                  AParams6,
                                  AParams7,
                                  AParams8,
                                  AParams9,
                                  AParams10 : String) : LongInt;
    Function DeleteMaster(ANoID : LongInt; ANama : String) : LongInt;
    Function GenerateStatus : String;
    procedure SetDBActive;
    procedure SetDBMasterActive;
    procedure SetTableActive(Sender: TObject; _TableName : String);
    procedure SetQuery(AStatus : Boolean; ASQL : String; _TableName : TKADaoTable);
    Function SetExecSQL(AStatus : Boolean; ASQL : String; ATable : TKADaoTable) : LongInt;
  end;

var
  dmTemp: TdmTemp;

implementation

uses
    UCommon, UMain;

{$R *.DFM}

const
     Opostropi = '''';

function IsLeapYear(AYear: Integer): Boolean;
begin
     Result := (AYear mod 4 = 0) and ((AYear mod 100 <> 0) or (AYear mod 400 = 0));
end;

function DaysPerMonth(AYear, AMonth: Integer): Integer;
const
  DaysInMonth: array[1..12] of Integer =
    (31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);
begin
  Result := DaysInMonth[AMonth];
  if (AMonth = 2) and IsLeapYear(AYear) then Inc(Result); { leap-year Feb is special }
end;

function IncDate(ADate: TDateTime; Days, Months, Years: Integer): TDateTime;
var
  D, M, Y: Word;
  Day, Month, Year: Longint;
begin
  DecodeDate(ADate, Y, M, D);
  Year := Y; Month := M; Day := D;
  Inc(Year, Years);
  Inc(Year, Months div 12);
  Inc(Month, Months mod 12);
  if Month < 1 then begin
    Inc(Month, 12);
    Dec(Year);
  end
  else if Month > 12 then begin
    Dec(Month, 12);
    Inc(Year);
  end;
  if Day > DaysPerMonth(Year, Month) then Day := DaysPerMonth(Year, Month);
  Result := EncodeDate(Year, Month, Day) + Days + Frac(ADate);
end;

Function TdmTemp.GantiDataMaster( AParams1, AParams2, AParams3, AParams4, AParams5,
                                  AParams6, AParams7, AParams8, AParams9, AParams10 : String) : LongInt;
const
     SQLuTInv = 'UPDATE TInv SET '+
                       'TInv.IDInventor = %s, TInv.Kode = ''%s'', TInv.Barcode = ''%s'', TInv.Nama = ''%s'', '+
                       'TInv.IDSatuan = %s, TInv.KodeSat = ''%s'', TInv.Konversi = %s, TInv.HargaJual = %s, '+
                       'TInv.HargaPokok = %s '+
                'WHERE TInv.NoID=%s';
     SQLaTInv = 'INSERT INTO TInv ( '+
                       'NoID, IDInventor, Kode, Barcode, Nama, IDSatuan, KodeSat, Konversi, HargaJual, HargaPokok ) '+
                'SELECT %s AS ANoID, %s AS AIDInv, ''%s'' AS AKode, ''%s'' AS ABarcode, ''%s'' AS ANama, %s AS Expr6, ''%s'' AS Expr7, %s AS Expr8, %s AS Expr9, %s AS Expr10';
var
   ASQL : String;
   StatusBefore : Boolean;
begin
     StatusBefore := DBMaster.Connected;
     if not StatusBefore then
     begin
          SetDBMasterActive;
     end;
     SetDefaultSystemDateTime;
     ASQL := Format(SQLuTInv, [ AParams2,
                                AParams3,
                                AParams4,
                                AParams5,
                                AParams6,
                                AParams7,
                                AParams8,
                                AParams9,
                                AParams10,
                                AParams1]);
     Result := SetExecSQL(True, ASQL, dmTemp.QAll);
     if Result <= 0 then
     begin
          ASQL := Format(SQLaTInv, [ AParams1,
                                     AParams2,
                                     AParams3,
                                     AParams4,
                                     AParams5,
                                     AParams6,
                                     AParams7,
                                     AParams8,
                                     AParams9,
                                     AParams10]);
          Result := SetExecSQL(True, ASQL, dmTemp.QAll);
          FrMain.SendToServer('01;', 'Append Inventor '+ AParams5 + ', ' +IntToStr(Result) +' record(s)');
          FrMain.eLog.Lines.Insert(0, HeaderTimeMail +' : '+ 'Append Inventor ''''' + AParams5 + ''''' OK. '+ IntToStr(Result) +' record(s)');
     end
     else
     begin
          FrMain.SendToServer('01;', 'Change Inventor '+ AParams5 + ', ' +IntToStr(Result) +' record(s)');
          FrMain.eLog.Lines.Insert(0, HeaderTimeMail +' : '+ 'Change Inventor ''''' + AParams5 + ''''' OK. '+ IntToStr(Result) +' record(s)');
     end;
     SetSystemDateTime;
     DBMaster.Connected := False;
end;

Function TdmTemp.DeleteMaster(ANoID : LongInt; ANama : String) : LongInt;
const
     SQLuTInv = 'DELETE TInv.* '+
                'FROM TInv '+
                'WHERE TInv.NoID = %d';
var
   ASQL : String;
   StatusBefore : Boolean;
begin
     StatusBefore := DBMaster.Connected;
     if not StatusBefore then
     begin
          SetDBMasterActive;
     end;
     SetDefaultSystemDateTime;
     ASQL := Format(SQLuTInv, [ANoID]);
     Result := SetExecSQL(True, ASQL, dmTemp.QAll);
     if Result > 0 then
     begin
          FrMain.SendToServer('01;', 'Delete Inventor '+ ANama + ', ' +IntToStr(Result) +' record(s)');
          FrMain.eLog.Lines.Insert(0, HeaderTimeMail +' : '+ 'Delete Inventor ''''' + ANama + ''''' OK. '+ IntToStr(Result) +' record(s)');
     end;
     SetSystemDateTime;
     DBMaster.Connected := False;
end;

Function TdmTemp.GenerateStatus : String;
const
     SQLSales = 'SELECT Count(MSales.NoID) AS CountOfNoID, Sum(MSales.HargaTotal) AS SumOfHargaTotal, Sum(MSales.UangMuka) AS SumOfUangMuka, Sum(MSales.Bank) AS SumOfBank '+
                'FROM MSales '+
                'WHERE (MSales.Tanggal >= #%s#) AND (MSales.Tanggal < #%s#)';
begin
     if not DBTemp.Connected then SetDBActive;
     SetDefaultSystemDateTime;
     QStatus.Active := False;
     QStatus.SQL.Clear;
     QStatus.SQL.Add(Format(SQLSales, [GetSQLDate(Now), GetSQLDate(IncDate(Date, 1, 0,0))]));
     QStatus.Active := True;
     if (TempNota       <> QStatus.Fields[0].AsInteger) OR
        (TempHargaTotal <> QStatus.Fields[1].AsFloat) OR
        (TempUangTunai  <> QStatus.Fields[2].AsFloat) OR
        (TempBank       <> QStatus.Fields[3].AsFloat) then
     begin
          TempNota       := QStatus.Fields[0].AsInteger;
          TempHargaTotal := QStatus.Fields[1].AsFloat;
          TempUangTunai  := QStatus.Fields[2].AsFloat;
          TempBank       := QStatus.Fields[3].AsFloat;
          Result := QStatus.Fields[0].AsString+';'+
                    QStatus.Fields[1].AsString+';'+
                    QStatus.Fields[2].AsString+';'+
                    QStatus.Fields[3].AsString;
          FrMain.SendToServer('02;', Result);
          FrMain.eLog.Lines.Insert(0, HeaderTimeMail +' : '+ 'Update data : Nota = '+ QStatus.Fields[0].AsString+
                                               ', Total = ' + QStatus.Fields[1].AsString+
                                               ', Tunai = ' + QStatus.Fields[2].AsString+
                                               ', Bank = '+ QStatus.Fields[3].AsString);
     end;
     QStatus.Active := False;
     SetSystemDateTime;
     DBTemp.Connected := False;
end;

Function TdmTemp.SetExecSQL(AStatus : Boolean; ASQL : String; ATable : TKADaoTable) : LongInt;
var
   TSQL : TStringList;
begin
     Result := 0;
     if AStatus then
     begin
          TSQL := TStringList.Create;
          TSQL.Add(ASQL);
          Result := (ATable as TKADaoTable).ExecSQL(TSQL);
          TSQL.Free;
     end
     else
     begin
          (ATable as TKADaoTable).Active := False;
     end;
end;

procedure TdmTemp.SetTableActive(Sender: TObject; _TableName : String);
begin
     Application.ProcessMessages;
     (Sender As TKADaoTable).Active := False;
     (Sender As TKADaoTable).TableName := _TableName;
     (Sender As TKADaoTable).Active := True;
end;

procedure TdmTemp.SetQuery(AStatus : Boolean; ASQL : String; _TableName : TKADaoTable);
begin
     Application.ProcessMessages;
     (_TableName As TKADaoTable).Active := False;
     (_TableName As TKADaoTable).SQL.Clear;
     (_TableName As TKADaoTable).SQL.Add(ASQL);
     (_TableName As TKADaoTable).Active := True;
end;

procedure TdmTemp.SetDBActive;
begin
     Application.ProcessMessages;
     DirektoriDatabase := ExtractFileDir(Application.ExeName) + '\Database\TempDB.mdb';
     if FileExists(DirektoriDatabase) then
     begin
          DBTemp.Connected := False;
          DBTemp.Database := DirektoriDatabase;
          DBTemp.DatabasePassword := 'andrewblack';
          try
          DBTemp.Connected := True;
          except
          end;
          if not DBTemp.Connected then
             FrMain.SendToServer('01;', 'Database failed to connect');
     end
     else
         FrMain.SendToServer('01;', 'Database Temporary not Found');
end;

procedure TdmTemp.SetDBMasterActive;
begin
     Application.ProcessMessages;
     DirektoriDBMaster := ExtractFileDir(Application.ExeName) + '\Database\DBMaster.mdb';
     if FileExists(DirektoriDBMaster) then
     begin
          DBMaster.Connected := False;
          DBMaster.Database := DirektoriDBMaster;
          DBMaster.DatabasePassword := 'andrewblack';
          try
             DBMaster.Connected := True;
          except
          end;
          if not DBMaster.Connected then
             FrMain.SendToServer('01;', 'Database Master failed to connect');
     end
     else
         FrMain.SendToServer('01;', 'Database Master not Found');
end;

end.
