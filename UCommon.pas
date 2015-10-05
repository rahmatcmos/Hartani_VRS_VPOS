unit UCommon;
interface
uses
  Graphics,
  Windows;

type
    Str15    = String[15];
    Str20    = String[20];
    Str50    = String[50];
const
  AboutStr =
           'Retail System' + #13#10#13#10 +
           'Design and Develop by:'+#13#10+
           'SAHA System'+#13#10+
           'Jl Gebang Lor 69A'+#13#10+
           'Surabaya, 60111'+#13#10+
           'Telp: 6231-5998277';

const
  DefaultReg = '\Software\SAHA System\POS Retail\';

var //MailSlot
    TempNota           : LongInt;
    TempHargaTotal,
    TempUangTunai,
    TempBank           : Single;
    HdrMessage        : String = '00';
    HdrRequestStatus  : String = '01';
    HdrGantiHarga     : String = '02';
    HdrUpdateMaster   : String = '03';

var // Caption
    TitleAppActive    : String = '..:: SAHA POS ::..';

    ACountNota     : Double;
    ASumNota       : Double;
var
    OtomatUpdate : Integer = 1;
    DirektoriDatabase : String;
    DirektoriDBMaster : String;

    NamaServer   : String = '.';
    NamaLocal   : String = '11';
    StatusServer : Boolean;

var // Warna Template
    // Warna Utama
    WarnaForm       : TColor = $00454545;//clBlack;//$00E0E0E0;//$00E3ECEE;

var //Default System
    _ShortDateFormat   : String[20];//= 'dd-mmm-yyyy';
    _DateSeparator     : Char;//'-';
    _LongTimeFormat    : String[20];//'hh:nn:ss';
    _TimeSeparator     : Char;//= ':';
    _DecimalSeparator  : Char;//= ',';
    _ThousandSeparator : Char;//=  '.';
    _CurrencyString    : String[5];// 'Rp. ';
    _CurrencyFormat    : Integer;//= 2;

const
    NBil : Array[0..9] of String[8] =
           ('','Satu','Dua','Tiga','Empat','Lima',
            'Enam','Tujuh','Delapan','Sembilan');

    NBulan : Array[1..12] of String[10] =
             ('Januari',
              'Febuari',
              'Maret',
              'April',
              'Mei',
              'Juni',
              'Juli',
              'Agustus',
              'September',
              'Oktober',
              'November',
              'Desember');
    NBulanPendek : Array[1..12] of String[3] =
             ('Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun',
              'Jul', 'Agt', 'Sep', 'Okt', 'Nov', 'Des');

Procedure GetSystemDateTime;
Procedure SetSystemDateTime;
Procedure SetDefaultSystemDateTime;
Function LeadZero(_Data : Integer) : String;
Function NamaHariLengkap : String;
Function NamaHariPendek : String;
Function HeaderTimeMail : String;
Function GetSQLTime(ATime : TDateTime) : String;
Function GetSQLDate(ADate : TDateTime) : String;
Function GetSQLServerDate(ADate : TDateTime) : String;
function AddZero(Digit : Integer; Data : LongInt) : String;
Function AddOpostropi(_Data : String) : String;

function ExtractDay(ADate: TDateTime): Word;
function ExtractMonth(ADate: TDateTime): Word;
function ExtractYear(ADate: TDateTime): Word;
function ExtractHour(ADate: TDateTime): Word;
function ExtractMinute(ADate: TDateTime): Word;
function ExtractSecond(ADate: TDateTime): Word;


implementation
uses
    SysUtils;

function ExtractDay(ADate: TDateTime): Word;
var
  M, Y: Word;
begin
  DecodeDate(ADate, Y, M, Result);
end;

function ExtractMonth(ADate: TDateTime): Word;
var
  D, Y: Word;
begin
  DecodeDate(ADate, Y, Result, D);
end;

function ExtractYear(ADate: TDateTime): Word;
var
  D, M: Word;
begin
  DecodeDate(ADate, Result, M, D);
end;

function ExtractHour(ADate: TDateTime): Word;
var
  AHour, AMinute, ASecond, AMSec: Word;
begin
     DecodeTime(ADate, AHour, AMinute, ASecond, AMSec);
     Result := AHour;
end;

function ExtractMinute(ADate: TDateTime): Word;
var
  AHour, AMinute, ASecond, AMSec: Word;
begin
     DecodeTime(ADate, AHour, AMinute, ASecond, AMSec);
     Result := AMinute;
end;

function ExtractSecond(ADate: TDateTime): Word;
var
  AHour, AMinute, ASecond, AMSec: Word;
begin
     DecodeTime(ADate, AHour, AMinute, ASecond, AMSec);
     Result := ASecond;
end;

Procedure GetSystemDateTime;
begin
    _ShortDateFormat   := ShortDateFormat;
    _DateSeparator     := DateSeparator;
    _LongTimeFormat    := LongTimeFormat;
    _TimeSeparator     := TimeSeparator;
    _DecimalSeparator  := DecimalSeparator;
    _ThousandSeparator := ThousandSeparator;
    _CurrencyString    := CurrencyString;
    _CurrencyFormat    := CurrencyFormat;
end;

Procedure SetSystemDateTime;
begin
     ShortDateFormat   := 'dd-mmm-yyyy';
     DateSeparator     := '-';
     LongTimeFormat    := 'hh:nn:ss';
     TimeSeparator     := ':';
     DecimalSeparator  := ',';
     ThousandSeparator :=  '.';
     CurrencyString    := 'Rp. ';
     CurrencyFormat    := 2;
end;

Procedure SetDefaultSystemDateTime;
begin
    ShortDateFormat   := _ShortDateFormat;
    DateSeparator     := _DateSeparator;
    LongTimeFormat    := _LongTimeFormat;
    TimeSeparator     := _TimeSeparator;
    DecimalSeparator  := _DecimalSeparator;
    ThousandSeparator := _ThousandSeparator;
    CurrencyString    := _CurrencyString;
    CurrencyFormat    := _CurrencyFormat;
end;

Function LeadZero(_Data : Integer) : String;
begin
     if Length(IntToStr(_Data)) < 2 then
        Result := '0' + IntToStr(_Data)
     else
        Result := IntToStr(_Data);
end;

Function NamaHariLengkap : String;
var
   Temp            : String;
begin
     Temp := '';
     case DayOfWeek(Date) of
      1 : Temp := 'Minggu';
      2 : Temp := 'Senin';
      3 : Temp := 'Selasa';
      4 : Temp := 'Rabu';
      5 : Temp := 'Kamis';
      6 : Temp := 'Jumat';
      7 : Temp := 'Sabtu';
     else
       Temp := 'Minggu';
     end;
     Result := Temp + ', ' + IntToStr(ExtractDay(Date)) + ' - ' +
               NBulan[ExtractMonth(Date)] + ' - '+ IntToStr(ExtractYear(Date));
end;

Function NamaHariPendek : String;
var
   Temp            : String;
begin
     Temp := '';
     case DayOfWeek(Date) of
      1 : Temp := 'Minggu';
      2 : Temp := 'Senin';
      3 : Temp := 'Selasa';
      4 : Temp := 'Rabu';
      5 : Temp := 'Kamis';
      6 : Temp := 'Jumat';
      7 : Temp := 'Sabtu';
     else
       Temp := 'Minggu';
     end;
     Result := Temp + ', ' + IntToStr(ExtractDay(Date)) + ' - ' +
               NBulanPendek[ExtractMonth(Date)] + ' - '+ IntToStr(ExtractYear(Date));
end;

Function HeaderTimeMail : String;
var
   Temp : String;
begin
     Temp := '';
     case DayOfWeek(Date) of
      1 : Temp := 'Minggu';
      2 : Temp := 'Senin';
      3 : Temp := 'Selasa';
      4 : Temp := 'Rabu';
      5 : Temp := 'Kamis';
      6 : Temp := 'Jumat';
      7 : Temp := 'Sabtu';
     else
       Temp := 'Minggu';
     end;
     Result := LeadZero(ExtractDay(Date)) + '-' +
               LeadZero(ExtractMonth(Date)) + '-'+
               LeadZero(ExtractYear(Date)) + '/'+
               LeadZero(ExtractHour(Time)) +':'+
               LeadZero(ExtractMinute(Time)) +':'+
               LeadZero(ExtractSecond(Time));
end;

Function GetSQLTime(ATime : TDateTime) : String;
var
   AHour,
   AMinute,
   ASecond,
   AMSec      : Word;
begin
     DecodeTime(ATime, AHour, AMinute, ASecond, AMSec);
     Result := IntToStr(AHour)+':'+IntToStr(AMinute)+':'+IntToStr(ASecond);
end;

Function GetSQLDate(ADate : TDateTime) : String;
var
   AYear,
   AMonth,
   ADay  : Word;
begin
     DecodeDate(ADate, AYear, AMonth, ADay);
     Result := IntToStr(AMonth)+'-'+IntToStr(ADay)+'-'+IntToStr(AYear);
end;

Function GetSQLServerDate(ADate : TDateTime) : String;
var
   AYear,
   AMonth,
   ADay  : Word;
begin
     DecodeDate(ADate, AYear, AMonth, ADay);
     Result := IntToStr(AYear)+'-'+IntToStr(AMonth)+'-'+IntToStr(ADay);
end;

{Function untuk conversi angka ke kata}

Function NolKiri(X : String; n : byte) : String;
var
   p  : Integer;
   s  : String[15];
begin
     p := Length (X);
     if n > p then
     begin
          fillchar(s, (n-p)+1,'0');
          s[0] := chr(n-p);
          x := s+x;
     end;
     Nolkiri := x;
end;

Function NolTrim(x : Str15) : string;
begin
     if (x[0] <> #0) then
       while (x[1] = '0') and (x[0] <> #0) do
         delete(x,1,1);
     noltrim := x;
end;

function AddZero(Digit : Integer; Data : LongInt) : String;
var
   Temp : String;
   i    : Integer;
begin
     Temp := '';
     for i := 0 to Digit-length(IntToStr(Data))-1 do
         Insert('0', Temp, 0);
     Result := Temp+IntToStr(Data);
end;

Function AddOpostropi(_Data : String) : String;
const
     Opostropi = '''';
begin
     Result := Opostropi+_Data+Opostropi;
end;

end.
