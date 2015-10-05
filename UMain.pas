unit UMain;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Menus, ExtCtrls, StdCtrls, G10Reg, G10Phantom, G10StatusBar,
  rmAppEvents, G10CMailSlot, Winsock;

type
  TFrMain = class(TForm)
    MailPOS: TG10CSecureMail;
    PopupMenu1: TPopupMenu;
    Show1: TMenuItem;
    Minimized1: TMenuItem;
    N1: TMenuItem;
    Close1: TMenuItem;
    TUpdate: TTimer;
    AReg: TG10Reg;
    Phantom: TG10Phantom;
    StatusBar: TG10StatusBar;
    rmApplicationEvents1: TrmApplicationEvents;
    eLog: TMemo;
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure MailPOSMessageAvail(Sender: TObject; Msg: String);
    procedure Close1Click(Sender: TObject);
    procedure Show1Click(Sender: TObject);
    procedure TUpdateTimer(Sender: TObject);
    procedure eLogChange(Sender: TObject);
    procedure rmApplicationEvents1Exception(Sender: TObject; E: Exception);
    procedure Minimized1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    AFirstRun : Boolean;
    NamaKomputer : String;
    IPKomputer : String;
    Procedure SendToServer(AHeader : String; AMsg : String);
    Procedure BrodcastData(AHeader : String; AMsg : String);
    procedure SetRegister;
  end;

var
  FrMain: TFrMain;

implementation

uses
    UCommon,
    UModulTemp;

{$R *.DFM}

procedure GetHostInfo(var Name, Address: string);
var
  WSAData: TWSAData;
  HostEnt: PHostEnt;
begin
  { no error checking...}
  WSAStartup(2, WSAData);
  SetLength(Name, 255);
  Gethostname(PChar(Name), 255);
  SetLength(Name, StrLen(PChar(Name)));
  HostEnt := gethostbyname(PChar(Name));
  with HostEnt^  do
    Address := Format('%d.%d.%d.%d',[
      Byte(h_addr^[0]),
      Byte(h_addr^[1]),
      Byte(h_addr^[2]),
      Byte(h_addr^[3])]);
  WSACleanup;
end;

Procedure TFrMain.SendToServer(AHeader : String; AMsg : String);
begin
     SendSecureMail(NamaServer, 'MailSvr', AHeader+NamaKomputer+';'+ AMsg);
end;

Procedure TFrMain.BrodcastData(AHeader : String; AMsg : String);
begin
     SendSecureMail('*', 'MailSvr', AHeader+NamaKomputer+';'+ AMsg);
end;

Function GetReg(AType : Byte; APath : String; AKey : String; ADefault : Variant) : Variant;
begin
     case AType  of
      1 : begin
               try
                   Result := FrMain.AReg.RInteger(APath, AKey);
                except
                      FrMain.AReg.WInteger(APath, AKey, ADefault);
                end;
                if Length(Result) <=0 then
                begin
                     FrMain.AReg.WInteger(APath, AKey, ADefault);
                     Result := FrMain.AReg.RInteger(APath, AKey);
                end;
          end;
      2 : begin //String;
                try
                   Result := FrMain.AReg.RDString(APath, AKey);
                except
                      FrMain.AReg.WEString(APath, AKey, ADefault);
                end;
                if Length(Result) <=0 then
                begin
                     FrMain.AReg.WEString(APath, AKey, ADefault);
                     Result := FrMain.AReg.RDString(APath, AKey);
                end;
          end;
      3 : begin //String;
                try
                   Result := FrMain.AReg.RString(APath, AKey);
                except
                   FrMain.AReg.WString(APath, AKey, ADefault);
                end;
                if Length(Result) <=0 then
                begin
                     FrMain.AReg.WString(APath, AKey, ADefault);
                     Result := FrMain.AReg.RString(APath, AKey);
                end;
          end;
     end;
end;

procedure TFrMain.SetRegister;
begin
     AReg.Active := True;
     OtomatUpdate := GetReg(1, 'Setting','Otomat Update', 5);
     NamaLocal    := GetReg(3, 'Setting', 'Nama Local', '11');
     NamaServer   := GetReg(3, 'Setting', 'SERVER', 'SERVER');
end;

procedure TFrMain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
     CanClose := False;
     SetDefaultSystemDateTime;
     SendToServer('04;', 'Application Terminated');
     CanClose := True;
end;

procedure TFrMain.FormCreate(Sender: TObject);
begin
     AFirstRun := True; 
     SetRegister;
     Phantom.MousePosX := Screen.Width;
     Phantom.MousePosY := Screen.Height;
     FrMain.Left := Screen.Width - FrMain.Width;
     FrMain.Top := Screen.Height - FrMain.Height-20;
     WindowState := wsMinimized;
     Phantom.Visible := False;

     GetSystemDateTime;
     SetSystemDateTime;
     Application.Title := TitleAppActive;
     FrMain.Caption := TitleAppActive;
     MailPOS.Active := True;
     TUpdate.Interval := OtomatUpdate * 10000;
     TUpdate.Enabled := True;
     //RunOnStartup.SetRunOnStartup('Point of Sales', Application.ExeName, False, True);
     TempNota := 0;
     TempHargaTotal := 0.00;
     TempUangTunai  := 0.00;
     TempBank       := 0.00;
     NamaKomputer := 'LOCALHOST';
     GetHostInfo(NamaKomputer, IPKomputer);
     StatusBar.Panels[0].Text := 'Computer Name = '+ NamaKomputer +' ('+IPKomputer+')';
     StatusBar.Panels[1].Text := 'Timer = ' + FloatToStr(TUpdate.Interval/1000) + ' seconds';
     BrodcastData('03;', 'Otomat Timer = ' + FloatToStr(TUpdate.Interval/1000) + ' seconds. Local Time = ' + HeaderTimeMail);
     eLog.Lines.Insert(0, HeaderTimeMail +' : '+ 'Initialized Broadcast');
end;

procedure TFrMain.MailPOSMessageAvail(Sender: TObject; Msg: String);
var
   i           : Integer;
   AParams     : Array[1..15] of String;
   ParamKe     : Integer;
   TempPesan   : String;
begin
     ParamKe := 1;
     TempPesan := Msg;
     for i := 1 to 15 do AParams[i] := '';
     for i := 1 to length(Msg) do
     begin
          if TempPesan[1] = ';' then
          begin
               ParamKe := ParamKe + 1;
               Delete(TempPesan, 1, 1);
          end
          else
          begin
               AParams[ParamKe] := AParams[ParamKe] + TempPesan[1];
               Delete(TempPesan, 1, 1);
          end;
     end;
     case StrToInt(AParams[1]) of
     10 : begin
               AReg.WInteger('Setting', 'Otomat Update', StrToInt(AParams[2]));
               TUpdate.Interval := StrToInt(AParams[2]) * 10000;
               TUpdate.Enabled := True;
               SendToServer('01;', 'Timer Set '+ FloatToStr(TUpdate.Interval/1000) +'. Local Time : ' + HeaderTimeMail);
               StatusBar.Panels[1].Text := 'Timer = ' + FloatToStr(TUpdate.Interval/1000) + ' seconds';
          end;
     //12 : RunOnStartup.SetRunOnStartup('Point of Sales', Application.ExeName, False, True);
     //13 : RunOnStartup.SetRunOnStartup('Point of Sales', Application.ExeName, False, False);
     20 : begin
               TUpdate.Enabled := False;
               if LowerCase(AParams[12]) = 'true' then
               begin
                    dmTemp.GantiDataMaster( AParams[2],
                                            AParams[3],
                                            AParams[4],
                                            AParams[5],
                                            AParams[6],
                                            AParams[7],
                                            AParams[8],
                                            AParams[9],
                                            AParams[10],
                                            AParams[11]);

               end
               else
               begin
                    dmTemp.DeleteMaster(StrToInt(AParams[2]), AParams[6]);
               end;
               TUpdate.Enabled := True;
          end;
     21 : begin // Set Server
               StatusServer := True; 
               NamaServer := AParams[2];
               NamaLocal := AParams[3];
               FrMain.AReg.WString('Setting', 'Nama Local', NamaLocal);
               FrMain.AReg.WString('Setting', 'Server', NamaServer);
               StatusBar.Panels[2].Text := 'Server = ' + NamaServer +' (up)';
          end;
     22 : begin // Set Server UP
               StatusServer := True;
               NamaServer := AParams[2];
               FrMain.AReg.WString('Setting', 'Server', NamaServer);
               SendToServer('01;', 'Online. Timer update '+ FloatToStr(TUpdate.Interval/1000) +'. Local Time : ' + HeaderTimeMail);
               StatusBar.Panels[2].Text := 'Server = ' + NamaServer +' (up)';
          end;
     23 : begin // Set Server Down
               StatusServer := False;
               StatusBar.Panels[2].Text := 'Server = ' + NamaServer +' (down)';
          end;
     24 : begin // Update Data when Server Up fo first Time
               StatusServer := True; 
               NamaServer := AParams[2];
               NamaLocal := AParams[3];
               FrMain.AReg.WString('Setting', 'Nama Local', NamaLocal);
               FrMain.AReg.WString('Setting', 'Server', NamaServer);
               StatusBar.Panels[2].Text := 'Server = ' + NamaServer +' (up)';
               SendToServer('01;', 'Online. Timer update '+ FloatToStr(TUpdate.Interval/1000) +'. Local Time : ' + HeaderTimeMail);
          end;
     end;
end;

procedure TFrMain.Close1Click(Sender: TObject);
begin
     Close;
end;

procedure TFrMain.Show1Click(Sender: TObject);
begin
     FrMain.Show;
     FrMain.WindowState := wsNormal;
end;

procedure TFrMain.TUpdateTimer(Sender: TObject);
begin
     if StatusServer = True then
        dmTemp.GenerateStatus;
end;

procedure TFrMain.eLogChange(Sender: TObject);
begin
     if eLog.Lines.Count >= 200 then
        eLog.Lines.Delete(200);
end;

procedure TFrMain.rmApplicationEvents1Exception(Sender: TObject;
  E: Exception);
begin
     SendToServer('01;', 'Exception : ' + E.Message);
end;

procedure TFrMain.Minimized1Click(Sender: TObject);
begin
     FrMain.Hide;
     FrMain.WindowState := wsMinimized;
end;

end.



