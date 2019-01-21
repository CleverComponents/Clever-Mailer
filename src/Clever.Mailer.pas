{
  Copyright (C) 2016 by Clever Components

  Author: Sergey Shirokov <admin@clevercomponents.com>

  Website: www.CleverComponents.com

  This file is part of Clever Components Mailer.

  Clever Components Mailer is free software:
  you can redistribute it and/or modify it under the terms of
  the GNU Lesser General Public License version 3
  as published by the Free Software Foundation and appearing in the
  included file COPYING.LESSER.

  Clever Components Mailer is distributed in the hope
  that it will be useful, but WITHOUT ANY WARRANTY; without even the
  implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
  See the GNU Lesser General Public License for more details.

  You should have received a copy of the GNU Lesser General Public License
  along with Clever Components Mailer. If not, see <http://www.gnu.org/licenses/>.

  The current version of Clever Components Mailer needs for
  the non-free library Clever Internet Suite. This is a drawback,
  and we suggest the task of changing
  the program so that it does the same job without the non-free library.
  Anyone who thinks of doing substantial further work on the program,
  first may free it from dependence on the non-free library.
}

unit Clever.Mailer;

interface

uses
  System.Classes, System.SysUtils, System.IniFiles, Winapi.Windows,
  clMailUtils, clEmailAddress, clMailMessage, clDkim, clSmtp, clUtils, clTranslator,
  clEncoder, clWUtils;

type
  TMailer = class
  private
    FReportFile: string;
    FTextBodyFileName: string;
    FFrom: string;
    FSubject: string;
    FRecipientsFileName: string;
    FPort: Integer;
    FHtmlBodyFileName: string;
    FPassword: string;
    FHost: string;
    FUserName: string;
    FMailAgent: string;
    FDomain: string;
    FSelector: string;
    FKeyFile: string;
    FSignedHeaderFields: string;
    FUnsubscribeUrlTemplate: string;
    FDefaultRecipient: string;
    FDefaultSalutation: string;

    procedure CheckSentStatus;
    procedure CheckConfiguration;
    procedure LoadMessageData(ATextTemplate, AHtmlTemplate, ARecipients: TStrings);
    procedure SendMessageTo(ASmtp: TclSmtp; ADkim: TclDkim; ATextTemplate, AHtmlTemplate: TStrings; const ARecipient: string);
    procedure WriteReport(const AMessage: string);
    procedure LoginSmtp(ASmtp: TclSmtp);
    function CreateDkim: TclDkim;
    function GetConfigurationFile(const AConfigFile: string): string;
    function EncodeUrl(const AUrl, ACharSet: string): string;
    function EncodeSubscriber(const AEmail: string): string;
    function GetUnsubscribeUrl(const AEmail: string): string;
    function MergeRecipient(ABodyTemplate: TStrings; const ARecipientName, AUnsubscribeUrl: string): string;
    function CanAddListUnsubscribe(const AEmail: string): Boolean;
    procedure SetListUnsubscribe(AMessage: TclMailMessage; const AFromEmail, AUnsubscribeUrl: string);
  public
    constructor Create;

    function GetVersionInfo: string;
    function GetConfigFileFormat: string;

    procedure ReadConfiguration(const AConfigFile: string);
    procedure SendMessage;

    property ReportFile: string read FReportFile write FReportFile;

    property MailAgent: string read FMailAgent write FMailAgent;

    property Host: string read FHost write FHost;
    property Port: Integer read FPort write FPort;
    property UserName: string read FUserName write FUserName;
    property Password: string read FPassword write FPassword;

    property Domain: string read FDomain write FDomain;
    property Selector: string read FSelector write FSelector;
    property KeyFile: string read FKeyFile write FKeyFile;
    property SignedHeaderFields: string read FSignedHeaderFields write FSignedHeaderFields;

    property Subject: string read FSubject write FSubject;
    property From: string read FFrom write FFrom;
    property TextBodyFileName: string read FTextBodyFileName write FTextBodyFileName;
    property HtmlBodyFileName: string read FHtmlBodyFileName write FHtmlBodyFileName;

    property RecipientsFileName: string read FRecipientsFileName write FRecipientsFileName;
    property DefaultRecipient: string read FDefaultRecipient write FDefaultRecipient;
    property DefaultSalutation: string read FDefaultSalutation write FDefaultSalutation;

    property UnsubscribeUrlTemplate: string read FUnsubscribeUrlTemplate write FUnsubscribeUrlTemplate;
  end;

const
  MailerVersion = '3.5';
  MailerTimestamp = '26 Dec 2018';
  InetSuiteVersion = '9.3 development';
  DefaultReportFile = 'sentreport.txt';

implementation

{ TMailer }

constructor TMailer.Create;
begin
  inherited Create();

  FReportFile := DefaultReportFile;
  FPort := DefaultSmtpPort;
  FMailAgent := 'Clever Components Mailer';
end;

function TMailer.GetVersionInfo: string;
begin
  Result := 'Clever Components Mailer, version ' + MailerVersion + ', ' + MailerTimestamp + ', Clever Internet Suite ' + InetSuiteVersion;
end;

function TMailer.GetConfigFileFormat: string;
begin
  Result :=
    '[SMTP]'#13#10 +
    'Host=sample.com'#13#10 +
    'Port=25'#13#10 +
    'User=username'#13#10 +
    'Pass=secret'#13#10 +
    '[DKIM]'#13#10 +
    'Domain=sample.com'#13#10 +
    'Selector=example'#13#10 +
    'KeyFile=dkim_private_key.txt'#13#10 +
    'SignedFields=Date,From,To,Subject,MIME-Version,Content-Type'#13#10 +
    '[MAIL]'#13#10 +
    'Subj=The subject line'#13#10 +
    'From=John <john@sample.com>'#13#10 +
    'Text=NewsletterText.txt'#13#10 +
    'Html=NewsletterHtml.htm'#13#10 +
    'UnsubscribeUrl=http://sample.com/unsubscribe.asp?u='#13#10 +
    '[RECIPIENTS]'#13#10 +
    'List=subscribers_list.txt'#13#10 +
    'DefaultRecipient=Subscriber'#13#10 +
    'DefaultSalutation=Good day,';
end;

function TMailer.CanAddListUnsubscribe(const AEmail: string): Boolean;
var
  addr: string;
begin
  addr := LowerCase(AEmail);
  Result :=
    (Pos('@mail.ru', addr) = 0) and
    (Pos('@list.ru', addr) = 0) and
    (Pos('@bk.ru', addr) = 0) and
    (Pos('@inbox.ru', addr) = 0);
end;

procedure TMailer.CheckConfiguration;
begin
  if (Host = '') or
     (Subject = '') or
     (From = '') or
     ((TextBodyFileName = '') and (HtmlBodyFileName = '')) or
     (RecipientsFileName = '') then
  begin
    raise Exception.Create('The format of configuration file is invalid, use as follows: '#13#10 + GetConfigFileFormat());
  end;
end;

function TMailer.GetConfigurationFile(const AConfigFile: string): string;
begin
  Result := AConfigFile;
  if (ExtractFileName(Result) = Result) then
  begin
    Result := AddTrailingBackSlash(GetCurrentDir()) + Result;
  end;

  if (not FileExists(Result)) then
  begin
    raise Exception.Create('The specified configuration file doesn''t exist: '#13#10 + AConfigFile);
  end;
end;

procedure TMailer.ReadConfiguration(const AConfigFile: string);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(GetConfigurationFile(AConfigFile));
  try
    FHost := ini.ReadString('SMTP', 'Host', '');
    FPort := ini.ReadInteger('SMTP', 'Port', DefaultSmtpPort);
    FUserName := ini.ReadString('SMTP', 'User', '');
    FPassword := ini.ReadString('SMTP', 'Pass', '');

    FDomain := ini.ReadString('DKIM', 'Domain', '');
    FSelector := ini.ReadString('DKIM', 'Selector', '');
    FKeyFile := ini.ReadString('DKIM', 'KeyFile', '');
    FSignedHeaderFields := ini.ReadString('DKIM', 'SignedFields', '');

    FSubject := ini.ReadString('MAIL', 'Subj', '');
    FFrom := ini.ReadString('MAIL', 'From', '');
    FTextBodyFileName := ini.ReadString('MAIL', 'Text', '');
    FHtmlBodyFileName := ini.ReadString('MAIL', 'Html', '');
    FUnsubscribeUrlTemplate := ini.ReadString('MAIL', 'UnsubscribeUrl', '');

    FRecipientsFileName := ini.ReadString('RECIPIENTS', 'List', '');
    FDefaultRecipient := ini.ReadString('RECIPIENTS', 'DefaultRecipient', '');
    FDefaultSalutation := ini.ReadString('RECIPIENTS', 'DefaultSalutation', '');

    FReportFile := AddTrailingBackSlash(ExtractFilePath(ini.FileName)) + DefaultReportFile;
  finally
    ini.Free();
  end;
end;

procedure TMailer.CheckSentStatus;
begin
  if FileExists(ReportFile) then
  begin
    raise Exception.Create('The newsletter has been already sent, see ' + ReportFile);
  end;
end;

procedure TMailer.LoadMessageData(ATextTemplate, AHtmlTemplate, ARecipients: TStrings);
begin
  if FileExists(TextBodyFileName) then
  begin
    ATextTemplate.LoadFromFile(TextBodyFileName, TEncoding.UTF8);
  end;

  if FileExists(HtmlBodyFileName) then
  begin
    AHtmlTemplate.LoadFromFile(HtmlBodyFileName, TEncoding.UTF8);
  end;

  if (ATextTemplate.Count = 0) and (AHtmlTemplate.Count = 0) then
  begin
    raise Exception.Create('The message body is empty, provide text, html or both files');
  end;

  if FileExists(RecipientsFileName) then
  begin
    ARecipients.LoadFromFile(RecipientsFileName, TEncoding.UTF8);
  end;

  if (ARecipients.Count = 0) then
  begin
    raise Exception.Create('The recipient list is empty, specify at least one message recipient');
  end;
end;

//TODO replace it in clUriUtils
var
  UnsafeUriChars: PclChar = nil;
  UnsafeUriCharsCount: Integer = 0;

procedure InitStaticVars;
const
  Chars = ' <>"#%{}|\^~[]`/=*-._';
begin
  UnsafeUriCharsCount := TclTranslator.GetByteCount(Chars, 'us-ascii');
  if (UnsafeUriCharsCount > 0) then
  begin
    GetMem(UnsafeUriChars, UnsafeUriCharsCount + 1);
    try
      TclTranslator.GetBytes(Chars, UnsafeUriChars, UnsafeUriCharsCount + 1, 'us-ascii');
      UnsafeUriChars[UnsafeUriCharsCount] := #0;
    except
      FreeMem(UnsafeUriChars);
      UnsafeUriChars := nil;
      UnsafeUriCharsCount := 0;
    end;
  end;
end;

function TMailer.EncodeUrl(const AUrl, ACharSet: string): string;
  function IsUnsafeChar(ACharCode: TclChar): Boolean;
  var
    i: Integer;
  begin
    for i := 0 to UnsafeUriCharsCount - 1 do
    begin
      if (UnsafeUriChars[i] = ACharCode) then
      begin
        Result := True;
        Exit;
      end;
    end;
    Result := False;
  end;


  function IsHexDigit(c: TclChar): Boolean;
  begin
    Result := (c in ['0'..'9']) or (c in ['a'..'f']) or (c in ['A'..'F']);
  end;

var
  i, size: Integer;
  encBytes: PclChar;
begin
  Result := '';

  size := TclTranslator.GetByteCount(AUrl, ACharSet);
  if (size > 0) then
  begin
    GetMem(encBytes, size);
    try
      TclTranslator.GetBytes(AUrl, encBytes, size, ACharSet);

      i := 0;
      while (i < size) do
      begin
        if (encBytes[i] = '%') and (i + 2 < size)
          and IsHexDigit(encBytes[i + 1]) and IsHexDigit(encBytes[i + 2]) then
        begin
          Result := Result + '%' + string(encBytes[i + 1]) + string(encBytes[i + 2]);
          Inc(i, 2);
        end else
        if IsUnsafeChar(encBytes[i]) or (encBytes[i] >= #$7F) or (encBytes[i] < #$20) then
        begin
          Result := Result + '%' + IntToHex(Integer(encBytes[i]), 2);
        end else
        begin
          Result := Result + string(encBytes[i]);
        end;
        Inc(i);
      end;
    finally
      FreeMem(encBytes);
    end;
  end else
  begin
    Result := StringReplace(Trim(AUrl), #32, '%20', [rfReplaceAll]);
  end;
end;

function TMailer.EncodeSubscriber(const AEmail: string): string;
begin
  Result := TclEncoder.EncodeToString(AEmail, cmBase64);
  Result := EncodeUrl(Result, 'UTF-8');
end;

function TMailer.GetUnsubscribeUrl(const AEmail: string): string;
begin
  Result := UnsubscribeUrlTemplate;
  if (Result <> '') then
  begin
    Result := Result + EncodeSubscriber(AEmail);
  end;
end;

function TMailer.MergeRecipient(ABodyTemplate: TStrings; const ARecipientName, AUnsubscribeUrl: string): string;
var
  s: string;
begin
  Result := ABodyTemplate.Text;

  s := ARecipientName;
  if Trim(s) = '' then
  begin
    s := DefaultRecipient;
  end;
  Result := StringReplace(Result, ':RECIPIENT', s, [rfReplaceAll]);

  s := ARecipientName;
  if Trim(s) = '' then
  begin
    s := DefaultSalutation;
  end else
  begin
    s := 'Dear ' + s;
  end;
  Result := StringReplace(Result, ':SALUTATION', s, [rfReplaceAll]);

  Result := StringReplace(Result, '_unsubscribe_', AUnsubscribeUrl, [rfReplaceAll]);
end;

procedure TMailer.WriteReport(const AMessage: string);
var
  hFile: THandle;
  cnt: Cardinal;
  buf: TclByteArray;
begin
  hFile := CreateFile(PChar(ReportFile), GENERIC_WRITE, FILE_SHARE_READ or FILE_SHARE_WRITE, nil,
    OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);
  if (hFile = INVALID_HANDLE_VALUE) then
  begin
    hFile := CreateFile(PChar(ReportFile), GENERIC_WRITE, FILE_SHARE_READ or FILE_SHARE_WRITE, nil,
      CREATE_NEW, FILE_ATTRIBUTE_NORMAL, 0);
  end;
  if (hFile <> INVALID_HANDLE_VALUE) then
  begin
    SetFilePointer(hFile, 0, nil, FILE_END);

    buf := TclTranslator.GetBytes(AMessage, 'UTF-8');
    WriteFile(hFile, buf[0], Length(buf), cnt, nil);
    CloseHandle(hFile);
  end;
end;

procedure TMailer.LoginSmtp(ASmtp: TclSmtp);
begin
  ASmtp.Server := Host;
  ASmtp.Port := Port;
  ASmtp.UserName := UserName;
  ASmtp.Password := Password;
  ASmtp.MailAgent := MailAgent;

  ASmtp.Open();
end;

function TMailer.CreateDkim: TclDkim;
begin
  if (Domain = '') or (Selector = '') or (KeyFile = '') then
  begin
    Result := nil;
    Exit;
  end;

  Result := TclDkim.Create(nil);
  try
    Result.Canonicalization := 'relaxed/relaxed';
    Result.Domain := Domain;
    Result.Selector := Selector;

    if (SignedHeaderFields <> '') then
    begin
      Result.SignedHeaderFields.CommaText := SignedHeaderFields;
    end;

    Result.ImportPrivateKey(KeyFile);
  except
    Result.Free();
    raise;
  end;
end;

procedure TMailer.SendMessage;
var
  smtp: TclSmtp;
  textTemplate, htmlTemplate, recipients: TStrings;
  dkim: TclDkim;
  i: Integer;
begin
  CheckConfiguration();
  CheckSentStatus();

  textTemplate := nil;
  htmlTemplate := nil;
  recipients := nil;
  smtp := nil;
  dkim := nil;
  try
    textTemplate := TStringList.Create();
    htmlTemplate := TStringList.Create();
    recipients := TStringList.Create();

    LoadMessageData(textTemplate, htmlTemplate, recipients);

    smtp := TclSmtp.Create(nil);

    LoginSmtp(smtp);

    dkim := CreateDkim();

    for i := 0 to recipients.Count - 1 do
    begin
      SendMessageTo(smtp, dkim, textTemplate, htmlTemplate, recipients[i]);
    end;
  finally
    dkim.Free();
    smtp.Free();
    recipients.Free();
    htmlTemplate.Free();
    textTemplate.Free();
  end;
end;

procedure TMailer.SetListUnsubscribe(AMessage: TclMailMessage; const AFromEmail, AUnsubscribeUrl: string);
begin
  AMessage.ListUnsubscribe := '<mailto:' + AFromEmail + '?subject=unsubscribe>, <' + AUnsubscribeUrl + '>';
end;

procedure TMailer.SendMessageTo(ASmtp: TclSmtp; ADkim: TclDkim; ATextTemplate, AHtmlTemplate: TStrings; const ARecipient: string);
var
  text, html, unsubscrUrl: string;
  msg: TclMailMessage;
  recipient: TclEmailAddressItem;
begin
  Write('Sending to ''' + ARecipient + '''... ');
  WriteReport('Sending to ''' + ARecipient + '''... ');
  try
    msg := nil;
    recipient := nil;
    try
      msg := TclMailMessage.Create(nil);
      recipient := TclEmailAddressItem.Create();

      recipient.FullAddress := ARecipient;

      unsubscrUrl := GetUnsubscribeUrl(recipient.Email);

      text := MergeRecipient(ATextTemplate, recipient.Name, unsubscrUrl);
      html := MergeRecipient(AHtmlTemplate, recipient.Name, unsubscrUrl);

      msg.Dkim := ADkim;

      msg.CharSet := 'UTF-8';
      msg.BuildMessage(text, html);
      msg.Subject := Subject;
      msg.From.FullAddress := From;
      msg.ReplyTo := msg.From.FullAddress;
      msg.ToList.Add(ARecipient);

      if CanAddListUnsubscribe(recipient.Email) then
      begin
        SetListUnsubscribe(msg, msg.From.Email, unsubscrUrl);
      end;

      ASmtp.Send(msg);
    finally
      recipient.Free();
      msg.Free();
    end;

    Writeln('Done.');
    WriteReport('Done.'#13#10);
  except
    on E: Exception do
    begin
      Writeln('Exception: ', E.Message);
      WriteReport('Exception: ' + E.Message + #13#10);

      ASmtp.Close();
      ASmtp.Open();
    end;
  end;
end;

initialization
  InitStaticVars();

finalization
  FreeMem(UnsafeUriChars);

end.
