unit UnitAutoEMail;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.NumberBox,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient,
  IdExplicitTLSClientServerBase, IdMessageClient, IdSMTPBase, IdSMTP,
  Vcl.ExtCtrls, Vcl.Menus, IdMessage, IdIOHandler, IdIOHandlerSocket,
  IdIOHandlerStack, IdSSL, IdSSLOpenSSL;

type
  TFormMain = class(TForm)
    TrayIcon: TTrayIcon;
    IdSMTP: TIdSMTP;
    EditHost: TEdit;
    LabelHostname: TLabel;
    EditUsername: TEdit;
    LabelUsername: TLabel;
    EditPassword: TEdit;
    LabelPassword: TLabel;
    LabelPort: TLabel;
    EditPort: TNumberBox;
    LabelSecurity: TLabel;
    ComboBoxSecurity: TComboBox;
    PopupMenu: TPopupMenu;
    MenuItemQuit: TMenuItem;
    MenuItemShow: TMenuItem;
    LabelSubject: TLabel;
    EditSubject: TEdit;
    MemoBody: TMemo;
    LabelBody: TLabel;
    LabelInterval: TLabel;
    EditInterval: TNumberBox;
    LabelRecipient: TLabel;
    EditRecipient: TEdit;
    LabelMinutes: TLabel;
    CheckBoxEnabled: TCheckBox;
    Timer: TTimer;
    IdMessage: TIdMessage;
    IdSSLIOHandler: TIdSSLIOHandlerSocketOpenSSL;
    EditSender: TEdit;
    LabelSender: TLabel;
    procedure MenuItemQuitClick(Sender: TObject);
    procedure MenuItemShowClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure EditPortChangeValue(Sender: TObject);
    procedure EditPasswordChange(Sender: TObject);
    procedure EditUsernameChange(Sender: TObject);
    procedure EditHostChange(Sender: TObject);
    procedure ComboBoxSecurityChange(Sender: TObject);
    procedure CheckBoxEnabledClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure TimerTimer(Sender: TObject);
    procedure EditIntervalChange(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure EditSenderChange(Sender: TObject);
    procedure EditSubjectChange(Sender: TObject);
    procedure MemoBodyChange(Sender: TObject);
  end;

var
  FormMain: TFormMain;

implementation

{$R *.dfm}

uses
  Registry;

procedure TFormMain.CheckBoxEnabledClick(Sender: TObject);
begin
  Timer.Enabled := CheckBoxEnabled.Checked;
  if Timer.Enabled then
    TimerTimer(Sender);
end;

procedure TFormMain.ComboBoxSecurityChange(Sender: TObject);
begin
  case ComboBoxSecurity.ItemIndex of
    0:
      begin
        EditPort.ValueInt := 25;
      end;
    1:
      begin
        EditPort.ValueInt := 587;
      end;
    2:
      begin
        EditPort.ValueInt := 465;
      end;
  end;
end;

procedure TFormMain.EditHostChange(Sender: TObject);
begin
  IdSMTP.Host := EditHost.Text;
end;

procedure TFormMain.EditIntervalChange(Sender: TObject);
begin
  Timer.Interval := EditInterval.ValueInt * 60 * 1000;
end;

procedure TFormMain.EditPasswordChange(Sender: TObject);
begin
  IdSMTP.Password := EditPassword.Text;
end;

procedure TFormMain.EditPortChangeValue(Sender: TObject);
begin
  IdSMTP.Port := EditPort.ValueInt;
end;

procedure TFormMain.EditSenderChange(Sender: TObject);
begin
  IdMessage.From.Address := EditSender.Text;
  IdMessage.From.Name := EditSender.Text;
end;

procedure TFormMain.EditSubjectChange(Sender: TObject);
begin
  IdMessage.Subject := EditSubject.Text;
end;

procedure TFormMain.EditUsernameChange(Sender: TObject);
begin
  IdSMTP.Username := EditUsername.Text;
end;

procedure TFormMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caMinimize;
end;

procedure TFormMain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if CheckBoxEnabled.Checked then
  begin
    Application.Minimize;
    ShowWindow(Application.Handle, SW_HIDE);
  end;

  CanClose := not CheckBoxEnabled.Checked;
end;

procedure TFormMain.FormCreate(Sender: TObject);
begin
  with TRegistry.Create do
  begin
    try
      Rootkey := HKEY_CURRENT_USER;
      if OpenKey('Software\SimpleAutoEmail\AutoEmail', False) then
      begin
        ComboBoxSecurity.ItemIndex := ReadInteger('Security');
        EditHost.Text := ReadString('Host');
        EditUsername.Text := ReadString('Username');
        EditPassword.Text := ReadString('Password');
        EditPort.ValueInt := ReadInteger('Port');
        EditRecipient.Text := ReadString('Recipient');
        EditSubject.Text := ReadString('Subject');
        EditSender.Text := ReadString('Sender');
        MemoBody.Lines.Text := ReadString('Body');
        EditInterval.ValueInt := ReadInteger('Interval');
        CheckBoxEnabled.Checked := ReadBool('Enabled');
      end;
    finally
      Free;
    end;
  end;
end;

procedure TFormMain.FormDestroy(Sender: TObject);
begin
  with TRegistry.Create do
  begin
    try
      Rootkey := HKEY_CURRENT_USER;
      if OpenKey('Software\GTA\AutoEmail', True) then
      begin
        WriteString('Host', EditHost.Text);
        WriteString('Username', EditUsername.Text);
        WriteString('Password', EditPassword.Text);
        WriteInteger('Port', EditPort.ValueInt);
        WriteInteger('Security', ComboBoxSecurity.ItemIndex);
        WriteString('Sender', EditSender.Text);
        WriteString('Recipient', EditRecipient.Text);
        WriteString('Subject', EditSubject.Text);
        WriteString('Body', MemoBody.Lines.Text);
        WriteInteger('Interval', EditInterval.ValueInt);
        WriteBool('Enabled', CheckBoxEnabled.Checked);
      end;
    finally
      Free;
    end;
  end;
end;

procedure TFormMain.MemoBodyChange(Sender: TObject);
begin
  IdMessage.Body := MemoBody.Lines;
end;

procedure TFormMain.MenuItemQuitClick(Sender: TObject);
begin
  Application.Terminate;
end;

procedure TFormMain.MenuItemShowClick(Sender: TObject);
begin
  Show;
  Application.Restore;
end;

procedure TFormMain.TimerTimer(Sender: TObject);
begin
  IdMessage.Recipients.Add.Address := EditRecipient.Text;
  try
    IdSMTP.Connect;
    IdSMTP.Authenticate;
  except
    on E:Exception do
    begin
      MessageDlg('Cannot authenticate: ' + E.Message, mtWarning, [mbOK], 0);
      Exit;
    end;
  end;

  IdSMTP.Send(IdMessage);
end;

end.
