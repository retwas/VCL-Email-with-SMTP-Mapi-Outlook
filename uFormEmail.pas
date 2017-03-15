unit uFormEmail;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls;

type
   TDataMessage = record
      SendTo: string;
      SendCC: string;
      SendBCC: string;
      Sender: string;
      MailObject: string;
      MailMessage: string;
   end;

   IMessageSender = interface
   ['{E76ABC78-18BF-48C5-9D48-21CDFC56A4C1}']
      /// <summary>pour envoyer un message en SMTP, en MAPI, ou avec Outlook</summary>
      procedure SendMessage;
   end;

   IMessageData = interface
   ['{83D98D4F-45D2-48E1-9C70-23C62EF24014}']
      /// <summary>utilisé pour récupérer les données</summary>
      function GetDataMessage: TDataMessage;
   end;

   IMessageDataSMTP = interface
   ['{83D98D4F-45D2-48E1-9C70-23C62EF24014}']
      /// <summary>pour récupérer les données spécifiques à l'envoi smtp</summary>
      procedure SetLogin(const aLogin: string);
      procedure SetPassword(const aPassword: string);
      procedure SetServer(const aServer: string);
      procedure SetPort(const aPort: integer);
      function GetLogin: string;
      function GetPassword: string;
      function GetServer: string;
      function GetPort: integer;
   end;

   TMessageSender = class(TInterfacedObject, IMessageSender, IMessageData)
   strict private
      FDataMessage: TDataMessage;
      function GetDataMessage: TDataMessage;
   public
      constructor Create(aDataMessage: TDataMessage);
      procedure SendMessage; virtual; abstract;

      property DataMessage: TDataMessage read GetDataMessage;
   end;

   TMessageSenderSMTP = class(TInterfacedObject, IMessageSender, IMessageData, IMessageDataSMTP)
   strict private
      FDataMessage: TDataMessage;
      FLogin: string;
      FPassword: string;
      FServer: string;
      FPort: integer;

      procedure SetLogin(const aLogin: string);
      procedure SetPassword(const aPassword: string);
      procedure SetServer(const aServer: string);
      procedure SetPort(const aPort: integer);
      function GetServer: string;
      function GetPort: integer;
      function GetLogin: string;
      function GetPassword: string;
      function GetDataMessage: TDataMessage;
   public
      constructor Create(aDataMessage: TDataMessage);
      procedure SendMessage; virtual; abstract;

      property DataMessage: TDataMessage read GetDataMessage;
      property Login:       string       read GetLogin    write SetLogin;
      property Password:    string       read GetPassword write SetPassword;
      property Server:      string       read GetServer   write SetServer;
      property Port:        integer      read GetPort     write SetPort;
   end;

   TMessageMapi = class(TMessageSender)
   public
      procedure SendMessage; override;
   end;

   TMessageOutlook = class(TMessageSender)
   public
      procedure SendMessage; override;
   end;

   TMessageSMTP = class(TMessageSenderSMTP)
   public
      procedure SendMessage; override;
   end;

  TFormEmail = class(TForm)
    edTo: TEdit;
    edCC: TEdit;
    edBCC: TEdit;
    edObject: TEdit;
    Memo: TMemo;
    RadioGroup: TRadioGroup;
    btnEnvoyer: TButton;
    edPassword: TEdit;
    edLogin: TEdit;
    edSender: TEdit;
    procedure btnEnvoyerClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  public
    { Déclarations publiques }
  end;

var
  FormEmail: TFormEmail;

implementation

uses
   // SMTP
   IdSMTP, IdExplicitTLSClientServerBase, IdMessage, IdSSLOpenSSL,
   // Outlook
   System.Win.ComObj,
   // MAPI
   Winapi.Mapi,
   Math, System.UITypes;

type
   ESMTPError    = class(Exception);
   EOutlookError = class(Exception);
   EMAPIError    = class(Exception);

const
   // correspondance avec le RadioGroup
   CST_SMTP    = 0;
   CST_OUTLOOK = 1;
   CST_MAPI    = 2;

{$R *.dfm}

procedure TFormEmail.btnEnvoyerClick(Sender: TObject);
var
   SendMessage: IMessageSender;
   DataMessage: TDataMessage;
   Cursor     : TCursor;
begin
   Cursor := Screen.Cursor;

   try
      Screen.Cursor := crHourGlass;

      try
         DataMessage.SendTo      := edTo.Text;
         DataMessage.SendCC      := edCC.Text;
         DataMessage.SendBCC     := edBCC.Text;
         DataMessage.Sender      := edSender.Text;
         DataMessage.MailObject  := edObject.Text;
         DataMessage.MailMessage := Memo.Text;

         case RadioGroup.ItemIndex of
            CST_SMTP :
            begin
               SendMessage := TMessageSMTP.Create(DataMessage);
               (SendMessage as IMessageDataSMTP).SetLogin(edLogin.Text);
               (SendMessage as IMessageDataSMTP).SetPassword(edPassword.Text);
               (SendMessage as IMessageDataSMTP).SetServer('smtp.office365.com');
               (SendMessage as IMessageDataSMTP).SetPort(587);
            end;
            CST_OUTLOOK : SendMessage := TMessageOutlook.Create(DataMessage);
            else
               SendMessage := TMessageMapi.Create(DataMessage);
         end;

         SendMessage.SendMessage;
      except
         on E : Exception do
            MessageDlg('Erreur lors de l''envoi de l''email.' + #10#13 + e.Message, mtError, [mbOK], 0);
      end;
   finally
      Screen.Cursor := Cursor;
   end;
end;

{ TMessageSMTP }

procedure TMessageSMTP.SendMessage;
var
   IdSMTP    : TIdSMTP;
   IdSSLIO   : TIdSSLIOHandlerSocketOpenSSL;
   IdMessage : TIdMessage;
begin
   inherited;

   // création des composants Indy
   IdSMTP    := TIdSMTP.Create(nil);
   IdMessage := TIdMessage.Create(nil);
   IdSSLIO   := TIdSSLIOHandlerSocketOpenSSL.Create(nil);

   try
      // adresse et port du serveur SMTP
      IdSMTP.AuthType := satDefault;
      IdSMTP.Host     := Server;
      IdSMTP.Port     := Port;

      // utilisation du mode sécurisé (ou pas si l'envoi en anonyme est autorisé)
      if (Login <> '') and (Password <> '') then
      begin
         IdSSLIO.SSLOptions.Method := sslvTLSv1;
         IdSMTP.Username  := Login;
         IdSMTP.Password  := Password;
         IdSMTP.IOHandler := IdSSLIO;
         IdSMTP.UseTLS    := utUseExplicitTLS;
      end;

      // connexion au serveur
      try
         IdSMTP.Connect;
      except
         on e : Exception do
            raise ESMTPError.Create('Erreur de connexion au serveur SMTP' + #10#13 + e.Message);
      end;

      IdMessage.Clear;

      // paramétrage de l'expéditeur
      IdMessage.From.Text           := DataMessage.Sender;
      IdMessage.ReplyTo.Add.Address := DataMessage.Sender;

      // ajout des destinataires
      if DataMessage.SendTo <> '' then
         IdMessage.Recipients.Add.Address := DataMessage.SendTo;

      if DataMessage.SendCC <> '' then
         IdMessage.CCList.Add.Address     := DataMessage.SendCC;

      if DataMessage.SendBCC <> '' then
         IdMessage.BccList.Add.Address    := DataMessage.SendBCC;

      // objet du message
      IdMessage.Subject := DataMessage.MailObject;

      // paramétrage de la date et de la priorité
      IdMessage.Date     := Now;
      IdMessage.Priority := mpNormal;

      // il est possible de paramétrer un l'accusé de lecture
      // idMessage.ReceiptRecipient.Address := adresse de l'expediteur;

      // corps du message
      IdMessage.Body.Text := DataMessage.MailMessage;

      try
         IdSMTP.Send(IdMessage);
      except
         on e : Exception do
            raise ESMTPError.Create('Erreur lors de l''envoi de l''email.' + #10#13 + e.Message);
      end;
   finally
      IdSMTP.Disconnect;

      FreeAndNil(IdSSLIO);
      FreeAndNil(IdMessage);
      FreeAndNil(IdSMTP);
   end;
end;

{ TMessageMapi }

procedure TMessageMapi.SendMessage;
var
   MapiMessage    : TMapiMessage;
   MapiDest       : PMapiRecipDesc;
   MapiExpediteur : TMapiRecipDesc;
   MAPI_Session   : Cardinal;
   MapiResult     : Cardinal;
   MAPIError      : DWord;
begin
   inherited;

   MapiDest := nil;

   try
      // FillChar permet de remplir une suite d'octets avec une valeur, ici 0
      FillChar(MapiMessage, Sizeof(TMapiMessage), 0);
      MapiMessage.lpszSubject  := PAnsiChar(AnsiString(DataMessage.MailObject));
      MapiMessage.lpszNoteText := PAnsiChar(AnsiString(DataMessage.MailMessage));

      // même traitement pour paramétrer l'expéditeur
      FillChar(MapiExpediteur, Sizeof(TMapiRecipDesc), 0);
      MapiExpediteur.lpszName    := PAnsiChar(AnsiString(DataMessage.Sender));
      MapiExpediteur.lpszAddress := PAnsiChar(AnsiString(DataMessage.Sender));
      MapiMessage.lpOriginator   := @MapiExpediteur;

      // paramétrage du nombre de destinataire
      MapiMessage.nRecipCount := IfThen(DataMessage.SendTo <> '', 1) + IfThen(DataMessage.SendCC <> '', 1) + IfThen(DataMessage.SendBCC <> '', 1);
      // et allocation de l mémoire nécessaire
      MapiDest                := AllocMem(SizeOf(TMapiRecipDesc) * MapiMessage.nRecipCount);
      // paramétrage des destinataire sur notre message MAPI
      MapiMessage.lpRecips    := MapiDest;

      if DataMessage.SendTo <> '' then
      begin
         MapiDest.lpszName     := PAnsiChar(AnsiString(DataMessage.SendTo));
         MapiDest.ulRecipClass := MAPI_TO;
      end;

      if DataMessage.SendCC <> '' then
      begin
         MapiDest.lpszName     := PAnsiChar(AnsiString(DataMessage.SendCC));
         MapiDest.ulRecipClass := MAPI_CC;
      end;

      if DataMessage.SendBCC <> '' then
      begin
         MapiDest.lpszName     := PAnsiChar(AnsiString(DataMessage.SendBCC));
         MapiDest.ulRecipClass := MAPI_BCC;
      end;

      // pour ajouter un accusé de lecture
      // MapiMessage.flFlags :=  MAPI_RECEIPT_REQUESTED;

      // pour ajouter des pièces jointes avec Fichier une variable de type PMapiFileDesc
      // qui est intialisée à nil au début de cette procédure
      // MapiMessage.nFileCount   := iNombreDeFichierAJoindre;
      // Fichier                  := AllocMem(SizeOf(TMapiFileDesc) * MapiMessage.nFileCount);
      // Fichier.nPosition        := 0;
      // Fichier.lpszPathName     := PAnsiChar(AnsiString(email.ListePieceJointe[iCompte].sCheminDoc));
      // MapiMessage.lpFiles      := Fichier;

      // récupération du client de messagerie
      MapiResult := MapiLogon(0, PAnsiChar(''), PAnsiChar(''), MAPI_LOGON_UI or MAPI_NEW_SESSION, 0, @MAPI_Session);

      if (MapiResult = SUCCESS_SUCCESS)then
      begin
         // nous pouvons envoyer le message
         MAPIError := MapiSendMail(0, 0, MapiMessage, MAPI_DIALOG or MAPI_LOGON_UI or MAPI_NEW_SESSION, 0);

         // liste des erreurs en MAPI : http://support.microsoft.com/kb/119647
         if MAPIError = SUCCESS_SUCCESS then
            ShowMessage('Message envoyé avec MAPI.')
         else
            raise EMAPIError.Create('Erreur MAPI avec le code ' + MapiResult.ToString + '.');
      end
      else
      begin
         raise EMAPIError.Create('Erreur MAPI avec le code ' + MapiResult.ToString + '.');
      end;
   finally
      MapiLogOff(MAPI_Session, 0, 0, 0);
      FreeMem(MapiDest);
   end;
end;

{ TMessageOutlook }

procedure TMessageOutlook.SendMessage;
const
   // constantes utilisée par outlook
   // https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
   CSTL_olMailItem        = 0;
   CSTL_olByValue         = 1;
   CSTL_olTo              = 1;
   CSTL_olCC              = 2;
   CSTL_olBCC             = 3;
   CSTL_olEditorWord      = 4;

   // constantes issues de Outlook2010.pas
   CSTL_olFormatUnspecified = $00000000;
   CSTL_olFormatPlain       = $00000001;
   CSTL_olFormatHTML        = $00000002;
   CSTL_olFormatRichText    = $00000003;

   // fonction permettant de recupérer ou de créer une instance de Outlook
   function GetOutlook(var bFind : boolean) : OLEVariant;
   begin
      bFind  := False;
      Result := Unassigned;

      try
         // récupération de la référence vers Outlook
         Result := GetActiveOleObject('Outlook.Application');
         bFind  := True;
      except
         try
            // création d'une nouvelle instance si
            // l'application n'a pas été trouvée
            Result := CreateOleObject('Outlook.Application');
            bFind  := True;
         except
            bFind := False;
         end;
      end;
   end;
var
   ovMailItem      : OLEVariant;
   ovOutlook       : OleVariant;
   bTrouve         : boolean;
   sSignature      : string;
   vDestinataire   : Variant;
begin
   inherited;

   try
      ovOutlook := GetOutlook(bTrouve);

      if not bTrouve then
      begin
         // si outlook est fermé ou ouvert en tant que administrateur
         raise EOutlookError.Create('Application Outlook non trouvée.')
      end
      else
      begin
         // création d'un email
         ovMailItem := ovOutlook.CreateItem(CSTL_olMailItem);

         // si il y a plusieurs profils dans outlook nous allons utiliser le premier
         if ovOutlook.Session.Accounts.Count > 0 then
            ovMailItem.sendUsingAccount := ovOutlook.Session.Accounts.Item(1);

         // mise en place du OlBodyFormat (olFormatRichText, olFormatHTML, olFormatPlain)
         // nous allons utiliser le texte html
         ovMailItem.BodyFormat := CSTL_olFormatHTML;

         // ajout des destinataires
         if DataMessage.SendTo <> '' then
         begin
            vDestinataire      := ovMailItem.Recipients.Add(DataMessage.SendTo);
            vDestinataire.Type := CSTL_olTo;
         end;

         if DataMessage.SendCC <> '' then
         begin
            vDestinataire      := ovMailItem.Recipients.Add(DataMessage.SendCC);
            vDestinataire.Type := CSTL_olCC;
         end;

         if DataMessage.SendBCC <> '' then
         begin
            vDestinataire      := ovMailItem.Recipients.Add(DataMessage.SendBCC);
            vDestinataire.Type := CSTL_olBCC;
         end;

         // ajouter un accusé de lecture
         // ovMailItem.ReadReceiptRequested := True;

         ovMailItem.Subject := DataMessage.MailObject;

         // il est possible d'ouvrir le message sans l'afficher
         // ceci permet par exemple de récupérer la signature
         // de l'utilisateur si elle est ajoutée automatiquement
         ovMailItem.Display(False);

         // récupération de la signature (si il y en a une)
         sSignature := ovMailItem.HTMLBody;

         // concaténation du texte du message et de la signature
         ovMailItem.HTMLBody := DataMessage.MailMessage + sSignature;

         // il est possible d'afficher l'email avant de l'envoyer avec ovMailItem.Display;
         ovMailItem.Send;
      end;
   finally
      ovMailItem := Unassigned;
   end;
end;

procedure TFormEmail.FormShow(Sender: TObject);
begin
   Memo.SetFocus;
end;

{ TMessageSender }

constructor TMessageSender.Create(aDataMessage: TDataMessage);
begin
   FDataMessage := aDataMessage;
end;

function TMessageSender.GetDataMessage: TDataMessage;
begin
   Result := FDataMessage;
end;

{ TMessageSenderSMTP }

constructor TMessageSenderSMTP.Create(aDataMessage: TDataMessage);
begin
   FDataMessage := aDataMessage;
end;

function TMessageSenderSMTP.GetDataMessage: TDataMessage;
begin
   Result := FDataMessage;
end;

function TMessageSenderSMTP.GetLogin: string;
begin
   Result := FLogin;
end;

function TMessageSenderSMTP.GetPassword: string;
begin
   Result := FPassword;
end;

function TMessageSenderSMTP.GetPort: integer;
begin
   Result := FPort;
end;

function TMessageSenderSMTP.GetServer: string;
begin
   Result := FServer;
end;

procedure TMessageSenderSMTP.SetLogin(const aLogin: string);
begin
   FLogin := aLogin;
end;

procedure TMessageSenderSMTP.SetPassword(const aPassword: string);
begin
   FPassword := aPassword;
end;

procedure TMessageSenderSMTP.SetPort(const aPort: integer);
begin
   FPort := aPort;
end;

procedure TMessageSenderSMTP.SetServer(const aServer: string);
begin
   FServer := aServer;
end;

end.
