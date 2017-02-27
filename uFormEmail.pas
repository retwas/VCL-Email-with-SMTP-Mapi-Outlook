unit uFormEmail;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls;

type
   TDataMessage = record
      SendTo:      string;
      SendCC:      string;
      SendBCC:     string;
      Sender:      string;
      MailObject:  string;
      MailMessage: string;
   end;

   IMessage = interface
   ['{E76ABC78-18BF-48C5-9D48-21CDFC56A4C1}']
      /// <summary>Permet d'envoyer un message en SMTP, en MAPI, ou avec Outlook</summary>
      procedure Send;
      /// <summary>Permet de renseigner un TDataMessage contenant les données de l'email à envoyer</summary>
      procedure SetDataMessage(aDataMessage: TDataMessage);
   end;

   TMessageMapi = class(TInterfacedObject, IMessage)
   strict private
      FDataMessage: TDataMessage;
   public
      procedure Send;
      procedure SetDataMessage(aDataMessage: TDataMessage);
   end;

   TMessageOutlook = class(TInterfacedObject, IMessage)
   strict private
      FDataMessage: TDataMessage;
   public
      procedure Send;
      procedure SetDataMessage(aDataMessage: TDataMessage);
   end;

   TMessageSMTP = class(TInterfacedObject, IMessage)
   strict private
      FDataMessage: TDataMessage;
      FLogin:       string;
      FPassword:    string;
      FServer:      string;
      FPort:        integer;
   public
      constructor Create(aLogin, aPassword: string);

      procedure Send;
      procedure SetDataMessage(aDataMessage: TDataMessage);
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
   Math;

type
   TMessageSMTPError    = class(Exception);
   TMessageOutlookError = class(Exception);
   TMessageMAPIError    = class(Exception);

const
   // correspondance avec le RadioGroup
   CST_SMTP    = 0;
   CST_OUTLOOK = 1;
   CST_MAPI    = 2;

{$R *.dfm}

procedure TFormEmail.btnEnvoyerClick(Sender: TObject);
var
   SendMessage: IMessage;
   DataMessage: TDataMessage;
   Cursor     : TCursor;
begin
   Cursor := Screen.Cursor;

   try
      Screen.Cursor := crHourGlass;

      try
         case RadioGroup.ItemIndex of
            CST_SMTP    : SendMessage := TMessageSMTP.Create(edLogin.Text, edPassword.Text);
            CST_OUTLOOK : SendMessage := TMessageOutlook.Create;
            else
               SendMessage := TMessageMapi.Create;
         end;

         DataMessage.SendTo      := edTo.Text;
         DataMessage.SendCC      := edCC.Text;
         DataMessage.SendBCC     := edBCC.Text;
         DataMessage.Sender      := edSender.Text;
         DataMessage.MailObject  := edObject.Text;
         DataMessage.MailMessage := Memo.Text;

         SendMessage.SetDataMessage(DataMessage);
         SendMessage.Send;
      except
         on E : Exception do
            MessageDlg('Erreur lors de l''envoi de l''email.' + #10#13 + e.Message, mtError, [mbOK], 0);
      end;
   finally
      Screen.Cursor := Cursor;
   end;
end;

{ TMessageSMTP }

constructor TMessageSMTP.Create(aLogin, aPassword : string);
begin
   inherited Create;

   FLogin    := aLogin;
   FPassword := aPassword;
   FServer   := 'smtp.office365.com';
   FPort     := 587;
end;

procedure TMessageSMTP.Send;
var
   IdSMTP    : TIdSMTP;
   IdSSLIO   : TIdSSLIOHandlerSocketOpenSSL;
   IdMessage : TIdMessage;
begin
   // création des composants Indy
   IdSMTP    := TIdSMTP.Create(nil);
   IdMessage := TIdMessage.Create(nil);
   IdSSLIO   := TIdSSLIOHandlerSocketOpenSSL.Create(nil);

   try
      // adresse et port du serveur SMTP
      IdSMTP.AuthType := satDefault;
      IdSMTP.Host     := FServer;
      IdSMTP.Port     := FPort;

      // utilisation du mode sécurisé (ou pas si l'envoi en anonyme est autorisé)
      if (FLogin <> '') and (FPassword <> '') then
      begin
         IdSSLIO.SSLOptions.Method := sslvTLSv1;
         IdSMTP.Username  := FLogin;
         IdSMTP.Password  := FPassword;
         IdSMTP.IOHandler := IdSSLIO;
         IdSMTP.UseTLS    := utUseExplicitTLS;
      end;

      // connexion au serveur
      try
         IdSMTP.Connect;
      except
         on e : Exception do
            raise TMessageSMTPError.Create('Erreur de connexion au serveur SMTP' + #10#13 + e.Message);
      end;

      IdMessage.Clear;

      // paramétrage de l'expéditeur
      IdMessage.From.Text           := FDataMessage.Sender;
      IdMessage.ReplyTo.Add.Address := FDataMessage.Sender;

      // ajout des destinataires
      if FDataMessage.SendTo <> '' then
         IdMessage.Recipients.Add.Address := FDataMessage.SendTo;

      if FDataMessage.SendCC <> '' then
         IdMessage.CCList.Add.Address     := FDataMessage.SendCC;

      if FDataMessage.SendBCC <> '' then
         IdMessage.BccList.Add.Address    := FDataMessage.SendBCC;

      // objet du message
      IdMessage.Subject := FDataMessage.MailObject;

      // paramétrage de la date et de la priorité
      IdMessage.Date     := Now;
      IdMessage.Priority := mpNormal;

      // il est possible de paramétrer un l'accusé de lecture
      // idMessage.ReceiptRecipient.Address := adresse de l'expediteur;

      // corps du message
      IdMessage.Body.Text := FDataMessage.MailMessage;

      try
         IdSMTP.Send(IdMessage);
      except
         on e : Exception do
            raise TMessageSMTPError.Create('Erreur lors de l''envoi de l''email.' + #10#13 + e.Message);
      end;
   finally
      IdSMTP.Disconnect;

      FreeAndNil(IdSSLIO);
      FreeAndNil(IdMessage);
      FreeAndNil(IdSMTP);
   end;
end;

procedure TMessageSMTP.SetDataMessage(aDataMessage: TDataMessage);
begin
   FDataMessage := aDataMessage;
end;

{ TMessageMapi }

procedure TMessageMapi.Send;
var
   MapiMessage    : TMapiMessage;
   MapiDest       : PMapiRecipDesc;
   MapiExpediteur : TMapiRecipDesc;
   MAPI_Session   : Cardinal;
   MapiResult     : Cardinal;
   MAPIError      : DWord;
begin
   MapiDest := nil;

   try
      // FillChar permet de remplir une suite d'octets avec une valeur, ici 0
      FillChar(MapiMessage, Sizeof(TMapiMessage), 0);
      MapiMessage.lpszSubject  := PAnsiChar(AnsiString(FDataMessage.MailObject));
      MapiMessage.lpszNoteText := PAnsiChar(AnsiString(FDataMessage.MailMessage));

      // même traitement pour paramétrer l'expéditeur
      FillChar(MapiExpediteur, Sizeof(TMapiRecipDesc), 0);
      MapiExpediteur.lpszName    := PAnsiChar(AnsiString(FDataMessage.Sender));
      MapiExpediteur.lpszAddress := PAnsiChar(AnsiString(FDataMessage.Sender));
      MapiMessage.lpOriginator   := @MapiExpediteur;

      // paramétrage du nombre de destinataire
      MapiMessage.nRecipCount := IfThen(FDataMessage.SendTo <> '', 1) + IfThen(FDataMessage.SendCC <> '', 1) + IfThen(FDataMessage.SendBCC <> '', 1);
      // et allocation de l mémoire nécessaire
      MapiDest                := AllocMem(SizeOf(TMapiRecipDesc) * MapiMessage.nRecipCount);
      // paramétrage des destinataire sur notre message MAPI
      MapiMessage.lpRecips    := MapiDest;

      if FDataMessage.SendTo <> '' then
      begin
         MapiDest.lpszName     := PAnsiChar(AnsiString(FDataMessage.SendTo));
         MapiDest.ulRecipClass := MAPI_TO;
      end;

      if FDataMessage.SendCC <> '' then
      begin
         MapiDest.lpszName     := PAnsiChar(AnsiString(FDataMessage.SendCC));
         MapiDest.ulRecipClass := MAPI_CC;
      end;

      if FDataMessage.SendBCC <> '' then
      begin
         MapiDest.lpszName     := PAnsiChar(AnsiString(FDataMessage.SendBCC));
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
            raise TMessageMAPIError.Create('Erreur MAPI avec le code ' + MapiResult.ToString + '.');
      end
      else
      begin
         raise TMessageMAPIError.Create('Erreur MAPI avec le code ' + MapiResult.ToString + '.');
      end;
   finally
      MapiLogOff(MAPI_Session, 0, 0, 0);
      FreeMem(MapiDest);
   end;
end;

procedure TMessageMapi.SetDataMessage(aDataMessage: TDataMessage);
begin
   FDataMessage := aDataMessage;
end;

{ TMessageOutlook }

procedure TMessageOutlook.Send;
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
   try
      ovOutlook := GetOutlook(bTrouve);

      if not bTrouve then
      begin
         // si outlook est fermé ou ouvert en tant que administrateur
         raise TMessageOutlookError.Create('Application Outlook non trouvée.')
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
         if FDataMessage.SendTo <> '' then
         begin
            vDestinataire      := ovMailItem.Recipients.Add(FDataMessage.SendTo);
            vDestinataire.Type := CSTL_olTo;
         end;

         if FDataMessage.SendCC <> '' then
         begin
            vDestinataire      := ovMailItem.Recipients.Add(FDataMessage.SendCC);
            vDestinataire.Type := CSTL_olCC;
         end;

         if FDataMessage.SendBCC <> '' then
         begin
            vDestinataire      := ovMailItem.Recipients.Add(FDataMessage.SendBCC);
            vDestinataire.Type := CSTL_olBCC;
         end;

         // ajouter un accusé de lecture
         // ovMailItem.ReadReceiptRequested := True;

         ovMailItem.Subject := FDataMessage.MailObject;

         // il est possible d'ouvrir le message sans l'afficher
         // ceci permet par exemple de récupérer la signature
         // de l'utilisateur si elle est ajoutée automatiquement
         ovMailItem.Display(False);

         // récupération de la signature (si il y en a une)
         sSignature := ovMailItem.HTMLBody;

         // concaténation du texte du message et de la signature
         ovMailItem.HTMLBody := FDataMessage.MailMessage + sSignature;

         // il est possible d'afficher l'email avant de l'envoyer avec ovMailItem.Display;
         ovMailItem.Send;
      end;
   finally
      ovMailItem := Unassigned;
   end;
end;

procedure TMessageOutlook.SetDataMessage(aDataMessage: TDataMessage);
begin
   FDataMessage := aDataMessage;
end;

end.
