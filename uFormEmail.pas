unit uFormEmail;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls;

type
   TSendType = (stSMTP, stMAPI, stOutlook);

   TMessage = class
   strict private
      FsA           : string;
      FsCC          : string;
      FsCCI         : string;
      FsObjet       : string;
      FsTexte       : string;
      FstMode       : TSendType;
      FsIdentifiant : string;
      FsMotDePasse  : string;
      FsExpediteur  : string;

      procedure EnvoyerAvecSMTP;
      procedure EnvoyerAvecOutlook;
      procedure EnvoyerAvecMAPI;
   public
      procedure Envoyer;

      property sA          : string    read FsA             write FsA;
      property sCC         : string    read FsCC            write FsCC;
      property sCCI        : string    read FsCCI           write FsCCI;
      property sObjet      : string    read FsObjet         write FsObjet;
      property sTexte      : string    read FsTexte         write FsTexte;
      property stMode      : TSendType read FstMode         write FstMode;
      property sIdentifiant: string    read FsIdentifiant   write FsIdentifiant;
      property sMotDePasse : string    read FsMotDePasse    write FsMotDePasse;
      property sExpediteur : string    read FsExpediteur    write FsExpediteur;
   end;


  TFormEmail = class(TForm)
    edA: TEdit;
    edCC: TEdit;
    edCCI: TEdit;
    edObjet: TEdit;
    Memo: TMemo;
    RadioGroup: TRadioGroup;
    btnEnvoyer: TButton;
    edMotDePasse: TEdit;
    edIdentifiant: TEdit;
    edExpediteur: TEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure btnEnvoyerClick(Sender: TObject);
  private
    FMessage : TMessage;
  public
    { Déclarations publiques }
  end;

var
  FormEmail: TFormEmail;

implementation

uses
   IdSMTP, IdExplicitTLSClientServerBase, IdMessage, IdSSLOpenSSL,
   System.Win.ComObj,
   Winapi.Mapi,
   Math;

const
   CST_SMTP    = 0;
   CST_OUTLOOK = 1;
   CST_MAPI    = 2;

{$R *.dfm}

procedure TFormEmail.btnEnvoyerClick(Sender: TObject);
begin
   FMessage.sA           := edA.Text;
   FMessage.sCC          := edCC.Text;
   FMessage.sCCI         := edCCI.Text;
   FMessage.sObjet       := edObjet.Text;
   FMessage.sTexte       := Memo.Text;
   FMessage.sIdentifiant := edIdentifiant.Text;
   FMessage.sMotDePasse  := edMotDePasse.Text;
   FMessage.sExpediteur  := edExpediteur.Text;

   if RadioGroup.ItemIndex = CST_SMTP then
      FMessage.stMode := stSMTP
   else if RadioGroup.ItemIndex = CST_OUTLOOK then
      FMessage.stMode := stOutlook
   else
      FMessage.stMode := stMAPI;

   FMessage.Envoyer;
end;

procedure TFormEmail.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
   if CanClose then
      FreeAndNil(FMessage);
end;

procedure TFormEmail.FormCreate(Sender: TObject);
begin
   FMessage             := TMessage.Create;
   RadioGroup.ItemIndex := CST_SMTP;
end;

{ TMessage }

procedure TMessage.Envoyer;
begin
   if FstMode = stSMTP then
      EnvoyerAvecSMTP
   else
      if FstMode = stOutlook then
         EnvoyerAvecOutlook
      else
         EnvoyerAvecMAPI;
end;

procedure TMessage.EnvoyerAvecMAPI;
var
   MapiMessage    : TMapiMessage;
   MapiDest       : PMapiRecipDesc;
   MapiExpediteur : TMapiRecipDesc;
   MAPI_Session   : Cardinal;
   MapiResult     : Cardinal;
   MAPIError      : DWord;
begin
   MapiDest := nil;
   // Fichier := nil

   try
      // FillChar permet de remplir une suite d'octets avec une valeur, ici 0
      FillChar(MapiMessage, Sizeof(TMapiMessage), 0);
      MapiMessage.lpszSubject  := PAnsiChar(AnsiString(FsObjet));
      MapiMessage.lpszNoteText := PAnsiChar(AnsiString(FsTexte));

      // même traitement pour paramétrer l'expéditeur
      FillChar(MapiExpediteur, Sizeof(TMapiRecipDesc), 0);
      MapiExpediteur.lpszName    := PAnsiChar(AnsiString(FsExpediteur));
      MapiExpediteur.lpszAddress := PAnsiChar(AnsiString(FsExpediteur));
      MapiMessage.lpOriginator   := @MapiExpediteur;

      // paramétrage du nombre de destinataire
      MapiMessage.nRecipCount := IfThen(FsA <> '', 1) + IfThen(FsCC <> '', 1) + IfThen(FsCCI <> '', 1);
      // et allocation de mémoire nécessaire
      MapiDest                := AllocMem(SizeOf(TMapiRecipDesc) * MapiMessage.nRecipCount);
      // paramétrage des destinataire sur notre message MAPI
      MapiMessage.lpRecips    := MapiDest;

      if FsA <> '' then
      begin
         MapiDest.lpszName     := PAnsiChar(AnsiString(FsA));
         MapiDest.ulRecipClass := MAPI_TO;
      end;

      if FsCC <> '' then
      begin
         MapiDest.lpszName     := PAnsiChar(AnsiString(FsCC));
         MapiDest.ulRecipClass := MAPI_CC;
      end;

      if FsCCI <> '' then
      begin
         MapiDest.lpszName     := PAnsiChar(AnsiString(FsCCI));
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
            ShowMessage('Message envoyé avec MAPI')
         else
            ShowMessage('Erreur MAPI avec le code ' + MapiResult.ToString);
      end
      else
      begin
         ShowMessage('Erreur MAPI avec le code ' + MapiResult.ToString);
      end;
   finally
      MapiLogOff(MAPI_Session, 0, 0, 0);
      FreeMem(MapiDest);
   end;
end;

procedure TMessage.EnvoyerAvecOutlook;
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
   function GetOutlookApp(var bFind : boolean) : OLEVariant;
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
      ovOutlook := GetOutlookApp(bTrouve);

      if not bTrouve then
         Showmessage('Application Outlook non trouvée.')
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
         if FsA <> '' then
         begin
            vDestinataire      := ovMailItem.Recipients.Add(FsA);
            vDestinataire.Type := CSTL_olTo;
         end;

         if FsCC <> '' then
         begin
            vDestinataire      := ovMailItem.Recipients.Add(FsCC);
            vDestinataire.Type := CSTL_olCC;
         end;

         if FsCCI <> '' then
         begin
            vDestinataire      := ovMailItem.Recipients.Add(FsCCI);
            vDestinataire.Type := CSTL_olBCC;
         end;

         // ajouter un accusé de lecture
         // ovMailItem.ReadReceiptRequested := True;

         ovMailItem.Subject := FsObjet;

         // il est possible d'ouvrir le message sans l'afficher
         // ceci permet par exemple de récupérer la signature
         // de l'utilisateur si elle est ajoutée automatiquement
         ovMailItem.Display(False);

         // récupération de la signature (si il y en a une)
         sSignature := ovMailItem.HTMLBody;

         // concaténation du texte du message et de la signature
         ovMailItem.HTMLBody := FsTexte + sSignature;

         // il est possible d'afficher l'email avant de l'envoyer avec ovMailItem.Display;
         ovMailItem.Send;
      end;
   finally
      ovMailItem := Unassigned;
   end;
end;

procedure TMessage.EnvoyerAvecSMTP;
var
   IdSMTP    : TIdSMTP;
   IdSSLIO   : TIdSSLIOHandlerSocketOpenSSL;
   IdMessage : TIdMessage;
begin
   // création des composants Indy
   idSMTP    := TIdSMTP.Create(nil);
   idMessage := TIdMessage.Create(nil);
   idSSLIO   := TIdSSLIOHandlerSocketOpenSSL.Create(nil);
   try
      // adresse et port du serveur SMTP
      idSMTP.AuthType := satDefault;
      idSMTP.Host     := 'smtp.office365.com';
      idSMTP.Port     := 587;

      // utilisation du mode sécurisé
      if (FsIdentifiant <> '') and (FsMotDePasse <> '') then
      begin
         idSSLIO.SSLOptions.Method := sslvTLSv1;
         idSMTP.Username  := FsIdentifiant;
         idSMTP.Password  := FsMotDePasse;
         IdSMTP.IOHandler := idSSLIO;
         idSMTP.UseTLS    := utUseExplicitTLS;
      end;

      // connexion au serveur
      try
         idSMTP.Connect;
      except
         on e : Exception do
         begin
            raise Exception.Create('SocketError : ' + e.Message);
         end;
      end;

      idMessage.Clear;

      // paramétrage de l'expéditeur
      idMessage.From.Text           := FsExpediteur;
      idMessage.ReplyTo.Add.Address := FsExpediteur;

      // ajout des destinataires
      if FsA <> '' then
         idMessage.Recipients.Add.Address := FsA;

      if FsCC <> '' then
         idMessage.CCList.Add.Address     := FsCC;

      if FsCCI <> '' then
         idMessage.BccList.Add.Address    := FsCCI;

      // objet du message
      IdMessage.Subject := FsObjet;

      // paramétrage de la date et de la priorité
      idMessage.Date     := Now;
      IdMessage.Priority := mpNormal;

      // il est possible de paramétrer un l'accusé de lecture
      // idMessage.ReceiptRecipient.Address := adresse de l'expediteur;

      // corps du message
      idMessage.Body.Text := FsTexte;

      try
         IdSMTP.Send(IdMessage);
      except
         on e : Exception do
            raise Exception.Create('Erreur SMTP : ' + e.Message);
      end;
   finally
      IdSMTP.Disconnect;

      FreeAndNil(idSSLIO);
      FreeAndNil(IdMessage);
      FreeAndNil(IdSMTP);
   end;
end;

end.
