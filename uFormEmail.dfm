object FormEmail: TFormEmail
  Left = 0
  Top = 0
  Caption = 'Envoyer des emails avec Delphi'
  ClientHeight = 510
  ClientWidth = 266
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object edA: TEdit
    AlignWithMargins = True
    Left = 3
    Top = 30
    Width = 260
    Height = 21
    Align = alTop
    TabOrder = 0
    TextHint = 'A'
  end
  object edCC: TEdit
    AlignWithMargins = True
    Left = 3
    Top = 57
    Width = 260
    Height = 21
    Align = alTop
    TabOrder = 1
    TextHint = 'CC'
  end
  object edCCI: TEdit
    AlignWithMargins = True
    Left = 3
    Top = 84
    Width = 260
    Height = 21
    Align = alTop
    TabOrder = 2
    TextHint = 'CCI'
  end
  object edObjet: TEdit
    AlignWithMargins = True
    Left = 3
    Top = 111
    Width = 260
    Height = 21
    Align = alTop
    TabOrder = 3
    TextHint = 'Objet'
  end
  object Memo: TMemo
    AlignWithMargins = True
    Left = 3
    Top = 192
    Width = 260
    Height = 185
    Align = alClient
    TabOrder = 4
  end
  object RadioGroup: TRadioGroup
    Left = 0
    Top = 380
    Width = 266
    Height = 105
    Align = alBottom
    Caption = 'Type d'#39'envoi'
    ItemIndex = 0
    Items.Strings = (
      'SMTP'
      'Outlook'
      'MAPI')
    TabOrder = 5
  end
  object btnEnvoyer: TButton
    Left = 0
    Top = 485
    Width = 266
    Height = 25
    Align = alBottom
    Caption = 'Envoyer'
    TabOrder = 6
    OnClick = btnEnvoyerClick
  end
  object edMotDePasse: TEdit
    AlignWithMargins = True
    Left = 3
    Top = 165
    Width = 260
    Height = 21
    Align = alTop
    PasswordChar = '*'
    TabOrder = 7
    Text = 'Vpij4b24$'
    TextHint = 'Objet'
  end
  object edIdentifiant: TEdit
    AlignWithMargins = True
    Left = 3
    Top = 138
    Width = 260
    Height = 21
    Align = alTop
    TabOrder = 8
    TextHint = 'Identifiant'
  end
  object edExpediteur: TEdit
    AlignWithMargins = True
    Left = 3
    Top = 3
    Width = 260
    Height = 21
    Align = alTop
    TabOrder = 9
    TextHint = 'Expediteur'
  end
end
