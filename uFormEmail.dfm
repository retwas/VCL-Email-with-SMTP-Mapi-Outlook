object FormEmail: TFormEmail
  Left = 0
  Top = 0
  Caption = 'Envoyer des emails avec Delphi'
  ClientHeight = 510
  ClientWidth = 354
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object edTo: TEdit
    AlignWithMargins = True
    Left = 3
    Top = 30
    Width = 348
    Height = 21
    Align = alTop
    TabOrder = 1
    TextHint = 'A'
    ExplicitWidth = 260
  end
  object edCC: TEdit
    AlignWithMargins = True
    Left = 3
    Top = 57
    Width = 348
    Height = 21
    Align = alTop
    TabOrder = 2
    TextHint = 'CC'
    ExplicitWidth = 260
  end
  object edBCC: TEdit
    AlignWithMargins = True
    Left = 3
    Top = 84
    Width = 348
    Height = 21
    Align = alTop
    TabOrder = 3
    TextHint = 'CCI'
    ExplicitWidth = 260
  end
  object edObject: TEdit
    AlignWithMargins = True
    Left = 3
    Top = 111
    Width = 348
    Height = 21
    Align = alTop
    TabOrder = 4
    TextHint = 'Objet'
    ExplicitWidth = 260
  end
  object Memo: TMemo
    AlignWithMargins = True
    Left = 3
    Top = 192
    Width = 348
    Height = 185
    Align = alClient
    TabOrder = 7
    ExplicitWidth = 260
  end
  object RadioGroup: TRadioGroup
    Left = 0
    Top = 380
    Width = 354
    Height = 105
    Align = alBottom
    Caption = 'Type d'#39'envoi'
    ItemIndex = 0
    Items.Strings = (
      'SMTP'
      'Outlook'
      'MAPI')
    TabOrder = 8
    ExplicitWidth = 266
  end
  object btnEnvoyer: TButton
    Left = 0
    Top = 485
    Width = 354
    Height = 25
    Align = alBottom
    Caption = 'Envoyer'
    TabOrder = 9
    OnClick = btnEnvoyerClick
    ExplicitWidth = 266
  end
  object edPassword: TEdit
    AlignWithMargins = True
    Left = 3
    Top = 165
    Width = 348
    Height = 21
    Align = alTop
    PasswordChar = '*'
    TabOrder = 6
    Text = '123456'
    TextHint = 'Objet'
    ExplicitWidth = 260
  end
  object edLogin: TEdit
    AlignWithMargins = True
    Left = 3
    Top = 138
    Width = 348
    Height = 21
    Align = alTop
    TabOrder = 5
    TextHint = 'Identifiant'
    ExplicitWidth = 260
  end
  object edSender: TEdit
    AlignWithMargins = True
    Left = 3
    Top = 3
    Width = 348
    Height = 21
    Align = alTop
    TabOrder = 0
    TextHint = 'Expediteur'
    ExplicitWidth = 260
  end
end
