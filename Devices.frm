VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A4749554-0441-4E64-8A03-3323601631C7}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmDevices 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Módulos/Devices  não Cadastrados"
   ClientHeight    =   5940
   ClientLeft      =   2265
   ClientTop       =   2565
   ClientWidth     =   12210
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Devices.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5940
   ScaleWidth      =   12210
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtLocal 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6240
      Width           =   4575
   End
   Begin TrueOleDBGrid80.TDBGrid tdbg1 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   8916
      _LayoutType     =   4
      _RowHeight      =   18
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   " Entity"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   " Descrição"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   " Tipo"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "sinal"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "ruído"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Comm"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   " Data"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   16
      Columns(7)._MaxComboItems=   5
      Columns(7).ValueItems(0)._DefaultItem=   0
      Columns(7).ValueItems(0).Value=   "0"
      Columns(7).ValueItems(0).Value.vt=   8
      Columns(7).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(7).ValueItems(0).DisplayValue(0)=   "bHQAADYJAABCTTYJAAAAAAAANgAAACgAAAAwAAAAEAAAAAEAGAAAAAAAAAkAAAAAAAAAAAAAAAAA"
      Columns(7).ValueItems(0).DisplayValue(1)=   "AAAAAAAAAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(2)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(3)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(4)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(5)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(6)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(7)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(8)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(9)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(10)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(11)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(12)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(13)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(14)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(15)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(16)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(17)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(18)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(19)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(20)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(21)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(22)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(23)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(24)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(25)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(26)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(27)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(28)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(29)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(30)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(31)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(32)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(33)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(34)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(35)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(36)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(37)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(38)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(39)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(40)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(0).DisplayValue(41)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8="
      Columns(7).ValueItems(0).DisplayValue.vt=   9
      Columns(7).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(7).ValueItems(1)._DefaultItem=   0
      Columns(7).ValueItems(1).Value=   "1"
      Columns(7).ValueItems(1).Value.vt=   8
      Columns(7).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(7).ValueItems(1).DisplayValue(0)=   "bHQAADYJAABCTTYJAAAAAAAANgAAACgAAAAwAAAAEAAAAAEAGAAAAAAAAAkAAAAAAAAAAAAAAAAA"
      Columns(7).ValueItems(1).DisplayValue(1)=   "AAAAAAAAAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(1).DisplayValue(2)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(1).DisplayValue(3)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(1).DisplayValue(4)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A//8A"
      Columns(7).ValueItems(1).DisplayValue(5)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(1).DisplayValue(6)=   "//8A//8A//8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(1).DisplayValue(7)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(1).DisplayValue(8)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(1).DisplayValue(9)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(1).DisplayValue(10)=   "AP8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(1).DisplayValue(11)=   "//8A//8A//8A//8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(1).DisplayValue(12)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(1).DisplayValue(13)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(1).DisplayValue(14)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(1).DisplayValue(15)=   "AP8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(1).DisplayValue(16)=   "//8A//8A//8A//8A//8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(1).DisplayValue(17)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(1).DisplayValue(18)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(1).DisplayValue(19)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(1).DisplayValue(20)=   "AP8AAP8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(1).DisplayValue(21)=   "//8A//8A//8A//8A//8A//8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(1).DisplayValue(22)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(1).DisplayValue(23)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8AAP8AAP8AAP8A"
      Columns(7).ValueItems(1).DisplayValue(24)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(1).DisplayValue(25)=   "AP8AAP8AAP8AAP8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(1).DisplayValue(26)=   "//8A//8A//8A//8A//8A//8A//8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(1).DisplayValue(27)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(1).DisplayValue(28)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8AAP8AAP8A"
      Columns(7).ValueItems(1).DisplayValue(29)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(1).DisplayValue(30)=   "AP8AAP8AAP8AAP8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(1).DisplayValue(31)=   "//8A//8A//8A//8A//8A//8A//8A//8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(1).DisplayValue(32)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A//8A//8A//8A//8A"
      Columns(7).ValueItems(1).DisplayValue(33)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8AAP8A"
      Columns(7).ValueItems(1).DisplayValue(34)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(1).DisplayValue(35)=   "AP8AAP8AAP8AAP8AAP8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(1).DisplayValue(36)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(1).DisplayValue(37)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A//8A//8A//8A//8A"
      Columns(7).ValueItems(1).DisplayValue(38)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(1).DisplayValue(39)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(7).ValueItems(1).DisplayValue(40)=   "AP8AAP8AAP8AAP8AAP8AAP8AAP8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(1).DisplayValue(41)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8="
      Columns(7).ValueItems(1).DisplayValue.vt=   9
      Columns(7).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(7).ValueItems(2)._DefaultItem=   0
      Columns(7).ValueItems(2).Value=   "2"
      Columns(7).ValueItems(2).Value.vt=   8
      Columns(7).ValueItems(2).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(7).ValueItems(2).DisplayValue(0)=   "bHQAADYJAABCTTYJAAAAAAAANgAAACgAAAAwAAAAEAAAAAEAGAAAAAAAAAkAAAAAAAAAAAAAAAAA"
      Columns(7).ValueItems(2).DisplayValue(1)=   "AAAAAAAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(2)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(3)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(4)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(5)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(6)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(7)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(8)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(9)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(10)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(11)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(12)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(13)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(14)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(15)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(16)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(17)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(18)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(19)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(20)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(21)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(22)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(23)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(24)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(25)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(26)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(27)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(28)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(29)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(30)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(31)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(32)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(33)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(34)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(35)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(36)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(37)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(38)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(39)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(40)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(2).DisplayValue(41)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8="
      Columns(7).ValueItems(2).DisplayValue.vt=   9
      Columns(7).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
      Columns(7).ValueItems(3)._DefaultItem=   0
      Columns(7).ValueItems(3).Value=   "3"
      Columns(7).ValueItems(3).Value.vt=   8
      Columns(7).ValueItems(3).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(7).ValueItems(3).DisplayValue(0)=   "bHQAADYJAABCTTYJAAAAAAAANgAAACgAAAAwAAAAEAAAAAEAGAAAAAAAAAkAAAAAAAAAAAAAAAAA"
      Columns(7).ValueItems(3).DisplayValue(1)=   "AAAAAACE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A"
      Columns(7).ValueItems(3).DisplayValue(2)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(3)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(3).DisplayValue(4)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(5)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(6)=   "//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(3).DisplayValue(7)=   "5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(8)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(3).DisplayValue(9)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(10)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(11)=   "//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(3).DisplayValue(12)=   "5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(13)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(3).DisplayValue(14)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(15)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(16)=   "//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(3).DisplayValue(17)=   "5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(18)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(3).DisplayValue(19)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(20)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(21)=   "//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(3).DisplayValue(22)=   "5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(23)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE"
      Columns(7).ValueItems(3).DisplayValue(24)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(25)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(26)=   "//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(3).DisplayValue(27)=   "5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(28)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE"
      Columns(7).ValueItems(3).DisplayValue(29)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(30)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(31)=   "//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(3).DisplayValue(32)=   "5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(33)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE"
      Columns(7).ValueItems(3).DisplayValue(34)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(35)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(36)=   "//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(3).DisplayValue(37)=   "5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(38)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E"
      Columns(7).ValueItems(3).DisplayValue(39)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(40)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(3).DisplayValue(41)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8="
      Columns(7).ValueItems(3).DisplayValue.vt=   9
      Columns(7).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
      Columns(7).ValueItems(4)._DefaultItem=   0
      Columns(7).ValueItems(4).Value=   "4"
      Columns(7).ValueItems(4).Value.vt=   8
      Columns(7).ValueItems(4).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(7).ValueItems(4).DisplayValue(0)=   "bHQAADYJAABCTTYJAAAAAAAANgAAACgAAAAwAAAAEAAAAAEAGAAAAAAAAAkAAAAAAAAAAAAAAAAA"
      Columns(7).ValueItems(4).DisplayValue(1)=   "AAAAAACE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(4).DisplayValue(2)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(4).DisplayValue(3)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(4).DisplayValue(4)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A"
      Columns(7).ValueItems(4).DisplayValue(5)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(4).DisplayValue(6)=   "//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(4).DisplayValue(7)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(4).DisplayValue(8)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(4).DisplayValue(9)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A"
      Columns(7).ValueItems(4).DisplayValue(10)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(4).DisplayValue(11)=   "//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(4).DisplayValue(12)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(4).DisplayValue(13)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(4).DisplayValue(14)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(7).ValueItems(4).DisplayValue(15)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(4).DisplayValue(16)=   "//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(4).DisplayValue(17)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(4).DisplayValue(18)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(4).DisplayValue(19)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(7).ValueItems(4).DisplayValue(20)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(4).DisplayValue(21)=   "//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(4).DisplayValue(22)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(4).DisplayValue(23)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE"
      Columns(7).ValueItems(4).DisplayValue(24)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(4).DisplayValue(25)=   "5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(4).DisplayValue(26)=   "//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(4).DisplayValue(27)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(4).DisplayValue(28)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE"
      Columns(7).ValueItems(4).DisplayValue(29)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(4).DisplayValue(30)=   "5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(4).DisplayValue(31)=   "//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(4).DisplayValue(32)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(4).DisplayValue(33)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE"
      Columns(7).ValueItems(4).DisplayValue(34)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(4).DisplayValue(35)=   "5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(4).DisplayValue(36)=   "//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(4).DisplayValue(37)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(4).DisplayValue(38)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E"
      Columns(7).ValueItems(4).DisplayValue(39)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(4).DisplayValue(40)=   "5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(4).DisplayValue(41)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8="
      Columns(7).ValueItems(4).DisplayValue.vt=   9
      Columns(7).ValueItems(4)._PropDict=   "_DefaultItem,517,2"
      Columns(7).ValueItems(5)._DefaultItem=   0
      Columns(7).ValueItems(5).Value=   "5"
      Columns(7).ValueItems(5).Value.vt=   8
      Columns(7).ValueItems(5).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(7).ValueItems(5).DisplayValue(0)=   "bHQAADYJAABCTTYJAAAAAAAANgAAACgAAAAwAAAAEAAAAAEAGAAAAAAAAAkAAAAAAAAAAAAAAAAA"
      Columns(7).ValueItems(5).DisplayValue(1)=   "AAAAAACE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(2)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A"
      Columns(7).ValueItems(5).DisplayValue(3)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(4)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(5)=   "5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(5).DisplayValue(6)=   "//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(7)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A"
      Columns(7).ValueItems(5).DisplayValue(8)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(9)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(10)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(5).DisplayValue(11)=   "//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(12)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A"
      Columns(7).ValueItems(5).DisplayValue(13)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(14)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(15)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(5).DisplayValue(16)=   "//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(17)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A"
      Columns(7).ValueItems(5).DisplayValue(18)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(19)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(20)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(5).DisplayValue(21)=   "//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(22)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(7).ValueItems(5).DisplayValue(23)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(24)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(25)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(5).DisplayValue(26)=   "//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(27)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(7).ValueItems(5).DisplayValue(28)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(29)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(30)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(5).DisplayValue(31)=   "//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(32)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(33)=   "5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE"
      Columns(7).ValueItems(5).DisplayValue(34)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(35)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(5).DisplayValue(36)=   "//8A//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(37)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(38)=   "5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//+E"
      Columns(7).ValueItems(5).DisplayValue(39)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(5).DisplayValue(40)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(5).DisplayValue(41)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8="
      Columns(7).ValueItems(5).DisplayValue.vt=   9
      Columns(7).ValueItems(5)._PropDict=   "_DefaultItem,517,2"
      Columns(7).ValueItems(6)._DefaultItem=   0
      Columns(7).ValueItems(6).Value=   "6"
      Columns(7).ValueItems(6).Value.vt=   8
      Columns(7).ValueItems(6).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(7).ValueItems(6).DisplayValue(0)=   "bHQAADYJAABCTTYJAAAAAAAANgAAACgAAAAwAAAAEAAAAAEAGAAAAAAAAAkAAAAAAAAAAAAAAAAA"
      Columns(7).ValueItems(6).DisplayValue(1)=   "AAAAAACE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(2)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(3)=   "5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(4)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(5)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A"
      Columns(7).ValueItems(6).DisplayValue(6)=   "//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(7)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(8)=   "5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(9)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(10)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A"
      Columns(7).ValueItems(6).DisplayValue(11)=   "//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(12)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(13)=   "5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(14)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(15)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A"
      Columns(7).ValueItems(6).DisplayValue(16)=   "//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(17)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(18)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(19)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(20)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A"
      Columns(7).ValueItems(6).DisplayValue(21)=   "//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(22)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(23)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(24)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(25)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A"
      Columns(7).ValueItems(6).DisplayValue(26)=   "//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(27)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(28)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//+E5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(29)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(30)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(31)=   "5wAA//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(32)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(33)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//+E5wCE"
      Columns(7).ValueItems(6).DisplayValue(34)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(35)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(36)=   "5wCE5wAA//8A//8A//8A//8A//8A//8A//+E5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(37)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(38)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//+E"
      Columns(7).ValueItems(6).DisplayValue(39)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(40)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(6).DisplayValue(41)=   "5wCE5wAA//8A//8A//8A//8A//8A//8A//8A//8="
      Columns(7).ValueItems(6).DisplayValue.vt=   9
      Columns(7).ValueItems(6)._PropDict=   "_DefaultItem,517,2"
      Columns(7).ValueItems(7)._DefaultItem=   0
      Columns(7).ValueItems(7).Value=   "7"
      Columns(7).ValueItems(7).Value.vt=   8
      Columns(7).ValueItems(7).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(7).ValueItems(7).DisplayValue(0)=   "bHQAADYJAABCTTYJAAAAAAAANgAAACgAAAAwAAAAEAAAAAEAGAAAAAAAAAkAAAAAAAAAAAAAAAAA"
      Columns(7).ValueItems(7).DisplayValue(1)=   "AAAAAACE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(2)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(3)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(4)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(5)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(6)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(7)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(8)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(9)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(10)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(11)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(12)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(13)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(14)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(15)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(16)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(17)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(18)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(19)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(20)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(21)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(22)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(23)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(24)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(25)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(26)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(27)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(28)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(29)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(30)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(31)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(32)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(33)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(34)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(35)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(36)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(37)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(38)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(39)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(40)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(7).ValueItems(7).DisplayValue(41)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wA="
      Columns(7).ValueItems(7).DisplayValue.vt=   9
      Columns(7).ValueItems(7)._PropDict=   "_DefaultItem,517,2"
      Columns(7).ValueItems(8)._DefaultItem=   0
      Columns(7).ValueItems(8).Value=   "8"
      Columns(7).ValueItems(8).Value.vt=   8
      Columns(7).ValueItems(8).DisplayValue=   "_"
      Columns(7).ValueItems(8).DisplayValue.vt=   8
      Columns(7).ValueItems(8)._PropDict=   "_DefaultItem,517,2"
      Columns(7).ValueItems.Count=   9
      Columns(7).Caption=   " Hora"
      Columns(7).DataField=   ""
      Columns(7).NumberFormat=   "Short Time"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   979
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).AllowColSelect=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1191"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1085"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=256"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=5080"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=4974"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=256"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=3598"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=3493"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=256"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=2090"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1984"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=65792"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=2011"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1905"
      Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=65792"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=1561"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1455"
      Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=65792"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(38)=   "Column(6).Width=3016"
      Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=2910"
      Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=65793"
      Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(44)=   "Column(7).Width=2037"
      Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=1931"
      Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=65793"
      Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=11.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=11.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      RowDividerStyle =   0
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   13160660
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=0,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=31"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34,.bgcolor=&H80000005&"
      _StyleDefs(16)  =   ":id=8,.fgcolor=&H8000000D&"
      _StyleDefs(17)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(18)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(19)  =   "RecordSelectorStyle:id=79,.parent=2,.namedParent=81"
      _StyleDefs(20)  =   "FilterBarStyle:id=82,.parent=1,.namedParent=84"
      _StyleDefs(21)  =   "Splits(0).Style:id=37,.parent=1"
      _StyleDefs(22)  =   "Splits(0).CaptionStyle:id=46,.parent=4"
      _StyleDefs(23)  =   "Splits(0).HeadingStyle:id=38,.parent=2"
      _StyleDefs(24)  =   "Splits(0).FooterStyle:id=39,.parent=3"
      _StyleDefs(25)  =   "Splits(0).InactiveStyle:id=40,.parent=5"
      _StyleDefs(26)  =   "Splits(0).SelectedStyle:id=42,.parent=6"
      _StyleDefs(27)  =   "Splits(0).EditorStyle:id=41,.parent=7"
      _StyleDefs(28)  =   "Splits(0).HighlightRowStyle:id=43,.parent=8"
      _StyleDefs(29)  =   "Splits(0).EvenRowStyle:id=44,.parent=9"
      _StyleDefs(30)  =   "Splits(0).OddRowStyle:id=45,.parent=10"
      _StyleDefs(31)  =   "Splits(0).RecordSelectorStyle:id=80,.parent=79"
      _StyleDefs(32)  =   "Splits(0).FilterBarStyle:id=83,.parent=82"
      _StyleDefs(33)  =   "Splits(0).Columns(0).Style:id=24,.parent=37,.alignment=0,.locked=0"
      _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=38,.alignment=0"
      _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=39,.alignment=3"
      _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=41"
      _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=54,.parent=37,.alignment=0,.locked=0"
      _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=38,.alignment=0"
      _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=39,.alignment=3"
      _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=41"
      _StyleDefs(41)  =   "Splits(0).Columns(2).Style:id=58,.parent=37,.alignment=0,.locked=0"
      _StyleDefs(42)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=38,.alignment=0"
      _StyleDefs(43)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=39,.alignment=3"
      _StyleDefs(44)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=41"
      _StyleDefs(45)  =   "Splits(0).Columns(3).Style:id=74,.parent=37"
      _StyleDefs(46)  =   "Splits(0).Columns(3).HeadingStyle:id=71,.parent=38"
      _StyleDefs(47)  =   "Splits(0).Columns(3).FooterStyle:id=72,.parent=39"
      _StyleDefs(48)  =   "Splits(0).Columns(3).EditorStyle:id=73,.parent=41"
      _StyleDefs(49)  =   "Splits(0).Columns(4).Style:id=62,.parent=37"
      _StyleDefs(50)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=38"
      _StyleDefs(51)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=39"
      _StyleDefs(52)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=41"
      _StyleDefs(53)  =   "Splits(0).Columns(5).Style:id=78,.parent=37"
      _StyleDefs(54)  =   "Splits(0).Columns(5).HeadingStyle:id=75,.parent=38"
      _StyleDefs(55)  =   "Splits(0).Columns(5).FooterStyle:id=76,.parent=39"
      _StyleDefs(56)  =   "Splits(0).Columns(5).EditorStyle:id=77,.parent=41"
      _StyleDefs(57)  =   "Splits(0).Columns(6).Style:id=70,.parent=37,.alignment=2,.locked=0"
      _StyleDefs(58)  =   "Splits(0).Columns(6).HeadingStyle:id=67,.parent=38,.alignment=0"
      _StyleDefs(59)  =   "Splits(0).Columns(6).FooterStyle:id=68,.parent=39,.alignment=0"
      _StyleDefs(60)  =   "Splits(0).Columns(6).EditorStyle:id=69,.parent=41"
      _StyleDefs(61)  =   "Splits(0).Columns(7).Style:id=66,.parent=37,.alignment=2,.locked=0"
      _StyleDefs(62)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=38,.alignment=0"
      _StyleDefs(63)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=39,.alignment=0"
      _StyleDefs(64)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=41"
      _StyleDefs(65)  =   "Named:id=29:Normal"
      _StyleDefs(66)  =   ":id=29,.parent=0"
      _StyleDefs(67)  =   "Named:id=30:Heading"
      _StyleDefs(68)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H808000&,.fgcolor=&H80000012&"
      _StyleDefs(69)  =   ":id=30,.wraptext=-1"
      _StyleDefs(70)  =   "Named:id=31:Footing"
      _StyleDefs(71)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(72)  =   "Named:id=32:Selected"
      _StyleDefs(73)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(74)  =   "Named:id=33:Caption"
      _StyleDefs(75)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(76)  =   "Named:id=34:HighlightRow"
      _StyleDefs(77)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(78)  =   "Named:id=35:EvenRow"
      _StyleDefs(79)  =   ":id=35,.parent=29,.bgcolor=&HFFFF&"
      _StyleDefs(80)  =   "Named:id=36:OddRow"
      _StyleDefs(81)  =   ":id=36,.parent=29"
      _StyleDefs(82)  =   "Named:id=81:RecordSelector"
      _StyleDefs(83)  =   ":id=81,.parent=30"
      _StyleDefs(84)  =   "Named:id=84:FilterBar"
      _StyleDefs(85)  =   ":id=84,.parent=29"
   End
   Begin VB.Label Label1 
      Caption         =   $"Devices.frx":0442
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   1320
      TabIndex        =   2
      Top             =   5280
      Width           =   8415
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdRefresh 
      Height          =   720
      Left            =   240
      ToolTipText     =   "Atualizar Dispositivos não Cadastrados"
      Top             =   5140
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Devices.frx":04F9
      Effects         =   "Devices.frx":12E9
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdPrint 
      Height          =   720
      Left            =   10680
      ToolTipText     =   "Imprimir Módulos não Cadastrados"
      Top             =   5160
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Devices.frx":1301
      Effects         =   "Devices.frx":24AC
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdExit 
      Height          =   720
      Left            =   11520
      ToolTipText     =   "Fechar Visualização de Módulos não Cadastrados"
      Top             =   5160
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Devices.frx":24C4
      Effects         =   "Devices.frx":31C9
   End
End
Attribute VB_Name = "frmDevices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mList As XArrayDB
Private cD As clsDevice

Private Sub cmdRefresh_Click()
    Set lstDevice = Nothing
    Set lstDevice = New Collection
    DoEvents
    Device_Refresh
End Sub

Private Sub cmdRefresh_MouseEnter()
   cmdRefresh.SetRedraw = False
   cmdRefresh.GrayScale = lvicSepia
   cmdRefresh.LightnessPct = -20
   cmdRefresh.SetRedraw = True
End Sub

Private Sub cmdRefresh_MouseExit()
   cmdRefresh.SetRedraw = False
   cmdRefresh.GrayScale = lvicNoGrayScale
   cmdRefresh.LightnessPct = 0
   cmdRefresh.SetRedraw = True
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdExit_MouseEnter()
   cmdExit.SetRedraw = False
   cmdExit.GrayScale = lvicSepia
   cmdExit.LightnessPct = -20
   cmdExit.SetRedraw = True
End Sub

Private Sub cmdExit_MouseExit()
   cmdExit.SetRedraw = False
   cmdExit.GrayScale = lvicNoGrayScale
   cmdExit.LightnessPct = 0
   cmdExit.SetRedraw = True
End Sub

Private Sub cmdPrint_Click()
   Load frmDir
   Set frmDir.fCaller = Me
   frmDir.Caption = "Diretório para salvar Módulos/Devices não Cadastrados"
   frmDir.Show vbModal

   Dim fileName As String
   fileName = txtLocal.Text & "\NotCadastro_" & CStr(Day(Now)) & "_" & CStr(Month(Now)) & ".txt"
    
   On Error GoTo FileError
   Dim fso As New FileSystemObject
   Dim txtfile As TextStream
   Set txtfile = fso.CreateTextFile(fileName, True)
   txtfile.WriteLine ("Módulos/Devices não Cadastrados. ")
   txtfile.WriteBlankLines (1)
   txtfile.WriteLine ("Data: " + CStr(Day(Now)) + "/" + CStr(Month(Now)) + "/" + CStr(Year(Now)))
   txtfile.WriteBlankLines (3)
   txtfile.WriteLine ("Módulo " & Chr$(9) & "tipo " & Chr$(9) & "nome " & Chr$(9) & _
                      "sinal " & Chr$(9) & "ruído " & Chr$(9) & "comm " & _
                      Chr$(9) & "Receptor " & Chr$(9) & "evento")
   For Each cD In lstDevice
      ' Write a line with a módulo/device information.
      If cD.tipo = "Sensor" Or cD.tipo = "Repeater" Then
         txtfile.WriteLine (cD.Serial & Chr$(9) & cD.tipo & Chr$(9) & cD.name & _
                            Chr$(9) & cD.level & Chr$(9) & cD.margin & Chr$(9) & _
                            cD.comm & Chr$(9) & cD.recep & Chr$(9) & cD.evDate)
      Else
         txtfile.WriteLine (cD.Serial & Chr$(9) & cD.tipo & Chr$(9) & cD.name & _
                            Chr$(9) & cD.level & Chr$(9) & cD.margin & Chr$(9) & _
                            cD.comm & Chr$(9) & Chr$(9) & Chr$(9) & cD.evDate)
      End If
   Next
   txtfile.WriteBlankLines (1)
   txtfile.Close
   MsgBox "Dados salvo no arquivo: " & fileName, sxInformation, sxProname
   Exit Sub
FileError:
   MsgBox "Diretório inválido/protegido ou arquivo existente!, sxExclamation, sxProname"
End Sub

Private Sub cmdPrint_MouseEnter()
   cmdPrint.SetRedraw = False
   cmdPrint.GrayScale = lvicSepia
   cmdPrint.LightnessPct = -20
   cmdPrint.SetRedraw = True
End Sub

Private Sub cmdPrint_MouseExit()
   cmdPrint.SetRedraw = False
   cmdPrint.GrayScale = lvicNoGrayScale
   cmdPrint.LightnessPct = 0
   cmdPrint.SetRedraw = True
End Sub

Private Sub Form_Activate()
    Device_Refresh
End Sub

Private Sub Device_Refresh()
   Dim mRow As Integer
   Dim mCol As Integer
   Set mList = Nothing
   If lstDevice.Count >= 1 Then
      cmdPrint.Enabled = True
      ' Allocate space for rows, 8 columns
      Set mList = New XArrayDB
      mList.ReDim 0, lstDevice.Count - 1, 0, 7
      mRow = 0
      For Each cD In lstDevice
         With cD
            On Error Resume Next
            mList(mRow, 0) = .sUID
            mList(mRow, 1) = .Serial
            mList(mRow, 2) = .name
            mList(mRow, 3) = .level
            mList(mRow, 4) = .margin
            mList(mRow, 5) = .comm
            mList(mRow, 6) = Format(.evDate, "dd/mm")
            mList(mRow, 7) = Format(.evDate, "hh:mm:ss")
            On Error GoTo 0
         End With
         mRow = mRow + 1
      Next
      mList.QuickSort 0, mRow - 1, 6, XORDER_DESCEND, XTYPE_DATE, 7, XORDER_DESCEND, XTYPE_DATE
      tdbg1.Array = mList
      tdbg1.ReBind
   Else
      cmdPrint.Enabled = False
   End If
End Sub

Private Sub Form_Load()
   Dim success As Long
   success = SetWindowPos(frmLastEvents.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   curhwnd = frmLastEvents.hWnd
   Left = (Screen.Width - Width) / 2   ' Center form horizontally.
   Top = (Screen.Height - Height) / 2  ' Center form vertically.
   tdbg1.EvenRowStyle.BackColor = &H80FFFF
   tdbg1.OddRowStyle.BackColor = &HC0FFFF
End Sub

