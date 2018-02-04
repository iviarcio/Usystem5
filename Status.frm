VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A4749554-0441-4E64-8A03-3323601631C7}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmStatus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Situação Corrente das Zonas"
   ClientHeight    =   5730
   ClientLeft      =   2265
   ClientTop       =   2565
   ClientWidth     =   12000
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
   Icon            =   "Status.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5730
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin TrueOleDBGrid80.TDBGrid tdbg1 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   8493
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
      Columns(1).Caption=   "No. Serial"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   " Descrição"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Local"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   " Status"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   16
      Columns(5)._MaxComboItems=   9
      Columns(5).ValueItems(0)._DefaultItem=   0
      Columns(5).ValueItems(0).Value=   "0"
      Columns(5).ValueItems(0).Value.vt=   8
      Columns(5).ValueItems(0).DisplayValue=   "-"
      Columns(5).ValueItems(0).DisplayValue.vt=   8
      Columns(5).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems(1)._DefaultItem=   0
      Columns(5).ValueItems(1).Value=   "1"
      Columns(5).ValueItems(1).Value.vt=   8
      Columns(5).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(5).ValueItems(1).DisplayValue(0)=   "bHQAADYJAABCTTYJAAAAAAAANgAAACgAAAAwAAAAEAAAAAEAGAAAAAAAAAkAAAAAAAAAAAAAAAAA"
      Columns(5).ValueItems(1).DisplayValue(1)=   "AAAAAAAAAP8AAP8AAP8AAP8AAP8AAP//////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(3)=   "//////////////////////////////////////////////8AAP8AAP8AAP8AAP8AAP8AAP//////"
      Columns(5).ValueItems(1).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(6)=   "//////////8AAP8AAP8AAP8AAP8AAP8AAP//////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(8)=   "//////////////////////////////////////////////////8AAP8AAP8AAP8AAP8AAP8AAP//"
      Columns(5).ValueItems(1).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(10)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(11)=   "//////////////8AAP8AAP8AAP8AAP8AAP8AAP//////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(12)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(13)=   "//////////////////////////////////////////////////////8AAP8AAP8AAP8AAP8AAP8A"
      Columns(5).ValueItems(1).DisplayValue(14)=   "AP//////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(15)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(16)=   "//////////////////8AAP8AAP8AAP8AAP8AAP8AAP//////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(17)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(18)=   "//////////////////////////////////////////////////////////8AAP8AAP8AAP8AAP8A"
      Columns(5).ValueItems(1).DisplayValue(19)=   "AP8AAP//////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(20)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(21)=   "//////////////////////8AAP8AAP8AAP8AAP8AAP8AAP//////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(22)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(23)=   "//////////////////////////////////////////////////////////////8AAP8AAP8AAP8A"
      Columns(5).ValueItems(1).DisplayValue(24)=   "AP8AAP8AAP//////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(25)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(26)=   "//////////////////////////8AAP8AAP8AAP8AAP8AAP8AAP//////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(27)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(28)=   "//////////////////////////////////////////////////////////////////8AAP8AAP8A"
      Columns(5).ValueItems(1).DisplayValue(29)=   "AP8AAP8AAP8AAP//////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(30)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(31)=   "//////////////////////////////8AAP8AAP8AAP8AAP8AAP8AAP//////////////////////"
      Columns(5).ValueItems(1).DisplayValue(32)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(33)=   "//////////////////////////////////////////////////////////////////////8AAP8A"
      Columns(5).ValueItems(1).DisplayValue(34)=   "AP8AAP8AAP8AAP8AAP//////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(35)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(36)=   "//////////////////////////////////8AAP8AAP8AAP8AAP8AAP8AAP//////////////////"
      Columns(5).ValueItems(1).DisplayValue(37)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(38)=   "//////////////////////////////////////////////////////////////////////////8A"
      Columns(5).ValueItems(1).DisplayValue(39)=   "AP8AAP8AAP8AAP8AAP8AAP//////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(40)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(41)=   "//////////////////////////////////////8="
      Columns(5).ValueItems(1).DisplayValue.vt=   9
      Columns(5).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems(2)._DefaultItem=   0
      Columns(5).ValueItems(2).Value=   "2"
      Columns(5).ValueItems(2).Value.vt=   8
      Columns(5).ValueItems(2).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(5).ValueItems(2).DisplayValue(0)=   "bHQAADYJAABCTTYJAAAAAAAANgAAACgAAAAwAAAAEAAAAAEAGAAAAAAAAAkAAAAAAAAAAAAAAAAA"
      Columns(5).ValueItems(2).DisplayValue(1)=   "AAAAAAAAAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP//////////////////"
      Columns(5).ValueItems(2).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(3)=   "//////////////////////////////////////////////8AAP8AAP8AAP8AAP8AAP8AAP8AAAAA"
      Columns(5).ValueItems(2).DisplayValue(4)=   "AP8AAP8AAP8AAP8AAP8AAP//////////////////////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(6)=   "//////////8AAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP//////////////"
      Columns(5).ValueItems(2).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(8)=   "//////////////////////////////////////////////////8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(5).ValueItems(2).DisplayValue(9)=   "AAAAAP8AAP8AAP8AAP8AAP8AAP//////////////////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(10)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(11)=   "//////////////8AAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP//////////"
      Columns(5).ValueItems(2).DisplayValue(12)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(13)=   "//////////////////////////////////////////////////////8AAP8AAP8AAP8AAP8AAP8A"
      Columns(5).ValueItems(2).DisplayValue(14)=   "AP8AAAAAAP8AAP8AAP8AAP8AAP8AAP//////////////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(15)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(16)=   "//////////////////8AAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP//////"
      Columns(5).ValueItems(2).DisplayValue(17)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(18)=   "//////////////////////////////////////////////////////////8AAP8AAP8AAP8AAP8A"
      Columns(5).ValueItems(2).DisplayValue(19)=   "AP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP//////////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(20)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(21)=   "//////////////////////8AAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP//"
      Columns(5).ValueItems(2).DisplayValue(22)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(23)=   "//////////////////////////////////////////////////////////////8AAP8AAP8AAP8A"
      Columns(5).ValueItems(2).DisplayValue(24)=   "AP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP//////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(25)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(26)=   "//////////////////////////8AAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8A"
      Columns(5).ValueItems(2).DisplayValue(27)=   "AP//////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(28)=   "//////////////////////////////////////////////////////////////////8AAP8AAP8A"
      Columns(5).ValueItems(2).DisplayValue(29)=   "AP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP//////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(30)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(31)=   "//////////////////////////////8AAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8A"
      Columns(5).ValueItems(2).DisplayValue(32)=   "AP8AAP//////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(33)=   "//////////////////////////////////////////////////////////////////////8AAP8A"
      Columns(5).ValueItems(2).DisplayValue(34)=   "AP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP//////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(35)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(36)=   "//////////////////////////////////8AAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8A"
      Columns(5).ValueItems(2).DisplayValue(37)=   "AP8AAP8AAP//////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(38)=   "//////////////////////////////////////////////////////////////////////////8A"
      Columns(5).ValueItems(2).DisplayValue(39)=   "AP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP//////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(40)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(41)=   "//////////////////////////////////////8="
      Columns(5).ValueItems(2).DisplayValue.vt=   9
      Columns(5).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems(3)._DefaultItem=   0
      Columns(5).ValueItems(3).Value=   "3"
      Columns(5).ValueItems(3).Value.vt=   8
      Columns(5).ValueItems(3).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(5).ValueItems(3).DisplayValue(0)=   "bHQAADYJAABCTTYJAAAAAAAANgAAACgAAAAwAAAAEAAAAAEAGAAAAAAAAAkAAAAAAAAAAAAAAAAA"
      Columns(5).ValueItems(3).DisplayValue(1)=   "AAAAAAAAAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8A"
      Columns(5).ValueItems(3).DisplayValue(2)=   "AP8AAP8AAP//////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(3).DisplayValue(3)=   "//////////////////////////////////////////////8AAP8AAP8AAP8AAP8AAP8AAP8AAAAA"
      Columns(5).ValueItems(3).DisplayValue(4)=   "AP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP//////////////////////////"
      Columns(5).ValueItems(3).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(3).DisplayValue(6)=   "//////////8AAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8A"
      Columns(5).ValueItems(3).DisplayValue(7)=   "AP8AAP8AAP8AAP//////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(3).DisplayValue(8)=   "//////////////////////////////////////////////////8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(5).ValueItems(3).DisplayValue(9)=   "AAAAAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP//////////////////////"
      Columns(5).ValueItems(3).DisplayValue(10)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(3).DisplayValue(11)=   "//////////////8AAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8A"
      Columns(5).ValueItems(3).DisplayValue(12)=   "AP8AAP8AAP8AAP8AAP//////////////////////////////////////////////////////////"
      Columns(5).ValueItems(3).DisplayValue(13)=   "//////////////////////////////////////////////////////8AAP8AAP8AAP8AAP8AAP8A"
      Columns(5).ValueItems(3).DisplayValue(14)=   "AP8AAAAAAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP//////////////////"
      Columns(5).ValueItems(3).DisplayValue(15)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(3).DisplayValue(16)=   "//////////////////8AAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP8AAAAA"
      Columns(5).ValueItems(3).DisplayValue(17)=   "AP8AAP8AAP8AAP8AAP8AAP//////////////////////////////////////////////////////"
      Columns(5).ValueItems(3).DisplayValue(18)=   "//////////////////////////////////////////////////////////8AAP8AAP8AAP8AAP8A"
      Columns(5).ValueItems(3).DisplayValue(19)=   "AP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP//////////////"
      Columns(5).ValueItems(3).DisplayValue(20)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(3).DisplayValue(21)=   "//////////////////////8AAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(5).ValueItems(3).DisplayValue(22)=   "AAAAAP8AAP8AAP8AAP8AAP8AAP//////////////////////////////////////////////////"
      Columns(5).ValueItems(3).DisplayValue(23)=   "//////////////////////////////////////////////////////////////8AAP8AAP8AAP8A"
      Columns(5).ValueItems(3).DisplayValue(24)=   "AP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP//////////"
      Columns(5).ValueItems(3).DisplayValue(25)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(3).DisplayValue(26)=   "//////////////////////////8AAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8A"
      Columns(5).ValueItems(3).DisplayValue(27)=   "AP8AAAAAAP8AAP8AAP8AAP8AAP8AAP//////////////////////////////////////////////"
      Columns(5).ValueItems(3).DisplayValue(28)=   "//////////////////////////////////////////////////////////////////8AAP8AAP8A"
      Columns(5).ValueItems(3).DisplayValue(29)=   "AP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP//////"
      Columns(5).ValueItems(3).DisplayValue(30)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(3).DisplayValue(31)=   "//////////////////////////////8AAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8A"
      Columns(5).ValueItems(3).DisplayValue(32)=   "AP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP//////////////////////////////////////////"
      Columns(5).ValueItems(3).DisplayValue(33)=   "//////////////////////////////////////////////////////////////////////8AAP8A"
      Columns(5).ValueItems(3).DisplayValue(34)=   "AP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP//"
      Columns(5).ValueItems(3).DisplayValue(35)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(3).DisplayValue(36)=   "//////////////////////////////////8AAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8A"
      Columns(5).ValueItems(3).DisplayValue(37)=   "AP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP//////////////////////////////////////"
      Columns(5).ValueItems(3).DisplayValue(38)=   "//////////////////////////////////////////////////////////////////////////8A"
      Columns(5).ValueItems(3).DisplayValue(39)=   "AP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8AAP8AAAAAAP8AAP8AAP8AAP8AAP8A"
      Columns(5).ValueItems(3).DisplayValue(40)=   "AP//////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(3).DisplayValue(41)=   "//////////////////////////////////////8="
      Columns(5).ValueItems(3).DisplayValue.vt=   9
      Columns(5).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems(4)._DefaultItem=   0
      Columns(5).ValueItems(4).Value=   "4"
      Columns(5).ValueItems(4).Value.vt=   8
      Columns(5).ValueItems(4).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(5).ValueItems(4).DisplayValue(0)=   "bHQAADYJAABCTTYJAAAAAAAANgAAACgAAAAwAAAAEAAAAAEAGAAAAAAAAAkAAAAAAAAAAAAAAAAA"
      Columns(5).ValueItems(4).DisplayValue(1)=   "AAAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(4).DisplayValue(2)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/////////////////////////////////////"
      Columns(5).ValueItems(4).DisplayValue(3)=   "//////////////////////////////////////////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE"
      Columns(5).ValueItems(4).DisplayValue(4)=   "5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(4).DisplayValue(5)=   "5wD/////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(4).DisplayValue(6)=   "//////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE"
      Columns(5).ValueItems(4).DisplayValue(7)=   "5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/////////////////////////////////"
      Columns(5).ValueItems(4).DisplayValue(8)=   "//////////////////////////////////////////////////+E5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(5).ValueItems(4).DisplayValue(9)=   "AACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(4).DisplayValue(10)=   "5wCE5wD/////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(4).DisplayValue(11)=   "//////////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE"
      Columns(5).ValueItems(4).DisplayValue(12)=   "5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/////////////////////////////"
      Columns(5).ValueItems(4).DisplayValue(13)=   "//////////////////////////////////////////////////////+E5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(4).DisplayValue(14)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(4).DisplayValue(15)=   "5wCE5wCE5wD/////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(4).DisplayValue(16)=   "//////////////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE"
      Columns(5).ValueItems(4).DisplayValue(17)=   "5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/////////////////////////"
      Columns(5).ValueItems(4).DisplayValue(18)=   "//////////////////////////////////////////////////////////+E5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(4).DisplayValue(19)=   "5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE"
      Columns(5).ValueItems(4).DisplayValue(20)=   "5wCE5wCE5wCE5wD/////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(4).DisplayValue(21)=   "//////////////////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(5).ValueItems(4).DisplayValue(22)=   "AACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/////////////////////"
      Columns(5).ValueItems(4).DisplayValue(23)=   "//////////////////////////////////////////////////////////////+E5wCE5wCE5wCE"
      Columns(5).ValueItems(4).DisplayValue(24)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE"
      Columns(5).ValueItems(4).DisplayValue(25)=   "5wCE5wCE5wCE5wCE5wD/////////////////////////////////////////////////////////"
      Columns(5).ValueItems(4).DisplayValue(26)=   "//////////////////////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(4).DisplayValue(27)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/////////////////"
      Columns(5).ValueItems(4).DisplayValue(28)=   "//////////////////////////////////////////////////////////////////+E5wCE5wCE"
      Columns(5).ValueItems(4).DisplayValue(29)=   "5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE"
      Columns(5).ValueItems(4).DisplayValue(30)=   "5wCE5wCE5wCE5wCE5wCE5wD/////////////////////////////////////////////////////"
      Columns(5).ValueItems(4).DisplayValue(31)=   "//////////////////////////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(4).DisplayValue(32)=   "5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/////////////"
      Columns(5).ValueItems(4).DisplayValue(33)=   "//////////////////////////////////////////////////////////////////////+E5wCE"
      Columns(5).ValueItems(4).DisplayValue(34)=   "5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(5).ValueItems(4).DisplayValue(35)=   "AACE5wCE5wCE5wCE5wCE5wCE5wD/////////////////////////////////////////////////"
      Columns(5).ValueItems(4).DisplayValue(36)=   "//////////////////////////////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(4).DisplayValue(37)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/////////"
      Columns(5).ValueItems(4).DisplayValue(38)=   "//////////////////////////////////////////////////////////////////////////+E"
      Columns(5).ValueItems(4).DisplayValue(39)=   "5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(4).DisplayValue(40)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/////////////////////////////////////////////"
      Columns(5).ValueItems(4).DisplayValue(41)=   "//////////////////////////////////////8="
      Columns(5).ValueItems(4).DisplayValue.vt=   9
      Columns(5).ValueItems(4)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems(5)._DefaultItem=   0
      Columns(5).ValueItems(5).Value=   "5"
      Columns(5).ValueItems(5).Value.vt=   8
      Columns(5).ValueItems(5).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(5).ValueItems(5).DisplayValue(0)=   "bHQAADYJAABCTTYJAAAAAAAANgAAACgAAAAwAAAAEAAAAAEAGAAAAAAAAAkAAAAAAAAAAAAAAAAA"
      Columns(5).ValueItems(5).DisplayValue(1)=   "AAAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(5).DisplayValue(2)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/////////"
      Columns(5).ValueItems(5).DisplayValue(3)=   "//////////////////////////////////////////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE"
      Columns(5).ValueItems(5).DisplayValue(4)=   "5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(5).DisplayValue(5)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/////////////////////////////////////////////"
      Columns(5).ValueItems(5).DisplayValue(6)=   "//////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE"
      Columns(5).ValueItems(5).DisplayValue(7)=   "5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/////"
      Columns(5).ValueItems(5).DisplayValue(8)=   "//////////////////////////////////////////////////+E5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(5).ValueItems(5).DisplayValue(9)=   "AACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(5).DisplayValue(10)=   "5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/////////////////////////////////////////"
      Columns(5).ValueItems(5).DisplayValue(11)=   "//////////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE"
      Columns(5).ValueItems(5).DisplayValue(12)=   "5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/"
      Columns(5).ValueItems(5).DisplayValue(13)=   "//////////////////////////////////////////////////////+E5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(5).DisplayValue(14)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(5).DisplayValue(15)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/////////////////////////////////////"
      Columns(5).ValueItems(5).DisplayValue(16)=   "//////////////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE"
      Columns(5).ValueItems(5).DisplayValue(17)=   "5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(5).DisplayValue(18)=   "5wD///////////////////////////////////////////////////////+E5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(5).DisplayValue(19)=   "5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE"
      Columns(5).ValueItems(5).DisplayValue(20)=   "5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/////////////////////////////////"
      Columns(5).ValueItems(5).DisplayValue(21)=   "//////////////////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(5).ValueItems(5).DisplayValue(22)=   "AACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(5).DisplayValue(23)=   "5wCE5wD///////////////////////////////////////////////////////+E5wCE5wCE5wCE"
      Columns(5).ValueItems(5).DisplayValue(24)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE"
      Columns(5).ValueItems(5).DisplayValue(25)=   "5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/////////////////////////////"
      Columns(5).ValueItems(5).DisplayValue(26)=   "//////////////////////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(5).DisplayValue(27)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(5).DisplayValue(28)=   "5wCE5wCE5wD///////////////////////////////////////////////////////+E5wCE5wCE"
      Columns(5).ValueItems(5).DisplayValue(29)=   "5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE"
      Columns(5).ValueItems(5).DisplayValue(30)=   "5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/////////////////////////"
      Columns(5).ValueItems(5).DisplayValue(31)=   "//////////////////////////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(5).DisplayValue(32)=   "5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE"
      Columns(5).ValueItems(5).DisplayValue(33)=   "5wCE5wCE5wCE5wD///////////////////////////////////////////////////////+E5wCE"
      Columns(5).ValueItems(5).DisplayValue(34)=   "5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(5).ValueItems(5).DisplayValue(35)=   "AACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/////////////////////"
      Columns(5).ValueItems(5).DisplayValue(36)=   "//////////////////////////////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(5).DisplayValue(37)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE"
      Columns(5).ValueItems(5).DisplayValue(38)=   "5wCE5wCE5wCE5wCE5wD///////////////////////////////////////////////////////+E"
      Columns(5).ValueItems(5).DisplayValue(39)=   "5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(5).DisplayValue(40)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/////////////////"
      Columns(5).ValueItems(5).DisplayValue(41)=   "//////////////////////////////////////8="
      Columns(5).ValueItems(5).DisplayValue.vt=   9
      Columns(5).ValueItems(5)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems(6)._DefaultItem=   0
      Columns(5).ValueItems(6).Value=   "6"
      Columns(5).ValueItems(6).Value.vt=   8
      Columns(5).ValueItems(6).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(5).ValueItems(6).DisplayValue(0)=   "bHQAADYJAABCTTYJAAAAAAAANgAAACgAAAAwAAAAEAAAAAEAGAAAAAAAAAkAAAAAAAAAAAAAAAAA"
      Columns(5).ValueItems(6).DisplayValue(1)=   "AAAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(6).DisplayValue(2)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE"
      Columns(5).ValueItems(6).DisplayValue(3)=   "5wCE5wCE5wCE5wCE5wD///////////////////////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE"
      Columns(5).ValueItems(6).DisplayValue(4)=   "5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(6).DisplayValue(5)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/////////////////"
      Columns(5).ValueItems(6).DisplayValue(6)=   "//////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE"
      Columns(5).ValueItems(6).DisplayValue(7)=   "5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE"
      Columns(5).ValueItems(6).DisplayValue(8)=   "5wCE5wCE5wCE5wCE5wCE5wD///////////////////////////+E5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(5).ValueItems(6).DisplayValue(9)=   "AACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(6).DisplayValue(10)=   "5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/////////////"
      Columns(5).ValueItems(6).DisplayValue(11)=   "//////////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE"
      Columns(5).ValueItems(6).DisplayValue(12)=   "5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(5).ValueItems(6).DisplayValue(13)=   "AACE5wCE5wCE5wCE5wCE5wCE5wD///////////////////////////+E5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(6).DisplayValue(14)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(6).DisplayValue(15)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/////////"
      Columns(5).ValueItems(6).DisplayValue(16)=   "//////////////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE"
      Columns(5).ValueItems(6).DisplayValue(17)=   "5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(6).DisplayValue(18)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD///////////////////////////+E5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(6).DisplayValue(19)=   "5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE"
      Columns(5).ValueItems(6).DisplayValue(20)=   "5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/////"
      Columns(5).ValueItems(6).DisplayValue(21)=   "//////////////////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(5).ValueItems(6).DisplayValue(22)=   "AACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(6).DisplayValue(23)=   "5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD///////////////////////////+E5wCE5wCE5wCE"
      Columns(5).ValueItems(6).DisplayValue(24)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE"
      Columns(5).ValueItems(6).DisplayValue(25)=   "5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD/"
      Columns(5).ValueItems(6).DisplayValue(26)=   "//////////////////////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(6).DisplayValue(27)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(6).DisplayValue(28)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD///////////////////////////+E5wCE5wCE"
      Columns(5).ValueItems(6).DisplayValue(29)=   "5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE"
      Columns(5).ValueItems(6).DisplayValue(30)=   "5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(6).DisplayValue(31)=   "5wD///////////////////////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(6).DisplayValue(32)=   "5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE"
      Columns(5).ValueItems(6).DisplayValue(33)=   "5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD///////////////////////////+E5wCE"
      Columns(5).ValueItems(6).DisplayValue(34)=   "5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(5).ValueItems(6).DisplayValue(35)=   "AACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(6).DisplayValue(36)=   "5wCE5wD///////////////////////////+E5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(6).DisplayValue(37)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE"
      Columns(5).ValueItems(6).DisplayValue(38)=   "5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wD///////////////////////////+E"
      Columns(5).ValueItems(6).DisplayValue(39)=   "5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(6).DisplayValue(40)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(6).DisplayValue(41)=   "5wCE5wCE5wD///////////////////////////8="
      Columns(5).ValueItems(6).DisplayValue.vt=   9
      Columns(5).ValueItems(6)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems(7)._DefaultItem=   0
      Columns(5).ValueItems(7).Value=   "7"
      Columns(5).ValueItems(7).Value.vt=   8
      Columns(5).ValueItems(7).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(5).ValueItems(7).DisplayValue(0)=   "bHQAADYJAABCTTYJAAAAAAAANgAAACgAAAAwAAAAEAAAAAEAGAAAAAAAAAkAAAAAAAAAAAAAAAAA"
      Columns(5).ValueItems(7).DisplayValue(1)=   "AAAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(2)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE"
      Columns(5).ValueItems(7).DisplayValue(3)=   "5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAAAACE"
      Columns(5).ValueItems(7).DisplayValue(4)=   "5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(5)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(6)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(7)=   "5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE"
      Columns(5).ValueItems(7).DisplayValue(8)=   "5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(5).ValueItems(7).DisplayValue(9)=   "AACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(10)=   "5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(11)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE"
      Columns(5).ValueItems(7).DisplayValue(12)=   "5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(5).ValueItems(7).DisplayValue(13)=   "AACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(14)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(15)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE"
      Columns(5).ValueItems(7).DisplayValue(16)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE"
      Columns(5).ValueItems(7).DisplayValue(17)=   "5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(18)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(19)=   "5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(20)=   "5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE"
      Columns(5).ValueItems(7).DisplayValue(21)=   "5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(5).ValueItems(7).DisplayValue(22)=   "AACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(23)=   "5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(24)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE"
      Columns(5).ValueItems(7).DisplayValue(25)=   "5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(5).ValueItems(7).DisplayValue(26)=   "AACE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(27)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(28)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(29)=   "5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE"
      Columns(5).ValueItems(7).DisplayValue(30)=   "5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(31)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(32)=   "5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(33)=   "5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(34)=   "5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(5).ValueItems(7).DisplayValue(35)=   "AACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(36)=   "5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(37)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE"
      Columns(5).ValueItems(7).DisplayValue(38)=   "5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(39)=   "5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(40)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(7).DisplayValue(41)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wA="
      Columns(5).ValueItems(7).DisplayValue.vt=   9
      Columns(5).ValueItems(7)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems(8)._DefaultItem=   0
      Columns(5).ValueItems(8).Value=   "8"
      Columns(5).ValueItems(8).Value.vt=   8
      Columns(5).ValueItems(8).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(5).ValueItems(8).DisplayValue(0)=   "bHQAALYKAABCTbYKAAAAAAAANgAAACgAAAA3AAAAEAAAAAEAGAAAAAAAgAoAAAAAAAAAAAAAAAAA"
      Columns(5).ValueItems(8).DisplayValue(1)=   "AAAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(2)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE"
      Columns(5).ValueItems(8).DisplayValue(3)=   "5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(5).ValueItems(8).DisplayValue(4)=   "AACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(5)=   "5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(6)=   "5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE"
      Columns(5).ValueItems(8).DisplayValue(7)=   "5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(8)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(9)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE"
      Columns(5).ValueItems(8).DisplayValue(10)=   "5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(5).ValueItems(8).DisplayValue(11)=   "AACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(12)=   "5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(13)=   "5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE"
      Columns(5).ValueItems(8).DisplayValue(14)=   "5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(15)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(16)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE"
      Columns(5).ValueItems(8).DisplayValue(17)=   "5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(5).ValueItems(8).DisplayValue(18)=   "AACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(19)=   "5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(20)=   "5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE"
      Columns(5).ValueItems(8).DisplayValue(21)=   "5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(22)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(23)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE"
      Columns(5).ValueItems(8).DisplayValue(24)=   "5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(5).ValueItems(8).DisplayValue(25)=   "AACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(26)=   "5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(27)=   "5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE"
      Columns(5).ValueItems(8).DisplayValue(28)=   "5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(29)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(30)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE"
      Columns(5).ValueItems(8).DisplayValue(31)=   "5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(5).ValueItems(8).DisplayValue(32)=   "AACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(33)=   "5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(34)=   "5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE"
      Columns(5).ValueItems(8).DisplayValue(35)=   "5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(36)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(37)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE"
      Columns(5).ValueItems(8).DisplayValue(38)=   "5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(5).ValueItems(8).DisplayValue(39)=   "AACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(40)=   "5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(41)=   "5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE"
      Columns(5).ValueItems(8).DisplayValue(42)=   "5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(43)=   "5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(44)=   "5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE"
      Columns(5).ValueItems(8).DisplayValue(45)=   "5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAA"
      Columns(5).ValueItems(8).DisplayValue(46)=   "AACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(47)=   "5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE5wCE5wCE5wCE5wAAAACE5wCE5wCE"
      Columns(5).ValueItems(8).DisplayValue(48)=   "5wCE5wCE5wCE5wAAAAA="
      Columns(5).ValueItems(8).DisplayValue.vt=   9
      Columns(5).ValueItems(8)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems.Count=   9
      Columns(5).Caption=   " Sinal"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   16
      Columns(6)._MaxComboItems=   5
      Columns(6).ValueItems(0)._DefaultItem=   0
      Columns(6).ValueItems(0).Value=   "0"
      Columns(6).ValueItems(0).Value.vt=   8
      Columns(6).ValueItems(0).DisplayValue=   "Ok"
      Columns(6).ValueItems(0).DisplayValue.vt=   8
      Columns(6).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(6).ValueItems(1)._DefaultItem=   0
      Columns(6).ValueItems(1).Value=   "1"
      Columns(6).ValueItems(1).Value.vt=   8
      Columns(6).ValueItems(1).DisplayValue=   "Fraca"
      Columns(6).ValueItems(1).DisplayValue.vt=   8
      Columns(6).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(6).ValueItems(2)._DefaultItem=   0
      Columns(6).ValueItems(2).Value=   "2"
      Columns(6).ValueItems(2).Value.vt=   8
      Columns(6).ValueItems(2).DisplayValue=   "_"
      Columns(6).ValueItems(2).DisplayValue.vt=   8
      Columns(6).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
      Columns(6).ValueItems.Count=   3
      Columns(6).Caption=   " Bat."
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).AllowColSelect=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1191"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1085"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=256"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=1931"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1826"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=256"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=6429"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=6324"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=256"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=5874"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=5768"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=256"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=2275"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2170"
      Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=65792"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=1958"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1852"
      Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=65792"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(38)=   "Column(6).Width=1164"
      Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1058"
      Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=65792"
      Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
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
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
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
      _StyleDefs(19)  =   "RecordSelectorStyle:id=25,.parent=2,.namedParent=27"
      _StyleDefs(20)  =   "FilterBarStyle:id=28,.parent=1,.namedParent=72"
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
      _StyleDefs(31)  =   "Splits(0).RecordSelectorStyle:id=26,.parent=25"
      _StyleDefs(32)  =   "Splits(0).FilterBarStyle:id=71,.parent=28"
      _StyleDefs(33)  =   "Splits(0).Columns(0).Style:id=24,.parent=37,.alignment=0,.locked=0"
      _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=38,.alignment=0"
      _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=39,.alignment=3"
      _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=41"
      _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=50,.parent=37,.alignment=0,.locked=0"
      _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=38,.alignment=0"
      _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=39,.alignment=3"
      _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=41"
      _StyleDefs(41)  =   "Splits(0).Columns(2).Style:id=54,.parent=37,.alignment=0,.locked=0"
      _StyleDefs(42)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=38,.alignment=0"
      _StyleDefs(43)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=39,.alignment=3"
      _StyleDefs(44)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=41"
      _StyleDefs(45)  =   "Splits(0).Columns(3).Style:id=58,.parent=37,.alignment=0,.locked=0"
      _StyleDefs(46)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=38,.alignment=0"
      _StyleDefs(47)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=39,.alignment=3"
      _StyleDefs(48)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=41"
      _StyleDefs(49)  =   "Splits(0).Columns(4).Style:id=70,.parent=37"
      _StyleDefs(50)  =   "Splits(0).Columns(4).HeadingStyle:id=67,.parent=38"
      _StyleDefs(51)  =   "Splits(0).Columns(4).FooterStyle:id=68,.parent=39"
      _StyleDefs(52)  =   "Splits(0).Columns(4).EditorStyle:id=69,.parent=41"
      _StyleDefs(53)  =   "Splits(0).Columns(5).Style:id=66,.parent=37"
      _StyleDefs(54)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=38"
      _StyleDefs(55)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=39"
      _StyleDefs(56)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=41"
      _StyleDefs(57)  =   "Splits(0).Columns(6).Style:id=62,.parent=37"
      _StyleDefs(58)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=38"
      _StyleDefs(59)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=39"
      _StyleDefs(60)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=41"
      _StyleDefs(61)  =   "Named:id=29:Normal"
      _StyleDefs(62)  =   ":id=29,.parent=0"
      _StyleDefs(63)  =   "Named:id=30:Heading"
      _StyleDefs(64)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H808000&,.fgcolor=&H80000012&"
      _StyleDefs(65)  =   ":id=30,.wraptext=-1"
      _StyleDefs(66)  =   "Named:id=31:Footing"
      _StyleDefs(67)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(68)  =   "Named:id=32:Selected"
      _StyleDefs(69)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(70)  =   "Named:id=33:Caption"
      _StyleDefs(71)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(72)  =   "Named:id=34:HighlightRow"
      _StyleDefs(73)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(74)  =   "Named:id=35:EvenRow"
      _StyleDefs(75)  =   ":id=35,.parent=29,.bgcolor=&HFFFF&"
      _StyleDefs(76)  =   "Named:id=36:OddRow"
      _StyleDefs(77)  =   ":id=36,.parent=29"
      _StyleDefs(78)  =   "Named:id=27:RecordSelector"
      _StyleDefs(79)  =   ":id=27,.parent=30"
      _StyleDefs(80)  =   "Named:id=72:FilterBar"
      _StyleDefs(81)  =   ":id=72,.parent=29"
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdExit 
      Height          =   720
      Left            =   11280
      ToolTipText     =   "Fechar Visualização de Situação de Zonas"
      Top             =   4920
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Status.frx":0442
      Effects         =   "Status.frx":1147
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl cmdPrint 
      Height          =   720
      Left            =   10440
      ToolTipText     =   "Visualizar Situação Corrente das Zonas"
      Top             =   4920
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Status.frx":115F
      Effects         =   "Status.frx":230A
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public localModule As New Collection
Public fZona As typeSensor
Private mList As XArrayDB

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
   Screen.MousePointer = vbHourglass
   'First prepare the table
   Dim lcmd As New ADODB.Command
   Set lcmd.ActiveConnection = cnDB
   lcmd.ActiveConnection.BeginTrans
   lcmd.CommandType = adCmdText
   lcmd.CommandText = "DELETE FROM StatusCorrente"
   lcmd.Execute
   Dim fP As clsPiso
   Dim fE As clsEntity
   Dim cM As clsModule
   If localModule.Count > 0 Then
      For Each cM In localModule
         With cM
            If .mTipo = fZona Then
               Set fE = lstEntity.Item(CStr(.mEntity))
               Set fP = lstPiso.Item(CStr(fE.floor))
               lcmd.CommandText = "INSERT INTO StatusCorrente (Numero_Sensor, Status, Sinal, Bateria, Tampa) VALUES ('" & _
                     .Serial_Number & "', " & .SZona & ", " & .NivelSinal & ", " & .SLowBat & ", " & .STampa & ")"
               lcmd.Execute
            End If
         End With
      Next
      lcmd.ActiveConnection.CommitTrans
      'Now Print the report
      
      Screen.MousePointer = vbHourglass
      Dim frm As New frmViewReport9
      frm.SetTipo = g_iRptSCZonas
      frm.WindowState = vbMaximized
      frm.Show
      Screen.MousePointer = vbDefault
      
   Else
      MsgBox "Não há nenhuma Zona do tipo selecionado!", sxExclamation, sxProname
   End If
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

Private Sub SetCaption()
    If fZona >= 0 And fZona <= 6 Then
        frmStatus.Caption = "Situação corrente dos Sensores de " + strTipo(fZona)
    Else
         frmStatus.Caption = "Situação corrente dos " + strTipo(fZona) + "s"
    End If
End Sub

Private Sub Form_Activate()
   Dim cM As clsModule
   Dim cE As clsEntity
   Dim mRow As Integer
   Dim mCol As Integer
   Set mList = Nothing
   
   SetCaption
   
   cmdPrint.Enabled = (localModule.Count >= 1)
   If localModule.Count >= 1 Then
      ' Allocate space for rows, 7 columns
      Set mList = New XArrayDB
      mList.ReDim 0, localModule.Count - 1, 0, 6
      mRow = 0
      For Each cM In localModule
         With cM
            If .mTipo = fZona Then
               mList(mRow, 0) = CStr(.mEntity)
               mList(mRow, 1) = CStr(.Serial_Number)
               mList(mRow, 2) = .mLocal
               On Error Resume Next
               Set cE = lstEntity.Item(CStr(.mEntity))
               mList(mRow, 3) = cE.vDescr
               On Error GoTo 0
               mList(mRow, 4) = strStatus(.SZona)
               mList(mRow, 5) = .NivelSinal
               mList(mRow, 6) = .SLowBat
               mRow = mRow + 1
            End If
         End With
      Next
      If mRow > 0 Then
         mList.ReDim 0, mRow - 1, 0, 6
         mList.QuickSort 0, mRow - 1, 1, XORDER_ASCEND, XTYPE_INTEGER, 2, XORDER_ASCEND, XTYPE_INTEGER
         cmdPrint.Enabled = True
      Else
         cmdPrint.Enabled = False
      End If
   End If
   tdbg1.Array = mList
   tdbg1.ReBind
End Sub

Private Sub Form_Load()
   Dim success As Long
   success = SetWindowPos(frmStatus.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   curhwnd = frmStatus.hWnd
   Left = (Screen.Width - Width) / 2   ' Center form horizontally.
   Top = (Screen.Height - Height) / 2  ' Center form vertically.
   tdbg1.EvenRowStyle.BackColor = &H80FFFF
   tdbg1.OddRowStyle.BackColor = &HC0FFFF
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mList = Nothing
End Sub

