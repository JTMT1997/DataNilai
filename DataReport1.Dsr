VERSION 5.00
Begin {78E93846-85FD-11D0-8487-00A0C90DC8A9} DataReport1 
   Bindings        =   "DataReport1.dsx":0000
   Caption         =   "DataReport1"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20160
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35560
   _ExtentY        =   19288
   _Version        =   393216
   _DesignerVersion=   100684101
   ReportWidth     =   10680
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   GridX           =   10
   GridY           =   10
   LeftMargin      =   1440
   RightMargin     =   1440
   TopMargin       =   1440
   BottomMargin    =   1440
   NumSections     =   5
   SectionCode0    =   1
   BeginProperty Section0 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "Section4"
      Object.Height          =   690
      NumControls     =   1
      ItemType0       =   3
      BeginProperty Item0 {1C13A8E1-A0B6-11D0-848E-00A0C90DC8A9} 
         _Version        =   393216
         Name            =   "Label1"
         Object.Left            =   3024
         Object.Top             =   144
         Object.Width           =   2730
         Object.Height          =   420
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bell MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Object.Caption         =   "Data Nilai"
      EndProperty
   EndProperty
   SectionCode1    =   2
   BeginProperty Section1 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "Section2"
      Object.Height          =   384
      NumControls     =   2
      ItemType0       =   3
      BeginProperty Item0 {1C13A8E1-A0B6-11D0-848E-00A0C90DC8A9} 
         _Version        =   393216
         Name            =   "Label2"
         Object.Width           =   1440
         Object.Height          =   240
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Object.Caption         =   "Nama"
      EndProperty
      ItemType1       =   3
      BeginProperty Item1 {1C13A8E1-A0B6-11D0-848E-00A0C90DC8A9} 
         _Version        =   393216
         Name            =   "Label3"
         Object.Left            =   2160
         Object.Width           =   1440
         Object.Height          =   240
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Object.Caption         =   "Nama"
      EndProperty
   EndProperty
   SectionCode2    =   4
   BeginProperty Section2 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "Section1"
      Object.Height          =   1440
      NumControls     =   0
   EndProperty
   SectionCode3    =   7
   BeginProperty Section3 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "Section3"
      Object.Height          =   360
      NumControls     =   0
   EndProperty
   SectionCode4    =   8
   BeginProperty Section4 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "Section5"
      Object.Height          =   3555
      NumControls     =   0
   EndProperty
End
Attribute VB_Name = "DataReport1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DataReport_Error(ByVal JobType As MSDataReportLib.AsyncTypeConstants, ByVal Cookie As Long, ByVal ErrObj As MSDataReportLib.RptError, ShowError As Boolean)
koneksi = "provider = microsoft.jet.oledb.4.0; data source = " & App.Path & "\dataNilai.mdb"
Adodc1.ConnectionString = koneksi
End Sub
