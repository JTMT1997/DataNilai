VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "Form Nilai"
   ClientHeight    =   10050
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   19530
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   10050
   ScaleWidth      =   19530
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3615
      Left            =   600
      TabIndex        =   35
      Top             =   6600
      Width           =   17175
      _ExtentX        =   30295
      _ExtentY        =   6376
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "zz"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   3615
      Left            =   18120
      TabIndex        =   34
      Top             =   6600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   6376
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   -2147483639
      HeadLines       =   1
      RowHeight       =   36
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Calligraphy"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H80000015&
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   11640
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Keluar"
      Height          =   495
      Left            =   14880
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5400
      Width           =   4455
   End
   Begin VB.TextBox Text12 
      Height          =   615
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000010&
      Caption         =   "Reset Form"
      Height          =   495
      Left            =   14880
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4320
      Width           =   4455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hapus Matkul"
      Height          =   495
      Left            =   14880
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3240
      Width           =   4455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Edit Matkul"
      Height          =   495
      Left            =   14880
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2160
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Tambah Matkul "
      Height          =   495
      Left            =   14880
      MaskColor       =   &H00FFFF00&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aksi"
      Height          =   5655
      Left            =   14400
      TabIndex        =   23
      Top             =   600
      Width           =   5535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   1920
      Top             =   8280
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H80000015&
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   11640
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H80000015&
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Konversi Nilai"
      Height          =   5535
      Left            =   10440
      TabIndex        =   17
      Top             =   600
      Width           =   2895
      Begin VB.TextBox Text10 
         BackColor       =   &H80000015&
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H80000015&
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3360
      Width           =   2415
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H80000015&
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H80000015&
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000015&
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   2160
      TabIndex        =   5
      Top             =   4920
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   2160
      TabIndex        =   4
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      Caption         =   "Nilai Murni"
      Height          =   5895
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   4095
      Begin VB.TextBox Text13 
         Height          =   405
         Left            =   1560
         TabIndex        =   31
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   1560
         TabIndex        =   3
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   1560
         TabIndex        =   2
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label12 
         BackColor       =   &H000000FF&
         Caption         =   "SKS"
         Height          =   495
         Left            =   360
         TabIndex        =   32
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H000000FF&
         Caption         =   "Nama Mata Kuliah"
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         Caption         =   "Tugas"
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H000000FF&
         Caption         =   "Kehadiran"
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H000000FF&
         Caption         =   "Uas"
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Caption         =   "UTS"
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   3120
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Persentase"
      Height          =   5775
      Left            =   5280
      TabIndex        =   1
      Top             =   480
      Width           =   3975
      Begin VB.Label Label9 
         BackColor       =   &H0000FFFF&
         Caption         =   "Jumlah Nilai"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackColor       =   &H0000FFFF&
         Caption         =   "Kehadiran"
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000FFFF&
         Caption         =   "UAS"
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label6 
         BackColor       =   &H0000FFFF&
         Caption         =   "UTS"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         Caption         =   "Tugas"
         Height          =   315
         Left            =   360
         TabIndex        =   13
         Top             =   600
         Width           =   570
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   615
      Left            =   18240
      Top             =   7560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim koneksi As String
Sub BukaDB()
Set conn = New ADODB.Connection
Set RS = New ADODB.Recordset
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dataNilai.mdb;"
End Sub
Sub readdata()
Dim result As String
result = "Select namaMataKuliah as NamaMataKuliah, bobotSKS as BobotSKS, tugasNM as NilaiTugas, utsNM as NilaiUTS, uasNM as NilaiUAS, kehadiranNM as nilaiKehadiran , tugasPs as PresensiNilaiTugas , utsPs as PresensiNilaiUTS, uasPs as PresensiNilaiUAS , kehadiranPs as PresensiNilaiKehadiran, jNilai as JumlahNilaiPersentase , kNilai as AngkaNilai, nAkhir as NilaiAkhir, bobotxNa as BobotNilaiSKS from dNilai"
conn.Execute (result)
Adodc1.RecordSource = result
Set DataGrid1.DataSource = Adodc1
Adodc1.Refresh
End Sub
Sub ambilIPK()
Dim result As String
result = "Select sum(bobotxNa) / sum(bobotSKS) as IPK from dNilai"
conn.Execute (result)
Adodc2.RecordSource = result
Set DataGrid2.DataSource = Adodc2
Adodc2.Refresh
End Sub
Sub KosongkanData()
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    Text7 = ""
    Text8 = ""
    Text9 = ""
    Text10 = ""
    Text11 = ""
    Text12 = ""
    Text13 = ""
    Text14 = ""
End Sub
Private Sub Command1_Click()
If MsgBox("Apakah Ingin keluar dari aplikasi? ", 36, "Informasi") = vbYes Then
            Unload Me
End If
End Sub

Private Sub Command2_Click()
 If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Or Text8 = "" Or Text9 = "" Or Text10 = "" Or Text11 = "" Or Text12 = "" Or Text13 = "" Or Text14 = "" Then
    
    MsgBox "Silahkan isi data terlebih dahulu!", vbQuestion, "Perhatian"
    
Else
    result = "select * from dNilai where namaMataKuliah = '" & Text12 & "' "
    Set RS = conn.Execute(result)
    If RS.EOF Then
        TambahData = "Insert into dNilai values('" & Text12 & "', '" & Text13 & "','" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "', '" & Text6 & "', '" & Text7 & "', '" & Text8 & "', '" & Text9 & "','" & Text10 & "', '" & Text11 & "' , '" & Text14 & "')"
        conn.Execute (TambahData)
        MsgBox "Tambah Data Berhasil", vbInformation, "Berhasil"
        Call KosongkanData
     Else
       
        MsgBox "Nama Mata Kuliah Sudah Tersedia", vbCritical, "Perhatian"
     End If
        Call readdata
        Call ambilIPK
End If
    
End Sub

Private Sub Command3_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Or Text8 = "" Or Text9 = "" Or Text10 = "" Or Text11 = "" Or Text12 = "" Or Text13 = "" Or Text14 = "" Then
        MsgBox "Pastikan data tidak kosong", vbExclamation
    Else
        
        Call BukaDB
        Dim UpdateData
        Data = MsgBox("Apkah anda yakin mengedit?", vbQuestion + vbYesNoCancel)
        If Data = vbYes Then

        UpdateData = "update dNilai SET bobotSKS = '" & Text13 & "', bobotxNa = '" & Text14 & "', tugasNM = '" & Text1 & "', utsNM = '" & Text2 & "', uasNM = '" & Text3 & "', kehadiranNM= '" & Text4 & "', tugasPs = '" & Text5 & "' , utsPs = '" & Text6 & "' , uasPs = '" & Text7 & "' , kehadiranPs= '" & Text8 & "' , jNilai= '" & Text9 & "' , kNilai= '" & Text10 & "' , nAkhir = '" & Text11 & "'where namaMataKuliah = '" & Text12 & "' "

        conn.Execute UpdateData
        MsgBox "Update Data Berhasil", vbInformation, "berhasil"
       
        ElseIf Data = vbCancel Then
        MsgBox "Data tidak berubah", vbCritical, "Data berhasil di cancel"
        Call KosongkanData
        
        ElseIf Data = vbNo Then
        MsgBox "Data tidak berubah", vbExclamation, "Gagal"
        
        Call KosongkanData
          End If
        Call readdata
        Call ambilIPK
    End If
End Sub
Private Sub Command4_Click()
 If Text12 = "" Then
    MsgBox ("Silahkan Isi Mata Kuliah Terlebih Dahulu")
 Else
    result = "select * from dNilai where namaMataKuliah = '" & Text12 & "' "
    Set RS = conn.Execute(result)
    If RS.EOF Then
       MsgBox "Nama Mata Kuliah Tidak Tersedia", vbCritical, "Perhatian"
        Call KosongkanData
     Else
     Jawab = MsgBox("Apa anda yakin?", vbQuestion + vbYesNo)
     If Jawab = vbYes Then
        HapusData = "delete from dNilai where namaMataKuliah = '" & Text12 & "'"
        conn.Execute HapusData
      
         MsgBox "Data Berhasil Di Hapus", vbInformation, "Berhasil"
         Call KosongkanData
     End If
     End If
        Call readdata
        Call ambilIPK
 End If
    
   
End Sub

Private Sub Command5_Click()
DataReport1.Show
End Sub

Private Sub Command6_Click()
Call KosongkanData
MsgBox "Form Berhasil di Reset", vbInformation, "Berhasil"
End Sub



Private Sub DataGrid1_Click()




Text12.Text = DataGrid1.Columns(0).Text
Text13.Text = DataGrid1.Columns(1).Text
Text1.Text = DataGrid1.Columns(2).Text
Text2.Text = DataGrid1.Columns(3).Text
Text3.Text = DataGrid1.Columns(4).Text
Text4.Text = DataGrid1.Columns(5).Text
Text5.Text = DataGrid1.Columns(6).Text
Text6.Text = DataGrid1.Columns(7).Text
Text7.Text = DataGrid1.Columns(8).Text
Text8.Text = DataGrid1.Columns(9).Text
Text9.Text = DataGrid1.Columns(10).Text
Text10.Text = DataGrid1.Columns(11).Text
Text11.Text = DataGrid1.Columns(12).Text


End Sub
Private Sub Form_Load()
Call KosongkanData

koneksi = "provider = microsoft.jet.oledb.4.0; data source = " & App.Path & "\dataNilai.mdb"
Adodc1.ConnectionString = koneksi
Adodc2.ConnectionString = koneksi
conn.Open koneksi
Call readdata
Call ambilIPK

End Sub

Private Sub Text1_Change()
texttukem = Val(Text1) * 20 / 100
Text5.Text = texttukem
Text9.Text = Val(Text5.Text) + Val(Text6.Text) + Val(Text7.Text) + Val(Text8.Text)
If Val(Text9.Text) > 79.99 Then
Text10.Text = "A"
Text11.Text = 4#
Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 76.99 Then
     Text10.Text = "A-"
     Text11.Text = 3.7
     Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 73.99 Then
    Text10.Text = "B+"
    Text11.Text = 3.3
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 70.99 Then
    Text10.Text = "B"
    Text11.Text = 3#
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 67.99 Then
    Text10.Text = "B-"
    Text11.Text = 2.7
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 63.99 Then
    Text10.Text = "C+"
    Text11.Text = 2.3
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 55.99 Then
    Text10.Text = "C"
    Text11.Text = 2#
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 45.99 Then
    Text10.Text = "D"
    Text11.Text = 1#
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) < 45.99 Then
    Text10.Text = "E"
    Text11.Text = 0
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
End If

End Sub

Private Sub Text2_Change()
texttukem = Val(Text2) * 20 / 100
Text6.Text = texttukem
Text9.Text = Val(Text5.Text) + Val(Text6.Text) + Val(Text7.Text) + Val(Text8.Text)
If Val(Text9.Text) > 79.99 Then
Text10.Text = "A"
Text11.Text = 4#
Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 76.99 Then
     Text10.Text = "A-"
     Text11.Text = 3.7
     Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 73.99 Then
    Text10.Text = "B+"
    Text11.Text = 3.3
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 70.99 Then
    Text10.Text = "B"
    Text11.Text = 3#
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 67.99 Then
    Text10.Text = "B-"
    Text11.Text = 2.7
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 63.99 Then
    Text10.Text = "C+"
    Text11.Text = 2.3
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 55.99 Then
    Text10.Text = "C"
    Text11.Text = 2#
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 45.99 Then
    Text10.Text = "D"
    Text11.Text = 1#
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) < 45.99 Then
    Text10.Text = "E"
    Text11.Text = 0
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
End If

End Sub

Private Sub Text3_Change()
texttukem = Val(Text3) * 30 / 100
Text7.Text = texttukem
Text9.Text = Val(Text5.Text) + Val(Text6.Text) + Val(Text7.Text) + Val(Text8.Text)
If Val(Text9.Text) > 79.99 Then
Text10.Text = "A"
Text11.Text = 4#
Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 76.99 Then
     Text10.Text = "A-"
     Text11.Text = 3.7
     Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 73.99 Then
    Text10.Text = "B+"
    Text11.Text = 3.3
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 70.99 Then
    Text10.Text = "B"
    Text11.Text = 3#
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 67.99 Then
    Text10.Text = "B-"
    Text11.Text = 2.7
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 63.99 Then
    Text10.Text = "C+"
    Text11.Text = 2.3
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 55.99 Then
    Text10.Text = "C"
    Text11.Text = 2#
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 45.99 Then
    Text10.Text = "D"
    Text11.Text = 1#
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) < 45.99 Then
    Text10.Text = "E"
    Text11.Text = 0
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
End If
End Sub

Private Sub Text4_Change()
texttukem = Val(Text4) * 30 / 100
Text8.Text = texttukem
Text9.Text = Val(Text5.Text) + Val(Text6.Text) + Val(Text7.Text) + Val(Text8.Text)
If Val(Text9.Text) > 79.99 Then
Text10.Text = "A"
Text11.Text = 4#
Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 76.99 Then
     Text10.Text = "A-"
     Text11.Text = 3.7
     Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 73.99 Then
    Text10.Text = "B+"
    Text11.Text = 3.3
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 70.99 Then
    Text10.Text = "B"
    Text11.Text = 3#
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 67.99 Then
    Text10.Text = "B-"
    Text11.Text = 2.7
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 63.99 Then
    Text10.Text = "C+"
    Text11.Text = 2.3
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 55.99 Then
    Text10.Text = "C"
    Text11.Text = 2#
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) > 45.99 Then
    Text10.Text = "D"
    Text11.Text = 1#
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
ElseIf Val(Text9.Text) < 45.99 Then
    Text10.Text = "E"
    Text11.Text = 0
    Text14.Text = Val(Text13.Text) * Val(Text11.Text)
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call BukaDB
        RS.Open "Select * From dNilai where namaMataKuliah = '" & Text12 & "'", KoneksiDB
        If Not RS.EOF Then
            Text13 = RS!bobotSKS
            Text1 = RS!tugasNM
            Text2 = RS!utsNM
            Text3 = RS!uasNM
            Text4 = RS!kehadiranNM
            Text5 = RS!tugasPs
            Text6 = RS!utsPs
            Text7 = RS!uasPs
            Text8 = RS!kehadiranPs
            Text9 = RS!jNilai
            Text10 = RS!kNilai
            Text11 = RS!nAkhir
        Else
        End If
    End If
End Sub
