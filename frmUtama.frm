VERSION 5.00
Begin VB.Form frmUtama 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tree"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmUtama.frx":0000
   ScaleHeight     =   7605
   ScaleWidth      =   17415
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pctTree 
      BackColor       =   &H0080FF80&
      Height          =   6255
      Left            =   120
      ScaleHeight     =   6195
      ScaleWidth      =   7995
      TabIndex        =   7
      Top             =   1200
      Width           =   8055
      Begin VB.Line lnTree 
         Index           =   0
         Visible         =   0   'False
         X1              =   2280
         X2              =   3120
         Y1              =   4320
         Y2              =   3000
      End
      Begin VB.Label lblTree 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.PictureBox pctBack 
      BackColor       =   &H0080FF80&
      Height          =   6255
      Left            =   8280
      ScaleHeight     =   6195
      ScaleWidth      =   8955
      TabIndex        =   5
      Top             =   1200
      Width           =   9015
      Begin VB.Line lnKode 
         Index           =   0
         Visible         =   0   'False
         X1              =   960
         X2              =   1440
         Y1              =   4320
         Y2              =   3240
      End
      Begin VB.Label lblKode 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.CommandButton cmdUndo 
         Caption         =   "&Undo"
         Height          =   495
         Left            =   7200
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdInput 
         Caption         =   "&Input"
         Default         =   -1  'True
         Height          =   495
         Left            =   5400
         MouseIcon       =   "frmUtama.frx":030A
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtKode 
         Height          =   285
         Left            =   840
         MaxLength       =   5
         TabIndex        =   2
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Kode   :"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Menentukan lebar garis root
Const Lebar = 2000 '3500

Const BatasAtas = 450
Const Max = 1000

Dim n As Integer
Dim Parent, Root As Boolean
Dim Isi As Integer
Dim NoRoot As Integer
Dim Pengali As Integer

Dim Tingkat, Akhir As Integer

Dim Tekan As Boolean

Dim RootKiri(Max), RootKanan(Max) As Integer

Dim nilaiX, nilaiY, Objek As Integer

'Dim Lebar As Integer 'Variabel untuk deklarasi lebar antar root dengan parentnya
Dim Kiri(Max), Kanan(Max), Atas(Max) As Integer
Dim Level(Max) As Integer
Dim Taruh As Boolean 'Variabel untuk mengetahui apakah sudah diinput

Private Sub cmdInput_Click()
    If Trim(txtKode.Text) <> "" Then
        Taruh = False
        Me.MousePointer = 99
    End If
End Sub

Private Sub cmdUndo_Click()
        lblKode(Objek).Left = nilaiX
        lblKode(Objek).Top = nilaiY

        'Menggambar Garis
        lnKode(Objek).X1 = lblKode(Objek).Left + (lblKode(Objek).Width / 2)
        lnKode(Objek).Y1 = lblKode(Objek).Top + (lblKode(Objek).Height / 2)
        
        'Mengatur garis Root
        If RootKiri(Objek) <> 0 Then
            lnKode(RootKiri(Objek)).X2 = lblKode(Objek).Left + (lblKode(Objek).Width / 2)
            lnKode(RootKiri(Objek)).Y2 = lblKode(Objek).Top + (lblKode(Objek).Height / 2)
        End If
        If RootKanan(Objek) <> 0 Then
            lnKode(RootKanan(Objek)).X2 = lblKode(Objek).Left + (lblKode(Objek).Width / 2)
            lnKode(RootKanan(Objek)).Y2 = lblKode(Objek).Top + (lblKode(Objek).Height / 2)
        End If
End Sub

Private Sub Form_Load()
    Rem Deklarasi Tree input
    Parent = False
    
    Taruh = False
    
    Rem Deklarasi Tree seimbang
    
    n = 1
    Isi = 0
    Pengali = 1
    Root = False
    NoRoot = 0
    
    Tingkat = 1
    Akhir = 1
    
    Tekan = False
    
    nilaiX = 0
    nilaiY = 0
End Sub

Private Sub lblKode_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Taruh = True Then Tekan = True
    
    'Memasukkan variabel-variabel Undo
    Objek = Index
    nilaiX = X
    nilaiY = Y
End Sub

Private Sub lblKode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Tekan = True Then
        'Me.Caption = CStr(X) & "," & CStr(Y)
        lblKode(Index).Left = X
        lblKode(Index).Top = Y

        'Menggambar Garis
        lnKode(Index).X1 = lblKode(Index).Left + (lblKode(Index).Width / 2)
        lnKode(Index).Y1 = lblKode(Index).Top + (lblKode(Index).Height / 2)
        
        'Mengatur garis Root
        If RootKiri(Index) <> 0 Then
            lnKode(RootKiri(Index)).X2 = lblKode(Index).Left + (lblKode(Index).Width / 2)
            lnKode(RootKiri(Index)).Y2 = lblKode(Index).Top + (lblKode(Index).Height / 2)
        End If
        If RootKanan(Index) <> 0 Then
            lnKode(RootKanan(Index)).X2 = lblKode(Index).Left + (lblKode(Index).Width / 2)
            lnKode(RootKanan(Index)).Y2 = lblKode(Index).Top + (lblKode(Index).Height / 2)
        End If
    End If
End Sub

Private Sub lblKode_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Tekan = False
End Sub

Sub IsiKanan()
    If Trim(txtKode.Text) <> "" Then
        If Root = False Then
            lblKode(0).Visible = True
            lblKode(0).Left = ((pctBack.Width - lblKode(0).Width) / 2)
            'lblKode(0).Left = pctBack.Left + (((pctBack.Width - pctBack.Left) - lblKode(0).Width) / 2)
            lblKode(0).Caption = txtKode.Text
            Root = True
        
            Tingkat = 4
            Akhir = 2
            RootKiri(0) = 0
            RootKanan(0) = 0
        Else
            Load lblKode(n)
        
            'Membuat Variabel Root Kiri dan Kanan
            RootKiri(n) = 0
            RootKanan(n) = 0
            
            'Menentukan posisi atas objek
            lblKode(n).Top = lblKode(NoRoot).Top + lblKode(NoRoot).Height + BatasAtas
        
            Isi = Isi + 1
            
            'Memeriksa tingkat root
            If n > Akhir Then
                Tingkat = Tingkat * 2
                Akhir = Tingkat - 2
                Pengali = Pengali + 1
            End If
        
            'Posisi root baru
            If Isi = 1 Then
                lblKode(n).Left = lblKode(NoRoot).Left - (Lebar / Pengali)
                RootKiri(NoRoot) = n
            Else
                lblKode(n).Left = lblKode(NoRoot).Left + (Lebar / Pengali)
                RootKanan(NoRoot) = n
            End If
        
            'Mengatur garis
            Load lnKode(n)
            lnKode(n).X1 = lblKode(n).Left + (lblKode(n).Width / 2)
            lnKode(n).Y1 = lblKode(n).Top + (lblKode(n).Height / 2)
            lnKode(n).X2 = lblKode(NoRoot).Left + (lblKode(NoRoot).Width / 2)
            lnKode(n).Y2 = lblKode(NoRoot).Top + (lblKode(NoRoot).Height / 2)
            lnKode(n).Visible = True
        
            lblKode(n).Caption = txtKode.Text
            lblKode(n).Visible = True
            n = n + 1
        
            'Memeriksa keadaan root berapa jumlah yang terisi
            If Isi = 2 Then
                Isi = 0
                NoRoot = NoRoot + 1
            End If
        End If
    End If
    
    'Mengosongkan textbox
    txtKode.Text = ""
    txtKode.SetFocus
End Sub

Private Sub lblTree_Click(Index As Integer)
    If Taruh = False And (Kiri(Index) = 0 Or Kanan(Index) = 0) Then
        Load lblTree(n)
        
        'Membuat Variabel Root Kiri, Kanan, dan Atas
        Kiri(n) = 0
        Kanan(n) = 0
        Atas(n) = Index
        Level(n) = Level(Atas(n)) + 1
            
        'Menentukan posisi atas objek
        lblTree(n).Top = lblTree(Atas(n)).Top + lblTree(Atas(n)).Height + BatasAtas
        
        'Memeriksa tingkat root
        'If n > Akhir Then
        '    Tingkat = Tingkat * 2
        '    Akhir = Tingkat - 2
        '    Pengali = Pengali + 1
        'End If
        
        'Posisi root baru
        If Kiri(Atas(n)) = 0 Then
            lblTree(n).Left = lblTree(Atas(n)).Left - (Lebar / Level(n))
            Kiri(Atas(n)) = n
        ElseIf Kanan(Atas(n)) = 0 Then
            lblTree(n).Left = lblTree(Atas(n)).Left + (Lebar / Level(n))
            Kanan(Atas(n)) = n
        End If
        
        'Mengatur garis
        Load lnTree(n)
        lnTree(n).X1 = lblTree(n).Left + (lblTree(n).Width / 2)
        lnTree(n).Y1 = lblTree(n).Top + (lblTree(n).Height / 2)
        lnTree(n).X2 = lblTree(Atas(n)).Left + (lblTree(Atas(n)).Width / 2)
        lnTree(n).Y2 = lblTree(Atas(n)).Top + (lblTree(Atas(n)).Height / 2)
        lnTree(n).Visible = True
        
        lblTree(n).Caption = txtKode.Text
        lblTree(n).Visible = True

        'Proses membuat tree sebelah kanan/tree seimbang
        IsiKanan
        
        'Objek telah diletakkan
        Taruh = True
        
        'Mengubah pointer ke semula
        Me.MousePointer = 0
    End If
End Sub

Private Sub lblTree_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Taruh = True Then Tekan = True
End Sub

Private Sub lblTree_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Tekan = True Then
        'Me.Caption = CStr(X) & "," & CStr(Y)
        lblTree(Index).Left = X
        lblTree(Index).Top = Y

        'Menggambar Garis
        lnTree(Index).X1 = lblTree(Index).Left + (lblTree(Index).Width / 2)
        lnTree(Index).Y1 = lblTree(Index).Top + (lblTree(Index).Height / 2)
        
        'Mengatur garis Root
        If Kiri(Index) <> 0 Then
            lnTree(Kiri(Index)).X2 = lblTree(Index).Left + (lblTree(Index).Width / 2)
            lnTree(Kiri(Index)).Y2 = lblTree(Index).Top + (lblTree(Index).Height / 2)
        End If
        If Kanan(Index) <> 0 Then
            lnTree(Kanan(Index)).X2 = lblTree(Index).Left + (lblTree(Index).Width / 2)
            lnTree(Kanan(Index)).Y2 = lblTree(Index).Top + (lblTree(Index).Height / 2)
        End If
    End If
End Sub

Private Sub lblTree_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Tekan = False
End Sub

Private Sub pctTree_Click()
    If Parent = False And Taruh = False Then
        lblTree(0).Visible = True
        lblTree(0).Left = ((pctTree.Width - lblTree(0).Width) / 2)
            'lblKode(0).Left = pctBack.Left + (((pctBack.Width - pctBack.Left) - lblKode(0).Width) / 2)
        lblTree(0).Caption = txtKode.Text
        
        Parent = True
        Taruh = True
        
        Level(0) = 0
        
        Kiri(0) = 0
        Kanan(0) = 0
        Atas(1) = 0
    
        'Proses membuat tree sebelah kanan/tree seimbang
        IsiKanan
        
        'Mengubah pointer ke semula
        Me.MousePointer = 0
    End If
End Sub
