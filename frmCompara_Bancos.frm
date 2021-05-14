VERSION 5.00
Begin VB.Form frmComparaBancos 
   Caption         =   "Compara Bancos"
   ClientHeight    =   11280
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15600
   Icon            =   "frmCompara_Bancos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11280
   ScaleWidth      =   15600
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrResetLink 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   14880
      Top             =   10800
   End
   Begin VB.TextBox txtServidorBSenha 
      Height          =   405
      Left            =   9600
      TabIndex        =   7
      Top             =   1080
      Width           =   4935
   End
   Begin VB.TextBox txtServidorBUsuario 
      Height          =   405
      Left            =   9600
      TabIndex        =   6
      Text            =   "sa"
      Top             =   600
      Width           =   4935
   End
   Begin VB.TextBox txtServidorASenha 
      Height          =   405
      Left            =   1080
      TabIndex        =   2
      Top             =   1080
      Width           =   4935
   End
   Begin VB.TextBox txtServidorAUsuario 
      Height          =   405
      Left            =   1080
      TabIndex        =   1
      Text            =   "sa"
      Top             =   600
      Width           =   4935
   End
   Begin VB.TextBox txtServidorB 
      Height          =   405
      Left            =   9600
      TabIndex        =   5
      Text            =   "IP\Nome_Instância,Porta"
      Top             =   120
      Width           =   4935
   End
   Begin VB.TextBox txtServidorA 
      Height          =   405
      Left            =   1080
      TabIndex        =   0
      Text            =   "IP\Nome_Instância,Porta"
      Top             =   120
      Width           =   4935
   End
   Begin VB.TextBox txtResultado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   4920
      Width           =   15255
   End
   Begin VB.ListBox lstBancosB 
      Height          =   2985
      Left            =   8640
      TabIndex        =   9
      Top             =   1800
      Width           =   6735
   End
   Begin VB.CommandButton btnListaBdB 
      Caption         =   "Lista BDs"
      Height          =   405
      Left            =   14640
      TabIndex        =   8
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton btnComparar 
      Caption         =   "Comparar Bancos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6960
      TabIndex        =   10
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton btnListaBdA 
      Caption         =   "Lista BDs"
      Height          =   405
      Left            =   6120
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.ListBox lstBancosA 
      Height          =   2985
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   6735
   End
   Begin VB.Label lblMouseOverColor 
      Caption         =   "lblMouseOverColor"
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   9720
      TabIndex        =   22
      Top             =   10920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblMouseOffColor 
      Caption         =   "lblMouseOffColor"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   12240
      TabIndex        =   21
      Top             =   10920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblLink 
      Caption         =   "https://github.com/silvairsoares/Compara_Bancos_SQL_Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   10920
      Width           =   6855
   End
   Begin VB.Label Label8 
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      TabIndex        =   19
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Usuário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      TabIndex        =   18
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Usuário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Banco B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      TabIndex        =   15
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Banco A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Servidor B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      TabIndex        =   13
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Servidor A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmComparaBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMAXIMIZED = 3

' Used to find the mouse position.
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Dim Conexao As New Connection
Dim ConexaoA As New Connection
Dim ConexaoB As New Connection

Private Sub btnComparar_Click()
    
    If lstBancosA.Text = "" Or lstBancosB.Text = "" Then
        MsgBox "Selecione 2 bancos de dados, para executar a comparação.", vbCritical + vbOKOnly
        Exit Sub
    End If
    
    txtResultado = ""
    DoEvents
    
    If Not ConectaBD("A", txtServidorA, lstBancosA.Text) Then Exit Sub
    If Not ConectaBD("B", txtServidorB, lstBancosB.Text) Then Exit Sub

    MsgBox "Este processo pode demorar um pouco!", vbExclamation

    Screen.MousePointer = 11
    ExecutaComparacao
    Screen.MousePointer = 0
    DoEvents
    
    If txtResultado = "" Then
        txtResultado = "Os bancos de dados comparados têm estruturas iguais."
        MsgBox "Verificação concluída com sucesso", vbInformation, "Ok"
    Else
        MsgBox "Verificação concluída com falhas", vbCritical, "Atenção"
    End If

End Sub

Private Sub ExecutaComparacao()

    Dim rsEstruturaA As New Recordset
    Dim rsEstruturaB As New Recordset

    rsEstruturaA.Open "SELECT " _
        & "TABLE_NAME Tabela " _
        & "FROM INFORMATION_SCHEMA.Columns " _
        & "WHERE TABLE_CATALOG = '" & lstBancosA.Text & "' GROUP BY TABLE_NAME", ConexaoA, adOpenStatic, adLockReadOnly

    rsEstruturaB.Open "SELECT " _
        & "TABLE_NAME Tabela " _
        & "FROM INFORMATION_SCHEMA.Columns " _
        & "WHERE TABLE_CATALOG = '" & lstBancosB.Text & "' GROUP BY TABLE_NAME", ConexaoB, adOpenStatic, adLockReadOnly

    Do While Not rsEstruturaA.EOF

        'Verifica a existência da tabela
        rsEstruturaB.Requery
        rsEstruturaB.Find " Tabela = '" & rsEstruturaA("Tabela") & "' "
        If rsEstruturaB.EOF Then
            Log "- Erro - Não encontrada a tabela """ & rsEstruturaA("Tabela") & """ no banco """ & lstBancosB & """"
        Else
            VerificaDetalhesDaTabela rsEstruturaA("Tabela")
        End If

        rsEstruturaA.MoveNext
    Loop


    RecordsetFinaliza rsEstruturaA
    RecordsetFinaliza rsEstruturaB
    
End Sub

Private Sub VerificaDetalhesDaTabela(Tabela As String)
    
    Dim rsEstruturaA As New Recordset
    Dim rsEstruturaB As New Recordset
    Dim strErro As String

    rsEstruturaA.Open "SELECT " _
        & "COLUMN_NAME Coluna, IS_NULLABLE PermiteNull, DATA_TYPE Tipo, CASE WHEN CHARACTER_MAXIMUM_LENGTH IS NULL THEN '' ELSE CHARACTER_MAXIMUM_LENGTH END Tamanho " _
        & "FROM INFORMATION_SCHEMA.Columns " _
        & "WHERE TABLE_CATALOG = '" & lstBancosA.Text & "' AND " _
        & "TABLE_NAME = '" & Tabela & "'", ConexaoA, adOpenStatic, adLockReadOnly

    rsEstruturaB.Open "SELECT " _
        & "COLUMN_NAME Coluna, IS_NULLABLE PermiteNull, DATA_TYPE Tipo, CASE WHEN CHARACTER_MAXIMUM_LENGTH IS NULL THEN '' ELSE CHARACTER_MAXIMUM_LENGTH END Tamanho " _
        & "FROM INFORMATION_SCHEMA.Columns " _
        & "WHERE TABLE_CATALOG = '" & lstBancosB.Text & "' AND " _
        & "TABLE_NAME = '" & Tabela & "'", ConexaoB, adOpenStatic, adLockReadOnly

    Do While Not rsEstruturaA.EOF

        'Verifica a existência da tabela
        rsEstruturaB.Requery
        rsEstruturaB.Find " Coluna = '" & rsEstruturaA("Coluna") & "' "
        If rsEstruturaB.EOF Then
            Log "- Erro - Não encontrada a coluna """ & rsEstruturaA("Coluna") & """ na tabela """ & Tabela & """ do banco """ & lstBancosB & """"
        Else
            If rsEstruturaA("PermiteNull") <> rsEstruturaB("PermiteNull") Then
                strErro = "- Erro - Configuração de permissão para valor nulo" & vbNewLine
                strErro = strErro & vbTab & "Tabela " & vbTab & "" & Tabela & "" & vbNewLine
                strErro = strErro & vbTab & "Coluna " & vbTab & "" & rsEstruturaB("Coluna") & "" & vbNewLine
                strErro = strErro & vbTab & "Banco A " & vbTab & "" & rsEstruturaA("PermiteNull") & "" & vbNewLine
                strErro = strErro & vbTab & "Banco B " & vbTab & "" & rsEstruturaB("PermiteNull") & "" & vbNewLine
                
                Log strErro
            End If
            If rsEstruturaA("Tipo") <> rsEstruturaB("Tipo") Then
                strErro = "- Erro - Tipo de dados" & vbNewLine
                strErro = strErro & vbTab & "Tabela " & vbTab & "" & Tabela & "" & vbNewLine
                strErro = strErro & vbTab & "Coluna " & vbTab & "" & rsEstruturaB("Coluna") & "" & vbNewLine
                strErro = strErro & vbTab & "Banco A " & vbTab & "" & rsEstruturaA("Tipo") & "" & vbNewLine
                strErro = strErro & vbTab & "Banco B " & vbTab & "" & rsEstruturaB("Tipo") & "" & vbNewLine
                
                Log strErro
            End If
            If rsEstruturaA("Tamanho") <> rsEstruturaB("Tamanho") Then
                strErro = "- Erro - Tamanho da coluna" & vbNewLine
                strErro = strErro & vbTab & "Tabela " & vbTab & "" & Tabela & "" & vbNewLine
                strErro = strErro & vbTab & "Coluna " & vbTab & "" & rsEstruturaB("Coluna") & "" & vbNewLine
                strErro = strErro & vbTab & "Banco A " & vbTab & "" & rsEstruturaA("Tamanho") & "" & vbNewLine
                strErro = strErro & vbTab & "Banco B " & vbTab & "" & rsEstruturaB("Tamanho") & "" & vbNewLine
                
                Log strErro
            End If
        End If

        rsEstruturaA.MoveNext
    Loop

    RecordsetFinaliza rsEstruturaA
    RecordsetFinaliza rsEstruturaB

End Sub

Private Sub Log(Mensagem As String)

    txtResultado = txtResultado & Mensagem & vbNewLine & String(243, "-") & vbNewLine
    
End Sub

Private Sub btnListaBdA_Click()

    Dim rsBancos As New Recordset
    If rsBancos.State = adStateOpen Then Set rsBancos = Nothing

    lstBancosA.Clear

    If ConectaBD("A", txtServidorA, "") Then
        rsBancos.Open "SELECT name FROM sys.databases WHERE (database_id > 6) ORDER BY name", ConexaoA, adOpenStatic, adLockReadOnly

        Do While Not (rsBancos.EOF)
            lstBancosA.AddItem rsBancos("name")
            rsBancos.MoveNext
        Loop

    End If

End Sub

Private Sub btnListaBdB_Click()

    Dim rsBancos As New Recordset
    If rsBancos.State = adStateOpen Then Set rsBancos = Nothing

    lstBancosB.Clear

    If ConectaBD("B", txtServidorB, "") Then
        rsBancos.Open "SELECT name FROM sys.databases WHERE (database_id > 6) ORDER BY name", ConexaoB, adOpenStatic, adLockReadOnly

        Do While Not (rsBancos.EOF)
            lstBancosB.AddItem rsBancos("name")
            rsBancos.MoveNext
        Loop

    End If

End Sub

Private Sub Form_Load()

    lstBancosA.Clear
    lstBancosB.Clear

End Sub

Private Function ConectaBD(Lado As String, Servidor As String, Banco As String) As Boolean
On Error GoTo TrataErro

    If Lado = "A" Then

        If ConexaoA.State = adStateOpen Then
            ConexaoA.Close
        End If

        If Banco <> "" Then
            ConexaoA.ConnectionString = "Provider=SQLOLEDB;Data Source=" & Servidor & ";Initial Catalog=" & Banco & "; User ID=" + txtServidorAUsuario + ";Password=" + txtServidorASenha + ""
        Else
            ConexaoA.ConnectionString = "Provider=SQLOLEDB;Data Source=" & Servidor & "; User ID=" + txtServidorAUsuario + ";Password=" + txtServidorASenha + ""
        End If

        ConexaoA.Open

    ElseIf Lado = "B" Then
        If ConexaoB.State = adStateOpen Then
            ConexaoB.Close
        End If

        If Banco <> "" Then
            ConexaoB.ConnectionString = "Provider=SQLOLEDB;Data Source=" & Servidor & ";Initial Catalog=" & Banco & "; User ID=" + txtServidorBUsuario + ";Password=" + txtServidorBSenha + ""
        Else
            ConexaoB.ConnectionString = "Provider=SQLOLEDB;Data Source=" & Servidor & "; User ID=" + txtServidorBUsuario + ";Password=" + txtServidorBSenha + ""
        End If

        ConexaoB.Open
    Else
        If Conexao.State = adStateOpen Then
            Conexao.Close
        End If

        Conexao.ConnectionString = "Provider=SQLOLEDB;Data Source=" & Servidor & "; User ID=sa;Password=keycod&"

        Conexao.Open
    End If
    
    ConectaBD = True
    
    Exit Function

TrataErro:
    MsgBox "Erro ao conectar no banco: " + Servidor + "/" + Banco + vbNewLine + Err.Description, vbCritical
End Function

Public Function RecordsetFinaliza(ByRef rsRecordset As Recordset)
On Error GoTo TrataErro
    
    If rsRecordset.State = adStateOpen Then
        rsRecordset.Close
    End If
    Set rsRecordset = Nothing
    Exit Function

TrataErro:
End Function

' Open the Web page.
Private Sub lblLink_Click()
    ShellExecute ByVal 0&, "open", _
        "https://github.com/silvairsoares/Compara_Bancos_SQL_Server/", _
        vbNullString, vbNullString, _
        SW_SHOWMAXIMIZED
End Sub


Private Sub lblLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' See if the link has the mouse over color.
    If Not lblLink.ForeColor = lblMouseOverColor.ForeColor Then
        ' Display the color.
        lblLink.ForeColor = lblMouseOverColor.ForeColor
        lblLink.Font.Underline = True

        ' Enable the timer.
        tmrResetLink.Enabled = True
    End If
End Sub


' See if the mouse is still over the label.
Private Sub tmrResetLink_Timer()
Dim pt As POINTAPI

    ' Get the cursor's position in screen coordinates.
    GetCursorPos pt

    ' Convert to form coordinates.
    ' Note that this converts the position in pixels
    ' and the form's ScaleMode is vbPixels.
    ScreenToClient hwnd, pt

    ' See if the point is under the text label.
    If pt.X < lblLink.Left Or pt.X > lblLink.Left + _
        lblLink.Width Or _
       pt.Y < lblLink.Top Or pt.Y > lblLink.Top + _
           lblLink.Height _
    Then
        ' The mouse is not over the label.
        ' Restore the original font.
        lblLink.ForeColor = lblMouseOffColor.ForeColor
        lblLink.Font.Underline = False

        ' Disable the timer.
        tmrResetLink.Enabled = False
    End If
End Sub


