VERSION 5.00
Begin VB.Form FrmMacros 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Configuracion de macros"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   3450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton RR 
      Caption         =   "Reiniciar macros"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   3135
   End
   Begin VB.CommandButton Cerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   1455
   End
   Begin VB.OptionButton EquiparObjeto 
      Caption         =   "Equipar objeto seleccionado"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   3375
   End
   Begin VB.OptionButton UsarItem 
      Caption         =   "Usar objeto seleccionado"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   3375
   End
   Begin VB.OptionButton EnviarHechizo 
      Caption         =   "Hechizo seleccionado"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   3135
   End
   Begin VB.TextBox Comando 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   3135
   End
   Begin VB.OptionButton EnviarComando 
      Caption         =   "Comando"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label ElObjetoE 
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label ElObjeto 
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label ElHechizo 
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "MACRO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   1125
   End
   Begin VB.Label ElMacro 
      AutoSize        =   -1  'True
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   150
   End
End
Attribute VB_Name = "FrmMacros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Aceptar_Click()
    If EnviarComando.value = True And Comando.Text = "" Then
        MsgBox "Por favor ingrese un comando."
        Exit Sub
    End If
    
    'comando / mensaje
    If EnviarComando.value = True Then
        Call WriteVar(App.Path & "\init\Macros\" & UserName & ".dat", "MACROS" & ElMacro.Caption, "Tipo", "1")
        Call WriteVar(App.Path & "\init\Macros\" & UserName & ".dat", "MACROS" & ElMacro.Caption, "Texto", Comando.Text)
        Call WriteVar(App.Path & "\init\Macros\" & UserName & ".dat", "MACROS" & ElMacro.Caption, "Index", "1")
    'hechizo
    ElseIf EnviarHechizo.value = True Then
        If frmMain.hlst.List(frmMain.hlst.listIndex) = "(None)" Then
            MsgBox "Por favor selecciona el hechizo."
            Exit Sub
        End If
        
        Call WriteVar(App.Path & "\init\Macros\" & UserName & ".dat", "MACROS" & ElMacro.Caption, "Tipo", "2")
        Call WriteVar(App.Path & "\init\Macros\" & UserName & ".dat", "MACROS" & ElMacro.Caption, "Texto", frmMain.hlst.List(frmMain.hlst.listIndex))
        Call WriteVar(App.Path & "\init\Macros\" & UserName & ".dat", "MACROS" & ElMacro.Caption, "Index", frmMain.hlst.listIndex + 1)
    'usar
    ElseIf UsarItem.value = True Then
        If Inventario.SelectedItem > 0 And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
            Call WriteVar(App.Path & "\init\Macros\" & UserName & ".dat", "MACROS" & ElMacro.Caption, "Tipo", "3")
            Call WriteVar(App.Path & "\init\Macros\" & UserName & ".dat", "MACROS" & ElMacro.Caption, "Texto", Inventario.ItemName(Inventario.SelectedItem))
            Call WriteVar(App.Path & "\init\Macros\" & UserName & ".dat", "MACROS" & ElMacro.Caption, "Index", Inventario.SelectedItem)
        End If
    'Equipar
    ElseIf EquiparObjeto.value = True Then
        If Inventario.SelectedItem > 0 And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
            Call WriteVar(App.Path & "\init\Macros\" & UserName & ".dat", "MACROS" & ElMacro.Caption, "Tipo", "4")
            Call WriteVar(App.Path & "\init\Macros\" & UserName & ".dat", "MACROS" & ElMacro.Caption, "Texto", Inventario.ItemName(Inventario.SelectedItem))
            Call WriteVar(App.Path & "\init\Macros\" & UserName & ".dat", "MACROS" & ElMacro.Caption, "Index", Inventario.SelectedItem)
        End If
    End If
    
    Call CargarMacros
    Unload Me
End Sub

Private Sub Cerrar_Click()
    Unload Me
End Sub

Private Sub RR_Click()
    Call NuevoMacros
    Call CargarMacros
    Unload Me
End Sub
