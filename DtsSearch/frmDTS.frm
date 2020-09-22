VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDts 
   Caption         =   "DTS"
   ClientHeight    =   6030
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   10590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   9240
      TabIndex        =   8
      Top             =   5280
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar sbrStatus 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   9
      Top             =   5670
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5768
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkCase 
      Caption         =   "Coincidir maiúsculas e minúsculas"
      Height          =   315
      Left            =   6780
      TabIndex        =   4
      Top             =   390
      Width           =   2805
   End
   Begin MSComctlLib.ListView lvwResult 
      Height          =   3855
      Left            =   150
      TabIndex        =   6
      Top             =   1260
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Package"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Task"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Property"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Text"
         Object.Width           =   14111
      EndProperty
   End
   Begin VB.ComboBox cboServers 
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Top             =   390
      Width           =   3135
   End
   Begin VB.TextBox txtTexto 
      Height          =   315
      Left            =   3420
      TabIndex        =   3
      Top             =   390
      Width           =   3135
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Procurar"
      Default         =   -1  'True
      Height          =   375
      Left            =   7920
      MousePointer    =   1  'Arrow
      TabIndex        =   7
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label lblResult 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resultado da pesquisa:"
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   960
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Procurar este texto:"
      Height          =   195
      Left            =   3420
      TabIndex        =   2
      Top             =   180
      Width           =   1380
   End
   Begin VB.Label lblServers 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SQL Server:"
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   870
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuSair 
         Caption         =   "Sai&r"
      End
   End
   Begin VB.Menu mnuEditar 
      Caption         =   "&Editar"
      Begin VB.Menu mnuEditarSeltudo 
         Caption         =   "Selecionar &tudo"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditarCopiar 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "frmDts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'ver lookups
'ver transformatios scripts

Private blnProcessar As Boolean

Private Sub cboServers_DropDown()
    If cboServers.ListCount = 0 Then
        Screen.MousePointer = vbHourglass
        CarregaCboServers
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    Dim objDTSAppl      As DTS.Application
    Dim colPkgInfo      As DTS.PackageInfos
    Dim objPkgInfo      As DTS.PackageInfo
    Dim ColPacks        As DTS.PackageSQLServer
    Dim strTexto       As String
    Dim strPack         As String

    If blnProcessar Then        'Indica que o processamento está acontecendo
        blnProcessar = False
        GoTo Fim
    End If

    cboServers.Enabled = False
    txtTexto.Enabled = False
    chkCase.Enabled = False
    mnuArquivo.Enabled = False
    mnuEditar.Enabled = False
    
    lvwResult.ListItems.Clear
    strTexto = txtTexto
    Set objDTSAppl = New DTS.Application
    
    Set ColPacks = objDTSAppl.GetPackageSQLServer(cboServers.Text, "", "", DTSSQLStgFlag_UseTrustedConnection)
    sbrStatus.Panels(1).Text = "Servidor " & cboServers.Text
    Set colPkgInfo = ColPacks.EnumPackageInfos("", True, "")
    
    Screen.MousePointer = vbHourglass
    Set objPkgInfo = colPkgInfo.Next
    blnProcessar = True
    cmdSearch.Caption = "P&arar"
    cmdSearch.Default = False
    cmdFechar.Enabled = False
    
    Do Until colPkgInfo.EOF
        DoEvents
        If Not blnProcessar Then
            Exit Do
        End If
        sbrStatus.Panels(2).Text = "Procurando " & objPkgInfo.Name
        SearchTask objPkgInfo.PackageID, txtTexto
        Set objPkgInfo = colPkgInfo.Next
    Loop
    cmdFechar.Enabled = True
    
    cmdSearch.Caption = "&Procurar"
    cmdSearch.Default = True
    
    If blnProcessar Then
        sbrStatus.Panels(2).Text = "Pronto"
    Else
        sbrStatus.Panels(2).Text = "Cancelado"
    End If
    sbrStatus.Panels(3).Text = ""
    blnProcessar = False
    Screen.MousePointer = vbDefault
    cboServers.Enabled = True
    txtTexto.Enabled = True
    chkCase.Enabled = True
    mnuArquivo.Enabled = True
    mnuEditar.Enabled = True
    
Fim:
    Set objDTSAppl = Nothing
    Set colPkgInfo = Nothing
    Set objPkgInfo = Nothing
    Set ColPacks = Nothing

End Sub

Private Sub SearchTask(ByVal strPackageId As String, strTextSearch As String)
'    Dim objPackage      As DTS.Package
    Dim objPackage2     As DTS.Package2
    
    Dim colTasks        As DTS.Tasks
    Dim objTask         As DTS.Task
    
    Dim objDataPumpTask2    As DTS.DataPumpTask2
    Dim objExecuteSQLTask2  As DTS.ExecuteSQLTask2
    
    Dim objTran         As DTS.Transformation2
    Dim objProp         As DTS.Property
    Dim objProp2        As DTS.Property
    Dim objLoo          As DTS.Lookup
    Dim blnCase         As Boolean
    
'    Set objPackage = New DTS.Package
'    objPackage.LoadFromSQLServer cboServers.Text, , , DTSSQLStgFlag_UseTrustedConnection, , strPackageId
    
    Set objPackage2 = New DTS.Package2
    objPackage2.LoadFromSQLServer cboServers.Text, , , DTSSQLStgFlag_UseTrustedConnection, , strPackageId
   
    Set colTasks = objPackage2.Tasks
    
    blnCase = (chkCase.Value = vbChecked)
    
    For Each objTask In colTasks
        DoEvents
        If Not blnProcessar Then
            Exit For
        End If
        
        For Each objProp In objTask.Properties
            If Not blnProcessar Then
                Exit For
            End If
            sbrStatus.Panels(3).Text = "Task: " & objTask.Name
            sbrStatus.Refresh

            
            Select Case objProp.Name
            Case "Name", "Description", "DestinationObjectName", "DestinationSQLStatement", "InputGlobalVariableNames", "SourceObjectName", "SourceSQLStatement", "SQLStatement"
                Busca objProp.Value, strTextSearch, blnCase, objPackage2.Name, objPackage2.Description, objTask.Name, objTask.Description, objProp.Name, objProp.Value
            End Select
            'Procura nas propriedades extendidas
            Select Case TypeName(objTask.CustomTask)
                Case "DataPumpTask2"
                    Set objDataPumpTask2 = objTask.CustomTask
                    For Each objTran In objDataPumpTask2.Transformations
                        'Busca nas propriedades da transformação
'                        For Each objProp2 In objTran.Properties
'                            Select Case objProp2.Name
'                            Case "DestinationColumns", "SourceColumns", "Name"
'                                Busca objProp2.Value, strTextSearch, blnCase, _
'                                        objPackage2.Name, objPackage2.Description, _
'                                        objDataPumpTask2.Name, objDataPumpTask2.Name, _
'                                        objProp2.Name, objProp2.Value
'                            End Select
'                        Next
                        'Busca em scripts da transformação
                        If objTran.TransformServerID = "DTSPump.DataPumpTransformScript" Then
                            Busca objTran.TransformServer.Text, strTextSearch, blnCase, objPackage2.Name, objPackage2.Description, objDataPumpTask2.Name, objDataPumpTask2.Description, "Transformation " & objTran.Name & " (Script Text)", objTran.TransformServer.Text
                        End If
                    Next
                    'Busca em lookups da task
                    For Each objLoo In objDataPumpTask2.Lookups
                        Busca objLoo.Query, strTextSearch, blnCase, objPackage2.Name, objPackage2.Description, objDataPumpTask2.Name, objDataPumpTask2.Description, "Lookup " & objLoo.Name & " (Query Text)", objLoo.Query
                    Next
                Case "ExecuteSQLTask2"
                    Set objExecuteSQLTask2 = objTask.CustomTask
                Case "ExecutePackageTask"
                
                Case "CreateProcessTask2"
                
                Case "DataDrivenQueryTask2"
    
                Case Else
                    Debug.Print "Task = " & TypeName(objTask.CustomTask)
            End Select
        Next
       
    Next

End Sub

Private Sub Busca(strTexto As String, strTextoProcurado As String, blnCase As Boolean, _
    strPackName As String, strPackDesc As String, strTaskName As String, strTaskDesc As String, _
    strPropName As String, strPropValue As String)
    
    Dim objItem         As ListItem
    
    If InStr(1, strTexto, strTextoProcurado, IIf(blnCase, vbBinaryCompare, vbTextCompare)) > 0 Then
        Set objItem = lvwResult.ListItems.Add(, , strPackName & IIf(Len(strPackDesc) > 0, " (" & strPackDesc & ")", vbNullString))
        objItem.ListSubItems.Add , , strTaskDesc & " (" & strTaskName & ")"
        objItem.ListSubItems.Add , , strPropName
        objItem.ListSubItems.Add , , strPropValue
        lvwResult.SelectedItem.EnsureVisible
    End If
End Sub

Private Sub CarregaCboServers()
    Dim objApl      As SQLDMO.Application
    Dim objServer   As SQLDMO.SQLServer
    Dim i As Integer
    
    On Error GoTo Erros
    
    Set objApl = New SQLDMO.Application

    cboServers.Clear
    For i = i To 1000
        cboServers.AddItem objApl.ListAvailableSQLServers(i)
    Next


Erros:
    '-2147199735 '    [SQL-DMO]The passed ordinal is out of range of the specified collection.
   
    Exit Sub
End Sub

Private Sub Form_Load()
    cboServers.Text = "Batel"
    Show
    txtTexto.SetFocus
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        cmdFechar.Move (ScaleWidth - cmdFechar.Width - 60), (ScaleHeight - cmdFechar.Height - sbrStatus.Height - 60)
        cmdSearch.Move (cmdFechar.Left - cmdSearch.Width - 60), cmdFechar.Top
        lvwResult.Move 120, lvwResult.Top, (Me.ScaleWidth - 180), (cmdFechar.Top - lvwResult.Top - 120)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If blnProcessar Then
        Cancel = True
    End If
End Sub

Private Sub lvwResult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Not blnProcessar Then
        lvwResult.SortKey = ColumnHeader.Index - 1
    End If
End Sub

Private Sub lvwResult_DblClick()
    Dim objItem As ListItem
    
    If Not blnProcessar Then
        Set objItem = lvwResult.SelectedItem
        frmDetalhe.ShowModal objItem.Text, objItem.SubItems(1), objItem.SubItems(2), objItem.SubItems(3), txtTexto
    End If
    
End Sub

Private Sub lvwResult_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not blnProcessar Then
        If Button = vbRightButton Then
            PopupMenu mnuEditar
        End If
    End If
End Sub

Private Sub mnuEditarCopiar_Click()
    Dim objItem As ListItem
    Dim intItem As Integer
    Dim strClip As String
    Dim strLinha As String
    

    For Each objItem In lvwResult.ListItems
        If objItem.Selected Then
            strLinha = "Package: " & vbTab & objItem.Text & vbCrLf & _
                       "Task:    " & vbTab & objItem.SubItems(1) & vbCrLf & _
                       "Property:" & vbTab & objItem.SubItems(2) & vbCrLf & _
                       "Text:    " & vbTab & objItem.SubItems(3) & vbCrLf
            strClip = strClip & strLinha
        End If
    Next
    
    Clipboard.Clear
    Clipboard.SetText (strClip)
    
End Sub

Private Sub mnuEditarSeltudo_Click()
    Dim intItem As Integer
    
    For intItem = 1 To lvwResult.ListItems.Count
        Set lvwResult.SelectedItem = lvwResult.ListItems(intItem)
    Next

End Sub

Private Sub mnuSair_Click()
    If Not blnProcessar Then
        Unload Me
    End If
End Sub
