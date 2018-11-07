Option Explicit On 
Imports System.Windows.Forms.SendKeys
Imports prjControl
Imports prjPrinterNet
Imports prjEmpresas

Public Class frmTelefonos
    Inherits System.Windows.Forms.Form
    Private mnEmpresa As Int32      ' empresa de gestion.
    Dim WithEvents moTelefono As clsTelefono
    Dim moBusTelefonos As clsBusTelefonos
    Public Enum EstadoVentana       ' estados posibles de la ventana
        Consulta = 1
        Mantenimiento = 2
        NuevoRegistro = 3
        Lineas = 4
        Salida = 5
    End Enum
    Dim mtEstado As EstadoVentana
    Dim mbPrimeraVez As Boolean
    ' variables de impresion ************************
    Private WithEvents moSelImpresora As prjControl.frmSelImpresora
    Dim moPrinter As New clsPrinter       ' objeto para imprimir
    Dim moImpresora As New prjPrinterNet.clsImpresora   ' Objeto Impresora
    Dim mnListado As Integer
    Dim msRaya As String
    Friend WithEvents lblExtension As System.Windows.Forms.Label
    Dim mnUltimoCodigo As Integer
    Friend WithEvents lblTipo As System.Windows.Forms.Label
    Dim msTipo As String = "M"

#Region " Código generado por el Diseñador de Windows Forms "

    Public Sub New()
        MyBase.New()

        'El Diseñador de Windows Forms requiere esta llamada.
        InitializeComponent()

        'Agregar cualquier inicialización después de la llamada a InitializeComponent()

    End Sub

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms requiere el siguiente procedimiento
    'Puede modificarse utilizando el Diseñador de Windows Forms. 
    'No lo modifique con el editor de código.
    Friend WithEvents lblInfo As System.Windows.Forms.Label
    Friend WithEvents lblPrograma As System.Windows.Forms.Label
    Friend WithEvents lblDescripcion As System.Windows.Forms.Label
    Friend WithEvents lblTitulo As System.Windows.Forms.Label
    Friend WithEvents lblTeclas As System.Windows.Forms.Label
    Friend WithEvents panCampos As System.Windows.Forms.Panel
    Friend WithEvents grdLineas As FlexCell.Grid
    Friend WithEvents tabMenu As prjToolBar.ctlToolBar
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtBusqueda As control.txtVisanfer
    Friend WithEvents lblNombre As System.Windows.Forms.Label
    Friend WithEvents lblTelefono As System.Windows.Forms.Label
    Friend WithEvents lblPrivado As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTelefonos))
        Me.lblInfo = New System.Windows.Forms.Label()
        Me.lblPrograma = New System.Windows.Forms.Label()
        Me.lblDescripcion = New System.Windows.Forms.Label()
        Me.lblTitulo = New System.Windows.Forms.Label()
        Me.lblTeclas = New System.Windows.Forms.Label()
        Me.panCampos = New System.Windows.Forms.Panel()
        Me.lblExtension = New System.Windows.Forms.Label()
        Me.lblPrivado = New System.Windows.Forms.Label()
        Me.lblTelefono = New System.Windows.Forms.Label()
        Me.lblNombre = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.txtBusqueda = New control.txtVisanfer()
        Me.grdLineas = New FlexCell.Grid()
        Me.tabMenu = New prjToolBar.ctlToolBar()
        Me.lblTipo = New System.Windows.Forms.Label()
        Me.panCampos.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblInfo
        '
        Me.lblInfo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblInfo.Location = New System.Drawing.Point(9, 688)
        Me.lblInfo.Name = "lblInfo"
        Me.lblInfo.Size = New System.Drawing.Size(1007, 16)
        Me.lblInfo.TabIndex = 64
        '
        'lblPrograma
        '
        Me.lblPrograma.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.lblPrograma.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPrograma.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPrograma.Location = New System.Drawing.Point(9, 7)
        Me.lblPrograma.Name = "lblPrograma"
        Me.lblPrograma.Size = New System.Drawing.Size(503, 32)
        Me.lblPrograma.TabIndex = 67
        Me.lblPrograma.Text = "GESTION"
        Me.lblPrograma.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblDescripcion
        '
        Me.lblDescripcion.BackColor = System.Drawing.SystemColors.ControlLight
        Me.lblDescripcion.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDescripcion.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDescripcion.Location = New System.Drawing.Point(9, 39)
        Me.lblDescripcion.Name = "lblDescripcion"
        Me.lblDescripcion.Size = New System.Drawing.Size(503, 24)
        Me.lblDescripcion.TabIndex = 68
        Me.lblDescripcion.Text = "Descripcion del proceso."
        Me.lblDescripcion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblTitulo
        '
        Me.lblTitulo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTitulo.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitulo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblTitulo.Location = New System.Drawing.Point(512, 7)
        Me.lblTitulo.Name = "lblTitulo"
        Me.lblTitulo.Size = New System.Drawing.Size(504, 56)
        Me.lblTitulo.TabIndex = 66
        Me.lblTitulo.Text = "VISANFER, S.A. - 2003"
        Me.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblTeclas
        '
        Me.lblTeclas.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.lblTeclas.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTeclas.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTeclas.Location = New System.Drawing.Point(9, 703)
        Me.lblTeclas.Name = "lblTeclas"
        Me.lblTeclas.Size = New System.Drawing.Size(1007, 33)
        Me.lblTeclas.TabIndex = 63
        Me.lblTeclas.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'panCampos
        '
        Me.panCampos.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panCampos.Controls.Add(Me.lblTipo)
        Me.panCampos.Controls.Add(Me.lblExtension)
        Me.panCampos.Controls.Add(Me.lblPrivado)
        Me.panCampos.Controls.Add(Me.lblTelefono)
        Me.panCampos.Controls.Add(Me.lblNombre)
        Me.panCampos.Controls.Add(Me.Label15)
        Me.panCampos.Controls.Add(Me.Label14)
        Me.panCampos.Controls.Add(Me.txtBusqueda)
        Me.panCampos.Controls.Add(Me.grdLineas)
        Me.panCampos.Location = New System.Drawing.Point(9, 87)
        Me.panCampos.Name = "panCampos"
        Me.panCampos.Size = New System.Drawing.Size(1007, 601)
        Me.panCampos.TabIndex = 65
        '
        'lblExtension
        '
        Me.lblExtension.BackColor = System.Drawing.SystemColors.Info
        Me.lblExtension.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblExtension.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblExtension.Location = New System.Drawing.Point(889, 544)
        Me.lblExtension.Name = "lblExtension"
        Me.lblExtension.Size = New System.Drawing.Size(103, 32)
        Me.lblExtension.TabIndex = 240
        Me.lblExtension.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblPrivado
        '
        Me.lblPrivado.BackColor = System.Drawing.Color.Yellow
        Me.lblPrivado.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPrivado.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPrivado.ForeColor = System.Drawing.Color.Red
        Me.lblPrivado.Location = New System.Drawing.Point(16, 576)
        Me.lblPrivado.Name = "lblPrivado"
        Me.lblPrivado.Size = New System.Drawing.Size(256, 16)
        Me.lblPrivado.TabIndex = 239
        Me.lblPrivado.Text = "NUMERO PRIVADO"
        Me.lblPrivado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblPrivado.Visible = False
        '
        'lblTelefono
        '
        Me.lblTelefono.BackColor = System.Drawing.SystemColors.Info
        Me.lblTelefono.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTelefono.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTelefono.Location = New System.Drawing.Point(16, 544)
        Me.lblTelefono.Name = "lblTelefono"
        Me.lblTelefono.Size = New System.Drawing.Size(256, 32)
        Me.lblTelefono.TabIndex = 238
        Me.lblTelefono.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblNombre
        '
        Me.lblNombre.BackColor = System.Drawing.SystemColors.Info
        Me.lblNombre.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblNombre.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNombre.Location = New System.Drawing.Point(272, 544)
        Me.lblNombre.Name = "lblNombre"
        Me.lblNombre.Size = New System.Drawing.Size(618, 32)
        Me.lblNombre.TabIndex = 237
        Me.lblNombre.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.SystemColors.Info
        Me.Label15.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(808, 25)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(184, 20)
        Me.Label15.TabIndex = 236
        Me.Label15.Text = "F3 - NUEVA BUSQUEDA"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.SystemColors.Info
        Me.Label14.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(16, 25)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(184, 20)
        Me.Label14.TabIndex = 235
        Me.Label14.Text = "CTRL + F9 - BUSQUEDA"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtBusqueda
        '
        Me.txtBusqueda.AutoSelec = False
        Me.txtBusqueda.BackColor = System.Drawing.Color.White
        Me.txtBusqueda.Blink = False
        Me.txtBusqueda.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtBusqueda.DesdeCodigo = CType(0, Long)
        Me.txtBusqueda.DesdeFecha = New Date(CType(0, Long))
        Me.txtBusqueda.ForeColor = System.Drawing.Color.Black
        Me.txtBusqueda.HastaCodigo = CType(0, Long)
        Me.txtBusqueda.HastaFecha = New Date(CType(0, Long))
        Me.txtBusqueda.Location = New System.Drawing.Point(200, 25)
        Me.txtBusqueda.MaxLength = 40
        Me.txtBusqueda.Name = "txtBusqueda"
        Me.txtBusqueda.Size = New System.Drawing.Size(608, 20)
        Me.txtBusqueda.TabIndex = 234
        Me.txtBusqueda.TabStop = False
        Me.txtBusqueda.tipo = control.txtVisanfer.TiposCaja.Alfanumerico
        Me.txtBusqueda.ValorMax = 999999999.0R
        '
        'grdLineas
        '
        Me.grdLineas.BackColorBkg = System.Drawing.SystemColors.ControlLightLight
        Me.grdLineas.BorderStyle = FlexCell.Grid.BorderStyleEnum.FixedSingle
        Me.grdLineas.CheckedImage = CType(resources.GetObject("grdLineas.CheckedImage"), System.Drawing.Bitmap)
        Me.grdLineas.DisplayRowNumber = True
        Me.grdLineas.ExtendLastCol = True
        Me.grdLineas.Location = New System.Drawing.Point(16, 48)
        Me.grdLineas.MultiSelect = False
        Me.grdLineas.Name = "grdLineas"
        Me.grdLineas.Rows = 2
        Me.grdLineas.ScrollBars = FlexCell.Grid.ScrollBarsEnum.Vertical
        Me.grdLineas.SelectionBorderColor = System.Drawing.Color.Gray
        Me.grdLineas.SelectionMode = FlexCell.Grid.SelectionModeEnum.ByRow
        Me.grdLineas.Size = New System.Drawing.Size(976, 494)
        Me.grdLineas.TabIndex = 7
        Me.grdLineas.UncheckedImage = CType(resources.GetObject("grdLineas.UncheckedImage"), System.Drawing.Bitmap)
        '
        'tabMenu
        '
        Me.tabMenu.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tabMenu.Location = New System.Drawing.Point(9, 63)
        Me.tabMenu.Name = "tabMenu"
        Me.tabMenu.Size = New System.Drawing.Size(1007, 24)
        Me.tabMenu.TabIndex = 62
        '
        'lblTipo
        '
        Me.lblTipo.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.lblTipo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTipo.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTipo.Location = New System.Drawing.Point(16, 2)
        Me.lblTipo.Name = "lblTipo"
        Me.lblTipo.Size = New System.Drawing.Size(976, 23)
        Me.lblTipo.TabIndex = 241
        Me.lblTipo.Text = "TELEFONOS MOVILES"
        Me.lblTipo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'frmTelefonos
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1018, 743)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblInfo)
        Me.Controls.Add(Me.lblPrograma)
        Me.Controls.Add(Me.lblDescripcion)
        Me.Controls.Add(Me.lblTitulo)
        Me.Controls.Add(Me.lblTeclas)
        Me.Controls.Add(Me.panCampos)
        Me.Controls.Add(Me.tabMenu)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmTelefonos"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Consulta de Telefonos"
        Me.panCampos.ResumeLayout(False)
        Me.panCampos.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Eventos del Menu "

    Private Sub tabMenu_evtEnter(ByVal loTab As prjToolBar.clsTabPage, ByVal e As System.EventArgs) Handles tabMenu.evtEnter
        lblDescripcion.Text = loTab.Tag
    End Sub

    Private Sub tabMenu_evtClick(ByVal loTab As prjToolBar.clsTabPage, ByVal e As System.EventArgs) Handles tabMenu.evtClick
        mrEjecutaAccion(loTab.Indice)
    End Sub

    Private Sub tabMenu_evtKeyDown(ByVal loTab As prjToolBar.clsTabPage, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tabMenu.evtKeyDown
        Dim loTabTemp As prjToolBar.clsTabPage

        Select Case e.KeyValue
            Case 13     ' Intro
                'mrEjecutaAccion(loTab.Indice)
            Case 27     ' Escape
                'End
                Me.Close()
            Case 77     ' moviles
                loTabTemp = tabMenu.mcolTabs(1)
                loTabTemp.moBoton.Focus()
                mrEjecutaAccion(1)
            Case 70     ' fijos
                loTabTemp = tabMenu.mcolTabs(2)
                loTabTemp.moBoton.Focus()
                mrEjecutaAccion(2)
            Case 83     ' salida
                loTabTemp = tabMenu.mcolTabs(3)
                loTabTemp.moBoton.Focus()
                mrEjecutaAccion(3)
        End Select

    End Sub

#End Region

#Region " Funciones y Rutinas varias "

    Public Sub mrCargar(ByVal lnEmpresa As Integer, ByRef loUsuario As clsUsuario)

        ' la llave actual de seguridad *****************************
        gnLlave = 0
        goUsuario = loUsuario
        mnEmpresa = lnEmpresa
        ' **** recupero los datos del profile **************
        modTelefonos.goProfile.mrRecuperaDatos()
        ' **************************************************


        ' ******** seleccion de la empresa contable *****************************
        Dim loEmpresa As New clsEmpresa
        loEmpresa.mnCodigo = mnEmpresa
        loEmpresa.mrRecuperaDatos()

        lblTitulo.Text = loEmpresa.msNombre
        ' **************************************************************************

        Me.ShowDialog()

    End Sub

    Private Sub mrEjecutaAccion(ByVal lnNumero As Integer)

        Select Case lnNumero
            Case 1
                msTipo = "M"
                mrConsulta()
                mrSituaFocoGrid(0)
            Case 2
                msTipo = "F"
                mrConsulta()
                mrSituaFocoGrid(0)
            Case 3
                Me.Close()
        End Select

    End Sub

    Private Sub mrPintaFormulario()
        Dim loEmpresa As New prjEmpresas.clsEmpresa
        Dim loTab As prjToolBar.clsTabPage

        ' pongo los datos de la empresa ********
        loEmpresa.mnCodigo = mnEmpresa
        loEmpresa.mrRecuperaDatos()
        lblPrograma.Text = "TELEFONOS VISANFER"
        lblTitulo.Text = loEmpresa.msNombre
        ' ****************************************************************
        loTab = New prjToolBar.clsTabPage
        loTab.Tag = "CONSULTA DE MOVILES"
        loTab.Titulo = "(M)OVILES"
        loTab.Ancho = 100
        tabMenu.mrAñadeTab(loTab)
        ' ****************************************************************
        loTab = New prjToolBar.clsTabPage
        loTab.Tag = "CONSULTA DE FIJOS"
        loTab.Titulo = "(F)IJOS"
        loTab.Ancho = 100
        tabMenu.mrAñadeTab(loTab)
        ' ****************************************************************
        loTab = New prjToolBar.clsTabPage
        loTab.Tag = "SALIDA DEL PROGRAMA."
        loTab.Titulo = "(S)ALIDA"
        loTab.Ancho = 100
        tabMenu.mrAñadeTab(loTab)
        ' ***** ahora pinto el tool en pantalla
        tabMenu.mrPintaTool()

        mrPreparaGrid()

    End Sub

    Private Sub mrFocoGrid()
        Dim lnI As Integer
        Dim lnY As Integer

        grdLineas.Focus()
        grdLineas.Cell(1, 2).SetFocus()
        lnI = grdLineas.ActiveCell.Row

        ' primero situo en la fila *********
        Select Case lnI
            Case Is = 1
            Case Is = 0
                Send("{DOWN}")
            Case Is > 1
                For lnY = 1 To lnI - 1
                    Send("{UP}")
                Next
        End Select
        Send("{ENTER}")

    End Sub

    Private Sub mrLeeTecla(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim loCaja As control.txtVisanfer
        Dim loGrid As FlexCell.Grid
        Dim lsControl As String = ""
        Dim lbSalta As Boolean = True

        If TypeOf sender Is control.txtVisanfer Then
            loCaja = sender
            lsControl = loCaja.Name
        End If
        If TypeOf sender Is FlexCell.Grid Then
            loGrid = sender
            lsControl = loGrid.Name
        End If

        Select Case e.KeyValue
            Case Keys.F9 And e.Control = True      'CONTROL + F9
                txtBusqueda.SelectAll()
                txtBusqueda.Focus()
            Case Keys.P And e.Control = True     'CONTROL + P
                mrImprimir()
            Case Keys.M And e.Control = True     'CONTROL + M
                mrMantenimiento()
                e.Handled = True
            Case Keys.Escape
                If lsControl = "grdLineas" Then
                    mrConsulta()
                    mtEstado = EstadoVentana.Salida
                    mrLimpiaFormulario()
                    tabMenu.evtFocus()
                End If

            Case Keys.Enter
                Select Case lsControl
                    Case "grdLineas"
                        If mtEstado <> EstadoVentana.Consulta Then
                            Select Case grdLineas.ActiveCell.Col
                                Case 4     ' ultima columna
                                    If grdLineas.ActiveCell.Row = grdLineas.Rows - 1 Then
                                        grdLineas.Rows = grdLineas.Rows + 1
                                    End If
                            End Select
                        End If
                    Case "txtBusqueda"
                        mrBuscarCadena()
                End Select
            Case Keys.F1
                If lsControl = "grdLineas" And mtEstado <> EstadoVentana.Consulta Then
                    mrInsercion()
                End If
            Case Keys.F2
                If lsControl = "grdLineas" And mtEstado <> EstadoVentana.Consulta Then mrBorrado()
            Case Keys.F3
                If lsControl = "grdLineas" Then mrBuscarCadena()
            Case Keys.F5
                mrGrabar(e)
        End Select

    End Sub

    Private Sub mrBuscarCadena()
        ' busca la cadena que hay en el txtbusqueda
        Dim lsCadena As String
        Dim lnPos As Integer
        Dim lnI As Integer
        Dim lbEncontrado As Boolean = False

        lsCadena = Trim(txtBusqueda.Text)
        If lsCadena <> "" Then
            lnPos = grdLineas.ActiveCell.Row + 1
            For lnI = lnPos To grdLineas.Rows - 1
                If InStr(grdLineas.Cell(lnI, 3).Text, lsCadena) > 0 Then
                    lbEncontrado = True
                    Exit For
                End If
            Next
            If Not lbEncontrado Then lnI = 1
            mrSituaFocoGrid(lnI - 1)
        End If

    End Sub

    Private Sub mrSituaFocoGrid(ByVal lnLinea As Integer)
        ' hago esta historia para poder situar el foco en la primera fila *****
        grdLineas.SelectionMode = FlexCell.Grid.SelectionModeEnum.ByCell
        grdLineas.Focus()
        grdLineas.Cell(lnLinea, 2).SetFocus()
        grdLineas.SelectionMode = FlexCell.Grid.SelectionModeEnum.ByRow
        Send("{DOWN}")
    End Sub

    Private Sub mrLimpiaFormulario()
        Dim loControl As Windows.Forms.Control
        Dim loCaja As control.txtVisanfer
        Dim loTextBox As TextBox
        Dim loLabel As System.Windows.Forms.Label
        Dim lnI As Integer
        Dim lnJ As Integer

        For Each loControl In panCampos.Controls
            If TypeOf loControl Is TextBox Then
                loTextBox = loControl
                loTextBox.Text = ""
            End If
            If TypeOf loControl Is control.txtVisanfer Then
                loCaja = loControl
                loCaja.Text = ""
            End If
            If TypeOf loControl Is System.Windows.Forms.Label Then
                loLabel = loControl
                If Mid(loLabel.Name, 1, 3) = "lbl" Then
                    loLabel.Text = ""
                End If
            End If
        Next

        grdLineas.Visible = False
        grdLineas.Rows = 1
        grdLineas.Rows = 34
        For lnI = 1 To grdLineas.Rows - 1
            For lnJ = 0 To grdLineas.Cols - 1
                grdLineas.Cell(lnI, lnJ).Text = ""
            Next
        Next
        grdLineas.Visible = True

    End Sub

    Private Sub mrGrabaLogMod()
        Dim lsLog As String

        ' gabo un log con el usuario que lo hace *********
        lsLog = " -U " & goUsuario.mnCodigo
        prjControl.mrGrabaLineaLog("ModificaTelefonos.log", lsLog)

    End Sub

    Private Sub mrGrabar(ByVal e As System.Windows.Forms.KeyEventArgs)
        ' graba el Telefono en la base de datos /******************

        If mtEstado = EstadoVentana.NuevoRegistro Or _
           mtEstado = EstadoVentana.Mantenimiento Then
            ' compruebo que los campos obligatorios estan cumplimentados
            Dim loTelefono As clsTelefono = Nothing
            Dim loTelefonoAux As clsTelefono
            Dim lnI As Integer

            mrGrabaLogMod()
            Cursor = Cursors.WaitCursor
            mrMoverCampos(2)    ' paso los valores al objeto
            For lnI = 1 To moBusTelefonos.mcolTelefonos.Count
                loTelefono = moBusTelefonos.mcolTelefonos(lnI)
                loTelefonoAux = New clsTelefono
                loTelefonoAux.mnEmpresa = loTelefono.mnEmpresa
                loTelefonoAux.mnCodigo = loTelefono.mnCodigo
                loTelefonoAux.msTipo = msTipo
                loTelefonoAux.mrRecuperaDatos()
                loTelefonoAux.msNumero = loTelefono.msNumero
                loTelefonoAux.msNombre = loTelefono.msNombre
                loTelefonoAux.msExtension = loTelefono.msExtension
                loTelefonoAux.mrGrabaDatos()
            Next
            For lnI = moBusTelefonos.mcolTelefonos.Count + 1 To mnUltimoCodigo
                loTelefonoAux = New clsTelefono
                loTelefonoAux.mnEmpresa = loTelefono.mnEmpresa
                loTelefonoAux.mnCodigo = lnI
                loTelefonoAux.msTipo = msTipo
                loTelefonoAux.mrBorrarDatos()
            Next
            moBusTelefonos.mrBuscaTelefonos(msTipo)
            ' despues de grabar resituo el foco en le inicio
            mrLimpiaFormulario()
            mrMoverCampos(1)
            mrConsulta()
            Cursor = Cursors.Default
        End If

    End Sub

    Private Sub mrImprimir()
        Dim loImpresion As New prjPrinterNet.frmImpresion
        Dim lsFormula As String

        Cursor = Cursors.WaitCursor
        lsFormula = "{telefonos.tip_tel}='" & msTipo & "'"
        loImpresion.msFormula = lsFormula
        loImpresion.msFichero = "rptTelefono.rpt"
        loImpresion.mnEmpresa = mnEmpresa
        loImpresion.msPapel = "A4"

        loImpresion.mrVisualizar()
        'loImpresion.mrListar()

        Cursor = Cursors.Default

    End Sub

    Private Function mfbObligatorios(ByVal e As System.Windows.Forms.KeyEventArgs) As Boolean
        ' Compruebo que los campos obligatorios estan terminados ****************
        ' los campos los controlo por el orden que he estimado oportun0 *********

        'For lnI = 1 To grdLineas.Rows - 1
        '    lnArticulo = mfnInteger(grdLineas.Cell(lnI, 2).Text)
        '    If lnArticulo > 0 Then
        '        Select Case lnArticulo
        '            Case 9999, 99999, 999999
        '            Case Else
        '                lnCantidad = mfnDouble(grdLineas.Cell(lnI, 7).Text)
        '                If lnCantidad = 0 Then
        '                    MsgBox("DEBE PONER ALGUNA CANTIDAD.", MsgBoxStyle.Exclamation, "Visanfer .Net")
        '                    grdLineas.Cell(lnI, 7).SetFocus()
        '                    grdLineas.Focus()
        '                    e.Handled = True        ' capturo el F5 para que no se ejecute otra vez
        '                    Return False
        '                End If
        '                lnAlmacen = mfnInteger(grdLineas.Cell(lnI, 8).Text)
        '                If lnAlmacen = 0 Then
        '                    MsgBox("DEBE PONER ALGUN ALMACEN.", MsgBoxStyle.Exclamation, "Visanfer .Net")
        '                    grdLineas.Cell(lnI, 8).SetFocus()
        '                    grdLineas.Focus()
        '                    e.Handled = True        ' capturo el F5 para que no se ejecute otra vez
        '                    Return False
        '                End If
        '        End Select
        '    End If
        'Next

        Return True

    End Function

    Private Sub mrInsercion()
        Dim lnLinea As Integer

        ' inserto en la linea actual ***********
        lnLinea = grdLineas.ActiveCell.Row
        grdLineas.InsertRow(lnLinea, 1)

    End Sub

    Private Sub mrBorrado()
        Dim lnLinea As Integer

        ' borro la linea actual ***********
        lnLinea = grdLineas.ActiveCell.Row
        grdLineas.Row(lnLinea).Delete()

    End Sub

    Private Sub mrConsulta()

        ' Relleno de los comandos de las teclas *************
        mtEstado = EstadoVentana.Consulta
        mrCargaTelefonos()
        ' ***************************************************
        grdLineas.EnterKeyMoveTo = FlexCell.Grid.MoveToEnum.NextRow
        grdLineas.SelectionBorderColor = Color.White
        grdLineas.SelectionMode = FlexCell.Grid.SelectionModeEnum.ByRow
        grdLineas.Column(1).Visible = True
        grdLineas.AllowUserSort = True
        ' ***************************************************
        lblTeclas.BackColor = Color.FromKnownColor(KnownColor.ActiveBorder)
        lblTeclas.ForeColor = Color.Black
        lblTeclas.Text = " CTRL+M-Mod. Datos           CTRL+P Imprimir              ESC-Salida"
        lblPrograma.Text = "TELEFONOS - CONSULTA"
        If msTipo = "F" Then
            lblTipo.Text = "TELEFONOS FIJOS"
            lblTipo.BackColor = Color.Yellow
        Else
            lblTipo.Text = "TELEFONOS MOVILES"
            lblTipo.BackColor = Color.Aquamarine
        End If
        'lblPrograma.BackColor = loColor.FromKnownColor(KnownColor.ActiveBorder)

    End Sub

    Private Sub mrCargaTelefonos()

        moBusTelefonos = New clsBusTelefonos
        moBusTelefonos.mrBuscaTelefonos(msTipo)
        mrMoverCampos(1)

    End Sub

    Private Sub mrMantenimiento()

        If goUsuario.mfbAccesoPermitido(50, True) Then
            If mtEstado = EstadoVentana.Mantenimiento Then
                mrConsulta()
            Else
                mtEstado = EstadoVentana.Mantenimiento ' entra en modo mantenimiento
                ' ************************************************************
                grdLineas.EnterKeyMoveTo = FlexCell.Grid.MoveToEnum.NextCol
                grdLineas.SelectionBorderColor = Color.Red
                grdLineas.SelectionMode = FlexCell.Grid.SelectionModeEnum.ByCell
                'grdLineas.Selection.BackColor = Color.Aqua
                grdLineas.Column(1).Visible = False
                grdLineas.AllowUserSort = False
                grdLineas.Cell(grdLineas.ActiveCell.Row, 2).SetFocus()
                grdLineas.Focus()
                ' ************************************************************
                lblTeclas.BackColor = Color.Tomato
                lblTeclas.ForeColor = Color.White
                lblTeclas.Text = " F1.-INSERTA LINEA       F2.-BORRA LINEA        " & _
                                    "      F5-.GRABA           ESC-.SALIDA"
                lblPrograma.Text = "TELEFONOS - MANTENIMIENTO"
                'lblPrograma.BackColor = Color.Tomato
            End If
        End If

    End Sub

    Private Sub mrPreparaGrid()

        grdLineas.Row(0).Visible = True
        grdLineas.Rows = 34

        With grdLineas
            .DisplayRowNumber = True
            .Column(1).Locked = False
            .Cols = 6

            '.Column(3).UserSortIndicator = FlexCell.Grid.SortIndicatorEnum.Ascending
            .DisplayFocusRect = False
            .AllowUserResizing = False
            .ExtendLastCol = False
            '.SelectionMode = FlexCell.Grid.SelectionModeEnum.ByCell
            .SelectionMode = FlexCell.Grid.SelectionModeEnum.ByRow
            .EnterKeyMoveTo = FlexCell.Grid.MoveToEnum.NextRow
            .FixedRowColStyle = FlexCell.Grid.FixedRowColStyleEnum.Flat
            .BorderStyle = FlexCell.Grid.BorderStyleEnum.FixedSingle
            .DateFormat = FlexCell.Grid.DateFormatEnum.DMY
            .BackColorSel = Color.Navy
            .BackColorFixed = Color.FromKnownColor(KnownColor.ControlLight)
            '.BackColorBkg = Color.FromKnownColor(KnownColor.InactiveCaption)
            .CellBorderColorFixed = Color.Black
            .GridColor = Color.FromArgb(148, 190, 231)

            .Cell(0, 1).Text = ""
            .Cell(0, 2).Text = "NUMERO"
            .Cell(0, 3).Text = "NOMBRE"
            .Cell(0, 4).Text = "EXTENSION"
            .Cell(0, 5).Text = "ESTADO"

            .Column(1).CellType = FlexCell.Grid.CellTypeEnum.TextBox
            .Column(2).CellType = FlexCell.Grid.CellTypeEnum.TextBox
            .Column(3).CellType = FlexCell.Grid.CellTypeEnum.TextBox
            .Column(4).CellType = FlexCell.Grid.CellTypeEnum.TextBox
            .Column(5).CellType = FlexCell.Grid.CellTypeEnum.TextBox

            '.Column(0).Visible = False
            .Column(1).Width = 0
            .Column(2).Width = 215
            .Column(3).Width = 620
            .Column(4).Width = 80
            .Column(5).Width = 60
            .Column(5).Visible = False

            .Column(2).Alignment = FlexCell.Grid.AlignmentEnum.CenterCenter
            .Column(3).Alignment = FlexCell.Grid.AlignmentEnum.LeftCenter

            .Column(2).MaxLength = 20
            .Column(3).MaxLength = 60
            .Column(4).MaxLength = 4
            .Column(5).MaxLength = 1
        End With

    End Sub

    Private Sub mrMoverCampos(ByVal lnTipo As Integer)
        Dim loTelefono As clsTelefono
        Dim lnI As Integer

        If lnTipo = 1 Then
            lnI = 1
            grdLineas.Visible = False
            If moBusTelefonos.mcolTelefonos.Count > grdLineas.Rows - 1 Then
                grdLineas.Rows = moBusTelefonos.mcolTelefonos.Count + 1
            End If
            For Each loTelefono In moBusTelefonos.mcolTelefonos
                grdLineas.Cell(lnI, 1).Text = loTelefono.mnCodigo
                grdLineas.Cell(lnI, 2).Text = loTelefono.msNumero
                grdLineas.Cell(lnI, 3).Text = loTelefono.msNombre
                grdLineas.Cell(lnI, 4).Text = loTelefono.msExtension
                grdLineas.Cell(lnI, 5).Text = ""
                If loTelefono.mbPrivado Then
                    grdLineas.Cell(lnI, 2).BackColor = Color.OrangeRed
                    grdLineas.Cell(lnI, 2).FontItalic = True
                Else
                    grdLineas.Cell(lnI, 2).BackColor = Color.White
                    grdLineas.Cell(lnI, 2).FontItalic = False
                End If
                lnI = lnI + 1
            Next
            'mnUltimoCodigo = lnI - 1
            mnUltimoCodigo = moBusTelefonos.mcolTelefonos.Count
            grdLineas.Visible = True
        Else
            moBusTelefonos.mcolTelefonos = New Collection
            For lnI = 1 To grdLineas.Rows - 1
                loTelefono = New clsTelefono
                loTelefono.mnEmpresa = mnEmpresa
                loTelefono.mnCodigo = lnI
                loTelefono.msTipo = msTipo
                loTelefono.msNumero = Trim(grdLineas.Cell(lnI, 2).Text)
                loTelefono.msNombre = Trim(grdLineas.Cell(lnI, 3).Text)
                loTelefono.msExtension = Trim(grdLineas.Cell(lnI, 4).Text)
                If loTelefono.msNumero <> "" Then moBusTelefonos.mcolTelefonos.Add(loTelefono, loTelefono.mpsCodigo)
            Next
        End If

    End Sub

    Private Sub mrPendiente(ByVal lnFila As Integer, ByVal lsValor As String)
        Dim lnI As Integer
        Dim loColor As Color

        If lsValor = "S" Then
            loColor = Color.Red
        Else
            loColor = Color.Black
        End If

        For lnI = 0 To 6
            'grdLineas.ColorCelda(lnI, lnFila).mnForeColor = loColor
        Next

    End Sub

    Private Sub mrAsignaEventos()
        AddHandler txtBusqueda.KeyDown, AddressOf mrLeeTecla
        AddHandler grdLineas.KeyDown, AddressOf mrLeeTecla
    End Sub

    Private Sub mrPintaNumero()
        Dim lnFila As Integer
        lnFila = grdLineas.ActiveCell.Row

        On Error Resume Next
        If grdLineas.Cell(lnFila, 2).BackColor.Name = "OrangeRed" Then
            lblTelefono.BackColor = Color.OrangeRed
            lblPrivado.Visible = True
        Else
            lblTelefono.BackColor = Color.FromName("Info")
            lblPrivado.Visible = False
        End If
        lblTelefono.Text = grdLineas.Cell(lnFila, 2).Text
        lblNombre.Text = grdLineas.Cell(lnFila, 3).Text
        lblExtension.Text = grdLineas.Cell(lnFila, 4).Text
        On Error GoTo 0

    End Sub

#End Region

#Region " Eventos de Formulario "

    Private Sub frmTelefonos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        mbPrimeraVez = True
        moTelefono = New clsTelefono
        mrAsignaEventos()
        mrPintaFormulario()
        tabMenu.evtFocus()

    End Sub

    Private Sub frmTelefonos_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyData = 131138 Then goUsuario.mrBloquear(gnLlave)
    End Sub

    Private Sub frmTelefonos_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        'Me.Text = goUsuario.msNombre
        If goUsuario Is Nothing Then goUsuario = New clsUsuario
        If goUsuario.mbEsNuevo Then goUsuario.mrBloquear(gnLlave)
        If Not goUsuario.msNombre Is Nothing Then Me.Text = goUsuario.msNombre

        lblTitulo.BackColor = mfoColorFondo(mnEmpresa)
        lblTitulo.ForeColor = mfoColorTexto(mnEmpresa)
        lblPrograma.BackColor = lblTitulo.BackColor
        lblPrograma.ForeColor = lblTitulo.ForeColor
    End Sub

#End Region

#Region " Eventos del Grid "

    Private Sub grdLineas_KeyPress(ByVal Sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles grdLineas.KeyPress
        ' CON ESTA RUTINA CONVIERTO LAS MINUSCULAS EN MAYUSCULAS ********************
        ' ESTO ES POR QUE EL MASKENUM.UPPER NO DEJA METER LOS ESPACIOS EN BLANCO ****
        Dim lnCodigo As Integer
        lnCodigo = Asc(e.KeyChar)
        Select Case lnCodigo
            Case 97 To 122
                e.Handled = True
                Send(Chr(lnCodigo - 32))
        End Select
    End Sub

    Private Sub grdLineas_RowColChange(ByVal Sender As Object, ByVal e As FlexCell.Grid.RowColChangeEventArgs) Handles grdLineas.RowColChange
        mrPintaNumero()
    End Sub

#End Region

End Class
