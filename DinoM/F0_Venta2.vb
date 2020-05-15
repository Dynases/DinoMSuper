Imports Logica.AccesoLogica
Imports Janus.Windows.GridEX
Imports DevComponents.DotNetBar
Imports System.IO
Imports DevComponents.DotNetBar.SuperGrid
Imports GMap.NET.MapProviders
Imports GMap.NET
Imports GMap.NET.WindowsForms.Markers
Imports GMap.NET.WindowsForms
Imports GMap.NET.WindowsForms.ToolTips
Imports System.Drawing
Imports DevComponents.DotNetBar.Controls
Imports System.Threading
Imports System.Drawing.Text
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Drawing.Printing
Imports CrystalDecisions.Shared
Imports Facturacion
Public Class F0_Venta2
#Region "Variables Globales"
    Dim _CodCliente As Integer = 0
    Dim _CodEmpleado As Integer = 0
    Dim OcultarFact As Integer = 0
    Dim _codeBar As Integer = 1
    Dim _dias As Integer = 0
    Public _nameButton As String
    Public _tab As SuperTabItem
    Public _modulo As SideNavItem
    Dim FilaSelectLote As DataRow = Nothing
    Dim Table_Producto As DataTable
    Dim G_Lote As Boolean = False '1=igual a mostrar las columnas de lote y fecha de Vencimiento

#End Region
#Region "Metodos Privados"
    Private Sub _IniciarTodo()
        L_prAbrirConexion(gs_Ip, gs_UsuarioSql, gs_ClaveSql, gs_NombreBD)
        MSuperTabControl.SelectedTabIndex = 0
        Me.WindowState = FormWindowState.Maximized

        _prValidarLote()
        _prCargarComboLibreriaSucursal(cbSucursal)
        _prCargarComboLibreria(cbCambioDolar, 7, 1)
        cbCambioDolar.Value = 1
        'lbTipoMoneda.Visible = False
        swMoneda.Visible = False
        P_prCargarVariablesIndispensables()
        _prCargarVenta()
        _prInhabiliitar()
        grVentas.Focus()
        Me.Text = "DESPACHOS"
        Dim blah As New Bitmap(New Bitmap(My.Resources.compra), 20, 20)
        Dim ico As Icon = Icon.FromHandle(blah.GetHicon())
        Me.Icon = ico
        _prAsignarPermisos()
        P_prCargarParametro()
        _prValidadFactura()
        _prCargarNameLabel()
        'Ocultar el botón Modificar
        btnModificar.Visible = False

    End Sub

    Public Sub _prCargarNameLabel()
        Dim dt As DataTable = L_fnNameLabel()
        If (dt.Rows.Count > 0) Then
            _codeBar = 1 'dt.Rows(0).Item("codeBar")
        End If
    End Sub
    Sub _prValidadFactura()
        'If (OcultarFact = 1) Then
        '    GroupPanelFactura2.Visible = False
        '    GroupPanelFactura.Visible = False
        'Else
        '    GroupPanelFactura2.Visible = True
        '    GroupPanelFactura.Visible = True
        'End If

    End Sub
    Public Sub _prValidarLote()
        Dim dt As DataTable = L_fnPorcUtilidad()
        If (dt.Rows.Count > 0) Then
            Dim lot As Integer = dt.Rows(0).Item("VerLote")
            OcultarFact = dt.Rows(0).Item("VerFactManual")
            If (lot = 1) Then
                G_Lote = True
            Else
                G_Lote = False
            End If

        End If
    End Sub
    Private Sub _prCargarComboLibreriaSucursal(mCombo As Janus.Windows.GridEX.EditControls.MultiColumnCombo)
        Dim dt As New DataTable
        dt = L_fnListarSucursales()
        With mCombo
            .DropDownList.Columns.Clear()
            .DropDownList.Columns.Add("aanumi").Width = 60
            .DropDownList.Columns("aanumi").Caption = "COD"
            .DropDownList.Columns.Add("aabdes").Width = 500
            .DropDownList.Columns("aabdes").Caption = "SUCURSAL"
            .ValueMember = "aanumi"
            .DisplayMember = "aabdes"
            .DataSource = dt
            .Refresh()
        End With
    End Sub
    Private Sub _prAsignarPermisos()

        Dim dtRolUsu As DataTable = L_prRolDetalleGeneral(gi_userRol, _nameButton)

        Dim show As Boolean = dtRolUsu.Rows(0).Item("ycshow")
        Dim add As Boolean = dtRolUsu.Rows(0).Item("ycadd")
        Dim modif As Boolean = dtRolUsu.Rows(0).Item("ycmod")
        Dim del As Boolean = dtRolUsu.Rows(0).Item("ycdel")

        If add = False Then
            btnNuevo.Visible = False
        End If
        If modif = False Then
            btnModificar.Visible = False
        End If
        If del = False Then
            btnEliminar.Visible = False
        End If
    End Sub
    Private Sub _prInhabiliitar()

        tbCodigo.ReadOnly = True
        tbCliente.ReadOnly = True
        tbVendedor.ReadOnly = True
        tbFechaVenta.IsInputReadOnly = True
        swMoneda.IsReadOnly = True
        tbFechaVenc.IsInputReadOnly = True
        swTipoVenta.IsReadOnly = True
        txtEstado.ReadOnly = True

        'Datos facturacion
        tbNroAutoriz.ReadOnly = True
        tbNroFactura.ReadOnly = True
        tbCodigoControl.ReadOnly = True
        dtiFechaFactura.IsInputReadOnly = True
        dtiFechaFactura.ButtonDropDown.Enabled = False

        btnModificar.Enabled = True
        btnGrabar.Enabled = False
        btnNuevo.Enabled = True
        'btnEliminar.Enabled = True

        If grVentas.GetValue("taest") = 1 Then
            btnEliminar.Enabled = True
        Else
            btnEliminar.Enabled = False
        End If

        tbSubTotal.IsInputReadOnly = True
        tbIce.IsInputReadOnly = True
        tbPrueba.IsInputReadOnly = True
        tbMontoBs.IsInputReadOnly = True
        tbMontoDolar.IsInputReadOnly = True
        tbMontoTarej.IsInputReadOnly = True
        'txtCambio1.IsInputReadOnly = True

        'txtMontoPagado1.IsInputReadOnly = True

        grVentas.Enabled = True
        PanelNavegacion.Enabled = True
        grdetalle.RootTable.Columns("img").Visible = False
        If (GPanelProductos.Visible = True) Then
            _DesHabilitarProductos()
        End If

        tbNit.ReadOnly = True
        TbNombre1.ReadOnly = True
        TbNombre2.ReadOnly = True
        cbSucursal.ReadOnly = True
        FilaSelectLote = Nothing
    End Sub
    Private Sub _prhabilitar()

        grVentas.Enabled = False
        tbCodigo.ReadOnly = False
        ''  tbCliente.ReadOnly = False  por que solo podra seleccionar Cliente
        ''  tbVendedor.ReadOnly = False
        tbFechaVenc.IsInputReadOnly = False
        swTipoVenta.IsReadOnly = False
        tbFechaVenta.IsInputReadOnly = False

        swMoneda.IsReadOnly = False

        btnGrabar.Enabled = True

        tbNit.ReadOnly = False
        TbNombre1.ReadOnly = False
        TbNombre2.ReadOnly = False

        'Datos facturacion
        tbNroAutoriz.ReadOnly = False
        tbNroFactura.ReadOnly = False
        tbCodigoControl.ReadOnly = False
        dtiFechaFactura.IsInputReadOnly = False

        tbMontoBs.IsInputReadOnly = False
        tbMontoDolar.IsInputReadOnly = False
        tbMontoTarej.IsInputReadOnly = False
        'txtCambio1.IsInputReadOnly = False
        'txtMontoPagado1.IsInputReadOnly = False

        If (tbCodigo.Text.Length > 0) Then
            cbSucursal.ReadOnly = True
        Else
            cbSucursal.ReadOnly = False

        End If
    End Sub
    Public Sub _prFiltrar()
        'cargo el buscador
        Dim _Mpos As Integer
        _prCargarVenta()
        If grVentas.RowCount > 0 Then
            _Mpos = 0
            grVentas.Row = _Mpos
        Else
            _Limpiar()
            LblPaginacion.Text = "0/0"
        End If
    End Sub
    Private Sub _Limpiar()

        tbCodigo.Clear()
        'tbCliente.Clear()
        'tbVendedor.Clear()

        swMoneda.Value = True

        '_CodCliente = 0
        '_CodEmpleado = 0
        tbFechaVenta.Value = Now.Date
        swTipoVenta.Value = True
        tbFechaVenc.Visible = False
        lbCredito.Visible = False
        _prCargarDetalleVenta(-1)
        MSuperTabControl.SelectedTabIndex = 0
        tbSubTotal.Value = 0
        tbPdesc.Value = 0
        tbMdesc.Value = 0
        tbIce.Value = 0
        tbPrueba.Value = 0
        tbMontoBs.Value = 0
        tbMontoDolar.Value = 0
        tbMontoTarej.Value = 0
        'txtCambio1.Value = 0
        'txtMontoPagado1.Value = 0
        txtCambio1.Text = "0.00"
        txtMontoPagado1.Text = "0.00"
        tbTotalBs.Text = "0.00"
        tbTotalDo.Text = "0.00"

        txtEstado.BackColor = Color.White
        txtEstado.Clear()

        With grdetalle.RootTable.Columns("img")
            .Width = 80
            .Caption = "Eliminar"
            .CellStyle.ImageHorizontalAlignment = ImageHorizontalAlignment.Center
            .Visible = True
        End With
        _prAddDetalleVenta()
        If (GPanelProductos.Visible = True) Then
            GPanelProductos.Visible = False

            PanelInferior.Visible = True
        End If
        'tbCliente.Focus()

        tbNit.Clear()
        TbNombre1.Clear()
        TbNombre2.Clear()
        tbNit.Select()
        tbNroAutoriz.Clear()
        tbNroFactura.Clear()
        tbCodigoControl.Clear()
        dtiFechaFactura.Value = Now.Date
        If (CType(cbSucursal.DataSource, DataTable).Rows.Count > 0) Then
            cbSucursal.SelectedIndex = 0
        End If
        FilaSelectLote = Nothing

        ' tbCliente.Focus()
        Table_Producto = Nothing
    End Sub
    Public Sub _prMostrarRegistro(_N As Integer)
        '' grVentas.Row = _N
        '     a.tanumi ,a.taalm ,a.tafdoc ,a.taven ,vendedor .yddesc as vendedor ,a.tatven ,a.tafvcr ,a.taclpr,
        'cliente.yddesc as cliente ,a.tamon ,IIF(tamon=1,'Boliviano','Dolar') as moneda,a.taest ,a.taobs ,
        'a.tadesc ,a.tafact ,a.tahact ,a.tauact,(Sum(b.tbptot)-a.tadesc ) as total,taproforma

        With grVentas
            cbSucursal.Value = .GetValue("taalm")
            tbCodigo.Text = .GetValue("tanumi")
            tbFechaVenta.Value = .GetValue("tafdoc")
            _CodEmpleado = .GetValue("taven")
            tbVendedor.Text = .GetValue("vendedor")
            _CodCliente = .GetValue("taclpr")
            tbCliente.Text = .GetValue("cliente")
            swMoneda.Value = .GetValue("tamon")
            tbFechaVenc.Value = .GetValue("tafvcr")
            swTipoVenta.Value = .GetValue("tatven")
            If grVentas.GetValue("taest") = 1 Then
                txtEstado.Text = "VIGENTE"
                txtEstado.BackColor = Color.Green
                btnEliminar.Enabled = True
            Else
                txtEstado.Text = "ANULADO"
                txtEstado.BackColor = Color.Red
                btnEliminar.Enabled = False
            End If
            'If (gb_FacturaEmite) Then
            Dim dt As DataTable = L_fnObtenerTabla("TFV001", "fvanitcli, fvadescli1, fvadescli2, fvaautoriz, fvanfac, fvaccont, fvafec", "fvanumi=" + tbCodigo.Text.Trim)
            If (dt.Rows.Count = 1) Then
                tbNit.Text = dt.Rows(0).Item("fvanitcli").ToString
                TbNombre1.Text = dt.Rows(0).Item("fvadescli1").ToString
                TbNombre2.Text = dt.Rows(0).Item("fvadescli2").ToString

                tbNroAutoriz.Text = dt.Rows(0).Item("fvaautoriz").ToString
                tbNroFactura.Text = dt.Rows(0).Item("fvanfac").ToString
                tbCodigoControl.Text = dt.Rows(0).Item("fvaccont").ToString
                dtiFechaFactura.Value = dt.Rows(0).Item("fvafec")
            Else
                tbNit.Clear()
                TbNombre1.Clear()
                TbNombre2.Clear()

                tbNroAutoriz.Clear()
                tbNroFactura.Clear()
                tbCodigoControl.Clear()
                dtiFechaFactura.Value = "2000/01/01"
            End If
            'End If

            lbFecha.Text = CType(.GetValue("tafact"), Date).ToString("dd/MM/yyyy")
            lbHora.Text = .GetValue("tahact").ToString
            lbUsuario.Text = .GetValue("tauact").ToString

        End With

        _prCargarDetalleVenta(tbCodigo.Text)
        tbMdesc.Value = grVentas.GetValue("tadesc")
        tbIce.Value = grVentas.GetValue("taice")
        _prCalcularPrecioTotal()
        'Calcular montos
        Dim tMonto As DataTable = L_fnMostrarMontos(tbCodigo.Text)
        If tMonto.Rows.Count > 0 Then
            tbMontoBs.Value = tMonto.Rows(0).Item("tgMontBs").ToString
            tbMontoDolar.Value = tMonto.Rows(0).Item("tgMontDol").ToString
            tbMontoTarej.Value = tMonto.Rows(0).Item("tgMontTare").ToString
            cbCambioDolar.Text = tMonto.Rows(0).Item("tgCambioDol").ToString

            If Convert.ToDecimal(tbTotalBs.Text) <> 0 And Convert.ToDecimal(txtMontoPagado1.Text) >= Convert.ToDecimal(tbTotalBs.Text) Then
                txtCambio1.Text = Convert.ToDecimal(txtMontoPagado1.Text) - Convert.ToDecimal(tbTotalBs.Text)
            Else
                txtCambio1.Text = "0.00"
            End If
        End If
        LblPaginacion.Text = Str(grVentas.Row + 1) + "/" + grVentas.RowCount.ToString

    End Sub

    Private Sub _prCargarDetalleVenta(_numi As String)
        Dim dt As New DataTable
        dt = L_fnDetalleVenta(_numi)
        grdetalle.DataSource = dt
        grdetalle.RetrieveStructure()
        grdetalle.AlternatingColors = True
        ' a.tbnumi ,a.tbtv1numi ,a.tbty5prod ,b.yfcdprod1 as producto,a.tbest ,a.tbcmin ,a.tbumin ,Umin .ycdes3 as unidad,a.tbpbas ,a.tbptot,a.tbdesc ,a.tbobs ,
        'a.tbfact ,a.tbhact ,a.tbuact

        With grdetalle.RootTable.Columns("tbnumi")
            .Width = 100
            .Caption = "CODIGO"
            .Visible = False
        End With

        With grdetalle.RootTable.Columns("tbtv1numi")
            .Width = 90
            .Visible = False
        End With
        With grdetalle.RootTable.Columns("tbty5prod")
            .Width = 90
            .Visible = False
        End With
        'If _codeBar = 2 Then
        '    With grdetalle.RootTable.Columns("yfcbarra")
        '        .Caption = "Cod.Barra"
        '        .Width = 100
        '        .Visible = True

        '    End With
        'Else
        '    With grdetalle.RootTable.Columns("yfcbarra")
        '        .Caption = "Cod.Barra"
        '        .Width = 100
        '        .Visible = False
        '    End With
        'End If

        With grdetalle.RootTable.Columns("Codigo")
            .Caption = "Código".ToUpper
            .Width = 100
            .Visible = False
        End With

        With grdetalle.RootTable.Columns("yfcbarra")
            .Caption = "Cod.Barra".ToUpper
            .Width = 155
            .Visible = gb_CodigoBarra

        End With

        With grdetalle.RootTable.Columns("producto")
            .Caption = "Productos".ToUpper
            .Width = 250
            .Visible = True

        End With
        With grdetalle.RootTable.Columns("tbest")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = False
        End With

        With grdetalle.RootTable.Columns("tbcmin")
            .Width = 130
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Far
            .Visible = True
            .FormatString = "0.00000"
            .Caption = "Cantidad".ToUpper
        End With
        With grdetalle.RootTable.Columns("tbumin")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Far
            .Visible = False
        End With
        With grdetalle.RootTable.Columns("unidad")
            .Width = 90
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Far
            .Visible = True
            .Caption = "Unidad".ToUpper
        End With
        With grdetalle.RootTable.Columns("tbpbas")
            .Width = 110
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Far
            .Visible = True
            .FormatString = "0.00"
            .Caption = "Precio U.".ToUpper
        End With
        With grdetalle.RootTable.Columns("tbptot")
            .Width = 100
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Far
            .Visible = False
            .FormatString = "0.00"
            .Caption = "Total".ToUpper
        End With
        With grdetalle.RootTable.Columns("tbporc")
            .Width = 100
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Far
            .Visible = False
            .FormatString = "0.00"
            .Caption = "P.Desc(%)".ToUpper
        End With
        With grdetalle.RootTable.Columns("tbdesc")
            .Width = 100
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Far
            .Visible = False
            .FormatString = "0.00"
            .Caption = "M.Desc".ToUpper
        End With
        With grdetalle.RootTable.Columns("tbtotdesc")
            .Width = 100
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Far
            .Visible = True
            .FormatString = "0.00"
            .Caption = "Total".ToUpper
        End With
        With grdetalle.RootTable.Columns("tbobs")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = False
        End With
        With grdetalle.RootTable.Columns("tbpcos")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Far
            .Visible = False
        End With
        With grdetalle.RootTable.Columns("tbptot2")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Far
            .Visible = False
        End With
        With grdetalle.RootTable.Columns("tbfact")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = False
        End With
        With grdetalle.RootTable.Columns("tbhact")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = False
        End With
        With grdetalle.RootTable.Columns("tbuact")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = False
        End With
        With grdetalle.RootTable.Columns("estado")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = False
        End With
        With grdetalle.RootTable.Columns("img")
            .Width = 80
            .Caption = "Eliminar".ToUpper
            .CellStyle.ImageHorizontalAlignment = ImageHorizontalAlignment.Center
            .Visible = False
        End With
        If (G_Lote = True) Then
            With grdetalle.RootTable.Columns("tblote")
                .Width = 120
                .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
                .Visible = True
                .Caption = "LOTE"
            End With
            With grdetalle.RootTable.Columns("tbfechaVenc")
                .Width = 120
                .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
                .Visible = True
                .Caption = "FECHA VENC."
                .FormatString = "yyyy/MM/dd"
            End With

        Else
            With grdetalle.RootTable.Columns("tblote")
                .Width = 120
                .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
                .Visible = False
                .Caption = "LOTE"
            End With
            With grdetalle.RootTable.Columns("tbfechaVenc")
                .Width = 120
                .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
                .Visible = False
                .Caption = "FECHA VENC."
                .FormatString = "yyyy/MM/dd"
            End With
        End If
        With grdetalle.RootTable.Columns("stock")
            .Width = 120
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = False
        End With
        With grdetalle
            .GroupByBoxVisible = False
            'diseño de la grilla
            .VisualStyle = VisualStyle.Office2007
        End With
    End Sub

    Private Sub _prCargarVenta()
        Dim dt As New DataTable
        dt = L_fnGeneralVenta()
        grVentas.DataSource = dt
        grVentas.RetrieveStructure()
        grVentas.AlternatingColors = True
        '   a.tamon ,IIF(tamon=1,'Boliviano','Dolar') as moneda,a.taest ,a.taobs ,
        'a.tadesc ,a.tafact ,a.tahact ,a.tauact,(Sum(b.tbptot)-a.tadesc ) as total

        With grVentas.RootTable.Columns("tanumi")
            .Width = 100
            .Caption = "CODIGO"
            .Visible = True

        End With

        With grVentas.RootTable.Columns("taalm")
            .Width = 90
            .Visible = False
        End With

        With grVentas.RootTable.Columns("taproforma")
            .Width = 90
            .Visible = False
        End With
        With grVentas.RootTable.Columns("tafdoc")
            .Width = 90
            .Visible = True
            .Caption = "FECHA"
        End With

        With grVentas.RootTable.Columns("taven")
            .Width = 160
            .Visible = False
        End With
        With grVentas.RootTable.Columns("vendedor")
            .Width = 250
            .Visible = True
            .Caption = "VENDEDOR".ToUpper
        End With


        With grVentas.RootTable.Columns("tatven")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = False
        End With

        With grVentas.RootTable.Columns("tafvcr")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = False
        End With
        With grVentas.RootTable.Columns("taclpr")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = False
        End With
        With grVentas.RootTable.Columns("cliente")
            .Width = 250
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = True
            .Caption = "CLIENTE"
        End With

        With grVentas.RootTable.Columns("tamon")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = False
        End With
        With grVentas.RootTable.Columns("moneda")
            .Width = 150
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = True
            .Caption = "MONEDA"
        End With
        With grVentas.RootTable.Columns("taobs")
            .Width = 200
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = True
            .Caption = "OBSERVACION"
        End With
        With grVentas.RootTable.Columns("tadesc")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = False
        End With
        With grVentas.RootTable.Columns("taest")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = False
        End With
        With grVentas.RootTable.Columns("taice")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = False
        End With
        With grVentas.RootTable.Columns("tafact")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = False
        End With
        With grVentas.RootTable.Columns("tahact")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = False
        End With
        With grVentas.RootTable.Columns("tauact")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = False
        End With
        With grVentas.RootTable.Columns("total")
            .Width = 150
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Far
            .Visible = True
            .Caption = "TOTAL"
            .FormatString = "0.00"
        End With
        With grVentas
            .DefaultFilterRowComparison = FilterConditionOperator.Contains
            .FilterMode = FilterMode.Automatic
            .FilterRowUpdateMode = FilterRowUpdateMode.WhenValueChanges
            .GroupByBoxVisible = False
            'diseño de la grilla

        End With

        If (dt.Rows.Count <= 0) Then
            _prCargarDetalleVenta(-1)
        End If
    End Sub

    Public Sub _prGuardar()
        If _ValidarCampos() = False Then
            Exit Sub
        End If

        If (tbCodigo.Text = String.Empty) Then
            _GuardarNuevo()
        Else
            If (tbCodigo.Text <> String.Empty) Then
                _prGuardarModificado()
                ''    _prInhabiliitar() RODRIGO RLA

            End If
        End If
    End Sub
    Public Sub actualizarSaldoSinLote(ByRef dt As DataTable)
        'b.yfcdprod1 ,a.iclot ,a.icfven  ,a.iccven 

        '      a.tbnumi ,a.tbtv1numi ,a.tbty5prod ,b.yfcdprod1 as producto,a.tbest ,a.tbcmin ,a.tbumin ,Umin .ycdes3 as unidad,a.tbpbas ,a.tbptot ,a.tbobs ,
        'a.tbpcos,a.tblote ,a.tbfechaVenc , a.tbptot2, a.tbfact ,a.tbhact ,a.tbuact,1 as estado,Cast(null as Image) as img,
        'Cast (0 as decimal (18,2)) as stock
        Dim _detalle As DataTable = CType(grdetalle.DataSource, DataTable)

        For i As Integer = 0 To dt.Rows.Count - 1 Step 1
            Dim sum As Integer = 0
            Dim codProducto As Integer = dt.Rows(i).Item("yfnumi")
            For j As Integer = 0 To grdetalle.RowCount - 1 Step 1
                grdetalle.Row = j
                Dim estado As Integer = grdetalle.GetValue("estado")
                If (estado = 0) Then
                    If (codProducto = grdetalle.GetValue("tbty5prod")) Then
                        sum = sum + grdetalle.GetValue("tbcmin")
                    End If
                End If
            Next
            dt.Rows(i).Item("stock") = dt.Rows(i).Item("stock") - sum
        Next

    End Sub

    Private Sub _prCargarProductos(_cliente As String)
        If (cbSucursal.SelectedIndex < 0) Then
            Return
        End If
        Dim dtname As DataTable = L_fnNameLabel()
        Dim dt As New DataTable

        If (G_Lote = True) Then
            dt = L_fnListarProductos(cbSucursal.Value, _cliente)  ''1=Almacen
            'Table_Producto = dt.Copy
        Else
            dt = L_fnListarProductosSinLote(cbSucursal.Value, _cliente, CType(grdetalle.DataSource, DataTable))  ''1=Almacen
            'Table_Producto = dt.Copy
        End If



        ''  actualizarSaldoSinLote(dt)
        grProductos.DataSource = dt
        grProductos.RetrieveStructure()
        grProductos.AlternatingColors = True

        '      a.yfnumi ,a.yfcprod ,a.yfcdprod1,a.yfcdprod2 ,a.yfgr1,gr1.ycdes3 as grupo1,a.yfgr2
        ',gr2.ycdes3 as grupo2 ,a.yfgr3,gr3.ycdes3 as grupo3,a.yfgr4 ,gr4 .ycdes3 as grupo4,a.yfumin ,Umin .ycdes3 as UnidMin
        ' ,b.yhprecio 

        With grProductos.RootTable.Columns("yfnumi")
            .Width = 100
            .Caption = "Código"
            .Visible = False

        End With
        With grProductos.RootTable.Columns("yfcprod")
            .Width = 100
            .Caption = "Código"
            .Visible = True
        End With
        With grProductos.RootTable.Columns("yfcbarra")
            .Width = 120
            .Caption = "Cod. Barra"
            .Visible = gb_CodigoBarra
        End With
        With grProductos.RootTable.Columns("yfcdprod1")
            .Width = 250
            .Visible = True
            .Caption = "Descripción"
        End With
        With grProductos.RootTable.Columns("yfcdprod2")
            .Width = 150
            .Visible = False
            .Caption = "Descripcion Corta"
        End With


        With grProductos.RootTable.Columns("yfgr1")
            .Width = 160
            .Visible = False
        End With
        If (dtname.Rows.Count > 0) Then

            With grProductos.RootTable.Columns("grupo1")
                .Width = 120
                .Caption = dtname.Rows(0).Item("Grupo 1").ToString
                .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
                .Visible = False
            End With
            With grProductos.RootTable.Columns("grupo2")
                .Width = 120
                .Caption = dtname.Rows(0).Item("Grupo 2").ToString
                .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
                .Visible = False
            End With

            With grProductos.RootTable.Columns("grupo3")
                .Width = 120
                .Caption = dtname.Rows(0).Item("Grupo 3").ToString
                .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
                .Visible = False
            End With
            With grProductos.RootTable.Columns("grupo4")
                .Width = 120
                .Caption = dtname.Rows(0).Item("Grupo 4").ToString
                .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
                .Visible = False
            End With
        Else
            With grProductos.RootTable.Columns("grupo1")
                .Width = 120
                .Caption = "Grupo 1"
                .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
                .Visible = False
            End With
            With grProductos.RootTable.Columns("grupo2")
                .Width = 120
                .Caption = "Grupo 2"
                .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
                .Visible = False
            End With
            With grProductos.RootTable.Columns("grupo3")
                .Width = 120
                .Caption = "Grupo 3"
                .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
                .Visible = False
            End With
            With grProductos.RootTable.Columns("grupo4")
                .Width = 120
                .Caption = "Grupo 4"
                .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
                .Visible = False
            End With
        End If


        With grProductos.RootTable.Columns("yfgr2")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = False
        End With

        With grProductos.RootTable.Columns("yfgr3")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = False
        End With
        With grProductos.RootTable.Columns("validacion")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = False
        End With
        With grProductos.RootTable.Columns("yfgr4")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = False
        End With


        With grProductos.RootTable.Columns("yfumin")
            .Width = 50
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = False
        End With
        With grProductos.RootTable.Columns("UnidMin")
            .Width = 120
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = True
            .Caption = "Unidad Min."
        End With
        With grProductos.RootTable.Columns("yhprecio")
            .Width = 120
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = True
            .Caption = "Precio"
            .FormatString = "0.00"
        End With
        With grProductos.RootTable.Columns("pcos")
            .Width = 120
            .CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Near
            .Visible = False
            .Caption = "Precio Costo"
            .FormatString = "0.00"
        End With
        With grProductos.RootTable.Columns("stock")
            .Width = 120
            .FormatString = "0.00"
            .Visible = True
            .Caption = "Stock"
        End With

        With grProductos
            .DefaultFilterRowComparison = FilterConditionOperator.Contains
            .FilterMode = FilterMode.Automatic
            .FilterRowUpdateMode = FilterRowUpdateMode.WhenValueChanges
            .GroupByBoxVisible = False
            'diseño de la grilla
            .VisualStyle = VisualStyle.Office2007
        End With
        _prAplicarCondiccionJanusSinLote()
    End Sub
    Public Sub _prAplicarCondiccionJanusSinLote()
        'Dim fc As GridEXFormatCondition
        'fc = New GridEXFormatCondition(grProductos.RootTable.Columns("stock"), ConditionOperator.Between, -9998 And 0)
        ''fc.FormatStyle.FontBold = TriState.True
        'fc.FormatStyle.ForeColor = Color.Red    'Color.Tan
        'grProductos.RootTable.FormatConditions.Add(fc)
        Dim fr As GridEXFormatCondition
        fr = New GridEXFormatCondition(grProductos.RootTable.Columns("validacion"), ConditionOperator.Equal, 1)
        fr.FormatStyle.ForeColor = Color.Red
        grProductos.RootTable.FormatConditions.Add(fr)
    End Sub


    Public Sub actualizarSaldo(ByRef dt As DataTable, CodProducto As Integer)
        'b.yfcdprod1 ,a.iclot ,a.icfven  ,a.iccven 

        '      a.tbnumi ,a.tbtv1numi ,a.tbty5prod ,b.yfcdprod1 as producto,a.tbest ,a.tbcmin ,a.tbumin ,Umin .ycdes3 as unidad,a.tbpbas ,a.tbptot ,a.tbobs ,
        'a.tbpcos,a.tblote ,a.tbfechaVenc , a.tbptot2, a.tbfact ,a.tbhact ,a.tbuact,1 as estado,Cast(null as Image) as img,
        'Cast (0 as decimal (18,2)) as stock
        Dim _detalle As DataTable = CType(grdetalle.DataSource, DataTable)

        For i As Integer = 0 To dt.Rows.Count - 1 Step 1
            Dim lote As String = dt.Rows(i).Item("iclot")
            Dim FechaVenc As Date = dt.Rows(i).Item("icfven")
            Dim sum As Integer = 0
            For j As Integer = 0 To _detalle.Rows.Count - 1
                Dim estado As Integer = _detalle.Rows(j).Item("estado")
                If (estado = 0) Then
                    If (lote = _detalle.Rows(j).Item("tblote") And
                        FechaVenc = _detalle.Rows(j).Item("tbfechaVenc") And CodProducto = _detalle.Rows(j).Item("tbty5prod")) Then
                        sum = sum + _detalle.Rows(j).Item("tbcmin")
                    End If
                End If
            Next
            dt.Rows(i).Item("iccven") = dt.Rows(i).Item("iccven") - sum
        Next

    End Sub

    Private Sub _prCargarLotesDeProductos(CodProducto As Integer, nameProducto As String)
        If (cbSucursal.SelectedIndex < 0) Then
            Return
        End If
        Dim dt As New DataTable
        GPanelProductos.Text = nameProducto
        dt = L_fnListarLotesPorProductoVenta(cbSucursal.Value, CodProducto)  ''1=Almacen
        actualizarSaldo(dt, CodProducto)
        grProductos.DataSource = dt
        grProductos.RetrieveStructure()
        grProductos.AlternatingColors = True
        With grProductos.RootTable.Columns("yfcdprod1")
            .Width = 150
            .Visible = False

        End With
        'b.yfcdprod1 ,a.iclot ,a.icfven  ,a.iccven 
        With grProductos.RootTable.Columns("iclot")
            .Width = 150
            .Caption = "LOTE"
            .Visible = True

        End With
        With grProductos.RootTable.Columns("icfven")
            .Width = 160
            .Caption = "FECHA VENCIMIENTO"
            .FormatString = "yyyy/MM/dd"
            .Visible = True

        End With

        With grProductos.RootTable.Columns("iccven")
            .Width = 150
            .Visible = True
            .Caption = "Stock"
            .FormatString = "0.00"
            .AggregateFunction = AggregateFunction.Sum
        End With
        With grProductos.RootTable.Columns("stockMinimo")
            .Width = 150
            .Visible = False

        End With
        With grProductos.RootTable.Columns("fechaVencimiento")
            .Width = 150
            .Visible = False

        End With

        With grProductos
            .DefaultFilterRowComparison = FilterConditionOperator.Contains
            .FilterMode = FilterMode.Automatic
            .FilterRowUpdateMode = FilterRowUpdateMode.WhenValueChanges
            .GroupByBoxVisible = False
            'diseño de la grilla
            .TotalRow = InheritableBoolean.True
            .TotalRowFormatStyle.BackColor = Color.Gold
            .TotalRowPosition = TotalRowPosition.BottomFixed
            .VisualStyle = VisualStyle.Office2007
        End With
        _prAplicarCondiccionJanusLote()

    End Sub
    Public Sub _prAplicarCondiccionJanusLote()

        Dim fc2 As GridEXFormatCondition
        fc2 = New GridEXFormatCondition(grProductos.RootTable.Columns("stockMinimo"), ConditionOperator.Equal, 1)
        fc2.FormatStyle.BackColor = Color.Red
        fc2.FormatStyle.FontBold = TriState.True
        fc2.FormatStyle.ForeColor = Color.White
        grProductos.RootTable.FormatConditions.Add(fc2)

        Dim fc As GridEXFormatCondition
        fc = New GridEXFormatCondition(grProductos.RootTable.Columns("fechaVencimiento"), ConditionOperator.Equal, 1)
        fc.FormatStyle.BackColor = Color.Gold
        fc.FormatStyle.FontBold = TriState.True
        fc.FormatStyle.ForeColor = Color.White
        grProductos.RootTable.FormatConditions.Add(fc)


    End Sub
    Private Sub _prAddDetalleVenta()
        '   a.tbnumi ,a.tbtv1numi ,a.tbty5prod ,b.yfcdprod1 as producto,a.tbest ,a.tbcmin ,a.tbumin ,Umin .ycdes3 as unidad,a.tbpbas ,a.tbptot ,a.tbobs ,
        'a.tbpcos,a.tblote ,a.tbfechaVenc , a.tbptot2, a.tbfact ,a.tbhact ,a.tbuact,1 as estado,Cast(null as Image) as img
        Dim Bin As New MemoryStream
        Dim img As New Bitmap(My.Resources.delete, 28, 28)
        img.Save(Bin, Imaging.ImageFormat.Png)
        CType(grdetalle.DataSource, DataTable).Rows.Add(_fnSiguienteNumi() + 1, 0, 0, "", 0, "", 0, 0, 0, "", 0, 0, 0, 0, 0, "", 0, "20170101", CDate("2017/01/01"), 0, Now.Date, "", "", 0, Bin.GetBuffer, 0)
    End Sub

    Public Function _fnSiguienteNumi()
        Dim dt As DataTable = CType(grdetalle.DataSource, DataTable)
        Dim rows() As DataRow = dt.Select("tbnumi=MAX(tbnumi)")
        If (rows.Count > 0) Then
            Return rows(rows.Count - 1).Item("tbnumi")
        End If
        Return 1
    End Function
    Public Function _fnAccesible()
        Return tbFechaVenta.IsInputReadOnly = False
    End Function
    Private Sub _HabilitarProductos()
        GPanelProductos.Visible = True
        PanelInferior.Visible = False
        _prCargarProductos(Str(_CodCliente))
        grProductos.Focus()
        grProductos.MoveTo(grProductos.FilterRow)
        grProductos.Col = 2
    End Sub
    Private Sub _HabilitarFocoDetalle(fila As Integer)
        _prCargarProductos(Str(_CodCliente))
        grdetalle.Focus()
        grdetalle.Row = fila
        grdetalle.Col = 2
    End Sub
    Private Sub _DesHabilitarProductos()
        GPanelProductos.Visible = False
        PanelInferior.Visible = True


        grdetalle.Select()
        grdetalle.Col = 5
        grdetalle.Row = grdetalle.RowCount - 1

    End Sub
    Public Sub _fnObtenerFilaDetalle(ByRef pos As Integer, numi As Integer)
        For i As Integer = 0 To CType(grdetalle.DataSource, DataTable).Rows.Count - 1 Step 1
            Dim _numi As Integer = CType(grdetalle.DataSource, DataTable).Rows(i).Item("tbnumi")
            If (_numi = numi) Then
                pos = i
                Return
            End If
        Next

    End Sub

    Public Sub _fnObtenerFilaDetalleProducto(ByRef pos As Integer, numi As Integer)
        For i As Integer = 0 To CType(grProductos.DataSource, DataTable).Rows.Count - 1 Step 1
            Dim _numi As Integer = CType(grProductos.DataSource, DataTable).Rows(i).Item("yfnumi")
            If (_numi = numi) Then
                pos = i
                Return
            End If
        Next

    End Sub

    Public Function _fnExisteProducto(idprod As Integer) As Boolean
        For i As Integer = 0 To CType(grdetalle.DataSource, DataTable).Rows.Count - 1 Step 1
            Dim _idprod As Integer = CType(grdetalle.DataSource, DataTable).Rows(i).Item("tbty5prod")
            Dim estado As Integer = CType(grdetalle.DataSource, DataTable).Rows(i).Item("estado")
            If (_idprod = idprod And estado >= 0) Then

                Return True
            End If
        Next
        Return False
    End Function

    Public Function _fnExisteProductoConLote(idprod As Integer, lote As String, fechaVenci As Date) As Boolean
        For i As Integer = 0 To CType(grdetalle.DataSource, DataTable).Rows.Count - 1 Step 1
            Dim _idprod As Integer = CType(grdetalle.DataSource, DataTable).Rows(i).Item("tbty5prod")
            Dim estado As Integer = CType(grdetalle.DataSource, DataTable).Rows(i).Item("estado")
            '          a.tbnumi ,a.tbtv1numi ,a.tbty5prod ,b.yfcdprod1 as producto,a.tbest ,a.tbcmin ,a.tbumin ,Umin .ycdes3 as unidad,a.tbpbas ,a.tbptot ,a.tbobs ,
            'a.tbpcos,a.tblote ,a.tbfechaVenc , a.tbptot2, a.tbfact ,a.tbhact ,a.tbuact,1 as estado,Cast(null as Image) as img,
            'Cast (0 as decimal (18,2)) as stock
            Dim _LoteDetalle As String = CType(grdetalle.DataSource, DataTable).Rows(i).Item("tblote")
            Dim _FechaVencDetalle As Date = CType(grdetalle.DataSource, DataTable).Rows(i).Item("tbfechaVenc")
            If (_idprod = idprod And estado >= 0 And lote = _LoteDetalle And fechaVenci = _FechaVencDetalle) Then

                Return True
            End If
        Next
        Return False
    End Function
    Public Sub P_PonerTotal(rowIndex As Integer)
        If (rowIndex < grdetalle.RowCount) Then
            'grdetalle.UpdateData()
            Dim lin As Integer = grdetalle.GetValue("tbnumi")
            Dim pos As Integer = -1
            _fnObtenerFilaDetalle(pos, lin)
            Dim cant As Double = grdetalle.GetValue("tbcmin")
            'Dim cantidad = Format(cant,"0.00")
            Dim uni As Double = grdetalle.GetValue("tbpbas")
            Dim cos As Double = grdetalle.GetValue("tbpcos")
            Dim MontoDesc As Double = grdetalle.GetValue("tbdesc")
            Dim dt As DataTable = CType(grdetalle.DataSource, DataTable)
            If (pos >= 0) Then
                Dim TotalUnitario As Double = cant * uni
                Dim TotalCosto As Double = cant * cos
                'grDetalle.SetValue("lcmdes", montodesc)

                CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbptot") = TotalUnitario
                CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbtotdesc") = TotalUnitario - MontoDesc

                grdetalle.SetValue("tbptot", TotalUnitario)
                grdetalle.SetValue("tbtotdesc", TotalUnitario - MontoDesc)

                CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbptot2") = TotalCosto
                grdetalle.SetValue("tbptot2", TotalCosto)

                Dim estado As Integer = CType(grdetalle.DataSource, DataTable).Rows(pos).Item("estado")
                If (estado = 1) Then
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("estado") = 2
                End If
            End If
            _prCalcularPrecioTotal()
        End If
    End Sub
    Public Sub _prCalcularPrecioTotal()
        'grdetalle.UpdateData()
        Dim montoDo As Decimal
        Dim montodesc As Double = tbMdesc.Value
        Dim pordesc As Double = ((montodesc * 100) / grdetalle.GetTotal(grdetalle.RootTable.Columns("tbtotdesc"), AggregateFunction.Sum))
        tbPdesc.Value = pordesc
        Dim total = Convert.ToDouble(Format(grdetalle.GetTotal(grdetalle.RootTable.Columns("tbtotdesc"), AggregateFunction.Sum), "0.00"))
        tbTotalBs.Text = total.ToString()
        montoDo = Convert.ToDecimal(tbTotalBs.Text) / IIf(cbCambioDolar.Text = "", 1, Convert.ToDecimal(cbCambioDolar.Text))
        tbTotalDo.Text = Format(montoDo, "0.00")
        tbIce.Value = grdetalle.GetTotal(grdetalle.RootTable.Columns("tbptot2"), AggregateFunction.Sum) * (gi_ICE / 100)
        'If (gb_FacturaIncluirICE = True) Then
        '    'tbTotal21.Value = grdetalle.GetTotal(grdetalle.RootTable.Columns("tbtotdesc"), AggregateFunction.Sum) - montodesc + tbIce.Value
        '    tbTotalBs.Text = grdetalle.GetTotal(grdetalle.RootTable.Columns("tbtotdesc"), AggregateFunction.Sum) - montodesc + tbIce.Value
        'Else
        '    'tbTotal2.Text = grdetalle.GetTotal(grdetalle.RootTable.Columns("tbtotdesc"), AggregateFunction.Sum) - montodesc
        '    'tbTotalBs.Text = Convert.ToDecimal(tbTotalBs.Text) - montodesc
        'End If
    End Sub
    Public Sub _prEliminarFila()
        If (grdetalle.Row >= 0) Then
            If (grdetalle.RowCount >= 2) Then
                Dim estado As Integer = grdetalle.GetValue("estado")
                Dim pos As Integer = -1
                Dim lin As Integer = grdetalle.GetValue("tbnumi")
                _fnObtenerFilaDetalle(pos, lin)
                If (estado = 0) Then
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("estado") = -2

                End If
                If (estado = 1) Then
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("estado") = -1
                End If
                grdetalle.RootTable.ApplyFilter(New Janus.Windows.GridEX.GridEXFilterCondition(grdetalle.RootTable.Columns("estado"), Janus.Windows.GridEX.ConditionOperator.GreaterThanOrEqualTo, 0))
                _prCalcularPrecioTotal()
                grdetalle.Select()
                grdetalle.Col = 5
                grdetalle.Row = grdetalle.RowCount - 1
            End If
        End If
        grdetalle.Refetch()
        grdetalle.Refresh()

    End Sub
    Public Function _ValidarCampos() As Boolean
        Try
            If (_CodCliente <= 0) Then
                Dim img As Bitmap = New Bitmap(My.Resources.mensaje, 50, 50)
                ToastNotification.Show(Me, "Por Favor Seleccione un Cliente con Ctrl+Enter".ToUpper, img, 2000, eToastGlowColor.Red, eToastPosition.BottomCenter)
                tbCliente.Focus()
                Return False

            End If
            If (_CodEmpleado <= 0) Then
                Dim img As Bitmap = New Bitmap(My.Resources.mensaje, 50, 50)
                ToastNotification.Show(Me, "Por Favor Seleccione un Vendedor con Ctrl+Enter".ToUpper, img, 2000, eToastGlowColor.Red, eToastPosition.BottomCenter)
                tbVendedor.Focus()
                Return False
            End If
            If (cbSucursal.SelectedIndex < 0) Then
                Dim img As Bitmap = New Bitmap(My.Resources.mensaje, 50, 50)
                ToastNotification.Show(Me, "Por Favor Seleccione una Sucursal".ToUpper, img, 2000, eToastGlowColor.Red, eToastPosition.BottomCenter)
                cbSucursal.Focus()
                Return False
            End If
            If (Convert.ToDecimal(txtMontoPagado1.Text) = 0) Then
                Throw New Exception("El monto Pagado debe ser mayor 0")
                Return False
            End If
            If (Convert.ToDecimal(txtMontoPagado1.Text) < Convert.ToDecimal(tbTotalBs.Text)) Then
                Throw New Exception("El monto Pagado debe ser mayor al monto Total")
                Return False
            End If
            'Validar datos de factura
            'If (TbNit.Text = String.Empty) Then
            '    Dim img As Bitmap = New Bitmap(My.Resources.Mensaje, 50, 50)
            '    ToastNotification.Show(Me, "Por Favor ponga el nit del cliente.".ToUpper, img, 2000, eToastGlowColor.Red, eToastPosition.BottomCenter)
            '    tbVendedor.Focus()
            '    Return False
            'End If

            'If (TbNombre1.Text = String.Empty) Then
            '    Dim img As Bitmap = New Bitmap(My.Resources.Mensaje, 50, 50)
            '    ToastNotification.Show(Me, "Por Favor ponga la razon social del cliente.".ToUpper, img, 2000, eToastGlowColor.Red, eToastPosition.BottomCenter)
            '    tbVendedor.Focus()
            '    Return False
            'End If

            If (grdetalle.RowCount = 1) Then
                grdetalle.Row = grdetalle.RowCount - 1
                If (grdetalle.GetValue("tbty5prod") = 0) Then
                    Dim img As Bitmap = New Bitmap(My.Resources.mensaje, 50, 50)
                    ToastNotification.Show(Me, "Por Favor Seleccione  un detalle de producto".ToUpper, img, 2000, eToastGlowColor.Red, eToastPosition.BottomCenter)
                    Return False
                End If

            End If
            Return True
        Catch ex As Exception
            MostrarMensajeError(ex.Message)
            Return False
        End Try

    End Function
    Private Sub _prInsertarMontoNuevo(ByRef tabla As DataTable)
        tabla.Rows.Add(0, tbMontoBs.Value, tbMontoDolar.Value, tbMontoTarej.Value, cbCambioDolar.Text, 0)
    End Sub
    Private Sub _prInsertarMontoModificar(ByRef tabla As DataTable)
        tabla.Rows.Add(tbCodigo.Text, tbMontoBs.Value, tbMontoDolar.Value, tbMontoTarej.Value, cbCambioDolar.Text, 2)
    End Sub
    Public Function rearmarDetalle() As DataTable
        Dim dt, dtDetalle, dtSaldos As DataTable
        Dim cantidadRepetido, contar, IdAux As Integer
        Dim ResultadoInventario = False

        dt = CType(grdetalle.DataSource, DataTable)
        'Ordena el detalle por codigo importante
        dt.DefaultView.Sort = "tbty5prod ASC"
        dt = dt.DefaultView.ToTable
        dtDetalle = dt.Copy
        dtDetalle.Clear()
        contar = 0
        Try
            For i As Integer = 0 To dt.Rows.Count - 1 Step 1
                Dim codProducto As Integer = dt.Rows(i).Item("tbty5prod")
                dt.DefaultView.RowFilter = "tbty5prod =  '" + codProducto.ToString() + "'"
                cantidadRepetido = dt.DefaultView.Count()
                If IdAux <> codProducto Then
                    contar = 1
                Else
                    contar += 1
                End If
                IdAux = codProducto

                'Evita llamar a saldo cada iteracion
                If contar = 1 Then
                    dtSaldos = L_fnObteniendoSaldosTI001(codProducto, 1)
                    dtSaldos.DefaultView.Sort = "icfven ASC"
                    dtSaldos = dtSaldos.DefaultView.ToTable
                End If
                'dtSaldos.DefaultView.RowFilter = "iccven >  '" + 0.ToString() + "'"
                'dtSaldos = dtSaldos.DefaultView.ToTable
                Dim cantidad As Double = dt.Rows(i).Item("tbcmin")
                Dim saldo As Double = cantidad
                Dim estado As Integer = dt.Rows(i).Item("estado")
                Dim k As Integer = 0
                If (estado >= 0) Then
                    If (dtSaldos.Rows.Count <= 0) Then
                        dtDetalle.ImportRow(dt.Rows(i))
                    Else
                        While (k <= dtSaldos.Rows.Count - 1 And saldo > 0)

                            Dim inventario As Double = dtSaldos.Rows(k).Item("iccven")
                            If (inventario >= saldo) Then
                                dtDetalle.ImportRow(dt.Rows(i))
                                Dim pos As Integer = dtDetalle.Rows.Count - 1

                                Dim precio As Double = dtDetalle.Rows(pos).Item("tbpbas")
                                Dim total As Decimal = CStr(Format(precio * saldo, "####0.00"))

                                dtDetalle.Rows(pos).Item("tbptot") = total
                                dtDetalle.Rows(pos).Item("tbtotdesc") = total
                                'CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbtotdesc") = total
                                'CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbcmin") = saldo
                                dtDetalle.Rows(pos).Item("tbcmin") = saldo

                                Dim precioCosto As Double = dtDetalle.Rows(pos).Item("tbpcos")
                                dtDetalle.Rows(pos).Item("tbptot2") = precioCosto * saldo
                                dtDetalle.Rows(pos).Item("tblote") = dtSaldos.Rows(k).Item("iclot")
                                dtDetalle.Rows(pos).Item("tbfechaVenc") = dtSaldos.Rows(k).Item("icfven")
                                dtSaldos.Rows(k).Item("iccven") = inventario - saldo
                                saldo = 0

                            Else
                                'Cuando el Invetanrio es menor
                                If (k <= dtSaldos.Rows.Count - 1 And inventario > 0) Then

                                    dtDetalle.ImportRow(dt.Rows(i))
                                    Dim pos As Integer = dtDetalle.Rows.Count - 1

                                    Dim precio As Double = dtDetalle.Rows(pos).Item("tbpbas")
                                    Dim total As Decimal = CStr(Format(precio * inventario, "####0.00"))
                                    dtDetalle.Rows(pos).Item("tbptot") = total
                                    dtDetalle.Rows(pos).Item("tbtotdesc") = total
                                    'CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbtotdesc") = total
                                    'CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbcmin") = inventario
                                    dtDetalle.Rows(pos).Item("tbcmin") = inventario

                                    Dim precioCosto As Double = dtDetalle.Rows(pos).Item("tbpcos")
                                    dtDetalle.Rows(pos).Item("tbptot2") = precioCosto * inventario
                                    dtDetalle.Rows(pos).Item("tblote") = dtSaldos.Rows(k).Item("iclot")
                                    dtDetalle.Rows(pos).Item("tbfechaVenc") = dtSaldos.Rows(k).Item("icfven")

                                    saldo = saldo - inventario
                                    'Actualiza el inventario en la Tabla
                                    dtSaldos.Rows(k).Item("iccven") = dtSaldos.Rows(k).Item("iccven") - inventario
                                End If
                            End If
                            k += 1
                        End While
                        If saldo <> 0 Then
                            dtDetalle.ImportRow(dt.Rows(i))
                            Dim pos As Integer = dtDetalle.Rows.Count - 1
                            Dim precio As Double = dtDetalle.Rows(pos).Item("tbpbas")
                            Dim total As Decimal = CStr(Format(precio * saldo, "####0.00"))
                            dtDetalle.Rows(pos).Item("tbptot") = total
                            dtDetalle.Rows(pos).Item("tbtotdesc") = total
                            dtDetalle.Rows(pos).Item("tbcmin") = saldo
                            Dim precioCosto As Double = dtDetalle.Rows(pos).Item("tbpcos")
                            dtDetalle.Rows(pos).Item("tbptot2") = precioCosto * saldo
                            dtDetalle.Rows(pos).Item("tblote") = dtSaldos.Rows(k - 1).Item("iclot")
                            dtDetalle.Rows(pos).Item("tbfechaVenc") = dtSaldos.Rows(k - 1).Item("icfven")
                            saldo = 0
                        End If
                    End If
                End If
            Next
            Return dtDetalle
        Catch ex As Exception
            MostrarMensajeError(ex.Message)
            Return dtDetalle
        End Try
    End Function
    Private Function _prExisteStockParaProducto() As Boolean
        Dim dtSaldos As DataTable
        Dim aux As Integer = 0
        Dim detalle = CType(grdetalle.DataSource, DataTable)
        Dim Lmensaje As List(Of String) = New List(Of String)
        detalle.DefaultView.RowFilter = "estado >= '" + 0.ToString() + "'"
        detalle = detalle.DefaultView.ToTable
        For Each fila As DataRow In detalle.Rows
            Dim idProducto = fila.Item("tbty5prod")
            If aux <> idProducto Then
                dtSaldos = L_fnObteniendoSaldosTI001(fila.Item("tbty5prod"), 1)
                Dim inventario = dtSaldos.Compute("SUM(iccven)", String.Empty)

                detalle.DefaultView.RowFilter = "tbty5prod = '" + fila.Item("tbty5prod").ToString() + "'"
                Dim productoRepeditos = detalle.DefaultView.ToTable
                Dim saldo = productoRepeditos.Compute("SUM(tbcmin)", String.Empty)
                'Dim saldo = fila.Item("tbcmin")
                If inventario < saldo Then
                    Dim mensaje As String = "No existe stock para el producto: " + fila.Item("producto") + " stock actual = " + inventario.ToString()
                    Lmensaje.Add(mensaje)
                    'Throw New Exception("No existe stock para el producto:" + fila.Item("producto") + " stock actual =" + inventario)
                End If
            End If
            aux = idProducto
            'dtSaldos.Select = "icfven ASC"
            'dtSaldos = dtSaldos.DefaultView.ToTable
        Next
        If Lmensaje.Count > 0 Then
            Dim mensaje = ""
            For Each item As String In Lmensaje
                mensaje = mensaje + "- " + item + vbCrLf
            Next
            MostrarMensajeError(mensaje)
            Return False
        End If
        Return True
    End Function

    Public Sub _GuardarNuevo()
        Try
            Dim numi As String = ""
            Dim tabla As DataTable = L_fnMostrarMontos(0)
            Dim factura = gb_FacturaEmite
            _prInsertarMontoNuevo(tabla)
            ''Verifica si existe estock para los productos
            'If _prExisteStockParaProducto() Then
            Dim dtDetalle As DataTable = rearmarDetalle()
                Dim res As Boolean = L_fnGrabarVenta(numi, "", tbFechaVenta.Value.ToString("yyyy/MM/dd"), _CodEmpleado, IIf(swTipoVenta.Value = True, 1, 0), IIf(swTipoVenta.Value = True, Now.Date.ToString("yyyy/MM/dd"), tbFechaVenc.Value.ToString("yyyy/MM/dd")), _CodCliente, IIf(swMoneda.Value = True, 1, 0), "", tbMdesc.Value, tbIce.Value, tbTotalBs.Text, dtDetalle, cbSucursal.Value, 0, tabla)
                If res Then
                'res = P_fnGrabarFacturarTFV001(numi)
                'Emite factura
                If (gb_FacturaEmite) Then
                    If _CodCliente <> 86 Then
                        P_fnGenerarFactura(numi)
                    Else
                        _prImiprimirNotaVenta(numi)
                    End If
                Else
                    _prImiprimirNotaVenta(numi)
                End If


                Dim img As Bitmap = New Bitmap(My.Resources.checked, 50, 50)
                    ToastNotification.Show(Me, "Código de Venta ".ToUpper + tbCodigo.Text + " Grabado con Exito.".ToUpper,
                                              img, 2000,
                                              eToastGlowColor.Green,
                                              eToastPosition.TopCenter
                                              )

                    _prCargarVenta()
                    _Limpiar()
                    Table_Producto = Nothing

                Else
                    Dim img As Bitmap = New Bitmap(My.Resources.cancel, 50, 50)
                    ToastNotification.Show(Me, "La Venta no pudo ser insertado".ToUpper, img, 2000, eToastGlowColor.Red, eToastPosition.BottomCenter)

                End If
            'End If
        Catch ex As Exception
            MostrarMensajeError(ex.Message)
        End Try


    End Sub
    Public Sub _prImiprimirNotaVenta(numi As String)
        Dim ef = New Efecto


        ef.tipo = 2
        ef.Context = "MENSAJE PRINCIPAL".ToUpper
        ef.Header = "¿desea imprimir la nota de venta?".ToUpper
        ef.ShowDialog()
        Dim bandera As Boolean = False
        bandera = ef.band
        If (bandera = True) Then
            P_GenerarReporte(numi)
        End If
    End Sub
    Public Sub _prImiprimirFacturaPreimpresa(numi As String)
        Dim ef = New Efecto


        ef.tipo = 2
        ef.Context = "MENSAJE PRINCIPAL".ToUpper
        ef.Header = "¿desea imprimir la factura Preimpresa?".ToUpper
        ef.ShowDialog()
        Dim bandera As Boolean = False
        bandera = ef.band
        If (bandera = True) Then
            P_GenerarReporteFactura(numi)
        End If
    End Sub
    Private Sub _prGuardarModificado()
        Dim tabla As DataTable = L_fnMostrarMontos(0)
        _prInsertarMontoModificar(tabla)
        Dim dtDetalle As DataTable = rearmarDetalle()
        Dim res As Boolean = L_fnModificarVenta(tbCodigo.Text, tbFechaVenta.Value.ToString("yyyy/MM/dd"), _CodEmpleado, IIf(swTipoVenta.Value = True, 1, 0), IIf(swTipoVenta.Value = True, Now.Date.ToString("yyyy/MM/dd"), tbFechaVenc.Value.ToString("yyyy/MM/dd")), _CodCliente, IIf(swMoneda.Value = True, 1, 0), "", tbMdesc.Value, tbIce.Value, tbTotalBs.Text, dtDetalle, cbSucursal.Value, 0, tabla)
        If res Then
            If (gb_FacturaEmite) Then
                L_fnEliminarDatos("TFV001", "fvanumi=" + tbCodigo.Text.Trim)
                L_fnEliminarDatos("TFV0011", "fvbnumi=" + tbCodigo.Text.Trim)
                P_fnGenerarFactura(tbCodigo.Text.Trim)
            End If
            '_prImiprimirNotaVenta(tbCodigo.Text)
            Dim img As Bitmap = New Bitmap(My.Resources.checked, 50, 50)
            ToastNotification.Show(Me, "Código de Venta ".ToUpper + tbCodigo.Text + " Modificado con Exito.".ToUpper,
                                      img, 2000,
                                      eToastGlowColor.Green,
                                      eToastPosition.TopCenter
                                      )
            _prCargarVenta()
            _prSalir()
        Else
            Dim img As Bitmap = New Bitmap(My.Resources.cancel, 50, 50)
            ToastNotification.Show(Me, "La Venta no pudo ser Modificada".ToUpper, img, 2000, eToastGlowColor.Red, eToastPosition.BottomCenter)

        End If
    End Sub
    Private Sub _prSalir()
        If btnGrabar.Enabled = True Then
            _prInhabiliitar()
            If grVentas.RowCount > 0 Then
                _prMostrarRegistro(0)
            End If
        Else
            _tab.Close()
            _modulo.Select()
        End If
    End Sub
    Public Sub _prCargarIconELiminar()
        For i As Integer = 0 To CType(grdetalle.DataSource, DataTable).Rows.Count - 1 Step 1
            Dim Bin As New MemoryStream
            Dim img As New Bitmap(My.Resources.delete, 28, 28)
            img.Save(Bin, Imaging.ImageFormat.Png)
            CType(grdetalle.DataSource, DataTable).Rows(i).Item("img") = Bin.GetBuffer
            grdetalle.RootTable.Columns("img").Visible = True
        Next

    End Sub
    Public Sub _PrimerRegistro()
        Dim _MPos As Integer
        If grVentas.RowCount > 0 Then
            _MPos = 0
            ''   _prMostrarRegistro(_MPos)
            grVentas.Row = _MPos
        End If
    End Sub
    Public Sub InsertarProductosSinLote()
        Dim pos As Integer = -1
        grdetalle.Row = grdetalle.RowCount - 1
        _fnObtenerFilaDetalle(pos, grdetalle.GetValue("tbnumi"))
        Dim existe As Boolean = _fnExisteProducto(grProductos.GetValue("yfnumi"))
        If ((pos >= 0) And (Not existe)) Then
            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbty5prod") = grProductos.GetValue("yfnumi")
            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("codigo") = grProductos.GetValue("yfcprod")
            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("yfcbarra") = grProductos.GetValue("yfcbarra")
            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("producto") = grProductos.GetValue("yfcdprod1")
            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbumin") = grProductos.GetValue("yfumin")
            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("unidad") = grProductos.GetValue("UnidMin")
            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbpbas") = grProductos.GetValue("yhprecio")
            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbptot") = grProductos.GetValue("yhprecio")
            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbtotdesc") = grProductos.GetValue("yhprecio")
            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbcmin") = 1
            If (gb_FacturaIncluirICE) Then
                CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbpcos") = grProductos.GetValue("pcos")
            Else
                CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbpcos") = 0
            End If
            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbptot2") = grProductos.GetValue("pcos")

            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("stock") = grProductos.GetValue("stock")
            _prCalcularPrecioTotal()
            _DesHabilitarProductos()
        Else
            If (existe) Then
                Dim img As Bitmap = New Bitmap(My.Resources.mensaje, 50, 50)
                ToastNotification.Show(Me, "El producto ya existe en el detalle".ToUpper, img, 2000, eToastGlowColor.Red, eToastPosition.BottomCenter)
            End If
        End If
    End Sub
    Public Sub InsertarProductosConLote()
        Dim pos As Integer = -1
        grdetalle.Row = grdetalle.RowCount - 1
        _fnObtenerFilaDetalleProducto(pos, grProductos.GetValue("yfnumi"))
        Dim posProducto As Integer = grProductos.Row
        FilaSelectLote = CType(grProductos.DataSource, DataTable).Rows(pos)


        If (grProductos.GetValue("stock") > 0) Then
            _prCargarLotesDeProductos(grProductos.GetValue("yfnumi"), grProductos.GetValue("yfcdprod1"))
        Else
            Dim img As Bitmap = New Bitmap(My.Resources.mensaje, 50, 50)
            ToastNotification.Show(Me, "El Producto: ".ToUpper + grProductos.GetValue("yfcdprod1") + " NO CUENTA CON STOCK DISPONIBLE", img, 5000, eToastGlowColor.Red, eToastPosition.BottomCenter)
            FilaSelectLote = Nothing
        End If

    End Sub
    Private Function P_fnGenerarFactura(numi As String) As Boolean
        Dim res As Boolean = False
        res = P_fnGrabarFacturarTFV001(numi) ' Grabar en la TFV001
        If (res) Then
            If (P_fnValidarFactura()) Then
                'Validar para facturar
                P_prImprimirFacturar(numi, True, True) '_Codigo de a tabla TV001
            Else
                'Volver todo al estada anterior
                ToastNotification.Show(Me, "No es posible facturar, vuelva a ingresar he intente nuevamente!!!".ToUpper,
                                       My.Resources.OK,
                                       5 * 1000,
                                       eToastGlowColor.Red,
                                       eToastPosition.MiddleCenter)
            End If

            If (Not tbNit.Text.Trim.Equals("0")) Then
                L_Grabar_Nit(tbNit.Text.Trim, TbNombre1.Text.Trim, TbNombre2.Text.Trim)
            Else
                L_Grabar_Nit(tbNit.Text, "S/N", "")
            End If
        End If

        Return res
    End Function

    Private Function P_fnGrabarFacturarTFV001(numi As String) As Boolean
        Dim a As Double = CDbl(Convert.ToDouble(tbTotalBs.Text) + tbMdesc.Value)
        'Dim b As Double = CDbl(IIf(IsDBNull(tbIce.Value), 0, tbIce.Value)) 'Ya esta calculado el 55% del ICE
        Dim b As Double = CDbl(0)
        Dim c As Double = CDbl("0")
        Dim d As Double = CDbl("0")
        Dim e As Double = a - b - c - d
        Dim f As Double = CDbl(tbMdesc.Value)
        Dim g As Double = e - f
        Dim h As Double = g * (gi_IVA / 100)

        Dim res As Boolean = False
        Dim _Hora As String = Now.Hour.ToString + ":" + Now.Minute.ToString
        'Grabado de Cabesera Factura
        L_Grabar_Factura(numi,
                        dtiFechaFactura.Value.ToString("yyyy/MM/dd"),
                        IIf(Val(tbNroFactura.Text) = 0, "0", tbNroFactura.Text),
                        IIf(Val(tbNroAutoriz.Text) = 0, "0", tbNroAutoriz.Text),
                        "1",
                        tbNit.Text.Trim,
                        _CodCliente,
                        TbNombre1.Text,
                        "",
                        CStr(Format(a, "####0.00")),
                        CStr(Format(b, "####0.00")),
                        CStr(Format(c, "####0.00")),
                        CStr(Format(d, "####0.00")),
                        CStr(Format(e, "####0.00")),
                        CStr(Format(f, "####0.00")),
                        CStr(Format(g, "####0.00")),
                        CStr(Format(h, "####0.00")),
                        "",
                        Now.Date.ToString("yyyy/MM/dd"),
                        "''",
                        cbSucursal.Value,
                        numi,
                       _Hora)

        'Grabar Nuevo y Modificado en la BDDiconDinoEco en la tabla TPA001
        If (tbCodigo.Text = String.Empty) Then
            L_Grabar_TPA001(numi, dtiFechaFactura.Value.ToString("yyyy/MM/dd"), "2", _CodCliente, TbNombre1.Text, "1",
                        "1", CStr(Format(g, "####0.00")), "1", "6.96", "0", "0")
        Else
            If (tbCodigo.Text <> String.Empty) Then
                L_Grabar_TPA001(numi, dtiFechaFactura.Value.ToString("yyyy/MM/dd"), "2", _CodCliente, TbNombre1.Text, "1",
                         "2", CStr(Format(g, "####0.00")), "1", "6.96", "0", "0")
            End If
        End If

        'Grabado de Detalle de Factura
        grProductos.Update()

        'Dim s As String = ""
        For Each fil As GridEXRow In grdetalle.GetRows
            If (Not fil.Cells("tbcmin").Value.ToString.Trim.Equals("") And
                Not fil.Cells("tbty5prod").Value.ToString.Trim.Equals("0")) Then
                's = fil.Cells("codP").Value
                's = fil.Cells("des").Value
                's = fil.Cells("can").Value
                's = fil.Cells("imp").Value
                L_Grabar_Factura_Detalle(numi.ToString,
                                        fil.Cells("tbty5prod").Value.ToString.Trim,
                                        fil.Cells("producto").Value.ToString.Trim,
                                        fil.Cells("tbcmin").Value.ToString.Trim,
                                        fil.Cells("tbpbas").Value.ToString.Trim,
                                        numi)
                res = True
            End If
        Next
        Return res
    End Function

    Private Function P_fnValidarFactura() As Boolean
        Return True
    End Function

    Private Sub P_prImprimirFacturar(numi As String, impFactura As Boolean, grabarPDF As Boolean)
        Dim _Fecha, _FechaAl As Date
        Dim _Ds, _Ds1, _Ds2, _Ds3 As New DataSet
        Dim _Autorizacion, _Nit, _Fechainv, _Total, _Key, _Cod_Control, _Hora,
            _Literal, _TotalDecimal, _TotalDecimal2 As String
        Dim I, _NumFac, _numidosif, _TotalCC As Integer
        Dim ice, _Desc, _TotalLi As Decimal
        Dim _VistaPrevia As Integer = 0


        _Desc = CDbl(tbMdesc.Value)
        If Not IsNothing(P_Global.Visualizador) Then
            P_Global.Visualizador.Close()
        End If

        '_Fecha = Now.Date '.ToString("dd/MM/yyyy")
        _Fecha = tbFechaVenta.Value
        _Hora = Now.Hour.ToString + ":" + Now.Minute.ToString
        _Ds1 = L_Dosificacion("1", "1", _Fecha)

        _Ds = L_Reporte_Factura(numi, numi)
        _Autorizacion = _Ds1.Tables(0).Rows(0).Item("sbautoriz").ToString
        _NumFac = CInt(_Ds1.Tables(0).Rows(0).Item("sbnfac")) + 1
        _Nit = _Ds.Tables(0).Rows(0).Item("fvanitcli").ToString
        _Fechainv = Microsoft.VisualBasic.Right(_Fecha.ToShortDateString, 4) +
                    Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(_Fecha.ToShortDateString, 5), 2) +
                    Microsoft.VisualBasic.Left(_Fecha.ToShortDateString, 2)
        _Total = _Ds.Tables(0).Rows(0).Item("fvatotal").ToString
        '_Total = _Ds.Tables(0).Rows(0).Item("fvastot").ToString
        ice = _Ds.Tables(0).Rows(0).Item("fvaimpsi")
        _numidosif = _Ds1.Tables(0).Rows(0).Item("sbnumi").ToString
        _Key = _Ds1.Tables(0).Rows(0).Item("sbkey")
        _FechaAl = _Ds1.Tables(0).Rows(0).Item("sbfal")

        Dim maxNFac As Integer = L_fnObtenerMaxIdTabla("TFV001", "fvanfac", "fvaautoriz = " + _Autorizacion)
        _NumFac = maxNFac + 1

        _TotalCC = Math.Round(CDbl(_Total), MidpointRounding.AwayFromZero)
        _Cod_Control = ControlCode.generateControlCode(_Autorizacion, _NumFac, _Nit, _Fechainv, CStr(_TotalCC), _Key)

        'Literal 
        _TotalLi = _Ds.Tables(0).Rows(0).Item("fvastot") - _Ds.Tables(0).Rows(0).Item("fvadesc")
        _TotalDecimal = _TotalLi - Math.Truncate(_TotalLi)
        _TotalDecimal2 = CDbl(_TotalDecimal) * 100

        'Dim li As String = Facturacion.ConvertirLiteral.A_fnConvertirLiteral(CDbl(_Total) - CDbl(_TotalDecimal)) + " con " + IIf(_TotalDecimal2.Equals("0"), "00", _TotalDecimal2) + "/100 Bolivianos"
        _Literal = Facturacion.ConvertirLiteral.A_fnConvertirLiteral(CDbl(_TotalLi) - CDbl(_TotalDecimal)) + "  " + IIf(_TotalDecimal2.Equals("0"), "00", _TotalDecimal2) + "/100 Bolivianos"
        _Ds2 = L_Reporte_Factura_Cia("2")
        QrFactura.Text = _Ds2.Tables(0).Rows(0).Item("scnit").ToString + "|" + Str(_NumFac).Trim + "|" + _Autorizacion + "|" + _Fecha + "|" + _Total + "|" + _TotalLi.ToString + "|" + _Cod_Control + "|" + tbNit.Text.Trim + "|" + ice.ToString + "|0|0|" + Str(_Desc).Trim

        L_Modificar_Factura("fvanumi = " + CStr(numi),
                            "",
                            CStr(_NumFac),
                            CStr(_Autorizacion),
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            _Cod_Control,
                            _FechaAl.ToString("yyyy/MM/dd"),
                            "",
                            "",
                            CStr(numi))

        _Ds = L_Reporte_Factura(numi, numi)

        For I = 0 To _Ds.Tables(0).Rows.Count - 1
            _Ds.Tables(0).Rows(I).Item("fvaimgqr") = P_fnImageToByteArray(QrFactura.Image)
        Next
        If (impFactura) Then
            _Ds3 = L_ObtenerRutaImpresora("1") ' Datos de Impresion de Facturación
            If (_Ds3.Tables(0).Rows(0).Item("cbvp")) Then 'Vista Previa de la Ventana de Vizualización 1 = True 0 = False
                P_Global.Visualizador = New Visualizador 'Comentar
            End If
            Dim objrep As Object = Nothing
            If (gi_FacturaTipo = 1) Then
                'objrep = New R_FacturaG
            ElseIf (gi_FacturaTipo = 2) Then
                objrep = New R_Factura_7_5x100
                'If (Not _Ds.Tables(0).Rows.Count = gi_FacturaCantidadItems) Then
                '    For index = _Ds.Tables(0).Rows.Count To gi_FacturaCantidadItems - 1
                '        'Insertamos la primera fila con el saldo Inicial
                '        Dim f As DataRow = _Ds.Tables(0).NewRow
                '        f.ItemArray() = _Ds.Tables(0).Rows(0).ItemArray
                '        f.Item("fvbcant") = -1
                '        _Ds.Tables(0).Rows.Add(f)
                '    Next
                'End If
            End If

            objrep.SetDataSource(_Ds.Tables(0))
            objrep.SetParameterValue("Hora", _Hora)
            objrep.SetParameterValue("Direccionpr", _Ds2.Tables(0).Rows(0).Item("scdir").ToString)
            objrep.SetParameterValue("Telefonopr", _Ds2.Tables(0).Rows(0).Item("sctelf").ToString)
            objrep.SetParameterValue("Literal1", _Literal)
            objrep.SetParameterValue("Literal2", " ")
            objrep.SetParameterValue("Literal3", " ")
            objrep.SetParameterValue("NroFactura", _NumFac)
            objrep.SetParameterValue("NroAutoriz", _Autorizacion)
            objrep.SetParameterValue("ENombre", _Ds2.Tables(0).Rows(0).Item("scneg").ToString) '?
            objrep.SetParameterValue("ECasaMatriz", _Ds2.Tables(0).Rows(0).Item("scsuc").ToString)
            objrep.SetParameterValue("ECiudadPais", _Ds2.Tables(0).Rows(0).Item("scpai").ToString)
            objrep.SetParameterValue("ESFC", _Ds1.Tables(0).Rows(0).Item("sbsfc").ToString)
            objrep.SetParameterValue("ENit", _Ds2.Tables(0).Rows(0).Item("scnit").ToString)
            objrep.SetParameterValue("EActividad", _Ds2.Tables(0).Rows(0).Item("scact").ToString)
            ' objrep.SetParameterValue("ESms", "''" + _Ds1.Tables(0).Rows(0).Item("sbnota").ToString + "''")
            'objrep.SetParameterValue("ESms2", "''" + _Ds1.Tables(0).Rows(0).Item("sbnota2").ToString + "''")
            objrep.SetParameterValue("EDuenho", _Ds2.Tables(0).Rows(0).Item("scnom").ToString) '?
            'objrep.SetParameterValue("URLImageLogo", gs_CarpetaRaiz + "\LogoFactura.jpg")
            'objrep.SetParameterValue("URLImageMarcaAgua", gs_CarpetaRaiz + "\MarcaAguaFactura.jpg")
            objrep.SetParameterValue("ENota", "''" + "ESTA FACTURA CONTRIBUYE AL DESARROLLO DEL PAÍS EL USO ILÍCITO DE ÉSTA SERÁ SANCIONADO DE ACUERDO A LA LEY" + "''")
            objrep.SetParameterValue("ELey", "''" + "Ley N° 453: " + "El proveedor debe brindar atención sin discriminación, con respeto, calidez y cordialidad a los usuarios y consumidores." + "''")

            If (_Ds3.Tables(0).Rows(0).Item("cbvp")) Then 'Vista Previa de la Ventana de Vizualización 1 = True 0 = False
                P_Global.Visualizador.CrGeneral.ReportSource = objrep 'Comentar
                P_Global.Visualizador.ShowDialog() 'Comentar
                P_Global.Visualizador.BringToFront() 'Comentar
            End If

            Dim pd As New PrintDocument()
            pd.PrinterSettings.PrinterName = _Ds3.Tables(0).Rows(0).Item("cbrut").ToString
            If (Not pd.PrinterSettings.IsValid) Then
                ToastNotification.Show(Me, "La Impresora ".ToUpper + _Ds3.Tables(0).Rows(0).Item("cbrut").ToString + Chr(13) + "No Existe".ToUpper,
                                       My.Resources.WARNING, 5 * 1000,
                                       eToastGlowColor.Blue, eToastPosition.BottomRight)
            Else
                objrep.PrintOptions.PrinterName = _Ds3.Tables(0).Rows(0).Item("cbrut").ToString '"EPSON TM-T20II Receipt5 (1)"
                objrep.PrintToPrinter(1, True, 0, 0)
                'crystalReportDocument.PrintOptions.PrinterName = "your printer name"
                'objrep.PrintTicket("EPSON TM-T20II Receipt")
            End If

            If (grabarPDF) Then
                'Copia de Factura en PDF
                If (Not Directory.Exists(gs_CarpetaRaiz + "\Facturas")) Then
                    Directory.CreateDirectory(gs_CarpetaRaiz + "\Facturas")
                End If
                objrep.ExportToDisk(ExportFormatType.PortableDocFormat, gs_CarpetaRaiz + "\Facturas\" + CStr(_NumFac) + "_" + CStr(_Autorizacion) + ".pdf")

            End If
        End If
                L_Actualiza_Dosificacion(_numidosif, _NumFac, numi)
    End Sub
    Private Sub ReimprimirFactura(numi As String, impFactura As Boolean, grabarPDF As Boolean)
        Try
            Dim _Fecha, _FechaAl As Date
            Dim _Ds, _Ds1, _Ds2, _Ds3 As New DataSet
            Dim _Autorizacion, _Nit, _Fechainv, _Total, _Key, _Cod_Control, _Hora,
                _Literal, _TotalDecimal, _TotalDecimal2 As String
            Dim I, _NumFac, _numidosif, _TotalCC As Integer
            Dim ice, _Desc, _TotalLi As Decimal
            Dim _VistaPrevia As Integer = 0


            _Desc = CDbl(tbMdesc.Value)
            If Not IsNothing(P_Global.Visualizador) Then
                P_Global.Visualizador.Close()
            End If

            _Ds1 = L_Dosificacion("1", "1", _Fecha)
            _Ds = L_Reporte_Factura(numi, numi)
            _Fecha = _Ds.Tables(0).Rows(0).Item("fvafec").ToString
            _Hora = _Ds.Tables(0).Rows(0).Item("fvahora").ToString
            _Autorizacion = _Ds1.Tables(0).Rows(0).Item("sbautoriz").ToString
            _NumFac = CInt(_Ds.Tables(0).Rows(0).Item("fvanfac").ToString)
            '_NumFac = CInt(_Ds1.Tables(0).Rows(0).Item("sbnfac"))
            _Nit = _Ds.Tables(0).Rows(0).Item("fvanitcli").ToString

            _Fechainv = Microsoft.VisualBasic.Right(_Fecha.ToShortDateString, 4) +
                        Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(_Fecha.ToShortDateString, 5), 2) +
                        Microsoft.VisualBasic.Left(_Fecha.ToShortDateString, 2)
            _Total = _Ds.Tables(0).Rows(0).Item("fvatotal").ToString
            '_Total = "45.82"
            ' _Total = _Ds.Tables(0).Rows(0).Item("fvastot").ToString
            ice = _Ds.Tables(0).Rows(0).Item("fvaimpsi")
            _numidosif = _Ds1.Tables(0).Rows(0).Item("sbnumi").ToString
            _Key = _Ds1.Tables(0).Rows(0).Item("sbkey")
            _FechaAl = _Ds1.Tables(0).Rows(0).Item("sbfal")

            'Dim maxNFac As Integer = L_fnObtenerMaxIdTabla("TFV001", "fvanfac", "fvaautoriz = " + _Autorizacion)
            '_NumFac = maxNFac + 1

            _TotalCC = Math.Round(CDbl(_Total), MidpointRounding.AwayFromZero)
            _Cod_Control = ControlCode.generateControlCode(_Autorizacion, _NumFac, _Nit, _Fechainv, CStr(_TotalCC), _Key)

            'Literal 
            '_TotalLi = "2740.38"
            _TotalLi = _Ds.Tables(0).Rows(0).Item("fvastot") - _Ds.Tables(0).Rows(0).Item("fvadesc")
            _TotalDecimal = _TotalLi - Math.Truncate(_TotalLi)
            '_TotalDecimal = "0.93"
            _TotalDecimal2 = CDbl(_TotalDecimal) * 100

            'Dim li As String = Facturacion.ConvertirLiteral.A_fnConvertirLiteral(CDbl(_Total) - CDbl(_TotalDecimal)) + " con " + IIf(_TotalDecimal2.Equals("0"), "00", _TotalDecimal2) + "/100 Bolivianos"
            _Literal = Facturacion.ConvertirLiteral.A_fnConvertirLiteral(CDbl(_TotalLi) - CDbl(_TotalDecimal)) + "  " + IIf(_TotalDecimal2.Equals("0"), "00", _TotalDecimal2) + "/100 Bolivianos"
            _Ds2 = L_Reporte_Factura_Cia("2")
            QrFactura.Text = _Ds2.Tables(0).Rows(0).Item("scnit").ToString + "|" + Str(_NumFac).Trim + "|" + _Autorizacion + "|" + _Fecha + "|" + _Total + "|" + _TotalLi.ToString + "|" + _Cod_Control + "|" + tbNit.Text.Trim + "|" + ice.ToString + "|0|0|" + Str(_Desc).Trim

            L_Modificar_Factura("fvanumi = " + CStr(numi),
                                "",
                                CStr(_NumFac),
                                CStr(_Autorizacion),
                                "",
                                "",
                                "",
                                "",
                                "",
                                "",
                                "",
                                "",
                                "",
                                "",
                                "",
                                "",
                                "",
                                _Cod_Control,
                                _FechaAl.ToString("yyyy/MM/dd"),
                                "",
                                "",
                                CStr(numi))

            _Ds = L_Reporte_Factura(numi, numi)

            For I = 0 To _Ds.Tables(0).Rows.Count - 1
                _Ds.Tables(0).Rows(I).Item("fvaimgqr") = P_fnImageToByteArray(QrFactura.Image)
            Next
            If (impFactura) Then
                _Ds3 = L_ObtenerRutaImpresora("1") ' Datos de Impresion de Facturación
                If (_Ds3.Tables(0).Rows(0).Item("cbvp")) Then 'Vista Previa de la Ventana de Vizualización 1 = True 0 = False
                    P_Global.Visualizador = New Visualizador 'Comentar
                End If
                Dim objrep As Object = Nothing
                If (gi_FacturaTipo = 1) Then
                    'objrep = New R_FacturaG
                ElseIf (gi_FacturaTipo = 2) Then
                    objrep = New R_Factura_7_5x100
                    'If (Not _Ds.Tables(0).Rows.Count = gi_FacturaCantidadItems) Then
                    '    For index = _Ds.Tables(0).Rows.Count To gi_FacturaCantidadItems - 1
                    '        'Insertamos la primera fila con el saldo Inicial
                    '        Dim f As DataRow = _Ds.Tables(0).NewRow
                    '        f.ItemArray() = _Ds.Tables(0).Rows(0).ItemArray
                    '        f.Item("fvbcant") = -1
                    '        _Ds.Tables(0).Rows.Add(f)
                    '    Next
                    'End If
                End If

                objrep.SetDataSource(_Ds.Tables(0))
                objrep.SetParameterValue("Hora", _Hora)
                objrep.SetParameterValue("Direccionpr", _Ds2.Tables(0).Rows(0).Item("scdir").ToString)
                objrep.SetParameterValue("Telefonopr", _Ds2.Tables(0).Rows(0).Item("sctelf").ToString)
                objrep.SetParameterValue("Literal1", _Literal)
                objrep.SetParameterValue("Literal2", " ")
                objrep.SetParameterValue("Literal3", " ")
                objrep.SetParameterValue("NroFactura", _NumFac)
                objrep.SetParameterValue("NroAutoriz", _Autorizacion)
                objrep.SetParameterValue("ENombre", _Ds2.Tables(0).Rows(0).Item("scneg").ToString) '?
                objrep.SetParameterValue("ECasaMatriz", _Ds2.Tables(0).Rows(0).Item("scsuc").ToString)
                objrep.SetParameterValue("ECiudadPais", _Ds2.Tables(0).Rows(0).Item("scpai").ToString)
                objrep.SetParameterValue("ESFC", _Ds1.Tables(0).Rows(0).Item("sbsfc").ToString)
                objrep.SetParameterValue("ENit", _Ds2.Tables(0).Rows(0).Item("scnit").ToString)
                objrep.SetParameterValue("EActividad", _Ds2.Tables(0).Rows(0).Item("scact").ToString)
                ' objrep.SetParameterValue("ESms", "''" + _Ds1.Tables(0).Rows(0).Item("sbnota").ToString + "''")
                'objrep.SetParameterValue("ESms2", "''" + _Ds1.Tables(0).Rows(0).Item("sbnota2").ToString + "''")
                objrep.SetParameterValue("EDuenho", _Ds2.Tables(0).Rows(0).Item("scnom").ToString) '?
                'objrep.SetParameterValue("URLImageLogo", gs_CarpetaRaiz + "\LogoFactura.jpg")
                'objrep.SetParameterValue("URLImageMarcaAgua", gs_CarpetaRaiz + "\MarcaAguaFactura.jpg")
                objrep.SetParameterValue("ENota", "''" + "ESTA FACTURA CONTRIBUYE AL DESARROLLO DEL PAÍS EL USO ILÍCITO DE ÉSTA SERÁ SANCIONADO DE ACUERDO A LA LEY" + "''")
                objrep.SetParameterValue("ELey", "''" + "Ley N° 453: " + "El proveedor debe brindar atención sin discriminación, con respeto, calidez y cordialidad a los usuarios y consumidores." + "''")

                If (_Ds3.Tables(0).Rows(0).Item("cbvp")) Then 'Vista Previa de la Ventana de Vizualización 1 = True 0 = False
                    P_Global.Visualizador.CrGeneral.ReportSource = objrep 'Comentar
                    P_Global.Visualizador.ShowDialog() 'Comentar
                    P_Global.Visualizador.BringToFront() 'Comentar
                Else
                    Dim pd As New PrintDocument()
                    pd.PrinterSettings.PrinterName = _Ds3.Tables(0).Rows(0).Item("cbrut").ToString
                    If (Not pd.PrinterSettings.IsValid) Then
                        ToastNotification.Show(Me, "La Impresora ".ToUpper + _Ds3.Tables(0).Rows(0).Item("cbrut").ToString + Chr(13) + "No Existe".ToUpper,
                                               My.Resources.WARNING, 5 * 1000,
                                               eToastGlowColor.Blue, eToastPosition.BottomRight)
                    Else
                        objrep.PrintOptions.PrinterName = _Ds3.Tables(0).Rows(0).Item("cbrut").ToString
                        '"EPSON TM-U330 Receipt"
                        objrep.PrintToPrinter(1, True, 0, 0)
                        'crystalReportDocument.PrintOptions.PrinterName = "your printer name"
                        'objrep.PrintTicket("EPSON TM-T20II Receipt")
                    End If
                End If
                'If (grabarPDF) Then
                '    'Copia de Factura en PDF
                '    If (Not Directory.Exists(gs_CarpetaRaiz + "\Facturas")) Then
                '        Directory.CreateDirectory(gs_CarpetaRaiz + "\Facturas")
                '    End If
                '    objrep.ExportToDisk(ExportFormatType.PortableDocFormat, gs_CarpetaRaiz + "\Facturas\" + CStr(_NumFac) + "_" + CStr(_Autorizacion) + ".pdf")
                'End If
            End If
            '  L_Actualiza_Dosificacion(_numidosif, _NumFac, numi)
        Catch ex As Exception
            MostrarMensajeError(ex.Message)
        End Try
    End Sub

    Public Function P_fnImageToByteArray(ByVal imageIn As Image) As Byte()
        Dim ms As New System.IO.MemoryStream()
        imageIn.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
        Return ms.ToArray()
    End Function


    Private Function P_fnValidarFacturaVigente() As Boolean
        Dim est As String = L_fnObtenerDatoTabla("TFV001", "fvaest", "fvanumi=" + tbCodigo.Text.Trim)
        Return (est.Equals("True"))
    End Function

    Private Sub P_prCargarVariablesIndispensables()
        If (gb_FacturaEmite) Then
            gi_IVA = CDbl(IIf(L_fnGetIVA().Rows(0).Item("scdebfis").ToString.Equals(""), gi_IVA, L_fnGetIVA().Rows(0).Item("scdebfis").ToString))
            gi_ICE = CDbl(IIf(L_fnGetICE().Rows(0).Item("scice").ToString.Equals(""), gi_ICE, L_fnGetICE().Rows(0).Item("scice").ToString))
        End If

    End Sub

    Private Sub P_prCargarParametro()
        'El sistema factura?

        'Si factura, preguntar si, Se incluye el Importe ICE / IEHD / TASAS?
        If (gb_FacturaEmite) Then
            lbIce.Visible = gb_FacturaIncluirICE
            tbIce.Visible = gb_FacturaIncluirICE
        Else
            lbIce.Visible = False
            tbIce.Visible = False
        End If

    End Sub
    Private Sub P_GenerarReporte(numi As String)
        Dim dt As DataTable = L_fnVentaNotaDeVenta(numi)
        If (gb_DetalleProducto) Then
            ponerDescripcionProducto(dt)
        End If
        Dim total As Decimal = dt.Compute("SUM(Total)", "")
        Dim totald As Double = (total / 6.96)
        Dim fechaven As String = dt.Rows(0).Item("fechaventa")
        If Not IsNothing(P_Global.Visualizador) Then
            P_Global.Visualizador.Close()
        End If
        Dim ParteEntera As Long
        Dim ParteDecimal As Decimal
        ParteEntera = Int(total)
        ParteDecimal = Math.Round(total - ParteEntera, 2)
        Dim li As String = Facturacion.ConvertirLiteral.A_fnConvertirLiteral(CDbl(ParteEntera)) + " con " +
        IIf(ParteDecimal.ToString.Equals("0"), "00", ParteDecimal.ToString) + "/100 Bolivianos"

        ParteEntera = Int(totald)
        ParteDecimal = Math.Round(totald - ParteEntera, 2)
        Dim lid As String = Facturacion.ConvertirLiteral.A_fnConvertirLiteral(CDbl(ParteEntera)) + " con " +
        IIf(ParteDecimal.ToString.Equals("0"), "00", ParteDecimal.ToString) + "/100 Dolares"
        Dim _Hora As String = Now.Hour.ToString + ":" + Now.Minute.ToString
        Dim _Ds2 = L_Reporte_Factura_Cia("2")
        Dim dt2 As DataTable = L_fnNameReporte()
        Dim ParEmp1 As String = ""
        Dim ParEmp2 As String = ""
        Dim ParEmp3 As String = ""
        Dim ParEmp4 As String = ""
        If (dt2.Rows.Count > 0) Then
            ParEmp1 = dt2.Rows(0).Item("Empresa1").ToString
            ParEmp2 = dt2.Rows(0).Item("Empresa2").ToString
            ParEmp3 = dt2.Rows(0).Item("Empresa3").ToString
            ParEmp4 = dt2.Rows(0).Item("Empresa4").ToString
        End If

        Dim _Ds3 = L_ObtenerRutaImpresora("1") ' Datos de Impresion de Facturación
        If (_Ds3.Tables(0).Rows(0).Item("cbvp")) Then 'Vista Previa de la Ventana de Vizualización 1 = True 0 = False
            P_Global.Visualizador = New Visualizador 'Comentar
        End If
        Dim _FechaAct As String
        Dim _FechaPar As String
        Dim _Fecha() As String
        Dim _Meses() As String = {"Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"}
        _FechaAct = fechaven
        _Fecha = Split(_FechaAct, "-")
        _FechaPar = "Cochabamba, " + _Fecha(0).Trim + " De " + _Meses(_Fecha(1) - 1).Trim + " Del " + _Fecha(2).Trim

        Dim objrep As New R_NotaVenta_7_5X100
        objrep.SetDataSource(dt)
            objrep.SetParameterValue("Literal1", li)
        'objrep.SetParameterValue("usuario", gs_user)
        objrep.SetParameterValue("ECasaMatriz", _Ds2.Tables(0).Rows(0).Item("scsuc").ToString)
            objrep.SetParameterValue("ECiudadPais", _Ds2.Tables(0).Rows(0).Item("scpai").ToString)
            objrep.SetParameterValue("EDuenho", _Ds2.Tables(0).Rows(0).Item("scnom").ToString) '?
            objrep.SetParameterValue("Direccionpr", _Ds2.Tables(0).Rows(0).Item("scdir").ToString)
        objrep.SetParameterValue("Hora", _Hora)
        objrep.SetParameterValue("ENombre", _Ds2.Tables(0).Rows(0).Item("scneg").ToString) '?

        If (_Ds3.Tables(0).Rows(0).Item("cbvp")) Then 'Vista Previa de la Ventana de Vizualización 1 = True 0 = False
            P_Global.Visualizador.CrGeneral.ReportSource = objrep 'Comentar
            P_Global.Visualizador.ShowDialog() 'Comentar
            P_Global.Visualizador.BringToFront() 'Comentar
        Else
            Dim pd As New PrintDocument()
            pd.PrinterSettings.PrinterName = _Ds3.Tables(0).Rows(0).Item("cbrut").ToString
            If (Not pd.PrinterSettings.IsValid) Then
                ToastNotification.Show(Me, "La Impresora ".ToUpper + _Ds3.Tables(0).Rows(0).Item("cbrut").ToString + Chr(13) + "No Existe".ToUpper,
                                       My.Resources.WARNING, 5 * 1000,
                                       eToastGlowColor.Blue, eToastPosition.BottomRight)
            Else
                objrep.PrintOptions.PrinterName = _Ds3.Tables(0).Rows(0).Item("cbrut").ToString
                objrep.PrintToPrinter(1, True, 0, 0)
            End If
        End If
    End Sub

    Private Sub ponerDescripcionProducto(ByRef dt As DataTable)
        For Each fila As DataRow In dt.Rows
            Dim numi As Integer = fila.Item("codProducto")
            Dim dtDP As DataTable = L_fnDetalleProducto(numi)
            Dim des As String = fila.Item("producto") + vbNewLine + vbNewLine
            For Each fila2 As DataRow In dtDP.Rows
                des = des + fila2.Item("yfadesc").ToString + vbNewLine
            Next
            fila.Item("producto") = des
        Next
    End Sub

    Private Sub P_GenerarReporteFactura(numi As String)
        Dim dt As DataTable = L_fnVentaFactura(numi)
        Dim total As Double = dt.Compute("SUM(Total)", "")

        Dim ParteEntera As Long
        Dim ParteDecimal As Double
        ParteEntera = Int(total)
        ParteDecimal = total - ParteEntera
        Dim li As String = Facturacion.ConvertirLiteral.A_fnConvertirLiteral(CDbl(ParteEntera)) + " con " +
        IIf(ParteDecimal.ToString.Equals("0"), "00", ParteDecimal.ToString) + "/100 Bolivianos"


        Dim objrep As New R_FacturaFarmacia
        '' GenerarNro(_dt)
        ''objrep.SetDataSource(Dt1Kardex)
        'imprimir
        If PrintDialog1.ShowDialog = DialogResult.OK Then
            objrep.SetDataSource(dt)
            objrep.SetParameterValue("TotalEscrito", li)
            objrep.SetParameterValue("nit", tbNit.Text)
            objrep.SetParameterValue("Total", total)
            objrep.SetParameterValue("cliente", TbNombre1.Text + " " + TbNombre2.Text)
            objrep.PrintOptions.PrinterName = PrintDialog1.PrinterSettings.PrinterName

            objrep.PrintToPrinter(1, False, 1, 1)
            objrep.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.DefaultPaperSize
        End If
        'objrep.SetDataSource(dt)
        'objrep.SetParameterValue("TotalEscrito", li)
        'objrep.SetParameterValue("nit", TbNit.Text)
        'objrep.SetParameterValue("Total", total)
        'P_Global.Visualizador.CrGeneral.ReportSource = objrep 'Comentar
        'P_Global.Visualizador.Show() 'Comentar
        'P_Global.Visualizador.BringToFront() 'Comentar



    End Sub

    Sub _prCargarProductoDeLaProforma(numiProforma As Integer)
        Dim dt As DataTable = L_fnListarProductoProforma(numiProforma)

        '        a.pbnumi ,a.pbtp1numi ,a.pbty5prod ,producto ,a.pbcmin ,a.pbumin ,Umin .ycdes3 as unidad,
        'a.pbpbas ,a.pbptot,a.pbporc,a.pbdesc ,a.pbtotdesc,
        'stock, pcosto
        If (dt.Rows.Count > 0) Then
            CType(grdetalle.DataSource, DataTable).Rows.Clear()
            For i As Integer = 0 To dt.Rows.Count - 1
                Dim numiproducto As Integer = dt.Rows(i).Item("pbty5prod")
                Dim nameproducto As String = dt.Rows(i).Item("producto")
                Dim lote As String = ""
                Dim FechaVenc As Date = Now.Date
                Dim cant As Double = dt.Rows(i).Item("pbcmin")
                Dim iccven As Double = 0
                _prPedirLotesProducto(lote, FechaVenc, iccven, numiproducto, nameproducto, cant)
                _prAddDetalleVenta()
                grdetalle.Row = grdetalle.RowCount - 1
                If (lote <> String.Empty) Then
                    If (cant <= iccven) Then

                        grdetalle.SetValue("tbptot", dt.Rows(i).Item("pbptot"))
                        grdetalle.SetValue("tbtotdesc", dt.Rows(i).Item("pbtotdesc"))
                        grdetalle.SetValue("tbdesc", dt.Rows(i).Item("pbdesc"))
                        grdetalle.SetValue("tbcmin", cant)
                        grdetalle.SetValue("tbptot2", dt.Rows(i).Item("pcosto") * cant)

                    Else
                        Dim tot As Double = dt.Rows(i).Item("pbpbas") * iccven
                        grdetalle.SetValue("tbptot", tot)
                        grdetalle.SetValue("tbtotdesc", tot)
                        grdetalle.SetValue("tbdesc", 0)
                        grdetalle.SetValue("tbcmin", iccven)
                        grdetalle.SetValue("tbptot2", dt.Rows(i).Item("pcosto") * iccven)
                    End If
                    grdetalle.SetValue("tbty5prod", numiproducto)
                    grdetalle.SetValue("producto", nameproducto)
                    grdetalle.SetValue("tbumin", dt.Rows(i).Item("pbumin"))
                    grdetalle.SetValue("unidad", dt.Rows(i).Item("unidad"))
                    grdetalle.SetValue("tbpbas", dt.Rows(i).Item("pbpbas"))


                    If (gb_FacturaIncluirICE) Then
                        grdetalle.SetValue("tbpcos", dt.Rows(i).Item("pcosto"))

                    Else
                        grdetalle.SetValue("tbpcos", 0)
                    End If

                    grdetalle.SetValue("tblote", lote)
                    grdetalle.SetValue("tbfechaVenc", FechaVenc)
                    grdetalle.SetValue("stock", iccven)

                End If

                'grdetalle.Refetch()
                'grdetalle.Refresh()


            Next

            grdetalle.Select()
            _prCalcularPrecioTotal()
        End If
    End Sub
    Public Sub _prPedirLotesProducto(ByRef lote As String, ByRef FechaVenc As Date, ByRef iccven As Double, CodProducto As Integer, nameProducto As String, cant As Integer)
        Dim dt As New DataTable
        dt = L_fnListarLotesPorProductoVenta(cbSucursal.Value, CodProducto)  ''1=Almacen
        'b.yfcdprod1 ,a.iclot ,a.icfven  ,a.iccven 
        Dim listEstCeldas As New List(Of Modelo.Celda)
        listEstCeldas.Add(New Modelo.Celda("yfcdprod1,", False, "", 150))
        listEstCeldas.Add(New Modelo.Celda("iclot", True, "LOTE", 150))
        listEstCeldas.Add(New Modelo.Celda("icfven", True, "FECHA VENCIMIENTO", 180, "dd/MM/yyyy"))
        listEstCeldas.Add(New Modelo.Celda("iccven", True, "Stock".ToUpper, 150, "0.00"))
        Dim ef = New Efecto
        ef.tipo = 3
        ef.dt = dt
        ef.SeleclCol = 2
        ef.listEstCeldas = listEstCeldas
        ef.alto = 50
        ef.ancho = 350
        ef.Context = "Producto ".ToUpper + nameProducto + "  cantidad=" + Str(cant)
        ef.ShowDialog()
        Dim bandera As Boolean = False
        bandera = ef.band
        'b.yfcdprod1 ,a.iclot ,a.icfven  ,a.iccven 
        If (bandera = True) Then
            Dim Row As Janus.Windows.GridEX.GridEXRow = ef.Row
            lote = Row.Cells("iclot").Value
            FechaVenc = Row.Cells("icfven").Value
            iccven = Row.Cells("iccven").Value
        End If
    End Sub
    Private Sub AsignarClienteEmpleado()
        Dim _tabla11 As DataTable = L_fnListarClientesUsuario(gi_userNumi)
        If _tabla11.Rows.Count > 0 Then
            tbCliente.Text = _tabla11.Rows(0).Item("yddesc")
            _CodCliente = _tabla11.Rows(0).Item("ydnumi") 'Codigo
            tbVendedor.Text = _tabla11.Rows(0).Item("vendedor") 'Codigo
            _CodEmpleado = _tabla11.Rows(0).Item("ydnumivend") 'Codigo
        Else
            Dim dt As DataTable
            dt = L_fnListarClientes()
            If dt.Rows.Count > 0 Then
                Dim fila As DataRow() = dt.Select("ydnumi =MIN(ydnumi)")
                tbCliente.Text = fila(0).ItemArray(3)
                _CodCliente = fila(0).ItemArray(0)
                tbVendedor.Text = fila(0).ItemArray(9)
                _CodEmpleado = fila(0).ItemArray(8)
            End If
        End If
    End Sub
    Private Sub MostrarMensajeError(mensaje As String)
        ToastNotification.Show(Me,
                               mensaje.ToUpper,
                               My.Resources.WARNING,
                               5000,
                               eToastGlowColor.Red,
                               eToastPosition.TopCenter)

    End Sub
    Private Sub MostrarMensajeOk(mensaje As String)
        ToastNotification.Show(Me,
                               mensaje.ToUpper,
                               My.Resources.OK,
                               5000,
                               eToastGlowColor.Green,
                               eToastPosition.TopCenter)
    End Sub
#End Region

#Region "Eventos Formulario"
    Private Sub F0_Ventas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        _IniciarTodo()
    End Sub
    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        _Limpiar()
        _prhabilitar()
        AsignarClienteEmpleado()
        btnNuevo.Enabled = False
        btnModificar.Enabled = False
        btnEliminar.Enabled = False
        btnGrabar.Enabled = True
        PanelNavegacion.Enabled = False
        tbNit.Select()
        'btnNuevo.Enabled = False
        'btnModificar.Enabled = False
        'btnEliminar.Enabled = False
        'GPanelProductos.Visible = False
        '_prhabilitar()

        '_Limpiar()
    End Sub
    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        _prSalir()

    End Sub

    Private Sub tbCliente_KeyDown(sender As Object, e As KeyEventArgs) Handles tbCliente.KeyDown
        If (_fnAccesible()) Then
            If e.KeyData = Keys.Control + Keys.Enter Then

                Dim dt As DataTable

                dt = L_fnListarClientes()
                '              a.ydnumi, a.ydcod, a.yddesc, a.yddctnum, a.yddirec
                ',a.ydtelf1 ,a.ydfnac 

                Dim listEstCeldas As New List(Of Modelo.Celda)
                listEstCeldas.Add(New Modelo.Celda("ydnumi,", True, "ID", 50))
                listEstCeldas.Add(New Modelo.Celda("ydcod", False, "ID", 50))
                listEstCeldas.Add(New Modelo.Celda("ydrazonsocial", True, "RAZÓN SOCIAL", 180))
                listEstCeldas.Add(New Modelo.Celda("yddesc", True, "NOMBRE", 280))
                listEstCeldas.Add(New Modelo.Celda("yddctnum", True, "N. Documento".ToUpper, 150))
                listEstCeldas.Add(New Modelo.Celda("yddirec", True, "DIRECCIÓN", 220))
                listEstCeldas.Add(New Modelo.Celda("ydtelf1", True, "Teléfono".ToUpper, 200))
                listEstCeldas.Add(New Modelo.Celda("ydfnac", True, "F.Nacimiento".ToUpper, 150, "MM/dd,YYYY"))
                listEstCeldas.Add(New Modelo.Celda("ydnumivend,", False, "ID", 50))
                listEstCeldas.Add(New Modelo.Celda("vendedor,", False, "ID", 50))
                listEstCeldas.Add(New Modelo.Celda("yddias", False, "CRED", 50))
                Dim ef = New Efecto
                ef.tipo = 3
                ef.dt = dt
                ef.SeleclCol = 2
                ef.listEstCeldas = listEstCeldas
                ef.alto = 50
                ef.ancho = 350
                ef.Context = "Seleccione Cliente".ToUpper
                ef.ShowDialog()
                Dim bandera As Boolean = False
                bandera = ef.band
                If (bandera = True) Then
                    Dim Row As Janus.Windows.GridEX.GridEXRow = ef.Row

                    _CodCliente = Row.Cells("ydnumi").Value
                    tbCliente.Text = Row.Cells("ydrazonsocial").Value
                    _dias = Row.Cells("yddias").Value

                    Dim numiVendedor As Integer = IIf(IsDBNull(Row.Cells("ydnumivend").Value), 0, Row.Cells("ydnumivend").Value)
                    If (numiVendedor > 0) Then
                        tbVendedor.Text = Row.Cells("vendedor").Value
                        _CodEmpleado = Row.Cells("ydnumivend").Value

                        grdetalle.Select()
                        Table_Producto = Nothing
                    Else
                        tbVendedor.Clear()
                        _CodEmpleado = 0
                        tbVendedor.Focus()
                        Table_Producto = Nothing

                    End If
                End If
            End If
        End If
    End Sub
    Private Sub tbVendedor_KeyDown(sender As Object, e As KeyEventArgs) Handles tbVendedor.KeyDown

        If (_fnAccesible()) Then
            If e.KeyData = Keys.Control + Keys.Enter Then

                Dim dt As DataTable

                dt = L_fnListarEmpleado()
                '              a.ydnumi, a.ydcod, a.yddesc, a.yddctnum, a.yddirec
                ',a.ydtelf1 ,a.ydfnac 

                Dim listEstCeldas As New List(Of Modelo.Celda)
                listEstCeldas.Add(New Modelo.Celda("ydnumi,", False, "ID", 50))
                listEstCeldas.Add(New Modelo.Celda("ydcod", True, "ID", 50))
                listEstCeldas.Add(New Modelo.Celda("yddesc", True, "NOMBRE", 280))
                listEstCeldas.Add(New Modelo.Celda("yddctnum", True, "N. Documento".ToUpper, 150))
                listEstCeldas.Add(New Modelo.Celda("yddirec", True, "DIRECCION", 220))
                listEstCeldas.Add(New Modelo.Celda("ydtelf1", True, "Telefono".ToUpper, 200))
                listEstCeldas.Add(New Modelo.Celda("ydfnac", True, "F.Nacimiento".ToUpper, 150, "MM/dd,YYYY"))
                Dim ef = New Efecto
                ef.tipo = 3
                ef.dt = dt
                ef.SeleclCol = 1
                ef.listEstCeldas = listEstCeldas
                ef.alto = 50
                ef.ancho = 350
                ef.Context = "Seleccione Vendedor".ToUpper
                ef.ShowDialog()
                Dim bandera As Boolean = False
                bandera = ef.band
                If (bandera = True) Then
                    Dim Row As Janus.Windows.GridEX.GridEXRow = ef.Row
                    If (IsNothing(Row)) Then
                        tbVendedor.Focus()
                        Return

                    End If
                    _CodEmpleado = Row.Cells("ydnumi").Value
                    tbVendedor.Text = Row.Cells("yddesc").Value
                    grdetalle.Select()

                End If

            End If

        End If
    End Sub
    Private Sub swTipoVenta_ValueChanged(sender As Object, e As EventArgs) Handles swTipoVenta.ValueChanged
        If (swTipoVenta.Value = False) Then
            lbCredito.Visible = True
            tbFechaVenc.Visible = True
            tbFechaVenc.Value = DateAdd(DateInterval.Day, _dias, Now.Date)
        Else
            lbCredito.Visible = False
            tbFechaVenc.Visible = False
        End If
    End Sub

    Private Sub grdetalle_EditingCell(sender As Object, e As EditingCellEventArgs) Handles grdetalle.EditingCell
        If (_fnAccesible()) Then
            'Habilitar solo las columnas de Precio, %, Monto y Observación
            'If (e.Column.Index = grdetalle.RootTable.Columns("yfcbarra").Index Or
            If (e.Column.Index = grdetalle.RootTable.Columns("yfcbarra").Index Or
                e.Column.Index = grdetalle.RootTable.Columns("tbcmin").Index Or
                e.Column.Index = grdetalle.RootTable.Columns("tbporc").Index Or
                e.Column.Index = grdetalle.RootTable.Columns("tbpbas").Index Or
                e.Column.Index = grdetalle.RootTable.Columns("tbdesc").Index) Then
                e.Cancel = False
            Else
                e.Cancel = True
            End If
        Else
            e.Cancel = True
        End If

    End Sub
    Private Sub grdetalle_KeyDown(sender As Object, e As KeyEventArgs) Handles grdetalle.KeyDown
        '        If (Not _fnAccesible()) Then
        '            Return
        '        End If
        '        If (e.KeyData = Keys.Enter) Then
        '            Dim f, c As Integer
        '            c = grdetalle.Col
        '            f = grdetalle.Row

        '            If (grdetalle.Col = grdetalle.RootTable.Columns("tbcmin").Index) Then
        '                If (grdetalle.GetValue("producto") <> String.Empty) Then
        '                    _prAddDetalleVenta()
        '                    _HabilitarProductos()
        '                Else
        '                    ToastNotification.Show(Me, "Seleccione un Producto Por Favor", My.Resources.WARNING, 3000, eToastGlowColor.Red, eToastPosition.TopCenter)
        '                End If

        '            End If
        '            If (grdetalle.Col = grdetalle.RootTable.Columns("producto").Index) Then
        '                If (grdetalle.GetValue("producto") <> String.Empty) Then
        '                    _prAddDetalleVenta()
        '                    _HabilitarProductos()
        '                Else
        '                    ToastNotification.Show(Me, "Seleccione un Producto Por Favor", My.Resources.WARNING, 3000, eToastGlowColor.Red, eToastPosition.TopCenter)
        '                End If

        '            End If
        '            If (grdetalle.Col = grdetalle.RootTable.Columns("yfcbarra").Index) Then
        '                If (grdetalle.GetValue("yfcbarra").ToString().Trim() <> String.Empty) Then
        '                    cargarProductos()
        '                    If (grdetalle.Row = grdetalle.RowCount - 1) Then
        '                        If (existeProducto(grdetalle.GetValue("yfcbarra").ToString)) Then
        '                            If (Not verificarExistenciaUnica(grdetalle.GetValue("yfcbarra").ToString)) Then
        '                                ponerProducto(grdetalle.GetValue("yfcbarra").ToString)
        '                                _prAddDetalleVenta()
        '                            Else
        '                                sumarCantidad(grdetalle.GetValue("yfcbarra").ToString)
        '                            End If
        '                        Else
        '                            grdetalle.DataChanged = False
        '                            ToastNotification.Show(Me, "El código de barra del producto no existe", My.Resources.WARNING, 3000, eToastGlowColor.Red, eToastPosition.TopCenter)
        '                        End If
        '                    Else
        '                        grdetalle.DataChanged = False
        '                        grdetalle.Row = grdetalle.RowCount - 1
        '                        grdetalle.Col = grdetalle.RootTable.Columns("yfcbarra").Index
        '                        ToastNotification.Show(Me, "El cursor debe situarse en la ultima fila", My.Resources.WARNING, 3000, eToastGlowColor.Red, eToastPosition.TopCenter)
        '                    End If
        '                Else
        '                    ToastNotification.Show(Me, "El código de barra no puede quedar vacio", My.Resources.WARNING, 3000, eToastGlowColor.Red, eToastPosition.TopCenter)
        '                End If

        '            End If
        '            'opcion para cargar la grilla con el codigo de barra
        '            'If (grdetalle.Col = grdetalle.RootTable.Columns("yfcbarra").Index) Then

        '            '    If (grdetalle.GetValue("yfcbarra") <> String.Empty) Then
        '            '        _buscarRegistro(grdetalle.GetValue("yfcbarra"))


        '            '        '_prAddDetalleVenta()
        '            '        '_HabilitarProductos()
        '            '        ' MsgBox("hola de la grilla" + grdetalle.GetValue("yfcbarra") + t.Container.ToString)
        '            '        'ojo
        '            '    Else
        '            '        ToastNotification.Show(Me, "Seleccione un Producto Por Favor", My.Resources.WARNING, 3000, eToastGlowColor.Red, eToastPosition.TopCenter)
        '            '    End If

        '            'End If
        'salirIf:
        '        End If

        '        If (e.KeyData = Keys.Control + Keys.Enter And grdetalle.Row >= 0 And
        '            grdetalle.Col = grdetalle.RootTable.Columns("producto").Index) Then
        '            Dim indexfil As Integer = grdetalle.Row
        '            Dim indexcol As Integer = grdetalle.Col
        '            _HabilitarProductos()

        '        End If
        '        If (e.KeyData = Keys.Escape And grdetalle.Row >= 0) Then

        '            _prEliminarFila()


        '        End If
        '        If (e.KeyData = Keys.Control + Keys.S) Then
        '            tbMontoBs.Select()
        '        End If

        Try
            If (Not _fnAccesible()) Then
                Return
            End If

            If (e.KeyData = Keys.Enter) Then
                Dim f, c As Integer
                c = grdetalle.Col
                f = grdetalle.Row

                If (grdetalle.Col = grdetalle.RootTable.Columns("tbcmin").Index) Then
                    If (grdetalle.GetValue("producto") <> String.Empty) Then
                        _prAddDetalleVenta()
                        _HabilitarProductos()
                    Else
                        ToastNotification.Show(Me, "Seleccione un Producto Por Favor", My.Resources.WARNING, 3000, eToastGlowColor.Red, eToastPosition.TopCenter)
                    End If

                End If
                If (grdetalle.Col = grdetalle.RootTable.Columns("producto").Index) Then
                    If (grdetalle.GetValue("producto") <> String.Empty) Then
                        _prAddDetalleVenta()
                        _HabilitarProductos()
                    Else
                        ToastNotification.Show(Me, "Seleccione un Producto Por Favor", My.Resources.WARNING, 3000, eToastGlowColor.Red, eToastPosition.TopCenter)
                    End If

                End If
                If (grdetalle.Col = grdetalle.RootTable.Columns("yfcbarra").Index) Then
                    If (grdetalle.GetValue("yfcbarra").ToString().Trim() <> String.Empty) Then
                        cargarProductos()
                        If (grdetalle.Row = grdetalle.RowCount - 1) Then
                            Dim codigoBarrasCompleto = grdetalle.GetValue("yfcbarra").ToString
                            Dim primerDigito As String = Mid(codigoBarrasCompleto, 1, 1)
                            If primerDigito = "2" Then
                                Dim codigoBarrasProducto As Integer
                                Dim totalEntero, totalDecimal, total2, total As Decimal
                                codigoBarrasProducto = Mid(codigoBarrasCompleto, 1, 6)
                                'CUANDO EL COD BARRA TENGA 6 DIGITOS  EJEM: 200001
                                If (existeProducto(codigoBarrasProducto)) Then

                                    totalEntero = Mid(codigoBarrasCompleto, 7, 4)
                                    totalDecimal = Mid(codigoBarrasCompleto, 11, 2)
                                    total2 = CDbl(totalEntero).ToString() + "." + CDbl(totalDecimal).ToString()
                                    total = CDbl(total2)
                                    If (Not verificarExistenciaUnica(codigoBarrasProducto)) Then
                                        ponerProducto2(codigoBarrasProducto, total, -1)
                                        _prAddDetalleVenta()
                                    Else
                                        ponerProducto2(codigoBarrasProducto, total, grdetalle.RowCount - 1)
                                        _prAddDetalleVenta()
                                    End If

                                Else
                                    ''CUANDO EL CODIGO DE BARRAS TENGA 7 DIGITOS EJEM: 2000001
                                    codigoBarrasProducto = Mid(codigoBarrasCompleto, 1, 7)
                                    If (existeProducto(codigoBarrasProducto)) Then
                                        totalEntero = Mid(codigoBarrasCompleto, 8, 3)
                                        totalDecimal = Mid(codigoBarrasCompleto, 11, 2)
                                        total2 = CDbl(totalEntero).ToString() + "." + CDbl(totalDecimal).ToString()
                                        total = CDbl(total2)
                                        If (Not verificarExistenciaUnica(codigoBarrasProducto)) Then
                                            ponerProducto2(codigoBarrasProducto, total, -1)
                                            _prAddDetalleVenta()
                                        Else
                                            ponerProducto2(codigoBarrasProducto, total, grdetalle.RowCount - 1)
                                            _prAddDetalleVenta()
                                        End If
                                    Else
                                        grdetalle.DataChanged = False
                                        ToastNotification.Show(Me, "El código de barra del producto no existe", My.Resources.WARNING, 3000, eToastGlowColor.Red, eToastPosition.TopCenter)
                                    End If
                                End If
                            Else

                                If (existeProducto(grdetalle.GetValue("yfcbarra").ToString)) Then
                                    If (Not verificarExistenciaUnica(grdetalle.GetValue("yfcbarra").ToString)) Then
                                        Dim resultado As Boolean = False
                                        ponerProducto(grdetalle.GetValue("yfcbarra").ToString, resultado)
                                        If resultado Then
                                            _prAddDetalleVenta()
                                        End If
                                    Else
                                        'If (grdetalle.GetValue("producto").ToString <> String.Empty) Then
                                        sumarCantidad(grdetalle.GetValue("yfcbarra").ToString)
                                        'Else
                                        '    Dim img As Bitmap = New Bitmap(My.Resources.mensaje, 50, 50)
                                        '    ToastNotification.Show(Me, "El Producto: NO CUENTA CON STOCK DISPONIBLE", img, 5000, eToastGlowColor.Red, eToastPosition.BottomCenter)
                                        '    FilaSelectLote = Nothing
                                        'End If
                                    End If
                                Else
                                    grdetalle.DataChanged = False
                                    ToastNotification.Show(Me, "El código de barra del producto no existe", My.Resources.WARNING, 3000, eToastGlowColor.Red, eToastPosition.TopCenter)
                                End If
                            End If
                        Else
                            grdetalle.DataChanged = False
                            grdetalle.Row = grdetalle.RowCount - 1
                            grdetalle.Col = grdetalle.RootTable.Columns("yfcbarra").Index
                            ToastNotification.Show(Me, "El cursor debe situarse en la ultima fila", My.Resources.WARNING, 3000, eToastGlowColor.Red, eToastPosition.TopCenter)
                        End If
                    Else
                        ToastNotification.Show(Me, "El código de barra no puede quedar vacio", My.Resources.WARNING, 3000, eToastGlowColor.Red, eToastPosition.TopCenter)
                    End If
                End If


                'opcion para cargar la grilla con el codigo de barra
                'If (grdetalle.Col = grdetalle.RootTable.Columns("yfcbarra").Index) Then

                '    If (grdetalle.GetValue("yfcbarra") <> String.Empty) Then
                '        _buscarRegistro(grdetalle.GetValue("yfcbarra"))


                '        '_prAddDetalleVenta()
                '        '_HabilitarProductos()
                '        ' MsgBox("hola de la grilla" + grdetalle.GetValue("yfcbarra") + t.Container.ToString)
                '        'ojo
                '    Else
                '        ToastNotification.Show(Me, "Seleccione un Producto Por Favor", My.Resources.WARNING, 3000, eToastGlowColor.Red, eToastPosition.TopCenter)
                '    End If

                'End If
salirIf:
            End If

            If (e.KeyData = Keys.Control + Keys.Enter And grdetalle.Row >= 0 And
                grdetalle.Col = grdetalle.RootTable.Columns("producto").Index) Then
                Dim indexfil As Integer = grdetalle.Row
                Dim indexcol As Integer = grdetalle.Col
                _HabilitarProductos()

            End If
            If (e.KeyData = Keys.Escape And grdetalle.Row >= 0) Then

                _prEliminarFila()


            End If
            If (e.KeyData = Keys.Control + Keys.S) Then
                tbMontoBs.Select()
            End If
        Catch ex As Exception
            MostrarMensajeError(ex.Message)
        End Try
    End Sub

    Private Sub ponerProducto2(codigo As String, total As Decimal, pos As Integer)
        Try
            grdetalle.DataChanged = True
            CType(grdetalle.DataSource, DataTable).AcceptChanges()
            Dim fila As DataRow() = Table_Producto.Select("yfcbarra='" + codigo.Trim + "'", "")
            If (fila.Count > 0) Then
                If pos = -1 Then
                    _fnObtenerFilaDetalle(pos, grdetalle.GetValue("tbnumi"))
                End If
                Dim cantidad = Format(total / CDbl(fila(0).ItemArray(15)), "0.00")
                Dim precio = fila(0).ItemArray(15)
                total = cantidad * precio
                CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbty5prod") = fila(0).ItemArray(0)
                CType(grdetalle.DataSource, DataTable).Rows(pos).Item("codigo") = fila(0).ItemArray(1)
                CType(grdetalle.DataSource, DataTable).Rows(pos).Item("yfcbarra") = fila(0).ItemArray(2)
                CType(grdetalle.DataSource, DataTable).Rows(pos).Item("producto") = fila(0).ItemArray(3)
                CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbumin") = fila(0).ItemArray(13)
                CType(grdetalle.DataSource, DataTable).Rows(pos).Item("unidad") = fila(0).ItemArray(14)
                'tbcmin
                CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbpbas") = precio
                CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbptot") = total
                CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbtotdesc") = total
                CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbcmin") = cantidad
                If (gb_FacturaIncluirICE) Then
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbpcos") = fila(0).ItemArray(17)
                Else
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbpcos") = 0
                End If
                ''Modif
                CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbpcos") = fila(0).ItemArray(16)
                CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbptot2") = fila(0).ItemArray(16) * cantidad
                '
                CType(grdetalle.DataSource, DataTable).Rows(pos).Item("stock") = fila(0).ItemArray(17) - cantidad
                'CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tblote") = grProductos.GetValue("iclot")
                'CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbfechaVenc") = grProductos.GetValue("icfven")
                'CType(grdetalle.DataSource, DataTable).Rows(pos).Item("stock") = grProductos.GetValue("iccven")
                _prCalcularPrecioTotal()
            End If
        Catch ex As Exception
            Throw New Exception
        End Try

    End Sub
    Private Sub cargarProductos()
        Dim dt As DataTable
        If (G_Lote = True) Then
            dt = L_fnListarProductos(cbSucursal.Value, Str(_CodCliente))  ''1=Almacen
            Table_Producto = dt.Copy
        Else
            dt = L_fnListarProductosSinLote(cbSucursal.Value, Str(_CodCliente), CType(grdetalle.DataSource, DataTable).Clone)  ''1=Almacen
            Table_Producto = dt.Copy
        End If
    End Sub
    Private Function existeProducto(codigo As String) As Boolean
        Return (Table_Producto.Select("yfcbarra='" + codigo.Trim() + "'", "").Count > 0)
    End Function

    Private Function verificarExistenciaUnica(codigo As String) As Boolean
        Dim cont As Integer = 0
        For Each fila As GridEXRow In grdetalle.GetRows()
            If (fila.Cells("yfcbarra").Value.ToString.Trim = codigo.Trim) Then
                cont += 1
            End If
        Next
        Return (cont >= 1)
    End Function

    Private Sub ponerProducto(codigo As String, ByRef resultado As Boolean)
        Try
            grdetalle.DataChanged = True
            CType(grdetalle.DataSource, DataTable).AcceptChanges()
            Dim fila As DataRow() = Table_Producto.Select("yfcbarra='" + codigo.Trim + "'", "")
            If (fila.Count > 0) Then
                'Si tiene stock > 0
                If fila(0).ItemArray(17) > 0 Then
                    Dim pos As Integer = -1
                    _fnObtenerFilaDetalle(pos, grdetalle.GetValue("tbnumi"))
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbty5prod") = fila(0).ItemArray(0)
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("codigo") = fila(0).ItemArray(1)
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("yfcbarra") = fila(0).ItemArray(2)
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("producto") = fila(0).ItemArray(3)
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbumin") = fila(0).ItemArray(13)
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("unidad") = fila(0).ItemArray(14)
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbpbas") = fila(0).ItemArray(15)
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbptot") = fila(0).ItemArray(15)
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbtotdesc") = fila(0).ItemArray(15)
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbcmin") = 1
                    If (gb_FacturaIncluirICE) Then
                        CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbpcos") = fila(0).ItemArray(17)
                    Else
                        CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbpcos") = 0
                    End If
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbpcos") = fila(0).ItemArray(16)
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbptot2") = fila(0).ItemArray(16)
                    'CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbptot2") = fila(0).ItemArray(17)
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("stock") = fila(0).ItemArray(17)
                    'CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tblote") = grProductos.GetValue("iclot")
                    'CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbfechaVenc") = grProductos.GetValue("icfven")
                    'CType(grdetalle.DataSource, DataTable).Rows(pos).Item("stock") = grProductos.GetValue("iccven")
                    _prCalcularPrecioTotal()
                    resultado = True
                Else
                    Dim img As Bitmap = New Bitmap(My.Resources.mensaje, 50, 50)
                    ToastNotification.Show(Me, "El Producto: ".ToUpper + fila(0).ItemArray(3) + " NO CUENTA CON STOCK DISPONIBLE", img, 5000, eToastGlowColor.Red, eToastPosition.BottomCenter)
                    FilaSelectLote = Nothing
                    resultado = False
                End If
            End If
        Catch ex As Exception
            MostrarMensajeError(ex.Message)
        End Try
    End Sub

    Private Sub sumarCantidad(codigo As String)
        Try
            Dim fila As DataRow() = CType(grdetalle.DataSource, DataTable).Select("yfcbarra='" + codigo.Trim + "'", "")
            If (fila.Count > 0) Then
                grdetalle.UpdateData()
                Dim pos1 As Integer = -1
                _fnObtenerFilaDetalle(pos1, fila(0).Item("tbnumi"))
                Dim cant As Integer = grdetalle.GetRow(pos1).Cells("tbcmin").Value + 1
                Dim stock As Integer = grdetalle.GetRow(pos1).Cells("stock").Value
                'If (cant > stock) Then
                Dim lin As Integer = grdetalle.GetRow(pos1).Cells("tbnumi").Value
                Dim pos2 As Integer = -1
                _fnObtenerFilaDetalle(pos2, lin)
                CType(grdetalle.DataSource, DataTable).Rows(pos2).Item("tbcmin") = cant
                CType(grdetalle.DataSource, DataTable).Rows(pos2).Item("tbptot") = CType(grdetalle.DataSource, DataTable).Rows(pos2).Item("tbpbas") * cant
                CType(grdetalle.DataSource, DataTable).Rows(pos2).Item("tbptot2") = grdetalle.GetRow(pos1).Cells("tbpcos").Value * cant
                CType(grdetalle.DataSource, DataTable).Rows(pos2).Item("tbtotdesc") = CType(grdetalle.DataSource, DataTable).Rows(pos2).Item("tbpbas") * cant
                Dim img As Bitmap = New Bitmap(My.Resources.mensaje, 50, 50)
                'ToastNotification.Show(Me, "La cantidad de la venta no debe ser mayor al del stock" & vbCrLf &
                '        "Stock=" + Str(stock).ToUpper, img, 2000, eToastGlowColor.Red, eToastPosition.BottomCenter)
                grdetalle.SetValue("yfcbarra", "")
                grdetalle.SetValue("tbcmin", 0)
                grdetalle.SetValue("tbptot", 0)
                grdetalle.SetValue("tbptot2", 0)
                grdetalle.DataChanged = True
                'grdetalle.Refetch()
                grdetalle.Refresh()
                '_prCalcularPrecioTotal()
                'Else
                '    If (cant = stock) Then
                '        'grdetalle.SelectedFormatStyle.ForeColor = Color.Blue
                '        'grdetalle.CurrentRow.Cells.Item(e.Column).FormatStyle = New GridEXFormatStyle
                '        'grdetalle.CurrentRow.Cells(e.Column).FormatStyle.BackColor = Color.Pink
                '        'grdetalle.CurrentRow.Cells.Item(e.Column).FormatStyle.BackColor = Color.DodgerBlue
                '        'grdetalle.CurrentRow.Cells.Item(e.Column).FormatStyle.ForeColor = Color.White
                '        'grdetalle.CurrentRow.Cells.Item(e.Column).FormatStyle.FontBold = TriState.True
                '    End If
                'End If

                _prCalcularPrecioTotal()
            End If
        Catch ex As Exception
            MostrarMensajeError(ex.Message)
        End Try
    End Sub

    Private Sub _buscarRegistro(cbarra As String)
        Dim _t As DataTable
        _t = L_fnListarProductosC(cbarra)
        If _t.Rows.Count > 0 Then
            CType(grdetalle.DataSource, DataTable).Rows(0).Item("producto") = _t.Rows(0).Item("yfcdprod1")
            CType(grdetalle.DataSource, DataTable).Rows(0).Item("tbcmin") = 1
            CType(grdetalle.DataSource, DataTable).Rows(0).Item("unidad") = _t.Rows(0).Item("uni")

        Else
            MsgBox("Codigo de Producto No Exite")
        End If
    End Sub
    Private Sub grProductos_KeyDown(sender As Object, e As KeyEventArgs) Handles grProductos.KeyDown
        Try
            If (Not _fnAccesible()) Then
                Return
            End If
            If (e.KeyData = Keys.Enter) Then
                Dim f, c As Integer
                c = grProductos.Col
                f = grProductos.Row
                If (f >= 0) Then

                    If (IsNothing(FilaSelectLote)) Then
                        ''''''''''''''''''''''''
                        If (G_Lote = True) Then
                            InsertarProductosConLote()
                        Else
                            InsertarProductosSinLote()
                        End If
                        '''''''''''''''
                    Else

                        '_fnExisteProductoConLote()
                        Dim pos As Integer = -1
                        grdetalle.Row = grdetalle.RowCount - 1
                        _fnObtenerFilaDetalle(pos, grdetalle.GetValue("tbnumi"))
                        Dim numiProd = FilaSelectLote.Item("yfnumi")
                        Dim lote As String = grProductos.GetValue("iclot")
                        Dim FechaVenc As Date = grProductos.GetValue("icfven")
                        If (Not _fnExisteProductoConLote(numiProd, lote, FechaVenc)) Then
                            'b.yfcdprod1, a.iclot, a.icfven, a.iccven
                            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbty5prod") = FilaSelectLote.Item("yfnumi")
                            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("codigo") = FilaSelectLote.Item("yfcprod")
                            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("yfcbarra") = FilaSelectLote.Item("yfcbarra")
                            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("producto") = FilaSelectLote.Item("yfcdprod1")
                            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbumin") = FilaSelectLote.Item("yfumin")
                            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("unidad") = FilaSelectLote.Item("UnidMin")
                            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbpbas") = FilaSelectLote.Item("yhprecio")
                            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbptot") = FilaSelectLote.Item("yhprecio")
                            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbtotdesc") = FilaSelectLote.Item("yhprecio")
                            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbcmin") = 1
                            'If (gb_FacturaIncluirICE) Then
                            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbpcos") = FilaSelectLote.Item("pcos")
                            'Else
                            '    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbpcos") = 0
                            'End If
                            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbptot2") = FilaSelectLote.Item("pcos")

                            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tblote") = grProductos.GetValue("iclot")
                            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbfechaVenc") = grProductos.GetValue("icfven")
                            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("stock") = grProductos.GetValue("iccven")
                            _prCalcularPrecioTotal()
                            _DesHabilitarProductos()
                            FilaSelectLote = Nothing
                        Else
                            Dim img As Bitmap = New Bitmap(My.Resources.mensaje, 50, 50)
                            ToastNotification.Show(Me, "El producto con el lote ya existe modifique su cantidad".ToUpper, img, 2000, eToastGlowColor.Red, eToastPosition.BottomCenter)
                        End If



                    End If

                End If
            End If
            If e.KeyData = Keys.Escape Then
                _DesHabilitarProductos()
                FilaSelectLote = Nothing
            End If
        Catch ex As Exception
            MostrarMensajeError(ex.Message)
        End Try
    End Sub
    Private Sub grdetalle_CellValueChanged(sender As Object, e As ColumnActionEventArgs) Handles grdetalle.CellValueChanged
        Try
            If (e.Column.Index = grdetalle.RootTable.Columns("tbcmin").Index) Or (e.Column.Index = grdetalle.RootTable.Columns("tbpbas").Index) Then
                If (Not IsNumeric(grdetalle.GetValue("tbcmin")) Or grdetalle.GetValue("tbcmin").ToString = String.Empty) Then

                    'grDetalle.GetRow(rowIndex).Cells("cant").Value = 1
                    '  grDetalle.CurrentRow.Cells.Item("cant").Value = 1
                    Dim lin As Integer = grdetalle.GetValue("tbnumi")
                    Dim pos As Integer = -1
                    _fnObtenerFilaDetalle(pos, lin)
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbcmin") = 1
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbptot") = CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbpbas")

                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbporc") = 0
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbdesc") = 0
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbtotdesc") = CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbpbas")
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbptot2") = CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbpcos")
                    'grdetalle.SetValue("tbcmin", 1)
                    'grdetalle.SetValue("tbptot", grdetalle.GetValue("tbpbas"))
                Else
                    If (grdetalle.GetValue("tbcmin") > 0) Then
                        Dim rowIndex As Integer = grdetalle.Row
                        Dim porcdesc As Double = grdetalle.GetValue("tbporc")
                        Dim montodesc As Double = ((grdetalle.GetValue("tbpbas") * grdetalle.GetValue("tbcmin")) * (porcdesc / 100))
                        Dim lin As Integer = grdetalle.GetValue("tbnumi")
                        Dim pos As Integer = -1
                        _fnObtenerFilaDetalle(pos, lin)
                        CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbdesc") = montodesc
                        grdetalle.SetValue("tbdesc", montodesc)
                        P_PonerTotal(rowIndex)

                    Else
                        Dim lin As Integer = grdetalle.GetValue("tbnumi")
                        Dim pos As Integer = -1
                        _fnObtenerFilaDetalle(pos, lin)
                        CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbcmin") = 1
                        CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbptot") = CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbpbas")
                        _prCalcularPrecioTotal()
                        'grdetalle.SetValue("tbcmin", 1)
                        'grdetalle.SetValue("tbptot", grdetalle.GetValue("tbpbas"))

                    End If
                End If
            End If
            '''''''''''''''''''''PORCENTAJE DE DESCUENTO '''''''''''''''''''''
            If (e.Column.Index = grdetalle.RootTable.Columns("tbporc").Index) Then
                If (Not IsNumeric(grdetalle.GetValue("tbporc")) Or grdetalle.GetValue("tbporc").ToString = String.Empty) Then

                    'grDetalle.GetRow(rowIndex).Cells("cant").Value = 1
                    '  grDetalle.CurrentRow.Cells.Item("cant").Value = 1
                    Dim lin As Integer = grdetalle.GetValue("tbnumi")
                    Dim pos As Integer = -1
                    _fnObtenerFilaDetalle(pos, lin)
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbporc") = 0
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbdesc") = 0
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbtotdesc") = CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbptot")
                    'grdetalle.SetValue("tbcmin", 1)
                    'grdetalle.SetValue("tbptot", grdetalle.GetValue("tbpbas"))
                Else
                    If (grdetalle.GetValue("tbporc") > 0 And grdetalle.GetValue("tbporc") <= 100) Then

                        Dim porcdesc As Double = grdetalle.GetValue("tbporc")
                        Dim montodesc As Double = (grdetalle.GetValue("tbptot") * (porcdesc / 100))
                        Dim lin As Integer = grdetalle.GetValue("tbnumi")
                        Dim pos As Integer = -1
                        _fnObtenerFilaDetalle(pos, lin)
                        CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbdesc") = montodesc
                        grdetalle.SetValue("tbdesc", montodesc)

                        Dim rowIndex As Integer = grdetalle.Row
                        P_PonerTotal(rowIndex)

                    Else
                        Dim lin As Integer = grdetalle.GetValue("tbnumi")
                        Dim pos As Integer = -1
                        _fnObtenerFilaDetalle(pos, lin)
                        CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbporc") = 0
                        CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbdesc") = 0
                        CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbtotdesc") = CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbptot")
                        grdetalle.SetValue("tbporc", 0)
                        grdetalle.SetValue("tbdesc", 0)
                        grdetalle.SetValue("tbtotdesc", grdetalle.GetValue("tbptot"))
                        _prCalcularPrecioTotal()
                        'grdetalle.SetValue("tbcmin", 1)
                        'grdetalle.SetValue("tbptot", grdetalle.GetValue("tbpbas"))

                    End If
                End If
            End If


            '''''''''''''''''''''MONTO DE DESCUENTO '''''''''''''''''''''
            If (e.Column.Index = grdetalle.RootTable.Columns("tbdesc").Index) Then
                If (Not IsNumeric(grdetalle.GetValue("tbdesc")) Or grdetalle.GetValue("tbdesc").ToString = String.Empty) Then

                    'grDetalle.GetRow(rowIndex).Cells("cant").Value = 1
                    '  grDetalle.CurrentRow.Cells.Item("cant").Value = 1
                    Dim lin As Integer = grdetalle.GetValue("tbnumi")
                    Dim pos As Integer = -1
                    _fnObtenerFilaDetalle(pos, lin)
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbporc") = 0
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbdesc") = 0
                    CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbtotdesc") = CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbptot")
                    'grdetalle.SetValue("tbcmin", 1)
                    'grdetalle.SetValue("tbptot", grdetalle.GetValue("tbpbas"))
                Else
                    If (grdetalle.GetValue("tbdesc") > 0 And grdetalle.GetValue("tbdesc") <= grdetalle.GetValue("tbptot")) Then

                        Dim montodesc As Double = grdetalle.GetValue("tbdesc")
                        Dim pordesc As Double = ((montodesc * 100) / grdetalle.GetValue("tbptot"))

                        Dim lin As Integer = grdetalle.GetValue("tbnumi")
                        Dim pos As Integer = -1
                        _fnObtenerFilaDetalle(pos, lin)
                        CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbdesc") = montodesc
                        CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbporc") = pordesc

                        grdetalle.SetValue("tbporc", pordesc)
                        Dim rowIndex As Integer = grdetalle.Row
                        P_PonerTotal(rowIndex)

                    Else
                        Dim lin As Integer = grdetalle.GetValue("tbnumi")
                        Dim pos As Integer = -1
                        _fnObtenerFilaDetalle(pos, lin)
                        CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbporc") = 0
                        CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbdesc") = 0
                        CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbtotdesc") = CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbptot")
                        grdetalle.SetValue("tbporc", 0)
                        grdetalle.SetValue("tbdesc", 0)
                        grdetalle.SetValue("tbtotdesc", grdetalle.GetValue("tbptot"))
                        _prCalcularPrecioTotal()
                        'grdetalle.SetValue("tbcmin", 1)
                        'grdetalle.SetValue("tbptot", grdetalle.GetValue("tbpbas"))

                    End If
                End If
            End If
        Catch ex As Exception
            MostrarMensajeError(ex.Message)
        End Try
    End Sub
    Private Sub tbPdesc_ValueChanged(sender As Object, e As EventArgs) Handles tbPdesc.ValueChanged
        Try
            If (tbPdesc.Focused) Then
                If (Not tbPdesc.Text = String.Empty And Not tbTotalBs.Text = String.Empty) Then
                    If (tbPdesc.Value = 0 Or tbPdesc.Value > 100) Then
                        tbPdesc.Value = 0
                        tbMdesc.Value = 0

                        _prCalcularPrecioTotal()


                    Else

                        Dim porcdesc As Double = tbPdesc.Value
                        Dim montodesc As Double = (grdetalle.GetTotal(grdetalle.RootTable.Columns("tbtotdesc"), AggregateFunction.Sum) * (porcdesc / 100))
                        tbMdesc.Value = montodesc

                        tbIce.Value = grdetalle.GetTotal(grdetalle.RootTable.Columns("tbptot2"), AggregateFunction.Sum) * (gi_ICE / 100)

                        If (gb_FacturaIncluirICE = True) Then
                            tbPrueba.Value = grdetalle.GetTotal(grdetalle.RootTable.Columns("tbtotdesc"), AggregateFunction.Sum) - montodesc + tbIce.Value
                        Else
                            tbPrueba.Value = grdetalle.GetTotal(grdetalle.RootTable.Columns("tbtotdesc"), AggregateFunction.Sum) - montodesc
                        End If
                    End If


                End If
                If (tbPdesc.Text = String.Empty) Then
                    tbMdesc.Value = 0

                End If
            End If
        Catch ex As Exception
            MostrarMensajeError(ex.Message)
        End Try
    End Sub

    Private Sub tbMdesc_ValueChanged(sender As Object, e As EventArgs) Handles tbMdesc.ValueChanged

        Try
            If (tbMdesc.Focused) Then

                Dim total As Double = Convert.ToDouble(tbTotalBs.Text)
                If (Not tbMdesc.Text = String.Empty And Not tbMdesc.Text = String.Empty) Then
                    If (tbMdesc.Value = 0 Or tbMdesc.Value > total) Then
                        tbMdesc.Value = 0
                        tbPdesc.Value = 0
                        _prCalcularPrecioTotal()
                    Else
                        Dim montodesc As Double = tbMdesc.Value
                        Dim pordesc As Double = ((montodesc * 100) / grdetalle.GetTotal(grdetalle.RootTable.Columns("tbtotdesc"), AggregateFunction.Sum))
                        tbPdesc.Value = pordesc
                        tbIce.Value = grdetalle.GetTotal(grdetalle.RootTable.Columns("tbptot2"), AggregateFunction.Sum) * (gi_ICE / 100)
                        If (gb_FacturaIncluirICE = True) Then
                            tbPrueba.Value = grdetalle.GetTotal(grdetalle.RootTable.Columns("tbtotdesc"), AggregateFunction.Sum) - montodesc + tbIce.Value
                        Else
                            tbPrueba.Value = grdetalle.GetTotal(grdetalle.RootTable.Columns("tbtotdesc"), AggregateFunction.Sum) - montodesc
                        End If

                    End If

                End If

                If (tbMdesc.Text = String.Empty) Then
                    tbMdesc.Value = 0

                End If
            End If
        Catch ex As Exception
            MostrarMensajeError(ex.Message)
        End Try

    End Sub

    Private Sub grdetalle_CellEdited_1(sender As Object, e As ColumnActionEventArgs) Handles grdetalle.CellEdited
        Try
            If (e.Column.Index = grdetalle.RootTable.Columns("tbcmin").Index) Then
                If (Not IsNumeric(grdetalle.GetValue("tbcmin")) Or grdetalle.GetValue("tbcmin").ToString = String.Empty) Then

                    grdetalle.SetValue("tbcmin", 1)
                    grdetalle.SetValue("tbptot", grdetalle.GetValue("tbpbas"))
                    grdetalle.SetValue("tbporc", 0)
                    grdetalle.SetValue("tbdesc", 0)
                    grdetalle.SetValue("tbtotdesc", grdetalle.GetValue("tbpbas"))


                Else
                    If (grdetalle.GetValue("tbcmin") > 0) Then

                        Dim cant As Integer = grdetalle.GetValue("tbcmin")

                        Dim stock As Double = grdetalle.GetValue("stock")
                        If (cant > stock) And stock <> -9999 Then
                            Dim lin As Integer = grdetalle.GetValue("tbnumi")
                            Dim pos As Integer = -1
                            _fnObtenerFilaDetalle(pos, lin)
                            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbcmin") = 1
                            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbptot") = CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbpbas")
                            CType(grdetalle.DataSource, DataTable).Rows(pos).Item("tbptot2") = grdetalle.GetValue("tbpcos") * 1
                            Dim img As Bitmap = New Bitmap(My.Resources.mensaje, 50, 50)
                            ToastNotification.Show(Me, "La cantidad de la venta no debe ser mayor al del stock" & vbCrLf &
                            "Stock=" + Str(stock).ToUpper, img, 2000, eToastGlowColor.Red, eToastPosition.BottomCenter)
                            grdetalle.SetValue("tbcmin", 1)
                            grdetalle.SetValue("tbptot", grdetalle.GetValue("tbpbas"))
                            grdetalle.SetValue("tbptot2", grdetalle.GetValue("tbpcos") * 1)

                            _prCalcularPrecioTotal()
                        Else
                            If (cant = stock) Then


                                'grdetalle.SelectedFormatStyle.ForeColor = Color.Blue
                                'grdetalle.CurrentRow.Cells.Item(e.Column).FormatStyle = New GridEXFormatStyle
                                'grdetalle.CurrentRow.Cells(e.Column).FormatStyle.BackColor = Color.Pink
                                'grdetalle.CurrentRow.Cells.Item(e.Column).FormatStyle.BackColor = Color.DodgerBlue
                                'grdetalle.CurrentRow.Cells.Item(e.Column).FormatStyle.ForeColor = Color.White
                                'grdetalle.CurrentRow.Cells.Item(e.Column).FormatStyle.FontBold = TriState.True
                            End If
                        End If

                    Else

                        grdetalle.SetValue("tbcmin", 1)
                        grdetalle.SetValue("tbptot", grdetalle.GetValue("tbpbas"))
                        grdetalle.SetValue("tbporc", 0)
                        grdetalle.SetValue("tbdesc", 0)
                        grdetalle.SetValue("tbtotdesc", grdetalle.GetValue("tbpbas"))

                    End If
                End If
            End If
        Catch ex As Exception
            MostrarMensajeError(ex.Message)
        End Try
    End Sub
    Private Sub grdetalle_MouseClick(sender As Object, e As MouseEventArgs) Handles grdetalle.MouseClick
        Try
            If (Not _fnAccesible()) Then
                Return
            End If
            If (grdetalle.RowCount >= 2) Then
                If (grdetalle.CurrentColumn.Index = grdetalle.RootTable.Columns("img").Index) Then
                    _prEliminarFila()
                End If
            End If
        Catch ex As Exception
            MostrarMensajeError(ex.Message)
        End Try

    End Sub
    Private Sub btnGrabar_Click(sender As Object, e As EventArgs) Handles btnGrabar.Click
        _prGuardar()
    End Sub

    Private Sub btnModificar_Click(sender As Object, e As EventArgs) Handles btnModificar.Click
        Try
            If (grVentas.RowCount > 0) Then
                If (gb_FacturaEmite) Then
                    If (Not P_fnValidarFacturaVigente()) Then
                        Dim img As Bitmap = New Bitmap(My.Resources.WARNING, 50, 50)

                        ToastNotification.Show(Me, "No se puede modificar la venta con codigo ".ToUpper + tbCodigo.Text + ", su factura esta anulada.".ToUpper,
                                                  img, 2000,
                                                  eToastGlowColor.Green,
                                                  eToastPosition.TopCenter)
                        Exit Sub
                    End If
                End If

                _prhabilitar()
                btnNuevo.Enabled = False
                btnModificar.Enabled = False
                btnEliminar.Enabled = False
                btnGrabar.Enabled = True

                PanelNavegacion.Enabled = False
                _prCargarIconELiminar()
            End If
        Catch ex As Exception
            MostrarMensajeError(ex.Message)
        End Try
    End Sub
    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        Try
            If _CodCliente <> 86 Then
                If (gb_FacturaEmite) Then
                    If (P_fnValidarFacturaVigente()) Then
                        Dim img As Bitmap = New Bitmap(My.Resources.WARNING, 50, 50)

                        ToastNotification.Show(Me, "No se puede anular la venta con código ".ToUpper + tbCodigo.Text + ", su factura esta vigente, por favor primero anule la factura".ToUpper,
                                                  img, 3000,
                                                  eToastGlowColor.Green,
                                                  eToastPosition.TopCenter)
                        Exit Sub
                    End If
                End If

                If (swTipoVenta.Value = False) Then
                    Dim res1 As Boolean = L_fnVerificarPagos(tbCodigo.Text)
                    If res1 Then
                        Dim img As Bitmap = New Bitmap(My.Resources.WARNING, 50, 50)
                        ToastNotification.Show(Me, "No se puede anular la venta con código ".ToUpper + tbCodigo.Text + ", porque tiene pagos realizados, por favor primero elimine los pagos correspondientes a esta venta".ToUpper,
                                                  img, 5000,
                                                  eToastGlowColor.Green,
                                                  eToastPosition.TopCenter)
                        Exit Sub
                    End If
                End If
            End If
            Dim result As Boolean = L_fnVerificarSiSeContabilizoVenta(tbCodigo.Text)
            If result Then
                Dim img As Bitmap = New Bitmap(My.Resources.cancel, 50, 50)
                ToastNotification.Show(Me, "La Venta no puede ser anulada porque ya fue contabilizada".ToUpper, img, 4500, eToastGlowColor.Red, eToastPosition.TopCenter)
            End If
            Dim ef = New Efecto
            ef.tipo = 2
            ef.Context = "¿esta seguro de eliminar el registro?".ToUpper
            ef.Header = "mensaje principal".ToUpper
            ef.ShowDialog()
            Dim bandera As Boolean = False
            bandera = ef.band
            If (bandera = True) Then
                Dim mensajeError As String = ""
                Dim res As Boolean = L_fnEliminarVenta(tbCodigo.Text, mensajeError)
                If res Then
                    Dim img As Bitmap = New Bitmap(My.Resources.checked, 50, 50)
                    ToastNotification.Show(Me, "Código de Venta ".ToUpper + tbCodigo.Text + " eliminado con Exito.".ToUpper,
                                              img, 2000,
                                              eToastGlowColor.Green,
                                              eToastPosition.TopCenter)
                    _prFiltrar()
                Else
                    Dim img As Bitmap = New Bitmap(My.Resources.cancel, 50, 50)
                    ToastNotification.Show(Me, mensajeError, img, 2000, eToastGlowColor.Red, eToastPosition.BottomCenter)
                End If
            End If
        Catch ex As Exception
            MostrarMensajeError(ex.Message)
        End Try
    End Sub

    Private Sub grVentas_SelectionChanged(sender As Object, e As EventArgs) Handles grVentas.SelectionChanged
        If (grVentas.RowCount >= 0 And grVentas.Row >= 0) Then
            _prMostrarRegistro(grVentas.Row)
        End If
    End Sub

    Private Sub btnSiguiente_Click(sender As Object, e As EventArgs) Handles btnSiguiente.Click
        Dim _pos As Integer = grVentas.Row
        If _pos < grVentas.RowCount - 1 And _pos >= 0 Then
            _pos = grVentas.Row + 1
            '' _prMostrarRegistro(_pos)
            grVentas.Row = _pos
        End If
    End Sub

    Private Sub btnUltimo_Click(sender As Object, e As EventArgs) Handles btnUltimo.Click
        Dim _pos As Integer = grVentas.Row
        If grVentas.RowCount > 0 Then
            _pos = grVentas.RowCount - 1
            ''  _prMostrarRegistro(_pos)
            grVentas.Row = _pos
        End If
    End Sub

    Private Sub btnAnterior_Click(sender As Object, e As EventArgs) Handles btnAnterior.Click
        Dim _MPos As Integer = grVentas.Row
        If _MPos > 0 And grVentas.RowCount > 0 Then
            _MPos = _MPos - 1
            ''  _prMostrarRegistro(_MPos)
            grVentas.Row = _MPos
        End If
    End Sub

    Private Sub btnPrimero_Click(sender As Object, e As EventArgs) Handles btnPrimero.Click
        _PrimerRegistro()
    End Sub
    Private Sub grVentas_KeyDown(sender As Object, e As KeyEventArgs) Handles grVentas.KeyDown
        If e.KeyData = Keys.Enter Then
            MSuperTabControl.SelectedTabIndex = 0
            grdetalle.Focus()

        End If
    End Sub

    Private Sub TbNit_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles tbNit.Validating
        Dim nom1, nom2 As String
        nom1 = ""
        nom2 = ""
        If (tbNit.Text.Trim = String.Empty) Then
            tbNit.Text = "0"
        End If
        L_Validar_Nit(tbNit.Text.Trim, nom1, nom2)
        If nom1 = "" Then
            TbNombre1.Focus()
        Else
            TbNombre1.Text = nom1
            TbNombre2.Text = nom2
        End If


    End Sub
    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Try
            If (Not _fnAccesible()) Then

                If (gb_FacturaEmite) Then
                    If tbCodigo.Text = String.Empty Then
                        Throw New Exception("Venta no encontrada")
                    End If
                    If _CodCliente <> 86 Then
                        _prImiprimirNotaVenta(tbCodigo.Text)
                        Return
                    ElseIf (Not P_fnValidarFacturaVigente()) Then

                        Dim img As Bitmap = New Bitmap(My.Resources.WARNING, 50, 50)

                        ToastNotification.Show(Me, "No se puede imprimir la factura con numero ".ToUpper + tbNroFactura.Text + ", su factura esta anulada".ToUpper,
                                              img, 3000,
                                              eToastGlowColor.Green,
                                              eToastPosition.TopCenter)
                        Exit Sub
                    End If
                    ReimprimirFactura(tbCodigo.Text, True, True)
                Else
                    _prImiprimirNotaVenta(tbCodigo.Text)
                End If
            End If
        Catch ex As Exception
            MostrarMensajeError(ex.Message)
        End Try
    End Sub

    Private Sub TbNit_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbNit.KeyPress
        g_prValidarTextBox(1, e)
    End Sub

    Private Sub swTipoVenta_Leave(sender As Object, e As EventArgs) Handles swTipoVenta.Leave
        grdetalle.Select()
    End Sub

    Private Sub cbSucursal_ValueChanged(sender As Object, e As EventArgs) Handles cbSucursal.ValueChanged
        Table_Producto = Nothing
    End Sub

    Private Sub tbNit_Leave(sender As Object, e As EventArgs) Handles tbNit.Leave, TbNombre1.Leave
        grdetalle.Select()
    End Sub
    Private Sub DoubleInput2_ValueChanged(sender As Object, e As EventArgs) Handles tbMontoBs.ValueChanged
        If tbMontoBs.Value <> 0 And tbMontoBs.Text <> String.Empty Then
            txtMontoPagado1.Text = tbMontoBs.Value + (tbMontoDolar.Value * IIf(cbCambioDolar.Text = "", 0, Convert.ToDecimal(cbCambioDolar.Text))) + tbMontoTarej.Value
            If Convert.ToDecimal(tbTotalBs.Text) <> 0 And Convert.ToDecimal(txtMontoPagado1.Text) >= Convert.ToDecimal(tbTotalBs.Text) Then
                txtCambio1.Text = Convert.ToDecimal(txtMontoPagado1.Text) - Convert.ToDecimal(tbTotalBs.Text)
            Else
                txtCambio1.Text = "0.00"
            End If
        End If
    End Sub
    Private Sub tbMontoDolar_ValueChanged(sender As Object, e As EventArgs) Handles tbMontoDolar.ValueChanged
        If tbMontoDolar.Value <> 0 And tbMontoDolar.Text <> String.Empty Then
            txtMontoPagado1.Text = tbMontoBs.Value + (tbMontoDolar.Value * IIf(cbCambioDolar.Text = "", 0, Convert.ToDecimal(cbCambioDolar.Text))) + tbMontoTarej.Value
            If Convert.ToDecimal(tbTotalBs.Text) <> 0 And Convert.ToDecimal(txtMontoPagado1.Text) >= Convert.ToDecimal(tbTotalBs.Text) Then
                txtCambio1.Text = Convert.ToDecimal(txtMontoPagado1.Text) - Convert.ToDecimal(tbTotalBs.Text)
            Else
                txtCambio1.Text = "0.00"
            End If
        End If
    End Sub

    Private Sub tbMontoTarej_ValueChanged(sender As Object, e As EventArgs) Handles tbMontoTarej.ValueChanged
        If tbMontoTarej.Value <> 0 And tbMontoTarej.Text <> String.Empty Then
            txtMontoPagado1.Text = tbMontoBs.Value + (tbMontoDolar.Value * IIf(cbCambioDolar.Text = "", 0, Convert.ToDecimal(cbCambioDolar.Text))) + tbMontoTarej.Value
            If Convert.ToDecimal(tbTotalBs.Text) <> 0 And Convert.ToDecimal(txtMontoPagado1.Text) >= Convert.ToDecimal(tbTotalBs.Text) Then
                txtCambio1.Text = Convert.ToDecimal(txtMontoPagado1.Text) - Convert.ToDecimal(tbTotalBs.Text)
            Else
                txtCambio1.Text = "0.00"
            End If
        End If
    End Sub
    Private Sub chbTarejetaa_CheckedChanged(sender As Object, e As EventArgs) Handles chbTarjeta.CheckedChanged
        If chbTarjeta.Checked Then
            tbMontoBs.Value = 0
            tbMontoDolar.Value = 0
            tbMontoTarej.Value = Convert.ToDecimal(tbTotalBs.Text)
        End If
    End Sub
    Private Sub cbCambioDolar_ValueChanged_1(sender As Object, e As EventArgs) Handles cbCambioDolar.ValueChanged
        If cbCambioDolar.SelectedIndex < 0 And cbCambioDolar.Text <> String.Empty Then
            btgrupo1.Visible = True
        Else
            btgrupo1.Visible = False
        End If
    End Sub
    Private Sub btgrupo1_Click(sender As Object, e As EventArgs) Handles btgrupo1.Click
        Dim numi As String = ""

        If L_prLibreriaGrabar(numi, "7", "1", cbCambioDolar.Text, "") Then
            _prCargarComboLibreria(cbCambioDolar, "7", "1")
            cbCambioDolar.SelectedIndex = CType(cbCambioDolar.DataSource, DataTable).Rows.Count - 1
        End If
    End Sub
    Private Sub _prCargarComboLibreria(mCombo As Janus.Windows.GridEX.EditControls.MultiColumnCombo, cod1 As String, cod2 As String)
        Dim dt As New DataTable
        dt = L_prLibreriaClienteLGeneral(cod1, cod2)
        With mCombo
            .DropDownList.Columns.Clear()
            .DropDownList.Columns.Add("yccod3").Width = 70
            .DropDownList.Columns("yccod3").Caption = "COD"
            .DropDownList.Columns.Add("ycdes3").Width = 200
            .DropDownList.Columns("ycdes3").Caption = "DESCRIPCION"
            .ValueMember = "yccod3"
            .DisplayMember = "ycdes3"
            .DataSource = dt
            .Refresh()
        End With
    End Sub

    Private Sub tbMontoBs_KeyDown(sender As Object, e As KeyEventArgs) Handles tbMontoBs.KeyDown
        If (e.KeyData = Keys.Control + Keys.A) Then
            _prGuardar()
        End If
    End Sub

    Private Sub tbMontoDolar_KeyDown(sender As Object, e As KeyEventArgs) Handles tbMontoDolar.KeyDown
        If (e.KeyData = Keys.Control + Keys.A) Then
            _prGuardar()
        End If
    End Sub
    Private Sub tbMontoTarej_KeyDown(sender As Object, e As KeyEventArgs) Handles tbMontoTarej.KeyDown
        If (e.KeyData = Keys.Control + Keys.A) Then
            _prGuardar()
        End If
    End Sub

    Private Sub txtMontoPagado1_KeyDown(sender As Object, e As KeyEventArgs) Handles txtMontoPagado1.KeyDown
        If (e.KeyData = Keys.Control + Keys.A) Then
            _prGuardar()
        End If
    End Sub

    Private Sub txtCambio1_KeyDown(sender As Object, e As KeyEventArgs) Handles txtCambio1.KeyDown
        If (e.KeyData = Keys.Control + Keys.A) Then
            _prGuardar()
        End If
    End Sub

    Private Sub GroupPanel1_Click(sender As Object, e As EventArgs) Handles GroupPanel1.Click

    End Sub

    Private Sub grdetalle_Enter(sender As Object, e As EventArgs) Handles grdetalle.Enter
        Try
            If (_fnAccesible()) Then
                If (_CodCliente <= 0) Then
                    ToastNotification.Show(Me, "           Antes de Continuar Por favor Seleccione un Cliente!!             ", My.Resources.WARNING, 4000, eToastGlowColor.Red, eToastPosition.TopCenter)
                    tbCliente.Focus()

                    Return
                End If
                If (_CodEmpleado <= 0) Then


                    ToastNotification.Show(Me, "           Antes de Continuar Por favor Seleccione un Vendedor!!             ", My.Resources.WARNING, 4000, eToastGlowColor.Red, eToastPosition.TopCenter)
                    tbVendedor.Focus()
                    Return

                End If

                grdetalle.Select()
                If _codeBar = 1 Then
                    If gb_CodigoBarra Then
                        grdetalle.Col = 3
                        grdetalle.Row = 0
                    Else
                        grdetalle.Col = 5
                        grdetalle.Row = 0
                    End If
                End If
            End If
        Catch ex As Exception
            MostrarMensajeError(ex.Message)
        End Try
    End Sub

    Private Sub btnDuplicar_Click(sender As Object, e As EventArgs) Handles btnDuplicar.Click
        Try
            Dim Nit As String = tbNit.Text
            Dim Razon_Social As String = TbNombre1.Text
            Dim Cliente As String = tbCliente.Text
            Dim Vend As String = tbVendedor.Text
            Dim TipoVenta As Boolean = swTipoVenta.Value
            Dim FechaVenc As Date = tbFechaVenc.Value
            Dim TotalBs As String = tbTotalBs.Text
            Dim TotalDo As String = tbTotalDo.Text

            Dim table As DataTable = grdetalle.DataSource

            btnNuevo.PerformClick()
            tbNit.Text = Nit
            TbNombre1.Text = Razon_Social
            tbCliente.Text = Cliente
            tbVendedor.Text = Vend
            swTipoVenta.Value = TipoVenta
            tbFechaVenc.Value = FechaVenc
            tbTotalBs.Text = TotalBs
            tbTotalDo.Text = TotalDo
            txtEstado.Clear()
            'txtEstado.Text = "VIGENTE"
            'grdetalle.DataSource = table

            'Dim _detalle1 As DataTable = grdetalle.DataSource

            'For j As Integer = 0 To grdetalle.RowCount - 1 Step 1
            '    grdetalle.Row = j
            '    _detalle1.Rows(j).Item("tbtv1numi") = 0
            '    _detalle1.Rows(j).Item("estado") = 0
            'Next
            'grdetalle.DataSource = _detalle1


            For j As Integer = 0 To table.Rows.Count - 1 Step 1

                table.Rows(j).Item("tbtv1numi") = 0
                table.Rows(j).Item("estado") = 0
            Next
            grdetalle.DataSource = table
            _prCargarIconELiminar()
        Catch ex As Exception
            MostrarMensajeError(ex.Message)
        End Try
    End Sub

    'Private Sub TbNombre1_KeyDown(sender As Object, e As KeyEventArgs) Handles TbNombre1.KeyDown
    '    If (e.KeyData = Keys.Enter) Then
    '        grdetalle.Col = 7
    '        grdetalle.Row = 0
    '        grdetalle.Select()

    '    End If
    'End Sub





#End Region
End Class