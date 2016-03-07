'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fDevoluciones_Clientes_Ingles"
'-------------------------------------------------------------------------------------------'
Partial Class fDevoluciones_Clientes_Ingles

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Devoluciones_Clientes.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Nom_Cli        As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Rif            As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Nit            As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Dir_Fis        As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Telefonos      As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Documento, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_DClientes.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_DClientes.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_DClientes.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_DClientes.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_DClientes.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_DClientes.Mon_Net  As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_DClientes.Por_Imp1 As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_DClientes.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_DClientes.Mon_Imp1 As  Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      Devoluciones_Clientes, ")
            loComandoSeleccionar.AppendLine("           Renglones_DClientes, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Devoluciones_Clientes.Documento     =   Renglones_DClientes.Documento AND ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Cod_Cli       =   Clientes.Cod_Cli AND ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Cod_For       =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Cod_Ven       =   Vendedores.Cod_Ven AND ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art                   =   Renglones_DClientes.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes          '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
      vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fDevoluciones_Clientes_Ingles", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfDevoluciones_Clientes_Ingles.ReportSource = loObjetoReporte

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub

    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload

        Try

            loObjetoReporte.Close()

        Catch loExcepcion As Exception

        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' Douglas Cortez: 06/05/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
' RAC: 24/03/2011: Modificacion de las Etiquetas Client por Customer en el archivo rpt.
'-------------------------------------------------------------------------------------------'
' MAT: 15/09/11: Ajuste de la vista de Diseño
'-------------------------------------------------------------------------------------------'
' MAT: 15/09/11: Eliminación del Pie de Página de eFactory según Requerimientos
'-------------------------------------------------------------------------------------------'
