'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_fPedidos"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_fPedidos

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("SELECT	Pedidos.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("       Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("       Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("       Clientes.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("       Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("       Pedidos.Documento,")
            loComandoSeleccionar.AppendLine("       Pedidos.Fec_Ini,      ")
            loComandoSeleccionar.AppendLine("       Pedidos.Comentario, ")
            loComandoSeleccionar.AppendLine("       Pedidos.Status, ")
            loComandoSeleccionar.AppendLine("		Renglones_Pedidos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("       Renglones_Pedidos.Comentario		AS Com_Ren,")
            loComandoSeleccionar.AppendLine("		Renglones_Pedidos.Notas, ")
            loComandoSeleccionar.AppendLine("       Articulos.Nom_Art,  ")
            loComandoSeleccionar.AppendLine("       Renglones_Pedidos.Can_Art1, ")
            loComandoSeleccionar.AppendLine("       Renglones_Pedidos.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("       Renglones_Pedidos.Precio1,")
            loComandoSeleccionar.AppendLine("       Renglones_Pedidos.Mon_Net           AS Neto,")
            loComandoSeleccionar.AppendLine("       Pedidos.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("       Pedidos.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("       Pedidos.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("       Pedidos.Mon_Net, ")
            loComandoSeleccionar.AppendLine("       (SELECT Nom_Usu FROM Factory_Global.dbo.Usuarios WHERE Cod_Usu COLLATE DATABASE_DEFAULT = Pedidos.Usu_Cre COLLATE DATABASE_DEFAULT) AS Usuario")
            loComandoSeleccionar.AppendLine("     ")
            loComandoSeleccionar.AppendLine("FROM Pedidos ")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_Pedidos ON Pedidos.Documento = Renglones_Pedidos.Documento ")
            loComandoSeleccionar.AppendLine("   JOIN Clientes ON Pedidos.Cod_Cli = Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("   JOIN Formas_Pagos ON Pedidos.Cod_For = Formas_Pagos.Cod_For ")
            loComandoSeleccionar.AppendLine("   JOIN Vendedores ON Pedidos.Cod_Ven = Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("   JOIN Articulos ON Articulos.Cod_Art = Renglones_Pedidos.Cod_Art ")
            loComandoSeleccionar.AppendLine("WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '-------------------------------------------------------------------------------------------'
            ' Carga la imagen del logo en cusReportes                                                   '
            '-------------------------------------------------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            '-------------------------------------------------------------------------------------------'
            ' Verificando si el select (tabla nº0) trae registros                                       '
            '-------------------------------------------------------------------------------------------'

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                          vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("CGS_fPedidos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvCGS_fPedidos.ReportSource = loObjetoReporte

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
' JJD: 27/12/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 19/12/09: Ajuste al formato del impuesto IVA
'-------------------------------------------------------------------------------------------'
' CMS: 17/03/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'
' CMS: 11/06/10: Proveedor Genarico
'-------------------------------------------------------------------------------------------'
' JJD: 11/03/5: Ajustes al formato para el cliente CEGASA
'-------------------------------------------------------------------------------------------'