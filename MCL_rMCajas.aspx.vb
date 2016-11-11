'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "MCL_rMCajas"
'-------------------------------------------------------------------------------------------'
Partial Class MCL_rMCajas

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            'Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @lcCajaDesde AS VARCHAR(10) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lfFechaDesde AS DATE = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lfFechaHasta AS DATE = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lnMonto AS DECIMAL(28,10) = " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lnSaldo AS DECIMAL(28,10) = " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Movimientos_Cajas.Documento					AS Documento,")
            loComandoSeleccionar.AppendLine("		Movimientos_Cajas.Cod_Caj					AS Cod_Caj,")
            loComandoSeleccionar.AppendLine("       Movimientos_Cajas.Comentario				AS Comentario,")
            loComandoSeleccionar.AppendLine("       Movimientos_Cajas.Fec_Ini					AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("       Movimientos_Cajas.Status					AS Status,")
            loComandoSeleccionar.AppendLine("       Movimientos_Cajas.Mon_Deb					AS Mon_Deb,")
            loComandoSeleccionar.AppendLine("       Movimientos_Cajas.Mon_Hab					AS Mon_Hab,")
            loComandoSeleccionar.AppendLine("       SUBSTRING(Movimientos_Cajas.Tip_Ori,1,20)	AS Tip_Ori,")
            loComandoSeleccionar.AppendLine("       Movimientos_Cajas.Doc_Ori					AS Doc_Ori,")
            loComandoSeleccionar.AppendLine("       Conceptos.Nom_Con							AS Nom_Con,")
            loComandoSeleccionar.AppendLine("       Cajas.Nom_Caj								AS Nombre,")
            loComandoSeleccionar.AppendLine("		@lnMonto									AS Monto,")
            loComandoSeleccionar.AppendLine("		@lnSaldo									AS Saldo,")
            loComandoSeleccionar.AppendLine("		@lfFechaDesde								AS Desde,")
            loComandoSeleccionar.AppendLine("		@lfFechaHasta								AS Hasta")
            loComandoSeleccionar.AppendLine("FROM Movimientos_Cajas")
            loComandoSeleccionar.AppendLine("    JOIN Cajas ON  Movimientos_Cajas.Cod_Caj = Cajas.Cod_Caj")
            loComandoSeleccionar.AppendLine("    JOIN Conceptos ON Movimientos_Cajas.Cod_Con = Conceptos.Cod_Con")
            loComandoSeleccionar.AppendLine("WHERE	Movimientos_Cajas.Fec_Ini BETWEEN @lfFechaDesde AND @lfFechaHasta")
            loComandoSeleccionar.AppendLine("	AND Movimientos_Cajas.Cod_Caj = @lcCajaDesde ")
            loComandoSeleccionar.AppendLine("   AND Movimientos_Cajas.Status        IN  (" & lcParametro2Desde & ")")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("MCL_rMCajas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvMCL_rMCajas.ReportSource = loObjetoReporte

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
' JJD: 21/02/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' GCR: 27/03/09: Ajustes al diseño
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
' CMS:  04/07/09: Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'
' JJD:  29/03/10: Inclusion de los datos de Tipo de Movimiento y Nombre de Cajas
'-------------------------------------------------------------------------------------------'
' CMS:  10/05/10: Se elimino la union con la tabla bancos cuando para el tercer select el cual 
'				  tiene como restriccion: Movimientos_Cajas.Tipo          IN  ('Tarjeta')
'-------------------------------------------------------------------------------------------'
' CMS:  11/05/09: Se corrigio la palabra Tickets por Ticket
'-------------------------------------------------------------------------------------------'
' MAT:  24/08/11 : Ajuste del Select y de la vista de diseño.
'-------------------------------------------------------------------------------------------'