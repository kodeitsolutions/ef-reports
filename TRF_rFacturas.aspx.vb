'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "TRF_rFacturas"
'-------------------------------------------------------------------------------------------'
Partial Class TRF_rFacturas
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim loConsulta As New StringBuilder()


            loConsulta.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro0Desde)
            loConsulta.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro0Hasta)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT   Facturas.Documento, ")
            loConsulta.AppendLine(" 		Facturas.Status,  ")
            loConsulta.AppendLine(" 		Facturas.Fec_Ini, ")
            loConsulta.AppendLine(" 		Facturas.Fec_Fin, ")
            loConsulta.AppendLine(" 		Facturas.Control, ")
            loConsulta.AppendLine(" 		Facturas.Cod_Cli, ")
            loConsulta.AppendLine(" 		Clientes.Nom_Cli, ")
            loConsulta.AppendLine("         Facturas.Mon_Imp1,")
            loConsulta.AppendLine("         Facturas.Mon_Net, ")
            loConsulta.AppendLine("         Facturas.Mon_Sal, ")
            loConsulta.AppendLine("		    CONCAT(CONVERT(VARCHAR(12),CAST(@ldFecha_Desde AS DATE),103), ' - ',  CONVERT(VARCHAR(12),CAST(@ldFecha_Hasta AS DATE),103))	AS Fecha")
            loConsulta.AppendLine("FROM Facturas")
            loConsulta.AppendLine(" JOIN Clientes ON Clientes.Cod_Cli = Facturas.Cod_Cli")
            loConsulta.AppendLine("WHERE Facturas.Status = 'Confirmado'")
            loConsulta.AppendLine(" AND Facturas.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loConsulta.AppendLine("ORDER BY Facturas.Documento")


            Dim loServicios As New cusDatos.goDatos
            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loConsulta.ToString, "curReportes")

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("TRF_rFacturas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvTRF_rFacturas.ReportSource = loObjetoReporte


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
' KDE: 27/09/2017: Codigo inicial
'-------------------------------------------------------------------------------------------'
