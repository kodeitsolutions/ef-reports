'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rReportes_Fecha_Creacion"
'-------------------------------------------------------------------------------------------'
Partial Class rReportes_Fecha_Creacion
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))

            'Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT	Factory_Global.dbo.reportes.cod_rep, ")
            loComandoSeleccionar.AppendLine("		Factory_Global.dbo.reportes.Nom_rep, ")
            loComandoSeleccionar.AppendLine("		Factory_Global.dbo.reportes.Tipo, ")
            loComandoSeleccionar.AppendLine("		Factory_Global.dbo.reportes.Registro, ")
            loComandoSeleccionar.AppendLine("		Factory_Global.dbo.reportes.Usu_cre, ")
            loComandoSeleccionar.AppendLine("		COALESCE(Factory_Global.dbo.usuarios.Nom_usu,'') nom_usu, ")
            loComandoSeleccionar.AppendLine("		SUM (CASE Factory_Global.dbo.reportes.Tipo ")
            loComandoSeleccionar.AppendLine("				WHEN 'Reporte' then 1 ")
            loComandoSeleccionar.AppendLine("				else 0 ")
            loComandoSeleccionar.AppendLine("			END) OVER()			reportes, ")
            loComandoSeleccionar.AppendLine("		SUM (CASE Factory_Global.dbo.reportes.Tipo ")
            loComandoSeleccionar.AppendLine("				WHEN 'formato' then 1 ")
            loComandoSeleccionar.AppendLine("				else 0 ")
            loComandoSeleccionar.AppendLine("			END) OVER()			formatos ")
            loComandoSeleccionar.AppendLine("FROM   Factory_Global.dbo.reportes ")
            loComandoSeleccionar.AppendLine("   LEFT JOIN Factory_Global.dbo.usuarios ON Factory_Global.dbo.reportes.Usu_cre = Factory_Global.dbo.usuarios.cod_usu ")
            loComandoSeleccionar.AppendLine("WHERE  Factory_Global.dbo.reportes.Registro BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("   AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("   AND Factory_Global.dbo.reportes.Usu_cre BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("   AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("   AND Factory_Global.dbo.reportes.Tipo IN(" & lcParametro2Desde & ")")
            loComandoSeleccionar.AppendLine("   AND Factory_Global.dbo.reportes.Modulo BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("   AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("   AND Factory_Global.dbo.reportes.Opcion BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("   AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("   AND Factory_Global.dbo.reportes.status IN(" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("   AND Factory_Global.dbo.reportes.Sistema IN(" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("ORDER BY Factory_Global.dbo.reportes.registro DESC")
            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos

            cusDatos.goDatos.pcNombreAplicativoExterno = "Framework"

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes_Fecha_Creacion")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rReportes_Fecha_Creacion", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrReportes_Fecha_Creacion.ReportSource = loObjetoReporte

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
' EAG:  29/09/15 : Codigo inicial
'-------------------------------------------------------------------------------------------'
