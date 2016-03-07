Imports System.Data
Imports cusAplicacion

Partial Class rDEliminaciones_Registros
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim lcComandoSeleccionar As New StringBuilder()


            lcComandoSeleccionar.AppendLine(" SELECT		Cod_Usu, ")
            lcComandoSeleccionar.AppendLine(" 				Registro,")
            lcComandoSeleccionar.AppendLine(" 				CONVERT(nchar(30), Registro,112) As Registro2,")
            lcComandoSeleccionar.AppendLine(" 				LOWER(Tabla) as Tabla, ")
            lcComandoSeleccionar.AppendLine(" 				Documento, ")
            lcComandoSeleccionar.AppendLine(" 				Codigo, ")
            lcComandoSeleccionar.AppendLine(" 				Detalle, ")
            lcComandoSeleccionar.AppendLine(" 				Equipo ")
            lcComandoSeleccionar.AppendLine(" INTO  #Temp ")
            lcComandoSeleccionar.AppendLine(" FROM			Auditorias ")
            lcComandoSeleccionar.AppendLine(" WHERE         Accion = 'Eliminar' ")
            lcComandoSeleccionar.AppendLine("   			AND Registro between " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine(" 				AND Cod_Usu between " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine(" 				AND " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine(" 				AND Tabla between " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine(" 				AND Opcion between " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
            lcComandoSeleccionar.AppendLine(" 				AND Documento between " & lcParametro4Desde)
            lcComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
            lcComandoSeleccionar.AppendLine(" 				AND Codigo between " & lcParametro5Desde)
            lcComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
            lcComandoSeleccionar.AppendLine(" 				AND Codigo between " & lcParametro6Desde)
            lcComandoSeleccionar.AppendLine(" 				AND " & lcParametro6Hasta)
            'lcComandoSeleccionar.AppendLine(" ORDER BY		Registro ")

            lcComandoSeleccionar.AppendLine(" SELECT * FROM #Temp ")
            lcComandoSeleccionar.AppendLine(" ORDER BY       Cod_Usu, " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rDEliminaciones_Registros", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrDEliminaciones_Registros.ReportSource = loObjetoReporte

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
' CMS: 23/08/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS:  17/08/09: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'
' JJD: 19/12/12: Se incluyo el filtro de la empresa
'-------------------------------------------------------------------------------------------'
