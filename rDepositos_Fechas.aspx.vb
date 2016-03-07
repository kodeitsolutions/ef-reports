Imports System.Data
Partial Class rDepositos_Fechas
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim lcComandoSeleccionar As New StringBuilder()

            lcComandoSeleccionar.AppendLine(" SELECT	Depositos.Documento, ")
            lcComandoSeleccionar.AppendLine("			Depositos.Num_Dep, ")
            lcComandoSeleccionar.AppendLine("			Depositos.Status, ")
            lcComandoSeleccionar.AppendLine("			Depositos.Fec_Ini, ")
            lcComandoSeleccionar.AppendLine("			Depositos.Cod_Cue, ")
            lcComandoSeleccionar.AppendLine("			Depositos.Mon_Efe, ")
            lcComandoSeleccionar.AppendLine("			Depositos.Mon_Che, ")
            lcComandoSeleccionar.AppendLine("			Depositos.Mon_Tar, ")
            lcComandoSeleccionar.AppendLine("			Depositos.Mon_Otr, ")
            lcComandoSeleccionar.AppendLine("			Depositos.Mon_Net, ")
            lcComandoSeleccionar.AppendLine("			Depositos.Cod_Con, ")
            lcComandoSeleccionar.AppendLine("			Depositos.Comentario, ")
            lcComandoSeleccionar.AppendLine("			Cuentas_Bancarias.Num_Cue, ")
            lcComandoSeleccionar.AppendLine("			Conceptos.Nom_Con ")
            lcComandoSeleccionar.AppendLine(" FROM		Depositos, ")
            lcComandoSeleccionar.AppendLine("			Cuentas_Bancarias, ")
            lcComandoSeleccionar.AppendLine("			Conceptos ")
            lcComandoSeleccionar.AppendLine(" WHERE		Depositos.Cod_Cue			=	Cuentas_Bancarias.Cod_Cue ")
            lcComandoSeleccionar.AppendLine("			And Depositos.Cod_Con		=	Conceptos.Cod_Con ")
            lcComandoSeleccionar.AppendLine("			And Depositos.Documento		Between " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("			And " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("			And Depositos.Fec_Ini		Between " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("			And " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("			And Depositos.Num_Dep		Between " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("			And " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("			And Depositos.Cod_Cue		Between " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("			And " & lcParametro3Hasta)
            lcComandoSeleccionar.AppendLine("			And Depositos.Cod_Con		Between " & lcParametro4Desde)
            lcComandoSeleccionar.AppendLine("			And " & lcParametro4Hasta)
            lcComandoSeleccionar.AppendLine("			And Depositos.Cod_Mon		Between " & lcParametro5Desde)
            lcComandoSeleccionar.AppendLine("			And " & lcParametro5Hasta)
            lcComandoSeleccionar.AppendLine("			And Depositos.Status		IN ( " & lcParametro6Desde & ")")
            lcComandoSeleccionar.AppendLine("			And Depositos.Cod_rev		Between " & lcParametro7Desde)
            lcComandoSeleccionar.AppendLine("			And " & lcParametro7Hasta)
            lcComandoSeleccionar.AppendLine("			And Depositos.Cod_Suc		Between " & lcParametro8Desde)
            lcComandoSeleccionar.AppendLine("			And " & lcParametro8Hasta)
            'lcComandoSeleccionar.AppendLine(" ORDER BY  Depositos.Fec_Ini, Depositos.Documento, Depositos.Num_Dep ")
            lcComandoSeleccionar.AppendLine("ORDER BY    CONVERT(nchar(30), Depositos.Fec_Ini,112), " & lcOrdenamiento)

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rDepositos_Fechas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrDepositos_Fechas.ReportSource = loObjetoReporte

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
' JJD: 14/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 14/03/09: Se ajusto el filtro del Status del Documento
'-------------------------------------------------------------------------------------------'
' GCR: 27/03/09: Ajustes al diseño.
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro revisión
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
' CMS:  11/08/09: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'