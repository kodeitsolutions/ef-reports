Imports System.Data
Partial Class rDepositos_Clientes
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
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Depositos.Documento, ")
            loComandoSeleccionar.AppendLine("			Depositos.Num_Dep, ")
            loComandoSeleccionar.AppendLine("			Depositos.Status, ")
            loComandoSeleccionar.AppendLine("			Depositos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Depositos.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("			Depositos.Mon_Efe, ")
            loComandoSeleccionar.AppendLine("			Depositos.Mon_Che, ")
            loComandoSeleccionar.AppendLine("			Depositos.Mon_Tar, ")
            loComandoSeleccionar.AppendLine("			Depositos.Mon_Otr, ")
            loComandoSeleccionar.AppendLine("			Depositos.Mon_Net, ")
            loComandoSeleccionar.AppendLine("			Depositos.Cod_Con, ")
            loComandoSeleccionar.AppendLine("			Depositos.Comentario, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Bancarias.Num_Cue, ")
            loComandoSeleccionar.AppendLine("			Conceptos.Nom_Con, ")
            loComandoSeleccionar.AppendLine("			Cobros.Documento            AS  Cobro, ")
            loComandoSeleccionar.AppendLine("			Cobros.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("			Clientes.Nom_Cli ")
            loComandoSeleccionar.AppendLine(" FROM		Depositos, ")
            loComandoSeleccionar.AppendLine("			Renglones_Depositos, ")
            loComandoSeleccionar.AppendLine("			Cobros, ")
            loComandoSeleccionar.AppendLine("			Clientes, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Bancarias, ")
            loComandoSeleccionar.AppendLine("			Conceptos ")
            loComandoSeleccionar.AppendLine(" WHERE		Depositos.Documento             =	Renglones_Depositos.Documento ")
            loComandoSeleccionar.AppendLine("			And Depositos.Cod_Cue		    =	Cuentas_Bancarias.Cod_Cue ")
            loComandoSeleccionar.AppendLine("			And Cobros.Documento            =	Renglones_Depositos.Doc_Ori ")
            loComandoSeleccionar.AppendLine("			And Renglones_Depositos.Tip_Ori =   'cobros' ")
            loComandoSeleccionar.AppendLine("			And Cobros.Cod_Cli              =	Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("			And Depositos.Cod_Con		    =	Conceptos.Cod_Con ")
            loComandoSeleccionar.AppendLine("			And Depositos.Documento		    Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			And Depositos.Fec_Ini		    Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			And Cobros.Cod_Cli              Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			And Depositos.Cod_Cue		    Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			And Depositos.Cod_Con		    Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			And Depositos.Cod_Mon		    Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("			And Depositos.Status		    IN ( " & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("			And Cobros.Documento            Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("			And Depositos.Cod_rev		    Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("			And Depositos.Cod_Suc		    Between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro9Hasta)

            'loComandoSeleccionar.AppendLine(" ORDER BY  16, 1, 15 ") Depositos.Documento, Depositos.Num_Dep, Depositos.Fec_Ini
            loComandoSeleccionar.AppendLine("ORDER BY   Cobros.Cod_Cli, " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString(), "curReportes")


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rDepositos_Clientes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrDepositos_Clientes.ReportSource = loObjetoReporte

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
' JJD: 14/03/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro revisión
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
' CMS:  11/08/09: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'