'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRGlobal_IVARetenidoV"
'-------------------------------------------------------------------------------------------'
Partial Class rRGlobal_IVARetenidoV
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()




            loComandoSeleccionar.AppendLine("SELECT			Retenciones_Documentos.Mon_Bas			AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Ret			AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("				Clientes.Cod_Cli						AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("				Clientes.Nom_Cli						AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("				Clientes.Rif							AS Rif,")
            loComandoSeleccionar.AppendLine("				Clientes.Nit							AS Nit,")
            loComandoSeleccionar.AppendLine("				CAST(Clientes.Dir_Fis AS CHAR(200))	AS Direccion")
            loComandoSeleccionar.AppendLine("INTO			#tmpRetencionesISLR")
            loComandoSeleccionar.AppendLine("FROM			Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("		LEFT JOIN	Cobros")
            loComandoSeleccionar.AppendLine("			ON	Cobros.documento = Cuentas_Cobrar.Doc_Ori")
            loComandoSeleccionar.AppendLine("		JOIN	Retenciones_Documentos ")
            loComandoSeleccionar.AppendLine("			ON	Retenciones_Documentos.Documento = Cobros.Documento")
            loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.doc_des = Cuentas_Cobrar.Documento")
            loComandoSeleccionar.AppendLine("		JOIN	Renglones_Cobros ")
            loComandoSeleccionar.AppendLine("			ON	Renglones_Cobros.Documento = Cobros.Documento")
            loComandoSeleccionar.AppendLine("			AND Renglones_Cobros.Doc_Ori = Retenciones_Documentos.Doc_Ori")
            loComandoSeleccionar.AppendLine("		LEFT JOIN	Cuentas_Cobrar AS Documentos								")
            loComandoSeleccionar.AppendLine("			ON	Documentos.Documento = Renglones_Cobros.Doc_Ori				")
            loComandoSeleccionar.AppendLine("			AND	Documentos.Cod_Tip = Renglones_Cobros.Cod_Tip 				")
            loComandoSeleccionar.AppendLine("		JOIN	Clientes")
            loComandoSeleccionar.AppendLine("			ON	Clientes.Cod_Cli = Cuentas_Cobrar.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE			Cuentas_Cobrar.Cod_Tip = 'RETIVA'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Cobrar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Cobrar.Tip_Ori = 'Cobros'")


            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Cli BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Tip BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Cla BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Per BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro4Hasta)

            loComandoSeleccionar.AppendLine("UNION ALL		") 

            loComandoSeleccionar.AppendLine("SELECT			Retenciones_Documentos.Mon_Bas			AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Ret			AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("				Clientes.Cod_Cli						AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("				Clientes.Nom_Cli						AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("				Clientes.Rif							AS Rif,")
            loComandoSeleccionar.AppendLine("				Clientes.Nit							AS Nit,")
            loComandoSeleccionar.AppendLine("				CAST(Clientes.Dir_Fis AS CHAR(200))	AS Direccion")
            loComandoSeleccionar.AppendLine("FROM			Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("		JOIN	Cuentas_Cobrar AS Documentos ")
            loComandoSeleccionar.AppendLine("			ON	Documentos.documento = Cuentas_Cobrar.Doc_Ori")
            loComandoSeleccionar.AppendLine("			AND Documentos.Cod_Tip = Cuentas_Cobrar.Cla_Ori")
            loComandoSeleccionar.AppendLine("		JOIN	Retenciones_Documentos")
            loComandoSeleccionar.AppendLine("			ON	Retenciones_Documentos.Doc_Des = Cuentas_Cobrar.Documento")
            loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.Doc_Ori = Cuentas_Cobrar.Doc_Ori")
            loComandoSeleccionar.AppendLine("		JOIN	Clientes ")
            loComandoSeleccionar.AppendLine("			ON	Clientes.Cod_Cli = Cuentas_Cobrar.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE			Cuentas_Cobrar.Cod_Tip = 'RETIVA'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Cobrar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Cobrar.Tip_Ori = 'Cuentas_Cobrar'")

            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Cli BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Tip BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Cla BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Per BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro4Hasta)

            loComandoSeleccionar.AppendLine("SELECT		SUM(Base_Retencion) AS Mon_Bas,")
            loComandoSeleccionar.AppendLine("			SUM(Monto_Retenido) AS Mon_NetRet,")
            loComandoSeleccionar.AppendLine("			Cod_Cli,")
            loComandoSeleccionar.AppendLine("			Nom_Cli,")
            loComandoSeleccionar.AppendLine("			Rif					AS Rif_Cli,")
            loComandoSeleccionar.AppendLine("			Nit,")
            loComandoSeleccionar.AppendLine("			Direccion")
            loComandoSeleccionar.AppendLine("FROM		#tmpRetencionesISLR")
            loComandoSeleccionar.AppendLine("GROUP BY	Cod_Cli, Nom_Cli, Rif, Nit, Direccion")

            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rRGlobal_IVARetenidoV", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrRGlobal_IVARetenidoV.ReportSource = loObjetoReporte


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
' RJG:  01/08/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
