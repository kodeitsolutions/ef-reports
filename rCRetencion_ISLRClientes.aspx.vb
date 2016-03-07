'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCRetencion_ISLRClientes"
'-------------------------------------------------------------------------------------------'
Partial Class rCRetencion_ISLRClientes
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


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()




            loComandoSeleccionar.AppendLine("SELECT			Cuentas_Cobrar.Tip_Ori				AS Tipo_Origen,")
            loComandoSeleccionar.AppendLine("				Cuentas_Cobrar.Fec_Ini				AS Fecha_Retencion,")
            loComandoSeleccionar.AppendLine("				Cuentas_Cobrar.Doc_Ori				AS Numero_Cobro,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Cod_Tip		AS Tipo_Documento,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Doc_Ori		AS Numero_Documento,")
            loComandoSeleccionar.AppendLine("				Renglones_Cobros.Mon_Net			AS Monto_Documento,")
            loComandoSeleccionar.AppendLine("				Renglones_Cobros.Mon_Abo			AS Monto_Abonado,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Sus		AS Sustraendo_Retenido,")
            loComandoSeleccionar.AppendLine("				RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret	AS Concepto,")
            
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("				Clientes.Cod_Cli					AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("				Clientes.Nom_Cli					AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("				Clientes.Rif						AS Rif,")
            loComandoSeleccionar.AppendLine("				Clientes.Nit						AS Nit,")
            loComandoSeleccionar.AppendLine("				Clientes.Dir_Fis					AS Direccion")
            loComandoSeleccionar.AppendLine("FROM			Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("		JOIN	Cobros ON Cobros.documento = Cuentas_Cobrar.Doc_Ori")
            loComandoSeleccionar.AppendLine("		JOIN	Retenciones_Documentos ON Retenciones_Documentos.Documento = Cobros.Documento")
            loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.Doc_Des = Cuentas_Cobrar.Documento")
            loComandoSeleccionar.AppendLine("		JOIN	Renglones_Cobros ON Renglones_Cobros.Documento = Cobros.Documento")
            loComandoSeleccionar.AppendLine("			AND Renglones_Cobros.Doc_Ori = Retenciones_Documentos.Doc_Ori")
            loComandoSeleccionar.AppendLine("		JOIN	Clientes ON Clientes.Cod_Cli = Cuentas_Cobrar.Cod_Cli")
            loComandoSeleccionar.AppendLine("		LEFT JOIN Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("WHERE			Cuentas_Cobrar.Cod_Tip = 'ISLR'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Cobrar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Cobrar.Tip_Ori = 'Cobros'")
											 
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Cli BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Cobros.Cod_Mon BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Cobros.Cod_Suc BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro3Hasta)

            loComandoSeleccionar.AppendLine("UNION ALL		")
            
            loComandoSeleccionar.AppendLine("SELECT			Cuentas_Cobrar.Tip_Ori				AS Tipo_Origen,")
            loComandoSeleccionar.AppendLine("				Cuentas_Cobrar.Fec_Ini				AS Fecha_Retencion,")
            loComandoSeleccionar.AppendLine("				''									AS Numero_Cobro,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Cod_Tip		AS Tipo_Documento,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Doc_Ori		AS Numero_Documento,")
            loComandoSeleccionar.AppendLine("				Documentos.Mon_Net					AS Monto_Documento,")
            loComandoSeleccionar.AppendLine("				Documentos.Mon_Net					AS Monto_Abonado,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Sus		AS Sustraendo_Retenido,")
            loComandoSeleccionar.AppendLine("				RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret	AS Concepto,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("				Clientes.Cod_Cli					AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("				Clientes.Nom_Cli					AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("				Clientes.Rif						AS Rif,")
            loComandoSeleccionar.AppendLine("				Clientes.Nit						AS Nit,")
            loComandoSeleccionar.AppendLine("				Clientes.Dir_Fis					AS Direccion")
            loComandoSeleccionar.AppendLine("FROM			Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("		JOIN	Cuentas_Cobrar AS Documentos ON Documentos.documento = Cuentas_Cobrar.Doc_Ori")
            loComandoSeleccionar.AppendLine("			AND Documentos.Cod_Tip = Cuentas_Cobrar.Cla_Ori")
            loComandoSeleccionar.AppendLine("		JOIN	Retenciones_Documentos ON Retenciones_Documentos.Doc_Des = Cuentas_Cobrar.Documento")
            loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.Doc_Ori = Cuentas_Cobrar.Doc_Ori")
            loComandoSeleccionar.AppendLine("		JOIN	Clientes ON Clientes.Cod_Cli = Cuentas_Cobrar.Cod_Cli")
            loComandoSeleccionar.AppendLine("		LEFT JOIN Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("WHERE			Cuentas_Cobrar.Cod_Tip = 'ISLR'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Cobrar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Cobrar.Tip_Ori = 'Cuentas_Cobrar'")

            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Cli BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Mon BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Suc BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro3Hasta)
            
            loComandoSeleccionar.AppendLine("ORDER BY " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº 0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCRetencion_ISLRClientes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrrCRetencion_ISLRClientes.ReportSource = loObjetoReporte


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
' CMS:  21/05/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS:  28/07/09: Se modofico la consulta de modo que se obtuvieron por separado los
'                 proveedores y los beneficiarios y luego se unieron los resultados.
'                 Verificacion de registros.
'                 Metodo de Ordenamiento
'-------------------------------------------------------------------------------------------'
' CMS:  29/07/09: Se Renonbre de Relación Global de ISLR Relativo a Relación Global de ISLR Retenido
'-------------------------------------------------------------------------------------------'