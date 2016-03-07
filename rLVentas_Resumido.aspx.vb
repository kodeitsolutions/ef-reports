'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLVentas_Resumido"
'-------------------------------------------------------------------------------------------'
Partial Class rLVentas_Resumido

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = cusAplicacion.goReportes.paParametrosFinales(3)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    CONVERT(nchar(10), Cuentas_Cobrar.Fec_Ini, 103)	AS	Fecha1, ")
            loComandoSeleccionar.AppendLine("           CONVERT(nchar(10), Cuentas_Cobrar.Fec_Ini, 112)	AS	Fecha2, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Documento, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fiscal, ")
            loComandoSeleccionar.AppendLine("			CASE")
            loComandoSeleccionar.AppendLine("				WHEN Cuentas_Cobrar.Cod_Tip = 'N/CR' THEN (Cuentas_Cobrar.Mon_Net * -1) ELSE Cuentas_Cobrar.Mon_Net")
            loComandoSeleccionar.AppendLine("			END AS Mon_Net,")
            loComandoSeleccionar.AppendLine("			CASE")
            loComandoSeleccionar.AppendLine("				WHEN Cuentas_Cobrar.Cod_Tip = 'N/CR' THEN (Cuentas_Cobrar.Mon_Bru * -1) ELSE Cuentas_Cobrar.Mon_Bru")
            loComandoSeleccionar.AppendLine("			END AS Mon_Bru,")
            loComandoSeleccionar.AppendLine("			CASE")
            loComandoSeleccionar.AppendLine("				WHEN Cuentas_Cobrar.Cod_Tip = 'N/CR' THEN ((Cuentas_Cobrar.Mon_Imp1) * -1) ELSE (Cuentas_Cobrar.Mon_Imp1)")
            loComandoSeleccionar.AppendLine("			END AS Mon_Imp,")
            loComandoSeleccionar.AppendLine("			CASE")
            loComandoSeleccionar.AppendLine("				WHEN Cuentas_Cobrar.Cod_Tip = 'N/CR' THEN (Cuentas_Cobrar.Mon_Des * -1) ELSE Cuentas_Cobrar.Mon_Des")
            loComandoSeleccionar.AppendLine("			END AS Mon_Des,")
            loComandoSeleccionar.AppendLine("			CASE")
            loComandoSeleccionar.AppendLine("				WHEN Cuentas_Cobrar.Cod_Tip = 'N/CR' THEN (Cuentas_Cobrar.Mon_Rec * -1) ELSE Cuentas_Cobrar.Mon_Rec")
            loComandoSeleccionar.AppendLine("			END AS Mon_Rec,")
            loComandoSeleccionar.AppendLine("			CASE")
            loComandoSeleccionar.AppendLine("				WHEN Cuentas_Cobrar.Cod_Tip = 'N/CR' THEN ((Cuentas_Cobrar.Mon_Otr1+Cuentas_Cobrar.Mon_Otr2+Cuentas_Cobrar.Mon_Otr3) * -1) ELSE (Cuentas_Cobrar.Mon_Otr1+Cuentas_Cobrar.Mon_Otr2+Cuentas_Cobrar.Mon_Otr3)")
            loComandoSeleccionar.AppendLine("			END AS Mon_Otr")
            loComandoSeleccionar.AppendLine(" INTO  #tmpTemporal ")
            loComandoSeleccionar.AppendLine(" FROM  Cuentas_Cobrar ")
            loComandoSeleccionar.AppendLine(" JOIN Clientes ON Cuentas_Cobrar.Cod_Cli    =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine(" WHERE     Cuentas_Cobrar.Fec_Ini    BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			AND Cuentas_Cobrar.Cod_Tip IN ('FACT','N/CR','N/DB','RETIVA')")
             loComandoSeleccionar.AppendLine("			AND Cuentas_Cobrar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Suc    BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev BETWEEN " & lcParametro2Desde)
            Else
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT BETWEEN " & lcParametro2Desde)
            End If
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)

          
            loComandoSeleccionar.AppendLine(" SELECT    Fecha1, ")
            loComandoSeleccionar.AppendLine("           Fecha2, ")
            loComandoSeleccionar.AppendLine("           MIN(Documento)  AS Doc_Ini, ")
            loComandoSeleccionar.AppendLine("           MAX(Documento)  AS Doc_Fin, ")
            loComandoSeleccionar.AppendLine("           SUM(Mon_Otr)	AS  Mon_Otr, ")
            loComandoSeleccionar.AppendLine("           SUM(CASE WHEN Fiscal = 1 THEN Mon_Des ELSE 0 END)	AS  Mon_DesC, ")
            loComandoSeleccionar.AppendLine("           SUM(CASE WHEN Fiscal = 1 THEN Mon_Rec ELSE 0 END)	AS  Mon_RecC, ")
            loComandoSeleccionar.AppendLine("           SUM(CASE WHEN Fiscal = 1 THEN Mon_Bru ELSE 0 END)	AS  Mon_BruC, ")
            loComandoSeleccionar.AppendLine("           SUM(CASE WHEN Fiscal = 1 THEN Mon_Imp ELSE 0 END)	AS  Mon_ImpC, ")
            loComandoSeleccionar.AppendLine("           SUM(CASE WHEN Fiscal = 0 THEN Mon_Des ELSE 0 END)	AS  Mon_DesNC, ")
            loComandoSeleccionar.AppendLine("           SUM(CASE WHEN Fiscal = 0 THEN Mon_Rec ELSE 0 END)	AS  Mon_RecNC, ")
            loComandoSeleccionar.AppendLine("           SUM(CASE WHEN Fiscal = 0 THEN Mon_Bru ELSE 0 END)	AS  Mon_BruNC, ")
            loComandoSeleccionar.AppendLine("           SUM(CASE WHEN Fiscal = 0 THEN Mon_Imp ELSE 0 END)	AS  Mon_ImpNC, ")
            loComandoSeleccionar.AppendLine("           SUM(Mon_Net) AS  Neto ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpTemporal ")
            loComandoSeleccionar.AppendLine(" GROUP BY  Fecha1, Fecha2 ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos
            
            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLVentas_Resumido", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLVentas_Resumido.ReportSource = loObjetoReporte

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
' JJD: 05/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' AAP:  29/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS:  06/08/09: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'
' CMS:  22/05/10: Filtro Revision y Tipo de revision. 
'-------------------------------------------------------------------------------------------'
' MAT:  05/09/11: Ajuste del Select y ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'