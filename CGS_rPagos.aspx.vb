Imports System.Data
Partial Class CGS_rPagos

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT  Pagos.Fec_Ini                           AS Fecha_Pago, ")
            loComandoSeleccionar.AppendLine("        Pagos.Cod_Pro                           AS Cod_Pro, ")
            loComandoSeleccionar.AppendLine("        Proveedores.Nom_Pro                     AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("        Pagos.Mon_Net                           AS Neto, ")
            loComandoSeleccionar.AppendLine("        Pagos.Comentario                        AS Comentario,")
            loComandoSeleccionar.AppendLine("        Detalles_Pagos.Tip_Ope                  AS Tip_Ope, ")
            loComandoSeleccionar.AppendLine("        Detalles_Pagos.Num_Doc                  AS Num_Doc, ")
            loComandoSeleccionar.AppendLine("        Detalles_Pagos.Cod_Caj                  AS Cod_Caj,")
            loComandoSeleccionar.AppendLine("        COALESCE((SELECT Nom_Caj")
            loComandoSeleccionar.AppendLine("         FROM Cajas")
            loComandoSeleccionar.AppendLine("        WHERE Cod_Caj = Detalles_Pagos.Cod_Caj),'') AS Nom_Caj, ")
            loComandoSeleccionar.AppendLine("        Detalles_Pagos.Cod_Ban                  AS Cod_Ban,  ")
            loComandoSeleccionar.AppendLine("        (SELECT Nom_Ban")
            loComandoSeleccionar.AppendLine("         FROM Bancos")
            loComandoSeleccionar.AppendLine("        WHERE Cod_Ban = Detalles_Pagos.Cod_Ban) AS Nom_Ban,")
            loComandoSeleccionar.AppendLine("        Detalles_Pagos.Cod_Cue                  AS Cod_Cue, ")
            loComandoSeleccionar.AppendLine("        COALESCE((SELECT Nom_Cue")
            loComandoSeleccionar.AppendLine("         FROM Cuentas_Bancarias")
            loComandoSeleccionar.AppendLine("        WHERE Cod_Cue = Detalles_Pagos.Cod_Cue),'') AS Nom_Cue")
            loComandoSeleccionar.AppendLine("FROM    Pagos ")
            loComandoSeleccionar.AppendLine("        JOIN  Detalles_Pagos ")
            loComandoSeleccionar.AppendLine("                ON Pagos.Documento = Detalles_Pagos.Documento ")
            loComandoSeleccionar.AppendLine("        LEFT JOIN Proveedores ")
            loComandoSeleccionar.AppendLine("                ON Pagos.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine(" WHERE     Pagos.Fec_Ini	        Between	" & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			AND Pagos.Cod_Pro           Between	" & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)

            loComandoSeleccionar.AppendLine("UNION ALL ")

            loComandoSeleccionar.AppendLine("SELECT  Ordenes_Pagos.Fec_Ini                    AS Fecha_Pago,")
            loComandoSeleccionar.AppendLine("        Ordenes_Pagos.Cod_Pro                    AS Cod_Pro, ")
            loComandoSeleccionar.AppendLine("        Proveedores.Nom_Pro                      AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("        Ordenes_Pagos.Mon_Net                    AS Neto, ")
            loComandoSeleccionar.AppendLine("        Ordenes_Pagos.Motivo                     AS Comentario,")
            loComandoSeleccionar.AppendLine("        Detalles_OPagos.Tip_Ope                  AS Tip_Ope, ")
            loComandoSeleccionar.AppendLine("        Detalles_OPagos.Num_Doc                  AS Num_Doc, ")
            loComandoSeleccionar.AppendLine("        Detalles_OPagos.Cod_Caj                  AS Cod_Caj,")
            loComandoSeleccionar.AppendLine("        COALESCE((SELECT Nom_Caj")
            loComandoSeleccionar.AppendLine("         FROM Cajas")
            loComandoSeleccionar.AppendLine("        WHERE Cod_Caj = Detalles_OPagos.Cod_Caj),'') AS Nom_Caj, ")
            loComandoSeleccionar.AppendLine("        Detalles_OPagos.Cod_Ban                  AS Cod_Ban,  ")
            loComandoSeleccionar.AppendLine("        (SELECT Nom_Ban")
            loComandoSeleccionar.AppendLine("         FROM Bancos")
            loComandoSeleccionar.AppendLine("        WHERE Cod_Ban = Detalles_OPagos.Cod_Ban) AS Nom_Ban,")
            loComandoSeleccionar.AppendLine("        Detalles_OPagos.Cod_Cue                  AS Cod_Cue, ")
            loComandoSeleccionar.AppendLine("        COALESCE((SELECT Nom_Cue")
            loComandoSeleccionar.AppendLine("         FROM Cuentas_Bancarias")
            loComandoSeleccionar.AppendLine("        WHERE Cod_Cue = Detalles_OPagos.Cod_Cue),'') AS Nom_Cue")
            loComandoSeleccionar.AppendLine("FROM    Ordenes_Pagos ")
            loComandoSeleccionar.AppendLine("        JOIN  Detalles_OPagos ")
            loComandoSeleccionar.AppendLine("                ON Ordenes_Pagos.Documento = Detalles_OPagos.Documento ")
            loComandoSeleccionar.AppendLine("        LEFT JOIN Proveedores ")
            loComandoSeleccionar.AppendLine("                ON Ordenes_Pagos.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine(" WHERE     Ordenes_Pagos.Fec_Ini	        Between	" & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			AND Ordenes_Pagos.Cod_Pro           Between	" & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)

            loComandoSeleccionar.AppendLine("ORDER BY Fecha_Pago")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rPagos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPagos_Proveedores.ReportSource = loObjetoReporte

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
' JJD: 06/12/08: Programacion inicial
'-------------------------------------------------------------------------------------------'
' YJP: 22/04/09: Corregir estatus, anexar combo
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
' CMS:  10/08/09: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'
' MAT:  17/03/11: Filtro "Concepto:", Mejora de la vista de diseño
'-------------------------------------------------------------------------------------------'
' RJG:  10/04/12: Se agregó el total de registros.											'
'-------------------------------------------------------------------------------------------'
