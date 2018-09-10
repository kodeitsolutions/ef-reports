'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports System.Globalization

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rPago_NominaBOD_TXT"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rPago_NominaBOD_TXT
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim lcRifEmpresaSQL As String = goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcRifEmpresa)
            Dim lcCodigoEmpresaSQL As String = goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcCodigo)

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @lcNomDesde AS VARCHAR(10) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcNomHasta AS VARCHAR(10) = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha AS DATE = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT ROW_NUMBER() OVER(ORDER BY Registros.Cedula ASC) AS Num,*")
            loComandoSeleccionar.AppendLine("FROM(")
            loComandoSeleccionar.AppendLine("	SELECT  SUBSTRING(Trabajadores.Cedula,1,1)                                  AS Cedula,")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Trabajadores.Cedula,2, LEN(RTRIM(Trabajadores.Cedula)))   AS Cod_Tra,")
            loComandoSeleccionar.AppendLine("			LOWER(RTRIM(Trabajadores.Nom_Tra))                                  AS Nom_Tra,")
            loComandoSeleccionar.AppendLine("			Trabajadores.Num_Cue                                                AS Num_Cue,")
            loComandoSeleccionar.AppendLine("			SUBSTRING(Trabajadores.Num_Cue,1,4)					                AS Cod_Banco,")
            loComandoSeleccionar.AppendLine("			CAST(Pagos.Mon_Net AS DECIMAL(28,2))				                AS Mon_Net,")
            loComandoSeleccionar.AppendLine("			@ldFecha											                AS Nomina,")
            loComandoSeleccionar.AppendLine("			COALESCE(RTRIM(Prop_Numero_Contrato.Val_Car), '')                   AS Numero_Contrato")
            loComandoSeleccionar.AppendLine("	FROM Trabajadores")
            loComandoSeleccionar.AppendLine("		JOIN (SELECT SUM(Recibos.Mon_Net) AS Mon_Net,")
            loComandoSeleccionar.AppendLine("					 Recibos.Cod_Tra,")
            loComandoSeleccionar.AppendLine("					 Recibos.Documento,")
            loComandoSeleccionar.AppendLine("					 Recibos.Fec_Ini,")
            loComandoSeleccionar.AppendLine("					 Recibos.Comentario")
            loComandoSeleccionar.AppendLine("			  FROM Recibos")
            loComandoSeleccionar.AppendLine("			  WHERE Recibos.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("					AND Recibos.Doc_Ori BETWEEN @lcNomDesde AND @lcNomHasta")
            loComandoSeleccionar.AppendLine("					AND Recibos.Tip_Ori = 'Nominas'")
            loComandoSeleccionar.AppendLine("			  GROUP BY Recibos.Cod_Tra, Recibos.Documento, Recibos.Fec_Ini, Recibos.Comentario")
            loComandoSeleccionar.AppendLine("			  ) AS Pagos")
            loComandoSeleccionar.AppendLine("			ON  Pagos.Cod_Tra = Trabajadores.Cod_Tra")
            loComandoSeleccionar.AppendLine("		LEFT JOIN Campos_Propiedades Prop_Numero_Contrato ON  Prop_Numero_Contrato.Cod_Reg = " & lcCodigoEmpresaSQL)
            loComandoSeleccionar.AppendLine("		    AND Prop_Numero_Contrato.Origen = 'Empresas'")
            loComandoSeleccionar.AppendLine("			AND Prop_Numero_Contrato.Cod_Pro = 'NUMCONBOD'")
            loComandoSeleccionar.AppendLine("	WHERE Pagos.Mon_Net > 0")
            loComandoSeleccionar.AppendLine("		AND Trabajadores.Cod_Ban = 'BOD' ")
            loComandoSeleccionar.AppendLine("		AND Trabajadores.Num_Cue <> '' ")
            loComandoSeleccionar.AppendLine("		AND Trabajadores.Tip_Pag = 'Transferencia' ) Registros")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")


            '-------------------------------------------------------------------------------------------------------
            ' Genera el archivo de texto    
            '-------------------------------------------------------------------------------------------------------
            Dim loLimpiarRIF As New Regex("[^a-zA-Z0-9]")
            Const Ceros As String = "0000000000000000000"

            Dim lcRif As String = loLimpiarRIF.Replace(goEmpresa.pcRifEmpresa, "")
            Dim lcFechaNomina As String = Strings.Format(CDate(cusAplicacion.goReportes.paParametrosIniciales(1)), "yyyyMMdd")
            Dim lcContrato As String = laDatosReporte.Tables(0).Rows(0).Item("Numero_Contrato")
            Dim lnTotalFilas As Integer = laDatosReporte.Tables(0).Rows.Count
            Dim lnTotalMonto As Decimal = 0
            For i As Integer = 0 To laDatosReporte.Tables(0).Rows.Count - 1
                lnTotalMonto += laDatosReporte.Tables(0).Rows(i).Item("Mon_Net")
            Next
            Dim lcTotalMonto As String = Strings.Right(Ceros & lnTotalMonto.ToString().Replace(".", ""), 17)

            Dim loSalida As New StringBuilder()

            loSalida.Append("01Nomina" & Space(14))
            loSalida.Append(lcRif)
            loSalida.Append(lcContrato)
            loSalida.Append("0")
            loSalida.Append(lcFechaNomina)
            loSalida.Append(lcFechaNomina)
            loSalida.Append(Strings.Right(Ceros & lnTotalFilas.ToString(), 6))
            loSalida.Append(lcTotalMonto)
            loSalida.Append("VEB")
            loSalida.AppendLine()

            For Each loFila As DataRow In laDatosReporte.Tables(0).Rows

                Dim lcCedula As String = CStr(loFila("Cedula")).Trim() & Strings.Right(Ceros & CStr(loFila("Cod_Tra")).Trim(), 9)
                Dim lcTrabajador As String = Strings.Left(CStr(loFila("Nom_Tra")) & Space(50), 60)
                lcTrabajador = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(lcTrabajador)

                Dim lcRenglon As String = Strings.Right(Ceros & Strings.Format(CDec(loFila("Num")), "0"), 9)

                Dim lcCuenta As String = CStr(loFila("Num_Cue")).Trim()
                Dim lcBanco As String = CStr(loFila("Cod_Banco")).Trim()
                Dim lcMonto_Neto As String = Strings.Right(Ceros & loLimpiarRIF.Replace(CDec(loFila("Mon_Net")), ""), 15)
                
                loSalida.Append("02")
                loSalida.Append(lcCedula)
                loSalida.Append(lcTrabajador)

                loSalida.Append(lcRenglon)
                loSalida.Append(Strings.Left("NOMINA" & Space(30), 30))

                loSalida.Append("CTA")
                loSalida.Append(lcCuenta)
                loSalida.Append(lcBanco)
                loSalida.Append(lcFechaNomina)
                loSalida.Append(lcMonto_Neto)
                loSalida.Append("VEB")
                loSalida.Append("000000000000000")

                loSalida.Append(Space(40))
                loSalida.Append("00000000000")

                loSalida.AppendLine()

            Next loFila

            'Me.mEscribirConsulta(loSalida.ToString())

            '-------------------------------------------------------------------------------------------------------
            ' Envia la salida a pantalla en un archivo descargable.
            '-------------------------------------------------------------------------------------------------------
            Me.Response.Clear()
            Me.Response.ContentEncoding = System.Text.Encoding.GetEncoding(1252)
            Me.Response.AppendHeader("content-disposition", "attachment; filename=NominaBOD_" & lcFechaNomina & ".txt")
            Me.Response.ContentType = "text/plain"
            Me.Response.Write(loSalida.ToString())
            Me.Response.End()

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 01/08/09: Codigo inicial (creado a partir de rCRetencion_IVAProveedores)				'
'-------------------------------------------------------------------------------------------'
' JJD: 08/06/11: Ajustes en la generacion de monto base y porcentaje.
'-------------------------------------------------------------------------------------------'
' MAT: 10/06/11: Generación del archivo del comprobante con cero retención
'-------------------------------------------------------------------------------------------'
' RJG: 16/04/13: Corrección de bug: algunos documentos aparecían duplicados.                
'-------------------------------------------------------------------------------------------'
' RJG: 04/06/13: Se agregó la eliminación de todos los carecteres no válidos del RIF.       '
'-------------------------------------------------------------------------------------------'
