'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rMCajas_Fechas_Resumido"
'-------------------------------------------------------------------------------------------'
Partial Class rMCajas_Fechas_Resumido
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
            Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            
            Dim lcParametro8Desde As String = CStr(cusAplicacion.goReportes.paParametrosIniciales(8)).Trim()
            Dim lcOrigen		  As String = CStr(cusAplicacion.goReportes.paParametrosIniciales(9)).Trim()
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(lcOrigen)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            lcParametro4Hasta = IIf(lcParametro4Hasta = "'zzzzzzz'", "'zzzzzzz'", lcParametro4Desde)

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT		Movimientos_Cajas.Documento, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Cod_Caj, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Comentario, ")
            loComandoSeleccionar.AppendLine("           SPACE(10) AS  Cod_Ban, ")
            loComandoSeleccionar.AppendLine("           SPACE(30) AS  Nom_Ban, ")
            loComandoSeleccionar.AppendLine("           SPACE(10) AS  Cod_Tar, ")
            loComandoSeleccionar.AppendLine("           SPACE(40) AS  Nom_Tar, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Status, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Referencia, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Cod_Con, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Movimientos_Cajas.Tip_Ori,1,20) AS Tip_Ori, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Tipo, ")
            loComandoSeleccionar.AppendLine("           Conceptos.Nom_Con, ")
            loComandoSeleccionar.AppendLine("           Cajas.Nom_Caj   AS  Nombre, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Movimientos_Cajas.Tipo = 'Efectivo' THEN 'Efectivo' ELSE 'Ticket' END) AS  Tip_Mov ")
            loComandoSeleccionar.AppendLine("FROM		Movimientos_Cajas ")
            loComandoSeleccionar.AppendLine(" 	JOIN	Cajas ON  Movimientos_Cajas.Cod_Caj		=   Cajas.Cod_Caj")
            loComandoSeleccionar.AppendLine(" 	JOIN	Conceptos ON Movimientos_Cajas.Cod_Con   =   Conceptos.Cod_Con")
            loComandoSeleccionar.AppendLine("WHERE		Movimientos_Cajas.Tipo          IN  ('Efectivo','Ticket') ")
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.Documento     BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.Fec_Ini       BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.Cod_Caj       BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.Cod_Mon       BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.Referencia    BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.Status        IN  (" & LcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.cod_rev       BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.Cod_Suc       BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro7Hasta)
            If (lcParametro8Desde = "Si") Then 
				loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Automatico = 0")
            End If
            If (lcOrigen <> "") Then 
				loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Tip_Ori = " & lcParametro9Desde)
            End If
            loComandoSeleccionar.AppendLine("   ")
            loComandoSeleccionar.AppendLine("   ")
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("   ")
            loComandoSeleccionar.AppendLine("   ")
            loComandoSeleccionar.AppendLine("SELECT    Movimientos_Cajas.Documento, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Cod_Caj, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Comentario, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Cod_Ban, ")
            loComandoSeleccionar.AppendLine("           Bancos.Nom_Ban, ")
            loComandoSeleccionar.AppendLine("           SPACE(10) AS  Cod_Tar, ")
            loComandoSeleccionar.AppendLine("           SPACE(40) AS  Nom_Tar, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Status, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Referencia, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Cod_Con, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Tipo, ")
            loComandoSeleccionar.AppendLine("           Conceptos.Nom_Con, ")
            loComandoSeleccionar.AppendLine("           Cajas.Nom_Caj		AS  Nombre, ")
            loComandoSeleccionar.AppendLine("           'Cheque'			AS  Tip_Mov ")
            loComandoSeleccionar.AppendLine("FROM		Movimientos_Cajas ")
            loComandoSeleccionar.AppendLine(" 	JOIN	Cajas ON  Movimientos_Cajas.Cod_Caj		=   Cajas.Cod_Caj")
            loComandoSeleccionar.AppendLine(" 	JOIN	Conceptos ON Movimientos_Cajas.Cod_Con   =   Conceptos.Cod_Con")
            loComandoSeleccionar.AppendLine(" 	JOIN	Bancos ON Movimientos_Cajas.Cod_Ban      =   Bancos.Cod_Ban")
            loComandoSeleccionar.AppendLine("WHERE		Movimientos_Cajas.Tipo          IN  ('CHEQUe') ")
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.Documento     BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.Fec_Ini       BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.Cod_Caj       BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.Cod_Mon       BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.Referencia    BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.Status        IN  (" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.cod_rev       BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.Cod_Suc       BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro7Hasta)
            If (lcParametro8Desde = "Si") Then 
				loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Automatico = 0")
            End If
            If (lcOrigen <> "") Then 
				loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Tip_Ori = " & lcParametro9Desde)
            End If
            loComandoSeleccionar.AppendLine("   ")
            loComandoSeleccionar.AppendLine("   ")
            loComandoSeleccionar.AppendLine(" UNION ALL ")
            loComandoSeleccionar.AppendLine("   ")
            loComandoSeleccionar.AppendLine("   ")
            loComandoSeleccionar.AppendLine(" SELECT    Movimientos_Cajas.Documento, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Cod_Caj, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Comentario, ")
			loComandoSeleccionar.AppendLine("           SPACE(10) AS  Cod_Ban, ")
            loComandoSeleccionar.AppendLine("           SPACE(10) AS  Nom_Ban, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Cod_Tar, ")
            loComandoSeleccionar.AppendLine("           Tarjetas.Nom_Tar, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Status, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Referencia, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Cod_Con, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Tipo, ")
            loComandoSeleccionar.AppendLine("           Conceptos.Nom_Con, ")
            loComandoSeleccionar.AppendLine("           Cajas.Nom_Caj    AS  Nombre, ")
            loComandoSeleccionar.AppendLine("           'Tarjeta' AS  Tip_Mov ")
            loComandoSeleccionar.AppendLine("FROM		Movimientos_Cajas ")
            loComandoSeleccionar.AppendLine("	JOIN 	Cajas ON  Movimientos_Cajas.Cod_Caj		=   Cajas.Cod_Caj")
            loComandoSeleccionar.AppendLine("	JOIN 	Conceptos ON Movimientos_Cajas.Cod_Con   =   Conceptos.Cod_Con")
            loComandoSeleccionar.AppendLine("	JOIN 	Tarjetas  ON Movimientos_Cajas.Cod_Tar   =   Tarjetas.Cod_Tar")
            loComandoSeleccionar.AppendLine("WHERE		Movimientos_Cajas.Tipo          IN  ('Tarjeta') ")
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.Documento     BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.Fec_Ini       BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.Cod_Caj       BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.Cod_Mon       BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.Referencia    BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.Status        IN  (" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.cod_rev       BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			AND Movimientos_Cajas.Cod_Suc       BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro7Hasta)
            If (lcParametro8Desde = "Si") Then 
				loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Automatico = 0")
            End If
            If (lcOrigen <> "") Then 
				loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Tip_Ori = " & lcParametro9Desde)
            End If
            loComandoSeleccionar.AppendLine(" ORDER BY  " & lcOrdenamiento & ", Movimientos_Cajas.Documento ASC")

			'me.mEscribirConsulta(loComandoSeleccionar.ToString)
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes          '
            '--------------------------------------------------'
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rMCajas_Fechas_Resumido", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrMCajas_Fechas_Resumido.ReportSource = loObjetoReporte

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
' RJG: 21/02/09: Codigo inicial (A partir de rMCajas_Fechas).								'
'-------------------------------------------------------------------------------------------'
