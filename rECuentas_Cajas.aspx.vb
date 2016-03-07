'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rECuentas_Cajas"
'-------------------------------------------------------------------------------------------'
Partial Class rECuentas_Cajas
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
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro7Desde As String = cusAplicacion.goReportes.paParametrosFinales(7)
            Dim lcParametro8Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("CREATE TABLE #tempSALDOINICIAL(Cod_Caj CHAR(10), Sal_Ini DECIMAL(28,10))")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("INSERT INTO #tempSALDOINICIAL(Cod_Caj, Sal_Ini)")
            loComandoSeleccionar.AppendLine("SELECT		Cod_Caj,")
            loComandoSeleccionar.AppendLine("			SUM(Mon_Deb-Mon_Hab) AS Sal_Ini     ")
            loComandoSeleccionar.AppendLine("FROM		Movimientos_Cajas     ")
            loComandoSeleccionar.AppendLine("WHERE			Fec_Ini < " & lcParametro1Desde & "  ")
            loComandoSeleccionar.AppendLine("	        AND	Cod_Caj		BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("	        AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Ban     BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND	Cod_Con		BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Status      IN ( " & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Cod_Suc     BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            If lcParametro7Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 		AND Cod_Rev BETWEEN " & lcParametro6Desde)
            Else
                loComandoSeleccionar.AppendLine(" 		AND Cod_Rev NOT BETWEEN " & lcParametro6Desde)
            End If

            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Tipo        IN ( " & lcParametro8Desde & ")")
            loComandoSeleccionar.AppendLine("GROUP BY	Cod_Caj     ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		ROW_NUMBER() OVER(PARTITION BY Movimientos_Cajas.Cod_Caj ORDER BY Movimientos_Cajas.Cod_Caj, Movimientos_Cajas.Fec_Ini) As Posicion, ")
            loComandoSeleccionar.AppendLine("			Movimientos_Cajas.Documento, ")
            loComandoSeleccionar.AppendLine("			Movimientos_Cajas.Tip_Doc, ")
            loComandoSeleccionar.AppendLine("			Movimientos_Cajas.Cod_Caj, ")
            loComandoSeleccionar.AppendLine("			Cajas.Nom_Caj, ")
            loComandoSeleccionar.AppendLine("			Movimientos_Cajas.Cod_Ban, ")
            loComandoSeleccionar.AppendLine("			Movimientos_Cajas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Movimientos_Cajas.Status, ")
            loComandoSeleccionar.AppendLine("			Movimientos_Cajas.Referencia, ")
            loComandoSeleccionar.AppendLine("			Movimientos_Cajas.Cod_Con, ")
            loComandoSeleccionar.AppendLine("			Conceptos.Nom_Con,")
            loComandoSeleccionar.AppendLine("			Movimientos_Cajas.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("			Movimientos_Cajas.Comentario,")
            loComandoSeleccionar.AppendLine("			Movimientos_Cajas.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("			Movimientos_Cajas.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("			Movimientos_Cajas.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("			Movimientos_Cajas.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("			Movimientos_Cajas.Tipo, ")
            loComandoSeleccionar.AppendLine("    		COALESCE(#tempSALDOINICIAL.Sal_Ini,0) AS Sal_Ini,	")
            loComandoSeleccionar.AppendLine("    		CAST(0 AS DECIMAL(28,10)) AS Sal_Doc")
            loComandoSeleccionar.AppendLine("INTO		#tempMOVIMIENTO")
            loComandoSeleccionar.AppendLine("FROM		Movimientos_Cajas ")
            loComandoSeleccionar.AppendLine("	JOIN	Cajas ")
            loComandoSeleccionar.AppendLine("		ON  Cajas.Cod_Caj = Movimientos_Cajas.Cod_Caj ")
            loComandoSeleccionar.AppendLine("	JOIN	Conceptos ")
            loComandoSeleccionar.AppendLine("		ON	Conceptos.Cod_Con = Movimientos_Cajas.Cod_Con")
            loComandoSeleccionar.AppendLine("LEFT JOIN #tempSALDOINICIAL")
            loComandoSeleccionar.AppendLine("		ON	#tempSALDOINICIAL.Cod_Caj = Cajas.Cod_Caj")
            loComandoSeleccionar.AppendLine("WHERE		Movimientos_Cajas.Cod_Caj       BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cajas.Fec_Ini       BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cajas.Cod_Ban       BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cajas.Cod_Con       BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cajas.Status        IN ( " & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cajas.Cod_Suc       BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro6Hasta)
            If lcParametro7Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 		AND Movimientos_Cajas.Cod_Rev BETWEEN " & lcParametro6Desde)
            Else
                loComandoSeleccionar.AppendLine(" 		AND Movimientos_Cajas.Cod_Rev NOT BETWEEN " & lcParametro6Desde)
            End If

            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Tipo        IN ( " & lcParametro8Desde & ")")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")


            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UPDATE #tempMOVIMIENTO")
            loComandoSeleccionar.AppendLine("SET		Sal_Doc = Sal_Ini + Acumulado.Movimiento")
            loComandoSeleccionar.AppendLine("FROM	(	")
            loComandoSeleccionar.AppendLine("			SELECT		A.Cod_Caj, ")
            loComandoSeleccionar.AppendLine("						A.Documento, ")
            loComandoSeleccionar.AppendLine("						SUM(B.Mon_Deb - B.Mon_Hab) AS Movimiento")
            loComandoSeleccionar.AppendLine("			FROM		#tempMOVIMIENTO A")
            loComandoSeleccionar.AppendLine("				JOIN	#tempMOVIMIENTO B")
            loComandoSeleccionar.AppendLine("					ON	B.Cod_Caj = A.Cod_Caj")
            loComandoSeleccionar.AppendLine("					AND	B.Posicion <= A.Posicion")
            loComandoSeleccionar.AppendLine("			GROUP BY A.Cod_Caj, A.Documento ")
            loComandoSeleccionar.AppendLine("		) AS Acumulado")
            loComandoSeleccionar.AppendLine("WHERE	#tempMOVIMIENTO.Cod_Caj = Acumulado.Cod_Caj")
            loComandoSeleccionar.AppendLine("	AND	#tempMOVIMIENTO.Documento = Acumulado.Documento")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		#tempMOVIMIENTO.*")
            loComandoSeleccionar.AppendLine("FROM		#tempMOVIMIENTO")
            loComandoSeleccionar.AppendLine("ORDER BY	#tempMOVIMIENTO." & lcOrdenamiento & ", #tempMOVIMIENTO.Fec_Ini ASC")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
			
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rECuentas_Cajas", laDatosReporte)


            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrECuentas_Cajas.ReportSource = loObjetoReporte

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
' MAT:  08/09/11: Código Inicial.															'
'-------------------------------------------------------------------------------------------'
' RJG:  13/10/12: Corrección del cálculo del acumulado por Caja.							'
'-------------------------------------------------------------------------------------------'
