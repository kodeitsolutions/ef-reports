'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCCobrar_cCobros"
'-------------------------------------------------------------------------------------------'
Partial Class rCCobrar_cCobros

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

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
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim loComandoSeleccionar As New StringBuilder()
			
            loComandoSeleccionar.AppendLine("SELECT")
            loComandoSeleccionar.AppendLine("		Retenciones_Documentos.Documento	AS Doc_Ret, ")
            loComandoSeleccionar.AppendLine("		Retenciones_Documentos.Doc_Ori	AS Doc_Ori, ")
            loComandoSeleccionar.AppendLine("		Retenciones_Documentos.Tip_Ori	AS Tip_Ori, ")
          	loComandoSeleccionar.AppendLine(" 		SUM(CASE	WHEN Retenciones_Documentos.Clase = 'IMPUESTO' THEN Retenciones_Documentos.Mon_Ret ELSE  0.00 END) AS  Mon_Ret_IVA, ")
			loComandoSeleccionar.AppendLine(" 		SUM(CASE	WHEN Retenciones_Documentos.Clase = 'ISLR' THEN Retenciones_Documentos.Mon_Ret ELSE  0.00 END) AS  Mon_Ret_ISLR, ")
			loComandoSeleccionar.AppendLine(" 		SUM(CASE	WHEN Retenciones_Documentos.Clase = 'PATENTE' THEN Retenciones_Documentos.Mon_Ret ELSE  0.00 END) AS  Mon_Ret_PATENTE ")
			loComandoSeleccionar.AppendLine("INTO	#tablaCobros1")
			loComandoSeleccionar.AppendLine("FROM   Cuentas_Cobrar,Renglones_Cobros,Retenciones_Documentos,Clientes")
			loComandoSeleccionar.AppendLine("WHERE	Cuentas_Cobrar.Documento	BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta & "")
			loComandoSeleccionar.AppendLine("		AND	Cuentas_Cobrar.Documento	=	Retenciones_Documentos.Doc_Ori")
			loComandoSeleccionar.AppendLine("		AND	Renglones_Cobros.Documento	=	Retenciones_Documentos.Documento")
			loComandoSeleccionar.AppendLine("		AND	Renglones_Cobros.Renglon	=	Retenciones_Documentos.Renglon")
			loComandoSeleccionar.AppendLine("		AND	Cuentas_Cobrar.Cod_Tip	=	Retenciones_Documentos.Cod_Tip")
			loComandoSeleccionar.AppendLine("		AND	Retenciones_Documentos.Clase IN ('IMPUESTO','ISLR','PATENTE')")
			loComandoSeleccionar.AppendLine("		AND	Cuentas_Cobrar.Cod_Cli		=	Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Fec_Ini		BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Cod_Tip		BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Cod_Cli		BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Cod_Ven		BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Clientes.Cod_Zon			BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Status		IN (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("		AND Clientes.Cod_Tip			BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Clientes.Cod_Cla			BETWEEN " & lcParametro8Desde & " AND " & lcParametro8Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Cod_Tra		BETWEEN " & lcParametro9Desde & " AND " & lcParametro9Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Cod_Mon		BETWEEN " & lcParametro10Desde & " AND " & lcParametro10Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Cod_Suc		BETWEEN " & lcParametro11Desde & " AND " & lcParametro11Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Cod_Rev		BETWEEN " & lcParametro12Desde & " AND " & lcParametro12Hasta & "")
			loComandoSeleccionar.AppendLine(" GROUP BY Retenciones_Documentos.Documento, Retenciones_Documentos.Doc_Ori,Retenciones_Documentos.Tip_Ori") 
			loComandoSeleccionar.AppendLine("")
			
            loComandoSeleccionar.AppendLine("SELECT")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Documento, ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("		Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Control, ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Tip_Doc, ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Tip_Ori,")
            loComandoSeleccionar.AppendLine("		(CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' THEN Cuentas_Cobrar.Mon_Bru *(-1) ELSE Cuentas_Cobrar.Mon_Bru END) AS Mon_Bru, ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("		(CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' THEN Cuentas_Cobrar.Mon_Net *(-1) ELSE Cuentas_Cobrar.Mon_Net END) AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("		(CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' THEN Cuentas_Cobrar.Mon_Sal *(-1) ELSE Cuentas_Cobrar.Mon_Sal END) AS Mon_Sal,  ")
            loComandoSeleccionar.AppendLine("		CASE    ")
            loComandoSeleccionar.AppendLine("			WHEN (DATALENGTH(Cuentas_Cobrar.Comentario) > 1) AND (DATALENGTH(Cuentas_Cobrar.Notas) > 1) THEN '- '+CAST(Cuentas_Cobrar.Comentario AS  VARCHAR(1000))+CHAR(13)+'- '+CAST(Cuentas_Cobrar.Notas AS  VARCHAR(1000)) ")
            loComandoSeleccionar.AppendLine("			WHEN (DATALENGTH(Cuentas_Cobrar.Comentario) > 1) AND (DATALENGTH(Cuentas_Cobrar.Notas) <= 1) THEN '- '+CAST(Cuentas_Cobrar.Comentario AS  VARCHAR(1000))   ")
            loComandoSeleccionar.AppendLine("			WHEN (DATALENGTH(Cuentas_Cobrar.Comentario) <= 1) AND (DATALENGTH(Cuentas_Cobrar.Notas) > 1) THEN '- '+CAST(Cuentas_Cobrar.Notas AS  VARCHAR(1000))   ")
            loComandoSeleccionar.AppendLine("			ELSE ''    ")
            loComandoSeleccionar.AppendLine("		END AS Comentario, ")
            loComandoSeleccionar.AppendLine("		Renglones_Cobros.Documento					AS	Num_Cob, ")
            loComandoSeleccionar.AppendLine("		Renglones_Cobros.Registro					AS	Fec_Cob, ")
            loComandoSeleccionar.AppendLine("		Renglones_Cobros.Mon_Net					AS	Mon_Cob, ")            
            loComandoSeleccionar.AppendLine("		Renglones_Cobros.Mon_Abo					AS	Mon_Abo, ")
            loComandoSeleccionar.AppendLine("		#tablaCobros1.Doc_Ret						AS	Doc_Ret, ")
            loComandoSeleccionar.AppendLine("		#tablaCobros1.Doc_Ori						AS	Origen, ")
			loComandoSeleccionar.AppendLine("		ISNULL (#tablaCobros1.Mon_Ret_IVA,0)		AS Mon_Ret_IVA, ")
			loComandoSeleccionar.AppendLine("		ISNULL (#tablaCobros1.Mon_Ret_ISLR,0)		AS Mon_Ret_ISLR, ")
			loComandoSeleccionar.AppendLine("		ISNULL (#tablaCobros1.Mon_Ret_PATENTE,0)	AS Mon_Ret_PATENTE, ")
            loComandoSeleccionar.AppendLine("		ISNULL (Descuentos_Documentos.Mon_Des,0)	AS	Mon_Des ")
            loComandoSeleccionar.AppendLine("INTO	#tablaCobros")
            loComandoSeleccionar.AppendLine("FROM   Clientes, Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("LEFT JOIN Renglones_Cobros ON (Cuentas_Cobrar.Documento	=	Renglones_Cobros.Doc_Ori AND Cuentas_Cobrar.Cod_Tip	=	Renglones_Cobros.Cod_Tip) ")
            loComandoSeleccionar.AppendLine("LEFT JOIN #tablaCobros1 ON (Renglones_Cobros.Documento	=	#tablaCobros1.Doc_Ret AND Cuentas_Cobrar.Documento	=	#tablaCobros1.Doc_Ori AND #tablaCobros1.Tip_Ori  =  'Cuentas_Cobrar') ")
			loComandoSeleccionar.AppendLine("LEFT JOIN Descuentos_Documentos ON (Renglones_Cobros.Documento	=	Descuentos_Documentos.Documento AND Renglones_Cobros.Doc_Ori	=	Descuentos_Documentos.Doc_Ori AND Cuentas_Cobrar.Documento	=	Descuentos_Documentos.Doc_Ori AND Descuentos_Documentos.Tip_Ori  =  'Cuentas_Cobrar') ")
			loComandoSeleccionar.AppendLine("LEFT JOIN Cobros ON (Renglones_Cobros.Documento = Cobros.Documento  AND Cobros.Automatico = 0) ")    
			loComandoSeleccionar.AppendLine("WHERE	Cuentas_Cobrar.Cod_Cli			=	Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("		AND NOT(Cuentas_Cobrar.Cod_Tip IN ('ISLR', 'RETIVA', 'RETPAT') AND  Cuentas_Cobrar.Tip_Ori IN ('Cobros'))")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Documento	BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Fec_Ini		BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Cod_Tip		BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Cod_Cli		BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Cod_Ven		BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Clientes.Cod_Zon			BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Status		IN (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("		AND Clientes.Cod_Tip			BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Clientes.Cod_Cla			BETWEEN " & lcParametro8Desde & " AND " & lcParametro8Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Cod_Tra		BETWEEN " & lcParametro9Desde & " AND " & lcParametro9Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Cod_Mon		BETWEEN " & lcParametro10Desde & " AND " & lcParametro10Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Cod_Suc		BETWEEN " & lcParametro11Desde & " AND " & lcParametro11Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Cod_Rev		BETWEEN " & lcParametro12Desde & " AND " & lcParametro12Hasta & "")
            loComandoSeleccionar.AppendLine("ORDER BY " & lcOrdenamiento & "")
            loComandoSeleccionar.AppendLine("")
            
            loComandoSeleccionar.AppendLine("SELECT")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Documento			AS	Documento, ")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Cod_Tip			AS	Cod_Tip, ")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Fec_Ini			AS	Fec_Ini, ")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Fec_Fin			AS	Fec_Fin, ")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Cod_Cli			AS	Cod_Cli, ")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Nom_Cli			AS	Nom_Cli, ")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Cod_Ven			AS	Cod_Ven, ")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Cod_Tra			AS	Cod_Tra, ")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Cod_Mon			AS	Cod_Mon, ")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Control			AS	Control, ")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Tip_Doc			AS	Tip_Doc, ")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Mon_Bru			AS	Mon_Bru, ")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Mon_Imp1			AS	Mon_Imp, ")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Mon_Net			AS	Mon_Net, ")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Mon_Sal			AS	Mon_Sal,  ")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Tip_Ori			AS  Tip_Ori,  ")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Comentario			AS	Comentario, ")
            loComandoSeleccionar.AppendLine("		Cobros.Status					AS	Est_Cob, ")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Num_Cob			AS	Num_Cob, ")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Fec_Cob			AS	Fec_Cob, ")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Mon_Abo			AS	Mon_Abo, ")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Mon_Des			AS	Mon_Des, ")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Mon_Ret_IVA		AS	Mon_Ret_IVA, ")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Mon_Ret_ISLR		AS	Mon_Ret_ISLR, ")
            loComandoSeleccionar.AppendLine("		#tablaCobros.Mon_Ret_PATENTE	AS	Mon_Ret_PATENTE, ")
            loComandoSeleccionar.AppendLine("		CAST(0 AS DECIMAL)				AS	Neto, ")
            loComandoSeleccionar.AppendLine("		CAST(0 AS DECIMAL)				AS	Saldo, ")
            loComandoSeleccionar.AppendLine("		(CASE WHEN Cobros.Status = 'Confirmado' THEN #tablaCobros.Mon_Net	ELSE 0.00 END)	AS	Mon_Cob")
            loComandoSeleccionar.AppendLine("FROM	#tablaCobros")
			loComandoSeleccionar.AppendLine("LEFT JOIN Cobros ON (Cobros.Documento= #tablaCobros.Num_Cob)")
            loComandoSeleccionar.AppendLine("")

            Dim loServicios As New cusDatos.goDatos
            
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
            
            If laDatosReporte.Tables(0).Rows.Count > 0   Then
					
					Dim Tabla As DataTable  = laDatosReporte.Tables(0)
					Dim Filas As Integer = Tabla.Rows.Count
					Dim Documento_Anterior As String = Tabla.Rows(0).Item("Documento")
					
					Tabla.Rows(0).Item("Neto")	=	Tabla.Rows(0).Item("Mon_Net")
					Tabla.Rows(0).Item("Saldo")	=	Tabla.Rows(0).Item("Mon_Sal")
					
					For i As Integer = 1 To Filas-1 
					
						If	(Documento_Anterior = Tabla.Rows(i).Item("Documento"))
					
							Tabla.Rows(i).Item("Neto")	=	0
							Tabla.Rows(i).Item("Saldo")	=	0
						Else
							Tabla.Rows(i).Item("Neto")	=	Tabla.Rows(i).Item("Mon_Net")
							Tabla.Rows(i).Item("Saldo")	=	Tabla.Rows(i).Item("Mon_Sal")
						End If
						
						Documento_Anterior  = Tabla.Rows(i).Item("Documento")
						
					Next i
					
			End If
            
           'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCCobrar_cCobros", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCCobrar_cCobros.ReportSource = loObjetoReporte

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
' MAT: 14/01/11: Programacion inicial a partir de (CPagar_cPagos)
'-------------------------------------------------------------------------------------------'

