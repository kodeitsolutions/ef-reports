'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rDocumentos_Vendedores"
'-------------------------------------------------------------------------------------------'
Partial Class rDocumentos_Vendedores
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
	
			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
			Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
			Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
			Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
			Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
			Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
			Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
			Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
			Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
			Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
			Dim lcParametro7Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
			
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
			Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine( "  SELECT		Cuentas_Cobrar.cod_ven, " )
			loComandoSeleccionar.AppendLine( " 				Vendedores.nom_ven, " )
			loComandoSeleccionar.AppendLine( " 				SUM(Cuentas_Cobrar.mon_bas1) AS mon_bas1, " )
			loComandoSeleccionar.AppendLine( " 				SUM(Cuentas_Cobrar.mon_imp1) AS mon_imp1, " ) 
			loComandoSeleccionar.AppendLine( " 				SUM(Cuentas_Cobrar.mon_net) AS mon_net, " ) 
			'loComandoSeleccionar.AppendLine("               SUM(Case when Cuentas_Cobrar.Tip_Doc = 'Credito' then Cuentas_Cobrar.Mon_Sal *(-1) Else Cuentas_Cobrar.Mon_Sal End) As Mon_Sal,  ")
			loComandoSeleccionar.AppendLine("               SUM(Cuentas_Cobrar.Mon_Sal) As Mon_Sal,  ")
		  	loComandoSeleccionar.AppendLine( "				Cuentas_Cobrar.Tip_doc " )
		  	loComandoSeleccionar.AppendLine( " FROM			Clientes, " )
			loComandoSeleccionar.AppendLine( " 				Cuentas_Cobrar, " )
			loComandoSeleccionar.AppendLine( " 				Vendedores, " )
			loComandoSeleccionar.AppendLine( " 				Tipos_Documentos " )
			
			loComandoSeleccionar.AppendLine( " WHERE 		Cuentas_Cobrar.Cod_Ven = Vendedores.Cod_Ven " )
            loComandoSeleccionar.AppendLine("               AND 			Cuentas_Cobrar.Cod_tip = Tipos_Documentos.Cod_tip ")
            loComandoSeleccionar.AppendLine("               AND 			Cuentas_Cobrar.Cod_cli = Clientes.Cod_cli ")
            loComandoSeleccionar.AppendLine("               AND 			Cuentas_Cobrar.Documento BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("               AND 			" & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("               AND 			Cuentas_Cobrar.cod_tip BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("               AND 			" & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("               AND 			Cuentas_Cobrar.Fec_Ini BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("               AND 			" & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("               AND 			Cuentas_Cobrar.Cod_cli BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("               AND 			" & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("               AND 			Cuentas_Cobrar.Cod_ven BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("               AND 			" & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("               AND 			Cuentas_Cobrar.Cod_Mon BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("               AND 			" & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("               AND 			Vendedores.cod_zon BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("               AND 			" & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("               AND             ((" & lcParametro7Desde & " = 'Si' AND Cuentas_Cobrar.Mon_Sal > 0)")
            loComandoSeleccionar.AppendLine("               OR              (" & lcParametro7Desde & " <> 'Si' AND (Cuentas_Cobrar.Mon_Sal >= 0 or Cuentas_Cobrar.Mon_Sal < 0)))")
            loComandoSeleccionar.AppendLine("               AND 			Cuentas_Cobrar.Status IN (" & lcParametro8Desde & ")")
            loComandoSeleccionar.AppendLine("               AND             Cuentas_Cobrar.Cod_Rev between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("               AND             " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY		Cuentas_Cobrar.Tip_doc, Cuentas_Cobrar.Cod_Ven, Vendedores.nom_ven ")
            loComandoSeleccionar.AppendLine("ORDER BY      Cuentas_Cobrar.Tip_doc, " & lcOrdenamiento)
            'loComandoSeleccionar.AppendLine( " ORDER BY		Cuentas_Cobrar.Cod_Ven, Vendedores.nom_ven ") 
'me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rDocumentos_Vendedores", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrDocumentos_Vendedores.ReportSource =	 loObjetoReporte

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
' YJP: 06/05/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' CMS:  15/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' CMS:  05/05/10: Se agrupo por tip_doc
'-------------------------------------------------------------------------------------------'