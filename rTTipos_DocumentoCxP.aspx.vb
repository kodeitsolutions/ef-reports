'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTTipos_DocumentoCxP"
'-------------------------------------------------------------------------------------------'
Partial Class rTTipos_DocumentoCxP
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("  SELECT		Cuentas_Pagar.cod_tip, ")
            loComandoSeleccionar.AppendLine(" 				Tipos_Documentos.nom_tip, ")
            loComandoSeleccionar.AppendLine(" 				count (Cuentas_Pagar.cod_tip) as Cant_Doc, ")
            loComandoSeleccionar.AppendLine(" 				SUM(Cuentas_Pagar.mon_bas1) AS mon_bas1, ")
            loComandoSeleccionar.AppendLine(" 				SUM(Cuentas_Pagar.mon_imp1) AS mon_imp1, ")
            loComandoSeleccionar.AppendLine(" 				SUM(Cuentas_Pagar.mon_net) AS mon_net, ")
            loComandoSeleccionar.AppendLine("               SUM(Case when Cuentas_Pagar.Tip_Doc = 'Credito' then Cuentas_Pagar.Mon_Sal *(-1) Else Cuentas_Pagar.Mon_Sal End) As Mon_Sal  ")

            loComandoSeleccionar.AppendLine(" FROM			Cuentas_Pagar, ")
            loComandoSeleccionar.AppendLine(" 				Tipos_Documentos, ")
            loComandoSeleccionar.AppendLine(" 				Proveedores ")

            loComandoSeleccionar.AppendLine(" WHERE 		Cuentas_Pagar.Cod_tip = Tipos_Documentos.Cod_tip")
            loComandoSeleccionar.AppendLine(" AND  			Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro")

            loComandoSeleccionar.AppendLine(" AND 			Cuentas_Pagar.Documento BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" AND 			" & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" AND 			Cuentas_Pagar.cod_tip BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" AND 			" & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" AND 			Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" AND 			" & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" AND 			Cuentas_Pagar.Cod_Pro BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" AND 			" & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" AND 			Cuentas_Pagar.Cod_ven BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" AND 			" & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" AND 			Cuentas_Pagar.Cod_Mon BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" AND 			" & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" AND 			Proveedores.Cod_Zon BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" AND 			" & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine(" AND 			Cuentas_Pagar.Status IN (" & lcParametro7Desde & ")")
            loComandoSeleccionar.AppendLine(" AND 			Cuentas_Pagar.Cod_rev BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine(" AND 			" & lcParametro9Hasta)
            
            loComandoSeleccionar.AppendLine("  AND ((" & lcParametro6Desde & " = 'Si' AND Cuentas_Pagar.Mon_Sal > 0)")
            loComandoSeleccionar.AppendLine("  OR (" & lcParametro6Desde & " <> 'Si' AND (Cuentas_Pagar.Mon_Sal >= 0 or Cuentas_Pagar.Mon_Sal < 0)))")
            loComandoSeleccionar.Append(" GROUP BY		Cuentas_Pagar.Cod_tip, Tipos_Documentos.nom_tip ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            'loComandoSeleccionar.Append( " ORDER BY		Cuentas_Pagar.Cod_tip, Tipos_Documentos.nom_tip  ") 

		
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTTipos_DocumentoCxP", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTTipos_DocumentoCxP.ReportSource = loObjetoReporte

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
' CMS: 09/05/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro Revisión
'-------------------------------------------------------------------------------------------'
' YJP: 13/04/10: Corrección del filtro de fechas
'-------------------------------------------------------------------------------------------'
