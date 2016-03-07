'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCPagar_Proveedores_CCSJ"
'-------------------------------------------------------------------------------------------'
Partial Class rCPagar_Proveedores_CCSJ
    Inherits vis2Formularios.frmReporte

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
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = cusAplicacion.goReportes.paParametrosIniciales(10)
            Dim lcParametro11Desde As String = cusAplicacion.goReportes.paParametrosIniciales(11)
            Dim lcParametro13Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(13))
            Dim lcParametro13Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(13))
            Dim lcParametro14Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(14))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.Appendline("  SELECT 	Cuentas_Pagar.Documento, ")
            loComandoSeleccionar.AppendLine("         	Cuentas_Pagar.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("         	Cuentas_Pagar.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("         	Cuentas_Pagar.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("         	Cuentas_Pagar.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("         	Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("         	Proveedores.Rif, ")
            loComandoSeleccionar.AppendLine("         	Proveedores.Telefonos, ")
            loComandoSeleccionar.AppendLine("         	Proveedores.Movil, ")
            loComandoSeleccionar.AppendLine("         	Proveedores.Correo, ")
            loComandoSeleccionar.AppendLine("         	Cuentas_Pagar.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("         	Cuentas_Pagar.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("         	Cuentas_Pagar.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("         	Cuentas_Pagar.Control, ")
            loComandoSeleccionar.AppendLine("         	SUBSTRING(Cuentas_Pagar.FACTURA,1,10) + ' - ' + SUBSTRING(Cuentas_Pagar.Referencia,1,10) AS Referencia, ")
            If lcParametro11Desde.ToString = "Si" Then
                loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Comentario, ")
            Else
                loComandoSeleccionar.AppendLine("		' '	AS Comentario, ")
            End If

            loComandoSeleccionar.AppendLine("			(CASE WHEN Tip_Doc = 'Credito' THEN Cuentas_Pagar.Mon_Bru *(-1) ELSE Cuentas_Pagar.Mon_Bru End) As Mon_Bru,  ")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Tip_Doc = 'Credito' THEN Cuentas_Pagar.Mon_Net *(-1) ELSE Cuentas_Pagar.Mon_Net End) As Mon_Net,  ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Tip_Doc = 'Credito' THEN Cuentas_Pagar.Mon_Sal *(-1) ELSE Cuentas_Pagar.Mon_Sal End) As Mon_Sal  ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpDocumentos ")
            loComandoSeleccionar.AppendLine(" FROM Tipos_Proveedores, Cuentas_Pagar ")
            loComandoSeleccionar.AppendLine(" JOIN Proveedores ON (Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro) ")
            loComandoSeleccionar.AppendLine(" JOIN Vendedores ON (Cuentas_Pagar.Cod_Ven = Vendedores.Cod_Ven)")
            loComandoSeleccionar.AppendLine(" JOIN Transportes ON (Cuentas_Pagar.Cod_Tra = Transportes.Cod_Tra) ")
            loComandoSeleccionar.AppendLine(" JOIN Monedas ON (Cuentas_Pagar.Cod_Mon = Monedas.Cod_Mon) ")
            loComandoSeleccionar.AppendLine(" WHERE			Tipos_Proveedores.Cod_Tip = Proveedores.Cod_Tip ")


            'Proveedores.Cod_Pro()


            loComandoSeleccionar.AppendLine("         	AND Cuentas_Pagar.Documento BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("         	AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("         	AND Cuentas_Pagar.Cod_Pro BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("         	AND Cuentas_Pagar.Cod_Ven BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("         	AND Cuentas_Pagar.Status IN ( " & lcParametro4Desde & " ) ")
            loComandoSeleccionar.AppendLine("         	AND Cuentas_Pagar.Cod_Tra BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("         	AND Cuentas_Pagar.Cod_Mon BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("         	AND Cuentas_Pagar.Cod_Tip BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("         	AND Cuentas_Pagar.Cod_Suc between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("         	AND Tipos_Proveedores.Cod_Tip Between " & lcParametro13Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro13Hasta)

            If lcParametro10Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cuentas_Pagar.Cod_Rev Between " & lcParametro9Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cuentas_Pagar.Cod_Rev NOT Between " & lcParametro9Desde)
            End If

            loComandoSeleccionar.AppendLine("         AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("		  AND ((" & lcParametro14Desde & " = 'Si' AND Cuentas_Pagar.Mon_Sal > 0)")
            loComandoSeleccionar.AppendLine("		  OR (" & lcParametro14Desde & " <> 'Si' AND (Cuentas_Pagar.Mon_Sal >= 0 or Cuentas_Pagar.Mon_Sal < 0)))")

            loComandoSeleccionar.AppendLine("ORDER BY    Cuentas_Pagar.Cod_Pro, " & lcOrdenamiento)

            'loComandoSeleccionar.AppendLine(" FULL JOIN Cuentas_Clientes ON (Cuentas_Clientes.Cod_Reg = Proveedores.Cod_Pro) ")
            'loComandoSeleccionar.AppendLine(" JOIN Bancos ON (Cuentas_Clientes.Cod_Ban = Bancos.Cod_Ban) ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpDocumentos.*, ")
            loComandoSeleccionar.AppendLine("         	Cuentas_Clientes.Tip_Cue, ")
            loComandoSeleccionar.AppendLine("         	Cuentas_Clientes.Num_Cue, ")
            loComandoSeleccionar.AppendLine("         	Cuentas_Clientes.Cod_Ban ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpDocumentos2 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpDocumentos LEFT JOIN Cuentas_Clientes ")
            loComandoSeleccionar.AppendLine("           ON  #tmpDocumentos.Cod_Pro =   Cuentas_Clientes.Cod_Reg ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpDocumentos2.*, ")
            loComandoSeleccionar.AppendLine("           Bancos.Nom_Ban ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpDocumentos2 LEFT JOIN Bancos ")
            loComandoSeleccionar.AppendLine("           ON  #tmpDocumentos2.Cod_Ban =   Bancos.Cod_Ban ")
            loComandoSeleccionar.AppendLine(" WHERE     SUBSTRING(#tmpDocumentos2.Comentario,1,9) <> 'REEMBOLSO' ")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCPagar_Proveedores_CCSJ", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCPagar_Proveedores_CCSJ.ReportSource = loObjetoReporte


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
' JJD: 22/09/08: Programacion inicial
'-------------------------------------------------------------------------------------------'
' CMS: 22/04/09: Estandarización del código y Corrección del estatus
'-------------------------------------------------------------------------------------------'
' CMS: 15/07/09: Metodo de Ordenamiento.
'                Verificación de Registros.
'                Multiplicación (*-1) al campo Mon_Net, Mon_Sal, Mon_Bru
'-------------------------------------------------------------------------------------------'
' CMS: 21/09/09: Se agregaron los filtros ¿Solo Con Saldo?, Sucursal, Revisión, Tipo de Revisión,
'				 Comentario
'-------------------------------------------------------------------------------------------'
' JJD: 12/10/09: Colocación de totales generales y contador de documentos.					'
'-------------------------------------------------------------------------------------------'
' RJG: 13/10/09: Corrección de contador de documentos. Totales generales intercalados.		'
'-------------------------------------------------------------------------------------------'
' MAT: 27/05/11: Ajuste del Select y Mejora de la vista de Diseño							'
'-------------------------------------------------------------------------------------------'
' MAT: 08/07/11: Ajuste Provisional para mon_sal > 0.01	(Stratos)							'
'-------------------------------------------------------------------------------------------'
' MAT: 08/07/11: Eliminación del filtro con Saldo, No aplica								'
'-------------------------------------------------------------------------------------------'
' RJG: 18/04/12: Eliminación del ajuste provisional mon_sal > 0.01 (No aplica).				'
'-------------------------------------------------------------------------------------------'
' JJD: 27/09/13: Adecuacion a formato de CCSJ
'-------------------------------------------------------------------------------------------'
' JJD: 04/10/13: Inclusion de los datos de Numero de Factura
'-------------------------------------------------------------------------------------------'
