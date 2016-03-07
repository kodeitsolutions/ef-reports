'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rMCuentas_Numeros"
'-------------------------------------------------------------------------------------------'
Partial Class rMCuentas_Numeros
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
	
		    Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = cusAplicacion.goReportes.paParametrosIniciales(8)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
          
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT		Movimientos_Cuentas.Documento, ")
			loComandoSeleccionar.AppendLine(" 				Movimientos_Cuentas.Cod_Tip, " )
            loComandoSeleccionar.AppendLine(" 				Movimientos_Cuentas.Cod_Cue, ")
            loComandoSeleccionar.AppendLine(" 				Movimientos_Cuentas.Referencia, ")
            loComandoSeleccionar.AppendLine(" 				Movimientos_Cuentas.Cod_Ban, ")
            loComandoSeleccionar.AppendLine(" 				Movimientos_Cuentas.Comentario, ")
            loComandoSeleccionar.AppendLine(" 				Bancos.Nom_Ban, ")
			loComandoSeleccionar.AppendLine(" 				Movimientos_Cuentas.Fec_Ini, " )
			loComandoSeleccionar.AppendLine(" 				Movimientos_Cuentas.Status, " )
			loComandoSeleccionar.AppendLine(" 				Movimientos_Cuentas.Referencia, " )
			loComandoSeleccionar.AppendLine(" 				Movimientos_Cuentas.Cod_Con, " )
			loComandoSeleccionar.AppendLine(" 				Movimientos_Cuentas.Cod_Mon, " )
			loComandoSeleccionar.AppendLine(" 				Movimientos_Cuentas.Mon_Deb, " )
			loComandoSeleccionar.AppendLine(" 				Movimientos_Cuentas.Mon_Hab, " )
			loComandoSeleccionar.AppendLine(" 				Movimientos_Cuentas.Tip_Ori, " )
			loComandoSeleccionar.AppendLine(" 				Movimientos_Cuentas.Doc_Ori, " )
			loComandoSeleccionar.AppendLine(" 				Movimientos_Cuentas.Tipo, " )
			loComandoSeleccionar.AppendLine(" 				Tipos_Movimientos.Nom_Tip," ) 
			loComandoSeleccionar.AppendLine(" 				Cuentas_Bancarias.Num_Cue, " )
			loComandoSeleccionar.AppendLine(" 				Conceptos.Nom_Con " )
            'loComandoSeleccionar.AppendLine(" From			Movimientos_Cuentas, " )
            'loComandoSeleccionar.AppendLine("				Tipos_Movimientos, " )
            'loComandoSeleccionar.AppendLine("				Cuentas_Bancarias, " )
            'loComandoSeleccionar.AppendLine("				Conceptos, ")
            'loComandoSeleccionar.AppendLine("				Bancos ")
            'loComandoSeleccionar.AppendLine(" WHERE			Movimientos_Cuentas.Cod_Tip = Tipos_Movimientos.Cod_Tip ")
            'loComandoSeleccionar.AppendLine(" 				AND Movimientos_Cuentas.Cod_Ban = Bancos.Cod_Ban ")
            'loComandoSeleccionar.AppendLine(" 				AND Movimientos_Cuentas.Cod_Cue = Cuentas_Bancarias.Cod_Cue " )
            'loComandoSeleccionar.AppendLine(" 				AND Movimientos_Cuentas.Cod_Con = Conceptos.Cod_Con " )
            'loComandoSeleccionar.AppendLine(" 				AND Movimientos_Cuentas.Documento between " & lcParametro0Desde)

            loComandoSeleccionar.AppendLine(" FROM			Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine(" JOIN Tipos_Movimientos ON Movimientos_Cuentas.Cod_Tip = Tipos_Movimientos.Cod_Tip ")
            'loComandoSeleccionar.AppendLine(" LEFT JOIN Bancos ON Movimientos_Cuentas.Cod_Ban = Bancos.Cod_Ban ")
            loComandoSeleccionar.AppendLine(" JOIN Cuentas_Bancarias ON Movimientos_Cuentas.Cod_Cue = Cuentas_Bancarias.Cod_Cue ")
            loComandoSeleccionar.AppendLine(" LEFT JOIN Bancos ON Cuentas_Bancarias.Cod_Ban = Bancos.Cod_Ban ")
            loComandoSeleccionar.AppendLine(" JOIN Conceptos ON Movimientos_Cuentas.Cod_Con = Conceptos.Cod_Con ")
            loComandoSeleccionar.AppendLine(" WHERE")

            loComandoSeleccionar.AppendLine(" 				Movimientos_Cuentas.Documento between " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Movimientos_Cuentas.Fec_Ini between " & lcParametro1Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Movimientos_Cuentas.Cod_Cue between " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Movimientos_Cuentas.Cod_Mon between " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Movimientos_Cuentas.Status IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine(" 				AND Movimientos_Cuentas.Tipo between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Movimientos_Cuentas.Cod_rev between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Movimientos_Cuentas.Cod_Suc between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY       Cuentas_Bancarias.Num_Cue, " & lcOrdenamiento)
            'loComandoSeleccionar.AppendLine(" ORDER BY		Movimientos_Cuentas.Documento ") 
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


            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rMCuentas_Numeros", laDatosReporte)

            If lcParametro8Desde.ToString = "Si" Then
                loObjetoReporte.DataDefinition.FormulaFields("Comentario_Notas").Text = 1
                CType(loObjetoReporte.ReportDefinition.ReportObjects("Line3"), CrystalDecisions.CrystalReports.Engine.LineObject).Right = 0
                CType(loObjetoReporte.ReportDefinition.ReportObjects("Line3"), CrystalDecisions.CrystalReports.Engine.LineObject).Left = 0
            Else
                loObjetoReporte.DataDefinition.FormulaFields("Comentario_Notas").Text = 0
                loObjetoReporte.ReportDefinition.ReportObjects("text17").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("text18").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("Comentario1").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("text17").Top = 0
                loObjetoReporte.ReportDefinition.ReportObjects("text18").Top = 0
                loObjetoReporte.ReportDefinition.ReportObjects("Comentario1").Top = 0
                loObjetoReporte.ReportDefinition.Sections("Section3").Height = 250
                loObjetoReporte.ReportDefinition.Sections(3).Height = 300
                loObjetoReporte.ReportDefinition.Sections("DetailSection2").Height = 10
            End If


            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrMCuentas_Numeros.ReportSource =	 loObjetoReporte	

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
' JJD: 25/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP:  04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' YJP:  14/05/09: Agregar filtro revisión
'-------------------------------------------------------------------------------------------'
' CMS:  16/06/09: Se agrego el estandar de en el encabezado del documento, Se Agrego el 
'                 filtro: Movimientos_Cuentas.Tipo
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
' CMS:  03/07/09: Metodo de Ordenamiento
'-------------------------------------------------------------------------------------------'
' CMS:  05/08/09: Verificacion de registros, se reescribio la consulta pra poder aplicar
'                 Left Join A la union de bancos y movimientos cuentas
'-------------------------------------------------------------------------------------------'
' CMS:  28/08/09: Filtro Comentario
'         Se modifico la union con la tabla bancos tal que la union es con Cuentas_Bancarias
'-------------------------------------------------------------------------------------------'