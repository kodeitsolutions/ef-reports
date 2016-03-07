'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCPagar_Renglones"
'-------------------------------------------------------------------------------------------'
Partial Class rCPagar_Renglones

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
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
                        
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = cusAplicacion.goReportes.paParametrosIniciales(8)
            Dim lcParametro9Desde As String = cusAplicacion.goReportes.paParametrosIniciales(9)
            

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()



            loComandoSeleccionar.AppendLine(" SELECT	Cuentas_Pagar.Documento, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Tasa, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Status, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Renglon                AS  Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Cod_Art                AS  Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Precio1                AS  Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Can_Art                AS  Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Cod_Uni                AS  Cod_Uni, ")
            'loComandoSeleccionar.AppendLine("           Renglones_Documentos.Mon_Net                AS  Mon_Net, ")
            loComandoSeleccionar.AppendLine("           (Case when Tip_Doc = 'Credito' then Cuentas_Pagar.Mon_Net *(-1) Else Cuentas_Pagar.Mon_Net End) As Mon_Net,  ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Comentario             AS  Comentario, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Articulos.Nom_Art,1,50)           AS  Nom_Art, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Vendedores.Nom_Ven,1,30)          AS  Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Transportes.Nom_Tra,1,30)         AS  Nom_Tra, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,30)        AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Proveedores.Nom_Pro,1,50)         AS  Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Cod_Tip                AS  Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Tipos_Documentos.Nom_Tip,1,40)    AS  Nom_Tip ")
            loComandoSeleccionar.AppendLine(" FROM      Cuentas_Pagar, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos, ")
            loComandoSeleccionar.AppendLine("           Proveedores, ")
            loComandoSeleccionar.AppendLine("           Articulos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Tipos_Documentos, ")
            loComandoSeleccionar.AppendLine("           Transportes ")
            loComandoSeleccionar.AppendLine(" WHERE     Cuentas_Pagar.Documento        =   Renglones_Documentos.Documento ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Tip      =   Renglones_Documentos.Cod_Tip ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Tip      =   Tipos_Documentos.Cod_Tip ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Pro      =   Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Ven      =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_For      =   Formas_Pagos.Cod_For ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Tra      =   Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Art           =   Renglones_Documentos.Cod_Art ")
            loComandoSeleccionar.AppendLine("           And Renglones_Documentos.Origen =   'Compras' ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Documento    Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Fec_Ini      Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Pro      Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Status       IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Tip      Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            
            loComandoSeleccionar.AppendLine("		  AND ((" & lcParametro5Desde & " = 'Si' AND Cuentas_Pagar.Mon_Sal > 0)")
            loComandoSeleccionar.AppendLine("		  OR (" & lcParametro5Desde & " <> 'Si' AND (Cuentas_Pagar.Mon_Sal >= 0 or Cuentas_Pagar.Mon_Sal < 0)))")
            loComandoSeleccionar.AppendLine("         AND Cuentas_Pagar.Cod_Suc between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("         AND " & lcParametro6Hasta)
           
            If lcParametro8Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cuentas_Pagar.Cod_Rev between " & lcParametro7Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cuentas_Pagar.Cod_Rev NOT between " & lcParametro7Desde)
            End If
            loComandoSeleccionar.AppendLine("         AND " & lcParametro7Hasta)
            
            'loComandoSeleccionar.AppendLine(" ORDER BY  Cuentas_Pagar.Cod_Tip, Renglones_Documentos.Renglon")
            loComandoSeleccionar.AppendLine("ORDER BY   Renglones_Documentos.Documento, " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCPagar_Renglones", laDatosReporte)
            
            If lcParametro9Desde.ToString = "Si" Then
                loObjetoReporte.DataDefinition.FormulaFields("Comentario_Notas").Text = 1
                CType(loObjetoReporte.ReportDefinition.ReportObjects("Line1"), CrystalDecisions.CrystalReports.Engine.LineObject).Right = 0
                CType(loObjetoReporte.ReportDefinition.ReportObjects("Line1"), CrystalDecisions.CrystalReports.Engine.LineObject).Left = 0
            Else
                loObjetoReporte.DataDefinition.FormulaFields("Comentario_Notas").Text = 0
                loObjetoReporte.ReportDefinition.ReportObjects("text14").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("text1").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("Comentario1").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("text14").Top = 0
                loObjetoReporte.ReportDefinition.ReportObjects("text1").Top = 0
                loObjetoReporte.ReportDefinition.ReportObjects("Comentario1").Top = 0
                loObjetoReporte.ReportDefinition.Sections(3).Height = 310
                loObjetoReporte.ReportDefinition.Sections("DetailSection2").Height = 3
            End If           

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCPagar_Renglones.ReportSource = loObjetoReporte

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
' JJD: 07/03/08: Codigo inicial.
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro revisión
'-------------------------------------------------------------------------------------------'
' CMS: 13/07/09: Agregar filtro Cop_Tip
'-------------------------------------------------------------------------------------------'
' CMS: 15/07/09: Verificación de Registros.
'                Multiplicación (*-1) al campo Mon_Net
'-------------------------------------------------------------------------------------------'