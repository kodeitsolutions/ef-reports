'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fReclamos_Sugerecias"
'-------------------------------------------------------------------------------------------'
Partial Class fReclamos_Sugerecias
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()


			loComandoSeleccionar.AppendLine(" SELECT")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Documento, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Status, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Fec_Ini, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Fec_Fin, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Fec_Reg, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Fec_Cie, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Tip_Doc, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Cod_Tip, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Cod_Cli, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Cod_Ven, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Cod_Ope, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Res_Cli, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Res_Emp, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Cuerpo, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Motivo, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Solucion, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Cierre, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Seguimiento, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Otros, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Articulos, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Tipo, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Clase, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Grupo, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Valor1, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Valor2, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Valor3, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Valor4, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Valor5, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Comentario, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Nivel, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Prioridad, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Logico1, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Logico2, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Logico3, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Logico4, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Logico5, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Caracter1, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Caracter2, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Caracter3, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Caracter4, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Caracter5, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Numerico1, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Numerico2, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Numerico3, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Numerico4, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Numerico5, ")
			loComandoSeleccionar.AppendLine(" 			Clientes.Nom_Cli, ")
			loComandoSeleccionar.AppendLine(" 			Vendedores.Nom_Ven, ")
			loComandoSeleccionar.AppendLine(" 			Operador.Nom_Ven AS Nom_Ope, ")
			loComandoSeleccionar.AppendLine(" 			Renglones_Reclamos.Renglon,  ")
			loComandoSeleccionar.AppendLine(" 			Renglones_Reclamos.Cod_Art,  ")
			loComandoSeleccionar.AppendLine(" 			Renglones_Reclamos.Nom_Art,  ")
			loComandoSeleccionar.AppendLine(" 			Renglones_Reclamos.Cod_Uni,   ")
			loComandoSeleccionar.AppendLine(" 			Renglones_Reclamos.Can_Art1,  ")
			loComandoSeleccionar.AppendLine(" 			Renglones_Reclamos.Cod_Ven AS Cod_Ven_Renglon,  ")
			loComandoSeleccionar.AppendLine(" 			Vendedor_Renglon.Nom_Ven AS Nom_Ven_Renglon, ")
			loComandoSeleccionar.AppendLine(" 			Tipos_Reclamos.Nom_Tip")
			loComandoSeleccionar.AppendLine(" FROM Reclamos")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Renglones_Reclamos ON Renglones_Reclamos.Documento = Reclamos.Documento")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Tipos_Reclamos ON Tipos_Reclamos.Cod_Tip = Reclamos.Cod_Tip")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Clientes ON Clientes.Cod_Cli = Reclamos.Cod_Cli")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Vendedores ON Vendedores.Cod_Ven = Reclamos.Cod_Ven")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Vendedores AS Operador ON Operador.Cod_Ven = Reclamos.Cod_Ven")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Vendedores AS Vendedor_Renglon ON Vendedor_Renglon.Cod_Ven = Reclamos.Cod_Ven")
			loComandoSeleccionar.AppendLine(" WHERE")	   			
            loComandoSeleccionar.AppendLine("           " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            
            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
            
           	'--------------------------------------------------'
			' Carga la imagen del logo en cusReportes            '
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

            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fReclamos_Sugerecias", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfReclamos_Sugerecias.ReportSource = loObjetoReporte

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

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' CMS:  17/09/09 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 18/03/10: Se aplico el metodo carga de imagen 
'-------------------------------------------------------------------------------------------'
