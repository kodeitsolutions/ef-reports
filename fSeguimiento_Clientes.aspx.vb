'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fSeguimiento_Clientes"
'-------------------------------------------------------------------------------------------'
Partial Class fSeguimiento_Clientes
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

			Dim lcNombreTabla As String = ""
			Dim lcCodigo As String = ""
			Dim lcNombre As String = ""
			Dim lcJoin As String = ""
			Dim lcJoin2 As String = ""
			Dim lcTitulo As String = ""
			Dim lcEncabezado As String = ""
            Dim loComandoSeleccionar As New StringBuilder()

			lcNombreTabla = System.Text.RegularExpressions.Regex.Match(cusAplicacion.goFormatos.pcCondicionPrincipal, "seguimientos\.tipo='([a-zA-Z0-9_]+)'",RegexOptions.IgnoreCase).Groups(1).Value.ToString
			
	lcNombreTabla = lcNombreTabla.ToLower

	lcTitulo = "Nombre:"
	
			Select Case lcNombreTabla
					Case "articulos"
					
							lcCodigo = " 				Articulos.Cod_Art As Codigo,		"
							lcNombre = " 				Articulos.Nom_Art As Nombre,		"
							lcJoin	 = " 	JOIN Articulos ON Articulos.Cod_Art = Seguimientos.cod_reg	 "
							lcTitulo = "Artículo:"
							lcEncabezado = "en Artículos"
							
					Case "clientes"
					
							lcCodigo = " 				Clientes.Cod_Cli As Codigo,		"
							lcNombre = " 				Clientes.Nom_Cli As Nombre,		"
							lcJoin	 = " 	JOIN Clientes ON Clientes.Cod_Cli = Seguimientos.Cod_Reg	 "
							lcTitulo = "Cliente:"
							lcEncabezado = "en Clientes"
							 
					Case "proveedores"
					
							lcCodigo = " 				Proveedores.Cod_Pro As Codigo,		"
							lcNombre = " 				Proveedores.Nom_Pro As Nombre,		"
							lcJoin	 = " 	JOIN Proveedores ON Proveedores.Cod_Pro = Seguimientos.cod_reg	 "
							lcTitulo = "Proveedor:"
							lcEncabezado = "en Proveedores"
					
					Case "prospectos"
					
							lcCodigo = " 				prospectos.Cod_Pro As Codigo,		"
							lcNombre = " 				prospectos.Nom_Pro As Nombre,		"
							lcJoin	 = " 	JOIN prospectos ON prospectos.Cod_Pro = Seguimientos.cod_reg	 "
							lcTitulo = "Prospecto:"
							lcEncabezado = "en Prospectos"
								 
					Case "vendedores"
					
							lcCodigo = " 				Vendedores.Cod_Ven As Codigo,		"
							lcNombre = " 				Vendedores.Nom_Ven As Nombre,		"
							'lcJoin	 = ""
							lcTitulo = "Vendedor:"
							lcEncabezado = "en Vendedores"
								 
					Case "movimientos_cajas"
					
							lcCodigo = " 				Movimientos_Cajas.Documento As Codigo,		"
							lcNombre = " 				(Movimientos_Cajas.Cod_Caj +' '+  Cajas.Nom_Caj) As Nombre,		"
							lcJoin	 = " 	JOIN Movimientos_Cajas ON Movimientos_Cajas.Documento = Seguimientos.cod_reg	 "
							lcJoin2	 = " 	JOIN Cajas ON Cajas.Cod_Caj = Movimientos_Cajas.Cod_Caj	 "
							lcTitulo = "Movimiento de Caja:"
							lcEncabezado = "en Movimientos de Cajas"							
									 
					Case "movimientos_cuentas"
					
							lcCodigo = " 				Movimientos_Cuentas.Documento As Codigo,		"
							lcNombre = " 				(Movimientos_Cuentas.Cod_Cue +' '+  Cuentas_Bancarias.Nom_Cue) As Nombre,		"
							lcJoin	 = " 	JOIN Movimientos_Cuentas ON Movimientos_Cuentas.Documento = Seguimientos.cod_reg	 "
							lcJoin2	 = " 	JOIN Cuentas_Bancarias ON Cuentas_Bancarias.Cod_Cue = Movimientos_Cuentas.Cod_Cue	 "
							lcTitulo = "Movimiento de Cuenta:"
							lcEncabezado = "en Movimientos de Cuentas"

					Case "usuarios_globales"
					
							lcCodigo = " 				Usuarios.Cod_Usu As Codigo,		"
							lcNombre = " 				Usuarios.Nom_Usu As Nombre,		"
							lcJoin	 = " 	JOIN Factory_Global.dbo.Usuarios AS USuarios ON Usuarios.Cod_Usu collate Modern_Spanish_CI_AS = Seguimientos.cod_reg collate Modern_Spanish_CI_AS"
							lcTitulo = "Usuario Global:"
							lcEncabezado = "en Usuarios Globales"							
							
					Case "sucursales"
					
							lcCodigo = " 				Sucursales.Cod_Suc As Codigo,		"
							lcNombre = " 				Sucursales.Nom_Suc As Nombre,		"
							lcJoin	 = " 	JOIN Sucursales ON Sucursales.Cod_Suc = Seguimientos.cod_reg	 "
							lcTitulo = "Sucursal:"
							lcEncabezado = "en Sucursales"
							
					Case "competencia"
					
							lcCodigo = " 				Competencia.Cod_Com As Codigo,		"
							lcNombre = " 				Competencia.Nom_Com As Nombre,		"
							lcJoin	 = " 	JOIN Competencia ON Competencia.Cod_Com = Seguimientos.cod_reg	 "
							lcTitulo = "Competencia:"
							lcEncabezado = "en Competencia"
					
			End Select 


			loComandoSeleccionar.AppendLine(" 	SELECT 		")
			loComandoSeleccionar.AppendLine(" 				Seguimientos.Cod_Ven,		")
			loComandoSeleccionar.AppendLine(" 				Seguimientos.Cod_Ope,		")
			loComandoSeleccionar.AppendLine(" 				Seguimientos.Fec_Ini,		")
			loComandoSeleccionar.AppendLine(" 				Seguimientos.Hor_Ini,		")
			loComandoSeleccionar.AppendLine(" 				Seguimientos.Lugar,		")
			loComandoSeleccionar.AppendLine(" 				Seguimientos.Contacto,		")
			loComandoSeleccionar.AppendLine(" 				Seguimientos.Accion,		")
			loComandoSeleccionar.AppendLine(" 				Seguimientos.Notas,		")
			loComandoSeleccionar.AppendLine(" 				Seguimientos.Comentario,		")
			
			loComandoSeleccionar.AppendLine(" 				"& lcCodigo &"		")
			loComandoSeleccionar.AppendLine(" 				"& lcNombre &"		")
			loComandoSeleccionar.AppendLine(" 				Operadores.Nom_Ven AS Nom_Ope,		")
			loComandoSeleccionar.AppendLine(" 				Vendedores.Nom_Ven		")
			loComandoSeleccionar.AppendLine(" 	FROM Seguimientos		")
			
			loComandoSeleccionar.AppendLine(" 				"& lcJoin &"		")
			loComandoSeleccionar.AppendLine(" 				"& lcJoin2 &"		")
			
			loComandoSeleccionar.AppendLine(" 	JOIN Vendedores ON Vendedores.Cod_Ven = Seguimientos.Cod_Ven ")
			loComandoSeleccionar.AppendLine(" 	JOIN Vendedores AS Operadores ON Operadores.Cod_Ven = Seguimientos.Cod_Ope ")
			loComandoSeleccionar.AppendLine(" WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			
			'me.mEscribirConsulta(loComandoSeleccionar.ToString)
			'me.mEscribirConsulta(lcNombreTabla)

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
					  
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fSeguimiento_Clientes", laDatosReporte)
             
            CType(loObjetoReporte.ReportDefinition.ReportObjects("text21"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "Seguimiento " & lcEncabezado
            CType(loObjetoReporte.ReportDefinition.ReportObjects("text10"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcTitulo
            
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfSeguimiento_Clientes.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' CMS:  19/03/10 : Codigo inicial															'
'-------------------------------------------------------------------------------------------'
' CMS:  26/06/10 : Se modifico para mostrar el seguimiento de las siguientes tablas:		'
'		articulos, clientes, proveedores, vendedores, movimientos_cajas, movimientos_cuentas'
'-------------------------------------------------------------------------------------------'

