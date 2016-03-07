'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
																	  
'-------------------------------------------------------------------------------------------'
' Inicio de clase "frmActualizacionCompetencia"
'-------------------------------------------------------------------------------------------'
Partial Class frmActualizacionCompetencia
	Inherits vis2formularios.frmActualizacionSimple

#Region "Declaraciones"

#End Region

#Region "Propiedades"

#End Region

#Region "Metodos"

'-------------------------------------------------------------------------------------------'
' Registra los controles cuando se carga la pagina
'-------------------------------------------------------------------------------------------'

	Protected Sub mCargarPagina(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
	
		If Not Page.IsPostBack() then
			
							
			Me.WbcBotonera.pnModoActualizacion = vis3Controles.wbcBotoneraFormularios.enuModoActualizacion.KN_ModoNormal
			
			Me.pcTablaPrincipal = "Competencia"		
			Me.mRegistrarIndice("cod_com")	
			
			Me.mRegistrarControl(Me.txtCod_Com,		"cod_com")
			Me.mRegistrarControl(Me.txtNom_Com,		"nom_com")
			Me.mRegistrarControl(Me.cboStatus,		"status")
			Me.mRegistrarControl(Me.txtRif,			"rif")
			Me.mRegistrarControl(Me.txtNit,			"nit")
			Me.mRegistrarControl(Me.txtDir_Fis,		"dir_fis")
			Me.mRegistrarControl(Me.txtDir_Exa,		"dir_exa")
			Me.mRegistrarControl(Me.txtDir_Otr,		"dir_otr")
			Me.mRegistrarControl(Me.txtTelefonos,	"telefonos")		
			Me.mRegistrarControl(Me.txtDirecto,		"directo")		
			Me.mRegistrarControl(Me.txtFax,			"fax")
			Me.mRegistrarControl(Me.txtMovil,		"movil")			
			Me.mRegistrarControl(Me.txtCorreo,		"correo")	
			Me.mRegistrarControl(Me.txtCorreo2,		"correo2")	
			Me.mRegistrarControl(Me.txtCorreo3,		"correo3")	
			Me.mRegistrarControl(Me.txtWeb,			"web")
			Me.mRegistrarControl(Me.txtKilometros,	"kilometros")
			Me.mRegistrarControl(Me.txtCod_Pai,		"cod_pai")
			Me.mRegistrarControl(Me.txtCod_Est,		"cod_est")
			Me.mRegistrarControl(Me.txtCod_Ciu,		"cod_ciu")
			Me.mRegistrarControl(Me.txtCod_Zon,		"cod_zon")
			Me.mRegistrarControl(Me.txtCod_Ven,		"cod_ven")
			Me.mRegistrarControl(Me.txtCod_For,		"cod_for")
			Me.mRegistrarControl(Me.txtCod_Tra,		"cod_tra")
			Me.mRegistrarControl(Me.cboTip_Pag,		"tip_pag")			
			Me.mRegistrarControl(Me.TxtTip_Emp,		"tip_emp")
			Me.mRegistrarControl(Me.cboAbc,			"abc")
			Me.mRegistrarControl(Me.chkMar_Rec,		"mar_rec")
			Me.mRegistrarControl(Me.txtCan_Emp,		"can_emp")
			Me.mRegistrarControl(Me.txtVen_Men,		"ven_men")
			Me.mRegistrarControl(Me.txtHor_Caj,		"hor_caj")
			Me.mRegistrarControl(Me.txtPos_Mer,		"pos_mer")
			Me.mRegistrarControl(Me.txtPerfil,		"perfil")
			Me.mRegistrarControl(Me.txtPun_Fue,		"pun_fue")
			Me.mRegistrarControl(Me.txtPun_Deb,		"pun_deb")
			Me.mRegistrarControl(Me.txtCliente,	"clientes")
			Me.mRegistrarControl(Me.txtDet_Cli,		"det_cli")
			Me.mRegistrarControl(Me.txtComentario,	"comentario")
			Me.mRegistrarControl(Me.txtTipo,		"tipo")
			Me.mRegistrarControl(Me.txtClase,		"clase")
			Me.mRegistrarControl(Me.txtGrupo,		"grupo")
			Me.mRegistrarControl(Me.TxtValor1,		"valor1")	
			Me.mRegistrarControl(Me.TxtValor2,		"valor2")
			Me.mRegistrarControl(Me.TxtValor3,		"valor3")
			Me.mRegistrarControl(Me.TxtValor4,		"valor4")
			Me.mRegistrarControl(Me.TxtValor5,		"valor5")	
			Me.mRegistrarControl(Me.TxtNivel,		"nivel")
			Me.mRegistrarControl(Me.cboPrioridad,	"prioridad")
			
			
			Me.txtCod_Zon.mConfigurarBusqueda("zonas",														_
												"cod_zon",													_
												"cod_zon,nom_zon,status",									_
												".,Código,Nombre,Estatus",									_
												"cod_zon,nom_zon",											_
												"../../Framework/Formularios/frmFormularioBusqueda.aspx",	_
												"cod_zon,nom_zon",											_
												"","")	

			Me.txtCod_Pai.mConfigurarBusqueda("paises",														_
												"cod_pai",													_
												"cod_pai,nom_pai,status",									_
												".,Código,Nombre,Estatus",									_
												"cod_pai,nom_pai",											_
												"../../Framework/Formularios/frmFormularioBusqueda.aspx",	_
												"cod_pai,nom_pai",											_
												"","")	

			Me.txtCod_Est.mConfigurarBusqueda("estados",													_
												"cod_est",													_
												"cod_est,nom_est,status",									_
												".,Código,Nombre,Estatus",									_
												"cod_est,nom_est",											_
												"../../Framework/Formularios/frmFormularioBusqueda.aspx",	_
												"cod_est,nom_est",											_
												"","")	

			Me.txtCod_Ciu.mConfigurarBusqueda("ciudades",												_
											"cod_ciu",													_
											"cod_ciu,nom_ciu,status",									_
											".,Código,Nombre,Estatus",									_
											"cod_ciu,nom_ciu",											_
											"../../Framework/Formularios/frmFormularioBusqueda.aspx",	_
											"cod_ciu,nom_ciu",											_
											"","")	

		    Me.txtCod_For.mConfigurarBusqueda("formas_pagos",											_
											"cod_for",													_
											"cod_for,nom_for,status",									_
											".,Código,Nombre,Estatus",									_
											"cod_for,nom_for",											_
											"../../Framework/Formularios/frmFormularioBusqueda.aspx",	_
											"cod_for,nom_for",											_
											"","")	
											
			Me.txtCod_Ven.mConfigurarBusqueda("vendedores",												_
											"cod_ven",													_
											"cod_ven,nom_ven,status",									_
											".,Código,Nombre,Estatus",									_
											"cod_ven,nom_ven",											_
											"../../Framework/Formularios/frmFormularioBusqueda.aspx",	_
											"cod_ven,nom_ven",											_
											"","")	
								
			
			
			Me.txtCod_Tra.mConfigurarBusqueda("transportes",											_
											"cod_tra",													_
											"cod_tra,nom_tra,status",									_
											".,Código,Nombre,Estatus",									_
											"cod_tra,nom_tra",											_
											"../../Framework/Formularios/frmFormularioBusqueda.aspx",	_
											"cod_tra,nom_tra",											_
											"","")									

		End If
		
		Me.poListaResumen	=	Me.grdDatos
		mActualizarListaResumen()
		
	End Sub

'-------------------------------------------------------------------------------------------'
' Invoca al metodo agregar registro de la botonera
'-------------------------------------------------------------------------------------------'

	Protected Sub WbcBotonera_mAgregar(ByVal sender As Object, ByVal e As System.EventArgs) Handles WbcBotonera.mAgregar
	
		Me.pgfContenedor.ActiveTabIndex = 0
		Me.mAgregarRegistro()		
		
	End Sub

'-------------------------------------------------------------------------------------------'
' Invoca al metodo editar registro de la botonera  y actualiza el grid
'-------------------------------------------------------------------------------------------'

	Protected Sub WbcBotonera_mEditar(ByVal sender As Object, ByVal e As System.EventArgs) Handles WbcBotonera.mEditar
		
		Me.pgfContenedor.ActiveTabIndex = 0
		Me.mActualizarListaResumen()  
		Me.mEditarRegistro()
		
	End Sub

'-------------------------------------------------------------------------------------------'
' Invoca al metodo editar registro de la botonera  y actualiza el grid
'-------------------------------------------------------------------------------------------'

	Protected Sub WbcBotonera_mEliminar(ByVal sender As Object, ByVal e As System.EventArgs) Handles WbcBotonera.mEliminar
		
		Me.pgfContenedor.ActiveTabIndex = 0
		Me.mEliminarRegistro()
		Me.mActualizarListaResumen()
		
	End Sub

'-------------------------------------------------------------------------------------------'
' Invoca al metodo eliminar registro de la botonera  y actualiza el grid
'-------------------------------------------------------------------------------------------'
	
	Protected Sub WbcBotonera_mAceptar(ByVal sender As Object, ByVal e As System.EventArgs) Handles WbcBotonera.mAceptar
		
		Me.pgfContenedor.ActiveTabIndex = 0	
		Me.mAceptarEdicion()
		Me.mActualizarListaResumen()
		
	End Sub

'-------------------------------------------------------------------------------------------'
' Invoca al metodo aceptar registro de la botonera  y actualiza el grid
'-------------------------------------------------------------------------------------------'

	Protected Sub WbcBotonera_mCancelar(ByVal sender As Object, ByVal e As System.EventArgs) Handles WbcBotonera.mCancelar
		
		Me.pgfContenedor.ActiveTabIndex = 0
		Me.mCancelarEdicion()
		Me.mActualizarListaResumen()
		
	End Sub

'-------------------------------------------------------------------------------------------'
' Invoca al metodo cancelar edicion de la botonera  y actualiza el grid
'-------------------------------------------------------------------------------------------'

 	Protected Sub mSeleccionarPaginaResumen(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles grdDatos.PageIndexChanging
	
		Me.grdDatos.PageIndex = e.NewPageIndex 
		mActualizarListaResumen()
		
	End Sub

'-------------------------------------------------------------------------------------------'
' Actualiza el grid cuando se activa el indice de la pagina
'-------------------------------------------------------------------------------------------'

	Protected Sub mSeleccionarRegistroResumen(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdDatos.SelectedIndexChanged
		
					
			Me.txtCod_Com.Text		 			= Trim(grdDatos.SelectedDataKey.Values("cod_com"))
			Me.txtNom_Com.Text					= Trim(grdDatos.SelectedDataKey.Values("nom_com"))
			Me.cboStatus.SelectedValue			= Trim(grdDatos.SelectedDataKey.Values("status"))
			Me.txtRif.Text						= Trim(grdDatos.SelectedDataKey.Values("rif"))
			Me.txtNit.Text						= Trim(grdDatos.SelectedDataKey.Values("nit"))
			Me.txtDir_Fis.Text					= Trim(grdDatos.SelectedDataKey.Values("dir_fis"))
			Me.txtDir_Exa.Text					= Trim(grdDatos.SelectedDataKey.Values("dir_exa"))
			Me.txtDir_Otr.Text					= Trim(grdDatos.SelectedDataKey.Values("dir_otr"))
			Me.txtTelefonos.Text				= Trim(grdDatos.SelectedDataKey.Values("telefonos"))
			Me.txtDirecto.Text					= Trim(grdDatos.SelectedDataKey.Values("directo"))
			Me.txtFax.Text						= Trim(grdDatos.SelectedDataKey.Values("fax"))
			Me.txtMovil.Text					= Trim(grdDatos.SelectedDataKey.Values("movil"))
			Me.txtCorreo.Text					= Trim(grdDatos.SelectedDataKey.Values("correo"	))
			Me.txtCorreo2.Text					= Trim(grdDatos.SelectedDataKey.Values("correo2"))
			Me.txtCorreo3.Text					= Trim(grdDatos.SelectedDataKey.Values("correo3"))
			Me.txtWeb.Text						= Trim(grdDatos.SelectedDataKey.Values("web"))
			Me.txtKilometros.Text  				= Trim(grdDatos.SelectedDataKey.Values("kilometros"))
			Me.txtCod_Pai.pcTexto("cod_pai")	= Trim(grdDatos.SelectedDataKey.Values("cod_pai"))
			Me.txtCod_Est.pcTexto("cod_est")	= Trim(grdDatos.SelectedDataKey.Values("cod_est"))
			Me.txtCod_Ciu.pcTexto("cod_ciu")	= Trim(grdDatos.SelectedDataKey.Values("cod_ciu"))
			Me.txtCod_Zon.pcTexto("cod_zon")	= Trim(grdDatos.SelectedDataKey.Values("cod_zon"))
			Me.txtCod_Ven.pcTexto("cod_ven")	= Trim(grdDatos.SelectedDataKey.Values("cod_ven"))
			Me.txtCod_For.pcTexto("cod_for")	= Trim(grdDatos.SelectedDataKey.Values("cod_for"))
			Me.txtCod_Tra.pcTexto("cod_tra")	= Trim(grdDatos.SelectedDataKey.Values("cod_tra"))
			Me.cboTip_Pag.SelectedValue			= Trim(grdDatos.SelectedDataKey.Values("tip_pag"))
			Me.TxtTip_Emp.Text 					= Trim(grdDatos.SelectedDataKey.Values("tip_emp"))
			Me.cboAbc.SelectedValue				= Trim(grdDatos.SelectedDataKey.Values("abc"))
			Me.chkMar_Rec.Checked				= Trim(grdDatos.SelectedDataKey.Values("mar_rec"))
			Me.txtCan_Emp.Text	  				= Trim(grdDatos.SelectedDataKey.Values("can_emp"))
			Me.txtVen_Men.pbValor  				= Trim(grdDatos.SelectedDataKey.Values("ven_men"))
			Me.txtHor_Caj.Text	  				= Trim(grdDatos.SelectedDataKey.Values("hor_caj"))
			Me.txtPos_Mer.pbValor				= Trim(grdDatos.SelectedDataKey.Values("pos_mer"))
			Me.txtPerfil.Text					= Trim(grdDatos.SelectedDataKey.Values("perfil"))
			Me.txtPun_Fue.Text			    	= Trim(grdDatos.SelectedDataKey.Values("pun_fue"))
			Me.txtPun_Deb.Text			    	= Trim(grdDatos.SelectedDataKey.Values("pun_deb"))
			Me.TxtCliente.pbValor				= Trim(grdDatos.SelectedDataKey.Values("clientes"))
			Me.txtDet_Cli.Text					= Trim(grdDatos.SelectedDataKey.Values("det_cli"))
			Me.txtComentario.Text				= Trim(grdDatos.SelectedDataKey.Values("comentario"))
			Me.txtTipo.Text						= Trim(grdDatos.SelectedDataKey.Values("tipo"))
			Me.txtClase.Text					= Trim(grdDatos.SelectedDataKey.Values("clase"))
			Me.txtGrupo.Text					= Trim(grdDatos.SelectedDataKey.Values("grupo"))
			Me.TxtValor1.pbValor				= Trim(grdDatos.SelectedDataKey.Values("valor1"	))		
			Me.TxtValor2.pbValor				= Trim(grdDatos.SelectedDataKey.Values("valor2"))		
			Me.TxtValor3.pbValor				= Trim(grdDatos.SelectedDataKey.Values("valor3"))		
			Me.TxtValor4.pbValor				= Trim(grdDatos.SelectedDataKey.Values("valor4"))
			Me.TxtValor5.pbValor				= Trim(grdDatos.SelectedDataKey.Values("valor5"	))
			Me.TxtNivel.pbValor					= grdDatos.SelectedDataKey.Values("nivel")
			Me.cboPrioridad.SelectedValue		= Trim(grdDatos.SelectedDataKey.Values("prioridad"))	
			
	

		Me.mRegistroSeleccionado(sender,e)
		Me.pgfContenedor.ActiveTabIndex = 0
		
	End Sub

'-------------------------------------------------------------------------------------------'
' Muestra en pantalla los valores del registro seleccionado en el grid
'-------------------------------------------------------------------------------------------'
	
	Public Overrides Sub mActualizarListaResumen()
		'-------------------------------------------------------------------------------------------'
		' Definiendo las variables
		'-------------------------------------------------------------------------------------------'
		Dim lnNumeroParametros					As		Integer = 48
		Dim laParametros(lnNumeroParametros)	As		String
		Dim lcNombreTablas						As		String
		Dim lcCondicionWhere					As		String
		Dim lcAgrupacionGroupBy					As		String
		Dim lcOrdenacionOrderBy					As		String
		
		Dim laDatosTiposIncumplimientos					As		DataSet
		Dim loSeleccionar						As New	cusDatos.goDatos
		
		'-------------------------------------------------------------------------------------------'
		' Colocar el string de los campos a buscar
		'-------------------------------------------------------------------------------------------'
		laParametros(0)		=	 "cod_com" 
		laParametros(1)		=	 "nom_com" 
		laParametros(2)		=	 "status" 
		laParametros(3)		=	 "rif"
		laParametros(4)		=	 "nit"
		laParametros(5)		=	 "dir_fis"
		laParametros(6)		=	 "dir_exa"
		laParametros(7)		=	 "dir_otr"
		laParametros(8)		=	 "telefonos"	
		laParametros(9)		=	 "directo"	
		laParametros(10)		=	"fax"
		laParametros(11)		=	"movil"		
		laParametros(12)		=	"correo"	
		laParametros(13)		=	"correo2"	
		laParametros(14)		=	"correo3"	
		laParametros(15)		=	"web"
		laParametros(16)		=	"kilometros"
		laParametros(17)		=	"cod_pai"
		laParametros(18)		=	"cod_est"
		laParametros(19)		=	"cod_ciu"
		laParametros(20)		=	"cod_zon"
		laParametros(21)		=	"cod_ven"
		laParametros(22)		=	"cod_for"
		laParametros(23)		=	"cod_tra"
		laParametros(24)		=	"tip_pag"	
		laParametros(25)		=	"tip_emp"
		laParametros(26)		=	"abc"
		laParametros(27)		=	"mar_rec"
		laParametros(28)		=	"can_emp"
		laParametros(29)		=	"ven_men"
		laParametros(30)		=	"hor_caj"
		laParametros(31)		=	"pos_mer"
		laParametros(32)		=	"perfil"
		laParametros(33)		=	"pun_fue"
		laParametros(34)		=	"pun_deb"
		laParametros(35)		=	"clientes"
		laParametros(36)		=	"det_cli"
		laParametros(37)		=	"comentario"
		laParametros(38)		=	"tipo"
		laParametros(39)		=	"clase"
		laParametros(40)		=	"grupo"
		laParametros(41)		=	"valor1"	
		laParametros(42)		=	"valor2"
		laParametros(43)		=	"valor3"
		laParametros(44)		=	"valor4"
		laParametros(45)		=	"valor5"	
		laParametros(46)		=	"nivel"
		laParametros(47)		=	"prioridad"
					


		'-------------------------------------------------------------------------------------------'
		' Colocar el string de la o las tablas separadas por comas
		'-------------------------------------------------------------------------------------------'
		lcNombreTablas		=	"Competencia"

		'-------------------------------------------------------------------------------------------'
		' Colocar el string de la o las condiciones
		'-------------------------------------------------------------------------------------------'
		lcCondicionWhere	=	""

		'-------------------------------------------------------------------------------------------'
		' Colocar el string de agrupacion
		'-------------------------------------------------------------------------------------------'
		lcAgrupacionGroupBy	=	""
		
		'-------------------------------------------------------------------------------------------'
		' Colocar el string de Ordenacion
		'-------------------------------------------------------------------------------------------'
		lcOrdenacionOrderBy =	"cod_com"

		'-------------------------------------------------------------------------------------------'
		' Colocar numero de parametros,
		'		  array de parametros, 
		'		  nombre de tabla, 
		'		  condicion		Where   (si aplica), 
		'		  agrupamiento	Group By(si aplica), 
		'		  ordenamiento	Order By(si aplica).
		'-------------------------------------------------------------------------------------------'
		' IMPORTANTE: Los argumentos que no apliquen, dejar en comillas ("").
		'-------------------------------------------------------------------------------------------'
		'goDatos.pcNombreAplicativoExterno = "Framework"
		laDatosTiposIncumplimientos		=	(loSeleccionar.mObtenerTodos(lnNumeroParametros,				_
															 laParametros,							_
															 lcNombreTablas,						_
															 lcCondicionWhere,						_
															 lcAgrupacionGroupBy,					_
															 lcOrdenacionOrderBy))
		Me.grdDatos.DataSource = laDatosTiposIncumplimientos 
		Me.grdDatos.DataBind()
		
	End Sub
	
#End Region
	
	Protected Sub wbcBotonera_mAntesAceptar(ByVal sender As Object, ByVal e As System.EventArgs, ByRef llCancelar As Boolean) _
	Handles wbcBotonera.mAntesAceptar

	'-------------------------------------------------------------------------------------------'
	'Validacion del Codigo de Competencia
	'-------------------------------------------------------------------------------------------'
		
		If Trim(Me.txtCod_Com.Text) = "" then

			Me.wbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia","El Código de la Competencia no puede dejarse en blanco.",wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia,"500px", "250px")

			llCancelar = True 
			Return

		End If

	'-------------------------------------------------------------------------------------------'
	'Validacion del Nombre de Competencia
	'-------------------------------------------------------------------------------------------'
				
		If Trim(Me.txtNom_Com.Text) = "" then

			Me.wbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia","El Nombre de la Competencia no puede dejarse en blanco.",wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia,"500px", "250px")

			llCancelar = True 
			Return

		End If

	'-------------------------------------------------------------------------------------------'
	'Validacion del ABC
	'-------------------------------------------------------------------------------------------'
		If Me.cboAbc.SelectedValue = "" then

			Me.wbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia","El ABC no puede estar vacío.",wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia,"500px", "250px")

			llCancelar = True 	
			Return
		End If

	'-------------------------------------------------------------------------------------------'
	'Validacion de la zona
	'-------------------------------------------------------------------------------------------'
		If Trim(Me.txtCod_Zon.pcTexto("cod_zon")) = "" then

			Me.wbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia","La Zona no puede estar vacía.",wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia,"500px", "250px")

			llCancelar = True 	
			Return
		End If
		
	'-------------------------------------------------------------------------------------------'
	'Validacion del RIF
	'-------------------------------------------------------------------------------------------'
		If Me.txtRif.Text.Trim() = "" then

			Me.wbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia", "El RIF no puede estar vacío.",wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia,"500px", "250px")

			llCancelar = True 	
			Return
			
		End If
		
		If Not CBool(goOpciones.mObtener("RIFDUPCLI", "L")) then
			
			Dim llRifEncontrado	As  Boolean = False
			Dim lcCompetencia	As String = goServicios.mObtenerCampoFormatoSQL(Me.txtCod_Com.Text)
			Dim lcRif			As String = goServicios.mObtenerCampoFormatoSQL(Me.txtRif.Text)
			Dim lcConsulta		As String = "SELECT cod_com FROM Competencia WHERE Rif=" & lcRif & " AND LTRIM(cod_com) <>" & lcCompetencia
			
			llRifEncontrado = (New goDatos().mObtenerTodosSinEsquema(lcConsulta, "competencia").Tables(0).Rows.Count > 0)
			
			If llRifEncontrado Then
			
				 Me.wbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia", _ 
																	"El RIF indicado ya está registrado.", _ 
																	wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia, _ 
																	"500px", "250px")
				llCancelar = True
				Return

			End If
			

		End If

	'-------------------------------------------------------------------------------------------'
	'Validacion del pais
	'-------------------------------------------------------------------------------------------'
		If Trim(Me.txtCod_Pai.pcTexto("cod_pai")) = "" then

			Me.wbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia","El Pais no puede estar vacío.",wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia,"500px", "250px")

			llCancelar = True 	
			Return
		End If

	'-------------------------------------------------------------------------------------------'
	'Validacion de la clase
	'-------------------------------------------------------------------------------------------'
		If Trim(Me.txtClase.Text) = "" then

			Me.wbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia","La Clase no puede estar vacía.",wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia,"500px", "250px")

			llCancelar = True 	
			Return
		End If
		
	'-------------------------------------------------------------------------------------------'
	'Validacion de la ciudad
	'-------------------------------------------------------------------------------------------'
		If Trim(Me.txtCod_Ciu.pcTexto("cod_ciu")) = "" then

			Me.wbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia","La Ciudad no puede estar vacía.",wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia,"500px", "250px")

			llCancelar = True 	
			Return
		End If
		
	'-------------------------------------------------------------------------------------------'
	'Validacion del tipo 
	'-------------------------------------------------------------------------------------------'
		If Trim(Me.txtTipo.Text) = ""  then

			Me.wbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia","El Tipo no puede estar vacío.",wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia,"500px", "250px")

			llCancelar = True 	
			Return
		End If

	'-------------------------------------------------------------------------------------------'
	'Validacion de la forma de pago
	'-------------------------------------------------------------------------------------------'
		If Trim(Me.txtCod_For.pcTexto("cod_for")) = ""  then

			Me.wbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia","La Forma de Pago no puede estar vacía.",wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia,"500px", "250px")

			llCancelar = True 	
			Return
		End If
	'-------------------------------------------------------------------------------------------'
	'Validacion del vendedor
	'-------------------------------------------------------------------------------------------'
		If Trim(Me.txtCod_Ven.pcTexto("cod_ven")) = ""  then

			Me.wbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia","El Vendedor no puede estar vacío.",wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia,"500px", "250px")

			llCancelar = True 	
			Return
		End If
	'-------------------------------------------------------------------------------------------'
	'Validacion del tipo de pago
	'-------------------------------------------------------------------------------------------'
		If Trim(Me.cboTip_Pag.SelectedValue) = "" then

			Me.wbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia","El Tipo de Pago no puede estar vacío.",wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia,"500px", "250px")

			llCancelar = True 	
			Return
		End If
	'-------------------------------------------------------------------------------------------'
	'Validacion del transporte
	'-------------------------------------------------------------------------------------------'
		If Trim(Me.txtCod_Tra.pcTexto("cod_tra")) = ""  then

			Me.wbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia","El Transporte no puede estar vacío.",wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia,"500px", "250px")

			llCancelar = True 	
			Return
		End If

	'-------------------------------------------------------------------------------------------'
	'Validacion de prioridad
	'-------------------------------------------------------------------------------------------'
		
		If Trim(Me.cboPrioridad.SelectedValue) = "" then

			Me.wbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia","La Prioridad no puede dejarse en blanco.",wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia,"500px", "250px")

			llCancelar = True 
			Return

		End If


	End Sub

	Protected Sub cmdImagenes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdImagenes.Click
		Dim	 lcNombreTabla			  As String
		Dim	 lcCondicionWhere		  As String
		Dim	 lcRegistroSeleccionado	  As String
		
		lcNombreTabla		 = "Competencia"
		lcCondicionWhere	 = "Cod_Com = '" & Me.txtCod_Com.Text & "'" 
		lcRegistroSeleccionado	 = lcNombreTabla & ": "	 &  Me.txtNom_Com.Text 
		
		
		Dim lcParametros As String   
		
		lcParametros  = "?pcNombretabla="  & lcNombreTabla 
		lcParametros += "&pcCodicionWhere="  & lcCondicionWhere 
		lcParametros += "&pcRegistroSeleccionado=" & lcRegistroSeleccionado 
		
		Me.wbcAdministradorVentanaModal.mMostrarVentanaModal("../../Framework/Formularios/frmOperacionSubirImagenes.aspx" & lcParametros,"800px","500px")
	   
	End Sub

	Protected Sub wbcBotonera_mBuscar(ByVal sender As Object, ByVal e As System.EventArgs) Handles wbcBotonera.mBuscar
		    goBusquedaRegistro.pcTablaBusqueda			=	"Competencia"
			goBusquedaRegistro.pcCamposIndice			=	"Cod_Com"
			goBusquedaRegistro.pcCamposMostrar			=	"Cod_Com,Nom_Com,comentario,status"
			goBusquedaRegistro.pcTitulosCamposMostrar	=	".,Código,Nombre,Comentario,Status"
			goBusquedaRegistro.pcCamposBusquedaTextual	=	"Cod_Com,Nom_Com"
			goBusquedaRegistro.plBusquedaFramework		=    False
	End Sub


	Protected Sub cmdAuditoria_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAuditoria.Click
	
		Dim lcAuditorias As New StringBuilder()
		
		lcAuditorias.Append("SELECT cod_usu,registro,tipo,tabla,opcion,accion,documento,codigo,detalle,equipo,cod_suc,cod_obj,comentario,notas,clave2,clave3,clave4,clave5 FROM auditorias WHERE tabla=")
		lcAuditorias.Append(goServicios.mObtenerCampoFormatoSQL(Me.pcTablaPrincipal))		
		lcAuditorias.Append(" AND codigo=")
		lcAuditorias.Append(goServicios.mObtenerCampoFormatoSQL(Me.poValorCampo("Cod_Com")))
		
		'lcAuditorias.Append(" AND clave2=")
		'lcAuditorias.Append(goServicios.mObtenerCampoFormatoSQL(Me.poValorCampo("cod_dep")))
		
		Session("Auditorias_Informacion")	= lcAuditorias
		Session("Auditorias_Tabla")			= Me.pcTablaPrincipal
		Session("Auditorias_Documento")		= "0"			   
		Session("Auditorias_Codigo")		= Me.poValorCampo("Cod_Com") '& "|" & Me.poValorCampo("cod_dep")
		Session("Auditorias_Claves")		= "Cod_Com" '& "|" & "cod_dep" 
																		
	 	Me.WbcAdministradorVentanaModal.mMostrarVentanaModal("../../Framework/Formularios/frmVerAuditorias.aspx", "700px", "480px", False)
	 
	End Sub

	Protected Sub cmdExportar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdExportar.Click
	 	
	 	Session("lcNombreTablaEncabezado")	= "Competencia"
		'Session("lcNombreTablaRenglones")	= "Competencia"
		Me.WbcAdministradorVentanaModal.mMostrarVentanaModal("../../Framework/Formularios/frmExportarActualizaciones.aspx", "330px", "250px", False)

	End Sub


	Protected Sub cmdComentarios_mTextoModificado(ByVal Sender As Object, ByVal e As System.EventArgs) Handles cmdComentarios.mTextoModificado
	 	Dim laParametros			As New ArrayList
		Dim lavalores				As New ArrayList
		Dim lcCadenaTransaccion		As New ArrayList
		Dim lcCondicion				As String  =	""

		Dim loObjetoTransaccionSQL	As New cusDatos.goDatos()
		
		laParametros.Add("comentario")
		lavalores.Add("'" & Me.cmdComentarios.pcContenido & "'")

	   lcCondicion = "Cod_Com = '" & Me.txtCod_Com.Text & "'"
		
	   lcCadenaTransaccion.Add(goServicios.mObtenerCadenaActualizarRegistro("Competencia",laParametros,laValores,lcCondicion))
		
		loObjetoTransaccionSQL.mEjecutarTransaccion(lcCadenaTransaccion)
		
		Me.poValorCampo("comentario") =	Me.cmdComentarios.pcContenido
	End Sub

	Protected Sub cmdCamposExtras_Click(ByVal sender As Object, ByVal e As System.EventArgs)
			
			Dim	 lcCondicionWhere	As String
			Dim	 lcClase	        As String   = "Normal"
			Dim	 lcOrigen	        As String   = "Competencia"
			
			lcCondicionWhere	 = "Cod_Com = '" & Trim(Me.txtCod_Com.Text) & "'" 
			
			
			Dim lcParametros As String   
			
			lcParametros = "?pcCodigoRegistro=" & Me.txtCod_Com.Text 
			lcParametros +=  "&pcClase="  & lcClase 
			lcParametros +=  "&pcCodicionWhere="  & lcCondicionWhere 
			lcParametros +=  "&pcDescripcionRegistro="  & Me.txtNom_Com.Text
			lcParametros +=  "&pcOrigen="  & lcOrigen  
			
			
			Me.wbcAdministradorVentanaModal.mMostrarVentanaModal("../../Framework/Formularios/frmCamposExtras.aspx" & lcParametros, "750px", "600px",False)


	End Sub

	Protected Sub cmdDuplicar_Click(ByVal sender As Object, ByVal e As System.EventArgs) _
	Handles cmdDuplicar.Click
		
		Me.mDuplicarRegistro("Cod_Com")
		
	End Sub
	
	
	
	Protected Sub cmdImportar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdImportar.Click
	
	 Session("lcNombreTablaEncabezado")	= "Competencia"
	'Session("lcNombreTablaRenglones")	= "renglones_pedidos"
	Me.WbcAdministradorVentanaModal.mMostrarVentanaModal("../../Framework/Formularios/frmImportarActualizaciones.aspx", "330px", "250px", False)
	
	End Sub
	
	
	Protected Sub cmdConsultas_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdConsultas.Click

		Session("lcTablaConsultas") = "Competencia"
		Me.WbcAdministradorVentanaModal.mMostrarVentanaModal("../../Framework/Formularios/frmConsultas.aspx", "650px", "500px",False)
		
	End Sub
	
	
	Protected Sub cmdComplementos_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdComplementos.Click

		Session("lcTablaComplementos") = "Competencia"
		Me.WbcAdministradorVentanaModal.mMostrarVentanaModal("../../Framework/Formularios/frmComplementos.aspx", "650px", "500px",False)

	End Sub
	
	
	Protected Sub cmdContable_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdContable.Click

		 Dim laListaParametosContable As New System.Collections.Generic.Dictionary(Of String, Object)
		
		laListaParametosContable.Add("Tabla",				"Competencia")
		laListaParametosContable.Add("CamposIndice",		New String(){"Cod_Com"})
		laListaParametosContable.Add("ValoresIndice",		New Object(){Trim(Me.txtCod_Com.Text)})
		laListaParametosContable.Add("Formulario",			TypeName(Me.Page()))
		laListaParametosContable.Add("SoloLectura",			False)
		
	 	Session("laListaParametosContable")	= laListaParametosContable
		
		Me.wbcAdministradorVentanaModal.mMostrarVentanaModal("../../Administrativo/Formularios/frmFormularioInformacionContable.aspx","900px","540px",False)




	End Sub
	
	Protected Sub cmdContactos_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdContactos.Click

	    Dim	 lcNombreTabla			As String
		
		lcNombreTabla			 = "competencia"
		
		Dim lcParametros As String   
		
		lcParametros  = "?pcCodigoRegistro=" & Me.txtCod_Com.Text 
		lcParametros +=  "&pcNombreTabla=" & lcNombreTabla 
		
		Me.wbcAdministradorVentanaModal.mMostrarVentanaModal("../../Framework/Formularios/frmActualizacionContactos.aspx" & lcParametros,"750px","600px",False)


	End Sub	

Protected Sub cmdAgrupaciones_Click(ByVal sender As Object, ByVal e As System.EventArgs) handles cmdAgrupaciones.Click

	    Dim	 lcCondicionActualizar	As String
	    Dim	 lcNombreTabla			As String
		
		lcCondicionActualizar	 = "cod_com = '" & Trim(Me.txtCod_Com.Text) & "'" 
		lcNombreTabla			 = "competencia"
		
		Dim lcParametros As String   
		
		lcParametros  = "?pcCodigoRegistro=" & Me.txtCod_Com.Text 
		lcParametros +=  "&pcNombreTabla=" & lcNombreTabla 
		lcParametros +=  "&pcCodicionActualizar="  & lcCondicionActualizar 
		lcParametros +=  "&pcDescripcionRegistro="  & Me.txtNom_Com.Text
		
		
		Me.wbcAdministradorVentanaModal.mMostrarVentanaModal("../../Administrativo/Formularios/frmAsignarAgrupaciones.aspx" & lcParametros,"600px","500px",False)


	End Sub	
	
	Protected Sub cmdSeguimientos_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSeguimientos.Click

		Dim	 lcNombreTabla	As String
		
		lcNombreTabla	= "competencia"
		
		Dim lcParametros As String   
		
		lcParametros  = "?pcCodigoRegistro=" & Me.txtCod_Com.Text 
		lcParametros +=  "&pcNombreTabla=" & lcNombreTabla 
		
		Me.wbcAdministradorVentanaModal.mMostrarVentanaModal("../../Administrativo/Formularios/frmActualizacionSeguimientos.aspx" & lcParametros,"750px","600px",False)



	End Sub

	
	Protected Sub cmdPrecios_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdPrecios.Click

		Dim lcParametros As String   
		
		lcParametros  = "?pcCodigoCliente=" & Me.txtCod_Com.Text 
		lcParametros +=  "&pcNombreCliente=" & Me.txtNom_Com.Text 
		
		Me.wbcAdministradorVentanaModal.mMostrarVentanaModal("../../Administrativo/Formularios/frmOperacionPreciosClientes.aspx" & lcParametros,"750px","600px",False)


	End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' CMS: 04/09/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 08/09/09: Se Agregaron los siguientes complementos: Contactos, Agrupaciones y Seguimiento
'-------------------------------------------------------------------------------------------'