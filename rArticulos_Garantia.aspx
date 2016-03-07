<%@ Page Language="VB" AutoEventWireup="false" CodeFile="rArticulos_Garantia.aspx.vb" Inherits="rArticulos_Garantia" %>

<%@ Register Assembly="vis2Controles" Namespace="vis2Controles" TagPrefix="vis2Controles" %>
<%@ Register Assembly="vis1Controles" Namespace="vis1Controles" TagPrefix="vis1Controles" %>
<%@ Register Assembly="vis3Controles" Namespace="vis3Controles" TagPrefix="vis3Controles" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>

<%@ Register Assembly="CrystalDecisions.Web, Version=10.2.3600.0, Culture=neutral, PublicKeyToken=692fbea5521e1304"
    Namespace="CrystalDecisions.Web" TagPrefix="CR" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Listado de Artículos con Garantía</title>
    <link href="~/Framework/cssEstilosFramework.css" rel="stylesheet" type="text/css" />
    <link href="~/Administrativo/cssEstilosAdministrativo.css" rel="stylesheet" type="text/css" />
    <link href="/aspnet_client/System_Web/2_0_50727/CrystalReportWebFormViewer3/css/default.css"
        rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
        <script type="text/javascript">
//        
//		if(!window.loServiciosImpresoraReportes){
//		    window.loServiciosImpresoraReportes = new Object();
//		    window.loServiciosImpresoraReportes.lnIntentos = 0;
//		}

//		window.loServiciosImpresoraReportes.mImprimirReporte = function(loEvento){
//			oEvento=loEvento||event;
//			oEvento=window.top.oServiciosCliente.mFormatearEvento(loEvento);
//			var loMarco = window.document.getElementById('wbcImpresoraDeReportes_ifrImprimir');
//			window.top.oServiciosCliente.mAgregarManejadorEvento(loMarco,'load',window.loServiciosImpresoraReportes.mIniciarImpresion);
//			window.setTimeout(function(){loMarco.src='../../administrativo/Reportes/rArticulos_Garantia.aspx?salida=impresora';}, 0);
//		}
//		
//		window.loServiciosImpresoraReportes.mReporteCargado = function(){
//			var loMarco = window.document.getElementById('wbcImpresoraDeReportes_ifrImprimir');
//			var loVentana = loMarco.contentWindow;
//			if(!loVentana){return false;};
//			var loDocumento = loVentana.document;
//			if(!loDocumento){ return false;};
//			var loHtml = loDocumento.childNodes[0];
//			if(!loHtml){ return false;};
//			var loBody = loHtml.childNodes[0];
//			if(!loBody){ return false;};
//			var loDpf = loBody.childNodes[0];
//			if(!loDpf){ return false;};
//			if(loDpf.type=='application/pdf'){
//				return true;
//			}else{
//				return false;
//			}
//		}
//		
//		window.loServiciosImpresoraReportes.mIniciarImpresion = function(loEvento){
//			oEvento=loEvento||event;
//			oEvento=window.top.oServiciosCliente.mFormatearEvento(loEvento);
//			
//			if(window.loServiciosImpresoraReportes.mReporteCargado()){
//			    var loMarco = window.document.getElementById('wbcImpresoraDeReportes_ifrImprimir');
//                window.setTimeout(function(){loMarco.contentWindow.print()},500);
//                
//                window.loServiciosImpresoraReportes.lnIntentos = 0; 
//                return;
//			}
//            if(window.loServiciosImpresoraReportes.lnIntentos<10){
//                window.loServiciosImpresoraReportes.lnIntentos += 1;
//                window.setTimeout(window.loServiciosImpresoraReportes.mIniciarImpresion,500);
//            }else{
//                window.loServiciosImpresoraReportes.lnIntentos = 0;
//                return;
//            }   
//		}
        </script>
    <div >
         <CR:CrystalReportViewer ID="crvrArticulos_Garantia" runat="server" AutoDataBind="true" EnableDatabaseLogonPrompt="False"
            EnableParameterPrompt="False" HasCrystalLogo="False" Height="50px" Width="350px"
            HasPrintButton="False" />
       <asp:ScriptManager ID="ScriptManager1" runat="server">
            <Scripts>
                <asp:ScriptReference Path="~/Framework/Librerias/jsServiciosCliente.js" />
                <asp:ScriptReference Path="~/Framework/Librerias/jsServiciosDatos.js" />
                <asp:ScriptReference Path="~/Framework/Librerias/jsServiciosFormato.js" />
            </Scripts>
        </asp:ScriptManager>
        <asp:UpdatePanel ID="udpReporte" runat="server">
            <ContentTemplate>
                <vis3Controles:wbcImpresoraReportes runat="server" ID="wbcImpresoraDeReportes" plMostrarBotonImprimir='True' />
                <vis3Controles:pnlVentanaModal ID="PnlVentanaModalPrincipal" runat="server" pcEstiloBotonCerrar="BotonCerrarVentanaModal"
                    pcEstiloFondo="FondoVentanaModal" pcEstiloMarco="MarcoVentanaModal" pcTextoBotonCerrar="Cerrar"
                    plMostrarBotonCerrar="false" poAlto="520px" poAncho="550px" Style="left: -16px;
                    top: 50px" />
                <vis3Controles:pnlMensajeModal ID="PnlMensajeModal" runat="server" pcEstiloContenido="ContenidoMensajeModal"
                    pcEstiloFondo="FondoVentanaModal" pcEstiloTitulo="TituloMensajeModal" pcEstiloVentana="MarcoMensajeModal"
                    poAlto="400px" poAncho="750px" poArriba="20%" poIzquierda="30%" />
                <vis3Controles:wbcAdministradorMensajeModal ID="WbcAdministradorMensajeModal" runat="server" />
                <vis3Controles:wbcAdministradorVentanaModal ID="WbcAdministradorVentanaModal" runat="server" />
            </ContentTemplate>
        </asp:UpdatePanel>
        <br />
    </div>
    </form>
</body>
</html>
