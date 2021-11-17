using Correos.SimuladorOfertas.Common;
using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.InOutHeavy;
using Correos.SimuladorOfertas.InOutLight;
using System;
using System.Collections.ObjectModel;
using System.Net;

namespace Correos.SimuladorOfertas.Business
{
    public class CredencialesBL
    {
        #region Métodos Públicos

        /// <summary>
        /// Método que comprueba qué productos deben actualizar su definición
        /// </summary>
        /// <param name="usuario">Código de usuario</param>
        /// <param name="pass">Contraseña</param>
        /// <returns>Listado de productos a los que se debe actualizar su definición</returns>
        public Collection<ProductoBE> ComprobarDescargaDefinicion(string usuario, string pass)
        {
            ProductoBL productoBL = new ProductoBL();


            Collection<ProductoBE> listaProductos = new Collection<ProductoBE>();
            listaProductos = productoBL.ObtenerListadoProductos();

            Collection<ProductoBE> listaFechaCambioProdRfc = new Collection<ProductoBE>();

            if (SSOHelper.Instance.LogarConSSO)
            {
                CommunicatorHeavy conectorSAP = new CommunicatorHeavy(SSOHelper.Instance.Usuario, pass);
                //SSOHelper.Instance.ActualizarCookiePortal();
                SSOHelper.Instance.InicializarWSHeavy(conectorSAP.Cliente);

                listaFechaCambioProdRfc = conectorSAP.ZCFechaCambioProdRfc(listaProductos);
                Collection<ProductoBE> result = productoBL.MarcarListadoProductosParaDescargaDefinicion(listaFechaCambioProdRfc);

                SSOHelper.Instance.LimpiarWSLight();
                return result;
            }
            else
            {
                CommunicatorHeavy conectorSAP = new CommunicatorHeavy(usuario, pass);
                //Ahora se comprueba que productos tienen una fechaValidezDesde menor que la que nos devuelve la llamada

                listaFechaCambioProdRfc = conectorSAP.ZCFechaCambioProdRfc(listaProductos);
                return productoBL.MarcarListadoProductosParaDescargaDefinicion(listaFechaCambioProdRfc);
            }
        }

        /// <summary>
        /// Indica si hay conectividad con la VPN de Correos
        /// </summary>
        /// <param name="usuario">Nombre del usuario</param>
        /// <param name="password">Contraseña del usuario</param>
        /// <returns>Indica si hay o no conexión</returns>
        public static ResultBE HayConexionVPN(string usuario, string password)
        {
            ResultBE objResult = new ResultBE();
            Boolean certificadoActivo = false;

            try
            {
                if (SSOHelper.Instance.LogarConSSO)
                {
                    CommunicatorLight conectorSAP = new CommunicatorLight(SSOHelper.Instance.Usuario, password);
                    ResultBE result = null;
                    
                    try
                    {
                        if (!SSOHelper.Instance.CookieGenerada)
                        {
                            SSOHelper.Instance.ActualizarCookiePortal(ref certificadoActivo, usuario, password);
                            if (!certificadoActivo)
                                throw new Exception();
                        }
                        
                        SSOHelper.Instance.InicializarWSLight(conectorSAP.Cliente);
                        
                        

                        result = conectorSAP.ZCCompruebaConexionRfc(SSOHelper.Instance.Usuario);
                        SSOHelper.Instance.LimpiarWSLight();
                    }
                    catch (WebException webException)
                    {
                        if (webException.Status == WebExceptionStatus.ProtocolError
                            && ((HttpWebResponse)webException.Response).StatusCode == HttpStatusCode.Unauthorized) // Cookie inválida
                        {
                            // Si ya hemos logado alguna vez con SSO
                            if (SSOHelper.Instance.CookieGenerada)
                            {
                                // Volvemos a actualizar la cookie
                                SSOHelper.Instance.ActualizarCookiePortal(ref certificadoActivo, usuario, password);                             
                                result = conectorSAP.ZCCompruebaConexionRfc(SSOHelper.Instance.Usuario);
                                SSOHelper.Instance.LimpiarWSLight();
                            }
                            else
                            {
                                result = new ResultBE() {Resultado = false, TextoError = SimuladorResources.ErrorSSO };
                            }
                        }
                        else
                        {
                            SSOHelper.Instance.LimpiarWSLight();
                            throw webException;
                        }
                    }
                    return result;
                }
                else
                {
                    CommunicatorLight conectorSAP = new CommunicatorLight(usuario, password);
                    return conectorSAP.ZCCompruebaConexionRfc(usuario);
                }
            }
            catch (Exception ex)
            {
                objResult.Resultado = false;

                if (!certificadoActivo)
                {
                    //Si las credenciales no son correctas en la cookie
                    objResult = new ResultBE() { Resultado = false, TextoError = SimuladorResources.ErrorSSO };
                }                
                else if (ex is System.ServiceModel.EndpointNotFoundException)
                {
                    objResult.TextoError = SimuladorResources.ErrorConexionVPNCorreos + Environment.NewLine + SimuladorResources.CompruebaConexionInternet;
                }
                else if (ex.InnerException != null && ex.InnerException.Message != SimuladorResources.ErrorAutenticacionExcepcion)
                {
                    objResult.TextoError = SimuladorResources.ErrorConexionVPNCorreos;
                }
                else if (ex.InnerException != null && ex.InnerException.Message == SimuladorResources.ErrorAutenticacionExcepcion && SSOHelper.Instance.CookieGenerada)
                {
                    //Si las credenciales no son correctas en la cookie
                    objResult = new ResultBE() { Resultado = false, TextoError = SimuladorResources.ErrorSSO };
                }
                else
                {
                    objResult.TextoError = SimuladorResources.ErrorAutenticacionSAP;
                }
                RegistrarAccionesSimulador.GuardarExcepcion(ex, false);
                return objResult;
            }
        }

        #endregion
    }
}
