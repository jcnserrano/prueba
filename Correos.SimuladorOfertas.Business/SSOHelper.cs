using Correos.SimuladorOfertas.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.Text;
namespace Correos.SimuladorOfertas.Business
{
    public class SSOHelper
    {
        private static SSOHelper _instance = null;
        public static SSOHelper Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new SSOHelper();
                }
                return _instance;
            }
        }
        /// <summary>
        /// Almacena el usuario logado
        /// </summary>
        public string Usuario { get; set; }
        /// <summary>
        /// Devuelve si ya se ha obtenido una cookie, aunque sea antigua
        /// </summary>
        public bool CookieGenerada { get { return !string.IsNullOrWhiteSpace(CookiePortal); } }
        /// <summary>
        /// Cookie obtenida del portal
        /// </summary>
        string CookiePortal { get; set; }
        /// <summary>
        /// Indica si el usuario debe logar con la cookie de Portal o con user&pass
        /// </summary>
        public bool LogarConSSO { get { return Utils.GetValorFromAppConfig(AppSettingsEnum.ModoLoginSSO).Equals("True"); } }
        /// <summary>
        /// Mantiene el OperationContextScope del WS Light para evitar perder la referencia
        /// </summary>
        private OperationContextScope OperationContextScopeLight { get; set; }
        /// <summary>
        /// Mantiene el OperationContextScope del WS Heavy para evitar perder la referencia
        /// </summary>
        private OperationContextScope OperationContextScopeHeavy { get; set; }
        public void ActualizarCookiePortal(ref bool certificadoActivo, string usuario = null, string password = null)
        {
            string url = Utils.GetValorFromAppConfig(AppSettingsEnum.RutaPortalSinCredenciales);
            
            if (!string.IsNullOrWhiteSpace(usuario) && !string.IsNullOrWhiteSpace(password))
            {
                //Para evitar problemas con los comodines de las contraseñas, codificamos los caracteres de la password a hexadecimal
                //password = Uri.EscapeDataString("deloitte2");
                //url = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaPortalConCredenciales), "E007639", "deloitte2");
                //this.Usuario = usuario;
            }
            
            HttpWebRequest httpwr = (HttpWebRequest)HttpWebRequest.Create(url);
                       
            /*
              ANTES DE SSO
             if (string.IsNullOrWhiteSpace(usuario) || string.IsNullOrWhiteSpace(password))
                {
                    // Utilizamos los credenciales de windows de la máquina
                    httpwr.Credentials = CredentialCache.DefaultNetworkCredentials;
                }
                httpwr.Credentials = new NetworkCredential("E007639", "Correos.13");
            */
            
            // Use the X509Store class to get a handle to the local certificate stores. "My" is the "Personal" store.
            X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            // Open the store to be able to read from it.
            store.Open(OpenFlags.OpenExistingOnly);

            // Use the X509Certificate2Collection class to get a list of certificates that match our criteria (in this case, we should only pull back one).
            X509Certificate2Collection collection = new X509Certificate2Collection(store.Certificates.Find(X509FindType.FindBySubjectName, "SAPSSO", true));

            if (collection.Count > 0)
            {
                certificadoActivo = true;

                // Associate the certificates with the request
                httpwr.ClientCertificates = collection;

                httpwr.CookieContainer = new CookieContainer();
                // Sin un UserAgent válido se recibe un error 400
                httpwr.UserAgent = "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.96 Safari/537.36";

                // Hay que hacer GetResponse para que GetCookieHeader no sea vacío
                WebResponse wr = httpwr.GetResponse();
            }
            else
            {
                certificadoActivo = false;
            }
          
            CookiePortal = httpwr.CookieContainer.GetCookieHeader(new Uri(url));
                        
        }
        public void InicializarWSLight(Correos.SimuladorOfertas.InOutLight.ServiceWCF_Light.ZBSP_SIMULADORClient client)
        {
            OperationContextScopeLight = new OperationContextScope(client.InnerChannel);
            HttpRequestMessageProperty request = new HttpRequestMessageProperty();
            request.Headers["Cookie"] = CookiePortal;
            OperationContext.Current.OutgoingMessageProperties[HttpRequestMessageProperty.Name] = request;
        }
        public void LimpiarWSLight()
        {
            OperationContextScopeLight.Dispose();
        }
        public void InicializarWSHeavy(Correos.SimuladorOfertas.InOutHeavy.ServiceWCF_Heavy.ZBSP_SIMULADORClient client)
        {
            OperationContextScopeHeavy = new OperationContextScope(client.InnerChannel);
            HttpRequestMessageProperty request = new HttpRequestMessageProperty();
            request.Headers["Cookie"] = CookiePortal;
            OperationContext.Current.OutgoingMessageProperties[HttpRequestMessageProperty.Name] = request;
        }
        public void LimpiarWSHeavy()
        {
            OperationContextScopeHeavy.Dispose();
        }
    }
}
