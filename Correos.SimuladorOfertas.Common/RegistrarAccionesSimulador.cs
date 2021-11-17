using Correos.SimuladorOfertas.DTOs;
using log4net;
using System;
using System.Collections;
using System.Globalization;
using System.Text;
using System.Windows.Forms;

namespace Correos.SimuladorOfertas.Common
{
    public static class RegistrarAccionesSimulador
    {
        #region Log4net
        /// <summary> 
        /// Objeto para la escritura en el log 
        /// </summary> 
        private static readonly ILog log = LogManager.GetLogger(typeof(RegistrarAccionesSimulador));
        #endregion

        #region Funciones Privadas

        #region ToLongString

        /// <summary>
        /// Formatea la excepción para escibirse en forma de texto
        /// </summary>
        /// <param name="ex"></param>
        /// <returns></returns>
        private static string ToLongString(Exception ex)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "Fecha : {0}", DateTime.Now.ToShortDateString()));
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "Hora : {0}", DateTime.Now.ToShortTimeString()));
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "Descripción : {0}", ex.Message));
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "Función Fallo : {0}", ex.TargetSite));
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "Traza Fallo : {0}", ex.StackTrace));
            sb.AppendLine("Información adicional : ");
            foreach (IDictionary item in ex.Data)
            {
                sb.AppendLine(Convert.ToString(item));
            }
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "Inner Exception : {0}", ex.InnerException != null ? ex.InnerException.Message : ""));

            sb.AppendLine("-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-");
            return sb.ToString();
        }

        /// <summary>
        /// Formatea la traza para ser usada en un documento de texto.
        /// </summary>
        /// <param name="trz"></param>
        /// <returns></returns>
        private static string ToLongString(AuditoriaDataTrazaBE trz)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "Usuario Window : {0}", trz.UsuarioPC));
            if (!string.IsNullOrEmpty(trz.UsuarioSAP))
            {
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "Usuario con Credencial SAP CRM : {0}", trz.UsuarioSAP));
            }
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "Fecha : {0}", trz.Fecha));
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "Hora : {0}", trz.Hora));
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "Descripción : {0}", trz.Accion));
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "Objeto afectado : {0}", trz.ObjetoAccion));
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "IP Acción : {0}", trz.IPAccion));
            sb.AppendLine("-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-");
            return sb.ToString();
        }

        /// <summary>
        /// Formatea la traza para ser usada en un documento de texto.
        /// </summary>
        /// <param name="trz"></param>
        /// <returns></returns>
        private static string ToLongString(string mensaje)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture, ">>>: {0}", mensaje));
            return sb.ToString();
        }

        #endregion

        #endregion

        #region Funciones Públicas

        /// <summary>
        /// Funcion que se encarga de guardar la excepcion en el fichero de log.
        /// </summary>
        /// <param name="ex"></param>
        public static void GuardarExcepcion(Exception ex)
        {
            GuardarExcepcion(ex, true);
        }

        /// <summary>
        /// Funcion que realmente se encarga de guardar la excepcion en el fichero de log.
        /// </summary>
        /// <param name="ex"></param>
        public static void GuardarExcepcion(Exception ex, bool mostrarMensaje)
        {
            //Por si la propia escritura en el fichero falla
            try
            {
                log.Error(ex.Message, ex);

                if (mostrarMensaje)
                {
                    //TODO: Cuando durante la actualizacion de la base de datos tenemos un error por ejemplo copiando las ofertas, la colección 'Application.OpenForms' estará vacía, así que dará un OutOfRange..., modificamos esto para que no suceda, revisar si se puede optimizar.
                    if (Application.OpenForms.Count == 0)
                    {
                        MessageBox.Show(SimuladorResources.ErrorExcepcionControladaSistema, SimuladorResources.Advertencia,
                            MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, 0);
                    }
                    else
                    {
                        MessageBox.Show(Application.OpenForms[0], SimuladorResources.ErrorExcepcionControladaSistema, SimuladorResources.Advertencia,
                            MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, 0);
                    }
                }
            }
            catch
            {
                MessageBox.Show(new Form() { TopMost = true }, SimuladorResources.ErrorGuardarFicheroSistema, SimuladorResources.Advertencia,
                        MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, 0);
            }
        }

        /// <summary>
        /// Funcion que guardar trazas de acción para auditoría en el fichero de log.
        /// </summary>
        /// <param name="trz"></param>
        public static void GuardarTraza(AuditoriaDataTrazaBE trz)
        {
            //Por si la propia escritura en el fichero falla
            try
            {
                log.Info(ToLongString(trz));                
            }
            catch
            {
                MessageBox.Show(SimuladorResources.ErrorGuardarFicheroSistema, SimuladorResources.Advertencia,
                        MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, 0);
            }
        }
        public static void GuardarTraza(string mensaje)
        {
            //Por si la propia escritura en el fichero falla
            try
            {
                log.Info(ToLongString(mensaje));
            }
            catch
            {
                //MessageBox.Show(SimuladorResources.ErrorGuardarFicheroSistema, SimuladorResources.Advertencia,
                //        MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, 0);
            }
        }
        #endregion
    }
}
