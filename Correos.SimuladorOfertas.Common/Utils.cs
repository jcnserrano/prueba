using System;
using System.Configuration;
using System.IO;
using System.Security.Cryptography;
using System.Threading;

namespace Correos.SimuladorOfertas.Common
{
    public static class Utils
    {
        public static string FicheroErrores
        {
            get
            {
                return Utils.ObtenerDirectorioBase() + "\\{0}";
            }
        }

        #region Obtener plantilla base Simulador

        /// <summary>
        /// Se encarga de obtener la ruta de la plantilla usada por defecto en la aplicación
        /// </summary>
        /// <returns></returns>
        public static string ObtenerDirectorioBase()
        {
            return string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaEjecucionSimulador), Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData));
        }

        /// <summary>
        /// Obtiene la ruta donde se guardan los xml serializados de lo que nos venga de SAP
        /// </summary>
        /// <returns></returns>
        public static string ObtenerRutaDescargaXML() 
        {
            return string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaDescargaFicherosXML), Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData));
        }

        /// <summary>
        /// Obtiene la ruta donde se guardan los xml serializados de lo que enviemos a SAP
        /// </summary>
        /// <returns></returns>
        public static string ObtenerRutaSubidaXML() 
        {
            return string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaSubidaFicherosXML), Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData));
        }


        #endregion

        #region GetValorFromAppConfig

        /// <summary> 
        /// Método que devuelve el valor de la clave pasada por parámetro del app.config 
        /// </summary> 
        /// <param name="key">Clave a obtener su valor</param> 
        /// <returns>Valor de la clave pasada por parámetro</returns> 
        public static string GetValorFromAppConfig(AppSettingsEnum key)
        {
            try
            {
                return ConfigurationManager.AppSettings[key.ToString()].ToString();
            }
            catch
            {
                return string.Empty;
            }
        }

        #endregion

        #region ObtenerTextoMensajesBGWEnum

        /// <summary>
        /// Funcion que devuelve el texto asociado al enumerado
        /// </summary>
        /// <param name="texto"></param>
        /// <returns></returns>
        public static string ObtenerTextoMensajesBGWEnum(int valor)
        {
            try
            {
                switch (valor)
                {
                    case (int)MensajesBGWEnum.BuscandoClientes:
                        return SimuladorResources.BuscandoClientes;

                    case (int)MensajesBGWEnum.EliminandoOfertaSeleccionada:
                        return SimuladorResources.EliminandoOfertaSeleccionada;

                    case (int)MensajesBGWEnum.GenerandoExcel:
                        return SimuladorResources.GenerandoExcel;

                    case (int)MensajesBGWEnum.GuardandoContenidoFicherosExcel:
                        return SimuladorResources.GuardandoContenidoFicherosExcel;

                    case (int)MensajesBGWEnum.GuardandoDefeinicionProductosOferta:
                        return SimuladorResources.GuardandoDefeinicionProductosOferta;

                    case (int)MensajesBGWEnum.GuardandoProductosOferta:
                        return SimuladorResources.GuardandoProductosOferta;

                    case (int)MensajesBGWEnum.ObtiendoOfertasClienteSAP:
                        return SimuladorResources.ObtiendoOfertasClienteSAP;

                    case (int)MensajesBGWEnum.PidiendoCredencialesUsuario:
                        return SimuladorResources.PidiendoCredencialesUsuario;

                    case (int)MensajesBGWEnum.SincronizandoDatosSAPCRM:
                        return SimuladorResources.SincronizandoDatosSAPCRM;

                    case (int)MensajesBGWEnum.VaciandoContenidoExcel:
                        return SimuladorResources.VaciandoContenidoExcel;

                    case (int)MensajesBGWEnum.ActualizandoStatusOferta:
                        return SimuladorResources.ActualizandoStatusOferta;

                    case (int)MensajesBGWEnum.ActualizandoStatusOfertas:
                        return SimuladorResources.ActualizandoStatusOfertas;
                        
                    case (int)MensajesBGWEnum.ActualizandoPotencialidadUsuario:
                        return SimuladorResources.ActualizandoPotencialidadUsuario;

                    case (int)MensajesBGWEnum.ActualizandoDefinicionesProductos:
                        return SimuladorResources.ActualizandoDefinicionesProductos;

                    case (int)MensajesBGWEnum.RealizandoCalculosTarifas:
                        return SimuladorResources.RealizandoCalculosTarifas;

                    case (int)MensajesBGWEnum.SincronizandoDefinicionProductos:
                        return SimuladorResources.SincronizandoDefinicionProductos;

                    case (int)MensajesBGWEnum.GenerandoInformeTarifas:
                        return SimuladorResources.GenerandoInformeTarifas;

                    case (int)MensajesBGWEnum.ActualizandoCoeficientesUsuario:
                        return SimuladorResources.ActualizandoCoeficientesUsuario;

                    case (int)MensajesBGWEnum.RealizandoGeneracionContrato:
                        return SimuladorResources.RealizandoGeneracionContrato;

                    case (int)MensajesBGWEnum.ActualizandoValoresCubicaje:
                        return SimuladorResources.ActualizandoValoresCubicaje;

                    case (int)MensajesBGWEnum.ActualizandoValoresAgrupacionesTipologia:
                        return SimuladorResources.ActualizandoValoresAgrupacionesTipologia;

                    case (int)MensajesBGWEnum.ActualizandoEstadoSincronizacionOfertas:
                        return SimuladorResources.ActualizandoEstadoSincronizacionOfertas;

                    case (int)MensajesBGWEnum.ActualizandoDefinicionesProductosTARIFAS:
                        return SimuladorResources.ActualizandoDefinicionesProductosTARIFAS;

                    default:
                        return String.Empty;
                }
            }
            catch
            {
                return String.Empty;
            }
        }

        #endregion


        #region ObtenerFechaDeString

        /// <summary>
        /// Obtiene el Datetime asociado a la fecha que se pasa por parámetro
        /// </summary>
        /// <param name="fecha"></param>
        /// <returns></returns>
        public static DateTime? ObtenerFechaDeString(string fecha) 
        {
            try 
            {
                string[] fechaSplit = fecha.Split('.');

                if (fechaSplit.Length.Equals(3))
                {

                    int intAnio, intMes, intDia;

                    int.TryParse(fechaSplit[2], out intAnio);
                    int.TryParse(fechaSplit[1], out intMes);
                    int.TryParse(fechaSplit[0], out intDia);

                    return new DateTime(intAnio, intMes, intDia);
                }
                else 
                {
                    return null;
                }

            }
            catch 
            {
                return null;
            }
        }       

        #endregion 

        #region ConvertFromDecimalToDateTimeParse

        /// <summary>
        /// Obtiene un datetime a partir de una fecha guardada en decimal
        /// </summary>
        /// <param name="fecha"></param>
        /// <returns></returns>
        public static DateTime ConvertFromDecimalToDateTimeParse(string fecha)
        {
            try
            {
                string anio, mes, dia;

                anio = fecha.Substring(0, 4);
                mes = fecha.Substring(4, 2);
                dia = fecha.Substring(6, 2);

                int intanio, intmes, intdia;
                int.TryParse(anio, out intanio);
                int.TryParse(mes, out intmes);
                int.TryParse(dia, out intdia);

                return new DateTime(intanio, intmes, intdia);

            }
            catch
            {
                return new DateTime();
            }
        }

        #endregion

        #region FechasIncorrectas

        /// <summary>
        /// se encarga de verificar que la fecha desde es menor o igual a la fecha hasta
        /// </summary>
        /// <returns></returns>
        public static bool FechasIncorrectas(DateTime desde, DateTime hasta)
        {

            if (desde.Year > hasta.Year)
            {
                //el año es posterior 
                return true;
            }
            else if (desde.Year < hasta.Year)
            {
                //el año es anterior.
                return false;
            }
            else
            {
                //a igualdad de años comparamos los meses
                if (desde.Month > hasta.Month)
                {
                    //el mes es posterior.
                    return true;
                }
                else if (desde.Month < hasta.Month)
                {
                    //el mes es anterior
                    return false;
                }
                else
                {
                    //a igualdad de año y mes comparamos el día
                    if (desde.Day > hasta.Day)
                    {
                        //el dia es posterior
                        return true;
                    }
                    else
                    {
                        //el dia es igual o anterior
                        return false;
                    }
                }
            }

        }

        #endregion

        #region ObtenerNombreMes
        /// <summary>
        /// Obtiene el nombre del mes que se pasa
        /// </summary>
        /// <param name="mes"></param>
        /// <returns></returns>
        public static string ObtenerNombreMes(int mes)
        {
            switch (mes)
            {
                case 1:
                    return "Enero";
                case 2:
                    return "Febrero";
                case 3:
                    return "Marzo";
                case 4:
                    return "Abril";
                case 5:
                    return "Mayo";
                case 6:
                    return "Junio";
                case 7:
                    return "Julio";
                case 8:
                    return "Agosto";
                case 9:
                    return "Septiembre";
                case 10:
                    return "Octubre";
                case 11:
                    return "Noviembre";
                case 12:
                    return "Diciembre";
                default:
                    return string.Empty;
            }

        }

        #endregion

        #region EsMismoDiaFechas

        /// <summary>
        /// Devuelve cierto si año,mes y día de las dots fechas son iguales
        /// </summary>
        /// <param name="dateTime1"></param>
        /// <param name="dateTime2"></param>
        /// <returns></returns>
        public static bool EsMismoDiaFechas(DateTime dateTime1, DateTime dateTime2)
        {
            return (dateTime1.Year.Equals(dateTime2.Year) && dateTime1.Month.Equals(dateTime2.Month) && dateTime1.Day.Equals(dateTime2.Day));
        }

        #endregion

        #region Cifrado
        public class Cifrado
        {
            // Change these keys
            private static byte[] Key = { 99, 217, 19, 11, 24, 26, 85, 42, 46, 184, 27, 162, 37, 112, 222, 209, 241, 24, 175, 144, 173, 53, 123, 29, 24, 26, 17, 218, 131, 1, 88, 239 };
            private static byte[] Vector = { 136, 32, 191, 66, 213, 23, 113, 119, 112, 121, 252, 112, 79, 32, 114, 66 };

            private static ICryptoTransform EncryptorTransform, DecryptorTransform;
            private static System.Text.UTF8Encoding UTFEncoder;

            private static void Inicializar()
            {
                RijndaelManaged rm = new RijndaelManaged();

                if (EncryptorTransform == null)
                {
                    EncryptorTransform = rm.CreateEncryptor(Key, Vector);
                }

                if (DecryptorTransform == null)
                {
                    DecryptorTransform = rm.CreateDecryptor(Key, Vector);
                }

                UTFEncoder = new System.Text.UTF8Encoding();
            }

            static public byte[] GenerarClave()
            {
                RijndaelManaged rm = new RijndaelManaged();
                rm.GenerateKey();
                return rm.Key;
            }

            static public byte[] GenerarVector()
            {
                RijndaelManaged rm = new RijndaelManaged();
                rm.GenerateIV();
                return rm.IV;
            }

            public static string Cifrar(string TextValue)
            {
                Inicializar();
                return ByteArrToString(Encrypt(TextValue));
            }

            /// Método interno de cifrado con bytes
            private static byte[] Encrypt(string TextValue)
            {
                Byte[] bytes = UTFEncoder.GetBytes(TextValue);

                MemoryStream memoryStream = new MemoryStream();

                #region Escribir valores no cifrados en el stream
                CryptoStream cs = new CryptoStream(memoryStream, EncryptorTransform, CryptoStreamMode.Write);
                cs.Write(bytes, 0, bytes.Length);
                cs.FlushFinalBlock();
                #endregion

                #region Read encrypted value back out of the stream
                memoryStream.Position = 0;
                byte[] encrypted = new byte[memoryStream.Length];
                memoryStream.Read(encrypted, 0, encrypted.Length);
                #endregion

                cs.Close();
                memoryStream.Close();

                return encrypted;
            }

            public static string Descifrar(string EncryptedString)
            {
                Inicializar();
                return Decrypt(StrToByteArray(EncryptedString));
            }

            /// Método interno de descifrado con bytes
            private static string Decrypt(byte[] EncryptedValue)
            {
                #region Escribir valores cifrados en el stream
                MemoryStream encryptedStream = new MemoryStream();
                CryptoStream decryptStream = new CryptoStream(encryptedStream, DecryptorTransform, CryptoStreamMode.Write);
                decryptStream.Write(EncryptedValue, 0, EncryptedValue.Length);
                decryptStream.FlushFinalBlock();
                #endregion

                #region Leer los valores no cifrados del stream
                encryptedStream.Position = 0;
                Byte[] decryptedBytes = new Byte[encryptedStream.Length];
                encryptedStream.Read(decryptedBytes, 0, decryptedBytes.Length);
                encryptedStream.Close();
                #endregion
                return UTFEncoder.GetString(decryptedBytes);
            }

            private static byte[] StrToByteArray(string str)
            {
                byte val;
                byte[] byteArr = new byte[str.Length / 3];
                int i = 0;
                int j = 0;
                do
                {
                    val = byte.Parse(str.Substring(i, 3));
                    byteArr[j++] = val;
                    i += 3;
                }
                while (i < str.Length);
                return byteArr;
            }

            private static string ByteArrToString(byte[] byteArr)
            {
                byte val;
                string tempStr = "";
                for (int i = 0; i <= byteArr.GetUpperBound(0); i++)
                {
                    val = byteArr[i];
                    if (val < (byte)10)
                        tempStr += "00" + val.ToString();
                    else if (val < (byte)100)
                        tempStr += "0" + val.ToString();
                    else
                        tempStr += val.ToString();
                }
                return tempStr;
            }
        }
        #endregion

        #region Transformar Valor Cubicaje para enviar/recibir a/desde SAP
        
        /// <summary>
        /// Transforma el valor de cubicaje desde nuestra BBDD a SAP ya que al "lince" que pensó el desarrollo de peso volumétrico en su día no tomó en cuenta ni la estructura de la tabla en SAP ni como se tenian que pasar los valores (sin la unidad de medida).
        /// <para>Así que toca eliminar la unidad de medida del valor a envíar.</para>
        /// </summary>
        /// <param name="valorCubicaje"></param>
        /// <returns></returns>
        public static string TransformarValorCubicajeParaSAP(string valorCubicaje)
        {
            if(!string.IsNullOrEmpty(valorCubicaje))
            {
                if (valorCubicaje.IndexOf("Sin") >= 0) 
                {
                    valorCubicaje = "SIN";
                }
                else if (valorCubicaje.IndexOf(" ") >= 0)
                {
                    int indice = valorCubicaje.IndexOf(" ");
                    valorCubicaje = valorCubicaje.Remove(indice);
                }
            }
            
            return valorCubicaje;
        }

        /// <summary>
        /// Transformamos el valor de cubicaje de la oferta desde SAP a nuestra BBDD
        /// </summary>
        /// <param name="valorCubicaje"></param>
        /// <returns></returns>
        /// <remarks>TODO: Gestionar con CRM la devolución de la unidad de medida, así como repensar la sincronización un poco mejor, ahora mismo nos basamos en ECOFIN</remarks>
        public static string TransformarValorCubicajeDesdeSAP(string valorCubicaje) 
        {
            if (!string.IsNullOrEmpty(valorCubicaje))
            {
                int factor;
                if (valorCubicaje.IndexOf("SIN") >= 0)
                {
                    valorCubicaje = "Sin volumétrico";
                }
                else if(int.TryParse(valorCubicaje, out factor))
                {
                    valorCubicaje = valorCubicaje + " Kg/m3";
                }
            }

            return valorCubicaje;
        }

        #endregion
    }
}
