using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections.ObjectModel;
using System.Text.RegularExpressions;

namespace Correos.SimuladorOfertas.Common.Extensions
{
    /// <summary>
    /// Agrupa funciones comunes para trabajar con Strings
    /// </summary>
    /// <unit>Common</unit>
    /// <project>Correos.SimuladorOfertas</project>
    /// <company>Correos</company>
    /// <date>01/10/2014</date>
    /// <author>Jose Antonio Pardos</author>
    /// <version>1.0</version>
    /// <history>
    /// 	Versión		|Date		        |Author					        Description
    /// 	1.0			01/10/2014		    Jose Antonio Pardos     		Versión inicial
    /// </history>  
    public static class StringUtil
    {
        #region Constantes
        
        public const string ORACLE_TRUE_VALUE = "1";
        public const string ORACLE_FALSE_VALUE = "0";
        public const int MAX_LENGTH_STRING = 2000;
        public const string SEPARATOR_IDS = ",";
        
        #endregion

        /// <summary>
        /// Para convertir el char = 1 de Oracle usado como Boolean. 1--> true
        /// </summary>
        /// <param name="s"></param>
        /// <returns> True</returns>
        public static bool ToBooleanFromOracleChar(this string s)
        {   
            if (s.CompareTo(ORACLE_TRUE_VALUE) == 0) return true;
            else return false;
        }

        /// <summary>
        /// Dada una cadena de carácteres que contiene ids separados por comas, 
        /// devuelve una colección de cadenas que contienen los identificadores concatenados. 
        /// Dichas cadenas no pueden superar la longitud de MAX_LENGTH_STRING (2000) caracteres
        /// </summary>
        /// <param name="chain">Cadena de carácteres con los ids separados por comas</param>
        /// <returns>Colección de cadenas con la concatenación de identificadores</returns>
        public static Collection<string> SplitIdentifiersChain(this string chain)
        {
            #region Tratamiento de la cadena de identificadores

            //Creamos una lista de identificadores (utilizaremos colecciones de listas cuyas longitudes no superen MAX_LENGTH_STRING) 
            //La lista de identificadores será la concatenación de los identificadores separados por comas
            //Nota: añadimos un separador al final de la cadena para simplificar el tratamiento
            Collection<string> collectionStrIds = new Collection<string>();
            string strRest = chain;
            while (strRest.Length >= MAX_LENGTH_STRING)
            {
                //Obtenermos la cadena a tratar en la iteración
                string strAux = strRest.Substring(0, MAX_LENGTH_STRING);
                int lastSeparatorPosition = strAux.LastIndexOf(SEPARATOR_IDS);

                //Añadimos la cadena tratada a la colección
                collectionStrIds.Add(strAux.Substring(0, lastSeparatorPosition).TrimEnd(SEPARATOR_IDS.ToCharArray()));
                //Establecemos el resto de la cadena a tratar
                strRest = strRest.Substring(lastSeparatorPosition + 1);
            }
            //Añadimos lo que queda de la cadena a la colección
            collectionStrIds.Add(strRest.TrimEnd(SEPARATOR_IDS.ToCharArray()));

            #endregion

            //Eliminamos las cadenas vacías
            collectionStrIds = new Collection<string>(collectionStrIds.Where(str => !String.IsNullOrEmpty(str)).ToList());

            return collectionStrIds;
        }

        /// <summary>
        /// Dada una cadena de caracteres, sustituye las vocales acentuadas por su equivalente sin tildar
        /// </summary>
        /// <param name="inputString">Cadena a tratar</param>
        /// <returns>Cadena tratada</returns>
        public static String RemoveAccents(this String inputString)
        {
            Regex replace_a_Accents = new Regex("[á|à|ä|â]", RegexOptions.Compiled);
            Regex replace_A_Accents = new Regex("[Á|À|Ä|Â]", RegexOptions.Compiled);
            Regex replace_e_Accents = new Regex("[é|è|ë|ê]", RegexOptions.Compiled);
            Regex replace_E_Accents = new Regex("[É|È|Ë|Ê]", RegexOptions.Compiled);
            Regex replace_i_Accents = new Regex("[í|ì|ï|î]", RegexOptions.Compiled);
            Regex replace_I_Accents = new Regex("[Í|Ì|Ï|Î]", RegexOptions.Compiled);
            Regex replace_o_Accents = new Regex("[ó|ò|ö|ô]", RegexOptions.Compiled);
            Regex replace_O_Accents = new Regex("[Ó|Ò|Ö|Ô]", RegexOptions.Compiled);
            Regex replace_u_Accents = new Regex("[ú|ù|ü|û]", RegexOptions.Compiled);
            Regex replace_U_Accents = new Regex("[Ú|Ù|Ü|Û]", RegexOptions.Compiled);
            inputString = replace_a_Accents.Replace(inputString, "a");
            inputString = replace_A_Accents.Replace(inputString, "A");
            inputString = replace_e_Accents.Replace(inputString, "e");
            inputString = replace_E_Accents.Replace(inputString, "E");
            inputString = replace_i_Accents.Replace(inputString, "i");
            inputString = replace_I_Accents.Replace(inputString, "I");
            inputString = replace_o_Accents.Replace(inputString, "o");
            inputString = replace_O_Accents.Replace(inputString, "O");
            inputString = replace_u_Accents.Replace(inputString, "u");
            inputString = replace_U_Accents.Replace(inputString, "U");
            return inputString;
        }

        /// <summary>
        /// Dadas dos cadenas a y b, devuelve true en caso de que sean iguales, gestionando el nullOrEmpty
        /// </summary>
        /// <param name="a">cadena a</param>
        /// <param name="b">cadena b</param>
        /// <returns>True si son iguales</returns>
        public static bool AreEqual(string a, string b)
        {
            if (string.IsNullOrEmpty(a))
            {
                return string.IsNullOrEmpty(b);
            }
            else
            {
                return string.Equals(a, b);
            }
        }

        /// <summary>
        /// Si todas son maýusculas, convierte la cadena a minúsculas poniendo la primera a mayúsculas
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string ToTitleCaseIfAllUpper(string str)
        {          
            if (str.Replace(" ", "").All(c => char.IsUpper(c)))
            {                
               str = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(str.ToLower());
            }
            return str;
        }


    }
}
