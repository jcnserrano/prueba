using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Correos.SimuladorOfertas.Common.Math
{
    /// <summary>
    /// Contiene métodos de ayuda para redondear números
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
    public static class MathHelper
    {
        #region Propiedades

        private const MidpointRounding MIDPOINTROUNDING = MidpointRounding.AwayFromZero;

        #endregion

        #region Métodos públicos

        /// <summary>
        /// Redondea el valor hacia el entero más cercano y más alejado de cero. 
        /// </summary>
        /// <param name="value">Número a redondear</param>
        /// <returns></returns>
        public static Decimal Round(decimal value)
        {
            return System.Math.Round(value, MIDPOINTROUNDING);
        }

        /// <summary>
        /// Redondea el valor hacia el número más cercano y más alejado de cero
        /// con el número de decimales que se le pasen como parámetro. 
        /// </summary>
        /// <param name="value">Número a redondear</param>
        /// <param name="decimals">Número de decimales</param>
        /// <returns></returns>
        public static Decimal Round(decimal value, int decimals)
        {
            return System.Math.Round(value, decimals, MIDPOINTROUNDING);
        }

        /// <summary>
        /// Redondea el valor hacia el entero más cercano y más alejado de cero. 
        /// </summary>
        /// <param name="value">Número a redondear</param>
        /// <returns></returns>
        public static double Round(double value)
        {
            return System.Math.Round(value, MIDPOINTROUNDING);
        }

        /// <summary>
        /// Redondea el valor hacia el número más cercano y más alejado de cero
        /// con el número de decimales que se le pasen como parámetro. 
        /// </summary>
        /// <param name="value">Número a redondear</param>
        /// <param name="decimals">Número de decimales</param>
        /// <returns></returns>
        public static double Round(double value, int decimals)
        {
            return System.Math.Round(value, decimals, MIDPOINTROUNDING);
        }

        #endregion
    }
}
