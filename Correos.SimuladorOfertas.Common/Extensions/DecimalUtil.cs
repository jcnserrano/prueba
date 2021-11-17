using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Correos.SimuladorOfertas.Common.Extensions
{
    /// <summary>
    /// Agrupa funciones comunes para trabajar con decimales
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
    public static class DecimalUtil
    {
        /// <summary>
        /// Transforma de kilowatios a metros cúbicos
        /// </summary>
        /// <param name="kwh">kilovatios a transformar</param>
        /// <param name="pcs">factor de conversión</param>
        /// <returns>metros cúbicos correspondientes a esos kilowatios</returns>
        public static Decimal KWhToM3(this Decimal kwh, Decimal pcs)
        {
            Decimal m3 = (pcs == 0) ?  0 : (kwh / pcs);
            return m3;    
        }

        /// <summary>
        /// Transforma de metros cúbicos a kilowatios
        /// </summary>
        /// <param name="m3">metros cúbicos a transformar</param>
        /// <param name="pcs">factor de conversión</param>
        /// <returns>kilowatios correspondientes a esos metros cúbicos</returns>
        public static Decimal M3ToKwh(this Decimal m3, Decimal pcs)
        {
            return m3 * pcs;
        }

        /// <summary>
        /// Convierte de KWh a MWh
        /// </summary>
        /// <param name="kwh">Kilowatios hora a convertir</param>
        /// <returns>MW hora equivalentes</returns>
        public static Decimal KWhToMWh(this decimal kwh)
        {
            return kwh / 1000;
        }

        /// <summary>
        /// Convierte de MWh a KWh
        /// </summary>
        /// <param name="mwh">Megawatios hora a convertir</param>
        /// <returns>KW hora equivalentes</returns>
        public static Decimal MWhToKWh(this decimal mwh)
        {
            return mwh * 1000;
        }

        /// <summary>
        /// Convierte de KWh a GWh
        /// </summary>
        /// <param name="kwh">KW hora a convertir</param>
        /// <returns>GW hora equivalentes</returns>
        public static Decimal KWhToGWh(this decimal kwh)
        {
            return kwh / 1000000;
        }

        /// <summary>
        /// Convierte de GWh a KWh
        /// </summary>
        /// <param name="gwh">GW hora a convertir</param>
        /// <returns>KW hora equivalentes</returns>
        public static Decimal GWhToKWh(this decimal gwh)
        {
            return gwh * 1000000;
        }
   
        /// <summary>
        /// Convierte el volumen que se encuentra a la temperatura indicada como temperatura origen a la temperatura indicada como temperatura de destino
        /// </summary>
        /// <param name="value">Volumen a convertir</param>
        /// <param name="sourceConversionFactor">Temperatura a la que se encuentra el volumen</param>
        /// <param name="targetConversionFactor">Temperatura de destino a la que se convertirá el volumen</param>
        /// <returns>Volumen convertido de la temperatura origen a la temperatura de destino</returns>
        public static Decimal KWhBetweenTemperatures(this decimal value, decimal sourceConversionFactor, decimal targetConversionFactor)
        {
            return (value / targetConversionFactor) * sourceConversionFactor;
        }

    }
}
