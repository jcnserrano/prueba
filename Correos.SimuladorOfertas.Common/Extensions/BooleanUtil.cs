using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Correos.SimuladorOfertas.Common.Extensions
{
    /// <summary>
    /// Agrupa funciones comunes para trabajar con Booleanos
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
    public static class BooleanUtils
    {
        /// <summary>
        /// Para convertir el Bool = true de .Net usado como Char = 1 de Oracle.
        /// </summary>
        /// <param name="s"></param>
        /// <returns>1--> True</returns>
        public static string ToOracleCharFromBoolean(this bool b)
        {
            if (b == true) return StringUtil.ORACLE_TRUE_VALUE;
            else return StringUtil.ORACLE_FALSE_VALUE;
        }
    }
}