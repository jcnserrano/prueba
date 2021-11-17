using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;

namespace Correos.SimuladorOfertas.Common.Extensions
{

    /// <summary>
    /// Agrupa funciones comunes para trabajar con Linq
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
    public static class LinqUtil
    {
        /// <summary>
        /// Ordena un enumerable según una expresión de ordenación
        /// </summary>
        /// <typeparam name="T">Tipo de objetos de la enumeración</typeparam>
        /// <param name="source">Objeto a ordenar</param>
        /// <param name="sortExpression">Expresión de ordenación</param>
        /// <returns>Enumeración ordenada</returns>
        public static IEnumerable<T> Sort<T>(this IEnumerable<T> source, string sortExpression)
        {
            string[] sortParts = sortExpression.Split(' ');
            var param = Expression.Parameter(typeof(T), string.Empty);
            try
            {
                var property = Expression.Property(param, sortParts[0]);
                var sortLambda = Expression.Lambda<Func<T, object>>(Expression.Convert(property, typeof(object)), param);

                if (sortParts.Length > 1 && sortParts[1].Equals("desc", StringComparison.OrdinalIgnoreCase))
                {
                    return source.AsQueryable<T>().OrderByDescending<T, object>(sortLambda);
                }
                return source.AsQueryable<T>().OrderBy<T, object>(sortLambda);
            }
            catch (ArgumentException)
            {
                return source;
            }
        } 
    }
}
