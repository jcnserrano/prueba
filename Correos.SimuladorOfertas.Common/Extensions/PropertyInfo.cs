using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;
using System.Reflection;

namespace Correos.SimuladorOfertas.Common.Extensions
{
    /// <summary>
    /// Agrupa funciones comunes para trabajar con PropertyInfo
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
    public static class PropertyInfoUtil
    {
        public static Func<object, T> GetValueGetter<T>(this PropertyInfo propertyInfo)
        {
            var instance = Expression.Parameter(typeof(Object), "i");
            var castedInstance = Expression.ConvertChecked(instance, propertyInfo.DeclaringType);
            var property = Expression.Property(castedInstance, propertyInfo);
            var convert = Expression.Convert(property, typeof(T));
            var expression = Expression.Lambda(convert, instance);
            return (Func<object, T>)expression.Compile();
        }

        public static Action<object, T> GetValueSetter<T>(this PropertyInfo propertyInfo)
        {
            var instance = Expression.Parameter(typeof(Object), "i");
            var castedInstance = Expression.ConvertChecked(instance, propertyInfo.DeclaringType);
            var argument = Expression.Parameter(typeof(T), "a");
            var setterCall = Expression.Call(
              castedInstance,
              propertyInfo.GetSetMethod(),
              Expression.Convert(argument, propertyInfo.PropertyType));
            return (Action<object, T>)Expression.Lambda(setterCall, instance, argument)
                              .Compile();
        }
    }
}
