using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Correos.SimuladorOfertas.Common.Extensions
{
    public static class ExtensionMethods
    {
        /// <summary>
        /// Método que convierte una lista System.Collections.Generic.List en System.Collection.ObjectModel.Collection
        /// </summary>
        /// <param name="items">Objeto List a convertir</param>
        /// <returns>Objeto Collection resultante</returns>
        public static Collection<T> ToCollection<T>(this List<T> items)
        {
            Collection<T> collection = new Collection<T>();

            for (int i = 0; i < items.Count; i++)
            {
                collection.Add(items[i]);
            }

            return collection;
        }
    }
}
