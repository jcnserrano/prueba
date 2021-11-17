using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.Persistence;
using System.Collections.ObjectModel;

namespace Correos.SimuladorOfertas.Business
{
    public class ListaPreciosBL
    {
        #region Métodos Obtener

        /// <summary>
        /// Método que obtiene los datos de la tabla maestra ListaPrecios
        /// </summary>
        /// <returns>Lista de entidades ListaPreciosBE</returns>
        public Collection<ListaPreciosBE> ObtenerListaPrecios()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ListaPreciosPersistence persistence = new ListaPreciosPersistence(uow);
                return persistence.ObtenerListaPrecios();
            }
        }

        #endregion
    }
}
