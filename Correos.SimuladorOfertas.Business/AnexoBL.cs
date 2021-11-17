using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.Persistence;
using System;
using System.Collections.ObjectModel;

namespace Correos.SimuladorOfertas.Business
{
    public class AnexoBL
    {
        /// <summary>
        /// Obtiene toda la información necesaria para el árbol de productos
        /// </summary>
        /// <returns></returns>
        public Collection<AnexoBE> ObtenerArbolProductos()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                AnexoPersistence objPersistencia = new AnexoPersistence(uow);
                return objPersistencia.ObtenerArbolProductos();
            }
        }

        /// <summary>
        /// Obtiene toda la información necesaria para el árbol de productos
        /// </summary>
        /// <returns></returns>
        public Collection<AnexoBE> ObtenerArbolProductosSincronizar(Guid idOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                AnexoPersistence objPersistencia = new AnexoPersistence(uow);
                return objPersistencia.ObtenerArbolProductosSincronizar(idOferta);
            }
        }
    }
}
