using System;
using System.Collections.ObjectModel;
using Correos.SimuladorOfertas.Common;
using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.Persistence;
using System.Diagnostics;
using System.Linq;
using System.Collections.Generic;

namespace Correos.SimuladorOfertas.Business
{
    public class TramoBL
    {
        #region Métodos Obtener

        /// <summary>
        /// Obtener todos los tramos para la lista de destinos
        /// </summary>
        /// <param name="listaIdsProductos"></param>
        /// <returns></returns>
        public Dictionary<Guid, Collection<TramoBE>> ObtenerTramosDestinos(Collection<Guid> listaIdsDestinos)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                TramoPersistence destinoPersistence = new TramoPersistence(uow);
                return destinoPersistence.ObtenerTramosDestinos(listaIdsDestinos);
            }
        }

        #endregion
    }
}
