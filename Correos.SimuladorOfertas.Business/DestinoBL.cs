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
    public class DestinoBL
    {
        #region Métodos Obtener

        public Dictionary<Guid, Collection<DestinoBE>> ObtenerDestinosProductos(Collection<Guid> listaIdsProductos)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                DestinoPersistence destinoPersistence = new DestinoPersistence(uow);
                return destinoPersistence.ObtenerDestinosProductos(listaIdsProductos);
            }
        }

        public List<DestinoBE> ObtenerDestinosProductoConDescripcion(Guid idProducto, String codProductoSAP)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                DestinoPersistence destinoPersistence = new DestinoPersistence(uow);
                return destinoPersistence.ObtenerDestinosProductoConDescripcion(idProducto, codProductoSAP);
            }
        }

        public List<DestinoBE> ObtenerDestinosProductoOfertaConDescripcion(Guid idProductoOferta, Guid idProducto, String codProductoSAP)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                DestinoPersistence destinoPersistence = new DestinoPersistence(uow);
                return destinoPersistence.ObtenerDestinosProductoOfertaConDescripcion(idProductoOferta,  idProducto, codProductoSAP);
            }
        }

        
        /// <summary>
        /// Método que obtiene los destinos con visibilidad para el producto pasado por parámetro
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <returns>Lista de DestinoBE</returns>
        public List<DestinoBE> ObtenerDestinosProductoConVisibilidad(Guid idProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                DestinoPersistence destinoPersistence = new DestinoPersistence(uow);
                return destinoPersistence.ObtenerDestinosProductoConVisibilidad(idProducto);
            }
        }

        /// <summary>
        /// Método que obtiene los destinos con visibilidad para el producto pasado por parámetro
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <returns>Lista de DestinoBE</returns>
        public List<DestinoBE> ObtenerDestinosProductoPorDefecto(Guid idProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                DestinoPersistence destinoPersistence = new DestinoPersistence(uow);
                return destinoPersistence.ObtenerDestinosProductoPorDefecto(idProducto);
            }
        }

        #endregion
    }
}
