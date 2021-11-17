using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.Persistence;
using System;
using System.Collections.ObjectModel;

namespace Correos.SimuladorOfertas.Business
{
    public class ConfiguracionGruposTramoBL
    {
        #region Métodos Obtener

        /// <summary>
        /// Método que obtiene la lista de ConfiguracionGruposTramoOfertaBE para el producto oferta seleccionado
        /// </summary>
        /// /// <param name="producto">Entidad de productoBE</param>
        /// <param name="idProductoOferta">Identificador del producto oferta</param>        
        /// <returns>Lista de ConfiguracionGruposTramoOfertaBE</returns>
        public Collection<GrupoTramoBE> ObtenerListaGruposTramoOferta(ProductoBE producto, Guid idProductoOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ConfiguracionGruposTramoOfertaPersistence persistence = new ConfiguracionGruposTramoOfertaPersistence(uow);
                return persistence.ObtenerListaGruposTramoOferta(producto, idProductoOferta);
            }
        }

        /// <summary>
        /// Método que obtiene una colección de registros de la tabla configuracionGruposTramoOferta de un listado de identificadores de productos oferta
        /// </summary>
        /// <param name="listaIdsProductoOferta">Listado de identificadores de productoOferta</param>
        /// <returns>Colección de entidades GrupoTramoBE</returns>
        public Collection<GrupoTramoBE> ObtenerListaConfiguracionGruposTramoOferta(Collection<Guid> listaIdsProductoOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ConfiguracionGruposTramoOfertaPersistence persistence = new ConfiguracionGruposTramoOfertaPersistence(uow);
                return persistence.ObtenerListaConfiguracionGruposTramoOferta(listaIdsProductoOferta);
            }
        }

        /// <summary>
        /// Método que obtiene una copia de las configuraciones de grupos de tramo de un producto oferta
        /// </summary>
        /// <param name="idProductoOfertaOrigen">Identificador del producto oferta origen</param>
        /// <param name="idProductoOfertaDestino">Identificador del producto oferta destino</param>
        /// <returns>Colección de entidades GrupoTramoBE</returns>
        public Collection<GrupoTramoBE> ObtenerCopiaListaConfiguracionGruposTramoOferta(Guid idProductoOfertaOrigen, Guid idProductoOfertaDestino)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ConfiguracionGruposTramoOfertaPersistence persistence = new ConfiguracionGruposTramoOfertaPersistence(uow);
                return persistence.ObtenerCopiaListaConfiguracionGruposTramoOferta(idProductoOfertaOrigen, idProductoOfertaDestino);
            }
        }

        #endregion
    }
}
