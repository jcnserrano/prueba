using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.Persistence;
using System;
using System.Collections.ObjectModel;

namespace Correos.SimuladorOfertas.Business
{
    public class ConfiguracionCaracteristicasBL
    {
        /// <summary>
        /// Obtiene las características de asociadas a un producto oferta 
        /// </summary>
        /// <param name="idProducto"></param>
        /// <param name="idProductoOferta"></param>
        /// <returns></returns>
        public Collection<ConfiguracionCaracteristicaBE> ObtenerCaracteristicasOferta(Guid idProducto, Guid idProductoOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                CaracteristicaPersistence persistence = new CaracteristicaPersistence(uow);
                return persistence.ObtenerCaracteristicaProducto(idProducto, idProductoOferta);
            }
        }

        /// <summary>
        /// Método que obtiene una colección de registros ConfiguracionCaracteristicaBE de los productos Oferta pasados por parámetro.
        /// </summary>
        /// <param name="listaIdsProductoOferta">Listado de identificadores de productos Oferta</param>
        /// <returns>Colección de entidades ConfiguracionCaracteristicaBE</returns>
        public Collection<ConfiguracionCaracteristicaBE> ObtenerConfiguracionCaracteristicasOferta(Collection<Guid> listaIdsProductoOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                CaracteristicaPersistence persistence = new CaracteristicaPersistence(uow);
                return persistence.ObtenerConfiguracionCaracteristicasOferta(listaIdsProductoOferta);
            }
        }

        /// <summary>
        /// Método que copia la configuración de características de un produto oferta
        /// </summary>
        /// <param name="idProductoOfertaOrigen">Identificador del producto oferta origen</param>
        /// <param name="idProductoOfertaDestino">Identificador del producto oferta destino</param>
        /// <returns>Colección de entidades ConfiguracionCaracteristicaBE</returns>
        public Collection<ConfiguracionCaracteristicaBE> CopiarConfiguracionCaracteristicasProductoOferta(Guid idProductoOfertaOrigen, Guid idProductoOfertaDestino)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                CaracteristicaPersistence persistence = new CaracteristicaPersistence(uow);
                return persistence.CopiarConfiguracionCaracteristicasProductoOferta(idProductoOfertaOrigen, idProductoOfertaDestino);
            }
        }

        /// <summary>
        /// Método que obtiene las características de un producto
        /// </summary>
        /// <param name="idproducto">identificador del producto</param>
        /// <returns>Colección de entidades ConfiguracionCaracteristicaBE</returns>
        public Collection<ConfiguracionCaracteristicaBE> ObtenerCaracteristicasProducto(Guid idproducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                return ObtenerCaracteristicasProducto(idproducto, uow);
            }
        }

        /// <summary>
        /// Método que obtiene las características de un producto
        /// </summary>
        /// <param name="idproducto">identificador del producto</param>
        /// <returns>Colección de entidades ConfiguracionCaracteristicaBE</returns>
        public Collection<ConfiguracionCaracteristicaBE> ObtenerCaracteristicasProducto(Guid idproducto, IUnitOfWork uow)
        {
                CaracteristicaPersistence persistence = new CaracteristicaPersistence(uow);
                return persistence.ObtenerCaracteristicasProducto(idproducto);
            }

        /// <summary>
        /// Método que obtiene los registros de la tabla CaracteristicaProducto de un producto
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <returns>Colección de entidades ConfiguracionCaracteristicaBE</returns>
        public Collection<ConfiguracionCaracteristicaBE> ObtenerRelacionCaracteristicasProducto(Guid idProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                return ObtenerRelacionCaracteristicasProducto(idProducto, uow);
            }
        }

        /// <summary>
        /// Método que obtiene los registros de la tabla CaracteristicaProducto de un producto
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <returns>Colección de entidades ConfiguracionCaracteristicaBE</returns>
        public Collection<ConfiguracionCaracteristicaBE> ObtenerRelacionCaracteristicasProducto(Guid idProducto, IUnitOfWork uow)
        {
            CaracteristicaPersistence persistence = new CaracteristicaPersistence(uow);
            return persistence.ObtenerRelacionCaracteristicasProducto(idProducto);
        }
    }
}
