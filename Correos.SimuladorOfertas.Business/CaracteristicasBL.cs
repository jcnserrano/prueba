using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections.ObjectModel;
using Correos.SimuladorOfertas.Persistence;
using Correos.SimuladorOfertas.DTOs;

namespace Correos.SimuladorOfertas.Business
{
    public class CaracteristicasBL
    {
        #region Métodos Obtener

        /// <summary>
        /// Método que obtiene las características de un valor añadido con su lista de valores
        /// </summary>
        /// <param name="idValorAnadido">identificador del valor añadido</param>
        /// <returns>Colección de entidades CaracteristicaBE</returns>
        public Collection<CaracteristicaBE> ObtenerCaracteristicasByidValorAnadido(Guid idValorAnadido)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                CaracteristicaPersistence caracteristicaPersistence = new CaracteristicaPersistence(uow);
                return caracteristicaPersistence.ObtenerCaracteristicasByidValorAnadido(idValorAnadido);
            }
        }

        /// <summary>
        /// Método que obtiene las características de un valor añadido con su lista de valores
        /// </summary>
        /// <param name="idValorAnadido">identificador del valor añadido</param>
        /// <returns>Colección de entidades CaracteristicaBE</returns>
        public Collection<ValoresBE> ObtenerValoresCaracteristica(Guid idCaracteristica)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                CaracteristicaPersistence caracteristicaPersistence = new CaracteristicaPersistence(uow);
                return caracteristicaPersistence.ObtenerValoresCaracteristica(idCaracteristica);
            }
        }

        /// <summary>
        /// Método que obtiene las características de un valor añadido con su lista de valores
        /// </summary>
        /// <param name="idValorAnadido">identificador del valor añadido</param>
        /// <returns>Colección de entidades CaracteristicaBE</returns>
        public Collection<CaracteristicaBE> ObtenerCaracteristicasProductoOferta(Guid idProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                CaracteristicaPersistence caracteristicaPersistence = new CaracteristicaPersistence(uow);
                return caracteristicaPersistence.ObtenerCaracteristicasDefinicionProducto(idProducto);
            }
        }

        

        #endregion
    }
}
