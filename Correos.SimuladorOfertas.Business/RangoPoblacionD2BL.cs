using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.Persistence;
using System;
using System.Collections.ObjectModel;

namespace Correos.SimuladorOfertas.Business
{
    public class RangoPoblacionD2BL
    {
        #region Métodos Obtener
        /// <summary>
        /// Obtiene el Registro del Rango de poblacionD2
        /// </summary>
        /// <param name="idProducto"></param>
        /// <returns></returns>
        public Collection<RangoPoblacionD2BE> ObtenerRangoPoblacionD2(Guid idProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                return ObtenerRangoPoblacionD2(idProducto, uow);
            }
        }

        public Collection<RangoPoblacionD2BE> ObtenerRangoPoblacionD2(Guid idProducto, IUnitOfWork uow)
        {
                RangoPoblacionD2Persistence persistence = new RangoPoblacionD2Persistence(uow);
                return persistence.ObtenerRangoPoblacionD2(idProducto);
            }
        #endregion
        #region Métodos Guardar
         /// <summary>
        /// Inserta un nuevo registro en base de datos
        /// </summary>
        /// <param name="rangoPoblacionD2"></param>
        public void InsertRangosPoblacionD2(RangoPoblacionD2BE rangoPoblacionD2) 
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                RangoPoblacionD2Persistence persistence = new RangoPoblacionD2Persistence(uow);
                persistence.InsertRangosPoblacionD2(rangoPoblacionD2);
            }
        }

        #endregion
    }
}
