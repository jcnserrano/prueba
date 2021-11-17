using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.Persistence;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

namespace Correos.SimuladorOfertas.Business
{
    public class GradosBL
    {
        /// <summary>
        /// Obtiene la lista de grados del producto que se pide
        /// </summary>
        /// <param name="idProducto"></param>
        /// <returns></returns>
        public Collection<GradoProductoInformacionBE> ObtenerGradosProducto(Guid idProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                return ObtenerGradosProducto(idProducto, uow);
            }
        }

        /// <summary>
        /// Obtiene la lista de grados del producto que se pide
        /// </summary>
        /// <param name="idProducto"></param>
        /// <returns></returns>
        public Collection<GradoProductoInformacionBE> ObtenerGradosProducto(Guid idProducto, IUnitOfWork uow)
        {
            DescuentoPersistence persistence = new DescuentoPersistence(uow);
            return persistence.ObtenerGradosProducto(idProducto);
        }

    }
}
