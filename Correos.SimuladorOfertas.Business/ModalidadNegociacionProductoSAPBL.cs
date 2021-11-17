using Correos.SimuladorOfertas.Persistence;
using System;

namespace Correos.SimuladorOfertas.Business
{
    public class ModalidadNegociacionProductoSAPBL
    {
        #region Métodos Obtener

        /// <summary>
        /// Devuelve el identificador de una modalidad de negociacion a partir de sus FK
        /// </summary>
        /// <param name="codProductoSAP"></param>
        /// <param name="codModalidadNegociacion"></param>
        /// <returns></returns>
        public Guid? ObtenerIdentificadorModalidadNegocioProductoByFK(string codProductoSAP, string codModalidadNegociacion)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ModalidadNegociacionProductoSAPPersistence objPersistencia = new ModalidadNegociacionProductoSAPPersistence(uow);
                return objPersistencia.ObtenerModalidaNegociacionByFK(codProductoSAP, codModalidadNegociacion);
            }
        }

        #endregion
    }
}
