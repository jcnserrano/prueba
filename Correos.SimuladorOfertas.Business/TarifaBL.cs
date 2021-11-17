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
    public class TarifaBL
    {
        #region Métodos Obtener

        /// <summary>
        /// Obtener todas las tarifas para la lista de tramos
        /// </summary>
        /// <param name="listaIdsProductos"></param>
        /// <returns></returns>
        public Dictionary<Guid, Collection<TarifaBE>> ObtenerTarifasTramos(Collection<Guid> listaIdsTramos)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                TarifaPersistence destinoPersistence = new TarifaPersistence(uow);
                return destinoPersistence.ObtenerTarifasTramos(listaIdsTramos);
            }
        }

        public Guid ObtenerIdTarifaValorAnadido(Guid idValorAnadido, Guid idValorAnadidoProducto, string combinacion, string tipoPrecioDe = "")
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                TarifaPersistence tarifaPersistence = new TarifaPersistence(uow);
                return tarifaPersistence.ObtenerIdTarifaValorAnadido(idValorAnadido, idValorAnadidoProducto, combinacion, tipoPrecioDe);
            }
        }

        public Guid ObtenerIdTarifaValorAnadidoProducto(Guid idValorAnadido, Guid idValorAnadidoProducto, Guid idProducto, string combinacion, String tipoTarifa, String tipoPrecioDe)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                TarifaPersistence tarifaPersistence = new TarifaPersistence(uow);
                return tarifaPersistence.ObtenerIdTarifaValorAnadidoProducto(idValorAnadido, idValorAnadidoProducto, idProducto, combinacion, tipoTarifa, tipoPrecioDe);
            }
        }

        public TarifaBE ObtenerTarifaValorAnadido(Guid idTarifaValorAnadido)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                TarifaPersistence tarifaPersistence = new TarifaPersistence(uow);
                return tarifaPersistence.ObtenerTarifaVAByIdTarifaValorAnadido(idTarifaValorAnadido);
            }
        }

        public Collection<TarifaBE> ObtenerTarifasValorAnyadidoProducto(Guid idProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                TarifaPersistence tarifaPersistence = new TarifaPersistence(uow);
                return tarifaPersistence.ObtenerTarifasValorAnyadidoProducto(idProducto);
            }
        }

        public Guid ObtenerIdValorAnadidoProductoByIdProductoCodVASAP(Guid idProducto, String codVASAP)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ValorAnadidoProductoPersistence vapPersistence = new ValorAnadidoProductoPersistence(uow);
                Guid idValorAnadidoProducto = vapPersistence.ObtenerIdValorAnadidoProductoByIdProductoCodVASAP(idProducto, codVASAP);
                return idValorAnadidoProducto;
            }            
        }


        #endregion
    }
}
