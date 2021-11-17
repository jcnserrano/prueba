using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.Persistence;
using System;
using System.Collections.ObjectModel;
using Correos.SimuladorOfertas.InOutLight;
using System.Linq;
using System.Collections.Generic;

namespace Correos.SimuladorOfertas.Business
{
    public class TipologiaClienteBL
    {
        #region Métodos Obtener

        /// <summary>
        /// Método que obtiene la tipología para el producto pasado por parámetro en función de su facturación
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <param name="facturacion">Facturación neta del producto</param>
        /// <returns></returns>
        public string ObtenerTipologiaProducto(Guid idProducto, double facturacion, string usuario)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                //Obtenemos la potencialidad del usuario para ese producto                
                PotencialidadPersistence persistencePotencial = new PotencialidadPersistence(uow);
                PotencialidadBE objPotencial = persistencePotencial.ObtenerPotencialidadProducto(idProducto, usuario);
                if (objPotencial != null)
                {
                    //obtiene la lista de tipologías para el producto segun la potencialidad
                    TipologiaClientePersistence persistence = new TipologiaClientePersistence(uow);

                    //La lista se ha obtenido ordenada por facturación de forma inversa
                    //Se va evaluando de tipologías más altas a las más bajas hasta encontrar la tipología correcta
                    foreach (TipologiaClienteBE tipologia in persistence.ObtenerTipologiasProducto(idProducto, objPotencial.Potencialidad))
                    {
                        decimal auxFacturacion = 0;
                        decimal.TryParse(facturacion.ToString(), out auxFacturacion);

                        //Si la facturación del producto es superior a la tipología actual, esta es la tipología correcta
                        if (auxFacturacion > tipologia.Facturacion)
                        {
                            return tipologia.TipoCliente;
                        }
                    }
                }

                return "E";
            }
        }

        /// <summary>
        /// Método que obtiene las tipologías de
        /// </summary>
        /// <param name="idProducto"></param>
        /// <returns></returns>
        public Collection<TipologiaClienteBE> ObtenerTipologiasProducto(Guid idProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                return ObtenerTipologiasProducto(idProducto, uow);
            }            
        }

        /// <summary>
        /// Método que obtiene las tipologías de
        /// </summary>
        /// <param name="idProducto"></param>
        /// <returns></returns>
        public Collection<TipologiaClienteBE> ObtenerTipologiasProducto(Guid idProducto, IUnitOfWork uow)
        {
            TipologiaClientePersistence tipologiaClientePersistence = new TipologiaClientePersistence(uow);
            return tipologiaClientePersistence.ObtenerTipologiasProducto(idProducto);
        }

        /// <summary>
        /// Método que obtiene la tipología para el producto pasado por parámetro en función de su facturación
        /// </summary>
        /// <param name="idValorAnadidoProducto">Identificador del valor añadido</param>
        /// <param name="facturacion">Facturación neta del producto</param>
        /// <returns></returns>
        public string ObtenerTipologiaValorAnadido(Guid idValorAnadidoProducto, double facturacion, string usuario)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                //Obtenemos la potencialidad del usuario para ese VA                
                PotencialidadPersistence persistencePotencial = new PotencialidadPersistence(uow);
                PotencialidadBE objPotencial = persistencePotencial.ObtenerPotencialidadVA(idValorAnadidoProducto, usuario);
                if (objPotencial != null)
                {
                    //Primeramente se obtiene la lista de tipologías para el valor añadido
                    TipologiaClientePersistence persistence = new TipologiaClientePersistence(uow);

                    //La lista se ha obtenido ordenada por facturación de forma inversa
                    //Se va evaluando de tipologías más altas a las más bajas hasta encontrar la tipología correcta
                    foreach (TipologiaClienteBE tipologia in persistence.ObtenerTipologiasValorAnadido(idValorAnadidoProducto, objPotencial.Potencialidad))
                    {
                        decimal auxFacturacion = 0;
                        decimal.TryParse(facturacion.ToString(), out auxFacturacion);

                        //Si la facturación del producto es superior a la tipología actual, esta es la tipología correcta
                        if (auxFacturacion > tipologia.Facturacion)
                        {
                            return tipologia.TipoCliente;
                        }
                    }
                }

                return "E";
            }
        }

        /// <summary>
        /// Método que devuelve una colección de registros de la tabla tipologíasValorAnadido de un producto
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <returns>Colección de entidades TipologiaClienteBE</returns>
        public Collection<TipologiaClienteBE> ObtenerTipologiasValorAnadidoByIdProducto(Guid idProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                return ObtenerTipologiasValorAnadidoByIdProducto(idProducto, uow);
            }
        }

        public Collection<TipologiaClienteBE> ObtenerTipologiasValorAnadidoByIdProducto(Guid idProducto, IUnitOfWork uow)
        {
            TipologiaClientePersistence tipologiaClientePersistence = new TipologiaClientePersistence(uow);
            return tipologiaClientePersistence.ObtenerTipologiasValorAnadidoByIdProducto(idProducto);
        }

        public Collection<AgrupacionTipologiaBE> ObtenerAgrupacionesTipologia()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                TipologiaClientePersistence tipologiaClientePersistence = new TipologiaClientePersistence(uow);
                return tipologiaClientePersistence.ObtenerAgrupacionesTipologia();
            }
        }

        /// <summary>
        /// <para>Método que obtiene las agrupaciones de tipologías de SAP y las guarda en la base de datos local.</para>
        /// <para>Se comprueba si ha habido cambios en los valores de las agrupaciones en SAP y se recalculan las ofertas si estas tienen productos afectados</para>
        /// </summary>
        /// <param name="usuario">Identificador del usuario</param>
        /// <param name="password">Contraseña del usuario</param>
        /// <param name="FechaHora">Fecha del sistema en la que se realiza la petición</param>
        /// <param name="uow">Objeto transaccional</param>
        public bool ObtenerAgrupacionesTipologiasDeSAP(string usuario, string password, DateTime FechaHora, IUnitOfWork uow, ref bool SeHanActualizadoProductosDeOfertasExistentes) 
        {
            //Obtenermos Agrupaciones de SAP
            //CommunicatorLight conectorSAP = new CommunicatorLight(usuario, password);
            //return this.GuardarAgrupacionesTipología(conectorSAP.ZCObtenerAgrupacionTipologiaRfc(usuario, FechaHora), uow, ref SeHanActualizadoProductosDeOfertasExistentes);

            if (SSOHelper.Instance.LogarConSSO)
            {
                CommunicatorLight conectorSAP = new CommunicatorLight(SSOHelper.Instance.Usuario, password);
                //SSOHelper.Instance.ActualizarCookiePortal();
                SSOHelper.Instance.InicializarWSLight(conectorSAP.Cliente);
               
                var resultado = this.GuardarAgrupacionesTipología(conectorSAP.ZCObtenerAgrupacionTipologiaRfc(usuario, FechaHora), uow, ref SeHanActualizadoProductosDeOfertasExistentes);
                SSOHelper.Instance.LimpiarWSLight();
                return resultado;
            }
            else
            {
                CommunicatorLight conectorSAP = new CommunicatorLight(usuario, password);
                return this.GuardarAgrupacionesTipología(conectorSAP.ZCObtenerAgrupacionTipologiaRfc(usuario, FechaHora), uow, ref SeHanActualizadoProductosDeOfertasExistentes);
            }
        }

        

        /// <summary>
        /// Método que guarda en la base de datos local las agrupaciones de tipologías.
        /// </summary>
        /// <param name="listaAgrupacionTipologiaCliente">Colección de agrupaciones a guardar</param>
        /// <param name="uow">Objeto transaccional</param>
        private bool GuardarAgrupacionesTipología(Collection<AgrupacionTipologiaBE> listaAgrupacionTipologiaCliente, IUnitOfWork uow, ref bool SeHanActualizadoProductosDeOfertasExistentes)
        {
            //Obtenemos la lista de agrupaciones de la base de datos
            TipologiaClientePersistence AgrupacionTipologiaClientePersistance = new TipologiaClientePersistence(uow);
            Collection<AgrupacionTipologiaBE> listaAgrupacionTipologiaClienteDB = AgrupacionTipologiaClientePersistance.ObtenerAgrupacionesTipologia();

            //Comprobamos si el número de registros es distinto en ambos entornos y además (si esto fuese cierto) comprobamos los registros...
            //...y en caso de que no sean coincidentes procedemos a actualizar la base de datos local y a establecer la necesidad de recálculo en los productos de las ofertas correspondientes
            if (!((listaAgrupacionTipologiaCliente.Count == listaAgrupacionTipologiaClienteDB.Count) && (listaAgrupacionTipologiaCliente.Except(listaAgrupacionTipologiaClienteDB).Count() == 0)))
            {
                //Obtenemos el listado de productos 
                List<string> ListaProductosSAP = listaAgrupacionTipologiaCliente.Select(x => x.Producto).ToList();
                List<string> ListaProductosDB = listaAgrupacionTipologiaClienteDB.Select(x => x.Producto).ToList();
                List<string> ListaCompletaProductos = new List<string>();
                ListaCompletaProductos.AddRange(ListaProductosSAP);
                ListaCompletaProductos.AddRange(ListaProductosDB);
                ListaCompletaProductos = ListaCompletaProductos.Distinct().ToList();

                //Actualizamos los productos de las ofertas que se ven afectados y ponemos este en pendiente de recálculo dentro de la oferta.
                ProductoOfertaBL productoOferta = new ProductoOfertaBL();

                foreach (string codProductoSap in ListaCompletaProductos)
                {
                    bool seactualizanproductosDeOfertas =  productoOferta.ActualizarProductoOfertaPorCodigoProductoSAP(codProductoSap);
                    if (seactualizanproductosDeOfertas) { SeHanActualizadoProductosDeOfertasExistentes = true; } 
                }

                //Borramos de la base da datos todas las agrupaciones de tipología 
                AgrupacionTipologiaClientePersistance.EliminarRegistrosAgrupacionTipologiaCliente();

                //Guardamos las nuevas agrupaciones de tipología en la DB
                foreach (AgrupacionTipologiaBE AgrupacionTipologiaCliente in listaAgrupacionTipologiaCliente)
                {
                    AgrupacionTipologiaBE AgrupacionTipologiaClienteDB = AgrupacionTipologiaClientePersistance.ObtenerAgrupacionesTipologia(AgrupacionTipologiaCliente.Anexo, AgrupacionTipologiaCliente.VA, AgrupacionTipologiaCliente.Producto, AgrupacionTipologiaCliente.idAgrupacion).FirstOrDefault();

                    if (AgrupacionTipologiaClienteDB == null)
                    {
                        AgrupacionTipologiaClienteDB = new AgrupacionTipologiaBE()
                        {
                            idAgrupacionProducto = Guid.NewGuid(),
                            Anexo = AgrupacionTipologiaCliente.Anexo,
                            idAgrupacion = AgrupacionTipologiaCliente.idAgrupacion,
                            Producto = AgrupacionTipologiaCliente.Producto,
                            VA = AgrupacionTipologiaCliente.VA
                        };

                        AgrupacionTipologiaClientePersistance.InsertarAgrupacionTipologia(AgrupacionTipologiaClienteDB);
                    }
                }

                return true;
            }
            else //No ha sido necesario actualizar las agrupaciones de tipología
            {
                return false;
            } 
        }

        #endregion
    }
}
