using Correos.SimuladorOfertas.Common;
using Correos.SimuladorOfertas.Common.Extensions;
using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.InOutLight;
using Correos.SimuladorOfertas.Persistence;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace Correos.SimuladorOfertas.Business
{
    public class OfertaBL
    {
        #region Propiedades

        public bool IsAltaNueva { get; set; }

        #endregion

        #region Métodos Publicos

        #region Obtener desde SAP

        /// <summary>
        /// Obtiene de SAP toda la información relativa a la oferta que ya existía previamente en SAP CRM
        /// </summary>
        /// <param name="objOferta"></param>
        /// <param name="usuario"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public ResultBE ReDescargarOfertaSAP(OfertaBE objOferta, string usuario, string password)
        {
            ResultBE objRespuesta = new ResultBE();

            try
            {
                #region variables usadas

                //Contiene los identificadores de las ofertas que quieren modificar
                Guid idOfertaAnterior = objOferta.idOferta;
                Guid idOfertaNueva = Guid.NewGuid();

                //lista de variables adaptadas a nuestro modelo de negocio
                Collection<ConfiguracionDestinoOfertaBE> listaDestinos = new Collection<ConfiguracionDestinoOfertaBE>();
                Collection<ConfiguracionTramoOfertaBE> listaTramos = new Collection<ConfiguracionTramoOfertaBE>();
                Collection<ConfiguracionPosicionSAPBE> listaPosiciones = new Collection<ConfiguracionPosicionSAPBE>();
                Collection<GrupoTramoBE> listaGruposTramo = new Collection<GrupoTramoBE>();
                Collection<ConfiguracionCaracteristicaBE> listaCaracteristicas = new Collection<ConfiguracionCaracteristicaBE>();
                Collection<ConfiguracionValorAnadidoBE> listaConfiguracionesValorAnadido = new Collection<ConfiguracionValorAnadidoBE>();

                Collection<ProductoOfertaBE> listaProductosOferta = new Collection<ProductoOfertaBE>();
                OfertaBE ofertaDescargada = new OfertaBE();

                #endregion

                #region Llamada al servicio

                ResultBE objResultado = null;
                if (SSOHelper.Instance.LogarConSSO)
                {
                    CommunicatorLight conectorSAP = new CommunicatorLight(SSOHelper.Instance.Usuario, password);
                    //SSOHelper.Instance.ActualizarCookiePortal();
                    SSOHelper.Instance.InicializarWSLight(conectorSAP.Cliente);
                    objResultado = conectorSAP.ZCPosicionesASimulaRfc(objOferta.CodOfertaSAP, out listaDestinos, out listaPosiciones,
                        out listaTramos, out listaGruposTramo, out ofertaDescargada, out listaCaracteristicas, out listaConfiguracionesValorAnadido);
                    SSOHelper.Instance.LimpiarWSLight();
                }
                else
                {
                CommunicatorLight conectorSAP = new CommunicatorLight(usuario, password);
                    objResultado = conectorSAP.ZCPosicionesASimulaRfc(objOferta.CodOfertaSAP, out listaDestinos, out listaPosiciones,
                    out listaTramos, out listaGruposTramo, out ofertaDescargada, out listaCaracteristicas, out listaConfiguracionesValorAnadido);
                }

                #endregion

                if (objResultado.Resultado)
                {
                    #region crear el esqueleto de la oferta.

                    using (IUnitOfWork uow = new UnitOfWork())
                    {
                        uow.DeshabilitarAutodeteccionDeCambios();

                        //clonar oferta
                        OfertaPersistence ofertaPersistence = new OfertaPersistence(uow);

                        //Antes de clonar la oferta, se actualiza los campos que deben modificarse.
                        objOferta.StatusSAP = ofertaDescargada.StatusSAP;
                        objOferta.DescripcionStatusSAP = ofertaDescargada.DescripcionStatusSAP;
                        objOferta.Estado = SimuladorResources.StatusSincronizado;
                        objOferta.FactorCubicaje = ofertaDescargada.FactorCubicaje;
                        objOferta.Descripcion = ofertaDescargada.Descripcion;
                        objOferta.ValidezDefinitiva = true;

                        ofertaPersistence.ClonarOferta(objOferta, idOfertaNueva);
                        objOferta.idOferta = idOfertaNueva;

                        ProductoPersistence objProductoPersistence = new ProductoPersistence(uow);
                        Collection<ReDescargaBE> listaProductosSAP = objProductoPersistence.ObtenerListaProductosAnexosModalidad();

                        foreach (ConfiguracionPosicionSAPBE item in listaPosiciones.Where(x => string.IsNullOrEmpty(x.CodProductoVASAP)))
                        {
                            ReDescargaBE objDescarga = listaProductosSAP.FirstOrDefault(x => x.CodAnexoSAP.Equals(item.CodAnexoSAP)
                                && x.CodProductoSAP.Equals(item.CodProductoSAP) && x.CodModalidadNegociacion.Equals(item.CodModalidadNegociacion));

                            List<DestinoBE> listadoDestinos = new DestinoBL().ObtenerDestinosProductoConVisibilidad(objDescarga.IdProducto);

                            if (objDescarga != null)
                            {
                                ProductoOfertaBE objProducto = new ProductoOfertaBE();
                                objProducto.idOferta = idOfertaNueva;
                                objProducto.idProductoOferta = Guid.NewGuid();
                                objProducto.idProducto = objDescarga.IdProducto;
                                objProducto.NumeroEnvios = item.NumEnvios;
                                objProducto.Posicion = item.Posicion;
                                objProducto.idModalidadNegociacion = objDescarga.IdModalidadNegociacion;
                                objProducto.CodModalidadNegociacion = objDescarga.CodModalidadNegociacion;
                                objProducto.Anexo = objDescarga.CodAnexoSAP;
                                objProducto.CodProductoSAP = objDescarga.CodModalidadNegociacion;
                                objProducto.StatusProducto = item.StatusPosicion;

                                objProducto.TarifaZA = item.TarifaZA;
                                objProducto.TarifaZB = item.TarifaZB;
                                objProducto.TarifaZC = item.TarifaZC;
                                objProducto.TarifaZD = item.TarifaZD;
                                objProducto.TarifaZE = item.TarifaZE;
                                objProducto.TarifaZX = item.TarifaZX;

                                //Impresión de ofertas
                                objProducto.Paises = item.Paises;
                                objProducto.PuntoAdmision = item.PuntoAdmision;
                                objProducto.Formato = item.Formato;
                                objProducto.ExpedicionesMultibulto = item.ExpedicionesMultiBulto;
                                objProducto.PorcentajeExpedicionesMonoBulto = item.PorcentajeExpedicionesMonoBulto;
                                objProducto.BultosPorExpedicion = item.BultosPorExpedicion;

                                objProducto.listaDestinosVisibles = listadoDestinos.Where(x => listaDestinos.Any(y => y.CodZona == x.CodDestinoSAP && y.CodProductoSAP.Equals(objDescarga.CodProductoSAP))).ToList(); 
                                // TODO Descomentar para evolutivo de peso volumétrico
                                objProducto.IndemnizacionPactada = item.IndemnizacionPactada ?? 0;
                                objProducto.IncrementoMinimo = item.IndemPactadaIncrementoMinimo ?? 0;
                                objProducto.EsReneg = item.EsReneg;
                                listaProductosOferta.Add(objProducto);
                                
                            }
                        }

                        ProductoOfertaPersistence productoOfertaPersistence = new ProductoOfertaPersistence(uow);
                        ConfiguracionDestinoOfertaPersistence confDestinoPersistence = new ConfiguracionDestinoOfertaPersistence(uow);
                        ConfiguracionValorAnadidoPersistence confVAPersistence = new ConfiguracionValorAnadidoPersistence(uow);
                        ConfiguracionListaPreciosPersistence confListaPreciosPersistence = new ConfiguracionListaPreciosPersistence(uow);
                        ConfiguracionPuntoOfertaPersistence confPuntoOfertaPersistence = new ConfiguracionPuntoOfertaPersistence(uow);
                        ConfiguracionGradoOfertaPersistence confGradoOfertaPersistence = new ConfiguracionGradoOfertaPersistence(uow);
                        CaracteristicaPersistence confCaracteristicaPersistence = new CaracteristicaPersistence(uow);

                        foreach (ProductoOfertaBE item in listaProductosOferta)
                        {
                            productoOfertaPersistence.GuardarProductoOferta(item);
                            confDestinoPersistence.InsertConfiguracionDestinoOfertaProductoOferta(item);
                            confVAPersistence.InsertConfiguracionValorAnadidoProductoOferta(item);
                            confListaPreciosPersistence.InsertConfiguracionListaPreciosProductoOferta(item.idProductoOferta);
                            confPuntoOfertaPersistence.InsertConfiguracionPuntosOferta(item);
                            confGradoOfertaPersistence.InsertConfiguracionGradosOferta(item);
                            confCaracteristicaPersistence.InsertConfiguracionCaracteristicaOfertaProductoOferta(item);
                        }

                        //Se guarda el contexto
                        uow.Save();

                        //Se guardan los datos de configuración
                        ProductoOfertaBL productoOfertaBL = new ProductoOfertaBL();
                        productoOfertaBL.GuardarDatosConfiguracion(listaProductosOferta, uow);

                        //Se vacía la lista almacenada en memoria
                        InformacionEstatica.ListaProductosOfertaBE = new Collection<ProductoOfertaBE>();

                        //Se guarda la configuración para Posiciones
                        PosicionPersistence posicionPersistence = new PosicionPersistence(uow);
                        posicionPersistence.GuardarConfiguracionPosicion(listaPosiciones, idOfertaNueva);

                        //Se guarda la configuración para ValorAñadido
                        ValorAnadidoProductoPersistence vaPersistencia = new ValorAnadidoProductoPersistence(uow);
                        vaPersistencia.GuardarConfiguracionValorAnadidoProducto(listaProductosOferta, listaPosiciones, idOfertaNueva, listaConfiguracionesValorAnadido);

                        //Se guarda la configuración para Destino
                        DestinoPersistence destinoPersistence = new DestinoPersistence(uow);
                        destinoPersistence.GuardarConfiguracionDestinoOferta(listaDestinos, idOfertaNueva, esRedescargaSAP: true);

                        //Se guarda la configuración para Tramo
                        TramoPersistence tramoPersistence = new TramoPersistence(uow);
                        tramoPersistence.GuardarConfiguracionTramoOferta(listaTramos, idOfertaNueva);

                        //Se guarda la configuración para ListaPrecios
                        ListaPreciosPersistence listaPreciosPersistence = new ListaPreciosPersistence(uow);
                        listaPreciosPersistence.GuardarConfiguracionListaPrecios(listaProductosOferta);

                        //Se guarda la lista de grupos de tramo
                        ConfiguracionGruposTramoOfertaPersistence gtPersistence = new ConfiguracionGruposTramoOfertaPersistence(uow);
                        gtPersistence.GuardarConfiguracionGruposTramoOferta(listaGruposTramo, idOfertaNueva);

                        foreach (var item in listaCaracteristicas)
                        {
                            confCaracteristicaPersistence.GuardarConfiguracionDescargarCaracteristicaOferta(item, idOfertaNueva);
                        }

                        //Se guarda el contexto
                        uow.Save();

                    }

                    #endregion

                    #region elimina la oferta anterior

                    OfertaBL objOfertaBL = new OfertaBL();
                    objOfertaBL.EliminarOferta(idOfertaAnterior);

                    #endregion
                }
            }
            catch (Exception ex)
            {
                RegistrarAccionesSimulador.GuardarExcepcion(ex, false);
                objRespuesta.Resultado = false;
                objRespuesta.TextoError = SimuladorResources.ErrorDescargarOfertaSAP;
            }

            return objRespuesta;
        }

        /// <summary>
        /// Obtiene el listado de ofertas entre fechas de SAP.
        /// </summary>
        /// <param name="usuario"></param>
        /// <param name="password"></param>
        /// <param name="fechaDesde"></param>
        /// <param name="fechaHasta"></param>
        /// <returns></returns>
        public Collection<OfertaChronoFichaBE> ObtenerListadoOfertasFechasSAP(string usuario, string password, DateTime fechaDesde, DateTime fechaHasta)
        {
            if (SSOHelper.Instance.LogarConSSO)
            {
                CommunicatorLight conectorSAP = new CommunicatorLight(SSOHelper.Instance.Usuario, password);
                //SSOHelper.Instance.ActualizarCookiePortal();
                SSOHelper.Instance.InicializarWSLight(conectorSAP.Cliente);
                Collection<OfertaChronoFichaBE> resultado = conectorSAP.ZCFichasCexRfc(fechaDesde, fechaHasta, usuario);
                SSOHelper.Instance.LimpiarWSLight();
                return resultado;
            }
            else
            {
            CommunicatorLight conectorSAP = new CommunicatorLight(usuario, password);
            return conectorSAP.ZCFichasCexRfc(fechaDesde, fechaHasta, usuario);
        }
        }

        /// <summary>
        /// Devuelve el listado de ofertas asociadas a un cliente
        /// </summary>
        /// <param name="usuario">Código usuario</param>
        /// <param name="password">Contraseña del usuario</param>
        /// <param name="fechaDesde"></param>
        /// <param name="fechaHasta"></param>
        /// <param name="cliente">Nombre del cliente a buscar</param>
        /// <param name="clienteBE">Entidad ClienteBE</param>
        public Collection<OfertaBE> CargarListadoOfertasGestor(string usuario, string password, string fechaDesde, string fechaHasta, string cliente, ClienteBE clienteBE, string numOferta)
        {
            if (SSOHelper.Instance.LogarConSSO)
            {
                CommunicatorLight conectorSAP = new CommunicatorLight(SSOHelper.Instance.Usuario, password);
                //SSOHelper.Instance.ActualizarCookiePortal();
                SSOHelper.Instance.InicializarWSLight(conectorSAP.Cliente);
                Collection<OfertaBE> resultado = conectorSAP.ZCOfertasGestorRfc(fechaDesde, fechaHasta, cliente, ref clienteBE, usuario, numOferta);
                SSOHelper.Instance.LimpiarWSLight();
                return resultado;
            }
            else
            {
                CommunicatorLight conectorSAP = new CommunicatorLight(usuario, password);
                return conectorSAP.ZCOfertasGestorRfc(fechaDesde, fechaHasta, cliente, ref clienteBE, usuario, numOferta);
            }
        }

        #endregion

        #region Obtener de BBDD

        /// <summary>
        /// Obtiene las Ofertas cuyo productos dentro de la oferta tienen un EstadoCalculo = 3
        /// </summary>
        /// <returns></returns>
        public Collection<OfertaBE> ObtenerOfertasAfectadasPorAgrupacionesTipologia()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                OfertaPersistence ofertaPersistence = new OfertaPersistence(uow);
                return ofertaPersistence.ObtenerOfertasAfectadasPorAgrupacionesTipologia();
            }
        }

        /// <summary>
        /// Devuelve cierto si hay listado de id iguales al codigo que le paso.
        /// </summary>
        /// <param name="codOfertaSAP"></param>
        /// <param name="identificadores"></param>
        /// <returns></returns>
        public bool ExisteOfertaConCodOfertaSAP(string codOfertaSAP, out Collection<Guid> identificadores)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                OfertaPersistence ofertaPersistence = new OfertaPersistence(uow);
                return ofertaPersistence.ExisteOfertaConCodOfertaSAP(codOfertaSAP, out identificadores);
            }
        }

        /// <summary>
        ///  Metodo para obtener una oferta a partir de su CodOfertaSAP
        /// </summary>
        /// <param name="codOfertaSAP">Código SAP de la oferta</param>
        /// <returns>OfertaBE cuyo ID corresponde con el parámetro de entrada</returns>
        public OfertaBE ObtenerOfertaPorCodOfertaSAP(string codOfertaSAP)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                OfertaPersistence ofertaPersistence = new OfertaPersistence(uow);
                return ofertaPersistence.ObtenerOfertaPorCodOfertaSAP(codOfertaSAP);
            }
        }

        /// <summary>
        /// Obtiene la entidad oferta a partir de su identificador de codigo de SAP
        /// </summary>
        /// <param name="codOfertaSAP"></param>
        /// <returns></returns>
        public OfertaBE ObtenerOfertaByIdOferta(Guid idOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                OfertaPersistence ofertaPersistence = new OfertaPersistence(uow);
                return ofertaPersistence.ObtenerOfertaPorIdOferta(idOferta);
            }
        }

        /// <summary>
        /// Método que obtiene todas las oferta almacenadas en base de datos
        /// </summary>
        /// <returns>Lista de OfertaBE</returns>
        public Collection<OfertaBE> ObtenerOfertas()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                OfertaPersistence ofertaPersistence = new OfertaPersistence(uow);
                return ofertaPersistence.ObtenerOfertas();
            }
        }

        /// <summary>
        /// Se obtienen todas las ofertas que tienen CodOfertaSAP
        /// </summary>
        /// <returns></returns>
        public Collection<string> ObtenerOfertasSincronizadas()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                OfertaPersistence ofertaPersistence = new OfertaPersistence(uow);
                return ofertaPersistence.ObtenerOfertasSincronizadas();
            }
        }

        /// <summary>
        /// Método que obtiene el número de ofertas de BD
        /// </summary>
        /// <returns></returns>
        public int obtenerNumeroOfertasBD()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                OfertaPersistence ofertaPersistence = new OfertaPersistence(uow);
                return ofertaPersistence.obtenerNumeroOfertasBD();
            }
        }

        /// <summary>
        /// Actualiza el estado de sincronización de las ofertas
        /// </summary>
        /// <param name="ofertas"></param>
        /// <returns></returns>
        public Collection<OfertaBE> SetEstadosSincronizacion(Collection<OfertaBE> ofertas)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                EstadoSincronizacionPersistence estadoPersistence = new EstadoSincronizacionPersistence(uow);

                //Obtenemos todos los estados de sincronización para no tener que consultar varias veces la BBDD
                var listadoEstadosSincro = estadoPersistence.ObtenerEstadosSincronizacion();

                foreach (var item in ofertas)
                {
                    //Si la oferta tiene estado de sincronización
                    if (listadoEstadosSincro.Any(t => t.idOferta.Equals(item.idOferta)))
                    {
                        var estado = listadoEstadosSincro.First(t=> t.idOferta.Equals(item.idOferta));

                        //Si la oferta ya se ha sincronizado, borramos el estado.
                        if (estado.EstadoOferta.Equals("Finalizada"))
                        {
                            item.EstadoSincronizacion = String.Empty;
                            estadoPersistence.EliminarEstadoSincronizacion(item.idOferta);
                        }
                        //Si no, actualizamos el estado.
                        else
                        {
                            item.EstadoSincronizacion = estado.EstadoOferta;
                        }
                    }
                    //Si no tiene estado de sincronización, lo borramos de memoria
                    else
                    {
                        item.EstadoSincronizacion = String.Empty;
                    }
                }                                
            }

            return ofertas;
        }

        #endregion

        #region Guardar

        /// <summary>
        /// Método que guarda en base de datos los datos relacionados de la oferta, sus productos y los datos de configuración
        /// </summary>
        /// <param name="oferta">Entidad OfertaBE</param>
        /// <param name="listaProductosOferta">Lista de ProductoOfertaBE</param>
        /// <param name="listaConfiguracionListaPreciosBE">Lista de ConfiguracionListaPreciosBE</param>
        /// <param name="listaConfiguracionDestinoOfertaBE">Lista de ConfiguracionDestinoOfertaBE</param>
        /// <param name="listaConfiguracionTramoOfertaBE">Lista de ConfiguracionTramoOfertaBE</param>
        /// <param name="listaConfiguracionValorAnadidoBE">Lista de ConfiguracionValorAnadidoBE</param>
        /// <param name="listaConfiguracionPuntoOfertaBE">Lista de ConfiguracionPuntoOfertaBE</param>
        /// <param name="listaConfiguracionGradoOfertaBE">Lista de ConfiguracionGradoOfertaBE</param>
        /// <param name="listaGruposTramoBE">Lista de GrupoTramoBE</param>
        public void GuardarOferta(OfertaBE oferta,
            Collection<ProductoOfertaBE> listaProductosOferta,
            Collection<ConfiguracionListaPreciosBE> listaConfiguracionListaPreciosBE,
            Collection<ConfiguracionDestinoOfertaBE> listaConfiguracionDestinoOfertaBE,
            Collection<ConfiguracionTramoOfertaBE> listaConfiguracionTramoOfertaBE,
            Collection<ConfiguracionValorAnadidoBE> listaConfiguracionValorAnadidoBE,
            Collection<ConfiguracionPuntoOfertaBE> listaConfiguracionPuntoOfertaBE,
            Collection<ConfiguracionGradoOfertaBE> listaConfiguracionGradoOfertaBE,
            Collection<GrupoTramoBE> listaGruposTramoBE,
            Collection<ConfiguracionCaracteristicaBE> listaCaracteristicaBE)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                uow.DeshabilitarAutodeteccionDeCambios();

                //Se guarda la oferta
                OfertaPersistence ofertaPersistence = new OfertaPersistence(uow);
                ofertaPersistence.GuardarOferta(oferta);

                //Se guardan los productos oferta
                ProductoOfertaPersistence productoOfertaPersistence = new ProductoOfertaPersistence(uow);
                foreach (ProductoOfertaBE item in listaProductosOferta)
                {
                    productoOfertaPersistence.GuardarProductoOferta(item);
                }

#if DEBUG
                uow.Save();
#endif
                //Se guardan las configuraciones de la lista de precios
                ConfiguracionListaPreciosPersistence listaPreciosPersistence = new ConfiguracionListaPreciosPersistence(uow);
                listaPreciosPersistence.GuardarConfiguracionListasPrecios(listaProductosOferta, listaConfiguracionListaPreciosBE);
#if DEBUG
                uow.Save();
#endif

                //Se guardan las configuraciones de la lista de destinos
                ConfiguracionDestinoOfertaPersistence destinoPersistence = new ConfiguracionDestinoOfertaPersistence(uow);
                destinoPersistence.GuardarConfiguracionDestinosOferta(listaProductosOferta, listaConfiguracionDestinoOfertaBE);
#if DEBUG
                uow.Save();
#endif

                //Se guardan las configuraciones de la lista de tramos
                ConfiguracionTramoOfertaPersistence tramoPersistence = new ConfiguracionTramoOfertaPersistence(uow);
                tramoPersistence.GuardarConfiguracionTramosOferta(listaProductosOferta, listaConfiguracionTramoOfertaBE);
#if DEBUG
                uow.Save();
#endif

                //Guarda la configuración de las características del producto.
                CaracteristicaPersistence caracteristicaPersistence = new CaracteristicaPersistence(uow);
                foreach (ConfiguracionCaracteristicaBE item in listaCaracteristicaBE)
                {
                    caracteristicaPersistence.GuardarConfiguracionCaracteristicasProducto(item);
                }
#if DEBUG
                uow.Save();
#endif

                //Se guardan las configuraciones de los valores añadidos
                ConfiguracionValorAnadidoPersistence vaPersistence = new ConfiguracionValorAnadidoPersistence(uow);


                //Borramos las configuraciones que ya no existen en el producto
                foreach(var productoOferta in listaProductosOferta)
                {
                    var listaConfigVABEProdOferta = vaPersistence.ObtenerListaConfiguracionVA(productoOferta.idProductoOferta);

                    foreach (var item in listaConfigVABEProdOferta)
                    {
                        vaPersistence.BorrarConfiguracionesVA(item);
                    }
                }
#if DEBUG
                uow.Save();
#endif

                foreach (ConfiguracionValorAnadidoBE item in listaConfiguracionValorAnadidoBE)
                {
                    vaPersistence.BorrarConfiguracionesVA(item);
                    vaPersistence.GuardarConfiguracionValorAnadido(item);
                    //if(item.NumeroEnvios > 0)
                    //    vaPersistence.GuardarConfiguracionValorAnadido(item);
                }
#if DEBUG
                uow.Save();
#endif

                //Se guardan las configuraciones de los puntos
                ConfiguracionPuntoOfertaPersistence puntoPersistence = new ConfiguracionPuntoOfertaPersistence(uow);
                foreach (ConfiguracionPuntoOfertaBE item in listaConfiguracionPuntoOfertaBE)
                {
                    puntoPersistence.GuardarConfiguracionPuntosOferta(item);
                }
#if DEBUG
                uow.Save();
#endif

                //Se guardan las configuraciones de los grados
                ConfiguracionGradoOfertaPersistence gradoPersistence = new ConfiguracionGradoOfertaPersistence(uow);
                foreach (ConfiguracionGradoOfertaBE item in listaConfiguracionGradoOfertaBE)
                {
                    gradoPersistence.GuardarConfiguracionGradosOferta(item);
                }
#if DEBUG
                uow.Save();
#endif

                //Grupos de Tramo
                ConfiguracionGruposTramoOfertaPersistence gruposTramoPersistence = new ConfiguracionGruposTramoOfertaPersistence(uow);

                //Primero se eliminan los Grupos de Tramo de los productos ya existentes
                gruposTramoPersistence.EliminarListaConfiguracionGruposTramoOferta(
                    listaProductosOferta.Where(x => x.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoGrupoTramo) ||
                        x.CodModalidadNegociacion.Equals(SimuladorResources.ValorPrecioCiertoGrupoTramo)).ToList<ProductoOfertaBE>().ToCollection<ProductoOfertaBE>());

                //Se insertan los nuevos Grupos de Tramo
                gruposTramoPersistence.InsertarConfiguracionGruposTramoOferta(listaGruposTramoBE);



                uow.Save();

                //-----------------------------------------------------------------
                // JCNS. VISIBILIDAD
                // ELIMINAR DESTINOS INVISIBLES CON DISTRIBUCION > 0
                // ELIMINAR TAMBIÉN TRAMOS Y GRUPO DE TRAMOS DE LOS DESTINOS
                //-----------------------------------------------------------------
                tramoPersistence.ActualizarConfiguracionTramos_DestinosInvisiblesConDistribucion(oferta.idOferta);
                gruposTramoPersistence.ActualizarConfiguracionGrupoTramos_DestinosInvisiblesConDistribucion(oferta.idOferta);
                destinoPersistence.ActualizarConfiguracionDestinos_DestinosInvisiblesConDistribucion(oferta.idOferta);

                //Se guarda el contexto
                //uow.Save();
            }
        }

        /// <summary>
        /// Método que guarda una oferta almacenada en memoria a una BD vacía
        /// </summary>
        /// <param name="oferta">Entidad OfertaBE</param>
        /// <param name="listaProductosOferta">Lista de ProductoOfertaBE</param>
        /// <param name="listaConfiguracionListaPreciosBE">Lista de ConfiguracionListaPreciosBE</param>
        /// <param name="listaConfiguracionDestinoOfertaBE">Lista de ConfiguracionDestinoOfertaBE</param>
        /// <param name="listaConfiguracionTramoOfertaBE">Lista de ConfiguracionTramoOfertaBE</param>
        /// <param name="listaConfiguracionValorAnadidoBE">Lista de ConfiguracionValorAnadidoBE</param>
        /// <param name="listaConfiguracionPuntoOfertaBE">Lista de ConfiguracionPuntoOfertaBE</param>
        /// <param name="listaConfiguracionGradoOfertaBE">Lista de ConfiguracionGradoOfertaBE</param>
        /// <param name="listaGruposTramoBE">Lista de GrupoTramoBE</param>
        /// /// <param name="listaGruposTramoBE">Lista de CaracteristicasBE</param>
        public void GuardarOfertaDeMemoria(OfertaBE oferta,
            Collection<ProductoOfertaBE> listaProductosOferta,
            Collection<ConfiguracionListaPreciosBE> listaConfiguracionListaPreciosBE,
            Collection<ConfiguracionDestinoOfertaBE> listaConfiguracionDestinoOfertaBE,
            Collection<ConfiguracionTramoOfertaBE> listaConfiguracionTramoOfertaBE,
            Collection<ConfiguracionValorAnadidoBE> listaConfiguracionValorAnadidoBE,
            Collection<ConfiguracionPuntoOfertaBE> listaConfiguracionPuntoOfertaBE,
            Collection<ConfiguracionGradoOfertaBE> listaConfiguracionGradoOfertaBE,
            Collection<GrupoTramoBE> listaGruposTramoBE,
            Collection<ConfiguracionCaracteristicaBE> listaCaracteristicaBE,
            Collection<ClienteBE> listaClientes,
            Collection<ConfiguracionValorAnadidoTarifaBE> listaConfigValorAnadidoTarifa,
            Collection<ConfiguracionValorAnadidoCaracteristicaBE> listaConfigValorAnadidoCaracteristica)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                // Deshabilitamos la autodetección de cambios para este uow porque la cantidad de Adds es muy elevada
                uow.DeshabilitarAutodeteccionDeCambios();

                //Se obtiene un listado de identificadores de los productosOferta de la oferta
                var productosOfertaIds = listaProductosOferta.Select(p => p.idProductoOferta).Distinct().ToList();

                #region Clientes
                //Se guarda el cliente de la oferta
                ClienteBL clienteBL = new ClienteBL();
                clienteBL.GuardarClienteDeMemoria(listaClientes.Where(x => x.idCliente.Equals(oferta.idCliente)).FirstOrDefault(), uow);

                #endregion

                #region Oferta
                //Se guarda la oferta
                OfertaPersistence ofertaPersistence = new OfertaPersistence(uow);
                ofertaPersistence.GuardarOfertaDeMemoria(oferta);
                #endregion

                #region Productos oferta
                //Se guardan los productos oferta
                ProductoOfertaPersistence productoOfertaPersistence = new ProductoOfertaPersistence(uow);
                foreach (ProductoOfertaBE item in listaProductosOferta)
                {
                    productoOfertaPersistence.GuardarProductoOfertaDeMemoria(item);
                }
                #endregion

                #region Configuraciones de lista de precios
                //Se guardan las configuraciones de la lista de precios
                ConfiguracionListaPreciosPersistence listaPreciosPersistence = new ConfiguracionListaPreciosPersistence(uow);
                listaPreciosPersistence.GuardarConfiguracionListasPreciosDeMemoria(listaProductosOferta, listaConfiguracionListaPreciosBE);
                #endregion

                #region Configuraciones de destinos
                //Se guardan las configuraciones de la lista de destinos
                ConfiguracionDestinoOfertaPersistence destinoPersistence = new ConfiguracionDestinoOfertaPersistence(uow);
                destinoPersistence.GuardarConfiguracionDestinosOfertaDeMemoria(listaProductosOferta, listaConfiguracionDestinoOfertaBE);
                #endregion

                #region Configuraciones lista de tramos

                //Se debe guardar el contexto ya que se realiza una insercción mediante una consulta a BD de los registros de configuración de tramos
                uow.Save();

                //Primero se guarda la estructura de los productos con valores inicializados a 0
                ProductoOfertaBL productoOfertaBL = new ProductoOfertaBL();
                productoOfertaBL.GuardarDatosConfiguracion(listaProductosOferta, uow);

                //Se guardan las configuraciones de la lista de tramos
                ConfiguracionTramoOfertaPersistence tramoPersistence = new ConfiguracionTramoOfertaPersistence(uow);
                Collection<ConfiguracionTramoOfertaBE> auxConfigTramoOfertaBE = listaConfiguracionTramoOfertaBE.Where(x => productosOfertaIds.Contains(x.idProductoOferta.Value)).ToList<ConfiguracionTramoOfertaBE>().ToCollection<ConfiguracionTramoOfertaBE>();
                tramoPersistence.GuardarConfiguracionTramosOferta(listaProductosOferta, auxConfigTramoOfertaBE);

                #endregion

                #region Configuraciones de valores añadidos
                //Se guardan las configuraciones de los valores añadidos
                ConfiguracionValorAnadidoPersistence vaPersistence = new ConfiguracionValorAnadidoPersistence(uow);
                Collection<ConfiguracionValorAnadidoBE> auxListaConfiguracionValorAnadidoBE = new Collection<ConfiguracionValorAnadidoBE>(listaConfiguracionValorAnadidoBE.Where(x => productosOfertaIds.Contains(x.idProductoOferta.Value)).ToList<ConfiguracionValorAnadidoBE>());
                foreach (ConfiguracionValorAnadidoBE item in auxListaConfiguracionValorAnadidoBE)
                {
                    vaPersistence.GuardarConfiguracionValorAnadidoDeMemoria(item);
                }
                #endregion

                #region ConfiguracionesPuntosOferta
                ConfiguracionPuntoOfertaPersistence puntoPersistence = new ConfiguracionPuntoOfertaPersistence(uow);
                //Se filtran los registros que pertenecen a la oferta               
                Collection<ConfiguracionPuntoOfertaBE> auxListaConfiguracionPuntosOferta = new Collection<ConfiguracionPuntoOfertaBE>(listaConfiguracionPuntoOfertaBE.Where(x => productosOfertaIds.Contains(x.idProductoOferta)).ToList<ConfiguracionPuntoOfertaBE>());

                //Se guardan las configuraciones de los puntos
                foreach (ConfiguracionPuntoOfertaBE item in auxListaConfiguracionPuntosOferta)
                {
                    puntoPersistence.GuardarConfiguracionPuntosOfertaDeMemoria(item);
                }
                #endregion

                #region Configuraciones de grados

                ConfiguracionGradoOfertaPersistence gradoPersistence = new ConfiguracionGradoOfertaPersistence(uow);
                //Se filtran los registros pertenecientes a la oferta
                Collection<ConfiguracionGradoOfertaBE> auxListaConfiguracionGradoOfertaBE = new Collection<ConfiguracionGradoOfertaBE>(listaConfiguracionGradoOfertaBE.Where(x => productosOfertaIds.Contains(x.idProductoOferta)).ToList<ConfiguracionGradoOfertaBE>());
                foreach (ConfiguracionGradoOfertaBE item in auxListaConfiguracionGradoOfertaBE)
                {
                    gradoPersistence.GuardarConfiguracionGradosOfertaDeMemoria(item);
                }

                #endregion

                #region Configuraciones de grupos de tramo
                //Grupos de Tramo
                ConfiguracionGruposTramoOfertaPersistence gruposTramoPersistence = new ConfiguracionGruposTramoOfertaPersistence(uow);
                Collection<GrupoTramoBE> auxListaGruposTramoBE = new Collection<GrupoTramoBE>(listaGruposTramoBE.Where(x => productosOfertaIds.Contains(x.idProductoOferta)).ToList<GrupoTramoBE>());
                //Se insertan los nuevos Grupos de Tramo
                gruposTramoPersistence.InsertarConfiguracionGruposTramoOferta(auxListaGruposTramoBE);
                #endregion

                #region Configuraciones de características

                //Guarda la configuración de las características del producto.
                CaracteristicaPersistence caracteristicaPersistence = new CaracteristicaPersistence(uow);
                Collection<ConfiguracionCaracteristicaBE> auxListaCaracteristicasBE = new Collection<ConfiguracionCaracteristicaBE>(listaCaracteristicaBE.Where(x => productosOfertaIds.Contains(x.idProductoOferta)).ToList<ConfiguracionCaracteristicaBE>());
                foreach (ConfiguracionCaracteristicaBE item in auxListaCaracteristicasBE)
                {
                    caracteristicaPersistence.GuardarConfiguracionCaracteristicasDeMemoria(item);
                }
                #endregion

                #region ConfiguracionVATarifa

                Collection<ConfiguracionValorAnadidoTarifaBE> auxListaConfiguracionValorAnadidoTarifaBE = new Collection<ConfiguracionValorAnadidoTarifaBE>(listaConfigValorAnadidoTarifa.Where(x => productosOfertaIds.Contains(x.idProductoOferta)).ToList<ConfiguracionValorAnadidoTarifaBE>());
                foreach (ConfiguracionValorAnadidoTarifaBE item in auxListaConfiguracionValorAnadidoTarifaBE)
                {
                    vaPersistence.GuardarCopiaConfiguracionVañorAnadidoTarifa(item, item.idConfiguracionValorAnadido.Value);
                }

                #endregion

                #region ConfiguracionVACaracteristica

                Collection<ConfiguracionValorAnadidoCaracteristicaBE> auxListaConfiguracionValorAnadidoCaracteristica = new Collection<ConfiguracionValorAnadidoCaracteristicaBE>(listaConfigValorAnadidoCaracteristica.Where(x => productosOfertaIds.Contains(x.idProductoOferta)).ToList<ConfiguracionValorAnadidoCaracteristicaBE>());
                foreach (ConfiguracionValorAnadidoCaracteristicaBE item in auxListaConfiguracionValorAnadidoCaracteristica)
                {
                    vaPersistence.GuardarCopiaConfiguracionVañorAnadidoCaracteristica(item, item.idConfiguracionValorAnadido.Value);
                }

                #endregion

                //Se guarda el contexto
                uow.Save();
            }
        }

        /// <summary>
        /// Método que guarda una oferta y su cliente asociado en base de datos
        /// </summary>
        /// <param name="oferta">Entidad OfertaBE a guardar</param>
        /// <param name="cliente">Entidad ClienteBE a guardar</param>
        /// <param name="listaProductos">Lista de productos que forman la oferta</param>
        /// <param name="usuarioLogin">User</param>
        /// <param name="passwordLogin">Password</param>
        /// <returns>Oferta guardada</returns>
        public OfertaBE GuardarOferta(OfertaBE oferta, ClienteBE cliente, Collection<ProductoOfertaBE> listaProductos, string usuarioLogin, string passwordLogin, bool esCopiaEsqueleto, DescargaOfertaBE configOfertaDescargada)        
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                uow.DeshabilitarAutodeteccionDeCambios();

                this.IsAltaNueva = oferta.idOferta.Equals(Guid.Empty) ? true : false;

                bool isEdicionOfertaSAP = false;

                if (!string.IsNullOrWhiteSpace(oferta.CodOfertaSAP) && oferta.idOferta.Equals(Guid.Empty))
                {
                    isEdicionOfertaSAP = true;

                    //Se actualizan los campos de estado y status
                    //oferta.DescripcionStatusSAP = SimuladorResources.StatusEnProceso;
                    //oferta.StatusSAP = SimuladorResources.CodigoEnProceso;
                    oferta.Estado = SimuladorResources.StatusSincronizado;
                }

                //Se guarda el cliente
                ClientePersistence clientePersistence = new ClientePersistence(uow);
                cliente = clientePersistence.GuardarCliente(cliente);

                //Se actualiza el id cliente
                oferta.idCliente = cliente.idCliente;
                oferta.CodClienteSAP = cliente.CodClienteSAP;

                //Se actualizan la condición de pago
                oferta.CodCondicionPago = cliente.CodCondicionPago;
                oferta.CondicionPago = cliente.CondicionPago;

                //Se guarda la oferta
                OfertaPersistence ofertaPersistence = new OfertaPersistence(uow);
                OfertaBE ofertaGuardada = ofertaPersistence.GuardarOferta(oferta);

                //Si es un alta nueva, se actualiza el idOferta de los elementos
                if (this.IsAltaNueva)
                {
                    foreach (ProductoOfertaBE item in listaProductos)
                    {
                        item.idOferta = ofertaGuardada.idOferta;
                    }
                }
                Collection<ProductoOfertaBE> listaProductosGuardados = new Collection<ProductoOfertaBE>();

                //Se guardan los productos de la oferta
                ProductoOfertaBL productoOfertaBL = new ProductoOfertaBL();
                listaProductosGuardados = productoOfertaBL.GuardarListaProductosOferta(ofertaGuardada, listaProductos, uow, esCopiaEsqueleto);

                //Se guarda el contexto
                uow.Save();

                //Se guardan los datos de configuración
                productoOfertaBL.GuardarDatosConfiguracion(listaProductosGuardados, uow);

                // Se llama a SAP para cargar la configuración
                if (isEdicionOfertaSAP)
                {
                    ConfiguracionProductosBL configProductosBL = new ConfiguracionProductosBL();
                    configProductosBL.ObtenerConfiguracionProductosFromSAP(oferta, listaProductosGuardados, usuarioLogin, passwordLogin, uow, configOfertaDescargada);

                    uow.Save();
                }

                ProductoOfertaBL objProductoOfertaBLDestinosVisibles = new ProductoOfertaBL();
                objProductoOfertaBLDestinosVisibles.ModificarConfiguracionDestinosNoVisibles(listaProductos, listaProductosGuardados);


                return ofertaGuardada;
            }
        }

    
        /// <summary>
        /// Obtiene la configuración de un producto en SAP CRM
        /// </summary>
        /// <param name="codOfertaSAP"></param>
        /// <param name="listaProductos"></param>
        /// <param name="usuarioLogin"></param>
        /// <param name="passwordLogin"></param>
        /// <returns></returns>
        public DescargaOfertaBE ObtenerProductosFromSAP(String codOfertaSAP, string usuarioLogin, string passwordLogin)
        {            
            DescargaOfertaBE resultadoDescargaOfertaBE;

            using (IUnitOfWork uow = new UnitOfWork())
            {
                // Se llama a SAP para cargar la configuración              
                ConfiguracionProductosBL configProductosBL = new ConfiguracionProductosBL();
                resultadoDescargaOfertaBE = configProductosBL.ObtenerProductosFromSAP(codOfertaSAP, usuarioLogin, passwordLogin, uow);
                                
            }

            return resultadoDescargaOfertaBE;
        }


        /// <summary>
        /// Método que realiza un copiado completo de una oferta
        /// </summary>
        /// <param name="ofertaOriginal">oferta origen</param>
        /// <param name="cliente">cliente seleccionado en el listado de clientes</param>
        /// <param name="listaProductosOferta">lista de productos de la oferta</param>
        /// <returns>Entidad OfertaBE</returns>
        public OfertaBE ClonarOferta(OfertaBE ofertaOriginal, ClienteBE cliente, Collection<ProductoOfertaBE> listaProductosOferta, bool esCopiaEsqueleto, ref Collection<String> listaProductosDesfasados)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                #region Guardado de cliente

                ClientePersistence clientePersistence = new ClientePersistence(uow);
                cliente = clientePersistence.GuardarCliente(cliente);

                #endregion

                #region Guardado de oferta

                //Se crea la nueva oferta copiando los datos de la original
                OfertaBE ofertaNueva = new OfertaBE();
                ofertaNueva.idOferta = Guid.NewGuid();
                ofertaNueva.CodOfertaSAP = string.Empty;
                ofertaNueva.Descripcion = ofertaOriginal.Descripcion;
                ofertaNueva.NumOferta = string.Empty;
                ofertaNueva.Usuario = ofertaOriginal.Usuario;
                ofertaNueva.Anio = ofertaOriginal.Anio;
                ofertaNueva.PersonaContacto = ofertaOriginal.PersonaContacto;
                ofertaNueva.CodPersonaContactoSAP = ofertaOriginal.CodPersonaContactoSAP;
                ofertaNueva.Estado = "No sincronizado";
                ofertaNueva.ValidezDe = DateTime.Now;
                //La fecha validezA será la fecha actual + la diferencia de fechas de la oferta original
                //double diferenciaDias = (ofertaOriginal.ValidezA.Value - ofertaOriginal.ValidezDe.Value).TotalDays + 1;

                DateTime fechaValidezDesdeProv = DateTime.Now.AddMonths(6);
                if (DateTime.Now.Year < fechaValidezDesdeProv.Year)
                {
                    fechaValidezDesdeProv = new DateTime(DateTime.Now.Year, 12, 31);
                }

                ofertaNueva.ValidezA = fechaValidezDesdeProv;
                ofertaNueva.FechaCreacion = DateTime.Now;
                ofertaNueva.GrupoOportunidades = ofertaOriginal.GrupoOportunidades;
                ofertaNueva.Origen = ofertaOriginal.Origen;
                ofertaNueva.Prioridad = ofertaOriginal.Prioridad;
                ofertaNueva.DescripcionStatusSAP = string.Empty;
                ofertaNueva.FechaModificacionStatusSAP = null;
                ofertaNueva.FacturacionBruta = ofertaOriginal.FacturacionBruta;
                ofertaNueva.FacturacionNeta = ofertaOriginal.FacturacionNeta;
                ofertaNueva.CodCondicionPago = ofertaOriginal.CodCondicionPago;
                ofertaNueva.CondicionPago = ofertaOriginal.CondicionPago;
                ofertaNueva.FechaUltimaModificacion = DateTime.Now;
                ofertaNueva.idCliente = cliente.idCliente;
                ofertaNueva.CodClienteSAP = cliente.CodClienteSAP;
                ofertaNueva.NombreCliente = cliente.Nombre;
                ofertaNueva.FactorCubicaje = ofertaOriginal.FactorCubicaje;

                //Se guarda la oferta
                OfertaPersistence ofertaPersistence = new OfertaPersistence(uow);
                ofertaNueva = ofertaPersistence.GuardarOferta(ofertaNueva);

                #endregion

                #region Guardar Productos de la oferta y obtener copia de datos de configuración

                #region declaración objetos

                //Listas de configuración
                Collection<ConfiguracionListaPreciosBE> listaConfigListaPrecios = new Collection<ConfiguracionListaPreciosBE>();
                Collection<ConfiguracionDestinoOfertaBE> listaConfigDestinoOferta = new Collection<ConfiguracionDestinoOfertaBE>();
                Collection<ConfiguracionTramoOfertaBE> listaConfigTramoOferta = new Collection<ConfiguracionTramoOfertaBE>();
                Collection<ConfiguracionPuntoOfertaBE> listaConfigPuntoOferta = new Collection<ConfiguracionPuntoOfertaBE>();
                Collection<ConfiguracionGradoOfertaBE> listaConfigGradoOferta = new Collection<ConfiguracionGradoOfertaBE>();
                Collection<ConfiguracionCaracteristicaBE> listaConfigCaracteristicas = new Collection<ConfiguracionCaracteristicaBE>();
                Collection<ConfiguracionValorAnadidoBE> listaConfigVAs = new Collection<ConfiguracionValorAnadidoBE>();
                Collection<GrupoTramoBE> listaConfigGrupoTramoOferta = new Collection<GrupoTramoBE>();
                Collection<ConfiguracionValorAnadidoTarifaBE> listaConfigVATarifa = new Collection<ConfiguracionValorAnadidoTarifaBE>();
                Collection<ConfiguracionValorAnadidoCaracteristicaBE> listaConfigVACaracteristica = new Collection<ConfiguracionValorAnadidoCaracteristicaBE>();

                //Clases de Business
                ConfiguracionProductosBL configProductoBL = new ConfiguracionProductosBL();
                ProductoBL productoBL = new ProductoBL();
                ConfiguracionCaracteristicasBL configCaracteristicasBL = new ConfiguracionCaracteristicasBL();
                ConfiguracionGruposTramoBL configGruposTramoBL = new ConfiguracionGruposTramoBL();
                ProductoOfertaBL productoOfertaBL = new ProductoOfertaBL();

                //Clases de persistencia
                ConfiguracionDestinoOfertaPersistence confDestinoPersistence = new ConfiguracionDestinoOfertaPersistence(uow);
                ConfiguracionTramoOfertaPersistence confTramoPersistence = new ConfiguracionTramoOfertaPersistence(uow);
                ConfiguracionValorAnadidoPersistence confVAPersistence = new ConfiguracionValorAnadidoPersistence(uow);
                ConfiguracionListaPreciosPersistence confListaPreciosPersistence = new ConfiguracionListaPreciosPersistence(uow);
                ConfiguracionPuntoOfertaPersistence confPuntoOfertaPersistence = new ConfiguracionPuntoOfertaPersistence(uow);
                ConfiguracionGradoOfertaPersistence confGradoOfertaPersistence = new ConfiguracionGradoOfertaPersistence(uow);
                CaracteristicaPersistence confCaracteristicaPersistence = new CaracteristicaPersistence(uow);
                ConfiguracionGruposTramoOfertaPersistence gruposTramoPersistence = new ConfiguracionGruposTramoOfertaPersistence(uow);
                ConfiguracionPuntoOfertaPersistence configPuntoOfertaPersistence = new ConfiguracionPuntoOfertaPersistence(uow);

                ProductoOfertaPersistence poPersistence = new ProductoOfertaPersistence(uow);
                ProductoPersistence productoPersistence = new ProductoPersistence(uow);

                Collection<ProductoOfertaBE> auxlistaProductosDefinicionVigente = new Collection<ProductoOfertaBE>();
                Collection<ProductoOfertaBE> auxlistaProductosDefinicionObsoleta = new Collection<ProductoOfertaBE>();

                //Identificador del producto oferta origen
                Guid idProductoOfertaOriginal = Guid.NewGuid();
                bool productoConDefinicionObsoleta = false;

                #endregion

                #region Guardado de los productos oferta

                //Se debe comprobar si las definiciones de los productos de la oferta son obsoletas
                //En el caso de ser obsoleta la definición de un producto se debe asociar al productoOferta el identificador del producto con la definición vigente
                foreach (ProductoOfertaBE productoOferta in listaProductosOferta)
                {
                    productoConDefinicionObsoleta = false;
                    //Se guarda en memoria el identificador del  productoOferta original
                    idProductoOfertaOriginal = productoOferta.idProductoOferta;

                    if (!poPersistence.EsDefinicionActual(productoOferta.idProducto))
                    {
                        ////Se obtiene el producto con la definición vigente
                        auxlistaProductosDefinicionObsoleta.Add(productoOferta);
                    }
                   
                        productoOferta.idOferta = ofertaNueva.idOferta;
                        productoOferta.idProductoOferta = Guid.NewGuid();
                        auxlistaProductosDefinicionVigente.Add(productoOferta);

                    //Guardado del nuevo producto oferta
                    poPersistence.GuardarCopiaProductoOferta(productoOferta);

                    //Se guarda el contexto
                    uow.Save();

                #endregion

                    #region Obtener Copia datos configuración

                    if (!productoConDefinicionObsoleta)
                    {
                        //Por cada producto se obtiene una copia de los datos de configuración con el nuevo idProducto
                        //Obtención copia de los datos de configuración
                        Collection<ConfiguracionListaPreciosBE> auxListaConfigListaPrecios = productoOfertaBL.CopiarConfiguracionListaPrecios(idProductoOfertaOriginal, productoOferta.idProductoOferta);
                        Collection<ConfiguracionDestinoOfertaBE> auxListaConfigDestinoOferta = configProductoBL.CopiarConfiguracionDestinoProductoOferta(idProductoOfertaOriginal, productoOferta.idProductoOferta);
                        Collection<ConfiguracionTramoOfertaBE> auxListaConfigTramoOferta = configProductoBL.CopiarConfiguracionTramoProductoOferta(idProductoOfertaOriginal, productoOferta.idProductoOferta);
                        Collection<ConfiguracionPuntoOfertaBE> auxListaConfigPuntoOferta = configProductoBL.CopiarConfiguracionPuntoProductoOferta(idProductoOfertaOriginal, productoOferta.idProductoOferta);
                        Collection<ConfiguracionGradoOfertaBE> auxListaConfigGradoOferta = productoBL.CopiarConfiguracionGradoProductoOferta(idProductoOfertaOriginal, productoOferta.idProductoOferta);
                        Collection<ConfiguracionCaracteristicaBE> auxListaConfigCaracteristicas = configCaracteristicasBL.CopiarConfiguracionCaracteristicasProductoOferta(idProductoOfertaOriginal, productoOferta.idProductoOferta);
                        Collection<ConfiguracionValorAnadidoBE> auxListaConfigVA = configProductoBL.ObtenerCopiaListaConfiguracionVA(idProductoOfertaOriginal, productoOferta.idProductoOferta);
                        Collection<GrupoTramoBE> auxListaConfigGrupoTramoOferta = configGruposTramoBL.ObtenerCopiaListaConfiguracionGruposTramoOferta(idProductoOfertaOriginal, productoOferta.idProductoOferta);

                        Collection<ConfiguracionValorAnadidoTarifaBE> auxListaConfigVATarifa = productoBL.ObtenerListaConfiguracionValorAnadidoTarifa(auxListaConfigVA);
                        Collection<ConfiguracionValorAnadidoCaracteristicaBE> auxListaConfigVACaracteristica = productoBL.ObtenerListaConfiguracionValorAnadidoCaracteristica(auxListaConfigVA);

                        //Se añaden los listados de datos de cada producto a las listas que contienen todos los datos de configuración a nivel de oferta
                        listaConfigListaPrecios = listaConfigListaPrecios.Concat(auxListaConfigListaPrecios).ToList<ConfiguracionListaPreciosBE>().ToCollection<ConfiguracionListaPreciosBE>();
                        listaConfigDestinoOferta = listaConfigDestinoOferta.Concat(auxListaConfigDestinoOferta).ToList<ConfiguracionDestinoOfertaBE>().ToCollection<ConfiguracionDestinoOfertaBE>();
                        listaConfigTramoOferta = listaConfigTramoOferta.Concat(auxListaConfigTramoOferta).ToList<ConfiguracionTramoOfertaBE>().ToCollection<ConfiguracionTramoOfertaBE>();
                        listaConfigPuntoOferta = listaConfigPuntoOferta.Concat(auxListaConfigPuntoOferta).ToList<ConfiguracionPuntoOfertaBE>().ToCollection<ConfiguracionPuntoOfertaBE>();
                        listaConfigGradoOferta = listaConfigGradoOferta.Concat(auxListaConfigGradoOferta).ToList<ConfiguracionGradoOfertaBE>().ToCollection<ConfiguracionGradoOfertaBE>();
                        listaConfigCaracteristicas = listaConfigCaracteristicas.Concat(auxListaConfigCaracteristicas).ToList<ConfiguracionCaracteristicaBE>().ToCollection<ConfiguracionCaracteristicaBE>(); ;
                        listaConfigVAs = listaConfigVAs.Concat(auxListaConfigVA).ToList<ConfiguracionValorAnadidoBE>().ToCollection<ConfiguracionValorAnadidoBE>();
                        listaConfigGrupoTramoOferta = listaConfigGrupoTramoOferta.Concat(auxListaConfigGrupoTramoOferta).ToList<GrupoTramoBE>().ToCollection<GrupoTramoBE>();
                        listaConfigVATarifa = listaConfigVATarifa.Concat(auxListaConfigVATarifa).ToList<ConfiguracionValorAnadidoTarifaBE>().ToCollection<ConfiguracionValorAnadidoTarifaBE>();
                        listaConfigVACaracteristica = listaConfigVACaracteristica.Concat(auxListaConfigVACaracteristica).ToList<ConfiguracionValorAnadidoCaracteristicaBE>().ToCollection<ConfiguracionValorAnadidoCaracteristicaBE>();

                    }

                    #endregion
                }

                #endregion

                #region Guardado de la estructura de configuraciones

                //Se guardan las configuraciones de todas las tablas de configuración inicializados a 0 para los productos que tienen definiciones obsoletas
                //productoOfertaBL.GuardarDatosConfiguracionCompleto(auxlistaProductosDefinicionObsoleta, uow);

                //Se guardan las configuraciones de tramos inicializados a 0 para los productos que tienen una definición vigente
                productoOfertaBL.GuardarDatosConfiguracion(auxlistaProductosDefinicionVigente, uow);

                #endregion

                #region Guardado de los registros de configuraciones con valor

                confDestinoPersistence.GuardarConfiguracionDestinosOferta(listaProductosOferta, listaConfigDestinoOferta);
                confTramoPersistence.GuardarConfiguracionTramosOferta(listaProductosOferta, listaConfigTramoOferta);
                confListaPreciosPersistence.GuardarConfiguracionListasPrecios(listaProductosOferta, listaConfigListaPrecios);
                configPuntoOfertaPersistence.GuardarConfiguracionPuntosOferta(listaProductosOferta, listaConfigPuntoOferta);
                
                //Se guardan los registros de configuracionVA y los registros en configVATarifa y configVACaracteristica
                confVAPersistence.GuardarConfiguracionVAsOferta(listaProductosOferta, listaConfigVAs, listaConfigVATarifa, listaConfigVACaracteristica);
                confGradoOfertaPersistence.GuardarConfiguracionGradosOferta(listaProductosOferta, listaConfigGradoOferta);
                confCaracteristicaPersistence.GuardarConfiguracionCaracteristicasOferta(listaProductosOferta, listaConfigCaracteristicas);
                gruposTramoPersistence.InsertarConfiguracionGruposTramoOferta(listaConfigGrupoTramoOferta);                

                #endregion

                //Se guarda el contexto
                uow.Save(); //Se guarda en RevisarDefinicionActualProductos

                //Intentamos copiar toda la información posible en los productos desfasados. Ya no se borra su configuración.
                System.Text.StringBuilder textoProdsDesfasados;
                auxlistaProductosDefinicionObsoleta = new ProductoOfertaBL().RevisarDefinicionActualProductos(ofertaNueva, listaProductosOferta, esCopiaEsqueleto, out textoProdsDesfasados);

                foreach (var item in auxlistaProductosDefinicionObsoleta)
                    listaProductosDesfasados.Add(item.CodProductoSAP);

                return ofertaNueva;
            }
        }

        #endregion

        #region Actualizar

        /// <summary>
        /// Recorre todas las ofertas que hay en base de datos y actualiza el status de las mismas.
        /// </summary>
        /// <param name="usuario"></param>
        /// <param name="password"></param>
        /// <param name="uow"></param>
        public void ActualizarStatusOfertasSAP(string usuario, string password)
        {
            //Obtener todas las ofertas
            Collection<string> listaCodOfertasSAP = this.ObtenerOfertasSincronizadas();

            if (listaCodOfertasSAP.Count > 0)
            {
                EstadoOfertaBL objEstadoOferta = new EstadoOfertaBL();
                objEstadoOferta.ActualizarEstadosSAP(usuario, password, listaCodOfertasSAP);
            }
        }

        /// <summary>
        /// Actualiza la tabla FactorCubicaje y recorre todas las ofertas que hay en base de datos actualizando el valor si es necesario
        /// </summary>
        /// <param name="usuario"></param>
        /// <param name="password"></param>
        public void ActualizarCubicajeOfertas(string usuario, string password)
        {
            return; // TODO QUITAR PARA ACTIVAR CUBICAJE. Se hace return porque en pro no existe esta función
            using (IUnitOfWork uow = new UnitOfWork())
            {
                CubicajePersistence cubicajePersistence = new CubicajePersistence(uow);
                Collection<KeyValuePair<string, bool>> cubicajes = null;
                if (SSOHelper.Instance.LogarConSSO)
                {
                    CommunicatorLight conectorSAP = new CommunicatorLight(SSOHelper.Instance.Usuario, password);
                    //SSOHelper.Instance.ActualizarCookiePortal();
                    SSOHelper.Instance.InicializarWSLight(conectorSAP.Cliente);
                    cubicajes = conectorSAP.ZCObtenerCubicajesRfc();
                    SSOHelper.Instance.LimpiarWSLight();
                }
                else
                {
                CommunicatorLight conectorSAP = new CommunicatorLight(usuario, password);
                    cubicajes = conectorSAP.ZCObtenerCubicajesRfc();
                }

                if (cubicajePersistence.ActualizarCubicajesDisponibles(cubicajes))
                {
                    OfertaPersistence persistencia = new OfertaPersistence(uow);
                    persistencia.ActualizarCubicajeOfertas(cubicajes.Select(x => x.Key).ToList().ToCollection(), cubicajes.FirstOrDefault(x => x.Value).Key);
                    uow.Save();
                }
            }
            }

        /// <summary>
        /// Actualiza el campo de cubicaje de una oferta concreta
        /// </summary>
        /// <param name="idOferta"></param>
        /// <param name="cubicajePorDefecto"></param>
        public void ActualizarCubicajeOferta(Guid idOferta, string cubicajePorDefecto) 
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                OfertaPersistence persistencia = new OfertaPersistence(uow);
                persistencia.ActualizarCubicajeOferta(idOferta, cubicajePorDefecto);
                uow.Save();
            }
        }

        /// <summary>
        /// Método para guardar el código de SAP devuelto por SAP al crear una nueva oferta
        /// </summary>
        /// <param name="idOFerta"></param>
        /// <param name="codSAP"></param>
        /// <returns></returns>
        private OfertaBE ActualizarEstadoOferta(Guid idOFerta, string codSAP)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                OfertaBE oferta = new OfertaBE();
                oferta = this.ObtenerOfertaByIdOferta(idOFerta);

                oferta.CodOfertaSAP = codSAP;

                if (codSAP != null && codSAP != String.Empty)
                {

                    oferta.ValidezDefinitiva = true;

                }

                if (oferta.StatusSAP == null)
                {
                    oferta.DescripcionStatusSAP = SimuladorResources.StatusEnProceso;
                    oferta.StatusSAP = SimuladorResources.CodigoEnProceso;
                }
                
                oferta.Estado = SimuladorResources.StatusSincronizado;

                OfertaPersistence persistencia = new OfertaPersistence(uow);
                persistencia.GuardarOferta(oferta);

                uow.Save();

                return oferta;
            }
        }

        #endregion

        #region Enviar SAP

        /// <summary>
        /// Metodo que construye la estructura de datos necesaria para enviarle toda la información de una oferta a SAP
        /// </summary>
        /// <param name="idOferta">Código de oferta a actualizar</param>
        /// <param name="listaProductos">Lista de productos a envíar a SAP</param>
        /// <returns>true/false</returns>
        public ResultadoCargaBE EnviarDatosOfertaSAP(string usuario, string password, Guid idOferta, Collection<ProductoOfertaBE> listaProductos, Collection<ProductoOfertaBE> listaVAs, string statusOferta, string codOferta)
        {
            ResultadoCargaBE objRespuestaValidacionesInternas = new ResultadoCargaBE();
            ProductoBL productoDev = new ProductoBL();

            //Si la oferta está en estado "Borrador" no se realizan las validaciones
            if (!statusOferta.Equals(SimuladorResources.CodigoEnBorrador))
            {
                //Validaciones de datos en nuestro lado antes de realizar la llamada a CRM
                objRespuestaValidacionesInternas = productoDev.ValidarDatos(idOferta, listaProductos, statusOferta, codOferta);
            }

            if (objRespuestaValidacionesInternas.errores.Count.Equals(0))
            {
                ResultadoCargaBE objRespuestaCarga = null;
                if (SSOHelper.Instance.LogarConSSO)
                {
                    CommunicatorLight conectorSAP = new CommunicatorLight(SSOHelper.Instance.Usuario, password);
                    //SSOHelper.Instance.ActualizarCookiePortal();
                    SSOHelper.Instance.InicializarWSLight(conectorSAP.Cliente);
                    objRespuestaCarga = conectorSAP.ZCCrearOportOfertaRfc(idOferta, listaProductos, listaVAs, usuario.ToUpper(), statusOferta);
                    SSOHelper.Instance.LimpiarWSLight();
                }
                else
                {
                    CommunicatorLight conectorSAP = new CommunicatorLight(usuario, password);
                    //Si no se está en modo administrador el usuario del SSO será el de windows
                    if (!Utils.GetValorFromAppConfig(AppSettingsEnum.ModoLoginAdmin).Equals("True"))
                    {
                        usuario = Environment.UserName.ToUpper();
                    }

#if DEBUG
                    if (Utils.GetValorFromAppConfig(AppSettingsEnum.UsuariosDebug).ToUpper().Contains(Environment.UserName.ToUpper()) == true)
                    {
                        usuario = "E007639";
                        usuario = Utils.GetValorFromAppConfig(AppSettingsEnum.UsuarioPrueba);
                    }
#endif

                    objRespuestaCarga = conectorSAP.ZCCrearOportOfertaRfc(idOferta, listaProductos, listaVAs, usuario.ToUpper(), statusOferta);
                }

                if (objRespuestaCarga.errores.Count.Equals(0))
                {
                    //Si es un guardado SINCRONO
                    if(objRespuestaCarga.estadoSincronizacion == null)
                    {
                        //Actualizamos el estado y el código de la oferta.
                        this.ActualizarEstadoOferta(idOferta, objRespuestaCarga.oferta);

                        //Borramos si existe el estado de la sincronización (Puede permanecer de previas sincros)
                        new EstadoSincronizacionBL().EliminarEstadoSincronizacion(idOferta);

                    }
                    //Si es un guardado ASINCRONO
                    else
                    {
                        //Actualizamos el estado de sincronización
                        new EstadoSincronizacionBL().InsertEstadoSincronizacion(objRespuestaCarga.estadoSincronizacion);
                        objRespuestaValidacionesInternas.estadoSincronizacion = objRespuestaCarga.estadoSincronizacion;
                    }
                }
                else
                {
                    return objRespuestaCarga;
                }
            }
            else
            {
                return objRespuestaValidacionesInternas;
            }

            return objRespuestaValidacionesInternas;
        }



        /// <summary>
        /// Metodo que construye la estructura de datos necesaria para enviarle toda la información de una oferta a SAP de manera asíncrona
        /// </summary>
        /// <param name="idOferta">Código de oferta a actualizar</param>
        /// <param name="listaProductos">Lista de productos a envíar a SAP</param>
        /// <returns>true/false</returns>
        public ResultadoCargaBE EnviarDatosOfertaSAPAsincrono(string usuario, string password, Guid idOferta, Collection<ProductoOfertaBE> listaProductos, Collection<ProductoOfertaBE> listaVAs, string statusOferta, string codOferta)
        {
            ResultadoCargaBE objRespuestaValidacionesInternas = new ResultadoCargaBE();
            ProductoBL productoDev = new ProductoBL();

            //Si la oferta está en estado "Borrador" no se realizan las validaciones
            if (!statusOferta.Equals(SimuladorResources.CodigoEnBorrador))
            {
                //Validaciones de datos en nuestro lado antes de realizar la llamada a CRM
                objRespuestaValidacionesInternas = productoDev.ValidarDatos(idOferta, listaProductos, statusOferta, codOferta);
            }

            if (objRespuestaValidacionesInternas.errores.Count.Equals(0))
            {
                ResultadoCargaBE objRespuestaCarga = null;
                if (SSOHelper.Instance.LogarConSSO)
                {
                    CommunicatorLight conectorSAP = new CommunicatorLight(SSOHelper.Instance.Usuario, password);
                    //SSOHelper.Instance.ActualizarCookiePortal();
                    SSOHelper.Instance.InicializarWSLight(conectorSAP.Cliente);
                    objRespuestaCarga = conectorSAP.ZCCrearOportOfertaRfcAsincrono(idOferta, listaProductos, listaVAs, usuario.ToUpper(), statusOferta);
                    SSOHelper.Instance.LimpiarWSLight();
                }
                else
                {
                    CommunicatorLight conectorSAP = new CommunicatorLight(usuario, password);
                    //Si no se está en modo administrador el usuario del SSO será el de windows
                    if (!Utils.GetValorFromAppConfig(AppSettingsEnum.ModoLoginAdmin).Equals("True"))
                    {
                            usuario = Environment.UserName.ToUpper();
                    }

                    objRespuestaCarga = conectorSAP.ZCCrearOportOfertaRfcAsincrono(idOferta, listaProductos, listaVAs, usuario.ToUpper(), statusOferta);
                }

                if (objRespuestaCarga.errores.Count.Equals(0))
                {
                    //Actualizamos el estado y el código de la oferta.
                    this.ActualizarEstadoOferta(idOferta, objRespuestaCarga.oferta);
                    new EstadoSincronizacionBL().InsertEstadoSincronizacion(objRespuestaCarga.estadoSincronizacion);
                }
                else
                {
                    return objRespuestaCarga;
                }
            }
            else
            {
                return objRespuestaValidacionesInternas;
            }

            return objRespuestaValidacionesInternas;
        }

        /// <summary>
        /// Metodo que construye la estructura de datos necesaria para obtener el estado de las sincronizaciones en SAP CRM
        /// </summary>
        /// <param name="usuario"></param>
        /// <param name="password"></param>
        /// <returns>true/false</returns>
        public void ActualizarEstadosSincronizaciones(string usuario, string password)
        {
            CommunicatorLight conectorSAP; 
            EstadoSincronizacionBL estadoSincroBL = new EstadoSincronizacionBL();
            ResultadoCargaBE resultado = new ResultadoCargaBE();
            
            //Nos conectamos
            if (SSOHelper.Instance.LogarConSSO)
            {
                conectorSAP = new CommunicatorLight(SSOHelper.Instance.Usuario, password);            
                SSOHelper.Instance.InicializarWSLight(conectorSAP.Cliente);
            }
            else
            {
                conectorSAP = new CommunicatorLight(usuario, password);
                //Si no se está en modo administrador el usuario del SSO será el de windows
                if (!Utils.GetValorFromAppConfig(AppSettingsEnum.ModoLoginAdmin).Equals("True"))                
                    usuario = Environment.UserName.ToUpper();                                    
            }

            //Recorremos todos los estados de sincronización y los vamos actualizando consultando a SAP CRM
            var listadoEstadosSincro = estadoSincroBL.ObtenerEstadosSincronizacion();

            foreach (var item in listadoEstadosSincro)
            {

                if (item.EstadoOferta != null && !item.EstadoOferta.Equals("Error"))
                {
                //Actualizamos el estado de sincronización de la oferta
                resultado = conectorSAP.ZCCrearOportOfertaRfcEstado(usuario.ToUpper(), item);
                estadoSincroBL.ActualizarEstadoSincronizacion(resultado.estadoSincronizacion.idOferta, resultado.estadoSincronizacion.idEstadoSincronizacion, resultado.estadoSincronizacion.EstadoOferta);

                //Se lanza un mensaje de error indicando la lista de errores detectados
                foreach (ErrorCargaBE error in resultado.errores)
                {
                    if (error.error == "El id unico enviado no existe")
                    {
                        error.error = "Se ha producto un error en la sincronización. Compruebe que la oferta no está bloqueada en SAP y pruebe a volver a sincronizarla.";
                    }

                    resultado.sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "Destino:{0} -- Error:{1} -- Fichero:{2} -- Oferta:{3} -- Oportunidad:{4} -- Posicion:{5} -- Producto:{6} -- Tramo:{7} -- TramoFin:{8} -- TramoIni:{9} -- VA:{10}",
                        error.destino, error.error, error.fichero, error.oferta, error.oportunidad, error.posicion, error.producto, error.tramo, error.tramoFin, error.tramoIni, error.VA));
                    //Añado el error un StringBuilder para mostrarselo al usuario
                    resultado.sbErroresSincronizarSAP.AppendLine(error.error);
                }

                if (resultado.sb.Length > 0)
                {
                    RegistrarAccionesSimulador.GuardarTraza(new AuditoriaDataTrazaBE(SimuladorResources.SincronizarCRM + " -> " + resultado.sb.ToString(),
                                SimuladorResources.AccionAccesoSAPCRM, usuario, password));

                        MessageBox.Show(string.Format(CultureInfo.InvariantCulture, SimuladorResources.MensajeErroresSincronizar, resultado.Cliente, resultado.Descripcion, Environment.NewLine + Environment.NewLine, resultado.sbErroresSincronizarSAP.ToString()), SimuladorResources.Advertencia, MessageBoxButtons.OK, MessageBoxIcon.Warning, 0);
                    } 

                //Actualizamos la oferta si se ha sincronizado
                if (!String.IsNullOrEmpty(resultado.oferta))
                {
                    new OfertaBL().ActualizarEstadoOferta(resultado.estadoSincronizacion.idOferta, resultado.oferta);
                }
            }
            }
            
            if (SSOHelper.Instance.LogarConSSO)
            {
                SSOHelper.Instance.LimpiarWSLight();
            }
        }
        
        #endregion

        #region Eliminar

        /// <summary>
        /// Método que elimina la oferta cuyo id se pasa por parámetro así como toda su información asociada
        /// </summary>
        /// <param name="idOferta">Identificador de la oferta a eliminar</param>
        public void EliminarOferta(Guid idOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                uow.DeshabilitarAutodeteccionDeCambios();

                //Se eliminan los productos 
                ProductoOfertaPersistence productoOfertaPersistence = new ProductoOfertaPersistence(uow);
                productoOfertaPersistence.EliminarProductosOferta(idOferta);

                //Se elimina la oferta
                OfertaPersistence ofertaPersistence = new OfertaPersistence(uow);
                ofertaPersistence.EliminarOferta(idOferta);

                //Se guarda el contexto
                uow.Save();
            }
        }

        #endregion

        public void FijarValidezDefinitiva (OfertaBE oferta){
            
            using (IUnitOfWork uow = new UnitOfWork())
            {
                OfertaPersistence objPersistencia = new OfertaPersistence(uow);
                objPersistencia.FijarValidezDefinitiva(oferta);
                uow.Save();
            }

        }

        public void ModificarFechaOferta(OfertaBE oferta, DateTime fechaDesde, DateTime fechaHasta)
        {

            using (IUnitOfWork uow = new UnitOfWork())
            {
                OfertaPersistence objPersistencia = new OfertaPersistence(uow);
                objPersistencia.ModificarFechaOferta(oferta, fechaDesde, fechaHasta);
                uow.Save();
            }

        }

        #endregion

        #region Métodos privados

        private void CopiarDatosConfiguraciónProducto(Guid idProductoOferta, Collection<ConfiguracionTramoOfertaBE> listaConfigTramoOferta, Collection<ConfiguracionDestinoOfertaBE> listaConfigDestinoOferta,
                                Collection<ConfiguracionPuntoOfertaBE> listaConfigPuntoOferta, Collection<ConfiguracionGradoOfertaBE> listaConfigGradoOferta,
                                Collection<ConfiguracionCaracteristicaBE> listaConfigCaracteristicas, Collection<ConfiguracionValorAnadidoBE> listaConfigVA,
                                Collection<GrupoTramoBE> listaConfigGrupoTramoOferta, Collection<ConfiguracionListaPreciosBE> listaConfigListaPrecios)
        {

        }


        #endregion
    }
}
