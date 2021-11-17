using System;

using System.Collections.ObjectModel;
using System.Linq;
using Correos.SimuladorOfertas.Common;
using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.Persistence;
using System.Text;
using System.Globalization;
using System.Collections.Generic;

namespace Correos.SimuladorOfertas.Business
{
    public class ProductoOfertaBL
    {
        #region Metodos publicos

        /// <summary>
        /// Obtiene el sumatorio de los envios de los productos de la oferta.
        /// </summary>
        /// <param name="idOferta"></param>
        /// <returns></returns>
        public decimal ObtenerSumatorioNumerosEnviosOferta(Guid idOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProductoOfertaPersistence objPersistencia = new ProductoOfertaPersistence(uow);
                return objPersistencia.ObtenerEnviosEnOferta(idOferta);
            }
        }

        /// <summary>
        /// Devuelve true si todos los tramos de los productos si está negociado en modo Precio Cierto tienen un valor
        /// </summary>
        /// <param name="productosSeleccionados"></param>
        /// <returns></returns>
        public ResultBE ComprobarPrecioCiertoTramosEnModalidad(Collection<ProductoOfertaBE> productosSeleccionados)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProductoOfertaPersistence objPersistencia = new ProductoOfertaPersistence(uow);
                return objPersistencia.ComprobarPrecioCiertoTramosEnModalidad(productosSeleccionados);
            }
        }

        /// <summary>
        /// Devuelve true si la lista cumple objetivos de % de cada destino y del sum de tramos de cada zona
        /// </summary>
        /// <param name="productosSeleccionados"></param>
        /// <returns></returns>
        public ResultBE ComprobarPorcentajesListaProductos(Collection<ProductoOfertaBE> productosSeleccionados)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProductoOfertaPersistence objPersistencia = new ProductoOfertaPersistence(uow);
                return objPersistencia.ComprobarPorcentajesProductosOferta(productosSeleccionados);
            }
        }

        /// <summary>
        /// Devuelve true si la lista cumple objetivos de % de cada destino y del sum de tramos de cada zona
        /// </summary>
        /// <param name="productosSeleccionados"></param>
        /// <returns></returns>
        public ResultBE ComprobarPuntosGradosPorcentajesListaProductos(ProductoOfertaBE productosSeleccionados)
        {
            ResultBE objRespuesta = new ResultBE();

            using (IUnitOfWork uow = new UnitOfWork())
            {
                decimal aux100 = 100;
                decimal auxValor = 0;

                //Se obtienen los grados del productooferta
                ConfiguracionGradoOfertaPersistence cgPersistence = new ConfiguracionGradoOfertaPersistence(uow);
                ConfiguracionGradoOfertaBE cgBE = cgPersistence.ObtenerConfiguracionGradoOferta(productosSeleccionados.idProductoOferta);

                if (cgBE != null)
                {
                    auxValor = cgBE.DistribucionG0.Value + cgBE.DistribucionG1.Value + cgBE.DistribucionG2.Value;

                    if (auxValor != aux100)
                    {
                        objRespuesta.Resultado = false;
                        objRespuesta.TextoError = string.Format(SimuladorResources.SumaDistribucionGradosDebeSumar100, productosSeleccionados.CodProductoSAP);
                        return objRespuesta;
                    }
                }

                auxValor = 0;

                //Se obtienen los puntos del productooferta
                ConfiguracionPuntoOfertaPersistence cpPersistence = new ConfiguracionPuntoOfertaPersistence(uow);
                ConfiguracionPuntoOfertaBE cpBE = cpPersistence.ObtenerConfiguracionPuntoOferta(productosSeleccionados.idProductoOferta);

                if (cpBE != null)
                {
                    auxValor = cpBE.DistribucionCAM.Value + cpBE.DistribucionRUR.Value + cpBE.DistribucionURB.Value;

                    if (auxValor != aux100)
                    {
                        objRespuesta.Resultado = false;
                        objRespuesta.TextoError = string.Format(SimuladorResources.SumaDistribucionPuntosDebeSumar100, productosSeleccionados.CodProductoSAP);
                        return objRespuesta;
                    }
                }
            }

            return objRespuesta;
        }

        /// <summary>
        /// Devuelve true si la lista cumple objetivos de % de cada destino y del sum de tramos de cada zona
        /// </summary>
        /// <param name="productosSeleccionados"></param>
        /// <returns></returns>
        public ResultBE ComprobarProductosSoloLectura(Collection<ProductoOfertaBE> listaProductosSeleccionados, OfertaBE oferta)
        {
            bool mostrarMensajeYNoEjecutar = false;
            ResultBE objRespuesta = new ResultBE();

                ProductoBL poBL = new ProductoBL();
                Collection<VisibilidadProductoBE> productosSoloLectura = poBL.ObtenerProductosSoloLectura();
                Collection<ProductoOfertaBE> productosSoloLecturaOferta = new Collection<ProductoOfertaBE>();

                foreach (ProductoOfertaBE item in listaProductosSeleccionados)
                {
                    VisibilidadProductoBE itemOferta = productosSoloLectura.FirstOrDefault(x => x.CodProductoSAP.Equals(item.CodProductoSAP) && x.SoloLectura);
                    if(itemOferta != null){
                        productosSoloLecturaOferta.Add(item);
                    }

                }

                if (productosSoloLecturaOferta.Count > 0)
                {
                    if ((oferta.CodOfertaSAP == null || oferta.CodOfertaSAP.Equals("-") || oferta.CodOfertaSAP.Equals(String.Empty) || oferta.StatusSAP == "E0013" || oferta.StatusSAP == "E0016"))
                    {
                        mostrarMensajeYNoEjecutar = true;
                    }

                    string mensaje = String.Empty;
                    string titulo = "Aviso de productos";
                    //MessageBoxButtons botones = MessageBoxButtons.OK;

                    if (!mostrarMensajeYNoEjecutar)
                    {
                        objRespuesta.Resultado = true;
                        objRespuesta.TextoError = "Su oferta contiene productos obsoletos: \n\n"; ;
                    }
                    else
                    {
                        objRespuesta.Resultado = false;
                        objRespuesta.TextoError = "No puede continuar debido a la presencia en la oferta de productos que ya no pueden ser tratados en el portal SAP.\n\n Por favor, edite la oportunidad para eliminar los productos: \n\n";
                    }

                    foreach (ProductoOfertaBE po in productosSoloLecturaOferta)
                    {
                        objRespuesta.TextoError += po.Anexo + "-" + po.CodProductoSAP + " '" + po.Descripcion + " - " + po.ModalidadNegociacion + "'. \n";

                    }
                    //DialogResult result = MessageBox.Show(mensaje, titulo, botones, MessageBoxIcon.Warning);

            }


            return objRespuesta;

        }

        /// <summary>
        /// Devuelve true si la lista cumple objetivos de % de cada destino y del sum de tramos de cada zona
        /// </summary>
        /// <param name="productosSeleccionados"></param>
        /// <returns></returns>
        public ResultBE ComprobarProductosSoloLectura(OfertaBE oferta)
        {
            bool mostrarMensajeYNoEjecutar = false;
            ResultBE objRespuesta = new ResultBE();

            ProductoBL poBL = new ProductoBL();
            Collection<VisibilidadProductoBE> productosSoloLectura = poBL.ObtenerProductosSoloLectura();
            Collection<ProductoOfertaBE> productosSoloLecturaOferta = new Collection<ProductoOfertaBE>();
            ProductoOfertaBL objProductoOfertaBL = new ProductoOfertaBL();
            Collection<ProductoOfertaBE> listaProductosSeleccionados = objProductoOfertaBL.ObtenerProductosEnOferta(oferta.idOferta);

            foreach (ProductoOfertaBE item in listaProductosSeleccionados)
            {
                VisibilidadProductoBE itemOferta = productosSoloLectura.FirstOrDefault(x => x.CodProductoSAP.Equals(item.CodProductoSAP));
                if (itemOferta != null)
                {
                    productosSoloLecturaOferta.Add(item);
                }

            }

            if (productosSoloLecturaOferta.Count > 0)
            {
                if ((oferta.CodOfertaSAP == null || oferta.CodOfertaSAP.Equals("-") || oferta.CodOfertaSAP.Equals(String.Empty) || oferta.StatusSAP == "E0013" || oferta.StatusSAP == "E0016"))
                {
                    mostrarMensajeYNoEjecutar = true;
                }

                string mensaje = String.Empty;
                string titulo = "Aviso de productos";
                //MessageBoxButtons botones = MessageBoxButtons.OK;

                if (!mostrarMensajeYNoEjecutar)
                {
                    objRespuesta.Resultado = true;
                    objRespuesta.TextoError = "Su oferta contiene productos obsoletos: \n\n"; ;
                }
                else
                {
                    objRespuesta.Resultado = false;
                    objRespuesta.TextoError = "No puede continuar debido a la presencia en la oferta de productos que ya no pueden ser tratados en el portal SAP.\n\n Por favor, edite la oferta para eliminar los productos: \n\n";
                }

                foreach (ProductoOfertaBE po in productosSoloLecturaOferta)
                {
                    objRespuesta.TextoError += po.Anexo + "-" + po.CodProductoSAP + " '" + po.Descripcion + " - " + po.ModalidadNegociacion + "'. \n";

                }
                //DialogResult result = MessageBox.Show(mensaje, titulo, botones, MessageBoxIcon.Warning);

            }


            return objRespuesta;

        }

        /// <summary>
        /// Comprueba que la selección de productos es correcta.
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public ResultBE ComprobarProductosSeleccionados(Collection<ProductoOfertaBE> list)
        {
            ResultBE objRespuesta = new ResultBE();
            RelacionProductosBL rpBl = new RelacionProductosBL();

            bool estaEnLaLista = true;

            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProductoPersistence productoPersistence = new ProductoPersistence(uow);

                //Recorrer listado productos
                foreach (ProductoOfertaBE objProductoOferta in list)
                {
                    var listaProdCampoA = rpBl.ObtenerRelacionProductos(objProductoOferta.CodProductoSAP, EsCampoA: true);
                    var listaProdCampoB = rpBl.ObtenerRelacionProductos(objProductoOferta.CodProductoSAP, EsCampoA: false);

                    if (objRespuesta.Resultado)
                    {

                    if (listaProdCampoA.Count > 0)
                    {   
                        foreach (var codSapA in listaProdCampoA)
                        {
                                //[MMunoz] Se realiza la comprobación a nivel General, no de Anexo.
                                //estaEnLaLista = estaEnLaLista & list.Any(t => t.CodProductoSAP.Equals(codSapA) && t.Anexo.Equals(objProductoOferta.Anexo));
                                estaEnLaLista = estaEnLaLista & list.Any(t => t.CodProductoSAP.Equals(codSapA));                                
                        }

                        if (estaEnLaLista)
                        {
                            objRespuesta.Resultado = true;
                        }

                        else
                        {
                            objRespuesta.Resultado = false;
                            objRespuesta.TextoError = SimuladorResources.ErrrorSeleccionProductos;
                        }
                    }

                    else if (listaProdCampoB.Count > 0)
                    {
                        objRespuesta.Resultado = false;
                        
                        foreach (var codSapB in listaProdCampoB)
                        {
                            if (!objRespuesta.Resultado)
                            {

                                if (list.Any(t => t.CodProductoSAP.Equals(codSapB)))
                                {
                                    objRespuesta.Resultado = true;
                                }
                                else
                                {
                                    objRespuesta.Resultado = false;
                                    objRespuesta.TextoError = SimuladorResources.ErrrorSeleccionProductos;
                                }
                            }
                        }
                    }
                    }
                    //if (estaEnLaLista)
                    //{
                    //    objRespuesta.Resultado = true;
                    //    return objRespuesta;
                    //}

                    //else
                    //{
                    //    objRespuesta.Resultado = false;
                    //    objRespuesta.TextoError = SimuladorResources.ErrrorSeleccionProductos;
                    //    return objRespuesta;
                    //}
                    //if (objProductoOferta.CodProductoSAP == SimuladorResources.S0134)
                    //{
                    //    estaS0134 = true;
                    //}
                    //if ((objProductoOferta.CodProductoSAP == SimuladorResources.S0132) || (objProductoOferta.CodProductoSAP == SimuladorResources.S0133)
                    //    || (objProductoOferta.CodProductoSAP == SimuladorResources.S0235) || (objProductoOferta.CodProductoSAP == SimuladorResources.S0236))
                    //{
                    //    estaResto = true;
                    //}
                }
                //if ((estaS0134 && estaResto) || (!estaS0134 && !estaResto))
                //{

                //}
                //else
                //{
                //    objRespuesta.Resultado = false;
                //    objRespuesta.TextoError = SimuladorResources.ErrrorSeleccionProductos;
                //    return objRespuesta;
                //}
            }

            return objRespuesta;
        }

        /// <summary>
        /// Devuelve todos los productos que hay en una oferta concreta.
        /// </summary>
        /// <param name="codigoOferta"></param>
        /// <returns></returns>
        public Collection<ProductoOfertaBE> ObtenerProductosEnOferta(Guid idOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProductoOfertaPersistence persistencia = new ProductoOfertaPersistence(uow);
                return persistencia.ObtenerProductosEnOferta(idOferta);
            }
        }


        /// <summary>
        /// Devuelve todos los idproductoOferta de un peroducto
        /// solo de las ofertas que estan en borrasor o en procedo
        /// </summary>
        /// <param name="codProductoSap">Código del producto a buscar</param>
        public Collection<ProductoOfertaBE> ObtenerProductosEnOfertaPorCodigoProductoSAP(string codProductoSap)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProductoOfertaPersistence persistencia = new ProductoOfertaPersistence(uow);
                return persistencia.ObtenerProductosEnOfertaPorCodigoProductoSAP(codProductoSap);
            }
        }


        /// <summary>
        /// Devuelve todos los idConfiguracionTramoOferta de un productoOferta
        /// solo de las ofertas que estan en borrasor o en procedo
        /// </summary>
        /// <param name="idProductoOferta">Código del producto a buscar</param>
        public Collection<ConfiguracionTramoOfertaBE> ObtenerProductoOfertaTramos(Guid idProductoOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ConfiguracionTramoOfertaPersistence persistencia = new ConfiguracionTramoOfertaPersistence(uow);
                return persistencia.ObtenerConfiguracionTramoOfertaPorProductoOferta(idProductoOferta);
            }
        }
        /// <summary>
        /// Obtenemos una lista de objetos de ProductoOfertaBE segun CodProductoSAP
        /// <para>Este método se usa para la actualización de agrupaciones de tipologias, cuando uno de los productos de una agrupación ha sido modificado</para>
        /// </summary>
        /// <param name="codProductoSap">Código del producto a buscar</param>
        public bool ActualizarProductoOfertaPorCodigoProductoSAP(string codProductoSap)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProductoOfertaPersistence persistencia = new ProductoOfertaPersistence(uow);
                return persistencia.ActualizarEstadoCalculoPorCodigoProductoSAP(codProductoSap);
            }
        }

        public bool LaOfertaTieneProductosAfectadosCambioAgrupacionesTipologia(Guid idoferta) 
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProductoOfertaPersistence persistencia = new ProductoOfertaPersistence(uow);
                return persistencia.LaOfertaTieneProductosAfectadosCambioAgrupacionesTipologia(idoferta);
            }
        }

        public void ModificarConfiguracionDestinosNoVisibles(Collection<ProductoOfertaBE> listaProductos, Collection<ProductoOfertaBE> listaProductosAGuardar)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ConfiguracionDestinoOfertaPersistence persistencia = new ConfiguracionDestinoOfertaPersistence(uow);
                persistencia.ModificarConfiguracionDestinosNoVisibles(listaProductos, listaProductosAGuardar);
                uow.Save();
            }
        }

        /// <summary>
        /// Método que guarda la lista de productos de la oferta
        /// </summary>
        /// <param name="oferta">Oferta sobre la que se van a guardar los productos</param>
        /// <param name="listaProductosOferta">Listado de productos a guardar</param>
        public Collection<ProductoOfertaBE> GuardarListaProductosOferta(OfertaBE oferta, Collection<ProductoOfertaBE> listaProductosOferta, IUnitOfWork uow, bool esCopiaEsqueleto = false)
        {
            ProductoOfertaPersistence productoOfertaPersistence = new ProductoOfertaPersistence(uow);
            ProductoPersistence productoPersistence = new ProductoPersistence(uow);
            ConfiguracionDestinoOfertaPersistence confDestinoPersistence = new ConfiguracionDestinoOfertaPersistence(uow);
            ConfiguracionValorAnadidoPersistence confVAPersistence = new ConfiguracionValorAnadidoPersistence(uow);
            ConfiguracionListaPreciosPersistence confListaPreciosPersistence = new ConfiguracionListaPreciosPersistence(uow);
            ConfiguracionPuntoOfertaPersistence confPuntoOfertaPersistence = new ConfiguracionPuntoOfertaPersistence(uow);
            ConfiguracionGradoOfertaPersistence confGradoOfertaPersistence = new ConfiguracionGradoOfertaPersistence(uow);
            CaracteristicaPersistence confCaracteristicaPersistence = new CaracteristicaPersistence(uow);
            ConfiguracionTramoOfertaPersistence confTramoPersistence = new ConfiguracionTramoOfertaPersistence(uow);
            ConfiguracionGruposTramoOfertaPersistence confGruposTramoOfertaPersistence = new ConfiguracionGruposTramoOfertaPersistence(uow);

            //obtengo el listado de productoOfertas
            Collection<ProductoOfertaBE> productosActualEnOferta = productoOfertaPersistence.ObtenerProductosEnOferta(oferta.idOferta);
            
            //Recorro tablas para hacer la diferencia entre los productos actuales a la oferta y los productos candidatos.
            foreach (ProductoOfertaBE objProdOferta in productosActualEnOferta)
            {
                //Check de productos que antes estaban en la BBDD y ahora no aparecen. Se eliminan
                ProductoOfertaBE existe = listaProductosOferta.FirstOrDefault(x => x.idProductoOferta == objProdOferta.idProductoOferta);

                if (existe == null)
                {
                    // Para eliminar un producto oferta, hay que eliminar antes todos los registros de las tablas de configuración                   
                    confDestinoPersistence.EliminarListaConfiguracionDestinoOferta(objProdOferta.idProductoOferta);
                    confTramoPersistence.EliminarListaConfiguracionTramoOferta(objProdOferta.idProductoOferta);
                    confVAPersistence.EliminarListaConfiguracionValorAnadido(objProdOferta.idProductoOferta);
                    confListaPreciosPersistence.EliminarListaConfiguracionListaPrecios(objProdOferta.idProductoOferta);
                    confPuntoOfertaPersistence.EliminarConfiguracionPuntosOferta(objProdOferta.idProductoOferta);
                    confGradoOfertaPersistence.EliminarConfiguracionGradosOferta(objProdOferta.idProductoOferta);
                    confGruposTramoOfertaPersistence.EliminarListaConfiguracionGruposTramoOferta(objProdOferta.idProductoOferta);
                    confCaracteristicaPersistence.EliminarListaCaracteristicasOferta(objProdOferta.idProductoOferta);

                    //Una vez eliminadas las dependencias, se elimina el producto oferta
                    productoOfertaPersistence.EliminarProductoOferta(objProdOferta);
                }
            }

            Collection<ProductoOfertaBE> listaProductos = new Collection<ProductoOfertaBE>();
            
            //Se añaden los ProductoOferta necesarios para completar la operación de añadir. En la interfaz gráfica ya se mantiene la coherencia de los
            //productos en la estructura del nodo, por lo que si los campos idProducto e idProductoOferta son Guid.Empty se tratan de productos a añadir.
            foreach (ProductoOfertaBE poTreeNode in listaProductosOferta)
            {
                //Se busca el producto cuya definición está vigente
                ProductoBE producto = productoPersistence.ObtenerProductoPorIdAnexoProducto(poTreeNode.idAnexoProducto);
                if (producto != null)
                {
                    ProductoOfertaBE productoOferta = productoOfertaPersistence.ObtenerProductoOfertaPorIdProductoIdOfertaIdModalidadNegociacion(producto.idProducto, oferta.idOferta, poTreeNode.idModalidadNegociacion);

                    //solo guardo los nuevos casos, los ya existentes solo los añado.(presupongo que ya estarán bien almacenados)
                    if ((poTreeNode.idProductoOferta.Equals(Guid.Empty) || poTreeNode.idProducto.Equals(Guid.Empty)) && productoOferta == null || esCopiaEsqueleto)
                    {                       
                        //Alta
                        productoOferta = new ProductoOfertaBE()
                        {
                            idProductoOferta = Guid.NewGuid(),
                            idProducto = producto.idProducto,
                            idOferta = oferta.idOferta,
                            NumeroEnvios = poTreeNode.NumeroEnvios,
                            idModalidadNegociacion = poTreeNode.idModalidadNegociacion,
                            StatusProducto = SimuladorResources.StatusEnProcesoPosicion,
                            CodProductoSAP = poTreeNode.CodProductoSAP,
                            Anexo = poTreeNode.Anexo,
                            CodModalidadNegociacion = poTreeNode.CodModalidadNegociacion,
                                Posicion = poTreeNode.Posicion,
                                listaDestinosVisibles = poTreeNode.listaDestinosVisibles,
                                EsReneg = poTreeNode.EsReneg
                        };

                        productoOfertaPersistence.GuardarProductoOferta(productoOferta);

                        listaProductos.Add(productoOferta);

                        //Se crean los registros en las tablas de configuración                        
                        confDestinoPersistence.InsertConfiguracionDestinoOfertaProductoOferta(productoOferta);                     
                        //confTramoPersistence.InsertConfiguracionTramoOfertaProductoOferta(productoOferta);
                        confVAPersistence.InsertConfiguracionValorAnadidoProductoOferta(productoOferta);
                        confListaPreciosPersistence.InsertConfiguracionListaPreciosProductoOferta(productoOferta.idProductoOferta);
                        confPuntoOfertaPersistence.InsertConfiguracionPuntosOferta(productoOferta);
                        confGradoOfertaPersistence.InsertConfiguracionGradosOferta(productoOferta);
                        confCaracteristicaPersistence.InsertConfiguracionCaracteristicaOfertaProductoOferta(productoOferta);
                                          
                    }
                    else if (poTreeNode != null && poTreeNode.CompProducto != null)// && !poTreeNode.CompProducto.coincideEstructura)
                    {
                        ComparacionProductosBE compProducto = poTreeNode.CompProducto;
                        productoOfertaPersistence.GuardarProductoOferta(poTreeNode, true);
                        confDestinoPersistence.ActualizarDefinicionConfiguracionDestinosOferta(poTreeNode.idProductoOferta, poTreeNode.idProducto, compProducto.destinosBorrar, compProducto.destinosAnyadir);
                        confTramoPersistence.ActualizarDefinicionConfiguracionTramoOferta(poTreeNode.idProductoOferta, poTreeNode.listaDestinosVisibles, compProducto);
                        confGruposTramoOfertaPersistence.ActualizarDefinicionConfigGruposTramos(poTreeNode.idProductoOferta, compProducto);
                        confVAPersistence.ActualizarDefinicionConfigValorAnadido(poTreeNode, compProducto);
                        confPuntoOfertaPersistence.ActualizarDefinicionConfigPuntosOferta(poTreeNode.idProductoOferta, poTreeNode.idProducto);
                        confGradoOfertaPersistence.ActualizarDefinicionConfiguracionGradosOferta(poTreeNode.idProductoOferta, poTreeNode.idProducto);
                        poTreeNode.CompProducto = null; //Ya no me hace falta la comparación
                        listaProductos.Add(poTreeNode);
                        //confCaracteristicaPersistence.EliminarListaCaracteristicasOferta(productoOferta.idProductoOferta); NO HACE FALTA
                        // confListaPreciosPersistence.EliminarListaConfiguracionListaPrecios(productoOferta.idProductoOferta); NO HACE FALTA                                            
                    } else if (poTreeNode.listaDestinosVisibles != null && poTreeNode.DestinosConfigurados)
                    {
                        confTramoPersistence.ActualizarConfiguracionDestinosVisibles(poTreeNode.idProductoOferta, poTreeNode.idProducto, poTreeNode.listaDestinosVisibles);
                    }
                }
            }

            return listaProductos;
        }

        public void GuardarDatosConfiguracion(Collection<ProductoOfertaBE> listaProductos, IUnitOfWork uow)
        {
            foreach (ProductoOfertaBE producto in listaProductos)
            {
                //Se crean los registros en las tablas de configuración
                //ConfiguracionDestinoOfertaPersistence confDestinoPersistence = new ConfiguracionDestinoOfertaPersistence(uow);
                //confDestinoPersistence.InsertConfiguracionDestinoOfertaProductoOferta(productoOferta);

                ConfiguracionTramoOfertaPersistence confTramoPersistence = new ConfiguracionTramoOfertaPersistence(uow);
                confTramoPersistence.InsertConfiguracionTramoOfertaProductoOferta(producto);

                //ConfiguracionValorAnadidoPersistence confVAPersistence = new ConfiguracionValorAnadidoPersistence(uow);
                //confVAPersistence.InsertConfiguracionValorAnadidoProductoOferta(productoOferta);

                //ConfiguracionListaPreciosPersistence confListaPreciosPersistence = new ConfiguracionListaPreciosPersistence(uow);
                //confListaPreciosPersistence.InsertConfiguracionListaPreciosProductoOferta(productoOferta.idProductoOferta);

                //ConfiguracionPuntoOfertaPersistence confPuntoOfertaPersistence = new ConfiguracionPuntoOfertaPersistence(uow);
                //confPuntoOfertaPersistence.InsertConfiguracionPuntosOferta(productoOferta);

                //ConfiguracionGradoOfertaPersistence confGradoOfertaPersistence = new ConfiguracionGradoOfertaPersistence(uow);
                //confGradoOfertaPersistence.InsertConfiguracionGradosOferta(productoOferta);
            }
        }


        /// <summary>
        /// Método que inserta los registros necesarios en la tabla ConfiguracionTramoOferta para todos los productos nuevos de la oferta
        /// </summary>
        /// <param name="idOferta"></param>
        public void InicializarConfigTramoOferta(Guid idOferta, IUnitOfWork uow)
        {
                ConfiguracionTramoOfertaPersistence confTramoPersistence = new ConfiguracionTramoOfertaPersistence(uow);
                confTramoPersistence.InicializarConfigTramoOferta(idOferta);              
        }

        /// <summary>
        /// Método que guarda en Base de datos todos los registros de configuración para un listado de productos oferta vacíos
        /// </summary>
        /// <param name="listaProductos">lista de productos oferta</param>
        /// <param name="uow">unit of work</param>
        public void GuardarDatosConfiguracionCompleto(Collection<ProductoOfertaBE> listaProductos, IUnitOfWork uow)
        {
            foreach (ProductoOfertaBE producto in listaProductos)
            {
                //Se crean los registros en las tablas de configuración
                ConfiguracionDestinoOfertaPersistence confDestinoPersistence = new ConfiguracionDestinoOfertaPersistence(uow);
                confDestinoPersistence.InsertConfiguracionDestinoOfertaProductoOferta(producto);

                ConfiguracionTramoOfertaPersistence confTramoPersistence = new ConfiguracionTramoOfertaPersistence(uow);
                confTramoPersistence.InsertConfiguracionTramoOfertaProductoOferta(producto);

                ConfiguracionValorAnadidoPersistence confVAPersistence = new ConfiguracionValorAnadidoPersistence(uow);
                confVAPersistence.InsertConfiguracionValorAnadidoProductoOferta(producto);

                ConfiguracionListaPreciosPersistence confListaPreciosPersistence = new ConfiguracionListaPreciosPersistence(uow);
                confListaPreciosPersistence.InsertConfiguracionListaPreciosProductoOferta(producto.idProductoOferta);

                ConfiguracionPuntoOfertaPersistence confPuntoOfertaPersistence = new ConfiguracionPuntoOfertaPersistence(uow);
                confPuntoOfertaPersistence.InsertConfiguracionPuntosOferta(producto);

                ConfiguracionGradoOfertaPersistence confGradoOfertaPersistence = new ConfiguracionGradoOfertaPersistence(uow);
                confGradoOfertaPersistence.InsertConfiguracionGradosOferta(producto);
            }

        }

        public void CopiarConfiguracionProductoOferta(Guid idProductoOriginal, Guid idProductoCopia, IUnitOfWork uow)
        {
           // //UNDONE Jose Maria COMPROBAR ESTA FUNCIÓN, SI NO ES VÁLIDA BORRARLA DE LA VERSIÓN FINAL
           // throw new NotImplementedException();

           // //foreach (ProductoOfertaBE producto in listaProductos)
           // //{
           //     //Se crean los registros en las tablas de configuración
           // ConfiguracionDestinoOfertaPersistence confDestinoPersistence = new ConfiguracionDestinoOfertaPersistence(uow);            
           // confDestinoPersistence.ObtenerConfiguracionDestinoOferta(idProductoOferta, iddestino);
           // //confDestinoPersistence.InsertConfiguracionDestinoOfertaProductoOferta(producto);

           // ConfiguracionTramoOfertaPersistence confTramoPersistence = new ConfiguracionTramoOfertaPersistence(uow);
           // confTramoPersistence.ObtenerConfiguracionTramoOfertaOptimizado(idProductoOferta, idTramo);
            
           // ConfiguracionValorAnadidoPersistence confVAPersistence = new ConfiguracionValorAnadidoPersistence(uow);
           // //ESTE ES COMPLICADO!!! ANALIZARLO BIEN. 
           // //ES NECESARIO CAMBIARLO?

           // ConfiguracionListaPreciosPersistence confListaPreciosPersistence = new ConfiguracionListaPreciosPersistence(uow);
           // confListaPreciosPersistence.ObtenerConfiguracionListaPrecios(idProducto)
           // //confListaPreciosPersistence.InsertConfiguracionListaPreciosProductoOferta(producto.idProductoOferta);

           // ConfiguracionPuntoOfertaPersistence confPuntoOfertaPersistence = new ConfiguracionPuntoOfertaPersistence(uow);
           // //ES NECESARIO CAMBIARLO?
           //// confPuntoOfertaPersistence.InsertConfiguracionPuntosOferta(producto);

           //     ConfiguracionGradoOfertaPersistence confGradoOfertaPersistence = new ConfiguracionGradoOfertaPersistence(uow);
           // //ES NECESARIO CAMBIARLO?
           //    // confGradoOfertaPersistence.InsertConfiguracionGradosOferta(producto);
           // //}
        }

        /// <summary>
        /// Método que obtiene la configuración de lista de precios para un producto Oferta
        /// </summary>
        /// <param name="idProductoOferta">Identificador del productoOferta</param>        
        /// <returns>Listado de entidades ConfiguracionListaPreciosBE</returns>
        public Collection<ConfiguracionListaPreciosBE> ObtenerListaPrecios(Guid idProductoOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                return new Collection<ConfiguracionListaPreciosBE>();
                //ConfiguracionListaPreciosPersistence persistence = new ConfiguracionListaPreciosPersistence(uow);
                //return persistence.ObtenerListaPrecios(idProductoOferta);
            }
        }

        /// <summary>
        /// Método que obtiene una colección de registros de la tabla configuracionListaPrecios de un conjunto de identificadores
        /// de productos oferta pasados por parámetro
        /// </summary>
        /// <param name="listaIdsProductoOferta">Listado de identificadores de productoOferta</param>
        /// <returns>Colección de entidades ConfiguracionListaPreciosBE</returns>
        public Collection<ConfiguracionListaPreciosBE> ObtenerConfiguracionListaPrecios(Collection<Guid> listaIdsProductoOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ConfiguracionListaPreciosPersistence persistence = new ConfiguracionListaPreciosPersistence(uow);
                return persistence.ObtenerConfiguracionListaPrecios(listaIdsProductoOferta);
            }
        }

        /// <summary>
        /// Método que copia la configuración de lista de precios de un producto oferta
        /// </summary>
        /// <param name="idProductoOriginal">Identificador del producto oferta original</param>
        /// <param name="idProductoDestino">Identificador del producto oferta destino</param>
        /// <returns>Colección de entidades ConfiguracionListaPreciosBE</returns>
        public Collection<ConfiguracionListaPreciosBE> CopiarConfiguracionListaPrecios(Guid idProductoOriginal, Guid idProductoDestino)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ConfiguracionListaPreciosPersistence persistence = new ConfiguracionListaPreciosPersistence(uow);
                return persistence.CopiarConfiguracionListaPrecios(idProductoOriginal, idProductoDestino);
            }
        }

        /// <summary>
        /// Nos indica si existe algún producto de la lista de ofertas que se ha visto afectado por la descarga de agrupaciones y todavía no ha sido actualizado
        /// </summary>
        /// <returns></returns>
        public bool ProductoOfertaConProductoAfectadoDescargaAgrupaciones() 
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProductoOfertaPersistence persistencia = new ProductoOfertaPersistence(uow);
                return persistencia.ProductoOfertaConProductoAfectadoDescargaAgrupaciones();
            }
        }

        #endregion

        #region Método Guardar

        /// <summary>
        /// Método que guarda en base de datos el registro pasado por parámetro
        /// </summary>
        /// <param name="productoOferta">Entidad a guardar</param>
        public void GuardarProductoOferta(ProductoOfertaBE productoOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProductoOfertaPersistence persistence = new ProductoOfertaPersistence(uow);
                persistence.GuardarProductoOferta(productoOferta);

                //Se guarda el contexto
                uow.Save();
            }
        }

        #endregion

        /// <summary>
        /// Verifica que los productos contenidos en la oferta tienen la definicion actualizada. En caso de detectar algun caso desfasado, lo sustituye por el producto con la nueva definicion.
        /// </summary>
        /// <param name="ofertaBE"></param>
        /// <param name="listaProductosOferta"></param>
        /// <param name="sbProductosDesfasados"></param>
        /// <returns></returns>
        public Collection<ProductoOfertaBE> RevisarDefinicionActualProductos(OfertaBE ofertaBE, Collection<ProductoOfertaBE> listaProductosOferta,  bool esCopiaEsqueleto, out StringBuilder sbProductosDesfasados)
        {
            Collection<ProductoOfertaBE> listaProductosActualOferta = new Collection<ProductoOfertaBE>();
            Collection<ProductoOfertaBE> productosReconfigurados= new Collection<ProductoOfertaBE>();
            sbProductosDesfasados = new StringBuilder();

            using (IUnitOfWork uow = new UnitOfWork())
            {
                //declaro el contexto con el que voy a trabajar....
                ProductoOfertaPersistence poPersistence = new ProductoOfertaPersistence(uow);

                foreach (ProductoOfertaBE itemAnterior in listaProductosOferta)
                {
                    //Verifico si el producto asociado tiene la definciion actual
                    if (!poPersistence.EsDefinicionActual(itemAnterior.idProducto))
                    {
                        sbProductosDesfasados.AppendLine(string.Format(CultureInfo.InvariantCulture, "{0}", itemAnterior.CodProductoSAP));

                        Guid idProductoActualizado = new ProductoBL().ObtenerGuidUltimaDefincionProducto(itemAnterior.CodProductoSAP, itemAnterior.Anexo);
                        //JM - Revisar con Manu los destinos visibles
                        ComparacionProductosBE cpObj = CompararDestinosTramosProductos(itemAnterior.idProducto, idProductoActualizado, itemAnterior.listaDestinosVisibles);

                        //Si coincide todo o una parte (Se asigna un nuevo GUID si coinciden los tramos y destinos)
                        if (cpObj.coincideEstructura || !cpObj.sonCompletamenteDiferentes)
                        {
                            itemAnterior.idProducto = idProductoActualizado;
                            itemAnterior.FacturacionNetaProducto = 0;
                            itemAnterior.FacturacionNetaPosicion = 0;
                            itemAnterior.FacturacionBrutaPosicion = 0;
                            itemAnterior.FacturacionBrutaProducto = 0;
                        }
                        //Si no coincide nada (sonCompletamenteDiferentes == true). Se borra todo.
                        else
                        {
                            itemAnterior.idProducto = Guid.Empty;
                            itemAnterior.idProductoOferta = Guid.Empty;
                            productosReconfigurados.Add(itemAnterior);
                        }

                        //Forzamos que haya que recalcular la oferta
                        itemAnterior.EstadoCalculo = 3;
                        itemAnterior.CompProducto = cpObj;
                    }
                    else
                    {
                        if (itemAnterior.idProductoOferta == Guid.Empty)
                        {
                            //Si es un producto nuevo, hay que añadirle la configuración para todos sus tramos
                            productosReconfigurados.Add(itemAnterior);
                        }
                    }

                    listaProductosActualOferta.Add(itemAnterior);
                }

                //en este punto ya tengo los productos como si los fuera a preparar para guardar.

                //Se guardan los productos de la oferta
                listaProductosActualOferta = this.GuardarListaProductosOferta(ofertaBE, listaProductosActualOferta, uow);
                
                //Se guarda el contexto
                uow.Save();

                //Eliminamos este trozo de codigo para resolver una incidencia -> duplicaba tramos
                //Se guardan los datos de configuración (Sólo lo hacemos para los productos para los que se ha reiniciado la configuración.                
                //if (esCopiaEsqueleto)
                //{
                //this.GuardarDatosConfiguracion(listaProductosActualOferta, uow);
                //}

                if (productosReconfigurados.Count > 0)
                    this.InicializarConfigTramoOferta(ofertaBE.idOferta, uow);

            }

            return listaProductosActualOferta;
        }

        /// <summary>
        /// Comparamos el esqueleto de destinos/tramos de dos productos
        /// </summary>
        /// <param name="idProductoAntiguo"></param>
        /// <param name="idProductoNuevo"></param>
        /// <param name="listadoCoincidentes">Listado de arrays destino/tramo coincidentes para ambos productos</param>
        /// <param name="listadoDiferentes">Listado de arrays destino/tramo que no aparecen en el nuevo producto</param>
        /// <returns></returns>
        public Boolean CompararDestinosTramosProductos(Guid idProductoAntiguo, Guid idProductoNuevo, ref List<String[]> listadoCoincidentes, ref List<String[]> listadoDiferentes)
        {
            listadoCoincidentes = new List<String[]>();
            listadoDiferentes = new List<String[]>();
            var esMismaEstructura = false;

            using (IUnitOfWork uow = new UnitOfWork())
            {
                var destinoPersistenceObj = new DestinoPersistence(uow);

                //Obtenemos un listado de tuplas destino/tramo para cada uno de los dos productos (Item 1 Destino, Item 2 Tramo)
                List<String[]> DestinosTramosAntiguo = destinoPersistenceObj.ObtenerListadoDestinosTramosByIdProducto(idProductoAntiguo);
                List<String[]> DestinosTramosNuevo = destinoPersistenceObj.ObtenerListadoDestinosTramosByIdProducto(idProductoNuevo);

                foreach (var tuplaAntigua in DestinosTramosAntiguo)
                {
                    if (DestinosTramosNuevo.Any(t => t[0].Equals(tuplaAntigua[0]) && t[1].Equals(tuplaAntigua[1])))
                    {
                        listadoCoincidentes.Add(tuplaAntigua);
                    }
                    else
                    {
                        listadoDiferentes.Add(tuplaAntigua);
                    }
                }     
                
                //Si ambos productos tienen los mismos destinos
                if (listadoCoincidentes.Count == DestinosTramosNuevo.Count)
                {
                    esMismaEstructura = true;
                }
                //Si los productos tienen destinos diferentes
                else
                {
                    esMismaEstructura = false;                     
                }

                return esMismaEstructura;
            }

        }

        /// <summary>
        /// Compara la definición de dos productos pasados por parámetro
        /// </summary>
        /// <param name="idProductoAntiguo"></param>
        /// <param name="idProductoNuevo"></param>
        /// <returns>Objeto con la información resultante de la comparación</returns>
        public ComparacionProductosBE CompararDestinosTramosProductos(Guid idProductoAntiguo, Guid idProductoNuevo, List<DestinoBE> listaDestinosVisible)
        {
            var cpObj = new ComparacionProductosBE();
            cpObj.idProductoAntiguo = idProductoAntiguo;
            cpObj.idProductoNuevo = idProductoNuevo;
            
            using (IUnitOfWork uow = new UnitOfWork())
            {
                var destinoPersistenceObj = new DestinoPersistence(uow);
                var tramoPersistenceObj = new TramoPersistence(uow);

                #region DESTINOS

                //Obtenemos los destinos de ambos productos
                cpObj.destinosAntiguos = destinoPersistenceObj.ObtenerDestinosTramosVisiblesProducto(idProductoAntiguo, listaDestinosVisible);
                cpObj.destinosNuevos = destinoPersistenceObj.ObtenerDestinosTramosVisiblesProducto(idProductoNuevo, listaDestinosVisible);

                //Obtenemos los destinos a borrar
                cpObj.destinosBorrar = new List<DestinoBE>();
                foreach (DestinoBE item in cpObj.destinosAntiguos)
                {
                    if (!cpObj.destinosNuevos.Any(t => t.CodDestinoSAP.Equals(item.CodDestinoSAP)))
                        cpObj.destinosBorrar.Add(item);
                }

                //Obtenemos los destinos a añadir
                cpObj.destinosAnyadir = new List<DestinoBE>();
                foreach (var item in cpObj.destinosNuevos)
                {
                    if (!cpObj.destinosAntiguos.Any(t => t.CodDestinoSAP.Equals(item.CodDestinoSAP)))
                        cpObj.destinosAnyadir.Add(item);
                }

                #endregion 

                #region TRAMOS

                //Obtenemos los tramos de ambos productos
                var primerDestinoAntiguo = cpObj.destinosAntiguos.FirstOrDefault();
                var primerDestinoNuevo = cpObj.destinosNuevos.FirstOrDefault();

                if (primerDestinoAntiguo != null && primerDestinoNuevo != null)
                {
                    //JCNS. v.1.10.43.2. S0360. TRAMOS. Trae todos los tramos antiguos para borrarlos. Ahora solo trae el primer destino                    
                    //Collection<TramoBE> tramosAntiguos = tramoPersistenceObj.ObtenerTramosByIdDestino(cpObj.destinosAntiguos.First().idDestino);
                    Collection<TramoBE> tramosAntiguos = tramoPersistenceObj.ObtenerTramosPorListaDestinos(cpObj.destinosAntiguos);
                    Collection<TramoBE> tramoNuevos = tramoPersistenceObj.ObtenerTramosByIdDestino(cpObj.destinosNuevos.First().idDestino);

                    //Obtenemos los tramos a borrar
                    cpObj.tramosBorrar = new List<TramoBE>();
                    foreach (TramoBE item in tramosAntiguos)
                    {
                        if (!tramoNuevos.Any(t => t.CodTramo.Equals(item.CodTramo)))
                            cpObj.tramosBorrar.Add(item);
                    }

                    //Obtenemos los tramos a añadir
                    cpObj.tramosAnyadir = new List<TramoBE>();
                    foreach (TramoBE item in tramoNuevos)
                    {
                        if (!tramosAntiguos.Any(t => t.CodTramo.Equals(item.CodTramo)))
                            cpObj.tramosAnyadir.Add(item);
                    }
                }

                #endregion 
               
                return cpObj;
            }

        }

        public bool comprobarsiesreneg(Guid idProductoOferta)
        {
            bool resultado = false;
            using (IUnitOfWork uow = new UnitOfWork())
            {
                var productoOfertaPersistenceObj = new ProductoOfertaPersistence(uow);
                resultado = productoOfertaPersistenceObj.obtenerSiEsRenegociacion(idProductoOferta);
            }
            return resultado;
        } 
    }
}
