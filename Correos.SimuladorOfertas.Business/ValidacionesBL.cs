using Correos.SimuladorOfertas.Common;
using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.Persistence;
using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;

namespace Correos.SimuladorOfertas.Business
{
    public class ValidacionesBL
    {
        #region Métodos Públicos

        /// <summary>
        /// Valida si existe solapamiento de tramos para un producto con distintas Modalidades de Negocio 
        /// </summary>
        /// <param name="idOferta"></param>
        /// <param name="listaProductosBE"></param>
        /// <returns></returns>
        public ResultBE ValidarSolapamientoTramosEnProducto(Guid idOferta, Collection<ProductoOfertaBE> listaProductosBE)
        {
            ResultBE result = new ResultBE();

            string MsgTramosSolapados = string.Empty;

            try
            {
                using (IUnitOfWork uow = new UnitOfWork())
                {
                    Collection<TramoSAPBE> todostramos = new Collection<TramoSAPBE>();
                    foreach (ProductoOfertaBE producto in listaProductosBE)
                    {
                        //Obtenemos los tramos del producto actual
                        TramoPersistence tramosOferta = new TramoPersistence(uow);
                        Collection<TramoSAPBE> ltramos = tramosOferta.ObtenerTramos(idOferta, producto.idProducto, producto.idModalidadNegociacion, 10);

                        foreach (TramoSAPBE tramoSAPBE in ltramos.Where(x=> x.Distribucion > 0 ))
                        {
                            todostramos.Add(tramoSAPBE);
                        }
                    }

                    var tramosXproductoyDestinoSolapados = todostramos.GroupBy(x => new { x.Anexo, x.CodProductoSAP, x.Tramo, x.Destino })
                                    .Select(g => new { g.Key, Count = g.Count() }).Where(z => z.Count > 1);
                    if (tramosXproductoyDestinoSolapados.Count() > 0)
                    {
                        StringBuilder sb = new StringBuilder();
                        foreach (var x in tramosXproductoyDestinoSolapados)
                        {
                            sb.Append("Anexo: " + x.Key.Anexo + " Producto: " + x.Key.CodProductoSAP + " Tramo: " + x.Key.Tramo + " Destino: " + x.Key.Destino + Environment.NewLine);
                        }
                        MsgTramosSolapados = sb.ToString();
                    }
                }

                if (MsgTramosSolapados.Length > 0)
                {
                    result.Resultado = false;
                    result.TextoError = SimuladorResources.ErrorSincronizarOfertaFalloSolapamientoTramos;
                    StringBuilder sbError = new StringBuilder();
                    sbError.AppendLine(SimuladorResources.ErrorSincronizarOfertaFalloSolapamientoTramos).Append(MsgTramosSolapados);
                    result.TextoError = sbError.ToString();
                    return result;
                }
            }
            catch (Exception ex)
            {
                RegistrarAccionesSimulador.GuardarExcepcion(ex);
                result.Resultado = false;
            }

            return result;
        }


        /// <summary>
        /// Recorre todos los productos de una oferta verificando que los descuentos de los destinos están comprendidos entre [-999,999]
        /// </summary>
        /// <param name="idOferta"></param>
        /// <param name="productosSeleccionados"></param>
        /// <returns></returns>
        public ResultBE ValidarDescuentosMaximosEnProducto(Guid idOferta, Collection<ProductoOfertaBE> productosSeleccionados)
        {
            ResultBE auxRespuesta = new ResultBE();
            foreach (ProductoOfertaBE objProdOferta in productosSeleccionados)
            {
                switch (objProdOferta.CodModalidadNegociacion)
                {
                    case "5DD":
                        //descuento destino
                        auxRespuesta = this.ValidarDescuentosMaximosDescuentoDestino(objProdOferta);
                        if (!auxRespuesta.Resultado)
                        {
                            return auxRespuesta;
                        }
                        break;
                    case "4DT":
                    //descuento tramo
                    case "3DG":
                        //grTramo descuento
                        auxRespuesta = this.ValidarDescuentosMaximosDescuentoTramo(objProdOferta);
                        if (!auxRespuesta.Resultado)
                        {
                            return auxRespuesta;
                        }
                        break;
                    case "2PT":
                    //PC
                    case "1PG":
                        //grTramo PC
                        auxRespuesta = this.ValidarDescuentosMaximosDescuentoPC(objProdOferta);
                        if (!auxRespuesta.Resultado)
                        {
                            return auxRespuesta;
                        }
                        break;
                }
            }
            return auxRespuesta;
        }

        #endregion

        #region Métodos privados

        /// <summary>
        /// valida que todos los destinos del producto cumplen los requisitos para ser sincronizados con SAP
        /// </summary>
        /// <param name="objProdOferta">identificador del producto dentro de la oferta</param>
        /// <returns></returns>
        private ResultBE ValidarDescuentosMaximosDescuentoDestino(ProductoOfertaBE objProdOferta)
        {
            decimal aux0 = 0;
            decimal maxValor = 999;
            decimal minValor = -999;

            ResultBE resultado = new ResultBE();

            if (objProdOferta != null && objProdOferta.PorcentajeFirma != null)
            {
                using (IUnitOfWork uow = new UnitOfWork())
                {
                    //Se obtienen todos los destinos con distribución
                    ConfiguracionDestinoOfertaPersistence persistence = new ConfiguracionDestinoOfertaPersistence(uow);
                    Collection<ConfiguracionDestinoOfertaBE> listaDestinos = persistence.ObtenerListaConfiguracionDestinosOferta(objProdOferta.idProductoOferta);

                    if (listaDestinos.Count.Equals(0))
                    {
                        //es un producto negociado a destino sin destino
                        if (objProdOferta.DescuentoFinal.HasValue &&
                            ((objProdOferta.DescuentoFinal.Value > maxValor) || (objProdOferta.DescuentoFinal.Value < minValor)))
                        {
                            resultado.Resultado = false;
                            resultado.TextoError = string.Format(CultureInfo.InvariantCulture, SimuladorResources.ProductoSinDestinosNoSincronizable, objProdOferta.Descripcion);
                            return resultado;
                        }
                    }
                    else
                    {
                        foreach (ConfiguracionDestinoOfertaBE objDestino in listaDestinos.Where(x => x.Distribucion.HasValue && x.Distribucion.Value > aux0 && x.DescuentoFinal.HasValue))
                        {
                            if ((objDestino.DescuentoFinal.Value > maxValor) || (objDestino.DescuentoFinal.Value < minValor))
                            {
                                resultado.Resultado = false;
                                resultado.TextoError = string.Format(CultureInfo.InvariantCulture, SimuladorResources.ProductoNoSincronizable, objDestino.CodZona, objProdOferta.Anexo, objProdOferta.Descripcion, string.Format("{0:0.00}", objDestino.DescuentoFinal), objProdOferta.ModalidadNegociacion);
                                return resultado;
                            }
                        }
                    }
                }
            }

            return resultado;
        }

        /// <summary>
        /// valida que todos los destinos del producto cumplen los requisitos para ser sincronizados con SAP
        /// </summary>
        /// <param name="objProdOferta">identificador del producto dentro de la oferta</param>
        /// <returns></returns>
        private ResultBE ValidarDescuentosMaximosDescuentoTramo(ProductoOfertaBE objProdOferta)
        {
            ResultBE resultado = new ResultBE();
            double numeroEnvios = 0;
            double.TryParse(objProdOferta.NumeroEnvios.ToString(), out numeroEnvios);
            decimal aux0 = 0;
            decimal maxValor = 999;
            decimal minValor = -999;

            if (objProdOferta != null && objProdOferta.PorcentajeFirma != null)
            {
                using (IUnitOfWork uow = new UnitOfWork())
                {
                    //Se obtienen todos los destinos con distribución
                    ConfiguracionDestinoOfertaPersistence persistence = new ConfiguracionDestinoOfertaPersistence(uow);
                    Collection<DestinoBE> listaDestinos = persistence.ObtenerListaConfiguracionOferta(objProdOferta.idProductoOferta);

                    foreach (DestinoBE objDestino in listaDestinos.Where(x => x.Distribucion.HasValue && x.Distribucion.Value > aux0))
                    {
                        double importeBrutoDestino = 0;
                        double importeNetoDestino = 0;
                        foreach (TramoBE tramoOferta in objDestino.Tramos.Where(x => x.Distribucion.HasValue && x.Distribucion.Value > aux0))
                        {
                            double numeroEnviosTramo = numeroEnvios * Convert.ToDouble(objDestino.Distribucion.Value / 100) * Convert.ToDouble(tramoOferta.Distribucion.Value / 100);
                            double importeBrutoTramo = 0;
                            double importeNetoTramo = 0;

                            importeBrutoTramo = (numeroEnviosTramo * Convert.ToDouble(tramoOferta.TarifaDB));
                            importeNetoTramo = numeroEnviosTramo * (Convert.ToDouble(tramoOferta.TarifaDB) - (Convert.ToDouble(tramoOferta.TarifaDB) * Convert.ToDouble(tramoOferta.DescuentoFinal / 100)));                            

                            //Se actualizan los importes generales del destino
                            importeBrutoDestino += importeBrutoTramo;
                            importeNetoDestino += double.IsNaN(importeNetoTramo) ? 0 : importeNetoTramo;
                        }

                        double descuentoUsuarioDestino = 100 - ((importeNetoDestino * 100) / importeBrutoDestino);

                        if ((descuentoUsuarioDestino < Convert.ToDouble(minValor)) || (descuentoUsuarioDestino > Convert.ToDouble(maxValor)))
                        {
                            resultado.Resultado = false;
                            resultado.TextoError = string.Format(CultureInfo.InvariantCulture, SimuladorResources.ProductoNoSincronizable, objDestino.CodDestinoSAP, objProdOferta.Anexo, objProdOferta.Descripcion, string.Format("{0:0.00}", descuentoUsuarioDestino), objProdOferta.ModalidadNegociacion);
                            return resultado;
                        }
                    }
                }
            }

            return resultado;
        }

        /// <summary>
        /// valida que todos los destinos del producto cumplen los requisitos para ser sincronizados con SAP
        /// </summary>
        /// <param name="objProdOferta">identificador del producto dentro de la oferta</param>
        /// <returns></returns>
        private ResultBE ValidarDescuentosMaximosDescuentoPC(ProductoOfertaBE objProdOferta)
        {
            ResultBE resultado = new ResultBE();
            double numeroEnvios = 0;
            double.TryParse(objProdOferta.NumeroEnvios.ToString(), out numeroEnvios);
            decimal aux0 = 0;
            double aux100 = 1;
            decimal maxValor = 999;
            decimal minValor = -999; 

            if (objProdOferta != null && objProdOferta.PorcentajeFirma != null)
            {
                using (IUnitOfWork uow = new UnitOfWork())
                {
                    //Se obtienen todos los destinos con distribución
                    ConfiguracionDestinoOfertaPersistence persistence = new ConfiguracionDestinoOfertaPersistence(uow);
                    Collection<DestinoBE> listaDestinos = persistence.ObtenerListaConfiguracionOferta(objProdOferta.idProductoOferta);

                    foreach (DestinoBE objDestino in listaDestinos.Where(x => x.Distribucion.HasValue && x.Distribucion.Value > aux0))
                    {
                        double importeBrutoDestino = 0;
                        double importeNetoDestino = 0;

                        foreach (TramoBE tramoOferta in objDestino.Tramos.Where(x => x.Distribucion.HasValue && x.Distribucion.Value > aux0))
                        {                            
                            double numeroEnviosTramo = numeroEnvios * Convert.ToDouble(objDestino.Distribucion.Value / 100) * Convert.ToDouble(tramoOferta.Distribucion.Value / 100);
                            double importeBrutoTramo = 0;
                            double importeNetoTramo = 0;

                            importeBrutoTramo = (numeroEnviosTramo * Convert.ToDouble(tramoOferta.Tarifa));

                            double auxDescuentoTramo = 0;
                            auxDescuentoTramo = ((Convert.ToDouble(tramoOferta.PrecioCierto) * 100) / Convert.ToDouble(tramoOferta.TarifaDB));

                            //Se valida a nivel de tramo que el descuento sea inferior a 999 e superior a -999
                            if (auxDescuentoTramo > Convert.ToDouble(maxValor) || auxDescuentoTramo < Convert.ToDouble(minValor))
                            {
                                resultado.Resultado = false;
                                resultado.TextoError = string.Format(CultureInfo.InvariantCulture, SimuladorResources.ProductoNoSincronizable, objDestino.CodDestinoSAP, objProdOferta.Anexo, objProdOferta.Descripcion, string.Format("{0:0.00}", auxDescuentoTramo), objProdOferta.ModalidadNegociacion);
                                return resultado;
                            }

                            //Como no puedo restar, multiplico por -1 y sumo
                            auxDescuentoTramo = (100 + (auxDescuentoTramo * -1)) / 100;
                            if (auxDescuentoTramo.Equals(aux100))
                            {
                                //Si el descuento aplicado es el 100%, al importe neto del tramo se le asigna el mismo
                                //valor que tenga el importe bruto para dicho tramo
                                importeNetoTramo = importeBrutoTramo;
                            }
                            else
                            {
                                importeNetoTramo = numeroEnviosTramo * (Convert.ToDouble(tramoOferta.TarifaDB) - (Convert.ToDouble(tramoOferta.TarifaDB) * auxDescuentoTramo));
                            }

                            //Se actualizan los importes generales del destino
                            importeBrutoDestino += importeBrutoTramo;
                            importeNetoDestino += double.IsNaN(importeNetoTramo) ? 0 : importeNetoTramo;
                        }

                        double descuentoUsuarioDestino = 100 - ((importeNetoDestino * 100) / importeBrutoDestino);

                        if ((descuentoUsuarioDestino < Convert.ToDouble(minValor)) || (descuentoUsuarioDestino > Convert.ToDouble(maxValor)))
                        {
                            resultado.Resultado = false;
                            resultado.TextoError = string.Format(CultureInfo.InvariantCulture, SimuladorResources.ProductoNoSincronizable, objDestino.CodDestinoSAP, objProdOferta.Anexo, objProdOferta.Descripcion, string.Format("{0:0.00}", descuentoUsuarioDestino), objProdOferta.ModalidadNegociacion);
                            return resultado;
                        }
                    }
                }
            }

            return resultado;
        }


        #endregion

    }
}
