using Correos.SimuladorOfertas.Common.Extensions;
using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.InOutLight;
using Correos.SimuladorOfertas.Persistence;
using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace Correos.SimuladorOfertas.Business
{
    public class CoeficienteBL
    {
        #region Métodos Públicos

        /// <summary>
        /// Método que comprueba si para todos los productos y valores añadidos se tiene su coeficiente asociado
        /// </summary>
        /// <param name="usuario">Identificador del usuario</param>
        /// <param name="password">Contraseña del usuario</param>
        /// <param name="uow">Objeto transaccional</param>
        public void ComprobarCoeficientesProductos(string usuario, string password, IUnitOfWork uow)
        {
            CoeficientePersistence persistence = new CoeficientePersistence(uow);
            Collection<CoeficienteBE> listaProductosVA = persistence.ObtenerProductosVASinCoeficiente(usuario);

            if (listaProductosVA != null && listaProductosVA.Count > 0)
            {
                if (SSOHelper.Instance.LogarConSSO)
                {
                    CommunicatorLight conectorSAP = new CommunicatorLight(SSOHelper.Instance.Usuario, password);
                    //SSOHelper.Instance.ActualizarCookiePortal();
                    SSOHelper.Instance.InicializarWSLight(conectorSAP.Cliente);
                    this.GuardarCoeficientes(conectorSAP.ZCObtenerCoeficienteRfc(listaProductosVA, usuario), usuario, uow);
                    SSOHelper.Instance.LimpiarWSLight();
                }
                else
                {
                    CommunicatorLight conectorSAP = new CommunicatorLight(usuario, password);
                    this.GuardarCoeficientes(conectorSAP.ZCObtenerCoeficienteRfc(listaProductosVA, usuario), usuario, uow);
                }
            }
        }

        /// <summary>
        /// Método que obtiene el coeficiente para un producto y usuario
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <param name="usuario">Identificador del usuario</param>
        /// <returns>Coeficiente para dicho producto y usuario</returns>
        public static double ObtenerValorCoeficienteProducto(Guid idProducto, string usuario)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                CoeficientePersistence persistence = new CoeficientePersistence(uow);
                CoeficienteBE coeficiente = persistence.ObtenerCoeficienteProducto(idProducto, usuario);

                double auxCoeficiente = 0;

                if (coeficiente != null && double.TryParse(coeficiente.Coeficiente.ToString(), out auxCoeficiente))
                {
                    return auxCoeficiente;
                }
                else
                {
                    return 0;
                }
            }
        }

        /// <summary>
        /// Método que obtiene el coeficiente para un valor añadido y usuario
        /// </summary>
        /// <param name="idValorAnadidoProducto">Identificador del producto</param>
        /// <param name="usuario">Identificador del usuario</param>
        /// <returns>Coeficiente para dicho valor añadido y usuario</returns>
        public static double ObtenerValorCoeficienteVA(Guid idValorAnadidoProducto, string usuario)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                CoeficientePersistence persistence = new CoeficientePersistence(uow);
                CoeficienteBE coeficiente = persistence.ObtenerCoeficienteVA(idValorAnadidoProducto, usuario);

                double auxCoeficiente = 0;

                if (coeficiente != null && double.TryParse(coeficiente.Coeficiente.ToString(), out auxCoeficiente))
                {
                    return auxCoeficiente;
                }
                else
                {
                    return 0;
                }
            }
        }

        /// <summary>
        /// Método que elimina los coeficientesProducto de un producto y un usuario
        /// </summary>
        /// <param name="idProducto">Identificador de producto</param>
        /// <param name="usuario">Código de usuario</param>
        public void EliminarCoeficientesProducto(Guid idProducto, string usuario, IUnitOfWork uow)
        {
            CoeficientePersistence persistence = new CoeficientePersistence(uow);
            persistence.EliminarCoeficientesProducto(idProducto, usuario);
        }

        /// <summary>
        /// Método que elimina los coeficientes de los valores añadidos de un producto y usuario
        /// </summary>
        /// <param name="idProducto">Identificador de producto</param>
        /// <param name="usuario">Código de usuario</param>
        public void EliminarCoeficientesVA(Guid idProducto, string usuario, IUnitOfWork uow)
        {
            CoeficientePersistence persistence = new CoeficientePersistence(uow);
            persistence.EliminarCoeficientesVA(idProducto, usuario);
        }

        #endregion

        #region Métodos Privados

        /// <summary>
        /// Método que guarda en base de datos toda la información de Coeficientes, tanto de productos como de valores añadidos
        /// </summary>
        /// <param name="listaCoeficientes">Lista de coeficientes</param>
        /// <param name="listaProductosBE">Lista de ProductoBE</param>
        /// <param name="usuario">Identificador del usuario</param>
        /// <param name="uow">Objeto transaccional</param>
        private void GuardarCoeficientes(Collection<CoeficienteBE> listaCoeficientes, string usuario, IUnitOfWork uow)
        {
            #region Productos

            // ------------------------
            // Se manejan los productos
            // ------------------------
            Collection<CoeficienteBE> listaCoeficientesProductos = listaCoeficientes.Where(x => string.IsNullOrEmpty(x.VA)).ToList<CoeficienteBE>().ToCollection<CoeficienteBE>();

            foreach (CoeficienteBE coeficienteProducto in listaCoeficientesProductos)
            {
                // Se gestiona el producto
                ProductoPersistence productoPersistence = new ProductoPersistence(uow);
                Collection<ProductoBE> productosDB = productoPersistence.ObtenerProductosByCodProductoSAP(coeficienteProducto.Producto);

                // El producto no puede ser nulo, no obstante se comprueba
                foreach (ProductoBE productoDB in productosDB)
                {
                    // Se gestiona la relación entre producto y usuario
                    CoeficientePersistence coeficientePersistence = new CoeficientePersistence(uow);
                    CoeficienteBE coeficienteProductoDB = coeficientePersistence.ObtenerCoeficienteProducto(productoDB.idProducto, usuario);

                    if (coeficienteProductoDB == null)
                    {
                        // Si no existe la relación, se crea
                        coeficienteProductoDB = new CoeficienteBE()
                        {
                            idCoeficienteProducto = Guid.NewGuid(),
                            Coeficiente = coeficienteProducto.Coeficiente,
                            idProducto = productoDB.idProducto,
                            Usuario = usuario
                        };
                    }

                    coeficientePersistence.InsertUpdateCoeficienteProducto(coeficienteProductoDB);
                }
            }

            #endregion

            #region Valores Añadidos

            // -------------------------------
            // Se manejan los valores añadidos
            // -------------------------------
            Collection<CoeficienteBE> listaCoeficienteVA = listaCoeficientes.Except(listaCoeficientesProductos).ToList<CoeficienteBE>().ToCollection<CoeficienteBE>();

            foreach (CoeficienteBE coeficienteVA in listaCoeficienteVA)
            {
                // Se gestiona el valor añadido
                ProductoPersistence productoPersistence = new ProductoPersistence(uow);

                Collection<ProductoBE> productosVADB = productoPersistence.ObtenerProductosByCodProductoSAP(coeficienteVA.Producto);

                // El producto no puede ser nulo, no obstante se comprueba 
                foreach (ProductoBE productoDB in productosVADB)
                {
                    ValorAnadidoProductoPersistence vapPersistence = new ValorAnadidoProductoPersistence(uow);
                    Guid idValorAnadidoProducto = vapPersistence.ObtenerIdValorAnadidoProductoByIdProductoCodVASAP(productoDB.idProducto, coeficienteVA.VA);

                    if (!idValorAnadidoProducto.Equals(Guid.Empty))
                    {
                        // Se gestiona la relación entre valor añadido y usuario
                        CoeficientePersistence coeficientePersistence = new CoeficientePersistence(uow);
                        CoeficienteBE coeficienteVADB = coeficientePersistence.ObtenerCoeficienteVA(idValorAnadidoProducto, usuario);

                        if (coeficienteVADB == null)
                        {
                            // Si no existe la relación, se crea
                            coeficienteVADB = new CoeficienteBE()
                            {
                                idCoeficienteVA = Guid.NewGuid(),
                                Coeficiente = coeficienteVA.Coeficiente,
                                idValorAnadidoProducto = idValorAnadidoProducto,
                                Usuario = usuario
                            };
                        }

                        coeficientePersistence.InsertUpdateCoeficienteVA(coeficienteVADB);
                    }
                }
            }

            #endregion
        }

        #endregion
    }
}
