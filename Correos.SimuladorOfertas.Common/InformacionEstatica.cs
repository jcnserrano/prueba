using System.Collections.ObjectModel;
using Correos.SimuladorOfertas.DTOs;

namespace Correos.SimuladorOfertas.Common
{
    public static class InformacionEstatica
    {
        #region Variables

        private static Collection<ProductoOfertaBE> _listaProductosOfertaBE;

        #endregion

        #region Propiedades

        /// <summary>
        /// Lista de ProductoOfertaBE
        /// </summary>
        public static Collection<ProductoOfertaBE> ListaProductosOfertaBE
        {
            get
            {
                if (_listaProductosOfertaBE == null)
                    _listaProductosOfertaBE = new Collection<ProductoOfertaBE>();

                return _listaProductosOfertaBE;
            }
            set { _listaProductosOfertaBE = value; }
        }

        #endregion
    }
}
