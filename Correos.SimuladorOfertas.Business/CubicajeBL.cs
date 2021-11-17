using Correos.SimuladorOfertas.Persistence;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

namespace Correos.SimuladorOfertas.Business
{
    public class CubicajeBL
    {
        #region Obtener
        public Collection<string> ObtenerPosiblesCubicajes()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                CubicajePersistence cubicajePersistence = new CubicajePersistence(uow);
                Collection<string> Cubicajes = cubicajePersistence.ObtenerPosiblesCubicajes();

                string modificarSinvolumetrico = Cubicajes.FirstOrDefault(valor => valor.IndexOf("SIN") >= 0);

                if (!string.IsNullOrEmpty(modificarSinvolumetrico))
                {
                    string sinVolumetrico = "Sin volumétrico";

                    Cubicajes.Remove(modificarSinvolumetrico);
                    Cubicajes.Add(sinVolumetrico);
                }

                return Cubicajes;
            }
        }

        public string ObtenerCubicajePorDefecto()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                CubicajePersistence cubicajePersistence = new CubicajePersistence(uow);

                string CubicajePorDefecto = cubicajePersistence.ObtenerCubicajePorDefecto();

                CubicajePorDefecto = CubicajePorDefecto.IndexOf("SIN") >= 0 ? "Sin volumétrico" : CubicajePorDefecto;

                return CubicajePorDefecto;
            }
        }
        #endregion
    }
}
