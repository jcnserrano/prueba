using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;

namespace Correos.SimuladorOfertas.Common.Extensions
{
    /// <summary>
    /// Agrupa funciones comunes para trabajar con DateTimes
    /// </summary>
    /// <unit>Common</unit>
    /// <project>Correos.SimuladorOfertas</project>
    /// <company>Correos</company>
    /// <date>01/10/2014</date>
    /// <author>Jose Antonio Pardos</author>
    /// <version>1.0</version>
    /// <history>
    /// 	Versión		|Date		        |Author					        Description
    /// 	1.0			01/10/2014		    Jose Antonio Pardos     		Versión inicial
    /// </history>  
    public static class DateTimeUtil
    {
        #region Fields
        /// <summary>
        /// Calendario sobre el que operarán los métodos implementados en esta clase
        /// </summary>
        static GregorianCalendar _gc = new GregorianCalendar();
        #endregion

        /// <summary>
        /// Método que determina la semana del mes a la que pertenece una fecha
        /// </summary>
        /// <param name="dateTime">Fecha de la cual determinar la funcionalidad</param>
        /// <returns>Número de la semana dentro del mes</returns>
        public static int GetWeekOfMonth(this System.DateTime dateTime)
        {
            System.DateTime first = new System.DateTime(dateTime.Year, dateTime.Month, 1);
            return dateTime.GetWeekOfYear() - first.GetWeekOfYear() + 1;
        }

        /// <summary>
        /// Método que determina la semana del año a la que pertenece una fecha
        /// </summary>
        /// <param name="dateTime">Fecha de la cual determinar la funcionalidad</param>
        /// <returns>número de semana dentro del año</returns>
        public static int GetWeekOfYear(this System.DateTime dateTime)
        {
            return _gc.GetWeekOfYear(dateTime, CalendarWeekRule.FirstDay, DayOfWeek.Sunday);
        }

        /// <summary>
        /// Método que determinará si el mes de una determinada fecha pertenece a la primera semana de dicho mes
        /// </summary>
        /// <param name="dateTime">Fecha de la cual determinar la funcionalidad</param>
        /// <returns>Booleano indicando si es la primera semana del mes</returns>
        public static Boolean IsFirstWeekOfMonth(this System.DateTime dateTime)
        {
            //Comprobamos si la semana de la fecha en el mes es la misma que la de la fecha perteneciente al primer día del mes
            if (dateTime.GetWeekOfMonth() == (new System.DateTime(dateTime.Year, dateTime.Month, 1)).GetWeekOfMonth())
                return true;
            else
                return false;
        }

        /// <summary>
        /// Método que determinará si el mes de una determinada fecha pertenece a la última semana de dicho mes
        /// </summary>
        /// <param name="dateTime">Fecha de la cual determinar la funcionalidad</param>
        /// <returns>Booleano indicando si es la última semana del mes</returns>
        public static Boolean IsLastWeekOfMonth(this System.DateTime dateTime)
        {
            //Comprobamos si la semana de la fecha en el mes es la misma que la de la fecha perteneciente al último día del mes
            if (dateTime.GetWeekOfMonth() ==
                             (new System.DateTime(dateTime.Year, dateTime.Month, System.DateTime.DaysInMonth(dateTime.Year, dateTime.Month)))
                             .GetWeekOfMonth())
                return true;
            else
                return false;

        }

        /// <summary>
        /// Devuelve la diferencia de meses entre dos fechas
        /// </summary>
        /// <param name="date">fecha de inincio</param>
        /// <param name="dateTime">fecha de fin</param>
        /// <returns>devuelve la diferencia de meses o cero si el fecha fin es menor que la fecha de inicio</returns>
        public static Decimal DiffMonths(this System.DateTime date, System.DateTime dateTime)
        {
            if (date <= dateTime)
            {
                decimal numMonth = (((dateTime.Year * 12) + dateTime.Month) - ((date.Year * 12) + date.Month)) + 1;    

                return numMonth;
            }
            else
            {
                return 0;
            }
        }

        /// <summary>
        /// Devuelve si la fecha es el último día del mes
        /// </summary>
        /// <param name="dateTime">fecha</param>
        /// <returns>devuelve true si la fecha corresponde al último día del mes</returns>
        public static Boolean IsLastDayMonth(this System.DateTime dateTime)
        {
            if (dateTime == System.DateTime.MaxValue)
            {
                return true;
            }
            else
            {
                int month = (dateTime.Month < 12) ? dateTime.Month + 1 : 1;

                return (new System.DateTime(dateTime.Year, month, 1).AddDays(-1).Day == dateTime.Day);
            }
        }

        /// <summary>
        /// Devuelve si la fecha es el primer día del mes
        /// </summary>
        /// <param name="dateTime">fecha</param>
        /// <returns>devuelve true si la fecha corresponde al primer día del mes</returns>
        public static Boolean IsFirstDayMonth(this System.DateTime dateTime)
        {
            return dateTime.Day == 1;
        }
    }
}
