using System;
using System.Collections.Generic;
using System.Linq;

namespace Mao.Relatorios.Classes
{
    public static class AppHelpers
    {
        public static TimeSpan SomarPeriodos<TSource>(this IEnumerable<TSource> collection, Func<TSource, TimeSpan> func)
        {
            return new TimeSpan(collection.Sum(item => func(item).Ticks));
        }


        /// <summary>
        /// Obter o mês do ano por extenso
        /// </summary>
        /// <param name="mes">valor do mês</param>
        /// <returns>retorna o mês especificado do ano por extenso</returns>
        public static string MesPorExtenso(int? mes)
        {
            if (mes == null) return null;
            if (String.IsNullOrEmpty(mes.ToString())) return null;
            if (mes == 0) return null;

            string mesExtenso = string.Empty;
            switch (mes)
            {
                case 1: mesExtenso = "Janeiro"; break;
                case 2: mesExtenso = "Fevereiro"; break;
                case 3: mesExtenso = "Março"; break;
                case 4: mesExtenso = "Abril"; break;
                case 5: mesExtenso = "Maio"; break;
                case 6: mesExtenso = "Junho"; break;
                case 7: mesExtenso = "Julho"; break;
                case 8: mesExtenso = "Agosto"; break;
                case 9: mesExtenso = "Setembro"; break;
                case 10: mesExtenso = "Outubro"; break;
                case 11: mesExtenso = "Novembro"; break;
                case 12: mesExtenso = "Dezembro"; break;
                default:
                    break;
            }
            return mesExtenso;
        }

        /// <summary>
        /// Obter o dia da semana por extenso
        /// </summary>
        /// <param name="dayOfWeek">data</param>
        /// <returns>retorna a string do dia da semana da data especificada</returns>
        public static string DiaSemanaPorExtenso(DateTime dayOfWeek)
        {
            if (dayOfWeek == DateTime.MinValue) return null;

            string result = string.Empty;

            switch (dayOfWeek.DayOfWeek)
            {
                case DayOfWeek.Sunday: result = "Domingo"; break;
                case DayOfWeek.Monday: result = "Segunda-Feira"; break;
                case DayOfWeek.Tuesday: result = "Terça-Feira"; break;
                case DayOfWeek.Wednesday: result = "Quarta-Feira"; break;
                case DayOfWeek.Thursday: result = "Quinta-Feira"; break;
                case DayOfWeek.Friday: result = "Sexta-Feira"; break;
                case DayOfWeek.Saturday: result = "Sábado"; break;
            }

            return result;
        }

        /// <summary>
        /// Obter a data da semana por extenso
        /// </summary>
        /// <param name="data">o valor da data</param>
        /// <returns>retorna dia da semana por extenso</returns>
        public static string DataMesPorExtenso(DateTime data)
        {
            if (data == DateTime.MinValue) return null;

            return $"{data.Day} de {MesPorExtenso(data.Month).ToLower()} de {data.Year}";
        }
    }
}
