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
    }
}
