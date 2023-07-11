using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharepointWorker
{
    public static class ListExtensions
    {
        public static IEnumerable<List<T>> Batch<T>(this List<T> list, int batchSize)
        {
            for (int i = 0; i < list.Count; i += batchSize)
                yield return list.GetRange(i, Math.Min(batchSize, list.Count - i));
        }
    }

}
