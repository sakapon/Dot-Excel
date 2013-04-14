using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Keiho.Apps.DotExcel
{
    public static class EnumerableHelper
    {
        public static IEnumerable<TSource> ForEach<TSource>(this IEnumerable<TSource> source, Action<TSource> action)
        {
            if (source == null) throw new ArgumentNullException("source");
            if (action == null) throw new ArgumentNullException("action");

            return source.Select(e =>
            {
                action(e);
                return e;
            });
        }

        public static void Execute<TSource>(this IEnumerable<TSource> source)
        {
            if (source == null) throw new ArgumentNullException("source");

            foreach (var item in source)
            {
            }
        }

        public static IEnumerable<IGrouping<TKey, TSource>> GroupBySequentially<TSource, TKey>(this IEnumerable<TSource> source, Func<TSource, TKey> keySelector)
        {
            if (source == null) throw new ArgumentNullException("source");
            if (keySelector == null) throw new ArgumentNullException("keySelector");

            var queue = new Queue<TSource>();
            var currentKey = default(TKey);
            var comparer = EqualityComparer<TKey>.Default;

            foreach (var item in source)
            {
                var key = keySelector(item);

                if (!comparer.Equals(key, currentKey))
                {
                    if (queue.Count != 0)
                    {
                        yield return new Grouping<TKey, TSource>(currentKey, queue.ToArray());
                        queue.Clear();
                    }
                    currentKey = key;
                }
                queue.Enqueue(item);
            }

            if (queue.Count != 0)
            {
                yield return new Grouping<TKey, TSource>(currentKey, queue.ToArray());
                queue.Clear();
            }
        }
    }

    public class Grouping<TKey, TElement> : IGrouping<TKey, TElement>
    {
        public TKey Key { get; private set; }
        protected IEnumerable<TElement> Values { get; private set; }

        public Grouping(TKey key, IEnumerable<TElement> values)
        {
            Key = key;
            Values = values;
        }

        public IEnumerator<TElement> GetEnumerator()
        {
            return Values.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
