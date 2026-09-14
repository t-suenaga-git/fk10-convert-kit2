using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Converter10
{
    public class SafeDictionary<TKey, TValue> : Dictionary<TKey, TValue>
    {
        public new TValue this[TKey key]
        {
            get
            {
                return TryGetValue(key, out var value) ? value : default;
            }
            set
            {
                base[key] = value;
            }
        }
    }
}
