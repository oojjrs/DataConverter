using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace DataConverter
{
    public static class Utility
    {
        public static IEnumerable<Type> GetTypes(Assembly assembly, string @namespace)
        {
            var types = assembly.GetTypes();

            return types.Where(t => t.Namespace == @namespace);
        }
    }
}
