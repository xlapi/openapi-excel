using System;
using System.Collections.Generic;

namespace openapi_excel.DI
{
    public class Resolver
    {
        private static Resolver _Instance { get; set; }
        private static object _Lock = new object();
        public static Resolver Instance
        {
            get
            {
                lock (_Lock)
                {
                    if (_Instance == null)
                    {
                        _Instance = new Resolver();
                    }
                    return _Instance;
                }
            }
        }

        public delegate object Creator(Resolver container);

        private readonly Dictionary<string, object> configuration
                       = new Dictionary<string, object>();
        private readonly Dictionary<Type, Creator> typeToCreator
                       = new Dictionary<Type, Creator>();

        public Dictionary<string, object> Configuration
        {
            get { return configuration; }
        }

        public void Register<T>(Creator creator)
        {
            typeToCreator.Add(typeof(T), creator);
        }

        public T Create<T>()
        {
            return (T)typeToCreator[typeof(T)](this);
        }

        public T GetConfiguration<T>(string name)
        {
            return (T)configuration[name];
        }
    }
}
