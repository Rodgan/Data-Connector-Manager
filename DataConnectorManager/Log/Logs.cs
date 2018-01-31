using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataConnectorManager
{
    public class Logs
    {
        private static List<Exception> Exceptions = new List<Exception>();

        public static void AddException(Exception excp)
        {
            Exceptions.Add(excp);
        }
        public static void ClearLog()
        {
            Exceptions.Clear();
        }
        public static Exception GetLastException()
        {
            return Exceptions.Count > 0 ? Exceptions.Last() : new Exception("There are no exceptions");
        }
        
    }
}
