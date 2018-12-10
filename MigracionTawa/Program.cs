using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;

namespace MigracionTawa {
    static class Program {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        static void Main() {
            try {
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[] { new MainTasks() };
                ServiceBase.Run(ServicesToRun);
            } catch (Exception e) {
                System.IO.File.WriteAllLines(@"C:\lel.txt", new string[] { e.ToString() });
            }
        }
    }
}
