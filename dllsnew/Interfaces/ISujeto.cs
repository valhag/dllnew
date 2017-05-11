using System;
using System.Collections.Generic;
//using System.Linq;
using System.Text;

namespace Interfaces
{
    public interface ISujeto
    {
        void Registrar(IObservador obs);
        void Registrar(); 
        void Notificar();
        void Notificar(double lavance);
        // test 
    }
}
