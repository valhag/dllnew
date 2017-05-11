using System;
using System.Collections.Generic;
using System.Text;

namespace LibreriaDoctos
{
    public class ClassRNLOB :  ClassRN
    {

        public ClassRNLOB()
        { 
            lbd = new ClassBDLOB();
        }
        

     // public ClassBDLOB lbd = new ClassBDLOB();

        public override  string mBuscarDoctoFlex(string aFolio, int aTipo, bool aRevisar)
        {
            return lbd.mBuscarDoctoAccess( aRevisar);
        }

        public override string mBuscarDoctosArchivo(string aArchivo)
        {
            return lbd.mBuscarDoctosArchivo(aArchivo );
        }

        public override string mBuscarDoctos(long aFolioinicial, long afoliofinal, int aTipo, bool aRevisar)
        {
                return lbd.mBuscarDoctos(aFolioinicial, afoliofinal, aTipo, aRevisar);
        }
    }
}
