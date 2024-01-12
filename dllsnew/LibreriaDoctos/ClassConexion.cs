using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Configuration;
using System.IO;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace LibreriaDoctos
{
    public class ClassConexion
    {
        //public string llaveregistry = "SOFTWARE\\Wow6432Node\\Computación en Acción, SA CV\\AdminPAQ";
        //public string llaveregistry = "SOFTWARE\\Computación en Acción, SA CV\\AdminPAQ";
        public string llaveregistry = "SOFTWARE\\Computación en Acción, SA CV\\CONTPAQ I Facturacion"; 
        public string llaveregistrycomercial = "SOFTWARE\\Computación en Acción, SA CV\\CONTPAQ I COMERCIAL";
        public string sError = "";
        public string aRutaExe = "";
        [DllImport("MGW_SDK.DLL")] static extern int fInicializaSDK();
        [DllImport("MGW_SDK.DLL")] private static extern void fTerminaSDK();
        [DllImport("MGW_SDK.DLL")] private static extern void fCierraEmpresa();
        [DllImport("MGW_SDK.DLL")] static extern int fAbreEmpresa(String aRuta);
        [DllImport("KERNEL32.DLL")] static extern int SetCurrentDirectory(string pPtrDirActual);
        [DllImport("MGW_SDK.DLL")] static extern int fBuscaProducto(String aCodigoProducto);
        [DllImport("MGW_SDK.DLL")] static extern int fInsertaDireccion();
        [DllImport("MGW_SDK.DLL")] static extern int fGuardaDireccion();
        [DllImport("MGW_SDK.DLL")] static extern int fSetDatoDireccion(string aCampo, string aValor);
        [DllImport("MGW_SDK.DLL")] static extern int fBuscaDireccionDocumento(long aIdDocumento, byte aValor);
        [DllImport("MGW_SDK.DLL")] static extern int fEditaDireccion();
        [DllImport("MGW_SDK.DLL")] static extern int fLeeDatoProducto(string aCampo, string aValor, long aLongitud);
        [DllImport("MGW_SDK.DLL")] static extern int fAfectaDocto_Param(string aConcepto, string aSerie, double aFolio, Boolean aAfecta);
        [DllImport("MGW_SDK.DLL")] static extern int fBuscarIdDocumento(int aIdDocumento);
        [DllImport("MGW_SDK.DLL")] static extern int fEditarDocumento();
        [DllImport("MGW_SDK.DLL")] static extern int fBuscarDocumento(string aConcepto, string aFolio, string aSerie);
        [DllImport("MGW_SDK.DLL")] static extern int fBorraDocumento();
        [DllImport("MGW_SDK.DLL")] static extern long fError(long aNumErrror, string aError, long aLen);
        [DllImport("MGW_SDK.DLL")] static extern long fLeeDatoDocumento(string aCampo, ref string aValor, long aLongitud);
        

        public string rutaorigen;
        public string rutadestino;
        public  string _NombreAplicacionCompleto = "InterfazAdmin.exe";
        public  string _NombreAplicacion = "InterfazAdmin";

      /*  public const string _NombreAplicacionCompleto = "Remisiones.exe";
        public const string _NombreAplicacion = "Remisiones";*/

       // public const string _NombreAplicacionCompleto = "Grid.exe";
       // public const string _NombreAplicacion = "Grid";

        public OleDbConnection _conexion ;
        public SqlConnection  _conexion1;
        
        public void borrar()
        { 
            RegistryKey hklp = Registry.LocalMachine;
            hklp = hklp.OpenSubKey(llaveregistry );
            Object obc = hklp.GetValue("DIRECTORIOBASE");
            string lruta1 = obc.ToString();
            string lruta2 = @lruta1;
            SetCurrentDirectory(obc.ToString());
        }
        public OleDbConnection  mAbrirConexionOrigen()
        {
            _conexion = null;
            rutaorigen = GetSettingValueFromAppConfigForDLL( "RutaEmpresaADM");
            if (rutaorigen != "c:\\" && rutaorigen != "LibreriaDoctos.RegEmpresa" && rutaorigen != "Ruta" && rutaorigen != "")
            {
                MessageBox.Show(rutaorigen);
                _conexion = new OleDbConnection();
                _conexion.ConnectionString = "Provider=vfpoledb.1;Data Source=" + rutaorigen;
                _conexion.Open();
            }
            return _conexion;
         
        }
        public OleDbConnection mAbrirConexionOrigen(int a)
        {

            rutaorigen = GetSettingValueFromAppConfigForDLL( "RutaEmpresaSamira");
            //rutaorigen = "c:\\compacw\\empresas\\adtala";
            //rutaorigen = Properties.Settings.Default.RutaEmpresaSamira;
             _conexion =new OleDbConnection();
            _conexion.ConnectionString = "Provider=vfpoledb.1;Data Source=" + rutaorigen ;
            _conexion.Open();
            
            RegistryKey hklp = Registry.LocalMachine;
            hklp = hklp.OpenSubKey(llaveregistry);
            Object obc = hklp.GetValue("DIRECTORIOBASE");
            if (obc == null)
            {
                sError = "No existe instalacion de Adminpaq en este computadora";
                return null;
            }
            SetCurrentDirectory(obc.ToString());
            
            fInicializaSDK();
            fAbreEmpresa(rutaorigen); 
            return _conexion;
        }
        
        public  void mCerrarConexionOrigen()
        {
            _conexion.Close();
        }

        public void mCerrarConexionOrigenComercial()
        {
            try
            {
                if (_conexion1 != null)
                if(_conexion1.State == System.Data.ConnectionState.Open)
                _conexion1.Close();
            }
            catch (Exception eee)
            { }
        }

        public void mCerrarConexionOrigen(int a)
        {
            
            RegistryKey hklp = Registry.LocalMachine;
            hklp = hklp.OpenSubKey(llaveregistry );
            Object obc = hklp.GetValue("DIRECTORIOBASE");
            SetCurrentDirectory(obc.ToString());
            
            _conexion.Close();
            fCierraEmpresa();
            fTerminaSDK();
        }

        
        public void mCerrarConexionDestino()
        {
            _conexion.Close();
        }

        public void mCerrarConexionGlobal()
        {
            _conexion.Close();
        }

        public OleDbConnection mAbrirRutaGlobal(out string amensaje)
        {
            amensaje = "";
            RegistryKey hklp = Registry.LocalMachine;
            

            Object obc = null;
            try
            {
                hklp = hklp.OpenSubKey(llaveregistry);
                obc = hklp.GetValue("DIRECTORIODATOS");
                
                int i = 0;
             //   i = 1 / i;
            }
            catch (Exception eeee)
            {
                amensaje = eeee.Message;
             //   obc = "c:\\compacw\\empresas";
              //  amensaje = "";
            }
                //amensaje = obc.ToString ();
            if (obc == null)
            {
                //amensaje = "No existe instalacion de Adminpaq en este computadora";
                return null;
            }
            _conexion = new OleDbConnection();
            _conexion.ConnectionString = "Provider=vfpoledb.1;Data Source=" + obc.ToString();
            //MessageBox.Show(_conexion.ConnectionString);
            // _conexion.ConnectionString = "Provider=vfpoledb;Data Source=" + obc.ToString();

            //_conexion.ConnectionString = "Provider=vfpoledb.1;Data Source=" + "\\toshiba-pc" + asc(92) +  "empresas";
            try
            {
                _conexion.Open();
            }
            catch (Exception eeee)
            {
                amensaje = eeee.Message;
            }
            return _conexion ;

        }
        public OleDbConnection mAbrirConexionAccess(out string msg)
        {
            msg = "";
                string rutaaccess = GetSettingValueFromAppConfigForDLL("RutaAccess");
                //msg = rutaaccess;
            _conexion = new OleDbConnection();
            _conexion.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + rutaaccess    + ";User Id=admin;Password=";
            //_conexion.Open();

            
            return _conexion;
 
        }

        public OleDbConnection mAbrirConexionDestinoComercial(int a)
        {
            rutadestino = GetSettingValueFromAppConfigForDLL("RutaEmpresaADM");


            RegistryKey hklp = Registry.LocalMachine;
            hklp = hklp.OpenSubKey(llaveregistry);
            Object obc = hklp.GetValue("DIRECTORIOBASE");
            string lruta1 = obc.ToString();
            string lruta2 = @lruta1;
            SetCurrentDirectory(obc.ToString());

            long lret;
            try
            {
                //fTerminaSDK();
                lret = fInicializaSDK();
            }
            catch (Exception eeeee)
            {
                fTerminaSDK();
                lret = fInicializaSDK();
            }
            lret = fAbreEmpresa(rutadestino);
            //fCierraEmpresa();
            //fTerminaSDK();
            return _conexion;

        }


        public SqlConnection mAbrirConexionSQLOrigen()
        {
            //            rutadestino = "c:\\compacw\\empresas\\adtala2";
            string rutadestino = GetSettingValueFromAppConfigForDLL("RutaEmpresaADM");

            string sempresa = rutadestino.Substring(rutadestino.LastIndexOf("\\") + 1);

            string server = GetSettingValueFromAppConfigForDLL("serverOrigen");
            string user = GetSettingValueFromAppConfigForDLL("userOrigen");
            string pwd = GetSettingValueFromAppConfigForDLL("passwordO");
            sempresa = GetSettingValueFromAppConfigForDLL("databaseOrigen");
            //string lruta3 = obc.ToString();
            string lruta4 = @rutadestino;
            _conexion1 = new SqlConnection();
            string Cadenaconexion1 = "data source =" + server + ";initial catalog = " + sempresa + ";user id = " + user + "; password = " + pwd + ";";
            _conexion1.ConnectionString = Cadenaconexion1;
            _conexion1.Open();

            return _conexion1;

        }

        

        public SqlConnection mAbrirConexionComercial(RegConexion datosconexion, bool incluyesdk)
        {
            //            rutadestino = "c:\\compacw\\empresas\\adtala2";
            string rutadestino = datosconexion.database;

            string sempresa = rutadestino.Substring(rutadestino.LastIndexOf("\\") + 1);

            string server = datosconexion.server;
            string user = datosconexion.usuario;
            string pwd = datosconexion.ps;
            //sempresa = GetSettingValueFromAppConfigForDLL("empresa");
            //string lruta3 = obc.ToString();
            string lruta4 = @rutadestino;
            _conexion1 = new SqlConnection();
            string Cadenaconexion1 = "data source =" + server + ";initial catalog = " + sempresa + ";user id = " + user + "; password = " + pwd + ";";
            _conexion1.ConnectionString = Cadenaconexion1;
            _conexion1.Open();

            if (incluyesdk == true)
            {

                RegistryKey hklp = Registry.LocalMachine;
                hklp = hklp.OpenSubKey(llaveregistrycomercial);
                Object obc = hklp.GetValue("DIRECTORIOBASE");
                string lruta1 = obc.ToString();
                string lruta2 = @lruta1;
                SetCurrentDirectory(obc.ToString());

                long lret;
                try
                {
                    //fTerminaSDK();
                    // lret = fInicializaSDK();
                       
           
                }
                catch (Exception eeeee)
                {
                    fTerminaSDK();
                    //  lret = fInicializaSDK();
                }
                //lret = fAbreEmpresa(rutadestino);
                //fCierraEmpresa();
                //fTerminaSDK();
            }
            return _conexion1;

        }


        public SqlConnection mAbrirConexionComercial(string cadenaconexion, bool incluyesdk)
        {
            //            rutadestino = "c:\\compacw\\empresas\\adtala2";
            
            //sempresa = GetSettingValueFromAppConfigForDLL("empresa");
            //string lruta3 = obc.ToString();
            string lruta4 = @rutadestino;
            _conexion1 = new SqlConnection();
            string Cadenaconexion1 = cadenaconexion;
            _conexion1.ConnectionString = Cadenaconexion1;
            _conexion1.Open();

            if (incluyesdk == true)
            {

                RegistryKey hklp = Registry.LocalMachine;
                hklp = hklp.OpenSubKey(llaveregistrycomercial);
                Object obc = hklp.GetValue("DIRECTORIOBASE");
                string lruta1 = obc.ToString();
                string lruta2 = @lruta1;
                SetCurrentDirectory(obc.ToString());

                long lret;
                try
                {
                    //fTerminaSDK();
                    // lret = fInicializaSDK();


                }
                catch (Exception eeeee)
                {
                    fTerminaSDK();
                    //  lret = fInicializaSDK();
                }
                //lret = fAbreEmpresa(rutadestino);
                //fCierraEmpresa();
                //fTerminaSDK();
            }
            return _conexion1;

        }

        public SqlConnection mAbrirConexionComercial(bool incluyesdk)
        {
            //            rutadestino = "c:\\compacw\\empresas\\adtala2";
            string  rutadestino = GetSettingValueFromAppConfigForDLL("RutaEmpresaADM");

            string sempresa = rutadestino.Substring(rutadestino.LastIndexOf("\\") + 1);

            string server = GetSettingValueFromAppConfigForDLL("server");
            string user = GetSettingValueFromAppConfigForDLL("user");
            string pwd = GetSettingValueFromAppConfigForDLL("password");
            //sempresa = GetSettingValueFromAppConfigForDLL("empresa");
            //string lruta3 = obc.ToString();
            string lruta4 = @rutadestino;
            _conexion1 = new SqlConnection();
            string Cadenaconexion1 = "data source =" + server + ";initial catalog = " + sempresa + ";user id = " + user + "; password = " + pwd + ";";
            _conexion1.ConnectionString = Cadenaconexion1;
            _conexion1.Open();

            if (incluyesdk == true)
            {

                RegistryKey hklp = Registry.LocalMachine;
                hklp = hklp.OpenSubKey(llaveregistrycomercial);
                Object obc = hklp.GetValue("DIRECTORIOBASE");
                string lruta1 = obc.ToString();
                string lruta2 = @lruta1;
                SetCurrentDirectory(obc.ToString());

                long lret;
                try
                {
                    //fTerminaSDK();
                   // lret = fInicializaSDK();
                }
                catch (Exception eeeee)
                {
                    fTerminaSDK();
                  //  lret = fInicializaSDK();
                }
                //lret = fAbreEmpresa(rutadestino);
                //fCierraEmpresa();
                //fTerminaSDK();
            }
            return _conexion1;

        }

        public OleDbConnection  mAbrirConexionDestino(int a)
        {
//            rutadestino = "c:\\compacw\\empresas\\adtala2";
            rutadestino = GetSettingValueFromAppConfigForDLL("RutaEmpresaADM");
            //string lruta3 = obc.ToString();
            string lruta4 = @rutadestino;
             _conexion =new OleDbConnection();
            _conexion.ConnectionString = "Provider=vfpoledb.1;Data Source=" + rutadestino ;
            _conexion.Open();
            
            RegistryKey hklp = Registry.LocalMachine;
            hklp = hklp.OpenSubKey(llaveregistry );
            Object obc = hklp.GetValue("DIRECTORIOBASE");
            string lruta1 = obc.ToString();
            string lruta2 = @lruta1;
            SetCurrentDirectory(obc.ToString());
            
            long lret;
            try
            {
                //fTerminaSDK();
                lret = fInicializaSDK();
            }
            catch (Exception eeeee)
            { fTerminaSDK();
            lret = fInicializaSDK();
            }
            lret = fAbreEmpresa(rutadestino);
            //fCierraEmpresa();
            //fTerminaSDK();
            return _conexion;
         
        }
        public OleDbConnection mAbrirConexionDestino()
        {
            _conexion = null;
            rutadestino = GetSettingValueFromAppConfigForDLL("RutaEmpresaADM");
            if (rutadestino != "c:\\" && rutadestino != "LibreriaDoctos.RegEmpresa")
            {
                _conexion = new OleDbConnection();
                _conexion.ConnectionString = "Provider=vfpoledb.1;Data Source=" + rutadestino;
                _conexion.Open();
            }
            return _conexion;

        }
        private string GetSettingValueFromAppConfigForDLL(string aNombreSetting)
        {
            string lrutadminpaq = Directory.GetCurrentDirectory();
            Directory.SetCurrentDirectory(lrutadminpaq);
            if (Directory.GetCurrentDirectory() != aRutaExe)
                Directory.SetCurrentDirectory(aRutaExe);

            string value ="";
            System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal);
            ClientSettingsSection userSettingsSection = (ClientSettingsSection)config.SectionGroups["userSettings"].Sections[_NombreAplicacion + ".Properties.Settings"];
            //SettingElement elemToDelete = null;

            foreach (SettingElement connStr in userSettingsSection.Settings)
            {
                if (connStr.Name == aNombreSetting)
                {
                    value = connStr.Value.ValueXml.InnerText;
                    break;
                }
            }
            if (lrutadminpaq != aRutaExe)
                Directory.SetCurrentDirectory(lrutadminpaq);
            
            return value;
        }
        


    }
}
