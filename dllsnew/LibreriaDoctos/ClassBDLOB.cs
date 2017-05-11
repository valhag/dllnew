using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient ;
using System.Data.OleDb;
using Interfaces;

namespace LibreriaDoctos
{
    public class ClassBDLOB : ClassBD
    {

        


        protected override bool  mActualizaDocumento(long liddocum, int adestino, double afolio)
        {
            //if (adestino > 0)
            //    miconexion.mAbrirConexionOrigen(1);
            //else
               // miconexion.mAbrirConexionDestino();
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            long cidfoldig;
            long cidconce;
            long ciddocum01 = 0;
            string cserie = "";
            double ctotal = 0;
            bool lrespuesta = false;

            long ctipocfd = 0;
            int cescfd = 0;
            string lCodigoConcepto;

            
            //
            if (_RegDoctoOrigen.cCodigoConcepto == ""  || string.IsNullOrEmpty (_RegDoctoOrigen.cCodigoConcepto) == true)
                lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();
            else
                lCodigoConcepto = _RegDoctoOrigen.cCodigoConcepto;


            lsql.CommandText = "select cescfd, cverfacele from mgw10006 where ccodigoc01 = '" + lCodigoConcepto + "'";
            lsql.Connection = miconexion._conexion;
            lreader = lsql.ExecuteReader();
            if (lreader.HasRows)
            {
                lreader.Read();
                cescfd = int.Parse(lreader["cescfd"].ToString());
                ctipocfd = long.Parse(lreader["cverfacele"].ToString());
            }

            lreader.Close();
            if (cescfd == 0)
                return true;
            string lcadena = "update mgw10008 set cescfd = 1 where ciddocum01 = " + liddocum;
            OleDbCommand lsql1 = new OleDbCommand(lcadena, miconexion._conexion);
            try
            {
                lsql1.ExecuteNonQuery();
                string lfecha = _RegDoctoOrigen.cFecha.ToString();
                DateTime ldate = DateTime.Parse(lfecha);
                lfecha = ldate.ToString("MM/dd/yyyy");
                lsql.CommandText = "select cidconce01, ciddocum01,cseriepo01,cescfd from mgw10006 where ccodigoc01 = '" + lCodigoConcepto + "'";
                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                _RegDoctoOrigen._RegMovtos.Clear();
                cserie = GetSettingValueFromAppConfigForDLL("SerieFactura").ToString().Trim();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    cidconce = long.Parse(lreader["cidconce01"].ToString());
                    ciddocum01 = long.Parse(lreader["ciddocum01"].ToString());
                    cserie = lreader["cseriepo01"].ToString();

                }
                else
                    cidconce = 1;
                lreader.Close();
                

                lsql.CommandText = "select ctotal from mgw10008 where ciddocum01 = " + liddocum;
                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                _RegDoctoOrigen._RegMovtos.Clear();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    ctotal = double.Parse(lreader["ctotal"].ToString());

                }
                else
                    cidconce = 1;
                lreader.Close();
                double x = afolio;
                
                lsql.CommandText = "select min(cidfoldig) as cidfoldig, min(cfolio) as cfolio, min(cserie) as cserie from mgw10045 where ciddocto = 0 and cestado = 0 and cidcptodoc = " + cidconce + " group by cidcptodoc";
                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                _RegDoctoOrigen._RegMovtos.Clear();
                //string cserie;
                if (lreader.HasRows)
                {
                    lreader.Read();
                    cidfoldig = long.Parse(lreader["cidfoldig"].ToString());
                    //x = double.Parse(lreader["cfolio"].ToString());
                    //cserie = lreader["cserie"].ToString();
                }
                else
                {
                    cidfoldig = 1;
                    //x = 1;

                }
                lreader.Close();
                 
                try
                {

                    if (ctipocfd == 1)
                        lcadena = "update mgw10045 set ciddocto=" + liddocum + ",cestado=1,cfechaemi=ctod('" + lfecha + "')" +
                        " where cidfoldig = " + cidfoldig + " and ciddocto = 0 ";
                    else
                        lcadena = "insert into mgw10045 (cidfoldig,ciddoctode,cidcptodoc,ciddocto,cserie,cfolio,cestado,centregado, cfechaemi,cestrad,ctotal) " +
                                 " values (" + liddocum + "," + ciddocum01 + "," + cidconce + "," + liddocum + ",'" + cserie.Trim() + "'," + x + ",1, 0, ctod('" + lfecha + "'),3," + ctotal + ")";
                    OleDbCommand lsql2 = new OleDbCommand(lcadena, miconexion._conexion);
                    lsql1.CommandText = "SET NULL OFF";
                    lsql1.ExecuteNonQuery();

                    long lcuantos = lsql2.ExecuteNonQuery();
                    if (lcuantos == 0)
                    {
                        lsql.CommandText = "select min(cidfoldig) as cidfoldig, min(cfolio) as cfolio, min(cserie) as cserie from mgw10045 where ciddocto = 0 and cidcptodoc = " + cidconce + " group by cidcptodoc";
                        lsql.Connection = miconexion._conexion;
                        lreader = lsql.ExecuteReader();
                        _RegDoctoOrigen._RegMovtos.Clear();
                        if (lreader.HasRows)
                        {
                            lreader.Read();
                            cidfoldig = long.Parse(lreader["cidfoldig"].ToString());
                            x = double.Parse(lreader["cfolio"].ToString());
                            cserie = lreader["cserie"].ToString();
                        }
                        else
                        {
                            cidfoldig = 1;
                            x = 1;

                        }
                        lreader.Close();

                        lcadena = "update mgw10045 set ciddocto=" + liddocum + ",cestado=1,cfechaemi=ctod('" + lfecha + "')" +
                        " where cidfoldig = " + cidfoldig;
                        lsql2.ExecuteNonQuery();
                    }
                    /*
                    lcadena = "update mgw10008 set cfolio=" + x + ",cseriedo01='" + cserie + "'" +
                    " where ciddocum01 = " + liddocum;
                    lsql2.CommandText = lcadena;
                    lsql2.ExecuteNonQuery();
                     * */
                    lrespuesta = true;
                }
                catch (Exception eeeee)
                {
                    OleDbCommand lsql3 = new OleDbCommand(lcadena, miconexion._conexion);
                    lcadena = "delete from  mgw10008 " +
                    " where ciddocum01 = " + liddocum;
                    lsql3.CommandText = lcadena;
                    lsql3.ExecuteNonQuery();

                    lcadena = "delete from  mgw10010 " +
                    " where ciddocum01 = " + liddocum;
                    lsql3.CommandText = lcadena;
                    lsql3.ExecuteNonQuery();
                    lrespuesta = false;
                }
            }
            catch (Exception eee)
            {
                lrespuesta = true;
            }
            finally
            {
                if (adestino > 0)
                {
                 //   if (adestino == 5)
                 //       miconexion.mCerrarConexionDestino();
            //        miconexion.mCerrarConexionOrigen(1);
                }
                //else
                //    miconexion.mCerrarConexionDestino();

            }
            //.mCerrarConexionDestino ();



            return lrespuesta;

        }

        protected override string mRegresarConsultaMovimientos(string aFuente, string lfolio, int atipo)
        {

            string aEmpresa = GetSettingValueFromAppConfigForDLL("Empresa");

            string aNombre = GetSettingValueFromAppConfigForDLL("Nombre");
            string aFecha = GetSettingValueFromAppConfigForDLL("Fecha");
            string lregresa = "";
            if (atipo == 1)
            {
                lregresa = " select f.[No_] as ccodigop01,f.[Unit Price]  as cprecioc01, " +
                         " case when h.[Currency Code] = '' then [VAT %] else 0 end as cporcent01,  '1' as ccodigoa01, " +
                    //" case when f.[Description] = '' then f.[Item Category Code] else f.[Description] end  as cnombrep01, " + 
                         " case when I.[Item Category Code] <> '' then I.[Item Category Code]  " +
                     "  else  " +
                     "       case when f.[Description] <> '' then f.[Description] " +
                  //   "       else  " +
                  //   "           case when f.[Description 2]  <> '' then f.[Description 2]  " +
                     "           else " +
                     "           f.[No_]  " +
                     "       end " +
                   //  "   end " +
                     " end  as cnombrep01, " +
                         " [Unit of Measure]  as Unidad,  " +
                         " case when h.[Currency Code] = '' then [Amount Including VAT] - Amount else 0 end as cimpuesto1, [Amount] as cneto,  [Line Amount] as ctotal, Quantity as unidades    " +
                         " , '' as ctextoextra2, '' as ctextoextra3, '' as creferen01 , isnull([Search Description],'') as ctextoextra1  " +
                         " from [Operadora Lob$Sales Invoice Line] f  " +
                         " join [Operadora Lob$Sales Invoice Header] h on h.[No_]= f.[Document No_] " +
                         " left join [Item] I on I.[No_] = f.[No_] " +
                " where convert(int,substring([Document No_],4,6)) = " + lfolio;


                // "SELECT C.id_cliente as cliente, o.id_punto_de_venta as cfecha, o.fechamov, o.puntodeventa, p.nombre, o.cantidad, o.importe, pr.nombreproducto, o.id_productos, c.nombrefiscal, c.direccion, c.colonia, m.nombre, e.nombreestados, c.codigopostal " +
                // , c.nombrefiscal
                lregresa = "SELECT o.id_productos as ccodigop01, o.importe as cprecioC01, 0 as cporcent01, '1' as ccodigoa01, pr.nombreproducto as cnombrep01, " + 
                       "'PZA' as Unidad,  0 as Cimpuesto1, o.cantidad * o.importe as cneto, o.cantidad * o.importe as ctotal, o.cantidad as unidades " +
                       " , '' as ctextoextra2, '' as ctextoextra3, '' as creferen01 , '' as ctextoextra1  " +
                       " FROM ((((tbl_operaciones AS o INNER JOIN tbl_puntosdeventa AS p ON o.id_punto_de_venta =p.id_pv02) INNER JOIN tbl_productos01 AS pr ON o.id_productos = pr.id_productos01) INNER JOIN tbl_clientes01 AS C ON p.empresa = C.nombrefiscal) INNER JOIN tbl_municipios AS m ON c.municipio = m.id_municipios) INNER JOIN tbl_estados AS e ON e.id_estados = c.estado " +
                       " WHERE o.id_punto_de_venta <> '' " +
                       " and o.id_punto_de_venta = p.id_pv02  " +
                       " and o.fechamov = #" + aFecha + "#" +
                       " and  " +
                       " p.Empresa ='" + aEmpresa + "'" +
                       " and p.nombre = '" + aNombre + "'" +
                       "ORDER BY o.fechamov DESC ";
                /*
                lregresa = " select f.[No_] as ccodigop01,avg(f.[Unit Price])  as cprecioc01, " +
                         " case when h.[Currency Code] = '' then avg([VAT %]) else 0 end as cporcent01,  '1' as ccodigoa01, " +
                         " case when I.[Item Category Code] <> '' then I.[Item Category Code] " +
                     "  else  " +
                     "       case when  f.[Description] <> '' then f.[Description]" +
                    // "       else  " +
                    // "           case when f.[Description 2]  <> '' then f.[Description 2]  " +
                     "           else " +
                     "           f.[No_]  " +
                     "       end " +
                    //"   end " +
                     " end  as cnombrep01, " +
                         " [Unit of Measure]  as Unidad,  " +
                         " case when h.[Currency Code] = '' then sum([Amount Including VAT] - Amount) else 0 end as cimpuesto1, sum([Amount]) as cneto,  sum([Line Amount]) as ctotal, sum(Quantity) as unidades    " +
                         " , '' as ctextoextra2, '' as ctextoextra3, '' as creferen01 , isnull([Search Description],'') as ctextoextra1  " +
                         " from [Operadora Lob$Sales Invoice Line] f  " +
                         " join [Operadora Lob$Sales Invoice Header] h on h.[No_]= f.[Document No_] " +
                         " left join [Item] I on I.[No_] = f.[No_] " +
                " where convert(int,substring([Document No_],4,6)) = " + lfolio +
                " and LEFT(f.[No_],1) <> '4'" +
                " group by f.[No_], h.[Currency Code], f.[Description],I.[Item Category Code] " +
                " ,  [Unit of Measure], [Search Description] " +
                " union " +
                " select f.[No_] as ccodigop01,f.[Unit Price]  as cprecioc01,   " +
                " case when h.[Currency Code] = '' then [VAT %] else 0 end as cporcent01,  '1' as ccodigoa01,   " +
                " case when I.[Item Category Code] <> '' then I.[Item Category Code]   else          " +
                " case when  f.[Description] <> '' then f.[Description]            " +
                " else            f.[No_]         end  end  as cnombrep01,  " +
                " case when substring(f.[No_],1,3)  = '400' then 'NO APLICA' " + 
                " when substring(f.[No_],1,3)  = '401' then 'NO APLICA'  " + 
                " when substring(f.[No_],1,3)  = '402' then 'NO APLICA' " + 
                " when substring(f.[No_],1,3)  = '403' then 'NO APLICA' " + 
                " when substring(f.[No_],1,3)  = '404' then 'NO APLICA' " + 
                " when substring(f.[No_],1,3)  = '405' then 'NO APLICA' " + 
                " when substring(f.[No_],1,3)  = '406' then 'NO APLICA' " + 
                " when substring(f.[No_],1,3)  = '406' then 'NO APLICA' " +
                " else " + 
                " [Unit of Measure] end as Unidad,    " +
                " case when h.[Currency Code] = '' then [Amount Including VAT] - Amount else 0 end as cimpuesto1,  " +
                " [Amount] as cneto,  [Line Amount] as ctotal, Quantity as unidades     ,  " +
                " '' as ctextoextra2, '' as ctextoextra3, '' as creferen01 ,  " +
                " isnull([Search Description],'') as ctextoextra1    " +
                " from [Operadora Lob$Sales Invoice Line] f    " +
                " join [Operadora Lob$Sales Invoice Header] h on h.[No_]= f.[Document No_]  left join [Item] I on I.[No_] = f.[No_]   " +
                " where convert(int,substring([Document No_],4,6)) = " + lfolio +
                " and LEFT(f.[No_],1) = '4'";

                */

                
            }
            // " join [Operadora Lob$Item] I on I.[No_] = f.[No_] " +
            //[Search Description] as ctextoextra1
            else
                lregresa = " select f.[No_] as ccodigop01,f.[Unit Price]  as cprecioc01, " +
                     " [VAT %] as cporcent01,  '1' as ccodigoa01, " +
                    //      " case when f.[Description] = '' then f.[Item Category Code] else f.[Description] end  as cnombrep01, 
                     " case when f.[Description] <> '' then f.[Description] " +
                     "  else  " +
                     "       case when I.[Item Category Code] <> '' then I.[Item Category Code] " +
                     "       else  " +
                     "           case when f.[Description 2]  <> '' then f.[Description 2]  " +
                     "           else " +
                     "           f.[No_]  " +
                     "       end " +
                     "   end " +
                     " end  as cnombrep01, " +
                     " [Unit of Measure]  as Unidad,  " +
                     " [Amount Including VAT] - Amount as cimpuesto1, [Amount] as cneto,  [Line Amount] as ctotal, Quantity as unidades    " +
                     " , '' as ctextoextra2, '' as ctextoextra3, '' as creferen01 , '' as ctextoextra1  " +
                     " from [Operadora Lob$Sales Cr_Memo Line] f  " +
                     " left join [Item] I on I.[No_] = f.[No_] " +
            " where convert(int,substring([Document No_],4,6)) = " + lfolio + " and LEFT([Document No_],2) = 'NC'";

            return lregresa;
        }

        protected override void mActualizarBarra(double valor)
        {

            double lporcentaje = 0.0D;
            lporcentaje = valor;
            Notificar(lporcentaje);
        }

        protected override Boolean mchecarvalido()
        {
            //if (_RegDoctoOrigen.cFecha > DateTime.Parse("2012/08/01"))
            //    return false;
            //else
                return true;
        }

        

        public override  string mBuscarDoctosArchivo(string aNombreArchivo)
        {
            //System.Data.OleDb.OleDbConnection conn = new OleDbConnection ( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + aNombreArchivo + ";");
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + aNombreArchivo + ";Extended Properties='Excel 12.0 xml;HDR=YES;'");

            conn.Open();
            System.Data.OleDb.OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = conn;
            cmd.CommandText = "SELECT * FROM [LAYOUT$]";
            cmd.ExecuteNonQuery();

            long xxx;
            xxx = 1000000;


            System.Data.OleDb.OleDbDataReader dr;
            dr = cmd.ExecuteReader();
            Boolean noseguir = false;
            _RegDoctos.Clear();
            if (dr.HasRows)
                while (noseguir == false)
                {
                    dr.Read();
                    _RegDoctoOrigen = null;
                    _RegDoctoOrigen = new RegDocto();
                    _RegDoctoOrigen.cFolio = xxx;
                    xxx = xxx + 1;
                    _RegDoctoOrigen.cCodigoConcepto = "";
                    try
                    {
                        _RegDoctoOrigen.cCodigoConcepto = dr["CODIGO CONCEPTO ADMIN"].ToString();
                    }
                    catch
                    { }
                    if (_RegDoctoOrigen.cCodigoConcepto == "")
                    {
                        noseguir = true;
                    }
                    else
                    {
                        _RegDoctoOrigen.cCodigoCliente = dr["RFC"].ToString();
                        _RegDoctoOrigen.cRFC = dr["RFC"].ToString();
                        _RegDoctoOrigen.cFecha = DateTime.Parse(dr["FECHA DE CAPTURA"].ToString());
                        //_RegDoctoOrigen.cFecha = DateTime.Today ;
                        _RegDoctoOrigen.cFecha = _RegDoctoOrigen.cFecha;
                        _RegDoctoOrigen.cRazonSocial = dr["RAZON SOCIAL"].ToString();
                        if (dr["RAZON SOCIAL"].ToString() == string.Empty)
                            _RegDoctoOrigen.cRazonSocial = "Cliente sin razon social";
                        _RegDoctoOrigen.cMoneda = "Pesos";

                        _RegDoctoOrigen.cTipoCambio = 1;
                        _RegDoctoOrigen.cReferencia = dr["REFERENCIA"].ToString().Substring(0, 10);

                        _RegDoctoOrigen.cTextoExtra1 = dr["OBSERVACIONES"].ToString();


                        _RegDoctoOrigen._RegDireccion.cNombreCalle = dr["CALLE"].ToString().Trim();

                        _RegDoctoOrigen._RegDireccion.cNumeroExterior = dr["NUMERO EXTERIOR"].ToString().Trim();
                        _RegDoctoOrigen._RegDireccion.cNumeroInterior = dr["NUMERO INTERIOR"].ToString().Trim();
                        _RegDoctoOrigen._RegDireccion.cColonia = dr["COLONIA"].ToString().Trim();
                        _RegDoctoOrigen._RegDireccion.cCodigoPostal = dr["CODIGO POSTAL"].ToString().PadLeft(5,'0') ;
                        //_RegDoctoOrigen._RegDireccion.cCodigoPostal = _RegDoctoOrigen._RegDireccion.cCodigoPostal.PadLeft(5, "0");
                        _RegDoctoOrigen._RegDireccion.cEstado = dr["ESTADO"].ToString().Trim();
                        _RegDoctoOrigen._RegDireccion.cPais = "MEXICO";
                        _RegDoctoOrigen._RegDireccion.cCiudad = dr["MUNICIPIO"].ToString().Trim();
                        _RegDoctoOrigen._RegDireccion.cEmail = dr["CORREO ELECTRONICO"].ToString().Trim();
                        _RegDoctoOrigen._RegDireccion.cEmail2 = dr["CORREO ELECTRONICO 2"].ToString().Trim();


                        RegMovto lRegmovto = new RegMovto();
                        lRegmovto.cCodigoProducto = dr["CODIGO SERVICIO ADMIN"].ToString();
                        //lRegmovto.cNombreProducto = dr["cnombrep01"].ToString();
                        //lRegmovto.cIdDocto = long.Parse(_RegDoctoOrigen.cIdDocto.ToString());
                        lRegmovto.cPrecio = decimal.Parse(dr["SUBTOTAL"].ToString());

                        lRegmovto.cImpuesto = decimal.Parse(dr["IVA"].ToString());
                        lRegmovto.cPorcent01 = decimal.Round(decimal.Parse(dr["IVA"].ToString()) * 100 / decimal.Parse(dr["SUBTOTAL"].ToString()), 2);
                        lRegmovto.cUnidades = 1;
                        lRegmovto.cTotal = decimal.Parse(dr["IVA"].ToString()) + decimal.Parse(dr["SUBTOTAL"].ToString()); ;
                        lRegmovto.cneto = decimal.Parse(dr["SUBTOTAL"].ToString());
                        lRegmovto.cCodigoAlmacen = "1";
                        lRegmovto.cNombreAlmacen = "1";
                        lRegmovto.cUnidad = "";
                        _RegDoctoOrigen._RegMovtos.Add(lRegmovto);
                        _RegDoctoOrigen.sMensaje = "";
                        _RegDoctos.Add(_RegDoctoOrigen);
                    }
                

                    
                }
            





            return "";
        }


    }
}
