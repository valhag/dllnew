using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient ;
using System.Data.OleDb;
using Interfaces;


namespace LibreriaDoctos
{
    public class ClassBDORA : ClassBD
    {

        public ClassBDORA()
        {
          //  miconexion = new ClassConexion();
            _con = null;
           _con = new SqlConnection  ();
        }

        protected override string mRegresarConsultaMovimientos(string aFuente, string lfolio, int atipo)
        {
            string lregresa = "";
            if (atipo == 1)
            {
                lregresa = " select f.[No_] as ccodigop01,f.[Unit Price]  as cprecioc01, " +
                         " case when h.[Currency Code] = '' then [VAT %] else 0 end as cporcent01,  '1' as ccodigoa01, " +
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

                lregresa = " SELECT p.prod_id AS CCODIGOP01, M.MOV_PRECIO AS cprecioc01, mov_iva as cporcent01, " +
" P.PROD_DESC AS cnombrep01, " +
" u.udm_desc as unidad, " +
" case " +
    " when mov_iva = 0 then 0  " +
    " when mov_iva <> 0 then (mov_iva/100)*m.mov_precio " +
  " end as cimpuesto1, " +
  " m.mov_cantidad * m.mov_precio as cneto,  " +
" case " +
    " when mov_iva = 0 then m.mov_cantidad * m.mov_precio " +
    " when mov_iva <> 0 then (m.mov_cantidad * m.mov_precio) + ((mov_iva/100)*m.mov_precio) " +
  " end as ctotal,  " +
  " m.mov_cantidad as unidades, " +
  " ' ' as ctextoextra2, ' ' as ctextoextra3, p.prod_desc_comercial as creferen01 , ' ' as ctextoextra1, " +
  " TO_CHAR(DL.LOTE_NUMERO)  AS LOTE,  " +
  " TO_CHAR(DL.LOTE_CADUCIDAD, 'YYYY/MM/DD') AS CADUCIDAD " +
" FROM DOC_MOVIMIENTO M " +
" JOIN DOC_DOCUMENTO D ON D.DOC_ID = M.DOC_ID " +
" JOIN cat_producto P ON p.prod_id = m.prod_id " +
" join cat_unidad_medida u on u.udm_id = m.udm_id " +
" join doc_mov_lote DML on M.MOV_ID = DML.MOVM_ID " +
" JOIN doc_lote DL on DML.LOTE_ID = DL.LOTE_ID " +
" where D.DOC_FOLIO = " + lfolio;



                

                
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

        protected override string mConsultaEncabezado (int aTipo,string aFolio)
        {
            string lregresa = "";
            switch (aTipo)
            {
                case 1:
                    lregresa = " select  isnull(f.[Bill-to Customer No_],'') as cliente, case when [Posting Date] = '1899-12-30 00:00:00.000' then GETDATE() else [Posting Date] end as cfecha,  " +
                     " convert(int,substring([No_],4,6)) as cfolio, isnull([Bill-to Address],'') as cnombrec01, '0' as cnumeroe01, '' as cnumeroi01,  " +
                     " CASE WHEN isnull([Bill-to Address 2],'') = '' THEN 'SIN COLONIA' ELSE [Bill-to Address 2]  END as ccolonia, isnull([Bill-to City],'') as cciudad, isnull([Bill-to County],'') as cestado, e.Name as cpais  " + 
                     " , isnull([VAT Registration No_],'') as crfc, isnull([Bill-to Name],'') as crazonso01, isnull([Bill-to Post Code],0) as ccodigop01,  " + 
      //               " isnull([Currency Code],'Pesos') as moneda, " +
                     " case when [Currency Code] = 'USD' then 'USD' else 'Pesos' end as moneda, " +
                     " case when [Currency Factor] = 0 then 1 else round(1/ [Currency Factor],4,10) end as tipocambio, 0 as condpago, '(Ninguno)' as agente, [No_] as creferen01  " +
                     " , f.[Ship-to Name] as textoextra1 " +
                     " from [Operadora Lob$Sales Invoice Header] f left join Country_Region e  " + 
                     " on e.Code = f.[Bill-to Country_Region Code]  " + 
                    " WHERE convert(int,substring([No_],4,6)) = " + aFolio;

                    lregresa = " select CL.CTE_CLAVE AS cliente, d.doc_fecha as fecha, d.doc_folio as cfolio, " +
                    " d.doc_fis_calle as cnombrec01, isnull(d.doc_fis_num_ext,0) as cnumeroe01, isnull(d.doc_fis_num,0) as cnumeroi01, " +
                    " d.doc_fis_colonia as ccolonia, c.cd_desc, e.edo_desc as cestado, p.pais_desc as cpais, " +
                    " cl.cte_rfc as crfc, cl.cte_razon_social as crazonso01, d.doc_fis_cp as ccodigop01, " +
                    " m.mon_desc as moneda, d.doc_tipo_cambio as tipocambio, ' ' as textoextra1 " +
                    " from doc_documento d join cat_ciudad c on d.cd_id_fis =c.cd_id " +
                    " join cat_estado e on c.edo_id = e.edo_id " +
                    " join cat_pais p on p.pais_id = e.pais_id " +
                    " join cat_cliente cl on cl.cte_id = d.cte_id " +
                    " join cat_moneda m on m.mon_id = d.mon_id " +
                    " where DOC_FOLIO = 59381";

                    break;
                case 2:
                    lregresa = " select  isnull(f.[Bill-to Customer No_],'') as cliente, case when [Posting Date] = '1899-12-30 00:00:00.000' then GETDATE() else [Posting Date] end as cfecha,  " +
                     " convert(int,right([No_],6)) as cfolio, isnull([Bill-to Address],'') as cnombrec01, '0' as cnumeroe01, '' as cnumeroi01,  " +
                     " isnull([Bill-to Address 2],'') as ccolonia, isnull([Bill-to City],'') as cciudad, isnull([Bill-to County],'') as cestado, 'Mexico' as cpais  " +
                     " , isnull([VAT Registration No_],'') as crfc, isnull([Bill-to Name],'') as crazonso01, isnull([Bill-to Post Code],0) as ccodigop01,  " +
                     " 'moneda' = 'Pesos' ,   " +
                     " '1' as tipocambio, 0 as condpago, '(Ninguno)' as agente, [No_] as creferen01  " +
                     " , '' as textoextra1 " +
                     " from [Operadora Lob$Sales Cr_Memo Header] f left join Country_Region e  " +
                     " on e.Code = f.[Bill-to Country_Region Code]  " +
                    " WHERE convert(int,right([No_],6))  = " + aFolio +
                    " and left([No_],2)  = 'NC' and " +
                    " LEFT([External Document No_],5) = 'NCVDV'";
                    break;
                case 4:
                    lregresa = " select  isnull(f.[Bill-to Customer No_],'') as cliente, case when [Posting Date] = '1899-12-30 00:00:00.000' then GETDATE() else [Posting Date] end as cfecha,  " +
                     " convert(int,right([No_],6)) as cfolio, isnull([Bill-to Address],'') as cnombrec01, '0' as cnumeroe01, '' as cnumeroi01,  " +
                     " isnull([Bill-to Address 2],'') as ccolonia, isnull([Bill-to City],'') as cciudad, isnull([Bill-to County],'') as cestado, 'Mexico' as cpais  " +
                     " , isnull([VAT Registration No_],'') as crfc, isnull([Bill-to Name],'') as crazonso01, isnull([Bill-to Post Code],0) as ccodigop01,  " +
                     " 'moneda' = 'Pesos' ,   " +
                     " '1' as tipocambio, 0 as condpago, '(Ninguno)' as agente, [No_] as creferen01  " +
                     ", importes.importe, importes.impuestos " +
                     " from [Operadora Lob$Sales Cr_Memo Header] f left join Country_Region e  " +
                     " on e.Code = f.[Bill-to Country_Region Code]  " +
                     " join " +
                     "( " +
                     "select [Document No_], sum(l.Amount) as importe, sum(l.[Amount Including VAT] ) - sum(l.Amount)  as impuestos " +
                     "from  [Operadora Lob$Sales Cr_Memo Line] l " +
                     "group by [Document No_] " +
                     ") as importes on importes.[Document No_] = f.[No_] " +
                    " WHERE convert(int,right([No_],6))  = " + aFolio +
                    " and left([No_],2)  = 'NC' and " +
                    " LEFT([External Document No_],5) = 'NCVBO'";

                    break;

                    // WHERE convert(int,right(l.[Document No_],6)) = " + aFolio + 

                case 5:
                    // globales


                    break;


            }
            return lregresa;

        }
        protected override Boolean mchecarvalido()
        {
            //if (_RegDoctoOrigen.cFecha > DateTime.Parse("2012/08/01"))
            //    return false;
            //else
                return true;
        }

        public override string mBuscarDoctos(long aFolio, long afoliofinal, int aTipo, Boolean aRevisar)
        {

            string lrespuesta = "";
            _RegDoctos.Clear();
            for (long i = aFolio; i <= afoliofinal; i++)
            {
                RegDocto lDocto = new RegDocto();
                _RegDoctoOrigen = null;
                _RegDoctoOrigen = new RegDocto();
                lrespuesta = mBuscarDoctoFlex(i.ToString(), aTipo, aRevisar);
                if (lrespuesta == string.Empty)
                {
                    _RegDoctoOrigen.sMensaje = "";
                    _RegDoctoOrigen.cFolio = i;
                }
                else
                {
                    _RegDoctoOrigen.sMensaje = lrespuesta;
                    _RegDoctoOrigen.cFolio = i;
                }
                lDocto = _RegDoctoOrigen;
                _RegDoctos.Add(lDocto);
            }
            return lrespuesta;
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


        public  string mLlenarDocto1(OleDbDataReader aReader, int atipo, string aFolio, string aFuente)
        {
            string lrespuesta = "";
            string lfolio = "0";
            if (atipo == 1 || atipo == 2)
            {
                lfolio = aReader["cfolio"].ToString();
                _RegDoctoOrigen.cFolio = long.Parse(lfolio);
            }
            if (aReader["cliente"].ToString() == string.Empty)
                return "Falta Codigo de cliente en documento " + aFolio;
            else
                _RegDoctoOrigen.cCodigoCliente = aReader["cliente"].ToString();

            _RegDoctoOrigen.cFecha = DateTime.Parse(aReader["cfecha"].ToString());
            if (mchecarvalido() == false)
                return "";



            //_RegDoctoOrigen.cFolio = long.Parse (aReader["cfolio"].ToString()) ;
            if (aReader["cRFC"].ToString() == string.Empty)
                return "Cliente sin RFC en documento " + aFolio;
            else
                if (!(aReader["cRFC"].ToString().Length == 12 || aReader["cRFC"].ToString().Length == 13))
                    return "El RFC tiene una longitud incorrecta en el documento " + aFolio;
                else
                    _RegDoctoOrigen.cRFC = aReader["cRFC"].ToString();


            if (atipo == 1)
            {
                _RegDoctoOrigen.cAgente = aReader["Agente"].ToString();
                _RegDoctoOrigen.cCond = aReader["condpago"].ToString();

            }
            if (aReader["cRazonso01"].ToString() == string.Empty)
                return "Cliente sin Razon Social en documento " + aFolio;
            else
                _RegDoctoOrigen.cRazonSocial = aReader["cRazonso01"].ToString();

            //IsDBNull(
            //aReader["cTextoExtra1"].isnull
            if (!aReader.IsDBNull(18))
                _RegDoctoOrigen.cTextoExtra1 = aReader[18].ToString();



            // UNA modificacion que aplica para flexo es que los datos del cliente se toman de adminpaq
            lrespuesta = mModificaDatosCliente();
            //lrespuesta = mModificaDatosClienteFlexo();
            if (lrespuesta != string.Empty)
                return lrespuesta;


            _RegDoctoOrigen.cMoneda = aReader["Moneda"].ToString();
            _RegDoctoOrigen.cTipoCambio = decimal.Parse(aReader["TipoCambio"].ToString());

            if (atipo != 1)
                _RegDoctoOrigen.cReferencia = aReader["cReferen01"].ToString();
            else
                _RegDoctoOrigen.cReferencia = aReader["cReferen01"].ToString();



            if (aReader["cnombrec01"].ToString().Trim() == string.Empty)
                _RegDoctoOrigen._RegDireccion.cNombreCalle = "Ninguna";
            else
                _RegDoctoOrigen._RegDireccion.cNombreCalle = aReader["cnombrec01"].ToString().Trim();

            _RegDoctoOrigen._RegDireccion.cNumeroExterior = aReader["cnumeroe01"].ToString().Trim();
            _RegDoctoOrigen._RegDireccion.cNumeroInterior = aReader["cnumeroi01"].ToString().Trim();
            _RegDoctoOrigen._RegDireccion.cColonia = aReader["ccolonia"].ToString().Trim();
            _RegDoctoOrigen._RegDireccion.cCodigoPostal = aReader["ccodigop01"].ToString().Trim();
            _RegDoctoOrigen._RegDireccion.cEstado = aReader["cestado"].ToString().Trim();
            _RegDoctoOrigen._RegDireccion.cPais = aReader["cpais"].ToString().Trim();
            _RegDoctoOrigen._RegDireccion.cCiudad = aReader["cciudad"].ToString().Trim();
            if (atipo == 3 || atipo == 4)
            {
                _RegDoctoOrigen.cNeto = double.Parse(aReader["importe"].ToString());
                _RegDoctoOrigen.cImpuestos = double.Parse(aReader["impuestos"].ToString().Trim());
            }


            SqlCommand lsql = new SqlCommand();
            SqlDataReader lreader;

            lsql.CommandText = mRegresarConsultaMovimientos(aFuente, lfolio, atipo);


            //lsql.Connection = (SqlConnection)_con;
            //aReader.Close();
            lreader = lsql.ExecuteReader();
            _RegDoctoOrigen._RegMovtos.Clear();
            if (lreader.HasRows)
            {
                while (lreader.Read())
                {
                    RegMovto lRegmovto = new RegMovto();
                    lRegmovto.cCodigoProducto = lreader["ccodigop01"].ToString();
                    lRegmovto.cNombreProducto = lreader["cnombrep01"].ToString();
                    lRegmovto.cIdDocto = long.Parse(_RegDoctoOrigen.cIdDocto.ToString());
                    lRegmovto.cPrecio = decimal.Parse(lreader["cprecioc01"].ToString());

                    lRegmovto.cImpuesto = decimal.Parse(lreader["cimpuesto1"].ToString());
                    lRegmovto.cPorcent01 = decimal.Parse(lreader["cPorcent01"].ToString());
                    if (aFuente != "Mercado")
                    {
                        lRegmovto.cUnidades = decimal.Parse(lreader["unidades"].ToString());
                        lRegmovto.cTotal = decimal.Parse(lreader["cTotal"].ToString());
                        lRegmovto.cneto = decimal.Parse(lreader["cneto"].ToString());
                        lRegmovto.cReferencia = lreader["creferen01"].ToString();
                        lRegmovto.ctextoextra1 = lreader["ctextoextra1"].ToString();
                        lRegmovto.ctextoextra2 = lreader["ctextoextra2"].ToString();
                        lRegmovto.ctextoextra3 = lreader["ctextoextra3"].ToString();

                    }
                    else
                        lRegmovto.cUnidades = decimal.Parse(lreader["cunidades"].ToString());
                    lRegmovto.cCodigoAlmacen = lreader["ccodigoa01"].ToString();
                    lRegmovto.cNombreAlmacen = lreader["ccodigoa01"].ToString();
                    lRegmovto.cUnidad = lreader["unidad"].ToString();
                    _RegDoctoOrigen._RegMovtos.Add(lRegmovto);
                }

            }
            else
            {

            }
            lreader.Close();
            return lrespuesta;
            //miconexion.mCerrarConexionOrigen(); 
        }


        public override bool mValidarConexionIntell(string aServidor, string aBd, string ausu, string apwd)
        {
            string Cadenaconexion = "data source =" + aServidor + ";initial catalog =" + aBd + ";user id = " + ausu + "; password = " + apwd + ";";
            String sConnectionString = "Provider=MSDAORA.1;User ID=" + ausu + ";password=" + apwd + ";" +
     " Data Source=" + aServidor + ";Persist Security Info=False";
            OleDbConnection _con = new OleDbConnection(sConnectionString);
            _con.ConnectionString = sConnectionString ;
            try
            {
                _con.Open();
                // si se conecto grabar los datos en el cnf
                _con.Close();
                return true;
            }
            catch (Exception ee)
            {
                return false;
            }
        }
        public string mBuscarDoctoFlex(string aFolio, int aTipo, Boolean aRevisar)
        {
//            OleDbConnection myConnection = new OleDbConnection(sConnectionString);
  //          OleDbCommand myCommand = new OleDbCommand(mySelectQuery, myConnection);


            SqlCommand lcmd = new SqlCommand();
            SqlDataReader lreader;
            string lRespuesta = "";
            if (aTipo == 0)
                return lRespuesta;
            _con.Open();
            lcmd.Connection = _con;

            lcmd.CommandText = mConsultaEncabezado(aTipo, aFolio);



            try
            {
                lreader = lcmd.ExecuteReader();
            }
            catch (Exception e)
            {
                lRespuesta = e.Message;
                _con.Close();
                return lRespuesta;
            }
            if (lreader.HasRows)
            {
                if (aRevisar == true)
                {
                    if (mBuscarADM(aFolio, aTipo) == true)
                    {
                        _con.Close();
                        return "Documento Ya existe en Adminpaq"; // documento ya existe
                    }
                }
                lreader.Read();
                //lRespuesta = mLlenarDocto(lreader, aTipo, aFolio, "Flex");
            }

            else
            {
                lRespuesta = "Documento no Encontrado"; // documento no encontrado
            }

            _con.Close();
            return lRespuesta;
        }


    }
}
