using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Xml;
using System.Data.SqlClient;
using System.Configuration;
namespace ImportToXLS.Controllers
{
    public class HomeController : Controller
    {
        //
        // GET: /Home/

        string query = "";
        string[] querys = { };
        string contenido;
        Boolean vHuboErrror = false;
        Dictionary<string, string> dictionary = new Dictionary<string, string>();
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file)
        {
            DataSet ds = new DataSet();
            if (Request.Files["file"].ContentLength > 0)
            {
                string fileExtension =
                                     System.IO.Path.GetExtension(Request.Files["file"].FileName);

                if (fileExtension == ".xls" || fileExtension == ".xlsx")
                {
                    string fileLocation = Server.MapPath("~/Content/") + Request.Files["file"].FileName;
                    if (System.IO.File.Exists(fileLocation))
                    {

                        System.IO.File.Delete(fileLocation);
                    }
                    Request.Files["file"].SaveAs(fileLocation);
                    string excelConnectionString = string.Empty;
                    excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileLocation + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                    //connection String for xls file format.
                    if (fileExtension == ".xls")
                    {
                        excelConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileLocation + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                    }
                    //connection String for xlsx file format.
                    else if (fileExtension == ".xlsx")
                    {

                        excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileLocation + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                    }
                    //Create Connection to Excel work book and add oledb namespace
                    OleDbConnection excelConnection = new OleDbConnection(excelConnectionString);
                    excelConnection.Open();
                    DataTable dt = new DataTable();

                    dt = excelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    if (dt == null)
                    {
                        return null;
                    }

                    String[] excelSheets = new String[dt.Rows.Count];
                    int t = 0;
                    //excel data saves in temp file here.
                    foreach (DataRow row in dt.Rows)
                    {
                        excelSheets[t] = row["TABLE_NAME"].ToString();
                        t++;
                    }
                    OleDbConnection excelConnection1 = new OleDbConnection(excelConnectionString);


                    string query = string.Format("Select * from [{0}]", excelSheets[0]);
                    using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter(query, excelConnection1))
                    {
                        dataAdapter.Fill(ds);
                        //agrego para que cierre el excel
                        excelConnection.Close();
                    }
                }
                if (fileExtension.ToString().ToLower().Equals(".xml"))
                {
                    string fileLocation = Server.MapPath("~/Content/") + Request.Files["FileUpload"].FileName;
                    if (System.IO.File.Exists(fileLocation))
                    {
                        System.IO.File.Delete(fileLocation);
                    }

                    Request.Files["FileUpload"].SaveAs(fileLocation);
                    XmlTextReader xmlreader = new XmlTextReader(fileLocation);
                    // DataSet ds = new DataSet();
                    ds.ReadXml(xmlreader);
                    xmlreader.Close();
                    
                }


                string conn = ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;
                SqlConnection con = new SqlConnection(conn);


                query =  " delete from Excel_MaestroArticulo ";
                con.Open();
                SqlCommand cmddelete = new SqlCommand(query, con);
                cmddelete.ExecuteNonQuery();
                con.Close();


                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {

                    try
                    {

                      

                        //si hay campo en blanco en el codigo sap, lo descarto.

                        if (ds.Tables[0].Rows[i][5].ToString().Replace("'", "") == "")
                        {

                            continue;
                        }

                        query = "";
                        query = "         INSERT INTO[dbo].[Excel_MaestroArticulo]           ([Descipcion]           ,[UnidadxCaja]           ,[PesoGrs]           ,[CodigoSap]           ,[CodigoBarras]           ,[ProductLine]           ,[GrupoMarcas]           ,[Marca]           ,[Tamaño]           ,[Segmentacion]           ,[GlobalRegional]           ,[NetoCompraCaja]           ,[NetoCompraUnidad]           ,[MargenDist]           ,[NetoVentaCaja]           ,[NetoVentaUnidad]           ,[PrecioSugeridoPublico]           ,[GananciaComerciante]           ,[MargenComerciante]           ,[DescripcionPrecioSugerido],  [Displays]           ,[CantidadBarraDisplay] ) ";
                        query = query + " VALUES           ('" + ds.Tables[0].Rows[i][0].ToString().Replace("'", "") + "','" + ds.Tables[0].Rows[i][1].ToString().Replace(",", ".") + "'," + ds.Tables[0].Rows[i][4].ToString().Replace(",", ".") + "," + ds.Tables[0].Rows[i][5].ToString() + ",'" + ds.Tables[0].Rows[i][6].ToString() + "','" + ds.Tables[0].Rows[i][7].ToString().Replace(",", ".").Replace("'", "") + "','" + ds.Tables[0].Rows[i][8].ToString().Replace(",", ".").Replace("'", "") + "','" + ds.Tables[0].Rows[i][9].ToString().Replace(",", ".").Replace("'", "") + "','" + ds.Tables[0].Rows[i][10].ToString().Replace(",", ".").Replace("'", "").Replace("'", "") + "','" + ds.Tables[0].Rows[i][11].ToString().Replace(",", ".").Replace("'", "").Replace("'", "") + "','" + ds.Tables[0].Rows[i][12].ToString().Replace(",", ".").Replace("'", "") + "'," + ds.Tables[0].Rows[i][13].ToString().Replace(",", ".") + "," + ds.Tables[0].Rows[i][14].ToString().Replace(",", ".") + "," + ds.Tables[0].Rows[i][15].ToString().Replace(",", ".") + "," + ds.Tables[0].Rows[i][16].ToString().Replace(",", ".") + "," + ds.Tables[0].Rows[i][17].ToString().Replace(",", ".") + "," + ds.Tables[0].Rows[i][18].ToString().Replace(",", ".") + "," + ds.Tables[0].Rows[i][19].ToString().Replace(",", ".") + "," + ds.Tables[0].Rows[i][20].ToString().Replace(",", ".") + ",'" + ds.Tables[0].Rows[i][21].ToString().Replace(",", ".").Replace("'", "") + "','" + ds.Tables[0].Rows[i][2].ToString().Replace(",", ".") + "','" + ds.Tables[0].Rows[i][3].ToString().Replace(",", ".") + "')";


                        //agrego al diccionario el articulo
                        dictionary.Add(ds.Tables[0].Rows[i][5].ToString(), query);

                        con.Open();
                        SqlCommand cmd = new SqlCommand(query, con);
                     
                        cmd.ExecuteNonQuery();
                        con.Close();


                        if (vHuboErrror == false)
                        {
                            contenido = "<HTML><head></head><title>Importacion Exitosa sin errores</title><body> <style> tr:nth-child(even) {    background-color: #dddddd;} table {    font-family: arial, sans-serif;    border-collapse: collapse;    width: 100%;} td, th {    border: 1px solid #dddddd;    text - align: left;                        padding: 8px;                    } </style> ";
                            contenido = contenido + " <b> Importacion exitosa  </b> ";
                            contenido = contenido + "<table>  <tr>    <th> Cod. Articulo </th >     <th> SQL </th>  </tr> ";
                        }
                        

                        
                        }

                    catch (Exception e)
                    {
                       vHuboErrror = true;
                       con.Close();

                        


                        // Store keys in a List.


                        contenido = "<HTML><head></head><title>Importacion con Errores</title><body> <style> tr:nth-child(even) {    background-color: #dddddd;} table {    font-family: arial, sans-serif;    border-collapse: collapse;    width: 100%;} td, th {    border: 1px solid #dddddd;    text - align: left;                        padding: 8px;                    } </style> ";
                        contenido = contenido + " <b> Articulos que no se importaron:  </b> ";
                        contenido = contenido + "<table>  <tr>    <th> Cod. Articulo </th >     <th> SQL </th>  </tr> ";

                        foreach (string key in dictionary.Keys)
                        {
                            var hola = key;
                            var hola2 = dictionary[key];


                            contenido = contenido + "<tr> ";
                            contenido = contenido + "<td> " + hola + " </td> ";
                            contenido = contenido + "<td> " + hola2 + "  </td> ";
                            contenido = contenido + "</tr> ";



                        }

                        contenido = contenido + "</table></body></HTML>";






                    }



                }

            }
            

          

            return Content(contenido);
        }



    }
}
