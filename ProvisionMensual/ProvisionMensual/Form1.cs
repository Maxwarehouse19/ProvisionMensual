using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Configuration;
using System.IO;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
//using System.Data;
using ClosedXML.Excel;
using System.Globalization;

namespace ProvisionMensual
{
    public partial class Form1 : Form
    {
        // variables para uso del programa
        // -------------------------------
        string RutaArchivoGeneracion            ="";
        string RutaArchivoVentasNoFacturado    ="";
        string SalesOrderNumber                 ="";
        string HoldCode                         ="";
        string TotalSales                       ="";
        string SalesSku                         ="";
        string SalesCategoryAtTimeOfSale        ="";
        string UomCode                          ="";
        string UomQuantity                      ="";
        string SalesStatus                      ="";
        string SalesOrderDate                   ="";
        string SalesChannelName                 ="";
        string CustomerName                     ="";
        string FulfillmentSku                   ="";
        string FulfillmentChannelName           ="";
        string FulfillmentChannelType           ="";
        string LinkedFulfillmentChannelName     ="";
        string FulfillmentLocationName          ="";
        string FulfillmentOrderNumber           ="";
        string Quantity                         ="";
        string Sku                              ="";
        string Title                            ="";
        string TotalCost                        ="";
        string Commission                       ="";
        string InventoryCost                    ="";
        string UnitCost                         ="";
        string ServiceCost                      ="";
        string EstimatedShippingCost            ="";
        string ShippingCost                     ="";
        string ShippingPrice                    ="";
        string OverheadCost                     ="";
        string PackageCost                      ="";
        string ProfitLoss                       ="";
        string Carrier                          ="";
        string ShippingServiceLevel             ="";
        string ShippedByUser                    ="";
        string ShippingWeight                   ="";
        string Length                           ="";
        string varWidth                         ="";
        string varHeight                        ="";
        string Weight                           ="";
        string StateRegion                      ="";
        string TrackingNum                      ="";
        string MfrName                          ="";
        string PricingRule                      = "";
        string ActualShippingCost               = "";
        string ActualShipping                   = "";
        string ShippingCostDifference           = "";
        int counter = 0;
        string line;
        int cantidad = 0;
        bool Encontro = false;
        string PalabraCompleta = "";
        int ContadorProgreso = 0;
        string ArchivoLog = "";
        string ReporteLog = "";
        
        bool FlgSihayFedex = false;
        bool FlgSihayUSPS = false;
        bool FlgSihayUPS = false;
        bool FlgSihayMI15 = false;
        bool FlgSihayAmazon = false;
        bool EncontroRegistro = false;
        bool FlgSihayEndicia = false;
        bool FlgSihayPITNEYBOWES = false;
        bool FlgSihayEstimatedDeliveryDate = false;

        string ArchivosSecundarios = ConfigurationManager.AppSettings["CarpetaArchivosSecundarios"];
        string pathString = "";
        System.IO.StreamReader fileRead = null;

        int contador = 1;
        string BodyExcel = "<html>";

        List<PedidoFedex> listaPedido = new List<PedidoFedex>();

        List<PedidoUSPS> listaPedidoUSPS = new List<PedidoUSPS>();

        List<PedidoUPS> listaPedidoUPS = new List<PedidoUPS>();

        public Form1()
        {
            try
            {
                InitializeComponent();
                //EjecutaProceso();
            }
            catch (Exception exp)
            {
                MessageBox.Show("Error: " + exp.Message);
            }
        }

        private static string GetConnectionString(string file,string Tipo)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            string extension = file.Split('.').Last();

            if (extension.ToUpper() == "XLS"  )
            {
                //Excel 2003 and Older
                props["Provider"] = "Microsoft.Jet.OLEDB.4.0";

                if (Tipo== "MASTER")
                    props["Extended Properties"] = "Excel 8.0";
                else
                    props["Extended Properties"] = "Excel 8.0";
            }
            else if (extension.ToUpper() == "XLSX")
            {
                //Excel 2007, 2010, 2012, 2013
                props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";

                if (Tipo == "MASTER")
                    props["Extended Properties"] = "Excel 12.0 XML";
                else
                    props["Extended Properties"] = "Excel 12.0 XML";
            }
            else
                throw new Exception(string.Format("error file: {0}", file));

            props["Data Source"] = file;

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }

        private static DataSet GetDataSetFromExcelFile(string file,string connectionString)
        {
            DataSet ds = new DataSet();

            //string connectionString = GetConnectionString(file,);

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                System.Data.DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data
                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();

                    if (!sheetName.EndsWith("$"))
                        continue;

                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.TableName = sheetName;

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);

                    ds.Tables.Add(dt);
                }

                cmd = null;
                conn.Close();
            }

            return ds;
        }

        private static DataSet GetDataSetFromExcelFileDetalle(string file, string connectionString)
        {
            DataSet ds = new DataSet();

            //string connectionString = GetConnectionString(file,);

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                System.Data.DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data
                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();

                    if (!sheetName.Contains(" "))
                    {

                        if (!sheetName.EndsWith("$"))
                            continue;
                    }
                    else {
                        if (sheetName.Contains("FilterDatabase"))
                            continue;
                    }

                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.TableName = sheetName;

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);

                    ds.Tables.Add(dt);
                }

                cmd = null;
                conn.Close();
            }

            return ds;
        }


        //  inserta fila a reporte
        // -----------------------
        private void InsertaEncabezadoReporte()
        {
            BodyExcel = "<html>";

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            RutaArchivoGeneracion = pathString +@"\"  ;//ConfigurationManager.AppSettings["RutaArchivoGeneracion"];
            RutaArchivoGeneracion = RutaArchivoGeneracion + "ReporteOutput" + DateTime.Now.ToString("yyyyMMddTHHmmss") + ".xls";
            using (System.IO.StreamWriter FileExcel = new System.IO.StreamWriter(RutaArchivoGeneracion, true))
            {

                BodyExcel = BodyExcel + "<body>";
                BodyExcel = BodyExcel + "<table>";
                BodyExcel = BodyExcel + @"<tr bgcolor= ""#CA2229"" style=""color:#ffffff"">";
                // Archivo DW
                // ----------
                BodyExcel = BodyExcel + "<td> SalesOrderNumber</td>";
                BodyExcel = BodyExcel + "<td> TotalSales</td>";
                BodyExcel = BodyExcel + "<td> HoldCode</td>";
                BodyExcel = BodyExcel + "<td> SalesSku</td>";
                BodyExcel = BodyExcel + "<td> SalesCategoryAtTimeOfSale</td>";
                BodyExcel = BodyExcel + "<td> UomCode</td>";
                BodyExcel = BodyExcel + "<td> UomQuantity</td>";
                BodyExcel = BodyExcel + "<td> SalesStatus</td>";
                BodyExcel = BodyExcel + "<td> SalesOrderDate</td>";
                BodyExcel = BodyExcel + "<td> SalesChannelName</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentSku</td>";
                BodyExcel = BodyExcel + "<td> CustomerName</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentChannelName</td>";
                BodyExcel = BodyExcel + "<td> LinkedFulfillmentChannelName</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentLocationName</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentChannelType</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentOrderNumber</td>";
                BodyExcel = BodyExcel + "<td> Quantity</td>";
                BodyExcel = BodyExcel + "<td> Sku</td>";
                BodyExcel = BodyExcel + "<td> Title</td>";
                BodyExcel = BodyExcel + "<td> TotalCost</td>";
                BodyExcel = BodyExcel + "<td> Commission</td>";
                BodyExcel = BodyExcel + "<td> InventoryCost</td>";
                BodyExcel = BodyExcel + "<td> UnitCost</td>";
                BodyExcel = BodyExcel + "<td> ServiceCost</td>";
                BodyExcel = BodyExcel + "<td> EstimatedShippingCost</td>";
                BodyExcel = BodyExcel + "<td> ShippingCost</td>";
                BodyExcel = BodyExcel + "<td> ShippingPrice</td>";
                BodyExcel = BodyExcel + "<td> OverheadCost</td>";
                BodyExcel = BodyExcel + "<td> PackageCost</td>";
                BodyExcel = BodyExcel + "<td> ProfitLoss</td>";
                BodyExcel = BodyExcel + "<td> Carrier</td>";
                BodyExcel = BodyExcel + "<td> ShippingServiceLevel</td>";
                BodyExcel = BodyExcel + "<td> ShippedByUser</td>";
                BodyExcel = BodyExcel + "<td> ShippingWeight</td>";
                BodyExcel = BodyExcel + "<td> Weight</td>";
                BodyExcel = BodyExcel + "<td> Width</td>";
                BodyExcel = BodyExcel + "<td> Length</td>";
                BodyExcel = BodyExcel + "<td> Height</td>";
                BodyExcel = BodyExcel + "<td> StateRegion</td>";
                BodyExcel = BodyExcel + "<td> TrackingNum</td>";
                BodyExcel = BodyExcel + "<td> MfrName</td>";
                BodyExcel = BodyExcel + "<td> PricingRule</td>";


                // Archivo Fedex
                // -------------
                //BodyExcel = BodyExcel + "<td>FullTrakingId</td>";
                BodyExcel = BodyExcel + "<td> Ground Tracking ID Prefix</td>";
                BodyExcel = BodyExcel + "<td> Express or Ground Tracking ID</td>";
                BodyExcel = BodyExcel + "<td> Net Charge Amount</td>";
                BodyExcel = BodyExcel + "<td> Service Type</td>";
                BodyExcel = BodyExcel + "<td> Ground Service</td>";
                BodyExcel = BodyExcel + "<td> Shipment Date</td>";
                BodyExcel = BodyExcel + "<td> POD Delivery Date</td>";
                BodyExcel = BodyExcel + "<td> Actual Weight Amount</td>";
                BodyExcel = BodyExcel + "<td> Rated Weight Amount</td>";
                BodyExcel = BodyExcel + "<td> Dim Length</td>";
                BodyExcel = BodyExcel + "<td> Dim Width</td>";
                BodyExcel = BodyExcel + "<td> Dim Height</td>";
                BodyExcel = BodyExcel + "<td> Dim Divisor</td>";
                BodyExcel = BodyExcel + "<td> Shipper State</td>";
                BodyExcel = BodyExcel + "<td> Zone Code</td>";
                BodyExcel = BodyExcel + "<td> Tendered Date</td>";

                // cargos fijos
                // ------------
                BodyExcel = BodyExcel + "<td>Earned Discount</td>";
                BodyExcel = BodyExcel + "<td>Fuel Surcharge</td>";
                BodyExcel = BodyExcel + "<td>Performance Pricing</td>";
                BodyExcel = BodyExcel + "<td>Delivery Area Surcharge Extended</td>";
                BodyExcel = BodyExcel + "<td>Delivery Area Surcharge</td>";
                BodyExcel = BodyExcel + "<td>USPS Non-Mach Surcharge</td>";
                BodyExcel = BodyExcel + "<td>Residential</td>";
                BodyExcel = BodyExcel + "<td>Grace Discount</td>";
                BodyExcel = BodyExcel + "<td>Declared Value</td>";
                BodyExcel = BodyExcel + "<td>DAS Extended Resi</td>";
                BodyExcel = BodyExcel + "<td>Additional Handling</td>";
                BodyExcel = BodyExcel + "<td>Parcel Re-Label Charge</td>";
                BodyExcel = BodyExcel + "<td>Indirect Signature</td>";
                BodyExcel = BodyExcel + "<td>DAS Resi</td>";
                BodyExcel = BodyExcel + "<td>Address Correction</td>";
                BodyExcel = BodyExcel + "<td>DAS Extended Comm</td>";
                BodyExcel = BodyExcel + "<td>Oversize Charge</td>";
                BodyExcel = BodyExcel + "<td>AHS - Dimensions</td>";

                // dato USPS
                BodyExcel = BodyExcel + "<td>Mail Class </td>";
                BodyExcel = BodyExcel + "<td>Tracking Number </td>";
                BodyExcel = BodyExcel + "<td>Total Postage Amt </td>";
                BodyExcel = BodyExcel + "<td>Delivery Date </td>";
                BodyExcel = BodyExcel + "<td>Weight </td>";
                BodyExcel = BodyExcel + "<td>Zone </td>";

                // dato UPS
                BodyExcel = BodyExcel + "<td> Service Type </td>";
                BodyExcel = BodyExcel + "<td>Tracking Number </td>";
                BodyExcel = BodyExcel + "<td>Net Charge Amount </td>";


 
                BodyExcel = BodyExcel + "</tr>";

                FileExcel.WriteLine(BodyExcel);
                BodyExcel = "";
            }

        }


        //  inserta fila a reporte
        // -----------------------
        private void InsertaEncabezadoVentasNoFacturado()
        {
            BodyExcel = "<html>";

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            RutaArchivoVentasNoFacturado = pathString + @"\";//ConfigurationManager.AppSettings["RutaArchivoGeneracion"];
            RutaArchivoVentasNoFacturado = RutaArchivoVentasNoFacturado + "OutputVentasNoFacturado" + DateTime.Now.ToString("yyyyMMddTHHmmss") + ".xls";
            using (System.IO.StreamWriter FileExcel = new System.IO.StreamWriter(RutaArchivoVentasNoFacturado, true))
            {

                BodyExcel = BodyExcel + "<body>";
                BodyExcel = BodyExcel + "<table>";
                BodyExcel = BodyExcel + @"<tr bgcolor= ""#CA2229"" style=""color:#ffffff"">";
                // Archivo DW
                // ----------
                BodyExcel = BodyExcel + "<td> SalesOrderNumber</td>";
                BodyExcel = BodyExcel + "<td> TotalSales</td>";
                BodyExcel = BodyExcel + "<td> HoldCode</td>";
                BodyExcel = BodyExcel + "<td> SalesSku</td>";
                BodyExcel = BodyExcel + "<td> SalesCategoryAtTimeOfSale</td>";
                BodyExcel = BodyExcel + "<td> UomCode</td>";
                BodyExcel = BodyExcel + "<td> UomQuantity</td>";
                BodyExcel = BodyExcel + "<td> SalesStatus</td>";
                BodyExcel = BodyExcel + "<td> SalesOrderDate</td>";
                BodyExcel = BodyExcel + "<td> SalesChannelName</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentSku</td>";
                BodyExcel = BodyExcel + "<td> CustomerName</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentChannelName</td>";
                BodyExcel = BodyExcel + "<td> LinkedFulfillmentChannelName</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentLocationName</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentChannelType</td>";
                BodyExcel = BodyExcel + "<td> FulfillmentOrderNumber</td>";
                BodyExcel = BodyExcel + "<td> Quantity</td>";
                BodyExcel = BodyExcel + "<td> Sku</td>";
                BodyExcel = BodyExcel + "<td> Title</td>";
                BodyExcel = BodyExcel + "<td> TotalCost</td>";
                BodyExcel = BodyExcel + "<td> Commission</td>";
                BodyExcel = BodyExcel + "<td> InventoryCost</td>";
                BodyExcel = BodyExcel + "<td> UnitCost</td>";
                BodyExcel = BodyExcel + "<td> ServiceCost</td>";
                BodyExcel = BodyExcel + "<td> EstimatedShippingCost</td>";
                BodyExcel = BodyExcel + "<td> ShippingCost</td>";
                BodyExcel = BodyExcel + "<td> ShippingPrice</td>";
                BodyExcel = BodyExcel + "<td> OverheadCost</td>";
                BodyExcel = BodyExcel + "<td> PackageCost</td>";
                BodyExcel = BodyExcel + "<td> ProfitLoss</td>";
                BodyExcel = BodyExcel + "<td> Carrier</td>";
                BodyExcel = BodyExcel + "<td> ShippingServiceLevel</td>";
                BodyExcel = BodyExcel + "<td> ShippedByUser</td>";
                BodyExcel = BodyExcel + "<td> ShippingWeight</td>";
                BodyExcel = BodyExcel + "<td> Weight</td>";
                BodyExcel = BodyExcel + "<td> Width</td>";
                BodyExcel = BodyExcel + "<td> Length</td>";
                BodyExcel = BodyExcel + "<td> Height</td>";
                BodyExcel = BodyExcel + "<td> StateRegion</td>";
                BodyExcel = BodyExcel + "<td> TrackingNum</td>";
                BodyExcel = BodyExcel + "<td> MfrName</td>";
                BodyExcel = BodyExcel + "<td> PricingRule</td>";

                BodyExcel = BodyExcel + "</tr>";

                FileExcel.WriteLine(BodyExcel);
                BodyExcel = "";
            }

        }

        // realiza la impresion del cargo enviado si la tuviera el reporte de fedex
        // -------------------------------------------------------------------------------
        private void ColumnaCargo(string NombreCargo, string TrackingIDChargeDescription, string TrackingIDChargeAmount, string TrackingIDChargeDescription1, string TrackingIDChargeAmount1, string TrackingIDChargeDescription2, string TrackingIDChargeAmount2, string TrackingIDChargeDescription3, string TrackingIDChargeAmount3, string TrackingIDChargeDescription4, string TrackingIDChargeAmount4, string TrackingIDChargeDescription5, string TrackingIDChargeAmount5, string TrackingIDChargeDescription6, string TrackingIDChargeAmount6, string TrackingIDChargeDescription7, string TrackingIDChargeAmount7, string TrackingIDChargeDescription8, string TrackingIDChargeAmount8, string TrackingIDChargeDescription9, string TrackingIDChargeAmount9, string TrackingIDChargeDescription10, string TrackingIDChargeAmount10, string TrackingIDChargeDescription11, string TrackingIDChargeAmount11, string TrackingIDChargeDescription12, string TrackingIDChargeAmount12, string TrackingIDChargeDescription13, string TrackingIDChargeAmount13, string TrackingIDChargeDescription14, string TrackingIDChargeAmount14, string TrackingIDChargeDescription15, string TrackingIDChargeAmount15, string TrackingIDChargeDescription16, string TrackingIDChargeAmount16, string TrackingIDChargeDescription17, string TrackingIDChargeAmount17, string TrackingIDChargeDescription18, string TrackingIDChargeAmount18, string TrackingIDChargeDescription19, string TrackingIDChargeAmount19, string TrackingIDChargeDescription20, string TrackingIDChargeAmount20, string TrackingIDChargeDescription21, string TrackingIDChargeAmount21, string TrackingIDChargeDescription22, string TrackingIDChargeAmount22, string TrackingIDChargeDescription23, string TrackingIDChargeAmount23, string TrackingIDChargeDescription24, string TrackingIDChargeAmount24)
        {
            if (TrackingIDChargeDescription == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount + "</td>";
            else if (TrackingIDChargeDescription1 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount1 + "</td>";
            else if (TrackingIDChargeDescription2 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount2 + "</td>";
            else if (TrackingIDChargeDescription3 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount3 + "</td>";
            else if (TrackingIDChargeDescription4 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount4 + "</td>";
            else if (TrackingIDChargeDescription5 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount5 + "</td>";
            else if (TrackingIDChargeDescription6 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount6 + "</td>";
            else if (TrackingIDChargeDescription7 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount7 + "</td>";
            else if (TrackingIDChargeDescription8 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount8 + "</td>";
            else if (TrackingIDChargeDescription9 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount9 + "</td>";
            else if (TrackingIDChargeDescription10 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount10 + "</td>";
            else if (TrackingIDChargeDescription11 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount11 + "</td>";
            else if (TrackingIDChargeDescription12 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount12 + "</td>";
            else if (TrackingIDChargeDescription13 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount13 + "</td>";
            else if (TrackingIDChargeDescription14 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount14 + "</td>";
            else if (TrackingIDChargeDescription15 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount15 + "</td>";
            else if (TrackingIDChargeDescription16 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount16 + "</td>";
            else if (TrackingIDChargeDescription17 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount17 + "</td>";
            else if (TrackingIDChargeDescription18 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount18 + "</td>";
            else if (TrackingIDChargeDescription19 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount19 + "</td>";
            else if (TrackingIDChargeDescription20 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount20 + "</td>";
            else if (TrackingIDChargeDescription21 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount21 + "</td>";
            else if (TrackingIDChargeDescription22 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount22 + "</td>";
            else if (TrackingIDChargeDescription23 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount23 + "</td>";
            else if (TrackingIDChargeDescription24 == NombreCargo)
                BodyExcel = BodyExcel + "<td>" + TrackingIDChargeAmount24 + "</td>";
            else
                BodyExcel = BodyExcel + "<td>" + " " + "</td>";
        }

        //  inserta fila a reporte
        // -----------------------
        private void InsertaFilaReporte()
        {
            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            using (System.IO.StreamWriter FileExcel = new System.IO.StreamWriter(RutaArchivoGeneracion, true))
            {

                var PedidoCollection = from s in listaPedido
                                       where s.FullTrakingId == TrackingNum
                                       select new
                                       {
                                           s.GroundTrackingIDPrefix,
                                           s.ExpressorGroundTrackingID,
                                           s.NetChargeAmount,
                                           s.ServiceType,
                                           s.GroundService,
                                           s.ShipmentDate,
                                           s.PODDeliveryDate,
                                           s.ActualWeightAmount,
                                           s.RatedWeightAmount,
                                           s.DimLength,
                                           s.DimWidth,
                                           s.DimHeight,
                                           s.DimDivisor,
                                           s.ShipperState,
                                           s.ZoneCode,
                                           s.TenderedDate,
                                           s.TrackingIDChargeDescription,
                                           s.TrackingIDChargeAmount,
                                           s.TrackingIDChargeDescription1,
                                           s.TrackingIDChargeAmount1,
                                           s.TrackingIDChargeDescription2,
                                           s.TrackingIDChargeAmount2,
                                           s.TrackingIDChargeDescription3,
                                           s.TrackingIDChargeAmount3,
                                           s.TrackingIDChargeDescription4,
                                           s.TrackingIDChargeAmount4,
                                           s.TrackingIDChargeDescription5,
                                           s.TrackingIDChargeAmount5,
                                           s.TrackingIDChargeDescription6,
                                           s.TrackingIDChargeAmount6,
                                           s.TrackingIDChargeDescription7,
                                           s.TrackingIDChargeAmount7,
                                           s.TrackingIDChargeDescription8,
                                           s.TrackingIDChargeAmount8,
                                           s.TrackingIDChargeDescription9,
                                           s.TrackingIDChargeAmount9,
                                           s.TrackingIDChargeDescription10,
                                           s.TrackingIDChargeAmount10,
                                           s.TrackingIDChargeDescription11,
                                           s.TrackingIDChargeAmount11,
                                           s.TrackingIDChargeDescription12,
                                           s.TrackingIDChargeAmount12,
                                           s.TrackingIDChargeDescription13,
                                           s.TrackingIDChargeAmount13,
                                           s.TrackingIDChargeDescription14,
                                           s.TrackingIDChargeAmount14,
                                           s.TrackingIDChargeDescription15,
                                           s.TrackingIDChargeAmount15,
                                           s.TrackingIDChargeDescription16,
                                           s.TrackingIDChargeAmount16,
                                           s.TrackingIDChargeDescription17,
                                           s.TrackingIDChargeAmount17,
                                           s.TrackingIDChargeDescription18,
                                           s.TrackingIDChargeAmount18,
                                           s.TrackingIDChargeDescription19,
                                           s.TrackingIDChargeAmount19,
                                           s.TrackingIDChargeDescription20,
                                           s.TrackingIDChargeAmount20,
                                           s.TrackingIDChargeDescription21,
                                           s.TrackingIDChargeAmount21,
                                           s.TrackingIDChargeDescription22,
                                           s.TrackingIDChargeAmount22,
                                           s.TrackingIDChargeDescription23,
                                           s.TrackingIDChargeAmount23,
                                           s.TrackingIDChargeDescription24,
                                           s.TrackingIDChargeAmount24
                                       };

                foreach (var Pedido in PedidoCollection)
                {

                    // arma la fila con el color de fondo que corresponde
                    // --------------------------------------------------
                    if (contador == 1)
                        BodyExcel = @"<tr bgcolor= ""#FF9F9F"" >";
                    else
                        BodyExcel = @"<tr bgcolor= ""#FFFFFF"" >";

                    // archivo base
                    // ------------
                    BodyExcel = BodyExcel + "<td>'" + SalesOrderNumber + "</td>";
                    BodyExcel = BodyExcel + "<td>" + HoldCode + "</td>";
                    BodyExcel = BodyExcel + "<td>" + TotalSales + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesSku + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesCategoryAtTimeOfSale + "</td>";
                    BodyExcel = BodyExcel + "<td>" + UomCode + "</td>";
                    BodyExcel = BodyExcel + "<td>" + UomQuantity + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesStatus + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesOrderDate + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesChannelName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + CustomerName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentSku + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentChannelName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentChannelType + "</td>";
                    BodyExcel = BodyExcel + "<td>" + LinkedFulfillmentChannelName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentLocationName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentOrderNumber + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Quantity + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Sku + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Title + "</td>";
                    BodyExcel = BodyExcel + "<td>" + TotalCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Commission + "</td>";
                    BodyExcel = BodyExcel + "<td>" + InventoryCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + UnitCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ServiceCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + EstimatedShippingCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingPrice + "</td>";
                    BodyExcel = BodyExcel + "<td>" + OverheadCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + PackageCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ProfitLoss + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Carrier + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingServiceLevel + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippedByUser + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingWeight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Length + "</td>";
                    BodyExcel = BodyExcel + "<td>" + varWidth + "</td>";
                    BodyExcel = BodyExcel + "<td>" + varHeight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Weight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + StateRegion + "</td>";
                    BodyExcel = BodyExcel + "<td>'" + TrackingNum + "</td>";
                    BodyExcel = BodyExcel + "<td>" + MfrName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + PricingRule + "</td>";

                    // archivo fedex
                    // -------------
                    BodyExcel = BodyExcel + "<td>'" + Pedido.GroundTrackingIDPrefix + "</td>";
                    BodyExcel = BodyExcel + "<td>'" + Pedido.ExpressorGroundTrackingID+"</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.NetChargeAmount+"</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.ServiceType + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.GroundService + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.ShipmentDate + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.PODDeliveryDate+"</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.ActualWeightAmount+"</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.RatedWeightAmount+"</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.DimLength + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.DimWidth + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.DimHeight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.DimDivisor + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.ShipperState + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.ZoneCode + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.TenderedDate + "</td>";

                    string NombreCargo = "Earned Discount";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Fuel Surcharge";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Performance Pricing";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Delivery Area Surcharge Extended";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Delivery Area Surcharge";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "USPS Non-Mach Surcharge";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Residential";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Grace Discount";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Declared Value";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "DAS Extended Resi";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Additional Handling";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Parcel Re-Label Charge";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Indirect Signature";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "DAS Resi";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    //NombreCargo = "DAS Resi";
                    //ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Address Correction";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "DAS Extended Comm";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "Oversize Charge";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    NombreCargo = "AHS - Dimensions";
                    ColumnaCargo(NombreCargo, Pedido.TrackingIDChargeDescription, Pedido.TrackingIDChargeAmount, Pedido.TrackingIDChargeDescription1, Pedido.TrackingIDChargeAmount1, Pedido.TrackingIDChargeDescription2, Pedido.TrackingIDChargeAmount2, Pedido.TrackingIDChargeDescription3, Pedido.TrackingIDChargeAmount3, Pedido.TrackingIDChargeDescription4, Pedido.TrackingIDChargeAmount4, Pedido.TrackingIDChargeDescription5, Pedido.TrackingIDChargeAmount5, Pedido.TrackingIDChargeDescription6, Pedido.TrackingIDChargeAmount6, Pedido.TrackingIDChargeDescription7, Pedido.TrackingIDChargeAmount7, Pedido.TrackingIDChargeDescription8, Pedido.TrackingIDChargeAmount8, Pedido.TrackingIDChargeDescription9, Pedido.TrackingIDChargeAmount9, Pedido.TrackingIDChargeDescription10, Pedido.TrackingIDChargeAmount10, Pedido.TrackingIDChargeDescription11, Pedido.TrackingIDChargeAmount11, Pedido.TrackingIDChargeDescription12, Pedido.TrackingIDChargeAmount12, Pedido.TrackingIDChargeDescription13, Pedido.TrackingIDChargeAmount13, Pedido.TrackingIDChargeDescription14, Pedido.TrackingIDChargeAmount14, Pedido.TrackingIDChargeDescription15, Pedido.TrackingIDChargeAmount15, Pedido.TrackingIDChargeDescription16, Pedido.TrackingIDChargeAmount16, Pedido.TrackingIDChargeDescription17, Pedido.TrackingIDChargeAmount17, Pedido.TrackingIDChargeDescription18, Pedido.TrackingIDChargeAmount18, Pedido.TrackingIDChargeDescription19, Pedido.TrackingIDChargeAmount19, Pedido.TrackingIDChargeDescription20, Pedido.TrackingIDChargeAmount20, Pedido.TrackingIDChargeDescription21, Pedido.TrackingIDChargeAmount21, Pedido.TrackingIDChargeDescription22, Pedido.TrackingIDChargeAmount22, Pedido.TrackingIDChargeDescription23, Pedido.TrackingIDChargeAmount23, Pedido.TrackingIDChargeDescription24, Pedido.TrackingIDChargeAmount24);

                    BodyExcel = BodyExcel + "</tr>";
                    break;
                }

                FileExcel.WriteLine(BodyExcel);
                BodyExcel= "";

                // incrementa contador para saber el color de linea que corresponde a la fila procesada
                // ------------------------------------------------------------------------------------
                contador = contador + 1;

                // solo se tienen dos colores por lo que si sobrepasa de 2 inicializa el contador
                // ------------------------------------------------------------------------------
                if (contador > 2)
                    contador = 1;
            }
        }

        //  inserta fila a reporte
        // -----------------------
        private void InsertaFilaReporteUSPS()
        {
            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            using (System.IO.StreamWriter FileExcel = new System.IO.StreamWriter(RutaArchivoGeneracion, true))
            {

               //var PedidoCollection = from s in listaPedidoUSPS
               //                       where s.TrackingNumber == TrackingNum
               //                       select new
               //                       {
               //                            s.GroundService,
               //                            s.TrackingNumber,
               //                            s.NetChargeAmount,
               //                            s.PODDeliveryDate,
               //                            s.RatedWeightAmount,
               //                            s.ZoneCode
               //                       };

                //foreach (var Pedido in PedidoCollection)
                //{
                //    // arma la fila con el color de fondo que corresponde
                //    // --------------------------------------------------
                //    if (contador == 1)
                //        BodyExcel = @"<tr bgcolor= ""#FF9F9F"" >";
                //    else
                //        BodyExcel = @"<tr bgcolor= ""#FFFFFF"" >";
                //
                //    // archivo base
                //    // ------------
                //    BodyExcel = BodyExcel + "<td>'" + SalesOrderNumber + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + HoldCode + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + TotalSales + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + SalesSku + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + SalesCategoryAtTimeOfSale + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + UomCode + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + UomQuantity + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + SalesStatus + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + SalesOrderDate + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + SalesChannelName + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + CustomerName + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + FulfillmentSku + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + FulfillmentChannelName + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + FulfillmentChannelType + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + LinkedFulfillmentChannelName + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + FulfillmentLocationName + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + FulfillmentOrderNumber + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Quantity + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Sku + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Title + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + TotalCost + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Commission + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + InventoryCost + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + UnitCost + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + ServiceCost + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + EstimatedShippingCost + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + ShippingCost + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + ShippingPrice + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + OverheadCost + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + PackageCost + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + ProfitLoss + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Carrier + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + ShippingServiceLevel + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + ShippedByUser + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + ShippingWeight + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Length + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + varWidth + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + varHeight + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Weight + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + StateRegion + "</td>";
                //    BodyExcel = BodyExcel + "<td>'" + TrackingNum + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + MfrName + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + PricingRule + "</td>";
                //
                //    // archivo fedex
                //    // -------------
                //    string vacio = ""; 
                //    BodyExcel = BodyExcel + "<td>'" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>'" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                //
                //    string NombreCargo = "Earned Discount";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Fuel Surcharge";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Performance Pricing";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Delivery Area Surcharge Extended";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Delivery Area Surcharge";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "USPS Non-Mach Surcharge";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Residential";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Grace Discount";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Declared Value";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "DAS Extended Resi";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Additional Handling";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Parcel Re-Label Charge";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Indirect Signature";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "DAS Resi";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "DAS Resi";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Address Correction";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "DAS Extended Comm";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "Oversize Charge";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    NombreCargo = "AHS - Dimensions";
                //    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);
                //
                //    // dato USPS
                //    BodyExcel = BodyExcel + "<td>" + Pedido.GroundService + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Pedido.TrackingNumber + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Pedido.NetChargeAmount + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Pedido.PODDeliveryDate + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Pedido.RatedWeightAmount + "</td>";
                //    BodyExcel = BodyExcel + "<td>" + Pedido.ZoneCode + "</td>";
                //
                //    BodyExcel = BodyExcel + "</tr>";
                //    break;
                //}

                FileExcel.WriteLine(BodyExcel);
                BodyExcel = "";

                // incrementa contador para saber el color de linea que corresponde a la fila procesada
                // ------------------------------------------------------------------------------------
                contador = contador + 1;

                // solo se tienen dos colores por lo que si sobrepasa de 2 inicializa el contador
                // ------------------------------------------------------------------------------
                if (contador > 2)
                    contador = 1;
            }
        }

        //  inserta fila a reporte
        // -----------------------
        private void InsertaFilaReporteUPS()
        {
            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            using (System.IO.StreamWriter FileExcel = new System.IO.StreamWriter(RutaArchivoGeneracion, true))
            {

                var PedidoCollection = from s in listaPedidoUPS
                                       where s.Campo30 == TrackingNum
                                       select new
                                       {
                                           s.Campo12,
                                           s.Campo30,
                                           s.Campo39
                                       };

                foreach (var Pedido in PedidoCollection)
                {
                    // arma la fila con el color de fondo que corresponde
                    // --------------------------------------------------
                    if (contador == 1)
                        BodyExcel = @"<tr bgcolor= ""#FF9F9F"" >";
                    else
                        BodyExcel = @"<tr bgcolor= ""#FFFFFF"" >";

                    // archivo base
                    // ------------
                    BodyExcel = BodyExcel + "<td>'" + SalesOrderNumber + "</td>";
                    BodyExcel = BodyExcel + "<td>" + HoldCode + "</td>";
                    BodyExcel = BodyExcel + "<td>" + TotalSales + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesSku + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesCategoryAtTimeOfSale + "</td>";
                    BodyExcel = BodyExcel + "<td>" + UomCode + "</td>";
                    BodyExcel = BodyExcel + "<td>" + UomQuantity + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesStatus + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesOrderDate + "</td>";
                    BodyExcel = BodyExcel + "<td>" + SalesChannelName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + CustomerName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentSku + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentChannelName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentChannelType + "</td>";
                    BodyExcel = BodyExcel + "<td>" + LinkedFulfillmentChannelName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentLocationName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + FulfillmentOrderNumber + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Quantity + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Sku + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Title + "</td>";
                    BodyExcel = BodyExcel + "<td>" + TotalCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Commission + "</td>";
                    BodyExcel = BodyExcel + "<td>" + InventoryCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + UnitCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ServiceCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + EstimatedShippingCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingPrice + "</td>";
                    BodyExcel = BodyExcel + "<td>" + OverheadCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + PackageCost + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ProfitLoss + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Carrier + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingServiceLevel + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippedByUser + "</td>";
                    BodyExcel = BodyExcel + "<td>" + ShippingWeight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Length + "</td>";
                    BodyExcel = BodyExcel + "<td>" + varWidth + "</td>";
                    BodyExcel = BodyExcel + "<td>" + varHeight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Weight + "</td>";
                    BodyExcel = BodyExcel + "<td>" + StateRegion + "</td>";
                    BodyExcel = BodyExcel + "<td>'" + TrackingNum + "</td>";
                    BodyExcel = BodyExcel + "<td>" + MfrName + "</td>";
                    BodyExcel = BodyExcel + "<td>" + PricingRule + "</td>";

                    // archivo fedex
                    // -------------
                    string vacio = "";
                    BodyExcel = BodyExcel + "<td>'" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>'" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";

                    string NombreCargo = "Earned Discount";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Fuel Surcharge";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Performance Pricing";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Delivery Area Surcharge Extended";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Delivery Area Surcharge";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "USPS Non-Mach Surcharge";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Residential";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Grace Discount";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Declared Value";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "DAS Extended Resi";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Additional Handling";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Parcel Re-Label Charge";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Indirect Signature";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "DAS Resi";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "DAS Resi";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Address Correction";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "DAS Extended Comm";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "Oversize Charge";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    NombreCargo = "AHS - Dimensions";
                    ColumnaCargo(NombreCargo, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio, vacio);

                    // dato USPS
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";
                    BodyExcel = BodyExcel + "<td>" + vacio + "</td>";

                    // dato UPS
                    BodyExcel = BodyExcel + "<td>" + Pedido.Campo12 + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.Campo30 + "</td>";
                    BodyExcel = BodyExcel + "<td>" + Pedido.Campo39 + "</td>";

                    BodyExcel = BodyExcel + "</tr>";
                    break;
                }

                FileExcel.WriteLine(BodyExcel);
                BodyExcel = "";

                // incrementa contador para saber el color de linea que corresponde a la fila procesada
                // ------------------------------------------------------------------------------------
                contador = contador + 1;

                // solo se tienen dos colores por lo que si sobrepasa de 2 inicializa el contador
                // ------------------------------------------------------------------------------
                if (contador > 2)
                    contador = 1;
            }
        }


        //  inserta fila a reporte
        // -----------------------
        private void VentasNoFacturado()
        {
            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            using (System.IO.StreamWriter FileExcel = new System.IO.StreamWriter(RutaArchivoVentasNoFacturado, true))
            {
                // arma la fila con el color de fondo que corresponde
                // --------------------------------------------------
                if (contador == 1)
                    BodyExcel = @"<tr bgcolor= ""#FF9F9F"" >";
                else
                    BodyExcel = @"<tr bgcolor= ""#FFFFFF"" >";

                // archivo base
                // ------------
                BodyExcel = BodyExcel + "<td>'" + SalesOrderNumber + "</td>";
                BodyExcel = BodyExcel + "<td>" + HoldCode + "</td>";
                BodyExcel = BodyExcel + "<td>" + TotalSales + "</td>";
                BodyExcel = BodyExcel + "<td>" + SalesSku + "</td>";
                BodyExcel = BodyExcel + "<td>" + SalesCategoryAtTimeOfSale + "</td>";
                BodyExcel = BodyExcel + "<td>" + UomCode + "</td>";
                BodyExcel = BodyExcel + "<td>" + UomQuantity + "</td>";
                BodyExcel = BodyExcel + "<td>" + SalesStatus + "</td>";
                BodyExcel = BodyExcel + "<td>" + SalesOrderDate + "</td>";
                BodyExcel = BodyExcel + "<td>" + SalesChannelName + "</td>";
                BodyExcel = BodyExcel + "<td>" + CustomerName + "</td>";
                BodyExcel = BodyExcel + "<td>" + FulfillmentSku + "</td>";
                BodyExcel = BodyExcel + "<td>" + FulfillmentChannelName + "</td>";
                BodyExcel = BodyExcel + "<td>" + FulfillmentChannelType + "</td>";
                BodyExcel = BodyExcel + "<td>" + LinkedFulfillmentChannelName + "</td>";
                BodyExcel = BodyExcel + "<td>" + FulfillmentLocationName + "</td>";
                BodyExcel = BodyExcel + "<td>" + FulfillmentOrderNumber + "</td>";
                BodyExcel = BodyExcel + "<td>" + Quantity + "</td>";
                BodyExcel = BodyExcel + "<td>" + Sku + "</td>";
                BodyExcel = BodyExcel + "<td>" + Title + "</td>";
                BodyExcel = BodyExcel + "<td>" + TotalCost + "</td>";
                BodyExcel = BodyExcel + "<td>" + Commission + "</td>";
                BodyExcel = BodyExcel + "<td>" + InventoryCost + "</td>";
                BodyExcel = BodyExcel + "<td>" + UnitCost + "</td>";
                BodyExcel = BodyExcel + "<td>" + ServiceCost + "</td>";
                BodyExcel = BodyExcel + "<td>" + EstimatedShippingCost + "</td>";
                BodyExcel = BodyExcel + "<td>" + ShippingCost + "</td>";
                BodyExcel = BodyExcel + "<td>" + ShippingPrice + "</td>";
                BodyExcel = BodyExcel + "<td>" + OverheadCost + "</td>";
                BodyExcel = BodyExcel + "<td>" + PackageCost + "</td>";
                BodyExcel = BodyExcel + "<td>" + ProfitLoss + "</td>";
                BodyExcel = BodyExcel + "<td>" + Carrier + "</td>";
                BodyExcel = BodyExcel + "<td>" + ShippingServiceLevel + "</td>";
                BodyExcel = BodyExcel + "<td>" + ShippedByUser + "</td>";
                BodyExcel = BodyExcel + "<td>" + ShippingWeight + "</td>";
                BodyExcel = BodyExcel + "<td>" + Length + "</td>";
                BodyExcel = BodyExcel + "<td>" + varWidth + "</td>";
                BodyExcel = BodyExcel + "<td>" + varHeight + "</td>";
                BodyExcel = BodyExcel + "<td>" + Weight + "</td>";
                BodyExcel = BodyExcel + "<td>" + StateRegion + "</td>";
                BodyExcel = BodyExcel + "<td>'" + TrackingNum + "</td>";
                BodyExcel = BodyExcel + "<td>" + MfrName + "</td>";
                BodyExcel = BodyExcel + "<td>" + PricingRule + "</td>";

                BodyExcel = BodyExcel + "</tr>";

                FileExcel.WriteLine(BodyExcel);
                BodyExcel = "";

                // incrementa contador para saber el color de linea que corresponde a la fila procesada
                // ------------------------------------------------------------------------------------
                contador = contador + 1;

                // solo se tienen dos colores por lo que si sobrepasa de 2 inicializa el contador
                // ------------------------------------------------------------------------------
                if (contador > 2)
                    contador = 1;
            }
        }


        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneValorRegistro(string[] valor/*DataRow row*/)
        {
            SalesOrderNumber             = valor[0];//Convert.ToString(row["SalesOrderNumber"]);
            HoldCode                     = valor[1];//Convert.ToString(row["HoldCode"]);
            TotalSales                   = valor[2];//Convert.ToString(row["TotalSales"]);
            SalesSku                     = valor[3];//Convert.ToString(row["SalesSku"]);
            SalesCategoryAtTimeOfSale    = valor[4];//Convert.ToString(row["SalesCategoryAtTimeOfSale"]);
            UomCode                      = valor[5];//Convert.ToString(row["UomCode"]);
            UomQuantity                  = valor[6];//Convert.ToString(row["UomQuantity"]);
            SalesStatus                  = valor[7];//Convert.ToString(row["SalesStatus"]);
            SalesOrderDate               = valor[8];//Convert.ToString(row["SalesOrderDate"]);
            SalesChannelName             = valor[9];//Convert.ToString(row["SalesChannelName"]);
            CustomerName                 = valor[10];//Convert.ToString(row["CustomerName"]);
            FulfillmentSku               = valor[11];//Convert.ToString(row["FulfillmentSku"]);
            FulfillmentChannelName       = valor[12];//Convert.ToString(row["FulfillmentChannelName"]);
            FulfillmentChannelType       = valor[13];//Convert.ToString(row["FulfillmentChannelType"]);
            LinkedFulfillmentChannelName = valor[14];//Convert.ToString(row["LinkedFulfillmentChannelName"]);
            FulfillmentLocationName      = valor[15];//Convert.ToString(row["FulfillmentLocationName"]);
            FulfillmentOrderNumber       = valor[16];//Convert.ToString(row["FulfillmentOrderNumber"]);
            Quantity                     = valor[17];//Convert.ToString(row["Quantity"]);
            Sku                          = valor[18];//Convert.ToString(row["Sku"]);
            Title                        = valor[19];//Convert.ToString(row["Title"]);
            TotalCost                    = valor[20];//Convert.ToString(row["TotalCost"]);
            Commission                   = valor[21];//Convert.ToString(row["Commission"]);
            InventoryCost                = valor[22];//Convert.ToString(row["InventoryCost"]);
            UnitCost                     = valor[23];//Convert.ToString(row["UnitCost"]);
            ServiceCost                  = valor[24];//Convert.ToString(row["ServiceCost"]);
            EstimatedShippingCost        = valor[25];//Convert.ToString(row["EstimatedShippingCost"]);
            ShippingCost                 = valor[26];//Convert.ToString(row["ShippingCost"]);
            ShippingPrice                = valor[27];//Convert.ToString(row["ShippingPrice"]);
            OverheadCost                 = valor[28];//Convert.ToString(row["OverheadCost"]);
            PackageCost                  = valor[29];//Convert.ToString(row["PackageCost"]);
            ProfitLoss                   = valor[30];//Convert.ToString(row["ProfitLoss"]);
            Carrier                      = valor[31];//Convert.ToString(row["Carrier"]);
            ShippingServiceLevel         = valor[32];//Convert.ToString(row["ShippingServiceLevel"]);
            ShippedByUser                = valor[33];//Convert.ToString(row["ShippedByUser"]);
            ShippingWeight               = valor[34];//Convert.ToString(row["ShippingWeight"]);
            Length                       = valor[35];//Convert.ToString(row["Length"]);
            varWidth                     = valor[36];//Convert.ToString(row["Width"]);
            varHeight                    = valor[37];//Convert.ToString(row["Height"]);
            Weight                       = valor[38];//Convert.ToString(row["Weight"]);
            StateRegion                  = valor[39];//Convert.ToString(row["StateRegion"]);
            TrackingNum                  = valor[40];//Convert.ToString(row["TrackingNum"]);
            MfrName                      = valor[41];//Convert.ToString(row["MfrName"]);
            PricingRule                  = valor[42];//Convert.ToString(row["PricingRule"]);
            ActualShippingCost           = valor[43];
            ActualShipping               = valor[44];
            ShippingCostDifference       = valor[45];
        }
        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneValorRegistroDetalle(DataRow row, ref PedidoFedex clsPedido)
        {
            clsPedido.GroundTrackingIDPrefix = Convert.ToString(row["Ground Tracking ID Prefix"]);
            clsPedido.ExpressorGroundTrackingID = Convert.ToString(row["Express or Ground Tracking ID"]);
            clsPedido.FullTrakingId = clsPedido.GroundTrackingIDPrefix + clsPedido.ExpressorGroundTrackingID;
            clsPedido.BilltoAccountNumber = Convert.ToString(row["Bill to Account Number"]);
            clsPedido.InvoiceDate = Convert.ToString(row["Invoice Date"]);
            clsPedido.InvoiceNumber = Convert.ToString(row["Invoice Number"]);
            clsPedido.StoreID = Convert.ToString(row["Store ID"]);
            clsPedido.OriginalAmountDue = Convert.ToString(row["Original Amount Due"]);
            clsPedido.CurrentBalance = Convert.ToString(row["Current Balance"]);
            clsPedido.Payor = Convert.ToString(row["Payor"]);
            clsPedido.TransportationChargeAmount = Convert.ToString(row["Transportation Charge Amount"]);
            clsPedido.NetChargeAmount = Convert.ToString(row["Net Charge Amount"]);
            clsPedido.ServiceType = Convert.ToString(row["Service Type"]);
            clsPedido.GroundService = Convert.ToString(row["Ground Service"]);
            clsPedido.ShipmentDate = Convert.ToString(row["Shipment Date"]);
            clsPedido.PODDeliveryDate = Convert.ToString(row["POD Delivery Date"]);
            clsPedido.PODDeliveryTime = Convert.ToString(row["POD Delivery Time"]);
            clsPedido.PODServiceAreaCode = Convert.ToString(row["POD Service Area Code"]);
            clsPedido.PODSignatureDescription = Convert.ToString(row["POD Signature Description"]);
            clsPedido.ActualWeightAmount = Convert.ToString(row["Actual Weight Amount"]);
            clsPedido.ActualWeightUnits = Convert.ToString(row["Actual Weight Units"]);
            clsPedido.RatedWeightAmount = Convert.ToString(row["Rated Weight Amount"]);
            clsPedido.RatedWeightUnits = Convert.ToString(row["Rated Weight Units"]);
            clsPedido.NumberofPieces = Convert.ToString(row["Number of Pieces"]);
            clsPedido.BundleNumber = Convert.ToString(row["Bundle Number"]);
            clsPedido.MeterNumber = Convert.ToString(row["Meter Number"]);
            clsPedido.TDMasterTrackingID = Convert.ToString(row["TDMasterTrackingID"]);
            clsPedido.ServicePackaging = Convert.ToString(row["Service Packaging"]);
            clsPedido.DimLength = Convert.ToString(row["Dim Length"]);
            clsPedido.DimWidth = Convert.ToString(row["Dim Width"]);
            clsPedido.DimHeight = Convert.ToString(row["Dim Height"]);
            clsPedido.DimDivisor = Convert.ToString(row["Dim Divisor"]);
            clsPedido.DimUnit = Convert.ToString(row["Dim Unit"]);
            clsPedido.RecipientName = Convert.ToString(row["Recipient Name"]);
            clsPedido.RecipientCompany = Convert.ToString(row["Recipient Company"]);
            clsPedido.RecipientAddressLine1 = Convert.ToString(row["Recipient Address Line 1"]);
            clsPedido.RecipientAddressLine2 = Convert.ToString(row["Recipient Address Line 2"]);
            clsPedido.RecipientCity = Convert.ToString(row["Recipient City"]);
            clsPedido.RecipientState = Convert.ToString(row["Recipient State"]);
            clsPedido.RecipientZipCode = Convert.ToString(row["Recipient Zip Code"]);
            clsPedido.ShipperCompany = Convert.ToString(row["Shipper Company"]);
            clsPedido.ShipperName = Convert.ToString(row["Shipper Name"]);
            clsPedido.ShipperAddressLine1 = Convert.ToString(row["Shipper Address Line 1"]);
            clsPedido.ShipperAddressLine2 = Convert.ToString(row["Shipper Address Line 2"]);
            clsPedido.ShipperCity = Convert.ToString(row["Shipper City"]);
            clsPedido.ShipperState = Convert.ToString(row["Shipper State"]);
            clsPedido.ShipperZipCode = Convert.ToString(row["Shipper Zip Code"]);
            clsPedido.OriginalCustomerReference = Convert.ToString(row["Original Customer Reference"]);
            clsPedido.OriginalDepartmentReferenceDescription = Convert.ToString(row["Original Department Reference Description"]);
            clsPedido.UpdatedCustomerReference = Convert.ToString(row["Updated Customer Reference"]);
            clsPedido.UpdatedDepartmentReferenceDescription = Convert.ToString(row["Updated Department Reference Description"]);
            clsPedido.OriginalRecipientAddressLine1 = Convert.ToString(row["Original Recipient Address Line 1"]);
            clsPedido.OriginalRecipientAddressLine2 = Convert.ToString(row["Original Recipient Address Line 2"]);
            clsPedido.OriginalRecipientCity = Convert.ToString(row["Original Recipient City"]);
            clsPedido.OriginalRecipientState = Convert.ToString(row["Original Recipient State"]);
            clsPedido.OriginalRecipientZipCode = Convert.ToString(row["Original Recipient Zip Code"]);
            clsPedido.ZoneCode = Convert.ToString(row["Zone Code"]);
            clsPedido.CostAllocation = Convert.ToString(row["Cost Allocation"]);
            clsPedido.AlternateAddressLine1 = Convert.ToString(row["Alternate Address Line 1"]);
            clsPedido.AlternateAddressLine2 = Convert.ToString(row["Alternate Address Line 2"]);
            clsPedido.AlternateCity = Convert.ToString(row["Alternate City"]);
            clsPedido.AlternateStateProvince = Convert.ToString(row["Alternate State Province"]);
            clsPedido.AlternateZipCode = Convert.ToString(row["Alternate Zip Code"]);
            clsPedido.CrossRefTrackingIDPrefix = Convert.ToString(row["CrossRefTrackingID Prefix"]);
            clsPedido.CrossRefTrackingID = Convert.ToString(row["CrossRefTrackingID"]);
            clsPedido.EntryDate = Convert.ToString(row["Entry Date"]);
            clsPedido.EntryNumber = Convert.ToString(row["Entry Number"]);
            clsPedido.CustomsValue = Convert.ToString(row["Customs Value"]);
            clsPedido.CustomsValueCurrencyCode = Convert.ToString(row["Customs Value Currency Code"]);
            clsPedido.DeclaredValue = Convert.ToString(row["Declared Value"]);
            clsPedido.DeclaredValueCurrencyCode = Convert.ToString(row["Declared Value Currency Code"]);
            clsPedido.CurrencyConversionDate = Convert.ToString(row["Currency Conversion Date"]);
            clsPedido.CurrencyConversionRate = Convert.ToString(row["Currency Conversion Rate"]);
            clsPedido.MultiweightNumber = Convert.ToString(row["Multiweight Number"]);
            clsPedido.MultiweightTotalMultiweightUnits = Convert.ToString(row["Multiweight Total Multiweight Units"]);
            clsPedido.MultiweightTotalMultiweightWeight = Convert.ToString(row["Multiweight Total Multiweight Weight"]);
            clsPedido.MultiweightTotalShipmentChargeAmount = Convert.ToString(row["Multiweight Total Shipment Charge Amount"]);
            clsPedido.MultiweightTotalShipmentWeight = Convert.ToString(row["Multiweight Total Shipment Weight"]);
            clsPedido.GroundTrackingIDAddressCorrectionDiscountChargeAmount = Convert.ToString(row["Ground Tracking ID Address Correction Discount Charge Amount"]);
            clsPedido.GroundTrackingIDAddressCorrectionGrossChargeAmount = Convert.ToString(row["Ground Tracking ID Address Correction Gross Charge Amount"]);
            clsPedido.RatedMethod = Convert.ToString(row["Rated Method"]);
            clsPedido.SortHub = Convert.ToString(row["Sort Hub"]);
            clsPedido.EstimatedWeight = Convert.ToString(row["Estimated Weight"]);
            clsPedido.EstimatedWeightUnit = Convert.ToString(row["Estimated Weight Unit"]);
            clsPedido.PostalClass = Convert.ToString(row["Postal Class"]);
            clsPedido.ProcessCategory = Convert.ToString(row["Process Category"]);
            clsPedido.PackageSize = Convert.ToString(row["Package Size"]);
            clsPedido.DeliveryConfirmation = Convert.ToString(row["Delivery Confirmation"]);
            clsPedido.TenderedDate = Convert.ToString(row["Tendered Date"]);
            clsPedido.TrackingIDChargeDescription = Convert.ToString(row["Tracking ID Charge Description"]);
            clsPedido.TrackingIDChargeAmount = Convert.ToString(row["Tracking ID Charge Amount"]);
            clsPedido.TrackingIDChargeDescription1 = Convert.ToString(row["Tracking ID Charge Description1"]);
            clsPedido.TrackingIDChargeAmount1 = Convert.ToString(row["Tracking ID Charge Amount1"]);
            clsPedido.TrackingIDChargeDescription2 = Convert.ToString(row["Tracking ID Charge Description2"]);
            clsPedido.TrackingIDChargeAmount2 = Convert.ToString(row["Tracking ID Charge Amount2"]);
            clsPedido.TrackingIDChargeDescription3 = Convert.ToString(row["Tracking ID Charge Description3"]);
            clsPedido.TrackingIDChargeAmount3 = Convert.ToString(row["Tracking ID Charge Amount3"]);
            clsPedido.TrackingIDChargeDescription4 = Convert.ToString(row["Tracking ID Charge Description4"]);
            clsPedido.TrackingIDChargeAmount4 = Convert.ToString(row["Tracking ID Charge Amount4"]);
            clsPedido.TrackingIDChargeDescription5 = Convert.ToString(row["Tracking ID Charge Description5"]);
            clsPedido.TrackingIDChargeAmount5 = Convert.ToString(row["Tracking ID Charge Amount5"]);
            clsPedido.TrackingIDChargeDescription6 = Convert.ToString(row["Tracking ID Charge Description6"]);
            clsPedido.TrackingIDChargeAmount6 = Convert.ToString(row["Tracking ID Charge Amount6"]);
            clsPedido.TrackingIDChargeDescription7 = Convert.ToString(row["Tracking ID Charge Description7"]);
            clsPedido.TrackingIDChargeAmount7 = Convert.ToString(row["Tracking ID Charge Amount7"]);
            clsPedido.TrackingIDChargeDescription8 = Convert.ToString(row["Tracking ID Charge Description8"]);
            clsPedido.TrackingIDChargeAmount8 = Convert.ToString(row["Tracking ID Charge Amount8"]);
            clsPedido.TrackingIDChargeDescription9 = Convert.ToString(row["Tracking ID Charge Description9"]);
            clsPedido.TrackingIDChargeAmount9 = Convert.ToString(row["Tracking ID Charge Amount9"]);
            clsPedido.TrackingIDChargeDescription10 = Convert.ToString(row["Tracking ID Charge Description10"]);
            clsPedido.TrackingIDChargeAmount10 = Convert.ToString(row["Tracking ID Charge Amount10"]);
            clsPedido.TrackingIDChargeDescription11 = Convert.ToString(row["Tracking ID Charge Description11"]);
            clsPedido.TrackingIDChargeAmount11 = Convert.ToString(row["Tracking ID Charge Amount11"]);
            clsPedido.TrackingIDChargeDescription12 = Convert.ToString(row["Tracking ID Charge Description12"]);
            clsPedido.TrackingIDChargeAmount12 = Convert.ToString(row["Tracking ID Charge Amount12"]);
            clsPedido.TrackingIDChargeDescription13 = Convert.ToString(row["Tracking ID Charge Description13"]);
            clsPedido.TrackingIDChargeAmount13 = Convert.ToString(row["Tracking ID Charge Amount13"]);
            clsPedido.TrackingIDChargeDescription14 = Convert.ToString(row["Tracking ID Charge Description14"]);
            clsPedido.TrackingIDChargeAmount14 = Convert.ToString(row["Tracking ID Charge Amount14"]);
            clsPedido.TrackingIDChargeDescription15 = Convert.ToString(row["Tracking ID Charge Description15"]);
            clsPedido.TrackingIDChargeAmount15 = Convert.ToString(row["Tracking ID Charge Amount15"]);
            clsPedido.TrackingIDChargeDescription16 = Convert.ToString(row["Tracking ID Charge Description16"]);
            clsPedido.TrackingIDChargeAmount16 = Convert.ToString(row["Tracking ID Charge Amount16"]);
            clsPedido.TrackingIDChargeDescription17 = Convert.ToString(row["Tracking ID Charge Description17"]);
            clsPedido.TrackingIDChargeAmount17 = Convert.ToString(row["Tracking ID Charge Amount17"]);
            clsPedido.TrackingIDChargeDescription18 = Convert.ToString(row["Tracking ID Charge Description18"]);
            clsPedido.TrackingIDChargeAmount18 = Convert.ToString(row["Tracking ID Charge Amount18"]);
            clsPedido.TrackingIDChargeDescription19 = Convert.ToString(row["Tracking ID Charge Description19"]);
            clsPedido.TrackingIDChargeAmount19 = Convert.ToString(row["Tracking ID Charge Amount19"]);
            clsPedido.TrackingIDChargeDescription20 = Convert.ToString(row["Tracking ID Charge Description20"]);
            clsPedido.TrackingIDChargeAmount20 = Convert.ToString(row["Tracking ID Charge Amount20"]);
            clsPedido.TrackingIDChargeDescription21 = Convert.ToString(row["Tracking ID Charge Description21"]);
            clsPedido.TrackingIDChargeAmount21 = Convert.ToString(row["Tracking ID Charge Amount21"]);
            clsPedido.TrackingIDChargeDescription22 = Convert.ToString(row["Tracking ID Charge Description22"]);
            clsPedido.TrackingIDChargeAmount22 = Convert.ToString(row["Tracking ID Charge Amount22"]);
            clsPedido.TrackingIDChargeDescription23 = Convert.ToString(row["Tracking ID Charge Description23"]);
            clsPedido.TrackingIDChargeAmount23 = Convert.ToString(row["Tracking ID Charge Amount23"]);
            clsPedido.TrackingIDChargeDescription24 = Convert.ToString(row["Tracking ID Charge Description24"]);
            clsPedido.TrackingIDChargeAmount24 = Convert.ToString(row["Tracking ID Charge Amount24"]);
            clsPedido.ShipmentNotes = Convert.ToString(row["Shipment Notes"]);
        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneValorRegistroDetalleUSPS(DataRow row, ref PedidoUSPS clsPedido )
        {
            clsPedido.AccountNumber = Convert.ToString(row["Account Number"]);
            clsPedido.ID = Convert.ToString(row["ID"]);
            //clsPedido.DateTime = Convert.ToString(row["Date/Time"]);
            clsPedido.Postmark = Convert.ToString(row["Postmark"]);
            clsPedido.Origin = Convert.ToString(row["Origin"]);
            clsPedido.Destination = Convert.ToString(row["Destination"]);
            clsPedido.Type = Convert.ToString(row["Type"]);
            clsPedido.MailClass = Convert.ToString(row["Mail Class"]);
            clsPedido.TrackingNumber = Convert.ToString(row["Tracking Number"]);
            clsPedido.DeclaredValue = Convert.ToString(row["Declared Value"]);
            clsPedido.TotalPostageAmt = Convert.ToString(row["Total Postage Amt"]);
            clsPedido.Balance = Convert.ToString(row["Balance"]);
            clsPedido.RefundStatus = Convert.ToString(row["Refund Status"]);
            clsPedido.GroupCode = Convert.ToString(row["Group Code"]);
            clsPedido.ReferenceID = Convert.ToString(row["Reference ID"]);
            clsPedido.DeliveryDate = Convert.ToString(row["Delivery Date"]);
            clsPedido.StatusCode = Convert.ToString(row["Status Code"]);
            clsPedido.StatusDescription = Convert.ToString(row["Status Description"]);
            clsPedido.Weight = Convert.ToString(row["Weight"]);
            clsPedido.OptionalServices = Convert.ToString(row["OptionalServices"]);
            clsPedido.DestinationName = Convert.ToString(row["Destination Name"]);
            clsPedido.DestinationCompanyName = Convert.ToString(row["Destination Company Name"]);
            clsPedido.DestinationAddress = Convert.ToString(row["Destination Address"]);
            clsPedido.DestinationCity = Convert.ToString(row["Destination City"]);
            clsPedido.DestinationState = Convert.ToString(row["Destination State"]);
            clsPedido.DestinationZip = Convert.ToString(row["Destination Zip"]);
            clsPedido.DestinationCountry = Convert.ToString(row["Destination Country"]);
            clsPedido.Phone = Convert.ToString(row["Phone"]);
            clsPedido.Email = Convert.ToString(row["Email"]);
            clsPedido.Reference2 = Convert.ToString(row["Reference2"]);
            clsPedido.Reference3 = Convert.ToString(row["Reference3"]);
            clsPedido.Reference4 = Convert.ToString(row["Reference4"]);
            clsPedido.PackageDescription = Convert.ToString(row["Package Description"]);
            clsPedido.Zone = Convert.ToString(row["Zone"]);
            clsPedido.IsCubic = Convert.ToString(row["IsCubic"]);
            clsPedido.CubicValue = Convert.ToString(row["Cubic Value"]);
            //clsPedido.AdjWeight = Convert.ToString(row["Adj. Weight"]);
            //clsPedido.AdjDimensions = Convert.ToString(row["Adj. Dimensions"]);
            //clsPedido.AdjFromZIP = Convert.ToString(row["Adj. From ZIP"]);
            //clsPedido.AdjToZIP = Convert.ToString(row["Adj. To ZIP"]);
            //clsPedido.AdjMailClass = Convert.ToString(row["Adj. Mail Class"]);

        }

    // obtiene el valor del registro actual
    // ------------------------------------
    private void ObtieneValorRegistroDetalleUPS(string[] valor,/*DataRow row, */ref PedidoUPS clsPedido)
    {
            clsPedido.Campo1 = valor[0];
            clsPedido.Campo2 = valor[1];
            clsPedido.Campo3 = valor[2];
            clsPedido.Campo4 = valor[3];
            clsPedido.Campo5 = valor[4];
            clsPedido.Campo6 = valor[5];
            clsPedido.Campo7 = valor[6];
            clsPedido.Campo8 = valor[7];
            clsPedido.Campo9 = valor[8];
            clsPedido.Campo10 = valor[9];
            clsPedido.Campo11 = valor[10];
            clsPedido.Campo12 = valor[11];
            clsPedido.Campo13 = valor[12];
            clsPedido.Campo14 = valor[13];
            clsPedido.Campo15 = valor[14];
            clsPedido.Campo16 = valor[15];
            clsPedido.Campo17 = valor[16];
            clsPedido.Campo18 = valor[17];
            clsPedido.Campo19 = valor[18];
            clsPedido.Campo20 = valor[19];
            clsPedido.Campo21 = valor[20];
            clsPedido.Campo22 = valor[21];
            clsPedido.Campo23 = valor[22];
            clsPedido.Campo24 = valor[23];
            clsPedido.Campo25 = valor[24];
            clsPedido.Campo26 = valor[25];
            clsPedido.Campo27 = valor[26];
            clsPedido.Campo28 = valor[27];
            clsPedido.Campo29 = valor[28];
            clsPedido.Campo30 = valor[29];
            clsPedido.Campo31 = valor[30];
            clsPedido.Campo32 = valor[31];
            clsPedido.Campo33 = valor[32];
            clsPedido.Campo34 = valor[33];
            clsPedido.Campo35 = valor[34];
            clsPedido.Campo36 = valor[35];
            clsPedido.Campo37 = valor[36];
            clsPedido.Campo38 = valor[37];
            clsPedido.Campo39 = valor[38];
            clsPedido.Campo40 = valor[39];
            clsPedido.Campo41 = valor[40];
            clsPedido.Campo42 = valor[41];
        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneValorRegistroDetalleEstimatedDeliveryDate(string[] valor,/*DataRow row, */ref ProvisionMensual.Clases.EstimatedDeliveryDate clsPedido)
        {
            clsPedido.SalesOrderNumber = valor[0];
            clsPedido.SalesOrderDate = valor[1];
            clsPedido.SalesChannelName = valor[2];
            clsPedido.ShopifyDeliveryDate = valor[3];
            clsPedido.NONShopifyDeliveryDate = valor[4];
        }

            // obtiene el valor del registro actual
            // ------------------------------------
            private void ObtieneValorRegistroDetallePITNEYBOEWS(string[] valor,/*DataRow row, */ref ProvisionMensual.Clases.PITNEYBOEWS clsPedido)
        {
            clsPedido.transactionType 			= valor[0];
            clsPedido.amount                    = (float)Convert.ToDecimal(valor[1]);
            clsPedido.transactionId             = valor[2];
            clsPedido.transactionDateTime       = Convert.ToDateTime(valor[3]);
            clsPedido.parcelTrackingNumber      = valor[4];
            clsPedido.status                    = valor[5];
            clsPedido.statusDate                = Convert.ToDateTime(valor[6]);
            clsPedido.refundDenialReason        = valor[7];
            clsPedido.service                   = valor[8];
            clsPedido.zone                      = (float)Convert.ToDecimal(valor[9]);
            clsPedido.weightInOunces            = (float)Convert.ToDecimal(valor[10]);
            clsPedido.packageType               = valor[11];
            clsPedido.packageLengthInInches     = (float)Convert.ToDecimal(valor[12]);
            clsPedido.packageWidthInInches      = (float)Convert.ToDecimal(valor[13]);
            clsPedido.packageHeightInInches     = (float)Convert.ToDecimal(valor[14]);
            clsPedido.specialServices           = valor[15];
            clsPedido.originationAddress        = valor[16];
            clsPedido.destinationAddress        = valor[17];
            clsPedido.destinationCountry        = valor[18];
            clsPedido.adjustmentReason          = valor[19];
            clsPedido.adjustmentId              = (float)Convert.ToDecimal(valor[20]);
            clsPedido.refundRequestor           = valor[21];
            clsPedido.postageBalance            = (float)Convert.ToDecimal(valor[22]);
            clsPedido.merchantRatePlan          = valor[23];
            clsPedido.packageIndicator          = valor[24];
            clsPedido.internationalCountryGroup = valor[25];
            clsPedido.dimensionalWeightOz       = (float)Convert.ToDecimal(valor[26]);
            clsPedido.valueOfGoods              = (float)Convert.ToDecimal(valor[27]);
            clsPedido.description               = valor[28];
            clsPedido.inductionPostalCode       = (float)Convert.ToDecimal(valor[29]);
            clsPedido.customMessage1            = valor[30];
            clsPedido.customMessage2            = valor[31];
        }

            // obtiene el valor del registro actual
            // ------------------------------------
            private void ObtieneValorRegistroDetalleAmazon(string[] valor, ref PedidoAmazon clsPedido)
        {
            clsPedido.date = valor[0];
            clsPedido.datetime = valor[1];
            clsPedido.settlementid = valor[2];
            clsPedido.type = valor[3];
            clsPedido.orderid = valor[4];
            clsPedido.sku = valor[5];
            clsPedido.description = valor[6];
            clsPedido.quantity = valor[7];
            clsPedido.marketplace = valor[8];
            clsPedido.fulfillment = valor[9];
            clsPedido.ordercity = valor[10];
            clsPedido.orderstate = valor[11];
            clsPedido.orderpostal = valor[12];
            clsPedido.taxcollectionmodel = valor[13];
            clsPedido.productsales = valor[14];
            clsPedido.productsalestax = valor[15];
            clsPedido.shippingcredits = valor[16];
            clsPedido.shippingcreditstax = valor[17];
            clsPedido.giftwrapcredits = valor[18];
            clsPedido.giftwrapcreditstax = valor[19];
            clsPedido.promotionalrebates = valor[20];
            clsPedido.promotionalrebatestax = valor[21];
            clsPedido.marketplacewithheldtax = valor[22];
            clsPedido.sellingfees = valor[23];
            clsPedido.fbafees = valor[24];
            clsPedido.othertransactionfees = valor[25];
            clsPedido.other = valor[26];
            clsPedido.total = valor[27];

        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneValorRegistroDetalleMI15(string[] valor, ref PesosDimensiones.MI15 clsPedido)
        {
            clsPedido.SHIPPINGDATE = valor[0];
            clsPedido.MANIFESTDATE = valor[1];
            clsPedido.PACKAGEID = valor[2];
            clsPedido.USPSTRACKINGNUMBER = valor[3];
            clsPedido.SEQUENCE = valor[4];
            clsPedido.COSTCENTER1 = valor[5];
            clsPedido.COSTCENTER2 = valor[6];
            clsPedido.COSTCENTER3 = valor[7];
            clsPedido.BILLEDWEIGHT = valor[8];
            clsPedido.WEIGHTTYPE = valor[9];
            clsPedido.ZIP = valor[10];
            clsPedido.ZONE = valor[11];
            clsPedido.SERVICE = valor[12];
            clsPedido.UPSMI = valor[13];
            clsPedido.USPS = valor[14];
            clsPedido.SAVINGS = valor[15];
            clsPedido.OVERLABELEDUSPSTRACKING = valor[16];
            clsPedido.ERRORREASON = valor[17];

        }

        // obtiene el valor del registro de box
        // ------------------------------------
        private void ObtieneValorRegistroDetalleBOX(DataRow row, ref PesosDimensiones.Box clsPedido)
        {
            clsPedido.STATE = Convert.ToString(row["STATE"]);
            clsPedido.POSTALCODE = Convert.ToString(row["POSTALCODE"]);
            clsPedido.SHIPPER = Convert.ToString(row["SHIPPER"]);
            clsPedido.PROSHIP_SHIPDATE = Convert.ToString(row["PROSHIP_SHIPDATE"]);
            clsPedido.PACKAGING_PLAINTEXT = Convert.ToString(row["PACKAGING_PLAINTEXT"]);
            clsPedido.WEIGHT = Convert.ToString(row["WEIGHT"]);
            clsPedido.DIMENSIONS = Convert.ToString(row["DIMENSIONS"]);
            clsPedido.TRACKING_NUMBER = Convert.ToString(row["TRACKING_NUMBER"]);
            clsPedido.CCN_SAP_ORDER_NUMBER = Convert.ToString(row["CCN_SAP_ORDER_NUMBER"]);
            clsPedido.CCN_ORDER_NUMBER = Convert.ToString(row["CCN_ORDER_NUMBER"]);
            clsPedido.CCN_COMPANY_CODE = Convert.ToString(row["CCN_COMPANY_CODE"]);
            clsPedido.CCN_STR_NUM = Convert.ToString(row["CCN_STR_NUM"]);
            clsPedido.CCN_DELIVERY_NUMBER = Convert.ToString(row["CCN_DELIVERY_NUMBER"]);
            clsPedido.SHIPPER_SYMBOL = Convert.ToString(row["SHIPPER_SYMBOL"]);
            clsPedido.OrderDate = Convert.ToString(row["Order Date"]);
            clsPedido.PROSHIP_SERVICE_PLAINTEXT = Convert.ToString(row["PROSHIP_SERVICE_PLAINTEXT"]);
            clsPedido.CCN_SHIP_TEXT = Convert.ToString(row["CCN_SHIP_TEXT"]);
        }

        // obtiene el valor del registro de EJDDimensions
        // ----------------------------------------------
        private void ObtieneValorRegistroDetalleEJDDimensions(DataRow row, ref PesosDimensiones.EJDDimensions clsPedido)
        {
            clsPedido.EvpSku = Convert.ToString(row["Evp Sku"]);
            clsPedido.Title = Convert.ToString(row["Title"]);
            clsPedido.EJDSku = Convert.ToString(row["EJD Sku"]);
            clsPedido.EJDUomCode = Convert.ToString(row["EJD Uom Code"]);
            clsPedido.EJDUomQuantity = Convert.ToString(row["EJD Uom Quantity"]);
            clsPedido.Length = Convert.ToString(row["Length"]);
            clsPedido.Height = Convert.ToString(row["Height"]);
            clsPedido.Width = Convert.ToString(row["Width"]);
            clsPedido.Weight = Convert.ToString(row["Weight"]);
        }

        // obtiene el valor del registro de EJDDimensions
        // ----------------------------------------------
        private void ObtieneValorRegistroDetalleJensenDimensions(DataRow row, ref PesosDimensiones.JensenDimensions clsPedido)
        {
            clsPedido.EvpSku = Convert.ToString(row["Evp Sku"]);
            clsPedido.Title = Convert.ToString(row["Title"]);
            clsPedido.JensenSku = Convert.ToString(row["Jensen Sku"]);
            clsPedido.UomCode = Convert.ToString(row["UomCode"]);
            clsPedido.UomQuantity = Convert.ToString(row["UomQuantity"]);
            clsPedido.Length = Convert.ToString(row["Length"]);
            clsPedido.Height = Convert.ToString(row["Height"]);
            clsPedido.Width = Convert.ToString(row["Width"]);
            clsPedido.Weight = Convert.ToString(row["Weight"]);
        }

        // obtiene el valor del registro de EJDDimensions
        // ----------------------------------------------
        private void ObtieneValorRegistroCancelados(DataRow row, ref ProvisionMensual.Cancelados clsPedido)
        {
            clsPedido.OrderDate             = Convert.ToDateTime(row["Order Date"]);
            clsPedido.PONumber              = Convert.ToString(row["PO Number"]);
            clsPedido.Status                = Convert.ToString(row["Status"]);
            clsPedido.Notes                 = Convert.ToString(row["Notes"]);
            clsPedido.Supplier              = Convert.ToString(row["Supplier"]);
            clsPedido.SupplierNumber        = Convert.ToString(row["Supplier Number"]);
            clsPedido.SupplierStatus         = Convert.ToString(row["Supplier Status"]);
            //clsPedido.Vacia                 = Convert.ToString(row["Supplier Status"]);
            clsPedido.ShipmentCount         = Convert.ToString(row["Shipment Count"]);
            clsPedido.Type                  = Convert.ToString(row["Type"]);
            clsPedido.PurchaseLocations     = Convert.ToString(row["Purchase Locations"]);
            clsPedido.ReceiveLocations      = Convert.ToString(row["Receive Locations"]);
            clsPedido.ItemSummary           = Convert.ToString(row["ItemSummary"]);
            clsPedido.ShippingServiceLevel  = Convert.ToString(row["ShippingServiceLevel"]);
            clsPedido.ShipTo                = Convert.ToString(row["Ship To"]);
            clsPedido.City                  = Convert.ToString(row["City"]);
            clsPedido.State                 = Convert.ToString(row["State"]);
            clsPedido.Country               = Convert.ToString(row["Country"]);
            clsPedido.PostalCode            = Convert.ToString(row["Postal Code"]);
            clsPedido.TotalWeight           = (float)Convert.ToDouble(row["TotalWeight"]);
            clsPedido.CreatedDate           = Convert.ToDateTime(row["Created Date"]);
            clsPedido.ExpectedDate          = Convert.ToDateTime(row["Expected Date"]);
            clsPedido.Total                 = (float)Convert.ToDouble(row["Total"]);
            clsPedido.FechaInsercion        = DateTime.Today; 
        }

        // Carga Archivo FEDEX
        // -------------------
        private void ObtieneDatosFedex(SqlConnection Conexion)
        {
            ProvisionMensualMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ProvisionMensualMaxwarehouse.Clases.ManejoBD();
            clsInsertaRegistro.EliminaRegistroFEDEX(Conexion);

            // abre todos los archivo secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------

            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString + @"\" + ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos secundarios Fedex " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                CreateText.WriteLine(Contenido);
                Contenido = "";
            }


            foreach (var Archivos in Directorios.GetFiles())
            {
                textBox1.Text = "Procesando Archivo: " + Archivos.Name;
                this.Refresh();
                this.Invalidate();

                //timer1.Enabled = true;
                //
                //if (progressBar1.Value == progressBar1.Maximum)
                //{
                //    progressBar1.Value = 0;
                //    timer1.Enabled = false;
                //}

                // obtiene datos del excel base
                // ----------------------------
                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["NombreArchivoBase"];

                if (file == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivoBase"];

                if (Archivos.Name == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContien"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }
                FlgSihayFedex = true;

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                }

                // obtiene record set
                // ------------------
                string connectionString = GetConnectionString(file, "DETALLE");

                var dataSet = GetDataSetFromExcelFileDetalle(file, connectionString);
                int conteoregistros = 0;

                // recorre registros obtenidos por la lectura del excel
                // ----------------------------------------------------
                foreach (DataRow row in dataSet.Tables[0].Rows)
                {
                    PedidoFedex clsPedido = new PedidoFedex();
                    // obtiene el valor del registro leido
                    // -----------------------------------
                    ObtieneValorRegistroDetalle(row, ref clsPedido);

                    //listaPedido.Add(clsPedido);
                    clsInsertaRegistro.InsertaBDFEDEX(clsPedido, Conexion);
                    conteoregistros += 1;
                }

                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString + @"\" + Archivos.Name;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);


            }

            textBox1.Text = "Fin Carga Archivos Fedex";
            this.Refresh();
            this.Invalidate();
        }


        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneDatosUSPS(SqlConnection Conexion)
        {
            ProvisionMensualMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ProvisionMensualMaxwarehouse.Clases.ManejoBD();
            clsInsertaRegistro.EliminaRegistroUSPS(Conexion);

            // abre todos los archivo secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            //string ArchivosSecundarios = ConfigurationManager.AppSettings["CarpetaArchivosSecundarios"];
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------
            //pathString = ArchivosSecundarios + "Output" + DateTime.Now.ToString("yyyyMMddTHHmmss");
            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString + @"\" + ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos secundarios USPS " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                CreateText.WriteLine(Contenido);
            }



            foreach (var Archivos in Directorios.GetFiles())
            {
                textBox1.Text = "Procesando Archivo: " + Archivos.Name;
                this.Refresh();
                this.Invalidate();

                //timer1.Enabled = true;
                //
                //if (progressBar1.Value == progressBar1.Maximum)
                //{
                //    progressBar1.Value = 0;
                //    timer1.Enabled = false;
                //}

                // obtiene datos del excel base
                // ----------------------------
                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["NombreArchivoBase"];

                if (file == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivoBase"];

                if (Archivos.Name == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContien"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienUSPS"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                FlgSihayUSPS = true;

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                }

                // obtiene record set
                // ------------------
                string connectionString = GetConnectionString(file, "DETALLE");

                var dataSet = GetDataSetFromExcelFileDetalle(file, connectionString);
                int conteoregistros = 0;

                // recorre registros obtenidos por la lectura del excel
                // ----------------------------------------------------
                foreach (DataRow row in dataSet.Tables[0].Rows)
                {
                    PedidoUSPS clsPedido = new PedidoUSPS();
                    // obtiene el valor del registro leido
                    // -----------------------------------
                    ObtieneValorRegistroDetalleUSPS(row, ref clsPedido);

                    // Valida no insertar registros de footer del archivo
                    // --------------------------------------------------
                    if (clsPedido.AccountNumber == "" || clsPedido.AccountNumber == "Transaction Type" || clsPedido.AccountNumber == "Adjustment" || clsPedido.AccountNumber == "Postage Print" || clsPedido.AccountNumber == "Postage Purchase" || clsPedido.AccountNumber == "Postage Refund")
                        continue;

                    clsInsertaRegistro.InsertaBDUSPS(clsPedido, Conexion);
                    //listaPedidoUSPS.Add(clsPedido);
                    conteoregistros += 1;
                }

                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString + @"\" + Archivos.Name;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);


            }

            textBox1.Text = "Fin Carga Archivos USPS";
            this.Refresh();
            this.Invalidate();
        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneDatosUPS(SqlConnection Conexion)
        {
            ProvisionMensualMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ProvisionMensualMaxwarehouse.Clases.ManejoBD();
            clsInsertaRegistro.EliminaRegistroUPS(Conexion);

            // abre todos los archiv++-o secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            //string ArchivosSecundarios = ConfigurationManager.AppSettings["CarpetaArchivosSecundarios"];
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------
            //pathString = ArchivosSecundarios + "Output" + DateTime.Now.ToString("yyyyMMddTHHmmss");
            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString + @"\" + ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos secundarios UPS " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                CreateText.WriteLine(Contenido);
            }

            foreach (var Archivos in Directorios.GetFiles())
            {
                textBox1.Text = "Procesando Archivo: " + Archivos.Name;
                this.Refresh();
                this.Invalidate();

                //timer1.Enabled = true;
                //
                //if (progressBar1.Value == progressBar1.Maximum)
                //{
                //    progressBar1.Value = 0;
                //    timer1.Enabled = false;
                //}

                // obtiene datos del excel base
                // ----------------------------
                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["NombreArchivoBase"];

                if (file == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivoBase"];

                if (Archivos.Name == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContien"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienUSPS"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                //filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienMI15"];
                //if (filemaster != "")
                //{
                //    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                //        continue;
                //}

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienUPS"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                FlgSihayUPS = true;

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                }

                // Read the file and display it line by line.  
                System.IO.StreamReader upsFile = new System.IO.StreamReader(file);
                string[] PalabraUps = new string[42];
                int contPala = 0;
                int conteoregistros = 0;

                while ((line = upsFile.ReadLine()) != null)
                {


                    string[] sa = line.Split(',');

                    if (sa.Length > 42)
                    {
                        if (!line.Contains("MISCELLANEOUS"))
                            //cantidad = sa.Length - 42;
                            continue;
                    }

                    contPala = 0;
                    foreach (string s in sa)
                    {
                        if (contPala < 42)
                        {
                            PalabraUps[contPala] = s;
                            contPala = contPala + 1;
                        }
                    }

                    PedidoUPS clsPedido = new PedidoUPS();
                    ObtieneValorRegistroDetalleUPS(PalabraUps, ref clsPedido);
                    //listaPedidoUPS.Add(clsPedido);
                    clsInsertaRegistro.InsertaBDUPS(clsPedido, Conexion);
                    conteoregistros += 1;
                }

                upsFile.Close();

                // obtiene record set
                // ------------------
                //string connectionString = GetConnectionString(file, "DETALLE");
                //var dataSet = GetDataSetFromExcelFileDetalle(file, connectionString);
                //int conteoregistros = 0;
                //
                //// recorre registros obtenidos por la lectura del excel
                //// ----------------------------------------------------
                //foreach (DataRow row in dataSet.Tables[0].Rows)
                //{
                //    PedidoUPS clsPedido = new PedidoUPS();
                //    // obtiene el valor del registro leido
                //    // -----------------------------------
                //    ObtieneValorRegistroDetalleUPS(row, ref clsPedido);
                //
                //    listaPedidoUPS.Add(clsPedido);
                //    conteoregistros += 1;
                //}

                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString + @"\" + Archivos.Name;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);


            }

            textBox1.Text = "Fin Carga Archivos UPS";
            this.Refresh();
            this.Invalidate();
        }



        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneDatosPITNEYBOWES(SqlConnection Conexion)
        {
            ProvisionMensualMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ProvisionMensualMaxwarehouse.Clases.ManejoBD();
            clsInsertaRegistro.EliminaRegistroPITNEYBOWES(Conexion);

            // abre todos los archiv++-o secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            //string ArchivosSecundarios = ConfigurationManager.AppSettings["CarpetaArchivosSecundarios"];
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------
            //pathString = ArchivosSecundarios + "Output" + DateTime.Now.ToString("yyyyMMddTHHmmss");
            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString + @"\" + ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos secundarios PITNEYBOWES " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                CreateText.WriteLine(Contenido);
            }

            foreach (var Archivos in Directorios.GetFiles())
            {
                textBox1.Text = "Procesando Archivo: " + Archivos.Name;
                this.Refresh();
                this.Invalidate();

                //timer1.Enabled = true;
                //
                //if (progressBar1.Value == progressBar1.Maximum)
                //{
                //    progressBar1.Value = 0;
                //    timer1.Enabled = false;
                //}

                // obtiene datos del excel base
                // ----------------------------
                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["NombreArchivoBase"];

                if (file == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivoBase"];

                if (Archivos.Name == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContien"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienUSPS"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                //filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienMI15"];
                //if (filemaster != "")
                //{
                //    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                //        continue;
                //}

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienPITNEYBOWES"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                FlgSihayPITNEYBOWES = true;

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                }

                // Read the file and display it line by line.  
                System.IO.StreamReader PITNEYBOWESFile = new System.IO.StreamReader(file);
                string[] PalabraPITNEYBOWES = new string[32];
                int contPala = 0;
                int conteoregistros = 0;

                while ((line = PITNEYBOWESFile.ReadLine()) != null)
                {
                    string[] sa = line.Split(',');
                    contPala = 0;
                    foreach (string s in sa)
                    {
                        if (contPala < 42)
                        {
                            PalabraPITNEYBOWES[contPala] = s;
                            contPala = contPala + 1;
                        }
                    }

                    ProvisionMensual.Clases.PITNEYBOEWS clsPedido = new ProvisionMensual.Clases.PITNEYBOEWS();
                    ObtieneValorRegistroDetallePITNEYBOEWS(PalabraPITNEYBOWES, ref clsPedido);
                    //listaPedidoUPS.Add(clsPedido);
                    clsInsertaRegistro.InsertaBDPITNEYBOEWS(clsPedido, Conexion);
                    conteoregistros += 1;
                }

                PITNEYBOWESFile.Close();

 
                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString + @"\" + Archivos.Name;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);


            }

            textBox1.Text = "Fin Carga Archivos UPS";
            this.Refresh();
            this.Invalidate();
        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneDatosEstimatedDeliveryDate(SqlConnection Conexion)
        {
            ProvisionMensualMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ProvisionMensualMaxwarehouse.Clases.ManejoBD();
            //clsInsertaRegistro.EliminaRegistroEstimatedDeliveryDate(Conexion);

            // abre todos los archiv++-o secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            //string ArchivosSecundarios = ConfigurationManager.AppSettings["CarpetaArchivosSecundarios"];
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------
            //pathString = ArchivosSecundarios + "Output" + DateTime.Now.ToString("yyyyMMddTHHmmss");
            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString + @"\" + ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos secundarios EstimatedDeliveryDate " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                CreateText.WriteLine(Contenido);
            }

            foreach (var Archivos in Directorios.GetFiles())
            {
                textBox1.Text = "Procesando Archivo: " + Archivos.Name;
                this.Refresh();
                this.Invalidate();

                //timer1.Enabled = true;
                //
                //if (progressBar1.Value == progressBar1.Maximum)
                //{
                //    progressBar1.Value = 0;
                //    timer1.Enabled = false;
                //}

                // obtiene datos del excel base
                // ----------------------------
                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["NombreArchivoBase"];

                if (file == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivoBase"];

                if (Archivos.Name == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContien"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienUSPS"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                //filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienMI15"];
                //if (filemaster != "")
                //{
                //    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                //        continue;
                //}

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienEstimatedDeliveryDate"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                FlgSihayEstimatedDeliveryDate = true;

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                }

                // Read the file and display it line by line.  
                System.IO.StreamReader EstimatedDeliveryDateFile = new System.IO.StreamReader(file);
                string[] PalabraEstimatedDeliveryDate = new string[5];
                int contPala = 0;
                int conteoregistros = 0;

                while ((line = EstimatedDeliveryDateFile.ReadLine()) != null)
                {
                    string[] sa = line.Split(',');
                    contPala = 0;
                    foreach (string s in sa)
                    {
                        if (contPala < 5)
                        {
                            PalabraEstimatedDeliveryDate[contPala] = s;
                            contPala = contPala + 1;
                        }
                    }

                    ProvisionMensual.Clases.EstimatedDeliveryDate clsPedido = new ProvisionMensual.Clases.EstimatedDeliveryDate();
                    ObtieneValorRegistroDetalleEstimatedDeliveryDate(PalabraEstimatedDeliveryDate, ref clsPedido);
                    clsInsertaRegistro.InsertaBDEstimatedDeliveryDate(clsPedido, Conexion);
                    conteoregistros += 1;
                }

                EstimatedDeliveryDateFile.Close();


                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString + @"\" + Archivos.Name;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);


            }

            textBox1.Text = "Fin Carga Archivos UPS";
            this.Refresh();
            this.Invalidate();
        }


        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneValorRegistroDetalleEndicia(string[] valor, ref PedidoEndicia clsPedido)
        {
            clsPedido.PrintDate = Convert.ToDateTime(valor[0]);
            clsPedido.AmountPaid = System.Convert.ToDecimal(valor[1]);
            clsPedido.AdjAmount = valor[2];
            clsPedido.QuotedAmount = System.Convert.ToDecimal(valor[3]);
            clsPedido.OriginZip = valor[4];
            clsPedido.Recipient = valor[5];
            clsPedido.Status = valor[6];
            clsPedido.TrackingNumber = valor[7];

            clsPedido.TrackingNumber = clsPedido.TrackingNumber.Replace("=", "");

            if (valor[7].Contains("/"))
                clsPedido.DateDelivered = Convert.ToDateTime(valor[8]);
            else
                clsPedido.DateDelivered = new DateTime(1000, 1, 1);

            clsPedido.Carrier = valor[9];
            clsPedido.ClassService = valor[10];
            clsPedido.ExtraServices = valor[11];
            //clsPedido.InsuredValue = System.Convert.ToDecimal(valor[12]);
            clsPedido.InsuranceID = valor[13];
            clsPedido.CostCode = valor[14];
            clsPedido.Weight = valor[15];

            if (valor[14].Contains("/"))
                clsPedido.ShipDate = Convert.ToDateTime(valor[16]);
            else
                clsPedido.ShipDate = new DateTime(1000, 1, 1);

            clsPedido.RefundType = valor[17];
            clsPedido.PrintedMessage = valor[18];
            clsPedido.User = valor[19];

            if (valor[18].Contains("/"))
                clsPedido.RefundRequestDate = Convert.ToDateTime(valor[20]);
            else
                clsPedido.RefundRequestDate = new DateTime(1000, 1, 1);

            clsPedido.RefundStatus = valor[21];
            clsPedido.RefundRequested = valor[22];
            clsPedido.Reference1 = valor[23];
            clsPedido.Reference2 = valor[24];
            clsPedido.Reference3 = valor[25];
            clsPedido.Reference4 = valor[26];
            clsPedido.OrderID = valor[27];

        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneDatosAmazon(SqlConnection Conexion)
        {
            ProvisionMensualMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ProvisionMensualMaxwarehouse.Clases.ManejoBD();
            clsInsertaRegistro.EliminaRegistroAmazon(Conexion);
            clsInsertaRegistro.EliminaRegistroAmazonRefunded(Conexion);

            // abre todos los archivo secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            //string ArchivosSecundarios = ConfigurationManager.AppSettings["CarpetaArchivosSecundarios"];
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------
            //pathString = ArchivosSecundarios + "Output" + DateTime.Now.ToString("yyyyMMddTHHmmss");
            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString + @"\" + ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos secundarios amazon " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                CreateText.WriteLine(Contenido);
            }

            int ContadorArchivos = 0;
            foreach (var Archivos in Directorios.GetFiles())
            {
                textBox1.Text = "Procesando Archivo: " + Archivos.Name;
                this.Refresh();
                this.Invalidate();


                // obtiene datos del excel base
                // ----------------------------
                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["NombreArchivoBase"];

                if (file == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivoBase"];

                if (Archivos.Name == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContien"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienUSPS"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienUPS"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienAmazonOriginal"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                FlgSihayAmazon = true;

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                }


                // elimina las primeras 7 lineas ya que son encabezado de amazon
                int Contadorlineas = 0;
                string RutaAmazon = ConfigurationManager.AppSettings["CarpetaArchivosSecundarios"];
                RutaAmazon = RutaAmazon + "Amazon_" + ContadorArchivos.ToString() + ".csv";
                int Inicio = 0;
                string PalabraSinComa = "";
                string PalabraFinal = "";

                using (StreamWriter fileWrite = new StreamWriter(RutaAmazon))
                {
                    using (StreamReader fielRead = new StreamReader(Archivos.FullName))
                    {
                        String line;

                        while ((line = fielRead.ReadLine()) != null)
                        {
                            if (Contadorlineas <= 7)
                            {
                                Contadorlineas = Contadorlineas + 1;
                                continue;
                            }
                            else
                            {

                                PalabraFinal = "";
                                string NuevaLinea = line.Replace(@"""""", "");

                                if (NuevaLinea.Contains(@""""))
                                {
                                    int Posicion = NuevaLinea.IndexOf(@"""", Inicio);

                                    int PosicionFinal = NuevaLinea.IndexOf(@"""", Posicion + 1);

                                    for (int ii = 0; ii < NuevaLinea.Length; ii++)
                                    {
                                        if (ii >= Posicion && ii <= PosicionFinal)
                                        {
                                            PalabraSinComa = NuevaLinea.Substring(ii, 1);

                                            if (PalabraSinComa == @"""" || PalabraSinComa == @",")
                                            {
                                                continue;
                                            }
                                            else
                                            {

                                                PalabraFinal = PalabraFinal + NuevaLinea.Substring(ii, 1);
                                            }
                                        }
                                        else
                                        {
                                            PalabraFinal = PalabraFinal + NuevaLinea.Substring(ii, 1);
                                        }
                                    }
                                }
                                else
                                {
                                    PalabraFinal = line;
                                }

                                // busca siguiente coincidencia de ,
                                // ---------------------------------
                                string PalabraFinal2 = "";
                                if (PalabraFinal.Contains(@""""))
                                {
                                    int Posicion = PalabraFinal.IndexOf(@"""", Inicio);

                                    int PosicionFinal = PalabraFinal.IndexOf(@"""", Posicion + 1);

                                    for (int ii = 0; ii < PalabraFinal.Length; ii++)
                                    {
                                        if (ii >= Posicion && ii <= PosicionFinal)
                                        {
                                            PalabraSinComa = PalabraFinal.Substring(ii, 1);

                                            if (PalabraSinComa == @"""" || PalabraSinComa == @",")
                                            {
                                                continue;
                                            }
                                            else
                                            {

                                                PalabraFinal2 = PalabraFinal2 + PalabraFinal.Substring(ii, 1);
                                            }
                                        }
                                        else
                                        {
                                            PalabraFinal2 = PalabraFinal2 + PalabraFinal.Substring(ii, 1);
                                        }
                                    }
                                }
                                else
                                {
                                    PalabraFinal2 = PalabraFinal;
                                }

                                // busca siguiente coincidencia de ,
                                // ---------------------------------
                                string PalabraFinal3 = "";
                                if (PalabraFinal2.Contains(@""""))
                                {
                                    int Posicion = PalabraFinal2.IndexOf(@"""", Inicio);

                                    int PosicionFinal = PalabraFinal2.IndexOf(@"""", Posicion + 1);

                                    for (int ii = 0; ii < PalabraFinal2.Length; ii++)
                                    {
                                        if (ii >= Posicion && ii <= PosicionFinal)
                                        {
                                            PalabraSinComa = PalabraFinal2.Substring(ii, 1);

                                            if (PalabraSinComa == @"""" || PalabraSinComa == @",")
                                            {
                                                continue;
                                            }
                                            else
                                            {

                                                PalabraFinal3 = PalabraFinal3 + PalabraFinal2.Substring(ii, 1);
                                            }
                                        }
                                        else
                                        {
                                            PalabraFinal3 = PalabraFinal3 + PalabraFinal2.Substring(ii, 1);
                                        }
                                    }
                                }
                                else
                                {
                                    PalabraFinal3 = PalabraFinal2;
                                }

                                // busca siguiente coincidencia de ,
                                // ---------------------------------
                                string PalabraFinal4 = "";
                                if (PalabraFinal3.Contains(@""""))
                                {
                                    int Posicion = PalabraFinal3.IndexOf(@"""", Inicio);

                                    int PosicionFinal = PalabraFinal3.IndexOf(@"""", Posicion + 1);

                                    for (int ii = 0; ii < PalabraFinal3.Length; ii++)
                                    {
                                        if (ii >= Posicion && ii <= PosicionFinal)
                                        {
                                            PalabraSinComa = PalabraFinal3.Substring(ii, 1);

                                            if (PalabraSinComa == @"""" || PalabraSinComa == @",")
                                            {
                                                continue;
                                            }
                                            else
                                            {

                                                PalabraFinal4 = PalabraFinal4 + PalabraFinal3.Substring(ii, 1);
                                            }
                                        }
                                        else
                                        {
                                            PalabraFinal4 = PalabraFinal4 + PalabraFinal3.Substring(ii, 1);
                                        }
                                    }
                                }
                                else
                                {
                                    PalabraFinal4 = PalabraFinal3;
                                }

                                string[] sa = PalabraFinal4.Split(',');
                                if (sa.Length > 47)
                                {
                                    cantidad = sa.Length - 47;
                                    continue;
                                }
                                fileWrite.WriteLine(PalabraFinal4);
                            }
                        }

                        fielRead.Close();
                    }

                    fileWrite.Close();
                }

                // Read the file and display it line by line.  
                System.IO.StreamReader amazonFile = new System.IO.StreamReader(RutaAmazon);

                string[] PalabraAmazon = new string[29];
                int contPala = 0;
                int conteoregistros = 0;

                while ((line = amazonFile.ReadLine()) != null)
                {
                    string LineaSinCaracteres = line.Replace("\"", "");
                    LineaSinCaracteres = LineaSinCaracteres.Replace("/", ",");
                    string[] sa = LineaSinCaracteres.Split(',');

                    if (sa.Length > 29)
                    {
                        //cantidad = sa.Length - 42;
                        continue;
                    }

                    contPala = 0;
                    foreach (string s in sa)
                    {
                        PalabraAmazon[contPala] = s;
                        contPala = contPala + 1;
                    }

                    // if (PalabraAmazon[2] != "Shipping Services")
                    //     continue;

                    if (PalabraAmazon[2] != "Shipping Services")
                    {
                        if (PalabraAmazon[2] == "Refund")
                        {

                            PedidoAmazon clsPedidorefund = new PedidoAmazon();
                            ObtieneValorRegistroDetalleAmazon(PalabraAmazon, ref clsPedidorefund);
                            //listaPedidoAmazon.Add(clsPedido);
                            clsInsertaRegistro.InsertaBDAMAZON(clsPedidorefund, Conexion, true);
                            conteoregistros += 1;
                        }

                        continue;
                    }

                    PedidoAmazon clsPedido = new PedidoAmazon();
                    ObtieneValorRegistroDetalleAmazon(PalabraAmazon, ref clsPedido);
                    //listaPedidoAmazon.Add(clsPedido);
                    clsInsertaRegistro.InsertaBDAMAZON(clsPedido, Conexion);
                    conteoregistros += 1;
                }

                amazonFile.Close();

                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString + @"\" + Archivos.Name;
                string RutaArchivoMoverAmazon = pathString + @"\" + "Amazon_" + ContadorArchivos.ToString() + ".csv"; ;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);
                System.IO.File.Move(RutaAmazon, RutaArchivoMoverAmazon);
                ContadorArchivos += 1;

            }

            textBox1.Text = "Fin Carga Archivos Amazon";
            this.Refresh();
            this.Invalidate();
        }

        private void ObtieneDatosEndicia(SqlConnection Conexion)
        {
            ProvisionMensualMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ProvisionMensualMaxwarehouse.Clases.ManejoBD();
            clsInsertaRegistro.EliminaRegistroEndicia(Conexion);

            // abre todos los archivo secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            //string ArchivosSecundarios = ConfigurationManager.AppSettings["CarpetaArchivosSecundarios"];
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------
            //pathString = ArchivosSecundarios + "Output" + DateTime.Now.ToString("yyyyMMddTHHmmss");
            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString + @"\" + ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos secundarios Endicia " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                CreateText.WriteLine(Contenido);
            }



            foreach (var Archivos in Directorios.GetFiles())
            {
                textBox1.Text = "Procesando Archivo: " + Archivos.Name;
                this.Refresh();
                this.Invalidate();

                // obtiene datos del excel base
                // ----------------------------
                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["NombreArchivoBase"];

                if (file == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivoBase"];

                if (Archivos.Name == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContien"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienEndicia"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                FlgSihayEndicia = true;

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                }

                // Read the file and display it line by line.  
                System.IO.StreamReader EndiciaFile = new System.IO.StreamReader(file);
                string[] PalabraEndicia = new string[29];
                int contPala = 0;
                int conteoregistros = 0;
                int Inicio = 0;
                string PalabraSinComa = "";
                string PalabraFinal = "";

                while ((line = EndiciaFile.ReadLine()) != null)
                {
                    PalabraFinal = "";
                    string NuevaLinea = line.Replace(@"""""", "");

                    if (NuevaLinea.Contains(@""""))
                    {
                        int Posicion = NuevaLinea.IndexOf(@"""", Inicio);

                        int PosicionFinal = NuevaLinea.IndexOf(@"""", Posicion + 1);

                        for (int ii = 0; ii < NuevaLinea.Length; ii++)
                        {
                            if (ii >= Posicion && ii <= PosicionFinal)
                            {
                                PalabraSinComa = NuevaLinea.Substring(ii, 1);

                                if (PalabraSinComa == @"""" || PalabraSinComa == @",")
                                {
                                    continue;
                                }
                                else
                                {

                                    PalabraFinal = PalabraFinal + NuevaLinea.Substring(ii, 1);
                                }
                            }
                            else
                            {
                                PalabraFinal = PalabraFinal + NuevaLinea.Substring(ii, 1);
                            }
                        }
                    }
                    else
                    {
                        PalabraFinal = line;
                    }

                    // busca siguiente coincidencia de ,
                    // ---------------------------------
                    string PalabraFinal2 = "";
                    if (PalabraFinal.Contains(@""""))
                    {
                        int Posicion = PalabraFinal.IndexOf(@"""", Inicio);

                        int PosicionFinal = PalabraFinal.IndexOf(@"""", Posicion + 1);

                        for (int ii = 0; ii < PalabraFinal.Length; ii++)
                        {
                            if (ii >= Posicion && ii <= PosicionFinal)
                            {
                                PalabraSinComa = PalabraFinal.Substring(ii, 1);

                                if (PalabraSinComa == @"""" || PalabraSinComa == @",")
                                {
                                    continue;
                                }
                                else
                                {

                                    PalabraFinal2 = PalabraFinal2 + PalabraFinal.Substring(ii, 1);
                                }
                            }
                            else
                            {
                                PalabraFinal2 = PalabraFinal2 + PalabraFinal.Substring(ii, 1);
                            }
                        }
                    }
                    else
                    {
                        PalabraFinal2 = PalabraFinal;
                    }

                    PalabraFinal2 = PalabraFinal2.Replace("$", "");
                    PalabraFinal2 = PalabraFinal2.Replace("\"", "");
                    string[] sa = PalabraFinal2.Split(',');

                    if (sa.Length > 29)
                    {
                        continue;
                    }

                    contPala = 0;
                    foreach (string s in sa)
                    {
                        string Valor = s.Replace("\"", "");

                        Valor = Valor.Replace("<None>", "");
                        PalabraEndicia[contPala] = Valor;
                        contPala = contPala + 1;
                    }

                    if (line.Contains("Print Date"))
                        continue;

                    PedidoEndicia clsPedido = new PedidoEndicia();
                    ObtieneValorRegistroDetalleEndicia(PalabraEndicia, ref clsPedido);
                    clsInsertaRegistro.InsertaBDEndicia(clsPedido, Conexion);
                    conteoregistros += 1;
                }

                EndiciaFile.Close();

                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString + @"\" + Archivos.Name;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);


            }

            textBox1.Text = "Fin Carga Archivos Endicia";
            this.Refresh();
            this.Invalidate();
        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void CargaDatosBOX(SqlConnection Conexion)
        {
            ProvisionMensualMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ProvisionMensualMaxwarehouse.Clases.ManejoBD();
            //clsInsertaRegistro.EliminaRegistroBOX(Conexion);

            // abre todos los archivo secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------
            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString + @"\" + ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos BOX " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                CreateText.WriteLine(Contenido);
                Contenido = "";
            }


            foreach (var Archivos in Directorios.GetFiles())
            {
                //textBox1.Text = "Procesando Archivo: " + Archivos.Name;
                this.Refresh();
                this.Invalidate();

                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["ArchivoContieneBOX"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                }

                // obtiene record set
                // ------------------
                string connectionString = GetConnectionString(file, "DETALLE");

                var dataSet = GetDataSetFromExcelFileDetalle(file, connectionString);
                int conteoregistros = 0;

                // recorre registros obtenidos por la lectura del excel
                // ----------------------------------------------------
                foreach (DataRow row in dataSet.Tables[0].Rows)
                {
                    PesosDimensiones.Box clsPedido = new PesosDimensiones.Box();

                    // obtiene el valor del registro leido
                    // -----------------------------------
                    ObtieneValorRegistroDetalleBOX(row, ref clsPedido);

                    //listaPedido.Add(clsPedido);
                    clsInsertaRegistro.InsertaBOX(clsPedido, Conexion);
                    conteoregistros += 1;
                }

                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString + @"\" + Archivos.Name;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);


            }

            //textBox1.Text = "Fin Carga Archivos Fedex";
            this.Refresh();
            this.Invalidate();
        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void CargaDatosEJDDimensions(SqlConnection Conexion)
        {
            ProvisionMensualMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ProvisionMensualMaxwarehouse.Clases.ManejoBD();
            clsInsertaRegistro.EliminaRegistroEJDDimensions(Conexion);

            // abre todos los archivo secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------

            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString + @"\" + ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos EJDDimensions " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                CreateText.WriteLine(Contenido);
                Contenido = "";
            }


            foreach (var Archivos in Directorios.GetFiles())
            {
                //textBox1.Text = "Procesando Archivo: " + Archivos.Name;
                this.Refresh();
                this.Invalidate();

                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["ArchivoContieneEJDDimensions"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                }

                // obtiene record set
                // ------------------
                string connectionString = GetConnectionString(file, "DETALLE");

                var dataSet = GetDataSetFromExcelFileDetalle(file, connectionString);
                int conteoregistros = 0;

                // recorre registros obtenidos por la lectura del excel
                // ----------------------------------------------------
                foreach (DataRow row in dataSet.Tables[0].Rows)
                {
                    PesosDimensiones.EJDDimensions clsPedido = new PesosDimensiones.EJDDimensions();

                    // obtiene el valor del registro leido
                    // -----------------------------------
                    ObtieneValorRegistroDetalleEJDDimensions(row, ref clsPedido);

                    //listaPedido.Add(clsPedido);
                    clsInsertaRegistro.InsertaEJDDimensions(clsPedido, Conexion);
                    conteoregistros += 1;
                }

                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString + @"\" + Archivos.Name;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);


            }

            //textBox1.Text = "Fin Carga Archivos EJD";
            this.Refresh();
            this.Invalidate();
        }


        // Cargadatos excel Jensen
        // -----------------------
        private void CargaDatosJensenDimensions(SqlConnection Conexion)
        {
            ProvisionMensualMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ProvisionMensualMaxwarehouse.Clases.ManejoBD();
            clsInsertaRegistro.EliminaRegistroJensenDimensions(Conexion);

            // abre todos los archivo secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------

            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString + @"\" + ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos JensenDimensions " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                CreateText.WriteLine(Contenido);
                Contenido = "";
            }


            foreach (var Archivos in Directorios.GetFiles())
            {
                //textBox1.Text = "Procesando Archivo: " + Archivos.Name;
                this.Refresh();
                this.Invalidate();

                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["ArchivoContieneJensenDimensions"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                }

                // obtiene record set
                // ------------------
                string connectionString = GetConnectionString(file, "DETALLE");

                var dataSet = GetDataSetFromExcelFileDetalle(file, connectionString);
                int conteoregistros = 0;

                // recorre registros obtenidos por la lectura del excel
                // ----------------------------------------------------
                foreach (DataRow row in dataSet.Tables[0].Rows)
                {
                    PesosDimensiones.JensenDimensions clsPedido = new PesosDimensiones.JensenDimensions();

                    // obtiene el valor del registro leido
                    // -----------------------------------
                    ObtieneValorRegistroDetalleJensenDimensions(row, ref clsPedido);

                    //listaPedido.Add(clsPedido);
                    clsInsertaRegistro.InsertaJensenDimensions(clsPedido, Conexion);
                    conteoregistros += 1;
                }

                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString + @"\" + Archivos.Name;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);


            }

            //textBox1.Text = "Fin Carga Archivos JensenDimensions";
            this.Refresh();
            this.Invalidate();
        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void ObtieneDatosMI15(SqlConnection Conexion)
        {
            ProvisionMensualMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ProvisionMensualMaxwarehouse.Clases.ManejoBD();
            //clsInsertaRegistro.EliminaRegistroM15(Conexion);

            // abre todos los archivo secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            //string ArchivosSecundarios = ConfigurationManager.AppSettings["CarpetaArchivosSecundarios"];
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------
            //pathString = ArchivosSecundarios + "Output" + DateTime.Now.ToString("yyyyMMddTHHmmss");
            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString + @"\" + ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos secundarios amazon " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                CreateText.WriteLine(Contenido);
            }

            int ContadorArchivos = 0;
            foreach (var Archivos in Directorios.GetFiles())
            {
                //textBox1.Text = "Procesando Archivo: " + Archivos.Name;
                this.Refresh();
                this.Invalidate();


                // obtiene datos del excel base
                // ----------------------------
                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["NombreArchivoBase"];

                if (file == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivoBase"];

                if (Archivos.Name == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContien"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienUSPS"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienAmazonOriginal"];
                if (filemaster != "")
                {
                    if (Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContienMI15"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                filemaster = ".csv";
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }

                FlgSihayMI15 = true;

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                }


                // elimina las primeras 7 lineas ya que son encabezado de amazon
                int Contadorlineas = 0;
                string RutaAmazon = ConfigurationManager.AppSettings["CarpetaArchivosSecundarios"];
                RutaAmazon = RutaAmazon + "MI15Final_" + ContadorArchivos.ToString() + ".csv";

                using (StreamWriter fileWrite = new StreamWriter(RutaAmazon))
                {
                    using (StreamReader fielRead = new StreamReader(Archivos.FullName))
                    {
                        String line;

                        while ((line = fielRead.ReadLine()) != null)
                        {
                            if (Contadorlineas <= 4)
                            {
                                Contadorlineas = Contadorlineas + 1;
                                continue;
                            }
                            else
                                fileWrite.WriteLine(line);
                        }

                        fielRead.Close();
                    }

                    fileWrite.Close();
                }

                // Read the file and display it line by line.  
                System.IO.StreamReader MI15File = new System.IO.StreamReader(RutaAmazon);

                string[] PalabraMI15 = new string[18];
                int contPala = 0;
                int conteoregistros = 0;

                while ((line = MI15File.ReadLine()) != null)
                {
                    string LineaSinCaracteres = line.Replace("\"", "");
                    //LineaSinCaracteres = LineaSinCaracteres.Replace("/", ",");
                    string[] sa = LineaSinCaracteres.Split(',');

                    if (sa.Length > 18)
                    {
                        //cantidad = sa.Length - 42;
                        continue;
                    }

                    contPala = 0;
                    foreach (string s in sa)
                    {
                        PalabraMI15[contPala] = s;
                        contPala = contPala + 1;
                    }


                    PesosDimensiones.MI15 clsPedido = new PesosDimensiones.MI15();
                    ObtieneValorRegistroDetalleMI15(PalabraMI15, ref clsPedido);
                    clsInsertaRegistro.InsertaMI15(clsPedido, Conexion);
                    conteoregistros += 1;
                }

                MI15File.Close();

                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString + @"\" + Archivos.Name;
                string RutaArchivoMoverMI15 = pathString + @"\" + "MI15Final_" + ContadorArchivos.ToString() + ".csv"; ;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);
                System.IO.File.Move(RutaAmazon, RutaArchivoMoverMI15);
                ContadorArchivos += 1;

            }

            //textBox1.Text = "Fin Carga Archivos MI15";
            this.Refresh();
            this.Invalidate();
        }

        // Realiza accion de boton unifica reportes
        // ----------------------------------------
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection Conexion = new SqlConnection();
                Conexion.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];

                SqlConnection ConexionGenerico1 = new SqlConnection();
                ConexionGenerico1.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                string sqlGenerico1 = "ProvisionMensual";
                SqlCommand commandGenerico1 = new SqlCommand(sqlGenerico1, ConexionGenerico1);
                commandGenerico1.CommandType = CommandType.StoredProcedure;
                commandGenerico1.CommandTimeout = 7200; //in seconds
                ConexionGenerico1.Open();
                commandGenerico1.ExecuteNonQuery();
                ConexionGenerico1.Close();
            }
            catch (SystemException exp)
            {
                MessageBox.Show("Error: " + exp.Message);
                

            }
        }

   
        private void MainForm_Load(object sender, EventArgs e)
        {
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
 
        }

         private void timer1_Tick(object sender, EventArgs e)
        {
            this.progressBar1.Increment(10);
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {
            ProgressBar Progebar = new ProgressBar();
            

        }

        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            //int sum = 0;

            for(int ii=1; ii<=100; ii++)
            {

                //Thread.Sleep(100);
                //sum = sum + 1;
                //backgroundWorker1.ReportProgress(ii);
                //
                //if (backgroundWorker1.CancellationPending)
                //{
                //
                //    e.Cancel = true;
                //    backgroundWorker1.ReportProgress(0);
                //}

                //textBox1.Text = "Cantidad Registros :" + ContadorProgreso;
                this.Refresh();
                this.Invalidate();
            }
        }

        private void BackgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }

        private void BackgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        // obtiene el valor del registro actual
        // ------------------------------------
        private void CargaCancelados(SqlConnection Conexion)
        {
            ProvisionMensualMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ProvisionMensualMaxwarehouse.Clases.ManejoBD();
            //clsInsertaRegistro.EliminaRegistroCancelados(Conexion);

            // abre todos los archivo secundarios para cargarlos en una lista y evaluar cuales se encuentran en los maestros
            // para unificarlos y poder generar un archivo de salida
            // -------------------------------------------------------------------------------------------------------------
            DirectoryInfo Directorios = new DirectoryInfo(ArchivosSecundarios);

            // creo directorio de corrida
            // --------------------------

            System.IO.Directory.CreateDirectory(pathString);

            //// realiza un archivo tipo Excel con la informacion del reporte
            //// ------------------------------------------------------------
            ArchivoLog = pathString + @"\" + ReporteLog;
            using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
            {
                string Contenido = "Inicia procesamiento de archivos secundarios Cancelados " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                CreateText.WriteLine(Contenido);
                Contenido = "";
            }


            foreach (var Archivos in Directorios.GetFiles())
            {
                //textBox1.Text = "Procesando Archivo: " + Archivos.Name;
                this.Refresh();
                this.Invalidate();

                //timer1.Enabled = true;
                //
                //if (progressBar1.Value == progressBar1.Maximum)
                //{
                //    progressBar1.Value = 0;
                //    timer1.Enabled = false;
                //}

                // obtiene datos del excel base
                // ----------------------------
                string file = Archivos.FullName;
                string filemaster = ConfigurationManager.AppSettings["NombreArchivoBase"];

                if (file == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivoBase"];

                if (Archivos.Name == filemaster)
                    continue;

                filemaster = ConfigurationManager.AppSettings["ArchivosSecundarioContien"];
                if (filemaster != "")
                {
                    if (!Archivos.Name.ToUpper().Contains(filemaster.ToUpper()))
                        continue;
                }
                FlgSihayFedex = true;

                //// Registra inicio de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Inicio Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                }

                // obtiene record set
                // ------------------
                string connectionString = GetConnectionString(file, "DETALLE");

                var dataSet = GetDataSetFromExcelFileDetalle(file, connectionString);
                int conteoregistros = 0;

                // recorre registros obtenidos por la lectura del excel
                // ----------------------------------------------------
                foreach (DataRow row in dataSet.Tables[0].Rows)
                {
                    ProvisionMensual.Cancelados clsPedido = new ProvisionMensual.Cancelados();

                    // obtiene el valor del registro leido
                    // -----------------------------------
                    ObtieneValorRegistroCancelados(row, ref clsPedido);

                    //listaPedido.Add(clsPedido);
                    clsInsertaRegistro.InsertaBDCancelados(clsPedido, Conexion);
                    conteoregistros += 1;
                }

                //// Registra fin de procesamiento de archivo
                //// ----------------------------------------
                using (System.IO.StreamWriter CreateText = new System.IO.StreamWriter(ArchivoLog, true))
                {
                    string Contenido = "Archivo: " + Archivos.Name + " Contiene: " + conteoregistros + " Registros";
                    CreateText.WriteLine(Contenido);
                    Contenido = "Fin Procesamiento Archivo: " + Archivos.Name + " " + DateTime.Now.ToString("yyyyMMddTHHmmss") + "\n\n";
                    CreateText.WriteLine(Contenido);
                    Contenido = "";
                    CreateText.WriteLine(Contenido);
                }

                string RutaArchivoMover = pathString + @"\" + Archivos.Name;
                System.IO.File.Move(Archivos.FullName, RutaArchivoMover);


            }

            //textBox1.Text = "Fin Carga Archivos Fedex";
            this.Refresh();
            this.Invalidate();
        }

        // Realiza accion de boton unifica reportes
        // ----------------------------------------
        private void EjecutaProceso()
        {
            try
            {
                SqlConnection Conexion = new SqlConnection();
                Conexion.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                ProvisionMensualMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ProvisionMensualMaxwarehouse.Clases.ManejoBD();
                ProvisionMensualMaxwarehouse.Clases.Logguer clsLogguer = new ProvisionMensualMaxwarehouse.Clases.Logguer();
                DateTime now = DateTime.Now;
                pathString = ArchivosSecundarios + "Output" + DateTime.Now.ToString("yyyyMMddTHHmmss");
                ReporteLog = "ReporteLog" + DateTime.Now.ToString("yyyyMMddTHHmmss") + ".txt";
                string pathOutPut1 = "";

                System.IO.Directory.CreateDirectory(pathString);

                //textBox1.Text = "Inicio Proceso";
                this.Refresh();
                this.Invalidate();

                now = DateTime.Now;
                clsLogguer.LogDuration(now, "Finaliza Insercion de BD DW...");
                //textBox1.Text = "Finaliza Insercion de BD DW...";

                // Inicia Carga de Fedex
                // ---------------------    
                ObtieneDatosFedex(Conexion);

                // Inicia carga de usps
                // --------------------
                ObtieneDatosUSPS(Conexion);

                // Inicia carga de UPS
                // -------------------
                ObtieneDatosUPS(Conexion);

                // Inicia carga Amazon
                // -------------------
                ObtieneDatosAmazon(Conexion);

                // Inicia carga Endicia
                // -------------------
                //textBox1.Text = "COMIENZA Carga Archivos Endicia";
                this.Refresh();
                this.Invalidate();

                ObtieneDatosEndicia(Conexion);

                //textBox1.Text = "Fin Carga Archivos Endicia";
                this.Refresh();
                this.Invalidate();

                //textBox1.Text = "Genera Excel";
                this.Refresh();
                this.Invalidate();
                using (var workbook = new XLWorkbook())
                {
                    // ejecuto sp que devuelve el crokis
                    // ----------------------------------
                    SqlConnection Conexion1 = new SqlConnection();

                    Conexion1.ConnectionString = ConfigurationManager.AppSettings["ConectionString"]; 

                    Conexion1.Open();
                    SqlCommand cmd = new SqlCommand("GeneraReporte", Conexion1);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 7200; //in seconds
                    SqlDataReader reader = cmd.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt.Load(reader);

                    var worksheet = workbook.Worksheets.Add("Output");

                    List<string[]> titles = new List<string[]> { new string[] { "SalesOrderNumber", "HoldCode", "TotalSales", "SalesSku", "SalesCategoryAtTimeOfSale", "UomCode", "UomQuantity", "SalesStatus", "SalesOrderDate", "SalesChannelName", "CustomerName", "FulfillmentSku", "FulfillmentChannelName", "FulfillmentChannelType", "LinkedFulfillmentChannelName", "FulfillmentLocationName", "FulfillmentOrderNumber", "Quantity", "Sku", "Title", "TotalCost", "Commission", "InventoryCost", "UnitCost", "ServiceCost", "EstimatedShippingCost", "ShippingCost", "ShippingPrice", "OverheadCost", "PackageCost", "ProfitLoss", "Carrier", "ShippingServiceLevel", "ShippedByUser", /*"ShippingWeight", "Length", "Width", "Height", "Weight", */"StateRegion", "TrackingNum", "MfrName", "PricingRule", "RequestedServiceLevel", "GroundTrackingIDPrefix", "ExpressorGroundTrackingID", "NetChargeAmount", "ServiceType", "GroundService", "ShipmentDate", "PODDeliveryDate", "ActualWeightAmount", "RatedWeightAmount", "DimLength", "DimWidth", "DimHeight", "DimDivisor", "ShipperState", "ZoneCode", "TenderedDate", "EarnedDiscount", "FuelSurcharge", "PerformancePricing", "DeliveryAreaSurchargeExtended", "DeliveryAreaSurcharge", "USPSNonMachSurcharge", "Residential", "GraceDiscount", "DeclaredValue", "DASExtendedResi", "AdditionalHandling", "ParcelReLabelCharge", "IndirectSignature", "DASResi", "AddressCorrection", "DASExtendedComm", "OversizeCharge", "AHSDimensions", "InvoiceDate", "InvoiceNumber", "MailClass", "TrackingNumberUSPS", "TotalPostageAmt", "DeliveryDate", "WeightUSPS", "Zone", "ServiceTypeUPS", "TrackingNumberUPS", "NetChargeAmountUPS", "order_idAmazon", "CarrierCargado" } };

                    worksheet.Cell(1, 1).InsertData(titles); //insert titles to one row

                    worksheet.Cell(2, 1).InsertData(dt);// inserta Contenido
                    string pathOutPut = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                    pathOutPut = pathOutPut + @"\Output" + DateTime.Now.ToString("MMddyyyy") + ".xlsx";
                    workbook.SaveAs(pathOutPut);

                    //Conexion1.Close();

                }

                //textBox1.Text = "Ejecuta SP No facturadas";
                this.Refresh();
                this.Invalidate();

                // Genera promedios FEDEX
                // ----------------------
                SqlConnection ConexionGenerico = new SqlConnection();
                ConexionGenerico.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                string sqlGenerico = "GeneraReporteNoFacturadas";
                SqlCommand commandGenerico = new SqlCommand(sqlGenerico, ConexionGenerico);
                commandGenerico.CommandType = CommandType.StoredProcedure;
                commandGenerico.CommandTimeout = 7200; //in seconds
                ConexionGenerico.Open();
                commandGenerico.ExecuteNonQuery();
                ConexionGenerico.Close();

                // Genera Reporte no facturadas
                // ----------------------------
                string PSNoFactutada = ConfigurationManager.AppSettings["RutaPSNoFacturada"];
                var proc7 = new System.Diagnostics.ProcessStartInfo();
                //string anyCommand;
                proc7.UseShellExecute = true;
                proc7.WorkingDirectory = @"C:\Windows\System32";
                proc7.FileName = @"C:\Windows\System32\cmd.exe";
                //proc1.Verb = "runas";
                proc7.Arguments = "/c " + "powershell -ExecutionPolicy Bypass -File \"" + PSNoFactutada + "\"";
                proc7.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                System.Diagnostics.Process.Start(proc7);

                // llena tabla de bianalitics
                // --------------------------
                //textBox1.Text = "Llena BI analitics";
                this.Refresh();
                this.Invalidate();
                SqlConnection Conexionbianalitics = new SqlConnection();
                Conexionbianalitics.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                string sqlbianalitics = "LlenaBIANALITICS";
                SqlCommand commandbianalitics = new SqlCommand(sqlbianalitics, Conexionbianalitics);
                commandbianalitics.CommandType = CommandType.StoredProcedure;
                commandbianalitics.CommandTimeout = 7200; //in seconds
                Conexionbianalitics.Open();
                commandbianalitics.ExecuteNonQuery();
                Conexionbianalitics.Close();

                // traslada facturas a historicos
                // ------------------------------
                //textBox1.Text = "Traslada historicos";
                this.Refresh();
                this.Invalidate();
                SqlConnection Conexionhistcarrier = new SqlConnection();
                Conexionhistcarrier.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                string sqlhistcarrier = "TrasladaCarrierHistorico";
                SqlCommand commandhistcarrier = new SqlCommand(sqlhistcarrier, Conexionhistcarrier);
                commandhistcarrier.CommandType = CommandType.StoredProcedure;
                commandhistcarrier.CommandTimeout = 7200; //in seconds
                Conexionhistcarrier.Open();
                commandhistcarrier.ExecuteNonQuery();
                Conexionhistcarrier.Close();

//                textBox1.Text = "Reporte Shipware";
                this.Refresh();
                this.Invalidate();
                using (var workbook = new XLWorkbook())
                {
                    // ejecuto sp que devuelve el crokis
                    // ----------------------------------
                    SqlConnection Conexion1 = new SqlConnection();

                    Conexion1.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];

                    Conexion1.Open();
                    SqlCommand cmd = new SqlCommand("ReporteShipware", Conexion1);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 7200; //in seconds
                    SqlDataReader reader = cmd.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt.Load(reader);

                    var worksheet = workbook.Worksheets.Add("Output");

                    List<string[]> titles = new List<string[]> { new string[] { "TRACKING_NUMBER", "Length", "Width", "Height", "Weight", "DimLength", "DimWidth", "DimHeight", "ActualWeightAmount", "AdditionalHandling", "OversizeCharge", "USPSNonMachSurcharge", "AHSDimensions" } };

                    worksheet.Cell(1, 1).InsertData(titles); //insert titles to one row

                    worksheet.Cell(2, 1).InsertData(dt);// inserta Contenido
                    string pathOutPut = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                    pathOutPut = pathOutPut + @"\ReporteShipware" + DateTime.Now.ToString("MMddyyyy") + ".xlsx";
                    workbook.SaveAs(pathOutPut);

                }

                // Genera no facturadas FEDEX
                // ----------------------
                SqlConnection ConexionGenerico1 = new SqlConnection();
                ConexionGenerico1.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                string sqlGenerico1 = "GeneraReporteNoFacturadas";
                SqlCommand commandGenerico1 = new SqlCommand(sqlGenerico1, ConexionGenerico1);
                commandGenerico1.CommandType = CommandType.StoredProcedure;
                commandGenerico1.CommandTimeout = 7200; //in seconds
                ConexionGenerico1.Open();
                commandGenerico1.ExecuteNonQuery();
                ConexionGenerico1.Close();

                //textBox1.Text = "Reporte Provision de Shipping";
                //this.Refresh();
                //this.Invalidate();
                //using (var workbook = new XLWorkbook())
                //{
                //    // ejecuto sp que devuelve el crokis
                //    // ----------------------------------
                //    SqlConnection Conexion1 = new SqlConnection();
                //
                //    Conexion1.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                //
                //    Conexion1.Open();
                //    SqlCommand cmd = new SqlCommand("ProvisionShipping", Conexion1);
                //    cmd.CommandType = CommandType.StoredProcedure;
                //    cmd.CommandTimeout = 7200; //in seconds
                //    SqlDataReader reader = cmd.ExecuteReader();
                //
                //    //Create a new DataTable.
                //    System.Data.DataTable dt = new System.Data.DataTable("Resultado");
                //
                //    //Load DataReader into the DataTable.
                //    dt.Load(reader);
                //
                //    var worksheet = workbook.Worksheets.Add("Output");
                //
                //    List<string[]> titles = new List<string[]> { new string[] { "TRACKING_NUMBER", "Length", "Width", "Height", "Weight", "DimLength", "DimWidth", "DimHeight", "ActualWeightAmount", "AdditionalHandling", "OversizeCharge", "USPSNonMachSurcharge", "AHSDimensions" } };
                //
                //    worksheet.Cell(1, 1).InsertData(titles); //insert titles to one row
                //
                //    worksheet.Cell(2, 1).InsertData(dt);// inserta Contenido
                //    string pathOutPut = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                //    pathOutPut = pathOutPut + @"\ReporteShipware" + DateTime.Now.ToString("MMddyyyy") + ".xlsx";
                //    workbook.SaveAs(pathOutPut);
                //
                //}

                //textBox1.Text = "Inicio Proceso";
                this.Refresh();
                this.Invalidate();

                now = DateTime.Now;
                clsLogguer.LogDuration(now, "Genera dimensiones Promedio...");
               // textBox1.Text = "Genera dimensiones Promedio";

                using (var workbook = new XLWorkbook())
                {
                    // ejecuto sp que devuelve el crokis
                    // ----------------------------------
                    SqlConnection Conexion1 = new SqlConnection();

                    Conexion1.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];

                    Conexion1.Open();
                    SqlCommand cmd = new SqlCommand("DimensionesCompleto", Conexion1);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 1200; //in seconds
                    SqlDataReader reader = cmd.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt.Load(reader);

                    var worksheet = workbook.Worksheets.Add("Output");

                    List<string[]> titles = new List<string[]> { new string[] { "Sku", "UomCode", "ActualWeightAmount", "DimLength", "DimWidth", "DimHeight" } };

                    worksheet.Cell(1, 1).InsertData(titles); //insert titles to one row

                    worksheet.Cell(2, 1).InsertData(dt);// inserta Contenido
                    string pathOutPut = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                    pathOutPut = pathOutPut + @"\PesosDimensiones" + DateTime.Now.ToString("MMddyyyy") + ".xlsx";
                    workbook.SaveAs(pathOutPut);
                }

                string ArchivoOrigenFTP = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                ArchivoOrigenFTP = ArchivoOrigenFTP + @"\PesosDimensiones" + DateTime.Now.ToString("MMddyyyy") + ".xlsx";
                string ArchivoFinalFTP = ConfigurationManager.AppSettings["DireccionFTPDimensiones"];
                ArchivoFinalFTP = ArchivoFinalFTP + @"\PesosDimensiones" + DateTime.Now.ToString("MMddyyyy") + ".xlsx        ";

                SubeFTP(ArchivoOrigenFTP, ArchivoFinalFTP, true);


                // mueve archivo a carpeta de vera
                // -------------------------------
               // textBox1.Text = "Mueve Archivo Vera";
                this.Refresh();
                this.Invalidate();
                System.Threading.Thread.Sleep(600000);
                string PathVera = ConfigurationManager.AppSettings["RutaArchivosOutputsVera"];
                PathVera = PathVera + @"\Output" + DateTime.Now.ToString("MMddyyyy") + ".xlsx";
                string PathDX = ConfigurationManager.AppSettings["RutaArchivosOutputsDX"];
                PathDX = PathDX + @"\Output" + DateTime.Now.ToString("MMddyyyy") + ".xlsx";

                pathOutPut1 = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                pathOutPut1 = pathOutPut1 + @"\Output" + DateTime.Now.ToString("MMddyyyy") + ".xlsx";

                System.IO.File.Copy(pathOutPut1, PathVera);
                System.IO.File.Copy(pathOutPut1, PathDX);

                System.Threading.Thread.Sleep(600000);
                //textBox1.Text = "MueveArchivoDX";
                this.Refresh();
                this.Invalidate();
                PathVera = ConfigurationManager.AppSettings["RutaArchivosOutputsVera"];
                PathVera = PathVera + @"/ReporteNoFacturadas" + DateTime.Now.ToString("MMddyyyy") + ".csv";
                PathDX = ConfigurationManager.AppSettings["RutaArchivosOutputsDX"];
                PathDX = PathDX + @"/ReporteNoFacturadas" + DateTime.Now.ToString("MMddyyyy") + ".csv";

                pathOutPut1 = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                pathOutPut1 = pathOutPut1 + @"\NoFacturadas" + DateTime.Now.ToString("MMddyyyy") + ".csv";

                System.IO.File.Copy(pathOutPut1, PathVera);
                System.IO.File.Copy(pathOutPut1, PathDX);

            }
            catch (Exception exp)
            {
                MessageBox.Show("Error: " + exp.Message);
            }
        }

        public void SubeFTP(string ArchivoOrigenFTP, string ArchivoFinalFTP, bool EsExcel =false)
        {
            string UserFTP = ConfigurationManager.AppSettings["UserFTP"];
            string PassFTP = ConfigurationManager.AppSettings["PassFTP"];

            if (System.IO.File.Exists(ArchivoOrigenFTP))

            {
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(ArchivoFinalFTP);
                request.Method = WebRequestMethods.Ftp.UploadFile;

                // This example assumes the FTP site uses anonymous logon.
                request.Credentials = new NetworkCredential(UserFTP, PassFTP);

                // finaliza conexion FTP
                // ---------------------
                request.KeepAlive = false;

                byte[] fileContents;
                // Si no es Excel utiliza este formateo de datos
                // ---------------------------------------------
                if (EsExcel == false)
                {
                    // Copy the contents of the file to the request stream.

                    using (StreamReader sourceStream = new StreamReader(ArchivoOrigenFTP))
                    {
                        fileContents = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd());
                    }

                    request.ContentLength = fileContents.Length;
                }
                else 
                {
                    // Copy the contents of the file to the request stream.
                    fileContents = File.ReadAllBytes(ArchivoOrigenFTP);

                    request.ContentLength = fileContents.Length;
                }


                using (Stream requestStream = request.GetRequestStream())
                {
                    requestStream.Write(fileContents, 0, fileContents.Length);
                }

                using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
                {
                    //Console.WriteLine($"Upload File Complete, status {response.StatusDescription}");

                    //MessageBox.Show("Subieron los archivos " + response.StatusDescription);
                    DateTime now = DateTime.Now;
                    ProvisionMensualMaxwarehouse.Clases.Logguer clsLogguer = new ProvisionMensualMaxwarehouse.Clases.Logguer();
                    clsLogguer.LogDuration(now, "Finaliza Carga de archivo"+ ArchivoOrigenFTP + " a direccion FTP " + ArchivoFinalFTP);
                    //textBox1.Text = "Finaliza Carga de archivo" + ArchivoOrigenFTP;
                }

            }

        }
        private void AsignaFechaInicioFin(DateTime Fecha1, DateTime Fecha2)
        {
 
        }


        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                SqlConnection Conexion = new SqlConnection();
                Conexion.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                ProvisionMensualMaxwarehouse.Clases.ManejoBD clsInsertaRegistro = new ProvisionMensualMaxwarehouse.Clases.ManejoBD();
                ProvisionMensualMaxwarehouse.Clases.Logguer clsLogguer = new ProvisionMensualMaxwarehouse.Clases.Logguer();


                // Inicia Carga de Fedex
                // ---------------------    
               // ObtieneDatosFedex(Conexion);
               //
               // // Inicia carga de usps
               // // --------------------
               // ObtieneDatosUSPS(Conexion);
               //
               // // Inicia carga de UPS
               // // -------------------
               // ObtieneDatosUPS(Conexion);
               //
               // // Inicia carga Amazon
               // // -------------------
               // ObtieneDatosAmazon(Conexion);
               //
               // // Inicia Carga PITNEYBOWES
               // // ------------------------
               // ObtieneDatosPITNEYBOWES(Conexion);
               //
               // // Inicia Carga estimated delivery date
               // // ------------------------------------
               // ObtieneDatosEstimatedDeliveryDate(Conexion);
               //
               // // Inicia Carga Cancelados
               // // -----------------------
               // CargaCancelados(Conexion);

                DateTime Fecha1;
                DateTime Fecha2;
                int Anio = 0;
                int Mes = 0;
                DateTime date = DateTime.Today;
                                                                      // date.AddDays(1);
                                                                      // verifica si es fin de mes
                if (comboBox2.Text != "SI")
                {
                    Anio = Convert.ToInt32(textBox7.Text);
                    Mes = Convert.ToInt16(comboBox1.SelectedItem);

                    int MesSigue0 = 0;
                    int anioproc0 = 0;
                    MesSigue0 = Mes;
                    MesSigue0 = MesSigue0 + 1;

                    if (MesSigue0 > 12)
                    {
                        MesSigue0 = 1;
                        anioproc0 = Anio + 1;
                    }
                    else
                    {
                        anioproc0 = Anio;
                    }

                    date = new DateTime(anioproc0, MesSigue0, 1);
                }
                else 
                {
                    Anio = Convert.ToInt32(textBox7.Text);
                    Mes = Convert.ToInt16(comboBox1.SelectedItem);

                }

                date = date.AddMonths(-1);

                int MesSigue = 0;
                int anioproc = 0;
                MesSigue = date.Month;
                MesSigue = MesSigue + 1;

                if (MesSigue > 12)
                {
                    MesSigue = 1;
                    anioproc = date.Year + 1;
                }
                else 
                {
                    anioproc = date.Year;
                }

                Fecha1 = new DateTime(date.Year, date.Month, 1);
                Fecha2 = new DateTime(anioproc, MesSigue, 1).AddDays(-1);

                string monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedMonthName(Fecha1.Month);

                DateTime FechaNombresProceso;
                DateTime FechaNombresAnterior ;
                FechaNombresProceso = new DateTime(Anio, Mes , 1);
                DateTime FechaIntermedia = FechaNombresProceso.AddMonths(-1);
                FechaNombresAnterior = new DateTime(FechaIntermedia.Year, FechaIntermedia.Month, 1); ;

                string monthNombreActual = FechaNombresProceso.AddMonths(1).ToString("MMM", CultureInfo.InvariantCulture);
                string monthNombreProceso = FechaNombresProceso.ToString("MMM", CultureInfo.InvariantCulture);
                string monthNombreAnterior = FechaNombresAnterior.ToString("MMM", CultureInfo.InvariantCulture);

                // valida si debe de reconstruir el DW o utilizar el que ya se tiene por defecto esto es por pruebas quitar y que siempre lo reconstruya
                if (checkBox1.Checked)
                {

                    SqlConnection ConexionDWD = new SqlConnection();
                    ConexionDWD.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    string sqlGenericoDWD = "DWDEPURADO";
                    SqlCommand commandDWD = new SqlCommand(sqlGenericoDWD, ConexionDWD);
                    commandDWD.CommandType = CommandType.StoredProcedure;
                    commandDWD.CommandTimeout = 10200; //in seconds
                    ConexionDWD.Open();
                    commandDWD.ExecuteNonQuery();
                    ConexionDWD.Close();
                }


                if (comboBox2.Text != "SI")
                {
                    SqlConnection ConexionFEDEX = new SqlConnection();
                    ConexionFEDEX.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    string sqlGenericofedex = "ObtienMontosFEDEX";
                    SqlCommand commandFEDEX = new SqlCommand(sqlGenericofedex, ConexionFEDEX);
                    commandFEDEX.Parameters.Add("@FechaInicio", SqlDbType.DateTime).Value = Fecha1;
                    commandFEDEX.Parameters.Add("@FechaFin", SqlDbType.DateTime).Value = Fecha2;
                    commandFEDEX.CommandType = CommandType.StoredProcedure;
                    commandFEDEX.CommandTimeout = 7200; //in seconds
                    ConexionFEDEX.Open();
                    commandFEDEX.ExecuteNonQuery();
                    ConexionFEDEX.Close();

                    // SqlConnection ConexionA = new SqlConnection();
                    //ConexionA.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    //string sqlGenericoA = "ObtienMontosAMAZON";
                    //SqlCommand commandA = new SqlCommand(sqlGenericoA, ConexionA);
                    //commandA.Parameters.Add("@FechaInicio", SqlDbType.DateTime).Value = Fecha1;
                    //commandA.Parameters.Add("@FechaFin", SqlDbType.DateTime).Value = Fecha2;
                    //commandA.CommandType = CommandType.StoredProcedure;
                    //commandA.CommandTimeout = 7200; //in seconds
                    //ConexionA.Open();
                    //commandA.ExecuteNonQuery();
                    //ConexionA.Close();
                }
                else
                {
                    SqlConnection ConexionFEDEX = new SqlConnection();
                    ConexionFEDEX.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    string sqlGenericofedex = "ObtienMontosFEDEXSEMANAL";
                    SqlCommand commandFEDEX = new SqlCommand(sqlGenericofedex, ConexionFEDEX);
                    commandFEDEX.Parameters.Add("@FechaInicio", SqlDbType.DateTime).Value = Fecha1;
                    commandFEDEX.Parameters.Add("@FechaFin", SqlDbType.DateTime).Value = Fecha2;
                    commandFEDEX.CommandType = CommandType.StoredProcedure;
                    commandFEDEX.CommandTimeout = 7200; //in seconds
                    ConexionFEDEX.Open();
                    commandFEDEX.ExecuteNonQuery();
                    ConexionFEDEX.Close();

                    //SqlConnection ConexionA = new SqlConnection();
                    //ConexionA.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    //string sqlGenericoA = "ObtienMontosAMAZONSEMANAL";
                    //SqlCommand commandA = new SqlCommand(sqlGenericoA, ConexionA);
                    //commandA.Parameters.Add("@FechaInicio", SqlDbType.DateTime).Value = Fecha1;
                    //commandA.Parameters.Add("@FechaFin", SqlDbType.DateTime).Value = Fecha2;
                    //commandA.CommandType = CommandType.StoredProcedure;
                    //commandA.CommandTimeout = 7200; //in seconds
                    //ConexionA.Open();
                    //commandA.ExecuteNonQuery();
                    //ConexionA.Close();

                }

                if (comboBox2.Text != "SI")
                {
                    SqlConnection ConexionGenerico1 = new SqlConnection();
                    ConexionGenerico1.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    string sqlGenerico1 = "ProvisionMensual";
                    SqlCommand commandGenerico1 = new SqlCommand(sqlGenerico1, ConexionGenerico1);
                    commandGenerico1.Parameters.Add("@FechaInicio", SqlDbType.DateTime).Value = Fecha1;
                    commandGenerico1.Parameters.Add("@FechaFin", SqlDbType.DateTime).Value = Fecha2;
                    commandGenerico1.Parameters.Add("@Mes", SqlDbType.VarChar).Value = "PROCESO";
                    commandGenerico1.CommandType = CommandType.StoredProcedure;
                    commandGenerico1.CommandTimeout = 7200; //in seconds
                    ConexionGenerico1.Open();
                    commandGenerico1.ExecuteNonQuery();
                    ConexionGenerico1.Close();
                }
                else
                {
                    SqlConnection ConexionGenerico1 = new SqlConnection();
                    ConexionGenerico1.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    string sqlGenerico1 = "ProvisionSEMANAL";
                    SqlCommand commandGenerico1 = new SqlCommand(sqlGenerico1, ConexionGenerico1);
                    commandGenerico1.Parameters.Add("@FechaInicio", SqlDbType.DateTime).Value = Fecha1;
                    commandGenerico1.Parameters.Add("@FechaFin", SqlDbType.DateTime).Value = Fecha2;
                    commandGenerico1.Parameters.Add("@Mes", SqlDbType.VarChar).Value = "PROCESO";
                    commandGenerico1.CommandType = CommandType.StoredProcedure;
                    commandGenerico1.CommandTimeout = 7200; //in seconds
                    ConexionGenerico1.Open();
                    commandGenerico1.ExecuteNonQuery();
                    ConexionGenerico1.Close();
                }

                using (var workbook = new XLWorkbook())
                {
                    // obtengo datos facturas
                    // ----------------------
                    SqlConnection Conexion1 = new SqlConnection();
                
                    Conexion1.ConnectionString = ConfigurationManager.AppSettings["ConectionString"]; 
                    Conexion1.Open();
                    SqlCommand cmd = new SqlCommand("DatosProvisionMensual", Conexion1);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 7200; //in seconds
                    cmd.Parameters.Add("@Operacion", SqlDbType.Int).Value = 5;
                    SqlDataReader reader = cmd.ExecuteReader();
                
                    //Create a new DataTable.
                    System.Data.DataTable dt = new System.Data.DataTable("Resultado");

                    //Conexion1.Close();

                    //Load DataReader into the DataTable.
                    dt.Load(reader);

                    SqlConnection Conexion0 = new SqlConnection();

                    //Conexion0.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    //Conexion0.Open();
                    //SqlCommand cmd0 = new SqlCommand("DatosProvisionMensual", Conexion0);
                    //cmd0.CommandType = CommandType.StoredProcedure;
                    //cmd0.CommandTimeout = 7200; //in seconds
                    //cmd0.Parameters.Add("@Operacion", SqlDbType.Int).Value = 6;
                    //SqlDataReader reader0 = cmd0.ExecuteReader();
                    //
                    ////Create a new DataTable.
                    //System.Data.DataTable dt0 = new System.Data.DataTable("Resultado");
                    //
                    ////Load DataReader into the DataTable.
                    //dt0.Load(reader0);

                    var worksheet = workbook.Worksheets.Add("PROVISION");

                    string Encabezado = "Invoice " + monthName+ " Carriers";
                    List<string[]> titles = new List<string[]> { new string[] {  Encabezado } };
                
                    worksheet.Cell(1, 2).InsertData(titles);
                    //Conexion0.Open();

                    //foreach (DataRow row in dt0.Rows)
                    //{
                    //    if (row["RequestedServiceLevel"].ToString() == "FEDEX")
                    //    {
                    //        worksheet.Cell(3, 3).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                    //        worksheet.Cell(3, 3).DataType = XLDataType.Number;
                    //        float Valor = (float)Convert.ToDouble(row["EstimatedShippingCost"].ToString()); 
                    //        worksheet.Cell(3, 3).SetValue(Valor);
                    //    }
                    //
                    //    if (row["RequestedServiceLevel"].ToString() == "UPS")
                    //    {
                    //        worksheet.Cell(4, 3).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                    //        worksheet.Cell(4, 3).DataType = XLDataType.Number;
                    //        float Valor = (float)Convert.ToDouble(row["EstimatedShippingCost"].ToString());
                    //        worksheet.Cell(4, 3).SetValue(Valor);
                    //    }
                    //
                    //    if (row["RequestedServiceLevel"].ToString() == "USPS")
                    //    {
                    //        worksheet.Cell(5, 3).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                    //        worksheet.Cell(5, 3).DataType = XLDataType.Number;
                    //        float Valor = (float)Convert.ToDouble(row["EstimatedShippingCost"].ToString());
                    //        worksheet.Cell(5, 3).SetValue(Valor);
                    //    }
                    //}
                    var rango = worksheet.Range("A3:B3"); rango.Style.Fill.BackgroundColor = XLColor.PowderBlue;
                    rango = worksheet.Range("A4:B4"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                    rango = worksheet.Range("A5:B5"); rango.Style.Fill.BackgroundColor = XLColor.PowderBlue;
                    rango = worksheet.Range("A6:B6"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                    rango = worksheet.Range("A7:B7"); rango.Style.Fill.BackgroundColor = XLColor.PowderBlue;
                    
                    if (comboBox2.Text != "SI")
                    {
                        rango = worksheet.Range("A8:B8"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                        rango = worksheet.Range("A9:B9"); rango.Style.Fill.BackgroundColor = XLColor.PowderBlue;
                        rango = worksheet.Range("A10:B10"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                        rango = worksheet.Range("A11:B11"); rango.Style.Fill.BackgroundColor = XLColor.PowderBlue;
                        rango = worksheet.Range("A12:B12"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                    }

                    if (dt.Rows.Count != 0)
                    {
                        worksheet.Cell(3, 1).InsertData(dt);// inserta Contenido
                    }
 
                    double Valor1 = Convert.ToDouble(textBox3.Text);

                    if (comboBox2.Text != "SI")
                    {
                       
                        worksheet.Cell(10, 2).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        worksheet.Cell(10, 2).DataType = XLDataType.Number;
                        string Etiqueta = "Returns (USPS)";
                        worksheet.Cell(10, 1).SetValue(Etiqueta);
                        worksheet.Cell(10, 2).SetValue(Valor1);

                        double Valor2 = Convert.ToDouble(textBox4.Text);
                        worksheet.Cell(11, 2).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        worksheet.Cell(11, 2).DataType = XLDataType.Number;
                        string Etiqueta2 = "Returns WALMART";
                        worksheet.Cell(11, 1).SetValue(Etiqueta2);
                        worksheet.Cell(11, 2).SetValue(Valor2);


                        worksheet.Cell(12, 2).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        worksheet.Cell(12, 2).DataType = XLDataType.Number;
                        string Etiqueta4 = "Provisión " + monthNombreProceso;
                        worksheet.Cell(12, 1).SetValue(Etiqueta4);
                        worksheet.Cell(12, 2).FormulaA1 = "=B24";

                        double Valor3 = Convert.ToDouble(textBox5.Text);
                        worksheet.Cell(13, 2).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        worksheet.Cell(13, 2).DataType = XLDataType.Number;
                        string Etiqueta3 = "Reversión Provisión " + monthNombreAnterior;
                        worksheet.Cell(13, 1).SetValue(Etiqueta3);
                        worksheet.Cell(13, 2).SetValue(Valor3);

                        worksheet.Cell(14, 2).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        worksheet.Cell(14, 2).DataType = XLDataType.Number;

                        worksheet.Cell(14, 2).FormulaA1 = "=SUM(B3:B13)";
                    }

                     // ejecuto sp que devuelve el crokis
                    // ----------------------------------
                    SqlConnection Conexion2 = new SqlConnection();

                    Conexion2.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion2.Open();
                    SqlCommand cmd2 = new SqlCommand("DatosProvisionMensual", Conexion2);
                    cmd2.CommandType = CommandType.StoredProcedure;
                    cmd2.CommandTimeout = 7200; //in seconds
                    cmd2.Parameters.Add("@Operacion", SqlDbType.Int).Value = 1;
                    SqlDataReader reader2 = cmd2.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt2 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt2.Load(reader2);
                    //Conexion2.Close();

                    rango = worksheet.Range("D2:V2"); rango.Style.Fill.BackgroundColor = XLColor.Coral;
                    rango.Style.Font.FontSize = 12;

                    rango = worksheet.Range("D3:V3"); rango.Style.Fill.BackgroundColor = XLColor.Wheat;
                    rango.Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                    rango.DataType = XLDataType.Number;

                    rango = worksheet.Range("D4:V4"); rango.Style.Fill.BackgroundColor = XLColor.SandyBrown;
                    rango.Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                    rango.DataType = XLDataType.Number;
                   
                    rango = worksheet.Range("D5:V5"); rango.Style.Fill.BackgroundColor = XLColor.Wheat;
                    rango.Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                    rango.DataType = XLDataType.Number;

                    rango = worksheet.Range("D6:V6"); rango.Style.Fill.BackgroundColor = XLColor.SandyBrown;
                    rango.Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                    rango.DataType = XLDataType.Number;

                    rango = worksheet.Range("D7:V7"); rango.Style.Fill.BackgroundColor = XLColor.Wheat;
                    rango.Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                    rango.DataType = XLDataType.Number;

                    if (comboBox2.Text != "SI")
                    {
                        rango = worksheet.Range("D10:R10"); rango.Style.Fill.BackgroundColor = XLColor.Gray;
                        rango.Style.Font.FontSize = 12;

                        rango = worksheet.Range("D11:R11"); rango.Style.Fill.BackgroundColor = XLColor.PastelGray;
                        rango.Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        rango.DataType = XLDataType.Number;

                        rango = worksheet.Range("D12:R12"); rango.Style.Fill.BackgroundColor = XLColor.Silver;
                        rango.Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        rango.DataType = XLDataType.Number;

                        rango = worksheet.Range("D13:R13"); rango.Style.Fill.BackgroundColor = XLColor.PastelGray;
                        rango.Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        rango.DataType = XLDataType.Number;
                    }

                    int Fila = 4;
                    foreach (DataRow row in dt2.Rows)
                    {
                        foreach (DataColumn column in dt2.Columns)
                        {
                            worksheet.Cell(2, Fila).SetValue(column.ColumnName);
                            Fila = Fila + 1;
                        }
                        break;
                    }

                    if (dt2.Rows.Count != 0)
                    {
                        worksheet.Cell(3, 4).InsertData(dt2);// inserta Contenido
                    }

                    string Labels = "";
                    // Etiqueta Total Facturas Carrier
                    // -------------------------------
                    if (comboBox2.Text != "SI")
                    {
                        rango = worksheet.Range("A16:B16"); rango.Style.Fill.BackgroundColor = XLColor.MistyRose;
                        rango = worksheet.Range("A17:B17"); rango.Style.Fill.BackgroundColor = XLColor.Melon;
                        rango = worksheet.Range("A18:B18"); rango.Style.Fill.BackgroundColor = XLColor.MistyRose;
                        rango = worksheet.Range("A19:B19"); rango.Style.Fill.BackgroundColor = XLColor.Melon;
                        rango = worksheet.Range("A20:B20"); rango.Style.Fill.BackgroundColor = XLColor.MistyRose;
                    }

                    if (comboBox2.Text != "SI")
                    {
                        worksheet.Cell(16, 1).SetValue("Total Facturas Carriers");

                        Labels = "Gasto " + monthNombreProceso + " Venta " + monthNombreAnterior;
                        worksheet.Cell(17, 1).SetValue(Labels);

                        Labels = "TFC - Cobro de " + monthNombreAnterior;
                        worksheet.Cell(18, 1).SetValue(Labels);

                        Labels = "Gasto " + monthNombreProceso + " Venta " + monthNombreProceso;
                        worksheet.Cell(19, 1).SetValue(Labels);


                        rango = worksheet.Range("A22:B22"); rango.Style.Fill.BackgroundColor = XLColor.CadetGrey;
                        rango = worksheet.Range("A23:B23"); rango.Style.Fill.BackgroundColor = XLColor.PastelGray;
                        rango = worksheet.Range("A24:B24"); rango.Style.Fill.BackgroundColor = XLColor.CadetGrey;

                        Labels = "Shipping " + monthNombreProceso + " Cobrado " + monthNombreActual;
                        worksheet.Cell(22, 1).SetValue(Labels);

                        Labels = "Shipping " + monthNombreProceso + " por cobrar ";
                        worksheet.Cell(23, 1).SetValue(Labels);
                        worksheet.Cell(24, 1).SetValue("Provision");


                        rango = worksheet.Range("A27:B27"); rango.Style.Fill.BackgroundColor = XLColor.PinkOrange;
                        rango = worksheet.Range("A28:B28"); rango.Style.Fill.BackgroundColor = XLColor.Sunset;

                        worksheet.Cell(27, 1).SetValue("Shipping No Cobrado Sales");
                        worksheet.Cell(28, 1).SetValue("Shipping No Cobrado");


                        rango = worksheet.Range("A31:B31"); rango.Style.Fill.BackgroundColor = XLColor.MikadoYellow;
                        rango = worksheet.Range("A32:B32"); rango.Style.Fill.BackgroundColor = XLColor.Jonquil;

                        worksheet.Cell(31, 1).SetValue("Shipping Cobrado");

                        Labels = "Total Shipping  " + monthNombreProceso;
                        worksheet.Cell(32, 1).SetValue(Labels);

                        worksheet.Cell(16, 2).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        worksheet.Cell(16, 2).DataType = XLDataType.Number;
                        worksheet.Cell(16, 2).FormulaA1 = "=SUM(B3:B9)";



                        DateTime date7 = new DateTime(date.Year, date.Month, 1);
                        SqlConnection Conexion3 = new SqlConnection();
                        Conexion3.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                        Conexion3.Open();
                        SqlCommand cmd3 = new SqlCommand("DatosProvisionMensual", Conexion3);
                        cmd3.CommandType = CommandType.StoredProcedure;
                        cmd3.CommandTimeout = 7200; //in seconds
                        cmd3.Parameters.Add("@Operacion", SqlDbType.Int).Value = 7;
                        cmd3.Parameters.Add("@FechaProceso", SqlDbType.DateTime).Value = date7;
                        SqlDataReader reader3 = cmd3.ExecuteReader();

                        //Create a new DataTable.
                        System.Data.DataTable dt3 = new System.Data.DataTable("Resultado");

                        //Load DataReader into the DataTable.
                        dt3.Load(reader3);
                        //Conexion3.Close();

                        if (dt3.Rows.Count != 0)
                        {

                            worksheet.Cell(17, 2).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                            worksheet.Cell(17, 2).DataType = XLDataType.Number;
                            worksheet.Cell(17, 2).InsertData(dt3);// inserta Contenido
                        }

                        worksheet.Cell(18, 2).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        worksheet.Cell(18, 2).DataType = XLDataType.Number;
                        worksheet.Cell(18, 2).FormulaA1 = "=B16-B17";
                    
                        DateTime date6 = new DateTime(date.Year, date.Month, 1);
                        SqlConnection Conexion4 = new SqlConnection();
                        Conexion4.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                        Conexion4.Open();
                        SqlCommand cmd4 = new SqlCommand("DatosProvisionMensual", Conexion4);
                        cmd4.CommandType = CommandType.StoredProcedure;
                        cmd4.CommandTimeout = 7200; //in seconds
                        cmd4.Parameters.Add("@Operacion", SqlDbType.Int).Value = 8;
                        cmd4.Parameters.Add("@FechaProceso", SqlDbType.DateTime).Value = date6;
                        SqlDataReader reader4 = cmd4.ExecuteReader();

                        //Create a new DataTable.
                        System.Data.DataTable dt4 = new System.Data.DataTable("Resultado");

                        //Load DataReader into the DataTable.
                        dt4.Load(reader4);
                        //Conexion4.Close();

                        if (dt4.Rows.Count != 0)
                        {

                            worksheet.Cell(19, 2).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                            worksheet.Cell(19, 2).DataType = XLDataType.Number;
                            worksheet.Cell(19, 2).InsertData(dt4);// inserta Contenido
                        }
                    }

                    SqlConnection Conexion11 = new SqlConnection();
                    Conexion11.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion11.Open();
                    SqlCommand cmd11 = new SqlCommand("DatosProvisionMensual", Conexion11);
                    cmd11.CommandType = CommandType.StoredProcedure;
                    cmd11.CommandTimeout = 7200; //in seconds
                    cmd11.Parameters.Add("@Operacion", SqlDbType.Int).Value = 11;
                    SqlDataReader reader11 = cmd11.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt11 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt11.Load(reader11);
                    //Conexion11.Close();

                    if (dt11.Rows.Count != 0)
                    {

                        worksheet.Cell(3, 16).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        worksheet.Cell(3, 16).DataType = XLDataType.Number;
                        worksheet.Cell(3, 16).InsertData(dt11);// inserta Contenido
                    }

                    SqlConnection Conexion14 = new SqlConnection();
                    Conexion14.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion14.Open();
                    SqlCommand cmd14 = new SqlCommand("DatosProvisionMensual", Conexion14);
                    cmd14.CommandType = CommandType.StoredProcedure;
                    cmd14.CommandTimeout = 7200; //in seconds
                    cmd14.Parameters.Add("@Operacion", SqlDbType.Int).Value = 13;
                    SqlDataReader reader14 = cmd14.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt14 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt14.Load(reader14);
                    //Conexion14.Close();

                    if (dt14.Rows.Count != 0)
                    {

                        worksheet.Cell(3, 12).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        worksheet.Cell(3, 12).DataType = XLDataType.Number;
                        worksheet.Cell(3, 12).InsertData(dt14);// inserta Contenido
                    }

                    SqlConnection Conexion15 = new SqlConnection();
                    Conexion15.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion15.Open();
                    SqlCommand cmd15 = new SqlCommand("DatosProvisionMensual", Conexion14);
                    cmd15.CommandType = CommandType.StoredProcedure;
                    cmd15.CommandTimeout = 7200; //in seconds
                    cmd15.Parameters.Add("@Operacion", SqlDbType.Int).Value = 12;
                    SqlDataReader reader15 = cmd15.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt15 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt15.Load(reader15);
                    //Conexion15.Close();

                    if (dt15.Rows.Count != 0)
                    {

                        worksheet.Cell(3, 14).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        worksheet.Cell(3, 14).DataType = XLDataType.Number;
                        worksheet.Cell(3, 14).InsertData(dt15);// inserta Contenido
                    }

                    SqlConnection Conexion22 = new SqlConnection();
                    Conexion22.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion22.Open();
                    SqlCommand cmd22 = new SqlCommand("DatosProvisionMensual", Conexion22);
                    cmd22.CommandType = CommandType.StoredProcedure;
                    cmd22.CommandTimeout = 7200; //in seconds
                    cmd22.Parameters.Add("@Operacion", SqlDbType.Int).Value = 15;
                    SqlDataReader reader22 = cmd22.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt22 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt22.Load(reader22);
                    //Conexion22.Close();

                    if (dt22.Rows.Count != 0)
                    {

                        worksheet.Cell(4, 14).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        worksheet.Cell(4, 14).DataType = XLDataType.Number;
                        worksheet.Cell(4, 14).InsertData(dt22);// inserta Contenido
                    }

                    SqlConnection Conexion29 = new SqlConnection();
                    Conexion29.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion29.Open();
                    SqlCommand cmd29 = new SqlCommand("DatosProvisionMensual", Conexion29);
                    cmd29.CommandType = CommandType.StoredProcedure;
                    cmd29.CommandTimeout = 7200; //in seconds
                    cmd29.Parameters.Add("@Operacion", SqlDbType.Int).Value = 20;
                    SqlDataReader reader29 = cmd29.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt29 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt29.Load(reader29);
                    //Conexion29.Close();

                    if (dt29.Rows.Count != 0)
                    {

                        worksheet.Cell(5, 14).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        worksheet.Cell(5, 14).DataType = XLDataType.Number;
                        worksheet.Cell(5, 14).InsertData(dt29);// inserta Contenido
                    }

                    SqlConnection Conexion30 = new SqlConnection();
                    Conexion30.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion30.Open();
                    SqlCommand cmd30 = new SqlCommand("DatosProvisionMensual", Conexion30);
                    cmd30.CommandType = CommandType.StoredProcedure;
                    cmd30.CommandTimeout = 7200; //in seconds
                    cmd30.Parameters.Add("@Operacion", SqlDbType.Int).Value = 21;
                    SqlDataReader reader30 = cmd30.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt30 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt30.Load(reader30);
                    //Conexion30.Close();

                    if (dt30.Rows.Count != 0)
                    {

                        worksheet.Cell(5, 17).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        worksheet.Cell(5, 17).DataType = XLDataType.Number;
                        worksheet.Cell(5, 17).InsertData(dt30);// inserta Contenido
                    }

                    SqlConnection Conexion17 = new SqlConnection();
                    Conexion17.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion17.Open();
                    SqlCommand cmd17 = new SqlCommand("DatosProvisionMensual", Conexion17);
                    cmd17.CommandType = CommandType.StoredProcedure;
                    cmd17.CommandTimeout = 7200; //in seconds
                    cmd17.Parameters.Add("@Operacion", SqlDbType.Int).Value = 14;
                    SqlDataReader reader17 = cmd17.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt17 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt17.Load(reader17);
                    //Conexion17.Close();

                    if (dt17.Rows.Count != 0)
                    {

                        worksheet.Cell(3, 17).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        worksheet.Cell(3, 17).DataType = XLDataType.Number;
                        worksheet.Cell(3, 17).InsertData(dt17);// inserta Contenido
                    }

                    SqlConnection Conexion23 = new SqlConnection();
                    Conexion23.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion23.Open();
                    SqlCommand cmd23 = new SqlCommand("DatosProvisionMensual", Conexion23);
                    cmd23.CommandType = CommandType.StoredProcedure;
                    cmd23.CommandTimeout = 7200; //in seconds
                    cmd23.Parameters.Add("@Operacion", SqlDbType.Int).Value = 16;
                    SqlDataReader reader23 = cmd23.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt23 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt23.Load(reader23);
                    //Conexion23.Close();

                    if (dt23.Rows.Count != 0)
                    {

                        worksheet.Cell(4, 17).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        worksheet.Cell(4, 17).DataType = XLDataType.Number;
                        worksheet.Cell(4, 17).InsertData(dt23);// inserta Contenido
                    }


                    //SqlConnection Conexion16 = new SqlConnection();
                    //Conexion16.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    //Conexion16.Open();
                    //SqlCommand cmd16 = new SqlCommand("DatosProvisionMensual", Conexion16);
                    //cmd16.CommandType = CommandType.StoredProcedure;
                    //cmd16.CommandTimeout = 7200; //in seconds
                    //cmd16.Parameters.Add("@Operacion", SqlDbType.Int).Value = 10;
                    //SqlDataReader reader16 = cmd16.ExecuteReader();
                    //
                    ////Create a new DataTable.
                    //System.Data.DataTable dt16 = new System.Data.DataTable("Resultado");
                    //
                    ////Load DataReader into the DataTable.
                    //dt16.Load(reader16);
                    //
                    //if (dt16.Rows.Count != 0)
                    //{
                    //    // SE DEBE CALCULAR EL PORCENTAJE DE NO COBRADO REPRESENTA DEL TOTAL DE VENTAS
                    //    worksheet.Cell(26, 3).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                    //    worksheet.Cell(26, 3).DataType = XLDataType.Number;
                    //    worksheet.Cell(26, 3).InsertData(dt16);// inserta Contenido
                    //}

                    worksheet.Cell(9, 19).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                    worksheet.Cell(9, 19).DataType = XLDataType.Number;
                    worksheet.Cell(9, 19).FormulaA1 = "=SUM(S3:S8)";

                    worksheet.Cell(9, 20).Style.NumberFormat.Format = "[$$-en-US] #,##0.000000";
                    worksheet.Cell(9, 20).DataType = XLDataType.Number;
                    worksheet.Cell(9, 20).FormulaA1 = "=(SUM(E3:I7)+R3)/S9";

                    worksheet.Cell(9, 21).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                    worksheet.Cell(9, 21).DataType = XLDataType.Number;
                    worksheet.Cell(9, 21).FormulaA1 = "=SUM(U3:U7)+SUM(B9:B11)";

                    worksheet.Cell(9, 22).Style.NumberFormat.Format = "[$$-en-US] #,##0.000000";
                    worksheet.Cell(9, 22).DataType = XLDataType.Number;
                    worksheet.Cell(9, 22).FormulaA1 = "=U9/S9";

                    if (comboBox2.Text != "SI")
                    {
                        worksheet.Cell(28, 3).Style.NumberFormat.Format = "[$$-en-US] #,##0.000000";
                        worksheet.Cell(28, 3).DataType = XLDataType.Number;
                        worksheet.Cell(28, 3).FormulaA1 = "=T9";

                        worksheet.Cell(20, 2).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        worksheet.Cell(20, 2).DataType = XLDataType.Number;
                        worksheet.Cell(20, 2).FormulaA1 = "=B18-B19";


                        worksheet.Cell(27, 2).FormulaA1 = "=I28";
                    }

                    Labels = "Sales";
                    worksheet.Cell(2, 19).SetValue(Labels);

                    SqlConnection Conexion25 = new SqlConnection();
                    Conexion25.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion25.Open();
                    SqlCommand cmd25 = new SqlCommand("DatosProvisionMensual", Conexion25);
                    cmd25.CommandType = CommandType.StoredProcedure;
                    cmd25.CommandTimeout = 7200; //in seconds
                    cmd25.Parameters.Add("@Operacion", SqlDbType.Int).Value = 17;
                    SqlDataReader reader25 = cmd25.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt25 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt25.Load(reader25);
                    //Conexion25.Close();

                    foreach (DataRow row in dt25.Rows)
                    {
                        if (row["Carrier"].ToString() == "FEDEX")
                        {
                            worksheet.Cell(3, 19).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                            worksheet.Cell(3, 19).DataType = XLDataType.Number;
                            float Valor = (float)Convert.ToDouble(row["Ventas"].ToString());
                            worksheet.Cell(3, 19).SetValue(Valor);
                        }

                        if (row["Carrier"].ToString() == "UPS")
                        {
                            worksheet.Cell(4, 19).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                            worksheet.Cell(4, 19).DataType = XLDataType.Number;
                            float Valor = (float)Convert.ToDouble(row["Ventas"].ToString());
                            worksheet.Cell(4, 19).SetValue(Valor);
                        }

                        if (row["Carrier"].ToString() == "ENDICIA")
                        {
                            worksheet.Cell(5, 19).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                            worksheet.Cell(5, 19).DataType = XLDataType.Number;
                            float Valor = (float)Convert.ToDouble(row["Ventas"].ToString());
                            worksheet.Cell(5, 19).SetValue(Valor);
                        }
                    }


                    DateTime date10 = new DateTime(date.Year, date.Month, 1);
                    SqlConnection Conexion26 = new SqlConnection();
                    Conexion26.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion26.Open();
                    SqlCommand cmd26 = new SqlCommand("DatosProvisionMensual", Conexion26);
                    cmd26.CommandType = CommandType.StoredProcedure;
                    cmd26.CommandTimeout = 7200; //in seconds
                    cmd26.Parameters.Add("@Operacion", SqlDbType.Int).Value = 18;
                    cmd26.Parameters.Add("@FechaProceso", SqlDbType.DateTime).Value = date10;
                    SqlDataReader reader26 = cmd26.ExecuteReader();
                    float TotalVentas = 0;

                    //Create a new DataTable.
                    System.Data.DataTable dt26 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt26.Load(reader26);
                    //Conexion26.Close();

                    foreach (DataRow row in dt26.Rows)
                    {

                        TotalVentas = (float)Convert.ToDouble(row["TotalVentas"].ToString());
                     
                    }

                    Labels = "Shipping % Over Sales";
                    worksheet.Cell(2, 20).SetValue(Labels);
                    worksheet.Cell(3, 20).FormulaA1 = "=SUM(E3+F3+G3+H3+I3+R3)/S3";
                    worksheet.Cell(4, 20).FormulaA1 = "=SUM(E4+F4+G4+H4+I4+R4)/S4";
                    worksheet.Cell(5, 20).FormulaA1 = "=SUM(E5+F5+G5+H5+I5+R5)/S5";

                    Labels = "Overall Spend";
                    worksheet.Cell(2, 21).SetValue(Labels);
                    worksheet.Cell(3, 21).FormulaA1 = "=SUM(E3+F3+G3+H3+I3+J3+K3+L3+M3+N3+O3+P3+Q3+R3)";
                    worksheet.Cell(4, 21).FormulaA1 = "=SUM(E4+F4+G4+H4+I4+J4+K4+L4+M4+N4+O4+P4+Q4+R4)";
                    worksheet.Cell(5, 21).FormulaA1 = "=SUM(E5+F5+G5+H5+I5+J5+K5+L5+M5+N5+O5+P5+Q5+R5)";

                    Labels = "Shipping % Overall Spend";
                    worksheet.Cell(2, 22).SetValue(Labels);
                    worksheet.Cell(3, 22).FormulaA1 = "=U3/S3";
                    worksheet.Cell(4, 22).FormulaA1 = "=U4/S4";
                    worksheet.Cell(5, 22).FormulaA1 = "=U5/S5";


                    DateTime Fecha5;
                    DateTime Fecha25;

                    //int MesSigue1 = 0;
                    //int anioproc1 = 0;
                    //MesSigue1 = date.AddMonths(1);
                    //MesSigue1 = MesSigue + 1;
                    //
                    //if (MesSigue1 > 12)
                    //{
                    //    MesSigue1 = 1;
                    //    anioproc1 = date.Year + 1;
                    //}
                    //else
                    //{
                    //    anioproc1 = date.Year;
                    //}
                    DateTime FechaComienzo = date.AddMonths(1);
                    DateTime date5 = new DateTime(FechaComienzo.Year, FechaComienzo.Month, 1);

                    Fecha5 = new DateTime(date5.Year, date5.Month, 1);
                    //int MesIncremento = date5.Month + 1;
                    //int AnioIncremento = date5.Year;
                    //
                    //if (MesIncremento > 12)
                    //{
                    //    MesIncremento = 1;
                    //    AnioIncremento = AnioIncremento + 1;
                    //}

                    DateTime FechaComienzo1 = date5.AddMonths(1);
                    Fecha25 = new DateTime(FechaComienzo1.Year, FechaComienzo1.Month, 1).AddDays(-1);
                    
                    //if (comboBox2.Text != "SI")
                    //{

                        SqlConnection ConexionFEDEX1 = new SqlConnection();
                        ConexionFEDEX1.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                        string sqlGenericofedex1 = "ObtienMontosFEDEX";
                        SqlCommand commandFEDEX1 = new SqlCommand(sqlGenericofedex1, ConexionFEDEX1);
                        commandFEDEX1.Parameters.Add("@FechaInicio", SqlDbType.DateTime).Value = Fecha5;
                        commandFEDEX1.Parameters.Add("@FechaFin", SqlDbType.DateTime).Value = Fecha25;
                        commandFEDEX1.CommandType = CommandType.StoredProcedure;
                        commandFEDEX1.CommandTimeout = 7200; //in seconds
                        ConexionFEDEX1.Open();
                        commandFEDEX1.ExecuteNonQuery();
                        ConexionFEDEX1.Close();



                        SqlConnection Conexion5 = new SqlConnection();
                        Conexion5.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                        SqlConnection ConexionGenerico5 = new SqlConnection();
                        ConexionGenerico5.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                        string sqlGenerico5 = "ProvisionMensual";
                        SqlCommand commandGenerico5 = new SqlCommand(sqlGenerico5, ConexionGenerico5);
                        commandGenerico5.Parameters.Add("@FechaInicio", SqlDbType.DateTime).Value = Fecha5;
                        commandGenerico5.Parameters.Add("@FechaFin", SqlDbType.DateTime).Value = Fecha25;
                        commandGenerico5.Parameters.Add("@Mes", SqlDbType.VarChar).Value = "ACTUAL";
                        commandGenerico5.CommandType = CommandType.StoredProcedure;
                        commandGenerico5.CommandTimeout = 7200; //in seconds
                        ConexionGenerico5.Open();
                        commandGenerico5.ExecuteNonQuery();
                        ConexionGenerico5.Close();

                    if (comboBox2.Text != "SI")
                    {

                        worksheet.Cell(28, 9).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        worksheet.Cell(28, 9).DataType = XLDataType.Number;
                        worksheet.Cell(28, 9).FormulaA1 = "=E28-G28";

                        // ejecuto sp que devuelve el crokis
                        // ----------------------------------
                        SqlConnection Conexion6 = new SqlConnection();
                        Conexion6.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                        Conexion6.Open();
                        SqlCommand cmd6 = new SqlCommand("DatosProvisionMensual", Conexion6);
                        cmd6.CommandType = CommandType.StoredProcedure;
                        cmd6.CommandTimeout = 7200; //in seconds
                        cmd6.Parameters.Add("@Operacion", SqlDbType.Int).Value = 1;
                        SqlDataReader reader6 = cmd6.ExecuteReader();

                        //Create a new DataTable.
                        System.Data.DataTable dt6 = new System.Data.DataTable("Resultado");

                        //Load DataReader into the DataTable.
                        dt6.Load(reader6);
                        //Conexion6.Close();

                        Fila = 4;
                        foreach (DataRow row in dt6.Rows)
                        {
                            foreach (DataColumn column in dt6.Columns)
                            {
                                worksheet.Cell(10, Fila).SetValue(column.ColumnName);
                                Fila = Fila + 1;
                            }
                            break;
                        }

                        if (dt6.Rows.Count != 0)
                        {
                            worksheet.Cell(11, 4).InsertData(dt6);// inserta Contenido
                        }

                        DateTime date9 = new DateTime(date.Year, date.Month, 1);

                        SqlConnection Conexion7 = new SqlConnection();
                        Conexion7.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                        Conexion7.Open();
                        SqlCommand cmd7 = new SqlCommand("DatosProvisionMensual", Conexion7);
                        cmd7.CommandType = CommandType.StoredProcedure;
                        cmd7.CommandTimeout = 7200; //in seconds
                        cmd7.Parameters.Add("@Operacion", SqlDbType.Int).Value = 8;
                        cmd7.Parameters.Add("@FechaProceso", SqlDbType.DateTime).Value = date9;
                        SqlDataReader reader7 = cmd7.ExecuteReader();

                        //Create a new DataTable.
                        System.Data.DataTable dt7 = new System.Data.DataTable("Resultado");

                        //Load DataReader into the DataTable.
                        dt7.Load(reader7);
                        //Conexion7.Close();

                        if (dt7.Rows.Count != 0)
                        {

                            worksheet.Cell(22, 2).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                            worksheet.Cell(22, 2).DataType = XLDataType.Number;
                            worksheet.Cell(22, 2).InsertData(dt7);// inserta Contenido
                        }

                        worksheet.Cell(31, 2).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        worksheet.Cell(31, 2).DataType = XLDataType.Number;
                        worksheet.Cell(31, 2).FormulaA1 = "=B19+B22";

                        worksheet.Cell(32, 2).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        worksheet.Cell(32, 2).DataType = XLDataType.Number;
                        worksheet.Cell(32, 2).FormulaA1 = "=B28+B31";

                        worksheet.Cell(27, 3).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        worksheet.Cell(27, 3).DataType = XLDataType.Number;
                        Labels = "100%";
                        worksheet.Cell(27, 3).SetValue(Labels);



                        worksheet.Cell(28, 2).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        worksheet.Cell(28, 2).DataType = XLDataType.Number;
                        worksheet.Cell(28, 2).FormulaA1 = "=(C28*B27)/C27";

                        worksheet.Cell(23, 2).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        worksheet.Cell(23, 2).DataType = XLDataType.Number;
                        worksheet.Cell(23, 2).FormulaA1 = "=B28";

                        worksheet.Cell(24, 2).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        worksheet.Cell(24, 2).DataType = XLDataType.Number;
                        worksheet.Cell(24, 2).FormulaA1 = "=B22+B23";
                    }

                        // SE DEBE CALCULAR EL PORCENTAJE DE NO COBRADO REPRESENTA DEL TOTAL DE VENTAS
                        Labels = "Freight";
                        worksheet.Cell(2, 11).SetValue(Labels);
                        Valor1 = Convert.ToDouble(textBox1.Text);
                        worksheet.Cell(3, 11).SetValue(Valor1);

                        Labels = "Returns";
                        worksheet.Cell(2, 12).SetValue(Labels);

                        Labels = "Adjustments";
                        worksheet.Cell(2, 13).SetValue(Labels);

                        Labels = "Cancelled PO´s";
                        worksheet.Cell(2, 14).SetValue(Labels);

                    if (comboBox2.Text != "SI")
                    {
                            DateTime date13 = new DateTime(date.Year, date.Month, 1);

                            SqlConnection Conexion36 = new SqlConnection();
                            Conexion36.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                            Conexion36.Open();
                            SqlCommand cmd36 = new SqlCommand("DatosProvisionMensual", Conexion36);
                            cmd36.CommandType = CommandType.StoredProcedure;
                            cmd36.CommandTimeout = 7200; //in seconds
                            cmd36.Parameters.Add("@Operacion", SqlDbType.Int).Value = 22;
                            cmd36.Parameters.Add("@FechaProceso", SqlDbType.DateTime).Value = date13;
                            SqlDataReader reader36 = cmd36.ExecuteReader();

                            //Create a new DataTable.
                            System.Data.DataTable dt36 = new System.Data.DataTable("Resultado");

                            //Load DataReader into the DataTable.
                            dt36.Load(reader36);

                            if (dt36.Rows.Count != 0)
                            {

                                worksheet.Cell(5, 13).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                                worksheet.Cell(5, 13).DataType = XLDataType.Number;
                                worksheet.Cell(5, 13).InsertData(dt36);// inserta Contenido
                            }

                        //Conexion36.Close();
                    }


                        Labels = "UPS Parcel";
                        worksheet.Cell(2, 15).SetValue(Labels);
                        Valor1 = Convert.ToDouble(textBox6.Text);
                        worksheet.Cell(3, 15).SetValue(Valor1);

                        Labels = "Weekly Fees";
                        worksheet.Cell(2, 16).SetValue(Labels);


                    if (comboBox2.Text != "SI")
                    {
                        Labels = "Weekly Fees";
                        worksheet.Cell(10, 16).SetValue(Labels);

                        SqlConnection Conexion20 = new SqlConnection();
                        Conexion20.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                        Conexion20.Open();
                        SqlCommand cmd20 = new SqlCommand("DatosProvisionMensual", Conexion20);
                        cmd20.CommandType = CommandType.StoredProcedure;
                        cmd20.CommandTimeout = 7200; //in seconds
                        cmd20.Parameters.Add("@Operacion", SqlDbType.Int).Value = 11;
                        SqlDataReader reader20 = cmd20.ExecuteReader();

                        //Create a new DataTable.
                        System.Data.DataTable dt20 = new System.Data.DataTable("Resultado");

                        //Load DataReader into the DataTable.
                        dt20.Load(reader20);
                        //Conexion20.Close();

                        if (dt20.Rows.Count != 0)
                        {

                            worksheet.Cell(11, 16).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                            worksheet.Cell(11, 16).DataType = XLDataType.Number;
                            worksheet.Cell(11, 16).InsertData(dt20);// inserta Contenido
                        }

                        Labels = "Replacements";
                        worksheet.Cell(10, 17).SetValue(Labels);

                        SqlConnection Conexion21 = new SqlConnection();
                        Conexion21.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                        Conexion21.Open();
                        SqlCommand cmd21 = new SqlCommand("DatosProvisionMensual", Conexion21);
                        cmd21.CommandType = CommandType.StoredProcedure;
                        cmd21.CommandTimeout = 7200; //in seconds
                        cmd21.Parameters.Add("@Operacion", SqlDbType.Int).Value = 14;
                        SqlDataReader reader21 = cmd21.ExecuteReader();

                        //Create a new DataTable.
                        System.Data.DataTable dt21 = new System.Data.DataTable("Resultado");

                        //Load DataReader into the DataTable.
                        dt21.Load(reader21);
                        //Conexion21.Close();

                        if (dt21.Rows.Count != 0)
                        {

                            worksheet.Cell(11, 17).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                            worksheet.Cell(11, 17).DataType = XLDataType.Number;
                            worksheet.Cell(11, 17).InsertData(dt21);// inserta Contenido
                        }
                    }
                        Labels = "Replacements";
                        worksheet.Cell(2, 17).SetValue(Labels);

                        Labels = "Shipware Sales";
                        worksheet.Cell(2, 18).SetValue(Labels);
                        worksheet.Cell(3, 18).FormulaA1 = "=((B3)*0.1685)*0.265";

                    if (comboBox2.Text != "SI")
                    {
                        Labels = "Shipware Sales";
                        worksheet.Cell(10, 18).SetValue(Labels);
                        worksheet.Cell(11, 18).FormulaA1 = "=((SUM(F11:Q11))*0.1685)*0.265";

                        Labels = "Freight";
                        worksheet.Cell(10, 11).SetValue(Labels);

                        Labels = "Returns";
                        worksheet.Cell(10, 12).SetValue(Labels);

                        SqlConnection Conexion18 = new SqlConnection();
                        Conexion18.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                        Conexion18.Open();
                        SqlCommand cmd18 = new SqlCommand("DatosProvisionMensual", Conexion18);
                        cmd18.CommandType = CommandType.StoredProcedure;
                        cmd18.CommandTimeout = 7200; //in seconds
                        cmd18.Parameters.Add("@Operacion", SqlDbType.Int).Value = 13;
                        SqlDataReader reader18 = cmd18.ExecuteReader();

                        //Create a new DataTable.
                        System.Data.DataTable dt18 = new System.Data.DataTable("Resultado");

                        //Load DataReader into the DataTable.
                        dt18.Load(reader18);
                        //Conexion18.Close();


                        if (dt18.Rows.Count != 0)
                        {

                            worksheet.Cell(11, 12).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                            worksheet.Cell(11, 12).DataType = XLDataType.Number;
                            worksheet.Cell(11, 12).InsertData(dt18);// inserta Contenido
                        }


                        Labels = "Adjustments";
                        worksheet.Cell(10, 13).SetValue(Labels);

                        Labels = "Cancelled PO´s";
                        worksheet.Cell(10, 14).SetValue(Labels);

                        SqlConnection Conexion19 = new SqlConnection();
                        Conexion19.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                        Conexion19.Open();
                        SqlCommand cmd19 = new SqlCommand("DatosProvisionMensual", Conexion19);
                        cmd19.CommandType = CommandType.StoredProcedure;
                        cmd19.CommandTimeout = 7200; //in seconds
                        cmd19.Parameters.Add("@Operacion", SqlDbType.Int).Value = 12;
                        SqlDataReader reader19 = cmd19.ExecuteReader();

                        //Create a new DataTable.
                        System.Data.DataTable dt19 = new System.Data.DataTable("Resultado");

                        //Load DataReader into the DataTable.
                        dt19.Load(reader19);
                        //Conexion19.Close();


                        if (dt19.Rows.Count != 0)
                        {

                            worksheet.Cell(11, 14).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                            worksheet.Cell(11, 14).DataType = XLDataType.Number;
                            worksheet.Cell(11, 14).InsertData(dt19);// inserta Contenido
                        }

                        Labels = "UPS Parcel";
                        worksheet.Cell(10, 15).SetValue(Labels);
                    }


                    worksheet.Cell(3, 2).FormulaA1 = "=SUM(E3:Q3)";
                    worksheet.Cell(4, 2).FormulaA1 = "=SUM(E4:Q4)";
                    worksheet.Cell(5, 2).FormulaA1 = "=SUM(E5:Q5)";

                    worksheet.Cell(3, 2).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                    worksheet.Cell(3, 2).DataType = XLDataType.Number;
                    worksheet.Cell(4, 2).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                    worksheet.Cell(4, 2).DataType = XLDataType.Number;
                    worksheet.Cell(5, 2).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                    worksheet.Cell(5, 2).DataType = XLDataType.Number;
                    worksheet.Cell(6, 2).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                    worksheet.Cell(6, 2).DataType = XLDataType.Number;
                    worksheet.Cell(7, 2).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                    worksheet.Cell(7, 2).DataType = XLDataType.Number;
                    worksheet.Cell(8, 2).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                    worksheet.Cell(8, 2).DataType = XLDataType.Number;
                    worksheet.Cell(9, 2).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                    worksheet.Cell(9, 2).DataType = XLDataType.Number;
                    


                    List<string[]> titles8 = new List<string[]> { new string[] {  "Mes", "Monto", "TotalSales", "Porcentaje" } };

                    if (comboBox2.Text != "SI")
                    {
                        rango = worksheet.Range("A36:D36"); rango.Style.Fill.BackgroundColor = XLColor.BlueGray;
                        rango = worksheet.Range("A37:D37"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                        rango = worksheet.Range("A38:D38"); rango.Style.Fill.BackgroundColor = XLColor.PowderBlue;
                        rango = worksheet.Range("A39:D39"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                        rango = worksheet.Range("A40:D40"); rango.Style.Fill.BackgroundColor = XLColor.PowderBlue;
                        rango = worksheet.Range("A41:D41"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                        rango = worksheet.Range("A42:D42"); rango.Style.Fill.BackgroundColor = XLColor.PowderBlue;

                        worksheet.Cell(36, 1).InsertData(titles8); //insert titles to one row
                    }
                    else
                    {
                        rango = worksheet.Range("A12:D12"); rango.Style.Fill.BackgroundColor = XLColor.BlueGray;
                        rango = worksheet.Range("A13:D13"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                        rango = worksheet.Range("A14:B14"); rango.Style.Fill.BackgroundColor = XLColor.PowderBlue;
                        rango = worksheet.Range("A15:B15"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                        rango = worksheet.Range("A16:B16"); rango.Style.Fill.BackgroundColor = XLColor.PowderBlue;
                        rango = worksheet.Range("A17:B17"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                        rango = worksheet.Range("A18:B18"); rango.Style.Fill.BackgroundColor = XLColor.PowderBlue;

                        worksheet.Cell(12, 1).InsertData(titles8); //insert titles to one row
                    }

                    DateTime date18 = new DateTime(date.Year, date.Month, 1);

                    SqlConnection Conexion39 = new SqlConnection();
                    Conexion39.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion39.Open();
                    SqlCommand cmd39 = new SqlCommand("DatosProvisionMensual", Conexion39);
                    cmd39.CommandType = CommandType.StoredProcedure;
                    cmd39.CommandTimeout = 7200; //in seconds
                    
                    if (comboBox2.Text != "SI")
                    {
                        cmd39.Parameters.Add("@Operacion", SqlDbType.Int).Value = 23;
                    }
                    else 
                    {
                        cmd39.Parameters.Add("@Operacion", SqlDbType.Int).Value = 231;
                    }

                    cmd39.Parameters.Add("@FechaProceso", SqlDbType.DateTime).Value = date18;
                    SqlDataReader reader39 = cmd39.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt39 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt39.Load(reader39);
                    //Conexion39.Close();

                    if (dt39.Rows.Count != 0)
                    {
                        if (comboBox2.Text != "SI")
                        {
                            worksheet.Cell(37, 1).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                            worksheet.Cell(37, 1).DataType = XLDataType.Number;
                            worksheet.Cell(37, 1).InsertData(dt39);// inserta Contenido
                        }
                        else
                        {
                            worksheet.Cell(13, 1).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                            worksheet.Cell(13, 1).DataType = XLDataType.Number;
                            worksheet.Cell(13, 1).InsertData(dt39);// inserta Contenido
                        }
                    }
                    

                    List<string[]> titles10 = new List<string[]> { new string[] { "SalesChannelName", "TotalSales", "Monto", "Procentaje", "Conteo" } };

                    if (comboBox2.Text != "SI")
                    {
                        rango = worksheet.Range("A46:E46"); rango.Style.Fill.BackgroundColor = XLColor.BlueGray;
                        rango = worksheet.Range("A47:E47"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                        rango = worksheet.Range("A48:E48"); rango.Style.Fill.BackgroundColor = XLColor.PowderBlue;
                        rango = worksheet.Range("A49:E49"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                        rango = worksheet.Range("A50:E50"); rango.Style.Fill.BackgroundColor = XLColor.PowderBlue;
                        rango = worksheet.Range("A51:E51"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;


                        worksheet.Cell(46, 1).InsertData(titles10); //insert titles to one row
                    }
                    else 
                    {
                        rango = worksheet.Range("A24:E24"); rango.Style.Fill.BackgroundColor = XLColor.BlueGray;
                        rango = worksheet.Range("A25:E25"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                        rango = worksheet.Range("A26:E26"); rango.Style.Fill.BackgroundColor = XLColor.PowderBlue;
                        rango = worksheet.Range("A27:E27"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                        rango = worksheet.Range("A28:E28"); rango.Style.Fill.BackgroundColor = XLColor.PowderBlue;
                        rango = worksheet.Range("A29:E29"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;

                        worksheet.Cell(24, 1).InsertData(titles10); //insert titles to one row
                    }

                    DateTime date19 = new DateTime(date.Year, date.Month, 1);

                    SqlConnection Conexion40 = new SqlConnection();
                    Conexion40.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion40.Open();
                    SqlCommand cmd40 = new SqlCommand("DatosProvisionMensual", Conexion40);
                    cmd40.CommandType = CommandType.StoredProcedure;
                    cmd40.CommandTimeout = 7200; //in seconds
                        
                    if (comboBox2.Text != "SI")
                    {
                        cmd40.Parameters.Add("@Operacion", SqlDbType.Int).Value = 24;
                    }
                    else
                    {
                        cmd40.Parameters.Add("@Operacion", SqlDbType.Int).Value = 241;
                    }

                    cmd40.Parameters.Add("@FechaProceso", SqlDbType.DateTime).Value = date19;
                    SqlDataReader reader40 = cmd40.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt40 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt40.Load(reader40);
                    //Conexion40.Close();

                    if (dt40.Rows.Count != 0)
                    {
                        if (comboBox2.Text != "SI")
                        {
                            worksheet.Cell(47, 1).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                            worksheet.Cell(47, 1).DataType = XLDataType.Number;
                            worksheet.Cell(47, 1).InsertData(dt40);// inserta Contenido
                        }
                        else
                        {
                            worksheet.Cell(25, 1).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                            worksheet.Cell(25, 1).DataType = XLDataType.Number;
                            worksheet.Cell(25, 1).InsertData(dt40);// inserta Contenido
                        }
                    }
                    

                    if (comboBox2.Text != "SI")
                    {

                        DateTime date14 = new DateTime(date.Year, date.Month, 1);

                        SqlConnection Conexion38 = new SqlConnection();
                        Conexion38.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                        Conexion38.Open();
                        SqlCommand cmd38 = new SqlCommand("DatosProvisionMensual", Conexion38);
                        cmd38.CommandType = CommandType.StoredProcedure;
                        cmd38.CommandTimeout = 7200; //in seconds
                        cmd38.Parameters.Add("@Operacion", SqlDbType.Int).Value = 25;
                        cmd38.Parameters.Add("@FechaProceso", SqlDbType.DateTime).Value = date14;
                        SqlDataReader reader38 = cmd38.ExecuteReader();

                        //Create a new DataTable.
                        System.Data.DataTable dt38 = new System.Data.DataTable("Resultado");

                        //Load DataReader into the DataTable.
                        dt38.Load(reader38);

                        if (dt38.Rows.Count != 0)
                        {

                            worksheet.Cell(28, 9).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                            worksheet.Cell(28, 9).DataType = XLDataType.Number;
                            worksheet.Cell(28, 9).InsertData(dt38);// inserta Contenido
                        }

                        worksheet.Cell(28, 7).SetValue(0);
                        worksheet.Cell(28, 8).SetValue(0);
                    }

                    worksheet.Column(1).AdjustToContents(3);
                    worksheet.Column(2).AdjustToContents(2);
                    worksheet.Column(3).AdjustToContents(2);

                    worksheet.Column(5).AdjustToContents(2);
                    worksheet.Column(6).AdjustToContents(2);
                    worksheet.Column(7).AdjustToContents(2);
                    worksheet.Column(8).AdjustToContents(2);
                    worksheet.Column(9).AdjustToContents(2);
                    worksheet.Column(10).AdjustToContents(2);
                    worksheet.Column(11).AdjustToContents(2);
                    worksheet.Column(12).AdjustToContents(2);
                    worksheet.Column(13).AdjustToContents(2);
                    worksheet.Column(14).AdjustToContents(2);
                    worksheet.Column(15).AdjustToContents(2);
                    worksheet.Column(16).AdjustToContents(2);
                    worksheet.Column(17).AdjustToContents(2);
                    worksheet.Column(18).AdjustToContents(2);
                    worksheet.Column(19).AdjustToContents(2);
                    worksheet.Column(20).AdjustToContents(2);
                    worksheet.Column(21).AdjustToContents(2);
                    worksheet.Column(22).AdjustToContents(2);
                    worksheet.Column(23).AdjustToContents(2);


                    var worksheet1 = workbook.Worksheets.Add("ZONAS");

                    List<string[]> titlesZONAS = new List<string[]> { new string[] { "FechaInicio","FechaFin" } };

                    worksheet1.Cell(1, 1).InsertData(titlesZONAS);

                    if (comboBox2.Text == "SI")
                    {

                        DateTime date14 = new DateTime(date.Year, date.Month, 1);

                        SqlConnection Conexion38 = new SqlConnection();
                        Conexion38.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                        Conexion38.Open();
                        SqlCommand cmd38 = new SqlCommand("GeneraZonaSemanal", Conexion38);
                        cmd38.CommandType = CommandType.StoredProcedure;
                        cmd38.CommandTimeout = 7200; //in seconds
                        SqlDataReader reader38 = cmd38.ExecuteReader();

                        //Create a new DataTable.
                        System.Data.DataTable dt38 = new System.Data.DataTable("Resultado");

                        //Load DataReader into the DataTable.
                        dt38.Load(reader38);

                        if (dt38.Rows.Count != 0)
                        {

                            //worksheet1.Cell(2, 1).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                            //worksheet1.Cell(2, 1).DataType = XLDataType.Number;
                            worksheet1.Cell(2, 1).InsertData(dt38);// inserta Contenido
                        }
                    }
                    else 
                    {
                        DateTime date14 = new DateTime(date.Year, date.Month, 1);

                        SqlConnection Conexion38 = new SqlConnection();
                        Conexion38.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                        Conexion38.Open();
                        SqlCommand cmd38 = new SqlCommand("GeneraZonamensual", Conexion38);
                        cmd38.CommandType = CommandType.StoredProcedure;
                        cmd38.CommandTimeout = 7200; //in seconds
                        SqlDataReader reader38 = cmd38.ExecuteReader();

                        //Create a new DataTable.
                        System.Data.DataTable dt38 = new System.Data.DataTable("Resultado");

                        //Load DataReader into the DataTable.
                        dt38.Load(reader38);

                        if (dt38.Rows.Count != 0)
                        {

                           // worksheet1.Cell(2, 1).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                            //worksheet1.Cell(2, 1).DataType = XLDataType.Number;
                            worksheet1.Cell(2, 1).InsertData(dt38);// inserta Contenido
                        }
                    }

                    /*pitney bowes*/
                    var worksheetpit = workbook.Worksheets.Add("PITNEYBOWESPROCESO");

                    List<string[]> titlespit = new List<string[]> { new string[] { "PO NUMBER", "TRACKING NUMBER", "Monto", "SalesOrderDate", "TotalSales", "Clasificacion", "SalesOrderNumber" } };

                    worksheetpit.Cell(1, 1).InsertData(titlespit);

                    SqlConnection Conexionpit = new SqlConnection();
                    Conexionpit.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexionpit.Open();
                    SqlCommand cmdpit = new SqlCommand("DetallePITNEYBOWESPROCESO", Conexionpit);
                    cmdpit.CommandType = CommandType.StoredProcedure;
                    cmdpit.CommandTimeout = 7200; //in seconds
                    SqlDataReader readerpit = cmdpit.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dtpit = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dtpit.Load(readerpit);

                    if (dtpit.Rows.Count != 0)
                    {

                        // worksheet1.Cell(2, 1).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        //worksheet1.Cell(2, 1).DataType = XLDataType.Number;
                        worksheetpit.Cell(2, 1).InsertData(dtpit);// inserta Contenido
                    }
                    /*pitney bowes*/

                    var worksheet2 = workbook.Worksheets.Add("FEDEXPROCESO");

                    List<string[]> titlesFEDEX = new List<string[]> { new string[] { "Track", "TRACKING_NUMBER", "CCN_ORDER_NUMBER", "Monto", "Mes", "TotalSales", "SalesOrderNumber","PO Original" } };

                    worksheet2.Cell(1, 1).InsertData(titlesFEDEX);

                    SqlConnection Conexion45 = new SqlConnection();
                    Conexion45.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion45.Open();
                    SqlCommand cmd45 = new SqlCommand("DetalleFEDEXPROCESO", Conexion45);
                    cmd45.CommandType = CommandType.StoredProcedure;
                    cmd45.CommandTimeout = 7200; //in seconds
                    SqlDataReader reader45 = cmd45.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt45 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt45.Load(reader45);

                    if (dt45.Rows.Count != 0)
                    {

                        // worksheet1.Cell(2, 1).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        //worksheet1.Cell(2, 1).DataType = XLDataType.Number;
                        worksheet2.Cell(2, 1).InsertData(dt45);// inserta Contenido
                    }

                    var worksheet3 = workbook.Worksheets.Add("UPSPROCESO");
                    
                    List<string[]> titlesUPS = new List<string[]> { new string[] { "PO NUMBER", "TRACKING NUMBER", "Monto", "SalesOrderDate", "TotalSales", "Clasificacion", "SalesOrderNumber" } };

                    worksheet3.Cell(1, 1).InsertData(titlesUPS);

                    SqlConnection Conexion46 = new SqlConnection();
                    Conexion46.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion46.Open();
                    SqlCommand cmd46 = new SqlCommand("DetalleUPSPROCESO", Conexion46);
                    cmd46.CommandType = CommandType.StoredProcedure;
                    cmd46.CommandTimeout = 7200; //in seconds
                    SqlDataReader reader46 = cmd46.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt46 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt46.Load(reader46);

                    if (dt46.Rows.Count != 0)
                    {

                        // worksheet1.Cell(2, 1).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        //worksheet1.Cell(2, 1).DataType = XLDataType.Number;
                        worksheet3.Cell(2, 1).InsertData(dt46);// inserta Contenido
                    }

                    var worksheet4 = workbook.Worksheets.Add("ENDICIAPROCESO");
                    
                    List<string[]> titlesENDICIA = new List<string[]> { new string[] { "PO NUMBER", "TRACKING NUMBER", "Monto", "SalesOrderDate", "TotalSales", "Clasificacion", "SalesOrderNumber" } };

                    worksheet4.Cell(1, 1).InsertData(titlesENDICIA);

                    SqlConnection Conexion47 = new SqlConnection();
                    Conexion47.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion47.Open();
                    SqlCommand cmd47 = new SqlCommand("DetalleENDICIAPROCESO", Conexion47);
                    cmd47.CommandType = CommandType.StoredProcedure;
                    cmd47.CommandTimeout = 7200; //in seconds
                    SqlDataReader reader47 = cmd47.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt47 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt47.Load(reader47);

                    if (dt47.Rows.Count != 0)
                    {

                        // worksheet1.Cell(2, 1).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                        //worksheet1.Cell(2, 1).DataType = XLDataType.Number;
                        worksheet4.Cell(2, 1).InsertData(dt47);// inserta Contenido
                    }

                    if (comboBox2.Text != "SI")
                    {

                        /*pitney bowes*/
                        var worksheetpitA = workbook.Worksheets.Add("PITNEYBOWESACTUAL");

                        List<string[]> titlespitA = new List<string[]> { new string[] { "PO NUMBER", "TRACKING NUMBER", "Monto", "SalesOrderDate", "TotalSales", "Clasificacion", "SalesOrderNumber" } };

                        worksheetpitA.Cell(1, 1).InsertData(titlespit);

                        SqlConnection ConexionpitA = new SqlConnection();
                        ConexionpitA.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                        ConexionpitA.Open();
                        SqlCommand cmdpitA = new SqlCommand("DetallePITNEYBOWESACTUAL", ConexionpitA);
                        cmdpitA.CommandType = CommandType.StoredProcedure;
                        cmdpitA.CommandTimeout = 7200; //in seconds
                        SqlDataReader readerpitA = cmdpitA.ExecuteReader();

                        //Create a new DataTable.
                        System.Data.DataTable dtpitA = new System.Data.DataTable("Resultado");

                        //Load DataReader into the DataTable.
                        dtpitA.Load(readerpitA);

                        if (dtpitA.Rows.Count != 0)
                        {

                            worksheetpitA.Cell(2, 1).InsertData(dtpitA);// inserta Contenido
                        }
                        /*pitney bowes*/

                        var worksheet5 = workbook.Worksheets.Add("FEDEXACTUAL");

                        List<string[]> titlesFEDEXactual = new List<string[]> { new string[] { "Track", "TRACKING_NUMBER", "CCN_ORDER_NUMBER", "Monto", "Mes", "TotalSales", "SalesOrderNumber","PO Original" } };

                        worksheet5.Cell(1, 1).InsertData(titlesFEDEXactual);

                        SqlConnection Conexion48 = new SqlConnection();
                        Conexion48.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                        Conexion48.Open();
                        SqlCommand cmd48 = new SqlCommand("DetalleFEDEXACTUAL", Conexion48);
                        cmd48.CommandType = CommandType.StoredProcedure;
                        cmd48.CommandTimeout = 7200; //in seconds
                        SqlDataReader reader48 = cmd48.ExecuteReader();

                        //Create a new DataTable.
                        System.Data.DataTable dt48 = new System.Data.DataTable("Resultado");

                        //Load DataReader into the DataTable.
                        dt48.Load(reader48);

                        if (dt48.Rows.Count != 0)
                        {

                            // worksheet1.Cell(2, 1).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                            //worksheet1.Cell(2, 1).DataType = XLDataType.Number;
                            worksheet5.Cell(2, 1).InsertData(dt48);// inserta Contenido
                        }

                        var worksheet6 = workbook.Worksheets.Add("UPSACTUAL");

                        List<string[]> titlesUPSACTUAL = new List<string[]> { new string[] { "PO NUMBER", "TRACKING NUMBER", "Monto", "SalesOrderDate", "TotalSales", "Clasificacion", "SalesOrderNumber" } };

                        worksheet6.Cell(1, 1).InsertData(titlesUPSACTUAL);

                        SqlConnection Conexion49 = new SqlConnection();
                        Conexion49.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                        Conexion49.Open();
                        SqlCommand cmd49 = new SqlCommand("DetalleUPSACTUAL", Conexion49);
                        cmd49.CommandType = CommandType.StoredProcedure;
                        cmd49.CommandTimeout = 7200; //in seconds
                        SqlDataReader reader49 = cmd49.ExecuteReader();

                        //Create a new DataTable.
                        System.Data.DataTable dt49 = new System.Data.DataTable("Resultado");

                        //Load DataReader into the DataTable.
                        dt49.Load(reader49);

                        if (dt49.Rows.Count != 0)
                        {

                            // worksheet1.Cell(2, 1).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                            //worksheet1.Cell(2, 1).DataType = XLDataType.Number;
                            worksheet6.Cell(2, 1).InsertData(dt49);// inserta Contenido
                        }

                        var worksheet7 = workbook.Worksheets.Add("ENDICIAACTUAL");

                        List<string[]> titlesENDICIAACTUAL = new List<string[]> { new string[] { "PO NUMBER", "TRACKING NUMBER", "Monto", "SalesOrderDate", "TotalSales", "Clasificacion", "SalesOrderNumber" } };

                        worksheet7.Cell(1, 1).InsertData(titlesENDICIAACTUAL);

                        SqlConnection Conexion50 = new SqlConnection();
                        Conexion50.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                        Conexion50.Open();
                        SqlCommand cmd50 = new SqlCommand("DetalleENDICIAACTUAL", Conexion50);
                        cmd50.CommandType = CommandType.StoredProcedure;
                        cmd50.CommandTimeout = 7200; //in seconds
                        SqlDataReader reader50 = cmd50.ExecuteReader();

                        //Create a new DataTable.
                        System.Data.DataTable dt50 = new System.Data.DataTable("Resultado");

                        //Load DataReader into the DataTable.
                        dt50.Load(reader50);

                        if (dt50.Rows.Count != 0)
                        {

                            // worksheet1.Cell(2, 1).Style.NumberFormat.Format = "[$$-en-US] #,##0.00";
                            //worksheet1.Cell(2, 1).DataType = XLDataType.Number;
                            worksheet7.Cell(2, 1).InsertData(dt50);// inserta Contenido
                        }
                    }

                    // analisis fedex
                    // --------------

                    // PROMEDIOS SERVICIOS
                    var worksheet8 = workbook.Worksheets.Add("PROMEDIOSERVICIO");
                    SqlCommand cmd51 = null;

                    List<string[]> titlesFEDEX1 = new List<string[]> { new string[] { "ServiceType", "Conteo", "NetChargeAmount", "PromedioServicio" } };

                    worksheet8.Cell(1, 1).InsertData(titlesFEDEX1);

                    SqlConnection Conexion51 = new SqlConnection();
                    Conexion51.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion51.Open();

                    // escoje el sp que le corresponde a mensual o semanal
                    // ---------------------------------------------------
                    if (comboBox2.Text != "SI")
                    {
                        cmd51 = new SqlCommand("AnalisisFEDEXMensual", Conexion51);
                    }
                    else
                    {
                        cmd51 = new SqlCommand("AnalisisFEDEXSemanal", Conexion51);
                    }

                    cmd51.CommandType = CommandType.StoredProcedure;
                    cmd51.CommandTimeout = 7200; //in seconds
                    cmd51.Parameters.Add("@Operacion", SqlDbType.Int).Value = 1;
                    
                    if (comboBox2.Text != "SI")
                    {
                        cmd51.Parameters.Add("@FechaInicio", SqlDbType.DateTime).Value = Fecha1;
                        cmd51.Parameters.Add("@FechaFin", SqlDbType.DateTime).Value = Fecha2;
                    }

                    // OBTIENE MONTOS MENORES A UNA LIBRA
                    // ----------------------------------
                    SqlDataReader reader51 = cmd51.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt51 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt51.Load(reader51);

                    if (dt51.Rows.Count != 0)
                    {
                        worksheet8.Cell(2, 1).InsertData(dt51);// inserta Contenido
                    }

                    var worksheet9 = workbook.Worksheets.Add("MENORA1LIBRA");
                    SqlCommand cmd52 = null;

                    List<string[]> titlesFEDEX2 = new List<string[]> { new string[] { "ServiceType", "Conteo", "NetChargeAmount", "PromedioServicio" } };

                    worksheet9.Cell(1, 1).InsertData(titlesFEDEX2);

                    SqlConnection Conexion52 = new SqlConnection();
                    Conexion52.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion52.Open();

                    // escoje el sp que le corresponde a mensual o semanal
                    // ---------------------------------------------------
                    if (comboBox2.Text != "SI")
                    {
                        cmd52 = new SqlCommand("AnalisisFEDEXMensual", Conexion52);
                    }
                    else
                    {
                        cmd52 = new SqlCommand("AnalisisFEDEXSemanal", Conexion52);
                    }

                    cmd52.CommandType = CommandType.StoredProcedure;
                    cmd52.CommandTimeout = 7200; //in seconds
                    cmd52.Parameters.Add("@Operacion", SqlDbType.Int).Value = 2;

                    if (comboBox2.Text != "SI")
                    {
                        cmd52.Parameters.Add("@FechaInicio", SqlDbType.DateTime).Value = Fecha1;
                        cmd52.Parameters.Add("@FechaFin", SqlDbType.DateTime).Value = Fecha2;
                    }


                    SqlDataReader reader52 = cmd52.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt52 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt52.Load(reader52);

                    if (dt52.Rows.Count != 0)
                    {
                        worksheet9.Cell(2, 1).InsertData(dt52);// inserta Contenido
                    }
                    // ----------------------------------------------------

                    /// OBTIENE PESO PROMEDIO
                    var worksheet10 = workbook.Worksheets.Add("PESOPROMEDIO");
                    SqlCommand cmd53 = null;

                    List<string[]> titlesFEDEX3 = new List<string[]> { new string[] { "ServiceType", "Conteo ", "NetChargeAmount", "PromedioServicio", "RatedWeightAmount", "DescuentoPromedio", "Tarifacompletapromedio", "Flatrate", "FlatratePromedio", "Tarifasinaccesorialspromedio ", "porcentajedescuento ", "Mayoresa20  ", "Porcentaje20 ", "Mayoresa10  ", "Porcentaje10" } };
                    

                    worksheet10.Cell(1, 1).InsertData(titlesFEDEX3);

                    SqlConnection Conexion53 = new SqlConnection();
                    Conexion53.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion53.Open();

                    // escoje el sp que le corresponde a mensual o semanal
                    // ---------------------------------------------------
                    if (comboBox2.Text != "SI")
                    {
                        cmd53 = new SqlCommand("AnalisisFEDEXMensual", Conexion53);
                    }
                    else
                    {
                        cmd53 = new SqlCommand("AnalisisFEDEXSemanal", Conexion53);
                    }

                    cmd53.CommandType = CommandType.StoredProcedure;
                    cmd53.CommandTimeout = 7200; //in seconds
                    cmd53.Parameters.Add("@Operacion", SqlDbType.Int).Value = 3;

                    if (comboBox2.Text != "SI")
                    {
                        cmd53.Parameters.Add("@FechaInicio", SqlDbType.DateTime).Value = Fecha1;
                        cmd53.Parameters.Add("@FechaFin", SqlDbType.DateTime).Value = Fecha2;
                    }


                    SqlDataReader reader53 = cmd53.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt53 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt53.Load(reader53);

                    if (dt53.Rows.Count != 0)
                    {
                        worksheet10.Cell(2, 1).InsertData(dt53);// inserta Contenido
                    }

                    // -----------------------------------------------------------------------
                    var worksheet11 = workbook.Worksheets.Add("CARGOSSEMANA");
                    SqlCommand cmd54 = null;

                    List<string[]> titlesFEDEX4 = new List<string[]> { new string[] { "ServiceType", "Conteo", "NetChargeAmount", "PromedioServicio", "EarnedDiscount", "FuelSurcharge", "PerformancePricing", "DeliveryAreaSurchargeExtended", "DeliveryAreaSurcharge", "USPSNonMachSurcharge", "Residential", "GraceDiscount", "DeclaredValue", "DASExtendedResi", "AdditionalHandling", "ParcelReLabelCharge", "IndirectSignature", "DASResi", "AddressCorrection", "DASExtendedComm", "OversizeCharge", "AHSDimensions", "PeakAHSCharge", "PeakOversizeCharge", "PeakSurcharge", "TemporarySurcharge ","WeeklyServiceChg", "DASAlaskaResi","WeekdayDelivery", "AdditionalHandling",  "AHSWeight","PrintReturnLabel","CourierPickupCharge","NDOCAutoComm" } };

                    worksheet11.Cell(1, 1).InsertData(titlesFEDEX4);

                    SqlConnection Conexion54 = new SqlConnection();
                    Conexion54.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion54.Open();

                    // escoje el sp que le corresponde a mensual o semanal
                    // ---------------------------------------------------
                    if (comboBox2.Text != "SI")
                    {
                        cmd54 = new SqlCommand("AnalisisFEDEXMensual", Conexion54);
                    }
                    else
                    {
                        cmd54 = new SqlCommand("AnalisisFEDEXSemanal", Conexion54);
                    }

                    cmd54.CommandType = CommandType.StoredProcedure;
                    cmd54.CommandTimeout = 7200; //in seconds
                    cmd54.Parameters.Add("@Operacion", SqlDbType.Int).Value = 4;

                    if (comboBox2.Text != "SI")
                    {
                        cmd54.Parameters.Add("@FechaInicio", SqlDbType.DateTime).Value = Fecha1;
                        cmd54.Parameters.Add("@FechaFin", SqlDbType.DateTime).Value = Fecha2;
                    }


                    SqlDataReader reader54 = cmd54.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt54 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt54.Load(reader54);

                    if (dt54.Rows.Count != 0)
                    {
                        worksheet11.Cell(2, 1).InsertData(dt54);// inserta Contenido
                    }

                    //---------------------------------------------------------------------
                    var worksheet12 = workbook.Worksheets.Add("UPSMAS1LIBRA");
                    SqlCommand cmd55 = null;

                    List<string[]> titlesFEDEX5 = new List<string[]> { new string[] { "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12", "F13", "F14", "F15", "F16", "F17", "F18", "F19", "F20", "F21", "F22", "F23", "F24", "F25", "F26", "F27", "F28", "F29", "F30", "F31", "F32", "F33", "F34", "F35", "F36", "F37", "F38", "F39", "F40", "F41", "F42", "FechaInsercion " } };

                    worksheet12.Cell(1, 1).InsertData(titlesFEDEX5);

                    SqlConnection Conexion55 = new SqlConnection();
                    Conexion55.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion55.Open();

                    // escoje el sp que le corresponde a mensual o semanal
                    // ---------------------------------------------------
                    if (comboBox2.Text != "SI")
                    {
                        cmd55 = new SqlCommand("AnalisisFEDEXMensual", Conexion55);
                    }
                    else
                    {
                        cmd55 = new SqlCommand("AnalisisFEDEXSemanal", Conexion55);
                    }

                    cmd55.CommandType = CommandType.StoredProcedure;
                    cmd55.CommandTimeout = 7200; //in seconds
                    cmd55.Parameters.Add("@Operacion", SqlDbType.Int).Value = 5;

                    if (comboBox2.Text != "SI")
                    {
                        cmd55.Parameters.Add("@FechaInicio", SqlDbType.DateTime).Value = Fecha1;
                        cmd55.Parameters.Add("@FechaFin", SqlDbType.DateTime).Value = Fecha2;
                    }


                    SqlDataReader reader55 = cmd55.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt55 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt55.Load(reader55);

                    if (dt55.Rows.Count != 0)
                    {
                        worksheet12.Cell(2, 1).InsertData(dt55);// inserta Contenido
                    }
                    //---------------------------------------------------------

                    var worksheet13 = workbook.Worksheets.Add("ENDICIAMAS1LIBRA");
                    SqlCommand cmd56 = null;

                    List<string[]> titlesFEDEX6 = new List<string[]> { new string[] { "PrintDate", "AmountPaid", "AdjAmount", "QuotedAmount", "Recipient", "Status", "TrackingNumber", "DateDelivered", "Carrier", "ClassService", "InsuredValue", "InsuranceID", "CostCode", "Weight", "ShipDate", "RefundType", "PrintedMessage", "User", "RefundRequestDate", "RefundStatus", "RefundRequested", "Reference1", "Reference2", "Reference3", "Reference4 " } };

                    worksheet13.Cell(1, 1).InsertData(titlesFEDEX6);

                    SqlConnection Conexion56 = new SqlConnection();
                    Conexion56.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion56.Open();

                    // escoje el sp que le corresponde a mensual o semanal
                    // ---------------------------------------------------
                    if (comboBox2.Text != "SI")
                    {
                        cmd56 = new SqlCommand("AnalisisFEDEXMensual", Conexion56);
                    }
                    else
                    {
                        cmd56 = new SqlCommand("AnalisisFEDEXSemanal", Conexion56);
                    }

                    cmd56.CommandType = CommandType.StoredProcedure;
                    cmd56.CommandTimeout = 7200; //in seconds
                    cmd56.Parameters.Add("@Operacion", SqlDbType.Int).Value = 6;

                    if (comboBox2.Text != "SI")
                    {
                        cmd56.Parameters.Add("@FechaInicio", SqlDbType.DateTime).Value = Fecha1;
                        cmd56.Parameters.Add("@FechaFin", SqlDbType.DateTime).Value = Fecha2;
                    }


                    SqlDataReader reader56 = cmd56.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt56 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt56.Load(reader56);

                    if (dt56.Rows.Count != 0)
                    {
                        worksheet13.Cell(2, 1).InsertData(dt56);// inserta Contenido
                    }


                    var worksheet14 = workbook.Worksheets.Add("ACCESORIAL");
                    SqlCommand cmd57 = null;

                    List<string[]> titlesFEDEX8 = new List<string[]> { new string[] { "ServiceType","NetChargeAmount","Accesorials","FLatRate","PromedioServicio","Conteo ","EarnedDiscount ","PerformancePricing  ","AverageFlateRate    ","AverageAccesorial   ","EarnedDiscount  ","PerformancePricing "} };

                    worksheet14.Cell(1, 1).InsertData(titlesFEDEX8);

                    SqlConnection Conexion57 = new SqlConnection();
                    Conexion57.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion57.Open();

                    // escoje el sp que le corresponde a mensual o semanal
                    // ---------------------------------------------------
                    if (comboBox2.Text != "SI")
                    {
                        cmd57 = new SqlCommand("ReportePorcentajesMensual", Conexion57);
                    }
                    else
                    {
                        cmd57 = new SqlCommand("ReportePorcentajes", Conexion57);
                    }

                    if (comboBox2.Text != "SI")
                    {
                        cmd57.Parameters.Add("@FechaInicio", SqlDbType.DateTime).Value = Fecha1;
                        cmd57.Parameters.Add("@FechaFin", SqlDbType.DateTime).Value = Fecha2;
                    }

                    cmd57.CommandType = CommandType.StoredProcedure;
                    cmd57.CommandTimeout = 7200; //in seconds

                    SqlDataReader reader57 = cmd57.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt57 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt57.Load(reader57);

                    if (dt57.Rows.Count != 0)
                    {
                        worksheet14.Cell(2, 1).InsertData(dt57);// inserta Contenido
                    }


                    //-----------------------------------------------------------------------
                    var worksheet15 = workbook.Worksheets.Add("DIASTRANSITO");
                    SqlCommand cmd58 = null;

                    List<string[]> titlesFEDEX9 = new List<string[]> { new string[] { "ServicesType", "DiasEntrega"} };

                    worksheet15.Cell(1, 1).InsertData(titlesFEDEX9);

                    SqlConnection Conexion58 = new SqlConnection();
                    Conexion58.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion58.Open();

                    // escoje el sp que le corresponde a mensual o semanal
                    // ---------------------------------------------------
                    if (comboBox2.Text != "SI")
                    {
                        cmd58 = new SqlCommand("DiasTransitoSemanal", Conexion58);
                    }
                    else
                    {
                        cmd58 = new SqlCommand("DiasTransitoSemanal", Conexion58);
                    }

                    cmd58.CommandType = CommandType.StoredProcedure;
                    cmd58.CommandTimeout = 7200; //in seconds

                    SqlDataReader reader58 = cmd58.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt58 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt58.Load(reader58);

                    if (dt58.Rows.Count != 0)
                    {
                        worksheet15.Cell(2, 1).InsertData(dt58);// inserta Contenido
                    }
                    //-------------------------------------------------------------------------
                    List<string[]> titulosfulfill = new List<string[]> { new string[] { "FechaInicio", "FechaFin" } };


                    if (comboBox2.Text != "SI")
                    {
                        worksheet1.Cell(52, 1).InsertData(titulosfulfill);
                    }
                    else
                    {
                        worksheet1.Cell(31, 1).InsertData(titulosfulfill);
                    }


                    SqlConnection Conexion61 = new SqlConnection();
                    Conexion61.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion61.Open();
                    SqlCommand cmd61 = new SqlCommand("DatosProvisionMensual", Conexion61);
                    cmd61.CommandType = CommandType.StoredProcedure;
                    cmd61.CommandTimeout = 7200; //in seconds

                    if (comboBox2.Text != "SI")
                    {
                        cmd61.Parameters.Add("@Operacion", SqlDbType.Int).Value = 26;
                    }
                    else
                    {
                        cmd61.Parameters.Add("@Operacion", SqlDbType.Int).Value = 261;
                    }

                    cmd61.Parameters.Add("@FechaProceso", SqlDbType.DateTime).Value = date19;
                    SqlDataReader reader61 = cmd61.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt61 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt61.Load(reader61);
                    //Conexion61.Close();

                    List<string[]> titlesFEDEX15 = new List<string[]> { new string[] { "FulfillmentChannelName", "LinkedFulfillmentChannelName", "Monto", "TotalSales", "Porcentaje", "CONTEO" } };

                    if (dt61.Rows.Count != 0)
                    {
                        if (comboBox2.Text != "SI")
                        {
                            worksheet.Cell(55, 1).InsertData(titlesFEDEX15);
                            worksheet.Cell(56, 1).InsertData(dt61);// inserta Contenido
                        }
                        else
                        {
                            worksheet.Cell(34, 1).InsertData(titlesFEDEX15);
                            worksheet.Cell(35, 1).InsertData(dt61);// inserta Contenido
                        }
                    }

                    if (comboBox2.Text != "SI")
                    {
                        rango = worksheet.Range("A55:F55"); rango.Style.Fill.BackgroundColor = XLColor.BlueGray;
                        rango = worksheet.Range("A56:F56"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                        rango = worksheet.Range("A57:F57"); rango.Style.Fill.BackgroundColor = XLColor.PowderBlue;
                        rango = worksheet.Range("A58:F58"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                        rango = worksheet.Range("A59:F59"); rango.Style.Fill.BackgroundColor = XLColor.PowderBlue;
                        rango = worksheet.Range("A60:F60"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                        rango = worksheet.Range("A61:F61"); rango.Style.Fill.BackgroundColor = XLColor.PowderBlue;
                        rango = worksheet.Range("A62:F62"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                        rango = worksheet.Range("A63:F63"); rango.Style.Fill.BackgroundColor = XLColor.PowderBlue;
                        rango = worksheet.Range("A64:F64"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;

                        worksheet.Cell(46, 1).InsertData(titles10); //insert titles to one row
                    }
                    else
                    {
                        rango = worksheet.Range("A34:F24"); rango.Style.Fill.BackgroundColor = XLColor.BlueGray;
                        rango = worksheet.Range("A35:F35"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                        rango = worksheet.Range("A36:F36"); rango.Style.Fill.BackgroundColor = XLColor.PowderBlue;
                        rango = worksheet.Range("A37:F37"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                        rango = worksheet.Range("A38:F38"); rango.Style.Fill.BackgroundColor = XLColor.PowderBlue;
                        rango = worksheet.Range("A39:F39"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                        rango = worksheet.Range("A40:F40"); rango.Style.Fill.BackgroundColor = XLColor.PowderBlue;
                        rango = worksheet.Range("A41:F41"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;
                        rango = worksheet.Range("A42:F42"); rango.Style.Fill.BackgroundColor = XLColor.PowderBlue;
                        rango = worksheet.Range("A43:F43"); rango.Style.Fill.BackgroundColor = XLColor.AliceBlue;

                        worksheet.Cell(24, 1).InsertData(titles10); //insert titles to one row
                    }

                    //-----------------------------------------------------------------------
                    string pathUPS = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                    
                    if (comboBox2.Text != "SI")
                    {
                        pathUPS = pathUPS + @"\ProvisionMensual" + monthName + DateTime.Now.ToString("MMddyyyyHHmmss") + ".xlsx";
                    }
                    else
                    {
                        pathUPS = pathUPS + @"\ProvisionSemanal" + monthName + DateTime.Now.ToString("MMddyyyyHHmmss") + ".xlsx";
                    }

                    workbook.SaveAs(pathUPS);
                }


                using (var workbook = new XLWorkbook())
                {

                    SqlConnection Conexion31 = new SqlConnection();
                    Conexion31.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion31.Open();
                    SqlCommand cmd31 = null;
                    var worksheet = workbook.Worksheets.Add("shippingbymarket");

                    if (comboBox2.Text == "SI")
                    {
                        cmd31 = new SqlCommand("sp_SHIPPINGBYMARKETsemanal", Conexion31);
                        List<string[]> titles = new List<string[]> { new string[] { "orderId", "totalSale", "sumOfQuantity", "EstimatedShiping", "date", "Channel", "sku", "SalesOrderNumber", "RealShipping", "ShipingSobreVentas", "EstvsReal", "TotalTracks", "RequestedServiceLevel", "ShippingServiceLevel", "FulfillmentLocationName", "ShipmentDate", "DeliveryDate", "DiasEntrega", "FulfillmentServiceLevel ","ZoneCode", "FulfillmentLocationName" } };
                        worksheet.Cell(1, 1).InsertData(titles); //insert titles to one row
                    }
                    else
                    {
                        cmd31 = new SqlCommand("sp_SHIPPINGBYMARKET", Conexion31);
                        List<string[]> titles = new List<string[]> { new string[] { "orderId", "totalSale", "sumOfQuantity", "EstimatedShiping", "date", "Channel", "sku", "RequestedServiceLevel", "ShippingServiceLevel", "FulfillmentLocationName", "SalesOrderNumber", "MontoTotal", "TotalTracks", "EstVsReal ", "ShipmentDate", "DeliveryDate", "DiasEntrega ", "FulfillmentServiceLevel", "ZoneCode", "FulfillmentLocationName" } };
                        worksheet.Cell(1, 1).InsertData(titles); //insert titles to one row
                    }
                    cmd31.CommandType = CommandType.StoredProcedure;
                    cmd31.CommandTimeout = 7200; //in seconds
                    cmd31.Parameters.Add("@BeginDate", SqlDbType.DateTime).Value = Fecha1;
                    cmd31.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = Fecha2;
                    SqlDataReader reader31 = cmd31.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt31 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt31.Load(reader31);

                    if (dt31.Rows.Count != 0)
                    {
                        worksheet.Cell(2, 1).InsertData(dt31);// inserta Contenido
                    }

                    //-------------------------------------------------------------------------
                    var worksheet16 = workbook.Worksheets.Add("RESUMENESTVSREAL");
                    SqlCommand cmd59 = null;

                    List<string[]> titlesFEDEX10 = new List<string[]> { new string[] { "Channel", "PromedioEstimatedShipping", "PromedioMontoTotal", "sumEstimatedShipping", "sumMontoTotal", "DifPorcentaje "} };

                    worksheet16.Cell(1, 1).InsertData(titlesFEDEX10);

                    SqlConnection Conexion59 = new SqlConnection();
                    Conexion59.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion59.Open();

                    // escoje el sp que le corresponde a mensual o semanal
                    // ---------------------------------------------------
                    if (comboBox2.Text != "SI")
                    {
                        cmd59 = new SqlCommand("EstimadovsRealResumenMensual", Conexion59);

                    }
                    else
                    {
                        cmd59 = new SqlCommand("EstimadovsRealResumensemanal", Conexion59);
                    }

                    cmd59.CommandType = CommandType.StoredProcedure;
                    cmd59.CommandTimeout = 7200; //in seconds
                    cmd59.Parameters.Add("@BeginDate", SqlDbType.DateTime).Value = Fecha1;
                    cmd59.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = Fecha2;

                    SqlDataReader reader59 = cmd59.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt59 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt59.Load(reader59);

                    if (dt59.Rows.Count != 0)
                    {
                        worksheet16.Cell(2, 1).InsertData(dt59);// inserta Contenido
                    }
                    //-------------------------------------------------------------------------
                    var worksheet17 = workbook.Worksheets.Add("CLASIFIESTVSREAL");
                    SqlCommand cmd60 = null;

                    List<string[]> titlesFEDEX11 = new List<string[]> { new string[] { "Clasificacion","Conteo","sumEstimatedShiping","sumMontoTotal","Amazon","Shopify","Walmart","sku "} };

                    worksheet17.Cell(1, 1).InsertData(titlesFEDEX11);

                    SqlConnection Conexion60 = new SqlConnection();
                    Conexion60.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion60.Open();

                    // escoje el sp que le corresponde a mensual o semanal
                    // ---------------------------------------------------
                    if (comboBox2.Text != "SI")
                    {
                        cmd60 = new SqlCommand("ClasiEstvsRealMensual", Conexion60);

                    }
                    else
                    {
                        cmd60 = new SqlCommand("ClasiEstvsRealsemanal", Conexion60);
                    }

                    cmd60.CommandType = CommandType.StoredProcedure;
                    cmd60.CommandTimeout = 7200; //in seconds
                    cmd60.Parameters.Add("@BeginDate", SqlDbType.DateTime).Value = Fecha1;
                    cmd60.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = Fecha2;

                    SqlDataReader reader60 = cmd60.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt60 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt60.Load(reader60);

                    if (dt60.Rows.Count != 0)
                    {
                        worksheet17.Cell(2, 1).InsertData(dt60);// inserta Contenido
                    }
                    //-------------------------------------------------------------------------
                    var worksheet5 = workbook.Worksheets.Add("Detalle");

                    List<string[]> titles99 = new List<string[]> { new string[] { "orderId", "RequestedServiceLevel", "ShippingServiceLevel", "sku", "Track", "Monto", "DimHeight", "DimLength", "DimWidth", "ActualWeightAmount", "RatedWeightAmount", "Girth", "ReglaFEDEX", "GirthEVP", "ReglaEVP " } };
                    worksheet5.Cell(1, 1).InsertData(titles99);

                    SqlConnection Conexion4 = new SqlConnection();
                    Conexion4.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion4.Open();
                    SqlCommand cmd4 = new SqlCommand("ClasiEstvsRealsemanalDetalle", Conexion4);
                    cmd4.CommandType = CommandType.StoredProcedure;
                    cmd4.CommandTimeout = 7200; //in seconds
                    SqlDataReader reader4 = cmd4.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt4 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt4.Load(reader4);

                    //Cantidad = dt4.Rows.Count;
                    if (dt4.Rows.Count != 0)
                    {
                        worksheet5.Cell(2, 1).InsertData(dt4);// inserta Contenido
                    }



                    string pathUPS = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                    pathUPS = pathUPS + @"\ShippingByMarket" + DateTime.Now.ToString("MMddyyyyHHmmss") + ".xlsx";
                    workbook.SaveAs(pathUPS);
                }



                if (checkBox1.Checked)
                {

                    SqlConnection Conexionlimite = new SqlConnection();
                    Conexionlimite.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    string sqllimite = "GeneraLimitesActualShippingCost";
                    SqlCommand commandlimite = new SqlCommand(sqllimite, Conexionlimite);
                    commandlimite.CommandType = CommandType.StoredProcedure;
                    commandlimite.CommandTimeout = 17200; //in seconds
                    Conexionlimite.Open();
                    commandlimite.ExecuteNonQuery();
                    Conexionlimite.Close();
                }

                // envia alerta de shipping
                // ------------------------
                string DireccionArchivo = "";
                int Cantidad = 0;
                using (var workbook = new XLWorkbook())
                {

                    var worksheet = workbook.Worksheets.Add("SKUPROMEDIO");

                    List<string[]> titles = new List<string[]> { new string[] { "SalesSku","DimHeightCeiling","DimLengthCeiling","DimWidthCeiling","ActualWeightAmountCeiling","ZoneCode","FulfillmentLocationName","Clasificacion","NetChargeAmount","promedioEstimatedShiping","sumTotalSales","Conteo","PorcentajeShipping"} };
                    worksheet.Cell(1, 1).InsertData(titles);

                    SqlConnection Conexion4 = new SqlConnection();
                    Conexion4.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion4.Open();
                    SqlCommand cmd4 = new SqlCommand("ReporteReglas", Conexion4);
                    cmd4.CommandType = CommandType.StoredProcedure;
                    cmd4.CommandTimeout = 7200; //in seconds
                    SqlDataReader reader4 = cmd4.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt4 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt4.Load(reader4);

                    Cantidad = dt4.Rows.Count;
                    if (dt4.Rows.Count != 0)
                    {
                        worksheet.Cell(2, 1).InsertData(dt4);// inserta Contenido
                    }

                    DireccionArchivo = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                    DireccionArchivo = DireccionArchivo + @"\ReporteReglas" + DateTime.Now.ToString("MMddyyyyHHmmss") + ".xlsx";
                    workbook.SaveAs(DireccionArchivo);
                }

                // envia alerta de shipping
                // ------------------------
                Cantidad = 0;
                using (var workbook = new XLWorkbook())
                {
                    
                    var worksheet = workbook.Worksheets.Add("Detalle");

                    List<string[]> titles66 = new List<string[]> { new string[] { "Track", "CCN_ORDER_NUMBER", "Monto", "TotalSales", "SalesOrderNumber", "RequestedServiceLevel", "ShippingServiceLevel", "ActualWeightAmount", "RatedWeightAmount", "sumOfQuantity" } };
                    worksheet.Cell(1, 1).InsertData(titles66);

                    SqlConnection Conexion4 = new SqlConnection();
                    Conexion4.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion4.Open();
                    SqlCommand cmd4 = new SqlCommand("RequestrealServices", Conexion4);
                    cmd4.CommandType = CommandType.StoredProcedure;
                    cmd4.CommandTimeout = 7200; //in seconds
                    SqlDataReader reader4 = cmd4.ExecuteReader();

                   
                    //Create a new DataTable.
                    System.Data.DataTable dt400 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt400.Load(reader4);

                    //Cantidad = dt4.Rows.Count;
                    if (dt400.Rows.Count != 0)
                    {
                        worksheet.Cell(2, 1).InsertData(dt400);// inserta Contenido
                    }


                    var worksheet1 = workbook.Worksheets.Add("Resumen");

                    List<string[]> titles1 = new List<string[]> { new string[] { "RequestedServiceLevel ","FedEx Ground®","FedEx SmartPost","FedEx Standard Overnight®","USPS – First Class","USPS – Standard Mail","FedEx 2Day®","FedEx Home Delivery®"  } };
                    worksheet1.Cell(1, 1).InsertData(titles1);

                    SqlConnection Conexion5 = new SqlConnection();
                    Conexion5.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion5.Open();
                    SqlCommand cmd5 = new SqlCommand("ResumenRequestrealServices", Conexion5);
                    cmd5.CommandType = CommandType.StoredProcedure;
                    cmd5.CommandTimeout = 7200; //in seconds
                    SqlDataReader reader5 = cmd5.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt5 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt5.Load(reader5);

                    //Cantidad = dt5.Rows.Count;
                    if (dt5.Rows.Count != 0)
                    {
                        worksheet1.Cell(2, 1).InsertData(dt5);// inserta Contenido
                    }


                    DireccionArchivo = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                    DireccionArchivo = DireccionArchivo + @"\RequestrealServices" + DateTime.Now.ToString("MMddyyyyHHmmss") + ".xlsx";
                    workbook.SaveAs(DireccionArchivo);
                }

                // envia alerta de shipping
                // ------------------------
                string DireccionArchivosku = "";
                int Cantidadsku = 0;
                using (var workbook = new XLWorkbook())
                {

                    var worksheet = workbook.Worksheets.Add("SKULOCATION");

                    List<string[]> titles = new List<string[]> { new string[] { "Sku", "SalesSku", "ColoradoActual", "DefaultActual", "Drop ShipActual", "Fredericksburg PAActual", "Gainesville GAActual", "Loxley ALActual", "Main WarehouseActual", "Prescott AZActual", "Princeton ILActual", "Sacramento CAActual", "ShedActual", "WAActual", "Wilton NYActual", "ColoradoRated", "DefaultRated", "Drop ShipRated", "Fredericksburg PARated", "Gainesville GARated", "Loxley ALRated", "Main WarehouseRated", "Prescott AZRated", "Princeton ILRated", "Sacramento CARated", "ShedRated", "WARated", "Wilton NYRated" } };
                    worksheet.Cell(1, 1).InsertData(titles);


                    SqlConnection Conexion4 = new SqlConnection();
                    Conexion4.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion4.Open();
                    SqlCommand cmd4 = new SqlCommand("SKULocation", Conexion4);
                    cmd4.CommandType = CommandType.StoredProcedure;
                    cmd4.CommandTimeout = 7200; //in seconds
                    SqlDataReader reader4 = cmd4.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt4 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt4.Load(reader4);

                    Cantidadsku = dt4.Rows.Count;
                    if (dt4.Rows.Count != 0)
                    {
                        worksheet.Cell(2, 1).InsertData(dt4);// inserta Contenido
                    }

                    DireccionArchivosku = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                    DireccionArchivosku = DireccionArchivosku + @"\ReporteSkuLocation" + DateTime.Now.ToString("MMddyyyyHHmmss") + ".xlsx";
                    workbook.SaveAs(DireccionArchivosku);
                }

                // envia alerta de shipping
                // ------------------------
                string DireccionArchivoskup = "";
                int Cantidadskup = 0;
                using (var workbook = new XLWorkbook())
                {

                    var worksheet = workbook.Worksheets.Add("ReporteB2B");

                    List<string[]> titles = new List<string[]> { new string[] { "orderId", "totalSale", "sumOfQuantity", "EstimatedShiping", "date", "Channel", "sku", "SalesOrderNumber", "RealShipping", "ShipingSobreVentas", "EstvsReal", "TotalTracks", "RequestedServiceLevel", "ShippingServiceLevel", "FulfillmentLocationName", "ShipmentDate", "DeliveryDate", "DiasEntrega", "FulfillmentServiceLevel", "ZoneCode" } };
                    worksheet.Cell(1, 1).InsertData(titles);

                    SqlConnection Conexion4 = new SqlConnection();
                    Conexion4.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion4.Open();
                    SqlCommand cmd4 = new SqlCommand("ReporteB2B", Conexion4);
                    cmd4.CommandType = CommandType.StoredProcedure;
                    cmd4.CommandTimeout = 7200; //in seconds
                    SqlDataReader reader4 = cmd4.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt4 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt4.Load(reader4);

                    Cantidadskup = dt4.Rows.Count;
                    if (dt4.Rows.Count != 0)
                    {
                        worksheet.Cell(2, 1).InsertData(dt4);// inserta Contenido
                    }

                    DireccionArchivoskup = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                    DireccionArchivoskup = DireccionArchivoskup + @"\ReporteB2B" + DateTime.Now.ToString("MMddyyyyHHmmss") + ".xlsx";
                    workbook.SaveAs(DireccionArchivoskup);
                }

                // Genera template amazon
                // ----------------------
                using (var workbook = new XLWorkbook())
                {

                    var worksheet = workbook.Worksheets.Add("TemplateAmazon");

                    List<string[]> titles = new List<string[]> { new string[] { "Template", "Loxley AL", "Sacramento CA", "WA", "Fredericksburg PA", "Prescott AZ", "Wilton NY", "Colorado", "Princeton IL", "Gainesville GA", "StateRegion", "MINIMO", "Shouldbe", "AmazonActual", "Remaininsametemplate" } };
                    worksheet.Cell(1, 1).InsertData(titles);

                    SqlConnection Conexion4 = new SqlConnection();
                    Conexion4.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion4.Open();
                    SqlCommand cmd4 = new SqlCommand("TemplateAmazon", Conexion4);
                    cmd4.CommandType = CommandType.StoredProcedure;
                    cmd4.CommandTimeout = 7200; //in seconds
                    SqlDataReader reader4 = cmd4.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt4 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt4.Load(reader4);

                    Cantidadskup = dt4.Rows.Count;
                    if (dt4.Rows.Count != 0)
                    {
                        worksheet.Cell(2, 1).InsertData(dt4);// inserta Contenido
                    }

                    DireccionArchivoskup = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                    DireccionArchivoskup = DireccionArchivoskup + @"\TemplateAmazon" + DateTime.Now.ToString("MMddyyyyHHmmss") + ".xlsx";
                    workbook.SaveAs(DireccionArchivoskup);
                }

                // Fecha Entrega
                // -------------
                using (var workbook = new XLWorkbook())
                {

                    var worksheet = workbook.Worksheets.Add("Resumen");

                    List<string[]> titles = new List<string[]> { new string[] { "ShippingServiceLevelTotal","Porcentaje","CantidadATiempo","CantidadTotal","DiasVentaEntrega","DiasVentaEntregaEstimada" } };
                    worksheet.Cell(1, 1).InsertData(titles);

                    List<string[]> titles10 = new List<string[]> { new string[] { "Channel", "Porcentaje", "CantidadATiempo", "CantidadTotal", "DiasVentaEntrega", "DiasVentaEntregaEstimada", "DiasVentaEnt", "DiasVentaEntregaEst" } };
                    worksheet.Cell(11, 1).InsertData(titles10);

                    SqlConnection Conexion4 = new SqlConnection();
                    Conexion4.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion4.Open();
                    SqlCommand cmd4 = new SqlCommand("ReporteFechaEntregaServicesType", Conexion4);
                    cmd4.CommandType = CommandType.StoredProcedure;
                    cmd4.CommandTimeout = 7200; //in seconds
                    SqlDataReader reader4 = cmd4.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt4 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt4.Load(reader4);

                    Cantidadskup = dt4.Rows.Count;
                    if (dt4.Rows.Count != 0)
                    {
                        worksheet.Cell(2, 1).InsertData(dt4);// inserta Contenido
                    }

                    SqlConnection Conexion5 = new SqlConnection();
                    Conexion5.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion5.Open();
                    SqlCommand cmd5 = new SqlCommand("ReporteFechaEntregaChannel", Conexion5);
                    cmd5.CommandType = CommandType.StoredProcedure;
                    cmd5.CommandTimeout = 7200; //in seconds
                    SqlDataReader reader5 = cmd5.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt5 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt5.Load(reader5);

                    Cantidadskup = dt5.Rows.Count;
                    if (dt5.Rows.Count != 0)
                    {
                        worksheet.Cell(12, 1).InsertData(dt5);// inserta Contenido
                    }


                    var worksheet1 = workbook.Worksheets.Add("Detalle");

                    List<string[]> titles1 = new List<string[]> { new string[] { "SalesOrderNumber","Channel","ShipmentDate","DeliveryDate","FECHAENTREGAESTIMADA","Analisis","SalesOrderDate","DiasMarketPlace","FulfillmentLocationName","StateRegion","ZoneCode","ServicesType" } };
                    worksheet1.Cell(1, 1).InsertData(titles);

                    SqlConnection Conexion6 = new SqlConnection();
                    Conexion6.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion6.Open();
                    SqlCommand cmd6 = new SqlCommand("ReporteDetalleFechaEntregaServices", Conexion6);
                    cmd6.CommandType = CommandType.StoredProcedure;
                    cmd6.CommandTimeout = 7200; //in seconds
                    SqlDataReader reader6 = cmd6.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt6 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt6.Load(reader6);

                    Cantidadskup = dt6.Rows.Count;
                    if (dt6.Rows.Count != 0)
                    {
                        worksheet1.Cell(2, 1).InsertData(dt6);// inserta Contenido
                    }

                    DireccionArchivoskup = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                    DireccionArchivoskup = DireccionArchivoskup + @"\FechaEntrega" + DateTime.Now.ToString("MMddyyyyHHmmss") + ".xlsx";
                    workbook.SaveAs(DireccionArchivoskup);
                }

                // envia alerta de shipping
                // ------------------------
                string DireccionArchivosoutput = "";
                int Cantidadregistros = 0;
                using (var workbook = new XLWorkbook())
                {

                    var worksheet = workbook.Worksheets.Add("SKUCOSTO");

                    List<string[]> titles = new List<string[]> { new string[] { "Sku", "TotalSales" } };
                    worksheet.Cell(1, 1).InsertData(titles);

                    SqlConnection Conexion4 = new SqlConnection();
                    Conexion4.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion4.Open();
                    SqlCommand cmd4 = new SqlCommand("ReporteSKUPromedioVenta", Conexion4);
                    cmd4.CommandType = CommandType.StoredProcedure;
                    cmd4.CommandTimeout = 7200; //in seconds
                    SqlDataReader reader4 = cmd4.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt4 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt4.Load(reader4);

                    Cantidadskup = dt4.Rows.Count;
                    if (dt4.Rows.Count != 0)
                    {
                        worksheet.Cell(2, 1).InsertData(dt4);// inserta Contenido
                    }

                    DireccionArchivoskup = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                    DireccionArchivoskup = DireccionArchivoskup + @"\ReporteSkuCosto" + DateTime.Now.ToString("MMddyyyyHHmmss") + ".xlsx";
                    workbook.SaveAs(DireccionArchivoskup);
                }

                using (var workbook = new XLWorkbook())
                {

                    var worksheet = workbook.Worksheets.Add("SKUNUEVOS");

                    List<string[]> titles = new List<string[]> { new string[] { "Sku", "DATE", "CantidadVentas", "TotalSales", "Monto" } };
                    worksheet.Cell(1, 1).InsertData(titles);


                    SqlConnection Conexion4 = new SqlConnection();
                    Conexion4.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                    Conexion4.Open();
                    SqlCommand cmd4 = new SqlCommand("ReporteSkuNuevos", Conexion4);
                    cmd4.CommandType = CommandType.StoredProcedure;
                    cmd4.CommandTimeout = 7200; //in seconds
                    SqlDataReader reader4 = cmd4.ExecuteReader();

                    //Create a new DataTable.
                    System.Data.DataTable dt4 = new System.Data.DataTable("Resultado");

                    //Load DataReader into the DataTable.
                    dt4.Load(reader4);

                    Cantidadsku = dt4.Rows.Count;
                    if (dt4.Rows.Count != 0)
                    {
                        worksheet.Cell(2, 1).InsertData(dt4);// inserta Contenido
                    }

                    DireccionArchivosku = ConfigurationManager.AppSettings["RutaArchivosOutputs"];
                    DireccionArchivosku = DireccionArchivosku + @"\ReporteSkuNuevos" + DateTime.Now.ToString("MMddyyyyHHmmss") + ".xlsx";
                    workbook.SaveAs(DireccionArchivosku);
                }
                // llena tabla de bianalitics
                // --------------------------
                //textBox1.Text = "Llena BI analitics";
                this.Refresh();
                this.Invalidate();
                SqlConnection Conexionbianalitics = new SqlConnection();
                Conexionbianalitics.ConnectionString = ConfigurationManager.AppSettings["ConectionString"];
                string sqlbianalitics = "LlenaBIANALITICS";
                SqlCommand commandbianalitics = new SqlCommand(sqlbianalitics, Conexionbianalitics);
                commandbianalitics.CommandType = CommandType.StoredProcedure;
                commandbianalitics.CommandTimeout = 8200; //in seconds
                Conexionbianalitics.Open();
                commandbianalitics.ExecuteNonQuery();
                Conexionbianalitics.Close();

 

            }
            catch (SystemException exp)
            {
                MessageBox.Show("Error: " + exp.Message);


            }
        }

        private void label6_Click(object sender, EventArgs e)
        {
        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
