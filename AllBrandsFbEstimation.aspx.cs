using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace DemoCrmH
{

   
    public partial class AllBrandsFbEstimation : System.Web.UI.Page
    {
        string endMonthdate;
        string endYesterdate;
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["crm_LC_MarketingConnectionString"].ConnectionString);

        OleDbConnection Econ;

        private void ExcelConn(string filepath)
        {

            string constr = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;""", filepath);

            Econ = new OleDbConnection(constr);
            Econ.Open();
        }

        private void InsertExceldata( string filename, string temp_table, string file_name)
        {

            try
            {

                string fullpath = Server.MapPath("/App_Data/") + filename;

                ExcelConn(fullpath);

                DataTable Sheets = new DataTable();

                Sheets = Econ.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                string CNRawData = Sheets.Rows[0]["TABLE_NAME"].ToString();

                string query = string.Format("Select * from [{0}]", CNRawData);

                OleDbCommand Ecom = new OleDbCommand(query, Econ);





                DataSet ds = new DataSet();

                OleDbDataAdapter oda = new OleDbDataAdapter(query, Econ);



                oda.Fill(ds);



                DataTable dt = ds.Tables[0];

                foreach (DataRow row in dt.Rows)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        string cell = row[i].ToString();
                        if (cell == "#DIV/0!")
                        {
                            cell = null;
                            row[i] = cell;
                        }


                    }

                }

                SqlBulkCopy objbulk = new SqlBulkCopy(con);

                objbulk.DestinationTableName = temp_table;


                objbulk.ColumnMappings.Add(0, 0);
                objbulk.ColumnMappings.Add(1, 1);

                objbulk.ColumnMappings.Add(2, 2);

                objbulk.ColumnMappings.Add(3, 3);

                objbulk.ColumnMappings.Add(4, 4);

                objbulk.ColumnMappings.Add(5, 5);
                objbulk.ColumnMappings.Add(6, 6);
                objbulk.ColumnMappings.Add(7, 7);
                objbulk.ColumnMappings.Add(8, 8);
                objbulk.ColumnMappings.Add(9, 9);
                objbulk.ColumnMappings.Add(10, 10);
                objbulk.ColumnMappings.Add(11, 11);
                objbulk.ColumnMappings.Add(12, 12);

                //  con.Open();

                objbulk.WriteToServer(dt);

                //    con.Close();
            }
            catch
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("There is an issue with file: " + file_name);
                sb.AppendLine("Please follow these steps and check:");
                sb.AppendLine("- All column titles are correct.");
                sb.AppendLine("- No empty lines.");
                sb.AppendLine("- Try to copy paste all content as values to a new Excel file.");
                sb.AppendLine("Please check that the columns are as the following also by small and capital letters:");
                sb.AppendLine("Ad Set Name / Delivery / Campaign ID / Ad Set ID / Leads / FTD / L2FTD / Max L2FTD / Bid / Cost all time / CPA all time / CPL all time / Profit per lead ");
                sb.AppendLine("- Contact the developer.");

                string message = sb.ToString().Replace(Environment.NewLine, "<br />");
                Response.ClearHeaders();
                Response.Write( message);
                Response.End();
            }

        }

        private void InsertExceldataCostCN(string filename, string cost_table, string file_name)
        {
            try
            {
                string fullpath = Server.MapPath("/App_Data/") + filename;

                ExcelConn(fullpath);

                DataTable Sheets = new DataTable();

                Sheets = Econ.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                string fortradesoct = Sheets.Rows[0]["TABLE_NAME"].ToString();
                

                string query = string.Format("Select * from [{0}]", fortradesoct);

                OleDbCommand Ecom = new OleDbCommand(query, Econ);

                DataSet ds = new DataSet();

                OleDbDataAdapter oda = new OleDbDataAdapter(query, Econ);



                oda.Fill(ds);



                DataTable dt = ds.Tables[0];

                foreach (DataRow row in dt.Rows)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        string cell = row[i].ToString();
                        if (cell == "#DIV/0!")
                        {
                            cell = null;
                            row[i] = cell;
                        }


                    }

                }


                string s = dt.Columns[7].DataType.Name.ToString();
                if (s != "String")
                {
                    dt.ConvertColumnType("Bid", typeof(string));
                }
                SqlBulkCopy objbulk = new SqlBulkCopy(con);

                objbulk.DestinationTableName = cost_table;


                objbulk.ColumnMappings.Add(0, 0);
                objbulk.ColumnMappings.Add(1, 1);

                objbulk.ColumnMappings.Add(2, 2);

                objbulk.ColumnMappings.Add(3, 3);

                objbulk.ColumnMappings.Add(4, 4);

                objbulk.ColumnMappings.Add(5, 5);
                objbulk.ColumnMappings.Add(6, 6);
                objbulk.ColumnMappings.Add(7, 7);
                objbulk.ColumnMappings.Add(8, 8);
                objbulk.ColumnMappings.Add(9, 9);
                objbulk.ColumnMappings.Add(10, 10);
                objbulk.ColumnMappings.Add(11, 11);

                objbulk.WriteToServer(dt);
            }
            catch
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("There is an issue with file:  '" + file_name + "'");
                sb.AppendLine("Please follow these steps and check:");
                sb.AppendLine("- All column titles are correct.");
                sb.AppendLine("- No empty lines.");
                sb.AppendLine("- Try to copy paste all content as values to a new Excel file.");
                sb.AppendLine("- Contact the developer.");

                string message = sb.ToString().Replace(Environment.NewLine, "<br />"); ;
                Response.ClearHeaders();
                Response.Write(message);
                Response.End();

            }

            //   con.Close();
        }
        private void InsertExceldataCost( string filename, string cost_table, string file_name)
        {
            try
            {
                // string fileName = Path.GetFileName(FileUpload2.PostedFile.FileName);
                string fullpath = Server.MapPath("/App_Data/") + filename;

                ExcelConn(fullpath);

                DataTable Sheets = new DataTable();

                Sheets = Econ.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                string fortradesoct = Sheets.Rows[0]["TABLE_NAME"].ToString();
                // fortradesoct = "1$";

                string query = string.Format("Select * from [{0}]", fortradesoct);

                OleDbCommand Ecom = new OleDbCommand(query, Econ);





                DataSet ds = new DataSet();

                OleDbDataAdapter oda = new OleDbDataAdapter(query, Econ);

                Econ.Close();

                oda.Fill(ds);



                DataTable dt = ds.Tables[0];

                string s = dt.Columns[15].DataType.Name.ToString();
                if (s != "String")
                {
                    dt.ConvertColumnType("Bid", typeof(string));
                }

                SqlBulkCopy objbulk = new SqlBulkCopy(con);

                objbulk.DestinationTableName = cost_table;


                objbulk.ColumnMappings.Add(0, 0);
                objbulk.ColumnMappings.Add(1, 1);

                objbulk.ColumnMappings.Add(2, 2);

                objbulk.ColumnMappings.Add(3, 3);

                objbulk.ColumnMappings.Add(4, 4);

                objbulk.ColumnMappings.Add(5, 5);
                objbulk.ColumnMappings.Add(6, 6);
                objbulk.ColumnMappings.Add(7, 7);
                objbulk.ColumnMappings.Add(8, 8);
                objbulk.ColumnMappings.Add(9, 9);
                objbulk.ColumnMappings.Add(10, 10);
                objbulk.ColumnMappings.Add(11, 11);
                objbulk.ColumnMappings.Add(12, 12);
                objbulk.ColumnMappings.Add(13, 13);
                objbulk.ColumnMappings.Add(14, 14);
                objbulk.ColumnMappings.Add(15, 15);
                objbulk.ColumnMappings.Add(16, 16);
                objbulk.ColumnMappings.Add(17, 17);
                objbulk.ColumnMappings.Add(18, 18);

               

                objbulk.WriteToServer(dt);
            }
            catch
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("There is an issue with file:  '" + file_name+ "'");
                sb.AppendLine("Please follow these steps and check:");
                sb.AppendLine("- All column titles are correct.");
                sb.AppendLine("- No empty lines.");
                sb.AppendLine("- Try to copy paste all content as values to a new Excel file.");
                sb.AppendLine("Please check that the columns are as the following also by small and capital letters:");
                sb.AppendLine("Ad set name / Day / Ad set ID / Delivery status / Delivery level / Campaign name / Result Type / Results / Reach / Impressions / Cost per result / Quality ranking / Engagement rate ranking / Conversion rate ranking / Amount spent (USD) / Bid / Bid Type / Reporting starts / Reporting ends");
                sb.AppendLine("- Contact the developer.");

                string message = sb.ToString().Replace(Environment.NewLine, "<br />"); ;
                Response.ClearHeaders();
                Response.Write(message);
                Response.End();

            }

         //   con.Close();

        }

        string[] sql_temp_tbl = {"#kapital","#fx","#viop","#my","#cn","#cnww","#fort_countries","#fort_countries_new","#fort_en_eu","#fort_en_ww", "#fort_en_ww1","#fort_en_cd","#fort_en_aus","#fort_cr","#fort_cr_eu","#fort_cr_ww","#fort_sl","#fort_sl_ww","#fort_mac","#fort_alb"};


        string [] temp_tables = {"[Marketing].[dbo].[FbEstKapitalRS$]","[Marketing].[dbo].[FbEstGcmFXTemp$]","[Marketing].[dbo].[FbEstGcmViopTemp$]","[Marketing].[dbo].[FbEstMYTemp$]"
                                    ,"[Marketing].[dbo].[FbEstCNTemp$]","[Marketing].[dbo].[FbEstCNWWTemp$]","[Marketing].[dbo].[FbEstFortradeCountriesTemp$]","[Marketing].[dbo].[FbEstFortradeCountriesNewTemp$]"
                                    ,"[Marketing].[dbo].[FbEstFortradeEnEuTemp$]","[Marketing].[dbo].[FbEstFortradeEnWWTemp$]","[Marketing].[dbo].[FbEstFortradeEnWW1Temp$]","[Marketing].[dbo].[FbEstFortradeCanadaTemp$]"
                                    ,"[Marketing].[dbo].[FbEstFortradeAustTemp$]","[Marketing].[dbo].[FbEstFortradeCRTemp$]","[Marketing].[dbo].[FbEstFortradeCREUTemp$]"
                                    ,"[Marketing].[dbo].[FbEstFortradeCRWWTemp$]","[Marketing].[dbo].[FbEstFortradeSLTemp$]","[Marketing].[dbo].[FbEstFortradeSLWWTemp$]"
                                    ,"[Marketing].[dbo].[FbEstFortradeMacedoniaTemp$]","[Marketing].[dbo].[FbEstFortradeAlbaniaTemp$]"};


        string[] cost_tables = {"[Marketing].[dbo].[FbEstKapitalCost$]", "[Marketing].[dbo].[FbEstGcmFxCost$]", "[Marketing].[dbo].[FbEstGcmViopCost$]", "[Marketing].[dbo].[FbEstMYCost$]"
                                  ,"[Marketing].[dbo].[FbEstCNCost$]","[Marketing].[dbo].[FbEstCNWWCost$]","[Marketing].[dbo].[FbEstFortradeCountriesCost$]","[Marketing].[dbo].[FbEstFortradeCountriesNewCost$]"
                                    ,"[Marketing].[dbo].[FbEstFortradeEnEuCost$]","[Marketing].[dbo].[FbEstFortradeEnWWCost$]","[Marketing].[dbo].[FbEstFortradeEnWW1Cost$]","[Marketing].[dbo].[FbEstFortradeCanadaCost$]"
                                    ,"[Marketing].[dbo].[FbEstFortradeAustCost$]","[Marketing].[dbo].[FbEstFortradeCRCost$]","[Marketing].[dbo].[FbEstFortradeCREUCost$]"
                                    ,"[Marketing].[dbo].[FbEstFortradeCRWWCost$]","[Marketing].[dbo].[FbEstFortradeSLCost$]","[Marketing].[dbo].[FbEstFortradeSLWWCost$]"
                                    ,"[Marketing].[dbo].[FbEstFortradeMacedoniaCost$]","[Marketing].[dbo].[FbEstFortradeAlbaniaCost$]"};





        protected void Page_Load(object sender, EventArgs e)
        {
            string maincon = ConfigurationManager.ConnectionStrings["crm_LC_MarketingConnectionString"].ConnectionString;
            SqlConnection sqlcon = new SqlConnection(maincon);

            string sqlquery = "";
           
            sqlquery += droptemptables(temp_tables, cost_tables);
            SqlCommand sqlcom = new SqlCommand(sqlquery, sqlcon);
            sqlcon.Open();
            sqlcom.ExecuteNonQuery();
        }

        public string droptemptables(string[] temp_tables, string[] cost_tables)
        {
            string sqlstr = "";

            for (int i = 0; i < temp_tables.Length; i++)
            {
                sqlstr += "truncate table " + temp_tables[i] + "; ";
            }

            for (int i = 0; i < cost_tables.Length; i++)
            {
                sqlstr += "truncate table " + cost_tables[i] + "; ";
            }
          

            return sqlstr;
        }

        protected void btnUploadXL_Click(object sender, EventArgs e)
        {
            endMonthdate = txtEndDate.Text;
            endYesterdate = txtYestEndDate.Text;
            DateTime endMonthDate = DateTime.Parse(endMonthdate);
            DateTime endYesterDate = DateTime.Parse(endYesterdate);
            string endMonth = endMonthDate.AddDays(1).ToString("yyyy-MM-dd");
            string endYesterday = endYesterDate.AddDays(1).ToString("yyyy-MM-dd");



            string filename_Kapital = Guid.NewGuid() + Path.GetExtension(kapital_temp.PostedFile.FileName);
            string filename_gcmforex_fx = Guid.NewGuid() + Path.GetExtension(gcmfrx_fx_temp.PostedFile.FileName);
            string filename_gcmforex_viop = Guid.NewGuid() + Path.GetExtension(gcmfrx_viop_temp.PostedFile.FileName);
            string filename_gcmAsia_my = Guid.NewGuid() + Path.GetExtension(gcmasia_my_temp.PostedFile.FileName);
            string filename_gcmAsia_cncn = Guid.NewGuid() + Path.GetExtension(gcmasia_cn_temp.PostedFile.FileName);
            string filename_gcmAsia_cnww = Guid.NewGuid() + Path.GetExtension(gcmasia_ww_temp.PostedFile.FileName);
            string filename_fortrade_countries = Guid.NewGuid() + Path.GetExtension(fort_en_count_temp.PostedFile.FileName);
            string filename_fortrade_countries_new = Guid.NewGuid() + Path.GetExtension(fort_en_count_new_temp.PostedFile.FileName);
            string filename_fortrade_en_eu = Guid.NewGuid() + Path.GetExtension(fort_en_eu_temp.PostedFile.FileName);
            string filename_fortrade_en_ww = Guid.NewGuid() + Path.GetExtension(fort_en_ww_temp.PostedFile.FileName);
            string filename_fortrade_en_ww1 = Guid.NewGuid() + Path.GetExtension(fort_en_ww_temp1.PostedFile.FileName);
            string filename_fortrade_canada = Guid.NewGuid() + Path.GetExtension(fort_canada_temp.PostedFile.FileName);
            string filename_fortrade_australia = Guid.NewGuid() + Path.GetExtension(fort_aust_temp.PostedFile.FileName);
            string filename_fortrade_cr = Guid.NewGuid() + Path.GetExtension(fort_cr_temp.PostedFile.FileName);
            string filename_fortrade_cr_eu = Guid.NewGuid() + Path.GetExtension(fort_cr_eu_temp.PostedFile.FileName);
            string filename_fortrade_cr_ww = Guid.NewGuid() + Path.GetExtension(fort_cr_ww_temp.PostedFile.FileName);
            string filename_fortrade_sl = Guid.NewGuid() + Path.GetExtension(fort_sl_temp.PostedFile.FileName);
            string filename_fortrade_sl_ww = Guid.NewGuid() + Path.GetExtension(fort_sl_ww_temp.PostedFile.FileName);
            string filename_fortrade_mcd = Guid.NewGuid() + Path.GetExtension(fort_mcd_temp.PostedFile.FileName);
            string filename_fortrade_al = Guid.NewGuid() + Path.GetExtension(fort_al_temp.PostedFile.FileName);


            string[] temp_arr = { filename_Kapital, filename_gcmforex_fx, filename_gcmforex_viop, filename_gcmAsia_my,
            filename_gcmAsia_cncn,filename_gcmAsia_cnww,filename_fortrade_countries,filename_fortrade_countries_new,filename_fortrade_en_eu,
            filename_fortrade_en_ww, filename_fortrade_en_ww1,filename_fortrade_canada,filename_fortrade_australia,filename_fortrade_cr,filename_fortrade_cr_eu,
            filename_fortrade_cr_ww,filename_fortrade_sl,filename_fortrade_sl_ww,filename_fortrade_mcd,filename_fortrade_al};

            System.Web.UI.WebControls.FileUpload [] temp_arr_upload = { kapital_temp, gcmfrx_fx_temp, gcmfrx_viop_temp, gcmasia_my_temp, gcmasia_cn_temp,
            gcmasia_ww_temp,fort_en_count_temp,fort_en_count_new_temp,fort_en_eu_temp,fort_en_ww_temp,fort_en_ww_temp1,
            fort_canada_temp,fort_aust_temp,fort_cr_temp,fort_cr_eu_temp,fort_cr_ww_temp,
            fort_sl_temp,fort_sl_ww_temp,fort_mcd_temp,fort_al_temp};
            //temp_arr_upload[0].FileName

            string filename_Kapital_cost = Guid.NewGuid() + Path.GetExtension(kapital_cost.PostedFile.FileName);
            string filename_gcmforex_fx_cost = Guid.NewGuid() + Path.GetExtension(gcmfrx_fx_cost.PostedFile.FileName);
            string filename_gcmforex_viop_cost = Guid.NewGuid() + Path.GetExtension(gcmfrx_viop_cost.PostedFile.FileName);
            string filename_gcmAsia_my_cost = Guid.NewGuid() + Path.GetExtension(gcmasia_my_cost.PostedFile.FileName);
            string filename_gcmAsia_cncn_cost = Guid.NewGuid() + Path.GetExtension(gcmasia_cn_cost.PostedFile.FileName);
            string filename_gcmAsia_cnww_cost = Guid.NewGuid() + Path.GetExtension(gcmasia_ww_cost.PostedFile.FileName);
            string filename_fortrade_countries_cost = Guid.NewGuid() + Path.GetExtension(fort_en_count_cost.PostedFile.FileName);
            string filename_fortrade_countries_new_cost = Guid.NewGuid() + Path.GetExtension(fort_en_count_new_cost.PostedFile.FileName);
            string filename_fortrade_en_eu_cost = Guid.NewGuid() + Path.GetExtension(fort_en_eu_cost.PostedFile.FileName);
            string filename_fortrade_en_ww_cost = Guid.NewGuid() + Path.GetExtension(fort_en_ww_cost.PostedFile.FileName);
            string filename_fortrade_en_ww_cost1 = Guid.NewGuid() + Path.GetExtension(fort_en_ww_cost1.PostedFile.FileName);
            string filename_fortrade_canada_cost = Guid.NewGuid() + Path.GetExtension(fort_canada_cost.PostedFile.FileName);
            string filename_fortrade_australia_cost = Guid.NewGuid() + Path.GetExtension(fort_aust_cost.PostedFile.FileName);
            string filename_fortrade_cr_cost = Guid.NewGuid() + Path.GetExtension(fort_cr_cost.PostedFile.FileName);
            string filename_fortrade_cr_eu_cost = Guid.NewGuid() + Path.GetExtension(fort_cr_eu_cost.PostedFile.FileName);
            string filename_fortrade_cr_ww_cost = Guid.NewGuid() + Path.GetExtension(fort_cr_ww_cost.PostedFile.FileName);
            string filename_fortrade_sl_cost = Guid.NewGuid() + Path.GetExtension(fort_sl_cost.PostedFile.FileName);
            string filename_fortrade_sl_ww_cost = Guid.NewGuid() + Path.GetExtension(fort_sl_ww_cost.PostedFile.FileName);
            string filename_fortrade_mcd_cost = Guid.NewGuid() + Path.GetExtension(fort_mcd_cost.PostedFile.FileName);
            string filename_fortrade_al_cost = Guid.NewGuid() + Path.GetExtension(fort_al_cost.PostedFile.FileName);


            string[] cost_arr = { filename_Kapital_cost, filename_gcmforex_fx_cost, filename_gcmforex_viop_cost, filename_gcmAsia_my_cost,
            filename_gcmAsia_cncn_cost,filename_gcmAsia_cnww_cost,filename_fortrade_countries_cost,filename_fortrade_countries_new_cost,filename_fortrade_en_eu_cost,
            filename_fortrade_en_ww_cost,filename_fortrade_en_ww_cost1,filename_fortrade_canada_cost,filename_fortrade_australia_cost,filename_fortrade_cr_cost,filename_fortrade_cr_eu_cost,
            filename_fortrade_cr_ww_cost,filename_fortrade_sl_cost,filename_fortrade_sl_ww_cost,filename_fortrade_mcd_cost,filename_fortrade_al_cost};

           


            System.Web.UI.WebControls.FileUpload[] cost_arr_upload = { kapital_cost, gcmfrx_fx_cost, gcmfrx_viop_cost, gcmasia_my_cost, gcmasia_cn_cost,
            gcmasia_ww_cost,fort_en_count_cost,fort_en_count_new_cost,fort_en_eu_cost,fort_en_ww_cost,fort_en_ww_cost1,
            fort_canada_cost,fort_aust_cost,fort_cr_cost,fort_cr_eu_cost,fort_cr_ww_cost,
            fort_sl_cost,fort_sl_ww_cost,fort_mcd_cost,fort_al_cost};
            //cost_arr_upload[0].FileName

            System.Web.UI.WebControls.GridView[] gridview = {gvKapital, gvGcmFx,gvGcmViop, gvGcmAsiaMy, gvGcmAsiaCnCn, gvGcmAsiaCnWw, gvFortEnCountries
            ,gvFortEnCountriesNew,gvFortEnEu,gvFortEnWw,gvFortEnWw1,gvCanada,gvAustraila,gvFortrCr,gvFortrCrEu,gvFortrCrWw,gvFortrSL,gvFortrSLWw,gvFortMacedonia,gvFortAlbania};


            string[] sp = { "[Marketing].[dbo].[sp_KapitalFbEst]","[Marketing].[dbo].[sp_gcmforex_fx]","[Marketing].[dbo].[sp_gcmforex_viop]","[Marketing].[dbo].[sp_gcm_my]",
"[Marketing].[dbo].[sp_gcm_cn]","[Marketing].[dbo].[sp_gcm_cnww]","[Marketing].[dbo].[sp_en_countries]","[Marketing].[dbo].[sp_en_countries_new]","[Marketing].[dbo].[sp_en_eu]", "[Marketing].[dbo].[sp_en_ww]","[Marketing].[dbo].[sp_en_ww1]","[Marketing].[dbo].[sp_en_cd]","[Marketing].[dbo].[sp_en_aus]",
         "[Marketing].[dbo].[sp_cr]", "[Marketing].[dbo].[sp_cr_eu]","[Marketing].[dbo].[sp_cr_ww]","[Marketing].[dbo].[sp_sl]","[Marketing].[dbo].[sp_sl_ww]","[Marketing].[dbo].[sp_mac]","[Marketing].[dbo].[sp_alb]"};

            con.Open();


            SqlCommand sqlcom = new SqlCommand();
         
           
            SqlDataAdapter sda = new SqlDataAdapter();
            DataTable dt = new DataTable();
            for (int i = 0; i < temp_arr.Length; i++)
            {
                if (temp_arr_upload[i].HasFile==false)
                {
                    gridview[i]. DataSource = new string[] { };
                    gridview[i].DataBind();
                    continue;
                }
                else
                {


                    temp_arr_upload[i].PostedFile.SaveAs(Path.Combine(Server.MapPath("/App_Data"), temp_arr[i]));
                    cost_arr_upload[i].PostedFile.SaveAs(Path.Combine(Server.MapPath("/App_Data"), cost_arr[i]));

                    bool b = (i == 3 || i == 4 || i == 5);

                    InsertExceldata(temp_arr[i], temp_tables[i], temp_arr_upload[i].FileName);
                    if (b)
                    {
                        InsertExceldataCostCN(cost_arr[i], cost_tables[i], temp_arr_upload[i].FileName);
                    }
                    else
                    {
                        InsertExceldataCost(cost_arr[i], cost_tables[i], cost_arr_upload[i].FileName);
                    }
                    sqlcom = new SqlCommand(sp[i], con);
                    sqlcom.CommandType = CommandType.StoredProcedure;
                    sqlcom.Parameters.Add(new SqlParameter("@StartMonth", txtStartDate.Text));
                    sqlcom.Parameters.Add(new SqlParameter("@EndMonth", endMonth));
                    sqlcom.Parameters.Add(new SqlParameter("@StartYes", txtYestStartDate.Text));
                    sqlcom.Parameters.Add(new SqlParameter("@EndYest", endYesterday));
                 
                 
                    sqlcom.CommandTimeout = 950;
                    sqlcom.ExecuteNonQuery();
                    sda = new SqlDataAdapter(sqlcom);
                    dt = new DataTable();
                    sda.Fill(dt);
                    //clone datatable     
                    DataTable dtCloned = dt.Clone();
                    //change data type of column
                    dtCloned.Columns[4].DataType = typeof(Int32);
                    //import row to cloned datatable
                    foreach (DataRow row in dt.Rows)
                    {
                        dtCloned.ImportRow(row);
                    }



                    SqlDataReader dr = sqlcom.ExecuteReader();
                    if (dr.Read())
                    {
                        gridview[i].DataSource = dtCloned;
                        gridview[i].DataBind();
                    }
                    else
                    {
                        // Label1.Text = "No Results";
                        gridview[i].DataSource = null;
                        gridview[i].DataBind();
                    }
                    dr.Close();
                    dt.Clear();
                }
            }
            con.Close();
            sqlcom.Dispose();
            sda.Dispose();


        }

        protected void btnExport1_Click(object sender, EventArgs e)
        {


            //DateTime time = DateTime.Now;
            //Response.Clear();
            //Response.AddHeader("content-disposition", "attachment; filename = " + time + ".xls");
            //Response.ContentType = "application/ms-excel";
            //Response.ContentEncoding = System.Text.Encoding.Unicode;
            //Response.BinaryWrite(System.Text.Encoding.Unicode.GetPreamble());

            //System.IO.StringWriter sw = new System.IO.StringWriter();
            //System.Web.UI.HtmlTextWriter hw = new HtmlTextWriter(sw);

            //GridView3.RenderControl(hw);

            //Response.Write(sw.ToString());
            //Response.End();

        }
    }
}
