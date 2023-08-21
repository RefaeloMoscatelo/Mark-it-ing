using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IdentityModel.Services;
namespace DemoCrmH
{

    public partial class Ads : System.Web.UI.Page
    {
        protected void Application_Error(object sender, EventArgs e)
        {
            var error = Server.GetLastError();
            var cryptoEx = error as CryptographicException;
            if (cryptoEx != null)
            {
                FederatedAuthentication.WSFederationAuthenticationModule.SignOut();
                Server.ClearError();
            }
        }
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings[""].ConnectionString);

        OleDbConnection Econ;



        protected void Page_Load(object sender, EventArgs e)
        {


            if (!IsPostBack)
            {

            }
            string maincon = ConfigurationManager.ConnectionStrings["crm_LC_MarketingConnectionString"].ConnectionString;
            SqlConnection sqlcon = new SqlConnection(maincon);

            string sqlquery = "truncate table [Marketing].[dbo].[AdsData$];truncate table [Marketing].[dbo].[AdGroup$];" +
                "truncate table [Marketing].[dbo].[AdsDataPiv$]";

            SqlCommand sqlcom = new SqlCommand(sqlquery, sqlcon);
            sqlcon.Open();
            sqlcom.ExecuteNonQuery();





        }
        private void ExcelConn(string filepath)
        {

            string constr = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;""", filepath);

            Econ = new OleDbConnection(constr);
            Econ.Open();
        }


        protected void btnupload_Click(object sender, EventArgs e)
        {

            con.Open();

            string filename = Guid.NewGuid() + Path.GetExtension(FileUpload1.PostedFile.FileName);

            string filepath = "/App_Data/" + filename;

            FileUpload1.PostedFile.SaveAs(Path.Combine(Server.MapPath("/App_Data"), filename));
            string fileName = Path.GetFileName(FileUpload1.PostedFile.FileName);
            string fullpath = Server.MapPath("/App_Data/") + filename;
            InsertExceldata(fullpath);

            string sqlquery = @"
                      delete  from [Marketing].[dbo].[AdsData$]
                      where [Offline Deposit_All conv#] is null;
                      update [Marketing].[dbo].[AdsData$] set [Description 1] = case when [Description 1]=' --' then [Description] else [Description 1] end;
                    insert into [Marketing].[dbo].[AdsDataPiv$]
                    SELECT distinct 
                    a.[Campaign]
                        ,a.[Ad Group]
                        ,[Keyword]

,[Final URL]= COALESCE([Final URL], '')
   ,  [Headline 1]= COALESCE([Headline 1], '')
  ,   [Headline 2]= COALESCE([Headline 2], '')
 ,     [Headline 3]= COALESCE([Headline 3], '')
      ,[Headline 4]= COALESCE([Headline 4], '')
      ,[Headline 5]= COALESCE([Headline 5], '')
      ,[Headline 6]= COALESCE([Headline 6], '')
      ,[Headline 7]= COALESCE([Headline 7], '')
      ,[Headline 8]= COALESCE([Headline 8], '')
      ,[Headline 9]= COALESCE([Headline 9], '')
      ,[Headline 10]= COALESCE([Headline 10], '')
      ,[Headline 11]= COALESCE([Headline 11], '')
      ,[Headline 12]= COALESCE([Headline 12], '')
      ,[Headline 13]= COALESCE([Headline 13], '')
      ,[Headline 14]= COALESCE([Headline 14], '')
      ,[Headline 15]= COALESCE([Headline 15], '')
      ,[Description 1]= COALESCE([Description 1], '')
      ,[Description 2]= COALESCE([Description 2], '')
      ,[Description 3]= COALESCE([Description 3], '')
      ,[Description 4]= COALESCE([Description 4], '')
      ,[Path 1]= COALESCE([Path 1], '')
      ,[Path 2]= COALESCE([Path 2], '')
,'' 
   ,b.[Ad type]
  FROM [Marketing].[dbo].[AdGroup$] a
 left join [Marketing].[dbo].[AdsData$] b on [Search term]=[Keyword];


 update [Marketing].[dbo].[AdsDataPiv$]	 set [Headline 1] = '' where [Headline 1]=' --';
  update [Marketing].[dbo].[AdsDataPiv$]    set [Headline 2] = '' where [Headline 2]=' --';
  update [Marketing].[dbo].[AdsDataPiv$]    set [Headline 3] = '' where [Headline 3]=' --';
 update [Marketing].[dbo].[AdsDataPiv$]    set [Headline 4] = '' where [Headline 4]=' --';
  update [Marketing].[dbo].[AdsDataPiv$]    set [Headline 5] = '' where [Headline 5]=' --';
 update [Marketing].[dbo].[AdsDataPiv$]   set [Headline 6] = '' where [Headline 6]=' --';
  update [Marketing].[dbo].[AdsDataPiv$]  set [Headline 7] = '' where [Headline 7]=' --';
 update [Marketing].[dbo].[AdsDataPiv$]   set [Headline 8] = '' where [Headline 8]=' --';
 update [Marketing].[dbo].[AdsDataPiv$]   set [Headline 9] = '' where [Headline 9]=' --';
  update [Marketing].[dbo].[AdsDataPiv$]   set [Headline 10] = '' where [Headline 10]=' --';
 update [Marketing].[dbo].[AdsDataPiv$]   set [Headline 11] = '' where [Headline 11]=' --';
 update [Marketing].[dbo].[AdsDataPiv$]   set [Headline 12] = '' where [Headline 12]=' --';
  update [Marketing].[dbo].[AdsDataPiv$]  set [Headline 13] = '' where [Headline 13]=' --';
 update [Marketing].[dbo].[AdsDataPiv$]    set [Headline 14] = '' where [Headline 14]=' --';
  update [Marketing].[dbo].[AdsDataPiv$]   set [Headline 15] = '' where [Headline 15]=' --';
  update [Marketing].[dbo].[AdsDataPiv$]   set [Description 1] = '' where [Description 1]=' --';
  update [Marketing].[dbo].[AdsDataPiv$]   set [Description 2] = '' where [Description 2]=' --';
  update [Marketing].[dbo].[AdsDataPiv$]  set [Description 3] = '' where [Description 3]=' --';
 update [Marketing].[dbo].[AdsDataPiv$]   set [Description 4] = '' where [Description 4]=' --';
 update [Marketing].[dbo].[AdsDataPiv$]     set [Path 1] = '' where [Path 1]=' --';
 update [Marketing].[dbo].[AdsDataPiv$]   set [Path 2] = '' where [Path 2]=' --';



  update [Marketing].[dbo].[AdsDataPiv$]   set [Ad found] = [Search term]
 from  [Marketing].[dbo].[AdsDataPiv$] a
 left join [Marketing].[dbo].[AdsData$] b on [Search term]=[Keyword];


select distinct * from [Marketing].[dbo].[AdsDataPiv$]
 ";


            SqlCommand cmd = new SqlCommand(sqlquery, con);
            cmd.CommandTimeout = 950;
            cmd.ExecuteNonQuery();

            SqlDataAdapter sqladapter = new SqlDataAdapter(cmd);

            DataTable dt = new DataTable();
            sqladapter.Fill(dt);


            GridView1.DataSource = dt;
            GridView1.DataBind();

            Label28.Text = "The process is done, please export to excel, this might take longer than usual due to data overload, please be patient";


        }

        protected void btnExportCS_Click(object sender, EventArgs e)
        {

            DateTime time = DateTime.Now;
            Response.Clear();
            Response.AddHeader("content-disposition", "attachment; filename = " + time + ".xls");
            Response.ContentType = "application/ms-excel";
            Response.ContentEncoding = System.Text.Encoding.Unicode;
            Response.BinaryWrite(System.Text.Encoding.Unicode.GetPreamble());

            StringWriter stringWriter = new StringWriter();
            HtmlTextWriter htmlWriter = new HtmlTextWriter(stringWriter);
            //if (ChckboxWithFixing.Checked)
            //{
            //    GridView2.RenderControl(htmlWriter);
            //}
            //else
            //{
            //    GridView3.RenderControl(htmlWriter);
            //}
            string s = stringWriter.ToString();
            Response.Write(s);
            Response.End();
        }

        public override void VerifyRenderingInServerForm(Control control)
        {
            return;
        }


        private void InsertExceldata(string filePath)
        {


            ExcelConn(filePath);

            DataTable Sheets = new DataTable();

            Sheets = Econ.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);




            string adsData = Sheets.Rows[0]["TABLE_NAME"].ToString();
            string adsGroup = Sheets.Rows[1]["TABLE_NAME"].ToString();





            string query = string.Format("Select * from [{0}]", adsData);

            string query1 = string.Format("Select * from [{0}]", adsGroup);

            OleDbCommand Ecom = new OleDbCommand(query, Econ);
            OleDbCommand Ecom1 = new OleDbCommand(query1, Econ);

            DataSet ds = new DataSet();

            OleDbDataAdapter oda = new OleDbDataAdapter(query, Econ);

           

           

            oda.Fill(ds);

            DataTable dt = ds.Tables[0];

            SqlBulkCopy objbulk = new SqlBulkCopy(con);


            AdsDataFunc(objbulk, dt);

            DataSet ds1 = new DataSet();

            OleDbDataAdapter oda1 = new OleDbDataAdapter(query1, Econ);

            oda1.Fill(ds1);

            DataTable dt1 = ds1.Tables[0];

            SqlBulkCopy objbulk1 = new SqlBulkCopy(con);
            AdsDataFunc(objbulk1,  dt1);
          

        }
        protected void AdsDataFunc(SqlBulkCopy objbulk, DataTable dt)
         {

            if (dt.Columns[0].ColumnName=="Campaign")
            {
                objbulk.DestinationTableName = "[Marketing].[dbo].[AdGroup$]";

                objbulk.ColumnMappings.Add(0, 0);
                objbulk.ColumnMappings.Add(1, 1);
                objbulk.ColumnMappings.Add(2, 2);
                objbulk.BatchSize = 10000;
                objbulk.BulkCopyTimeout = 0;
                objbulk.WriteToServer(dt);
                //con.Close();
            }
            else
            {
                objbulk.DestinationTableName = "[Marketing].[dbo].[AdsData$]";

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
                objbulk.ColumnMappings.Add(19, 19);
                objbulk.ColumnMappings.Add(20, 20);
                objbulk.ColumnMappings.Add(21, 21);
                objbulk.ColumnMappings.Add(22, 22);
                objbulk.ColumnMappings.Add(23, 23);
                objbulk.ColumnMappings.Add(24, 24);
                objbulk.ColumnMappings.Add(25, 25);
                objbulk.ColumnMappings.Add(26, 26);
                objbulk.ColumnMappings.Add(27, 27);
                objbulk.ColumnMappings.Add(28, 28);
                objbulk.ColumnMappings.Add(29, 29);




                //con.Open();
                objbulk.BatchSize = 10000;
                objbulk.BulkCopyTimeout = 0;
                objbulk.WriteToServer(dt);
               }
           
            }

        protected void Button1_Click(object sender, EventArgs e)
        {
           
        }

        
        protected void Button2_Click(object sender, EventArgs e)
        {
            DateTime time = DateTime.Now;
            Response.Clear();
            Response.AddHeader("content-disposition", "attachment; filename = " + time + ".xls");
            Response.ContentType = "application/ms-excel";
            Response.ContentEncoding = System.Text.Encoding.Unicode;
            Response.BinaryWrite(System.Text.Encoding.Unicode.GetPreamble());

            StringWriter stringWriter = new StringWriter();
            HtmlTextWriter htmlWriter = new HtmlTextWriter(stringWriter);
            
                GridView1.RenderControl(htmlWriter);
          
            string s = stringWriter.ToString();
            Response.Write(s);
            Response.End();
        }

        protected void btnFillVal_Click(object sender, EventArgs e)
        {

        }

        protected void btnfilval_Click(object sender, EventArgs e)
        {

        }

        protected void chkMY_CheckedChanged(object sender, EventArgs e)
        {
         //   Label10.Text = "MY";
            
        }
        
    }

}
