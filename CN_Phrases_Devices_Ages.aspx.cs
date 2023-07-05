using ClosedXML.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace DemoCrmH
{
    public partial class CN_Phrases_Devices_Ages : System.Web.UI.Page
    {

        

        string[] temporary_tables = { "#desktop_CN_Dep", "#mobile_CN_Dep", "#tablet_CN_Dep", "#mobtab_CN_Dep" };
        string[] temporary_tablesWW = { "#desktop_CN_Dep_MY", "#mobile_CN_Dep_MY", "#tablet_CN_Dep_MY", "#mobtab_CN_Dep_MY" };
        string[] final_tables = { "#final_Tbl_CN_desktop", "#final_Tbl_CN_mobile", "#final_Tbl_CN_tablet", "#final_Tbl_CN_mobtab" };
        string[] final_tablesww = { "#final_Tbl_CN_desktopww", "#final_Tbl_CN_mobileww", "#final_Tbl_CN_tabletww", "#final_Tbl_CN_mobtabww" };

        string[] tbljoin = { "#desktop_join", "#mobile_join", "#tablet_join", "#mobtab_join" };
        string[] final_tables8 = { "#desktop_phrases_final8", "#mobile_phrases_final8", "#tablet_phrases_final8", "#mobtab_phrases_final8" };

        string[] tbljoinWW = { "#desktop_joinWW", "#mobile_joinWW", "#tablet_joinWW", "#mobtab_joinWW" };
        string[] final_tables8WWW = { "#desktop_phrases_final8WW", "#mobile_phrases_final8WW", "#tablet_phrases_final8WW", "#mobtab_phrases_final8WW" };

        string[] temporary_tables8WW = { "#desktop_phrasesWW", "#mobile_phrasesWW", "#tablet_phrasesWW", "#mobtab_phrasesWW" };

        string[] final_tables_ages = { "#final_Tbl_CN_desktopages", "#final_Tbl_CN_mobileages", "#final_Tbl_CN_tabletages", "#final_Tbl_CN_mobtabages" };
        string[] final_tablesww_ages = { "#final_Tbl_CN_desktopwwages", "#final_Tbl_CN_mobilewwages", "#final_Tbl_CN_tabletwwages", "#final_TblCN_mobtabwwages" };
        string[] final_tables_ages_m = { "#final_Tbl_CN_desktopagesm", "#final_Tbl_CN_mobileagesm", "#final_Tbl_CN_tabletagesm", "#final_Tbl_CN_mobtabagesm" };


        string[] final_tables_ages_my = { "#final_Tbl_CN_desktop_agesmy", "#final_Tbl_CN_mobile_agesmy", "#final_Tbl_CN_tablet_agesmy", "#final_Tbl_CN_mobtab_agesmy" };
        string[] final_tablesww_ages_my = { "#final_Tbl_CN_desktopww_agesmy", "#final_Tbl_CN_mobileww_agesmy", "#final_Tbl_CN_tabletww_agesmy", "#final_Tbl_CN_mobtabww_agesmy" };
        string[] final_tables_ages_m_my = { "#final_Tbl_CN_desktop_agesmmy", "#final_Tbl_CN_mobile_agesmmy", "#final_Tbl_CN_tablet_agesmmy", "#final_Tbl_CN_mobtab_agesmmy" };





        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["crm_LC_MarketingConnectionString"].ConnectionString);

        OleDbConnection Econ;



        protected void Page_Load(object sender, EventArgs e)
        {

            //load data
            if (!IsPostBack)
            {
                //cn
                string selectSql = @"select * from [Marketing].[dbo].[SaveDataTblDevices$] where [ID]='CN_Phrase';";
                SqlCommand com = new SqlCommand(selectSql, con);

                try
                {
                    con.Open();

                    using (SqlDataReader read = com.ExecuteReader())
                    {
                        while (read.Read())
                        {
                            //Change the fields from db
                            txtCampaignNameSL1.Text = (read["CampDesktop"].ToString());
                            txtCampaignNameSL2.Text = (read["CampMobile"].ToString());
                            txtCampaignNameSL3.Text = (read["CampTablet"].ToString());
                            txtCampaignNameSL4.Text = (read["CampMobTab"].ToString());
                            txtDesiredCpaFtdBigDesktop.Text = (read["DesCpaBig1Desktop"].ToString());
                            txtDesiredCpaFtdBigMobile.Text = (read["DesCpaBig1Mobile"].ToString());
                            txtDesiredCpaFtdBigTablet.Text = (read["DesCpaBig1Tablet"].ToString());
                            txtDesiredCpaFtdBigMobTab.Text = (read["DesCpaBig1MobTab"].ToString());

                            txtDesiredCpaFtdEqDesktop.Text = (read["DesCpaEqDesktop"].ToString());
                            txtDesiredCpaFtdEqMobile.Text = (read["DesCpaEqMobile"].ToString());
                            txtDesiredCpaFtdEqTablet.Text = (read["DesCpaEqTablet"].ToString());
                            txtDesiredCpaFtdEqMobTab.Text = (read["DesCpaEqMobTab"].ToString());

                            txtDeviceAdjDesktop.Text = (read["DeviceAdjDesktop"].ToString());
                            txtDeviceAdjMobile.Text = (read["DeviceAdjMobile"].ToString());
                            txtDeviceAdjTablet.Text = (read["DeviceAdjTablet"].ToString());
                            txtDeviceAdjMobTab.Text = (read["DeviceAdjMobTab"].ToString());


                            txtMaxCpaBidFtdBigDesktop.Text = (read["MaxCpaBig1Desktop"].ToString());
                            txtMaxCpaBidFtdBigMobile.Text = (read["MaxCpaBig1Mobile"].ToString());
                            txtMaxCpaBidFtdBigTablet.Text = (read["MaxCpaBig1Tablet"].ToString());
                            txtMaxCpaBidFtdBigMobTab.Text = (read["MaxCpaBig1MobTab"].ToString());


                            txtMaxCpaBidFtdEqDesktop.Text = (read["MaxCpaEqDesktop"].ToString());
                            txtMaxCpaBidFtdEqMobile.Text = (read["MaxCpaEqMobile"].ToString());
                            txtMaxCpaBidFtdEqTablet.Text = (read["MaxCpaEqTablet"].ToString());
                            txtMaxCpaBidFtdEqMobTab.Text = (read["MaxCpaEqMobTab"].ToString());



                            txtMaxCpcBidFtdBigDesktop.Text = (read["MaxCpcBig1Desktop"].ToString());
                            txtMaxCpcBidFtdBigMobile.Text = (read["MaxCpcBig1Mobile"].ToString());
                            txtMaxCpcBidFtdBigTablet.Text = (read["MaxCpcBig1Tablet"].ToString());
                            txtMaxCpcBidFtdBigMobTab.Text = (read["MaxCpcBig1MobTab"].ToString());


                            txtMaxCpcBidFtdEqDesktop.Text = (read["MaxCpcEqDesktop"].ToString());
                            txtMaxCpcBidFtdEqMobile.Text = (read["MaxCpcEqMobile"].ToString());
                            txtMaxCpcBidFtdEqTablet.Text = (read["MaxCpcEqTablet"].ToString());
                            txtMaxCpcBidFtdEqMobTab.Text = (read["MaxCpcEqMobTab"].ToString());




                        }
                    }


                    //cn my
                    string selectSql_e = @"select * from [Marketing].[dbo].[SaveDataTblDevices$] where [ID]='CN_Phrase_MY';";
                    SqlCommand com1 = new SqlCommand(selectSql_e, con);



                    using (SqlDataReader read = com1.ExecuteReader())
                    {
                        while (read.Read())
                        {

                            txtCampaignNameSLWW1.Text = (read["CampDesktop"].ToString());
                            txtCampaignNameSLWW2.Text = (read["CampMobile"].ToString());
                            txtCampaignNameSLWW3.Text = (read["CampTablet"].ToString());
                            txtCampaignNameSLWW4.Text = (read["CampMobTab"].ToString());

                            txtDesiredCpaFtdBigDesktop1.Text = (read["DesCpaBig1Desktop"].ToString());
                            txtDesiredCpaFtdBigMobile1.Text = (read["DesCpaBig1Mobile"].ToString());
                            txtDesiredCpaFtdBigTablet1.Text = (read["DesCpaBig1Tablet"].ToString());
                            txtDesiredCpaFtdBigMobTab1.Text = (read["DesCpaBig1MobTab"].ToString());

                            txtDesiredCpaFtdEqDesktop1.Text = (read["DesCpaEqDesktop"].ToString());
                            txtDesiredCpaFtdEqMobile1.Text = (read["DesCpaEqMobile"].ToString());
                            txtDesiredCpaFtdEqTablet1.Text = (read["DesCpaEqTablet"].ToString());
                            txtDesiredCpaFtdEqMobTab1.Text = (read["DesCpaEqMobTab"].ToString());

                            txtDeviceAdjDesktop1.Text = (read["DeviceAdjDesktop"].ToString());
                            txtDeviceAdjMobile1.Text = (read["DeviceAdjMobile"].ToString());
                            txtDeviceAdjTablet1.Text = (read["DeviceAdjTablet"].ToString());
                            txtDeviceAdjMobTab1.Text = (read["DeviceAdjMobTab"].ToString());


                            txtMaxCpaBidFtdBigDesktop1.Text = (read["MaxCpaBig1Desktop"].ToString());
                            txtMaxCpaBidFtdBigMobile1.Text = (read["MaxCpaBig1Mobile"].ToString());
                            txtMaxCpaBidFtdBigTablet1.Text = (read["MaxCpaBig1Tablet"].ToString());
                            txtMaxCpaBidFtdBigMobTab1.Text = (read["MaxCpaBig1MobTab"].ToString());


                            txtMaxCpaBidFtdEqDesktop1.Text = (read["MaxCpaEqDesktop"].ToString());
                            txtMaxCpaBidFtdEqMobile1.Text = (read["MaxCpaEqMobile"].ToString());
                            txtMaxCpaBidFtdEqTablet1.Text = (read["MaxCpaEqTablet"].ToString());
                            txtMaxCpaBidFtdEqMobTab1.Text = (read["MaxCpaEqMobTab"].ToString());



                            txtMaxCpcBidFtdBigDesktop1.Text = (read["MaxCpcBig1Desktop"].ToString());
                            txtMaxCpcBidFtdBigMobile1.Text = (read["MaxCpcBig1Mobile"].ToString());
                            txtMaxCpcBidFtdBigTablet1.Text = (read["MaxCpcBig1Tablet"].ToString());
                            txtMaxCpcBidFtdBigMobTab1.Text = (read["MaxCpcBig1MobTab"].ToString());


                            txtMaxCpcBidFtdEqDesktop1.Text = (read["MaxCpcEqDesktop"].ToString());
                            txtMaxCpcBidFtdEqMobile1.Text = (read["MaxCpcEqMobile"].ToString());
                            txtMaxCpcBidFtdEqTablet1.Text = (read["MaxCpcEqTablet"].ToString());
                            txtMaxCpcBidFtdEqMobTab1.Text = (read["MaxCpcEqMobTab"].ToString());
                        }
                    }

                    //cn ages
                    string selectSqlAges = @"select * from [Marketing].[dbo].[SaveDataTblDevices$] where [ID]='CN_Phrase_Ages';";
                    SqlCommand comAges = new SqlCommand(selectSqlAges, con);
                    using (SqlDataReader read = comAges.ExecuteReader())
                    {
                        while (read.Read())
                        {
                       
                            txt25_34_CN.Text = (read["25_34"].ToString());
                            txt35_44_CN.Text = (read["35_44"].ToString());
                            txt45_54_CN.Text = (read["45_54"].ToString());
                            txt55_64_CN.Text = (read["55_64"].ToString());
                            txt65_CN.Text = (read["65 or more"].ToString());
                            txt_unknown_CN.Text = (read["Unknown"].ToString());

                            txtCampaignNameSLages1.Text = (read["CampDesktop"].ToString());
                            txtCampaignNameSLages2.Text = (read["CampMobile"].ToString());
                            txtCampaignNameSLages3.Text = (read["CampTablet"].ToString());
                            txtCampaignNameSLages4.Text = (read["CampMobTab"].ToString());
                            txtDesiredCpaFtdBigDesktopAges.Text = (read["DesCpaBig1Desktop"].ToString());
                            txtDesiredCpaFtdBigMobileAges.Text = (read["DesCpaBig1Mobile"].ToString());
                            txtDesiredCpaFtdBigTabletAges.Text = (read["DesCpaBig1Tablet"].ToString());
                            txtDesiredCpaFtdBigMobTabAges.Text = (read["DesCpaBig1MobTab"].ToString());

                            txtDesiredCpaFtdEqDesktopAges.Text = (read["DesCpaEqDesktop"].ToString());
                            txtDesiredCpaFtdEqMobileAges.Text = (read["DesCpaEqMobile"].ToString());
                            txtDesiredCpaFtdEqTabletAges.Text = (read["DesCpaEqTablet"].ToString());
                            txtDesiredCpaFtdEqMobTabAges.Text = (read["DesCpaEqMobTab"].ToString());

                            txtDeviceAdjDesktopAges.Text = (read["DeviceAdjDesktop"].ToString());
                            txtDeviceAdjMobileAges.Text = (read["DeviceAdjMobile"].ToString());
                            txtDeviceAdjTabletAges.Text = (read["DeviceAdjTablet"].ToString());
                            txtDeviceAdjMobTabAges.Text = (read["DeviceAdjMobTab"].ToString());


                            txtMaxCpaBidFtdBigDesktopAges.Text = (read["MaxCpaBig1Desktop"].ToString());
                            txtMaxCpaBidFtdBigMobileAges.Text = (read["MaxCpaBig1Mobile"].ToString());
                            txtMaxCpaBidFtdBigTabletAges.Text = (read["MaxCpaBig1Tablet"].ToString());
                            txtMaxCpaBidFtdBigMobTabAges.Text = (read["MaxCpaBig1MobTab"].ToString());


                            txtMaxCpaBidFtdEqDesktopAges.Text = (read["MaxCpaEqDesktop"].ToString());
                            txtMaxCpaBidFtdEqMobileAges.Text = (read["MaxCpaEqMobile"].ToString());
                            txtMaxCpaBidFtdEqTabletAges.Text = (read["MaxCpaEqTablet"].ToString());
                            txtMaxCpaBidFtdEqMobTabAges.Text = (read["MaxCpaEqMobTab"].ToString());



                            txtMaxCpcBidFtdBigDesktopAges.Text = (read["MaxCpcBig1Desktop"].ToString());
                            txtMaxCpcBidFtdBigMobileAges.Text = (read["MaxCpcBig1Mobile"].ToString());
                            txtMaxCpcBidFtdBigTabletAges.Text = (read["MaxCpcBig1Tablet"].ToString());
                            txtMaxCpcBidFtdBigMobTabAges.Text = (read["MaxCpcBig1MobTab"].ToString());


                            txtMaxCpcBidFtdEqDesktopAges.Text = (read["MaxCpcEqDesktop"].ToString());
                            txtMaxCpcBidFtdEqMobileAges.Text = (read["MaxCpcEqMobile"].ToString());
                            txtMaxCpcBidFtdEqTabletAges.Text = (read["MaxCpcEqTablet"].ToString());
                            txtMaxCpcBidFtdEqMobTabAges.Text = (read["MaxCpcEqMobTab"].ToString());




                        }
                    }


                    //my ages
                    string selectSql_eAges = @"select * from [Marketing].[dbo].[SaveDataTblDevices$] where [ID]='CN_Phrase_MY_Ages';";
                    SqlCommand com1Ages = new SqlCommand(selectSql_eAges, con);



                    using (SqlDataReader read = com1Ages.ExecuteReader())
                    {
                        while (read.Read())
                        {
                            txt25_34.Text = (read["25_34"].ToString());
                            txt35_44.Text = (read["35_44"].ToString());
                            txt45_54.Text = (read["45_54"].ToString());
                            txt55_64.Text = (read["55_64"].ToString());
                            txt65.Text = (read["65 or more"].ToString());
                            txt_unknown.Text = (read["Unknown"].ToString());

                            txtCampaignNameSLWW1Ages.Text = (read["CampDesktop"].ToString());
                            txtCampaignNameSLWW2Ages.Text = (read["CampMobile"].ToString());
                            txtCampaignNameSLWW3Ages.Text = (read["CampTablet"].ToString());
                            txtCampaignNameSLWW4Ages.Text = (read["CampMobTab"].ToString());

                            txtDesiredCpaFtdBigDesktop1Ages.Text = (read["DesCpaBig1Desktop"].ToString());
                            txtDesiredCpaFtdBigMobile1Ages.Text = (read["DesCpaBig1Mobile"].ToString());
                            txtDesiredCpaFtdBigTablet1Ages.Text = (read["DesCpaBig1Tablet"].ToString());
                            txtDesiredCpaFtdBigMobTab1Ages.Text = (read["DesCpaBig1MobTab"].ToString());

                            txtDesiredCpaFtdEqDesktop1Ages.Text = (read["DesCpaEqDesktop"].ToString());
                            txtDesiredCpaFtdEqMobile1Ages.Text = (read["DesCpaEqMobile"].ToString());
                            txtDesiredCpaFtdEqTablet1Ages.Text = (read["DesCpaEqTablet"].ToString());
                            txtDesiredCpaFtdEqMobTab1Ages.Text = (read["DesCpaEqMobTab"].ToString());

                            txtDeviceAdjDesktop1Ages.Text = (read["DeviceAdjDesktop"].ToString());
                            txtDeviceAdjMobile1Ages.Text = (read["DeviceAdjMobile"].ToString());
                            txtDeviceAdjTablet1Ages.Text = (read["DeviceAdjTablet"].ToString());
                            txtDeviceAdjMobTab1Ages.Text = (read["DeviceAdjMobTab"].ToString());

                 
                            txtMaxCpaBidFtdBigDesktop1Ages.Text = (read["MaxCpaBig1Desktop"].ToString());
                            txtMaxCpaBidFtdBigMobile1Ages.Text = (read["MaxCpaBig1Mobile"].ToString());
                            txtMaxCpaBidFtdBigTablet1Ages.Text = (read["MaxCpaBig1Tablet"].ToString());
                            txtMaxCpaBidFtdBigMobTab1Ages.Text = (read["MaxCpaBig1MobTab"].ToString());

                    
                            txtMaxCpaBidFtdEqDesktop1Ages.Text = (read["MaxCpaEqDesktop"].ToString());
                            txtMaxCpaBidFtdEqMobile1Ages.Text = (read["MaxCpaEqMobile"].ToString());
                            txtMaxCpaBidFtdEqTablet1Ages.Text = (read["MaxCpaEqTablet"].ToString());
                            txtMaxCpaBidFtdEqMobTab1Ages.Text = (read["MaxCpaEqMobTab"].ToString());



                            txtMaxCpcBidFtdBigDesktop1Ages.Text = (read["MaxCpcBig1Desktop"].ToString());
                            txtMaxCpcBidFtdBigMobile1Ages.Text = (read["MaxCpcBig1Mobile"].ToString());
                            txtMaxCpcBidFtdBigTablet1Ages.Text = (read["MaxCpcBig1Tablet"].ToString());
                            txtMaxCpcBidFtdBigMobTab1Ages.Text = (read["MaxCpcBig1MobTab"].ToString());


                            txtMaxCpcBidFtdEqDesktop1Ages.Text = (read["MaxCpcEqDesktop"].ToString());
                            txtMaxCpcBidFtdEqMobile1Ages.Text = (read["MaxCpcEqMobile"].ToString());
                            txtMaxCpcBidFtdEqTablet1Ages.Text = (read["MaxCpcEqTablet"].ToString());
                            txtMaxCpcBidFtdEqMobTab1Ages.Text = (read["MaxCpcEqMobTab"].ToString());
                        }
                    }

                }




                finally
                {
                    con.Close();
                }
            }









            string maincon = ConfigurationManager.ConnectionStrings["crm_LC_MarketingConnectionString"].ConnectionString;
            SqlConnection sqlcon = new SqlConnection(maincon);

              string sqlquery = "truncate table [Marketing].[dbo].[CN_RawData$];truncate table [Marketing].[dbo].[Phrases_final_CN_CampaignsI$];truncate table [Marketing].[dbo].[Phrases_final_CN_AND_MY$];truncate table [Marketing].[dbo].[Phrases_final_CN_MY_CampaignsI$];DROP TABLE IF EXISTS #CN_Pivot;DROP TABLE IF EXISTS #CN_Pivot_Post;DROP TABLE IF EXISTS #CN_Pivot_Post_MY" +
                ";truncate table [Marketing].[dbo].[phrases8tables_CN$];truncate table [Marketing].[dbo].[CN_Pivot$];truncate table [Marketing].[dbo].[AgesCompareCN$] ;truncate table [Marketing].[dbo].[Phrases_final_CN_CampaignsAgesMY$];truncate table [Marketing].[dbo].[Phrases_final_CN_CampaignsAges$]; UPDATE [Marketing].[dbo].[ages] set [value]=null; UPDATE [Marketing].[dbo].[ages_my] set [value]=null;" +
                "truncate table [Marketing].[dbo].[AgesWithFixing$];truncate table [Marketing].[dbo].[AgesNoFixing$];truncate table [Marketing].[dbo].[ages_4tbl];" +
                "truncate table [Marketing].[dbo].[phrasesPivot_final_CN$];truncate table [Marketing].[dbo].[phrasesPivot_finalWW_CN$];truncate table [Marketing].[dbo].[CN_PivotRaw$];" +
                "";
            sqlquery += droptemptables(temporary_tables, temporary_tablesWW, final_tables, final_tablesww, final_tables_ages, final_tablesww_ages, final_tables_ages_m, final_tables_ages_my, final_tablesww_ages_my, final_tables_ages_m_my, tbljoin, final_tables8, tbljoinWW, final_tables8WWW, temporary_tables8WW);
            SqlCommand sqlcom = new SqlCommand(sqlquery, sqlcon);
            sqlcon.Open();
            sqlcom.ExecuteNonQuery();
            sqlcom.ExecuteNonQuery();
            //GridView1.DataSource = null;
            //GridView1.Columns.Clear();
        }
        public string droptemptables(string[] tables, string[] tablesWW, string[] final_tables, string[] final_tablesww, string[] final_tables_ages, string[] final_tablesww_ages, string[] final_tables_ages_m, string[] final_tables_ages_my, string[] final_tablesww_ages_my, string[] final_tables_ages_m_my, string[] tbljoin, string[] final_tables8, string[] tbljoinWW, string[] final_tables8WWW,string [] temporary_tables8WW)
        {
            string sqlstr = "";

            for (int i = 0; i < tables.Length; i++)
            {
                sqlstr += "DROP TABLE IF EXISTS " + tables[i] + "; ";
            }

            for (int i = 0; i < tablesWW.Length; i++)
            {
                sqlstr += "DROP TABLE IF EXISTS " + tablesWW[i] + "; ";
            }
            for (int i = 0; i < final_tables.Length; i++)
            {
                sqlstr += "DROP TABLE IF EXISTS " + tablesWW[i] + "; ";
            }
            for (int i = 0; i < final_tablesww.Length; i++)
            {
                sqlstr += "DROP TABLE IF EXISTS " + tablesWW[i] + "; ";
            }
            for (int i = 0; i < final_tables_ages.Length; i++)
            {
                sqlstr += "DROP TABLE IF EXISTS " + final_tables_ages[i] + "; ";
            }
            for (int i = 0; i < final_tablesww_ages.Length; i++)
            {
                sqlstr += "DROP TABLE IF EXISTS " + final_tablesww_ages[i] + "; ";
            }
            for (int i = 0; i < final_tables_ages_m.Length; i++)
            {
                sqlstr += "DROP TABLE IF EXISTS " + final_tables_ages_m[i] + "; ";
            }
            for (int i = 0; i < final_tables_ages_my.Length; i++)
            {
                sqlstr += "DROP TABLE IF EXISTS " + final_tables_ages_my[i] + "; ";
            }
            for (int i = 0; i < final_tablesww_ages_my.Length; i++)
            {
                sqlstr += "DROP TABLE IF EXISTS " + final_tablesww_ages_my[i] + "; ";
            }
            for (int i = 0; i < final_tables_ages_m_my.Length; i++)
            {
                sqlstr += "DROP TABLE IF EXISTS " + final_tables_ages_m_my[i] + "; ";
            }
            for (int i = 0; i < tbljoin.Length; i++)
            {
                sqlstr += "DROP TABLE IF EXISTS " + tbljoin[i] + "; ";
            }
            for (int i = 0; i < temporary_tables8WW.Length; i++)
            {
                sqlstr += "DROP TABLE IF EXISTS " + temporary_tables8WW[i] + "; ";
            }
            for (int i = 0; i < final_tables8.Length; i++)
            {
                sqlstr += "DROP TABLE IF EXISTS " + final_tables8[i] + "; ";
            }
            for (int i = 0; i < tbljoinWW.Length; i++)
            {
                sqlstr += "DROP TABLE IF EXISTS " + tbljoinWW[i] + "; ";
            }
            for (int i = 0; i < final_tables8WWW.Length; i++)
            {
                sqlstr += "DROP TABLE IF EXISTS " + final_tables8WWW[i] + "; ";
            }
            return sqlstr;
        }


        private void ExcelConn(string filepath)
        {

            string constr = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;""", filepath);

            Econ = new OleDbConnection(constr);
            Econ.Open();
        }
        private void InsertExceldata(string fileepath, string filename)
        {
            string fileName = Path.GetFileName(FileUpload1.PostedFile.FileName);
            string fullpath = Server.MapPath("/App_Data/") + filename;

            ExcelConn(fullpath);

            DataTable Sheets = new DataTable();

            Sheets = Econ.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            string fortradest = Sheets.Rows[0]["TABLE_NAME"].ToString();

            string query = string.Format("Select * from [{0}]", fortradest);

            OleDbCommand Ecom = new OleDbCommand(query, Econ);

            DataSet ds = new DataSet();

            OleDbDataAdapter oda = new OleDbDataAdapter(query, Econ);

            oda.Fill(ds);

            DataTable dt = ds.Tables[0];

            SqlBulkCopy objbulk = new SqlBulkCopy(con);

            objbulk.DestinationTableName = "[Marketing].[dbo].[CN_RawData$]";

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


          //  con.Open();
            objbulk.BatchSize = 10000;
            objbulk.BulkCopyTimeout = 0;
            objbulk.WriteToServer(dt);

        }


        public string GeneralMainTable( string campaignat, string deviceAdjust, string CamapignName, string final_table, string desired_cap_big, string desired_cap_eq, string device)
        {
            string SLquery = @"

SELECT distinct '" + campaignat + "'+[Search_term] as Campaign@Searchterm, [Search_term]"
                                         + @",[Cost]
                                      ,[Clicks]
                                      ,[Demo CRM]
                                      ,[FTD Approved]
                                      ,[All conv value]
                                      ,[L2FTD]
                                      ,[Conv rate]
                                      ,[CPL]
                                      ,[CPC]
                                      ,[ROI]
                                      ,[Target CPA by desired CPA]
                                      ,[Target CPA by desired CPA*2]
                                      ,[Target CPA by ROI]
                                      ,[Target CPC by desired CPA]
                                      ,[Target CPC by desired CPA*2]
                                      ,[Target CPC by ROI]
                                      ,[Max CPA]
                                      ,[Max CPC]
                                      ,[Device adjustment CPA]
                                      ,[Device Adjustment CPC]
                                      ,[Target CPA]
                                      ,[Target CPC], 
                                    '" + CamapignName + "' as [Campaign name]," + "[Search_term] as [Ad group name]"
                                + @"into " + final_table + "  FROM [Marketing].[dbo].[phrasesPivot_final_CN$];"

                             + @" alter table " + final_table + " add [type] nvarchar (500); update " + final_table + " set [type]= '" + device + "' + [Search_term]; "

                            +@"    update " + final_table + " set [Target CPA by desired CPA]= CASE when [FTD Approved]>1 THEN(" + desired_cap_big + " *[L2FTD]) else (" + desired_cap_eq + "*[L2FTD]) END,"
                                + @"[Target CPA by desired CPA*2] = (CASE when [FTD Approved] > 1 THEN(" + desired_cap_big + " *[L2FTD]) else (" + desired_cap_eq + " *[L2FTD]) END)*2" + ","
                                + @"[Target CPA by ROI]=([ROI]*[CPL])/2 ,
                                  [Target CPC by desired CPA] = CASE when [FTD Approved]>1 THEN(" + desired_cap_big + "*[L2FTD]*[Conv rate]) else (" + desired_cap_eq + "*[L2FTD]*[Conv rate]) END ,"
                             + @"[Target CPC by desired CPA*2] = (CASE when   [FTD Approved]>1 THEN(" + desired_cap_big + "*[L2FTD]*[Conv rate]) else (" + desired_cap_eq + "*[L2FTD]*[Conv rate]) END)*2 ,"
                             + @"[Target CPC by ROI]=([ROI]*[CPC])/2;"

           + @"



                update " + final_table + " set "
             + @"[Max CPA]= IIF(IIF([Target CPA by desired CPA] <[Target CPA by ROI],[Target CPA by ROI],[Target CPA by desired CPA]) <[Target CPA by desired CPA*2],IIF([Target CPA by desired CPA] <[Target CPA by ROI],[Target CPA by ROI],[Target CPA by desired CPA]),[Target CPA by desired CPA*2]) 
            ,[Max CPC]= IIF(IIF([Target CPC by desired CPA]<[Target CPC by ROI], [Target CPC by ROI], [Target CPC by desired CPA])<[Target CPC by desired CPA*2],IIF([Target CPC by desired CPA]<[Target CPC by ROI], [Target CPC by ROI], [Target CPC by desired CPA]),[Target CPC by desired CPA*2]);"


            + @"


                    update " + final_table + " set "
            + @" [Device adjustment CPA]= (IIF(IIF([Target CPA by desired CPA]<[Target CPA by ROI],[Target CPA by ROI],[Target CPA by desired CPA])<[Target CPA by desired CPA*2],IIF([Target CPA by desired CPA]<[Target CPA by ROI], [Target CPA by ROI], [Target CPA by desired CPA]),[Target CPA by desired CPA*2]))* " + deviceAdjust + ""
            + @",[Device Adjustment CPC]=(IIF(IIF([Target CPC by desired CPA]<[Target CPC by ROI],[Target CPC by ROI],[Target CPC by desired CPA])<[Target CPC by desired CPA*2],IIF([Target CPC by desired CPA]<[Target CPC by ROI],[Target CPC by ROI],[Target CPC by desired CPA]),[Target CPC by desired CPA*2]))* " + deviceAdjust + ",[Target CPA]='' ,[Target CPC]='' ";

            return SLquery;

        }

        public string GeneralMainTableAges(string campaignat, string deviceAdjust, string CamapignName, string final_table, string desired_cap_big, string desired_cap_eq, string device)
        {
            string SLquery = @"

SELECT distinct '" + campaignat + "'+[Search_term] as Campaign@Searchterm, [Search_term] as [Search term]"
                                         + @",[Cost]
                                      ,[Clicks]
                                      ,[Demo CRM]
                                      ,[FTD Approved]
                                      ,[All conv value]
                                      ,[L2FTD]
                                      ,[Conv rate]
                                      ,[CPL]
                                      ,[CPC]
                                      ,[ROI]
                                      ,[Target CPA by desired CPA]
                                      ,[Target CPA by desired CPA*2]
                                      ,[Target CPA by ROI]
                                      ,[Target CPC by desired CPA]
                                      ,[Target CPC by desired CPA*2]
                                      ,[Target CPC by ROI]
                                      ,[Max CPA]
                                      ,[Max CPC]
                                      ,[Device adjustment CPA]
                                      ,[Device Adjustment CPC]
                                      ,[Target CPA]
                                      ,[Target CPC], 
                                    '" + CamapignName + "' as [Campaign name]," + "[Search_term] as [Ad group name]"
                                + @"into " + final_table + "  FROM [Marketing].[dbo].[phrasesPivot_final_CN$];"

                             + @" alter table " + final_table + " add [type] nvarchar (500); update " + final_table + " set [type]= '" + device + "' + [Search term]; "

                            + @"    update " + final_table + " set [Target CPA by desired CPA]= CASE when [FTD Approved]>1 THEN(" + desired_cap_big + " *[L2FTD]) else (" + desired_cap_eq + "*[L2FTD]) END,"
                                + @"[Target CPA by desired CPA*2] = (CASE when [FTD Approved] > 1 THEN(" + desired_cap_big + " *[L2FTD]) else (" + desired_cap_eq + " *[L2FTD]) END)*2" + ","
                                + @"[Target CPA by ROI]=([ROI]*[CPL])/2 ,
                                  [Target CPC by desired CPA] = CASE when [FTD Approved]>1 THEN(" + desired_cap_big + "*[L2FTD]*[Conv rate]) else (" + desired_cap_eq + "*[L2FTD]*[Conv rate]) END ,"
                             + @"[Target CPC by desired CPA*2] = (CASE when   [FTD Approved]>1 THEN(" + desired_cap_big + "*[L2FTD]*[Conv rate]) else (" + desired_cap_eq + "*[L2FTD]*[Conv rate]) END)*2 ,"
                             + @"[Target CPC by ROI]=([ROI]*[CPC])/2;"

           + @"



                update " + final_table + " set "
             + @"[Max CPA]= IIF(IIF([Target CPA by desired CPA] <[Target CPA by ROI],[Target CPA by ROI],[Target CPA by desired CPA]) <[Target CPA by desired CPA*2],IIF([Target CPA by desired CPA] <[Target CPA by ROI],[Target CPA by ROI],[Target CPA by desired CPA]),[Target CPA by desired CPA*2]) 
            ,[Max CPC]= IIF(IIF([Target CPC by desired CPA]<[Target CPC by ROI], [Target CPC by ROI], [Target CPC by desired CPA])<[Target CPC by desired CPA*2],IIF([Target CPC by desired CPA]<[Target CPC by ROI], [Target CPC by ROI], [Target CPC by desired CPA]),[Target CPC by desired CPA*2]);"


            + @"


                    update " + final_table + " set "
            + @" [Device adjustment CPA]= (IIF(IIF([Target CPA by desired CPA]<[Target CPA by ROI],[Target CPA by ROI],[Target CPA by desired CPA])<[Target CPA by desired CPA*2],IIF([Target CPA by desired CPA]<[Target CPA by ROI], [Target CPA by ROI], [Target CPA by desired CPA]),[Target CPA by desired CPA*2]))* " + deviceAdjust + ""
            + @",[Device Adjustment CPC]=(IIF(IIF([Target CPC by desired CPA]<[Target CPC by ROI],[Target CPC by ROI],[Target CPC by desired CPA])<[Target CPC by desired CPA*2],IIF([Target CPC by desired CPA]<[Target CPC by ROI],[Target CPC by ROI],[Target CPC by desired CPA]),[Target CPC by desired CPA*2]))* " + deviceAdjust + ",[Target CPA]='' ,[Target CPC]='' ";

            return SLquery;

        }


        public string GeneralMainTableTargetCpaCpc(string site, string cpa_bideq1, string cpa_bidbig1, string cpc_bideq1, string cpc_big1, string final_table)
        {
            string TargetCpaCpc = @" 

                                        update " + final_table + ""
                                     + @" set   [Target CPA]  =case  when  [FTD Approved]<=1 and [Device adjustment CPA]  >= " + cpa_bideq1 + "  then " + cpa_bideq1 + "  when  [FTD Approved]<=1 and [Device adjustment CPA]  < " + cpa_bideq1 + "  then [Device adjustment CPA]"
                                      + @" when  [FTD Approved]>1 and[Device adjustment CPA] >= " + cpa_bidbig1 + "  then  " + cpa_bidbig1 + ""
                                      + @"  when  [FTD Approved]>1 and[Device adjustment CPA] <= " + cpa_bidbig1 + "  then [Device adjustment CPA] end,"
                                      + @" [Target CPC]= case  when  [FTD Approved]<=1 and [Device adjustment CPC]  >= " + cpc_bideq1 + "  then " + cpc_bideq1 + ""
                                      + @"   when  [FTD Approved]<=1 and [Device adjustment CPC]  <= " + cpc_bideq1 + "  then [Device adjustment CPC] "
                                      + @" when  [FTD Approved]>1 and [Device adjustment CPC] >= " + cpc_big1 + " then " + cpc_big1 + "   when  [FTD Approved]>1 and [Device adjustment CPC] <= " + cpc_big1 + "  then [Device adjustment CPC] end; insert into " + site + " select * from " + final_table + "";

            return TargetCpaCpc;

        }



        public string GeneralMainTableTargetCpaCpcAges(string cpa_bideq1, string cpa_bidbig1, string cpc_bideq1, string cpc_big1, string final_table, string site, string final_table_m, string cpa_bid_big1, string cpc_bid_big1)
        {
            string TargetCpaCpc =
         @"   SELECT distinct a.[campaign@searchterm],a.[Search term],a.[cost],a.[clicks],a.[Demo CRM],a.[FTD Approved],a.[All conv value],a.[L2FTD],a.[Conv rate],a.[CPL],a.[CPC],a.[ROI]
                                        ,a.[Target CPA by desired CPA],a.[Target CPA by desired CPA*2],a.[Target CPA by ROI],a.[Target CPC by desired CPA],a.[Target CPC by desired CPA*2],a.[Target CPC by ROI]
                                        ,a.[Max CPA],a.[Max CPC],a.[Device adjustment CPA],a.[Device Adjustment CPC],a.[Target CPA],a.[Target CPC],
                                        a.[Campaign name],a.[Ad group name] + ' phrase ' +  [Age] as 'Ad group name', b.[Age],b.[value]," + cpa_bid_big1 + " as [cpabig1]," + cpa_bideq1 + " as [cpaeq1]," + cpc_bid_big1 + " as[cpcbig1]," + cpc_bideq1 + " as [cpceq1], [type] "
                                + @" into " + final_table_m + "  FROM " + final_table + " a full outer join [Marketing].[dbo].[ages] b on 1 = 1; delete from " + final_table_m + " where [value] is null";

            TargetCpaCpc += @" insert into " + site + " select* from " + final_table_m + ";";

            return TargetCpaCpc;

        }


        public string GeneralMainTableTargetCpaCpcAgesMY(string cpa_bideq1, string cpa_bidbig1, string cpc_bideq1, string cpc_big1, string final_table, string site, string final_table_m, string cpa_bid_big1, string cpc_bid_big1)
        {
            string TargetCpaCpc =
         @"   SELECT distinct a.[campaign@searchterm],a.[Search term],a.[cost],a.[clicks],a.[Demo CRM],a.[FTD Approved],a.[All conv value],a.[L2FTD],a.[Conv rate],a.[CPL],a.[CPC],a.[ROI]
                                        ,a.[Target CPA by desired CPA],a.[Target CPA by desired CPA*2],a.[Target CPA by ROI],a.[Target CPC by desired CPA],a.[Target CPC by desired CPA*2],a.[Target CPC by ROI]
                                        ,a.[Max CPA],a.[Max CPC],a.[Device adjustment CPA],a.[Device Adjustment CPC],a.[Target CPA],a.[Target CPC],
                                        a.[Campaign name],a.[Ad group name] + ' phrase ' +  [Age] as 'Ad group name', b.[Age],b.[value]," + cpa_bid_big1 + " as [cpabig1]," + cpa_bideq1 + " as [cpaeq1]," + cpc_bid_big1 + " as[cpcbig1]," + cpc_bideq1 + " as [cpceq1], [type] "
                                + @" into " + final_table_m + "  FROM " + final_table + " a full outer join [Marketing].[dbo].[ages_my] b on 1 = 1; delete from " + final_table_m + " where [value] is null;";
          

            TargetCpaCpc += @" insert into " + site + " select* from " + final_table_m + ";";

            return TargetCpaCpc;

        }


        protected void btnupload_Click(object sender, EventArgs e)
        {


            con.Open();
           





         //   if (!chkMY.Checked)
        //    {
                //cn
                string query = @" UPDATE [Marketing].[dbo].[SaveDataTblDevices$]
                            SET [CampDesktop]='" + txtCampaignNameSL1.Text + "'"
                                + @",[CampMobile]='" + txtCampaignNameSL2.Text + "'"
                                + @",[CampTablet] = '" + txtCampaignNameSL3.Text + "'"
                                + @",[CampMobTab]='" + txtCampaignNameSL4.Text + "'"

                                + @",[DesCpaBig1Desktop]='" + txtDesiredCpaFtdBigDesktop.Text + "'"
                                + @",[DesCpaBig1Mobile]='" + txtDesiredCpaFtdBigMobile.Text + "'"
                                + @",[DesCpaBig1Tablet]='" + txtDesiredCpaFtdBigTablet.Text + "'"
                                + @",[DesCpaBig1MobTab]='" + txtDesiredCpaFtdBigMobTab.Text + "'"

                                + @",[DesCpaEqDesktop]='" + txtDesiredCpaFtdEqDesktop.Text + "'"
                                + @",[DesCpaEqMobile]='" + txtDesiredCpaFtdEqMobile.Text + "'"
                                + @",[DesCpaEqTablet]='" + txtDesiredCpaFtdEqTablet.Text + "'"
                                + @",[DesCpaEqMobTab]='" + txtDesiredCpaFtdEqMobTab.Text + "'"


                                + @",[DeviceAdjDesktop]='" + txtDeviceAdjDesktop.Text + "'"
                                + @",[DeviceAdjMobile]='" + txtDeviceAdjMobile.Text + "'"
                                + @",[DeviceAdjTablet]='" + txtDeviceAdjTablet.Text + "'"
                                + @",[DeviceAdjMobTab]='" + txtDeviceAdjMobTab.Text + "'"


                                + @",[MaxCpaBig1Desktop]='" + txtMaxCpaBidFtdBigDesktop.Text + "'"
                                + @",[MaxCpaBig1Mobile]='" + txtMaxCpaBidFtdBigMobile.Text + "'"
                                + @",[MaxCpaBig1Tablet]='" + txtMaxCpaBidFtdBigTablet.Text + "'"
                                + @",[MaxCpaBig1MobTab]='" + txtMaxCpaBidFtdBigMobTab.Text + "'"

                                + @",[MaxCpaEqDesktop]='" + txtMaxCpaBidFtdEqDesktop.Text + "'"
                                + @",[MaxCpaEqMobile]='" + txtMaxCpaBidFtdEqMobile.Text + "'"
                                + @",[MaxCpaEqTablet]='" + txtMaxCpaBidFtdEqTablet.Text + "'"
                                + @",[MaxCpaEqMobTab]='" + txtMaxCpaBidFtdEqMobTab.Text + "'"


                                + @",[MaxCpcBig1Desktop]='" + txtMaxCpcBidFtdBigDesktop.Text + "'"
                                + @",[MaxCpcBig1Mobile]='" + txtMaxCpcBidFtdBigMobile.Text + "'"
                                + @",[MaxCpcBig1Tablet]='" + txtMaxCpcBidFtdBigTablet.Text + "'"
                                + @",[MaxCpcBig1MobTab]='" + txtMaxCpcBidFtdBigMobTab.Text + "'"

                                + @",[MaxCpcEqDesktop]='" + txtMaxCpcBidFtdEqDesktop.Text + "'"
                                + @",[MaxCpcEqMobile]='" + txtMaxCpcBidFtdEqMobile.Text + "'"
                                + @",[MaxCpcEqTablet]='" + txtMaxCpcBidFtdEqTablet.Text + "'"
                                + @",[MaxCpcEqMobTab]='" + txtMaxCpcBidFtdEqMobTab.Text + "' WHERE[ID] = 'CN_Phrase';";


                SqlCommand command = new SqlCommand(query, con);



                command.ExecuteNonQuery();
                command.ExecuteScalar();




                //cn my
                string query1 = @" UPDATE [Marketing].[dbo].[SaveDataTblDevices$]
                           SET [CampDesktop]='" + txtCampaignNameSLWW1.Text + "'"
                            + @",[CampMobile]='" + txtCampaignNameSLWW2.Text + "'"
                            + @",[CampTablet] = '" + txtCampaignNameSLWW3.Text + "'"
                            + @",[CampMobTab]='" + txtCampaignNameSLWW4.Text + "'"

                            + @",[DesCpaBig1Desktop]='" + txtDesiredCpaFtdBigDesktop1.Text + "'"
                            + @",[DesCpaBig1Mobile]='" + txtDesiredCpaFtdBigMobile1.Text + "'"
                            + @",[DesCpaBig1Tablet]='" + txtDesiredCpaFtdBigTablet1.Text + "'"
                            + @",[DesCpaBig1MobTab]='" + txtDesiredCpaFtdBigMobTab1.Text + "'"

                            + @",[DesCpaEqDesktop]='" + txtDesiredCpaFtdEqDesktop1.Text + "'"
                            + @",[DesCpaEqMobile]='" + txtDesiredCpaFtdEqMobile1.Text + "'"
                            + @",[DesCpaEqTablet]='" + txtDesiredCpaFtdEqTablet1.Text + "'"
                            + @",[DesCpaEqMobTab]='" + txtDesiredCpaFtdEqMobTab1.Text + "'"


                            + @",[DeviceAdjDesktop]='" + txtDeviceAdjDesktop1.Text + "'"
                            + @",[DeviceAdjMobile]='" + txtDeviceAdjMobile1.Text + "'"
                            + @",[DeviceAdjTablet]='" + txtDeviceAdjTablet1.Text + "'"
                            + @",[DeviceAdjMobTab]='" + txtDeviceAdjMobTab1.Text + "'"


                            + @",[MaxCpaBig1Desktop]='" + txtMaxCpaBidFtdBigDesktop1.Text + "'"
                            + @",[MaxCpaBig1Mobile]='" + txtMaxCpaBidFtdBigMobile1.Text + "'"
                            + @",[MaxCpaBig1Tablet]='" + txtMaxCpaBidFtdBigTablet1.Text + "'"
                            + @",[MaxCpaBig1MobTab]='" + txtMaxCpaBidFtdBigMobTab1.Text + "'"

                            + @",[MaxCpaEqDesktop]='" + txtMaxCpaBidFtdEqDesktop1.Text + "'"
                            + @",[MaxCpaEqMobile]='" + txtMaxCpaBidFtdEqMobile1.Text + "'"
                            + @",[MaxCpaEqTablet]='" + txtMaxCpaBidFtdEqTablet1.Text + "'"
                            + @",[MaxCpaEqMobTab]='" + txtMaxCpaBidFtdEqMobTab1.Text + "'"


                            + @",[MaxCpcBig1Desktop]='" + txtMaxCpcBidFtdBigDesktop1.Text + "'"
                            + @",[MaxCpcBig1Mobile]='" + txtMaxCpcBidFtdBigMobile1.Text + "'"
                            + @",[MaxCpcBig1Tablet]='" + txtMaxCpcBidFtdBigTablet1.Text + "'"
                            + @",[MaxCpcBig1MobTab]='" + txtMaxCpcBidFtdBigMobTab1.Text + "'"

                            + @",[MaxCpcEqDesktop]='" + txtMaxCpcBidFtdEqDesktop1.Text + "'"
                            + @",[MaxCpcEqMobile]='" + txtMaxCpcBidFtdEqMobile1.Text + "'"
                            + @",[MaxCpcEqTablet]='" + txtMaxCpcBidFtdEqTablet1.Text + "'"
                            + @",[MaxCpcEqMobTab]='" + txtMaxCpcBidFtdEqMobTab1.Text + "' WHERE[ID] = 'CN_Phrase_MY';";

             
                                        SqlCommand command1 = new SqlCommand(query1, con);



                                        command1.ExecuteNonQuery();
                                        command1.ExecuteScalar();




                //ages
                //cn
                string queryCnAges = @"   UPDATE [Marketing].[dbo].[ages] set [value] = NULLIF('" + txt25_34_CN.Text + "', '') where Age='25-34';"
                                + @" UPDATE [Marketing].[dbo].[ages] set [value] =NULLIF('" + txt35_44_CN.Text + "', '') where Age='35-44';"
                                + @"  UPDATE [Marketing].[dbo].[ages] set [value]=NULLIF('" + txt45_54_CN.Text + "', '') where Age='45-54';"
                                + @"  UPDATE [Marketing].[dbo].[ages] set [value]=NULLIF('" + txt55_64_CN.Text + "', '') where Age='55-64';"
                                + @"  UPDATE [Marketing].[dbo].[ages] set [value]=NULLIF('" + txt65_CN.Text + "', '') where Age='65 or more';"
                                + @"  UPDATE [Marketing].[dbo].[ages] set [value]=NULLIF('" + txt_unknown_CN.Text + "', '') where Age='Unknown';"


                +@" UPDATE [Marketing].[dbo].[SaveDataTblDevices$]
                            SET [CampDesktop]='" + txtCampaignNameSLages1.Text + "'"
                                + @",[CampMobile]='" + txtCampaignNameSLages2.Text + "'"
                                + @",[CampTablet] = '" + txtCampaignNameSLages3.Text + "'"
                                + @",[CampMobTab]='" + txtCampaignNameSLages4.Text + "'"


                               + @",[25_34]=NULLIF('" + txt25_34_CN.Text + "', '')"
                              + @",[35_44]=NULLIF('" + txt35_44_CN.Text + "', '')"
                              + @",[45_54]=NULLIF('" + txt45_54_CN.Text + "', '')"
                              + @",[55_64]=NULLIF('" + txt55_64_CN.Text + "', '')"
                              + @",[65 or more]=NULLIF('" + txt65_CN.Text + "', '')"
                              + @",[Unknown]=NULLIF('" + txt_unknown_CN.Text + "', '')"

                                + @",[DesCpaBig1Desktop]='" + txtDesiredCpaFtdBigDesktopAges.Text + "'"
                                + @",[DesCpaBig1Mobile]='" + txtDesiredCpaFtdBigMobileAges.Text + "'"
                                + @",[DesCpaBig1Tablet]='" + txtDesiredCpaFtdBigTabletAges.Text + "'"
                                + @",[DesCpaBig1MobTab]='" + txtDesiredCpaFtdBigMobTabAges.Text + "'"

                                + @",[DesCpaEqDesktop]='" + txtDesiredCpaFtdEqDesktopAges.Text + "'"
                                + @",[DesCpaEqMobile]='" + txtDesiredCpaFtdEqMobileAges.Text + "'"
                                + @",[DesCpaEqTablet]='" + txtDesiredCpaFtdEqTabletAges.Text + "'"
                                + @",[DesCpaEqMobTab]='" + txtDesiredCpaFtdEqMobTabAges.Text + "'"


                                + @",[DeviceAdjDesktop]='" + txtDeviceAdjDesktopAges.Text + "'"
                                + @",[DeviceAdjMobile]='" + txtDeviceAdjMobileAges.Text + "'"
                                + @",[DeviceAdjTablet]='" + txtDeviceAdjTabletAges.Text + "'"
                                + @",[DeviceAdjMobTab]='" + txtDeviceAdjMobTabAges.Text + "'"


                                + @",[MaxCpaBig1Desktop]='" + txtMaxCpaBidFtdBigDesktopAges.Text + "'"
                                + @",[MaxCpaBig1Mobile]='" + txtMaxCpaBidFtdBigMobileAges.Text + "'"
                                + @",[MaxCpaBig1Tablet]='" + txtMaxCpaBidFtdBigTabletAges.Text + "'"
                                + @",[MaxCpaBig1MobTab]='" + txtMaxCpaBidFtdBigMobTabAges.Text + "'"

                                + @",[MaxCpaEqDesktop]='" + txtMaxCpaBidFtdEqDesktopAges.Text + "'"
                                + @",[MaxCpaEqMobile]='" + txtMaxCpaBidFtdEqMobileAges.Text + "'"
                                + @",[MaxCpaEqTablet]='" + txtMaxCpaBidFtdEqTabletAges.Text + "'"
                                + @",[MaxCpaEqMobTab]='" + txtMaxCpaBidFtdEqMobTabAges.Text + "'"


                                + @",[MaxCpcBig1Desktop]='" + txtMaxCpcBidFtdBigDesktopAges.Text + "'"
                                + @",[MaxCpcBig1Mobile]='" + txtMaxCpcBidFtdBigMobileAges.Text + "'"
                                + @",[MaxCpcBig1Tablet]='" + txtMaxCpcBidFtdBigTabletAges.Text + "'"
                                + @",[MaxCpcBig1MobTab]='" + txtMaxCpcBidFtdBigMobTabAges.Text + "'"

                                + @",[MaxCpcEqDesktop]='" + txtMaxCpcBidFtdEqDesktopAges.Text + "'"
                                + @",[MaxCpcEqMobile]='" + txtMaxCpcBidFtdEqMobileAges.Text + "'"
                                + @",[MaxCpcEqTablet]='" + txtMaxCpcBidFtdEqTabletAges.Text + "'"
                                + @",[MaxCpcEqMobTab]='" + txtMaxCpcBidFtdEqMobTabAges.Text + "' WHERE[ID] = 'CN_Phrase_Ages';";


                SqlCommand commandcnAges = new SqlCommand(queryCnAges, con);



                commandcnAges.ExecuteNonQuery();
                commandcnAges.ExecuteScalar();




                //cn my
                string queryCnAgesMY = @"  UPDATE [Marketing].[dbo].[ages_my] set [value] = NULLIF('" + txt25_34.Text + "', '') where Age = '25-34'; "
                                 + @" UPDATE [Marketing].[dbo].[ages_my] set [value] =NULLIF('" + txt35_44.Text + "', '') where Age='35-44';"
                                 + @"  UPDATE [Marketing].[dbo].[ages_my] set [value]=NULLIF('" + txt45_54.Text + "', '') where Age='45-54';"
                                 + @"  UPDATE [Marketing].[dbo].[ages_my] set [value]=NULLIF('" + txt55_64.Text + "', '') where Age='55-64';"
                                 + @"  UPDATE [Marketing].[dbo].[ages_my] set [value]=NULLIF('" + txt65.Text + "', '') where Age='65 or more';"
                                 + @"  UPDATE [Marketing].[dbo].[ages_my] set [value]=NULLIF('" + txt_unknown.Text + "', '') where Age='Unknown';"


                            +@"UPDATE [Marketing].[dbo].[SaveDataTblDevices$]
                           SET [CampDesktop]='" + txtCampaignNameSLWW1Ages.Text + "'"
                            + @",[CampMobile]='" + txtCampaignNameSLWW2Ages.Text + "'"
                            + @",[CampTablet] = '" + txtCampaignNameSLWW3Ages.Text + "'"
                            + @",[CampMobTab]='" + txtCampaignNameSLWW4Ages.Text + "'"

                                 + @",[25_34]=NULLIF('" + txt25_34.Text + "', '')"
                              + @",[35_44]=NULLIF('" + txt35_44.Text + "', '')"
                              + @",[45_54]=NULLIF('" + txt45_54.Text + "', '')"
                              + @",[55_64]=NULLIF('" + txt55_64.Text + "', '')"
                              + @",[65 or more]=NULLIF('" + txt65.Text + "', '')"
                              + @",[Unknown]=NULLIF('" + txt_unknown.Text + "', '')"

                            + @",[DesCpaBig1Desktop]='" + txtDesiredCpaFtdBigDesktop1Ages.Text + "'"
                            + @",[DesCpaBig1Mobile]='" + txtDesiredCpaFtdBigMobile1Ages.Text + "'"
                            + @",[DesCpaBig1Tablet]='" + txtDesiredCpaFtdBigTablet1Ages.Text + "'"
                            + @",[DesCpaBig1MobTab]='" + txtDesiredCpaFtdBigMobTab1Ages.Text + "'"

                            + @",[DesCpaEqDesktop]='" + txtDesiredCpaFtdEqDesktop1Ages.Text + "'"
                            + @",[DesCpaEqMobile]='" + txtDesiredCpaFtdEqMobile1Ages.Text + "'"
                            + @",[DesCpaEqTablet]='" + txtDesiredCpaFtdEqTablet1Ages.Text + "'"
                            + @",[DesCpaEqMobTab]='" + txtDesiredCpaFtdEqMobTab1Ages.Text + "'"


                            + @",[DeviceAdjDesktop]='" + txtDeviceAdjDesktop1Ages.Text + "'"
                            + @",[DeviceAdjMobile]='" + txtDeviceAdjMobile1Ages.Text + "'"
                            + @",[DeviceAdjTablet]='" + txtDeviceAdjTablet1Ages.Text + "'"
                            + @",[DeviceAdjMobTab]='" + txtDeviceAdjMobTab1Ages.Text + "'"


                            + @",[MaxCpaBig1Desktop]='" + txtMaxCpaBidFtdBigDesktop1Ages.Text + "'"
                            + @",[MaxCpaBig1Mobile]='" + txtMaxCpaBidFtdBigMobile1Ages.Text + "'"
                            + @",[MaxCpaBig1Tablet]='" + txtMaxCpaBidFtdBigTablet1Ages.Text + "'"
                            + @",[MaxCpaBig1MobTab]='" + txtMaxCpaBidFtdBigMobTab1Ages.Text + "'"
                                                        
                            + @",[MaxCpaEqDesktop]='" + txtMaxCpaBidFtdEqDesktop1Ages.Text + "'"
                            + @",[MaxCpaEqMobile]='" + txtMaxCpaBidFtdEqMobile1Ages.Text + "'"
                            + @",[MaxCpaEqTablet]='" + txtMaxCpaBidFtdEqTablet1Ages.Text + "'"
                            + @",[MaxCpaEqMobTab]='" + txtMaxCpaBidFtdEqMobTab1Ages.Text + "'"


                            + @",[MaxCpcBig1Desktop]='" + txtMaxCpcBidFtdBigDesktop1Ages.Text + "'"
                            + @",[MaxCpcBig1Mobile]='" + txtMaxCpcBidFtdBigMobile1Ages.Text + "'"
                            + @",[MaxCpcBig1Tablet]='" + txtMaxCpcBidFtdBigTablet1Ages.Text + "'"
                            + @",[MaxCpcBig1MobTab]='" + txtMaxCpcBidFtdBigMobTab1Ages.Text + "'"

                            + @",[MaxCpcEqDesktop]='" + txtMaxCpcBidFtdEqDesktop1Ages.Text + "'"
                            + @",[MaxCpcEqMobile]='" + txtMaxCpcBidFtdEqMobile1Ages.Text + "'"
                            + @",[MaxCpcEqTablet]='" + txtMaxCpcBidFtdEqTablet1Ages.Text + "'"
                            + @",[MaxCpcEqMobTab]='" + txtMaxCpcBidFtdEqMobTab1Ages.Text + "' WHERE[ID] = 'CN_Phrase_MY_Ages';";


                SqlCommand queryCnMyAges = new SqlCommand(queryCnAgesMY, con);



                queryCnMyAges.ExecuteNonQuery();
                queryCnMyAges.ExecuteScalar();



        //    }
           


            string desktop_camp = txtCampaignNameSL1.Text + "@";
            string mobile_camp = txtCampaignNameSL2.Text + "@";
            string tablet_camp = txtCampaignNameSL3.Text + "@";
            string mobtab_camp = txtCampaignNameSL4.Text + "@";
            string desktop_campww = txtCampaignNameSLWW1.Text + "@";
            string mobile_campww = txtCampaignNameSLWW2.Text + "@";
            string tablet_campww = txtCampaignNameSLWW3.Text + "@";
            string mobtab_campww = txtCampaignNameSLWW4.Text + "@";

            string filename = Guid.NewGuid() + Path.GetExtension(FileUpload1.PostedFile.FileName);

            string filepath = "/App_Data/" + filename;

            FileUpload1.PostedFile.SaveAs(Path.Combine(Server.MapPath("/App_Data"), filename));

            InsertExceldata(filepath, filename);

            // con.Open();




            string sqlquery = @"insert into [Marketing].[dbo].[CN_PivotRaw$] select [Device],[Search term],[Match type],[Campaign],[Cost],[Clicks],[Demo Crm],[FTD Approved],[All conv. value],'' as [Search_term]  from [Marketing].[dbo].[CN_Opt_OldAccounts$];" +

                                    @"insert into [Marketing].[dbo].[CN_PivotRaw$] select  [Device],[Search term],[Match type] ,[Campaign],[Cost],[Clicks],[Demo CRM],[FTD Approved],[All conv. value],'' as [Search_term]  from [Marketing].[dbo].[CN_RawData$];

                                    UPDATE [Marketing].[dbo].[CN_PivotRaw$]
                                    SET    [FTD Approved] =  CASE  
                                    WHEN [All conv. value] > 30 and [FTD Approved]=0  
                                    THEN  1
                                    ELSE  [FTD Approved]
                                    end
                                    UPDATE [Marketing].[dbo].[CN_PivotRaw$]
                                    set  [Demo CRM]  =  CASE  
                                    WHEN [FTD Approved] > 0   and   [Demo CRM] =0  
                                    THEN  [FTD Approved]
                                    ELSE   [Demo CRM]
                                    END 
                                    UPDATE [Marketing].[dbo].[CN_PivotRaw$]
                                    set [Clicks]  =  CASE  
                                    WHEN    [Demo CRM]     > 0   and [Clicks]=0  
                                    THEN    [Demo CRM]    
                                    ELSE   [Clicks]
                                    END 

                                    UPDATE [Marketing].[dbo].[CN_PivotRaw$]
                                    set [Cost]  =  CASE  
                                    WHEN  [Clicks] > 0   and [Cost]=0  
                                    THEN  1
                                    ELSE   [Cost] 
                                    END ";

            //if (chkMY.Checked)
            //{
            //    sqlquery += @"  
            //                        UPDATE [Marketing].[dbo].[CN_PivotRaw$]
            //                      set[All conv. value] = CASE
            //                                            WHEN[FTD Approved] > 0   and[All conv. value] = 0
            //                                            THEN[FTD Approved] * 50
            //                                            ELSE[All conv. value]
            //                                        END";
            //}


            sqlquery += @";


update [Marketing].[dbo].[CN_PivotRaw$] set [Search term] =
'#'+replace([Search term],' ','#')+'#'

--update [Marketing].[dbo].[CN_PivotRaw$] set Search term =
--[Search term] where [FTD Approved]>0;
                                

SELECT  distinct
[Search term]
,sum([Cost]) as 'cost'
,sum([Clicks]) as 'clicks'
,sum([Demo CRM]) as 'Demo CRM'
,sum([FTD Approved]) as 'FTD Approved'
,sum([All conv. value]) as 'All conv value'
into #CN_Pivot_Post							
FROM  [Marketing].[dbo].[CN_PivotRaw$]
group by [Search term],[Search term];


                                  
SELECT  
[Search term],
[Cost] as 'cost',
[Clicks] as 'clicks',
[Demo CRM] as 'Demo CRM',
[FTD Approved] as 'FTD Approved',
[All conv value] as 'All conv value'
into #CN_Pivot_Phrases
FROM #CN_Pivot_Post

select distinct [Search term]
into  #CN_Pivot_Phrases_ST
from #CN_Pivot_Phrases 
where  [FTD Approved]>0;;

insert into [Marketing].[dbo].[phrasesPivot_final_CN$]
SELECT   b.[Search term],								  		 
sum(a.[cost]),sum(a.[clicks]),sum(a.[Demo CRM]) ,sum(a.[FTD Approved]),sum(a.[All conv value]),'','','','','','','','','','','','','','','','',''
FROM #CN_Pivot_Phrases a
left  join #CN_Pivot_Phrases_ST b on a.[Search term]  like N'%'+ b.[Search term]+ '%'
group by b.[Search term];
                                    
update  [Marketing].[dbo].[phrasesPivot_final_CN$] set [Search_term] =
replace([Search_term],'#',' ') ;
update  [Marketing].[dbo].[phrasesPivot_final_CN$] set [Search_term] =LTRIM(RTRIM('  '+[Search_term]+' '));

                                    
--update   [Marketing].[dbo].[CN_PivotRaw$] set [Search term] =
--replace([Search term],'#',' ') ;
--update   [Marketing].[dbo].[CN_PivotRaw$] set [Search term] =LTRIM(RTRIM('  '+[Search term]+' '));


	 delete from  [Marketing].[dbo].[phrasesPivot_final_CN$] where [FTD Approved]=0;


update [Marketing].[dbo].[phrasesPivot_final_CN$]
set [L2FTD]= [FTD Approved]/[Demo CRM]
,[Conv rate]= [Demo CRM]/[Clicks]
,[CPL]=[Cost]/[Demo CRM]
,[CPC]=[Cost]/[Clicks]
,[ROI]=[All conv value]/[Cost];";




            string[] site = { "[Marketing].[dbo].[Phrases_final_CN_CampaignsI$]", "[Marketing].[dbo].[Phrases_final_CN_MY_CampaignsI$]" };


            string[] devices = { desktop_camp, mobile_camp, tablet_camp, mobtab_camp };
            string[] devicesWW = { desktop_campww, mobile_campww, tablet_campww, mobtab_campww };
            string[] devices_conditions = { "='computers'", "='Mobile phones'", "='Tablets'", "in ('Mobile phones','Tablets')" };
            string[] devices_conditions_SLnWW = { "([Campaign] not like '%WW%' and [Campaign] not like '%MY%') ", "([Campaign]  like '%MY%')" };
           
            string sqlquery8tables = "";
            string[] desired_cpa = { txtDesiredCpaFtdBigDesktop.Text, txtDesiredCpaFtdBigMobile.Text, txtDesiredCpaFtdBigTablet.Text, txtDesiredCpaFtdBigMobTab.Text };
            string[] cpa_bidbig1 = { txtMaxCpaBidFtdBigDesktop.Text, txtMaxCpaBidFtdBigMobile.Text, txtMaxCpaBidFtdBigTablet.Text, txtMaxCpaBidFtdBigMobTab.Text };
            string[] cpc_bidbig1 = { txtMaxCpcBidFtdBigDesktop.Text, txtMaxCpcBidFtdBigMobile.Text, txtMaxCpcBidFtdBigTablet.Text, txtMaxCpcBidFtdBigMobTab.Text };

            string[] desired_cpaWW = { txtDesiredCpaFtdBigDesktop1.Text, txtDesiredCpaFtdBigMobile1.Text, txtDesiredCpaFtdBigTablet1.Text, txtDesiredCpaFtdBigMobTab1.Text };
            string[] cpa_bidbig1WW = { txtMaxCpaBidFtdBigDesktop1.Text, txtMaxCpaBidFtdBigMobile1.Text, txtMaxCpaBidFtdBigTablet1.Text, txtMaxCpaBidFtdBigMobTab1.Text };
            string[] cpc_bidbig1WW = { txtMaxCpcBidFtdBigDesktop1.Text, txtMaxCpcBidFtdBigMobile1.Text, txtMaxCpcBidFtdBigTablet1.Text, txtMaxCpcBidFtdBigMobTab1.Text };


            string[] devices_types = { "desktop@", "mobile@", "tablet@", "mobtab@" };

            string ages_text = txt25_34_CN.Text + txt35_44_CN.Text + txt45_54_CN.Text + txt55_64_CN.Text + txt65_CN.Text + txt_unknown_CN.Text;
            string ages_text_my = txt25_34.Text + txt35_44.Text + txt45_54.Text + txt55_64.Text + txt65.Text + txt_unknown.Text;
            bool is_ages_empty = ages_text == "";
            bool is_ages_empty_my = ages_text_my == "";
            string[] ages_fields = { txt25_34_CN.Text, txt35_44_CN.Text, txt45_54_CN.Text, txt55_64_CN.Text, txt65_CN.Text, txt_unknown_CN.Text };
            string[] ages_labels = { lbl25_34.Text, lbl35_44.Text, lbl45_54.Text, lbl55_64.Text, lbl65.Text, lblUnknown.Text };
            ArrayList ages_selected = new ArrayList();
            ArrayList ages_lbles = new ArrayList();

            for (int i = 0; i < ages_fields.Length; i++)
            {

                if (ages_fields[i] != "")
                {
                    ages_selected.Add(ages_fields[i]);
                    ages_lbles.Add(ages_labels[i]);
                }

            }

            //cn cn
            if (!String.IsNullOrEmpty(txtCampaignNameSL1.Text))
            {

                sqlquery += GeneralMainTable( desktop_camp, txtDeviceAdjDesktop.Text, txtCampaignNameSL1.Text, final_tables[0], txtDesiredCpaFtdBigDesktop.Text, txtDesiredCpaFtdEqDesktop.Text, devices_types[0]);
                sqlquery += GeneralMainTableTargetCpaCpc(site[0],txtMaxCpaBidFtdEqDesktop.Text, txtMaxCpaBidFtdBigDesktop.Text, txtMaxCpcBidFtdEqDesktop.Text, txtMaxCpcBidFtdBigDesktop.Text, final_tables[0]);
                sqlquery8tables += insert8tablesSL(final_tables8[0], tbljoin[0], temporary_tables[0], devices[0], devices_conditions[0], devices_conditions_SLnWW[0], desired_cpa[0], cpa_bidbig1[0], cpc_bidbig1[0]);


            }
            if (!String.IsNullOrEmpty(txtCampaignNameSL2.Text))
            {

                sqlquery += GeneralMainTable(mobile_camp, txtDeviceAdjMobile.Text, txtCampaignNameSL2.Text, final_tables[1], txtDesiredCpaFtdBigMobile.Text, txtDesiredCpaFtdEqMobile.Text, devices_types[1]);
                sqlquery += GeneralMainTableTargetCpaCpc(site[0], txtMaxCpaBidFtdEqMobile.Text, txtMaxCpaBidFtdBigMobile.Text, txtMaxCpcBidFtdEqMobile.Text, txtMaxCpcBidFtdBigMobile.Text, final_tables[1]);
                sqlquery8tables += insert8tablesSL(final_tables8[1], tbljoin[1], temporary_tables[1], devices[1], devices_conditions[1], devices_conditions_SLnWW[0], desired_cpa[1], cpa_bidbig1[1], cpc_bidbig1[1]);

            }
            if (!String.IsNullOrEmpty(txtCampaignNameSL3.Text))
            {
                sqlquery += GeneralMainTable(tablet_camp, txtDeviceAdjTablet.Text, txtCampaignNameSL3.Text, final_tables[2], txtDesiredCpaFtdBigTablet.Text, txtDesiredCpaFtdEqTablet.Text, devices_types[2]);
                sqlquery += GeneralMainTableTargetCpaCpc(site[0], txtMaxCpaBidFtdEqTablet.Text, txtMaxCpaBidFtdBigTablet.Text, txtMaxCpcBidFtdEqTablet.Text, txtMaxCpcBidFtdBigTablet.Text, final_tables[2]);
                sqlquery8tables += insert8tablesSL(final_tables8[2], tbljoin[2], temporary_tables[2], devices[2], devices_conditions[2], devices_conditions_SLnWW[0], desired_cpa[2], cpa_bidbig1[2], cpc_bidbig1[2]);

            }

            if (!String.IsNullOrEmpty(txtCampaignNameSL4.Text))
            {

                sqlquery += GeneralMainTable(mobtab_camp, txtDeviceAdjMobTab.Text, txtCampaignNameSL4.Text, final_tables[3], txtDesiredCpaFtdBigMobTab.Text, txtDesiredCpaFtdEqMobTab.Text, devices_types[3]);
                sqlquery += GeneralMainTableTargetCpaCpc(site[0], txtMaxCpaBidFtdEqMobTab.Text, txtMaxCpaBidFtdBigMobTab.Text, txtMaxCpcBidFtdEqMobTab.Text, txtMaxCpcBidFtdBigMobTab.Text, final_tables[3]);
                sqlquery8tables += insert8tablesSL(final_tables8[3], tbljoin[3], temporary_tables[3], devices[3], devices_conditions[3], devices_conditions_SLnWW[0], desired_cpa[3], cpa_bidbig1[3], cpc_bidbig1[3]);

            }

            //cn my
            if (!String.IsNullOrEmpty(txtCampaignNameSLWW1.Text))
            {
                sqlquery += GeneralMainTable( desktop_campww, txtDeviceAdjDesktop1.Text, txtCampaignNameSLWW1.Text, final_tablesww[0], txtDesiredCpaFtdBigDesktop1.Text, txtDesiredCpaFtdEqDesktop1.Text, devices_types[0]);
                sqlquery += GeneralMainTableTargetCpaCpc(site[1], txtMaxCpaBidFtdEqDesktop1.Text, txtMaxCpaBidFtdBigDesktop1.Text, txtMaxCpcBidFtdEqDesktop1.Text, txtMaxCpcBidFtdBigDesktop1.Text, final_tablesww[0]);
                sqlquery8tables += insert8tablesSL(final_tables8WWW[0], tbljoinWW[0], temporary_tables8WW[0], devicesWW[0], devices_conditions[0], devices_conditions_SLnWW[1], desired_cpaWW[0], cpa_bidbig1WW[0], cpc_bidbig1WW[0]);
            }
            if (!String.IsNullOrEmpty(txtCampaignNameSLWW2.Text))
            {
                sqlquery += GeneralMainTable(mobile_campww, txtDeviceAdjMobile1.Text, txtCampaignNameSLWW2.Text, final_tablesww[1], txtDesiredCpaFtdBigMobile1.Text, txtDesiredCpaFtdEqMobile1.Text, devices_types[1]);
                sqlquery += GeneralMainTableTargetCpaCpc(site[1], txtMaxCpaBidFtdEqMobile1.Text, txtMaxCpaBidFtdBigMobile1.Text, txtMaxCpcBidFtdEqMobile1.Text, txtMaxCpcBidFtdBigMobile1.Text, final_tablesww[1]);
                sqlquery8tables += insert8tablesSL(final_tables8WWW[1], tbljoinWW[1], temporary_tables8WW[1], devicesWW[1], devices_conditions[1], devices_conditions_SLnWW[1], desired_cpaWW[1], cpa_bidbig1WW[1], cpc_bidbig1WW[1]);
            }

            if (!String.IsNullOrEmpty(txtCampaignNameSLWW3.Text))
            {
                sqlquery += GeneralMainTable(tablet_campww, txtDeviceAdjTablet1.Text, txtCampaignNameSLWW3.Text, final_tablesww[2], txtDesiredCpaFtdBigTablet1.Text, txtDesiredCpaFtdEqTablet1.Text, devices_types[2]);
                sqlquery += GeneralMainTableTargetCpaCpc(site[1], txtMaxCpaBidFtdEqDesktop1.Text, txtMaxCpaBidFtdBigDesktop1.Text, txtMaxCpcBidFtdEqDesktop1.Text, txtMaxCpcBidFtdBigDesktop1.Text, final_tablesww[2]);
                sqlquery8tables += insert8tablesSL(final_tables8WWW[2], tbljoinWW[2], temporary_tables8WW[2], devicesWW[2], devices_conditions[2], devices_conditions_SLnWW[1], desired_cpaWW[2], cpa_bidbig1WW[2], cpc_bidbig1WW[2]);
            }
            if (!String.IsNullOrEmpty(txtCampaignNameSLWW4.Text))
            {
                sqlquery += GeneralMainTable(mobtab_campww, txtDeviceAdjMobTab1.Text, txtCampaignNameSLWW4.Text, final_tablesww[3], txtDesiredCpaFtdBigMobTab1.Text, txtDesiredCpaFtdEqMobTab1.Text, devices_types[3]);
                sqlquery += GeneralMainTableTargetCpaCpc(site[1], txtMaxCpaBidFtdEqDesktop1.Text, txtMaxCpaBidFtdBigDesktop1.Text, txtMaxCpcBidFtdEqDesktop1.Text, txtMaxCpcBidFtdBigDesktop1.Text, final_tablesww[3]);
                sqlquery8tables += insert8tablesSL(final_tables8WWW[3], tbljoinWW[3], temporary_tables8WW[3], devicesWW[3], devices_conditions[3], devices_conditions_SLnWW[1], desired_cpaWW[3], cpa_bidbig1WW[3], cpc_bidbig1WW[3]);
            }




            //cn cn ages
            string site_ages = string.Format("[Marketing].[dbo].[Phrases_final_CN_CampaignsAges$]");
            if (!is_ages_empty)
            {
                string desktop_camp_ages = txtCampaignNameSLages1.Text + "@";
                string mobile_camp_ages = txtCampaignNameSLages2.Text + "@";
                string tablet_camp_ages = txtCampaignNameSLages3.Text + "@";
                string mobtab_camp_ages = txtCampaignNameSLages4.Text + "@";

                string[] devicesages = { desktop_camp_ages, mobile_camp_ages, tablet_camp_ages, mobtab_camp_ages };

                string[] desired_cpa_ages = { txtDesiredCpaFtdBigDesktopAges.Text, txtDesiredCpaFtdBigMobileAges.Text, txtDesiredCpaFtdBigTabletAges.Text, txtDesiredCpaFtdBigMobTabAges.Text };
                string[] cpa_bidbig1_ages = { txtMaxCpaBidFtdBigDesktopAges.Text, txtMaxCpaBidFtdBigMobileAges.Text, txtMaxCpaBidFtdBigTabletAges.Text, txtMaxCpaBidFtdBigMobTabAges.Text };
                string[] cpc_bidbig1_ages = { txtMaxCpcBidFtdBigDesktopAges.Text, txtMaxCpcBidFtdBigMobileAges.Text, txtMaxCpcBidFtdBigTabletAges.Text, txtMaxCpcBidFtdBigMobTabAges.Text };

              

                if (!String.IsNullOrEmpty(txtCampaignNameSLages1.Text))
                {

                    sqlquery += GeneralMainTableAges(desktop_camp_ages, txtDeviceAdjDesktopAges.Text, txtCampaignNameSLages1.Text, final_tables_ages[0], txtDesiredCpaFtdBigDesktopAges.Text, txtDesiredCpaFtdEqDesktopAges.Text,  devices_types[0]);
                    sqlquery += GeneralMainTableTargetCpaCpcAges(txtMaxCpaBidFtdEqDesktopAges.Text, txtMaxCpaBidFtdBigDesktopAges.Text, txtMaxCpcBidFtdEqDesktopAges.Text, txtMaxCpcBidFtdBigDesktopAges.Text, final_tables_ages[0], site_ages, final_tables_ages_m[0], cpa_bidbig1_ages[0], cpc_bidbig1_ages[0]);
               }
                if (!String.IsNullOrEmpty(txtCampaignNameSLages2.Text))
                {

                    sqlquery += GeneralMainTableAges(mobile_camp_ages, txtDeviceAdjMobileAges.Text, txtCampaignNameSLages2.Text, final_tables_ages[1], txtDesiredCpaFtdBigMobileAges.Text, txtDesiredCpaFtdEqMobileAges.Text, devices_types[1]);
                    sqlquery += GeneralMainTableTargetCpaCpcAges(txtMaxCpaBidFtdEqMobileAges.Text, txtMaxCpaBidFtdBigMobileAges.Text, txtMaxCpcBidFtdEqMobileAges.Text, txtMaxCpcBidFtdBigMobileAges.Text, final_tables_ages[1], site_ages, final_tables_ages_m[1], cpa_bidbig1_ages[1], cpc_bidbig1_ages[1]);

                }
                if (!String.IsNullOrEmpty(txtCampaignNameSLages3.Text))
                {
                    sqlquery += GeneralMainTableAges(tablet_camp_ages, txtDeviceAdjTabletAges.Text, txtCampaignNameSLages3.Text, final_tables_ages[2], txtDesiredCpaFtdBigTabletAges.Text, txtDesiredCpaFtdEqTabletAges.Text, devices_types[2]);
                    sqlquery += GeneralMainTableTargetCpaCpcAges(txtMaxCpaBidFtdEqTabletAges.Text, txtMaxCpaBidFtdBigTabletAges.Text, txtMaxCpcBidFtdEqTabletAges.Text, txtMaxCpcBidFtdBigTabletAges.Text, final_tables_ages[2], site_ages, final_tables_ages_m[2], cpa_bidbig1_ages[2], cpc_bidbig1_ages[2]);

                }
                if (!String.IsNullOrEmpty(txtCampaignNameSLages4.Text))
                {

                    sqlquery += GeneralMainTableAges(mobtab_camp_ages, txtDeviceAdjMobTabAges.Text, txtCampaignNameSLages4.Text, final_tables_ages[3], txtDesiredCpaFtdBigMobTabAges.Text, txtDesiredCpaFtdEqMobTabAges.Text, devices_types[3]);
                    sqlquery += GeneralMainTableTargetCpaCpcAges(txtMaxCpaBidFtdEqMobTabAges.Text, txtMaxCpaBidFtdBigMobTabAges.Text, txtMaxCpcBidFtdEqMobTabAges.Text, txtMaxCpcBidFtdBigMobTabAges.Text, final_tables_ages[3], site_ages, final_tables_ages_m[3], cpa_bidbig1_ages[3], cpc_bidbig1_ages[3]);

                }

            }








            string site_ages_my = string.Format("[Marketing].[dbo].[Phrases_final_CN_CampaignsAgesMY$]");

            //cn my ages

            if (!is_ages_empty_my)
            {
                string desktop_camp_ages_my = txtCampaignNameSLWW1Ages.Text + "@";
                string mobile_camp_ages_my = txtCampaignNameSLWW2Ages.Text + "@";
                string tablet_camp_ages_my = txtCampaignNameSLWW3Ages.Text + "@";
                string mobtab_camp_ages_my = txtCampaignNameSLWW4Ages.Text + "@";

                string[] devicesages = { desktop_camp_ages_my, mobile_camp_ages_my, tablet_camp_ages_my, mobtab_camp_ages_my};

                string[] desired_cpaWW_ages = { txtDesiredCpaFtdBigDesktop1Ages.Text, txtDesiredCpaFtdBigMobile1Ages.Text, txtDesiredCpaFtdBigTablet1Ages.Text, txtDesiredCpaFtdBigMobTab1Ages.Text };
                string[] cpa_bidbig1WW_ages = { txtMaxCpaBidFtdBigDesktop1Ages.Text, txtMaxCpaBidFtdBigMobile1Ages.Text, txtMaxCpaBidFtdBigTablet1Ages.Text, txtMaxCpaBidFtdBigMobTab1Ages.Text };
                string[] cpc_bidbig1WW_ages = { txtMaxCpcBidFtdBigDesktop1Ages.Text, txtMaxCpcBidFtdBigMobile1Ages.Text, txtMaxCpcBidFtdBigTablet1Ages.Text, txtMaxCpcBidFtdBigMobTab1Ages.Text };

               

                if (!String.IsNullOrEmpty(txtCampaignNameSLWW1Ages.Text))
                {

                    sqlquery += GeneralMainTableAges(desktop_camp_ages_my, txtDeviceAdjDesktop1Ages.Text, txtCampaignNameSLWW1Ages.Text, final_tables_ages_my[0], txtDesiredCpaFtdBigDesktop1Ages.Text, txtDesiredCpaFtdEqDesktop1Ages.Text,  devices_types[0]);
                    sqlquery += GeneralMainTableTargetCpaCpcAgesMY(txtMaxCpaBidFtdEqDesktop1Ages.Text, txtMaxCpaBidFtdBigDesktop1Ages.Text, txtMaxCpcBidFtdEqDesktop1Ages.Text, txtMaxCpcBidFtdBigDesktop1Ages.Text, final_tables_ages_my[0], site_ages_my, final_tables_ages_m_my[0], cpa_bidbig1WW_ages[0], cpc_bidbig1WW_ages[0]);
                }
                if (!String.IsNullOrEmpty(txtCampaignNameSLWW2Ages.Text))
                {

                    sqlquery += GeneralMainTableAges(mobile_camp_ages_my, txtDeviceAdjMobile1Ages.Text, txtCampaignNameSLWW2Ages.Text, final_tables_ages_my[1], txtDesiredCpaFtdBigMobile1Ages.Text, txtDesiredCpaFtdEqMobile1Ages.Text, devices_types[1]);
                    sqlquery += GeneralMainTableTargetCpaCpcAgesMY(txtMaxCpaBidFtdEqMobile1Ages.Text, txtMaxCpaBidFtdBigMobile1Ages.Text, txtMaxCpcBidFtdEqMobile1Ages.Text, txtMaxCpcBidFtdBigMobile1Ages.Text, final_tables_ages_my[1], site_ages_my, final_tables_ages_m_my[1], cpa_bidbig1WW_ages[1], cpc_bidbig1WW_ages[1]);

                }
                if (!String.IsNullOrEmpty(txtCampaignNameSLWW3Ages.Text))
                {
                    sqlquery += GeneralMainTableAges(tablet_camp_ages_my, txtDeviceAdjTablet1Ages.Text, txtCampaignNameSLWW3Ages.Text, final_tables_ages_my[2], txtDesiredCpaFtdBigTablet1Ages.Text, txtDesiredCpaFtdEqTablet1Ages.Text, devices_types[2]);
                    sqlquery += GeneralMainTableTargetCpaCpcAgesMY(txtMaxCpaBidFtdEqTablet1Ages.Text, txtMaxCpaBidFtdBigTablet1Ages.Text, txtMaxCpcBidFtdEqTablet1Ages.Text, txtMaxCpcBidFtdBigTablet1Ages.Text, final_tables_ages_my[2], site_ages_my, final_tables_ages_m_my[2], cpa_bidbig1WW_ages[2], cpc_bidbig1WW_ages[2]);

                }
                if (!String.IsNullOrEmpty(txtCampaignNameSLWW4Ages.Text))
                {

                    sqlquery += GeneralMainTableAges(mobtab_camp_ages_my, txtDeviceAdjMobTab1Ages.Text, txtCampaignNameSLWW4Ages.Text, final_tables_ages_my[3], txtDesiredCpaFtdBigMobTab1Ages.Text, txtDesiredCpaFtdEqMobTab1Ages.Text, devices_types[3]);
                    sqlquery += GeneralMainTableTargetCpaCpcAgesMY(txtMaxCpaBidFtdEqMobTab1Ages.Text, txtMaxCpaBidFtdBigMobTab1Ages.Text, txtMaxCpcBidFtdEqMobTab1Ages.Text, txtMaxCpcBidFtdBigMobTab1Ages.Text, final_tables_ages_my[3], site_ages_my, final_tables_ages_m_my[3], cpa_bidbig1WW_ages[3], cpc_bidbig1WW_ages[3]);

                }

            }

            sqlquery8tables += @"update [Marketing].[dbo].[phrases8tables_CN$] set [Search term] =
                            replace([Search term],'#',' ') ;
                            update  [Marketing].[dbo].[phrases8tables_CN$] set [Search term] =LTRIM(RTRIM('  '+[Search term]+' '));
                            update  [Marketing].[dbo].[phrases8tables_CN$] set [campaign@searchterm]= [campaign@searchterm]+[Search term];";

            SqlCommand final = new SqlCommand(sqlquery, con);
                final.CommandTimeout = 950;
                final.ExecuteNonQuery();

           SqlCommand sqlcom3 = new SqlCommand(sqlquery8tables, con);
            sqlcom3.CommandTimeout = 6000;
            sqlcom3.ExecuteNonQuery();
            string sqlquery1 = "";

            sqlquery1 += @"insert into [Marketing].[dbo].[AgesCompareCN$] SELECT distinct [type],
                                    CASE
                                        WHEN b.[Target CPA] is not null THEN b.[Target CPA]  
                                        ELSE a.[Target CPA]
                                    END AS 'Final CPA bid',
                                    CASE
                                        WHEN b.[Target CPC] is not null THEN b.[Target CPC]  
                                        ELSE a.[Target CPC]
                                    END AS 'Final CPC bid',a.[Target CPA],a.[Target CPC]
                                  
                                  FROM [Marketing].[dbo].[Phrases_final_CN_CampaignsI$] a
                                    left join [Marketing].[dbo].[phrases8tables_CN$] b on a.[campaign@searchterm]=b.[campaign@searchterm];


                                 delete from [Marketing].[dbo].[Phrases_final_CN_CampaignsI$] where [campaign@searchterm] is null;
                           

                                    insert into [Marketing].[dbo].[AgesCompare$] SELECT distinct [type],
                                    CASE
                                        WHEN b.[Target CPA] is not null THEN b.[Target CPA]  
                                        ELSE a.[Target CPA]
                                    END AS 'Final CPA bid',
                                    CASE
                                        WHEN b.[Target CPC] is not null THEN b.[Target CPC]  
                                        ELSE a.[Target CPC]
                                    END AS 'Final CPC bid',a.[Target CPA],a.[Target CPC]
                                  
                                  FROM [Marketing].[dbo].[Phrases_final_CN_MY_CampaignsI$] a
                                    left join [Marketing].[dbo].[phrases8tables_CN$] b on a.[campaign@searchterm]=b.[campaign@searchterm];




                                  delete from [Marketing].[dbo].[Phrases_final_CN_MY_CampaignsI$] where [campaign@searchterm] is null;




insert into [Marketing].[dbo].[Phrases_final_CN_AND_MY$]
                                SELECT distinct a.[campaign@searchterm],a.[Search_term],a.[cost],a.[clicks],a.[Demo CRM],a.[FTD Approved],a.[All conv value],a.[L2FTD],a.[Conv rate],a.[CPL],a.[CPC],a.[ROI]
                                    ,a.[Target CPA by desired CPA],a.[Target CPA by desired CPA*2],a.[Target CPA by ROI],a.[Target CPC by desired CPA],a.[Target CPC by desired CPA*2],a.[Target CPC by ROI]
                                    ,a.[Max CPA],a.[Max CPC],a.[Device adjustment CPA],a.[Device Adjustment CPC],a.[Target CPA],a.[Target CPC],
                                    CASE
                                        WHEN b.[Target CPA] is not null THEN b.[Target CPA]  
                                        ELSE a.[Target CPA]
                                    END AS 'Final CPA bid',
                                    CASE
                                        WHEN b.[Target CPC] is not null THEN b.[Target CPC]  
                                        ELSE a.[Target CPC]
                                    END AS 'Final CPC bid',
                                    a.[Campaign name],a.[Ad group name]+ ' phrase' as 'Ad group name',b.[Target CPA] as 'fixing cpa',b.[Target CPC] as 'fixing cpc'
                                    FROM [Marketing].[dbo].[Phrases_final_CN_CampaignsI$] a
                                    left join [Marketing].[dbo].[phrases8tables_CN$] b on a.[campaign@searchterm]=b.[campaign@searchterm]
                                    order by a.[Search_term];



insert into [Marketing].[dbo].[Phrases_final_CN_AND_MY$]
                                SELECT distinct a.[campaign@searchterm],a.[Search_term],a.[cost],a.[clicks],a.[Demo CRM],a.[FTD Approved],a.[All conv value],a.[L2FTD],a.[Conv rate],a.[CPL],a.[CPC],a.[ROI]
                                    ,a.[Target CPA by desired CPA],a.[Target CPA by desired CPA*2],a.[Target CPA by ROI],a.[Target CPC by desired CPA],a.[Target CPC by desired CPA*2],a.[Target CPC by ROI]
                                    ,a.[Max CPA],a.[Max CPC],a.[Device adjustment CPA],a.[Device Adjustment CPC],a.[Target CPA],a.[Target CPC],
                                    CASE
                                        WHEN b.[Target CPA] is not null THEN b.[Target CPA]  
                                        ELSE a.[Target CPA]
                                    END AS 'Final CPA bid',
                                    CASE
                                        WHEN b.[Target CPC] is not null THEN b.[Target CPC]  
                                        ELSE a.[Target CPC]
                                    END AS 'Final CPC bid',
                                    a.[Campaign name],a.[Ad group name]+ ' phrase' as 'Ad group name',b.[Target CPA] as 'fixing cpa',b.[Target CPC] as 'fixing cpc'
                                    FROM [Marketing].[dbo].[Phrases_final_CN_MY_CampaignsI$] a
                                    left join [Marketing].[dbo].[phrases8tables_CN$] b on a.[campaign@searchterm]=b.[campaign@searchterm]
                                    order by a.[Search_term];


select  distinct * from [Marketing].[dbo].[Phrases_final_CN_AND_MY$];";






            string sqlquery_ages = @"
                                    SELECT distinct a.[type],a.[campaign@searchterm],a.[Search_term],a.[cost],a.[clicks],a.[Demo CRM],a.[FTD Approved],a.[All conv value],a.[L2FTD],a.[Conv rate],a.[CPL],a.[CPC],a.[ROI]
                                    ,a.[Target CPA by desired CPA],a.[Target CPA by desired CPA*2],a.[Target CPA by ROI],a.[Target CPC by desired CPA],a.[Target CPC by desired CPA*2],a.[Target CPC by ROI]
                                    ,a.[Max CPA],a.[Max CPC],a.[Device adjustment CPA],a.[Device Adjustment CPC],
                                   --[value],b.[Final CPA bid] as 'originial final cpa',b.[Final CPC bid] as 'originial final cpc',a.[cpabig1],a.[cpaeq1],a.[cpcbig1],a.[cpceq1],
                                    CASE
                                        WHEN a.[FTD Approved]<=1 and (b.[Final CPA bid]*[value])>=a.[cpaeq1] THEN a.[cpaeq1]
                                        WHEN a.[FTD Approved]<=1 and (b.[Final CPA bid]*[value])<=a.[cpaeq1] THEN (b.[Final CPA bid]*[value])
                                        WHEN a.[FTD Approved]>1 and (b.[Final CPA bid]*[value])>=a.[cpabig1] THEN a.[cpabig1]
										WHEN a.[FTD Approved]>1 and (b.[Final CPA bid]*[value])<=a.[cpabig1] THEN (b.[Final CPA bid]*[value])
                                      END AS 'Final CPA bid',
                                   CASE
                                        WHEN a.[FTD Approved]<=1 and (b.[Final CPC bid]*[value])>=a.[cpceq1] THEN a.[cpceq1]
                                        WHEN a.[FTD Approved]<=1 and (b.[Final CPC bid]*[value])<=a.[cpceq1] THEN (b.[Final CPC bid]*[value])
                                        WHEN a.[FTD Approved]>1 and (b.[Final CPC bid]*[value])>=a.[cpcbig1] THEN a.[cpcbig1]
										WHEN a.[FTD Approved]>1 and (b.[Final CPC bid]*[value])<=a.[cpcbig1] THEN (b.[Final CPC bid]*[value])
                                      END AS 'Final CPC bid',
                                    a.[Campaign name],a.[Ad group name],a.[Age_f] as 'Age'
                                    from  [Marketing].[dbo].[Phrases_final_CN_CampaignsAges$] a
                                    left join [Marketing].[dbo].[AgesCompareCN$] b on a.[type]=b.[type]
                                    where [value] is not null
									order by a.[Search_term];";



 string sqlquery_ages_my= @"  SELECT distinct a.[type],a.[campaign@searchterm],a.[Search_term],a.[cost],a.[clicks],a.[Demo CRM],a.[FTD Approved],a.[All conv value],a.[L2FTD],a.[Conv rate],a.[CPL],a.[CPC],a.[ROI]
                                    ,a.[Target CPA by desired CPA],a.[Target CPA by desired CPA*2],a.[Target CPA by ROI],a.[Target CPC by desired CPA],a.[Target CPC by desired CPA*2],a.[Target CPC by ROI]
                                    ,a.[Max CPA],a.[Max CPC],a.[Device adjustment CPA],a.[Device Adjustment CPC],
                                   --[value],b.[Final CPA bid] as 'originial final cpa',b.[Final CPC bid] as 'originial final cpc',a.[cpabig1],a.[cpaeq1],a.[cpcbig1],a.[cpceq1],
                                    CASE
                                        WHEN a.[FTD Approved]<=1 and (b.[Final CPA bid]*[value])>=a.[cpaeq1] THEN a.[cpaeq1]
                                        WHEN a.[FTD Approved]<=1 and (b.[Final CPA bid]*[value])<=a.[cpaeq1] THEN (b.[Final CPA bid]*[value])
                                        WHEN a.[FTD Approved]>1 and (b.[Final CPA bid]*[value])>=a.[cpabig1] THEN a.[cpabig1]
										WHEN a.[FTD Approved]>1 and (b.[Final CPA bid]*[value])<=a.[cpabig1] THEN (b.[Final CPA bid]*[value])
                                      END AS 'Final CPA bid',
                                   CASE
                                        WHEN a.[FTD Approved]<=1 and (b.[Final CPC bid]*[value])>=a.[cpceq1] THEN a.[cpceq1]
                                        WHEN a.[FTD Approved]<=1 and (b.[Final CPC bid]*[value])<=a.[cpceq1] THEN (b.[Final CPC bid]*[value])
                                        WHEN a.[FTD Approved]>1 and (b.[Final CPC bid]*[value])>=a.[cpcbig1] THEN a.[cpcbig1]
										WHEN a.[FTD Approved]>1 and (b.[Final CPC bid]*[value])<=a.[cpcbig1] THEN (b.[Final CPC bid]*[value])
                                      END AS 'Final CPC bid',
                                    a.[Campaign name],a.[Ad group name],a.[Age_f] as 'Age'
                                    from  [Marketing].[dbo].[Phrases_final_CN_CampaignsAgesMY$] a
                                    left join [Marketing].[dbo].[AgesCompare$] b on a.[type]=b.[type]
                                    where [value] is not null
									order by a.[Search_term];";

            DataTable dtAll = new DataTable();
            SqlCommand cmd = new SqlCommand(sqlquery1, con);
            cmd.CommandTimeout = 950;
            cmd.ExecuteNonQuery();
            SqlDataAdapter sqladapter = new SqlDataAdapter(cmd);

            DataTable dt = new DataTable();
            sqladapter.Fill(dt);


            SqlCommand cmd_ages = new SqlCommand(sqlquery_ages, con);
            cmd_ages.CommandTimeout = 950;
            cmd_ages.ExecuteNonQuery();
            SqlDataAdapter sqladapter_ages = new SqlDataAdapter(cmd_ages);

            DataTable dt_ages = new DataTable();
            sqladapter_ages.Fill(dt_ages);

            dtAll = dt.Copy();
            dtAll.Merge(dt_ages, false);

            SqlCommand cmd_ages_my = new SqlCommand(sqlquery_ages_my, con);
            cmd_ages_my.CommandTimeout = 950;
            cmd_ages_my.ExecuteNonQuery();
            SqlDataAdapter sqladapter_ages_my = new SqlDataAdapter(cmd_ages_my);

            DataTable dt_ages_my = new DataTable();
            sqladapter_ages_my.Fill(dt_ages_my);

            DataTable general = new DataTable();
            general = dtAll.Copy();
            general.Merge(dt_ages_my, false);



            GridView1.DataSource = general;
            GridView1.DataBind();
            string sqlfinalqueryShortVersion = "";
            string sqlfinalqueryShortVersion1 = "";


            sqlfinalqueryShortVersion = @"select distinct a.[Campaign name],a.[Ad group name] as 'Ad group name', a.[Search_term] as 'Keyword','phrase' as 'Match type',CASE
                                                    WHEN b.[Target CPA] is not null THEN b.[Target CPA]  
                                                    ELSE a.[Target CPA]
                                                END AS 'CPA bid',
                                                CASE
                                                    WHEN b.[Target CPC] is not null THEN b.[Target CPC]  
                                                    ELSE a.[Target CPC]
                                                END AS 'MAX CPC'
                                             FROM [Marketing].[dbo].[Phrases_final_CN_AND_MY$] a
                                    left join [Marketing].[dbo].[phrases8tables_CN$] b on a.[campaign@searchterm]=b.[campaign@searchterm]
                                                order by a.[Search_term]";

            string sqlfinalqueryShortVersionAges = @"insert into [Marketing].[dbo].[AgesWithFixing$]
SELECT distinct  a.[Campaign name],a.[Ad group name],a.[Search_term] as 'Keyword','phrase' as 'Match type',
                                    CASE
                                        WHEN a.[FTD Approved]<=1 and(b.[Final CPA bid]*[value])>=a.[cpaeq1] THEN a.[cpaeq1]
                                       WHEN a.[FTD Approved]<=1 and(b.[Final CPA bid]*[value])<=a.[cpaeq1] THEN(b.[Final CPA bid]*[value])
                                        WHEN a.[FTD Approved]>1 and (b.[Final CPA bid]*[value])>=a.[cpabig1] THEN a.[cpabig1]
                                        WHEN a.[FTD Approved]>1 and(b.[Final CPA bid]*[value])<=a.[cpabig1] THEN(b.[Final CPA bid]*[value])
                                      END  AS 'CPA bid',
                                   CASE
                                        WHEN a.[FTD Approved]<= 1 and (b.[Final CPC bid]*[value])>=a.[cpceq1] THEN a.[cpceq1]
                                         WHEN a.[FTD Approved]<=1 and(b.[Final CPC bid]*[value])<=a.[cpceq1] THEN(b.[Final CPC bid]*[value])
                                        WHEN a.[FTD Approved]>1 and (b.[Final CPC bid]*[value])>=a.[cpcbig1] THEN a.[cpcbig1]
                                        WHEN a.[FTD Approved]>1 and(b.[Final CPC bid]*[value])<=a.[cpcbig1] THEN(b.[Final CPC bid]*[value])
                                      END AS 'MAX CPC',a.[Age_f] as 'Age'
                                   from [Marketing].[dbo].[Phrases_final_CN_CampaignsAges$] a
                                    left join [Marketing].[dbo].[AgesCompareCN$] b on a.[type]=b.[type]
                                    where [value] is not null
                                    order by a.[Search_term]; 

insert into [Marketing].[dbo].[AgesWithFixing$]
SELECT distinct  a.[Campaign name],a.[Ad group name],a.[Search_term] as 'Keyword','phrase' as 'Match type',
                                    CASE
                                        WHEN a.[FTD Approved]<=1 and(b.[Final CPA bid]*[value])>=a.[cpaeq1] THEN a.[cpaeq1]
                                       WHEN a.[FTD Approved]<=1 and(b.[Final CPA bid]*[value])<=a.[cpaeq1] THEN(b.[Final CPA bid]*[value])
                                        WHEN a.[FTD Approved]>1 and (b.[Final CPA bid]*[value])>=a.[cpabig1] THEN a.[cpabig1]
                                        WHEN a.[FTD Approved]>1 and(b.[Final CPA bid]*[value])<=a.[cpabig1] THEN(b.[Final CPA bid]*[value])
                                      END  AS 'CPA bid',
                                   CASE
                                        WHEN a.[FTD Approved]<= 1 and (b.[Final CPC bid]*[value])>=a.[cpceq1] THEN a.[cpceq1]
                                         WHEN a.[FTD Approved]<=1 and(b.[Final CPC bid]*[value])<=a.[cpceq1] THEN(b.[Final CPC bid]*[value])
                                        WHEN a.[FTD Approved]>1 and (b.[Final CPC bid]*[value])>=a.[cpcbig1] THEN a.[cpcbig1]
                                        WHEN a.[FTD Approved]>1 and(b.[Final CPC bid]*[value])<=a.[cpcbig1] THEN(b.[Final CPC bid]*[value])
                                      END AS 'MAX CPC',a.[Age_f] as 'Age'
                                   from [Marketing].[dbo].[Phrases_final_CN_CampaignsAgesMY$] a
                                    left join [Marketing].[dbo].[AgesCompare$] b on a.[type]=b.[type]
                                    where [value] is not null
                                    order by a.[Search_term]; 



select  distinct * from [Marketing].[dbo].[AgesWithFixing$]";



            SqlCommand finalShortFixing = new SqlCommand(sqlfinalqueryShortVersion, con);
            finalShortFixing.ExecuteNonQuery();
            SqlDataAdapter finaladapterShortFix = new SqlDataAdapter(finalShortFixing);
            DataTable finaldtShortFixing = new DataTable();
            finaladapterShortFix.Fill(finaldtShortFixing);

            SqlCommand finalShortFixingAges = new SqlCommand(sqlfinalqueryShortVersionAges, con);
            finalShortFixingAges.ExecuteNonQuery();
            SqlDataAdapter finaladapterShortFixAges = new SqlDataAdapter(finalShortFixingAges);
            DataTable finaldtShortFixingAges = new DataTable();
            finaladapterShortFixAges.Fill(finaldtShortFixingAges);

            DataTable dtAll_short = new DataTable();
            dtAll_short = finaldtShortFixing.Copy();
            dtAll_short.Merge(finaldtShortFixingAges, false);

            GridView2.DataSource = dtAll_short;
            GridView2.DataBind();

            sqlfinalqueryShortVersion1 = @"

select distinct a.[Campaign name],a.[Ad group name] as 'Ad group name', a.[Search_term] as 'Keyword','phrase' as 'Match type',
                                               a.[Target CPA] as 'CPA bid',a.[Target CPC] as 'MAX CPC'
                                              FROM [Marketing].[dbo].[Phrases_final_CN_AND_MY$] a
                                            left join [Marketing].[dbo].[phrases8tables_CN$] b on a.[campaign@searchterm]=b.[campaign@searchterm]
                                                order by a.[Search_term]";


            string sqlfinalqueryShortVersionAgesWithoutFixing = @"
insert into [Marketing].[dbo].[AgesNoFixing$] SELECT distinct  c.[Campaign name],c.[Ad group name],c.[Search_term] as 'Keyword','phrase' as 'Match type',
                                                            CASE
                                                            WHEN a.[FTD Approved]<=1 and(a.[Target CPA]*[value])>=c.[cpaeq1] THEN c.[cpaeq1]
                                                            WHEN a.[FTD Approved]<=1 and(a.[Target CPA]*[value])<=c.[cpaeq1] THEN(a.[Target CPA]*[value])
                                                            WHEN a.[FTD Approved]>1 and (a.[Target CPA]*[value])>=c.[cpabig1] THEN c.[cpabig1]
                                                            WHEN a.[FTD Approved]>1 and(a.[Target CPA]*[value])<=c.[cpabig1] THEN(a.[Target CPA]*[value])
                                                            END  AS 'CPA bid',
                                                            CASE
                                                            WHEN a.[FTD Approved]<= 1 and (a.[Target CPC]*[value])>=c.[cpceq1] THEN c.[cpceq1]
                                                            WHEN a.[FTD Approved]<=1 and(a.[Target CPC]*[value])<=c.[cpceq1] THEN(a.[Target CPC]*[value])
                                                            WHEN a.[FTD Approved]>1 and (a.[Target CPC]*[value])>=c.[cpcbig1] THEN c.[cpcbig1]
                                                            WHEN a.[FTD Approved]>1 and(a.[Target CPC]*[value])<=c.[cpcbig1] THEN(a.[Target CPC]*[value])
                                                            END AS 'MAX CPC',c.[Age_f] as 'Age'
                                                            from [Marketing].[dbo].[Phrases_final_CN_CampaignsAges$] c
                                                            left join [Marketing].[dbo].[Phrases_final_CN_CampaignsI$] a on a.[type]=c.[type]
                                                            left join [Marketing].[dbo].[AgesCompareCN$] b on a.[type]=b.[type]
                                                            where [value] is not null;

insert into [Marketing].[dbo].[AgesNoFixing$] SELECT distinct  c.[Campaign name],c.[Ad group name],c.[Search_term] as 'Keyword','phrase' as 'Match type',
                                                            CASE
                                                            WHEN a.[FTD Approved]<=1 and(a.[Target CPA]*[value])>=c.[cpaeq1] THEN c.[cpaeq1]
                                                            WHEN a.[FTD Approved]<=1 and(a.[Target CPA]*[value])<=c.[cpaeq1] THEN(a.[Target CPA]*[value])
                                                            WHEN a.[FTD Approved]>1 and (a.[Target CPA]*[value])>=c.[cpabig1] THEN c.[cpabig1]
                                                            WHEN a.[FTD Approved]>1 and(a.[Target CPA]*[value])<=c.[cpabig1] THEN(a.[Target CPA]*[value])
                                                            END  AS 'CPA bid',
                                                            CASE
                                                            WHEN a.[FTD Approved]<= 1 and (a.[Target CPC]*[value])>=c.[cpceq1] THEN c.[cpceq1]
                                                            WHEN a.[FTD Approved]<=1 and(a.[Target CPC]*[value])<=c.[cpceq1] THEN(a.[Target CPC]*[value])
                                                            WHEN a.[FTD Approved]>1 and (a.[Target CPC]*[value])>=c.[cpcbig1] THEN c.[cpcbig1]
                                                            WHEN a.[FTD Approved]>1 and(a.[Target CPC]*[value])<=c.[cpcbig1] THEN(a.[Target CPC]*[value])
                                                            END AS 'MAX CPC',c.[Age_f] as 'Age'
                                                            from [Marketing].[dbo].[Phrases_final_CN_CampaignsAgesMY$] c
                                                            left join [Marketing].[dbo].[Phrases_final_CN_MY_CampaignsI$] a on a.[type]=c.[type]

                                                            left join [Marketing].[dbo].[AgesCompare$] b on a.[type]=b.[type]
                                                            where [value] is not null;
                                    select distinct * from [Marketing].[dbo].[AgesNoFixing$];";


            DataTable dtAll_short_WithoutFixing = new DataTable();
            SqlCommand finalShort = new SqlCommand(sqlfinalqueryShortVersion1, con);
            finalShort.ExecuteNonQuery();
            SqlDataAdapter finaladapterShort = new SqlDataAdapter(finalShort);
            DataTable finaldtShort = new DataTable();
            finaladapterShort.Fill(finaldtShort);

            SqlCommand finalShortAges = new SqlCommand(sqlfinalqueryShortVersionAgesWithoutFixing, con);
            finalShortAges.ExecuteNonQuery();
            SqlDataAdapter finaladapterShortAges = new SqlDataAdapter(finalShortAges);
            DataTable finaldtShortAges = new DataTable();
            finaladapterShortAges.Fill(finaldtShortAges);

            DataTable dtAll_shortWithoutFix = new DataTable();
            dtAll_shortWithoutFix = finaldtShort.Copy();
            dtAll_shortWithoutFix.Merge(finaldtShortAges, false);

            GridView3.DataSource = dtAll_shortWithoutFix;
            GridView3.DataBind();

            string tbl4 = @"insert into [Marketing].[dbo].[ages_4tbl] SELECT distinct  [Campaign name],[Ad group name],[Age] FROM [Marketing].[dbo].[Deposit_final_CN_CampaignsAges$] a
                            full outer join [Marketing].[dbo].[ages] b on  1=1
                            where [Age]!=[Age_f]
							and a.[value] is not null
							and b.[value] is not null
							order by [Ad group name];


insert into [Marketing].[dbo].[ages_4tbl] SELECT distinct  [Campaign name],[Ad group name],[Age] FROM [Marketing].[dbo].[Deposit_final_CN_CampaignsAgesMY$] a
                            full outer join [Marketing].[dbo].[ages_my] b on  1=1
                            where [Age]!=[Age_f]
							and a.[value] is not null
							and b.[value] is not null
							order by [Ad group name];

select * from [Marketing].[dbo].[ages_4tbl];
";


            SqlCommand tbl4th = new SqlCommand(tbl4, con);
            tbl4th.ExecuteNonQuery();
            SqlDataAdapter tbl4da = new SqlDataAdapter(tbl4th);
            DataTable tbl4dt = new DataTable();
            tbl4da.Fill(tbl4dt);
            gvAges.DataSource = tbl4dt;
            gvAges.DataBind();

            Label28.Text = "The process is done, file will be downloaded automatically is few seconds";

            dtAll.TableName = "full results";
            dtAll_short.TableName = "short version with fixing";
            dtAll_shortWithoutFix.TableName = "short version without fixing";
            tbl4dt.TableName = "ages";

            DataSet ds = new DataSet();
            ds.Tables.Add(dtAll);
            ds.Tables.Add(dtAll_short);
            ds.Tables.Add(dtAll_shortWithoutFix);
            ds.Tables.Add(tbl4dt);

            DateTime time = DateTime.Now;

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(ds);
                wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wb.Style.Font.Bold = true;
                Response.Clear();
                Response.Buffer = true;
                Response.Charset = "";
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //  Response.AddHeader("content-disposition", "attachment;filename= results.xlsx");
                Response.AddHeader("content-disposition", "attachment; filename = " + time + ".xlsx");

                using (MemoryStream MyMemoryStream = new MemoryStream())
                {
                    wb.SaveAs(MyMemoryStream);
                    MyMemoryStream.WriteTo(Response.OutputStream);
                    Response.Flush();
                    Response.End();

                }
            }
        }



        //protected void btnExportCS_Click1(object sender, EventArgs e)
        //{
        //    DateTime time = DateTime.Now;
        //    Response.Clear();
        //    Response.AddHeader("content-disposition", "attachment; filename = " + time + ".xls");
        //    Response.ContentType = "application/ms-excel";
        //    Response.ContentEncoding = System.Text.Encoding.Unicode;
        //    Response.BinaryWrite(System.Text.Encoding.Unicode.GetPreamble());

        //    StringWriter stringWriter = new StringWriter();
        //    HtmlTextWriter htmlWriter = new HtmlTextWriter(stringWriter);

        //    GridView1.RenderControl(htmlWriter);

        //    string s = stringWriter.ToString();
        //    Response.Write(s);
        //    Response.End();
        //}
        //protected void btnFillValues_Click(object sender, EventArgs e)
        //{

        //}

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

        public override void VerifyRenderingInServerForm(Control control)
        {
            return;
        }

        //public string insert8tablesSL(string finaltbl, string tbljoin, string tbl, string device, string device_condition, string device_condition_and)
        //{
        //    string tables8SL = @"  SELECT  distinct  '" + device + "'+[Search term] as Campaign@Searchterm"
        //                                             + @",[Search term]
        //                                           ,sum([Cost]) as 'cost'
        //                                          ,sum([Clicks]) as 'clicks'
        //                                          ,sum([Demo CRM]) as 'Demo CRM'
        //                                          ,sum([FTD Approved]) as 'FTD Approved'
        //                                          ,sum([All conv. value]) as 'All conv value'	
        //                                          into  " + tbl + " from  [Marketing].[dbo].[CN_PivotRaw$]" +
        //                                              " where [Device] " + device_condition + " and " + device_condition_and + ""
        //                                              + @"  group by [Search term];"

        //                                            //+ @"--update " + tbl + " set [Search term] ='#' + replace([Search term],' ','#')+'#';"
        //                                            + @" select distinct [Search term] into  " + tbljoin + "  from  " + tbl + " where [FTD Approved]>0;"



        //                                            + @" SELECT '" + device + "' as campaign@searchterm,b.[Search term],"
        //                                           + @" sum(a.[cost]) as [Cost],sum(a.[clicks]) as [Clicks],sum([Demo CRM]) as 'Demo CRM' ,sum([FTD Approved]) as 'FTD Approved',sum(a.[All conv value]) as [All conv value]
        //                                        into  " + finaltbl + " FROM " + tbl + " a " +
        //                                           "left  join " + tbljoin + " b on a.[Search term]  like N'%'+ b.[Search term]+ '%' group by b.[Search term];" +



        //                                           @"ALTER TABLE  " + finaltbl + " ADD "
        //                                          + @"[L2FTD]  float,[Conv rate]  float,[CPL]  float,[CPC]  float,[ROI]  float,[Target CPA by desired CPA]  float,[Target CPA by desired CPA*2]  float,[Target CPA by ROI]  float,[Target CPC by desired CPA]  float,
        //                                       [Target CPC by desired CPA*2]  float,[Target CPC by ROI]  float,[Max CPA]  float,[Max CPC]  float,[Device adjustment CPA] float,[Device adjustment CPC] float, [Target CPA] float,[Target CPC] float;" +


        //                                            @" update " + finaltbl + " set [L2FTD] = [FTD Approved] /[Demo CRM],[Conv rate]= [Demo CRM]/[Clicks]"
        //                                          + @",[CPL]=[Cost]/[Demo CRM]
        //                                     ,[CPC]=[Cost]/[Clicks]
        //                                     ,[ROI]=[All conv value]/[Cost];
        //                                       // " + @" update   " + finaltbl + "  set [Target CPA by desired CPA]= CASE when[FTD Approved]>1 THEN(" + txtDesiredFtdBig1.Text + " *[L2FTD]) else (" + txtDesiredFtdBig1.Text + "*[L2FTD]) END,"
        //                                             // + @"[Target CPA by desired CPA*2] = (CASE when[FTD Approved] > 1 THEN(" + txtDesiredFtdBig1.Text + " *[L2FTD]) else (" + txtDesiredFtdBig1.Text + " *[L2FTD]) END)*2" + ","
        //                                             // + @"[Target CPA by ROI]=([ROI]*[CPL])/2 ,
        //                                             // [Target CPC by desired CPA] = CASE when[FTD Approved]>1 THEN(" + txtDesiredFtdBig1.Text + "*[L2FTD]*[Conv rate]) else (" + txtDesiredFtdBig1.Text + "*[L2FTD]*[Conv rate]) END ,"
        //                                             //+ @"[Target CPC by desired CPA*2] = (CASE when  [FTD Approved]>1 THEN(" + txtDesiredFtdBig1.Text + "*[L2FTD]*[Conv rate]) else (" + txtDesiredFtdBig1.Text + "*[L2FTD]*[Conv rate]) END)*2 ,"
        //                                             // + @"[Target CPC by ROI]=([ROI]*[CPC])/2;"

        //                                             + @"update  " + finaltbl + "  set [Max CPA] = IIF(IIF([Target CPA by desired CPA] <[Target CPA by ROI],[Target CPA by ROI],[Target CPA by desired CPA]) <[Target CPA by desired CPA*2],IIF([Target CPA by desired CPA] <[Target CPA by ROI],[Target CPA by ROI],[Target CPA by desired CPA]),[Target CPA by desired CPA*2])," +
        //                                             " [Max CPC] = IIF(IIF([Target CPC by desired CPA]<[Target CPC by ROI], [Target CPC by ROI], [Target CPC by desired CPA])<[Target CPC by desired CPA*2],IIF([Target CPC by desired CPA]<[Target CPC by ROI], [Target CPC by ROI], [Target CPC by desired CPA]),[Target CPC by desired CPA*2])," +
        //                                             " [Device adjustment CPA] = (IIF(IIF([Target CPA by desired CPA]<[Target CPA by ROI],[Target CPA by ROI],[Target CPA by desired CPA])<[Target CPA by desired CPA*2],IIF([Target CPA by desired CPA]<[Target CPA by ROI], [Target CPA by ROI], [Target CPA by desired CPA]),[Target CPA by desired CPA*2]))," +
        //                                             " [Device Adjustment CPC] = (IIF(IIF([Target CPC by desired CPA]<[Target CPC by ROI],[Target CPC by ROI],[Target CPC by desired CPA])<[Target CPC by desired CPA*2],IIF([Target CPC by desired CPA]<[Target CPC by ROI], [Target CPC by ROI], [Target CPC by desired CPA]),[Target CPC by desired CPA*2]));" +
        //                                        //           " update    " + finaltbl + " set      [Target CPA] = IIF([Device adjustment CPA]>=" + txtMaxCpaBidBig1.Text + "," + txtMaxCpaBidBig1.Text + ", [Device adjustment CPA])" +
        //                                        //          @",[Target CPC] = IIF([Device adjustment CPC]>=" + txtMaxCpcBidBig1.Text + "," + txtMaxCpcBidBig1.Text + ",[Device adjustment CPC]);" +


        //                                        @" insert into [Marketing].[dbo].[phrases8tables_CN$] select * from  " + finaltbl + ";";



        //    return tables8SL;

        //  }
        public string insert8tablesSL(string final_tables, string tbljoin, string tbl, string device, string device_condition, string device_condition_and, string desired_cap_big, string cpa_bidbig1, string cpc_big1)
        {
            string tables8_CN = @"  SELECT  distinct  '" + device + "'+[Search term] as Campaign@Searchterm"
                                                     + @",[Search term]
                                                       ,sum([Cost]) as 'cost'
                                                      ,sum([Clicks]) as 'clicks'
                                                      ,sum([Demo CRM]) as 'Demo CRM'
                                                      ,sum([FTD Approved]) as 'FTD Approved'
                                                      ,sum([All conv. value]) as 'All conv value'	
                                                      into  " + tbl + " from  [Marketing].[dbo].[CN_PivotRaw$]" +
                                                      " where [Device] " + device_condition + " and " + device_condition_and + ""
                                                      + @"  group by [Search term];"


                                                      + @" select distinct [Search term] into  " + tbljoin + "  from  " + tbl + " where[FTD Approved]>0;"


                                                      + @" SELECT '" + device + "' as campaign@searchterm,b.[Search term],"
                                                       + @" sum(a.[cost]) as [Cost],sum(a.[clicks]) as [Clicks],sum([Demo CRM]) as 'Demo CRM' ,sum([FTD Approved]) as 'FTD Approved',sum(a.[All conv value]) as [All conv value]
                                                    into  " + final_tables + " FROM " + tbl + " a " +
                                                       "left  join " + tbljoin + " b on a.[Search term]  like N'%'+ b.[Search term]+ '%' group by b.[Search term];" +


                                                 @"ALTER TABLE  " + final_tables + " ADD "
                                                  + @"[L2FTD]  float,[Conv rate]  float,[CPL]  float,[CPC]  float,[ROI]  float,[Target CPA by desired CPA]  float,[Target CPA by desired CPA*2]  float,[Target CPA by ROI]  float,[Target CPC by desired CPA]  float,
                                                   [Target CPC by desired CPA*2]  float,[Target CPC by ROI]  float,[Max CPA]  float,[Max CPC]  float,[Device adjustment CPA] float,[Device adjustment CPC] float, [Target CPA] float,[Target CPC] float;"


                                                       + @" update " + final_tables + " set [L2FTD] =  [FTD Approved] /[Demo CRM],[Conv rate]= [Demo CRM]/[Clicks]"
                                                      + @",[CPL]=[Cost]/[Demo CRM]
                                                     ,[CPC]=[Cost]/[Clicks]
                                                     ,[ROI]=[All conv value]/[Cost];
                                                        " + @" update   " + final_tables + "  set [Target CPA by desired CPA]= CASE when [FTD Approved]>1 THEN(" + desired_cap_big + " *[L2FTD]) else (" + desired_cap_big + "*[L2FTD]) END,"
                                                        + @"[Target CPA by desired CPA*2] = (CASE when [FTD Approved] > 1 THEN(" + desired_cap_big + " *[L2FTD]) else (" + desired_cap_big + " *[L2FTD]) END)*2" + ","
                                                        + @"[Target CPA by ROI]=([ROI]*[CPL])/2 ,
                                                        [Target CPC by desired CPA] = CASE when [FTD Approved]>1 THEN(" + desired_cap_big + "*[L2FTD]*[Conv rate]) else (" + desired_cap_big + "*[L2FTD]*[Conv rate]) END ,"
                                                       + @"[Target CPC by desired CPA*2] = (CASE when   [FTD Approved]>1 THEN(" + desired_cap_big + "*[L2FTD]*[Conv rate]) else (" + desired_cap_big + "*[L2FTD]*[Conv rate]) END)*2 ,"
                                                        + @"[Target CPC by ROI]=([ROI]*[CPC])/2;"

                                                     + @"update  " + final_tables + "  set [Max CPA] = IIF(IIF([Target CPA by desired CPA] <[Target CPA by ROI],[Target CPA by ROI],[Target CPA by desired CPA]) <[Target CPA by desired CPA*2],IIF([Target CPA by desired CPA] <[Target CPA by ROI],[Target CPA by ROI],[Target CPA by desired CPA]),[Target CPA by desired CPA*2])," +
                                                     " [Max CPC] = IIF(IIF([Target CPC by desired CPA]<[Target CPC by ROI], [Target CPC by ROI], [Target CPC by desired CPA])<[Target CPC by desired CPA*2],IIF([Target CPC by desired CPA]<[Target CPC by ROI], [Target CPC by ROI], [Target CPC by desired CPA]),[Target CPC by desired CPA*2])," +
                                                     " [Device adjustment CPA] = (IIF(IIF([Target CPA by desired CPA]<[Target CPA by ROI],[Target CPA by ROI],[Target CPA by desired CPA])<[Target CPA by desired CPA*2],IIF([Target CPA by desired CPA]<[Target CPA by ROI], [Target CPA by ROI], [Target CPA by desired CPA]),[Target CPA by desired CPA*2]))," +
                                                     " [Device Adjustment CPC] = (IIF(IIF([Target CPC by desired CPA]<[Target CPC by ROI],[Target CPC by ROI],[Target CPC by desired CPA])<[Target CPC by desired CPA*2],IIF([Target CPC by desired CPA]<[Target CPC by ROI], [Target CPC by ROI], [Target CPC by desired CPA]),[Target CPC by desired CPA*2]));" +
                                                 " update    " + final_tables + " set      [Target CPA] = IIF([Device adjustment CPA]>=" + cpa_bidbig1 + "," + cpa_bidbig1 + ", [Device adjustment CPA])" +
                                                @",[Target CPC] = IIF([Device adjustment CPC]>=" + cpc_big1 + "," + cpc_big1 + ",[Device adjustment CPC]);" +


                                                @" insert into [Marketing].[dbo].[phrases8tables_CN$] select * from  " + final_tables + ";";

            return tables8_CN;

        }

        //   public string insert8tablesSL_WW(string finaltbl, string tbljoin, string tbl, string device, string device_condition, string device_condition_and)
        //{
        //    string tables8SL_WW = @"  SELECT  distinct  '" + device + "'+[Search term] as Campaign@Searchterm"
        //                                             + @",[Search term]
        //                                              ,sum([Cost]) as 'cost'
        //                                              ,sum([Clicks]) as 'clicks'
        //                                              ,sum([Demo CRM]) as 'Demo CRM'
        //                                              ,sum([FTD Approved]) as 'FTD Approved'
        //                                              ,sum([All conv. value]) as 'All conv value'
        //                                              into  " + tbl + " from   [Marketing].[dbo].[CN_PivotRaw$]" +
        //                                              " where [Device] " + device_condition + " and " + device_condition_and + ""
        //                                              + @"  group by [Search term];"

        //                                            //+ @"--update " + tbl + " set[Search term] ='#' + replace([Search term],' ','#')+'#';"
        //                                            + @" select distinct [Search term] into  " + tbljoin + "  from  " + tbl + " where [FTD Approved]>0;"



        //                                            + @" SELECT '" + device + "' as campaign@searchterm,b.[Search term],"
        //                                           + @" sum(a.[cost]) as [Cost],sum(a.[clicks]) as [Clicks],sum(a.[Demo CRM]) as [Demo CRM] ,sum(a.[FTD Approved]) as [FTD Approved],sum(a.[All conv value]) as [All conv value]
        //                                            into  " + finaltbl + " FROM " + tbl + " a " +
        //                                           "left  join " + tbljoin + " b on a.[Search term]  like N'%'+ b.[Search term]+ '%' group by b.[Search term];" +



        //                                           @"ALTER TABLE  " + finaltbl + " ADD "
        //                                          + @"[L2FTD]  float,[Conv rate]  float,[CPL]  float,[CPC]  float,[ROI]  float,[Target CPA by desired CPA]  float,[Target CPA by desired CPA*2]  float,[Target CPA by ROI]  float,[Target CPC by desired CPA]  float,
        //                                           [Target CPC by desired CPA*2]  float,[Target CPC by ROI]  float,[Max CPA]  float,[Max CPC]  float,[Device adjustment CPA] float,[Device adjustment CPC] float, [Target CPA] float,[Target CPC] float;"


        //                                           + @" update " + finaltbl + " set [L2FTD] = [FTD Approved] /[Demo CRM],[Conv rate]= [Demo CRM]/[Clicks]"
        //                                          + @",[CPL]=[Cost]/[Demo CRM]
        //                                         ,[CPC]=[Cost]/[Clicks]
        //                                         ,[ROI]=[All conv value]/[Cost];
        //                                            " + @" update   " + finaltbl + "  set [Target CPA by desired CPA]= CASE when[FTD Approved]>1 THEN(" + txtDesiredFtdBigSLWW1.Text + " *[L2FTD]) else (" + txtDesiredFtdBigSLWW1.Text + "*[L2FTD]) END,"
        //                                            + @"[Target CPA by desired CPA*2] = (CASE when[FTD Approved] > 1 THEN(" + txtDesiredFtdBigSLWW1.Text + " *[L2FTD]) else (" + txtDesiredFtdBigSLWW1.Text + " *[L2FTD]) END)*2" + ","
        //                                            + @"[Target CPA by ROI]=([ROI]*[CPL])/2 ,
        //                                            [Target CPC by desired CPA] = CASE when[FTD Approved]>1 THEN(" + txtDesiredFtdBigSLWW1.Text + "*[L2FTD]*[Conv rate]) else (" + txtDesiredFtdBigSLWW1.Text + "*[L2FTD]*[Conv rate]) END ,"
        //                                           + @"[Target CPC by desired CPA*2] = (CASE when  [FTD Approved]>1 THEN(" + txtDesiredFtdBigSLWW1.Text + "*[L2FTD]*[Conv rate]) else (" + txtDesiredFtdBigSLWW1.Text + "*[L2FTD]*[Conv rate]) END)*2 ,"
        //                                            + @"[Target CPC by ROI]=([ROI]*[CPC])/2;"

        //                                             + @"update  " + finaltbl + "  set [Max CPA] = IIF(IIF([Target CPA by desired CPA] <[Target CPA by ROI],[Target CPA by ROI],[Target CPA by desired CPA]) <[Target CPA by desired CPA*2],IIF([Target CPA by desired CPA] <[Target CPA by ROI],[Target CPA by ROI],[Target CPA by desired CPA]),[Target CPA by desired CPA*2])," +
        //                                             " [Max CPC] = IIF(IIF([Target CPC by desired CPA]<[Target CPC by ROI], [Target CPC by ROI], [Target CPC by desired CPA])<[Target CPC by desired CPA*2],IIF([Target CPC by desired CPA]<[Target CPC by ROI], [Target CPC by ROI], [Target CPC by desired CPA]),[Target CPC by desired CPA*2])," +
        //                                             " [Device adjustment CPA] = (IIF(IIF([Target CPA by desired CPA]<[Target CPA by ROI],[Target CPA by ROI],[Target CPA by desired CPA])<[Target CPA by desired CPA*2],IIF([Target CPA by desired CPA]<[Target CPA by ROI], [Target CPA by ROI], [Target CPA by desired CPA]),[Target CPA by desired CPA*2]))," +
        //                                             " [Device Adjustment CPC] = (IIF(IIF([Target CPC by desired CPA]<[Target CPC by ROI],[Target CPC by ROI],[Target CPC by desired CPA])<[Target CPC by desired CPA*2],IIF([Target CPC by desired CPA]<[Target CPC by ROI], [Target CPC by ROI], [Target CPC by desired CPA]),[Target CPC by desired CPA*2]));" +
        //                                             " update    " + finaltbl + " set      [Target CPA] = IIF([Device adjustment CPA]>=" + txtMaxCpaBidBigSLWW1.Text + "," + txtMaxCpaBidBigSLWW1.Text + ", [Device adjustment CPA])" +

        //                                      //       @",[Target CPC] = IIF([Device adjustment CPC]>=" + txtMaxCpcBidBigSLWW1.Text + "," + txtMaxCpcBidBigSLWW1.Text + ",[Device adjustment CPC]);" +
        //                                              @" insert into [Marketing].[dbo].[phrases8tables_CN$]  select* from  " + finaltbl + ";";

        //    return tables8SL_WW;

        //}

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
            //  if (ChckboxWithFixing.Checked)
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

        protected void chkMY_CheckedChanged(object sender, EventArgs e)
        {
            string queryMY_my = @" UPDATE [Marketing].[dbo].[SaveDataTblDevices$]
                           SET [CampDesktop]='" + txtCampaignNameSLWW1Ages.Text + "'"
                        + @",[CampMobile]='" + txtCampaignNameSLWW2Ages.Text + "'"
                        + @",[CampTablet] = '" + txtCampaignNameSLWW3Ages.Text + "'"
                        + @",[CampMobTab]='" + txtCampaignNameSLWW4Ages.Text + "'"

                        + @",[DesCpaBig1Desktop]='" + txtDesiredCpaFtdBigDesktop1Ages.Text + "'"
                        + @",[DesCpaBig1Mobile]='" + txtDesiredCpaFtdBigMobile1Ages.Text + "'"
                        + @",[DesCpaBig1Tablet]='" + txtDesiredCpaFtdBigTablet1Ages.Text + "'"
                        + @",[DesCpaBig1MobTab]='" + txtDesiredCpaFtdBigMobTab1Ages.Text + "'"

                        + @",[DesCpaEqDesktop]='" + txtDesiredCpaFtdEqDesktop1Ages.Text + "'"
                        + @",[DesCpaEqMobile]='" + txtDesiredCpaFtdEqMobile1Ages.Text + "'"
                        + @",[DesCpaEqTablet]='" + txtDesiredCpaFtdEqTablet1Ages.Text + "'"
                        + @",[DesCpaEqMobTab]='" + txtDesiredCpaFtdEqMobTab1Ages.Text + "'"


                        + @",[DeviceAdjDesktop]='" + txtDeviceAdjDesktop1Ages.Text + "'"
                        + @",[DeviceAdjMobile]='" + txtDeviceAdjMobile1Ages.Text + "'"
                        + @",[DeviceAdjTablet]='" + txtDeviceAdjTablet1Ages.Text + "'"
                        + @",[DeviceAdjMobTab]='" + txtDeviceAdjMobTab1Ages.Text + "'"


                        + @",[MaxCpaBig1Desktop]='" + txtMaxCpaBidFtdBigDesktop1Ages.Text + "'"
                        + @",[MaxCpaBig1Mobile]='" + txtMaxCpaBidFtdBigMobile1Ages.Text + "'"
                        + @",[MaxCpaBig1Tablet]='" + txtMaxCpaBidFtdBigTablet1Ages.Text + "'"
                        + @",[MaxCpaBig1MobTab]='" + txtMaxCpaBidFtdBigTablet1Ages.Text + "'"

                        + @",[MaxCpaEqDesktop]='" + txtMaxCpaBidFtdEqDesktop1Ages.Text + "'"
                        + @",[MaxCpaEqMobile]='" + txtMaxCpaBidFtdEqMobile1Ages.Text + "'"
                        + @",[MaxCpaEqTablet]='" + txtMaxCpaBidFtdEqTablet1Ages.Text + "'"
                        + @",[MaxCpaEqMobTab]='" + txtMaxCpaBidFtdEqMobTab1Ages.Text + "'"


                        + @",[MaxCpcBig1Desktop]='" + txtMaxCpcBidFtdBigDesktop1Ages.Text + "'"
                        + @",[MaxCpcBig1Mobile]='" + txtMaxCpcBidFtdBigMobile1Ages.Text + "'"
                        + @",[MaxCpcBig1Tablet]='" + txtMaxCpcBidFtdBigTablet1Ages.Text + "'"
                        + @",[MaxCpcBig1MobTab]='" + txtMaxCpcBidFtdBigMobTab1Ages.Text + "'"

                        + @",[MaxCpcEqDesktop]='" + txtMaxCpcBidFtdEqDesktop1Ages.Text + "'"
                        + @",[MaxCpcEqMobile]='" + txtMaxCpcBidFtdEqMobile1Ages.Text + "'"
                        + @",[MaxCpcEqTablet]='" + txtMaxCpcBidFtdEqTablet1Ages.Text + "'"
                        + @",[MaxCpcEqMobTab]='" + txtMaxCpcBidFtdEqMobTab1Ages.Text + "' WHERE[ID] = 'MY_Dep_MY';";


            SqlCommand queryMYMY = new SqlCommand(queryMY_my, con);



            queryMYMY.ExecuteNonQuery();
            queryMYMY.ExecuteScalar();




            string queryMYMYAges = @" UPDATE [Marketing].[dbo].[SaveDataTblDevices$]
                           SET [CampDesktop]='" + txtCampaignNameSLWW1Ages.Text + "'"
                    + @",[CampMobile]='" + txtCampaignNameSLWW2Ages.Text + "'"
                    + @",[CampTablet] = '" + txtCampaignNameSLWW3Ages.Text + "'"
                    + @",[CampMobTab]='" + txtCampaignNameSLWW4Ages.Text + "'"

                    + @",[DesCpaBig1Desktop]='" + txtDesiredCpaFtdBigDesktop1Ages.Text + "'"
                    + @",[DesCpaBig1Mobile]='" + txtDesiredCpaFtdBigMobile1Ages.Text + "'"
                    + @",[DesCpaBig1Tablet]='" + txtDesiredCpaFtdBigTablet1Ages.Text + "'"
                    + @",[DesCpaBig1MobTab]='" + txtDesiredCpaFtdBigMobTab1Ages.Text + "'"

                    + @",[DesCpaEqDesktop]='" + txtDesiredCpaFtdEqDesktop1Ages.Text + "'"
                    + @",[DesCpaEqMobile]='" + txtDesiredCpaFtdEqMobile1Ages.Text + "'"
                    + @",[DesCpaEqTablet]='" + txtDesiredCpaFtdEqTablet1Ages.Text + "'"
                    + @",[DesCpaEqMobTab]='" + txtDesiredCpaFtdEqMobTab1Ages.Text + "'"


                    + @",[DeviceAdjDesktop]='" + txtDeviceAdjDesktop1Ages.Text + "'"
                    + @",[DeviceAdjMobile]='" + txtDeviceAdjMobile1Ages.Text + "'"
                    + @",[DeviceAdjTablet]='" + txtDeviceAdjTablet1Ages.Text + "'"
                    + @",[DeviceAdjMobTab]='" + txtDeviceAdjMobTab1Ages.Text + "'"


                    + @",[MaxCpaBig1Desktop]='" + txtMaxCpaBidFtdBigDesktop1Ages.Text + "'"
                    + @",[MaxCpaBig1Mobile]='" + txtMaxCpaBidFtdBigMobile1Ages.Text + "'"
                    + @",[MaxCpaBig1Tablet]='" + txtMaxCpaBidFtdBigTablet1Ages.Text + "'"
                    + @",[MaxCpaBig1MobTab]='" + txtMaxCpaBidFtdBigTablet1Ages.Text + "'"

                    + @",[MaxCpaEqDesktop]='" + txtMaxCpaBidFtdEqDesktop1Ages.Text + "'"
                    + @",[MaxCpaEqMobile]='" + txtMaxCpaBidFtdEqMobile1Ages.Text + "'"
                    + @",[MaxCpaEqTablet]='" + txtMaxCpaBidFtdEqTablet1Ages.Text + "'"
                    + @",[MaxCpaEqMobTab]='" + txtMaxCpaBidFtdEqMobTab1Ages.Text + "'"


                    + @",[MaxCpcBig1Desktop]='" + txtMaxCpcBidFtdBigDesktop1Ages.Text + "'"
                    + @",[MaxCpcBig1Mobile]='" + txtMaxCpcBidFtdBigMobile1Ages.Text + "'"
                    + @",[MaxCpcBig1Tablet]='" + txtMaxCpcBidFtdBigTablet1Ages.Text + "'"
                    + @",[MaxCpcBig1MobTab]='" + txtMaxCpcBidFtdBigMobTab1Ages.Text + "'"

                    + @",[MaxCpcEqDesktop]='" + txtMaxCpcBidFtdEqDesktop1Ages.Text + "'"
                    + @",[MaxCpcEqMobile]='" + txtMaxCpcBidFtdEqMobile1Ages.Text + "'"
                    + @",[MaxCpcEqTablet]='" + txtMaxCpcBidFtdEqTablet1Ages.Text + "'"
                    + @",[MaxCpcEqMobTab]='" + txtMaxCpcBidFtdEqMobTab1Ages.Text + "' WHERE[ID] = 'MY_Dep_MY_Ages';";


            SqlCommand queryMYMyAges = new SqlCommand(queryMYMYAges, con);



            queryMYMyAges.ExecuteNonQuery();
            queryMYMyAges.ExecuteScalar();
        }

      }
    }

