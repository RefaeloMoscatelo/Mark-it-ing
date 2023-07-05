using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.Collections;
using System.Text;

namespace DemoCrmH
{
    public partial class DemoCrmH : System.Web.UI.Page
    {
        protected bool IsLower(string value)
        {
            // Consider string to be lowercase if it has no uppercase letters.
            for (int i = 0; i < value.Length; i++)
            {
                if (char.IsUpper(value[i])|| value.Contains("test"))
                {
                    return false;
                }
            }
            return true;
        }

        

        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void BtnSearch_Click(object sender, EventArgs e)
        {
            string endMonthdate;
            endMonthdate = txtEndDate.Text;
            DateTime endMonthDate = DateTime.Parse(endMonthdate);
            string endDate = endMonthDate.AddDays(1).ToString("yyyy-MM-dd");

            string maincon = ConfigurationManager.ConnectionStrings["crm_LC_"].ConnectionString;
            string brand_type = "";
            string brand = ddlSite1.SelectedValue;
            string db_brand = "df-0863-E811-80D8-00dfgdfg5056970BD3','dfg-7B33-E911-80DA-dfgdfg','2D826510-00B3-E511-80C8-005056A42D66','34534-A542-EA11-80E8-005056970BD3','34534545-3453-EA11-80E8-34534543','A9273C91-F352-EA11-80E1-005056972925";
            string brandCidNumber = "";
            string cd = "cd5";

            switch (brand)
            {
                case "fort":
                    Lab2.Text = "frade";
                    brand_type = "12345-31";
                    brandCidNumber = "31";
                    break;
                case "gcmasia":
                    db_brand = "55666-7953-E911-80E0-8787878','3453-0863-3453-80D8-23452','54324-00B3-34535-80C8-34535";
                    Lab2.Text = "pcm";
                    brand_type = "7334285-27";
                    brandCidNumber = "27";
                    cd = "cd2";
                    break;
                case "kapitalRS":
                    db_brand = "ter-0863-E811-erte-005056970BD3','ertert-D8D7-E611-80E0-23333','ewr-95DF-EA11-80EF-rrrr";
                    Lab2.Text = "kapi";
                    brand_type = "4444-11";
                    brandCidNumber = "11";
                    cd = "cd2";
                    break;
                case "gcmforex":
                    maincon = ConfigurationManager.ConnectionStrings["crm_"].ConnectionString;
                    db_brand = "434-FCB2-E511-80C6-005056A44066";
                    Lab2.Text = "gforex";
                    brand_type = "34534-26";
                    brandCidNumber = "555";
                    cd = "cd3";
                    break;
                default:

                    break;
            }
            string sqlquery = "";
            // string maincon = ConfigurationManager.ConnectionStrings["crm_LC_MarketingConnectionString"].ConnectionString;
            // maincon = ConfigurationManager.ConnectionStrings["crm_GCM_MarketingConnectionString"].ConnectionString;
            SqlConnection sqlcon = new SqlConnection(maincon);
            //string sqlquery = "select [AccountBase].[lv_cid], [AccountBase].[lv_TLID],[AccountBase].[EMailAddress1], [lv_siteBase].[lv_name] from [AccountBase] join [lv_siteBase] on [AccountBase].[lv_siteId]=[lv_siteBase].[lv_siteId] where [lv_name] in('" + db_brand + "') and  [AccountBase].[lv_TLID] is not null and [lv_contains_test] = 0 and [AccountBase].[lv_created_in_office] = 0 and [CreatedOn] BETWEEN '" + txtStartDate.Text + "'  AND '" + txtEndDate.Text + " 00:00' order by [AccountBase].[lv_cid]";
            if (brand=="gcmforex")
            {
                 sqlquery= @" select   [lv_cid], [lv_TLID] from [dbo].[AccountBase]   
                                     where[lv_siteId] in ('" + db_brand + "')        "+
                                     " and [CreatedOn] BETWEEN '" + txtStartDate.Text + "'  AND '" + endDate + "'"+
                                  " and ((lv_accountstatus != '4' and lv_accountstatus != '4') or ([lv_viop_account_status] != '772400000' and [lv_viop_account_status] is not null))" +
                                      "and lv_contains_test = 0  and lv_created_in_office = 0  and lv_tlid is not null order by [lv_cid]";
            }
            else
            {
                 sqlquery = @"select  [lv_cid], [lv_TLID] from [dbo].[AccountBase] a
                                join[dbo].[SystemUserBase] b on b.[SystemUserId] = a.ownerid
                                where[lv_siteId] in ('" + db_brand + "')" +
                              "and[lv_TLID] is not null and[lv_contains_test] = 0 and[lv_created_in_office] = 0" +
                              "and lv_accountstatus not in (4, 5, 6)" +
                              "and b.[FullName] not like '%owner%'" +
                              "and a.[CreatedOn] BETWEEN '" + txtStartDate.Text + "'  AND '" + endDate + " 00:00' order by [lv_cid]";
            }
          
            //if (CheckBox1.Checked)
            //{
            //    sqlquery = "select  a.[lv_cid], a.[lv_TLID], s.[lv_name] from[dbo].[AccountBase] a join[dbo].[lv_siteBase] s on a.[lv_siteid] = s.[lv_siteid] where a.[lv_siteId] in('" + db_brand + "')  and a.[lv_TLID] is not null and a.[lv_contains_test] = 0 and a.[lv_created_in_office] = 0  and a.[CreatedOn] >= dateadd(day, datediff(day, 0, getdate()) - 1, 0)  order by a.[lv_cid]";
            // }
            SqlCommand sqlcom = new SqlCommand(sqlquery, sqlcon);
            sqlcon.Open();
            SqlDataAdapter sda = new SqlDataAdapter(sqlcom);
            DataTable dt = new DataTable();
            sda.Fill(dt);
            SqlDataReader dr = sqlcom.ExecuteReader();
            if (dr.Read())
            {
                GridView1.DataSource = dt;
                GridView1.DataBind();
            }
            else
            {
                // Lab1.Text = GridView1.Rows.Count.ToString();
            }
            int count = GridView1.Rows.Count;
            Lab1.Text = count.ToString() + " Rows";

            ArrayList arrayLinks = new ArrayList();


            List<String> cidList = new List<string>();
            ArrayList arraycid = new ArrayList();
            GridView1.Rows.OfType<GridViewRow>().ToList().ForEach(db => cidList.Add(db.Cells[0].Text));
            List<String> commaList = new List<string>();
            ArrayList arrayGclid = new ArrayList();
            GridView1.Rows.OfType<GridViewRow>().ToList().ForEach(a => arrayGclid.Add(a.Cells[1].Text));
            commaList = cidList.Where(x => x.Contains(",")).ToList();
            if (commaList.Count>0)
            {
              
                LabCommaError1.Text += "<h2>CID Contain Comma List</h2>";
                LabCommaError1.Text += "<ul id='CommaLetters'>";
                for (int i = 0; i < commaList.Count(); i++){
                    
                    LabCommaError1.Text += "<li>"+commaList[i].ToString()+ "</li>";
                    
                }
                LabCommaError1.Text += "</ul>";
            }
            DateTime cid_time = DateTime.Now;
            string q = cid_time.ToString("yyyyMMdd.hhmmss");
          
            List<String> smallLetterList = new List<string>();
           
            LabSmallLetters.Text += "<ul id='smallLetters'>";
               
                for (int i = 0; i < arrayGclid.Count; i++)
                {
                    if (IsLower(arrayGclid[i].ToString())|| arrayGclid[i].ToString().Contains("test") || arrayGclid[i].ToString().Contains(","))
                    {
                        LabSmallLetters.Text += "<li>" + arrayGclid[i].ToString() + "</li>";
                    arrayGclid.RemoveAt(i);
                    cidList.RemoveAt(i);
                    }
                }
                LabSmallLetters.Text += "</ul>";
            
           
          


            for (int i = 0; i < cidList.Count; i++)
            {
                if (cidList[i].ToString().StartsWith("GA1"))
                {
                    cidList[i]=cidList[i].ToString().Substring(6);
                }

                if (cidList[i].ToString() == "&nbsp;")
                {
                    string d = i+ brandCidNumber + q ;
                    cidList[i] = d ;
                }
                arrayLinks.Add(@"<a href='https://www.google-analytics.com/collect?v=1&t=event&tid=UA-" + brand_type + "&cid=" + cidList[i] + "&ni=1&ec=TP&ea=Demo&el=Conversion&gclid=" + arrayGclid[i] + "&" + cd + "=" + arrayGclid[i] + "' target='_blank' class='ck'>" + i + "</a>");
            //    arrayLinks.Add(@"https://www.google-analytics.com/collect?v=1&t=event&tid=UA-" + brand_type + "&cid=" + cidList[i] + "&ni=1&ec=TP&ea=Demo&el=Conversion&gclid=" + arrayGclid[i] + "&" + cd + "=" + arrayGclid[i]+"");
                // GridView1.Rows.OfType<GridViewRow>().ToList().ForEach(c >= arrayLinks.Add(c.Cells[2].Text));
            }
            string temporary = string.Empty;
            //foreach (String var in arrayLinks)
            //{
            //    temporary = temporary + var.ToString() + "<br />";
            //}
           

            StringBuilder s = new StringBuilder();

            for (int i = 0; i < arrayLinks.Count; i++)
            {
                s.Append(arrayLinks[i] + "<br />");
            }


            //foreach (string x in arrayLinks)
            //{
            //    s.Append(x + "<br />");
            //}

            Label1.Text = s.ToString();

            sqlcon.Close();
        }

        protected void DDlSite_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            LabSmallLetters.Text = "";
            LabCommaError1.Text = "";
        }
        
        protected void Unnamed_SelectedIndexChanged(object sender, EventArgs e)
        {
            LabSmallLetters.Text = "";
            LabCommaError1.Text = "";
        }

        protected void Unnamed_SelectedIndexChanged1(object sender, EventArgs e)
        {
            LabSmallLetters.Text = "";
            LabCommaError1.Text = "";
        }
    }


}