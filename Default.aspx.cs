using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using System.Data;
using iTextSharp.text;
using System.IO;
using iTextSharp.text.html.simpleparser;
using iTextSharp.text.pdf;
using iTextSharp.text.html;
using System.Data.SqlClient;
using System.Configuration;

public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e) 
    {
        if (!IsPostBack)
        {

            BaseFont fontthai_Bold = BaseFont.CreateFont(Server.MapPath("/Fonts/THSarabun Bold.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            Font bold = new Font(fontthai_Bold, 16);

            BaseFont fontthai = BaseFont.CreateFont(Server.MapPath("/Fonts/THSarabun.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            Font normal = new Font(fontthai, 16);
            string ConString = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;

            SqlConnection con = new SqlConnection(ConString);

            SqlCommand cmd = new SqlCommand();

            cmd.CommandText = "SELECT EN_borrow, Date_borrow, Name, Activity, place, Get_Equipment, Return_Equipment FROM borroww";

            cmd.Connection = con;

            con.Open();

            DataTable dt = new DataTable();

            dt.Load(cmd.ExecuteReader());

            if (dt.Rows.Count > 0)
            {

                //Bind Datatable to Labels

                lblEmployeeId.Text = dt.Rows[0]["EN_borrow"].ToString();

                lblFirstName.Text = dt.Rows[0]["Date_borrow"].ToString();

                lblLastName.Text = dt.Rows[0]["Name"].ToString();

                lblCity.Text = dt.Rows[0]["Activity"].ToString();

                lblState.Text = dt.Rows[0]["place"].ToString();

                lblPostalCode.Text = dt.Rows[0]["Get_Equipment"].ToString();

                lblCountry.Text = dt.Rows[0]["Return_Equipment"].ToString();

            }

            con.Close();

        }


    }

    protected void btnExport_Click(object sender, EventArgs e)
    {
        Response.ContentType = "application/pdf";

        Response.AddHeader("content-disposition", "attachment;filename=Panel.pdf");

        Response.Cache.SetCacheability(HttpCacheability.NoCache);



        StringWriter stringWriter = new StringWriter();

        HtmlTextWriter htmlTextWriter = new HtmlTextWriter(stringWriter);

        employeelistDiv.RenderControl(htmlTextWriter);



        StringReader stringReader = new StringReader(stringWriter.ToString());

        Document Doc = new Document(PageSize.A4, 10f, 10f, 100f, 0f);

        HTMLWorker htmlparser = new HTMLWorker(Doc);

        PdfWriter.GetInstance(Doc, Response.OutputStream);



        Doc.Open();

        htmlparser.Parse(stringReader);

        Doc.Close();

        Response.Write(Doc);

        Response.End();
    }

    public override void VerifyRenderingInServerForm(Control control)
    {

    } 
}
