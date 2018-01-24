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

public partial class Default2 : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            //Populate DataTable
            DataTable dt = new DataTable();

            dt.Columns.Add("Name");
            dt.Columns.Add("Age");
            dt.Columns.Add("City");
            dt.Columns.Add("Country");
            dt.Rows.Add();
            dt.Rows[0]["Name"] = "Mudassar Khan";
            dt.Rows[0]["Age"] = "27";
            dt.Rows[0]["City"] = "Mumbai";
            dt.Rows[0]["Country"] = "India";

            //Bind Datatable to Labels
            lblName.Text = dt.Rows[0]["Name"].ToString();
            lblAge.Text = dt.Rows[0]["Age"].ToString();
            lblCity.Text = dt.Rows[0]["City"].ToString();
            lblCountry.Text = dt.Rows[0]["Country"].ToString();

        }
    }
    protected void btnExport_Click(object sender, EventArgs e)
    {
        Response.ContentType = "application/pdf";
        Response.AddHeader("content-disposition", "attachment;filename=Panel.pdf");
        Response.Cache.SetCacheability(HttpCacheability.NoCache);

        //font
        BaseFont EnCodefont = BaseFont.CreateFont(Server.MapPath("/Fonts/THSarabun.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
        Font Nfont = new Font(EnCodefont, 18, Font.NORMAL);
       
        //image
        string imageFile = Server.MapPath(".") + "/images/ff.png";
        iTextSharp.text.Image myImage = iTextSharp.text.Image.GetInstance(imageFile);

        StringWriter sw = new StringWriter();
        HtmlTextWriter hw = new HtmlTextWriter(sw);
        pnlPerson.RenderControl(hw);
        StringReader sr = new StringReader(sw.ToString());
        Document pdfDoc = new Document(PageSize.A4);
        HTMLWorker htmlparser = new HTMLWorker(pdfDoc);
        PdfWriter.GetInstance(pdfDoc, Response.OutputStream);
        pdfDoc.Open();


      

        
        //addimage
        myImage.ScaleToFit(50.0F, 230.0F);
        myImage.SpacingBefore = 50.0F;
        myImage.SpacingAfter = 10.0F;
        myImage.Alignment = 3;
        pdfDoc.Add(myImage);    
   

        //add font
      //  pdfDoc.Add(new Paragraph("แบบขอยืมคอมพิวเตอร์ /อุปกรณ์ต่อพ่วง และอุปกรณ์สื่อสารโทรคมนาคม", Nfont));
        Paragraph head1 = new Paragraph("แบบขอยืมคอมพิวเตอร์ /อุปกรณ์ต่อพ่วง และอุปกรณ์สื่อสารโทรคมนาคม\nสำหรับงานเฉพาะกิจ(งานบริการยืมชั่วคราว)", Nfont); //สร้างหัวข้อใหม่
        head1.Alignment = Element.ALIGN_CENTER;
        pdfDoc.Add(head1);


        htmlparser.Parse(sr);
        pdfDoc.Close();
        Response.Write(pdfDoc);
        Response.End();
    }
}