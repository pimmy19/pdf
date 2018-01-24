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
using System.Data.SqlClient;
using System.Configuration;
using System.Web.Configuration;

public partial class Default3 : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
      
    }
    
  
    private PdfPTable GetHeader() 
    {
        BaseFont fontthai_Bold = BaseFont.CreateFont(Server.MapPath("/Fonts/THSarabun Bold.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
        Font bold = new Font(fontthai_Bold, 17);
       
         BaseFont fontthai = BaseFont.CreateFont(Server.MapPath("/Fonts/THSarabun.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
         Font normal = new Font(fontthai, 16);
         Font smallnormal = new Font(fontthai, 14);

            PdfPTable headerTable = new PdfPTable(3);
            headerTable.TotalWidth = 530f;
            headerTable.HorizontalAlignment = 0;
            headerTable.SpacingAfter = 20;
            //headerTable.DefaultCell.Border = Rectangle.NO_BORDER;

            float[] headerTableColWidth = new float[3];
            headerTableColWidth[0] = 60f;
            headerTableColWidth[1] = 290f;
            headerTableColWidth[2] = 90f;
            headerTable.SetWidths(headerTableColWidth);
            headerTable.LockedWidth = true;

            string imageFile = Server.MapPath(".") + "/images/ff.png";
            iTextSharp.text.Image myImage = iTextSharp.text.Image.GetInstance(imageFile);
            myImage.ScaleAbsolute(60,70);

            PdfPCell headerTableCell_0 = new PdfPCell(myImage);
            headerTableCell_0.HorizontalAlignment = Element.ALIGN_LEFT;
            headerTableCell_0.Border = Rectangle.NO_BORDER;
            headerTable.AddCell(headerTableCell_0);

            PdfPCell headerTableCell_1 = new PdfPCell(new Phrase("แบบขอยืมคอมพิวเตอร์ /อุปกรณ์ต่อพ่วง และอุปกรณ์สื่อสารโทรคมนาคม\n \t \t \t \t \t \t สำหรับงานเฉพาะกิจ(งานบริการยืมชั่วคราว)", bold));
            headerTableCell_1.HorizontalAlignment = Element.ALIGN_LEFT;
            headerTableCell_1.VerticalAlignment = Element.ALIGN_BOTTOM;
            headerTableCell_1.Border = Rectangle.NO_BORDER;
            headerTable.AddCell(headerTableCell_1);

            Phrase p = new Phrase();

            p.Add(new Chunk("\t \t \t \t \t \tศูนย์คอมพิวเตอร์\n", smallnormal));
            p.Add(new Chunk("เลขทียืม:", smallnormal));
            p.Add(new Chunk("\nวันที่:", smallnormal));
            p.Add(new Chunk("\nเวลา:", smallnormal));
            PdfPCell headerTableCell_2 = new PdfPCell(p);           
            headerTableCell_2.HorizontalAlignment = Element.ALIGN_LEFT;
            headerTableCell_2.VerticalAlignment = Element.ALIGN_BOTTOM;
            headerTableCell_2.VerticalAlignment = Element.ALIGN_TOP;
                         
            headerTable.AddCell(headerTableCell_2);

            return headerTable;
       }

    private PdfPTable GetHeaderDetail()
     {
         BaseFont fontthai_Bold = BaseFont.CreateFont(Server.MapPath("/Fonts/THSarabun Bold.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
         Font bold = new Font(fontthai_Bold, 16);

         BaseFont fontthai = BaseFont.CreateFont(Server.MapPath("/Fonts/THSarabun.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
         Font normal = new Font(fontthai, 16);

            PdfPTable table = new PdfPTable(2);
            table.TotalWidth = 530f;
            table.HorizontalAlignment = 0;
            table.SpacingAfter = 10;

            float[] tableWidths = new float[2];
            tableWidths[0] = 400f;
            tableWidths[1] = 130f;

            table.SetWidths(tableWidths);
            table.LockedWidth = true;

           Chunk blank = new Chunk("", normal);

            Phrase p = new Phrase();

            p.Add(new Chunk("", normal));           
            PdfPCell cell0 = new PdfPCell(p);
            cell0.Border = Rectangle.NO_BORDER;
            table.AddCell(cell0);

            p = new Phrase();
            p.Add(new Chunk("วันที่:\t \t" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "\t \tน.", normal));

          

            PdfPCell cell1 = new PdfPCell(p);
            cell1.Border = Rectangle.NO_BORDER;
            table.AddCell(cell1);                                                         
                    
            return table;
    } 

//    private DataTable GetData(string query)
//{
//    string conString = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
//    SqlCommand cmd = new SqlCommand(query);
//    using (SqlConnection con = new SqlConnection(conString))
//    {
//        using (SqlDataAdapter sda = new SqlDataAdapter())
//        {
//            cmd.Connection = con;
 
//            sda.SelectCommand = cmd;
//            using (DataTable dt = new DataTable())
//            {
//                sda.Fill(dt);
//                return dt;
//            }
//        }
//    }
//}

   
    private Paragraph GetBody() {
        BaseFont fontthai_Bold = BaseFont.CreateFont(Server.MapPath("/Fonts/THSarabun Bold.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
        Font bold = new Font(fontthai_Bold, 16);

        BaseFont fontthai = BaseFont.CreateFont(Server.MapPath("/Fonts/THSarabun.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
        Font normal = new Font(fontthai, 16);

        string ConString = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;
        SqlConnection con = new SqlConnection(ConString);
        SqlCommand cmd = new SqlCommand();
        SqlDataReader myRead;
        cmd.CommandText = "SELECT TOP 1 position FROM user";
        cmd.Connection = con;
        con.Open();
        myRead = cmd.ExecuteReader();
        myRead.Read();


        string position = myRead.GetString(myRead.GetOrdinal("position"));
      //  string Activity = myReader.GetString(myReader.GetOrdinal("Activity"));
        //string name = myReader.GetString(myReader.GetOrdinal("Name"));
        //string sAddress1 = myReader.GetString(myReader.GetOrdinal("Address Line 1"));
        //string sAddress2 = myReader.GetString(myReader.GetOrdinal("Address Line 2"));
        //string sPost = myReader.GetString(myReader.GetOrdinal("Post Code"));
        //string sTelephone = myReader.GetString(myReader.GetOrdinal("Telephone Number"));
        //string sEmail = myReader.GetString(myReader.GetOrdinal("Email Address"));
        //string sRegistration = myReader.GetString(myReader.GetOrdinal("Registration Number"));

                Paragraph para = new Paragraph();
                para.FirstLineIndent = 38.1f;
                para.Add(new Phrase("ข้าพเจ้านาย/นาง/น.ส.\t", normal));
                para.Add(new Phrase(position + "\t \t", normal));
                para.Add(new Phrase("ตำแหน่ง", normal));
               // para.Add(new Phrase(Activity, normal));
                para.Add(new Phrase("................................................................\n", normal));
                para.Add(new Phrase("สังกัดหน่วยงาน", normal));
                para.Add(new Phrase("..............................................................................", normal));
                para.Add(new Phrase("โทรศัพท์ที่ติดต่อได้", normal));
                para.Add(new Phrase("................................................................\n", normal));
                para.Add(new Phrase("อีเมล์", normal));
                para.Add(new Phrase("..........................................................................................\n", normal));
           
            
            con.Close();


          //  para.Add(new Phrase(“ขออนุมัติให้นักศึกษาจำนวน ” + master.StudentCount + ” คน เดินทางไปราชการต่างประเทศระหว่างวันที่ ” + master.StartDate + ” ถึงวันที่ ” + master.EndDate + ” รวม ” + master.PeriodDay + ” วัน เพื่อดำเนินกิจกรรมดังต่อไปนี้”, normal));

            return para;
    }
           /*    private PdfPTable BodyData() {
                   BaseFont fontthai_Bold = BaseFont.CreateFont(Server.MapPath("/Fonts/THSarabun Bold.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                   Font bold = new Font(fontthai_Bold, 16);

                   BaseFont fontthai = BaseFont.CreateFont(Server.MapPath("/Fonts/THSarabun.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                   Font normal = new Font(fontthai, 16);
                   PdfPTable table = new PdfPTable(2);
                   table.TotalWidth = 430f;
                   table.HorizontalAlignment = 0;
                   table.SpacingAfter = 10;
                   table.HorizontalAlignment = Element.ALIGN_CENTER;

                   float[] tableWidths = new float[2];
                   tableWidths[0] = 215f;
                   tableWidths[1] = 215f;

                   table.SetWidths(tableWidths);
                   table.LockedWidth = true;

                   Chunk blank = new Chunk("", normal);

                   Phrase p = new Phrase();

                   p.Add(new Chunk("ตารางยืมตารางยืมตารางยืมตารางยืมตารางยืมตารางยืมตารางยืมตารางยืมตารางยืมตารางยืม\n", normal));
                   p.Add(new Chunk(blank));


                   PdfPCell cell0 = new PdfPCell(p);
                   cell0.Border = Rectangle.NO_BORDER;
                   table.AddCell(cell0);
                   p = new Phrase();
                   p.Add(new Chunk("ตารางยืมตารางยืมตารางยืมตารางยืมตารางยืมตารางยืมตารางยืมตารางยืมตารางยืมตารางยืมตารางยืม", normal));
                   p.Add(new Chunk(blank));
                   // p.Add(new Chunk(“2012”, normal));

                   PdfPCell cell1 = new PdfPCell(p);
                   cell1.Border = Rectangle.NO_BORDER;
                   table.AddCell(cell1);


                 

            return table;
     } */
     private Paragraph GetFooter()
               {
                   BaseFont fontthai_Bold = BaseFont.CreateFont(Server.MapPath("/Fonts/THSarabun Bold.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                   Font bold = new Font(fontthai_Bold, 16);

                   BaseFont fontthai = BaseFont.CreateFont(Server.MapPath("/Fonts/THSarabun.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                   Font normal = new Font(fontthai, 16);
                   Paragraph para = new Paragraph();

                   String ConString = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;
                   SqlConnection con = new SqlConnection(ConString);
                   SqlCommand cmd = new SqlCommand();
                   SqlDataReader myReader;
                   cmd.CommandText = "SELECT EN_borrow, Date_borrow, Name, Activity, place, Get_Equipment, Return_Equipment,Total_use,assign_to,Department,EN_Equipment FROM borroww";
                   cmd.Connection = con;
                   con.Open();
                   myReader = cmd.ExecuteReader();
                   myReader.Read();


                   
                   string Activity = myReader.GetString(myReader.GetOrdinal("Activity"));
                   string place = myReader.GetString(myReader.GetOrdinal("place"));
                   string Get_Equipment = myReader.GetString(myReader.GetOrdinal("Get_Equipment"));
                   string Return_Equipment = myReader.GetString(myReader.GetOrdinal("Return_Equipment"));
                   string Total_use = myReader.GetString(myReader.GetOrdinal("Total_use"));
                   string assign_to = myReader.GetString(myReader.GetOrdinal("assign_to"));
                   string Department = myReader.GetString(myReader.GetOrdinal("Department"));
                   string EN_Equipment = myReader.GetString(myReader.GetOrdinal("EN_Equipment"));
                   



                   para.Add(new Phrase("เพื่อใช้ในกิจกรรม\t \t", normal));
                   para.Add(new Phrase(Activity, normal));
                   para.Add(new Phrase(" \t \t \t"));
                   para.Add(new Phrase("สถานที่" + "\t \t", normal));
                   para.Add(new Phrase(place, normal));
                   para.Add(new Phrase("\n"));
                   para.Add(new Phrase("ทั้งนี้ขอรับอุปกรณ์ในวันที่" + "\t \t", normal));
                   para.Add(new Phrase(Get_Equipment, normal));
                   para.Add(new Phrase(" \t \t \t"));
                   para.Add(new Phrase("และส่งคืนในวันที่" + "\t \t", normal));
                   para.Add(new Phrase(Return_Equipment, normal));
                   para.Add(new Phrase(" \t \t \t"));
                   para.Add(new Phrase("รวมระยะเวลายืม" + "\t \t", normal));
                   para.Add(new Phrase(Total_use, normal));
                   para.Add(new Phrase(" \t \t \t"));
                   para.Add(new Phrase("วัน", normal)); 
                   para.Add(new Phrase("\n"));
                   para.Add(new Phrase("โดยหมอบหมายให้" +"\t \t", normal));
                   para.Add(new Phrase(assign_to, normal));
                  
                   para.Add(new Phrase(" หน่วยงาน"+"\t \t", normal));
                   para.Add(new Phrase(Department, normal));
                   para.Add(new Phrase("\n"));
                   para.Add(new Phrase("หมายเลขบัตรประจำประชาชน"+"\t \t", normal));
                   para.Add(new Phrase(EN_Equipment, normal));
                             
                   para.Add(new Phrase("เป็นผู้รับมอบหมายอุปกรณ์ดังกล่าวแทนข้าพเจ้า\n", normal));
                   para.Add(new Phrase("ขออนุญาติยืม", bold));
                   para.Add(new Phrase("\n"));
                   para.Add(new Phrase("\n"));
                   //  para.Add(new Phrase(“ขออนุมัติให้นักศึกษาจำนวน ” + master.StudentCount + ” คน เดินทางไปราชการต่างประเทศระหว่างวันที่ ” + master.StartDate + ” ถึงวันที่ ” + master.EndDate + ” รวม ” + master.PeriodDay + ” วัน เพื่อดำเนินกิจกรรมดังต่อไปนี้”, normal));

                   return para;
               } 

    protected void btnExport_Click(object sender, EventArgs e)
    {
        GridView1.AllowPaging = false;
        GridView1.DataBind();        
        
        Response.ContentType = "application/pdf";
        Response.AddHeader("content-disposition", "attachment;filename=Borrow.pdf");
        Response.Cache.SetCacheability(HttpCacheability.NoCache);

        //font
        BaseFont fontthai_Bold = BaseFont.CreateFont(Server.MapPath("/Fonts/THSarabun Bold.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
        Font bold = new Font(fontthai_Bold, 16);

        BaseFont fontthai = BaseFont.CreateFont(Server.MapPath("/Fonts/THSarabun.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
        Font normal = new Font(fontthai, 16);
        
        //image
        //string imageFile = Server.MapPath(".") + "/images/ff.png";
        //iTextSharp.text.Image myImage = iTextSharp.text.Image.GetInstance(imageFile);
       
        StringWriter sw = new StringWriter();
        HtmlTextWriter hw = new HtmlTextWriter(sw);
       // pnlPerson.RenderControl(hw);
        StringReader sr = new StringReader(sw.ToString());
        Document pdfDoc = new Document(PageSize.A4);
        HTMLWorker htmlparser = new HTMLWorker(pdfDoc);
        PdfWriter.GetInstance(pdfDoc, Response.OutputStream);
        pdfDoc.Open();
        BaseFont fontthai1 = BaseFont.CreateFont(Server.MapPath("/Fonts/THSarabun.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
        Font normal1 = new Font(fontthai1, 16);
        int cols = GridView1.Columns.Count;
        int rows = GridView1.Rows.Count;

        PdfPTable table = new PdfPTable(cols);
        table.SpacingBefore = 4;
        table.SpacingAfter = 8;
        table.WidthPercentage = 100;

        for (int j = 0; j < GridView1.Columns.Count; j++)
        {
            table.AddCell(new Phrase(GridView1.Columns[j].HeaderText, normal));
        }

        table.HeaderRows = 1;

        for (int i = 0; i < GridView1.Rows.Count; i++)
        {
            for (int k = 0; k < GridView1.Columns.Count; k++)
            {
                string cellValue = GridView1.Rows[i].Cells[k].Text;
                if (cellValue != null)
                {
                    table.AddCell(new Phrase(cellValue, normal));
                }
            }
        }


        pdfDoc.Add(GetHeader());
        pdfDoc.Add(GetHeaderDetail());

        //addimage
      //  myImage.ScaleToFit(50.0F, 230.0F);
     //   myImage.SpacingBefore = 50.0F;
       // myImage.SpacingAfter = 10.0F;
       // myImage.Alignment = 3;
       // pdfDoc.Add(myImage);


        //add font
        //  pdfDoc.Add(new Paragraph("แบบขอยืมคอมพิวเตอร์ /อุปกรณ์ต่อพ่วง และอุปกรณ์สื่อสารโทรคมนาคม", Nfont));
      //  Paragraph head1 = new Paragraph("แบบขอยืมคอมพิวเตอร์ /อุปกรณ์ต่อพ่วง และอุปกรณ์สื่อสารโทรคมนาคม\nสำหรับงานเฉพาะกิจ(งานบริการยืมชั่วคราว)", h1); //สร้างหัวข้อใหม่
       // head1.Alignment = Element.ALIGN_CENTER;
     //   pdfDoc.Add(head1);
        pdfDoc.Add(new Paragraph("เรียน ผู้อำนวยการศูนย์คอมพิวเตอร์", bold));
        pdfDoc.Add(GetBody());
        pdfDoc.Add(GetFooter());
       // pdfDoc.Add(BodyData());
        pdfDoc.Add(table);
        
        

        htmlparser.Parse(sr);
        Response.Write(pdfDoc);
        pdfDoc.Add(new Paragraph("\t \t \t \t \t \t \t \t \t \t \t \t \t \t จึงเรียนมาเพื่อโปรอนุญาต", normal));
        pdfDoc.Add(new Paragraph("\t \t \t \t \t \t \t \t \t \t \t \t \t \t ลงชื่อ...............................................ผู้ขอยืม\t \t \t \t \t \t \t \t \t \t ลงชื่อ...............................................หัวหน้าหน่วยงาน", normal));
        pdfDoc.Add(new Paragraph("\t \t \t \t \t \t \t \t \t \t \t \t \t \t \t \t \t \t \t \t(.................................................) \t \t \t \t \t \t \t \t \t \t \t \t \t \t \t \t \t \t \t \t \t \t \t (...............................................)", normal));
        pdfDoc.Add(new Paragraph("\t \t \t \t \t \t \t \t \t \t \t \t \t \t \t \t \t \t \t วันที่............................................   \t \t \t \t \t \t \t \t \t \t \t \t \t \t \t \t      วันที่..........................................", normal));
        pdfDoc.Add(new Paragraph("หมายเหตุ", bold));
        pdfDoc.Add(new Paragraph("ศูนย์บริการคอมพิวเตอร์สามารถให้บริการยืมคอมพิวเตอร์และอุปกรณ์ต่างๆไดไม่เกิน 10 วัน(นับทั้งวันทำการและวันหยุด)\nเนื่องจากมีผู้ขอใช้บริการจนวนมาก และโปรดแจ้งแบบฟอร์มแล้วหน้าอย่างน้อย 3 วันทำการ", normal));
        
        pdfDoc.Close();
        GridView1.AllowPaging = true;
        GridView1.DataBind();
        Response.End();
    }
    //public override void VerifyRenderingInServerForm(Control control)
    //{

    //}
}
