using McDermott.Lib.Office;
using System;
using System.Data;

namespace WebOpenXML
{
    public partial class Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }
     
        private DataSet CreateSampleData()
        {
            //  Create a sample DataSet, containing three DataTables.
            //  (Later, this will save to Excel as three Excel worksheets.)
            //
            DataSet ds = new DataSet();

            //  Create the first table of sample data
            DataTable dt1 = new DataTable("Drivers");
            dt1.Columns.Add("UserID", Type.GetType("System.Decimal"));
            dt1.Columns.Add("Surname", Type.GetType("System.String"));
            dt1.Columns.Add("Forename", Type.GetType("System.String"));
            dt1.Columns.Add("Sex", Type.GetType("System.String"));
            dt1.Columns.Add("Date of Birth", Type.GetType("System.DateTime"));

            dt1.Rows.Add(new object[] { 1, "James", "Brown", "M", new DateTime(1962, 3, 19) });
            dt1.Rows.Add(new object[] { 2, "Edward", "Jones", "M", new DateTime(1939, 7, 12) });
            dt1.Rows.Add(new object[] { 3, "Janet", "Spender", "F", new DateTime(1996, 1, 7) });
            dt1.Rows.Add(new object[] { 4, "Maria", "Percy", "F", null });
            dt1.Rows.Add(new object[] { 5, "Malcolm", "Marvelous", "M", new DateTime(1973, 5, 7) });
            ds.Tables.Add(dt1);


            //  Create the second table of sample data
            DataTable dt2 = new DataTable("Vehicles");
            dt2.Columns.Add("Vehicle ID", Type.GetType("System.Decimal"));
            dt2.Columns.Add("Make", Type.GetType("System.String"));
            dt2.Columns.Add("Model", Type.GetType("System.String"));

            dt2.Rows.Add(new object[] { 1001, "Ford", "Banana" });
            dt2.Rows.Add(new object[] { 1002, "GM", "Thunderbird" });
            dt2.Rows.Add(new object[] { 1003, "Porsche", "Rocket" });
            dt2.Rows.Add(new object[] { 1004, "Toyota", "Gas guzzler" });
            dt2.Rows.Add(new object[] { 1005, "Fiat", "Spangly" });
            dt2.Rows.Add(new object[] { 1006, "Peugeot", "Lawnmower" });
            dt2.Rows.Add(new object[] { 1007, "Jaguar", "Freeloader" });
            dt2.Rows.Add(new object[] { 1008, "Aston Martin", "Caravanette" });
            dt2.Rows.Add(new object[] { 1009, "Mercedes-Benz", "Hitchhiker" });
            dt2.Rows.Add(new object[] { 1010, "Renault", "Sausage" });
            dt2.Rows.Add(new object[] { 1011, /*char.ConvertFromUtf32(12) + */ "Saab", "Chickennuggetmobile" });
            ds.Tables.Add(dt2);


            //  Create the third table of sample data
            DataTable dt3 = new DataTable("Vehicle owners");
            dt3.Columns.Add("User ID", Type.GetType("System.Decimal"));
            dt3.Columns.Add("Vehicle_ID", Type.GetType("System.Decimal"));

            dt3.Rows.Add(new object[] { 1, 1002 });
            dt3.Rows.Add(new object[] { 2, 1000 });
            dt3.Rows.Add(new object[] { 3, 1010 });
            dt3.Rows.Add(new object[] { 5, 1006 });
            dt3.Rows.Add(new object[] { 6, 1007 });
            ds.Tables.Add(dt3);

            return ds;
        }

        protected void btnExportToExcel_Click(object sender, EventArgs e)
        {
            try
            {
                //Set initial data
                string fileName = "Sample.xlsx";
                string title = "My Report Title";
                string comments = "This is report comments bla bla bla bla bla bla bla";
                DataSet ds = CreateSampleData();
                //Process Dataset
                ExcelReport myreport = new ExcelReport(ds, fileName, title, DateTime.Now, comments);
                byte[] data = myreport.GenerateExcelReport();
                Response.ClearContent();
                Response.Clear();
                Response.Buffer = true;
                Response.Charset = "";
                
                Response.Cache.SetCacheability(System.Web.HttpCacheability.NoCache);
                Response.AddHeader("content-disposition", "attachment; filename=" + fileName);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                
                Response.BinaryWrite(data);
                Response.Flush();

                System.Web.HttpContext.Current.Response.Flush();
                System.Web.HttpContext.Current.Response.SuppressContent = true;
                System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest();

            }
            catch (Exception ex)
            {
                string s = ex.Message;
            }
        }
    }
}