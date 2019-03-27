using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using System.Net.Mail;
using System.Net;

namespace Email_code
{
    class Program
    {

        static void Main(string[] args)
        {

            string connectionString = "Data Source=192.168.10.118;User ID=srihari;Password=Mouri@123;Initial Catalog=MT_SOW_DEV;Persist Security Info=True";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
               
                SendEmail(conn, "USA", "SOW");
                //SendEmail(conn, "IND", "SOW");
                //SendEmail(conn, "USA", "PO");
                //SendEmail(conn, "IND", "PO");
                //SendEmail(conn, "SA", "SOW");
                //SendEmail(conn, "SA", "PO");

            }
        }

        private static void SendEmail(SqlConnection conn, string locParameter, string type)
        {
            string USAMailTo = System.Configuration.ConfigurationManager.AppSettings["USAMailTo"];
            string INDMailTo = System.Configuration.ConfigurationManager.AppSettings["INDMailTo"];
            string SAMailTo = System.Configuration.ConfigurationManager.AppSettings["SAMailTo"];
            string MailFrom = System.Configuration.ConfigurationManager.AppSettings["MailFrom"];
            string MailFromPwd = System.Configuration.ConfigurationManager.AppSettings["MailFromPwd"];
            string Contracttype = string.Empty;
            DataTable dt = new DataTable();
            using (SqlCommand cmd = new SqlCommand("CMT_Emailservice", conn))
            {
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Location", locParameter);
                cmd.Parameters.AddWithValue("@Type", type);
                if (type == "SOW")
                {
                    Contracttype = "SOWCode";
                }
                else
                    Contracttype = "POCode";
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);
                foreach (DataRow dr in dt.Rows)
                {
                    using (MailMessage mail = new MailMessage())
                    {
                        mail.From = new MailAddress(MailFrom, "MOURI Tech-CMT");
                        mail.To.Add(INDMailTo);
                        if (locParameter == "USA")
                            mail.CC.Add(USAMailTo);
                        else if(locParameter == "IND")
                            mail.CC.Add(INDMailTo);
                        else
                            mail.CC.Add(SAMailTo);

                        mail.Subject = "Expiration of Service Contract";
                        mail.Body = "Dear user, <br />  <br /> The following  " + type + " <b>" + Convert.ToString(dr[Contracttype]) + "</b>" + " of Customer " + "<b>" + Convert.ToString(dr["CustomerName"]) + "</b>" + " <br />  will expire soon dated as " + Convert.ToDateTime(dr["EndDate"]).ToString("MM-dd-yyyy") + "(MM-dd-yyyy)<br /> for details please login into CMT application " + "test" + " <br />  <br /> Disclaimer :- This is an auto generated e-mail. Kindly do not reply on this mail. <br />  <br /> Regards, <br /> MOURI Tech Pvt Ltd- CMT";

                        mail.IsBodyHtml = true;
                        SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
                        smtp.Credentials = new NetworkCredential(MailFrom, MailFromPwd);
                        smtp.EnableSsl = true;
                        smtp.Send(mail);

                    }
                }
            }
        }
    }
}
