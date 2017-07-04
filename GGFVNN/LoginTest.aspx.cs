using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace GGFVNN
{
    public partial class LoginTest : System.Web.UI.Page
    {
        static string strConnectString = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["GGFConnectionString"].ToString();
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Login_Authenticate(object sender, AuthenticateEventArgs e)
        {
            using (SqlConnection connection=new SqlConnection(strConnectString))
            {
                SqlCommand command = new SqlCommand(
                  "SELECT [username],[link]  FROM [GGFCustomer]  where id=@id and [Password]=@Password",
                  connection);
                command.Parameters.Add("@id", SqlDbType.NVarChar).Value = Login.UserName;
                command.Parameters.Add("@Password", SqlDbType.NVarChar).Value = Login.Password;
                connection.Open();

                SqlDataReader reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        
                        Session["username"] = reader.GetString(0);
                        Response.Redirect(reader.GetString(1));
                        //Console.WriteLine("{0}\t{1}", reader.GetInt32(0),
                        //    reader.GetString(1));
                    }

                }
                else
                {
                    Console.WriteLine("No rows found.");
                }
                reader.Close();
            }
        }
    }
}