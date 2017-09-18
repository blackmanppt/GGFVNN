using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace GGFVNN.VNN
{
    public partial class VNNindex : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (Session["username"] == null)
                Response.Redirect("~/LoginIndex.aspx");
            else if (Session["username"].ToString() != "VNN")
                Response.Redirect("~/LoginIndex.aspx");
        }
    }
}