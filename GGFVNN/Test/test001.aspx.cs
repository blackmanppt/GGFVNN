using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace GGFGAMA.Test
{
    public partial class test001 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            int i2 = 2, i3 = 7;
            float fx =  (float)i2 / (float)i3;
            Label1.Text = ((float)i2 / (float)i3 * 100).ToString("0.00");
        }

        protected void Button2_Click(object sender, EventArgs e)
        {
            Regex reg = new Regex("V[0-9]{4}");
            
            if (reg.IsMatch(TextBox1.Text)&&TextBox1.Text.Length==5)
            {
                Label2.Text = "true";
            }
            else
            {
                Label2.Text = "false";
            }
        }
    }
}