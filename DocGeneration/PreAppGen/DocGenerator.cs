using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using EllieMae.Encompass.ComponentModel;
using EllieMae.Encompass.Forms;

namespace PreAppGen
{
    [Plugin]
    public class DocGenerator : Form
    {
        // Custom Input Form Button
        private Button btnGenPreApp = null;
        // Class to Design Pre-Application Worksheet
        private CreatePreApp preApp = new CreatePreApp();

        public override void CreateControls ()
        {
            this.btnGenPreApp = (Button)FindControl("btnGenPreApp");
            this.btnGenPreApp.Click += new EventHandler(btnGenPreApp_Click);

            base.CreateControls();
        }

        private void btnGenPreApp_Click (object sender, EventArgs e)
        {
            preApp.CreatePDF();
        }
    }
}
