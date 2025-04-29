using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EVEDRI1
{
    public partial class Logs : Form
    {
        
        
        public Logs()
        {
            InitializeComponent();
            Mylogs mylogs = new Mylogs();
            mylogs.showlogs(logsDatagidview);


        }

        private void logsDatagidview_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
