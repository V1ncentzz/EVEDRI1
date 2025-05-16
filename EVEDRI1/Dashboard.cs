using System.IO;
using Spire.Xls;
using System;
using System.Drawing;
using System.Windows.Forms;
using static System.Net.WebRequestMethods;


namespace EVEDRI1
{
    public partial class Dashboard : Form
    {
        public Dashboard()
        {
            InitializeComponent();


            if (!string.IsNullOrEmpty(Props.ProfilePath) && File.Exists(Props.ProfilePath))
            {
                pcbProfile.Image = Image.FromFile(Props.ProfilePath);
            }
            else
            {
                pcbProfile.Image = null;
            }

            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\HF\Desktop\Main\Book1.xlsx");

            Worksheet activeSheet = book.Worksheets[0];
            Worksheet inactiveSheet = book.Worksheets[1];

            int activeCount = activeSheet.LastRow - 1;
            int inactiveCount = inactiveSheet.LastRow - 1;

            lblActiveNum.Text = activeCount.ToString();
            lblInactiveNum.Text = inactiveCount.ToString();

            int male = 0, female = 0;
            int orange = 0, blue = 0, yellow = 0, red = 0;
            int basketball = 0, volleyball = 0, badminton = 0, tennis = 0, soccer = 0, baseball = 0;
            int bsit = 0, bstm = 0, beed = 0;

            void CountFromSheet(Worksheet sheet)
            {
                for (int i = 2; i <= sheet.LastRow; i++)
                {

                    string gender = sheet.Range[i, 2].Text.ToLower();
                    if (gender == "male") male++;
                    else if (gender == "female") female++;

                    string color = sheet.Range[i, 8].Text.ToLower();
                    if (color == "orange") orange++;
                    if (color == "blue") blue++;
                    if (color == "yellow")  yellow++;
                    if (color == "red") red++;

               
                    string values = sheet.Range[i, 3].Value;
                    string[] data = values.Split(' ');
                    foreach (var hobby in data)
                    {
                        if (hobby.Contains("Volleyball")) volleyball++;
                        if (hobby.Contains("Tennis")) tennis++;
                        if (hobby.Contains("Badminton")) badminton++;
                        if (hobby.Contains("Basketball")) basketball++;
                        if (hobby.Contains("Soccer")) soccer++;
                        if (hobby.Contains("Baseball")) baseball++;
                    }



                    string course = sheet.Range[i, 12].Text.ToLower();
                    if (course == "bsit") bsit++;
                    else if (course == "bstm") bstm++;
                    else if (course == "beed") beed++;
                }
            }

            CountFromSheet(activeSheet);
            CountFromSheet(inactiveSheet);

            // Assign values to labels
            lblMaleNum.Text = male.ToString();
            lblFemaleNum.Text = female.ToString();

            lblOrangeNum.Text = orange.ToString();
            lblBlueNum.Text = blue.ToString();
            lblRedNum.Text = red.ToString();
            lblYellowNum.Text = yellow.ToString();

            lblVolleyballNum.Text = volleyball.ToString();
            lblTennisNum.Text = tennis.ToString();
            lblBadmintonNum.Text = badminton.ToString();
            lblBasketballNum.Text = basketball.ToString();
            lblSoccerNum.Text = soccer.ToString();
            lblBaseballNum.Text = baseball.ToString();

            lblBsitNum.Text = bsit.ToString();
            lblBstmNum.Text = bstm.ToString();
            lblBeedNum.Text = beed.ToString();


            lblName.Text = Props.DisplayName;
            lblDate.Text = DateTime.Now.ToString("MM/dd/yyyy");
        }

        private void btnActive_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.Show();
            this.Hide();
        }

        private void btnInactive_Click(object sender, EventArgs e)
        {
            Form3 form3 = new Form3();
            form3.Show();
            this.Hide();
        }

        private void btnLogs_Click(object sender, EventArgs e)
        {
            Logs log = new Logs();
            log.Show();
            this.Hide();
        }

        private void btnLogout_Click(object sender, EventArgs e)
        {
            Login login = new Login();
            login.Show();
            this.Hide();
        }

        private void btnClosewindow_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnAddstudent_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();
            form1.Show();
            this.Hide();
        }
    }
}





    

