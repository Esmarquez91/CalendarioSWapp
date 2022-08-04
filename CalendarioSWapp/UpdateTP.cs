using System.Windows.Forms;
using CalendarioSWapp.ClasesCalendar;


namespace CalendarioSWapp
{
    public partial class UpdateTP : Form
    {
        public UpdateTP(int ID, string Fi, string Hi, string Ff, string Hf)
        {
            InitializeComponent();
            TPId = ID.ToString();

            TBUpdate.Text = Fi + " - " + Hi + " - " + Ff + " - " + Hf;
        }

        string TPId;
        private void TBUpdate_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Return)
            {
                BDcalendar.ActualizarFecha(TPId, TBUpdate.Text);
                this.Close();
            }
            else if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }
    }
}
