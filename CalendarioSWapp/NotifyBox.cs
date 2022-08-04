using System;
using System.Windows.Forms;
using CalendarioSWapp.ClasesCalendar;

namespace CalendarioSWapp
{
    public partial class NotifyBox : Form
    {
        public NotifyBox()
        {
            InitializeComponent();
        }

        private void BtnStart_Click(object sender, EventArgs e)
        {
            FuncionesCalendar.NotifyText = "STARTED";
            this.Close();
        }

        private void BtnComplete_Click(object sender, EventArgs e)
        {
            FuncionesCalendar.NotifyText = "COMPLETED";
            this.Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            FuncionesCalendar.NotifyText = "CANCELLED";
            this.Close();
        }
    }
}
