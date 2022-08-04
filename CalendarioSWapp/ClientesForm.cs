using CalendarioSWapp.ClasesCalendar;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CalendarioSWapp
{
    public partial class ClientesForm : Form
    {
        public ClientesForm(string IDSW, string Ticket)
        {
            InitializeComponent();
            tID = IDSW;
            BDcalendar.BuscarIDenBDClientes(tID, DGClientes);
            LabelResultados.Text = DGClientes.Rows.Count.ToString() + " circuitos";
            TicketSelected.Text = Ticket;
            
        }

        string tID = "";

        private void BtnFiltrar_Click(object sender, EventArgs e)
        {
            BDcalendar.FiltroTablaClientes(DGClientes, TxtBoxFiltro.Text, tID);
            //LabelResultados.Text = DGClientes.Rows.Count.ToString();
        }

        private void ClientesForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }

        private void TxtBoxFiltro_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }

        private void DGClientes_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }

        private void BtnFiltrar_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }
    }
}
