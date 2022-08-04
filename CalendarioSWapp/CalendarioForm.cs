using CalendarioSWapp.ClasesCalendar;
using System;
using System.Data;
using System.Data.SqlTypes;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Globalization;

namespace CalendarioSWapp
{
    public partial class CalendarioForm : Form
    {
        public CalendarioForm()
        {
            InitializeComponent();
            FilesCalendar.ObtenerDireccionesCalendar();
            BDcalendar.connSTr = FilesCalendar.AccesoBD;
            
        }
        int MonthValue = 0;
        private void Form1_Load(object sender, EventArgs e)
        {
            
            foreach (Control L in TLPanelCalendar.Controls)
            {
                if (L is ListView)
                {
                    ((ListView)L).View = View.Details;
                    //((ListView)L).LabelEdit = true;
                    //((ListView)L).AllowColumnReorder = true;
                    ((ListView)L).GridLines = true;
                }
            }
            FuncionesCalendar.EnumerarCalendario(TLPanelCalendar, monthCalendar1);
            BDcalendar.BuscarTPenBD(TLPanelCalendar, "I");
            BDcalendar.BuscarTPenBD(TLPanelCalendar, "F");
            TITULOmes.Text = FuncionesCalendar.mesCalendario;
            LVToday.View = View.Details;
            LVToday.GridLines = true;
            LVToday.Columns.Add("Today", 85, HorizontalAlignment.Left);
            LVToday.Columns.Add("State", 63, HorizontalAlignment.Left);
            FuncionesCalendar.PlaceToday(TLPanelCalendar, monthCalendar1,LVToday);
            MonthValue = monthCalendar1.SelectionStart.Month;
        }

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {
            if(MonthValue != monthCalendar1.SelectionStart.Month)
            {
                MonthValue = monthCalendar1.SelectionStart.Month;
                FuncionesCalendar.ClearControls(TLPanelCalendar);
                FuncionesCalendar.EnumerarCalendario(TLPanelCalendar, monthCalendar1);
                BDcalendar.BuscarTPenBD(TLPanelCalendar, "I");
                BDcalendar.BuscarTPenBD(TLPanelCalendar, "F");
                TITULOmes.Text = FuncionesCalendar.mesCalendario;

                //FuncionesCalendar.PlaceToday(TLPanelCalendar, monthCalendar1,LVToday);
            }
        }
        private void BtnFiltroCRT_Click(object sender, EventArgs e)
        {
            if (TxtBoxCRT.Text == "Activar")
            {
                TxtBoxCRT.Text = "";
                //BtnNotif.Visible = true;
                BtnDeleteTP.Visible = true;
                //BtnUpdate.Visible = true;
            }
            else if(TxtBoxCRT.Text.Length==10)
            {
                if (TxtBoxCRT.Text.Substring(0,3) == "CRT")
                {
                    BDcalendar.FiltrarSWX(TLPanelCalendar, monthCalendar1, TxtBoxCRT);
                }
                else if (TxtBoxCRT.Text.Substring(0, 3) == "SWX")
                {
                    BDcalendar.FiltrarSWX(TLPanelCalendar, monthCalendar1, TxtBoxCRT, "SWX");
                }
                else
                {
                    MessageBox.Show("Ingresar un CRT o un Ticket SWX");
                }
            }
            else if(TxtBoxCRT.Text=="")
            {
                MonthValue = monthCalendar1.SelectionStart.Month;
                FuncionesCalendar.ClearControls(TLPanelCalendar);
                FuncionesCalendar.EnumerarCalendario(TLPanelCalendar, monthCalendar1);
                BDcalendar.BuscarTPenBD(TLPanelCalendar, "I");
                BDcalendar.BuscarTPenBD(TLPanelCalendar, "F");
                TITULOmes.Text = FuncionesCalendar.mesCalendario;
            }
            
        }
        private void LV1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV1.GetItemAt(e.X, e.Y).Text);
        }
        private void LV2_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV2.GetItemAt(e.X, e.Y).Text);
        }
        private void LV3_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV3.GetItemAt(e.X, e.Y).Text);
        }
        private void LV4_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV4.GetItemAt(e.X, e.Y).Text);
        }
        private void LV5_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV5.GetItemAt(e.X, e.Y).Text);
        }
        private void LV6_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV6.GetItemAt(e.X, e.Y).Text);
        }
        private void LV7_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV7.GetItemAt(e.X, e.Y).Text);
        }
        private void LV8_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV8.GetItemAt(e.X, e.Y).Text);
        }
        private void LV9_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV9.GetItemAt(e.X, e.Y).Text);
        }
        private void LV10_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV10.GetItemAt(e.X, e.Y).Text);
        }
        private void LV11_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV11.GetItemAt(e.X, e.Y).Text);
        }
        private void LV12_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV12.GetItemAt(e.X, e.Y).Text);
        }
        private void LV13_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV13.GetItemAt(e.X, e.Y).Text);
        }
        private void LV14_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV14.GetItemAt(e.X, e.Y).Text);
        }
        private void LV15_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV15.GetItemAt(e.X, e.Y).Text);
        }
        private void LV16_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV16.GetItemAt(e.X, e.Y).Text);
        }
        private void LV17_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV17.GetItemAt(e.X, e.Y).Text);
        }
        private void LV18_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV18.GetItemAt(e.X, e.Y).Text);
        }
        private void LV19_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV19.GetItemAt(e.X, e.Y).Text);
        }
        private void LV20_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV20.GetItemAt(e.X, e.Y).Text);
        }
        private void LV21_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV21.GetItemAt(e.X, e.Y).Text);
        }
        private void LV22_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV22.GetItemAt(e.X, e.Y).Text);
        }
        private void LV23_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV23.GetItemAt(e.X, e.Y).Text);
        }
        private void LV24_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV24.GetItemAt(e.X, e.Y).Text);
        }
        private void LV25_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV25.GetItemAt(e.X, e.Y).Text);
        }
        private void LV26_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV26.GetItemAt(e.X, e.Y).Text);
        }
        private void LV27_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV27.GetItemAt(e.X, e.Y).Text);
        }
        private void LV28_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV28.GetItemAt(e.X, e.Y).Text);
        }
        private void LV29_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV29.GetItemAt(e.X, e.Y).Text);
        }
        private void LV30_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV30.GetItemAt(e.X, e.Y).Text);
        }
        private void LV31_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV31.GetItemAt(e.X, e.Y).Text);
        }
        private void LV32_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV32.GetItemAt(e.X, e.Y).Text);
        }
        private void LV33_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV33.GetItemAt(e.X, e.Y).Text);
        }
        private void LV34_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV34.GetItemAt(e.X, e.Y).Text);
        }
        private void LV35_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LV35.GetItemAt(e.X, e.Y).Text);
        }
        private void CalendarioForm_SizeChanged(object sender, EventArgs e)
        {
            if(this.WindowState == FormWindowState.Maximized)
            {
                LabelNMC.Font = new Font("Microsoft Sans Serif", 15, FontStyle.Bold);
                TITULOmes.Font = new Font("Microsoft Sans Serif", 13, FontStyle.Bold | FontStyle.Underline);
            }
            else
            {
                LabelNMC.Font = new Font("Microsoft Sans Serif", 12, FontStyle.Bold);
                TITULOmes.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold | FontStyle.Underline);
            }
            
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            ClientesForm Clientes = new ClientesForm(dataGridView1.CurrentRow.Cells["Id"].Value.ToString(), dataGridView1.CurrentRow.Cells["Ticket"].Value.ToString());
            Clientes.StartPosition = FormStartPosition.CenterScreen;
            Clientes.ShowDialog();
        }

        private void BtnDeleteTP_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Eliminar el ticket seleccionado", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                {
                    if (dataGridView1.DataSource != null)
                    {
                        int ID = Convert.ToInt32(dataGridView1.CurrentRow.Cells["Id"].Value);
                        BDcalendar.EliminarSWX(ID);
                        dataGridView1.Rows.Remove(dataGridView1.CurrentRow);
                        //ActualizarCalendario();
                        FuncionesCalendar.ClearControls(TLPanelCalendar);
                        FuncionesCalendar.EnumerarCalendario(TLPanelCalendar, monthCalendar1);
                        BDcalendar.BuscarTPenBD(TLPanelCalendar, "I");
                        BDcalendar.BuscarTPenBD(TLPanelCalendar, "F");
                    }

                }
                else
                {
                    //NADA
                }
            }
            catch (Exception ExcepcionDelete)
            {
                MessageBox.Show("ExcepcionDelete:\n\r-Error al intentar eliminar SWX\r\n" + ExcepcionDelete.Message);
            }
        }

        private void BtnUpdate_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                string SWX = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                UpdateTP FormUpdate = new UpdateTP(Convert.ToInt32(dataGridView1.CurrentRow.Cells["Id"].Value), dataGridView1.CurrentRow.Cells["Fecha inicio"].Value.ToString().Split(' ')[0], dataGridView1.CurrentRow.Cells["Hora inicio"].Value.ToString(), dataGridView1.CurrentRow.Cells["Fecha fin"].Value.ToString().Split(' ')[0], dataGridView1.CurrentRow.Cells["Hora fin"].Value.ToString());
                
                FormUpdate.StartPosition = FormStartPosition.CenterScreen;
                FormUpdate.ShowDialog();

                if(BDcalendar.Actualizado == "Si")
                {
                    FuncionesCalendar.ClearControls(TLPanelCalendar);
                    FuncionesCalendar.EnumerarCalendario(TLPanelCalendar, monthCalendar1);
                    BDcalendar.BuscarTPenBD(TLPanelCalendar, "I");
                    BDcalendar.BuscarTPenBD(TLPanelCalendar, "F");
                    TITULOmes.Text = FuncionesCalendar.mesCalendario;
                    BDcalendar.MostrarDescripcion(dataGridView1, SWX);
                    BDcalendar.Actualizado = "No";
                }
                


            }
        }
        private void BtnNotif_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int ID = Convert.ToInt32(dataGridView1.CurrentRow.Cells["Id"].Value);
                BDcalendar.GetTableSWXComplete(ID);
                string AsuntoNotify = "";
                string Fi = "";
                string Ff = "";
                string FechaNoFormat = "";
                string FechaSiFormat = "";
                DateTime FormatoDT;

                NotifyBox FormBox = new NotifyBox();
                FormBox.ShowDialog();
                if (FuncionesCalendar.NotifyText != "")
                {
                    try
                    {
                        foreach (DataRow fila in BDcalendar.DTComplete.Rows)
                        {
                            //Fi = "Start Time: " + fila["Fecha"].ToString().Split(' ')[0] + " " + fila["Fecha"].ToString().Split(' ')[2].Substring(0,4) + " UTC";
                            //Ff = "End Time: " + fila["Fecha"].ToString().Split(' ')[4] + " " + fila["Fecha"].ToString().Split(' ')[6].Substring(0,4) + " UTC";

                            FechaNoFormat = fila["Fecha"].ToString().Split(' ')[0];
                            FormatoDT = DateTime.ParseExact(FechaNoFormat, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                            FechaSiFormat = FormatoDT.ToString("dd/MMMM/yyyy", CultureInfo.GetCultureInfo("en-GB"));
                            Fi = "Start Time: " + FechaSiFormat + " " + fila["Fecha"].ToString().Split(' ')[2].Substring(0, 5) + " UTC";
                            FechaNoFormat = fila["Fecha"].ToString().Split(' ')[4];
                            FormatoDT = DateTime.ParseExact(FechaNoFormat, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                            FechaSiFormat = FormatoDT.ToString("dd/MMMM/yyyy", CultureInfo.GetCultureInfo("en-GB"));
                            Ff = "End Time: " + FechaSiFormat + " " + fila["Fecha"].ToString().Split(' ')[6].Substring(0, 5) + " UTC";
                            AsuntoNotify = "[" + FuncionesCalendar.NotifyText + "] Maintenance Notification: " + fila["Ticket"] + " - " + fila["Cliente"];
                            FuncionesCalendar.NotifyMail(fila["Cliente"].ToString(), AsuntoNotify, fila["Ticket"].ToString(), Fi, Ff, fila["Impact"].ToString(), FuncionesCalendar.NotifyText);
                        }
                        MessageBox.Show("Correos Generados");
                        FuncionesCalendar.NotifyText = "";
                    }
                    catch(Exception E)
                    {
                        MessageBox.Show(E.ToString());
                    }
                    
                }
            }
        }

        private void BtnFiltroCRT_MouseHover(object sender, EventArgs e)
        {
            ToolTip ToolTip1 = new ToolTip();
            ToolTip1.SetToolTip(this.BtnFiltroCRT, "CRT o Ticket SWX");
        }


        private void LVToday_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            if ((LVToday.Items != null) && (monthCalendar1.SelectionStart.Month == DateTime.Today.Month))
            {
                string operador = "+";
                int DifHora = 0;
                int horalocal = Convert.ToInt32(DateTime.Now.ToString("HH"));
                int horaUTC = 0;

                if (TimeZoneInfo.Local.ToString().Substring(0, 5) != "(UTC)")
                {
                    operador = TimeZoneInfo.Local.ToString().Substring(4, 1);
                    DifHora = Convert.ToInt32(TimeZoneInfo.Local.ToString().Split(Convert.ToChar(operador))[1].Substring(0, 2));
                }

                if (operador == "-")
                {
                    horaUTC = horalocal + DifHora;
                    if (horaUTC > 23) { horaUTC = horaUTC - 24; }
                }
                else if (operador == "+")
                {
                    horaUTC = horalocal - DifHora;
                    if (horaUTC > 23) { horaUTC = horaUTC - 24; }
                }

                foreach (ListViewItem Itm in LVToday.Items)
                {
                    if (((Convert.ToInt32(Itm.SubItems[1].Text.Split(':')[1].TrimStart(' ')) - horaUTC) == 1) || (Convert.ToInt32(Itm.SubItems[1].Text.Split(':')[1].TrimStart(' ')) - horaUTC) == -1 || (Convert.ToInt32(Itm.SubItems[1].Text.Split(':')[1].TrimStart(' ')) - horaUTC) == 0)
                    {
                        Itm.BackColor = Color.CornflowerBlue;
                    }
                    else
                    {
                        Itm.BackColor = SystemColors.Window;
                    }
                }
            }



        }

        private void LVToday_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            BDcalendar.MostrarDescripcion(dataGridView1, LVToday.GetItemAt(e.X, e.Y).Text);
        }

        private void TxtBoxCRT_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                BtnFiltroCRT.PerformClick();
            }
        }
    }
}
