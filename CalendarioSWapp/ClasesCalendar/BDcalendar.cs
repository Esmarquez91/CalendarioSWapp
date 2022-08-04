using System;
using System.Linq;
using System.Data.SqlClient;
using CalendarioSWapp.ClasesCalendar;
using System.Data;
using System.Windows.Forms;

namespace CalendarioSWapp.ClasesCalendar
{
    public class BDcalendar
    {
        public static string connSTr;
        public static SqlConnection conn;
        public static SqlCommand comando;
        public static DataTable dt = new DataTable();
        public static DataTable dt2 = new DataTable();
        public static DataTable dt3 = new DataTable();
        public static string Actualizado = "No";

        public static void BuscarTPenBD(TableLayoutPanel TL1, string IF)
        {
            dt.Clear();
            conn = new SqlConnection(connSTr);
            string mesCalendario = FuncionesCalendar.mesCalendario;
            string anocalendario = FuncionesCalendar.anocalendario;
            int mesCalendarionum = FuncionesCalendar.mesCalendarionum;
            string query;
            if (IF == "I") { IF = "Inicio"; }
            else if (IF == "F") { IF = "Fin"; }

            try
            {
                conn.Open();
                if ((mesCalendario == "Febrero") && (DateTime.IsLeapYear(Convert.ToInt32(anocalendario))))
                {
                    query = "SELECT * FROM trabajosprogramados WHERE [Fecha " + IF + "] BETWEEN '" + anocalendario +"-"+ mesCalendarionum + "-01' AND '" + anocalendario +"-"+ mesCalendarionum + "-29'";
                }
                else if (mesCalendario == "Febrero")
                {
                    query = "SELECT * FROM trabajosprogramados WHERE [Fecha " + IF + "] BETWEEN '" + anocalendario +"-"+ mesCalendarionum + "-01' AND '" + anocalendario +"-"+ mesCalendarionum + "-28'";
                }
                else if ((mesCalendario == "Abril") || (mesCalendario == "Junio") || (mesCalendario == "Septiembre") || (mesCalendario == "Noviembre"))
                {
                    query = "SELECT * FROM trabajosprogramados WHERE [Fecha " + IF + "] BETWEEN '" + anocalendario +"-"+ mesCalendarionum + "-01' AND '" + anocalendario +"-"+ mesCalendarionum + "-30'";
                }
                else
                {
                    query = "SELECT * FROM trabajosprogramados WHERE [Fecha " + IF + "] BETWEEN '" + anocalendario +"-"+ mesCalendarionum + "-01' AND '" + anocalendario +"-"+ mesCalendarionum + "-31'";
                }
                comando = new SqlCommand(query, conn);
                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = comando;
                da.Fill(dt);


                //FiltroCRT
                if (FILTRARCRT == 1)
                {
                    DataView dvFiltroCRT = new DataView(dt);
                    dvFiltroCRT.RowFilter = DVFILTER;
                    DataTable dt5 = dvFiltroCRT.ToTable();
                    dt = dt5;
                }


                int lineas = dt.Rows.Count;
                

                for (int i = 0; i < lineas; i++)
                {
                    string Ticket = dt.Rows[i][1].ToString();
                    string FechaInicio = dt.Rows[i][2].ToString();
                    string HoraInicio = dt.Rows[i][3].ToString();
                    string EstadoI = "i";
                    string FechaFin = dt.Rows[i][4].ToString();
                    string HoraFin = dt.Rows[i][5].ToString();
                    string EstadoF = "f";
                    if (IF == "Inicio")
                    {
                        FuncionesCalendar.UbicarTPenCalendar(TL1, Ticket, FechaInicio, HoraInicio, EstadoI);
                    }
                    else if (IF == "Fin")
                    {
                        FuncionesCalendar.UbicarTPenCalendar(TL1, Ticket, FechaFin, HoraFin, EstadoF);
                    }
                }
                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error " + ex);
                conn.Close();
            }

        }

        public static int FILTRARCRT = 0;
        public static string DVFILTER = "";
        public static void FiltrarSWX(TableLayoutPanel TL1, MonthCalendar m1, TextBox FILTRO, string SWX = null)
        {
            dt.Clear();
            conn = new SqlConnection(connSTr);
            if (FILTRO.Text != "")
            {
                try
                {
                    FILTRARCRT = 1;
                    conn.Open();
                    string query = "";
                    if (SWX == "SWX")
                    {
                        query = "Select Id from Clientes WHERE Ticket = '" + FILTRO.Text + "'";
                    }
                    else
                    {
                        query = "Select Id from Clientes WHERE CRT = '" + FILTRO.Text + "'";
                    }
                    
                    comando = new SqlCommand(query,conn);
                    SqlDataAdapter da = new SqlDataAdapter();
                    da.SelectCommand = comando;
                    da.Fill(dt);
                    conn.Close();
                    if (dt.Rows.Count > 0)
                    {
                        foreach (DataRow R in dt.Rows)
                        {
                            DVFILTER = DVFILTER + "Id = " + R[0].ToString() + " OR ";
                        }
                        DVFILTER = DVFILTER.Remove(DVFILTER.Length - 4);
                        FuncionesCalendar.ClearControls(TL1);
                        FuncionesCalendar.EnumerarCalendario(TL1, m1);
                        BDcalendar.BuscarTPenBD(TL1, "I");
                        BDcalendar.BuscarTPenBD(TL1, "F");

                        DVFILTER = "";
                        FILTRARCRT = 0;
                    }
                    else
                    {
                        MessageBox.Show("No hay trabajos registrados para este Circuito");
                        FILTRARCRT = 0;
                    }

                }
                catch (Exception ex10)
                {
                    conn.Close();
                    MessageBox.Show("Error al filtrar el CRT \r\n " + ex10);
                }

            }
            else
            {
                FILTRARCRT = 0;
                FuncionesCalendar.ClearControls(TL1);
                FuncionesCalendar.EnumerarCalendario(TL1, m1);
                BDcalendar.BuscarTPenBD(TL1, "I");
                BDcalendar.BuscarTPenBD(TL1, "F");
            }


        }



        public static void MostrarDescripcion(DataGridView DGV1, string SWX)
        {
            try
            {
                dt2.Clear();
                conn = new SqlConnection(connSTr);
                conn.Open();
                string query = "Select * FROM trabajosprogramados WHERE Ticket = '" + SWX + "'";
                comando = new SqlCommand(query, conn);
                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = comando;
                da.Fill(dt2);
                conn.Close();
                DGV1.DataSource = dt2;
                DGV1.Columns[0].Visible = false;
                DGV1.Columns[1].Width = 75;
                DGV1.Columns[2].Width = 87;
                DGV1.Columns[3].Width = 81;
                DGV1.Columns[4].Width = 74;
                DGV1.Columns[5].Width = 67;
                DGV1.Columns[6].Width = 267;
                //var result= MessageBox.Show(dt2.Rows[0][6].ToString(), "Copiar Descripción", MessageBoxButtons.YesNo);
                //if(result == DialogResult.Yes)
                //{
                //    Clipboard.SetText(dt2.Rows[0][6].ToString());
                //}
                //else { /*No copia*/}
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error " + ex);
                conn.Close();
            }
        }

        public static void BuscarIDenBDClientes(string tID,DataGridView DG3)
        {
            try
            {
                dt3.Clear();
                conn = new SqlConnection(connSTr);
                conn.Open();
                string query = "Select Cliente, CRT from Clientes WHERE Id=" + tID + "";
                comando = new SqlCommand(query, conn);
                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = comando;
                da.Fill(dt3);
                DG3.DataSource = dt3;
                conn.Close();
            }
            catch
            {
                conn.Close();
            }
        }

        public static void ActualizarFecha(string ID, string Fecha)
        {
            try
            {
                //01/11/2020 - 09:00:00 - 01/11/2020 - 15:00:00
                //string NewFechaInicio = Fecha.Split('-')[2].Split(' ')[0] + "-" + NombreMesTOnum(Fecha.Split('-')[1]) + "-" + Fecha.Split('-')[0];
                //string NewHoraInicio = Fecha.Split('-')[2].Split(' ')[1];
                //string NewFechaFin = Fecha.Split('-')[5].Split(' ')[0] + "-" + NombreMesTOnum(Fecha.Split('-')[4]) + "-" + Fecha.Split('-')[3].TrimStart(' ');
                //string NewHoraFin = Fecha.Split('-')[5].Split(' ')[1];

                string NewFechaInicio = Fecha.Split('-')[0].Split('/')[2].Split(' ')[0] + "-" + Fecha.Split('-')[0].Split('/')[1] + "-" + Fecha.Split('-')[0].Split('/')[0];
                string NewHoraInicio = Fecha.Split('-')[1].TrimStart(' ');
                string NewFechaFin = Fecha.Split('-')[2].Split('/')[2].Split(' ')[0] + "-" + Fecha.Split('-')[0].Split('/')[1] + "-" + Fecha.Split('-')[0].Split('/')[0];
                string NewHoraFin = Fecha.Split('-')[3].TrimStart(' ');

                conn = new SqlConnection(connSTr);
                conn.Open();
                string query = "UPDATE trabajosprogramados SET [Fecha inicio] = '" + NewFechaInicio + "', [Hora inicio] = '" + NewHoraInicio + "', [Fecha fin] = '" + NewFechaFin + "', [Hora fin]= '" + NewHoraFin + "' WHERE Id = " + ID + "";
                comando = new SqlCommand(query, conn);
                comando.ExecuteNonQuery();
                MessageBox.Show("Actualizado");
                Actualizado = "Si";
                conn.Close();
            }
            catch
            {
                conn.Close();
            }
        }

        public static string NombreMesTOnum(string Mes)
        {
            string result = "0";
            Mes = Mes.TrimEnd('.');
            if (Mes == "set") { Mes = "sep"; }
            switch (Mes)
            {
                case "ene": result = "1"; break;
                case "feb": result = "2"; break;
                case "mar": result = "3"; break;
                case "abr": result = "4"; break;
                case "may": result = "5"; break;
                case "jun": result = "6"; break;
                case "jul": result = "7"; break;
                case "ago": result = "8"; break;
                case "sep": result = "9"; break;
                case "oct": result = "10"; break;
                case "nov": result = "11"; break;
                case "dic": result = "12"; break;
            }


            switch (Mes)
            {
                case "Jan": result = "1"; break;
                case "Feb": result = "2"; break;
                case "Mar": result = "3"; break;
                case "Apr": result = "4"; break;
                case "May": result = "5"; break;
                case "Jun": result = "6"; break;
                case "Jul": result = "7"; break;
                case "Aug": result = "8"; break;
                case "Sep": result = "9"; break;
                case "Oct": result = "10"; break;
                case "Nov": result = "11"; break;
                case "Dec": result = "12"; break;
            }

            return result;
        }

        public static void FiltroTablaClientes(DataGridView DG4, string Filtro,string tID)
        {
            if (Filtro.Length > 0)
            {
                dt3.Clear();
                conn = new SqlConnection(connSTr);
                conn.Open();
                string query = "Select Cliente, CRT from Clientes WHERE Id=" + tID + " AND CRT= '" + Filtro + "'";
                comando = new SqlCommand(query, conn);
                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = comando;
                da.Fill(dt3);
                DG4.DataSource = dt3;
                conn.Close();
            }
            else
            {
                BuscarIDenBDClientes(tID, DG4);
            }
        }

        public static void EliminarSWX(int ID_SWX)
        {
            try
            {
                conn = new SqlConnection(connSTr);
                conn.Open();
                string query = "Delete from trabajosprogramados where ID= @ID";
                comando = new SqlCommand(query, conn);
                comando.Parameters.AddWithValue("@ID", ID_SWX);
                comando.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception ExBorrado)
            {
                MessageBox.Show("Excepción ExBorrado:\n\r-" + ExBorrado.Message);
                conn.Close();
            }
            try
            {
                conn = new SqlConnection(connSTr);
                conn.Open();
                string query = "Delete from Clientes where ID= @ID";
                comando = new SqlCommand(query, conn);
                comando.Parameters.AddWithValue("@ID", ID_SWX);
                comando.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error " + ex);
                conn.Close();
            }
        }

        public static DataTable DTComplete;
        public static DataTable DTtp;
        public static DataTable DTcl;
        public static void GetTableSWXComplete(int ID_SWX)
        {
            DTComplete = new DataTable();
            DTtp = new DataTable();
            DTcl = new DataTable();
            
            try
            {
                DTComplete = new DataTable();
                DTComplete.Columns.Add("Ticket");
                DTComplete.Columns.Add("Cliente");
                DTComplete.Columns.Add("Fecha");
                DTComplete.Columns.Add("Impact");
                

                conn = new SqlConnection(connSTr);
                conn.Open();
                string query = "SELECT * FROM trabajosprogramados where ID= " + ID_SWX;
                comando = new SqlCommand(query, conn);
                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = comando;
                da.Fill(DTtp);
                conn.Close();
                conn.Open();
                query = "SELECT DISTINCT ID,Ticket,Cliente FROM Clientes where ID = " + ID_SWX;
                comando = new SqlCommand(query, conn);
                da = new SqlDataAdapter();
                da.SelectCommand = comando;
                da.Fill(DTcl);
                conn.Close();


                foreach(DataRow R in DTcl.Rows)
                {
                    DataRow DTR1 = DTComplete.NewRow();
                    DTR1["Ticket"] = R["Ticket"];
                    DTR1["Cliente"] = R["Cliente"];
                    DTR1["Fecha"] = DTtp.Rows[0]["Fecha Inicio"].ToString() + " " + DTtp.Rows[0]["Hora Inicio"] + " | " + DTtp.Rows[0]["Fecha fin"] + " " + DTtp.Rows[0]["Hora Fin"];
                    DTR1["Impact"] = DTtp.Rows[0]["Impact"];
                    DTComplete.Rows.Add(DTR1);
                }


            }
            catch (Exception ExBorrado)
            {
                MessageBox.Show("Excepción ExBorrado:\n\r-" + ExBorrado.ToString());
                conn.Close();
            }
        }

    }
}
