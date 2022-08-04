using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace CalendarioSWapp.ClasesCalendar
{
    public class FuncionesCalendar
    {
        public static string NotifyText = "";
        public static string ObtenerMes(string Mes)
        {
            switch (Mes)
            {
                case "January": Mes = "Enero"; break;
                case "january": Mes = "Enero"; break;
                case "enero": Mes = "Enero"; break;
                case "February": Mes = "Febrero"; break;
                case "february": Mes = "Febrero"; break;
                case "febrero": Mes = "Febrero"; break;
                case "March": Mes = "Marzo"; break;
                case "march": Mes = "Marzo"; break;
                case "marzo": Mes = "Marzo"; break;
                case "April": Mes = "Abril"; break;
                case "april": Mes = "Abril"; break;
                case "abril": Mes = "Abril"; break;
                case "May": Mes = "Mayo"; break;
                case "may": Mes = "Mayo"; break;
                case "mayo": Mes = "Mayo"; break;
                case "June": Mes = "Junio"; break;
                case "june": Mes = "Junio"; break;
                case "junio": Mes = "Junio"; break;
                case "July": Mes = "Julio"; break;
                case "july": Mes = "Julio"; break;
                case "julio": Mes = "Julio"; break;
                case "August": Mes = "Agosto"; break;
                case "august": Mes = "Agosto"; break;
                case "agosto": Mes = "Agosto"; break;
                case "September": Mes = "Septiembre"; break;
                case "september": Mes = "Septiembre"; break;
                case "septiembre": Mes = "Septiembre"; break;
                case "setiembre": Mes = "Septiembre"; break;
                case "Setiembre": Mes = "Septiembre"; break;
                case "October": Mes = "Octubre"; break;
                case "october": Mes = "Octubre"; break;
                case "octubre": Mes = "Octubre"; break;
                case "November": Mes = "Noviembre"; break;
                case "november": Mes = "Noviembre"; break;
                case "noviembre": Mes = "Noviembre"; break;
                case "December": Mes = "Diciembre"; break;
                case "december": Mes = "Diciembre"; break;
                case "diciembre": Mes = "Diciembre"; break;
            }
            return Mes;
                
        }

        public static void ClearControls(TableLayoutPanel TL1)
        {
            foreach (Control LV in TL1.Controls)
            {
                if (LV is ListView)
                {
                    (LV as ListView).Columns.Clear();
                    (LV as ListView).Items.Clear();
                    (LV as ListView).BackColor = SystemColors.Window;
                    (LV as ListView).BorderStyle = BorderStyle.FixedSingle;
                }
            }
        }

        public static void PlaceToday(TableLayoutPanel TL1, MonthCalendar M1, ListView LT)
        {
            if (DateTime.Today.ToString("MM") == M1.SelectionRange.Start.ToString("MM"))
            {
                int AddDay = 0;
                //ListView Temporal = new ListView();

                switch (primerdiadelmes)
                {
                    case "domingo":
                        AddDay = 0; break;
                    case "lunes":
                        AddDay = 1; break;
                    case "martes":
                        AddDay = 2; break;
                    case "miércoles":
                        AddDay = 3; break;
                    case "jueves":
                        AddDay = 4; break;
                    case "viernes":
                        AddDay = 5; break;
                    case "sábado":
                        AddDay = 6; break;
                }

                //MessageBox.Show("AddDay = " + AddDay.ToString() + " - Primer dia del mes: " + primerdiadelmes);
                foreach (Control L in TL1.Controls)
                {
                    //if (L.Name == "LV" + (1 + AddDay).ToString())
                    if (L.Name == "LV" + (DateTime.Today.Day + AddDay).ToString())
                    {
                        (L as ListView).BackColor = Color.CornflowerBlue;
                        //MessageBox.Show((L as ListView).Name);
                        foreach (ListViewItem Li in (L as ListView).Items)
                        {
                            ListViewItem Temporal = new ListViewItem(Li.Text);
                            Temporal.SubItems.Add(Li.SubItems[1]);
                            LT.Items.Add(Temporal);
                        }
                        break;
                    }
                }

            }

            



        }

        public static string diaCalendario, mesCalendario, anocalendario, primerdiadelmes;
        public static int mesCalendarionum;
        public static void EnumerarCalendario(TableLayoutPanel TL1, MonthCalendar Calendario1)
        {
            diaCalendario = Calendario1.SelectionRange.Start.ToString("dddd");
            mesCalendario = Calendario1.SelectionRange.Start.ToString("MMMM");
            mesCalendarionum = Convert.ToInt32(Calendario1.SelectionRange.Start.ToString("MM"));
            anocalendario = Calendario1.SelectionRange.Start.ToString("yyyy");
            DateTime PrimerDia = new DateTime(Convert.ToInt32(anocalendario), mesCalendarionum, 1);
            primerdiadelmes = PrimerDia.ToString("dddd");
            mesCalendario = ObtenerMes(mesCalendario);
            if ((mesCalendario == "Febrero") && (DateTime.IsLeapYear(Convert.ToInt32(anocalendario)))) //Año bisiesto.
            {
                
                switch (primerdiadelmes)
                {
                    case "domingo":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 1 && nombredeLV <= 29) { ((ListView)L).Columns.Add((nombredeLV).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                    case "lunes":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 2 && nombredeLV <= 30) { ((ListView)L).Columns.Add((nombredeLV - 1).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                    case "martes":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 3 && nombredeLV <= 31) { ((ListView)L).Columns.Add((nombredeLV - 2).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                    case "miércoles":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 4 && nombredeLV <= 32) { ((ListView)L).Columns.Add((nombredeLV - 3).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                    case "jueves":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 5 && nombredeLV <= 33) { ((ListView)L).Columns.Add((nombredeLV - 4).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                    case "viernes":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 6 && nombredeLV <= 34) { ((ListView)L).Columns.Add((nombredeLV - 5).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                    case "sábado":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 7 && nombredeLV <= 35) { ((ListView)L).Columns.Add((nombredeLV - 6).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;

                }
            }
            else if (mesCalendario == "Febrero")
            {
                switch (primerdiadelmes)
                {
                    case "domingo":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 1 && nombredeLV <= 28) { ((ListView)L).Columns.Add((nombredeLV).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                    case "lunes":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 2 && nombredeLV <= 29) { ((ListView)L).Columns.Add((nombredeLV - 1).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                    case "martes":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 3 && nombredeLV <= 30) { ((ListView)L).Columns.Add((nombredeLV - 2).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                    case "miércoles":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 4 && nombredeLV <= 31) { ((ListView)L).Columns.Add((nombredeLV - 3).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                    case "jueves":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 5 && nombredeLV <= 32) { ((ListView)L).Columns.Add((nombredeLV - 4).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                    case "viernes":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 6 && nombredeLV <= 33) { ((ListView)L).Columns.Add((nombredeLV - 5).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                    case "sábado":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 7 && nombredeLV <= 34) { ((ListView)L).Columns.Add((nombredeLV - 6).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                }
            }
            else if ((mesCalendario == "Abril") || (mesCalendario == "Junio") || (mesCalendario == "Septiembre") || (mesCalendario == "Noviembre"))
            {
                switch (primerdiadelmes)
                {
                    case "domingo":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 1 && nombredeLV <= 30) { ((ListView)L).Columns.Add((nombredeLV).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                    case "lunes":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 2 && nombredeLV <= 31) { ((ListView)L).Columns.Add((nombredeLV - 1).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                    case "martes":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 3 && nombredeLV <= 32) { ((ListView)L).Columns.Add((nombredeLV - 2).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                    case "miércoles":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 4 && nombredeLV <= 33) { ((ListView)L).Columns.Add((nombredeLV - 3).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                    case "jueves":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 5 && nombredeLV <= 34) { ((ListView)L).Columns.Add((nombredeLV - 4).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                    case "viernes":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 6 && nombredeLV <= 35) { ((ListView)L).Columns.Add((nombredeLV - 5).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                    case "sábado":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 7 && nombredeLV <= 35) { ((ListView)L).Columns.Add((nombredeLV - 6).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                                if (nombredeLV == 1){((ListView)L).Columns.Add((nombredeLV + 29).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                }
            }
            else
            {
                switch (primerdiadelmes)
                {
                    case "domingo":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 1 && nombredeLV <= 31) { ((ListView)L).Columns.Add((nombredeLV).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                    case "lunes":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 2 && nombredeLV <= 32) { ((ListView)L).Columns.Add((nombredeLV - 1).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                    case "martes":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 3 && nombredeLV <= 33) { ((ListView)L).Columns.Add((nombredeLV - 2).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                    case "miércoles":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 4 && nombredeLV <= 34) { ((ListView)L).Columns.Add((nombredeLV - 3).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }
                        break;
                    case "jueves":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 5 && nombredeLV <= 35) { ((ListView)L).Columns.Add((nombredeLV-4).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }break;
                    case "viernes":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 6 && nombredeLV <= 35) { ((ListView)L).Columns.Add((nombredeLV - 5).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                                if (nombredeLV == 1) { ((ListView)L).Columns.Add((nombredeLV + 30).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                            }
                        }break;
                    case "sábado":
                        foreach (Control L in TL1.Controls)
                        {
                            int nombredeLV = 0;
                            if (L is ListView)
                            {
                                if ((L.Name.Where(char.IsDigit).ToArray().Length > 1))
                                { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()) * 10 + Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[1].ToString()); }
                                else { nombredeLV = Convert.ToInt32((L.Name.Where(char.IsDigit).ToArray())[0].ToString()); }
                                if (nombredeLV >= 7 && nombredeLV <= 35) { ((ListView)L).Columns.Add((nombredeLV - 6).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left); }
                                if(nombredeLV >= 1 && nombredeLV <= 2) { ((ListView)L).Columns.Add((nombredeLV + 29).ToString(), 78, HorizontalAlignment.Left); ((ListView)L).Columns.Add("", 49, HorizontalAlignment.Left);}
                            }
                        }break;
                }
            }

        }

        public static string itemLV1, itemLV2 = "";
        public static string Dia1, Dia2 = "";
        public static void UbicarTPenCalendar(TableLayoutPanel TL1, string ticket, string fecha, string hora, string Estado)
        {
            string FirstDay = primerdiadelmes;

            fecha = fecha.Split(' ')[0];

            if (hora.Split(' ').Length > 1) { hora = hora.Split(' ')[1].Split(':')[0] + ":" + hora.Split(' ')[1].Split(':')[1]; }
            else if(hora.Split(':').Length>2)
            { hora = hora.Substring(0, 5); }

            int dd = Convert.ToInt32(fecha.Split('/')[0]);
            int mm = Convert.ToInt32(fecha.Split('/')[1]);
            int yy = Convert.ToInt32(fecha.Split('/')[2]);
            int FijarDia = 0;

            switch (FirstDay)
            {
                case "domingo":
                    FijarDia = dd + 0; break;
                case "lunes":
                    FijarDia = dd + 1; break;
                case "martes":
                    FijarDia = dd + 2; break;
                case "miércoles":
                    FijarDia = dd + 3; break;
                case "jueves":
                    FijarDia = dd + 4; break;
                case "viernes":
                    FijarDia = dd + 5; break;
                case "sábado":
                    FijarDia = dd + 6; break;
            }
            if (FijarDia > 35) { FijarDia = FijarDia - 35; }

            foreach (Control L in TL1.Controls)
            {
                if (L is ListView)
                {
                    Dia1 = "LV" + FijarDia.ToString();
                    if (Dia1 != Dia2) { itemLV2 = ""; }
                    if (L.Name == Dia1)
                    {
                        itemLV1 = ticket + " " + Estado + ": " + hora;
                        if (itemLV1 == itemLV2)
                        {
                            //No se agrega nada.
                        }
                        else
                        {
                            ListViewItem item1 = new ListViewItem(ticket);
                            item1.SubItems.Add(Estado + ": " + hora);
                            (L as ListView).Items.Add(item1);
                            if (Estado == "i") { (L as ListView).Items[(L as ListView).Items.Count-1].BackColor = Color.FromArgb(193, 255, 196); }
                            else { (L as ListView).Items[(L as ListView).Items.Count - 1].BackColor = Color.FromArgb(255, 196, 196); }
                            
                        }
                        Dia2 = Dia1;
                        itemLV2= itemLV1;
                    }
                }
            }
            
        }

                

        public static void NotifyMail(string cliente2, string Asunto, string TicketID, string FechaI, string FechaF, string Impacto,string TipoCorreo)
        {
            
            try
            {
                
                Outlook._Application _app = new Outlook.Application();
                Outlook.MailItem mail = (Outlook.MailItem)_app.CreateItem(Outlook.OlItemType.olMailItem);
                mail.To = "sworks@telxius.com";
                mail.Subject = Asunto;

                if ((cliente2 == "FACEBOOK AMERICA") || (cliente2 == "FACEBOOK") || (cliente2 == "NQT-INTERNEXA") || (cliente2 == "INTERNEXA PERU") || (cliente2 == "INTERNEXA") || (cliente2 == "LAN NAUTILUS CAPACIDAD") || (cliente2 == "LAN") || (cliente2 == "ANTEL") || (cliente2 == "GOOGLE INTERNATIONAL LLC") || (cliente2 == "GOOGLE") || (cliente2 == "MICROSOFT") || (cliente2 == "AMAZON") || (cliente2 == "BLACKBURN") || (cliente2 == "AT&T") || (cliente2 == "ATT"))
                {
                    if ((cliente2 == "GOOGLE INTERNATIONAL LLC") || (cliente2 == "GOOGLE") || (cliente2 == "MICROSOFT") || (cliente2 == "AMAZON") || (cliente2 == "BLACKBURN") || (cliente2 == "ATT") || (cliente2 == "AT&T"))
                    {
                        if ((cliente2 == "ATT") || (cliente2 == "AT&T"))
                        {
                            mail.CC = "customerservice.capacity@telxius.com; manageronduty@telxius.com; juanantonio.bravo@telxius.com; alberto.leivaocana@telxius.com; vinicius.mantovani@telxius.com;";
                        }
                        mail.CC = "customerservice.capacity@telxius.com; manageronduty@telxius.com; juanantonio.bravo@telxius.com; alberto.leivaocana@telxius.com;";
                    }
                    else
                    {
                        mail.CC = "customerservice.capacity@telxius.com; manageronduty@telxius.com; juanantonio.bravo@telxius.com";
                    }
                }
                else if ((cliente2 == "VIVO") || (cliente2 == "TELEFONICA BRASIL-CAPACIDAD") || (cliente2 == "TELEFONICA EMPRESAS BRASIL") || (cliente2 == "ALOO") || (cliente2 == "ALOO TELECOM") || (cliente2 == "RED CLARA") || (cliente2 == "ELLALINK") || (cliente2 == "WIRELINK") || (cliente2 == "WIRELINK SOBRAL") || (cliente2 == "SOFTCOM"))
                {
                    mail.CC = "customerservice.capacity@telxius.com; manageronduty@telxius.com; vinicius.mantovani@telxius.com;";
                }
                else if ((cliente2 == "PRT") || (cliente2 == "PRT LARGA DISTANCIA INC") || (cliente2 == "PRT LARGA DISTANCIA INC.") || (cliente2 == "PRT LARGA DISTANCIA"))
                {
                    mail.CC = "customerservice.capacity@telxius.com; manageronduty@telxius.com; netwSchedWorksNotif.businesssolutions@telefonica.com";
                }
                else
                {
                    mail.CC = "customerservice.capacity@telxius.com; manageronduty@telxius.com;";
                }

                string Notify = "Notificación";
                string Estado = "State";
                if (TipoCorreo == "STARTED") { Notify = "We inform you that this scheduled work is confirmed and will begin at the scheduled time. // Le informamos que este trabajo está confirmado y empezará a la hora programada."; Estado = "To begin // Por empezar"; }
                else if (TipoCorreo == "COMPLETED") { Notify = "We inform you that this scheduled work has concluded. Could you confirm the status of your services? // Le informamos que este trabajo programado ha concluido. Podría validar el estado de sus servicios?"; Estado = "Completed // Completado"; }
                else if (TipoCorreo == "CANCELLED") { Notify = "We inform you that this scheduled work has been cancelled. // Le informamos que este trabajo programado ha sido cancelado."; Estado = "Cancelled // Cancelado"; }

                //string html1 = @"<html>
                //                <body style=""font-size:100%"">
                //                <p>Dear Sirs // Estimados Señores,</p>
                //                <p><b> " + cliente2 + " </b></p>" +
                //                "<p>Please find below our Scheduled Work Notification // Abajo encontrará nuestra Notificación de Trabajo Programado:</p>" +
                //                "<p><b>NOTIFICATION NUMBER // NUMERO DE NOTIFICACION: </b> " + TicketID + "</p>" +
                //                "<p><b>NOTIFICATION TYPE // TIPO DE NOTIFICACION: </b> Completed // Completado </p>" +
                //                "<p><b>SERVICE IMPACT // IMPACTO EN SERVICIOS: </b>" + Impacto + "</p>" +
                //                "<p><b>SCHEDULE // VENTANA(S) DE TRABAJO (UTC):</b></p>" +
                //                scheduleHTML;
                string html1 = @"<html>
                                <body style=""font-size:100%"">
                                <p>Dear Sirs // Estimados Señores,</p>
                                <p><b>" + cliente2 + "</b></p>" +
                                "<p>"+Notify+"</p>" +
                                "<p><b>NOTIFICATION NUMBER // NUMERO DE NOTIFICACION: </b> " + TicketID + "</p>" +
                                "<p><b>NOTIFICATION TYPE // TIPO DE NOTIFICACION: </b>" +Estado+ "</p>" +
                                "<p><b>SCHEDULE // VENTANA DE TRABAJO:</b></p>" +
                                "<p>" + FechaI + "</p>" +
                                "<p>" + FechaF + "</p>";

                string html2 = @"<html>
                                <body style=""font-size:90%"">
                                <p style="" Color:#666666;"">[English]<br>
                                If you shall experience any problem with your services due to the performing of this task, please contact our Capacity Services NOC at +511 411 0070 or mailing to customerservice.capacity@telxius.com.<br>
                                We would like to apologize for any inconvenience caused by this maintenance to you and your customers. If any additional question appears, do not hesitate to contact us. Our personnel will be permanently available for you.</p>
                                <p style="" Color:#666666;"">[Español]<br>
                                Si experimenta algún problema con su servicio debido a la realización de esta actividad programada, por favor comuníquese a nuestro NOC de Capacity Services vía el +511 411 0070 o al correo customerservice.capacity@telxius.com.<br>
                                Lamentamos los inconvenientes que este mantenimiento pueda causarle a usted o sus clientes. Si una nueva consulta surgiera, no dude en comunicarse con nosotros, nuestro personal estará permanentemente disponible para usted.</p>
                                <br>";

                string firmaHTML = @"<table width=""450"" height=""120"" style =""border-style:none;"">
                                    <tr>
                                    <td width=""120"" style=""border-style:none; padding-top:15px"" valign=""top"">
                                    <a><img src=""data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAgQAAABiCAMAAAAY7jvKAAAAwFBMVEX///8ALjsAlKEAKzgAJjQkQEvn7O3CzdAAIDAAIjEdOkW2wMLR2tzy9fYAMT/V3N4RO0c4T1iNl5t3hYpeb3cAGyxUbnYAj52hsLUAFym03OAAGSsdmKQAmqYAHi/C5ejs7/Dq9fY8WWIADiOBlZtPYWiTpKlGq7UAAB4uSVOsuLwAipl/kZdqfoU2pK/g8fNxusKLyM+k09iGxs0vT1q9yMtMaXIVQE1DXWdYaXBxgIZGqLKcz9XP6Opltb4sRU8ZOQuuAAAQgUlEQVR4nO1dDXeiOhMGiQgqaFs/YN2yVC2tbOnu3av2497d/v9/9YIfCMkMSVCrL9fnnD1njyVDPh4yycxkoigXXHDBBRdccMEFF1zwX4J1t4NwkS8pfhy1cmcFvd1Yw6J+b3DQ1hGB2wc6+d/9DiyHX7MO+kgR7v5+uE3x/POLSJlfD7XhDn9+A48grVh3iY81ZdsSC3kAL4J1M1aq5ThBEDjhotHRxd7mR01b1dawo8zrGuPtzyhUuxkwb+lEqp0+MF/s/hDObERidwrV1b/ppgXUbijTEWv8fh7WMhgOn7mzwd2fXJHasMYUcbpF/aLa9YAVawX1tCX2DHgAgDOz0yLdSOgr0Bs3r01bI6YRwySrAWrOo2DBIYM+9+JnN9CMcfpwQyPp7yg0zb2hBPovRkYgISkLwj7REIlaf8DWzHrN1kzri3VdFj/zAxoPKZcFD0yRv6gnFiOsFZuKapMWI7bXTwvFD/SfBCo/9UimiNHlzwat97pr5Pt4xSHTcI1ZL1jgJSMv1wY37eox4VJgPcyU8Ju8QK27+d23i/oO6DjHzQtyfYGuy+G5RmN4W1yCoU1t+JN6hN8v5I3+7BaT/BOPbW7d9cd8ESMqfr4zHXkFX61GjH6X7eMNKFqb23d16vx5YAUv/4Va91QnTTYUbkzg8pv30hOKogwoQaa0QnhgSFAb/ltU4BtQ4G/qGX6/aDP6q30y8k8YU27dgz4ldFw0FcTqxuDXy0AmU32Uf5Dcb/6w4IncwnzPCezMqMq4LbhV1HvfmKq9UiTAmoDjL+a7jvGtoADw+HdqZag3+Z1t0+P1blJ9xvmsFZY32kvBqsBR+yKfrPYIE6mTn7x3JGgJTgQq6eUFvlAF+5sPeGrA5TdSGBJYV5Qg45rbcxR+A4Nae8af/5slzfCBekaEBCrV11aP4jPVZxBuKN5oXVSF6BF/FlhjBE+mbZoE28FoiS0J4hLLnMBOlyaBs/4DhwT3CoUDkABQ8cD8ngJQBrWvtKXg/EjQ6Bb2bBYePJlWmwR3t9BcAG39E7DryNrwF/3Q2ZGg1RUdKXQtUm0SKL8ADgyf4T0zsIJglMH5kWDhiupt9b9KAlgh0Fv/FaAFxC07aZwZCVqFG28axyNB3s5zZiQAFcIQUgiMmQhmS5ndwfFI4AsY9LJdCJOA2R1sh1SYBNR2ByVB4MHlN+89xu4gwReABDXAcCimDBQhO8HLp5GgRz3FAWafyJvldkMqTAIjbwNFSRC6cHnqvbueOwwJlD9CCuEHxBXQpDDnWwzHPtWUY5EgLPywWGAkoLb1k+1OUpwE+b0nSgK90FjEegYORYI7YNHPmIBElYGSWLN5U0GfruixSGCpUsoAJ0G+TaTrb34XJQH5yM99KAmsXgFticnYsg5FAuULNBU85HcIkNJAnE1WZJqJlw6YiFfOO4PM6d3HsUgQyCmDAqNr4HpkDdOYzNI3sSQgENwZFQyAkkDR710TlEEM94UNKTgYCQR2CODyETMnKO0wuI7RZPo4+fU6AJpyHBL4tHF2+6iWOJP7o1HfiDtcy3iHXNSF1JkO1oimLT/9lSGBthywiBya9TgJFGsRASKS94aATftwJAAVwjBrCgRoghsWN/DpFaJm+8ijRyKBAypY4tkfV8soCFthGETL8VXd1taxBSrRxANaVmBIYIoF9xSQQA6HI4Hy7TtAgj+Zv4srgx10lgSYn+9IJKA9bCvJk6XTzg+1324518u57T6O8RAuGCwJ+D7wBOdIAlgh7CzCkL24yNm4wslJAFkszDEy4VsxfMl5oGIkABXC7fZb/wUoAzqUhMXJSQAs3Y3I5wqWQaVIoPyGpoLNQP8AGMIPRjw9CVjrm0CYghyqRQJQIWzW/0AUwXeBuOSTkyBilu5oBcqiYiS4A0LN1vElPwBl8IcnTjk9CawxvSTYq4NAVIwEsEJIAg6BOYKJJIFwchIwVgKzZE/jqBoJ4IDDH8o3YCIQOqRychLQHV0iEpeHypHAegbnfdZpIKQMzoAETCSBeVEHXEBRI1//AL+JHUA8OQmYmUATOKAih+qRQPkXUgjs7MCEFcI4OQlYz4HJFyuHCpJAARQCpCHEcHISAEehjIEvWHsxVJEE4DkECreip9FPTQLWTpAMUv2gi8MqkgDeIeQngsJjalmcnARgvJ42ilB3sTwqSQKuQoDDCkGcnAQLOJqAePaTRA6EQrToqJVKkAByGudIgEaSMDg5CdC4Z82w74ODbBQONhMYQQdHAWOPQgKOQkDCCkGcnAQFocYa8SbjltiAFeFgJNC6H3UMH/dsopMtjkMCBTyYlioDiWn09CRoFMZvE6/7Fvji7YFwMBIU5rwh7hgTdCQSFO0QhjJ5qk5PAoUWS3c88dzX0JdoE43DkaAYE2zJcCQSQJ7jLQcEzURrnAEJGvyIcGJe7TEdfBYJ0P44FgmUfxAWCJuJ1jgDEtC5hkBo/Xi3INWyHT6LBGgalqOR4PdXmAS3cvuqcyCBLpRYKt4tLMvR4NNIwCT42eBoJIBiChMIOZB3OAcSKB2hNDWqZtqlOu/TSFBHSHo8EsCxZv9ICjkLEihtwe7WjFlBCjsMVSbBHagQhA3Ga5wHCZS2K3pekMi7FapMAvB4Yu0796hBDmdCAkXvCeYq0R5FcmjmUGkSwApBIM48g3MhgWIFquDBVCBZZDEqvDBUsAQmUnvEsyFBXJUbQ0wnTPhJNHOo7hZxBTCBiViE6QZnRAJFadwTodmAyHnzqmssWgNUCILhhSucFQni8YomJr/ntYmUwaCiZuO7dA8ATgXi4QT7kYBN4ktDlgTJFQMfJnc6EKBfBhV1IP1Mc0/sqxDOjgRxlVo9b8TJZ2bKmAs+w5U8e337ZFfyl9owzW0MKgThEMP9SMAk7GNQhgTJq5zerHA+IHPRBiqfFFTiFzXnGCS4zQSO3IFTgfAOYR8SaK9cP0VJEiQVi5cHroluF0YSq4JKxhiuPv40/wQYaybsT5YggRLRI4rti3coT4IEjadXA9ELpoTJqIok2Ix6ahSCUhzWvgqajGRIQN93IDCi+5FAUfzFgIDbBfZeFhwVJMHWQpRO+RaY0UpQIciQgB5Rlb4ziF9ElgQxrGkdWB3IZDGoIAnSmKJ0DwDuEArvRtlBhgT0NSaqx+2UA5AgruMSSCjtii8KqkeCjNcoNQrBOwSh+VKGBAE9FgS4/S2Pd3otWYYEzG1iCZDrTyBUjgTZ5FWpUQhOcchLYbiCDAnYjIMeh2gNdqddLoL8idEIArdwbXEwEiBXrnDBkkDaE5pD7qP/vt0DADkqagXJTDOQIcGC+R45n0aDST6gdctFiDHdKNOPpUlAn5iWIF5x7alr1yRBjXZ6nQXsVBYQKEMCfUKTgNwXTQUOe/8m6nLngd6ZyOQ4K00Cum+kndhb0CQggz1O2NHzfmo4hBWCQJSRDAkU6uJBtTjB0NRj93baVcmDZcxlIxIDUpoE9L2IAnZyGHTWVgELCw72xtyt4i+rEKRIwEYF4/48/xUKJC/djcz29BPWBDr9AWsfJceOvmpV9cqGz8MjvR1n8FYMvjtRigRTdlyZqzG2cq9A039pZfjGdKP4Iq0sCZhrcktvD+hrcvdI2AnN+TsfQjmFIEWCNpSLGrzztKXCVv9JiWDh1ZuZJaYrLqksCdgDcppdrvq0wT0WVHYqgD72f1LzMKgQuO5EKRIw8+OKBV32mwSWhBvhJRdES4ZTEh6k0iRgdJBqllvRsZk4SOGl0TigVPfZC3DAJOg8hSBFAqBXkhL9cSvfNe8jJBrAAKbTxrXT4HSI1WMsFDI5zkqTAMiiQsaNEjQATl6TWVhCEJTBOpfGHFYInPgSORJ0mE3iukFklskuo0fYPVEmsKMMVc+zZ/dPLTyFfQe4tat4c5pHaRKwhpH4xe7LU9jo6Bh8SBB0c5bmdt/DRRuWgtUIUgb58HJYIRS7E+VIAEzLqzJeZr/WqWNXR5I6K9p6IUnUFjHd0SxyWmy+jyTOCHipjM2tNAnoS/Y2rTWI2kSTVdTHzDU6UAbnlSDTVLuwEOQQNngRFuUkgs6r4+5Ey7csizGIxEuW5GoJpEwDurlMszOzfMvGAkFIHTqMuptbNNMgzY+rQfQULNbfQ9uJxk0CyvMkEluVJwF6cWBRrOEEyMDH+N54glxozQDuDOg7LUCFgLgTO9eDcQK2Xsmvb+9IuiAguQyxM/fROGiwsPEBXVtDacukE4npuZMVXM9EYko0Ex41EOVJ4Eje2LiGx+7/WnJxy2py8wdbHb4ySAAmMAEVQsf2yIpyQAevBwJWCzpTwsxw1ooAK+GmUfc+JA9SuyJdJOOBKU8CdmcqBJcRZPGvJabhMUKgGDLoTgtQIUDuxKjwzucESD/TK2Z36e+aCjn+V9AmEaxhypFAzg9VngS8VDoIABvGk/ScMqErCZ03gy84AhOYAGwRuCt57oP98p5d6WpuxkjQgEKA1o/ZmKmtHAk8KX/8HiRowNshDgDDYkd6TunTAROQCQA+ZwQqBNadKDA9YTtx/3431JqZ4UDBktBEb7ErRQLNlNpi70EC2bu814D8GtfcqZcCbReH4scwAwCUBJ1VCGgKyR3QvaI/33aM2c0M7hQ9W95f4rvOUiQw5HIU7EMC2aNo6/oBJFjthWVAkcCClAG684Ny3t7S7kQREqjY0LXr6/YY451mtp4wC5FmvBd8t2VIAKy+C7EPCZRQLJVODqCHk4my4rUyTwJIGeBOgR9AApMhrRD2IoGiz5O5zcvY0fVXNtZgIwa93nqFEiQw+Aef8tiLBEogcEKWriHo5l5wjtZRyJMAVAYFh0tAhUC5E/cjgaL3+hrJ7B62cwMLUi/29smvvIw3v1Aii/1IEO+HZFmAxDosujILjDwJpJRBAvZCJGYZuScJFOVmlFkBh4jjOK8xQLAha5xa9ZcHuCZXzo+7QHc9CLBYB30sISi3xUDuPysAqBDya8N9FoZr+Lv/Tg1MGD8mHbr5pKhSWok43QUtxZR043aWE6laorEn1vRRmAZu1jAOXHLFS2P+C3A6533Ke2wRGUSY0tQeBULAHNG8ZeoqrWmZoBTaSVYiNCR8laimWpBDo9PjZ1/YCMl0P+AP4B8xAwyH1NKQjqEEehwxFtGtGmMbYPNFaCcXjkXX36aNpwAoxCBvrnNLzCZ+OHYNE4mVoUC6RZIaPQN1iWQbmzt8z3gNhjX+YVNmWUDbF/n5hMXC4Bo2Yh4go1dBzeuHMxc7fJxCI4Z7UzYky4q/49TJZ2IWbB50531uT/qGSQji+lu/YAQG3GURPo27I7dIEHHnOSHfhvkBHT6LHCr5WcuWGg5pZ5M/6xc2RfM0oS63nJv6aCUqO2Ka6d3LTNztIJq5iRSYAHGfqL19bkGxwkFzjdnyplU+5l/vtMNptBzPZ00M8+VUoKK+3gmDm0EsqA4LoVv7JXdL9te/xA6d//75/HWL54d/mUL6dDlHW9Js1p+Et1F6xxnU7fhbNQ3D8DzP1Oy4tGxPJ8ED86aaCMl8IMQ0DaLWB0C0yX8N1u9vW4jfbvS5sDqL0HGCBM6i/DfbbjnBU+91/mLbdvdlPh5E10GrfIz+Bf+3sCzf9/X4n49GOF1wwQUXXHDBBRdccAGN/wHc7pSutHbvggAAAABJRU5ErkJggg=="" width=""150"" height=""30""></a>
                                    </td>
                                    <td width=""250"" style=""border-style:none; padding-left:15px; padding-top:0px; font-family: Helvetica, Arial, sans-serif; font-size:12px; line-height:0px;"" valign=""top"">
                                    <p style=""margin-bottom:-8em; font-size:12px; Color: #008B95; ""><b>Scheduled Works Team</b></p>
                                    <div><b> Service Operations | Telxius Cable </b></div>
                                    <p style="" font-size:11px; Color:#777777;"">5895 Paseo de la República Av. Leuro bldg. off 903<br>
                                    Miraflores 15047 - Lima - Perú<br>
                                    T  +51 1 411 0079<br>
                                    M  +51 945 300 791<br>
                                    sworks@telxius.com | telxius.com</p>
                                    </td>
                                    </tr>
                                    </table>
                                    ";
                //string firmaHTML = @"<table width=""400"" height=""120"" border=""0"" cellspacing=""0"" cellpadding=""0"" style =""border-style:none;"">
                //                    <tr>
                //                    <td width=""120"" style=""border-style:none"" align =""left"" valign=""top"">
                //                    <a><img src=""data:image/png;base64,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"" width=""120"" height=""30"" style=""padding-top:25px;""></a>
                //                    </td>
                //                    <td width=""220"" style=""padding-left:15px; padding-top:0px; font-family: Helvetica, Arial, sans-serif; font-size:10px; border:0px solid #719695; border-left: none; border-right: none; line-height:0px;"" valign=""top"">
                //                    <p style=""margin-bottom:-5em; font-size:12px; Color: #008B95; ""><b>Scheduled Works Team</b></p>
                //                    <div><b> Service Operations | Telxius Cable </b></div>
                //                    <p style="" Color:#777777;"">5895 Paseo de la República Av. Leuro bldg. off 903<br>
                //                    Miraflores 15047 - Lima - Perú<br>
                //                    T  +51 1 411 0079<br>
                //                    M  +51 945 300 791<br>
                //                    sworks@telxius.com | telxius.com</p>
                //                    </td>
                //                    </tr>
                //                    </table>
                //                    ";

                mail.HTMLBody = html1 + html2 + firmaHTML;
                mail.Importance = Outlook.OlImportance.olImportanceNormal;

                string TIPOIMPACT = Impacto.Split(' ')[0];
                if ((TIPOIMPACT == "PROTECTION") && ((cliente2.Split(' ')[0] == "Neutrona") || (cliente2.Split(' ')[0] == "NEUTRONA")))
                {
                    //GenerarAlerta(R, Color.Red, TicketID + " - Notificación PL para Neutrona omitida");
                }
                else
                {
                    ((Outlook._MailItem)mail).Save();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex);
            }
        }

    }
}
