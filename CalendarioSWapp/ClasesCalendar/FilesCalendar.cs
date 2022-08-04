using System.IO;

namespace CalendarioSWapp.ClasesCalendar
{
    public class FilesCalendar
    {
        public static string ArchivoXLSX, ArchivoCSV, ArchivoFinal, csvData, AccesoBD;

        public static void ObtenerDireccionesCalendar()
        {
            try
            {
                string[] Direcciones = File.ReadAllLines("C:\\SWprogram\\direcciones.txt");
                AccesoBD = Direcciones[3];
            }
            catch
            {

            }
        }

    }
}