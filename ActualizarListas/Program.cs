using System;
using System.IO;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.Text;
using SpreadsheetLight;
using System.Diagnostics;

namespace ActualizarListas
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> alLista = new List<string>();
            List<string> archivos = new List<string>();
            List<string> hojas = new List<string>();
            List<string> renglon = new List<string>();
            List<string> columna = new List<string>();
            if (listarArchivos(archivos, hojas, renglon, columna))
                if (leerLista(alLista))
                    if (llenarFormatos(alLista, archivos, hojas, renglon, columna))
                    {
                        Console.WriteLine("todos los formatos fueron actualizados");
                        Console.ReadKey();
                    }
                    else
                    {
                        Console.WriteLine("algunos formatos no pudieron ser actualizados");
                        Console.ReadKey();
                    }
        }
        
        static private bool listarArchivos(List<string> archivos, List<string> hojas, List<string> renglon, List<string> columna)
        {//http://spreadsheetlight.com/sample-code/
            System.IO.IOException res = null;
            SLDocument sl = null;
            try
            {
                sl = new SLDocument(@"..\archivos.xlsx", "Hoja1");
                int cont = 1;
                string archivo;
                while (!string.IsNullOrWhiteSpace(archivo = sl.GetCellValueAsString(++cont, 1)))
                {
                    archivos.Add(archivo);
                    hojas.Add(sl.GetCellValueAsString(cont, 2));
                    renglon.Add(sl.GetCellValueAsString(cont, 3));
                    columna.Add(sl.GetCellValueAsString(cont, 4));
                }
                sl.CloseWithoutSaving();
            }
            catch (System.IO.IOException e)
            {
                res = e;
                Console.WriteLine("el archivo \"archivos.xlsx\" no existe o se encuentra abierto");
                Console.ReadKey();
            }
            return res == null;
        }

        static private bool leerLista(List<string> alLista)
        {//http://spreadsheetlight.com/sample-code/
            System.IO.IOException res = null;
            SLDocument sl = null;
            try
            {
                sl = new SLDocument(@"..\lista.xlsx", "Hoja1");
                int alNumero = 0;
                string alNombre;
                while (!string.IsNullOrWhiteSpace(alNombre = sl.GetCellValueAsString(++alNumero, 1)))
                {
                    alLista.Add(alNombre);
                }
                sl.CloseWithoutSaving();
            }
            catch (System.IO.IOException e)
            {
                res = e;
                Console.WriteLine("el archivo \"lista.xlsx\" no existe o se encuentra abierto");
                Console.ReadKey();
            }
            return res == null;
        }

        static private bool llenarFormatos(List<string> alLista, List<string> archivos, List<string> hojas, List<string> renglon, List<string> columna)
        {//http://spreadsheetlight.com/sample-code/
            System.IO.IOException res = null;
            SLDocument sl = null;
            for (int i = 0; i < archivos.Count; i++)
                try
                {
                    sl = new SLDocument(@"..\" + archivos.ElementAt(i)+".xlsx", hojas.ElementAt(i));
                    int alNumero = int.Parse(renglon.ElementAt(i));
                    for (int j = 0; j < alLista.Count; j++)
                        sl.SetCellValue(alNumero++, int.Parse(columna.ElementAt(i)), alLista.ElementAt(j));
                    string strNombre;
                    while (!string.IsNullOrWhiteSpace(strNombre = sl.GetCellValueAsString(alNumero, int.Parse(columna.ElementAt(i)))))
                        sl.SetCellValue(alNumero++, int.Parse(columna.ElementAt(i)), "");
                    sl.Save();
                }
                catch (System.IO.IOException e)
                {
                    res = e;
                    Console.WriteLine("el archivo \"" + archivos.ElementAt(i) + ".xlsx\" no existe o se encuentra abierto");
                    Console.ReadKey();
                }
            return res == null;
        }
    }
}
