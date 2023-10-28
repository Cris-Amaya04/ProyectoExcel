using Microsoft.AspNetCore.Mvc;
using ProyectoExcel.Models;
using System.Diagnostics;
//librerias que permiten leer un tipo de archivo Excel
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using ProyectoExcel.Models.ViewModels;
using EFCore.BulkExtensions;

namespace ProyectoExcel.Controllers
{
    public class HomeController : Controller
    {
        private readonly BDExcelContext _context;

        public HomeController(BDExcelContext context)
        {
            _context = context; 
        }

        public IActionResult Index()
        {
            return View();
        }

        //Método que recibirá el archivo excel, lee la información y lo devuelve en formato de lista
        [HttpPost]
        public IActionResult MostrarDatos([FromForm] IFormFile archivoExcel)
        {

            //Leer el archivo Excel en forma de memoria
            Stream stream = archivoExcel.OpenReadStream();

            //Abre el archivo Excel
            IWorkbook miExcel = null;

            if(Path.GetExtension(archivoExcel.FileName) == ".xlsx")
            {
                //Archivos del formato xlsx serán leídos con este recurso
                miExcel = new XSSFWorkbook(stream);
            }
            else
            {
                miExcel = new HSSFWorkbook(stream);
            }

            //Accede al archivo de miExcel y obtiene solo la primera hoja
            ISheet hojaExcel = miExcel.GetSheetAt(0);

            //obtiene la cantidad de filas de la hoja de Excel
            int cantidadFilas = hojaExcel.LastRowNum;

            //Crear lista del viewModel creado, ya que ahí se almacenará toda la información que viene del Excel
            List<VMContacto> lista = new List<VMContacto>();

            //Realiza el recorrido de cada una de las filas que se encuentran en la hoja, comienza en 1 porque no se quiere leer la información de las columnas
            for(int i = 1; i <= cantidadFilas; i++)
            {
                //Devuelve una fila por cada iteración
                IRow fila = hojaExcel.GetRow(i);

                lista.Add(new VMContacto
                {
                    //Celda es la columna
                    Nombre =fila.GetCell(0).ToString(),
                    Apellido = fila.GetCell(1).ToString(),
                    Telefono = fila.GetCell(2).ToString(),
                    Correo = fila.GetCell(3).ToString()
                });
            }

            return StatusCode(StatusCodes.Status200OK, lista);
        }


        [HttpPost]
        public IActionResult EnviarDatos([FromForm] IFormFile archivoExcel)
        {
            //Leer el archivo Excel en forma de memoria
            Stream stream = archivoExcel.OpenReadStream();

            //Abre el archivo Excel
            IWorkbook miExcel = null;

            if (Path.GetExtension(archivoExcel.FileName) == ".xlsx")
            {
                //Archivos del formato xlsx serán leídos con este recurso
                miExcel = new XSSFWorkbook(stream);
            }
            else
            {
                miExcel = new HSSFWorkbook(stream);
            }

            //Accede al archivo de miExcel y obtiene solo la primera hoja
            ISheet hojaExcel = miExcel.GetSheetAt(0);

            //obtiene la cantidad de filas de la hoja de Excel
            int cantidadFilas = hojaExcel.LastRowNum;

            //Crear lista de la clase creada con el Scaffolding, que es la información que contiene la tabla de la base de datos
            List<Contacto> lista = new List<Contacto>();

            //Realiza el recorrido de cada una de las filas que se encuentran en la hoja, comienza en 1 porque no se quiere leer la información de las columnas
            for (int i = 1; i <= cantidadFilas; i++)
            {
                //Devuelve una fila por cada iteración
                IRow fila = hojaExcel.GetRow(i);

                lista.Add(new Contacto
                {
                    //Celda es la columna
                    Nombre = fila.GetCell(0).ToString(),
                    Apellido = fila.GetCell(1).ToString(),
                    Telefono = fila.GetCell(2).ToString(),
                    Correo = fila.GetCell(3).ToString()
                });
            }

            //Permite utilizar la función de enviar o registrar listas o información masiva de una tabla
            _context.BulkInsert(lista);

            return StatusCode(StatusCodes.Status200OK, new { mensaje = "ok"});
        }


        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}