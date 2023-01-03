using Microsoft.AspNetCore.Mvc;
using ReadImportSQL_ExcelFile.Models;
using System.Diagnostics;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using ReadImportSQL_ExcelFile.Models.ViewModels;
using EFCore.BulkExtensions;
using Microsoft.EntityFrameworkCore;

namespace ReadImportSQL_ExcelFile.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly ImportExcelDBContext _dbContext;

        public HomeController(ILogger<HomeController> logger, ImportExcelDBContext context )
        {
            _logger = logger;
            _dbContext = context;
        }




        public IActionResult Index()
        {
            return View();
        }


        private ISheet ProcesaExcel(IFormFile formFile)
        {
            Stream stream = formFile.OpenReadStream();
            IWorkbook MiExcel = Path.GetExtension(formFile.FileName).ToUpper() == ".XLSX" ? new XSSFWorkbook(stream) : new HSSFWorkbook(stream);
            ISheet HojaExcel = MiExcel.GetSheetAt(0);
            return HojaExcel;
        }


        [HttpPost]
        public IActionResult MostrarDatos([FromForm] IFormFile ArchivoExcel)
        {
            var HojaExcel = ProcesaExcel(ArchivoExcel);
            int cantidadFilas = HojaExcel.LastRowNum;

            List<ContactoVM> records = new List<ContactoVM>();
            for (int i = 1; i <= cantidadFilas; i++)
            {
                IRow fila = HojaExcel.GetRow(i);
                records.Add(new ContactoVM
                {
                    Nombre = fila.GetCell(0).ToString(),
                    Apellido = fila.GetCell(1).ToString(),
                    Telefono = fila.GetCell(2).ToString(),
                    Email = fila.GetCell(3).ToString(),
                });
            }
            return StatusCode(StatusCodes.Status200OK, records);
        }



        [HttpPost]
        public IActionResult EnviarDatosDB([FromForm] IFormFile ArchivoExcel)
        {


            Stream stream = ArchivoExcel.OpenReadStream();

            IWorkbook MiExcel = null;

            if (Path.GetExtension(ArchivoExcel.FileName) == ".xlsx")
            {
                MiExcel = new XSSFWorkbook(stream);
            }
            else
            {
                MiExcel = new HSSFWorkbook(stream);
            }

            ISheet HojaExcel = MiExcel.GetSheetAt(0);

            int cantidadFilas = HojaExcel.LastRowNum;

            List<Contacto> lista = new List<Contacto>();

            for (int i = 1; i <= cantidadFilas; i++)
            {
                IRow fila = HojaExcel.GetRow(i);

                lista.Add(new Contacto
                {
                    Nombre = fila.GetCell(0).ToString(),
                    Apellido = fila.GetCell(1).ToString(),
                    Telefono = fila.GetCell(2).ToString(),
                    Correo = fila.GetCell(3).ToString(),
                });
            }

            _dbContext.BulkInsert(lista);
            return StatusCode(StatusCodes.Status200OK, new { mensaje = "OK" });
            
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