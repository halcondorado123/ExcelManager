using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Reflection.Metadata;
using System.Reflection.PortableExecutable;
using System.Xml.Linq;
using System.Security.Cryptography.Xml;
using System.Collections.Generic;
using System.Data;
using System.Reflection.Emit;
using System.Linq;
using ExcelManager.Models;
using System.ComponentModel;


namespace PruebaExcel01.Controllers
{

    // Implementar bloqueo de la hoja de Excel para evitar que la hoja se modifique, alterando la información del estudiante
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
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

        public ActionResult ExportToExcel()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Exportar()
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {

                var fechaHoraActual = DateTime.Now;

                string formatoFecha = "yyyyMMdd_HHmmss";

                #region infoDownloadDoc

                // Crea el nombre de mi documento 
                // Si llegaste hasta el final, modificar la manera en la que se va a llamar el documento: acta_[numeroCedula]

                string nombreArchivo = "Datos_Docs_Est_" + fechaHoraActual.ToString(formatoFecha) + ".xlsx";
                string nombreHoja = "docs_titulos_01" + fechaHoraActual;


                // Crear Hoja xlsx

                #endregion

                #region Cellsizes
                var worksheet = package.Workbook.Worksheets.Add("SummaryDocs");

                ApplyBackgroundColorToRange(worksheet, 1, 1, 200, 200, System.Drawing.Color.White);

                int[] columnIndices = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10,
                                       11, 12, 13, 14, 15, 16, 17, 18, 19, 20,
                                       21, 22, 23, 24};

                double[] columnWidths = { 17.00, 35.00, 10.00, 10.00, 15.00, 15.00, 14.89, 11.00, 35.00, 16.89,
                                        13.00, 15.00, 15.00, 14.22, 11.00, 35.00, 14.00, 14.00, 15.00, 15.00,
                                        18.00, 16.56, 14.89, 13.56,};

                for (int i = 0; i < columnIndices.Length; i++)
                {
                    int columnIndex = columnIndices[i];
                    double columnWidth = columnWidths[i];
                    worksheet.Column(columnIndex).Width = columnWidth;
                }

                int[] rowIndices = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10,
                                    11, 12, 13, 14, 15, 16, 17, 18, 19, 20,
                                    21, 22, 23, 24, 25, 26, 27, 28, 29, 30,
                                    31, 32, 33, 34, 35, 36, 37, 38, 39, 40,
                                    41, 42, 43, 44, 45, 46, 47, 48, 49, 50,
                                    51, 52, 53, 54, 55, 56, 57, 58, 59, 60,
                                    61, 62, 63, 64, 65, 66, 67, 68, 69, 70,
                                    71, 72, 73, 74, 75, 76, 77, 78, 79, 80,
                                    81, 82, 83, 84, 85, 86, 87, 88, 89, 90,
                                    91, 92, 93, 94, 95, 96, 97, 98, 99, 100,
                                    101, 102, 103, 104, 105, 106, 107, 108, 109, 110
                };

                double[] rowHeights = {12.60, 13.20, 12.60, 12.60, 25.20, 12.60, 12.60, 25.20, 12.60, 38.40,
                                        37.20, 12.60, 12.60, 21.00, 24.00, 15.60, 12.60, 33.60, 12.60, 25.20,
                                        12.60, 12.60, 12.60, 12.60, 12.60, 12.60, 12.60, 12.60, 12.60, 12.60,
                                        51.00, 12.60, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00,
                                        26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00,
                                        26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00,
                                        26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00,
                                        26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00,
                                        26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00,
                                        26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00,
                                        26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00, 26.00};

                for (int i = 0; i < rowIndices.Length; i++)
                {
                    int rowIndex = rowIndices[i];
                    double rowHeight = rowHeights[i];
                    worksheet.Row(rowIndex).Height = rowHeight;
                }

                #endregion

                #region DocumentContent

                List<SubjectsME> subjects = new List<SubjectsME>();

                AddLogo(worksheet);
                AddHeaderInfo(worksheet);

                AddFormUserContent(worksheet);
                AddSenaStructureInformation(worksheet);
                ConstructionHeaderTable(worksheet);
                ContentTable(worksheet);


                #endregion


                var filePath = @"C:\Users\Jhonattan_Casallas\Downloads\" + nombreArchivo;
                //var filePath = @"C:\Users\Usuario\Downloads\" + nombreArchivo;
                package.SaveAs(new System.IO.FileInfo(filePath));

                return File(filePath, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", nombreArchivo);

            }
        }


        // Propiedades generales del texto
        public static void ContentCenter(ExcelRange range)
        {

            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            range.Style.WrapText = true;
        }

        // DerechaContenido
        public static void ContentRight(ExcelRange range)
        {
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            range.Style.WrapText = true;
        }

        // IzquierdaContenido
        public static void ContentLeft(ExcelRange range)
        {
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            range.Style.WrapText = true;
        }

        // Aplicar Negrilla
        public static void FontWeightBold(ExcelRange range, bool bold = true)
        {
            range.Style.Font.Bold = bold;
        }

        public static void WhiteColor(ExcelRange range, ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin)
        {
            range.Style.Font.Color.SetColor(System.Drawing.Color.White);
        }

        // Propiedades de celda

        // Bordes Tabla (4 bordes)
        public static void ApplyBorders(ExcelRange range, ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin)
        {
            range.Style.Border.Top.Style = borderStyle;
            range.Style.Border.Left.Style = borderStyle;
            range.Style.Border.Right.Style = borderStyle;
            range.Style.Border.Bottom.Style = borderStyle;
        }

        // Bordes Tipo Firma (Borde inferior)
        public static void ApplySignatureBorders(ExcelRange range, ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin)
        {
            range.Style.Border.Bottom.Style = borderStyle;
        }

        // Unir celdas

        public static void MergedCellsHorizontally(ExcelWorksheet worksheet, int startRow, int endRow, int startColumn, int endColumn, bool merge = true)
        {
            for (int row = startRow; row <= endRow; row++)
            {
                ExcelRange rangeToMerge = worksheet.Cells[row, startColumn, row, endColumn];
                rangeToMerge.Style.WrapText = true;
                rangeToMerge.Merge = merge;
            }
        }
        public static void MergedCells(ExcelRange range, bool merge = true)
        {
            range.Merge = merge;
        }

        public static void WrapText(ExcelRange range)
        {
            range.Style.WrapText = true;
        }


        // Aplicar bordes + Centrado Texto
        public static void CellCenter(ExcelRange range, ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin)
        {
            ApplyBorders(range);
            ContentCenter(range);
            WrapText(range);
        }

        // Aplicar bordes + Texto Derecha
        public static void CellRight(ExcelRange range, ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin)
        {
            ApplyBorders(range);
            ContentRight(range);
            WrapText(range);
        }

        // Aplicar bordes + Texto Izquierda
        public static void CellLeft(ExcelRange range, ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin)
        {
            ApplyBorders(range);
            ContentLeft(range);
            WrapText(range);
        }
        public static void ApplyBackgroundColorToRange(ExcelWorksheet worksheet, int startRow, int startColumn, int endRow, int endColumn, System.Drawing.Color color)
        {
            var range = worksheet.Cells[startRow, startColumn, endRow, endColumn];
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(color);
        }

        #region GetCellAddition


        private ExcelRange GetTotalSubjects(ExcelWorksheet worksheet, int row, int column)
        {
            return worksheet.Cells[row, column];
        }


        #endregion

        private void AssignSubjectsToWorksheet(ExcelWorksheet worksheet, int row, int column, SubjectsME subject)
        {
            worksheet.Cells[row, column].Value = subject.Numero[0];
            worksheet.Cells[row, column + 1].Value = subject.Asignatura[0];
            worksheet.Cells[row, column + 2].Value = subject.Creditos[0];
            worksheet.Cells[row, column + 3].Value = subject.Semestre[0];
            worksheet.Cells[row, column + 4].Value = subject.CalificacionNumerica[0];
            worksheet.Cells[row, column + 5].Value = subject.CalificacionLiteral[0];
            worksheet.Cells[row, column + 6].Value = subject.Nivel[0];
        }

        // Seleccion de celdas por rango
        public static ExcelRange GetExcelRange(ExcelWorksheet worksheet, int startRow, int startColumn, int endRow, int endColumn)
        {
            return worksheet.Cells[startRow, startColumn, endRow, endColumn];
        }


        #region ImagesDocument
        private void AddLogo(ExcelWorksheet worksheet)
        {
            string imagePathLogo = "C:\\Users\\Jhonattan_Casallas\\Desktop\\EnsayoExcelTecnologo\\PruebaExcel_Version02\\Img_sample\\log1.png";
            //string imagePathLogo = "C:\\Users\\Usuario\\Desktop\\PruebaExcel01\\Img_sample\\log1.png";
            int widthLogoInPixels = 300;
            int heightLogoInPixels = 110;

            var pictureLogo = worksheet.Drawings.AddPicture("Logo", new FileInfo(imagePathLogo));

            pictureLogo.SetPosition(1, -20, 1, -50);
            pictureLogo.SetSize(widthLogoInPixels, heightLogoInPixels);
            pictureLogo.Locked = true;
        }

        private void SignaMark(ExcelWorksheet worksheet, int startRow, int startColumn, int endRow, int endColumn)
        {
            ExcelRange signatureMark = GetExcelRange(worksheet, startRow, startColumn, endRow, endColumn);
            ApplySignatureBorders(signatureMark);
            MergedCells(signatureMark);
        }

        private void ProgramManagerSignature(ExcelWorksheet worksheet, int row, int column)
        {
            string imagePathSignature = "C:\\Users\\Jhonattan_Casallas\\Desktop\\EnsayoExcelTecnologo\\PruebaExcel_Version02\\Img_sample\\lennon_signature.jpg";
            int widthSignatureInPixels = 230;
            int heightSignatureInPixels = 70;

            // Agregar imagen
            var pictureSignature = worksheet.Drawings.AddPicture("Firma", new FileInfo(imagePathSignature));

            // Establecer tamaño
            pictureSignature.SetSize(widthSignatureInPixels, heightSignatureInPixels);

            pictureSignature.SetPosition(row, column);

            // Bloquear la imagen si es necesario
            pictureSignature.Locked = true;
        }

        //private void StudentSignature(ExcelWorksheet worksheet)
        //{
        //    // Valor quemado, por evaluar la opcion de firma generada por parte del estudiante
        //    string imagePathSignature = "C:\\Users\\Jhonattan_Casallas\\Desktop\\EnsayoExcelTecnologo\\PruebaExcel_Version02\\Img_sample\\2560px-Freddie_Mercury_signature.svg.png";
        //    //string imagePathSignature = "C:\\Users\\Usuario\\Desktop\\PruebaExcel01\\Img_sample\\2560px-Freddie_Mercury_signature.svg.png";
        //    int widthSignatureInPixels = 230;
        //    int heightSignatureInPixels = 70;

        //    var pictureSignature = worksheet.Drawings.AddPicture("FirmaStu", new FileInfo(imagePathSignature));

        //    pictureSignature.SetPosition(73, -80, 8, -30);
        //    pictureSignature.SetSize(widthSignatureInPixels, heightSignatureInPixels);
        //    pictureSignature.Locked = true;

        //}

        #endregion



        private void AddHeaderInfo(ExcelWorksheet worksheet)
        {
            string fechaRegistro = "Fecha Registro / Hora Registro";
            DateTime fechaHoraActual = DateTime.Now;

            ExcelRange logoSpace = GetExcelRange(worksheet, 1, 1, 1, 2);
            logoSpace.Value = fechaRegistro;
            MergedCells(logoSpace);
            ContentCenter(logoSpace);
            WhiteColor(logoSpace);

            ExcelRange logoDate = GetExcelRange(worksheet, 2, 1, 5, 2);
            logoDate.Value = fechaHoraActual.ToString("yyyy-MM-dd \n" + fechaHoraActual.ToString("HH:mm:ss"));
            MergedCells(logoDate);
            ContentCenter(logoDate);
            WhiteColor(logoDate);

            ExcelRange headerRectangle = GetExcelRange(worksheet, 1, 3, 3, 21);
            headerRectangle.Value = "ACTA DE RECONOCIMIENTO DE TITULO";
            CellCenter(headerRectangle);
            MergedCells(headerRectangle);
            FontWeightBold(headerRectangle);
        }

        private void AddFormUserContent(ExcelWorksheet worksheet)
        {
            ExcelRange formMainUser = GetExcelRange(worksheet, 4, 1, 20, 21);
            ContentCenter(formMainUser);
            WrapText(formMainUser);

            #region FirstDivision

            var labelFormA = worksheet.Cells[5, 6];
            labelFormA.Value = "INTERNA";
            MergedCells(labelFormA);
            FontWeightBold(labelFormA);
            ContentCenter(labelFormA);

            var txtbox1 = worksheet.Cells[5, 7];
            ApplyBorders(txtbox1);
            // TARGET //


            var labelFormB = worksheet.Cells[5, 11];
            labelFormB.Value = "EXTERNA";
            MergedCells(labelFormB);
            FontWeightBold(labelFormB);
            ContentCenter(labelFormB);

            var txtbox2 = worksheet.Cells[5, 12];
            ApplyBorders(txtbox2);
            // TARGET //               

            var labelFormC = worksheet.Cells[5, 16];
            labelFormC.Value = "PERIODO ACADÉMICO";
            MergedCells(labelFormC);
            FontWeightBold(labelFormC);
            ContentCenter(labelFormC);

            var txtbox3 = worksheet.Cells[5, 17];
            ApplyBorders(txtbox3);
            // TARGET //               


            #endregion

            #region SecondDivision

            ExcelRange labelFormD = GetExcelRange(worksheet, 8, 3, 8, 4);
            labelFormD.Value = "MODALIDAD:";
            MergedCells(labelFormD);
            FontWeightBold(labelFormD);
            ContentCenter(labelFormD);

            var labelFormE = worksheet.Cells[8, 6];
            labelFormE.Value = "DISTANCIA";
            MergedCells(labelFormE);
            FontWeightBold(labelFormE);
            ContentCenter(labelFormE);

            var txtbox4 = worksheet.Cells[8, 7];
            ApplyBorders(txtbox4);
            // TARGET //               

            var labelFormF = worksheet.Cells[8, 11];
            labelFormF.Value = "PRESENCIAL";
            MergedCells(labelFormF);
            FontWeightBold(labelFormF);
            ContentCenter(labelFormF);

            var txtbox5 = worksheet.Cells[8, 12];
            ApplyBorders(txtbox5);
            // TARGET //               

            var labelFormG = worksheet.Cells[8, 16];
            labelFormG.Value = "VIRTUAL";
            MergedCells(labelFormG);
            FontWeightBold(labelFormG);
            ContentCenter(labelFormG);

            var txtbox6 = worksheet.Cells[8, 17];
            ApplyBorders(txtbox6);
            // TARGET //               

            #endregion

            #region ThirdDivision

            var txtbox7 = worksheet.Cells[10, 2];
            ApplySignatureBorders(txtbox7);
            // TARGET               

            ExcelRange labelFormH = GetExcelRange(worksheet, 11, 2, 12, 2);
            labelFormH.Value = "REGIONAL - SEDE O CUNAD";
            MergedCells(labelFormH);
            FontWeightBold(labelFormH);
            ContentCenter(labelFormH);

            var txtbox8 = worksheet.Cells[10, 9];
            ApplySignatureBorders(txtbox8);
            // TARGET //

            ExcelRange labelFormI = GetExcelRange(worksheet, 11, 9, 12, 9);
            labelFormI.Value = "FECHA DEL RECONOCIMIENTO";
            MergedCells(labelFormI);
            FontWeightBold(labelFormI);
            ContentCenter(labelFormI);

            var txtbox9 = worksheet.Cells[10, 16];
            ApplySignatureBorders(txtbox9);
            // TARGET //               

            ExcelRange labelFormJ = GetExcelRange(worksheet, 11, 16, 12, 16);
            labelFormJ.Value = "PLAN DE ESTUDIOS A APLICAR";
            MergedCells(labelFormJ);
            FontWeightBold(labelFormJ);
            ContentCenter(labelFormJ);

            var txtbox10 = worksheet.Cells[10, 21];
            ApplySignatureBorders(txtbox10);
            // TARGET //               

            ExcelRange labelFormK = GetExcelRange(worksheet, 11, 21, 12, 21);
            labelFormK.Value = "CÓDIGO DEL PLAN DE ESTUDIOS A APLICAR";
            MergedCells(labelFormK);
            FontWeightBold(labelFormK);
            ContentCenter(labelFormK);

            #endregion

            #region FourthDivision

            var txtbox11 = worksheet.Cells[14, 2];
            ApplySignatureBorders(txtbox11);
            // TARGET //               

            ExcelRange labelFormL = GetExcelRange(worksheet, 15, 2, 16, 2);
            labelFormL.Value = "APELLIDOS Y NOMBRES DEL ESTUDIANTE";
            MergedCells(labelFormL);
            FontWeightBold(labelFormL);
            ContentCenter(labelFormL);

            var txtbox12 = worksheet.Cells[14, 9];
            ApplySignatureBorders(txtbox12);
            // TARGET //               

            ExcelRange labelFormM = GetExcelRange(worksheet, 15, 9, 16, 9);
            labelFormM.Value = "DOCUMENTO DE IDENTIDAD";
            MergedCells(labelFormM);
            FontWeightBold(labelFormM);
            ContentCenter(labelFormM);

            var txtbox13 = worksheet.Cells[14, 16];
            ApplySignatureBorders(txtbox13);
            // TARGET //               

            ExcelRange labelFormO = GetExcelRange(worksheet, 15, 16, 16, 16);
            labelFormO.Value = "CORREO ELECTRONICO";
            MergedCells(labelFormO);
            FontWeightBold(labelFormO);
            ContentCenter(labelFormO);

            var txtbox14 = worksheet.Cells[14, 21];
            ApplySignatureBorders(txtbox14);
            // TARGET //               

            ExcelRange labelFormP = GetExcelRange(worksheet, 15, 21, 16, 21);
            labelFormP.Value = "TELEFONO FIJO - CELULAR";
            MergedCells(labelFormP);
            FontWeightBold(labelFormP);
            ContentCenter(labelFormP);


            #endregion

            #region FifthDivision

            var txtbox15 = worksheet.Cells[18, 2];
            ApplySignatureBorders(txtbox15);
            // TARGET //               

            ExcelRange labelFormQ = GetExcelRange(worksheet, 19, 2, 20, 2);
            labelFormQ.Value = "INSTITUCIÓN DE DONDE PROVIENE";
            MergedCells(labelFormQ);
            FontWeightBold(labelFormQ);
            ContentCenter(labelFormQ);

            var txtbox16 = worksheet.Cells[18, 9];
            ApplySignatureBorders(txtbox16);
            // TARGET //               

            ExcelRange labelFormR = GetExcelRange(worksheet, 19, 9, 20, 9);
            labelFormR.Value = "PROGRAMA CURSADO";
            MergedCells(labelFormR);
            FontWeightBold(labelFormR);
            ContentCenter(labelFormR);

            var txtbox17 = worksheet.Cells[18, 16];
            ApplySignatureBorders(txtbox17);
            // TARGET //               

            ExcelRange labelFormS = GetExcelRange(worksheet, 19, 16, 20, 16);
            labelFormS.Value = "PROGRAMA A CURSAR";
            MergedCells(labelFormS);
            FontWeightBold(labelFormS);
            ContentCenter(labelFormS);

            #endregion
        }

        private void AddSenaStructureInformation(ExcelWorksheet worksheet)
        {
            var numbSena = worksheet.Cells[22, 1];
            int startRow = 22;
            for (int i = 1; i <= 7; i++)
            {
                var cell = worksheet.Cells[startRow + i - 1, 1];
                cell.Value = i;
                ContentRight(cell);
            }

            var numbSenaLastCell = worksheet.Cells[28, 1];
            ApplySignatureBorders(numbSenaLastCell);


            ExcelRange labelTitle1 = GetExcelRange(worksheet, 21, 1, 21, 21);
            labelTitle1.Value = "ESTRUCTURA CURRICULAR SENA";
            CellCenter(labelTitle1);
            MergedCells(labelTitle1);
            FontWeightBold(labelTitle1);

            ExcelRange context1 = GetExcelRange(worksheet, 22, 2, 22, 21);
            context1.Value = "ADMITIR AL USUARIO EN LA RED DE SERVICIOS DE SALUD SEGÚN NIVELES DE ATENCIÓN Y NORMATIVA VIGENTE.";
            ContentLeft(context1);
            MergedCells(context1);

            ExcelRange context2 = GetExcelRange(worksheet, 23, 2, 23, 21);
            context2.Value = "AFILIAR A LA POBLACIÓN AL SISTEMA GENERAL DE SEGURIDAD SOCIAL EN SALUD SEGÚN NORMATIVIDAD VIGENTE.";
            MergedCells(context2);
            ContentLeft(context2);

            ExcelRange context3 = GetExcelRange(worksheet, 24, 2, 24, 21);
            context3.Value = "FACTURAR LA PRESTACIÓN DE LOS SERVICIOS DE SALUD SEGÚN NORMATIVIDAD Y CONTRATACIÓN";
            MergedCells(context3);
            ContentLeft(context3);

            ExcelRange context4 = GetExcelRange(worksheet, 25, 2, 25, 21);
            context4.Value = "MANEJAR VALORES E INGRESOS RELACIONADOS CON LA OPERACIÓN DEL ESTABLECIMIENTO. (EQUIVALE A LA NORMA NTS 005 DEL MINCOMERCIO, INDUSTRIA Y TURISMO)";
            MergedCells(context4);
            ContentLeft(context4);

            ExcelRange context5 = GetExcelRange(worksheet, 26, 2, 26, 21);
            context5.Value = "ORIENTAR AL USUARIO EN RELACIÓN CON SUS NECESIDADES Y EXPECTATIVAS DE ACUERDO CON POLÍTICAS INSTITUCIONALES Y NORMAS DE SALUD VIGENTES.";
            MergedCells(context5);
            ContentLeft(context5);

            ExcelRange context6 = GetExcelRange(worksheet, 27, 2, 27, 21);
            context6.Value = "PROMOVER LA INTERACCION IDONEA CONSIGO MISMO, CON LOS DEMAS Y CON LA NATURALEZA EN LOS CONTEXTOS LABORAL Y SOCIAL.";
            MergedCells(context6);
            ContentLeft(context6);

            ExcelRange context7 = GetExcelRange(worksheet, 28, 2, 28, 21);
            context7.Value = "RESULTADOS DE APRENDIZAJE ETAPA PRACTICA";
            ApplySignatureBorders(context7);
            MergedCells(context7);
            ContentLeft(context7);

        }

        private void ConstructionHeaderTable(ExcelWorksheet worksheet)
        {
            ExcelRange HeaderBar = GetExcelRange(worksheet, 31, 1, 32, 21);
            CellCenter(HeaderBar);
            FontWeightBold(HeaderBar);

            #region CellsMainTable

            ExcelRange CellMT1 = GetExcelRange(worksheet, 31, 1, 32, 1);
            CellMT1.Value = "No";
            MergedCells(CellMT1);


            ExcelRange CellMT2 = GetExcelRange(worksheet, 31, 2, 32, 2);
            CellMT2.Value = "ASIGNATURA Y/O CRÉDITO HOMOLOGADO";
            MergedCells(CellMT2);

            ExcelRange CellMT3 = GetExcelRange(worksheet, 31, 3, 31, 4);
            CellMT3.Value = "SISTEMA";
            MergedCells(CellMT3);

            ExcelRange CellMT4 = GetExcelRange(worksheet, 31, 5, 32, 5);
            CellMT4.Value = "CALIFICACIÓN NUMERICA";
            MergedCells(CellMT4);

            ExcelRange CellMT5 = GetExcelRange(worksheet, 31, 6, 32, 6);
            CellMT5.Value = "CALIFICACION LITERAL";
            MergedCells(CellMT5);

            ExcelRange CellMT6 = GetExcelRange(worksheet, 31, 7, 32, 7);    // OK
            CellMT6.Value = "NIVEL";
            MergedCells(CellMT6);

            ExcelRange CellMT7 = GetExcelRange(worksheet, 31, 8, 32, 8);
            CellMT7.Value = "No";
            MergedCells(CellMT7);

            ExcelRange CellMT8 = GetExcelRange(worksheet, 31, 9, 32, 9);
            CellMT8.Value = "ASIGNATURA Y/O CRÉDITO HOMOLOGADO";
            MergedCells(CellMT8);

            ExcelRange CellMT9 = GetExcelRange(worksheet, 31, 10, 31, 11);
            CellMT9.Value = "SISTEMA";
            MergedCells(CellMT9);

            ExcelRange CellMT10 = GetExcelRange(worksheet, 31, 12, 32, 12);
            CellMT10.Value = "CALIFICACIÓN NUMERICA";
            MergedCells(CellMT10);

            ExcelRange CellMT11 = GetExcelRange(worksheet, 31, 13, 32, 13);
            CellMT11.Value = "CALIFICACION LITERAL";
            MergedCells(CellMT11);

            ExcelRange CellMT12 = GetExcelRange(worksheet, 31, 14, 32, 14);     // OK
            CellMT12.Value = "NIVEL";
            MergedCells(CellMT12);

            ExcelRange CellMT13 = GetExcelRange(worksheet, 31, 15, 32, 15);
            CellMT13.Value = "No";
            MergedCells(CellMT13);

            ExcelRange CellMT14 = GetExcelRange(worksheet, 31, 16, 32, 16);
            CellMT14.Value = "ASIGNATURA Y/O CRÉDITO HOMOLOGADO";
            MergedCells(CellMT14);

            ExcelRange CellMT15 = GetExcelRange(worksheet, 31, 17, 31, 18);
            CellMT15.Value = "SISTEMA";
            MergedCells(CellMT15);

            ExcelRange CellMT16 = GetExcelRange(worksheet, 31, 19, 32, 19);
            CellMT16.Value = "CALIFICACIÓN NUMERICA";
            MergedCells(CellMT16);

            ExcelRange CellMT17 = GetExcelRange(worksheet, 31, 20, 32, 20);
            CellMT17.Value = "CALIFICACION LITERAL";
            MergedCells(CellMT17);

            ExcelRange CellMT18 = GetExcelRange(worksheet, 31, 21, 32, 21);
            CellMT18.Value = "NIVEL";
            MergedCells(CellMT18);

            #endregion

            #region SubCellsMainTable

            var subCellMT1 = worksheet.Cells[32, 3];
            subCellMT1.Value = "Créditos";

            var subCellMT2 = worksheet.Cells[32, 4];
            subCellMT2.Value = "Semestre";

            var subCellMT3 = worksheet.Cells[32, 10];
            subCellMT3.Value = "Créditos";

            var subCellMT4 = worksheet.Cells[32, 11];
            subCellMT4.Value = "Semestre";

            var subCellMT5 = worksheet.Cells[32, 17];
            subCellMT5.Value = "Créditos";

            var subCellMT6 = worksheet.Cells[32, 18];
            subCellMT6.Value = "Semestre";

            #endregion

        }

        public static void ProcessScenario(ExcelWorksheet worksheet, int startRow, int endRow, int startColumn,
                                           int endColumn, string label, string sumRange, string sumResult)
        {
            ExcelRange celdasMaterias = GetExcelRange(worksheet, startRow, startColumn, endRow, endColumn);
            CellCenter(celdasMaterias);

            ExcelRange LabelTotals = GetExcelRange(worksheet, endRow + 1, startColumn, endRow + 1, startColumn + 1);
            LabelTotals.Value = label;
            MergedCells(LabelTotals);
            CellCenter(LabelTotals);
            FontWeightBold(LabelTotals);

            var rangeToSum = worksheet.Cells[sumRange];
            var result = worksheet.Cells[sumResult];
            result.Formula = $"SUM({rangeToSum.Address})";
            CellCenter(result);
            FontWeightBold(result);
        }

        public static void CellTotalSum(ExcelWorksheet worksheet, int startRow, int startColumn, int endRow, int endColumn)
        {
            ExcelRange cellTotalSubject = GetExcelRange(worksheet, startRow, startColumn, endRow, endColumn);
            MergedCells(cellTotalSubject);
            FontWeightBold(cellTotalSubject);
            CellCenter(cellTotalSubject);

        }

        public static void TotalSumColumns(ExcelWorksheet worksheet, int startRow, int endRow, int startColumn,
                                   int endColumn, string sumRange, string sumResult)
        {
            ExcelRange celdasMaterias = GetExcelRange(worksheet, startRow, startColumn, endRow, endColumn);
            var rangeToSum = worksheet.Cells[sumRange];
            var result = worksheet.Cells[sumResult];
            result.Formula = $"SUM({rangeToSum.Address})";
            ContentCenter(result);
            FontWeightBold(result);
        }






        public static void AddSingleParraf(ExcelWorksheet worksheet, int startRow, int endRow, int startColumn, int endColumn, string[] labels)
        {
            ExcelRange labelTotalCredits = GetExcelRange(worksheet, startRow, startColumn, endRow, endColumn);
            MergedCellsHorizontally(worksheet, startRow, endRow, startColumn, endColumn);

            for (int i = 0; i < labels.Length; i++)
            {

                ExcelRange labelTotals = GetExcelRange(worksheet, startRow + i, startColumn, endRow, endColumn);
                labelTotals.Value = labels[i];
                FontWeightBold(labelTotals);
                CellRight(labelTotals);
            }
        }

        public static void AddSingleParrafSignature(ExcelWorksheet worksheet, int row, int column, string[] labels)
        {
            for (int i = 0; i < labels.Length; i++)
            {
                var labelTotals = worksheet.Cells[row + i, column];
                labelTotals.Value = labels[i];
                ContentLeft(labelTotals);
            }
        }

        public static void AddTitleDocument(ExcelWorksheet worksheet, int startRow, int endRow, int startColumn, int endColumn, string labels)
        {
            ExcelRange titleCell = GetExcelRange(worksheet, startRow, endRow, startColumn, endColumn);
            titleCell.Value = labels;
            MergedCells(titleCell);
            ContentLeft(titleCell);
            FontWeightBold(titleCell);
        }


        public static void AddParagraph(ExcelWorksheet worksheet, int startRow, int endRow, int startColumn, int endColumn, List<string> conditionLabels)
        {
            ExcelRange labelTotalCredits = GetExcelRange(worksheet, startRow, startColumn, endRow, endColumn);
            MergedCellsHorizontally(worksheet, startRow, endRow, startColumn, endColumn);

            for (int i = 0; i < conditionLabels.Count; i++)
            {

                ExcelRange conditions = GetExcelRange(worksheet, startRow + i, startColumn, endRow, endColumn);
                conditions.Value = conditionLabels[i];
                ContentLeft(conditions);
            }
        }

        // Obtener la lista de asignaturas
        private void ContentTable(ExcelWorksheet worksheet)
        {

            List<SubjectsME> subjects = GetSubjects.SubjectGenerator();
            ContentText labelProvider = new ContentText();
            string[] labels = labelProvider.GetLabelsTotalCredits();
            string[] programManagerInfo = labelProvider.AddProgramManagerInfo();
            //string[] programStudentInfo = labelProvider.AddProgramStudentInfo();
            string condition = labelProvider.TitleCondition();
            string impCondition = labelProvider.TitleImportantCondition();
            string constancy = labelProvider.TextInConstancy();
            List<string> conditionLabels = labelProvider.ConditionLabels();
            List<string> importantCondition = labelProvider.ImportantConditionLabels();


            int row = 33;
            int column = 1;
            int subjectCount = 0;
            int totalSubjects = subjects.Count;
            int SubjectsPerColumn = 10;



            if (totalSubjects <= 30)
            {
                SubjectsPerColumn = 10;
                ExcelRange cellsTiny = GetExcelRange(worksheet, 33, 1, 42, 21);
                CellCenter(cellsTiny);
                ProcessScenario(worksheet, 33, 42, 1, 2, "TOTALES", "C33:C42", "C43");
                ProcessScenario(worksheet, 33, 42, 1, 2, "TOTALES", "J33:J42", "J43");
                ProcessScenario(worksheet, 33, 42, 1, 2, "TOTALES", "Q33:Q42", "Q43");
                TotalSumColumns(worksheet, 43, 43, 3, 15, "C43:Q43", "E46");

                BorderTable(worksheet, 47, 5, 50, 7);
                TotalCreditsLabels(worksheet, 47, 5, 47, 6, 47, 7);
                CellTotalSum(worksheet, 46, 5, 46, 7);
                ZeroDefault(worksheet, 48, 5, 50, 7);

                AddSingleParraf(worksheet, 48, 50, 1, 4, labels);
                AddTitleDocument(worksheet, 52, 1, 52, 4, condition);
                AddParagraph(worksheet, 53, 56, 1, 21, conditionLabels);

                AddTitleDocument(worksheet, 58, 1, 58, 4, impCondition);
                AddParagraph(worksheet, 59, 61, 1, 21, importantCondition);

                AddTitleDocument(worksheet, 64, 1, 64, 4, constancy);

                SignaMark(worksheet, 68, 1, 68, 2);
                ProgramManagerSignature(worksheet, 96 + 1800, 1);
                AddSingleParrafSignature(worksheet, 69, 1, programManagerInfo);
                //AddSingleParrafSignature(worksheet, 69, 10, programStudentInfo);
            }
            else if (totalSubjects > 30 && totalSubjects <= 48)
            {
                SubjectsPerColumn = 16;
                ExcelRange cellsMedium = GetExcelRange(worksheet, 33, 1, 48, 21);
                CellCenter(cellsMedium);
                ProcessScenario(worksheet, 33, 48, 1, 2, "TOTALES", "C33:C48", "C49");
                ProcessScenario(worksheet, 33, 48, 1, 2, "TOTALES", "J33:J48", "J49");
                ProcessScenario(worksheet, 33, 48, 1, 2, "TOTALES", "Q33:Q48", "Q49");
                TotalSumColumns(worksheet, 73, 73, 3, 15, "C49:Q49", "E51");

                BorderTable(worksheet, 52, 5, 55, 7);
                TotalCreditsLabels(worksheet, 52, 5, 52, 6, 52, 7);
                CellTotalSum(worksheet, 51, 5, 51, 7);
                ZeroDefault(worksheet, 53, 5, 55, 7);

                AddSingleParraf(worksheet, 53, 55, 1, 4, labels);
                AddTitleDocument(worksheet, 57, 1, 57, 4, condition);
                AddParagraph(worksheet, 58, 61, 1, 21, conditionLabels);

                AddTitleDocument(worksheet, 63, 1, 63, 4, impCondition);
                AddParagraph(worksheet, 64, 66, 1, 21, importantCondition);

                AddTitleDocument(worksheet, 69, 1, 69, 4, constancy);

                SignaMark(worksheet, 73, 1, 73, 2);
                ProgramManagerSignature(worksheet, 96 + 1950, 1);
                AddSingleParrafSignature(worksheet, 74, 1, programManagerInfo);
                //AddSingleParrafSignature(worksheet, 74, 10, programStudentInfo);

            }
            else if (totalSubjects > 48 && totalSubjects <= 72)
            {
                SubjectsPerColumn = 24;

                ExcelRange cellsBig = GetExcelRange(worksheet, 33, 1, 56, 21);
                CellCenter(cellsBig);
                ProcessScenario(worksheet, 33, 56, 1, 2, "TOTALES", "C33:C56", "C57");
                ProcessScenario(worksheet, 33, 56, 1, 2, "TOTALES", "J33:J56", "J57");
                ProcessScenario(worksheet, 33, 56, 1, 2, "TOTALES", "Q33:Q56", "Q57");
                TotalSumColumns(worksheet, 57, 57, 3, 15, "C57:Q57", "E60");

                BorderTable(worksheet, 61, 5, 64, 7);
                TotalCreditsLabels(worksheet, 61, 5, 61, 6, 61, 7);
                CellTotalSum(worksheet, 60, 5, 60, 7);
                ZeroDefault(worksheet, 62, 5, 64, 7);

                AddSingleParraf(worksheet, 62, 64, 1, 4, labels);
                AddTitleDocument(worksheet, 66, 1, 66, 4, condition);
                AddParagraph(worksheet, 67, 70, 1, 21, conditionLabels);

                AddTitleDocument(worksheet, 72, 1, 72, 4, impCondition);
                AddParagraph(worksheet, 73, 75, 1, 21, importantCondition);

                AddTitleDocument(worksheet, 78, 1, 78, 4, constancy);

                SignaMark(worksheet, 82, 1, 82, 2);
                ProgramManagerSignature(worksheet, 96 + 2270, 1);
                AddSingleParrafSignature(worksheet, 83, 1, programManagerInfo);
                //AddSingleParrafSignature(worksheet, 83, 10, programStudentInfo);
            }
            else if (totalSubjects > 72 && totalSubjects <= 90)
            {
                SubjectsPerColumn = 30; // Cambiar el límite si hay entre 51 y 100 arreglos

                ExcelRange cellsExtraBig = GetExcelRange(worksheet, 33, 1, 62, 21);
                CellCenter(cellsExtraBig);
                ProcessScenario(worksheet, 33, 62, 1, 2, "TOTALES", "C33:C62", "C63");
                ProcessScenario(worksheet, 33, 62, 1, 2, "TOTALES", "J33:J62", "J63");
                ProcessScenario(worksheet, 33, 62, 1, 2, "TOTALES", "Q33:Q62", "Q63");
                TotalSumColumns(worksheet, 63, 63, 3, 15, "C63:Q63", "E66");

                BorderTable(worksheet, 67, 5, 70, 7);
                TotalCreditsLabels(worksheet, 67, 5, 67, 6, 67, 7);
                CellTotalSum(worksheet, 66, 5, 66, 7);
                ZeroDefault(worksheet, 68, 5, 70, 7);

                AddSingleParraf(worksheet, 68, 70, 1, 4, labels);
                AddTitleDocument(worksheet, 72, 1, 72, 4, condition);
                AddParagraph(worksheet, 73, 76, 1, 21, conditionLabels);

                AddTitleDocument(worksheet, 78, 1, 78, 4, impCondition);
                AddParagraph(worksheet, 79, 81, 1, 21, importantCondition);

                AddTitleDocument(worksheet, 84, 1, 84, 4, constancy);

                SignaMark(worksheet, 88, 1, 88, 2);
                ProgramManagerSignature(worksheet, 96 + 2470, 1);
                AddSingleParrafSignature(worksheet, 89, 1, programManagerInfo);
                //AddSingleParrafSignature(worksheet, 89, 10, programStudentInfo);

            }
            else if (totalSubjects > 90 && totalSubjects <= 120)
            {
                SubjectsPerColumn = 40;
                ExcelRange cellsMaximum = GetExcelRange(worksheet, 33, 1, 72, 21);
                CellCenter(cellsMaximum);
                ProcessScenario(worksheet, 33, 72, 1, 2, "TOTALES", "C33:C72", "C73");
                ProcessScenario(worksheet, 33, 72, 1, 2, "TOTALES", "J33:J72", "J73");
                ProcessScenario(worksheet, 33, 72, 1, 2, "TOTALES", "Q33:Q72", "Q73");
                TotalSumColumns(worksheet, 73, 73, 3, 15, "C73:Q73", "E76");

                BorderTable(worksheet, 77, 5, 80, 7);
                TotalCreditsLabels(worksheet, 77, 5, 77, 6, 77, 7);
                CellTotalSum(worksheet, 76, 5, 76, 7);
                ZeroDefault(worksheet, 78, 5, 80, 7);

                AddSingleParraf(worksheet, 78, 80, 1, 4, labels);
                AddTitleDocument(worksheet, 82, 1, 82, 4, condition);
                AddParagraph(worksheet, 83, 86, 1, 21, conditionLabels);

                AddTitleDocument(worksheet, 88, 1, 88, 4, impCondition);
                AddParagraph(worksheet, 89, 91, 1, 21, importantCondition);

                AddTitleDocument(worksheet, 94, 1, 94, 4, constancy);

                SignaMark(worksheet, 98, 1, 98, 2);
                ProgramManagerSignature(worksheet, 96 + 2815, 1);
                AddSingleParrafSignature(worksheet, 99, 1, programManagerInfo);
                //AddSingleParrafSignature(worksheet, 99, 10, programStudentInfo);
            }


            foreach (var subject in subjects)
            {
                if (subjectCount >= SubjectsPerColumn)
                {
                    column += 7;
                    row = 33;
                    subjectCount = 0;
                }

                AssignSubjectsToWorksheet(worksheet, row, column, subject);

                row++;
                subjectCount++;
            }
        }

        private void BorderTable(ExcelWorksheet worksheet, int startRow, int startColumn, int endRow, int endColumn)
        {
            ExcelRange totalValuesSubjects = worksheet.Cells[startRow, startColumn, endRow, endColumn];
            CellCenter(totalValuesSubjects);
        }

        // int rowTotal, int columnTotal

        private void TotalCreditsLabels(ExcelWorksheet worksheet, int row1, int column1, int row2, int column2, int row3, int column3)
        {

            ContentText contentText = new ContentText();

            var cellAproLabel = worksheet.Cells[row1, column1];
            cellAproLabel.Value = contentText.TotalCreditsValue()[0];
            FontWeightBold(cellAproLabel);

            var cellPendingLabel = worksheet.Cells[row2, column2];
            cellPendingLabel.Value = contentText.TotalCreditsValue()[1];
            FontWeightBold(cellPendingLabel);

            var cellTotalCredLabel = worksheet.Cells[row3, column3];
            cellTotalCredLabel.Value = contentText.TotalCreditsValue()[2];
            FontWeightBold(cellTotalCredLabel);

        }

        private void ZeroDefault(ExcelWorksheet worksheet, int startRow, int startColumn, int endRow, int endColumn)
        {
            for (int row = startRow; row <= endRow; row++)
            {
                for (int column = startColumn; column <= endColumn; column++)
                {
                    var cell = worksheet.Cells[row, column];
                    cell.Value = 0;
                }
            }
        }



        //    ExcelRange celdasMateriasTotales = GetExcelRange(worksheet, 53, 5, 55, 7);
        //    CellCenter(celdasMateriasTotales);

        //    // IMPORTANTE: CODIGO PROPENSO A SER MODIFICADO //

        //    var cellsToSetZero = new List<ExcelRange>
        //        {
        //            worksheet.Cells["G53"],
        //            worksheet.Cells["E53"],
        //            worksheet.Cells["G54"],
        //            worksheet.Cells["E54"],
        //            worksheet.Cells["G55"],
        //            worksheet.Cells["E55"]
        //        };

        //    foreach (var cell in cellsToSetZero)
        //    {
        //        cell.Value = 0;
        //    }

        //    // FORMULAS TO DETERMINE CREDITS

        //    // CREDITOS TECNICO PROFESIONAL (1)
        //    var totalTechnicalCredits = worksheet.Cells["G53"]; // 3
        //    var totalTechnicalApproved = worksheet.Cells["E53"]; // 1
        //    var totalTechnicalPending = worksheet.Cells["F53"]; // 2
        //    totalTechnicalPending.Formula = $"({totalTechnicalCredits.Address}) - ({totalTechnicalApproved.Address})";

        //    // TOTAL CREDITOS APROBADOS (2)
        //    var totalTechnologistCredits = worksheet.Cells["G54"];
        //    var totalTechnologistApproved = worksheet.Cells["E54"];
        //    var totalTechnologistPending = worksheet.Cells["F54"];
        //    totalTechnologistPending.Formula = $"({totalTechnologistCredits.Address}) - ({totalTechnologistApproved.Address})";

        //    // TOTAL CREDITOS PENDIENTES (3)
        //    var totalProfessionalCredits = worksheet.Cells["G55"];
        //    var totalProfessionalApproved = worksheet.Cells["E55"];
        //    var totalProfessionalPending = worksheet.Cells["F55"];
        //    totalProfessionalPending.Formula = $"({totalProfessionalCredits.Address}) - ({totalProfessionalApproved.Address})";
        //}




        private void ApplyCellsFooter(ExcelWorksheet worksheet)
        {
            ExcelRange tableFooter = worksheet.Cells[78, 1, 79, 14];
            CellLeft(tableFooter);

            var cellFooter1 = worksheet.Cells[78, 1];
            cellFooter1.Value = "ELABORÓ: "; // Convertir y generar valor dinámico

            var cellFooter2 = worksheet.Cells[79, 1];
            cellFooter2.Value = "FECHA: "; // Convertir y generar valor dinámico

            ExcelRange cellSpaceFooter1 = worksheet.Cells[78, 2, 78, 4];
            MergedCells(cellSpaceFooter1);
            // TARGET

            ExcelRange cellSpaceFooter2 = worksheet.Cells[79, 2, 79, 4];
            MergedCells(cellSpaceFooter2);
            // TARGET

            var cellFooter3 = worksheet.Cells[78, 5];
            cellFooter3.Value = "ELABORÓ: "; // Convertir y generar valor dinámico

            var cellFooter4 = worksheet.Cells[79, 5];
            cellFooter4.Value = "FECHA: "; // Convertir y generar valor dinámico

            ExcelRange cellSpaceFooter3 = worksheet.Cells[78, 6, 78, 9];
            MergedCells(cellSpaceFooter3);
            // TARGET

            ExcelRange cellSpaceFooter4 = worksheet.Cells[79, 6, 79, 9];
            MergedCells(cellSpaceFooter4);
            // TARGET

            var cellFooter5 = worksheet.Cells[78, 10];
            cellFooter5.Value = "ELABORÓ: "; // Convertir y generar valor dinámico

            var cellFooter6 = worksheet.Cells[79, 10];
            cellFooter6.Value = "FECHA: "; // Convertir y generar valor dinámico

            ExcelRange cellSpaceFooter5 = worksheet.Cells[78, 11, 78, 14];
            MergedCells(cellSpaceFooter5);
            // TARGET

            ExcelRange cellSpaceFooter6 = worksheet.Cells[79, 11, 79, 14];
            MergedCells(cellSpaceFooter6);
            // TARGET
        }


    }


}
