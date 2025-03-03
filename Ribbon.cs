using Dapper;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using ProjectIP_2.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace ProjectIP_2
{
    public partial class Ribbon
    {
        Form form;
        string connectionString = "Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=DB_AgentieTurism;Integrated Security=True";
        Excel.Application excel;
        private float euro = 4.93f;
        FormStatistici formStatistici;
        int rowExcel = 0;
        private void RibbonPPT_Load(object sender, RibbonUIEventArgs e)
        {
            excel = new Excel.Application();
        }

        private void buttonVizualizeazaOferte_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application application = new PowerPoint.Application();
            PowerPoint.Presentation presentation = application.Presentations.Add(Office.MsoTriState.msoTrue);

            int nrs = presentation.Slides.Count;
            PowerPoint.CustomLayout customLayout = presentation.SlideMaster.CustomLayouts[PowerPoint.PpSlideLayout.ppLayoutTitle];
            PowerPoint.Slide slidePPT = presentation.Slides.AddSlide(nrs + 1, customLayout);


            slidePPT.Shapes.Title.TextFrame.TextRange.Text = "Agentie Turism";
            slidePPT.Shapes[2].TextFrame.TextRange.Text = "“The world is a book and those who do not travel read only one page.”";
            slidePPT.Shapes[2].TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft;

            byte[] imageTravel = File.ReadAllBytes("C:\\Users\\aramo\\Desktop\\images\\travel.png");
            var tempTravel = Path.GetTempFileName();
            File.WriteAllBytes(tempTravel, imageTravel);
            try
            {
                var slideBackgroundTravel = slidePPT.Shapes.AddPicture(tempTravel, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0, presentation.PageSetup.SlideWidth, presentation.PageSetup.SlideHeight);
                slideBackgroundTravel.ZOrder(MsoZOrderCmd.msoSendToBack);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error setting background image: " + ex.Message);
            }

            finally
            {
                File.Delete(tempTravel);
            }

            SqlConnection conn = new SqlConnection(connectionString);
            conn.Open();
            List<Oferta> oferte = conn.Query<Oferta>("SELECT * FROM Oferte WHERE StatusOferta=0").ToList();

            foreach (Oferta oferta in oferte)
            {
                Slide slide = presentation.Slides.Add(presentation.Slides.Count + 1, PpSlideLayout.ppLayoutBlank);

                var title = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 290, 30, 400, 60);
                title.TextFrame.TextRange.Text = oferta.Titlu;
                title.TextFrame.TextRange.Font.Color.RGB = 16777215;
                title.TextFrame.TextRange.Font.Size = 50;
                title.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                title.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;



                TimeSpan durataZile = oferta.DataIntoarcere - oferta.DataPlecare;
                int zile = (int)durataZile.TotalDays;

                var durata = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 350, 80, 400, 50);
                durata.TextFrame.TextRange.Text = "Durata Sejur: " + zile + " zile";
                durata.TextFrame.TextRange.Font.Size = 20;


                var hotel = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 180, 110, 400, 60);
                hotel.TextFrame.TextRange.Text = "Hotel " + oferta.NumeHotel;
                hotel.TextFrame.TextRange.Font.Size = 50;
                hotel.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;

                var hotelBackground = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 50, 120, 400, 50);
                hotelBackground.Fill.ForeColor.RGB = ColorTranslator.ToOle(ColorTranslator.FromHtml("#DBF4FE"));
                hotelBackground.Line.Visible = MsoTriState.msoFalse;
                hotel.ZOrder(MsoZOrderCmd.msoBringToFront);

                var description = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 50, 180, 400, 50);
                description.TextFrame.TextRange.Text = "Detalii: " + oferta.Descriere;
                hotel.TextFrame.TextRange.Font.Size = 20;

                var descriptionBackground = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 50, 180, 400, 50);
                descriptionBackground.Fill.ForeColor.RGB = ColorTranslator.ToOle(ColorTranslator.FromHtml("#91DEFE"));
                descriptionBackground.Line.Visible = MsoTriState.msoFalse;
                description.ZOrder(MsoZOrderCmd.msoBringToFront);

                var nrPers = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 50, 240, 400, 50);
                nrPers.TextFrame.TextRange.Text = "Numar Persoane : " + oferta.NrAdulti + " adulti";
                nrPers.TextFrame.TextRange.Font.Size = 20;

                var nrPersBackground = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 50, 240, 400, 50);
                nrPersBackground.Fill.ForeColor.RGB = ColorTranslator.ToOle(ColorTranslator.FromHtml("#67D3FF"));
                nrPersBackground.Line.Visible = MsoTriState.msoFalse;
                nrPers.ZOrder(MsoZOrderCmd.msoBringToFront);

                byte[] imageAirplane = File.ReadAllBytes("C:\\Users\\aramo\\Desktop\\images\\border_airplane.png");
                var tempBorderAirplane = Path.GetTempFileName();
                File.WriteAllBytes(tempBorderAirplane, imageAirplane);
                var imageShapeAirplane = slide.Shapes.AddPicture(tempBorderAirplane, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0, 200, 50);


                var price = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 660, 80, 300, 50);
                price.TextFrame.TextRange.Text = oferta.Pret.ToString();
                price.TextFrame.TextRange.Font.Size = 60;
                price.TextFrame.TextRange.Font.Color.RGB = 16777215;

                var euro = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 680, 140, 300, 50);
                euro.TextFrame.TextRange.Text = "Euro";
                euro.TextFrame.TextRange.Font.Size = 40;
                euro.TextFrame.TextRange.Font.Color.RGB = 16777215;

                byte[] imagePrice = File.ReadAllBytes("C:\\Users\\aramo\\Desktop\\images\\border_price.png");
                var tempBorder = Path.GetTempFileName();
                File.WriteAllBytes(tempBorder, imagePrice);
                var imageShape = slide.Shapes.AddPicture(tempBorder, MsoTriState.msoFalse, MsoTriState.msoTrue, 600, 10, 250, 250);

                byte[] imageCalendar = File.ReadAllBytes("C:\\Users\\aramo\\Desktop\\images\\calendarr.png");
                var tempCalendar = Path.GetTempFileName();
                File.WriteAllBytes(tempCalendar, imageCalendar);
                var imageShapeCalendar = slide.Shapes.AddPicture(tempCalendar, MsoTriState.msoFalse, MsoTriState.msoTrue, 50, 300, 250, 250);


                var dataPlecare = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 350, 370, 500, 50);
                dataPlecare.TextFrame.TextRange.Text = "Data Plecare: " + oferta.DataPlecare;
                dataPlecare.TextFrame.TextRange.Font.Size = 20;

                var backgroundDate = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 330, 350, 400, 120);
                backgroundDate.Fill.ForeColor.RGB = ColorTranslator.ToOle(ColorTranslator.FromHtml("#CFFAFF"));
                backgroundDate.Line.Visible = MsoTriState.msoFalse;
                dataPlecare.ZOrder(MsoZOrderCmd.msoBringToFront);

                var dataIntoarcere = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 350, 430, 500, 50);
                dataIntoarcere.TextFrame.TextRange.Text = "Data Intoarcere: " + oferta.DataIntoarcere;
                dataIntoarcere.TextFrame.TextRange.Font.Size = 20;


                byte[] imageBytes = oferta.Imagine;
                var tempFilePath = Path.GetTempFileName();

                File.WriteAllBytes(tempFilePath, imageBytes);

                try
                {
                    var slideBackground = slide.Shapes.AddPicture(tempFilePath, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0, presentation.PageSetup.SlideWidth, presentation.PageSetup.SlideHeight);
                    slideBackground.ZOrder(MsoZOrderCmd.msoSendToBack);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error setting background image: " + ex.Message);
                }

                finally
                {
                    File.Delete(tempFilePath);
                }
            }
            conn.Close();



        }

        private void buttonDeschideExcel_Click(object sender, RibbonControlEventArgs e)
        {
            excel.Visible = true;
            Excel.Workbook workbook = excel.Workbooks.Add();

            Excel.Worksheet worksheet = workbook.Sheets.Add();
            worksheet.Name = "Oferte";

            SqlConnection conn = new SqlConnection(connectionString);
            conn.Open();
            List<Oferta> oferte = conn.Query<Oferta>("SELECT * FROM Oferte").ToList();

            worksheet.Cells[1, 1] = "Nr. Crt";
            worksheet.Cells[1, 2] = "Titlu";
            worksheet.Cells[1, 3] = "Data Plecare";
            worksheet.Cells[1, 4] = "Data Intoarcere";
            worksheet.Cells[1, 5] = "Hotel";
            worksheet.Cells[1, 6] = "Numar Adulti";
            worksheet.Cells[1, 7] = "Pret";
            worksheet.Cells[1, 8] = "Descriere";
            worksheet.Cells[1, 9] = "Status";

            int row = 2;
            foreach (Oferta oferta in oferte)
            {
                worksheet.Cells[row, 1] = oferta.IdOferta;
                worksheet.Cells[row, 2] = oferta.Titlu;
                worksheet.Cells[row, 3] = oferta.DataPlecare;
                worksheet.Cells[row, 4] = oferta.DataIntoarcere;
                worksheet.Cells[row, 5] = oferta.NumeHotel;
                worksheet.Cells[row, 6] = oferta.NrAdulti;
                worksheet.Cells[row, 7] = oferta.Pret;
                worksheet.Cells[row, 8] = oferta.Descriere;
                worksheet.Cells[row, 9] = oferta.StatusOferta.ToString();
                row++;
            }

            Excel.Worksheet worksheet2 = workbook.Sheets.Add();
            worksheet2.Name = "Clienti";

            worksheet2.Cells[1, 1] = "Nume";
            worksheet2.Cells[1, 2] = "Prenume";
            worksheet2.Cells[1, 3] = "CNP";
            worksheet2.Cells[1, 4] = "Data Nasterii";
            worksheet2.Cells[1, 5] = "Email";
            worksheet2.Cells[1, 6] = "Telefon";
            worksheet2.Cells[1, 7] = "Judet";
            worksheet2.Cells[1, 8] = "Localitate";
            worksheet2.Cells[1, 9] = "Adresa";

            SqlConnection conn2 = new SqlConnection(connectionString);
            conn2.Open();
            List<Client> clienti = conn2.Query<Client>("SELECT * FROM Clienti").ToList();

            int row2 = 2;
            foreach (Client client in clienti)
            {
                worksheet2.Cells[row2, 1] = client.Nume;
                worksheet2.Cells[row2, 2] = client.Prenume;
                worksheet2.Cells[row2, 3] = client.CNP;
                worksheet2.Cells[row2, 4] = client.DataNasterii;
                worksheet2.Cells[row2, 5] = client.Email;
                worksheet2.Cells[row2, 6] = client.Telefon;
                worksheet2.Cells[row2, 7] = client.Judet;
                worksheet2.Cells[row2, 8] = client.Localitate;
                worksheet2.Cells[row2, 9] = client.Adresa;
                row2++;
            }


            Excel.Worksheet worksheet3 = workbook.Sheets.Add();
            worksheet3.Name = "Rezervari";

            worksheet3.Cells[1, 1] = "Nume";
            worksheet3.Cells[1, 2] = "Prenume";
            worksheet3.Cells[1, 3] = "CNP";
            worksheet3.Cells[1, 4] = "Email";
            worksheet3.Cells[1, 5] = "Telefon";
            worksheet3.Cells[1, 6] = "Oferta";
            worksheet3.Cells[1, 7] = "Data Semnare";
            worksheet3.Cells[1, 8] = "Pret Total";
            worksheet3.Cells[1, 9] = "Avans";

            Excel.Range columnC = worksheet3.Range["C:C"];
            columnC.NumberFormat = "0";

            Excel.Range columnE = worksheet3.Range["E:E"];
            columnE.NumberFormat = "@";

            Excel.Range columnCClienti = worksheet2.Range["C:C"];
            columnCClienti.NumberFormat = "0";

            Excel.Range columnEClienti = worksheet2.Range["E:E"];
            columnEClienti.NumberFormat = "@";

            Excel.Style headerStyle = workbook.Styles.Add("HeaderStyle");
            headerStyle.Font.Bold = true;
            headerStyle.Font.Color = Excel.XlRgbColor.rgbWhite;
            headerStyle.Interior.Color = Excel.XlRgbColor.rgbDarkBlue;


            Excel.Range headerRange = worksheet3.Range["A1:I1"];
            headerRange.Style = headerStyle;
            Excel.Range cells = worksheet3.Cells;
            cells.Range["A2:K2"].ColumnWidth = 20;


            Excel.Range headerRange2 = worksheet2.Range["A1:I1"];
            headerRange2.Style = headerStyle;
            Excel.Range cells2 = worksheet2.Cells;
            cells2.Range["A2:K2"].ColumnWidth = 20;

            Excel.Range headerRange3 = worksheet.Range["A1:I1"];
            headerRange3.Style = headerStyle;
            Excel.Range cells3 = worksheet.Cells;
            cells3.Range["A2:K2"].ColumnWidth = 20;


            SqlConnection conn3 = new SqlConnection(connectionString);
            conn3.Open();
            List<Contract> contracte = conn3.Query<Contract>("SELECT * FROM Contracte").ToList();
            int row3 = 2;

            foreach (Contract contract in contracte)
            {
                SqlConnection conn4 = new SqlConnection(connectionString);
                conn4.Open();
                Oferta oferta = conn4.QueryFirstOrDefault<Oferta>("SELECT * FROM Oferte WHERE IdOferta=@idOferta", new { idOferta = contract.IdOferta });
                Client client = conn4.QueryFirstOrDefault<Client>("SELECT * FROM Clienti WHERE IdClient=@idClient", new { idClient = contract.IdClient });
                worksheet3.Cells[row3, 1] = client.Nume;
                worksheet3.Cells[row3, 2] = client.Prenume;
                worksheet3.Cells[row3, 3] = client.CNP;
                worksheet3.Cells[row3, 4] = client.Email;
                worksheet3.Cells[row3, 5] = client.Telefon;
                worksheet3.Cells[row3, 6] = oferta.Titlu;
                worksheet3.Cells[row3, 7] = contract.DataSemnare;
                worksheet3.Cells[row3, 8] = oferta.Pret;
                worksheet3.Cells[row3, 9] = contract.Avans;
                worksheet3.Cells[row3, 10] = "-";
                worksheet3.Cells[row3, 11] = oferta.Pret - contract.Avans;
                row3++;
            }
            row3--;
            string rangeBorder = "A2:K" + row3;

            Excel.Range range = worksheet3.Range[rangeBorder];
            Excel.Borders borders = range.Borders;
            borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            borders.Weight = Excel.XlBorderWeight.xlMedium;

            worksheet3.Cells[1, 10] = "Procent Reducere";
            worksheet3.Cells[1, 11] = "Suma neachitata";

            worksheet3.SelectionChange += Worksheet_SelectionChange;

            string cellRange1 = "H2:H" + row3;
            Excel.Range sumRange1 = worksheet3.Range[cellRange1];
            Excel.Range totalCell1 = worksheet3.Range["H" + (row3 + 1)];
            totalCell1.Formula = "=SUM(" + sumRange1.Address + ")";

            string cellRange2 = "I2:I" + row3;
            Excel.Range sumRange2 = worksheet3.Range[cellRange2];
            Excel.Range totalCell2 = worksheet3.Range["I" + (row3 + 1)];
            totalCell2.Formula = "=SUM(" + sumRange2.Address + ")";


            string cellRange3 = "K2:K" + row3;
            Excel.Range sumRange3 = worksheet3.Range[cellRange3];
            Excel.Range totalCell3 = worksheet3.Range["K" + (row3 + 1)];
            totalCell3.Formula = "=SUM(" + sumRange3.Address + ")";

            rowExcel = row3;
            rowExcel++;


            double a = Double.Parse(totalCell1.Text);
            double b = Double.Parse(totalCell2.Text);
            double c = Double.Parse(totalCell3.Text);

            double reduceri = a - (b + c);

            formStatistici = new FormStatistici(a, reduceri);
            formStatistici.Show();



            workbook.Activate();
        }
        private void Worksheet_SelectionChange(Excel.Range target)
        {

            Excel.Range selectedRange = (Excel.Range)excel.Selection;
            Excel.Worksheet worksheetExcel = excel.ActiveSheet;
            Excel.Range desiredRange = worksheetExcel.Range["A2:I" + rowExcel];

            if (selectedRange != null && desiredRange != null)
            {
                Excel.Range intersectRange = selectedRange.Application.Intersect(selectedRange, desiredRange);

                if (intersectRange != null)
                {
                    int rowCount = intersectRange.Rows.Count;
                    int colCount = intersectRange.Columns.Count;

                    FormExcel form = new FormExcel();
                    form.ShowDialog();
                    double percentageExcel = form.percentage;

                    if (rowCount > 1)
                    {
                        string address = target.Address;

                        string[] cellAddresses = address.Split(':');

                        string secondCellAddress = cellAddresses[1];
                        string rowNumberString2 = secondCellAddress.Substring(3);

                        string firstCellAddress = cellAddresses[0];
                        string rowNumberString = firstCellAddress.Substring(3);

                        int nr1 = int.Parse(rowNumberString);
                        int nr2 = int.Parse(rowNumberString2);

                        for (int i = nr1; i <= nr2; i++)
                        {
                            string cellRegion = "J" + i;
                            Excel.Worksheet worksheet = excel.ActiveSheet;
                            Excel.Range cell = worksheet.Range[cellRegion];
                            cell.Value = percentageExcel;
                            double total = worksheet.Range["H" + i].Value;
                            double avans = worksheet.Range["I" + i].Value;
                            double sumaNeachitata = total - avans;
                            double totalDeAchitat = sumaNeachitata - (sumaNeachitata * percentageExcel / 100);

                            string cellRegionTotal = "K" + i;
                            Excel.Range cellTotal = worksheet.Range[cellRegionTotal];
                            cellTotal.Value = totalDeAchitat;
                        }
                    }
                    else
                    {
                        if (!target.Address.Contains(":"))
                        {
                            int rowNumber = int.Parse(target.Address.Substring(3));
                            string cellRegion = "J" + rowNumber;
                            Excel.Worksheet worksheet = excel.ActiveSheet;
                            Excel.Range cell = worksheet.Range[cellRegion];
                            cell.Value = percentageExcel;
                            double total = worksheet.Range["H" + rowNumber].Value;
                            double avans = worksheet.Range["I" + rowNumber].Value;
                            double sumaNeachitata = total - avans;
                            double totalDeAchitat = sumaNeachitata - (sumaNeachitata * percentageExcel / 100);

                            string cellRegionTotal = "K" + rowNumber;
                            Excel.Range cellTotal = worksheet.Range[cellRegionTotal];
                            cellTotal.Value = totalDeAchitat;
                        }
                    }

                    Excel.Worksheet currentWorksheet = excel.ActiveSheet;
                    double cellValue1 = currentWorksheet.Range["H" + rowExcel].Value;
                    double cellValue2 = currentWorksheet.Range["I" + rowExcel].Value;
                    double cellValue3 = currentWorksheet.Range["K" + rowExcel].Value;


                    double reduceri = cellValue1 - (cellValue2 + cellValue3);

                    update(cellValue1, reduceri);


                }
                else
                {
                    MessageBox.Show("Selectați o adresă validă.");
                }
            }


        }

        private void update(double a, double b)
        {
            formStatistici.labelTotal.Invoke((MethodInvoker)delegate
            {
                formStatistici.labelTotal.Text = a.ToString();
            });
            formStatistici.labelReduceri.Invoke((MethodInvoker)delegate
            {
                formStatistici.labelReduceri.Text = b.ToString();
            });
            double procent = (b / a) * 100;
            formStatistici.labelPierderi.Invoke((MethodInvoker)delegate
            {
                formStatistici.labelPierderi.Text = procent.ToString();
            });
            formStatistici.labelPierderi.Invoke((MethodInvoker)delegate
            {
                formStatistici.labelRLei.Text = "" + (b * euro);
            });


        }

        private void buttonGenereazaContract_Click(object sender, RibbonControlEventArgs e)
        {
            FormOferta form = new FormOferta();
            form.ShowDialog();
        }
    }
}
