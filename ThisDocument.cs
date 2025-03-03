using ProjectIP_2.Models;
using System;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;

namespace ProjectIP_2
{
    public partial class ThisDocument
    {
        private Form form;
        int idClient = 0;
        int idOferta = 0;
        Button btnAddClient;
        string connectionString = "Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=DB_AgentieTurism;Integrated Security=True";
        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            btnAddClient = new Button();
            btnAddClient.Click += new EventHandler(add_click);
            Globals.ThisDocument.ActionsPane.Controls.Add(btnAddClient);

            //ADAUGAREA OFERTELOR IN BAZA DE DATE
            //adaugaOferte();

            btnAddClient.Text = "Adauga Client";
        }

        private void add_click(object sender, System.EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(rNume.Text) ||
                string.IsNullOrWhiteSpace(rPrenume.Text) ||
                string.IsNullOrWhiteSpace(rJudet.Text) ||
                string.IsNullOrWhiteSpace(rTelefon.Text) ||
                string.IsNullOrWhiteSpace(rEmail.Text) ||
                string.IsNullOrWhiteSpace(rAdresa.Text) ||
                string.IsNullOrWhiteSpace(rLocalitate.Text) ||
                string.IsNullOrWhiteSpace(rCNP.Text))
            {
                MessageBox.Show("Completati toate campurile!");
            }
            else if (!DateTime.TryParse(rDate.Text, out DateTime dateOfBirth))
            {
                MessageBox.Show("Format Invalid pentru campul Data Nasterii");
            }
            else if (rCNP.Text.Length != 13)
            {
                MessageBox.Show("CNP invalid");
            }
            else if (!rEmail.Text.Contains("@"))
            {
                MessageBox.Show("Email invalid");
            }
            else
            {
                SqlConnection conn = new SqlConnection(connectionString);
                conn.Open();
                //string sqlQuery = "INSERT INTO Clienti (IdClient,Nume, Prenume, DataNasterii, CNP, Email, Telefon, Judet, Localitate, Adresa) " +
                //      "VALUES ('" + id.ToString() + "','Oana', 'Ramona', '2001-05-29', '6010529440033', 'oanaandreea19@stud.ase.ro', '0731831884', 'Ilfov', 'Copaceni', 'Strada Principala Nr 348')";
                Client client = new Client();
                client.Nume = rNume.Text;
                client.Prenume = rPrenume.Text;
                client.Judet = rJudet.Text;
                client.Telefon = rTelefon.Text;
                client.Email = rEmail.Text;
                client.Adresa = rAdresa.Text;
                client.Localitate = rLocalitate.Text;
                client.CNP = rCNP.Text;
                client.DataNasterii = DateTime.Parse(rDate.Text);
                string maxIdQuery = "SELECT MAX(IdClient) FROM Clienti";
                SqlCommand maxIdCommand = new SqlCommand(maxIdQuery, conn);
                int maxId = (int)maxIdCommand.ExecuteScalar();
                idClient = maxId + 1;

                string sqlQuery = "INSERT INTO Clienti (IdClient,Nume, Prenume, DataNasterii, CNP, Email, Telefon, Judet, Localitate, Adresa) " +
                      "VALUES ('" + idClient.ToString() + "','" + client.Nume + "', '" + client.Prenume + "', '" + client.DataNasterii.ToString("yyyy-MM-dd") + "', '" + client.CNP + "', '" + client.Email + "', '" + client.Telefon + "', '" + client.Judet + "', '" + client.Localitate + "', '" + client.Adresa + "')";

                SqlCommand sqlCommand = new SqlCommand(sqlQuery, conn);
                sqlCommand.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Datele au fost adaugate");

                rNume.Text = "";
                rPrenume.Text = "";
                rAdresa.Text = "";
                rCNP.Text = "";
                rEmail.Text = "";
                rLocalitate.Text = "";
                rTelefon.Text = "";
                rJudet.Text = "";


            }
        }



        private void adaugaOferte()
        {
            SqlConnection conn = new SqlConnection(connectionString);
            conn.Open();

            Oferta oferta1 = new Oferta();
            oferta1.Titlu = "Vacanta Antalya";
            oferta1.NrAdulti = 2;
            oferta1.DataPlecare = new DateTime(2023, 5, 6, 9, 0, 0);
            oferta1.DataIntoarcere = new DateTime(2023, 5, 13, 18, 30, 0);
            oferta1.Pret = 2000;
            oferta1.NumeHotel = "Turkish";
            oferta1.Descriere = "All Inclusive, Transport avion din Bucuresti,7 nopti cazare";
            oferta1.StatusOferta = Status.Disponibil;
            byte[] imageData1 = File.ReadAllBytes("C:\\Users\\aramo\\Desktop\\images\\antalya.jpg");
            oferta1.Imagine = imageData1;

            Oferta oferta2 = new Oferta();
            oferta2.Titlu = "Vacanta Grecia";
            oferta2.NrAdulti = 2;
            oferta2.DataPlecare = new DateTime(2023, 7, 10, 8, 0, 0);
            oferta2.DataIntoarcere = new DateTime(2023, 7, 15, 20, 30, 0);
            oferta2.Pret = 2500;
            oferta2.NumeHotel = "Imperial";
            oferta2.Descriere = "All Inclusive, Transport avion din Bucuresti,5 nopti cazare";
            oferta2.StatusOferta = Status.Disponibil;
            byte[] imageData2 = File.ReadAllBytes("C:\\Users\\aramo\\Desktop\\images\\grecia.jpg");
            oferta2.Imagine = imageData2;


            Oferta oferta3 = new Oferta();
            oferta3.Titlu = "Vacanta Bulgaria";
            oferta3.NrAdulti = 2;
            oferta3.DataPlecare = new DateTime(2023, 7, 20, 8, 0, 0);
            oferta3.DataIntoarcere = new DateTime(2023, 7, 25, 20, 30, 0);
            oferta3.Pret = 1500;
            oferta3.NumeHotel = "Royal";
            oferta3.Descriere = "All Inclusive, Transport avion din Bucuresti,5 nopti cazare";
            oferta3.StatusOferta = Status.Disponibil;
            byte[] imageData3 = File.ReadAllBytes("C:\\Users\\aramo\\Desktop\\images\\bulgaria.jpg");
            oferta3.Imagine = imageData3;

            Oferta[] oferte = new Oferta[3];
            oferte[0] = oferta1;
            oferte[1] = oferta2;
            oferte[2] = oferta3;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand())
                {
                    command.Connection = connection;
                    int id = 1;
                    command.CommandText = "INSERT INTO Oferte (IdOferta, Titlu, DataPlecare, DataIntoarcere, NumeHotel, NrAdulti, Pret, Descriere, StatusOferta, Imagine) VALUES (@IdOferta, @Titlu, @DataPlecare, @DataIntoarcere, @NumeHotel, @NrAdulti, @Pret, @Descriere, @StatusOferta, @Imagine)";
                    foreach (var oferta in oferte)
                    {
                        command.Parameters.AddWithValue("@IdOferta", id);
                        command.Parameters.AddWithValue("@Titlu", oferta.Titlu);
                        command.Parameters.AddWithValue("@DataPlecare", oferta.DataPlecare);
                        command.Parameters.AddWithValue("@DataIntoarcere", oferta.DataIntoarcere);
                        command.Parameters.AddWithValue("@NumeHotel", oferta.NumeHotel);
                        command.Parameters.AddWithValue("@NrAdulti", oferta.NrAdulti);
                        command.Parameters.AddWithValue("@Pret", oferta.Pret);
                        command.Parameters.AddWithValue("@Descriere", oferta.Descriere);
                        command.Parameters.AddWithValue("@StatusOferta", oferta.StatusOferta);
                        command.Parameters.AddWithValue("@Imagine", oferta.Imagine);

                        command.ExecuteNonQuery();
                        command.Parameters.Clear();
                        id++;
                    }
                }
            }


            MessageBox.Show("Ofertele au fost adaugate");
        }


        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {

        }

        #region VSTO Designer generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(this.ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(this.ThisDocument_Shutdown);

        }

        #endregion

    }
}
