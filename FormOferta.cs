using Dapper;
using Microsoft.Office.Interop.Word;
using ProjectIP_2.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace ProjectIP_2
{
    public partial class FormOferta : Form
    {
        string connectionString = "Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=DB_AgentieTurism;Integrated Security=True";

        public FormOferta()
        {
            InitializeComponent();
            getValues();
        }

        private void getValues()
        {
            SqlConnection conn = new SqlConnection(connectionString);
            conn.Open();
            List<Client> clients = conn.Query<Client>("SELECT * FROM Clienti").ToList();
            foreach (Client client in clients)
            {
                String fullName = client.Nume.Trim() + " " + client.Prenume.Trim();
                comboBoxClienti.Items.Add(fullName);
            }

            List<Oferta> oferte = conn.Query<Oferta>("SELECT * FROM Oferte WHERE StatusOferta=0").ToList();
            foreach (Oferta oferta in oferte)
            {
                String numeOferta = oferta.Titlu.Trim();
                comboBoxOferte.Items.Add(numeOferta);
            }

            conn.Close();
        }

        private void buttonGenereazaContract_Click(object sender, EventArgs e)
        {
            if (comboBoxClienti.SelectedIndex != -1 && comboBoxOferte.SelectedIndex != -1)
            {
                string selectedClient = comboBoxClienti.SelectedItem.ToString();
                string selectedOffer = comboBoxOferte.SelectedItem.ToString();

                DialogResult result = MessageBox.Show("Doriti sa rezervati sejurul pentru clientul " + selectedClient + "?", "Confirmation", MessageBoxButtons.OKCancel);

                if (result == DialogResult.OK)
                {
                    MessageBox.Show("Se genereaza contractul ...");
                    this.Close();
                    genereazaContract(selectedClient, selectedOffer);

                }
                else if (result == DialogResult.Cancel)
                {
                    this.Close();
                }
            }
            else
            {
                MessageBox.Show("Selectati clientul si sejurul!");
            }
        }



        private void genereazaContract(string clientName, string offer)
        {
            string[] name = clientName.Split(' ');

            if (name.Length > 0)
            {
                string firstName = name[0];
                string lastName = String.Join(" ", name.Skip(1));


                SqlConnection conn = new SqlConnection(connectionString);
                conn.Open();
                Client client = conn.QueryFirstOrDefault<Client>("SELECT * FROM Clienti WHERE Nume=@firstName AND Prenume=@lastName", new { firstName = firstName, lastName = lastName });
                Oferta oferta = conn.QueryFirstOrDefault<Oferta>("SELECT * FROM Oferte WHERE Titlu=@titlu", new { titlu = offer });

                Word.Application wordApp = new Word.Application();

                Word.Document doc = wordApp.Documents.Add();

                Word.Paragraph title1 = doc.Content.Paragraphs.Add();
                title1.Range.Text = "CONTRACT";
                title1.Range.Font.Name = "Times New Roman";
                title1.Range.Font.Size = 20;
                title1.Range.Font.Bold = 1;
                title1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                title1.Range.InsertParagraphAfter();

                Word.Paragraph title = doc.Content.Paragraphs.Add();
                title.Range.Text = "de comercializare a pachetelor de servicii de calatorie";
                title.Range.Font.Name = "Times New Roman";
                title.Range.Font.Size = 20;
                title.Range.Font.Bold = 1;
                title.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                title.Range.InsertParagraphAfter();

                Word.Paragraph para1 = doc.Content.Paragraphs.Add();
                para1.Range.Text = "HAPPY TOUR S.R.L. cu sediul in Bucuresti, Polona Business Center, Receptia 2, Etaj 4, Strada Polona 68-72, Sector 1, tel. 021 / 307 06 30, fax. 307 06 40, e-mail: office@happytour.ro, " +
                    "inregistrata la Registrul Comertului sub nr. J40 / 23452/ 13 dec. 1994, CIF RO 6842431, cont bancar deschis la Banca Transilvania, Cont RON - RO59 BTRL 0440 1202 1411 70XX, Cont EUR - RO08 BTRL 0440 4202 1411 70XX" +
                    " cu licenta de turism nr. 1400 / 11.03. 2019, reprezentata prin Dl. Javier Garcia del Valle – Administrator – denumita in continuare AGENTIE ORGANIZATOARE si Calatorul:";

                para1.Range.Font.Name = "Times New Roman";
                para1.Range.Font.Size = 9;
                para1.Range.Font.Bold = 0;
                para1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                para1.Range.InsertParagraphAfter();

                Word.Paragraph paraClient = doc.Content.Paragraphs.Add();
                paraClient.Range.Text = client.Nume.ToUpper().Trim() + " " + client.Prenume.ToUpper().Trim() + ", cu CNP-ul: "
                    + client.CNP.Trim() + ", domiciliat in " + client.Judet.Trim() + "," + client.Localitate.Trim() + "," + client.Adresa.Trim() + ", cu datele de contact: " + client.Telefon.Trim() + "," + client.Email.Trim()
                    + ",au convenit la incheierea prezentului contract pentru rezervarea urmatorului sejur: " +
                    oferta.Titlu.Trim().ToUpper() + ", la hotelul " + oferta.NumeHotel.Trim() + ", pentru " + oferta.NrAdulti + " adulti," + " cu urmatoarele detalii: " + oferta.Descriere.Trim() + ". Data plecare: " + oferta.DataPlecare + ". Data intoarcere: " + oferta.DataIntoarcere + ".";
                paraClient.Range.Font.Size = 9;
                paraClient.Range.Font.Bold = 1;
                paraClient.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paraClient.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                paraClient.Range.InsertParagraphAfter();
                paraClient.Format.SpaceAfter = 6;


                Word.Paragraph para2 = doc.Content.Paragraphs.Add();
                para2.Range.Text = "I. OBIECTUL CONTRACTULUI";
                para2.Range.Font.Name = "Times New Roman";
                para2.Range.Font.Size = 12;
                para2.Range.Font.Bold = 1;
                para2.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para2.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                para2.Range.InsertParagraphAfter();
                para2.Format.SpaceAfter = 6;

                Word.Paragraph para3 = doc.Content.Paragraphs.Add();
                para3.Range.Text = "1.1 Il constituie vânzarea de catre Agentie a pachetului de servicii de calatorie/a unui serviciu de calatorie si/sau a unor servicii asociate inscrise in voucher, bilet de odihna, tratament, bilet de excursie, alt inscris anexat prezentului contract si eliberarea documentelor de plata si calatorie;";
                para3.Range.Font.Name = "Times New Roman";
                para3.Range.Font.Size = 9;
                para3.Range.Font.Bold = 0;
                para3.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para3.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                para3.Range.InsertParagraphAfter();
                para3.Format.SpaceAfter = 6;

                Word.Paragraph para4 = doc.Content.Paragraphs.Add();
                para4.Range.Text = "1.2 Caracteristicile pachetului de servicii de calatorie sunt prezentate in oferta, iar aceasta este parte integranta a acestui contract.";
                para4.Range.Font.Name = "Times New Roman";
                para4.Range.Font.Size = 9;
                para4.Range.Font.Bold = 0;
                para4.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para4.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                para4.Range.InsertParagraphAfter();
                para4.Format.SpaceAfter = 6;

                Word.Paragraph para5 = doc.Content.Paragraphs.Add();
                para5.Range.Text = "1.2 Caracteristicile pachetului de servicii de calatorie sunt prezentate in oferta, iar aceasta este parte integranta a acestui contract.";
                para5.Range.Font.Name = "Times New Roman";
                para5.Range.Font.Size = 9;
                para5.Range.Font.Bold = 0;
                para5.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para5.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                para5.Range.InsertParagraphAfter();
                para5.Format.SpaceAfter = 6;

                Word.Paragraph para6 = doc.Content.Paragraphs.Add();
                para6.Range.Text = "II. INCHEIEREA SI DURATA CONTRACTULUI";
                para6.Range.Font.Name = "Times New Roman";
                para6.Range.Font.Size = 12;
                para6.Range.Font.Bold = 1;
                para6.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para6.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                para6.Range.InsertParagraphAfter();
                para6.Format.SpaceAfter = 6;

                Word.Paragraph para7 = doc.Content.Paragraphs.Add();
                para7.Range.Text = "2.Contractul se încheie, după caz, în oricare din următoarele situaţii:";
                para7.Range.Font.Name = "Times New Roman";
                para7.Range.Font.Size = 9;
                para7.Range.Font.Bold = 0;
                para7.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para7.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                para7.Range.InsertParagraphAfter();
                para7.Format.SpaceAfter = 6;

                Word.Paragraph para8 = doc.Content.Paragraphs.Add();
                para8.Range.Text = "2.1 In momentul semnarii lui de catre Calator sau prin acceptarea conditiilor precontractuale de servicii de calatorie, inclusiv in cazul celor achizitionate la distanta prin telefon si/sau mijloace electronice.";
                para8.Range.Font.Name = "Times New Roman";
                para8.Range.Font.Size = 9;
                para8.Range.Font.Bold = 0;
                para8.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para8.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                para8.Range.InsertParagraphAfter();
                para8.Format.SpaceAfter = 6;

                Word.Paragraph para9 = doc.Content.Paragraphs.Add();
                para9.Range.Text = "2.2 In momentul în care calatorul primeşte confirmarea scrisă a rezervării de la Agenţie. Este responsabilitatea agentiei de turism de a informa calatorul prin orice mijloace convenite cu acesta (telefon, mail, fax etc.) daca rezervarea solicitata este confirmata. Pentru procesarea unei rezervari de servicii, Agentia poate solicita un avans cuprins intre 20 - 50 % din pretul pachetului sau plata integrala a contravalorii pachetului, in functie de data la care calatorul solicita serviciile.";
                para9.Range.Font.Name = "Times New Roman";
                para9.Range.Font.Size = 9;
                para9.Range.Font.Bold = 0;
                para9.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para9.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                para9.Range.InsertParagraphAfter();
                para9.Format.SpaceAfter = 6;

                Word.Paragraph para10 = doc.Content.Paragraphs.Add();
                para10.Range.Text = "2.3 In momentul eliberarii documentelor de calatorie (voucher, bilet de odihna si/sau tratament, bilet de excursie, etc.), inclusiv in format electronic, in cazul in care pachetele de servicii de calatorie fac parte din oferta standard a agentiei de turism sau exista deja confirmarea de rezervare din partea altor prestatori.";
                para10.Range.Font.Name = "Times New Roman";
                para10.Range.Font.Size = 9;
                para10.Range.Font.Bold = 0;
                para10.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para10.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.InsertParagraphAfter();
                para10.Format.SpaceAfter = 6;

                Word.Paragraph para11 = doc.Content.Paragraphs.Add();
                para11.Range.Text = "2.4 Contractul încetează, de drept, odată cu finalizarea prestării efective a pachetului de servicii turistice înscris în documentele de călătorie.";
                para11.Range.Font.Name = "Times New Roman";
                para11.Range.Font.Size = 9;
                para11.Range.Font.Bold = 0;
                para11.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para11.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                para11.Range.InsertParagraphAfter();
                para11.Format.SpaceAfter = 6;

                Word.Paragraph para12 = doc.Content.Paragraphs.Add();
                para12.Range.Text = "III. PRETUL CONTRACTULUI SI MODALITATI DE PLATA";
                para12.Range.Font.Name = "Times New Roman";
                para12.Range.Font.Size = 12;
                para12.Range.Font.Bold = 1;
                para12.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para12.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                para12.Range.InsertParagraphAfter();
                para12.Format.SpaceAfter = 6;

                Word.Paragraph para13 = doc.Content.Paragraphs.Add();
                para13.Range.Text = "Preţul total al contractului este de " + oferta.Pret + " Euro" + " si include toate taxele, comisioanele, tarifele si orice alte costuri suportate de Agentie.";
                para13.Range.Font.Name = "Times New Roman";
                para13.Range.Font.Size = 9;
                para13.Range.Font.Bold = 0;
                para13.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para13.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                para13.Range.InsertParagraphAfter();
                para13.Format.SpaceAfter = 6;

                Word.Paragraph para14 = doc.Content.Paragraphs.Add();
                para14.Range.Text = "Clientul trebuie sa realizeze plata unui avans de 10% din pretul pachetului achizitionat, sau dupa caz, plata integrala, conform conditiilor mentionate in oferta transmisa electronic si/sau conform conditiilor ofertelor prezentate pe site-ul Agentiei www.happytour.ro. Platile se pot face in EUR /USD/ RON. Pentru platile in lei, acestea se calculeaza folosind cursul de schimb valutar al bancii comerciale a Agentiei, Banca Transilvania.";
                para14.Range.Font.Name = "Times New Roman";
                para14.Range.Font.Size = 9;
                para14.Range.Font.Bold = 0;
                para14.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para14.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                para14.Range.InsertParagraphAfter();
                para14.Format.SpaceAfter = 6;

                Word.Paragraph para15 = doc.Content.Paragraphs.Add();
                para15.Range.Text = "Conditiile de plata difera in functie de tipul pachetului de servicii de calatorie achizitionat, de tipul de oferta si vor fi trecute in contract sau in anexele aferente acestuia, adica in informatiile precontractuale care fac parte intergranta din acest contract.";
                para15.Range.Font.Name = "Times New Roman";
                para15.Range.Font.Size = 9;
                para15.Range.Font.Bold = 0;
                para15.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para15.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                para15.Range.InsertParagraphAfter();
                para15.Format.SpaceAfter = 6;

                Word.Paragraph para16 = doc.Content.Paragraphs.Add();
                para16.Range.Text = "IV. PRELUCRAREA DATELOR CU CARACTER PERSONAL";
                para16.Range.Font.Name = "Times New Roman";
                para16.Range.Font.Size = 12;
                para16.Range.Font.Bold = 1;
                para16.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para16.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                para16.Range.InsertParagraphAfter();
                para16.Format.SpaceAfter = 6;

                Word.Paragraph para17 = doc.Content.Paragraphs.Add();
                para17.Range.Text = "Prelucrarea datelor cu caracter personal reprezinta orice operatiune sau set de operatiuni care se efectueaza asupra datelor dumneavoastra cu caracter personal.HAPPY TOUR SRL poate prelucra urmatoarele date cu caracter personal: nume, prenume, numar telefon, adresa domiciliu, adresa de e-mail, serie si nr. carte de identitate, serie si nr. pasaport, CNP, data nasterii, varsta copiilor, apartenenta la sindicate, locul de munca, numele companiei (daca este aplicabil), numarul de inregistrare TVA (daca este cazul).";
                para17.Range.Font.Name = "Times New Roman";
                para17.Range.Font.Size = 9;
                para17.Range.Font.Bold = 0;
                para17.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para17.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                para17.Range.InsertParagraphAfter();
                para17.Format.SpaceAfter = 6;

                Word.Paragraph para18 = doc.Content.Paragraphs.Add();
                para18.Range.Text = "V. DISPOZITII FINALE";
                para18.Range.Font.Name = "Times New Roman";
                para18.Range.Font.Size = 12;
                para18.Range.Font.Bold = 1;
                para18.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para18.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                para18.Range.InsertParagraphAfter();
                para18.Format.SpaceAfter = 6;

                Word.Paragraph para19 = doc.Content.Paragraphs.Add();
                para19.Range.Text = "Prezentul contract a fost incheiat in doua exemplare, câte unul pentru fiecare parte. Comercializarea pachetelor de servicii de calatorie se va face in conformitate cu prevederile prezentului contract si cu respectarea prevederilor Ordonantei Guvernului nr. 2/2018 privind pachetele de servicii de calatorie si serviciile de calatorie associate, precum si a tuturor celorlalte acte normative incidente.";
                para19.Range.Font.Name = "Times New Roman";
                para19.Range.Font.Size = 9;
                para19.Range.Font.Bold = 0;
                para19.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para19.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                para19.Range.InsertParagraphAfter();
                para19.Format.SpaceAfter = 6;

                Word.Paragraph para20 = doc.Content.Paragraphs.Add();
                para20.Range.Text = "Comercializarea pachetelor de servicii de calatorie se va face in conformitate cu prevederile prezentului contract si cu respectarea prevederilor Ordonantei Guvernului nr. 2/2018 privind pachetele de servicii de calatorie si serviciile de calatorie associate, precum si a tuturor celorlalte  acte normative incidente.";
                para20.Range.Font.Name = "Times New Roman";
                para20.Range.Font.Size = 9;
                para20.Range.Font.Bold = 0;
                para20.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para20.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                para20.Range.InsertParagraphAfter();
                para20.Format.SpaceAfter = 6;

                Word.Paragraph para21 = doc.Content.Paragraphs.Add();
                para21.Range.Text = "Toate unitatile de cazare, precum si mijloacele de transport sunt clasificate de catre organismele abilitate ale tarilor de destinatie, conform procedurilor interne si normativelor locale acolo unde acestea exista, care difera de la o tara la alta si de la un tip de destinatie la altul.";
                para21.Range.Font.Size = 9;
                para21.Range.Font.Bold = 0;
                para21.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para21.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                para21.Format.SpaceAfter = 6;
                para21.Range.InsertParagraphAfter();

                Word.Paragraph para22 = doc.Content.Paragraphs.Add();
                para22.Range.Text = "Calatorul declara ca Agentia l-a informat complet cu privire la conditiile de comercializare a pachetelor de servicii de calatorie in conformitate cu prevederile Ordonantei Guvernului nr. 2/2018. Prin semnarea acestui contract, sau prin acceptarea pachetelor de servicii de calatorie inclusiv in cazul celor achizitionate la distanta prin mijloace electronice, calatorul isi exprima acordul si luarea la cunostinta cu privire la conditiile generale de comercializare a pachetelor de servicii de calatorie, in conformitate cu oferta Agentiei.";
                para22.Range.Font.Size = 9;
                para22.Range.Font.Bold = 0;
                para22.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para22.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                para22.Range.InsertParagraphAfter();

                Word.Paragraph para23 = doc.Content.Paragraphs.Add();

                Table table = doc.Tables.Add(para23.Range, 2, 2);
                table.Columns[1].Width = 300;
                table.Columns[2].Width = 300;

                table.Cell(1, 1).Range.Text = "Client: " + client.Nume.Trim() + " " + client.Prenume.Trim();
                table.Cell(1, 2).Range.Text = "Companie: Happy Tour";

                var cell1 = table.Cell(2, 2);
                var shape1 = cell1.Range.InlineShapes.AddPicture("C:\\Users\\aramo\\Desktop\\images\\stamp.png");
                shape1.Width = 100;
                shape1.Height = 100;

                wordApp.Visible = true;

                SqlConnection conn2 = new SqlConnection(connectionString);
                conn2.Open();
                string maxIdQuery = "SELECT MAX(IdContract) FROM Contracte";
                SqlCommand maxIdCommand = new SqlCommand(maxIdQuery, conn2);
                object result = maxIdCommand.ExecuteScalar();
                int idContract = 1;
                if (result != null && result != DBNull.Value)
                {
                    int maxId = Convert.ToInt32(result);
                    idContract = maxId + 1;
                }


                string fileName = "Contract" + idContract + ".docx";
                string filePath = Path.Combine(@"C:\Users\aramo\Desktop\contracte", fileName);

                string docText = doc.WordOpenXML;
                byte[] fileBytes = Encoding.UTF8.GetBytes(docText);

                doc.SaveAs2(filePath);

                using (SqlConnection conn3 = new SqlConnection(connectionString))
                {
                    conn3.Open();
                    double avans = oferta.Pret * 0.1;

                    string sqlQuery = "INSERT INTO Contracte (IdContract, IdClient, IdOferta, DataSemnare, Avans, Document) " +
                                       "VALUES (@IdContract, @IdClient, @IdOferta, @DataSemnare, @Avans, @Document)";

                    SqlCommand sqlCommand = new SqlCommand(sqlQuery, conn3);
                    sqlCommand.Parameters.AddWithValue("@IdContract", idContract);
                    sqlCommand.Parameters.AddWithValue("@IdClient", client.IdClient);
                    sqlCommand.Parameters.AddWithValue("@IdOferta", oferta.IdOferta);
                    sqlCommand.Parameters.AddWithValue("@DataSemnare", DateTime.Now.ToString("yyyy-MM-dd"));
                    sqlCommand.Parameters.AddWithValue("@Avans", avans);
                    sqlCommand.Parameters.AddWithValue("@Document", fileBytes);

                    sqlCommand.ExecuteNonQuery();
                    conn3.Close();
                }

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    string sql = "UPDATE Oferte SET StatusOferta = @Status WHERE Titlu = @Titlu";
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@Status", 1);
                        command.Parameters.AddWithValue("@Titlu", oferta.Titlu);

                        int rowsAffected = command.ExecuteNonQuery();
                    }
                }



            }
        }
    }
}
