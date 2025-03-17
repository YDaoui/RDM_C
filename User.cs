using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Windows.Forms;

namespace ExcelFormConverter
{
    public partial class MainForm : Form
    {
        private Excel.Workbook workbook;
        private Excel.Worksheet databaseSheet;

        public MainForm()
        {
            InitializeComponent();
            InitializeExcel();
            InitializeFormComponents();
        }

        private void InitializeExcel()
        {
            Excel.Application excelApp = new Excel.Application();
            workbook = excelApp.Workbooks.Open("Chemin/vers/votre/fichier.xlsx");
            databaseSheet = workbook.Sheets["Database"];
        }

        private void InitializeFormComponents()
        {
            // Initialisation similaire à UserForm_Initialize
            labelDate.Text = DateTime.Now.ToString("dddd dd MMMM yyyy");
            
            // Remplissage des ComboBox
            comboActionExistant.Items.AddRange(new object[] { "Appel", "Mail", "Proposition" });
            comboResumeNouveau.Items.AddRange(new object[] { "Clôt", "Rappel", "Suivie", "Traité", "Injoignable" });
            
            // Charger les données depuis Excel
            LoadDatabaseData();
        }

        private void LoadDatabaseData()
        {
            int lastRow = databaseSheet.Cells[databaseSheet.Rows.Count, "D"].End[Excel.XlDirection.xlUp].Row;
            
            for (int i = 10; i <= lastRow; i++)
            {
                string nom = databaseSheet.Cells[i, 4].Value?.ToString();
                string numFiche = databaseSheet.Cells[i, 5].Value?.ToString();
                if (!string.IsNullOrEmpty(nom))
                {
                    comboBox1.Items.Add($"{nom} - {numFiche}");
                }
            }
        }

        // Événement pour le bouton d'enregistrement
        private void btnEnregistrer_Click(object sender, EventArgs e)
        {
            try
            {
                int lastRow = databaseSheet.Cells[databaseSheet.Rows.Count, "C"].End[Excel.XlDirection.xlUp].Row + 1;
                if (lastRow < 10) lastRow = 10;

                databaseSheet.Cells[lastRow, 3] = textBox1.Text;
                databaseSheet.Cells[lastRow, 4] = textBox2.Text;
                databaseSheet.Cells[lastRow, 5] = textBox3.Text;

                MessageBox.Show("Les données ont été enregistrées avec succès !");
                ClearInputs();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur : {ex.Message}");
            }
        }

        private void ClearInputs()
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
        }

        // Événement pour la recherche
        private void textBoxSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string searchText = textBoxSearch.Text;
                // Implémenter la logique de recherche
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            workbook.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(databaseSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
        }
    }
}