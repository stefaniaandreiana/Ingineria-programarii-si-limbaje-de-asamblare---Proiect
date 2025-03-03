using System.Windows.Forms;

namespace ProjectIP_2
{
    public partial class FormStatistici : Form
    {
        private double valProfit;
        private double valReduceri;
        private float euro = 4.93f;
        public FormStatistici(double p, double r)
        {
            InitializeComponent();
            valProfit = p;
            valReduceri = r;
            labelTotal.Text = p.ToString();
            labelReduceri.Text = r.ToString();
            labelTLei.Text = "" + (p * euro);
            labelRLei.Text = "" + (r * euro);
            double procent = (r / p) * 100;
            labelPierderi.Text = procent.ToString();
        }
    }
}
