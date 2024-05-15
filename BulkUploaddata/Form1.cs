namespace BulkUploaddata
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string schema = textBox1.Text;
            string table = textBox2.Text;

            op.formQuery(schema, table);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
