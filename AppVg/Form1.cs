using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Data.OleDb;
using System.Threading;
using Xceed.Words.NET;
using Xceed.Document.NET;

namespace AppVg
{

    public partial class Form1 : Form
    {
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        private string pathTempDocxP;

        public Form1()
        {
            InitializeComponent();
            pSubMenuProp.Width = 167;
            pSubMenuProp.Visible = false;

            var pathTempFolder = Path.Combine(System.AppContext.BaseDirectory, "TempFilesApp");
            var pathTempDocx = pathTempFolder + "\\tempProp.docx";
            pathTempDocxP = pathTempDocx;

            if (!Directory.Exists(pathTempFolder))
            {
                Directory.CreateDirectory(pathTempFolder);
            }

            if (!File.Exists(pathTempDocx))
            {
                File.Create(pathTempDocx);
            }

        }


        #region top bar
        private void bunifuImageButton2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void topPanel_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private void btnMini_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void btnMaxi_Click(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Maximized)
            {
                this.WindowState = FormWindowState.Normal;
            }
            else
            {
                //this.WindowState = FormWindowState.Maximized;
                this.MaximizedBounds = Screen.FromHandle(this.Handle).WorkingArea;
                this.WindowState = FormWindowState.Maximized;
            }
        }

        #endregion

        #region menu

        bool menuExpanded = true;

        private void timer1_Tick(object sender, EventArgs e)
        {
            //if (!bunifuTransition1.IsCompleted) { return; }
            //if (pSideMenu.ClientRectangle.Contains(PointToClient(Control.MousePosition)))
            //{
            //    if (!menuExpanded)
            //    {
            //        menuExpanded = true;
            //        pSideMenu.Visible = false;
            //        pSideMenu.Width = 140;
            //        bunifuTransition1.Show(pSideMenu);
            //    }
            //}
            //else
            //{
            //    if (menuExpanded)
            //    {
            //        menuExpanded = false;
            //        pSideMenu.Visible = false;
            //        pSideMenu.Width = 50;
            //        bunifuTransition1.Show(pSideMenu);
            //    }
            //}
        }

        private void openSubM(Panel subMenu)
        {
            if (subMenu.Visible == false)
            {
                bunifuTransition1.Show(subMenu);
            }
            else
            {
                bunifuTransition1.Hide(subMenu);
            }
        }
        private void closeAllSubMenu()
        {
            bunifuTransition1.Hide(pSubMenuProp);
        }

        private void btnDash_Click(object sender, EventArgs e)
        {
            closeAllSubMenu();
            pages.PageIndex = 0;
        }

        private void btnProp_Click(object sender, EventArgs e)
        {
            openSubM(pSubMenuProp);
        }

        private void bunifuImageButton1_Click(object sender, EventArgs e)
        {
            if (!menuExpanded)
            {
                menuExpanded = true;
                pSideMenu.Visible = false;
                pSideMenu.Width = 140;
                bunifuTransition1.Show(pSideMenu);
            }
            else
            {
                menuExpanded = false;
                pSideMenu.Visible = false;
                pSideMenu.Width = 50;
                bunifuTransition1.Show(pSideMenu);

            }
            
        }

        private void btnDadosProp_Click(object sender, EventArgs e)
        {
            closeAllSubMenu();
            pages.PageIndex = 1;
        }

        private void btnDadosCliente_Click(object sender, EventArgs e)
        {
            closeAllSubMenu();
            pages.PageIndex = 2;
        }

        private void bunifuFlatButton1_Click(object sender, EventArgs e)
        {
            pages.PageIndex = 3;
        }

        private void bunifuImageButton3_Click(object sender, EventArgs e)
        {
            pages.PageIndex = 2;
        }

        private void btnDadosPiscina_Click(object sender, EventArgs e)
        {
            closeAllSubMenu();
            pages.PageIndex = 4;
        }

        private void btnEquipamentos_Click(object sender, EventArgs e)
        {
            pages.PageIndex = 5;
            Thread.Sleep(1000);
            closeAllSubMenu();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            pages.PageIndex = 6;
            Thread.Sleep(1000);
            closeAllSubMenu();
        }

        private string pathFolder = Path.Combine(System.AppContext.BaseDirectory, "DBs");
        private string pathXLSDB = Path.Combine(System.AppContext.BaseDirectory, "DBs") + "\\xlsDB.xlsx";

        private void pages_SelectedIndexChanged(object sender, EventArgs e)
        {
            int currentPageIndex = pages.SelectedIndex;

            switch (currentPageIndex)
            {
                //DADOS PROPOSTA
                case 1:
                    if (string.IsNullOrWhiteSpace(tboxNumProp.Text))
                    {
                        tboxNumProp.Text = DateTime.Now.ToString("dd/MM/yy").Replace(@"/", "") + DateTime.Now.ToString("HH:mm").Replace(@":", "");
                    }
                    if (string.IsNullOrWhiteSpace(tboxData.Text))
                    {
                        tboxData.Text = DateTime.Now.ToString("dd/MM/yyyy");
                    }

                    break;

                //CLIENTE
                case 2:
                    break;

                //PROCURAR CLIENTE
                case 3:


                    if (!Directory.Exists(pathFolder))
                    {
                        Directory.CreateDirectory(pathFolder);
                    }

                    String name = "Clientes";
                    String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    pathXLSDB +
                                    ";Extended Properties='Excel 8.0;HDR=YES;';";

                    OleDbConnection con = new OleDbConnection(constr);
                    OleDbCommand oconn = new OleDbCommand("Select * From [" + name + "$]", con);
                    con.Open();

                    OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                    DataTable data = new DataTable();
                    sda.Fill(data);
                    dgvClientes.DataSource = data;

                    dgvClientes.RowHeadersVisible = false;
                    dgvClientes.ReadOnly = true;
                    break;

                //PISCINA
                case 4:

                    ddMoradaInst.AddItem(tboxMoradaCliente.Text);
                    ddMoradaInst.AddItem(tboxMorada2Cliente.Text);
                    tboxMoradaInst.Enabled = false;


                    break;

                //EQUIPAMENTO
                case 5:
                    ddDgvFilter.selectedIndex = 0;
                    fillDgv();


                    break;

                //EXPORT
                case 6:
                    //try
                    //{
                    //    object readOnly = false;
                    //    object visible = true;
                    //    object save = false;
                    //    object fileName = pathTempDocxP;
                    //    object newTemplate = false;
                    //    object docType = 0;
                    //    object missing = Type.Missing;

                    //    Microsoft.Office.Interop.Word._Document document;
                    //    Microsoft.Office.Interop.Word._Application application = new Microsoft.Office.Interop.Word.Application() { Visible = false };
                    //    document = application.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing,
                    //         ref missing, ref missing, ref missing, ref missing, ref visible, ref missing, ref missing, ref missing, ref missing);
                    //    document.ActiveWindow.Selection.WholeStory();
                    //    document.ActiveWindow.Selection.Copy();
                    //    IDataObject dataObject = Clipboard.GetDataObject();
                    //    rtboxPropPrev.Rtf = dataObject.GetData(DataFormats.Rtf).ToString();
                    //    application.Quit(ref missing, ref missing, ref missing);

                    //}
                    //catch (Exception exc)
                    //{
                    //    MessageBox.Show(exc.ToString());
                    //}

                    break;

                default:
                    break;
            }

        }

        #endregion

        #region Dados proposta
        //Dados da Proposta
        private void tboxNumProp_TextChange(object sender, EventArgs e)
        {
            if (tsNomeProp.Value == false)
            {
                tboxNomeProp.Text = tboxNumProp.Text;
            }
        }

        private void tsNomeProp_OnValuechange(object sender, EventArgs e)
        {
            if (tsNomeProp.Value == true)
            {
                tboxNomeProp.Enabled = true;
            }
            else
            {
                tboxNomeProp.Enabled = false;
                tboxNomeProp.Text = tboxNumProp.Text;
            }
        }

        private void tsNumProp_OnValuechange(object sender, EventArgs e)
        {
            if (tsNumProp.Value == true)
            {
                tboxNumProp.Enabled = true;
            }
            else
            {
                tboxNumProp.Enabled = false;
                tboxNumProp.Text = DateTime.Now.ToString("dd/MM/yy").Replace(@"/", "") + DateTime.Now.ToString("HH:mm").Replace(@":", "");
            }
        }

        private void tsData_OnValuechange(object sender, EventArgs e)
        {
            if (tsData.Value == true)
            {
                tboxData.Enabled = true;
            }
            else
            {
                tboxData.Enabled = false;
                tboxData.Text = DateTime.Now.ToString("dd/MM/yyyy");
            }
        }


        #endregion

        #region Cliente
        private void dgvClientes_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            string morada = dgvClientes.CurrentRow.Cells[9].Value.ToString() + ", " +
                dgvClientes.CurrentRow.Cells[3].Value.ToString() + ", " +
                dgvClientes.CurrentRow.Cells[8].Value.ToString();
            tboxNomeCliente.Text = dgvClientes.CurrentRow.Cells[1].Value.ToString();
            tboxNumCliente.Text = dgvClientes.CurrentRow.Cells[0].Value.ToString();
            tboxContactoCliente.Text = dgvClientes.CurrentRow.Cells[4].Value.ToString();
            tboxMailCliente.Text = dgvClientes.CurrentRow.Cells[13].Value.ToString();
            tboxMoradaCliente.Text = morada;
            tboxMorada2Cliente.Text = dgvClientes.CurrentRow.Cells[6].Value.ToString();
            tboxCoord.Text = dgvClientes.CurrentRow.Cells[12].Value.ToString();
            pages.PageIndex = 2;
        }

        private string imgClientePath = "";
        private void btnSearchImg_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            // image filters  
            open.Filter = "Image Files(*.jpg; *.jpeg; *.gif; *.bmp)|*.jpg; *.jpeg; *.gif; *.bmp";
            if (open.ShowDialog() == DialogResult.OK)
            {
                // display image in picture box  
                imgCliente.Image = new Bitmap(open.FileName);
                // image file path  
                imgClientePath = open.FileName;
            }
        }

        //Search client
        private void bunifuTextBox5_TextChanged(object sender, EventArgs e)
        {
            (dgvClientes.DataSource as DataTable).DefaultView.RowFilter = string.Format("Nome LIKE '%{0}%'", bunifuTextBox5.Text);

        }
        #endregion

        #region PISCINA

        #region PISCINA KEYPRESS

        private void tboxComp_KeyPress(object sender, KeyPressEventArgs e)
        {
            onlyNum(sender, e);
        }

        private void tboxLarg_KeyPress(object sender, KeyPressEventArgs e)
        {
            onlyNum(sender, e);

        }

        private void tboxProfMin_KeyPress(object sender, KeyPressEventArgs e)
        {
            onlyNum(sender, e);

        }

        private void tboxProfMax_KeyPress(object sender, KeyPressEventArgs e)
        {
            onlyNum(sender, e);

        }

        private void tboxProfMed_KeyPress(object sender, KeyPressEventArgs e)
        {
            onlyNum(sender, e);

        }

        private void tboxEscadas_KeyPress(object sender, KeyPressEventArgs e)
        {
            onlyNum(sender, e);

        }

        private void tboxSuperf_KeyPress(object sender, KeyPressEventArgs e)
        {
            onlyNum(sender, e);

        }

        private void tboxVol_KeyPress(object sender, KeyPressEventArgs e)
        {
            onlyNum(sender, e);

        }

        private void tboxArea_KeyPress(object sender, KeyPressEventArgs e)
        {
            onlyNum(sender, e);

        }

        private void tboxPerim_KeyPress(object sender, KeyPressEventArgs e)
        {
            onlyNum(sender, e);

        }

        private void ddMoradaInst_onItemSelected(object sender, EventArgs e)
        {
            if (ddMoradaInst.selectedValue == "Outra")
            {
                tboxMoradaInst.Enabled = true;
            }
            else
            {
                tboxMoradaInst.Text = "";
                tboxMoradaInst.Enabled = false;
            }
        }

        #endregion

        #region Contas

        private void ddFormato_onItemSelected(object sender, EventArgs e)
        {
            changeValues();
        }

        private void tboxComp_TextChanged(object sender, EventArgs e)
        {
            changeValues();
        }

        private void tboxLarg_TextChanged(object sender, EventArgs e)
        {
            changeValues();
        }

        private void tboxProfMin_TextChanged(object sender, EventArgs e)
        {
            changeValues();
        }

        private void tboxProfMax_TextChanged(object sender, EventArgs e)
        {
            changeValues();
        }
        #endregion



        #endregion

        #region EQUIPAMENTO

        private void tboxSearchDGV_TextChanged(object sender, EventArgs e)
        {
            (dgvDB.DataSource as DataTable).DefaultView.RowFilter = string.Format("Equipamento LIKE '%{0}%'", tboxSearchDGV.Text);
        }

        private void ddDgvFilter_onItemSelected(object sender, EventArgs e)
        {
            fillDgv();
        }

        private void dgvProp_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress -= new KeyPressEventHandler(Column2_KeyPress);
            if (dgvProp.CurrentCell.ColumnIndex == 2)
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress += new KeyPressEventHandler(Column2_KeyPress);
                }
            }
        }

        private void Column2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void dgvDB_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            addEquipamento();
        }

        private void bunifuFlatButton3_Click(object sender, EventArgs e)
        {
            try
            {
                dgvProp.Rows.RemoveAt(dgvProp.CurrentCell.RowIndex);
            }
            catch (Exception)
            {
                MessageBox.Show("Selecione um equipamento!");

            }
        }

        private void btnAddEquip_Click(object sender, EventArgs e)
        {
            addEquipamento();
        }

        private void addEquipamento()
        {
            DataGridViewRow r = dgvDB.CurrentRow;

            int index = dgvProp.Rows.Add(r.Cells[1].Value, r.Cells[2].Value, r.Cells[3].Value, r.Cells[4].Value, r.Cells[5].Value, r.Cells[6].Value, r.Cells[7].Value, r.Cells[8].Value);

            for (int row = 0; row < dgvProp.RowCount - 1; row++)
            {
                if (dgvProp.Rows[row].Cells[1].Value == dgvProp.Rows[index].Cells[1].Value && row != index)
                {
                    DataGridViewRow rowDuplicate = dgvProp.Rows[index];
                    dgvProp.Rows.Remove(rowDuplicate);
                    dgvProp[7, row].Value = Convert.ToInt32(dgvProp[7, row].Value) + 1;

                    dgvProp[8, row].Value = Convert.ToInt32(dgvProp[7, row].Value) * Convert.ToInt32(dgvProp[6, row].Value);
                }
            }
        }

        #endregion

        #region EXPORTAR
        private void btnExportProp_Click(object sender, EventArgs e)
        {
            //Preview 
            string filePath = "";
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.InitialDirectory = @"C:\";
                sfd.RestoreDirectory = true;
                sfd.FileName = tboxNomeProp.Text;
                sfd.Filter = "docx files (*.docx)|*.docx";
                sfd.DefaultExt = "docx";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    filePath = sfd.FileName;

                }
            }

            if (filePath != "")
            {
                string fileName = filePath;
                var doc = DocX.Create(fileName);
                //var imgLogoScale = System.Drawing.Image.FromFile("E:\\Projects\\AppVg\\Resources\\images\\logoVG.png");

                var path = Path.GetTempPath();
                Properties.Resources.logoVGscaled.Save(path + "\\logoVGscaled.png");

                //LOGO VG
                Xceed.Document.NET.Image img = doc.AddImage(path + "\\logoVGscaled.png");
                Picture pLogo = img.CreatePicture();
                Paragraph parLogo = doc.InsertParagraph("");
                parLogo.Alignment = Alignment.center;
                parLogo.AppendPicture(pLogo);

                doc.InsertParagraph();
                doc.InsertParagraph();
                doc.InsertParagraph();
                doc.InsertParagraph();

                //NUM PROPOSTA
                doc.InsertParagraph("Proposta VG - " + tboxNumProp.Text);

                doc.InsertParagraph();

                //DATA PROPOSTA
                doc.InsertParagraph("Data: " + tboxData.Text);

                doc.InsertParagraph();
                doc.InsertParagraph();
                
                //IMG CLIENTE
                if (imgClientePath != "")
                {
                    Xceed.Document.NET.Image imgCliente = doc.AddImage(imgClientePath);
                    Picture pCliente = imgCliente.CreatePicture();
                    Paragraph par = doc.InsertParagraph("Cliente: ");
                    par.AppendPicture(pCliente);
                    doc.InsertParagraph();
                }

                //DADOS CLIENTE
                doc.InsertParagraph("Nome: " + tboxNomeCliente.Text);
                doc.InsertParagraph();
                doc.InsertParagraph("Tlm: " + tboxContactoCliente.Text);
                doc.InsertParagraph();
                doc.InsertParagraph("E-mail: " + tboxMailCliente.Text);
                doc.InsertParagraph();
                doc.InsertParagraph("Morada: " + tboxMoradaCliente.Text);
                doc.InsertParagraph();
                doc.InsertParagraph(tboxTextoAbertura.Text);

                doc.InsertSectionPageBreak();

                //DADOS PISCINA
                doc.InsertParagraph("A proposta, está baseada nos elementos fornecidos à VITORGEST, e têm as seguintes características:");
                doc.InsertParagraph();
                doc.InsertParagraph();

                Table t = doc.AddTable(2, 6);
                t.Alignment = Alignment.center;
                t.Design = TableDesign.MediumGrid3Accent5;

                t.Rows[0].Cells[0].Paragraphs.First().Append("Forma");
                t.Rows[0].Cells[1].Paragraphs.First().Append("Dimensões");
                t.Rows[0].Cells[2].Paragraphs.First().Append("Área");
                t.Rows[0].Cells[3].Paragraphs.First().Append("Prof. Max.");
                t.Rows[0].Cells[4].Paragraphs.First().Append("Prof. Min.");
                t.Rows[0].Cells[5].Paragraphs.First().Append("Volume");

                t.Rows[1].Cells[0].Paragraphs.First().Append(ddFormato.selectedValue);
                t.Rows[1].Cells[1].Paragraphs.First().Append(tboxComp.Text + " x " + tboxLarg.Text);
                t.Rows[1].Cells[2].Paragraphs.First().Append(tboxArea.Text + "m^2");
                t.Rows[1].Cells[3].Paragraphs.First().Append(tboxProfMax.Text + "m");
                t.Rows[1].Cells[4].Paragraphs.First().Append(tboxProfMin.Text + "m");
                t.Rows[1].Cells[5].Paragraphs.First().Append(tboxVol.Text + "m^3");

                doc.InsertTable(t);

                doc.InsertParagraph();
                doc.InsertParagraph("Localização da Obra: " + ddMoradaInst.selectedValue);
                doc.InsertParagraph();

                doc.InsertSectionPageBreak();

                //EQUIPAMENTO
                doc.InsertParagraph();
                doc.InsertParagraph("Fornecimento de Equipamento e Instalação");
                doc.InsertParagraph();

                Table tEquipamento = doc.AddTable(dgvProp.Rows.Count, 3);
                tEquipamento.Alignment = Alignment.center;
                tEquipamento.Design = TableDesign.MediumGrid3Accent5;

                for (int i = 0; i < dgvProp.Rows.Count; i++)
                {
                    for (int j = 1; j < 3; j++)
                    {
                        tEquipamento.Rows[i].Cells[j].Paragraphs.First().Append(dgvProp[j, i].Value.ToString());
                    }
                }

                doc.InsertTable(tEquipamento);

                //VALORES
                doc.InsertParagraph();
                doc.InsertParagraph("I- Valores");
                doc.InsertParagraph("Transporte incluído.");
                doc.InsertParagraph();

                Table tValores = doc.AddTable(4, 2);
                tValores.Alignment = Alignment.center;
                tValores.Design = TableDesign.MediumGrid3Accent5;

                tValores.Rows[0].Cells[0].Paragraphs.First().Append("Equipamento");
                tValores.Rows[1].Cells[0].Paragraphs.First().Append("Revestimento");
                tValores.Rows[2].Cells[0].Paragraphs.First().Append("Tratamento de Água");
                tValores.Rows[3].Cells[0].Paragraphs.First().Append("Total");

                int valorEquipamento= 0;
                int valorRevestimento = 0;
                int valorTratamento= 0;
                for (int i = 0; i < dgvProp.Rows.Count; i++)
                {
                    valorEquipamento += Convert.ToInt32(dgvProp[7, i].Value.ToString());
                }

                int valorTotal = valorEquipamento + valorRevestimento + valorTratamento;

                tValores.Rows[0].Cells[1].Paragraphs.First().Append(valorEquipamento.ToString());
                tValores.Rows[1].Cells[1].Paragraphs.First().Append(valorRevestimento.ToString());
                tValores.Rows[2].Cells[1].Paragraphs.First().Append(valorTratamento.ToString());
                tValores.Rows[3].Cells[1].Paragraphs.First().Append(valorTotal.ToString());
                doc.InsertTable(tValores);
                doc.InsertParagraph("A este valor acresce a taxa de IVA em vigor.");

                //CONDICÕES
                doc.InsertParagraph();
                doc.InsertParagraph("II- Condições");
                doc.InsertParagraph("30% Com adjudicação");
                doc.InsertParagraph("30% Com a entrega do equipamento");
                doc.InsertParagraph("30% Na conclusão da instalação/revestimento");
                doc.InsertParagraph("10% No Final");
                doc.InsertParagraph();
                
                //GARANTIAS
                doc.InsertParagraph();
                doc.InsertParagraph("III- Garantias*");

                Table tGarantia = doc.AddTable(5, 2);
                tGarantia.Alignment = Alignment.center;
                tGarantia.Design = TableDesign.MediumGrid3Accent5;

                tGarantia.Rows[0].Cells[0].Paragraphs.First().Append("Filtros e Bombas");
                tGarantia.Rows[1].Cells[0].Paragraphs.First().Append("Equipamento de limpeza robot");
                tGarantia.Rows[2].Cells[0].Paragraphs.First().Append("Tratamento de água");
                tGarantia.Rows[3].Cells[0].Paragraphs.First().Append("Quadros Elétricos e Equipamentos");
                tGarantia.Rows[4].Cells[0].Paragraphs.First().Append("Revestimento em Tela Armada (estanquicidade)");

                tGarantia.Rows[0].Cells[1].Paragraphs.First().Append("3 Anos");
                tGarantia.Rows[1].Cells[1].Paragraphs.First().Append("3 Anos");
                tGarantia.Rows[2].Cells[1].Paragraphs.First().Append("3 Anos");
                tGarantia.Rows[3].Cells[1].Paragraphs.First().Append("2 Anos");
                tGarantia.Rows[4].Cells[1].Paragraphs.First().Append("10 Anos");

                doc.InsertTable(tGarantia);
                doc.InsertParagraph("*(Conforme norma dos fabricantes)");
                doc.InsertParagraph();

                //VALIDADE
                doc.InsertParagraph();
                doc.InsertParagraph("IV- Validade");
                doc.InsertParagraph("Esta proposta é válida por 60 dias");
                doc.InsertParagraph();

                //ADJUDICAÇÃO
                doc.InsertParagraph();
                doc.InsertParagraph("V- Adjudicação");
                doc.InsertParagraph("Esta proposta é válida por 60 dias");

                doc.Save();

                MessageBox.Show("Proposta criada com sucesso!");
            }


        }

        #endregion

        #region OUTRAS FUNÇÕES
        public static System.Drawing.Image ScaleImage(System.Drawing.Image image, int maxWidth, int maxHeight)
        {
            var ratioX = (double)maxWidth / image.Width;
            var ratioY = (double)maxHeight / image.Height;
            var ratio = Math.Min(ratioX, ratioY);

            var newWidth = (int)(image.Width * ratio);
            var newHeight = (int)(image.Height * ratio);

            var newImage = new Bitmap(newWidth, newHeight);

            using (var graphics = Graphics.FromImage(newImage))
                graphics.DrawImage(image, 0, 0, newWidth, newHeight);

            return newImage;
        }

        private void dgvProp_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            for (int row = 0; row < dgvProp.RowCount - 1; row++)
            {
                dgvProp[8, row].Value = Convert.ToInt32(dgvProp[7, row].Value) * Convert.ToInt32(dgvProp[6, row].Value);
            }
        }

        private void fillDgv()
        {
            String name = ddDgvFilter.selectedValue.ToString();
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            pathXLSDB +
                            ";Extended Properties='Excel 8.0;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select * From [" + name + "$]", con);
            con.Open();

            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataTable data = new DataTable();
            sda.Fill(data);
            dgvDB.DataSource = data;

            if (dgvProp.Columns.Count < dgvDB.Columns.Count)
            {
                foreach (DataGridViewColumn c in dgvDB.Columns)
                {
                    dgvProp.Columns.Add(c.Clone() as DataGridViewColumn);
                }
                dgvDB.ReadOnly = true;
                dgvProp.ReadOnly = false;

                for (int i = 0; i < dgvProp.Columns.Count; i++)
                {
                    if (i != 7)
                    {
                        dgvProp.Columns[i].ReadOnly = true;
                    }
                }
                
                dgvDB.RowHeadersVisible = false;
                dgvProp.RowHeadersVisible = false;

                dgvDB.Columns["img"].Visible = false;
                dgvDB.Columns["Marca"].Visible = false;
                dgvDB.Columns["Obs"].Visible = false;
                dgvDB.Columns["Quantidade"].Visible = false;
                dgvDB.Columns["Total"].Visible = false;
                dgvDB.Columns["Preco"].Visible = false;
                //dgvDB.Columns["f7"].Visible = false;

                dgvProp.Columns["img"].Visible = false;
                dgvProp.Columns["Marca"].Visible = false;
                dgvProp.Columns["Obs"].Visible = false;
                //dgvProp.Columns["f7"].Visible = false;


            }
        }

        private void onlyNum(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '.')
            {
                e.Handled = true;
            }

            if (e.KeyChar == '.'
                && (sender as TextBox).Text.IndexOf('.') > -1)
            {
                e.Handled = true;
            }
        }

        private void changeValues()
        {

            switch (ddFormato.selectedIndex)
            {
                //Rectangulo
                case 0:
                    labelRaio.Visible = false;
                    labelComp.Visible = true; 

                    tboxComp.Enabled = true;
                    tboxLarg.Enabled = true;
                    tboxProfMax.Enabled = true;
                    tboxProfMin.Enabled = true;
                    if (!String.IsNullOrEmpty(tboxComp.Text) && !String.IsNullOrEmpty(tboxLarg.Text) && !String.IsNullOrEmpty(tboxProfMin.Text) && !String.IsNullOrEmpty(tboxProfMax.Text))
                    {
                        double comprimento = Double.Parse(tboxComp.Text);
                        double largura = Double.Parse(tboxLarg.Text);


                        double profMin = Double.Parse(tboxProfMin.Text);
                        double profMax = Double.Parse(tboxProfMax.Text);
                        double profMed = (profMin + profMax) / 2;

                        double superficie = comprimento * largura;

                        double volume = comprimento * largura * profMed;

                        double area = ((comprimento * profMed * 2) + (largura * profMed * 2) + superficie) * 1.1;

                        double perimetro = (comprimento + largura) * 2;

                        tboxSuperf.Text = superficie.ToString();
                        tboxVol.Text = volume.ToString();
                        tboxProfMed.Text = profMed.ToString();
                        tboxArea.Text = area.ToString();
                        tboxPerim.Text = perimetro.ToString();

                    }
                    break;

                //circulo
                case 1:
                    labelRaio.Visible = true;
                    labelComp.Visible = false;

                    tboxComp.Enabled = true;
                    tboxLarg.Enabled = false;
                    tboxProfMax.Enabled = true;
                    tboxProfMin.Enabled = true;

                    double profMinC = Double.Parse(tboxProfMin.Text);
                    double profMaxC = Double.Parse(tboxProfMax.Text);
                    double profMedC = (profMinC + profMaxC) / 2;
                    double raio = Double.Parse(tboxComp.Text);
                    double areaC = Math.PI * (raio * raio);
                    double volC = areaC * profMedC;
                    double perimetroC = 2 * Math.PI * raio;
                    double areaCtotal = (2 * Math.PI * raio * (profMedC + raio)) * 1.1;

                    tboxProfMed.Text = Math.Round(profMedC, 2).ToString();
                    tboxSuperf.Text = Math.Round(areaC, 2).ToString();
                    tboxVol.Text = Math.Round(volC, 2).ToString();
                    tboxPerim.Text = Math.Round(perimetroC, 2).ToString();
                    tboxArea.Text = Math.Round(areaCtotal, 2).ToString();



                    break;

                default:
                    labelRaio.Visible = false;
                    labelComp.Visible = true;

                    tboxComp.Enabled = false;
                    tboxLarg.Enabled = false;
                    tboxProfMax.Enabled = false;
                    tboxProfMin.Enabled = false;

                    break;
            }

           
        }










        #endregion

       
    }
}
