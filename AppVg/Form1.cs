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

        public Form1()
        {
            InitializeComponent();
            pSubMenuProp.Width = 167;
            pSubMenuProp.Visible = false;

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


                    tboxMoradaInst.Enabled = false;


                    break;

                //EQUIPAMENTO
                case 5:
                    ddDgvFilter.selectedIndex = 0;
                    fillDgv();


                    break;

                //EXPORT
                case 6:
                    string propPrev = "Número da Proposta: " + tboxNumProp.Text + "\n" +
                         "Data: " + tboxData.Text + "\n" +
                         "Texto de Abertura: " + tboxTextoAbertura.Text + "\n" +
                         "Cliente: " + tboxNomeCliente.Text + "\n" +
                         "Contacto: " + tboxContactoCliente.Text + "\n";
                    //Teste

                    rtboxPropPrev.Text = propPrev;

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
            string morada = dgvClientes.CurrentRow.Cells[3].Value.ToString() + ", " +
                dgvClientes.CurrentRow.Cells[9].Value.ToString() + ", " +
                dgvClientes.CurrentRow.Cells[8].Value.ToString();

            tboxNomeCliente.Text = dgvClientes.CurrentRow.Cells[1].Value.ToString();
            tboxNumCliente.Text = dgvClientes.CurrentRow.Cells[0].Value.ToString();
            tboxContactoCliente.Text = dgvClientes.CurrentRow.Cells[4].Value.ToString();
            tboxMailCliente.Text = dgvClientes.CurrentRow.Cells[13].Value.ToString();
            tboxMoradaCliente.Text = morada;
            tboxMorada2Cliente.Text = dgvClientes.CurrentRow.Cells[6].Value.ToString();
            tboxCoord.Text = dgvClientes.CurrentRow.Cells[12].Value.ToString();
            pages.PageIndex = 2;

            ddMoradaInst.Clear();

            ddMoradaInst.AddItem(tboxMoradaCliente.Text);
            ddMoradaInst.AddItem(tboxMorada2Cliente.Text);
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
            (dgvDB.DataSource as DataTable).DefaultView.RowFilter = string.Format("Nome LIKE '%{0}%'", tboxSearchDGV.Text);
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
            try
            {
                DataGridViewRow r = dgvDB.CurrentRow;

                int index = dgvProp.Rows.Add(r.Cells[0].Value, r.Cells[1].Value, r.Cells[2].Value, r.Cells[3].Value, r.Cells[4].Value, r.Cells[5].Value, r.Cells[6].Value, r.Cells[7].Value, r.Cells[8].Value);

                for (int row = 0; row < dgvProp.RowCount - 1; row++)
                {
                    if (dgvProp.Rows[row].Cells[0].Value == dgvProp.Rows[index].Cells[0].Value && row != index)
                    {
                        DataGridViewRow rowDuplicate = dgvProp.Rows[index];
                        dgvProp.Rows.Remove(rowDuplicate);
                        dgvProp[6, row].Value = Convert.ToInt32(dgvProp[6, row].Value) + 1;

                        dgvProp[7, row].Value = Convert.ToDouble(dgvProp[6, row].Value) * Convert.ToDouble(dgvProp[5, row].Value);
                    }
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }
            
        }

        #endregion

        #region EXPORTAR

        string filePathPdf = "";
        private void btnExportProp_Click(object sender, EventArgs e)
        {
            try
            {
                
                //Prop file path
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
                        filePathPdf = filePath;
                    }
                }

                if (filePath != "")
                {
                    string fileName = filePath;
                    var doc = DocX.Create(fileName);
                    //var imgLogoScale = System.Drawing.Image.FromFile("E:\\Projects\\AppVg\\Resources\\images\\logoVG.png");

                    var path = Path.GetTempPath();
                    Properties.Resources.logoVGscaled.Save(path + "\\logoVGscaled.png");



                    doc.DifferentOddAndEvenPages = false;
                    doc.DifferentFirstPage = false;
                    doc.AddHeaders();
                    doc.AddFooters();

                    //HEADER

                    Table tHeader = doc.AddTable(1, 2);
                    tHeader.Alignment = Alignment.center;
                    tHeader.Design = TableDesign.Custom;

                    tHeader.Rows[0].Cells[0].Paragraphs.First().Append("Original").Bold().Alignment = Alignment.left;
                    tHeader.Rows[0].Cells[1].Paragraphs.First().Append("Proposta VG-" + tboxNumProp.Text).Alignment = Alignment.right;
                    tHeader.SetColumnWidth(0, 250);
                    tHeader.SetColumnWidth(1, 250);

                    doc.Headers.Odd.InsertTable(tHeader);

                    //FOOTER
                    doc.Footers.Odd.InsertParagraph("Rua dos Aventureiros Lote 19ª – Parque das Nações – 1990-024 Lisboa\nTel. 218 940 990 - Telm. 935 809 381 - Fax. 218 940 992 - Email: info@vitorgest.pt").Alignment = Alignment.center;

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

                    //NUM PROP E DATA

                    Table tNumData = doc.AddTable(1, 2);
                    tNumData.Alignment = Alignment.center;
                    tNumData.Design = TableDesign.Custom;

                    tNumData.Rows[0].Cells[0].Paragraphs.First().Append("Proposta VG - " + tboxNumProp.Text).Bold().FontSize(12).Alignment = Alignment.left;
                    tNumData.Rows[0].Cells[1].Paragraphs.First().Append("Data: " + tboxData.Text).Alignment = Alignment.right;
                    tNumData.SetColumnWidth(0, 250);
                    tNumData.SetColumnWidth(1, 250);

                    doc.InsertTable(tNumData);

                    doc.InsertParagraph();
                    doc.InsertParagraph();
                
                    //IMG CLIENTE
                    if (imgClientePath != "")
                    {
                        var imgCliente = doc.AddImage(imgClientePath);
                        Picture pCliente = imgCliente.CreatePicture(100, 100);
                        Paragraph par = doc.InsertParagraph();
                        par.AppendPicture(pCliente);
                        doc.InsertParagraph();
                    }

                    //DADOS CLIENTE
                    doc.InsertParagraph("Cliente: ").Bold().Append(tboxNomeCliente.Text);
                    doc.InsertParagraph();
                    doc.InsertParagraph("TLM: ").Bold().Append(tboxContactoCliente.Text);
                    doc.InsertParagraph();
                    doc.InsertParagraph("EMAIL: ").Bold().Append(tboxMailCliente.Text);
                    doc.InsertParagraph();
                    doc.InsertParagraph("MORADA: ").Bold().Append(tboxMoradaCliente.Text);
                    doc.InsertParagraph();
                    doc.InsertParagraph();
                    doc.InsertParagraph(tboxTextoAbertura.Text);
                    doc.InsertParagraph("Vitor Filipe").Font(new Xceed.Document.NET.Font("Freestyle Script")).Color(Color.Blue).Italic().FontSize(20);

                    doc.InsertParagraph().InsertPageBreakAfterSelf();

                    //DADOS PISCINA
                    doc.InsertParagraph("A proposta, está baseada nos elementos fornecidos à VITORGEST, e têm as seguintes características:");
                    doc.InsertParagraph();
                    doc.InsertParagraph();

                    Table t = doc.AddTable(2, 8);
                    t.Alignment = Alignment.center;
                    t.Design = TableDesign.LightListAccent1;

                    t.Rows[0].Cells[0].Paragraphs.First().Append("Designação");
                    t.Rows[0].Cells[1].Paragraphs.First().Append("Número");
                    t.Rows[0].Cells[2].Paragraphs.First().Append("Área");
                    t.Rows[0].Cells[3].Paragraphs.First().Append("Forma");
                    t.Rows[0].Cells[4].Paragraphs.First().Append("Prof. Máx.");
                    t.Rows[0].Cells[5].Paragraphs.First().Append("Prof. Min.");
                    t.Rows[0].Cells[6].Paragraphs.First().Append("Volume");
                    t.Rows[0].Cells[7].Paragraphs.First().Append("Tipo de Circulação");

                    t.Rows[1].Cells[0].Paragraphs.First().Append("Piscina\n" + tboxComp.Text + "x" + tboxLarg.Text);
                    t.Rows[1].Cells[1].Paragraphs.First().Append("1");
                    t.Rows[1].Cells[2].Paragraphs.First().Append(tboxSuperf.Text + "m^2");
                    if (ddFormato.selectedIndex == -1)
                    {
                        t.Rows[1].Cells[3].Paragraphs.First().Append("Não definido");
                    }
                    else
                    {
                        t.Rows[1].Cells[3].Paragraphs.First().Append(ddFormato.selectedValue);
                    }
                    t.Rows[1].Cells[4].Paragraphs.First().Append(tboxProfMax.Text + "m");
                    t.Rows[1].Cells[5].Paragraphs.First().Append(tboxProfMin.Text + "m");
                    t.Rows[1].Cells[6].Paragraphs.First().Append(tboxVol.Text + "m^3");

                    if (ddCirculação.selectedIndex == -1)
                    {
                        t.Rows[1].Cells[7].Paragraphs.First().Append("Não definido");
                    }
                    else
                    {
                        t.Rows[1].Cells[7].Paragraphs.First().Append(ddFormato.selectedValue);
                    }

                    doc.InsertTable(t);

                    doc.InsertParagraph();
                    if(ddMoradaInst.selectedIndex == -1)
                    {
                        doc.InsertParagraph("Localização da Obra: ").Append("Não definido").FontSize(12);
                    }
                    else if (ddMoradaInst.selectedValue == "Outra")
                    {
                        doc.InsertParagraph("Localização da Obra: ").Append(tboxMoradaInst.Text).FontSize(12);
                    }
                    else
                    {
                        doc.InsertParagraph("Localização da Obra: ").Append(ddMoradaInst.selectedValue).FontSize(12);
                    }
                    doc.InsertParagraph();

                    doc.InsertParagraph("Escavação: ").Append(ddEscavacao.selectedValue).FontSize(12);
                    doc.InsertParagraph();
                    doc.InsertParagraph("Estrutura: ").Append(ddEstrutura.selectedValue).FontSize(12);
                    doc.InsertParagraph();
                    doc.InsertParagraph("Revestimento: ").Append(ddRevestimento.selectedValue).FontSize(12);
                    doc.InsertParagraph();
                    doc.InsertParagraph("Equipamento de Filtração: ").Append("Será fornecido e instalado todo o equipamento previsto nesta proposta de acordo com as normas em vigor.").FontSize(12);
                    doc.InsertParagraph();
                    doc.InsertParagraph("Equipamento Suplementar: ").Append("Ver proposta em anexo.").FontSize(12);
                    doc.InsertParagraph();
                    doc.InsertParagraph("Ligações á rede: ").Append(ddLigRed.selectedValue).FontSize(12);
                    doc.InsertParagraph();
                    doc.InsertParagraph("Casa das Máquinas: ").Append(ddCMaqui.selectedValue).FontSize(12);
                    doc.InsertParagraph();
                    doc.InsertParagraph("Tanque de Compensação: ").Append(ddTComp.selectedValue).FontSize(12);
                    doc.InsertParagraph();
                    doc.InsertParagraph("Acabamento final: ").Append("Enchimento da piscina (água a cargo do Cliente). Colocação dos equipamentos em funcionamento. Testes e ensaios. Entrega da piscina pronta a utilizar.").FontSize(12);

                    doc.InsertParagraph().InsertPageBreakAfterSelf();

                    //EQUIPAMENTO
                    doc.InsertParagraph();
                    doc.InsertParagraph("Fornecimento de Equipamento e Instalação").FontSize(13);
                    doc.InsertParagraph();

                    //TABELAS

                    string imgDBpath = Path.Combine(System.AppContext.BaseDirectory, "imgDB");

                    for (int i = 0; i < dgvProp.Rows.Count; i++)
                    {
                        Table tEqui = doc.AddTable(1, 2);
                        tEqui.Design = TableDesign.Custom;
                        tEqui.Alignment = Alignment.center;

                        doc.InsertParagraph(dgvProp[1, i].Value.ToString()).Bold().Alignment = Alignment.center;

                        Xceed.Document.NET.Image equiImg = doc.AddImage(imgDBpath + "\\" + dgvProp[8, i].Value.ToString() + ".png");
                        Picture pEqui = equiImg.CreatePicture();
                        tEqui.Rows[0].Cells[0].Paragraphs.First().AppendPicture(pEqui).Alignment = Alignment.right;

                        string[] equiDesc = dgvProp[2, i].Value.ToString().Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

                        var bulletedList = doc.AddList(equiDesc[0], 0, ListItemType.Bulleted);
                        for (int j = 1; j < equiDesc.Length; j++)
                        {
                            doc.AddListItem(bulletedList, equiDesc[j]);
                        }

                        tEqui.Rows[0].Cells[1].Paragraphs.First().InsertListAfterSelf(bulletedList);

                        tEqui.SetColumnWidth(0, 150);
                        tEqui.SetColumnWidth(1, 250);
                        doc.InsertTable(tEqui);

                        doc.InsertParagraph();
                        doc.InsertParagraph();
                        doc.InsertParagraph();
                    }

                    //VALORES
                    doc.InsertParagraph();
                    doc.InsertParagraph("I- VALORES").FontSize(13);
                    doc.InsertParagraph();
                    doc.InsertParagraph("Transporte incluído.");
                    doc.InsertParagraph();

                    Table tValores = doc.AddTable(4, 2);
                    tValores.Alignment = Alignment.center;
                    tValores.Design = TableDesign.MediumGrid3Accent5;

                    tValores.Rows[0].Cells[0].Paragraphs.First().Append("Equipamento");
                    tValores.Rows[1].Cells[0].Paragraphs.First().Append("Revestimento");
                    tValores.Rows[2].Cells[0].Paragraphs.First().Append("Tratamento de Água");
                    tValores.Rows[3].Cells[0].Paragraphs.First().Append("Total");

                    double valorEquipamento= 0;
                    double valorRevestimento = 0;
                    double valorTratamento= 0;
                    for (int i = 0; i < dgvProp.Rows.Count; i++)
                    {
                        valorEquipamento += Convert.ToDouble(dgvProp[7, i].Value.ToString());
                    }

                    double valorTotal = valorEquipamento + valorRevestimento + valorTratamento;

                    tValores.Rows[0].Cells[1].Paragraphs.First().Append(valorEquipamento.ToString());
                    tValores.Rows[1].Cells[1].Paragraphs.First().Append(valorRevestimento.ToString());
                    tValores.Rows[2].Cells[1].Paragraphs.First().Append(valorTratamento.ToString());
                    tValores.Rows[3].Cells[1].Paragraphs.First().Append(valorTotal.ToString());
                    doc.InsertTable(tValores);
                    doc.InsertParagraph("A este valor acresce a taxa de IVA em vigor.");

                    //CONDICÕES
                    doc.InsertParagraph();
                    doc.InsertParagraph("II- CONDIÇÕES").FontSize(13);
                    doc.InsertParagraph();
                    doc.InsertParagraph("30% Com adjudicação");
                    doc.InsertParagraph("30% Com a entrega do equipamento");
                    doc.InsertParagraph("30% Na conclusão da instalação/revestimento");
                    doc.InsertParagraph("10% No Final");
                    doc.InsertParagraph();
                
                    //GARANTIAS
                    doc.InsertParagraph();
                    doc.InsertParagraph("III- GARANTIAS*").FontSize(13);
                    doc.InsertParagraph();

                    string garantia1 = "Filtros e Bombas..................................................................";
                    string garantia2 = "Equipamento de limpeza robot...........................................";
                    string garantia3 = "Tratamento de água............................................................";
                    string garantia4 = "Quadros Elétricos e Equipamentos.....................................";
                    string garantia5 = "Revestimento em Tela Armada (estanquicidade).............";

                    doc.InsertParagraph(garantia1).Append("3 Anos").Bold();
                    doc.InsertParagraph(garantia2).Append("3 Anos").Bold();
                    doc.InsertParagraph(garantia3).Append("2 Anos").Bold();
                    doc.InsertParagraph(garantia4).Append("2 Anos").Bold();
                    doc.InsertParagraph(garantia5).Append("10 Anos").Bold();

                    doc.InsertParagraph("*(Conforme norma dos fabricantes)");
                    doc.InsertParagraph();

                    //VALIDADE
                    doc.InsertParagraph();
                    doc.InsertParagraph("IV- VALIDADE").FontSize(13);
                    doc.InsertParagraph();
                    doc.InsertParagraph("Esta proposta é válida por 60 dias");
                    doc.InsertParagraph();

                    //ADJUDICAÇÃO
                    doc.InsertParagraph();
                    doc.InsertParagraph("V- ADJUDICAÇÃO").FontSize(13);
                    doc.InsertParagraph();

                    Table tAdju = doc.AddTable(3, 2);
                    tAdju.Alignment = Alignment.center;
                    tAdju.Design = TableDesign.Custom;
                    tAdju.Rows[0].Cells[0].Paragraphs.First().Append("VITORGEST, Lda.\n\n");
                    tAdju.Rows[0].Cells[1].Paragraphs.First().Append("Cliente\n\n");
                    tAdju.Rows[1].Cells[0].Paragraphs.First().Append("Vitor Filipe").Font( new Xceed.Document.NET.Font("Freestyle Script")).Color(Color.Blue).Italic().FontSize(20);
                    tAdju.Rows[1].Cells[1].Paragraphs.First().Append("_______________________________");
                    tAdju.Rows[2].Cells[0].Paragraphs.First().Append("(T - 935 809 380)");
                    tAdju.Rows[2].Cells[1].Paragraphs.First().Append("(concordo com as condições desta proposta)");
                    doc.InsertTable(tAdju);

                    doc.InsertParagraph();
                    doc.InsertParagraph();
                    doc.InsertParagraph("Lisboa,____de_______________ de _____");

                    doc.InsertParagraph().InsertPageBreakAfterSelf();
                    doc.InsertParagraph();
                    doc.InsertParagraph();

                    string[] lastPageText = System.IO.File.ReadAllLines(Path.Combine(System.AppContext.BaseDirectory, "DBs") + "\\lastPage.txt");

                    foreach(string line in lastPageText)
                    {
                        doc.InsertParagraph(line).FontSize(10);
                    }

                    doc.InsertParagraph("Vitor Filipe").Font(new Xceed.Document.NET.Font("Freestyle Script")).Color(Color.Blue).Italic().FontSize(20);
                    doc.InsertParagraph("(T - 935 809 380)").FontSize(10);

                    doc.Save();

                    MessageBox.Show("Proposta criada com sucesso!");
                }

                }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
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

                dgvDB.Columns["Marca"].Visible = false;
                dgvDB.Columns["Obs"].Visible = false;
                dgvDB.Columns["Quantidade"].Visible = false;
                dgvDB.Columns["Total"].Visible = false;
                dgvDB.Columns["Preco"].Visible = false;
                dgvDB.Columns[8].Visible = false;

                dgvProp.Columns["Marca"].Visible = false;
                dgvProp.Columns["Obs"].Visible = false;
                dgvProp.Columns[8].Visible = false;

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

            try
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

                        tboxComp.Enabled = true;
                        tboxLarg.Enabled = true;
                        tboxProfMax.Enabled = true;
                        tboxProfMin.Enabled = true;

                        break;
                }

            }
            catch (Exception excep)
            {

            }



           
        }


        #endregion


    }
}
