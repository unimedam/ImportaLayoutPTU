using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;

namespace prjImporta_LayoutPTU {

    public partial class frmLayout_PTU : Form {
        // Criar conexão com o banco
        private OleDbConnection msConnection = new OleDbConnection();
        private string conStringSeven = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\" + Environment.UserName.ToString() + "\\AppData\\Local\\cardPresso\\DATABASE\\internalDatabase.mdb;";
        private string conStringXP = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Documents and Settings\\" + Environment.UserName.ToString() + "\\Configurações locais\\Dados de aplicativos\\cardPresso\\DATABASE\\internalDatabase.mdb;";


        // Auxiliares
        public OpenFileDialog OFDialog = new OpenFileDialog();

        // Função para verificar registros existentes,
        public void verificaRegistros() {
            try {
                msConnection.Open();
                OleDbCommand msComand = new OleDbCommand();
                msComand.Connection = msConnection;
                msComand.CommandText = "select * from Layout_PTU";

                OleDbDataReader reader = msComand.ExecuteReader();

                int count = 0;
                while (reader.Read()) {
                    count++;
                }

                if (count >= 1) {
                    lblDados.Text = "Quantidade de registros no banco: " + count.ToString();
                } else {
                    lblDados.Text = "Não existe registros no banco.";
                }

                msConnection.Close();

            } catch (Exception ex) {
                MessageBox.Show("Erro ao verificar registros, mensagem do banco: " + ex.ToString());
            }
        }

        public frmLayout_PTU() {
            InitializeComponent();
            // Verifica a versão do Windows e atribui a String de conexão para aquele OS .
            if (int.Parse(Environment.OSVersion.ToString().Substring(21, 1)) >= 6) // Vindows 7 e superiores
            {
                msConnection.ConnectionString = conStringSeven; // Inicializa string de conexão   
            }
            else // Vista e inferiores
            {
                msConnection.ConnectionString = conStringXP;
            }
        }


        private void frmLayout_PTU_Load(object sender, EventArgs e) {

            // Testar conexão com o banco
            try {
                msConnection.Open();
                lblStatus.Text = "Conectado";
                msConnection.Close();

                //Verificar se existe registros
                verificaRegistros();
            }
            catch (Exception ex) {
                lblStatus.Text = "Não conectado";
                MessageBox.Show("Não foi possivel fazer a conexão com o banco,\n verifique se o caminho esta correto. \n \n O aplicativo será encerrado.",
                    "Erro !",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                MessageBox.Show("Erro de banco: \n \n" + ex.Message, "Error !", MessageBoxButtons.OK, MessageBoxIcon.Error) ;

                btnLocalizar.Enabled = false;
                btnApagar.Enabled = false;
                btnImportar.Enabled = false;

                this.Close();
            }

        }

        private void btnLocalizar_Click(object sender, EventArgs e) {
            // Pega o caminho do arquivo verificando se esta no formato .TXT
            string nomeArquivo = "";
            do {
                // Pega o caminho do arquivo e o nome
                if (OFDialog.ShowDialog() != DialogResult.Cancel) {
                    txtCaminhoArquivoImportado.Text = OFDialog.FileName;
                    string[] path = txtCaminhoArquivoImportado.ToString().Split('\\'); // Quebra o caminho do arquivo em um vetor para pegar somente o index que conter o nome.
                    nomeArquivo = path[path.Length - 1]; //pega o ultimo index que é referente ao nome do arquivo.
                    //nomeArquivo = OFDialog.
                }
                else {
                    txtCaminhoArquivoImportado.Text = "";
                    nomeArquivo = "";
                    break;
                }

                // Se o arquivo for diferente exibe a mensagem de erro.
                if (nomeArquivo.Substring(nomeArquivo.Length - 3, 3) != "TXT" && nomeArquivo.Substring(nomeArquivo.Length - 3, 3) != "txt") {
                    //MessageBox.Show(nomeArquivo);
                    MessageBox.Show(null, "Arquivo no formato invalido. \nO formato do arquivo deve ser .TXT !!!", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            } while (nomeArquivo.Substring(nomeArquivo.Length - 3, 3) != "TXT" && nomeArquivo.Substring(nomeArquivo.Length - 3, 3) != "txt");
        }

        private void btnImportar_Click(object sender, EventArgs e) {
            // Verifica se o campo que contem o caminho do arquivo esta preenchido.
            if (txtCaminhoArquivoImportado.Text == "") {
                MessageBox.Show(null, "Favor informar o caminho do arquivo.", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Atribuir o arquivo a uma variavel do sistema.
            string caminhoArquivo = txtCaminhoArquivoImportado.Text;
            string[] linhasArquivo = File.ReadAllLines(caminhoArquivo); // Vetor que pega cada linha do arquivo.
            int qtdeLinhasArquivo = File.ReadAllLines(caminhoArquivo).Count();

            /*======================================= Começar a Importação =======================================*/

            // Variaveis com os dados a serem importados.
            int id;
            string plano, tipo_contrato, cartao_benef, dt_nascimento, acomodacao, dt_vigencia,
                    dt_validade, nome_benef, rede, tipo_prestador, atendimento, regulamentacao,
                    abrangencia, via, cobertura, contratante, segmentacao, trilha1, trilha2,
                    trilha3, cod_ans_operadora, cod_ans_produto, cod_ans_op_cont, desc_abrangencia, carencia1,
                    carencia2, carencia3, carencia4, carencia5, carencia6, carencia7, carencia8, sem_carencia,
                    nome_administradora, cnSaude;
            string carencia11, carencia21, carencia31, carencia41, carencia51, carencia61, carencia71, carencia81;

            // inserir no banco
            try {
                msConnection.Open();
                OleDbCommand msComando = new OleDbCommand();
                msComando.Connection = msConnection;

                // Laço para pegar os valores e atribuir as variaveis.
                for (int i = 0; i < qtdeLinhasArquivo; i++) {
                    // ID
                    id = i + 1;

                    /* =================   Frente  ================= */
                    /* ================= Cabeçalho ================= */
                    /* =================  Linha 1  ================= */

                    // Plano
                    plano = linhasArquivo[i].Substring(0, 60).ToString();

                    /* ================= Linha 2 ================= */
                    // Tipo Contrato
                    tipo_contrato = linhasArquivo[i].Substring(60, 30).ToString().ToUpper();

                    /* ================= Corpo Cartao ================= */
                    /* =================    Linha 1   ================= */

                    // Cartão Beneficiario
                    cartao_benef = linhasArquivo[i].Substring(90, 20).ToString();

                    /* =================    Linha 2   ================= */

                    // Data Nascimento
                    dt_nascimento = linhasArquivo[i].Substring(110, 14).ToString();

                    // Acomodação
                    acomodacao = linhasArquivo[i].Substring(124, 17).ToString().ToUpper();

                    // Vigencia
                    dt_vigencia = linhasArquivo[i].Substring(141, 14).ToString();

                    // Validade Cartao
                    dt_validade = linhasArquivo[i].Substring(155, 10).ToString();

                    /* =================    Linha 3   ================= */

                    // Nome Beneficiario
                    nome_benef = linhasArquivo[i].Substring(165, 35).ToString();

                    // Rede Preferenciada
                    rede = linhasArquivo[i].Substring(200, 4).ToString();

                    // Tipo Prestador
                    tipo_prestador = linhasArquivo[i].Substring(205, 8).ToString();

                    /* =================    Linha 4   ================= */

                    // Local Atendimento
                    atendimento = linhasArquivo[i].Substring(213, 9).ToString();

                    // Tipo Regulamentação 
                    regulamentacao = linhasArquivo[i].Substring(222, 23).ToString().ToUpper();

                    // Abrangencia
                    abrangencia = linhasArquivo[i].Substring(245, 21).ToString().ToUpper();

                    // Via
                    via = linhasArquivo[i].Substring(266, 2).ToString();

                    /* =================    Linha 5   ================= */

                    // CPT
                    cobertura = linhasArquivo[i].Substring(268, 31).ToString();

                    // Empresa
                    contratante = linhasArquivo[i].Substring(299, 18).ToString();

                    /* =================    Linha 6   ================= */

                    // Segmentação
                    segmentacao = linhasArquivo[i].Substring(317, 56).ToString();

                    /* =================   Verso  ================= */
                    /* ================= Trilha 1 ================= */
                    trilha1 = linhasArquivo[i].Substring(373, 28).ToString();

                    /* ================= Trilha 2 ================= */
                    trilha2 = linhasArquivo[i].Substring(401, 40).ToString();

                    /* ================= Trilha 3 ================= */
                    trilha3 = linhasArquivo[i].Substring(441, 104).ToString();

                    /* =================   Informações  ================= */
                    // Codigo da operadora na ANS
                    cod_ans_operadora = linhasArquivo[i].Substring(545, 6).ToString();

                    // Codigo do produto na ANS
                    cod_ans_produto = linhasArquivo[i].Substring(551, 30).ToString();

                    // NUMERO REGISTRO ANS NA OPERADORA CONTRATADA
                    cod_ans_op_cont = linhasArquivo[i].Substring(581, 26).ToString();

                    // Descrição Abrangencia
                    desc_abrangencia = linhasArquivo[i].Substring(607, 180).ToString();

                    // Mensagens Carencia
                    carencia1 = linhasArquivo[i].Substring(787, 60).ToString();
                    carencia11 = linhasArquivo[i].Substring(847, 8).ToString();
                    carencia2 = linhasArquivo[i].Substring(855, 60).ToString();
                    carencia21 = linhasArquivo[i].Substring(915, 8).ToString();
                    carencia3 = linhasArquivo[i].Substring(923, 60).ToString();
                    carencia31 = linhasArquivo[i].Substring(983, 8).ToString();
                    carencia4 = linhasArquivo[i].Substring(991, 60).ToString();
                    carencia41 = linhasArquivo[i].Substring(1051, 8).ToString();
                    carencia5 = linhasArquivo[i].Substring(1059, 60).ToString();
                    carencia51 = linhasArquivo[i].Substring(1119, 8).ToString();
                    carencia6 = linhasArquivo[i].Substring(1127, 60).ToString();
                    carencia61 = linhasArquivo[i].Substring(1187, 8).ToString();
                    carencia7 = linhasArquivo[i].Substring(1195, 60).ToString();
                    carencia71 = linhasArquivo[i].Substring(1255, 8).ToString();
                    carencia8 = linhasArquivo[i].Substring(1263, 60).ToString();
                    carencia81 = linhasArquivo[i].Substring(1323, 8).ToString();
                    
                    
                    // Sem Carencia
                    sem_carencia = linhasArquivo[i].Substring(1331, 23).ToString();

                    // Nome administradora
                    nome_administradora = linhasArquivo[i].Substring(1354, 25).ToString();

                    // Cartão Nacional Saúde
                    cnSaude = linhasArquivo[i].Substring(1379, 15).ToString();


                    // Insrir os valores
                    msComando.CommandText = "insert into Layout_PTU values(" + id + " , '" + plano + "', '" + tipo_contrato + "', '" +
                                            cartao_benef + "', '" + dt_nascimento + "', '" + acomodacao + "', '" + dt_vigencia + "', '" + dt_validade + "', '" + nome_benef
                                            + "', '" + rede + "', '" + tipo_prestador + "', '" + atendimento + "', '" + regulamentacao + "', '" + abrangencia + "', '" + via
                                            + "', '" + cobertura + "', '" + contratante + "', '" + segmentacao + "', '" + trilha1 + "', '" + trilha2 + "', '" + trilha3
                                            + "', '" + cod_ans_operadora + "', '" + cod_ans_produto + "', '" + cod_ans_op_cont + "', '" + desc_abrangencia + "', '" + carencia1
                                            + "', '" + carencia11 + "', '" + carencia2 + "', '" + carencia21 + "', '" + carencia3 + "', '" + carencia31
                                            + "', '" + carencia4 + "', '" + carencia41 + "', '" + carencia5 + "', '" + carencia51 + "', '" + carencia6 + "', '" +carencia61
                                            + "', '" + carencia7 + "', '" + carencia71 + "', '" + carencia8 + "', '" + carencia81 + "', '" + sem_carencia + "', '" + nome_administradora + "', '"+ cnSaude + "');";

                    msComando.ExecuteNonQuery();
                }
                msConnection.Close();
            }
            catch (Exception ex) {
                MessageBox.Show("Erro na importação, mesagem do banco: " + ex.ToString());
            } finally {
                MessageBox.Show(null, "Arquivo importado com sucesso !", "Informação", MessageBoxButtons.OK,MessageBoxIcon.Information);
            }

            //Verificar se existe registros
            verificaRegistros();
        }

        private void btnApagar_Click(object sender, EventArgs e) {

            try {
                OleDbCommand msComand = new OleDbCommand();
                msConnection.Open();
                msComand.Connection = msConnection;
                msComand.CommandText = "delete from Layout_PTU";

                msComand.ExecuteNonQuery();

                msConnection.Close();
                MessageBox.Show("Base de dados atual apagada.");

            } catch(Exception ex) {
                MessageBox.Show("Erro: " + ex.ToString());
            }

            //Verificar se existe registros
            verificaRegistros();
        }
    }
}