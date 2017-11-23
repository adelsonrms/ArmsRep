using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace UI.WF
{
    public partial class Form1 : Form
    {

        private BLL.ClienteBLL clientBLL = new BLL.ClienteBLL();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CarregaListaClientes();

        }

        private void CarregaListaClientes()
        {
            try
            {
                gvClientes.DataSource = clientBLL.ExibirTodos();
                lblmsg.Text = string.Format("A lista foi obtida com sucesso \n Qtd Clientes Cadastrados : {0}", gvClientes.Rows.Count);
                lblDetalhes.Visible = false;
            }
            catch (Exception ex)
            {
                lblDetalhes.Visible = true;
                lblmsg.Text = string.Format("Ocorreu um erro ao carregar a lista de Lientes !");
                lblDetalhes.Text = ex.Message;
            }
        }

        private void ConsultarClientePorID(int id)
        {
            try
            {
                var cliente = clientBLL.GetClienteByID(id);
                txtNome.Text = cliente.Nome;
                txtEndereco.Text = cliente.Endereco;
                txtEmail.Text = cliente.Email;
                txtTelefone.Text = cliente.Telefone;
                txtObservacao.Text = cliente.Observacoes;
            }
            catch (Exception ex)
            {
                lblDetalhes.Visible = true;
                lblmsg.Text = string.Format("Ocorreu um erro ao obter informações sobreo o cliente procurado !");
                lblDetalhes.Text = ex.Message;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ConsultarClientePorID (int.Parse(txtID.Text));
        }
    }
}
