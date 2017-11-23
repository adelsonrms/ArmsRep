using System;

namespace UI
{
    public partial class Default : System.Web.UI.Page
    {

        private BLL.ClienteBLL clientBLL = new BLL.ClienteBLL();
        
        protected void Page_Load(object sender, EventArgs e)
        {
            CarregaListaClientes();
        }

        protected void btnLocalizar_Click(object sender, EventArgs e)
        {
            ConsultarClientePorID(id: int.Parse(txtID.Text));
        }

        private void CarregaListaClientes()
        {
            try
            {
                gvClientes.DataSource = clientBLL.ExibirTodos();
                gvClientes.DataBind();
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
    }
}