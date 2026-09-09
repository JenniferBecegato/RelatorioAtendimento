using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace RelatorioAtendimento
{
    public partial class Form1 : Form
    {
        // =========================================================
        // CORES
        // =========================================================
        private readonly Color CorSidebar = Color.FromArgb(10, 67, 56);
        private readonly Color CorSidebarHover = Color.FromArgb(18, 93, 76);
        private readonly Color CorVerde = Color.FromArgb(23, 145, 92);
        private readonly Color CorVerdeEscuro = Color.FromArgb(9, 102, 73);
        private readonly Color CorFundo = Color.FromArgb(245, 247, 250);
        private readonly Color CorTexto = Color.FromArgb(28, 42, 57);
        private readonly Color CorCinza = Color.FromArgb(105, 115, 125);
        private readonly Color CorBorda = Color.FromArgb(220, 225, 230);
        private readonly Color CorVermelho = Color.FromArgb(210, 54, 54);
        private readonly Color CorAzul = Color.FromArgb(41, 112, 211);
        private readonly Color CorAmarelo = Color.FromArgb(215, 158, 30);

        private static readonly CultureInfo PtBr = CultureInfo.GetCultureInfo("pt-BR");

        // =========================================================
        // ARQUIVOS
        // =========================================================
        private string caminhoMesRetrasado = string.Empty;
        private string caminhoMesPassado = string.Empty;

        // Relatórios mantidos em memória depois de clicar em Comparar.
        private RelatorioMensal? relatorioAnteriorAtual;
        private RelatorioMensal? relatorioAtualAtual;

        // Navegação
        private Panel painelPaginas = null!;
        private AppPage paginaAtual = AppPage.Comparativo;
        private readonly Dictionary<AppPage, Button> botoesMenu = new Dictionary<AppPage, Button>();

        private enum AppPage
        {
            Comparativo,
            Resumo,
            Atendentes,
            Motivos,
            Marketing,
            Qualidade,
            Exportar
        }

        // =========================================================
        // CONTROLES
        // =========================================================
        private TextBox txtMesRetrasado = null!;
        private TextBox txtMesPassado = null!;
        private Label lblPeriodo1 = null!;
        private Label lblPeriodo2 = null!;
        private Label lblStatus = null!;

        private DataGridView dgvAtendentes = null!;
        private ComparisonBarChart graficoMotivos = null!;

        private Label lblResumoTexto = null!;
        private Label lblAlertasTexto = null!;
        private Label lblOportunidadesTexto = null!;
        private Label lblAcoesTexto = null!;
        private Label lblDestaquesTexto = null!;
        private Label lblIntegridadeTexto = null!;

        private KpiView kpiConversas = null!;
        private KpiView kpiFinalizados = null!;
        private KpiView kpiTaxaFinalizacao = null!;
        private KpiView kpiReagendamentos = null!;
        private KpiView kpiNovosContatos = null!;
        private KpiView kpiNotaMedia = null!;

        public Form1()
        {
            InitializeComponent();
            CriarTela();
            LimparDados(false);
        }

        // =========================================================
        // TELA
        // =========================================================
        private void CriarTela()
        {
            SuspendLayout();
            Controls.Clear();

            Text = "Relatório de Atendimento - Comparativo Mensal";
            StartPosition = FormStartPosition.CenterScreen;
            WindowState = FormWindowState.Maximized;
            MinimumSize = new Size(1200, 720);
            BackColor = CorFundo;
            Font = new Font("Segoe UI", 9F);
            DoubleBuffered = true;

            TableLayoutPanel estrutura = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = CorFundo,
                Margin = Padding.Empty,
                Padding = Padding.Empty
            };

            estrutura.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 205));
            estrutura.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            estrutura.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            estrutura.Controls.Add(CriarSidebar(), 0, 0);

            painelPaginas = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = CorFundo
            };
            estrutura.Controls.Add(painelPaginas, 1, 0);

            Controls.Add(estrutura);

            MostrarPagina(AppPage.Comparativo);
            ResumeLayout(true);
        }

        private Control CriarSidebar()
        {
            Panel sidebar = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = CorSidebar
            };

            Panel cabecalho = new Panel
            {
                Dock = DockStyle.Top,
                Height = 105,
                BackColor = CorSidebar
            };

            cabecalho.Controls.Add(new Label
            {
                Text = "RELATÓRIO\nDE ATENDIMENTO",
                ForeColor = Color.White,
                Font = new Font("Segoe UI Semibold", 13F, FontStyle.Bold),
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter
            });

            FlowLayoutPanel menu = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 430,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                BackColor = CorSidebar,
                Padding = new Padding(0, 8, 0, 0)
            };

            menu.Controls.Add(CriarBotaoMenu("▥   Comparativo", AppPage.Comparativo));
            menu.Controls.Add(CriarBotaoMenu("▣   Resumo Executivo", AppPage.Resumo));
            menu.Controls.Add(CriarBotaoMenu("●   Atendentes", AppPage.Atendentes));
            menu.Controls.Add(CriarBotaoMenu("◆   Motivos", AppPage.Motivos));
            menu.Controls.Add(CriarBotaoMenu("▲   Marketing", AppPage.Marketing));
            menu.Controls.Add(CriarBotaoMenu("✓   Qualidade dos Dados", AppPage.Qualidade));
            menu.Controls.Add(CriarBotaoMenu("⇩   Exportar", AppPage.Exportar));

            Panel rodape = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 150,
                BackColor = CorSidebar
            };

            rodape.Controls.Add(new Label
            {
                Text = "Transformando dados\nem informações para\nmelhores decisões.",
                ForeColor = Color.FromArgb(200, 225, 218),
                Font = new Font("Segoe UI", 10F),
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter
            });

            sidebar.Controls.Add(rodape);
            sidebar.Controls.Add(menu);
            sidebar.Controls.Add(cabecalho);

            return sidebar;
        }

        private Button CriarBotaoMenu(string texto, AppPage pagina)
        {
            Button btn = new Button
            {
                Width = 205,
                Height = 49,
                Text = texto,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(22, 0, 0, 0),
                ForeColor = Color.White,
                BackColor = CorSidebar,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F),
                Cursor = Cursors.Hand,
                Tag = pagina
            };

            btn.FlatAppearance.BorderSize = 0;

            btn.Click += (s, e) => MostrarPagina(pagina);

            btn.MouseEnter += (s, e) =>
            {
                if (paginaAtual != pagina)
                    btn.BackColor = CorSidebarHover;
            };

            btn.MouseLeave += (s, e) =>
            {
                if (paginaAtual != pagina)
                    btn.BackColor = CorSidebar;
            };

            botoesMenu[pagina] = btn;
            return btn;
        }

        private Control CriarConteudoPrincipal()
        {
            TableLayoutPanel principal = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = CorFundo,
                ColumnCount = 1,
                RowCount = 6,
                Padding = new Padding(14, 10, 14, 10)
            };

            principal.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            principal.RowStyles.Add(new RowStyle(SizeType.Absolute, 72));
            principal.RowStyles.Add(new RowStyle(SizeType.Absolute, 105));
            principal.RowStyles.Add(new RowStyle(SizeType.Absolute, 124));
            principal.RowStyles.Add(new RowStyle(SizeType.Absolute, 215));
            principal.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            principal.RowStyles.Add(new RowStyle(SizeType.Absolute, 55));

            principal.Controls.Add(CriarCabecalho(), 0, 0);
            principal.Controls.Add(CriarAreaImportacao(), 0, 1);
            principal.Controls.Add(CriarKpis(), 0, 2);
            principal.Controls.Add(CriarAnalise(), 0, 3);
            principal.Controls.Add(CriarParteInferior(), 0, 4);
            principal.Controls.Add(CriarRodape(), 0, 5);

            return principal;
        }

        private Control CriarCabecalho()
        {
            Panel panel = new Panel { Dock = DockStyle.Fill };

            Label titulo = new Label
            {
                Text = "Relatório de Atendimento - Comparativo Mensal",
                Font = new Font("Segoe UI Semibold", 20F, FontStyle.Bold),
                ForeColor = CorTexto,
                AutoSize = true,
                Location = new Point(8, 5)
            };

            Label subtitulo = new Label
            {
                Text = "Compare dois arquivos do Excel e descubra os principais insights para a gestão do atendimento.",
                Font = new Font("Segoe UI", 10.5F),
                ForeColor = CorCinza,
                AutoSize = true,
                Location = new Point(11, 43)
            };

            Label data = new Label
            {
                Text = "Versão 1.0.0\n" + DateTime.Now.ToString("dd/MM/yyyy  HH:mm"),
                Font = new Font("Segoe UI", 8.5F),
                ForeColor = CorCinza,
                AutoSize = false,
                Width = 190,
                Height = 46,
                TextAlign = ContentAlignment.TopRight,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };

            panel.Controls.Add(titulo);
            panel.Controls.Add(subtitulo);
            panel.Controls.Add(data);

            panel.Resize += (s, e) =>
            {
                data.Location = new Point(Math.Max(0, panel.ClientSize.Width - data.Width - 10), 7);
            };

            return panel;
        }

        // =========================================================
        // IMPORTAÇÃO
        // =========================================================
        private Control CriarAreaImportacao()
        {
            TableLayoutPanel area = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                Padding = new Padding(0, 4, 0, 6)
            };

            area.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42F));
            area.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42F));
            area.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16F));

            area.Controls.Add(
                CriarCardImportacao("1. Mês retrasado", out txtMesRetrasado, out lblPeriodo1, EscolherMesRetrasado),
                0, 0);

            area.Controls.Add(
                CriarCardImportacao("2. Mês passado", out txtMesPassado, out lblPeriodo2, EscolherMesPassado),
                1, 0);

            Button btnComparar = new Button
            {
                Text = "Comparar",
                Dock = DockStyle.Fill,
                Margin = new Padding(8, 10, 4, 10),
                BackColor = CorVerde,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI Semibold", 12F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };

            btnComparar.FlatAppearance.BorderSize = 0;
            btnComparar.Click += BtnComparar_Click;

            area.Controls.Add(btnComparar, 2, 0);
            return area;
        }

        private Control CriarCardImportacao(
            string titulo,
            out TextBox txtArquivo,
            out Label lblPeriodo,
            EventHandler evento)
        {
            RoundedPanel card = new RoundedPanel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(4),
                BackColor = Color.White,
                BorderColor = CorBorda,
                Radius = 10
            };

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(12, 7, 12, 7)
            };

            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            Label lblTitulo = new Label
            {
                Text = titulo,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold),
                ForeColor = CorTexto,
                TextAlign = ContentAlignment.MiddleLeft
            };

            TableLayoutPanel linhaArquivo = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1
            };

            linhaArquivo.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            linhaArquivo.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 105));

            txtArquivo = new TextBox
            {
                Text = "Nenhum arquivo selecionado",
                ReadOnly = true,
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9F)
            };

            Button btnArquivo = new Button
            {
                Text = "Selecionar",
                Dock = DockStyle.Fill,
                Margin = new Padding(6, 0, 0, 0),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(239, 247, 243),
                ForeColor = CorVerdeEscuro,
                Cursor = Cursors.Hand
            };

            btnArquivo.FlatAppearance.BorderColor = Color.FromArgb(190, 220, 205);
            btnArquivo.Click += evento;

            linhaArquivo.Controls.Add(txtArquivo, 0, 0);
            linhaArquivo.Controls.Add(btnArquivo, 1, 0);

            lblPeriodo = new Label
            {
                Text = string.Empty,
                Dock = DockStyle.Fill,
                ForeColor = CorCinza,
                Font = new Font("Segoe UI", 8.3F),
                TextAlign = ContentAlignment.MiddleLeft
            };

            layout.Controls.Add(lblTitulo, 0, 0);
            layout.Controls.Add(linhaArquivo, 0, 1);
            layout.Controls.Add(lblPeriodo, 0, 2);
            card.Controls.Add(layout);

            return card;
        }

        // =========================================================
        // KPI
        // =========================================================
        private Control CriarKpis()
        {
            TableLayoutPanel painel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 6,
                RowCount = 1,
                Padding = new Padding(0, 2, 0, 6)
            };

            for (int i = 0; i < 6; i++)
                painel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16.666F));

            painel.Controls.Add(CriarKpi("Conversas recebidas", out kpiConversas), 0, 0);
            painel.Controls.Add(CriarKpi("Atendimentos finalizados", out kpiFinalizados), 1, 0);
            painel.Controls.Add(CriarKpi("Taxa de finalização", out kpiTaxaFinalizacao), 2, 0);
            painel.Controls.Add(CriarKpi("Reagendamentos", out kpiReagendamentos), 3, 0);
            painel.Controls.Add(CriarKpi("Novos contatos", out kpiNovosContatos), 4, 0);
            painel.Controls.Add(CriarKpi("Nota média", out kpiNotaMedia), 5, 0);

            return painel;
        }

        private Control CriarKpi(string titulo, out KpiView kpi)
        {
            RoundedPanel card = new RoundedPanel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(4),
                Radius = 10,
                BorderColor = CorBorda,
                BackColor = Color.White
            };

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 4,
                ColumnCount = 1,
                Padding = new Padding(13, 9, 10, 6)
            };

            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            Label lblTitulo = new Label
            {
                Text = titulo,
                Dock = DockStyle.Fill,
                ForeColor = CorTexto,
                Font = new Font("Segoe UI", 9F),
                TextAlign = ContentAlignment.MiddleLeft
            };

            Label lblValor = new Label
            {
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(17, 33, 48),
                Font = new Font("Segoe UI Semibold", 16F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };

            Label lblVariacao = new Label
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };

            Label lblAnterior = new Label
            {
                Dock = DockStyle.Fill,
                ForeColor = CorCinza,
                Font = new Font("Segoe UI", 8F),
                TextAlign = ContentAlignment.TopLeft
            };

            layout.Controls.Add(lblTitulo, 0, 0);
            layout.Controls.Add(lblValor, 0, 1);
            layout.Controls.Add(lblVariacao, 0, 2);
            layout.Controls.Add(lblAnterior, 0, 3);
            card.Controls.Add(layout);

            kpi = new KpiView(lblValor, lblVariacao, lblAnterior);
            return card;
        }

        // =========================================================
        // ANÁLISES
        // =========================================================
        private Control CriarAnalise()
        {
            TableLayoutPanel area = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 1,
                Padding = new Padding(0, 3, 0, 6)
            };

            area.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 24F));
            area.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 24F));
            area.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 24F));
            area.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 28F));

            area.Controls.Add(
                CriarCardAnalise("▣  Resumo Executivo", Color.FromArgb(236, 249, 242), CorVerdeEscuro, out lblResumoTexto),
                0, 0);

            area.Controls.Add(
                CriarCardAnalise("⚠  Principais Alertas", Color.FromArgb(253, 239, 239), CorVermelho, out lblAlertasTexto),
                1, 0);

            area.Controls.Add(
                CriarCardAnalise("●  Oportunidades", Color.FromArgb(237, 250, 241), CorVerdeEscuro, out lblOportunidadesTexto),
                2, 0);

            area.Controls.Add(
                CriarCardAnalise("◎  Ações Recomendadas", Color.FromArgb(238, 245, 253), CorAzul, out lblAcoesTexto),
                3, 0);

            return area;
        }

        private Control CriarCardAnalise(string titulo, Color fundo, Color corTitulo, out Label lblTexto)
        {
            RoundedPanel card = new RoundedPanel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(4),
                Radius = 10,
                BorderColor = CorBorda,
                BackColor = fundo
            };

            Label lblTitulo = new Label
            {
                Dock = DockStyle.Top,
                Height = 37,
                Text = titulo,
                Font = new Font("Segoe UI Semibold", 11F, FontStyle.Bold),
                ForeColor = corTitulo,
                Padding = new Padding(12, 10, 0, 0)
            };

            lblTexto = new Label
            {
                Dock = DockStyle.Fill,
                Text = string.Empty,
                ForeColor = CorTexto,
                Font = new Font("Segoe UI", 8.8F),
                Padding = new Padding(13, 4, 10, 7)
            };

            card.Controls.Add(lblTexto);
            card.Controls.Add(lblTitulo);
            return card;
        }

        // =========================================================
        // PARTE INFERIOR
        // =========================================================
        private Control CriarParteInferior()
        {
            TableLayoutPanel area = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1
            };

            area.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34F));
            area.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 44F));
            area.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 22F));

            area.Controls.Add(CriarGrafico(), 0, 0);
            area.Controls.Add(CriarTabelaAtendentes(), 1, 0);
            area.Controls.Add(CriarPainelDireito(), 2, 0);
            return area;
        }

        private Control CriarGrafico()
        {
            RoundedPanel card = new RoundedPanel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(4),
                Radius = 10,
                BorderColor = CorBorda,
                BackColor = Color.White
            };

            Label titulo = new Label
            {
                Text = "▥  Motivos de Atendimento",
                Dock = DockStyle.Top,
                Height = 35,
                Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold),
                ForeColor = CorTexto,
                Padding = new Padding(12, 10, 0, 0)
            };

            graficoMotivos = new ComparisonBarChart
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White
            };

            card.Controls.Add(graficoMotivos);
            card.Controls.Add(titulo);
            return card;
        }

        private Control CriarTabelaAtendentes()
        {
            RoundedPanel card = new RoundedPanel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(4),
                Radius = 10,
                BorderColor = CorBorda,
                BackColor = Color.White
            };

            Label titulo = new Label
            {
                Text = "●  Desempenho por Atendente",
                Dock = DockStyle.Top,
                Height = 35,
                Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold),
                ForeColor = CorTexto,
                Padding = new Padding(12, 10, 0, 0)
            };

            dgvAtendentes = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ColumnHeadersHeight = 32,
                EnableHeadersVisualStyles = false
            };

            dgvAtendentes.RowTemplate.Height = 28;
            dgvAtendentes.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(239, 243, 247);
            dgvAtendentes.ColumnHeadersDefaultCellStyle.ForeColor = CorTexto;
            dgvAtendentes.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI Semibold", 8.2F, FontStyle.Bold);
            dgvAtendentes.DefaultCellStyle.Font = new Font("Segoe UI", 8.2F);
            dgvAtendentes.DefaultCellStyle.ForeColor = CorTexto;
            dgvAtendentes.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 240, 234);
            dgvAtendentes.DefaultCellStyle.SelectionForeColor = CorTexto;
            dgvAtendentes.GridColor = Color.FromArgb(230, 233, 237);

            dgvAtendentes.Columns.Add("Nome", "Atendente");
            dgvAtendentes.Columns.Add("Anterior", "Mês retrasado");
            dgvAtendentes.Columns.Add("Atual", "Mês passado");
            dgvAtendentes.Columns.Add("Variacao", "Variação");
            dgvAtendentes.Columns.Add("Nota", "Nota");
            dgvAtendentes.Columns.Add("TMR", "TMR");
            dgvAtendentes.Columns.Add("Diagnostico", "Diagnóstico");
            dgvAtendentes.CellFormatting += DgvAtendentes_CellFormatting;

            card.Controls.Add(dgvAtendentes);
            card.Controls.Add(titulo);
            return card;
        }

        private void DgvAtendentes_CellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex < 0 || dgvAtendentes.Columns[e.ColumnIndex].Name != "Variacao")
                return;

            string valor = Convert.ToString(e.Value) ?? string.Empty;

            if (valor.StartsWith("▲"))
            {
                e.CellStyle.ForeColor = CorVerde;
                e.CellStyle.Font = new Font("Segoe UI Semibold", 8.2F, FontStyle.Bold);
            }
            else if (valor.StartsWith("▼"))
            {
                e.CellStyle.ForeColor = CorVermelho;
                e.CellStyle.Font = new Font("Segoe UI Semibold", 8.2F, FontStyle.Bold);
            }
            else if (valor.Contains("Novo", StringComparison.OrdinalIgnoreCase))
            {
                e.CellStyle.ForeColor = CorAzul;
                e.CellStyle.Font = new Font("Segoe UI Semibold", 8.2F, FontStyle.Bold);
            }
        }

        private Control CriarPainelDireito()
        {
            TableLayoutPanel area = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1
            };

            area.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            area.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            area.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));

            area.Controls.Add(
                CriarCardAnalise("★  Destaques da Equipe", Color.White, CorVerdeEscuro, out lblDestaquesTexto),
                0, 0);

            area.Controls.Add(
                CriarCardAnalise("⚠  Integridade dos Dados", Color.White, CorAmarelo, out lblIntegridadeTexto),
                0, 1);

            return area;
        }

        // =========================================================
        // RODAPÉ / LIMPAR
        // =========================================================
        private Control CriarRodape()
        {
            Panel rodape = new Panel { Dock = DockStyle.Fill };

            lblStatus = new Label
            {
                AutoSize = true,
                ForeColor = CorVerdeEscuro,
                Location = new Point(5, 18)
            };

            FlowLayoutPanel botoes = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                Width = 530,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(0, 7, 0, 0)
            };

            Button btnFechar = CriarBotaoRodape("Fechar", Color.White, CorTexto);
            Button btnExcel = CriarBotaoRodape("Exportar Excel", Color.White, CorVerdeEscuro);
            Button btnPdf = CriarBotaoRodape("Exportar PDF", CorVerde, Color.White);
            Button btnLimpar = CriarBotaoRodape("Limpar dados", Color.White, CorVermelho);

            btnFechar.Click += (s, e) => Close();
            btnLimpar.Click += (s, e) => LimparDados();

            btnExcel.Click += (s, e) => MessageBox.Show(
                "A exportação será ligada depois que finalizarmos a leitura e o relatório.",
                "Exportar Excel",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            btnPdf.Click += (s, e) => MessageBox.Show(
                "A exportação será ligada depois que finalizarmos a leitura e o relatório.",
                "Exportar PDF",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            botoes.Controls.Add(btnFechar);
            botoes.Controls.Add(btnExcel);
            botoes.Controls.Add(btnPdf);
            botoes.Controls.Add(btnLimpar);

            rodape.Controls.Add(lblStatus);
            rodape.Controls.Add(botoes);
            return rodape;
        }

        private Button CriarBotaoRodape(string textoBotao, Color fundo, Color corTexto)
        {
            Button btn = new Button
            {
                Text = textoBotao,
                Width = 120,
                Height = 38,
                Margin = new Padding(5),
                BackColor = fundo,
                ForeColor = corTexto,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI Semibold", 9F),
                Cursor = Cursors.Hand
            };

            btn.FlatAppearance.BorderColor = CorBorda;
            return btn;
        }


        // =========================================================
        // NAVEGAÇÃO ENTRE AS ABAS DO MENU
        // =========================================================
        private void MostrarPagina(AppPage pagina)
        {
            paginaAtual = pagina;
            AtualizarMenuAtivo();

            if (painelPaginas == null)
                return;

            painelPaginas.SuspendLayout();
            painelPaginas.Controls.Clear();

            Control conteudo = pagina switch
            {
                AppPage.Comparativo => CriarConteudoPrincipal(),
                AppPage.Resumo => CriarPaginaResumoDetalhado(),
                AppPage.Atendentes => CriarPaginaAtendentesDetalhado(),
                AppPage.Motivos => CriarPaginaMotivosDetalhado(),
                AppPage.Marketing => CriarPaginaMarketingDetalhado(),
                AppPage.Qualidade => CriarPaginaQualidadeDetalhado(),
                AppPage.Exportar => CriarPaginaExportarDetalhado(),
                _ => CriarConteudoPrincipal()
            };

            conteudo.Dock = DockStyle.Fill;
            painelPaginas.Controls.Add(conteudo);
            painelPaginas.ResumeLayout(true);

            if (pagina == AppPage.Comparativo)
                AtualizarPaginaComparativo();
        }

        private void AtualizarMenuAtivo()
        {
            foreach (KeyValuePair<AppPage, Button> item in botoesMenu)
            {
                bool ativo = item.Key == paginaAtual;

                item.Value.BackColor = ativo ? CorSidebarHover : CorSidebar;
                item.Value.Font = new Font(
                    "Segoe UI",
                    10F,
                    ativo ? FontStyle.Bold : FontStyle.Regular);
            }
        }

        private void AtualizarPaginaComparativo()
        {
            if (txtMesRetrasado != null)
            {
                txtMesRetrasado.Text = string.IsNullOrWhiteSpace(caminhoMesRetrasado)
                    ? "Nenhum arquivo selecionado"
                    : Path.GetFileName(caminhoMesRetrasado);
            }

            if (txtMesPassado != null)
            {
                txtMesPassado.Text = string.IsNullOrWhiteSpace(caminhoMesPassado)
                    ? "Nenhum arquivo selecionado"
                    : Path.GetFileName(caminhoMesPassado);
            }

            if (relatorioAnteriorAtual != null && relatorioAtualAtual != null)
            {
                PreencherDashboard(relatorioAnteriorAtual, relatorioAtualAtual);

                lblPeriodo1.Text = relatorioAnteriorAtual.Periodo;
                lblPeriodo2.Text = relatorioAtualAtual.Periodo;

                lblStatus.Text =
                    $"●  Comparação concluída: {relatorioAnteriorAtual.Periodo} x {relatorioAtualAtual.Periodo}.";
            }
            else
            {
                LimparResultados();

                if (lblPeriodo1 != null)
                    lblPeriodo1.Text = string.Empty;

                if (lblPeriodo2 != null)
                    lblPeriodo2.Text = string.Empty;

                AtualizarStatusArquivos();
            }
        }

        // =========================================================
        // BASE DAS TELAS DETALHADAS
        // =========================================================
        private Panel CriarPaginaDetalheBase(
            string titulo,
            string subtitulo,
            int alturaConteudo,
            out TableLayoutPanel corpo)
        {
            Panel scroll = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = CorFundo
            };

            TableLayoutPanel corpoLocal = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Width = 1000,
                Height = alturaConteudo,
                ColumnCount = 1,
                RowCount = 1,
                Padding = new Padding(20, 12, 20, 15),
                BackColor = CorFundo
            };

            corpoLocal.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            corpoLocal.RowStyles.Add(new RowStyle(SizeType.Absolute, 78F));

            corpoLocal.Controls.Add(CriarCabecalhoDetalhe(titulo, subtitulo), 0, 0);

            scroll.Controls.Add(corpoLocal);

            // Não capturamos o parâmetro OUT dentro da lambda.
            // Isso evita o erro CS1628.
            scroll.Resize += (s, e) =>
            {
                corpoLocal.Width = Math.Max(950, scroll.ClientSize.Width - 28);
            };

            corpo = corpoLocal;

            return scroll;
        }

        private Control CriarCabecalhoDetalhe(string titulo, string subtitulo)
        {
            Panel panel = new Panel
            {
                Dock = DockStyle.Fill
            };

            Label lblTitulo = new Label
            {
                AutoSize = true,
                Text = titulo,
                Font = new Font("Segoe UI Semibold", 19F, FontStyle.Bold),
                ForeColor = CorTexto,
                Location = new Point(8, 4)
            };

            Label lblSubtitulo = new Label
            {
                AutoSize = true,
                Text = subtitulo,
                Font = new Font("Segoe UI", 9.5F),
                ForeColor = CorCinza,
                Location = new Point(10, 43)
            };

            Label periodo = new Label
            {
                Width = 240,
                Height = 42,
                TextAlign = ContentAlignment.TopRight,
                ForeColor = CorCinza,
                Font = new Font("Segoe UI", 8F),
                Text = relatorioAnteriorAtual != null && relatorioAtualAtual != null
                    ? $"{relatorioAnteriorAtual.Periodo}  ×  {relatorioAtualAtual.Periodo}"
                    : "Aguardando comparação"
            };

            panel.Controls.Add(lblTitulo);
            panel.Controls.Add(lblSubtitulo);
            panel.Controls.Add(periodo);

            panel.Resize += (s, e) =>
            {
                periodo.Location = new Point(
                    Math.Max(0, panel.ClientSize.Width - periodo.Width - 8),
                    8);
            };

            return panel;
        }

        private Control CriarAvisoSemComparacao()
        {
            RoundedPanel card = new RoundedPanel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(4),
                BackColor = Color.White,
                BorderColor = CorBorda,
                Radius = 10
            };

            card.Controls.Add(new Label
            {
                Dock = DockStyle.Fill,
                Text =
                    "Nenhuma comparação disponível.\n\n" +
                    "Clique em Comparativo, selecione os dois arquivos Excel e depois clique em Comparar.",
                ForeColor = CorCinza,
                Font = new Font("Segoe UI", 10F),
                TextAlign = ContentAlignment.MiddleCenter
            });

            return card;
        }

        private RoundedPanel CriarCardDetalhe()
        {
            return new RoundedPanel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(4),
                BackColor = Color.White,
                BorderColor = CorBorda,
                Radius = 10
            };
        }

        private Control CriarCardTextoDetalhe(
            string titulo,
            string texto,
            Color corTitulo,
            Color? fundo = null)
        {
            RoundedPanel card = CriarCardDetalhe();

            if (fundo.HasValue)
                card.BackColor = fundo.Value;

            Label lblTitulo = new Label
            {
                Dock = DockStyle.Top,
                Height = 40,
                Text = titulo,
                Padding = new Padding(13, 11, 0, 0),
                ForeColor = corTitulo,
                Font = new Font("Segoe UI Semibold", 10.5F, FontStyle.Bold)
            };

            Label lblTexto = new Label
            {
                Dock = DockStyle.Fill,
                Text = texto,
                Padding = new Padding(13, 6, 13, 10),
                ForeColor = CorTexto,
                Font = new Font("Segoe UI", 9F),
                AutoEllipsis = true
            };

            card.Controls.Add(lblTexto);
            card.Controls.Add(lblTitulo);

            return card;
        }

        private DataGridView CriarGridDetalhe()
        {
            DataGridView grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                GridColor = Color.FromArgb(230, 233, 237),
                ColumnHeadersHeight = 34
            };

            grid.RowTemplate.Height = 29;

            grid.ColumnHeadersDefaultCellStyle.BackColor =
                Color.FromArgb(239, 243, 247);

            grid.ColumnHeadersDefaultCellStyle.ForeColor = CorTexto;

            grid.ColumnHeadersDefaultCellStyle.Font =
                new Font("Segoe UI Semibold", 8.3F, FontStyle.Bold);

            grid.DefaultCellStyle.ForeColor = CorTexto;
            grid.DefaultCellStyle.Font = new Font("Segoe UI", 8.3F);

            grid.DefaultCellStyle.SelectionBackColor =
                Color.FromArgb(225, 240, 234);

            grid.DefaultCellStyle.SelectionForeColor = CorTexto;

            return grid;
        }

        private Control CriarCardComGrid(string titulo, DataGridView grid)
        {
            RoundedPanel card = CriarCardDetalhe();

            Label lblTitulo = new Label
            {
                Dock = DockStyle.Top,
                Height = 40,
                Text = titulo,
                Padding = new Padding(13, 11, 0, 0),
                ForeColor = CorTexto,
                Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold)
            };

            card.Controls.Add(grid);
            card.Controls.Add(lblTitulo);
            return card;
        }

        // =========================================================
        // RESUMO EXECUTIVO DETALHADO
        // =========================================================
        private Control CriarPaginaResumoDetalhado()
        {
            Panel pagina = CriarPaginaDetalheBase(
                "Resumo Executivo",
                "Uma leitura mais completa do que mudou e do que merece atenção.",
                760,
                out TableLayoutPanel corpo);

            corpo.RowCount = 4;
            corpo.RowStyles.Add(new RowStyle(SizeType.Absolute, 125F));
            corpo.RowStyles.Add(new RowStyle(SizeType.Absolute, 255F));
            corpo.RowStyles.Add(new RowStyle(SizeType.Absolute, 285F));

            if (relatorioAnteriorAtual == null || relatorioAtualAtual == null)
            {
                corpo.Controls.Add(CriarAvisoSemComparacao(), 0, 1);
                return pagina;
            }

            RelatorioMensal anterior = relatorioAnteriorAtual;
            RelatorioMensal atual = relatorioAtualAtual;

            double taxaAnterior = CalcularTaxa(anterior.Finalizados, anterior.Conversas);
            double taxaAtual = CalcularTaxa(atual.Finalizados, atual.Conversas);

            TableLayoutPanel indicadores = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 1,
                Padding = new Padding(0, 2, 0, 4)
            };

            for (int i = 0; i < 4; i++)
                indicadores.ColumnStyles.Add(
                    new ColumnStyle(SizeType.Percent, 25F));

            indicadores.Controls.Add(
                CriarMiniIndicadorDetalhe(
                    "Marketing",
                    anterior.ContatosMarketing,
                    atual.ContatosMarketing,
                    false),
                0,
                0);

            indicadores.Controls.Add(
                CriarMiniIndicadorDetalhe(
                    "Mensagens template",
                    anterior.MensagensTemplate,
                    atual.MensagensTemplate,
                    true),
                1,
                0);

            indicadores.Controls.Add(
                CriarMiniIndicadorDetalhe(
                    "Encerrados por inatividade",
                    ValorMotivo(anterior, "Encerrado por inatividade do cliente"),
                    ValorMotivo(atual, "Encerrado por inatividade do cliente"),
                    true),
                2,
                0);

            indicadores.Controls.Add(
                CriarMiniIndicadorDetalhe(
                    "Orçamentos de fórmula",
                    ValorMotivo(anterior, "Orçamento de fórmula"),
                    ValorMotivo(atual, "Orçamento de fórmula"),
                    false),
                3,
                0);

            corpo.Controls.Add(indicadores, 0, 1);

            TableLayoutPanel analises = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1
            };

            analises.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40F));
            analises.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F));
            analises.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F));

            analises.Controls.Add(
                CriarCardTextoDetalhe(
                    "Resumo da comparação",
                    GerarResumoExecutivo(anterior, atual, taxaAnterior, taxaAtual),
                    CorVerdeEscuro),
                0,
                0);

            analises.Controls.Add(
                CriarCardTextoDetalhe(
                    "Principais alertas",
                    GerarAlertas(anterior, atual, taxaAnterior, taxaAtual),
                    CorVermelho,
                    Color.FromArgb(253, 239, 239)),
                1,
                0);

            analises.Controls.Add(
                CriarCardTextoDetalhe(
                    "Oportunidades",
                    GerarOportunidades(anterior, atual),
                    CorVerdeEscuro,
                    Color.FromArgb(237, 250, 241)),
                2,
                0);

            corpo.Controls.Add(analises, 0, 2);

            DataGridView grid = CriarGridDetalhe();

            grid.Columns.Add("Indicador", "Indicador");
            grid.Columns.Add("Anterior", anterior.Periodo);
            grid.Columns.Add("Atual", atual.Periodo);
            grid.Columns.Add("Variacao", "Variação");
            grid.Columns.Add("Leitura", "Leitura gerencial");

            DataGridViewColumn? colIndicador = grid.Columns["Indicador"];
            DataGridViewColumn? colLeituraResumo = grid.Columns["Leitura"];

            if (colIndicador != null)
                colIndicador.FillWeight = 130;

            if (colLeituraResumo != null)
                colLeituraResumo.FillWeight = 190;

            AdicionarIndicadorResumo(
                grid,
                "Conversas recebidas",
                anterior.Conversas,
                atual.Conversas,
                "Demanda total recebida.");

            AdicionarIndicadorResumo(
                grid,
                "Atendimentos finalizados",
                anterior.Finalizados,
                atual.Finalizados,
                "Capacidade de fechamento da equipe.");

            grid.Rows.Add(
                "Taxa de finalização",
                taxaAnterior.ToString("0.0", PtBr) + "%",
                taxaAtual.ToString("0.0", PtBr) + "%",
                FormatarDeltaPp(taxaAnterior, taxaAtual),
                "Compare a evolução da demanda com a capacidade de finalizar.");

            AdicionarIndicadorResumo(
                grid,
                "Novos contatos",
                anterior.NovosContatos,
                atual.NovosContatos,
                "Entrada de novos clientes/contatos.");

            AdicionarIndicadorResumo(
                grid,
                "Reagendamentos",
                anterior.Reagendamentos,
                atual.Reagendamentos,
                "Crescimento relevante deve ser investigado por motivo.");

            grid.Rows.Add(
                "Nota média",
                anterior.NotaMedia > 0 ? anterior.NotaMedia.ToString("0.00", PtBr) : "—",
                atual.NotaMedia > 0 ? atual.NotaMedia.ToString("0.00", PtBr) : "—",
                anterior.NotaMedia > 0 && atual.NotaMedia > 0
                    ? FormatarDeltaSimples(anterior.NotaMedia, atual.NotaMedia)
                    : "—",
                "Percepção de qualidade do atendimento.");

            corpo.Controls.Add(
                CriarCardComGrid(
                    "Comparação dos principais indicadores",
                    grid),
                0,
                3);

            return pagina;
        }

        private Control CriarMiniIndicadorDetalhe(
            string titulo,
            int anterior,
            int atual,
            bool aumentoPodeSerRuim)
        {
            RoundedPanel card = CriarCardDetalhe();
            double? variacao = CalcularVariacaoPercentual(anterior, atual);

            Color cor = CorCinza;

            if (variacao.HasValue)
            {
                if (aumentoPodeSerRuim)
                    cor = variacao.Value > 0 ? CorVermelho : CorVerde;
                else
                    cor = variacao.Value >= 0 ? CorVerde : CorVermelho;
            }

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 4,
                ColumnCount = 1,
                Padding = new Padding(13, 8, 10, 7)
            };

            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 24F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 22F));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            layout.Controls.Add(new Label
            {
                Dock = DockStyle.Fill,
                Text = titulo,
                ForeColor = CorTexto,
                Font = new Font("Segoe UI", 8.5F)
            }, 0, 0);

            layout.Controls.Add(new Label
            {
                Dock = DockStyle.Fill,
                Text = atual.ToString("N0", PtBr),
                ForeColor = CorTexto,
                Font = new Font("Segoe UI Semibold", 14F, FontStyle.Bold)
            }, 0, 1);

            layout.Controls.Add(new Label
            {
                Dock = DockStyle.Fill,
                Text = FormatarVariacaoDetalhe(variacao),
                ForeColor = cor,
                Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold)
            }, 0, 2);

            layout.Controls.Add(new Label
            {
                Dock = DockStyle.Fill,
                Text = "Anterior: " + anterior.ToString("N0", PtBr),
                ForeColor = CorCinza,
                Font = new Font("Segoe UI", 7.8F)
            }, 0, 3);

            card.Controls.Add(layout);
            return card;
        }

        private void AdicionarIndicadorResumo(
            DataGridView grid,
            string indicador,
            int anterior,
            int atual,
            string leitura)
        {
            grid.Rows.Add(
                indicador,
                anterior.ToString("N0", PtBr),
                atual.ToString("N0", PtBr),
                FormatarVariacaoDetalhe(
                    CalcularVariacaoPercentual(anterior, atual)),
                leitura);
        }

        // =========================================================
        // ATENDENTES DETALHADO
        // =========================================================
        private Control CriarPaginaAtendentesDetalhado()
        {
            Panel pagina = CriarPaginaDetalheBase(
                "Atendentes",
                "Compare volume, nota, TMR, TME e TMA individualmente.",
                770,
                out TableLayoutPanel corpo);

            corpo.RowCount = 3;
            corpo.RowStyles.Add(new RowStyle(SizeType.Absolute, 125F));
            corpo.RowStyles.Add(new RowStyle(SizeType.Absolute, 555F));

            if (relatorioAnteriorAtual == null || relatorioAtualAtual == null)
            {
                corpo.Controls.Add(CriarAvisoSemComparacao(), 0, 1);
                return pagina;
            }

            RelatorioMensal anterior = relatorioAnteriorAtual;
            RelatorioMensal atual = relatorioAtualAtual;

            List<DetalheAtendente> linhas =
                MontarComparacaoAtendentes(anterior, atual);

            TableLayoutPanel destaques = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 1,
                Padding = new Padding(0, 2, 0, 5)
            };

            for (int i = 0; i < 4; i++)
                destaques.ColumnStyles.Add(
                    new ColumnStyle(SizeType.Percent, 25F));

            DetalheAtendente? maiorEvolucao =
                linhas
                    .Where(x => x.Variacao.HasValue)
                    .OrderByDescending(x => x.Variacao)
                    .FirstOrDefault();

            DetalheAtendente? maiorQueda =
                linhas
                    .Where(x => x.Variacao.HasValue)
                    .OrderBy(x => x.Variacao)
                    .FirstOrDefault();

            DetalheAtendente? melhorNota =
                linhas
                    .Where(x => x.NotaAtual.HasValue)
                    .OrderByDescending(x => x.NotaAtual)
                    .ThenByDescending(x => x.Atual)
                    .FirstOrDefault();

            DetalheAtendente? melhorTmr =
                linhas
                    .Where(x => x.TmrAnterior.HasValue && x.TmrAtual.HasValue)
                    .OrderBy(x => x.TmrAtual!.Value - x.TmrAnterior!.Value)
                    .FirstOrDefault();

            destaques.Controls.Add(
                CriarCardDestaqueDetalhe(
                    "Maior evolução de volume",
                    maiorEvolucao?.Nome ?? "—",
                    FormatarVariacaoDetalhe(maiorEvolucao?.Variacao),
                    CorVerde),
                0,
                0);

            destaques.Controls.Add(
                CriarCardDestaqueDetalhe(
                    "Maior queda de volume",
                    maiorQueda?.Nome ?? "—",
                    FormatarVariacaoDetalhe(maiorQueda?.Variacao),
                    CorVermelho),
                1,
                0);

            destaques.Controls.Add(
                CriarCardDestaqueDetalhe(
                    "Melhor nota atual",
                    melhorNota?.Nome ?? "—",
                    melhorNota?.NotaAtual?.ToString("0.00", PtBr) ?? "—",
                    CorAmarelo),
                2,
                0);

            destaques.Controls.Add(
                CriarCardDestaqueDetalhe(
                    "Maior melhora no TMR",
                    melhorTmr?.Nome ?? "—",
                    melhorTmr == null
                        ? "—"
                        : FormatarDiferencaTempoDetalhe(
                            melhorTmr.TmrAnterior,
                            melhorTmr.TmrAtual),
                    CorAzul),
                3,
                0);

            corpo.Controls.Add(destaques, 0, 1);

            DataGridView grid = CriarGridDetalhe();
            grid.AutoSizeColumnsMode =
                DataGridViewAutoSizeColumnsMode.DisplayedCells;

            grid.Columns.Add("Atendente", "Atendente");
            grid.Columns.Add("Anterior", anterior.Periodo + "\nFinalizados");
            grid.Columns.Add("Atual", atual.Periodo + "\nFinalizados");
            grid.Columns.Add("Variacao", "Variação");
            grid.Columns.Add("NotaAnterior", "Nota anterior");
            grid.Columns.Add("NotaAtual", "Nota atual");
            grid.Columns.Add("TMRAnterior", "TMR anterior");
            grid.Columns.Add("TMRAtual", "TMR atual");
            grid.Columns.Add("TMEAnterior", "TME anterior");
            grid.Columns.Add("TMEAtual", "TME atual");
            grid.Columns.Add("TMAAnterior", "TMA anterior");
            grid.Columns.Add("TMAAtual", "TMA atual");
            grid.Columns.Add("Diagnostico", "Diagnóstico");

            DataGridViewColumn? colDiagnostico = grid.Columns["Diagnostico"];

            if (colDiagnostico != null)
            {
                colDiagnostico.AutoSizeMode =
                    DataGridViewAutoSizeColumnMode.Fill;

                colDiagnostico.MinimumWidth = 210;
            }

            foreach (DetalheAtendente linha in linhas.OrderByDescending(x => x.Atual))
            {
                grid.Rows.Add(
                    linha.Nome,
                    linha.Anterior.ToString("N0", PtBr),
                    linha.Atual.ToString("N0", PtBr),
                    FormatarVariacaoDetalhe(linha.Variacao),
                    linha.NotaAnterior?.ToString("0.00", PtBr) ?? "—",
                    linha.NotaAtual?.ToString("0.00", PtBr) ?? "—",
                    FormatarTempoNullable(linha.TmrAnterior),
                    FormatarTempoNullable(linha.TmrAtual),
                    FormatarTempoNullable(linha.TmeAnterior),
                    FormatarTempoNullable(linha.TmeAtual),
                    FormatarTempoNullable(linha.TmaAnterior),
                    FormatarTempoNullable(linha.TmaAtual),
                    linha.Diagnostico);
            }

            corpo.Controls.Add(
                CriarCardComGrid(
                    "Comparação completa da equipe",
                    grid),
                0,
                2);

            return pagina;
        }

        private List<DetalheAtendente> MontarComparacaoAtendentes(
            RelatorioMensal anterior,
            RelatorioMensal atual)
        {
            HashSet<string> nomes =
                new HashSet<string>(
                    anterior.FinalizadosPorAtendente.Keys,
                    StringComparer.OrdinalIgnoreCase);

            nomes.UnionWith(atual.FinalizadosPorAtendente.Keys);
            nomes.UnionWith(anterior.Notas.Keys);
            nomes.UnionWith(atual.Notas.Keys);
            nomes.UnionWith(anterior.TMR.Keys);
            nomes.UnionWith(atual.TMR.Keys);
            nomes.UnionWith(anterior.TME.Keys);
            nomes.UnionWith(atual.TME.Keys);
            nomes.UnionWith(anterior.TMA.Keys);
            nomes.UnionWith(atual.TMA.Keys);

            List<DetalheAtendente> lista = new List<DetalheAtendente>();

            foreach (string nome in nomes)
            {
                int ant =
                    anterior.FinalizadosPorAtendente.TryGetValue(
                        nome,
                        out int a)
                        ? a
                        : 0;

                int atu =
                    atual.FinalizadosPorAtendente.TryGetValue(
                        nome,
                        out int b)
                        ? b
                        : 0;

                NotaAtendente? notaAnt =
                    anterior.Notas.TryGetValue(
                        nome,
                        out NotaAtendente? na)
                        ? na
                        : null;

                NotaAtendente? notaAtu =
                    atual.Notas.TryGetValue(
                        nome,
                        out NotaAtendente? nb)
                        ? nb
                        : null;

                TimeSpan? tmrAnt =
                    anterior.TMR.TryGetValue(
                        nome,
                        out TimeSpan t1)
                        ? t1
                        : null;

                TimeSpan? tmrAtu =
                    atual.TMR.TryGetValue(
                        nome,
                        out TimeSpan t2)
                        ? t2
                        : null;

                TimeSpan? tmeAnt =
                    anterior.TME.TryGetValue(
                        nome,
                        out TimeSpan e1)
                        ? e1
                        : null;

                TimeSpan? tmeAtu =
                    atual.TME.TryGetValue(
                        nome,
                        out TimeSpan e2)
                        ? e2
                        : null;

                TimeSpan? tmaAnt =
                    anterior.TMA.TryGetValue(
                        nome,
                        out TimeSpan m1)
                        ? m1
                        : null;

                TimeSpan? tmaAtu =
                    atual.TMA.TryGetValue(
                        nome,
                        out TimeSpan m2)
                        ? m2
                        : null;

                double? variacao =
                    ant > 0
                        ? ((atu - ant) / (double)ant) * 100d
                        : null;

                lista.Add(new DetalheAtendente
                {
                    Nome = nome,
                    Anterior = ant,
                    Atual = atu,
                    Variacao = variacao,
                    NotaAnterior =
                        notaAnt != null && notaAnt.Votos > 0
                            ? notaAnt.Media
                            : null,
                    NotaAtual =
                        notaAtu != null && notaAtu.Votos > 0
                            ? notaAtu.Media
                            : null,
                    TmrAnterior = tmrAnt,
                    TmrAtual = tmrAtu,
                    TmeAnterior = tmeAnt,
                    TmeAtual = tmeAtu,
                    TmaAnterior = tmaAnt,
                    TmaAtual = tmaAtu,
                    Diagnostico =
                        GerarDiagnosticoAtendente(
                            ant,
                            atu,
                            notaAnt,
                            notaAtu,
                            tmrAnt,
                            tmrAtu)
                });
            }

            return lista;
        }

        private Control CriarCardDestaqueDetalhe(
            string titulo,
            string nome,
            string valor,
            Color corValor)
        {
            RoundedPanel card = CriarCardDetalhe();

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(13, 9, 10, 8)
            };

            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 24F));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F));

            layout.Controls.Add(new Label
            {
                Dock = DockStyle.Fill,
                Text = titulo,
                ForeColor = CorCinza,
                Font = new Font("Segoe UI", 8.3F)
            }, 0, 0);

            layout.Controls.Add(new Label
            {
                Dock = DockStyle.Fill,
                Text = nome,
                ForeColor = CorTexto,
                Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true
            }, 0, 1);

            layout.Controls.Add(new Label
            {
                Dock = DockStyle.Fill,
                Text = valor,
                ForeColor = corValor,
                Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Bold)
            }, 0, 2);

            card.Controls.Add(layout);
            return card;
        }

        // =========================================================
        // MOTIVOS DETALHADO
        // =========================================================
        private Control CriarPaginaMotivosDetalhado()
        {
            Panel pagina = CriarPaginaDetalheBase(
                "Motivos de Atendimento",
                "Veja em detalhes quais motivos aumentaram, diminuíram e onde existem oportunidades.",
                820,
                out TableLayoutPanel corpo);

            corpo.RowCount = 3;
            corpo.RowStyles.Add(new RowStyle(SizeType.Absolute, 325F));
            corpo.RowStyles.Add(new RowStyle(SizeType.Absolute, 405F));

            if (relatorioAnteriorAtual == null || relatorioAtualAtual == null)
            {
                corpo.Controls.Add(CriarAvisoSemComparacao(), 0, 1);
                return pagina;
            }

            RelatorioMensal anterior = relatorioAnteriorAtual;
            RelatorioMensal atual = relatorioAtualAtual;

            List<string> nomes =
                anterior.Motivos.Keys
                    .Union(atual.Motivos.Keys, StringComparer.OrdinalIgnoreCase)
                    .OrderByDescending(
                        nome => Math.Max(
                            anterior.Motivos.TryGetValue(nome, out int a)
                                ? a
                                : 0,
                            atual.Motivos.TryGetValue(nome, out int b)
                                ? b
                                : 0))
                    .ToList();

            List<string> top = nomes.Take(8).ToList();

            ComparisonBarChart chart = new ComparisonBarChart
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Categories =
                    top.Select(AbreviarMotivo).ToArray(),
                PreviousValues =
                    top.Select(
                        m => anterior.Motivos.TryGetValue(
                            m,
                            out int v)
                            ? (double)v
                            : 0d)
                        .ToArray(),
                CurrentValues =
                    top.Select(
                        m => atual.Motivos.TryGetValue(
                            m,
                            out int v)
                            ? (double)v
                            : 0d)
                        .ToArray(),
                PreviousLabel = anterior.Periodo,
                CurrentLabel = atual.Periodo
            };

            RoundedPanel graficoCard = CriarCardDetalhe();

            graficoCard.Controls.Add(chart);

            graficoCard.Controls.Add(new Label
            {
                Dock = DockStyle.Top,
                Height = 40,
                Text =
                    $"Principais motivos - {anterior.Periodo} x {atual.Periodo}",
                Padding = new Padding(13, 11, 0, 0),
                ForeColor = CorTexto,
                Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold)
            });

            corpo.Controls.Add(graficoCard, 0, 1);

            DataGridView grid = CriarGridDetalhe();

            grid.Columns.Add("Motivo", "Motivo");
            grid.Columns.Add("Anterior", anterior.Periodo);
            grid.Columns.Add("Atual", atual.Periodo);
            grid.Columns.Add("Diferenca", "Diferença");
            grid.Columns.Add("Variacao", "Variação");
            grid.Columns.Add("Leitura", "Leitura construtiva");

            DataGridViewColumn? colMotivo = grid.Columns["Motivo"];
            DataGridViewColumn? colLeituraMotivo = grid.Columns["Leitura"];

            if (colMotivo != null)
                colMotivo.FillWeight = 160;

            if (colLeituraMotivo != null)
                colLeituraMotivo.FillWeight = 190;

            foreach (string nome in nomes)
            {
                int ant =
                    anterior.Motivos.TryGetValue(
                        nome,
                        out int a)
                        ? a
                        : 0;

                int atu =
                    atual.Motivos.TryGetValue(
                        nome,
                        out int b)
                        ? b
                        : 0;

                double? variacao =
                    CalcularVariacaoPercentual(
                        ant,
                        atu);

                grid.Rows.Add(
                    nome,
                    ant.ToString("N0", PtBr),
                    atu.ToString("N0", PtBr),
                    (atu - ant).ToString("+0;-0;0", PtBr),
                    FormatarVariacaoDetalhe(variacao),
                    GerarLeituraMotivoDetalhe(
                        nome,
                        variacao));
            }

            corpo.Controls.Add(
                CriarCardComGrid(
                    "Tabela completa dos motivos",
                    grid),
                0,
                2);

            return pagina;
        }

        private string GerarLeituraMotivoDetalhe(
            string motivo,
            double? variacao)
        {
            if (!variacao.HasValue)
                return "Novo motivo ou sem base anterior.";

            string normalizado = Normalizar(motivo);

            if (normalizado.Contains("inatividade") &&
                variacao.Value > 10)
            {
                return
                    "Aumento merece análise do tempo de resposta, " +
                    "abandono do cliente e fluxo de encerramento.";
            }

            if (normalizado.Contains("orcamento") &&
                variacao.Value > 10)
            {
                return
                    "Boa oportunidade comercial; acompanhar quantos " +
                    "orçamentos se transformam em pedidos.";
            }

            if (variacao.Value >= 20)
                return "Crescimento relevante no período.";

            if (variacao.Value <= -20)
                return "Queda relevante; verificar mudança de demanda ou processo.";

            return "Variação moderada.";
        }

        // =========================================================
        // MARKETING DETALHADO
        // =========================================================
        private Control CriarPaginaMarketingDetalhado()
        {
            Panel pagina = CriarPaginaDetalheBase(
                "Marketing",
                "Compare contatos de campanhas, novos contatos e mensagens template.",
                790,
                out TableLayoutPanel corpo);

            corpo.RowCount = 4;
            corpo.RowStyles.Add(new RowStyle(SizeType.Absolute, 125F));
            corpo.RowStyles.Add(new RowStyle(SizeType.Absolute, 265F));
            corpo.RowStyles.Add(new RowStyle(SizeType.Absolute, 310F));

            if (relatorioAnteriorAtual == null || relatorioAtualAtual == null)
            {
                corpo.Controls.Add(CriarAvisoSemComparacao(), 0, 1);
                return pagina;
            }

            RelatorioMensal anterior = relatorioAnteriorAtual;
            RelatorioMensal atual = relatorioAtualAtual;

            TableLayoutPanel indicadores = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 1,
                Padding = new Padding(0, 2, 0, 5)
            };

            for (int i = 0; i < 4; i++)
                indicadores.ColumnStyles.Add(
                    new ColumnStyle(SizeType.Percent, 25F));

            indicadores.Controls.Add(
                CriarMiniIndicadorDetalhe(
                    "Contatos via marketing",
                    anterior.ContatosMarketing,
                    atual.ContatosMarketing,
                    false),
                0,
                0);

            indicadores.Controls.Add(
                CriarMiniIndicadorDetalhe(
                    "Novos contatos",
                    anterior.NovosContatos,
                    atual.NovosContatos,
                    false),
                1,
                0);

            indicadores.Controls.Add(
                CriarMiniIndicadorDetalhe(
                    "Mensagens template",
                    anterior.MensagensTemplate,
                    atual.MensagensTemplate,
                    true),
                2,
                0);

            indicadores.Controls.Add(
                CriarMiniIndicadorDetalhe(
                    "Reagendamentos",
                    anterior.Reagendamentos,
                    atual.Reagendamentos,
                    true),
                3,
                0);

            corpo.Controls.Add(indicadores, 0, 1);

            List<string> campanhas =
                anterior.MarketingCriativos.Keys
                    .Union(
                        atual.MarketingCriativos.Keys,
                        StringComparer.OrdinalIgnoreCase)
                    .OrderByDescending(
                        nome => Math.Max(
                            anterior.MarketingCriativos.TryGetValue(
                                nome,
                                out int a)
                                ? a
                                : 0,
                            atual.MarketingCriativos.TryGetValue(
                                nome,
                                out int b)
                                ? b
                                : 0))
                    .ToList();

            List<string> top = campanhas.Take(7).ToList();

            ComparisonBarChart chart = new ComparisonBarChart
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Categories =
                    top.Select(
                        x => x.Length <= 18
                            ? x
                            : x.Substring(0, 17) + "…")
                        .ToArray(),
                PreviousValues =
                    top.Select(
                        x => anterior.MarketingCriativos.TryGetValue(
                            x,
                            out int v)
                            ? (double)v
                            : 0d)
                        .ToArray(),
                CurrentValues =
                    top.Select(
                        x => atual.MarketingCriativos.TryGetValue(
                            x,
                            out int v)
                            ? (double)v
                            : 0d)
                        .ToArray(),
                PreviousLabel = anterior.Periodo,
                CurrentLabel = atual.Periodo
            };

            RoundedPanel chartCard = CriarCardDetalhe();
            chartCard.Controls.Add(chart);

            chartCard.Controls.Add(new Label
            {
                Dock = DockStyle.Top,
                Height = 40,
                Text = "Contatos por criativo/campanha",
                Padding = new Padding(13, 11, 0, 0),
                ForeColor = CorTexto,
                Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold)
            });

            corpo.Controls.Add(chartCard, 0, 2);

            DataGridView grid = CriarGridDetalhe();

            grid.Columns.Add("Criativo", "Criativo / campanha");
            grid.Columns.Add("Anterior", anterior.Periodo);
            grid.Columns.Add("Atual", atual.Periodo);
            grid.Columns.Add("Diferenca", "Diferença");
            grid.Columns.Add("Variacao", "Variação");
            grid.Columns.Add("Observacao", "Observação");

            DataGridViewColumn? colCriativo = grid.Columns["Criativo"];
            DataGridViewColumn? colObservacaoMarketing = grid.Columns["Observacao"];

            if (colCriativo != null)
                colCriativo.FillWeight = 170;

            if (colObservacaoMarketing != null)
                colObservacaoMarketing.FillWeight = 180;

            foreach (string campanha in campanhas)
            {
                int ant =
                    anterior.MarketingCriativos.TryGetValue(
                        campanha,
                        out int a)
                        ? a
                        : 0;

                int atu =
                    atual.MarketingCriativos.TryGetValue(
                        campanha,
                        out int b)
                        ? b
                        : 0;

                double? variacao =
                    CalcularVariacaoPercentual(
                        ant,
                        atu);

                string observacao =
                    !variacao.HasValue
                        ? "Novo criativo no período atual."
                        : variacao.Value >= 20
                            ? "Crescimento relevante."
                            : variacao.Value <= -20
                                ? "Queda relevante; verificar campanha e período."
                                : "Resultado relativamente estável.";

                grid.Rows.Add(
                    campanha,
                    ant.ToString("N0", PtBr),
                    atu.ToString("N0", PtBr),
                    (atu - ant).ToString("+0;-0;0", PtBr),
                    FormatarVariacaoDetalhe(variacao),
                    observacao);
            }

            corpo.Controls.Add(
                CriarCardComGrid(
                    "Comparação detalhada das campanhas",
                    grid),
                0,
                3);

            return pagina;
        }

        // =========================================================
        // QUALIDADE DOS DADOS
        // =========================================================
        private Control CriarPaginaQualidadeDetalhado()
        {
            Panel pagina = CriarPaginaDetalheBase(
                "Qualidade dos Dados",
                "Confira ausências, divergências e limitações antes de usar os números para decisões.",
                710,
                out TableLayoutPanel corpo);

            corpo.RowCount = 3;
            corpo.RowStyles.Add(new RowStyle(SizeType.Absolute, 205F));
            corpo.RowStyles.Add(new RowStyle(SizeType.Absolute, 420F));

            if (relatorioAnteriorAtual == null || relatorioAtualAtual == null)
            {
                corpo.Controls.Add(CriarAvisoSemComparacao(), 0, 1);
                return pagina;
            }

            RelatorioMensal anterior = relatorioAnteriorAtual;
            RelatorioMensal atual = relatorioAtualAtual;

            TableLayoutPanel topo = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1
            };

            topo.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 50F));

            topo.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 50F));

            topo.Controls.Add(
                CriarCardTextoDetalhe(
                    "Leitura do mês retrasado",
                    CriarResumoQualidade(anterior),
                    CorAmarelo),
                0,
                0);

            topo.Controls.Add(
                CriarCardTextoDetalhe(
                    "Leitura do mês passado",
                    CriarResumoQualidade(atual),
                    CorAmarelo),
                1,
                0);

            corpo.Controls.Add(topo, 0, 1);

            DataGridView grid = CriarGridDetalhe();

            grid.Columns.Add("Item", "Verificação");
            grid.Columns.Add("Anterior", anterior.Periodo);
            grid.Columns.Add("Atual", atual.Periodo);
            grid.Columns.Add("Status", "Status");
            grid.Columns.Add("Observacao", "Observação");

            DataGridViewColumn? colItem = grid.Columns["Item"];
            DataGridViewColumn? colObservacaoQualidade = grid.Columns["Observacao"];

            if (colItem != null)
                colItem.FillWeight = 150;

            if (colObservacaoQualidade != null)
                colObservacaoQualidade.FillWeight = 210;

            grid.Rows.Add(
                "TMA disponível",
                anterior.TmaDisponivel ? "Sim" : "Não",
                atual.TmaDisponivel ? "Sim" : "Não",
                atual.TmaDisponivel ? "OK" : "Atenção",
                atual.TmaDisponivel
                    ? "Indicador disponível para comparação."
                    : "O TMA do período atual não possui valores.");

            grid.Rows.Add(
                "Total de atendimentos da relação x finalizados",
                anterior.AtendimentosRelacao.HasValue
                    ? $"{anterior.AtendimentosRelacao.Value:N0} x {anterior.Finalizados:N0}"
                    : "—",
                atual.AtendimentosRelacao.HasValue
                    ? $"{atual.AtendimentosRelacao.Value:N0} x {atual.Finalizados:N0}"
                    : "—",
                "Informativo",
                "Os dois números podem representar etapas diferentes; a diferença fica visível para conferência.");

            string divergAnterior =
                ObterDivergenciaPossivelVenda(anterior);

            string divergAtual =
                ObterDivergenciaPossivelVenda(atual);

            grid.Rows.Add(
                "Orçamento de fórmula em seções diferentes",
                divergAnterior,
                divergAtual,
                divergAtual == "Sem divergência"
                    ? "OK"
                    : "Atenção",
                "Quando o mesmo indicador aparece com valores diferentes, o sistema não escolhe automaticamente qual é o correto.");

            grid.Rows.Add(
                "Nota média geral",
                anterior.NotaMedia > 0
                    ? anterior.NotaMedia.ToString("0.00", PtBr)
                    : "Ausente",
                atual.NotaMedia > 0
                    ? atual.NotaMedia.ToString("0.00", PtBr)
                    : "Ausente",
                anterior.NotaMedia > 0 && atual.NotaMedia > 0
                    ? "OK"
                    : "Atenção",
                "A média geral é usada no KPI de satisfação.");

            corpo.Controls.Add(
                CriarCardComGrid(
                    "Checklist de integridade",
                    grid),
                0,
                2);

            return pagina;
        }

        private string CriarResumoQualidade(RelatorioMensal relatorio)
        {
            List<string> linhas = new List<string>();

            linhas.Add(
                relatorio.Conversas > 0
                    ? "✓ Total de conversas localizado."
                    : "⚠ Total de conversas não localizado.");

            linhas.Add(
                relatorio.Finalizados > 0
                    ? "✓ Total de finalizados localizado."
                    : "⚠ Total de finalizados não localizado.");

            linhas.Add(
                relatorio.TMR.Count > 0
                    ? "✓ TMR disponível."
                    : "⚠ TMR sem dados.");

            linhas.Add(
                relatorio.TME.Count > 0
                    ? "✓ TME disponível."
                    : "⚠ TME sem dados.");

            linhas.Add(
                relatorio.TmaDisponivel
                    ? "✓ TMA disponível."
                    : "⚠ TMA sem dados.");

            if (ObterDivergenciaPossivelVenda(relatorio) != "Sem divergência")
            {
                linhas.Add(
                    "⚠ Existe divergência em 'Orçamento de fórmula'.");
            }

            return string.Join(
                Environment.NewLine + Environment.NewLine,
                linhas);
        }

        private string ObterDivergenciaPossivelVenda(
            RelatorioMensal relatorio)
        {
            int principal =
                ValorMotivo(
                    relatorio,
                    "Orçamento de fórmula");

            if (relatorio.PossiveisVendas.TryGetValue(
                    "Orçamento de fórmula",
                    out int possivel) &&
                principal != possivel)
            {
                return
                    $"{principal.ToString("N0", PtBr)} x " +
                    $"{possivel.ToString("N0", PtBr)}";
            }

            return "Sem divergência";
        }

        // =========================================================
        // EXPORTAÇÃO
        // =========================================================
        private Control CriarPaginaExportarDetalhado()
        {
            Panel pagina = CriarPaginaDetalheBase(
                "Exportar",
                "Gere o relatório consolidado depois de realizar a comparação.",
                570,
                out TableLayoutPanel corpo);

            corpo.RowCount = 3;
            corpo.RowStyles.Add(new RowStyle(SizeType.Absolute, 230F));
            corpo.RowStyles.Add(new RowStyle(SizeType.Absolute, 250F));

            if (relatorioAnteriorAtual == null || relatorioAtualAtual == null)
            {
                corpo.Controls.Add(CriarAvisoSemComparacao(), 0, 1);
                return pagina;
            }

            TableLayoutPanel cards = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Padding = new Padding(0, 10, 0, 10)
            };

            cards.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 50F));

            cards.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 50F));

            cards.Controls.Add(
                CriarCardAcaoDetalhe(
                    "Exportar Excel",
                    "Gera um arquivo consolidado com resumo, motivos, atendentes, marketing e qualidade.",
                    "Exportar Excel",
                    () =>
                    {
                        MessageBox.Show(
                            "A exportação para Excel pode ser ligada na próxima etapa.",
                            "Exportar Excel",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }),
                0,
                0);

            cards.Controls.Add(
                CriarCardAcaoDetalhe(
                    "Exportar PDF",
                    "Gera uma versão visual do relatório para apresentação e arquivamento.",
                    "Exportar PDF",
                    () =>
                    {
                        MessageBox.Show(
                            "A exportação para PDF pode ser ligada na próxima etapa.",
                            "Exportar PDF",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }),
                1,
                0);

            corpo.Controls.Add(cards, 0, 1);

            corpo.Controls.Add(
                CriarCardTextoDetalhe(
                    "Conteúdo sugerido do relatório",
                    "• Resumo executivo e KPIs\n\n" +
                    "• Comparação dos motivos de atendimento\n\n" +
                    "• Desempenho completo dos atendentes\n\n" +
                    "• Marketing e mensagens template\n\n" +
                    "• Qualidade dos dados\n\n" +
                    "• Alertas e ações recomendadas",
                    CorVerdeEscuro),
                0,
                2);

            return pagina;
        }

        private Control CriarCardAcaoDetalhe(
            string titulo,
            string descricao,
            string textoBotao,
            Action acao)
        {
            RoundedPanel card = CriarCardDetalhe();

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(18, 14, 18, 14)
            };

            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44F));

            layout.Controls.Add(new Label
            {
                Dock = DockStyle.Fill,
                Text = titulo,
                ForeColor = CorTexto,
                Font = new Font("Segoe UI Semibold", 11F, FontStyle.Bold)
            }, 0, 0);

            layout.Controls.Add(new Label
            {
                Dock = DockStyle.Fill,
                Text = descricao,
                ForeColor = CorCinza,
                Font = new Font("Segoe UI", 9F)
            }, 0, 1);

            Button btn = new Button
            {
                Text = textoBotao,
                Width = 145,
                Height = 38,
                Dock = DockStyle.Left,
                BackColor = CorVerde,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI Semibold", 9F),
                Cursor = Cursors.Hand
            };

            btn.FlatAppearance.BorderSize = 0;
            btn.Click += (s, e) => acao();

            layout.Controls.Add(btn, 0, 2);
            card.Controls.Add(layout);

            return card;
        }

        // =========================================================
        // HELPERS DAS TELAS DETALHADAS
        // =========================================================
        private static string FormatarVariacaoDetalhe(
            double? variacao)
        {
            if (!variacao.HasValue)
                return "Novo";

            if (Math.Abs(variacao.Value) < 0.05)
                return "0,0%";

            return
                $"{(variacao.Value > 0 ? "▲" : "▼")} " +
                $"{Math.Abs(variacao.Value).ToString("0.0", PtBr)}%";
        }

        private static string FormatarDeltaPp(
            double anterior,
            double atual)
        {
            double delta = atual - anterior;

            return
                $"{(delta >= 0 ? "▲" : "▼")} " +
                $"{Math.Abs(delta).ToString("0.0", PtBr)} p.p.";
        }

        private static string FormatarDeltaSimples(
            double anterior,
            double atual)
        {
            double delta = atual - anterior;

            return
                $"{(delta >= 0 ? "▲" : "▼")} " +
                $"{Math.Abs(delta).ToString("0.00", PtBr)}";
        }

        private static string FormatarTempoNullable(
            TimeSpan? tempo)
        {
            return tempo.HasValue
                ? FormatarTempo(tempo.Value)
                : "—";
        }

        private static string FormatarDiferencaTempoDetalhe(
            TimeSpan? anterior,
            TimeSpan? atual)
        {
            if (!anterior.HasValue || !atual.HasValue)
                return "—";

            TimeSpan diferenca =
                atual.Value - anterior.Value;

            string seta =
                diferenca.TotalSeconds <= 0
                    ? "▼"
                    : "▲";

            diferenca = diferenca.Duration();

            if (diferenca.TotalHours >= 1)
            {
                return
                    $"{seta} {(int)diferenca.TotalHours}h " +
                    $"{diferenca.Minutes:00}m";
            }

            return
                $"{seta} {diferenca.Minutes}m " +
                $"{diferenca.Seconds:00}s";
        }

        private sealed class DetalheAtendente
        {
            public string Nome { get; set; } = string.Empty;
            public int Anterior { get; set; }
            public int Atual { get; set; }
            public double? Variacao { get; set; }
            public double? NotaAnterior { get; set; }
            public double? NotaAtual { get; set; }
            public TimeSpan? TmrAnterior { get; set; }
            public TimeSpan? TmrAtual { get; set; }
            public TimeSpan? TmeAnterior { get; set; }
            public TimeSpan? TmeAtual { get; set; }
            public TimeSpan? TmaAnterior { get; set; }
            public TimeSpan? TmaAtual { get; set; }
            public string Diagnostico { get; set; } = string.Empty;
        }

        // =========================================================
        // SELEÇÃO DOS ARQUIVOS
        // =========================================================
        private void EscolherMesRetrasado(object? sender, EventArgs e)
        {
            using OpenFileDialog ofd = new OpenFileDialog
            {
                Title = "Selecione a planilha do mês retrasado",
                Filter = "Arquivos Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
                CheckFileExists = true,
                Multiselect = false
            };

            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            caminhoMesRetrasado = ofd.FileName;
            relatorioAnteriorAtual = null;
            relatorioAtualAtual = null;
            LimparResultados();

            txtMesRetrasado.Text = Path.GetFileName(ofd.FileName);
            lblPeriodo1.Text = "Arquivo selecionado";
            lblPeriodo2.Text = string.Empty;
            AtualizarStatusArquivos();
        }

        private void EscolherMesPassado(object? sender, EventArgs e)
        {
            using OpenFileDialog ofd = new OpenFileDialog
            {
                Title = "Selecione a planilha do mês passado",
                Filter = "Arquivos Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
                CheckFileExists = true,
                Multiselect = false
            };

            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            caminhoMesPassado = ofd.FileName;
            relatorioAnteriorAtual = null;
            relatorioAtualAtual = null;
            LimparResultados();

            txtMesPassado.Text = Path.GetFileName(ofd.FileName);
            lblPeriodo2.Text = "Arquivo selecionado";
            lblPeriodo1.Text = string.IsNullOrWhiteSpace(caminhoMesRetrasado)
                ? string.Empty
                : "Arquivo selecionado";
            AtualizarStatusArquivos();
        }

        private void AtualizarStatusArquivos()
        {
            bool primeiro = !string.IsNullOrWhiteSpace(caminhoMesRetrasado);
            bool segundo = !string.IsNullOrWhiteSpace(caminhoMesPassado);

            if (primeiro && segundo)
                lblStatus.Text = "●  Os dois arquivos estão selecionados. Clique em Comparar.";
            else if (primeiro || segundo)
                lblStatus.Text = "●  Um arquivo selecionado. Selecione o segundo.";
            else
                lblStatus.Text = "●  Aguardando os dois arquivos.";
        }

        // =========================================================
        // COMPARAR - AGORA LÊ OS DOIS EXCEL DE VERDADE
        // =========================================================
        private void BtnComparar_Click(object? sender, EventArgs e)
        {
            if (!ValidarArquivos())
                return;

            try
            {
                Cursor = Cursors.WaitCursor;
                lblStatus.Text = "●  Lendo e comparando as planilhas...";
                Application.DoEvents();

                RelatorioMensal anterior = LerRelatorioExcel(caminhoMesRetrasado);
                RelatorioMensal atual = LerRelatorioExcel(caminhoMesPassado);

                relatorioAnteriorAtual = anterior;
                relatorioAtualAtual = atual;

                PreencherDashboard(anterior, atual);

                lblPeriodo1.Text = anterior.Periodo;
                lblPeriodo2.Text = atual.Periodo;
                lblStatus.Text = $"●  Comparação concluída: {anterior.Periodo} x {atual.Periodo}.";
            }
            catch (Exception ex)
            {
                LimparResultados();
                lblStatus.Text = "●  Não foi possível concluir a comparação.";

                MessageBox.Show(
                    "Não foi possível ler os arquivos.\n\n" +
                    "Verifique se as planilhas mantêm as abas e os títulos do relatório.\n\n" +
                    "Detalhe: " + ex.Message,
                    "Erro ao comparar",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private bool ValidarArquivos()
        {
            if (string.IsNullOrWhiteSpace(caminhoMesRetrasado) || !File.Exists(caminhoMesRetrasado))
            {
                MessageBox.Show("Selecione a planilha do mês retrasado.", "Arquivo necessário",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (string.IsNullOrWhiteSpace(caminhoMesPassado) || !File.Exists(caminhoMesPassado))
            {
                MessageBox.Show("Selecione a planilha do mês passado.", "Arquivo necessário",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (string.Equals(caminhoMesRetrasado, caminhoMesPassado, StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("Selecione dois arquivos diferentes.", "Arquivos iguais",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        // =========================================================
        // LEITURA DO RELATÓRIO
        // =========================================================
        private RelatorioMensal LerRelatorioExcel(string caminho)
        {
            using XLWorkbook workbook = new XLWorkbook(caminho);

            RelatorioMensal r = new RelatorioMensal();

            IXLWorksheet wsRelacao = ObterAbaObrigatoria(workbook, "Relação Conversas-Atendimento");
            IXLWorksheet wsNovos = ObterAbaObrigatoria(workbook, "Mensagens por número");
            IXLWorksheet wsMarketing = ObterAbaObrigatoria(workbook, "Marketing");
            IXLWorksheet wsTemplate = ObterAbaObrigatoria(workbook, "Mensagem Template");
            IXLWorksheet wsMotivos = ObterAbaObrigatoria(workbook, "Motivos Atendimento");
            IXLWorksheet wsFinalizados = ObterAbaObrigatoria(workbook, "Atendimentos finalizados");
            IXLWorksheet wsTmr = ObterAbaObrigatoria(workbook, "TMR");

            IXLWorksheet? wsTme = ObterAbaOpcional(workbook, "TME");
            IXLWorksheet? wsTma = ObterAbaOpcional(workbook, "TMA");
            IXLWorksheet? wsMedia = ObterAbaOpcional(workbook, "MEDIA - IND", "Média individual");

            r.DataReferencia = DetectarDataReferencia(wsRelacao) ?? DateTime.MinValue;
            r.Periodo = r.DataReferencia == DateTime.MinValue
                ? Path.GetFileNameWithoutExtension(caminho)
                : CultureInfo.GetCultureInfo("pt-BR").DateTimeFormat.GetMonthName(r.DataReferencia.Month)
                    .FirstCharToUpper() + "/" + r.DataReferencia.Year;

            List<int> totaisRelacao = LerTotaisMes(wsRelacao);
            r.Conversas = totaisRelacao.Count > 0 ? totaisRelacao[0] : 0;
            r.AtendimentosRelacao = totaisRelacao.Count > 1 ? totaisRelacao[1] : null;

            r.NovosContatos = LerPrimeiroTotalColunaAB(wsNovos);
            r.ContatosMarketing = LerTotalMarketing(wsMarketing);
            r.MarketingCriativos = LerCriativosMarketing(wsMarketing);
            r.Reagendamentos = LerValorPorLabel(wsTemplate, "reagendamento");
            r.MensagensTemplate = LerValorPorLabel(wsTemplate, "total geral");

            r.Motivos = LerTabelaMotivosPrincipal(wsMotivos);
            r.PossiveisVendas = LerSecaoMotivos(wsMotivos, "possivel venda");

            (r.Finalizados, r.FinalizadosPorAtendente) = LerFinalizados(wsFinalizados);
            r.TMR = LerTempos(wsTmr);
            r.TME = wsTme == null
                ? new Dictionary<string, TimeSpan>(StringComparer.OrdinalIgnoreCase)
                : LerTempos(wsTme);

            r.TMA = wsTma == null
                ? new Dictionary<string, TimeSpan>(StringComparer.OrdinalIgnoreCase)
                : LerTempos(wsTma);

            r.TmaDisponivel = r.TMA.Count > 0;

            if (wsMedia != null)
            {
                r.Notas = LerNotasAtendentes(wsMedia);
                r.NotaMedia = LerMediaGeral(wsMedia);
            }

            return r;
        }

        private static IXLWorksheet ObterAbaObrigatoria(XLWorkbook workbook, string nome)
        {
            if (workbook.TryGetWorksheet(nome, out IXLWorksheet? ws))
                return ws;

            throw new InvalidOperationException($"A aba '{nome}' não foi encontrada.");
        }

        private static IXLWorksheet? ObterAbaOpcional(XLWorkbook workbook, params string[] nomes)
        {
            foreach (string nome in nomes)
            {
                if (workbook.TryGetWorksheet(nome, out IXLWorksheet? ws))
                    return ws;
            }

            return null;
        }

        private static DateTime? DetectarDataReferencia(IXLWorksheet ws)
        {
            IXLRange? used = ws.RangeUsed();
            if (used == null)
                return null;

            foreach (IXLCell cell in used.CellsUsed())
            {
                if (cell.TryGetValue<DateTime>(out DateTime dt) && dt.Year >= 2020 && dt.Year <= 2100)
                    return dt;

                string texto = cell.GetFormattedString().Trim();
                if (DateTime.TryParse(texto, PtBr, DateTimeStyles.None, out dt) && dt.Year >= 2020 && dt.Year <= 2100)
                    return dt;
            }

            return null;
        }

        private static List<int> LerTotaisMes(IXLWorksheet ws)
        {
            List<(int row, int col, int valor)> encontrados =
                new List<(int row, int col, int valor)>();

            IXLRange? used = ws.RangeUsed();

            if (used == null)
                return new List<int>();

            foreach (IXLCell cell in used.CellsUsed())
            {
                string texto = Normalizar(cell.GetFormattedString());

                if (!texto.Contains("total mes"))
                    continue;

                int row = cell.Address.RowNumber;
                int col = cell.Address.ColumnNumber;
                int? valor = LerInteiro(ws.Cell(row, col + 1));

                if (valor.HasValue)
                    encontrados.Add((row, col, valor.Value));
            }

            return encontrados
                .OrderBy(x => x.row)
                .ThenBy(x => x.col)
                .Select(x => x.valor)
                .ToList();
        }

        private static int LerPrimeiroTotalColunaAB(IXLWorksheet ws)
        {
            IXLRange? used = ws.RangeUsed();
            if (used == null)
                return 0;

            foreach (IXLRangeRow row in used.RowsUsed())
            {
                string label = Normalizar(row.Cell(1).GetFormattedString());
                if (label == "total" || label.StartsWith("total "))
                {
                    int? valor = LerInteiro(row.Cell(2));
                    if (valor.HasValue)
                        return valor.Value;
                }
            }

            return 0;
        }

        private static int LerTotalMarketing(IXLWorksheet ws)
        {
            IXLRange? used = ws.RangeUsed();
            if (used == null)
                return 0;

            foreach (IXLRangeRow row in used.RowsUsed())
            {
                string label = Normalizar(row.Cell(1).GetFormattedString());
                if (!label.StartsWith("total"))
                    continue;

                int? valor = LerInteiro(row.Cell(3));
                if (valor.HasValue)
                    return valor.Value;
            }

            return 0;
        }

        private static Dictionary<string, int> LerCriativosMarketing(IXLWorksheet ws)
        {
            Dictionary<string, int> itens =
                new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            IXLRange? used = ws.RangeUsed();

            if (used == null)
                return itens;

            bool iniciou = false;

            foreach (IXLRangeRow row in used.RowsUsed())
            {
                string colB =
                    row.Cell(2).GetFormattedString().Trim();

                string normB =
                    Normalizar(colB);

                string normA =
                    Normalizar(
                        row.Cell(1).GetFormattedString());

                if (!iniciou)
                {
                    if (normB == "criativos")
                        iniciou = true;

                    continue;
                }

                if (normA.StartsWith("total"))
                    break;

                if (string.IsNullOrWhiteSpace(colB))
                    continue;

                int? valor =
                    LerInteiro(row.Cell(3));

                if (valor.HasValue)
                    itens[colB] = valor.Value;
            }

            return itens;
        }

        private static int LerValorPorLabel(IXLWorksheet ws, string trecho)
        {
            string procurado = Normalizar(trecho);
            IXLRange? used = ws.RangeUsed();

            if (used == null)
                return 0;

            foreach (IXLRangeRow row in used.RowsUsed())
            {
                string label = Normalizar(row.Cell(1).GetFormattedString());

                if (!label.Contains(procurado))
                    continue;

                int? valor = LerInteiro(row.Cell(2));
                if (valor.HasValue)
                    return valor.Value;
            }

            return 0;
        }

        private static Dictionary<string, int> LerTabelaMotivosPrincipal(IXLWorksheet ws)
        {
            Dictionary<string, int> dados = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            IXLRange? used = ws.RangeUsed();

            if (used == null)
                return dados;

            bool iniciou = false;

            foreach (IXLRangeRow row in used.RowsUsed())
            {
                string nome = row.Cell(1).GetFormattedString().Trim();
                string norm = Normalizar(nome);

                if (!iniciou)
                {
                    if (norm == "motivo")
                        iniciou = true;
                    continue;
                }

                if (norm.StartsWith("total"))
                    break;

                int? valor = LerInteiro(row.Cell(2));
                if (!string.IsNullOrWhiteSpace(nome) && valor.HasValue)
                    dados[nome] = valor.Value;
            }

            return dados;
        }

        private static Dictionary<string, int> LerSecaoMotivos(IXLWorksheet ws, string tituloSecao)
        {
            Dictionary<string, int> dados = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            IXLRange? used = ws.RangeUsed();

            if (used == null)
                return dados;

            bool iniciou = false;
            string alvo = Normalizar(tituloSecao);

            foreach (IXLRangeRow row in used.RowsUsed())
            {
                string nome = row.Cell(1).GetFormattedString().Trim();
                string norm = Normalizar(nome);

                if (!iniciou)
                {
                    if (norm.StartsWith(alvo))
                        iniciou = true;
                    continue;
                }

                if (norm.StartsWith("total"))
                    break;

                int? valor = LerInteiro(row.Cell(2));
                if (!string.IsNullOrWhiteSpace(nome) && valor.HasValue)
                    dados[nome] = valor.Value;
            }

            return dados;
        }

        private static (int total, Dictionary<string, int> porAtendente) LerFinalizados(IXLWorksheet ws)
        {
            Dictionary<string, int> agentes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            int? total = null;
            IXLRange? used = ws.RangeUsed();

            if (used == null)
                return (0, agentes);

            foreach (IXLRangeRow row in used.RowsUsed())
            {
                string nome = row.Cell(1).GetFormattedString().Trim();
                string norm = Normalizar(nome);
                int? valor = LerInteiro(row.Cell(2));

                if (norm.Contains("total geral") && valor.HasValue)
                {
                    total = valor.Value;
                    continue;
                }

                if (string.IsNullOrWhiteSpace(nome) || !valor.HasValue)
                    continue;

                if (norm.Contains("atendimentos finalizados") || norm.StartsWith("total"))
                    continue;

                agentes[nome] = valor.Value;
            }

            return (total ?? agentes.Values.Sum(), agentes);
        }

        private static Dictionary<string, TimeSpan> LerTempos(IXLWorksheet ws)
        {
            Dictionary<string, TimeSpan> dados = new Dictionary<string, TimeSpan>(StringComparer.OrdinalIgnoreCase);
            IXLRange? used = ws.RangeUsed();

            if (used == null)
                return dados;

            foreach (IXLRangeRow row in used.RowsUsed())
            {
                string nome = row.Cell(1).GetFormattedString().Trim();
                string norm = Normalizar(nome);

                if (string.IsNullOrWhiteSpace(nome) || norm == "agente")
                    continue;

                TimeSpan? tempo = LerTempo(row.Cell(2));
                if (tempo.HasValue)
                    dados[nome] = tempo.Value;
            }

            return dados;
        }

        private static Dictionary<string, NotaAtendente> LerNotasAtendentes(IXLWorksheet ws)
        {
            Dictionary<string, NotaAtendente> notas = new Dictionary<string, NotaAtendente>(StringComparer.OrdinalIgnoreCase);
            IXLRange? used = ws.RangeUsed();

            if (used == null)
                return notas;

            bool iniciou = false;

            foreach (IXLRangeRow row in used.RowsUsed())
            {
                string nome = row.Cell(1).GetFormattedString().Trim();
                string norm = Normalizar(nome);

                if (!iniciou)
                {
                    if (norm == "atendentes")
                        iniciou = true;
                    continue;
                }

                if (string.IsNullOrWhiteSpace(nome))
                {
                    if (notas.Count > 0)
                        break;
                    continue;
                }

                if (norm.StartsWith("atendentes wl") || norm.StartsWith("atendentes guergel") || norm.StartsWith("atendentes gurgel"))
                    break;

                double? media = LerDouble(row.Cell(2));
                int? votos = LerInteiro(row.Cell(3));

                if (media.HasValue)
                {
                    notas[nome] = new NotaAtendente
                    {
                        Media = media.Value,
                        Votos = votos ?? 0
                    };
                }
            }

            return notas;
        }

        private static double LerMediaGeral(IXLWorksheet ws)
        {
            IXLRange? used = ws.RangeUsed();
            if (used == null)
                return 0;

            foreach (IXLCell cell in used.CellsUsed())
            {
                string texto = Normalizar(cell.GetFormattedString());
                if (!texto.Contains("media geral"))
                    continue;

                int row = cell.Address.RowNumber;
                int col = cell.Address.ColumnNumber;

                for (int offset = 1; offset <= 3; offset++)
                {
                    double? valor = LerDouble(ws.Cell(row, col + offset));
                    if (valor.HasValue)
                        return valor.Value;
                }
            }

            return 0;
        }

        private static int? LerInteiro(IXLCell cell)
        {
            double? d = LerDouble(cell);
            if (!d.HasValue)
                return null;

            return Convert.ToInt32(Math.Round(d.Value, MidpointRounding.AwayFromZero));
        }

        private static double? LerDouble(IXLCell cell)
        {
            if (cell.TryGetValue<double>(out double numero))
                return numero;

            string texto = cell.GetFormattedString().Trim();

            if (double.TryParse(texto, NumberStyles.Any, PtBr, out numero))
                return numero;

            if (double.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out numero))
                return numero;

            return null;
        }

        private static TimeSpan? LerTempo(IXLCell cell)
        {
            if (cell.TryGetValue<TimeSpan>(out TimeSpan ts))
                return ts;

            string texto = cell.GetFormattedString().Trim();
            if (string.IsNullOrWhiteSpace(texto))
                return null;

            if (TimeSpan.TryParse(texto, CultureInfo.InvariantCulture, out ts) ||
                TimeSpan.TryParse(texto, PtBr, out ts))
                return ts;

            Match match = Regex.Match(texto, @"^(?<dias>\d+)d\s+(?<h>\d+):(?<m>\d+):(?<s>\d+)$", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                int dias = int.Parse(match.Groups["dias"].Value);
                int h = int.Parse(match.Groups["h"].Value);
                int m = int.Parse(match.Groups["m"].Value);
                int s = int.Parse(match.Groups["s"].Value);
                return new TimeSpan(dias, h, m, s);
            }

            return null;
        }

        // =========================================================
        // PREENCHER DASHBOARD
        // =========================================================
        private void PreencherDashboard(RelatorioMensal anterior, RelatorioMensal atual)
        {
            double taxaAnterior = CalcularTaxa(anterior.Finalizados, anterior.Conversas);
            double taxaAtual = CalcularTaxa(atual.Finalizados, atual.Conversas);

            AtualizarKpiNumero(kpiConversas, atual.Conversas, anterior.Conversas, anterior.Periodo);
            AtualizarKpiNumero(kpiFinalizados, atual.Finalizados, anterior.Finalizados, anterior.Periodo);
            AtualizarKpiPontosPercentuais(kpiTaxaFinalizacao, taxaAtual, taxaAnterior, anterior.Periodo);
            AtualizarKpiNumero(kpiReagendamentos, atual.Reagendamentos, anterior.Reagendamentos, anterior.Periodo);
            AtualizarKpiNumero(kpiNovosContatos, atual.NovosContatos, anterior.NovosContatos, anterior.Periodo);
            AtualizarKpiNota(kpiNotaMedia, atual.NotaMedia, anterior.NotaMedia, anterior.Periodo);

            lblResumoTexto.Text = GerarResumoExecutivo(anterior, atual, taxaAnterior, taxaAtual);
            lblAlertasTexto.Text = GerarAlertas(anterior, atual, taxaAnterior, taxaAtual);
            lblOportunidadesTexto.Text = GerarOportunidades(anterior, atual);
            lblAcoesTexto.Text = GerarAcoes(anterior, atual, taxaAnterior, taxaAtual);
            lblDestaquesTexto.Text = GerarDestaques(anterior, atual);
            lblIntegridadeTexto.Text = GerarIntegridade(anterior, atual);

            PreencherGraficoMotivos(anterior, atual);
            PreencherTabelaAtendentes(anterior, atual);
        }

        private void AtualizarKpiNumero(KpiView kpi, int atual, int anterior, string periodoAnterior)
        {
            kpi.Valor.Text = atual.ToString("N0", PtBr);
            kpi.Anterior.Text = $"vs. {anterior.ToString("N0", PtBr)} em {periodoAnterior}";

            double? variacao = CalcularVariacaoPercentual(anterior, atual);
            AplicarVariacao(kpi.Variacao, variacao, "%");
        }

        private void AtualizarKpiPontosPercentuais(KpiView kpi, double atual, double anterior, string periodoAnterior)
        {
            kpi.Valor.Text = atual.ToString("0.0", PtBr) + "%";
            kpi.Anterior.Text = $"vs. {anterior.ToString("0.0", PtBr)}% em {periodoAnterior}";

            double delta = atual - anterior;
            string seta = delta > 0.0001 ? "▲" : delta < -0.0001 ? "▼" : "•";
            kpi.Variacao.Text = $"{seta} {Math.Abs(delta).ToString("0.0", PtBr)} p.p.";
            kpi.Variacao.ForeColor = delta >= 0 ? CorVerde : CorVermelho;
        }

        private void AtualizarKpiNota(KpiView kpi, double atual, double anterior, string periodoAnterior)
        {
            kpi.Valor.Text = atual > 0 ? atual.ToString("0.00", PtBr) : "—";
            kpi.Anterior.Text = anterior > 0 ? $"vs. {anterior.ToString("0.00", PtBr)} em {periodoAnterior}" : string.Empty;

            if (atual <= 0 || anterior <= 0)
            {
                kpi.Variacao.Text = string.Empty;
                return;
            }

            double delta = atual - anterior;
            string seta = delta > 0.0001 ? "▲" : delta < -0.0001 ? "▼" : "•";
            kpi.Variacao.Text = $"{seta} {Math.Abs(delta).ToString("0.00", PtBr)}";
            kpi.Variacao.ForeColor = delta >= 0 ? CorVerde : CorVermelho;
        }

        private void AplicarVariacao(Label label, double? variacao, string sufixo)
        {
            if (!variacao.HasValue)
            {
                label.Text = string.Empty;
                return;
            }

            double v = variacao.Value;
            string seta = v > 0.0001 ? "▲" : v < -0.0001 ? "▼" : "•";
            label.Text = $"{seta} {Math.Abs(v).ToString("0.0", PtBr)}{sufixo}";
            label.ForeColor = v >= 0 ? CorVerde : CorVermelho;
        }

        private void PreencherGraficoMotivos(RelatorioMensal anterior, RelatorioMensal atual)
        {
            List<string> top = atual.Motivos
                .OrderByDescending(x => x.Value)
                .Take(6)
                .Select(x => x.Key)
                .ToList();

            graficoMotivos.Categories = top.Select(AbreviarMotivo).ToArray();
            graficoMotivos.PreviousValues = top.Select(m => anterior.Motivos.TryGetValue(m, out int v) ? (double)v : 0d).ToArray();
            graficoMotivos.CurrentValues = top.Select(m => atual.Motivos.TryGetValue(m, out int v) ? (double)v : 0d).ToArray();
            graficoMotivos.PreviousLabel = anterior.Periodo;
            graficoMotivos.CurrentLabel = atual.Periodo;
            graficoMotivos.Invalidate();
        }

        private void PreencherTabelaAtendentes(RelatorioMensal anterior, RelatorioMensal atual)
        {
            dgvAtendentes.Rows.Clear();
            dgvAtendentes.Columns["Anterior"].HeaderText = anterior.Periodo;
            dgvAtendentes.Columns["Atual"].HeaderText = atual.Periodo;

            HashSet<string> nomes = new HashSet<string>(anterior.FinalizadosPorAtendente.Keys, StringComparer.OrdinalIgnoreCase);
            nomes.UnionWith(atual.FinalizadosPorAtendente.Keys);

            var linhas = nomes
                .Select(nome =>
                {
                    int ant = anterior.FinalizadosPorAtendente.TryGetValue(nome, out int a) ? a : 0;
                    int atu = atual.FinalizadosPorAtendente.TryGetValue(nome, out int b) ? b : 0;

                    double? variacao = ant > 0 ? ((atu - ant) / (double)ant) * 100d : null;

                    NotaAtendente? notaAtual = atual.Notas.TryGetValue(nome, out NotaAtendente? nAtual) ? nAtual : null;
                    NotaAtendente? notaAnterior = anterior.Notas.TryGetValue(nome, out NotaAtendente? nAnterior) ? nAnterior : null;

                    TimeSpan? tmrAtual = atual.TMR.TryGetValue(nome, out TimeSpan ta) ? ta : null;
                    TimeSpan? tmrAnterior = anterior.TMR.TryGetValue(nome, out TimeSpan tb) ? tb : null;

                    return new
                    {
                        Nome = nome,
                        Anterior = ant,
                        Atual = atu,
                        Variacao = variacao,
                        NotaAtual = notaAtual,
                        NotaAnterior = notaAnterior,
                        TmrAtual = tmrAtual,
                        TmrAnterior = tmrAnterior,
                        Diagnostico = GerarDiagnosticoAtendente(ant, atu, notaAnterior, notaAtual, tmrAnterior, tmrAtual)
                    };
                })
                .OrderByDescending(x => x.Atual)
                .ThenBy(x => x.Nome)
                .ToList();

            foreach (var linha in linhas)
            {
                string textoVariacao;

                if (linha.Anterior == 0 && linha.Atual > 0)
                    textoVariacao = "Novo";
                else if (!linha.Variacao.HasValue)
                    textoVariacao = "—";
                else
                {
                    string seta = linha.Variacao.Value >= 0 ? "▲" : "▼";
                    textoVariacao = $"{seta} {Math.Abs(linha.Variacao.Value).ToString("0.0", PtBr)}%";
                }

                dgvAtendentes.Rows.Add(
                    linha.Nome,
                    linha.Anterior.ToString("N0", PtBr),
                    linha.Atual.ToString("N0", PtBr),
                    textoVariacao,
                    linha.NotaAtual != null ? linha.NotaAtual.Media.ToString("0.00", PtBr) : "—",
                    linha.TmrAtual.HasValue ? FormatarTempo(linha.TmrAtual.Value) : "—",
                    linha.Diagnostico);
            }
        }

        // =========================================================
        // CONSIDERAÇÕES AUTOMÁTICAS
        // =========================================================
        private string GerarResumoExecutivo(RelatorioMensal anterior, RelatorioMensal atual, double taxaAnterior, double taxaAtual)
        {
            double? varConversas = CalcularVariacaoPercentual(anterior.Conversas, atual.Conversas);
            double? varFinalizados = CalcularVariacaoPercentual(anterior.Finalizados, atual.Finalizados);
            double? varNovos = CalcularVariacaoPercentual(anterior.NovosContatos, atual.NovosContatos);
            double deltaTaxa = taxaAtual - taxaAnterior;

            List<string> partes = new List<string>();

            if (varConversas.HasValue)
                partes.Add($"O volume de conversas {(varConversas >= 0 ? "cresceu" : "caiu")} {Math.Abs(varConversas.Value).ToString("0.0", PtBr)}%.");

            if (varFinalizados.HasValue)
                partes.Add($"Os atendimentos finalizados {(varFinalizados >= 0 ? "cresceram" : "caíram")} {Math.Abs(varFinalizados.Value).ToString("0.0", PtBr)}%.");

            if (varNovos.HasValue)
                partes.Add($"Novos contatos {(varNovos >= 0 ? "aumentaram" : "diminuíram")} {Math.Abs(varNovos.Value).ToString("0.0", PtBr)}%.");

            partes.Add($"A taxa de finalização ficou em {taxaAtual.ToString("0.0", PtBr)}% ({(deltaTaxa >= 0 ? "+" : "")}{deltaTaxa.ToString("0.0", PtBr)} p.p.).");

            if (atual.NotaMedia > 0)
                partes.Add($"A nota média do atendimento foi {atual.NotaMedia.ToString("0.00", PtBr)}.");

            return string.Join(Environment.NewLine + Environment.NewLine, partes.Take(5));
        }

        private string GerarAlertas(RelatorioMensal anterior, RelatorioMensal atual, double taxaAnterior, double taxaAtual)
        {
            List<string> alertas = new List<string>();

            double deltaTaxa = taxaAtual - taxaAnterior;
            if (deltaTaxa < -0.5)
                alertas.Add($"● Taxa de finalização caiu {Math.Abs(deltaTaxa).ToString("0.0", PtBr)} p.p.");

            int inatAnt = ValorMotivo(anterior, "Encerrado por inatividade do cliente");
            int inatAtual = ValorMotivo(atual, "Encerrado por inatividade do cliente");
            double? varInatividade = CalcularVariacaoPercentual(inatAnt, inatAtual);
            if (varInatividade > 10)
                alertas.Add($"● Encerramentos por inatividade aumentaram {varInatividade.Value.ToString("0.0", PtBr)}%.");

            double? varReag = CalcularVariacaoPercentual(anterior.Reagendamentos, atual.Reagendamentos);
            if (varReag > 15)
                alertas.Add($"● Reagendamentos cresceram {varReag.Value.ToString("0.0", PtBr)}%.");

            int tmrPiorou = ContarTmrComPiora(anterior, atual, TimeSpan.FromMinutes(10));
            if (tmrPiorou > 0)
                alertas.Add($"● {tmrPiorou} atendente(s) tiveram aumento superior a 10 min no TMR.");

            if (alertas.Count == 0)
                alertas.Add("✓ Nenhum alerta relevante foi identificado pelos critérios atuais.");

            return string.Join(Environment.NewLine + Environment.NewLine, alertas.Take(5));
        }

        private string GerarOportunidades(RelatorioMensal anterior, RelatorioMensal atual)
        {
            List<string> itens = new List<string>();

            double? varMarketing = CalcularVariacaoPercentual(anterior.ContatosMarketing, atual.ContatosMarketing);
            if (varMarketing.HasValue && varMarketing > 0)
                itens.Add($"✓ Contatos de marketing cresceram {varMarketing.Value.ToString("0.0", PtBr)}%.");

            int orcAnt = ValorMotivo(anterior, "Orçamento de fórmula");
            int orcAtual = ValorMotivo(atual, "Orçamento de fórmula");
            double? varOrc = CalcularVariacaoPercentual(orcAnt, orcAtual);
            if (varOrc.HasValue && varOrc > 0)
                itens.Add($"✓ Orçamentos de fórmula aumentaram {varOrc.Value.ToString("0.0", PtBr)}%.");

            double? varNovos = CalcularVariacaoPercentual(anterior.NovosContatos, atual.NovosContatos);
            if (varNovos.HasValue && varNovos > 0)
                itens.Add($"✓ Novos contatos cresceram {varNovos.Value.ToString("0.0", PtBr)}%.");

            if (atual.NotaMedia >= 4.7)
                itens.Add($"✓ Satisfação permanece em nível elevado ({atual.NotaMedia.ToString("0.00", PtBr)}).");

            if (itens.Count == 0)
                itens.Add("• Não houve oportunidade com crescimento relevante pelos critérios atuais.");

            return string.Join(Environment.NewLine + Environment.NewLine, itens.Take(5));
        }

        private string GerarAcoes(RelatorioMensal anterior, RelatorioMensal atual, double taxaAnterior, double taxaAtual)
        {
            List<string> acoes = new List<string>();

            if (taxaAtual < taxaAnterior - 0.5)
                acoes.Add("Revisar horários de maior fila e distribuição entre atendentes.");

            int inatAnt = ValorMotivo(anterior, "Encerrado por inatividade do cliente");
            int inatAtual = ValorMotivo(atual, "Encerrado por inatividade do cliente");
            if (CalcularVariacaoPercentual(inatAnt, inatAtual) > 10)
                acoes.Add("Investigar por que mais conversas estão encerrando por inatividade.");

            if (ContarTmrComPiora(anterior, atual, TimeSpan.FromMinutes(10)) > 0)
                acoes.Add("Acompanhar os atendentes com maior aumento no TMR e verificar carga/fila.");

            if (CalcularVariacaoPercentual(anterior.Reagendamentos, atual.Reagendamentos) > 15)
                acoes.Add("Separar os reagendamentos por motivo para entender a origem do aumento.");

            int orcAtual = ValorMotivo(atual, "Orçamento de fórmula");
            if (orcAtual > 0)
                acoes.Add("Acompanhar o funil orçamento → pedido → venda para medir conversão.");

            if (acoes.Count == 0)
                acoes.Add("Manter acompanhamento dos principais indicadores e repetir a comparação no próximo mês.");

            return string.Join(Environment.NewLine + Environment.NewLine,
                acoes.Take(5).Select((x, i) => $"{i + 1}. {x}"));
        }

        private string GerarDestaques(RelatorioMensal anterior, RelatorioMensal atual)
        {
            List<string> linhas = new List<string>();

            var evolucoes = atual.FinalizadosPorAtendente
                .Where(x => anterior.FinalizadosPorAtendente.TryGetValue(x.Key, out int ant) && ant > 0)
                .Select(x => new
                {
                    Nome = x.Key,
                    Var = ((x.Value - anterior.FinalizadosPorAtendente[x.Key]) / (double)anterior.FinalizadosPorAtendente[x.Key]) * 100d
                })
                .OrderByDescending(x => x.Var)
                .FirstOrDefault();

            if (evolucoes != null && evolucoes.Var > 0)
            {
                linhas.Add("Maior evolução de volume");
                linhas.Add($"{evolucoes.Nome}  ▲ {evolucoes.Var.ToString("0.0", PtBr)}%");
                linhas.Add(string.Empty);
            }

            var melhorNota = atual.Notas
                .Where(x => x.Value.Votos >= 10)
                .OrderByDescending(x => x.Value.Media)
                .ThenByDescending(x => x.Value.Votos)
                .FirstOrDefault();

            if (!string.IsNullOrWhiteSpace(melhorNota.Key))
            {
                linhas.Add("Melhor nota (mín. 10 avaliações)");
                linhas.Add($"{melhorNota.Key}  ★ {melhorNota.Value.Media.ToString("0.00", PtBr)}");
                linhas.Add(string.Empty);
            }

            var maiorVolume = atual.FinalizadosPorAtendente.OrderByDescending(x => x.Value).FirstOrDefault();
            if (!string.IsNullOrWhiteSpace(maiorVolume.Key))
            {
                linhas.Add("Maior volume no mês");
                linhas.Add($"{maiorVolume.Key}  {maiorVolume.Value.ToString("N0", PtBr)} atendimentos");
            }

            return string.Join(Environment.NewLine, linhas);
        }

        private string GerarIntegridade(RelatorioMensal anterior, RelatorioMensal atual)
        {
            List<string> linhas = new List<string>();

            if (!atual.TmaDisponivel)
                linhas.Add("⚠ TMA não disponível no mês atual.");

            foreach (var item in atual.PossiveisVendas)
            {
                if (atual.Motivos.TryGetValue(item.Key, out int principal) && principal != item.Value)
                {
                    linhas.Add($"⚠ Divergência em '{item.Key}': {principal.ToString("N0", PtBr)} x {item.Value.ToString("N0", PtBr)}.");
                }
            }

            if (atual.AtendimentosRelacao.HasValue && atual.AtendimentosRelacao.Value != atual.Finalizados)
            {
                linhas.Add($"⚠ Total diário ({atual.AtendimentosRelacao.Value.ToString("N0", PtBr)}) difere do total finalizado ({atual.Finalizados.ToString("N0", PtBr)}).");
            }

            if (linhas.Count == 0)
                linhas.Add("✓ Não foram encontradas divergências relevantes nos dados utilizados.");

            return string.Join(Environment.NewLine + Environment.NewLine, linhas.Take(5));
        }

        private static string GerarDiagnosticoAtendente(
            int anterior,
            int atual,
            NotaAtendente? notaAnterior,
            NotaAtendente? notaAtual,
            TimeSpan? tmrAnterior,
            TimeSpan? tmrAtual)
        {
            double? varVolume = anterior > 0 ? ((atual - anterior) / (double)anterior) * 100d : null;
            double? deltaNota = notaAnterior != null && notaAtual != null ? notaAtual.Media - notaAnterior.Media : null;
            TimeSpan? deltaTmr = tmrAnterior.HasValue && tmrAtual.HasValue ? tmrAtual.Value - tmrAnterior.Value : null;

            if (anterior == 0 && atual > 0)
                return "Novo no comparativo";

            if (varVolume >= 15 && deltaTmr.HasValue && deltaTmr.Value <= TimeSpan.FromMinutes(5) && (!deltaNota.HasValue || deltaNota.Value >= -0.10))
                return "Boa evolução";

            if (varVolume >= 15 && deltaTmr.HasValue && deltaTmr.Value > TimeSpan.FromMinutes(10))
                return "Volume ↑ / TMR piorou";

            if (varVolume <= -20 && ((deltaNota.HasValue && deltaNota.Value > 0) || (deltaTmr.HasValue && deltaTmr.Value < TimeSpan.Zero)))
                return "Volume ↓ / qualidade melhorou";

            if (deltaTmr.HasValue && deltaTmr.Value > TimeSpan.FromMinutes(15))
                return "Atenção ao TMR";

            if (deltaNota.HasValue && deltaNota.Value <= -0.20)
                return "Queda na nota";

            if (varVolume.HasValue && Math.Abs(varVolume.Value) <= 10)
                return "Estável";

            if (varVolume < -20)
                return "Queda de volume: verificar escala";

            return "Acompanhar evolução";
        }

        // =========================================================
        // LIMPAR
        // =========================================================
        private void LimparDados(bool atualizarStatus = true)
        {
            caminhoMesRetrasado = string.Empty;
            caminhoMesPassado = string.Empty;
            relatorioAnteriorAtual = null;
            relatorioAtualAtual = null;

            if (txtMesRetrasado != null)
                txtMesRetrasado.Text = "Nenhum arquivo selecionado";

            if (txtMesPassado != null)
                txtMesPassado.Text = "Nenhum arquivo selecionado";

            if (lblPeriodo1 != null)
                lblPeriodo1.Text = string.Empty;

            if (lblPeriodo2 != null)
                lblPeriodo2.Text = string.Empty;

            LimparResultados();

            if (atualizarStatus && lblStatus != null)
                lblStatus.Text = "●  Dados limpos. Selecione os dois arquivos novamente.";
            else if (lblStatus != null)
                lblStatus.Text = "●  Aguardando os dois arquivos.";
        }

        private void LimparResultados()
        {
            LimparKpi(kpiConversas);
            LimparKpi(kpiFinalizados);
            LimparKpi(kpiTaxaFinalizacao);
            LimparKpi(kpiReagendamentos);
            LimparKpi(kpiNovosContatos);
            LimparKpi(kpiNotaMedia);

            if (lblResumoTexto != null) lblResumoTexto.Text = string.Empty;
            if (lblAlertasTexto != null) lblAlertasTexto.Text = string.Empty;
            if (lblOportunidadesTexto != null) lblOportunidadesTexto.Text = string.Empty;
            if (lblAcoesTexto != null) lblAcoesTexto.Text = string.Empty;
            if (lblDestaquesTexto != null) lblDestaquesTexto.Text = string.Empty;
            if (lblIntegridadeTexto != null) lblIntegridadeTexto.Text = string.Empty;

            if (dgvAtendentes != null)
            {
                dgvAtendentes.Rows.Clear();
                dgvAtendentes.Columns["Anterior"].HeaderText = "Mês retrasado";
                dgvAtendentes.Columns["Atual"].HeaderText = "Mês passado";
            }

            if (graficoMotivos != null)
            {
                graficoMotivos.Categories = Array.Empty<string>();
                graficoMotivos.PreviousValues = Array.Empty<double>();
                graficoMotivos.CurrentValues = Array.Empty<double>();
                graficoMotivos.PreviousLabel = "Mês retrasado";
                graficoMotivos.CurrentLabel = "Mês passado";
                graficoMotivos.Invalidate();
            }
        }

        private static void LimparKpi(KpiView? kpi)
        {
            if (kpi == null)
                return;

            kpi.Valor.Text = string.Empty;
            kpi.Variacao.Text = string.Empty;
            kpi.Anterior.Text = string.Empty;
        }

        // =========================================================
        // HELPERS
        // =========================================================
        private static double CalcularTaxa(int parte, int total)
        {
            return total <= 0 ? 0 : (parte / (double)total) * 100d;
        }

        private static double? CalcularVariacaoPercentual(int anterior, int atual)
        {
            if (anterior == 0)
                return null;

            return ((atual - anterior) / (double)anterior) * 100d;
        }

        private static int ValorMotivo(RelatorioMensal r, string motivo)
        {
            return r.Motivos.TryGetValue(motivo, out int valor) ? valor : 0;
        }

        private static int ContarTmrComPiora(RelatorioMensal anterior, RelatorioMensal atual, TimeSpan limite)
        {
            int total = 0;

            foreach (var item in atual.TMR)
            {
                if (anterior.TMR.TryGetValue(item.Key, out TimeSpan ant) && item.Value - ant > limite)
                    total++;
            }

            return total;
        }

        private static string AbreviarMotivo(string motivo)
        {
            string norm = Normalizar(motivo);

            if (norm.Contains("atendimento ja realizado")) return "Já realizado";
            if (norm.Contains("pedido de formula")) return "Pedido fórmula";
            if (norm.Contains("alteracao no pedido")) return "Alteração";
            if (norm.Contains("encerrado por inatividade")) return "Inatividade";
            if (norm.Contains("orcamento de formula")) return "Orçamento";
            if (norm.Contains("duvida de formula")) return "Dúvida fórmula";
            if (norm.Contains("avaliacao")) return "Avaliação";
            if (norm.Contains("entrega")) return "Entrega";
            if (norm.Contains("repeticao")) return "Repetição";

            return motivo.Length <= 14 ? motivo : motivo[..14] + "…";
        }

        private static string FormatarTempo(TimeSpan tempo)
        {
            if (tempo.TotalDays >= 1)
                return $"{(int)tempo.TotalDays}d {tempo.Hours:00}h {tempo.Minutes:00}m";

            if (tempo.TotalHours >= 1)
                return $"{(int)tempo.TotalHours}h {tempo.Minutes:00}m";

            return $"{tempo.Minutes}m {tempo.Seconds:00}s";
        }

        private static string Normalizar(string? texto)
        {
            if (string.IsNullOrWhiteSpace(texto))
                return string.Empty;

            string formD = texto.Trim().ToLowerInvariant().Normalize(NormalizationForm.FormD);
            StringBuilder sb = new StringBuilder();

            foreach (char c in formD)
            {
                UnicodeCategory categoria = CharUnicodeInfo.GetUnicodeCategory(c);
                if (categoria != UnicodeCategory.NonSpacingMark)
                    sb.Append(c);
            }

            return Regex.Replace(sb.ToString().Normalize(NormalizationForm.FormC), @"\s+", " ").Trim();
        }
    }

    // =============================================================
    // MODELOS
    // =============================================================
    public sealed class KpiView
    {
        public Label Valor { get; }
        public Label Variacao { get; }
        public Label Anterior { get; }

        public KpiView(Label valor, Label variacao, Label anterior)
        {
            Valor = valor;
            Variacao = variacao;
            Anterior = anterior;
        }
    }

    public sealed class NotaAtendente
    {
        public double Media { get; set; }
        public int Votos { get; set; }
    }

    public sealed class RelatorioMensal
    {
        public DateTime DataReferencia { get; set; }
        public string Periodo { get; set; } = string.Empty;

        public int Conversas { get; set; }
        public int Finalizados { get; set; }
        public int? AtendimentosRelacao { get; set; }
        public int Reagendamentos { get; set; }
        public int NovosContatos { get; set; }
        public int ContatosMarketing { get; set; }
        public int MensagensTemplate { get; set; }
        public double NotaMedia { get; set; }
        public bool TmaDisponivel { get; set; }

        public Dictionary<string, int> Motivos { get; set; } = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, int> PossiveisVendas { get; set; } = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, int> MarketingCriativos { get; set; } = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, int> FinalizadosPorAtendente { get; set; } = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, TimeSpan> TMR { get; set; } = new Dictionary<string, TimeSpan>(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, TimeSpan> TME { get; set; } = new Dictionary<string, TimeSpan>(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, TimeSpan> TMA { get; set; } = new Dictionary<string, TimeSpan>(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, NotaAtendente> Notas { get; set; } = new Dictionary<string, NotaAtendente>(StringComparer.OrdinalIgnoreCase);
    }

    internal static class StringExtensions
    {
        public static string FirstCharToUpper(this string texto)
        {
            if (string.IsNullOrWhiteSpace(texto))
                return texto;

            return char.ToUpper(texto[0], PtCulture) + texto[1..];
        }

        private static readonly CultureInfo PtCulture = CultureInfo.GetCultureInfo("pt-BR");
    }

    // =============================================================
    // PAINEL ARREDONDADO
    // =============================================================
    public class RoundedPanel : Panel
    {
        private int _radius = 10;
        private Color _borderColor = Color.LightGray;
        private int _borderWidth = 1;

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int Radius
        {
            get => _radius;
            set
            {
                _radius = Math.Max(0, value);
                AtualizarRegiao();
                Invalidate();
            }
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Color BorderColor
        {
            get => _borderColor;
            set
            {
                _borderColor = value;
                Invalidate();
            }
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int BorderWidth
        {
            get => _borderWidth;
            set
            {
                _borderWidth = Math.Max(0, value);
                Invalidate();
            }
        }

        public RoundedPanel()
        {
            SetStyle(
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.UserPaint |
                ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.ResizeRedraw,
                true);
        }

        protected override void OnResize(EventArgs eventargs)
        {
            base.OnResize(eventargs);
            AtualizarRegiao();
            Invalidate();
        }

        private void AtualizarRegiao()
        {
            if (Width <= 0 || Height <= 0)
                return;

            Rectangle rect = new Rectangle(0, 0, Width, Height);
            int radius = Math.Min(Radius, Math.Min(Width, Height) / 2);

            using GraphicsPath path = CriarPath(rect, radius);
            Region? antiga = Region;
            Region = new Region(path);
            antiga?.Dispose();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            Rectangle rect = new Rectangle(0, 0, Math.Max(1, Width - 1), Math.Max(1, Height - 1));
            int radius = Math.Min(Radius, Math.Min(rect.Width, rect.Height) / 2);

            using GraphicsPath path = CriarPath(rect, radius);
            using SolidBrush brush = new SolidBrush(BackColor);
            e.Graphics.FillPath(brush, path);

            if (BorderWidth > 0)
            {
                using Pen pen = new Pen(BorderColor, BorderWidth);
                e.Graphics.DrawPath(pen, path);
            }

            base.OnPaint(e);
        }

        private static GraphicsPath CriarPath(Rectangle rect, int radius)
        {
            GraphicsPath path = new GraphicsPath();

            if (radius <= 0)
            {
                path.AddRectangle(rect);
                path.CloseFigure();
                return path;
            }

            int diameter = radius * 2;
            Rectangle arc = new Rectangle(rect.Left, rect.Top, diameter, diameter);

            path.AddArc(arc, 180, 90);
            arc.X = rect.Right - diameter;
            path.AddArc(arc, 270, 90);
            arc.Y = rect.Bottom - diameter;
            path.AddArc(arc, 0, 90);
            arc.X = rect.Left;
            path.AddArc(arc, 90, 90);
            path.CloseFigure();

            return path;
        }
    }

    // =============================================================
    // GRÁFICO DE BARRAS
    // =============================================================
    public class ComparisonBarChart : Control
    {
        private string[] _categories = Array.Empty<string>();
        private double[] _previousValues = Array.Empty<double>();
        private double[] _currentValues = Array.Empty<double>();
        private string _previousLabel = "Mês retrasado";
        private string _currentLabel = "Mês passado";

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string[] Categories
        {
            get => _categories;
            set { _categories = value ?? Array.Empty<string>(); Invalidate(); }
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public double[] PreviousValues
        {
            get => _previousValues;
            set { _previousValues = value ?? Array.Empty<double>(); Invalidate(); }
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public double[] CurrentValues
        {
            get => _currentValues;
            set { _currentValues = value ?? Array.Empty<double>(); Invalidate(); }
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string PreviousLabel
        {
            get => _previousLabel;
            set { _previousLabel = value ?? string.Empty; Invalidate(); }
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string CurrentLabel
        {
            get => _currentLabel;
            set { _currentLabel = value ?? string.Empty; Invalidate(); }
        }

        public ComparisonBarChart()
        {
            SetStyle(
                ControlStyles.UserPaint |
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.ResizeRedraw,
                true);

            BackColor = Color.White;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            if (Categories.Length == 0 ||
                PreviousValues.Length != Categories.Length ||
                CurrentValues.Length != Categories.Length ||
                Width < 100 || Height < 100)
            {
                return;
            }

            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            int esquerda = 48;
            int direita = 15;
            int topo = 42;
            int baixo = 52;

            Rectangle area = new Rectangle(
                esquerda,
                topo,
                Math.Max(10, Width - esquerda - direita),
                Math.Max(10, Height - topo - baixo));

            double max = Math.Max(PreviousValues.Max(), CurrentValues.Max());
            if (max <= 0) max = 1;
            max *= 1.15;

            using Pen gridPen = new Pen(Color.FromArgb(230, 234, 238), 1F);
            using Font fonteLegenda = new Font("Segoe UI", 7.2F);
            using Font fonteValor = new Font("Segoe UI Semibold", 7F);
            using Font fonteCategoria = new Font("Segoe UI", 6.8F);
            using SolidBrush brushAnterior = new SolidBrush(Color.FromArgb(39, 111, 86));
            using SolidBrush brushAtual = new SolidBrush(Color.FromArgb(111, 181, 137));
            using SolidBrush brushTexto = new SolidBrush(Color.FromArgb(60, 70, 80));

            for (int i = 0; i <= 4; i++)
            {
                float y = area.Top + area.Height / 4F * i;
                g.DrawLine(gridPen, area.Left, y, area.Right, y);

                double valor = max - max / 4D * i;
                g.DrawString(((int)valor).ToString("N0", CultureInfo.GetCultureInfo("pt-BR")),
                    fonteCategoria, brushTexto, 2F, y - 7F);
            }

            g.FillRectangle(brushAnterior, area.Left, 13, 10, 10);
            g.DrawString(PreviousLabel, fonteLegenda, brushTexto, area.Left + 15, 10);

            g.FillRectangle(brushAtual, area.Left + 125, 13, 10, 10);
            g.DrawString(CurrentLabel, fonteLegenda, brushTexto, area.Left + 140, 10);

            float grupo = area.Width / (float)Categories.Length;
            float larguraBarra = Math.Max(8F, grupo * 0.27F);

            for (int i = 0; i < Categories.Length; i++)
            {
                float centro = area.Left + grupo * i + grupo / 2F;
                float alturaAnterior = (float)(PreviousValues[i] / max * area.Height);
                float alturaAtual = (float)(CurrentValues[i] / max * area.Height);

                RectangleF anterior = new RectangleF(
                    centro - larguraBarra - 2F,
                    area.Bottom - alturaAnterior,
                    larguraBarra,
                    alturaAnterior);

                RectangleF atual = new RectangleF(
                    centro + 2F,
                    area.Bottom - alturaAtual,
                    larguraBarra,
                    alturaAtual);

                g.FillRectangle(brushAnterior, anterior);
                g.FillRectangle(brushAtual, atual);

                string txtAnterior = PreviousValues[i].ToString("N0", CultureInfo.GetCultureInfo("pt-BR"));
                string txtAtual = CurrentValues[i].ToString("N0", CultureInfo.GetCultureInfo("pt-BR"));

                SizeF sizeAnt = g.MeasureString(txtAnterior, fonteValor);
                SizeF sizeAtu = g.MeasureString(txtAtual, fonteValor);

                g.DrawString(txtAnterior, fonteValor, brushTexto,
                    anterior.X + anterior.Width / 2F - sizeAnt.Width / 2F,
                    anterior.Y - 16F);

                g.DrawString(txtAtual, fonteValor, brushTexto,
                    atual.X + atual.Width / 2F - sizeAtu.Width / 2F,
                    atual.Y - 16F);

                string categoria = Categories[i];
                SizeF sizeCat = g.MeasureString(categoria, fonteCategoria);
                g.DrawString(categoria, fonteCategoria, brushTexto,
                    centro - sizeCat.Width / 2F,
                    area.Bottom + 8F);
            }
        }
    }
}
