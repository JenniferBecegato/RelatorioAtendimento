#nullable disable
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using Rectangle = System.Drawing.Rectangle;
using Graphics = System.Drawing.Graphics;
using Bitmap = System.Drawing.Bitmap;
using Pen = System.Drawing.Pen;

namespace RelatorioAtendimento
{
    public partial class Form1 : Form
    {
        // ============================================================
        // CORES
        // ============================================================
        private readonly Color CorFundo = Color.FromArgb(232, 236, 240);
        private readonly Color CorSidebar = Color.FromArgb(9, 76, 64);
        private readonly Color CorSidebarAtivo = Color.FromArgb(16, 116, 94);
        private readonly Color CorVerde = Color.FromArgb(20, 150, 96);
        private readonly Color CorVerdeClaro = Color.FromArgb(234, 248, 240);
        private readonly Color CorVermelho = Color.FromArgb(214, 55, 55);
        private readonly Color CorVermelhoClaro = Color.FromArgb(253, 239, 239);
        private readonly Color CorAzul = Color.FromArgb(37, 99, 235);
        private readonly Color CorAzulClaro = Color.FromArgb(238, 244, 255);
        private readonly Color CorAmarelo = Color.FromArgb(194, 132, 18);
        private readonly Color CorAmareloClaro = Color.FromArgb(255, 249, 229);
        private readonly Color CorTexto = Color.FromArgb(20, 35, 50);
        private readonly Color CorTextoSecundario = Color.FromArgb(92, 105, 120);
        private readonly Color CorBorda = Color.FromArgb(221, 227, 233);

        // ============================================================
        // ESTADO
        // ============================================================
        private string caminhoAnterior = "";
        private string caminhoAtual = "";
        private DadosMes dadosAnterior;
        private DadosMes dadosAtual;
        private bool comparacaoRealizada = false;

        // ============================================================
        // CONTROLES PRINCIPAIS
        // ============================================================
        private Panel pnlConteudo;
        private TextBox txtAnterior;
        private TextBox txtAtual;
        private Label lblPeriodoAnterior;
        private Label lblPeriodoAtual;
        private Label lblStatus;
        private FlowLayoutPanel pnlMenu;
        private readonly Dictionary<string, Button> botoesMenu = new();

        public Form1()
        {
            InitializeComponent();
            MontarTela();
        }

        // ============================================================
        // TELA BASE
        // ============================================================
        private void MontarTela()
        {
            SuspendLayout();
            Controls.Clear();

            Text = "Relatório de Atendimento - Comparativo Mensal";
            StartPosition = FormStartPosition.CenterScreen;
            WindowState = FormWindowState.Maximized;
            MinimumSize = new Size(1180, 720);
            BackColor = CorFundo;
            Font = new Font("Segoe UI", 9F);
            AutoScaleMode = AutoScaleMode.Dpi;

            // Estrutura simples e estável: menu fixo à esquerda e conteúdo ocupando o restante.
            var estrutura = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = CorFundo
            };

            var areaDireita = CriarAreaDireita();
            areaDireita.Dock = DockStyle.Fill;

            var sidebar = CriarSidebar();
            sidebar.Dock = DockStyle.Left;
            sidebar.Width = 210;

            // A ordem garante que o painel esquerdo reserve os 210px e o conteúdo use o restante.
            estrutura.Controls.Add(areaDireita);
            estrutura.Controls.Add(sidebar);

            Controls.Add(estrutura);

            MostrarPagina("Comparativo");
            ResumeLayout(true);
        }

        private Control CriarSidebar()
        {
            var sidebar = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = CorSidebar,
                ColumnCount = 1,
                RowCount = 3,
                Margin = Padding.Empty,
                Padding = Padding.Empty
            };

            sidebar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            sidebar.RowStyles.Add(new RowStyle(SizeType.Absolute, 112F));
            sidebar.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            sidebar.RowStyles.Add(new RowStyle(SizeType.Absolute, 145F));

            var titulo = new Label
            {
                Dock = DockStyle.Fill,
                Text = "RELATÓRIO\nDE ATENDIMENTO",
                ForeColor = Color.White,
                Font = new Font("Segoe UI Semibold", 13F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Margin = Padding.Empty
            };

            pnlMenu = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = false,
                BackColor = CorSidebar,
                Padding = new Padding(0, 4, 0, 0),
                Margin = Padding.Empty
            };

            AdicionarBotaoMenu("Comparativo", "▣");
            AdicionarBotaoMenu("Atendentes", "●");
            AdicionarBotaoMenu("Motivos", "◆");
            AdicionarBotaoMenu("Marketing", "▲");
            AdicionarBotaoMenu("Qualidade", "✓");
            AdicionarBotaoMenu("Exportar", "⇩");

            var dica = new Label
            {
                Dock = DockStyle.Fill,
                Text = "VISÃO RÁPIDA\nna tela principal\n\nDETALHES\nno menu acima",
                ForeColor = Color.FromArgb(210, 235, 229),
                Font = new Font("Segoe UI", 9.2F),
                TextAlign = ContentAlignment.MiddleCenter,
                Margin = Padding.Empty
            };

            sidebar.Controls.Add(titulo, 0, 0);
            sidebar.Controls.Add(pnlMenu, 0, 1);
            sidebar.Controls.Add(dica, 0, 2);

            return sidebar;
        }

        private void AdicionarBotaoMenu(string nome, string icone)
        {
            var btn = new Button
            {
                Name = "btn" + nome,
                Text = $"{icone}   {nome}",
                Width = 210,
                Height = 50,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 0 },
                BackColor = CorSidebar,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10F),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(22, 0, 0, 0),
                Cursor = Cursors.Hand,
                Tag = nome
            };

            btn.Click += (_, _) => MostrarPagina(nome);
            btn.MouseEnter += (_, _) =>
            {
                if (btn.BackColor != CorSidebarAtivo)
                    btn.BackColor = Color.FromArgb(12, 91, 76);
            };
            btn.MouseLeave += (_, _) =>
            {
                if ((string)btn.Tag != PaginaAtual)
                    btn.BackColor = CorSidebar;
            };

            botoesMenu[nome] = btn;
            pnlMenu.Controls.Add(btn);
        }

        private string PaginaAtual { get; set; } = "Comparativo";

        private Control CriarAreaDireita()
        {
            var baseDireita = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                BackColor = CorFundo,
                Padding = new Padding(18, 14, 18, 12)
            };

            baseDireita.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            baseDireita.RowStyles.Add(new RowStyle(SizeType.Absolute, 74));
            baseDireita.RowStyles.Add(new RowStyle(SizeType.Absolute, 94));
            baseDireita.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            baseDireita.Controls.Add(CriarCabecalho(), 0, 0);
            baseDireita.Controls.Add(CriarImportacao(), 0, 1);

            pnlConteudo = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = CorFundo
            };
            baseDireita.Controls.Add(pnlConteudo, 0, 2);

            return baseDireita;
        }

        private Control CriarCabecalho()
        {
            var p = new Panel { Dock = DockStyle.Fill };

            var titulo = new Label
            {
                AutoSize = true,
                Location = new Point(4, 2),
                Text = "Relatório de Atendimento - Comparativo Mensal",
                Font = new Font("Segoe UI Semibold", 20F, FontStyle.Bold),
                ForeColor = CorTexto
            };

            var subtitulo = new Label
            {
                AutoSize = true,
                Location = new Point(7, 42),
                Text = "Simples: veja quem melhorou, quem precisa de atenção e o que mudou no mês.",
                Font = new Font("Segoe UI", 10F),
                ForeColor = CorTextoSecundario
            };

            var versao = new Label
            {
                Width = 190,
                Height = 42,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                TextAlign = ContentAlignment.TopRight,
                Text = $"Versão 2.0\n{DateTime.Now:dd/MM/yyyy HH:mm}",
                Font = new Font("Segoe UI", 7.7F),
                ForeColor = CorTextoSecundario
            };

            p.Controls.Add(titulo);
            p.Controls.Add(subtitulo);
            p.Controls.Add(versao);

            p.Resize += (_, _) => versao.Location = new Point(Math.Max(0, p.ClientSize.Width - 195), 3);

            return p;
        }

        private Control CriarImportacao()
        {
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                BackColor = CorFundo,
                Padding = new Padding(0, 3, 0, 6),
                Margin = Padding.Empty
            };

            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 43.5F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 43.5F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 13F));

            var cardAnterior = CriarCardArquivo(
                "1. Mês retrasado",
                out txtAnterior,
                out lblPeriodoAnterior,
                SelecionarAnterior);

            var cardAtual = CriarCardArquivo(
                "2. Mês passado",
                out txtAtual,
                out lblPeriodoAtual,
                SelecionarAtual);

            layout.Controls.Add(cardAnterior, 0, 0);
            layout.Controls.Add(cardAtual, 1, 0);

            // Terceiro bloco: botão Comparar dentro de um card com contorno.
            var btnComparar = new Button
            {
                Dock = DockStyle.Fill,
                Text = "Comparar",
                BackColor = CorVerde,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI Semibold", 10.5F, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Margin = new Padding(8, 10, 2, 10)
            };

            btnComparar.FlatAppearance.BorderSize = 0;
            btnComparar.Click += BtnComparar_Click;

            layout.Controls.Add(btnComparar, 2, 0);

            return layout;
        }

        private Control CriarCardArquivo(
            string titulo,
            out TextBox txt,
            out Label periodo,
            EventHandler evento)
        {
            var card = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Margin = new Padding(0, 3, 4, 3),
                Padding = new Padding(11, 7, 11, 7)
            };

            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Margin = Padding.Empty,
                Padding = Padding.Empty
            };

            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 19F));
            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 29F));
            grid.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            var lbl = new Label
            {
                Dock = DockStyle.Fill,
                Text = titulo,
                Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold),
                ForeColor = CorTexto,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = Padding.Empty
            };

            var linha = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Margin = Padding.Empty,
                Padding = Padding.Empty
            };

            linha.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            linha.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 96F));

            txt = new TextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Text = "Nenhum arquivo selecionado",
                Font = new Font("Segoe UI", 8.4F),
                Margin = new Padding(0, 2, 0, 2)
            };

            var btn = new Button
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(7, 2, 0, 2),
                Text = "Selecionar",
                BackColor = Color.FromArgb(241, 248, 245),
                ForeColor = CorSidebar,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Font = new Font("Segoe UI", 8.3F)
            };

            btn.FlatAppearance.BorderSize = 0;
            btn.Click += evento;

            linha.Controls.Add(txt, 0, 0);
            linha.Controls.Add(btn, 1, 0);

            periodo = new Label
            {
                Dock = DockStyle.Fill,
                Text = "Período: aguardando arquivo...",
                ForeColor = CorTextoSecundario,
                Font = new Font("Segoe UI", 7.7F),
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = Padding.Empty
            };

            grid.Controls.Add(lbl, 0, 0);
            grid.Controls.Add(linha, 0, 1);
            grid.Controls.Add(periodo, 0, 2);

            card.Controls.Add(grid);
            return card;
        }

        // ============================================================
        // NAVEGAÇÃO
        // ============================================================
        private void MostrarPagina(string pagina)
        {
            PaginaAtual = pagina;

            foreach (var par in botoesMenu)
                par.Value.BackColor = par.Key == pagina ? CorSidebarAtivo : CorSidebar;

            pnlConteudo.Controls.Clear();

            Control paginaControl = pagina switch
            {
                "Comparativo" => CriarPaginaComparativo(),
                "Atendentes" => CriarPaginaAtendentes(),
                "Motivos" => CriarPaginaMotivos(),
                "Marketing" => CriarPaginaMarketing(),
                "Qualidade" => CriarPaginaQualidade(),
                "Exportar" => CriarPaginaExportar(),
                _ => CriarPaginaComparativo()
            };

            paginaControl.Dock = DockStyle.Fill;
            pnlConteudo.Controls.Add(paginaControl);
        }

        // ============================================================
        // PÁGINA PRINCIPAL - SIMPLES
        // ============================================================
        private Control CriarPaginaComparativo()
        {
            var corpo = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                BackColor = CorFundo,
                Padding = new Padding(0, 8, 0, 4),
                Margin = Padding.Empty
            };

            corpo.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            corpo.RowStyles.Add(new RowStyle(SizeType.Absolute, 96F));   // KPIs
            corpo.RowStyles.Add(new RowStyle(SizeType.Absolute, 116F));  // TMR
            corpo.RowStyles.Add(new RowStyle(SizeType.Absolute, 116F));  // TME
            corpo.RowStyles.Add(new RowStyle(SizeType.Absolute, 116F));  // Atendimentos + Notas
            corpo.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));   // Resumo
            corpo.RowStyles.Add(new RowStyle(SizeType.Absolute, 58F));   // Rodapé

            corpo.Controls.Add(CriarLinhaKpisSimples(), 0, 0);
            corpo.Controls.Add(CriarLinhaMudancasTmr(), 0, 1);
            corpo.Controls.Add(CriarLinhaMudancasTme(), 0, 2);
            corpo.Controls.Add(CriarLinhaEquipe(), 0, 3);
            corpo.Controls.Add(CriarResumoSimples(), 0, 4);
            corpo.Controls.Add(CriarBarraInferior(), 0, 5);

            return corpo;
        }

        private Control CriarLinhaKpisSimples()
        {
            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 1
            };
            for (int i = 0; i < 4; i++)
                grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));

            if (!comparacaoRealizada)
            {
                grid.Controls.Add(CriarKpiVazio("Conversas"), 0, 0);
                grid.Controls.Add(CriarKpiVazio("Finalizados"), 1, 0);
                grid.Controls.Add(CriarKpiVazio("Novos contatos"), 2, 0);
                grid.Controls.Add(CriarKpiVazio("Nota geral"), 3, 0);
                return grid;
            }

            grid.Controls.Add(CriarKpiComparacao(
                "Conversas",
                dadosAnterior.Conversas,
                dadosAtual.Conversas,
                false), 0, 0);

            grid.Controls.Add(CriarKpiComparacao(
                "Finalizados",
                dadosAnterior.Finalizados,
                dadosAtual.Finalizados,
                false), 1, 0);

            grid.Controls.Add(CriarKpiComparacao(
                "Novos contatos",
                dadosAnterior.NovosContatos,
                dadosAtual.NovosContatos,
                false), 2, 0);

            grid.Controls.Add(CriarKpiNota(), 3, 0);
            return grid;
        }

        private Control CriarKpiVazio(string titulo)
        {
            var card = CriarCard(Color.White);
            card.Margin = new Padding(5);

            var lbl = new Label
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(14, 12, 10, 8),
                Text = $"{titulo}\n\n—",
                Font = new Font("Segoe UI Semibold", 10F),
                ForeColor = CorTexto
            };
            card.Controls.Add(lbl);
            return card;
        }

        private Control CriarKpiComparacao(string titulo, double anterior, double atual, bool percentual)
        {
            double variacao = CalcularPercentual(anterior, atual);
            Color cor = variacao >= 0 ? CorVerde : CorVermelho;
            string seta = variacao >= 0 ? "▲" : "▼";

            var card = CriarCard(Color.White);
            card.Margin = new Padding(5);

            string atualTexto = percentual
                ? $"{atual:N1}%"
                : $"{atual:N0}";

            string antTexto = percentual
                ? $"{anterior:N1}%"
                : $"{anterior:N0}";

            var tituloLbl = new Label
            {
                AutoSize = true,
                Location = new Point(14, 10),
                Text = titulo,
                Font = new Font("Segoe UI", 9F),
                ForeColor = CorTextoSecundario
            };

            var valor = new Label
            {
                AutoSize = true,
                Location = new Point(14, 32),
                Text = atualTexto,
                Font = new Font("Segoe UI Semibold", 15F, FontStyle.Bold),
                ForeColor = CorTexto
            };

            var varLbl = new Label
            {
                AutoSize = true,
                Location = new Point(14, 65),
                Text = $"{seta} {Math.Abs(variacao):N1}%   (antes {antTexto})",
                Font = new Font("Segoe UI Semibold", 8.8F, FontStyle.Bold),
                ForeColor = cor
            };

            card.Controls.Add(tituloLbl);
            card.Controls.Add(valor);
            card.Controls.Add(varLbl);
            return card;
        }

        private Control CriarKpiNota()
        {
            double dif = dadosAtual.NotaGeral - dadosAnterior.NotaGeral;
            Color cor = Math.Abs(dif) < 0.01 ? CorTextoSecundario : dif > 0 ? CorVerde : CorVermelho;
            string seta = Math.Abs(dif) < 0.01 ? "●" : dif > 0 ? "▲" : "▼";

            var card = CriarCard(Color.White);
            card.Margin = new Padding(5);

            var t = new Label
            {
                AutoSize = true,
                Location = new Point(14, 10),
                Text = "Nota geral",
                Font = new Font("Segoe UI", 9F),
                ForeColor = CorTextoSecundario
            };

            var v = new Label
            {
                AutoSize = true,
                Location = new Point(14, 32),
                Text = dadosAtual.NotaGeral > 0 ? dadosAtual.NotaGeral.ToString("N2") : "—",
                Font = new Font("Segoe UI Semibold", 15F, FontStyle.Bold),
                ForeColor = CorTexto
            };

            var d = new Label
            {
                AutoSize = true,
                Location = new Point(14, 65),
                Text = $"{seta} {Math.Abs(dif):N2}   (antes {dadosAnterior.NotaGeral:N2})",
                Font = new Font("Segoe UI Semibold", 8.8F, FontStyle.Bold),
                ForeColor = cor
            };

            card.Controls.Add(t);
            card.Controls.Add(v);
            card.Controls.Add(d);
            return card;
        }

        private Control CriarLinhaMudancasTmr()
        {
            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1
            };

            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            if (!comparacaoRealizada)
            {
                grid.Controls.Add(
                    CriarCardFrase(
                        "⏱  TMR PIOROU",
                        "Aguardando comparação.",
                        CorVermelhoClaro,
                        CorVermelho),
                    0, 0);

                grid.Controls.Add(
                    CriarCardFrase(
                        "✓  TMR MELHOROU",
                        "Aguardando comparação.",
                        CorVerdeClaro,
                        CorVerde),
                    1, 0);

                return grid;
            }

            var mudancas = CompararTempos(
                dadosAnterior.TMR,
                dadosAtual.TMR,
                60);

            var pioraram = mudancas
                .Where(x => x.DiferencaSegundos > 0)
                .OrderByDescending(x => x.DiferencaSegundos)

                .Select(x => NomeCurto(x.Nome))
                .ToList();

            var melhoraram = mudancas
                .Where(x => x.DiferencaSegundos < 0)
                .OrderBy(x => x.DiferencaSegundos)

                .Select(x => NomeCurto(x.Nome))
                .ToList();

            string textoPiorou = pioraram.Count == 0
                ? "Nenhuma piora relevante no TMR."
                : $"Pioraram: {JuntarNomes(pioraram)}.";

            string textoMelhorou = melhoraram.Count == 0
                ? "Nenhuma melhora relevante no TMR."
                : $"Melhoraram: {JuntarNomes(melhoraram)}.";

            grid.Controls.Add(
                CriarCardFrase(
                    "⏱  TMR PIOROU",
                    textoPiorou,
                    CorVermelhoClaro,
                    CorVermelho),
                0, 0);

            grid.Controls.Add(
                CriarCardFrase(
                    "✓  TMR MELHOROU",
                    textoMelhorou,
                    CorVerdeClaro,
                    CorVerde),
                1, 0);

            return grid;
        }

        private Control CriarLinhaMudancasTme()
        {
            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1
            };

            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            if (!comparacaoRealizada)
            {
                grid.Controls.Add(
                    CriarCardFrase(
                        "⏳  TME PIOROU",
                        "Aguardando comparação.",
                        CorVermelhoClaro,
                        CorVermelho),
                    0, 0);

                grid.Controls.Add(
                    CriarCardFrase(
                        "✓  TME MELHOROU",
                        "Aguardando comparação.",
                        CorVerdeClaro,
                        CorVerde),
                    1, 0);

                return grid;
            }

            var mudancas = CompararTempos(
                dadosAnterior.TME,
                dadosAtual.TME,
                60);

            var pioraram = mudancas
                .Where(x => x.DiferencaSegundos > 0)
                .OrderByDescending(x => x.DiferencaSegundos)

                .Select(x => NomeCurto(x.Nome))
                .ToList();

            var melhoraram = mudancas
                .Where(x => x.DiferencaSegundos < 0)
                .OrderBy(x => x.DiferencaSegundos)

                .Select(x => NomeCurto(x.Nome))
                .ToList();

            string textoPiorou = pioraram.Count == 0
                ? "Nenhuma piora relevante no TME."
                : $"Pioraram: {JuntarNomes(pioraram)}.";

            string textoMelhorou = melhoraram.Count == 0
                ? "Nenhuma melhora relevante no TME."
                : $"Melhoraram: {JuntarNomes(melhoraram)}.";

            grid.Controls.Add(
                CriarCardFrase(
                    "⏳  TME PIOROU",
                    textoPiorou,
                    CorVermelhoClaro,
                    CorVermelho),
                0, 0);

            grid.Controls.Add(
                CriarCardFrase(
                    "✓  TME MELHOROU",
                    textoMelhorou,
                    CorVerdeClaro,
                    CorVerde),
                1, 0);

            return grid;
        }

        private Control CriarLinhaEquipe()
        {
            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1
            };

            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            if (!comparacaoRealizada)
            {
                grid.Controls.Add(
                    CriarCardFrase(
                        "📞  ATENDIMENTOS",
                        "Aguardando comparação.",
                        CorAzulClaro,
                        CorAzul),
                    0, 0);

                grid.Controls.Add(
                    CriarCardFrase(
                        "⭐  NOTAS",
                        "Aguardando comparação.",
                        CorAmareloClaro,
                        CorAmarelo),
                    1, 0);

                return grid;
            }

            var atendimentos = CompararNumeros(
                dadosAnterior.AtendentesFinalizados,
                dadosAtual.AtendentesFinalizados);

            var aumentaram = atendimentos
                .Where(x => x.Percentual >= 5)
                .OrderByDescending(x => x.Percentual)

                .Select(x => NomeCurto(x.Nome))
                .ToList();

            var diminuiram = atendimentos
                .Where(x => x.Percentual <= -5)
                .OrderBy(x => x.Percentual)

                .Select(x => NomeCurto(x.Nome))
                .ToList();

            string textoAtendimentos = "";

            if (aumentaram.Count > 0)
                textoAtendimentos += $"Aumentaram: {JuntarNomes(aumentaram)}.";

            if (diminuiram.Count > 0)
                textoAtendimentos +=
                    (textoAtendimentos.Length > 0 ? "\n" : "") +
                    $"Diminuíram: {JuntarNomes(diminuiram)}.";

            if (textoAtendimentos.Length == 0)
                textoAtendimentos = "Sem mudanças relevantes no volume de atendimentos.";

            var notas = CompararNotas(
                dadosAnterior.Notas,
                dadosAtual.Notas);

            var notasMelhoraram = notas
                .Where(x => x.Diferenca >= 0.03)
                .OrderByDescending(x => x.Diferenca)

                .Select(x => NomeCurto(x.Nome))
                .ToList();

            var notasPioraram = notas
                .Where(x => x.Diferenca <= -0.03)
                .OrderBy(x => x.Diferenca)

                .Select(x => NomeCurto(x.Nome))
                .ToList();

            string textoNotas = "";

            if (notasMelhoraram.Count > 0)
                textoNotas += $"Melhoraram: {JuntarNomes(notasMelhoraram)}.";

            if (notasPioraram.Count > 0)
                textoNotas +=
                    (textoNotas.Length > 0 ? "\n" : "") +
                    $"Pioraram: {JuntarNomes(notasPioraram)}.";

            if (textoNotas.Length == 0)
                textoNotas = "As notas ficaram estáveis.";

            grid.Controls.Add(
                CriarCardFrase(
                    "📞  ATENDIMENTOS",
                    textoAtendimentos,
                    CorAzulClaro,
                    CorAzul),
                0, 0);

            grid.Controls.Add(
                CriarCardFrase(
                    "⭐  NOTAS",
                    textoNotas,
                    CorAmareloClaro,
                    CorAmarelo),
                1, 0);

            return grid;
        }

        private Control CriarCardFrase(
            string titulo,
            string texto,
            Color fundo,
            Color corTitulo)
        {
            var card = CriarCard(fundo);
            card.Margin = new Padding(4);

            var tituloLabel = new Label
            {
                Dock = DockStyle.Top,
                Height = 34,
                Text = titulo,
                Padding = new Padding(14, 9, 8, 0),
                Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold),
                ForeColor = corTitulo
            };

            var areaTexto = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = false,
                BackColor = fundo,
                Padding = Padding.Empty
            };

            var textoLabel = new Label
            {
                AutoSize = true,
                Text = texto,
                Location = new Point(14, 8),
                Font = new Font("Segoe UI Semibold", 10.2F),
                ForeColor = CorTexto
            };

            void AtualizarScroll()
            {
                if (areaTexto.ClientSize.Width <= 0 ||
                    areaTexto.ClientSize.Height <= 0)
                {
                    return;
                }

                int larguraDisponivel =
                    Math.Max(120, areaTexto.ClientSize.Width - 28);

                // Mantém o texto dentro da largura do card.
                textoLabel.MaximumSize =
                    new Size(larguraDisponivel, 0);

                // Calcula a altura real necessária para o texto.
                Size tamanhoNecessario =
                    TextRenderer.MeasureText(
                        textoLabel.Text,
                        textoLabel.Font,
                        new Size(larguraDisponivel, int.MaxValue),
                        TextFormatFlags.WordBreak |
                        TextFormatFlags.NoPadding);

                int alturaNecessaria =
                    textoLabel.Top +
                    tamanhoNecessario.Height +
                    8;

                bool precisaScroll =
                    alturaNecessaria > areaTexto.ClientSize.Height;

                // Só ativa a barra quando o conteúdo realmente não cabe.
                areaTexto.AutoScroll = precisaScroll;

                if (precisaScroll)
                {
                    areaTexto.AutoScrollMinSize =
                        new Size(0, alturaNecessaria);
                }
                else
                {
                    areaTexto.AutoScrollMinSize = Size.Empty;

                    // Garante que o texto volte para a posição original
                    // caso o usuário tenha rolado antes.
                    areaTexto.AutoScrollPosition = Point.Empty;
                }
            }

            areaTexto.Controls.Add(textoLabel);
            card.Controls.Add(areaTexto);
            card.Controls.Add(tituloLabel);

            areaTexto.Resize += (_, _) => AtualizarScroll();
            textoLabel.TextChanged += (_, _) => AtualizarScroll();

            // Executa também depois que o layout terminar.
            card.HandleCreated += (_, _) =>
            {
                card.BeginInvoke(new Action(AtualizarScroll));
            };

            return card;
        }

        private Control CriarResumoSimples()
        {
            var card = CriarCard(Color.White);
            card.Margin = new Padding(4);

            var titulo = new Label
            {
                Dock = DockStyle.Top,
                Height = 38,
                Text = "📌  RESUMO DO MÊS",
                Padding = new Padding(14, 11, 0, 0),
                Font = new Font("Segoe UI Semibold", 10.5F, FontStyle.Bold),
                ForeColor = CorSidebar
            };

            string texto = comparacaoRealizada
                ? GerarResumoCurto()
                : "Importe os dois arquivos e clique em COMPARAR.\n" +
                  "Aqui aparecerá um resumo simples das principais mudanças do mês.";

            var areaResumo = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = false,
                BackColor = Color.White,
                Padding = Padding.Empty
            };

            var resumo = new Label
            {
                AutoSize = true,
                Text = texto,
                Location = new Point(15, 9),
                Font = new Font("Segoe UI", 9.7F),
                ForeColor = CorTexto
            };

            void AtualizarScrollResumo()
            {
                if (areaResumo.ClientSize.Width <= 0 ||
                    areaResumo.ClientSize.Height <= 0)
                {
                    return;
                }

                int larguraDisponivel =
                    Math.Max(200, areaResumo.ClientSize.Width - 32);

                resumo.MaximumSize =
                    new Size(larguraDisponivel, 0);

                Size tamanhoNecessario =
                    TextRenderer.MeasureText(
                        resumo.Text,
                        resumo.Font,
                        new Size(larguraDisponivel, int.MaxValue),
                        TextFormatFlags.WordBreak |
                        TextFormatFlags.NoPadding);

                int alturaNecessaria =
                    resumo.Top +
                    tamanhoNecessario.Height +
                    10;

                bool precisaScroll =
                    alturaNecessaria > areaResumo.ClientSize.Height;

                areaResumo.AutoScroll = precisaScroll;

                if (precisaScroll)
                {
                    areaResumo.AutoScrollMinSize =
                        new Size(0, alturaNecessaria);
                }
                else
                {
                    areaResumo.AutoScrollMinSize = Size.Empty;
                    areaResumo.AutoScrollPosition = Point.Empty;
                }
            }

            areaResumo.Controls.Add(resumo);
            card.Controls.Add(areaResumo);
            card.Controls.Add(titulo);

            areaResumo.Resize += (_, _) => AtualizarScrollResumo();
            resumo.TextChanged += (_, _) => AtualizarScrollResumo();

            card.HandleCreated += (_, _) =>
            {
                card.BeginInvoke(new Action(AtualizarScrollResumo));
            };

            return card;
        }

        private Control CriarBarraInferior()
        {
            var p = new Panel { Dock = DockStyle.Fill };

            lblStatus = new Label
            {
                AutoSize = true,
                Location = new Point(5, 19),
                Text = comparacaoRealizada
                    ? $"● Comparação concluída: {dadosAnterior.Periodo} x {dadosAtual.Periodo}"
                    : "● Selecione os dois arquivos e clique em comparar.",
                ForeColor = CorSidebar,
                Font = new Font("Segoe UI", 8.3F)
            };

            var botoes = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                Width = 275,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(0, 6, 0, 0)
            };

            var btnLimpar = CriarBotaoAcao("Limpar dados", Color.White, CorVermelho, 125);
            btnLimpar.Click += (_, _) => LimparDados();

            var btnDetalhes = CriarBotaoAcao("Ver atendentes", CorSidebar, Color.White, 130);
            btnDetalhes.Click += (_, _) => MostrarPagina("Atendentes");

            botoes.Controls.Add(btnDetalhes);
            botoes.Controls.Add(btnLimpar);

            p.Controls.Add(lblStatus);
            p.Controls.Add(botoes);
            return p;
        }

        // ============================================================
        // PÁGINA ATENDENTES
        // ============================================================
        private Control CriarPaginaAtendentes()
        {
            var pagina = CriarPaginaDetalhe("Atendentes", "Aqui ficam os números detalhados. A tela principal continua simples.");

            if (!comparacaoRealizada)
            {
                pagina.Controls.Add(CriarMensagemVazia());
                return pagina;
            }

            var grid = CriarDataGrid();
            grid.Columns.Add("Nome", "Atendente");
            grid.Columns.Add("Ant", dadosAnterior.Periodo);
            grid.Columns.Add("Atual", dadosAtual.Periodo);
            grid.Columns.Add("Var", "Variação");
            grid.Columns.Add("TmrAnt", "TMR anterior");
            grid.Columns.Add("TmrAtual", "TMR atual");
            grid.Columns.Add("TmeAnt", "TME anterior");
            grid.Columns.Add("TmeAtual", "TME atual");
            grid.Columns.Add("NotaAnt", "Nota anterior");
            grid.Columns.Add("NotaAtual", "Nota atual");
            grid.Columns.Add("Leitura", "Leitura rápida");

            var nomes = dadosAnterior.AtendentesFinalizados.Keys
                .Union(dadosAtual.AtendentesFinalizados.Keys)
                .OrderBy(x => NomeCurto(x))
                .ToList();

            foreach (var chave in nomes)
            {
                double ant = Valor(dadosAnterior.AtendentesFinalizados, chave);
                double atual = Valor(dadosAtual.AtendentesFinalizados, chave);
                double pct = CalcularPercentual(ant, atual);

                TimeSpan? tmrAnt = Tempo(dadosAnterior.TMR, chave);
                TimeSpan? tmrAt = Tempo(dadosAtual.TMR, chave);
                TimeSpan? tmeAnt = Tempo(dadosAnterior.TME, chave);
                TimeSpan? tmeAt = Tempo(dadosAtual.TME, chave);

                double notaAnt = Valor(dadosAnterior.Notas, chave);
                double notaAt = Valor(dadosAtual.Notas, chave);

                string leitura = GerarLeituraAtendente(
                    ant,
                    atual,
                    tmrAnt,
                    tmrAt,
                    tmeAnt,
                    tmeAt,
                    notaAnt,
                    notaAt);

                int r = grid.Rows.Add(
                    NomeExibicao(chave),
                    ant.ToString("N0"),
                    atual.ToString("N0"),
                    FormatarPercentualComSeta(pct),
                    FormatarTempo(tmrAnt),
                    FormatarTempo(tmrAt),
                    FormatarTempo(tmeAnt),
                    FormatarTempo(tmeAt),
                    notaAnt > 0 ? notaAnt.ToString("N2") : "—",
                    notaAt > 0 ? notaAt.ToString("N2") : "—",
                    leitura);

                var linha = grid.Rows[r];
                linha.Cells["Var"].Style.ForeColor = pct >= 0 ? CorVerde : CorVermelho;
            }

            // =====================================================
            // TABELA SEM SCROLL HORIZONTAL
            // =====================================================
            // Todas as colunas se ajustam à largura disponível.
            // A coluna "Leitura rápida" recebe mais espaço e quebra
            // o texto em mais de uma linha quando necessário.
            grid.ScrollBars = ScrollBars.Vertical;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            grid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            grid.RowTemplate.MinimumHeight = 28;

            grid.DefaultCellStyle.WrapMode = DataGridViewTriState.False;

            grid.Columns["Nome"].FillWeight = 150;
            grid.Columns["Ant"].FillWeight = 72;
            grid.Columns["Atual"].FillWeight = 72;
            grid.Columns["Var"].FillWeight = 70;

            grid.Columns["TmrAnt"].FillWeight = 83;
            grid.Columns["TmrAtual"].FillWeight = 83;

            grid.Columns["TmeAnt"].FillWeight = 83;
            grid.Columns["TmeAtual"].FillWeight = 83;

            grid.Columns["NotaAnt"].FillWeight = 68;
            grid.Columns["NotaAtual"].FillWeight = 68;

            // Maior espaço para o texto explicativo.
            grid.Columns["Leitura"].FillWeight = 200;
            grid.Columns["Leitura"].MinimumWidth = 180;
            grid.Columns["Leitura"].DefaultCellStyle.WrapMode =
                DataGridViewTriState.True;

            grid.Columns["Nome"].MinimumWidth = 130;
            grid.Columns["Ant"].MinimumWidth = 65;
            grid.Columns["Atual"].MinimumWidth = 65;
            grid.Columns["Var"].MinimumWidth = 65;

            grid.Columns["TmrAnt"].MinimumWidth = 72;
            grid.Columns["TmrAtual"].MinimumWidth = 72;
            grid.Columns["TmeAnt"].MinimumWidth = 72;
            grid.Columns["TmeAtual"].MinimumWidth = 72;

            grid.Columns["NotaAnt"].MinimumWidth = 62;
            grid.Columns["NotaAtual"].MinimumWidth = 62;

            pagina.Controls.Add(grid);
            return pagina;
        }

        // ============================================================
        // PÁGINA MOTIVOS
        // ============================================================
        private Control CriarPaginaMotivos()
        {
            var pagina = CriarPaginaDetalhe("Motivos de atendimento", "Compare cada motivo sem misturar com a tela principal.");

            if (!comparacaoRealizada)
            {
                pagina.Controls.Add(CriarMensagemVazia());
                return pagina;
            }

            var grid = CriarDataGrid();
            grid.Columns.Add("Motivo", "Motivo");
            grid.Columns.Add("Ant", dadosAnterior.Periodo);
            grid.Columns.Add("Atual", dadosAtual.Periodo);
            grid.Columns.Add("Dif", "Diferença");
            grid.Columns.Add("Var", "Variação");
            grid.Columns.Add("Leitura", "Leitura rápida");

            var motivos = dadosAnterior.Motivos.Keys
                .Union(dadosAtual.Motivos.Keys)
                .OrderByDescending(x => Math.Max(Valor(dadosAnterior.Motivos, x), Valor(dadosAtual.Motivos, x)))
                .ToList();

            foreach (var motivo in motivos)
            {
                double ant = Valor(dadosAnterior.Motivos, motivo);
                double atual = Valor(dadosAtual.Motivos, motivo);
                double dif = atual - ant;
                double pct = CalcularPercentual(ant, atual);

                int r = grid.Rows.Add(
                    motivo,
                    ant.ToString("N0"),
                    atual.ToString("N0"),
                    $"{dif:+0;-0;0}",
                    FormatarPercentualComSeta(pct),
                    dif > 0 ? "Aumentou" : dif < 0 ? "Reduziu" : "Estável");

                grid.Rows[r].Cells["Var"].Style.ForeColor =
                    dif > 0 ? CorAzul : dif < 0 ? CorVerde : CorTextoSecundario;
            }

            grid.Columns["Motivo"].FillWeight = 190;
            grid.Columns["Leitura"].FillWeight = 120;

            pagina.Controls.Add(grid);
            return pagina;
        }

        // ============================================================
        // PÁGINA MARKETING
        // ============================================================
        private Control CriarPaginaMarketing()
        {
            var pagina = CriarPaginaDetalhe("Marketing", "Veja quais criativos geraram mais contatos e como mudou de um mês para o outro.");

            if (!comparacaoRealizada)
            {
                pagina.Controls.Add(CriarMensagemVazia());
                return pagina;
            }

            var area = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                BackColor = CorFundo
            };
            area.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            area.RowStyles.Add(new RowStyle(SizeType.Absolute, 104));
            area.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var topo = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(0, 2, 0, 4)
            };

            topo.Controls.Add(CriarMiniCard(
                "Contatos de marketing",
                $"{dadosAnterior.MarketingTotal:N0} → {dadosAtual.MarketingTotal:N0}",
                CalcularPercentual(dadosAnterior.MarketingTotal, dadosAtual.MarketingTotal)));

            topo.Controls.Add(CriarMiniCard(
                "Novos contatos",
                $"{dadosAnterior.NovosContatos:N0} → {dadosAtual.NovosContatos:N0}",
                CalcularPercentual(dadosAnterior.NovosContatos, dadosAtual.NovosContatos)));

            topo.Controls.Add(CriarMiniCard(
                "Reagendamentos",
                $"{dadosAnterior.Reagendamentos:N0} → {dadosAtual.Reagendamentos:N0}",
                CalcularPercentual(dadosAnterior.Reagendamentos, dadosAtual.Reagendamentos)));

            var grid = CriarDataGrid();
            grid.Columns.Add("Campanha", "Criativo / campanha");
            grid.Columns.Add("Ant", dadosAnterior.Periodo);
            grid.Columns.Add("Atual", dadosAtual.Periodo);
            grid.Columns.Add("Var", "Variação");

            var campanhas = dadosAnterior.MarketingCriativos.Keys
                .Union(dadosAtual.MarketingCriativos.Keys)
                .OrderByDescending(x => Math.Max(
                    Valor(dadosAnterior.MarketingCriativos, x),
                    Valor(dadosAtual.MarketingCriativos, x)))
                .ToList();

            foreach (var c in campanhas)
            {
                double ant = Valor(dadosAnterior.MarketingCriativos, c);
                double atual = Valor(dadosAtual.MarketingCriativos, c);
                double pct = CalcularPercentual(ant, atual);

                int r = grid.Rows.Add(
                    c,
                    ant.ToString("N0"),
                    atual.ToString("N0"),
                    FormatarPercentualComSeta(pct));

                grid.Rows[r].Cells["Var"].Style.ForeColor =
                    pct >= 0 ? CorVerde : CorVermelho;
            }

            grid.Columns["Campanha"].FillWeight = 250;

            area.Controls.Add(topo, 0, 0);
            area.Controls.Add(grid, 0, 1);
            pagina.Controls.Add(area);

            return pagina;
        }

        // ============================================================
        // PÁGINA QUALIDADE
        // ============================================================
        private Control CriarPaginaQualidade()
        {
            var pagina = CriarPaginaDetalhe("Qualidade dos dados", "Mostra apenas problemas que podem atrapalhar a comparação.");

            if (!comparacaoRealizada)
            {
                pagina.Controls.Add(CriarMensagemVazia());
                return pagina;
            }

            var lista = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                Padding = new Padding(0, 10, 0, 0)
            };

            var alertas = GerarAlertasQualidade();

            if (alertas.Count == 0)
            {
                lista.Controls.Add(CriarAlertaLinha("✓", "Nenhuma inconsistência importante encontrada.", CorVerdeClaro, CorVerde));
            }
            else
            {
                foreach (var a in alertas)
                    lista.Controls.Add(CriarAlertaLinha("⚠", a, CorAmareloClaro, CorAmarelo));
            }

            pagina.Controls.Add(lista);
            return pagina;
        }

        // ============================================================
        // PÁGINA EXPORTAR
        // ============================================================
        private Control CriarPaginaExportar()
        {
            var pagina = CriarPaginaDetalhe(
                "Exportar",
                "Escolha qual tela deseja salvar em PDF. A aparência do relatório será mantida.");

            var area = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = CorFundo,
                Padding = new Padding(18, 20, 18, 18)
            };

            var card = new Panel
            {
                Width = 620,
                Height = 285,
                BackColor = Color.White,
                Location = new Point(18, 18),
                Padding = new Padding(22)
            };

            var lblStatusExportacao = new Label
            {
                AutoSize = true,
                Location = new Point(22, 20),
                Text = comparacaoRealizada
                    ? $"Comparação pronta: {dadosAnterior.Periodo} x {dadosAtual.Periodo}"
                    : "Primeiro importe os dois arquivos e clique em COMPARAR.",
                Font = new Font("Segoe UI Semibold", 10.5F),
                ForeColor = CorTexto
            };

            var lblEscolha = new Label
            {
                AutoSize = true,
                Location = new Point(22, 65),
                Text = "Qual tela deseja exportar?",
                Font = new Font("Segoe UI Semibold", 9.5F),
                ForeColor = CorTexto
            };

            var cboTela = new ComboBox
            {
                Location = new Point(22, 90),
                Width = 360,
                Height = 32,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9.5F),
                BackColor = Color.White
            };

            cboTela.Items.AddRange(new object[]
            {
                "Comparativo",
                "Atendentes",
                "Motivos",
                "Marketing",
                "Todos"
            });

            cboTela.SelectedIndex = 0;

            var lblAjuda = new Label
            {
                Location = new Point(22, 132),
                Width = 555,
                Height = 42,
                Text = "Ao escolher “Todos”, será criado um único PDF com uma página para cada tela: " +
                       "Comparativo, Atendentes, Motivos e Marketing.",
                Font = new Font("Segoe UI", 8.7F),
                ForeColor = CorTextoSecundario
            };

            var btnPdf = CriarBotaoAcao(
                "Exportar PDF",
                CorVerde,
                Color.White,
                180);

            btnPdf.Location = new Point(22, 195);
            btnPdf.Height = 44;
            btnPdf.Enabled = comparacaoRealizada;
            btnPdf.Anchor = AnchorStyles.Left | AnchorStyles.Top;

            btnPdf.Click += (_, _) =>
            {
                string opcao =
                    cboTela.SelectedItem?.ToString()
                    ?? "Comparativo";

                ExportarTelasParaPdf(opcao);
            };

            var btnLimpar = CriarBotaoAcao(
                "Limpar dados",
                Color.White,
                CorVermelho,
                145);

            btnLimpar.Location = new Point(214, 197);
            btnLimpar.Height = 40;
            btnLimpar.Anchor = AnchorStyles.Left | AnchorStyles.Top;
            btnLimpar.Click += (_, _) => LimparDados();

            card.Controls.Add(lblStatusExportacao);
            card.Controls.Add(lblEscolha);
            card.Controls.Add(cboTela);
            card.Controls.Add(lblAjuda);
            card.Controls.Add(btnPdf);
            card.Controls.Add(btnLimpar);

            area.Controls.Add(card);
            pagina.Controls.Add(area);

            return pagina;
        }

        // ============================================================
        // COMPONENTES AUXILIARES DE TELA
        // ============================================================
        private Panel CriarPaginaDetalhe(string titulo, string subtitulo)
        {
            var p = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = CorFundo,
                // O conteúdo que for adicionado com Dock=Fill começa abaixo do cabeçalho.
                Padding = new Padding(4, 70, 4, 4)
            };

            var cab = new Panel
            {
                Location = new Point(4, 4),
                Height = 62,
                Width = Math.Max(300, p.ClientSize.Width - 8),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = CorFundo
            };

            var t = new Label
            {
                AutoSize = true,
                Location = new Point(5, 5),
                Text = titulo,
                Font = new Font("Segoe UI Semibold", 16F, FontStyle.Bold),
                ForeColor = CorTexto
            };

            var s = new Label
            {
                AutoSize = true,
                Location = new Point(7, 37),
                Text = subtitulo,
                Font = new Font("Segoe UI", 9F),
                ForeColor = CorTextoSecundario
            };

            cab.Controls.Add(t);
            cab.Controls.Add(s);
            p.Controls.Add(cab);
            return p;
        }

        private Control CriarMensagemVazia()
        {
            return new Label
            {
                Dock = DockStyle.Fill,
                Text = "Nenhuma comparação realizada.\n\nSelecione os dois arquivos acima e clique em COMPARAR.",
                Font = new Font("Segoe UI", 12F),
                ForeColor = CorTextoSecundario,
                TextAlign = ContentAlignment.MiddleCenter
            };
        }

        private Panel CriarScroll()
        {
            return new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = CorFundo
            };
        }

        private Panel CriarCard(Color fundo)
        {
            return new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = fundo,
                BorderStyle = BorderStyle.None
            };
        }

        private Button CriarBotaoAcao(string texto, Color fundo, Color corTexto, int largura)
        {
            var b = new Button
            {
                Width = largura,
                Height = 38,
                Text = texto,
                BackColor = fundo,
                ForeColor = corTexto,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI Semibold", 9F),
                Cursor = Cursors.Hand,
                Margin = new Padding(5)
            };
            b.FlatAppearance.BorderColor = CorBorda;
            return b;
        }

        private Control CriarMiniCard(string titulo, string valor, double percentual)
        {
            var p = CriarCard(Color.White);
            p.Dock = DockStyle.None;
            p.Width = 240;
            p.Height = 88;
            p.Margin = new Padding(4, 3, 4, 3);

            var t = new Label
            {
                AutoSize = true,
                Location = new Point(12, 10),
                Text = titulo,
                Font = new Font("Segoe UI", 8.8F),
                ForeColor = CorTextoSecundario
            };
            var v = new Label
            {
                AutoSize = true,
                Location = new Point(12, 33),
                Text = valor,
                Font = new Font("Segoe UI Semibold", 13F, FontStyle.Bold),
                ForeColor = CorTexto
            };
            var d = new Label
            {
                AutoSize = true,
                Location = new Point(12, 62),
                Text = FormatarPercentualComSeta(percentual),
                Font = new Font("Segoe UI Semibold", 8.5F),
                ForeColor = percentual >= 0 ? CorVerde : CorVermelho
            };

            p.Controls.Add(t);
            p.Controls.Add(v);
            p.Controls.Add(d);
            return p;
        }

        private Control CriarAlertaLinha(string icone, string texto, Color fundo, Color cor)
        {
            var p = new Panel
            {
                Width = 900,
                Height = 56,
                BackColor = fundo,
                BorderStyle = BorderStyle.None,
                Margin = new Padding(4)
            };

            var l = new Label
            {
                Dock = DockStyle.Fill,
                Text = $"{icone}  {texto}",
                Padding = new Padding(14, 16, 10, 0),
                Font = new Font("Segoe UI Semibold", 9.5F),
                ForeColor = cor
            };
            p.Controls.Add(l);
            return p;
        }

        private DataGridView CriarDataGrid()
        {
            var g = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                MultiSelect = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 34
            };

            g.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(233, 238, 243);
            g.ColumnHeadersDefaultCellStyle.ForeColor = CorTexto;
            g.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI Semibold", 8.5F, FontStyle.Bold);
            g.DefaultCellStyle.Font = new Font("Segoe UI", 8.5F);
            g.DefaultCellStyle.ForeColor = CorTexto;
            g.DefaultCellStyle.SelectionBackColor = Color.FromArgb(226, 241, 235);
            g.DefaultCellStyle.SelectionForeColor = CorTexto;
            g.GridColor = CorBorda;
            g.RowTemplate.Height = 29;
            return g;
        }

        // ============================================================
        // IMPORTAÇÃO / COMPARAÇÃO
        // ============================================================
        private void SelecionarAnterior(object sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog
            {
                Title = "Selecione o Excel do mês retrasado",
                Filter = "Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
                Multiselect = false,
                CheckFileExists = true
            };

            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            caminhoAnterior = ofd.FileName;
            txtAnterior.Text = Path.GetFileName(ofd.FileName);
            lblPeriodoAnterior.Text = "Arquivo selecionado. Clique em COMPARAR.";
        }

        private void SelecionarAtual(object sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog
            {
                Title = "Selecione o Excel do mês passado",
                Filter = "Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
                Multiselect = false,
                CheckFileExists = true
            };

            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            caminhoAtual = ofd.FileName;
            txtAtual.Text = Path.GetFileName(ofd.FileName);
            lblPeriodoAtual.Text = "Arquivo selecionado. Clique em COMPARAR.";
        }

        private void BtnComparar_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(caminhoAnterior) || !File.Exists(caminhoAnterior))
            {
                MessageBox.Show("Selecione o arquivo do mês retrasado.", "Arquivo necessário",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(caminhoAtual) || !File.Exists(caminhoAtual))
            {
                MessageBox.Show("Selecione o arquivo do mês passado.", "Arquivo necessário",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                Cursor = Cursors.WaitCursor;

                dadosAnterior = LerArquivo(caminhoAnterior);
                dadosAtual = LerArquivo(caminhoAtual);

                comparacaoRealizada = true;

                lblPeriodoAnterior.Text = $"Período: {dadosAnterior.Periodo}";
                lblPeriodoAtual.Text = $"Período: {dadosAtual.Periodo}";

                MostrarPagina("Comparativo");
            }
            catch (Exception ex)
            {
                comparacaoRealizada = false;
                MessageBox.Show(
                    "Não consegui ler uma das planilhas.\n\n" +
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

        private void LimparDados()
        {
            caminhoAnterior = "";
            caminhoAtual = "";
            dadosAnterior = null;
            dadosAtual = null;
            comparacaoRealizada = false;

            txtAnterior.Text = "Nenhum arquivo selecionado";
            txtAtual.Text = "Nenhum arquivo selecionado";
            lblPeriodoAnterior.Text = "Período: aguardando arquivo...";
            lblPeriodoAtual.Text = "Período: aguardando arquivo...";

            MostrarPagina("Comparativo");
        }

        // ============================================================
        // LEITURA DO EXCEL
        // ============================================================
        private DadosMes LerArquivo(string caminho)
        {
            using var wb = new XLWorkbook(caminho);

            var d = new DadosMes
            {
                Arquivo = caminho,
                Periodo = DescobrirPeriodo(wb, caminho)
            };

            // Conversas
            var wsRelacao = ObterAba(wb, "Relação Conversas-Atendimento");
            if (wsRelacao != null)
                d.Conversas = ProcurarValorAposRotulo(wsRelacao, "TOTAL MÊS", colunasADireita: 2);

            // Novos contatos
            var wsMensagens = ObterAba(wb, "Mensagens por número");
            if (wsMensagens != null)
                d.NovosContatos = ProcurarPrimeiroTotal(wsMensagens, 1, 2, 45);

            // Finalizados
            var wsFinal = ObterAba(wb, "Atendimentos finalizados");
            if (wsFinal != null)
            {
                d.AtendentesFinalizados = LerListaNomeValor(wsFinal, pararEmTotal: true);
                d.Finalizados = ProcurarUltimoTotal(wsFinal);
            }

            // Motivos
            var wsMotivos = ObterAba(wb, "Motivos Atendimento");
            if (wsMotivos != null)
            {
                d.Motivos = LerMotivosPrincipais(wsMotivos);
                d.Inatividade = ValorPorDescricao(d.Motivos, "Encerrado por inatividade do cliente");
                d.OrcamentoFormula = ValorPorDescricao(d.Motivos, "Orçamento de fórmula");

                d.OrcamentoFormulaPossivelVenda = ProcurarValorNaSecao(
                    wsMotivos, "Possível venda", "Orçamento de fórmula");
            }

            // Templates
            var wsTemplate = ObterAba(wb, "Mensagem Template");
            if (wsTemplate != null)
            {
                d.Reagendamentos = ProcurarPorInicioDeTexto(
                    wsTemplate, new[] { "Reagendamento", "Reagendamentos entregues" });

                d.TemplatesTotal = ProcurarUltimoTotal(wsTemplate);
            }

            // Marketing
            var wsMarketing = ObterAba(wb, "Marketing");
            if (wsMarketing != null)
            {
                d.MarketingCriativos = LerMarketing(wsMarketing);
                d.MarketingTotal = ProcurarUltimoTotal(wsMarketing);
            }

            // TMR / TME / TMA
            var wsTmr = ObterAba(wb, "TMR");
            if (wsTmr != null) d.TMR = LerNomeTempo(wsTmr);

            var wsTme = ObterAba(wb, "TME");
            if (wsTme != null) d.TME = LerNomeTempo(wsTme);

            var wsTma = ObterAba(wb, "TMA");
            if (wsTma != null) d.TMA = LerNomeTempo(wsTma);

            // Notas
            var wsNotas = ObterAba(wb, "MEDIA - IND") ?? ObterAba(wb, "Média individual");
            if (wsNotas != null)
            {
                d.Notas = LerNotas(wsNotas);
                d.NotaGeral = AjustarNota(ProcurarValorAoLado(wsNotas, "Média geral (W.L + Gurgel)"));
                d.NotaWL = AjustarNota(ProcurarNotaUnidade(wsNotas, "W.L"));
                d.NotaGurgel = AjustarNota(ProcurarNotaUnidade(wsNotas, "Gurgel"));
            }

            return d;
        }

        private IXLWorksheet ObterAba(XLWorkbook wb, string nome)
        {
            string alvo = NormalizarTexto(nome);
            return wb.Worksheets.FirstOrDefault(w => NormalizarTexto(w.Name) == alvo);
        }

        private string DescobrirPeriodo(XLWorkbook wb, string caminho)
        {
            var ws = ObterAba(wb, "Relação Conversas-Atendimento");
            if (ws != null)
            {
                var range = ws.RangeUsed();
                if (range != null)
                {
                    foreach (var cell in range.Cells())
                    {
                        if (cell.TryGetValue<DateTime>(out var dt) &&
                            dt.Year >= 2020 && dt.Year <= 2100)
                        {
                            return CultureInfo.GetCultureInfo("pt-BR")
                                .TextInfo.ToTitleCase(dt.ToString("MMMM/yyyy", new CultureInfo("pt-BR")));
                        }

                        string txt = cell.GetFormattedString().Trim();
                        if (DateTime.TryParseExact(txt,
                            new[] { "dd/MM/yyyy", "d/M/yyyy", "dd/MM/yy" },
                            new CultureInfo("pt-BR"),
                            DateTimeStyles.None,
                            out dt))
                        {
                            return CultureInfo.GetCultureInfo("pt-BR")
                                .TextInfo.ToTitleCase(dt.ToString("MMMM/yyyy", new CultureInfo("pt-BR")));
                        }
                    }
                }
            }

            return Path.GetFileNameWithoutExtension(caminho);
        }

        private double ProcurarValorAposRotulo(IXLWorksheet ws, string rotulo, int colunasADireita)
        {
            string alvo = NormalizarTexto(rotulo);
            var used = ws.RangeUsed();
            if (used == null) return 0;

            foreach (var c in used.Cells())
            {
                string txt = NormalizarTexto(c.GetFormattedString());
                if (!txt.Contains(alvo)) continue;

                for (int i = 1; i <= colunasADireita; i++)
                {
                    var n = LerNumero(c.CellRight(i));
                    if (n.HasValue) return n.Value;
                }
            }

            return 0;
        }

        private double ProcurarPrimeiroTotal(IXLWorksheet ws, int colRotulo, int colValor, int limiteLinhas)
        {
            int max = Math.Min(limiteLinhas, ws.LastRowUsed()?.RowNumber() ?? limiteLinhas);

            for (int r = 1; r <= max; r++)
            {
                string txt = NormalizarTexto(ws.Cell(r, colRotulo).GetFormattedString());
                if (txt == "total" || txt.StartsWith("total "))
                {
                    var v = LerNumero(ws.Cell(r, colValor));
                    if (v.HasValue) return v.Value;
                }
            }

            return 0;
        }

        private double ProcurarUltimoTotal(IXLWorksheet ws)
        {
            double ultimo = 0;
            var used = ws.RangeUsed();
            if (used == null) return 0;

            int totalColunas = used.ColumnCount();

            foreach (var row in used.Rows())
            {
                for (int c = 1; c <= Math.Min(3, totalColunas); c++)
                {
                    var cell = row.Cell(c);
                    string txt = NormalizarTexto(cell.GetFormattedString());

                    if (txt.StartsWith("total"))
                    {
                        for (int x = c + 1; x <= Math.Min(c + 2, totalColunas); x++)
                        {
                            var v = LerNumero(row.Cell(x));
                            if (v.HasValue)
                            {
                                ultimo = v.Value;
                                break;
                            }
                        }
                    }
                }
            }

            return ultimo;
        }

        private Dictionary<string, double> LerListaNomeValor(IXLWorksheet ws, bool pararEmTotal)
        {
            var dict = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            int ultima = ws.LastRowUsed()?.RowNumber() ?? 0;

            for (int r = 1; r <= ultima; r++)
            {
                string nome = ws.Cell(r, 1).GetFormattedString().Trim();
                if (string.IsNullOrWhiteSpace(nome)) continue;

                string norm = NormalizarTexto(nome);

                if (pararEmTotal && norm.StartsWith("total"))
                    break;

                if (norm.Contains("atendimentos finalizados"))
                    continue;

                var valor = LerNumero(ws.Cell(r, 2));
                if (!valor.HasValue) continue;

                string chave = NormalizarNome(nome);
                if (!dict.ContainsKey(chave))
                    dict[chave] = valor.Value;
            }

            return dict;
        }

        private Dictionary<string, double> LerMotivosPrincipais(IXLWorksheet ws)
        {
            var dict = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            bool iniciou = false;
            int ultima = ws.LastRowUsed()?.RowNumber() ?? 0;

            for (int r = 1; r <= ultima; r++)
            {
                string desc = ws.Cell(r, 1).GetFormattedString().Trim();
                string norm = NormalizarTexto(desc);

                if (norm == "motivo")
                {
                    iniciou = true;
                    continue;
                }

                if (!iniciou) continue;
                if (norm.StartsWith("total")) break;

                var v = LerNumero(ws.Cell(r, 2));
                if (!v.HasValue || string.IsNullOrWhiteSpace(desc)) continue;

                dict[desc] = v.Value;
            }

            return dict;
        }

        private double ProcurarValorNaSecao(IXLWorksheet ws, string secao, string item)
        {
            bool dentro = false;
            string secaoNorm = NormalizarTexto(secao);
            string itemNorm = NormalizarTexto(item);
            int ultima = ws.LastRowUsed()?.RowNumber() ?? 0;

            for (int r = 1; r <= ultima; r++)
            {
                string a = NormalizarTexto(ws.Cell(r, 1).GetFormattedString());

                if (a.Contains(secaoNorm))
                {
                    dentro = true;
                    continue;
                }

                if (!dentro) continue;

                if (a.StartsWith("total"))
                    break;

                if (a == itemNorm)
                {
                    var v = LerNumero(ws.Cell(r, 2));
                    return v ?? 0;
                }
            }

            return 0;
        }

        private double ProcurarPorInicioDeTexto(IXLWorksheet ws, IEnumerable<string> rotulos)
        {
            var alvos = rotulos.Select(NormalizarTexto).ToList();
            int ultima = ws.LastRowUsed()?.RowNumber() ?? 0;

            for (int r = 1; r <= ultima; r++)
            {
                string txt = NormalizarTexto(ws.Cell(r, 1).GetFormattedString());
                if (alvos.Any(x => txt.StartsWith(x)))
                {
                    var v = LerNumero(ws.Cell(r, 2));
                    if (v.HasValue) return v.Value;
                }
            }

            return 0;
        }

        private Dictionary<string, double> LerMarketing(IXLWorksheet ws)
        {
            var dict = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            int ultima = ws.LastRowUsed()?.RowNumber() ?? 0;

            for (int r = 1; r <= ultima; r++)
            {
                string criativo = ws.Cell(r, 2).GetFormattedString().Trim();
                if (string.IsNullOrWhiteSpace(criativo)) continue;

                string norm = NormalizarTexto(criativo);
                if (norm == "marketing" || norm == "criativos")
                    continue;

                var v = LerNumero(ws.Cell(r, 3));
                if (!v.HasValue) continue;

                dict[criativo] = v.Value;
            }

            return dict;
        }

        private Dictionary<string, TimeSpan?> LerNomeTempo(IXLWorksheet ws)
        {
            var dict = new Dictionary<string, TimeSpan?>(StringComparer.OrdinalIgnoreCase);
            int ultima = ws.LastRowUsed()?.RowNumber() ?? 0;

            // A/B contém a lista completa; D/E é apenas uma separação por unidade.
            for (int r = 1; r <= ultima; r++)
            {
                string nome = ws.Cell(r, 1).GetFormattedString().Trim();
                if (string.IsNullOrWhiteSpace(nome)) continue;

                string norm = NormalizarTexto(nome);
                if (norm.StartsWith("agente")) continue;

                string tempoTxt = ws.Cell(r, 2).GetFormattedString().Trim();
                TimeSpan? tempo = ParseTempo(tempoTxt);

                string chave = NormalizarNome(nome);
                if (!dict.ContainsKey(chave))
                    dict[chave] = tempo;
            }

            return dict;
        }

        private Dictionary<string, double> LerNotas(IXLWorksheet ws)
        {
            var dict = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            int ultima = ws.LastRowUsed()?.RowNumber() ?? 0;

            for (int r = 1; r <= ultima; r++)
            {
                string nome = ws.Cell(r, 1).GetFormattedString().Trim();
                string norm = NormalizarTexto(nome);

                if (string.IsNullOrWhiteSpace(nome)) continue;
                if (norm.StartsWith("atendentes wl") ||
                    norm.StartsWith("atendentes gu") ||
                    norm.StartsWith("media unidade") ||
                    norm.StartsWith("nota") ||
                    norm.StartsWith("total"))
                    continue;

                var nota = LerNumero(ws.Cell(r, 2));

                if (!nota.HasValue)
                    continue;

                // Corrige casos em que "4.88" é interpretado como 488,
                // "4.93" como 493 etc.
                double notaCorrigida = AjustarNota(nota.Value);

                if (notaCorrigida < 0 || notaCorrigida > 5)
                    continue;

                // Só aceita nomes que parecem atendentes.
                if (!(nome.Contains("-") ||
                      nome.Contains("W. Luiz", StringComparison.OrdinalIgnoreCase) ||
                      nome.Contains("Gurgel", StringComparison.OrdinalIgnoreCase) ||
                      nome.Contains("Maria Angela", StringComparison.OrdinalIgnoreCase)))
                    continue;

                string chave = NormalizarNome(nome);

                if (!dict.ContainsKey(chave))
                    dict[chave] = notaCorrigida;
            }

            return dict;
        }

        private double ProcurarValorAoLado(IXLWorksheet ws, string rotulo)
        {
            string alvo = NormalizarTexto(rotulo);
            var used = ws.RangeUsed();
            if (used == null) return 0;

            foreach (var c in used.Cells())
            {
                if (NormalizarTexto(c.GetFormattedString()).Contains(alvo))
                {
                    var v = LerNumero(c.CellRight());
                    if (v.HasValue) return v.Value;
                }
            }

            return 0;
        }

        private double ProcurarNotaUnidade(IXLWorksheet ws, string unidade)
        {
            string alvo = NormalizarTexto(unidade);
            var used = ws.RangeUsed();
            if (used == null) return 0;

            int totalColunas = used.ColumnCount();

            foreach (var row in used.Rows())
            {
                for (int c = 1; c <= totalColunas - 1; c++)
                {
                    string txt = NormalizarTexto(row.Cell(c).GetFormattedString());
                    if (txt == alvo)
                    {
                        var v = LerNumero(row.Cell(c + 1));
                        if (v.HasValue && v.Value >= 0 && v.Value <= 5)
                            return v.Value;
                    }
                }
            }

            return 0;
        }

        private double? LerNumero(IXLCell cell)
        {
            if (cell == null || cell.IsEmpty())
                return null;

            if (cell.TryGetValue<double>(out var d))
                return d;

            string txt = cell.GetFormattedString()
                .Replace("R$", "", StringComparison.OrdinalIgnoreCase)
                .Replace("%", "")
                .Trim();

            // Algumas notas das planilhas estão gravadas como texto "4.85".
            // Se houver somente ponto, trata primeiro como decimal no padrão Invariant.
            if (txt.Contains(".") && !txt.Contains(","))
            {
                if (double.TryParse(txt, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                    return d;

                if (double.TryParse(txt, NumberStyles.Any, new CultureInfo("pt-BR"), out d))
                    return d;
            }
            else
            {
                if (double.TryParse(txt, NumberStyles.Any, new CultureInfo("pt-BR"), out d))
                    return d;

                if (double.TryParse(txt, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                    return d;
            }

            return null;
        }

        private double AjustarNota(double nota)
        {
            // Defesa contra valores textuais como 4.85 interpretados como 485.
            while (nota > 5 && nota <= 5000)
                nota /= 10.0;

            return nota;
        }

        private TimeSpan? ParseTempo(string valor)
        {
            if (string.IsNullOrWhiteSpace(valor))
                return null;

            valor = valor.Trim();

            // Ex.: 2d 1:41:9
            if (valor.Contains("d ", StringComparison.OrdinalIgnoreCase))
            {
                var partes = valor.Split('d', 2);
                if (int.TryParse(partes[0].Trim(), out int dias) &&
                    TimeSpan.TryParse(partes[1].Trim(), out var resto))
                    return TimeSpan.FromDays(dias) + resto;
            }

            if (TimeSpan.TryParse(valor, new CultureInfo("pt-BR"), out var ts))
                return ts;

            if (TimeSpan.TryParse(valor, CultureInfo.InvariantCulture, out ts))
                return ts;

            return null;
        }

        // ============================================================
        // RESUMOS AUTOMÁTICOS
        // ============================================================
        private string GerarResumoCurto()
        {
            var linhas = new List<string>();

            // =====================================================
            // VOLUME GERAL
            // =====================================================
            double variacaoConversas =
                CalcularPercentual(
                    dadosAnterior.Conversas,
                    dadosAtual.Conversas);

            double variacaoFinalizados =
                CalcularPercentual(
                    dadosAnterior.Finalizados,
                    dadosAtual.Finalizados);

            linhas.Add(
                variacaoConversas >= 0
                    ? $"• Conversas: aumentaram {Math.Abs(variacaoConversas):N1}% " +
                      $"({dadosAnterior.Conversas:N0} → {dadosAtual.Conversas:N0})."
                    : $"• Conversas: diminuíram {Math.Abs(variacaoConversas):N1}% " +
                      $"({dadosAnterior.Conversas:N0} → {dadosAtual.Conversas:N0}).");

            linhas.Add(
                variacaoFinalizados >= 0
                    ? $"• Finalizados: aumentaram {Math.Abs(variacaoFinalizados):N1}% " +
                      $"({dadosAnterior.Finalizados:N0} → {dadosAtual.Finalizados:N0})."
                    : $"• Finalizados: diminuíram {Math.Abs(variacaoFinalizados):N1}% " +
                      $"({dadosAnterior.Finalizados:N0} → {dadosAtual.Finalizados:N0}).");

            // Taxa de finalização
            if (dadosAnterior.Conversas > 0 &&
                dadosAtual.Conversas > 0)
            {
                double taxaAnterior =
                    dadosAnterior.Finalizados /
                    dadosAnterior.Conversas *
                    100.0;

                double taxaAtual =
                    dadosAtual.Finalizados /
                    dadosAtual.Conversas *
                    100.0;

                double diferencaPp =
                    taxaAtual - taxaAnterior;

                string leituraTaxa =
                    diferencaPp > 0.05
                        ? "melhorou"
                        : diferencaPp < -0.05
                            ? "caiu"
                            : "ficou estável";

                linhas.Add(
                    $"• Taxa de finalização: {leituraTaxa} " +
                    $"({taxaAnterior:N1}% → {taxaAtual:N1}%; " +
                    $"{diferencaPp:+0.0;-0.0;0.0} p.p.).");
            }

            // =====================================================
            // CONTATOS / MARKETING
            // =====================================================
            double variacaoNovos =
                CalcularPercentual(
                    dadosAnterior.NovosContatos,
                    dadosAtual.NovosContatos);

            linhas.Add(
                variacaoNovos >= 0
                    ? $"• Novos contatos: aumentaram {Math.Abs(variacaoNovos):N1}% " +
                      $"({dadosAnterior.NovosContatos:N0} → {dadosAtual.NovosContatos:N0})."
                    : $"• Novos contatos: diminuíram {Math.Abs(variacaoNovos):N1}% " +
                      $"({dadosAnterior.NovosContatos:N0} → {dadosAtual.NovosContatos:N0}).");

            if (dadosAnterior.MarketingTotal > 0 ||
                dadosAtual.MarketingTotal > 0)
            {
                double variacaoMarketing =
                    CalcularPercentual(
                        dadosAnterior.MarketingTotal,
                        dadosAtual.MarketingTotal);

                linhas.Add(
                    variacaoMarketing >= 0
                        ? $"• Marketing: contatos aumentaram {Math.Abs(variacaoMarketing):N1}% " +
                          $"({dadosAnterior.MarketingTotal:N0} → {dadosAtual.MarketingTotal:N0})."
                        : $"• Marketing: contatos diminuíram {Math.Abs(variacaoMarketing):N1}% " +
                          $"({dadosAnterior.MarketingTotal:N0} → {dadosAtual.MarketingTotal:N0}).");
            }

            // =====================================================
            // PONTOS DE ATENÇÃO
            // =====================================================
            if (dadosAnterior.Reagendamentos > 0 ||
                dadosAtual.Reagendamentos > 0)
            {
                double variacaoReag =
                    CalcularPercentual(
                        dadosAnterior.Reagendamentos,
                        dadosAtual.Reagendamentos);

                linhas.Add(
                    variacaoReag >= 0
                        ? $"• Reagendamentos: aumentaram {Math.Abs(variacaoReag):N1}% " +
                          $"({dadosAnterior.Reagendamentos:N0} → {dadosAtual.Reagendamentos:N0})."
                        : $"• Reagendamentos: diminuíram {Math.Abs(variacaoReag):N1}% " +
                          $"({dadosAnterior.Reagendamentos:N0} → {dadosAtual.Reagendamentos:N0}).");
            }

            if (dadosAnterior.Inatividade > 0 ||
                dadosAtual.Inatividade > 0)
            {
                double variacaoInat =
                    CalcularPercentual(
                        dadosAnterior.Inatividade,
                        dadosAtual.Inatividade);

                linhas.Add(
                    variacaoInat >= 0
                        ? $"• Inatividade: aumentou {Math.Abs(variacaoInat):N1}% " +
                          $"({dadosAnterior.Inatividade:N0} → {dadosAtual.Inatividade:N0})."
                        : $"• Inatividade: diminuiu {Math.Abs(variacaoInat):N1}% " +
                          $"({dadosAnterior.Inatividade:N0} → {dadosAtual.Inatividade:N0}).");
            }

            if (dadosAnterior.OrcamentoFormula > 0 ||
                dadosAtual.OrcamentoFormula > 0)
            {
                double variacaoOrc =
                    CalcularPercentual(
                        dadosAnterior.OrcamentoFormula,
                        dadosAtual.OrcamentoFormula);

                linhas.Add(
                    variacaoOrc >= 0
                        ? $"• Orçamentos de fórmula: aumentaram {Math.Abs(variacaoOrc):N1}% " +
                          $"({dadosAnterior.OrcamentoFormula:N0} → {dadosAtual.OrcamentoFormula:N0})."
                        : $"• Orçamentos de fórmula: diminuíram {Math.Abs(variacaoOrc):N1}% " +
                          $"({dadosAnterior.OrcamentoFormula:N0} → {dadosAtual.OrcamentoFormula:N0}).");
            }

            // =====================================================
            // EQUIPE - RESUMO EM QUANTIDADE
            // =====================================================
            var mudancasTmr =
                CompararTempos(
                    dadosAnterior.TMR,
                    dadosAtual.TMR,
                    60);

            int tmrPioraram =
                mudancasTmr.Count(x => x.DiferencaSegundos > 0);

            int tmrMelhoraram =
                mudancasTmr.Count(x => x.DiferencaSegundos < 0);

            if (mudancasTmr.Count > 0)
            {
                linhas.Add(
                    $"• TMR da equipe: {tmrMelhoraram} melhoraram e " +
                    $"{tmrPioraram} pioraram.");
            }

            var mudancasTme =
                CompararTempos(
                    dadosAnterior.TME,
                    dadosAtual.TME,
                    60);

            int tmePioraram =
                mudancasTme.Count(x => x.DiferencaSegundos > 0);

            int tmeMelhoraram =
                mudancasTme.Count(x => x.DiferencaSegundos < 0);

            if (mudancasTme.Count > 0)
            {
                linhas.Add(
                    $"• TME da equipe: {tmeMelhoraram} melhoraram e " +
                    $"{tmePioraram} pioraram.");
            }

            var mudancasNotas =
                CompararNotas(
                    dadosAnterior.Notas,
                    dadosAtual.Notas);

            int notasMelhoraram =
                mudancasNotas.Count(x => x.Diferenca >= 0.03);

            int notasPioraram =
                mudancasNotas.Count(x => x.Diferenca <= -0.03);

            int notasEstaveis =
                mudancasNotas.Count -
                notasMelhoraram -
                notasPioraram;

            if (mudancasNotas.Count > 0)
            {
                linhas.Add(
                    $"• Notas individuais: {notasMelhoraram} melhoraram, " +
                    $"{notasPioraram} pioraram e {notasEstaveis} ficaram estáveis.");
            }

            // =====================================================
            // NOTA GERAL
            // =====================================================
            if (dadosAnterior.NotaGeral > 0 &&
                dadosAtual.NotaGeral > 0)
            {
                double diferencaNota =
                    dadosAtual.NotaGeral -
                    dadosAnterior.NotaGeral;

                string leituraNota =
                    diferencaNota >= 0.03
                        ? "melhorou"
                        : diferencaNota <= -0.03
                            ? "caiu"
                            : "ficou estável";

                linhas.Add(
                    $"• Nota geral: {leituraNota} " +
                    $"({dadosAnterior.NotaGeral:N2} → {dadosAtual.NotaGeral:N2}).");
            }

            // =====================================================
            // LEITURA GERENCIAL CURTA
            // =====================================================
            if (variacaoConversas > variacaoFinalizados + 1.0)
            {
                linhas.Add(
                    "• Atenção: a demanda cresceu mais do que os atendimentos finalizados.");
            }

            return string.Join(
                Environment.NewLine,
                linhas);
        }

        private string GerarLeituraAtendente(
            double ant,
            double atual,
            TimeSpan? tmrAnt,
            TimeSpan? tmrAt,
            TimeSpan? tmeAnt,
            TimeSpan? tmeAt,
            double notaAnt,
            double notaAt)
        {
            var partes = new List<string>();

            // =====================================================
            // ATENDIMENTOS
            // =====================================================
            double pct = CalcularPercentual(ant, atual);

            if (ant == 0 && atual == 0)
            {
                partes.Add("atendimentos estáveis");
            }
            else if (pct >= 5)
            {
                partes.Add("atendimentos aumentaram");
            }
            else if (pct <= -5)
            {
                partes.Add("atendimentos diminuíram");
            }
            else
            {
                partes.Add("atendimentos estáveis");
            }

            // =====================================================
            // TMR
            // Menor tempo = melhora.
            // Diferenças menores que 1 minuto são tratadas como estáveis.
            // =====================================================
            if (tmrAnt.HasValue && tmrAt.HasValue)
            {
                double diferencaTmrSegundos =
                    (tmrAt.Value - tmrAnt.Value).TotalSeconds;

                if (diferencaTmrSegundos >= 60)
                    partes.Add("TMR piorou");
                else if (diferencaTmrSegundos <= -60)
                    partes.Add("TMR melhorou");
                else
                    partes.Add("TMR estável");
            }
            else
            {
                partes.Add("TMR sem comparação");
            }

            // =====================================================
            // TME
            // Menor tempo = melhora.
            // Diferenças menores que 1 minuto são tratadas como estáveis.
            // =====================================================
            if (tmeAnt.HasValue && tmeAt.HasValue)
            {
                double diferencaTmeSegundos =
                    (tmeAt.Value - tmeAnt.Value).TotalSeconds;

                if (diferencaTmeSegundos >= 60)
                    partes.Add("TME piorou");
                else if (diferencaTmeSegundos <= -60)
                    partes.Add("TME melhorou");
                else
                    partes.Add("TME estável");
            }
            else
            {
                partes.Add("TME sem comparação");
            }

            // =====================================================
            // NOTA
            // =====================================================
            if (notaAnt > 0 && notaAt > 0)
            {
                double diferencaNota = notaAt - notaAnt;

                if (diferencaNota >= 0.03)
                    partes.Add("nota melhorou");
                else if (diferencaNota <= -0.03)
                    partes.Add("nota piorou");
                else
                    partes.Add("nota estável");
            }
            else
            {
                partes.Add("nota sem comparação");
            }

            return string.Join(" / ", partes);
        }

        private List<string> GerarAlertasQualidade()
        {
            var alertas = new List<string>();

            if (dadosAnterior.TMA.Count > 0 && dadosAtual.TMA.Values.Count(x => x.HasValue) == 0)
                alertas.Add($"TMA existe em {dadosAnterior.Periodo}, mas está vazio em {dadosAtual.Periodo}.");

            if (dadosAtual.OrcamentoFormula > 0 &&
                dadosAtual.OrcamentoFormulaPossivelVenda > 0 &&
                Math.Abs(dadosAtual.OrcamentoFormula - dadosAtual.OrcamentoFormulaPossivelVenda) > 0.01)
            {
                alertas.Add(
                    $"\"Orçamento de fórmula\" aparece com dois valores em {dadosAtual.Periodo}: " +
                    $"{dadosAtual.OrcamentoFormula:N0} e {dadosAtual.OrcamentoFormulaPossivelVenda:N0}.");
            }

            if (dadosAtual.Finalizados > 0 && dadosAtual.Conversas > 0)
            {
                double taxa = dadosAtual.Finalizados / dadosAtual.Conversas * 100.0;
                if (taxa > 100)
                    alertas.Add("A quantidade de finalizados é maior que a quantidade de conversas. Verifique os totais.");
            }

            if (dadosAtual.NotaGeral == 0)
                alertas.Add("Não foi possível localizar a nota média geral do mês passado.");

            if (dadosAtual.TMR.Count == 0)
                alertas.Add("Não foi possível localizar os dados de TMR.");

            return alertas;
        }

        // ============================================================
        // EXPORTAR PDF
        // ============================================================
        // IMPORTANTE:
        // O PDF NÃO é uma captura da tela.
        // Ele é desenhado a partir dos dados carregados nas planilhas.
        // Por isso, tabelas e resumos podem continuar em novas páginas
        // e nenhuma informação fica escondida por scroll.
        private void ExportarTelasParaPdf(string opcao)
        {
            if (!comparacaoRealizada)
            {
                MessageBox.Show(
                    "Primeiro importe os dois arquivos e clique em COMPARAR.",
                    "Exportar PDF",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            string[] telas =
                opcao == "Todos"
                    ? new[]
                    {
                        "Comparativo",
                        "Atendentes",
                        "Motivos",
                        "Marketing"
                    }
                    : new[] { opcao };

            using var sfd = new SaveFileDialog
            {
                Title = "Salvar relatório em PDF",
                Filter = "Arquivo PDF (*.pdf)|*.pdf",
                DefaultExt = "pdf",
                AddExtension = true,
                FileName =
                    $"Relatorio_Atendimento_" +
                    $"{dadosAnterior.Periodo.Replace("/", "-")}_" +
                    $"{dadosAtual.Periodo.Replace("/", "-")}.pdf"
            };

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            try
            {
                Cursor = Cursors.WaitCursor;

                List<PaginaRelatorioPdf> paginas =
                    MontarPaginasRelatorioPdf(telas);

                ImprimirRelatorioPdf(
                    paginas,
                    sfd.FileName);

                MessageBox.Show(
                    "PDF gerado com sucesso.\n\n" +
                    "Todas as informações foram exportadas, " +
                    "inclusive os conteúdos que na tela precisam de scroll.",
                    "Exportar PDF",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Não foi possível gerar o PDF.\n\n" +
                    "Detalhe: " + ex.Message,
                    "Exportar PDF",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private List<PaginaRelatorioPdf> MontarPaginasRelatorioPdf(
            string[] telas)
        {
            var paginas =
                new List<PaginaRelatorioPdf>();

            foreach (string tela in telas)
            {
                switch (tela)
                {
                    case "Comparativo":
                        // Duas páginas para manter o visual limpo
                        // e trazer todo o resumo sem cortes.
                        paginas.Add(
                            new PaginaRelatorioPdf
                            {
                                Tipo = "Comparativo1"
                            });

                        paginas.Add(
                            new PaginaRelatorioPdf
                            {
                                Tipo = "Comparativo2"
                            });
                        break;

                    case "Atendentes":
                        {
                            int total =
                                dadosAnterior.AtendentesFinalizados.Keys
                                    .Union(
                                        dadosAtual.AtendentesFinalizados.Keys)
                                    .Count();

                            const int linhasPorPagina = 20;

                            if (total == 0)
                            {
                                paginas.Add(
                                    new PaginaRelatorioPdf
                                    {
                                        Tipo = "Atendentes",
                                        Inicio = 0,
                                        Quantidade = 0
                                    });
                            }
                            else
                            {
                                for (
                                    int inicio = 0;
                                    inicio < total;
                                    inicio += linhasPorPagina)
                                {
                                    paginas.Add(
                                        new PaginaRelatorioPdf
                                        {
                                            Tipo = "Atendentes",
                                            Inicio = inicio,
                                            Quantidade =
                                                Math.Min(
                                                    linhasPorPagina,
                                                    total - inicio)
                                        });
                                }
                            }

                            break;
                        }

                    case "Motivos":
                        {
                            int total =
                                dadosAnterior.Motivos.Keys
                                    .Union(dadosAtual.Motivos.Keys)
                                    .Count();

                            const int linhasPorPagina = 22;

                            if (total == 0)
                            {
                                paginas.Add(
                                    new PaginaRelatorioPdf
                                    {
                                        Tipo = "Motivos",
                                        Inicio = 0,
                                        Quantidade = 0
                                    });
                            }
                            else
                            {
                                for (
                                    int inicio = 0;
                                    inicio < total;
                                    inicio += linhasPorPagina)
                                {
                                    paginas.Add(
                                        new PaginaRelatorioPdf
                                        {
                                            Tipo = "Motivos",
                                            Inicio = inicio,
                                            Quantidade =
                                                Math.Min(
                                                    linhasPorPagina,
                                                    total - inicio)
                                        });
                                }
                            }

                            break;
                        }

                    case "Marketing":
                        {
                            int total =
                                dadosAnterior.MarketingCriativos.Keys
                                    .Union(
                                        dadosAtual.MarketingCriativos.Keys)
                                    .Count();

                            // A primeira página reserva espaço para os 3 KPIs.
                            const int primeiraPagina = 15;
                            const int proximasPaginas = 22;

                            if (total == 0)
                            {
                                paginas.Add(
                                    new PaginaRelatorioPdf
                                    {
                                        Tipo = "Marketing",
                                        Inicio = 0,
                                        Quantidade = 0,
                                        PrimeiraPagina = true
                                    });
                            }
                            else
                            {
                                int inicio = 0;

                                paginas.Add(
                                    new PaginaRelatorioPdf
                                    {
                                        Tipo = "Marketing",
                                        Inicio = 0,
                                        Quantidade =
                                            Math.Min(
                                                primeiraPagina,
                                                total),
                                        PrimeiraPagina = true
                                    });

                                inicio += primeiraPagina;

                                while (inicio < total)
                                {
                                    paginas.Add(
                                        new PaginaRelatorioPdf
                                        {
                                            Tipo = "Marketing",
                                            Inicio = inicio,
                                            Quantidade =
                                                Math.Min(
                                                    proximasPaginas,
                                                    total - inicio),
                                            PrimeiraPagina = false
                                        });

                                    inicio += proximasPaginas;
                                }
                            }

                            break;
                        }
                }
            }

            return paginas;
        }

        private void ImprimirRelatorioPdf(
            List<PaginaRelatorioPdf> paginas,
            string caminhoPdf)
        {
            if (paginas.Count == 0)
            {
                throw new InvalidOperationException(
                    "Nenhuma página foi preparada para o relatório.");
            }

            string impressoraPdf = "";

            foreach (
                string impressora
                in PrinterSettings.InstalledPrinters)
            {
                if (impressora.Equals(
                        "Microsoft Print to PDF",
                        StringComparison.OrdinalIgnoreCase))
                {
                    impressoraPdf = impressora;
                    break;
                }

                if (string.IsNullOrWhiteSpace(impressoraPdf) &&
                    impressora.Contains(
                        "Print to PDF",
                        StringComparison.OrdinalIgnoreCase))
                {
                    impressoraPdf = impressora;
                }
            }

            if (string.IsNullOrWhiteSpace(impressoraPdf))
            {
                throw new InvalidOperationException(
                    "A impressora virtual 'Microsoft Print to PDF' " +
                    "não foi encontrada no Windows.");
            }

            int indicePagina = 0;

            using var documento =
                new PrintDocument();

            documento.PrinterSettings.PrinterName =
                impressoraPdf;

            documento.PrinterSettings.PrintToFile =
                true;

            documento.PrinterSettings.PrintFileName =
                caminhoPdf;

            // PDF em A4 RETRATO.
            documento.DefaultPageSettings.Landscape =
                false;

            documento.DefaultPageSettings.PaperSize =
                new PaperSize(
                    "A4",
                    827,
                    1169);

            documento.DefaultPageSettings.Margins =
                new Margins(
                    28,
                    28,
                    28,
                    28);

            documento.PrintController =
                new StandardPrintController();

            documento.PrintPage += (_, e) =>
            {
                PaginaRelatorioPdf pagina =
                    paginas[indicePagina];

                DesenharPaginaRelatorio(
                    e.Graphics,
                    e.MarginBounds,
                    pagina,
                    indicePagina + 1,
                    paginas.Count);

                indicePagina++;

                e.HasMorePages =
                    indicePagina < paginas.Count;
            };

            documento.Print();
        }

        private void DesenharPaginaRelatorio(
            Graphics g,
            Rectangle area,
            PaginaRelatorioPdf pagina,
            int numeroPagina,
            int totalPaginas)
        {
            g.Clear(Color.White);

            switch (pagina.Tipo)
            {
                case "Comparativo1":
                    DesenharComparativoPdfPagina1(
                        g,
                        area);
                    break;

                case "Comparativo2":
                    DesenharComparativoPdfPagina2(
                        g,
                        area);
                    break;

                case "Atendentes":
                    DesenharAtendentesPdf(
                        g,
                        area,
                        pagina.Inicio,
                        pagina.Quantidade);
                    break;

                case "Motivos":
                    DesenharMotivosPdf(
                        g,
                        area,
                        pagina.Inicio,
                        pagina.Quantidade);
                    break;

                case "Marketing":
                    DesenharMarketingPdf(
                        g,
                        area,
                        pagina.Inicio,
                        pagina.Quantidade,
                        pagina.PrimeiraPagina);
                    break;
            }

            DesenharRodapePdf(
                g,
                area,
                numeroPagina,
                totalPaginas);
        }

        private int DesenharCabecalhoPdf(
            Graphics g,
            Rectangle area,
            string titulo,
            string subtitulo)
        {
            using var fonteTitulo =
                new Font(
                    "Segoe UI",
                    16F,
                    FontStyle.Bold);

            using var fonteSub =
                new Font(
                    "Segoe UI",
                    8.5F);

            using var fontePeriodo =
                new Font(
                    "Segoe UI Semibold",
                    8.5F);

            using var brushTitulo =
                new SolidBrush(CorTexto);

            using var brushSub =
                new SolidBrush(CorTextoSecundario);

            g.DrawString(
                titulo,
                fonteTitulo,
                brushTitulo,
                area.Left,
                area.Top);

            g.DrawString(
                subtitulo,
                fonteSub,
                brushSub,
                area.Left,
                area.Top + 31);

            string periodo =
                $"{dadosAnterior.Periodo}  x  {dadosAtual.Periodo}";

            SizeF tamPeriodo =
                g.MeasureString(
                    periodo,
                    fontePeriodo);

            g.DrawString(
                periodo,
                fontePeriodo,
                brushTitulo,
                area.Right -
                tamPeriodo.Width,
                area.Top + 8);

            return area.Top + 55;
        }

        private void DesenharRodapePdf(
            Graphics g,
            Rectangle area,
            int paginaAtual,
            int totalPaginas)
        {
            using var fonte =
                new Font(
                    "Segoe UI",
                    7.5F);

            using var brush =
                new SolidBrush(
                    CorTextoSecundario);

            string texto =
                $"Relatório de Atendimento  •  " +
                $"{dadosAnterior.Periodo} x {dadosAtual.Periodo}  •  " +
                $"Página {paginaAtual} de {totalPaginas}";

            g.DrawString(
                texto,
                fonte,
                brush,
                area.Left,
                area.Bottom - 12);
        }

        private void DesenharComparativoPdfPagina1(
            Graphics g,
            Rectangle area)
        {
            int y =
                DesenharCabecalhoPdf(
                    g,
                    area,
                    "Relatório de Atendimento - Comparativo Mensal",
                    "Visão geral dos principais indicadores e tempos da equipe.");

            int espaco = 8;
            int larguraCard =
                (area.Width - espaco) / 2;

            // =====================================================
            // KPIs - 2 x 2 para ficar legível no A4 retrato
            // =====================================================
            int alturaKpi = 74;

            double varConversas =
                CalcularPercentual(
                    dadosAnterior.Conversas,
                    dadosAtual.Conversas);

            double varFinalizados =
                CalcularPercentual(
                    dadosAnterior.Finalizados,
                    dadosAtual.Finalizados);

            double varNovos =
                CalcularPercentual(
                    dadosAnterior.NovosContatos,
                    dadosAtual.NovosContatos);

            double difNota =
                dadosAtual.NotaGeral -
                dadosAnterior.NotaGeral;

            DesenharKpiPdf(
                g,
                new Rectangle(
                    area.Left,
                    y,
                    larguraCard,
                    alturaKpi),
                "Conversas",
                dadosAtual.Conversas.ToString("N0"),
                $"{FormatarPercentualComSeta(varConversas)}  " +
                $"(antes {dadosAnterior.Conversas:N0})",
                varConversas >= 0
                    ? CorVerde
                    : CorVermelho);

            DesenharKpiPdf(
                g,
                new Rectangle(
                    area.Left +
                    larguraCard +
                    espaco,
                    y,
                    larguraCard,
                    alturaKpi),
                "Finalizados",
                dadosAtual.Finalizados.ToString("N0"),
                $"{FormatarPercentualComSeta(varFinalizados)}  " +
                $"(antes {dadosAnterior.Finalizados:N0})",
                varFinalizados >= 0
                    ? CorVerde
                    : CorVermelho);

            y += alturaKpi + espaco;

            DesenharKpiPdf(
                g,
                new Rectangle(
                    area.Left,
                    y,
                    larguraCard,
                    alturaKpi),
                "Novos contatos",
                dadosAtual.NovosContatos.ToString("N0"),
                $"{FormatarPercentualComSeta(varNovos)}  " +
                $"(antes {dadosAnterior.NovosContatos:N0})",
                varNovos >= 0
                    ? CorVerde
                    : CorVermelho);

            DesenharKpiPdf(
                g,
                new Rectangle(
                    area.Left +
                    larguraCard +
                    espaco,
                    y,
                    larguraCard,
                    alturaKpi),
                "Nota geral",
                dadosAtual.NotaGeral > 0
                    ? dadosAtual.NotaGeral.ToString("N2")
                    : "—",
                $"{(difNota >= 0 ? "▲" : "▼")} " +
                $"{Math.Abs(difNota):N2}  " +
                $"(antes {dadosAnterior.NotaGeral:N2})",
                difNota >= 0
                    ? CorVerde
                    : CorVermelho);

            y += alturaKpi + 12;

            // =====================================================
            // TMR - cards em largura total
            // =====================================================
            var mudancasTmr =
                CompararTempos(
                    dadosAnterior.TMR,
                    dadosAtual.TMR,
                    60);

            string tmrPioraram =
                JuntarNomes(
                    mudancasTmr
                        .Where(
                            x =>
                                x.DiferencaSegundos > 0)
                        .OrderByDescending(
                            x =>
                                x.DiferencaSegundos)
                        .Select(
                            x =>
                                NomeCurto(x.Nome))
                        .ToList());

            string tmrMelhoraram =
                JuntarNomes(
                    mudancasTmr
                        .Where(
                            x =>
                                x.DiferencaSegundos < 0)
                        .OrderBy(
                            x =>
                                x.DiferencaSegundos)
                        .Select(
                            x =>
                                NomeCurto(x.Nome))
                        .ToList());

            int alturaTempo = 112;

            DesenharCardTextoPdf(
                g,
                new Rectangle(
                    area.Left,
                    y,
                    area.Width,
                    alturaTempo),
                CorVermelhoClaro,
                CorVermelho,
                "TMR PIOROU",
                string.IsNullOrWhiteSpace(tmrPioraram)
                    ? "Nenhuma piora relevante."
                    : "Pioraram: " + tmrPioraram + ".",
                8.1F);

            y += alturaTempo + 8;

            DesenharCardTextoPdf(
                g,
                new Rectangle(
                    area.Left,
                    y,
                    area.Width,
                    alturaTempo),
                CorVerdeClaro,
                CorVerde,
                "TMR MELHOROU",
                string.IsNullOrWhiteSpace(tmrMelhoraram)
                    ? "Nenhuma melhora relevante."
                    : "Melhoraram: " + tmrMelhoraram + ".",
                8.1F);

            y += alturaTempo + 12;

            // =====================================================
            // TME - cards em largura total
            // =====================================================
            var mudancasTme =
                CompararTempos(
                    dadosAnterior.TME,
                    dadosAtual.TME,
                    60);

            string tmePioraram =
                JuntarNomes(
                    mudancasTme
                        .Where(
                            x =>
                                x.DiferencaSegundos > 0)
                        .OrderByDescending(
                            x =>
                                x.DiferencaSegundos)
                        .Select(
                            x =>
                                NomeCurto(x.Nome))
                        .ToList());

            string tmeMelhoraram =
                JuntarNomes(
                    mudancasTme
                        .Where(
                            x =>
                                x.DiferencaSegundos < 0)
                        .OrderBy(
                            x =>
                                x.DiferencaSegundos)
                        .Select(
                            x =>
                                NomeCurto(x.Nome))
                        .ToList());

            DesenharCardTextoPdf(
                g,
                new Rectangle(
                    area.Left,
                    y,
                    area.Width,
                    alturaTempo),
                CorVermelhoClaro,
                CorVermelho,
                "TME PIOROU",
                string.IsNullOrWhiteSpace(tmePioraram)
                    ? "Nenhuma piora relevante."
                    : "Pioraram: " + tmePioraram + ".",
                8.1F);

            y += alturaTempo + 8;

            DesenharCardTextoPdf(
                g,
                new Rectangle(
                    area.Left,
                    y,
                    area.Width,
                    alturaTempo),
                CorVerdeClaro,
                CorVerde,
                "TME MELHOROU",
                string.IsNullOrWhiteSpace(tmeMelhoraram)
                    ? "Nenhuma melhora relevante."
                    : "Melhoraram: " + tmeMelhoraram + ".",
                8.1F);
        }

        private void DesenharComparativoPdfPagina2(
            Graphics g,
            Rectangle area)
        {
            int y =
                DesenharCabecalhoPdf(
                    g,
                    area,
                    "Comparativo - Equipe e Resumo do Mês",
                    "Mudanças de volume, notas e leitura gerencial.");

            var atendimentos =
                CompararNumeros(
                    dadosAnterior.AtendentesFinalizados,
                    dadosAtual.AtendentesFinalizados);

            string aumentaram =
                JuntarNomes(
                    atendimentos
                        .Where(
                            x =>
                                x.Percentual >= 5)
                        .OrderByDescending(
                            x =>
                                x.Percentual)
                        .Select(
                            x =>
                                NomeCurto(x.Nome))
                        .ToList());

            string diminuiram =
                JuntarNomes(
                    atendimentos
                        .Where(
                            x =>
                                x.Percentual <= -5)
                        .OrderBy(
                            x =>
                                x.Percentual)
                        .Select(
                            x =>
                                NomeCurto(x.Nome))
                        .ToList());

            string textoAtendimentos =
                (string.IsNullOrWhiteSpace(aumentaram)
                    ? ""
                    : "Aumentaram: " +
                      aumentaram +
                      ".") +
                (string.IsNullOrWhiteSpace(diminuiram)
                    ? ""
                    : Environment.NewLine +
                      "Diminuíram: " +
                      diminuiram +
                      ".");

            if (string.IsNullOrWhiteSpace(
                    textoAtendimentos))
            {
                textoAtendimentos =
                    "Sem mudanças relevantes.";
            }

            var notas =
                CompararNotas(
                    dadosAnterior.Notas,
                    dadosAtual.Notas);

            string notasMelhoraram =
                JuntarNomes(
                    notas
                        .Where(
                            x =>
                                x.Diferenca >= 0.03)
                        .OrderByDescending(
                            x =>
                                x.Diferenca)
                        .Select(
                            x =>
                                NomeCurto(x.Nome))
                        .ToList());

            string notasPioraram =
                JuntarNomes(
                    notas
                        .Where(
                            x =>
                                x.Diferenca <= -0.03)
                        .OrderBy(
                            x =>
                                x.Diferenca)
                        .Select(
                            x =>
                                NomeCurto(x.Nome))
                        .ToList());

            string textoNotas =
                (string.IsNullOrWhiteSpace(
                    notasMelhoraram)
                    ? ""
                    : "Melhoraram: " +
                      notasMelhoraram +
                      ".") +
                (string.IsNullOrWhiteSpace(
                    notasPioraram)
                    ? ""
                    : Environment.NewLine +
                      "Pioraram: " +
                      notasPioraram +
                      ".");

            if (string.IsNullOrWhiteSpace(
                    textoNotas))
            {
                textoNotas =
                    "As notas ficaram estáveis.";
            }

            int alturaEquipe = 150;

            DesenharCardTextoPdf(
                g,
                new Rectangle(
                    area.Left,
                    y,
                    area.Width,
                    alturaEquipe),
                CorAzulClaro,
                CorAzul,
                "ATENDIMENTOS",
                textoAtendimentos,
                8.2F);

            y += alturaEquipe + 10;

            DesenharCardTextoPdf(
                g,
                new Rectangle(
                    area.Left,
                    y,
                    area.Width,
                    alturaEquipe),
                CorAmareloClaro,
                CorAmarelo,
                "NOTAS",
                textoNotas,
                8.2F);

            y += alturaEquipe + 12;

            int alturaResumo =
                area.Bottom -
                y -
                25;

            DesenharCardTextoPdf(
                g,
                new Rectangle(
                    area.Left,
                    y,
                    area.Width,
                    Math.Max(
                        280,
                        alturaResumo)),
                Color.White,
                CorSidebar,
                "RESUMO DO MÊS",
                GerarResumoCurto(),
                8.1F);
        }

        private void DesenharAtendentesPdf(
            Graphics g,
            Rectangle area,
            int inicio,
            int quantidade)
        {
            int y =
                DesenharCabecalhoPdf(
                    g,
                    area,
                    "Atendentes",
                    "Comparação detalhada por atendente.");

            var nomes =
                dadosAnterior.AtendentesFinalizados.Keys
                    .Union(
                        dadosAtual.AtendentesFinalizados.Keys)
                    .OrderBy(
                        x =>
                            NomeCurto(x))
                    .ToList();

            // Tabela inspirada diretamente no modal do programa.
            // As larguras foram ajustadas para caber em A4 retrato.
            var colunas =
                new[]
                {
                    new ColunaTabelaPdf
                    {
                        Titulo = "Atendente",
                        Peso = 1.35F
                    },
                    new ColunaTabelaPdf
                    {
                        Titulo = dadosAnterior.Periodo,
                        Peso = 0.55F
                    },
                    new ColunaTabelaPdf
                    {
                        Titulo = dadosAtual.Periodo,
                        Peso = 0.55F
                    },
                    new ColunaTabelaPdf
                    {
                        Titulo = "Variação",
                        Peso = 0.58F
                    },
                    new ColunaTabelaPdf
                    {
                        Titulo = "TMR ant.",
                        Peso = 0.64F
                    },
                    new ColunaTabelaPdf
                    {
                        Titulo = "TMR atual",
                        Peso = 0.64F
                    },
                    new ColunaTabelaPdf
                    {
                        Titulo = "TME ant.",
                        Peso = 0.64F
                    },
                    new ColunaTabelaPdf
                    {
                        Titulo = "TME atual",
                        Peso = 0.64F
                    },
                    new ColunaTabelaPdf
                    {
                        Titulo = "Nota ant.",
                        Peso = 0.50F
                    },
                    new ColunaTabelaPdf
                    {
                        Titulo = "Nota atual",
                        Peso = 0.50F
                    },
                    new ColunaTabelaPdf
                    {
                        Titulo = "Leitura rápida",
                        Peso = 2.25F
                    }
                };

            int[] larguras =
                CalcularLargurasTabela(
                    area.Width,
                    colunas);

            const int alturaCabecalho = 28;
            const int alturaLinha = 42;

            y =
                DesenharCabecalhoTabelaPdfCompacto(
                    g,
                    area.Left,
                    y,
                    larguras,
                    colunas,
                    alturaCabecalho);

            int fim =
                Math.Min(
                    inicio + quantidade,
                    nomes.Count);

            for (
                int i = inicio;
                i < fim;
                i++)
            {
                string chave =
                    nomes[i];

                double ant =
                    Valor(
                        dadosAnterior.AtendentesFinalizados,
                        chave);

                double atual =
                    Valor(
                        dadosAtual.AtendentesFinalizados,
                        chave);

                double pct =
                    CalcularPercentual(
                        ant,
                        atual);

                TimeSpan? tmrAnt =
                    Tempo(
                        dadosAnterior.TMR,
                        chave);

                TimeSpan? tmrAtual =
                    Tempo(
                        dadosAtual.TMR,
                        chave);

                TimeSpan? tmeAnt =
                    Tempo(
                        dadosAnterior.TME,
                        chave);

                TimeSpan? tmeAtual =
                    Tempo(
                        dadosAtual.TME,
                        chave);

                double notaAnt =
                    Valor(
                        dadosAnterior.Notas,
                        chave);

                double notaAtual =
                    Valor(
                        dadosAtual.Notas,
                        chave);

                string leitura =
                    GerarLeituraAtendente(
                        ant,
                        atual,
                        tmrAnt,
                        tmrAtual,
                        tmeAnt,
                        tmeAtual,
                        notaAnt,
                        notaAtual);

                string[] valores =
                {
                    NomeExibicao(chave),
                    ant.ToString("N0"),
                    atual.ToString("N0"),
                    FormatarPercentualComSeta(pct),
                    FormatarTempo(tmrAnt),
                    FormatarTempo(tmrAtual),
                    FormatarTempo(tmeAnt),
                    FormatarTempo(tmeAtual),
                    notaAnt > 0
                        ? notaAnt.ToString("N2")
                        : "—",
                    notaAtual > 0
                        ? notaAtual.ToString("N2")
                        : "—",
                    leitura
                };

                y =
                    DesenharLinhaTabelaPdf(
                        g,
                        area.Left,
                        y,
                        larguras,
                        valores,
                        alturaLinha,
                        i % 2 == 0
                            ? Color.White
                            : Color.FromArgb(
                                247,
                                249,
                                251),
                        5.6F);
            }
        }

        private int DesenharCabecalhoTabelaPdfCompacto(
            Graphics g,
            int x,
            int y,
            int[] larguras,
            ColunaTabelaPdf[] colunas,
            int altura)
        {
            using var brush =
                new SolidBrush(
                    Color.FromArgb(
                        233,
                        238,
                        243));

            using var pen =
                new Pen(
                    Color.FromArgb(
                        205,
                        212,
                        220));

            using var fonte =
                new Font(
                    "Segoe UI Semibold",
                    5.7F,
                    FontStyle.Bold);

            using var brushTexto =
                new SolidBrush(CorTexto);

            int atualX = x;

            for (
                int i = 0;
                i < colunas.Length;
                i++)
            {
                Rectangle rect =
                    new Rectangle(
                        atualX,
                        y,
                        larguras[i],
                        altura);

                g.FillRectangle(
                    brush,
                    rect);

                g.DrawRectangle(
                    pen,
                    rect);

                var areaTexto =
                    new RectangleF(
                        rect.Left + 3,
                        rect.Top + 4,
                        rect.Width - 6,
                        rect.Height - 8);

                using var formato =
                    new StringFormat
                    {
                        Trimming =
                            StringTrimming.EllipsisWord,
                        FormatFlags =
                            StringFormatFlags.LineLimit
                    };

                g.DrawString(
                    colunas[i].Titulo,
                    fonte,
                    brushTexto,
                    areaTexto,
                    formato);

                atualX +=
                    larguras[i];
            }

            return y + altura;
        }

        private void DesenharMotivosPdf(
            Graphics g,
            Rectangle area,
            int inicio,
            int quantidade)
        {
            int y =
                DesenharCabecalhoPdf(
                    g,
                    area,
                    "Motivos de atendimento",
                    "Comparação completa dos motivos entre os dois meses.");

            var motivos =
                dadosAnterior.Motivos.Keys
                    .Union(
                        dadosAtual.Motivos.Keys)
                    .OrderByDescending(
                        x =>
                            Math.Max(
                                Valor(
                                    dadosAnterior.Motivos,
                                    x),
                                Valor(
                                    dadosAtual.Motivos,
                                    x)))
                    .ToList();

            var colunas =
                new[]
                {
                    new ColunaTabelaPdf
                    {
                        Titulo = "Motivo",
                        Peso = 3.4F
                    },
                    new ColunaTabelaPdf
                    {
                        Titulo = dadosAnterior.Periodo,
                        Peso = 1F
                    },
                    new ColunaTabelaPdf
                    {
                        Titulo = dadosAtual.Periodo,
                        Peso = 1F
                    },
                    new ColunaTabelaPdf
                    {
                        Titulo = "Diferença",
                        Peso = 1F
                    },
                    new ColunaTabelaPdf
                    {
                        Titulo = "Variação",
                        Peso = 1F
                    },
                    new ColunaTabelaPdf
                    {
                        Titulo = "Leitura",
                        Peso = 1.3F
                    }
                };

            int[] larguras =
                CalcularLargurasTabela(
                    area.Width,
                    colunas);

            const int alturaCabecalho = 28;
            const int alturaLinha = 25;

            y =
                DesenharCabecalhoTabelaPdf(
                    g,
                    area.Left,
                    y,
                    larguras,
                    colunas,
                    alturaCabecalho);

            int fim =
                Math.Min(
                    inicio + quantidade,
                    motivos.Count);

            for (
                int i = inicio;
                i < fim;
                i++)
            {
                string motivo =
                    motivos[i];

                double ant =
                    Valor(
                        dadosAnterior.Motivos,
                        motivo);

                double atual =
                    Valor(
                        dadosAtual.Motivos,
                        motivo);

                double dif =
                    atual - ant;

                double pct =
                    CalcularPercentual(
                        ant,
                        atual);

                string[] valores =
                {
                    motivo,
                    ant.ToString("N0"),
                    atual.ToString("N0"),
                    $"{dif:+0;-0;0}",
                    FormatarPercentualComSeta(pct),
                    dif > 0
                        ? "Aumentou"
                        : dif < 0
                            ? "Reduziu"
                            : "Estável"
                };

                y =
                    DesenharLinhaTabelaPdf(
                        g,
                        area.Left,
                        y,
                        larguras,
                        valores,
                        alturaLinha,
                        i % 2 == 0
                            ? Color.White
                            : Color.FromArgb(
                                247,
                                249,
                                251),
                        7.4F);
            }
        }

        private void DesenharMarketingPdf(
            Graphics g,
            Rectangle area,
            int inicio,
            int quantidade,
            bool primeiraPagina)
        {
            int y =
                DesenharCabecalhoPdf(
                    g,
                    area,
                    "Marketing",
                    "Comparação dos contatos e criativos de marketing.");

            if (primeiraPagina)
            {
                int espaco = 8;

                int larguraCard =
                    (area.Width -
                     espaco * 2) / 3;

                int alturaCard = 74;

                DesenharKpiPdf(
                    g,
                    new Rectangle(
                        area.Left,
                        y,
                        larguraCard,
                        alturaCard),
                    "Contatos de marketing",
                    $"{dadosAnterior.MarketingTotal:N0} → " +
                    $"{dadosAtual.MarketingTotal:N0}",
                    FormatarPercentualComSeta(
                        CalcularPercentual(
                            dadosAnterior.MarketingTotal,
                            dadosAtual.MarketingTotal)),
                    CalcularPercentual(
                        dadosAnterior.MarketingTotal,
                        dadosAtual.MarketingTotal) >= 0
                            ? CorVerde
                            : CorVermelho);

                DesenharKpiPdf(
                    g,
                    new Rectangle(
                        area.Left +
                        larguraCard +
                        espaco,
                        y,
                        larguraCard,
                        alturaCard),
                    "Novos contatos",
                    $"{dadosAnterior.NovosContatos:N0} → " +
                    $"{dadosAtual.NovosContatos:N0}",
                    FormatarPercentualComSeta(
                        CalcularPercentual(
                            dadosAnterior.NovosContatos,
                            dadosAtual.NovosContatos)),
                    CalcularPercentual(
                        dadosAnterior.NovosContatos,
                        dadosAtual.NovosContatos) >= 0
                            ? CorVerde
                            : CorVermelho);

                DesenharKpiPdf(
                    g,
                    new Rectangle(
                        area.Left +
                        (larguraCard + espaco) * 2,
                        y,
                        larguraCard,
                        alturaCard),
                    "Reagendamentos",
                    $"{dadosAnterior.Reagendamentos:N0} → " +
                    $"{dadosAtual.Reagendamentos:N0}",
                    FormatarPercentualComSeta(
                        CalcularPercentual(
                            dadosAnterior.Reagendamentos,
                            dadosAtual.Reagendamentos)),
                    CalcularPercentual(
                        dadosAnterior.Reagendamentos,
                        dadosAtual.Reagendamentos) <= 0
                            ? CorVerde
                            : CorVermelho);

                y += alturaCard + 12;
            }

            var campanhas =
                dadosAnterior.MarketingCriativos.Keys
                    .Union(
                        dadosAtual.MarketingCriativos.Keys)
                    .OrderByDescending(
                        x =>
                            Math.Max(
                                Valor(
                                    dadosAnterior.MarketingCriativos,
                                    x),
                                Valor(
                                    dadosAtual.MarketingCriativos,
                                    x)))
                    .ToList();

            var colunas =
                new[]
                {
                    new ColunaTabelaPdf
                    {
                        Titulo = "Criativo / campanha",
                        Peso = 3.4F
                    },
                    new ColunaTabelaPdf
                    {
                        Titulo = dadosAnterior.Periodo,
                        Peso = 1F
                    },
                    new ColunaTabelaPdf
                    {
                        Titulo = dadosAtual.Periodo,
                        Peso = 1F
                    },
                    new ColunaTabelaPdf
                    {
                        Titulo = "Variação",
                        Peso = 1F
                    }
                };

            int[] larguras =
                CalcularLargurasTabela(
                    area.Width,
                    colunas);

            y =
                DesenharCabecalhoTabelaPdf(
                    g,
                    area.Left,
                    y,
                    larguras,
                    colunas,
                    28);

            int fim =
                Math.Min(
                    inicio + quantidade,
                    campanhas.Count);

            for (
                int i = inicio;
                i < fim;
                i++)
            {
                string campanha =
                    campanhas[i];

                double ant =
                    Valor(
                        dadosAnterior.MarketingCriativos,
                        campanha);

                double atual =
                    Valor(
                        dadosAtual.MarketingCriativos,
                        campanha);

                double pct =
                    CalcularPercentual(
                        ant,
                        atual);

                string[] valores =
                {
                    campanha,
                    ant.ToString("N0"),
                    atual.ToString("N0"),
                    FormatarPercentualComSeta(pct)
                };

                y =
                    DesenharLinhaTabelaPdf(
                        g,
                        area.Left,
                        y,
                        larguras,
                        valores,
                        25,
                        i % 2 == 0
                            ? Color.White
                            : Color.FromArgb(
                                247,
                                249,
                                251),
                        7.4F);
            }
        }

        private void DesenharKpiPdf(
            Graphics g,
            Rectangle rect,
            string titulo,
            string valor,
            string variacao,
            Color corVariacao)
        {
            using var brushFundo =
                new SolidBrush(Color.White);

            using var fonteTitulo =
                new Font(
                    "Segoe UI",
                    7.5F);

            using var fonteValor =
                new Font(
                    "Segoe UI Semibold",
                    13F,
                    FontStyle.Bold);

            using var fonteVar =
                new Font(
                    "Segoe UI Semibold",
                    7.2F);

            using var brushTitulo =
                new SolidBrush(
                    CorTextoSecundario);

            using var brushValor =
                new SolidBrush(CorTexto);

            using var brushVar =
                new SolidBrush(corVariacao);

            g.FillRectangle(
                brushFundo,
                rect);

            g.DrawString(
                titulo,
                fonteTitulo,
                brushTitulo,
                rect.Left + 10,
                rect.Top + 8);

            g.DrawString(
                valor,
                fonteValor,
                brushValor,
                rect.Left + 10,
                rect.Top + 27);

            g.DrawString(
                variacao,
                fonteVar,
                brushVar,
                rect.Left + 10,
                rect.Top + 56);
        }

        private void DesenharCardTextoPdf(
            Graphics g,
            Rectangle rect,
            Color fundo,
            Color corTitulo,
            string titulo,
            string texto,
            float tamanhoTexto = 7.8F)
        {
            using var brushFundo =
                new SolidBrush(fundo);

            using var fonteTitulo =
                new Font(
                    "Segoe UI Semibold",
                    9.5F,
                    FontStyle.Bold);

            using var fonteTexto =
                new Font(
                    "Segoe UI",
                    tamanhoTexto);

            using var brushTitulo =
                new SolidBrush(corTitulo);

            using var brushTexto =
                new SolidBrush(CorTexto);

            g.FillRectangle(
                brushFundo,
                rect);

            g.DrawString(
                titulo,
                fonteTitulo,
                brushTitulo,
                rect.Left + 12,
                rect.Top + 10);

            var areaTexto =
                new RectangleF(
                    rect.Left + 12,
                    rect.Top + 34,
                    rect.Width - 24,
                    rect.Height - 44);

            using var formato =
                new StringFormat
                {
                    Trimming =
                        StringTrimming.Word,
                    FormatFlags =
                        StringFormatFlags.LineLimit
                };

            g.DrawString(
                texto,
                fonteTexto,
                brushTexto,
                areaTexto,
                formato);
        }

        private int DesenharCabecalhoTabelaPdf(
            Graphics g,
            int x,
            int y,
            int[] larguras,
            ColunaTabelaPdf[] colunas,
            int altura)
        {
            using var brush =
                new SolidBrush(
                    Color.FromArgb(
                        233,
                        238,
                        243));

            using var pen =
                new Pen(
                    Color.FromArgb(
                        205,
                        212,
                        220));

            using var fonte =
                new Font(
                    "Segoe UI Semibold",
                    7F,
                    FontStyle.Bold);

            using var brushTexto =
                new SolidBrush(CorTexto);

            int atualX = x;

            for (
                int i = 0;
                i < colunas.Length;
                i++)
            {
                Rectangle rect =
                    new Rectangle(
                        atualX,
                        y,
                        larguras[i],
                        altura);

                g.FillRectangle(
                    brush,
                    rect);

                g.DrawRectangle(
                    pen,
                    rect);

                var areaTexto =
                    new RectangleF(
                        rect.Left + 4,
                        rect.Top + 5,
                        rect.Width - 8,
                        rect.Height - 8);

                g.DrawString(
                    colunas[i].Titulo,
                    fonte,
                    brushTexto,
                    areaTexto);

                atualX +=
                    larguras[i];
            }

            return y + altura;
        }

        private int DesenharLinhaTabelaPdf(
            Graphics g,
            int x,
            int y,
            int[] larguras,
            string[] valores,
            int altura,
            Color fundo,
            float tamanhoFonte)
        {
            using var brushFundo =
                new SolidBrush(fundo);

            using var pen =
                new Pen(
                    Color.FromArgb(
                        220,
                        225,
                        230));

            using var fonte =
                new Font(
                    "Segoe UI",
                    tamanhoFonte);

            using var brushTexto =
                new SolidBrush(CorTexto);

            int atualX = x;

            for (
                int i = 0;
                i < larguras.Length;
                i++)
            {
                Rectangle rect =
                    new Rectangle(
                        atualX,
                        y,
                        larguras[i],
                        altura);

                g.FillRectangle(
                    brushFundo,
                    rect);

                g.DrawRectangle(
                    pen,
                    rect);

                var textoRect =
                    new RectangleF(
                        rect.Left + 4,
                        rect.Top + 4,
                        rect.Width - 8,
                        rect.Height - 8);

                using var formato =
                    new StringFormat
                    {
                        Trimming =
                            StringTrimming.EllipsisWord,
                        FormatFlags =
                            StringFormatFlags.LineLimit
                    };

                g.DrawString(
                    i < valores.Length
                        ? valores[i]
                        : "",
                    fonte,
                    brushTexto,
                    textoRect,
                    formato);

                atualX +=
                    larguras[i];
            }

            return y + altura;
        }

        private int[] CalcularLargurasTabela(
            int larguraTotal,
            ColunaTabelaPdf[] colunas)
        {
            float pesoTotal =
                colunas.Sum(
                    x =>
                        x.Peso);

            var larguras =
                new int[colunas.Length];

            int usado = 0;

            for (
                int i = 0;
                i < colunas.Length;
                i++)
            {
                if (i ==
                    colunas.Length - 1)
                {
                    larguras[i] =
                        larguraTotal - usado;
                }
                else
                {
                    larguras[i] =
                        (int)Math.Round(
                            larguraTotal *
                            colunas[i].Peso /
                            pesoTotal);

                    usado +=
                        larguras[i];
                }
            }

            return larguras;
        }

        // ============================================================
        // EXPORTAR EXCEL
        // ============================================================
        private void ExportarExcel()
        {
            if (!comparacaoRealizada)
                return;

            using var sfd = new SaveFileDialog
            {
                Title = "Salvar comparativo",
                Filter = "Excel (*.xlsx)|*.xlsx",
                FileName = $"Comparativo_{dadosAnterior.Periodo.Replace("/", "-")}_{dadosAtual.Periodo.Replace("/", "-")}.xlsx"
            };

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            try
            {
                using var wb = new XLWorkbook();

                var resumo = wb.Worksheets.Add("Resumo");
                resumo.Cell("A1").Value = "Relatório de Atendimento - Comparativo";
                resumo.Cell("A2").Value = $"{dadosAnterior.Periodo} x {dadosAtual.Periodo}";
                resumo.Cell("A4").Value = "Indicador";
                resumo.Cell("B4").Value = dadosAnterior.Periodo;
                resumo.Cell("C4").Value = dadosAtual.Periodo;
                resumo.Cell("D4").Value = "Variação";

                var itens = new[]
                {
                    ("Conversas", dadosAnterior.Conversas, dadosAtual.Conversas),
                    ("Finalizados", dadosAnterior.Finalizados, dadosAtual.Finalizados),
                    ("Novos contatos", dadosAnterior.NovosContatos, dadosAtual.NovosContatos),
                    ("Reagendamentos", dadosAnterior.Reagendamentos, dadosAtual.Reagendamentos),
                    ("Inatividade", dadosAnterior.Inatividade, dadosAtual.Inatividade),
                    ("Marketing", dadosAnterior.MarketingTotal, dadosAtual.MarketingTotal),
                    ("Templates", dadosAnterior.TemplatesTotal, dadosAtual.TemplatesTotal)
                };

                int linha = 5;
                foreach (var item in itens)
                {
                    resumo.Cell(linha, 1).Value = item.Item1;
                    resumo.Cell(linha, 2).Value = item.Item2;
                    resumo.Cell(linha, 3).Value = item.Item3;
                    resumo.Cell(linha, 4).Value = CalcularPercentual(item.Item2, item.Item3) / 100.0;
                    resumo.Cell(linha, 4).Style.NumberFormat.Format = "0.0%";
                    linha++;
                }

                resumo.Cell(linha + 1, 1).Value = "Resumo curto";
                resumo.Cell(linha + 2, 1).Value = GerarResumoCurto();

                var at = wb.Worksheets.Add("Atendentes");
                string[] cab = { "Atendente", dadosAnterior.Periodo, dadosAtual.Periodo, "Variação", "TMR anterior", "TMR atual", "TME anterior", "TME atual", "Nota anterior", "Nota atual" };
                for (int c = 0; c < cab.Length; c++) at.Cell(1, c + 1).Value = cab[c];

                linha = 2;
                foreach (var chave in dadosAnterior.AtendentesFinalizados.Keys.Union(dadosAtual.AtendentesFinalizados.Keys).OrderBy(NomeCurto))
                {
                    double ant = Valor(dadosAnterior.AtendentesFinalizados, chave);
                    double atual = Valor(dadosAtual.AtendentesFinalizados, chave);
                    at.Cell(linha, 1).Value = NomeExibicao(chave);
                    at.Cell(linha, 2).Value = ant;
                    at.Cell(linha, 3).Value = atual;
                    at.Cell(linha, 4).Value = CalcularPercentual(ant, atual) / 100.0;
                    at.Cell(linha, 4).Style.NumberFormat.Format = "0.0%";
                    at.Cell(linha, 5).Value = FormatarTempo(Tempo(dadosAnterior.TMR, chave));
                    at.Cell(linha, 6).Value = FormatarTempo(Tempo(dadosAtual.TMR, chave));
                    at.Cell(linha, 7).Value = FormatarTempo(Tempo(dadosAnterior.TME, chave));
                    at.Cell(linha, 8).Value = FormatarTempo(Tempo(dadosAtual.TME, chave));
                    at.Cell(linha, 9).Value = Valor(dadosAnterior.Notas, chave);
                    at.Cell(linha, 10).Value = Valor(dadosAtual.Notas, chave);
                    linha++;
                }

                var mot = wb.Worksheets.Add("Motivos");
                mot.Cell("A1").Value = "Motivo";
                mot.Cell("B1").Value = dadosAnterior.Periodo;
                mot.Cell("C1").Value = dadosAtual.Periodo;
                mot.Cell("D1").Value = "Variação";

                linha = 2;
                foreach (var m in dadosAnterior.Motivos.Keys.Union(dadosAtual.Motivos.Keys))
                {
                    double ant = Valor(dadosAnterior.Motivos, m);
                    double atual = Valor(dadosAtual.Motivos, m);
                    mot.Cell(linha, 1).Value = m;
                    mot.Cell(linha, 2).Value = ant;
                    mot.Cell(linha, 3).Value = atual;
                    mot.Cell(linha, 4).Value = CalcularPercentual(ant, atual) / 100.0;
                    mot.Cell(linha, 4).Style.NumberFormat.Format = "0.0%";
                    linha++;
                }

                foreach (var ws in wb.Worksheets)
                {
                    var used = ws.RangeUsed();
                    if (used != null)
                    {
                        used.Style.Font.FontName = "Segoe UI";
                        ws.Columns().AdjustToContents();
                    }
                }

                wb.SaveAs(sfd.FileName);

                MessageBox.Show("Comparativo exportado com sucesso.", "Exportar",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao exportar:\n" + ex.Message, "Exportar",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ============================================================
        // COMPARAÇÕES
        // ============================================================
        private List<MudancaTempo> CompararTempos(
            Dictionary<string, TimeSpan?> ant,
            Dictionary<string, TimeSpan?> atual,
            double limiteSegundos)
        {
            var lista = new List<MudancaTempo>();

            foreach (var chave in ant.Keys.Intersect(atual.Keys))
            {
                var a = ant[chave];
                var b = atual[chave];

                if (!a.HasValue || !b.HasValue) continue;

                double dif = (b.Value - a.Value).TotalSeconds;
                if (Math.Abs(dif) < limiteSegundos) continue;

                lista.Add(new MudancaTempo
                {
                    Nome = NomeExibicao(chave),
                    Anterior = a.Value,
                    Atual = b.Value,
                    DiferencaSegundos = dif
                });
            }

            return lista;
        }

        private List<MudancaNumero> CompararNumeros(
            Dictionary<string, double> ant,
            Dictionary<string, double> atual)
        {
            var lista = new List<MudancaNumero>();

            foreach (var chave in ant.Keys.Union(atual.Keys))
            {
                double a = Valor(ant, chave);
                double b = Valor(atual, chave);
                lista.Add(new MudancaNumero
                {
                    Nome = NomeExibicao(chave),
                    Anterior = a,
                    Atual = b,
                    Percentual = CalcularPercentual(a, b)
                });
            }

            return lista;
        }

        private List<MudancaNota> CompararNotas(
            Dictionary<string, double> ant,
            Dictionary<string, double> atual)
        {
            var lista = new List<MudancaNota>();

            foreach (var chave in ant.Keys.Intersect(atual.Keys))
            {
                double a = ant[chave];
                double b = atual[chave];

                lista.Add(new MudancaNota
                {
                    Nome = NomeExibicao(chave),
                    Anterior = a,
                    Atual = b,
                    Diferenca = b - a
                });
            }

            return lista;
        }

        // ============================================================
        // HELPERS
        // ============================================================
        private string NormalizarTexto(string texto)
        {
            if (string.IsNullOrWhiteSpace(texto)) return "";

            texto = texto.Trim().ToLowerInvariant();

            string formD = texto.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder();

            foreach (char ch in formD)
            {
                var cat = CharUnicodeInfo.GetUnicodeCategory(ch);
                if (cat != UnicodeCategory.NonSpacingMark)
                    sb.Append(ch);
            }

            return sb.ToString()
                .Normalize(NormalizationForm.FormC)
                .Replace(":", "")
                .Trim();
        }

        private string NormalizarNome(string nome)
        {
            string n = NormalizarTexto(nome);

            n = n.Replace("dr. gurgel", "gurgel")
                 .Replace("w. luiz", "wl")
                 .Replace("w luiz", "wl")
                 .Replace("washington luiz", "wl");

            return n.Trim();
        }

        private string NomeExibicao(string chave)
        {
            if (string.IsNullOrWhiteSpace(chave))
                return "";

            return CultureInfo.GetCultureInfo("pt-BR")
                .TextInfo
                .ToTitleCase(chave);
        }

        private string NomeCurto(string nome)
        {
            if (string.IsNullOrWhiteSpace(nome)) return "";

            string n = nome;
            n = n.Replace(" - dr. gurgel", "", StringComparison.OrdinalIgnoreCase)
                 .Replace(" - gurgel", "", StringComparison.OrdinalIgnoreCase)
                 .Replace("-gurgel", "", StringComparison.OrdinalIgnoreCase)
                 .Replace(" - wl", "", StringComparison.OrdinalIgnoreCase)
                 .Replace(" - w. luiz", "", StringComparison.OrdinalIgnoreCase)
                 .Replace(" w. luiz", "", StringComparison.OrdinalIgnoreCase);

            n = n.Trim();

            // nomes normalizados estão em minúsculo: coloca iniciais em maiúscula
            return CultureInfo.GetCultureInfo("pt-BR").TextInfo.ToTitleCase(n);
        }

        private string JuntarNomes(List<string> nomes)
        {
            nomes = nomes.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct().ToList();

            if (nomes.Count == 0) return "";
            if (nomes.Count == 1) return nomes[0];
            if (nomes.Count == 2) return $"{nomes[0]} e {nomes[1]}";

            return string.Join(", ", nomes.Take(nomes.Count - 1)) + " e " + nomes.Last();
        }

        private double CalcularPercentual(double anterior, double atual)
        {
            if (Math.Abs(anterior) < 0.000001)
                return atual > 0 ? 100 : 0;

            return (atual - anterior) / anterior * 100.0;
        }

        private string FormatarPercentualComSeta(double p)
        {
            if (Math.Abs(p) < 0.05) return "● 0,0%";
            return p > 0 ? $"▲ {Math.Abs(p):N1}%" : $"▼ {Math.Abs(p):N1}%";
        }

        private string FormatarTempo(TimeSpan? t)
        {
            if (!t.HasValue) return "—";

            if (t.Value.TotalDays >= 1)
                return $"{(int)t.Value.TotalDays}d {t.Value.Hours:00}h {t.Value.Minutes:00}m";

            if (t.Value.TotalHours >= 1)
                return $"{(int)t.Value.TotalHours}h {t.Value.Minutes:00}m";

            return $"{t.Value.Minutes}m {t.Value.Seconds:00}s";
        }

        private double Valor(Dictionary<string, double> dict, string chave)
        {
            return dict != null && dict.TryGetValue(chave, out var v) ? v : 0;
        }

        private TimeSpan? Tempo(Dictionary<string, TimeSpan?> dict, string chave)
        {
            return dict != null && dict.TryGetValue(chave, out var v) ? v : null;
        }

        private double ValorPorDescricao(Dictionary<string, double> dict, string descricao)
        {
            string alvo = NormalizarTexto(descricao);
            foreach (var kv in dict)
                if (NormalizarTexto(kv.Key) == alvo)
                    return kv.Value;
            return 0;
        }

        // ============================================================
        // MODELOS
        // ============================================================
        private class PaginaRelatorioPdf
        {
            public string Tipo { get; set; } = "";
            public int Inicio { get; set; }
            public int Quantidade { get; set; }
            public bool PrimeiraPagina { get; set; }
        }

        private class ColunaTabelaPdf
        {
            public string Titulo { get; set; } = "";
            public float Peso { get; set; }
        }

        private class DadosMes
        {
            public string Arquivo { get; set; } = "";
            public string Periodo { get; set; } = "";

            public double Conversas { get; set; }
            public double Finalizados { get; set; }
            public double NovosContatos { get; set; }
            public double Reagendamentos { get; set; }
            public double TemplatesTotal { get; set; }
            public double Inatividade { get; set; }
            public double OrcamentoFormula { get; set; }
            public double OrcamentoFormulaPossivelVenda { get; set; }

            public double MarketingTotal { get; set; }

            public double NotaGeral { get; set; }
            public double NotaWL { get; set; }
            public double NotaGurgel { get; set; }

            public Dictionary<string, double> AtendentesFinalizados { get; set; } =
                new(StringComparer.OrdinalIgnoreCase);

            public Dictionary<string, double> Motivos { get; set; } =
                new(StringComparer.OrdinalIgnoreCase);

            public Dictionary<string, double> MarketingCriativos { get; set; } =
                new(StringComparer.OrdinalIgnoreCase);

            public Dictionary<string, TimeSpan?> TMR { get; set; } =
                new(StringComparer.OrdinalIgnoreCase);

            public Dictionary<string, TimeSpan?> TME { get; set; } =
                new(StringComparer.OrdinalIgnoreCase);

            public Dictionary<string, TimeSpan?> TMA { get; set; } =
                new(StringComparer.OrdinalIgnoreCase);

            public Dictionary<string, double> Notas { get; set; } =
                new(StringComparer.OrdinalIgnoreCase);
        }

        private class MudancaTempo
        {
            public string Nome { get; set; } = "";
            public TimeSpan Anterior { get; set; }
            public TimeSpan Atual { get; set; }
            public double DiferencaSegundos { get; set; }
        }

        private class MudancaNumero
        {
            public string Nome { get; set; } = "";
            public double Anterior { get; set; }
            public double Atual { get; set; }
            public double Percentual { get; set; }
        }

        private class MudancaNota
        {
            public string Nome { get; set; } = "";
            public double Anterior { get; set; }
            public double Atual { get; set; }
            public double Diferenca { get; set; }
        }
    }
}
