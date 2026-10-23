"""
Aplicação Desktop para Automação de Mala Direta de Documentos do Word.
Substitui tags no formato {{NOME_VARIAVEL}} nos modelos modelo_atesto.docx e modelo_relatorio.docx.
Organiza os arquivos gerados em diretórios aninhados por Ano e Mês (<ANO>/<MES>/).
"""

import os
import sys
import datetime
import tempfile
from tkinter import messagebox
import customtkinter as ctk
import docx
import win32com.client
import pythoncom


def obter_diretorio_base() -> str:
    """
    Retorna o diretório base onde os arquivos de modelo e o executável estão localizados.
    Lida com a execução normal do Python e com o binário congelado pelo PyInstaller.

    :return: Caminho absoluto do diretório raiz da aplicação.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def substituir_texto_em_paragrafo(paragrafo, mapa_substituicoes: dict):
    """
    Substitui chaves em um parágrafo preservando a formatação original dos rótulos.
    Garante que os textos preenchidos não herdem indevidamente o negrito de rótulos adjacentes.

    :param paragrafo: Objeto Paragraph do python-docx.
    :param mapa_substituicoes: Dicionário contendo { '{{CHAVE}}': 'VALOR' }.
    """
    texto_completo = "".join(run.text for run in paragrafo.runs)
    if not texto_completo:
        return

    # Verifica se há alguma chave a ser substituída neste parágrafo
    chaves_presentes = [ch for ch in mapa_substituicoes.keys() if ch in texto_completo]
    if not chaves_presentes:
        return

    # Caso 1: A chave está contida inteiramente dentro de um único run
    for run in paragrafo.runs:
        for chave in chaves_presentes:
            if chave in run.text:
                run.text = run.text.replace(chave, str(mapa_substituicoes[chave]))
                # Desativa o negrito explicitamente no valor preenchido
                run.bold = False

    # Atualiza o texto consolidado para ver se restou alguma chave dividida entre múltiplos runs
    texto_atualizado = "".join(run.text for run in paragrafo.runs)
    chaves_restantes = [ch for ch in mapa_substituicoes.keys() if ch in texto_atualizado]

    # Caso 2: Fragmentação entre runs (a tag começou em um run e terminou em outro)
    if chaves_restantes:
        for chave in chaves_restantes:
            texto_atualizado = texto_atualizado.replace(chave, str(mapa_substituicoes[chave]))

        if paragrafo.runs:
            paragrafo.runs[0].text = texto_atualizado
            paragrafo.runs[0].bold = False
            for run in paragrafo.runs[1:]:
                run.text = ""


def converter_docx_para_pdf_direto(caminho_docx: str, caminho_pdf: str):
    """
    Converte um arquivo .docx para .pdf usando diretamente a automação COM do Microsoft Word.
    Não utiliza bibliotecas de barra de progresso (como tqdm), o que elimina 100% dos erros de
    'NoneType object has no attribute write' comuns em executáveis compildos com --windowed.

    :param caminho_docx: Caminho absoluto do arquivo Word.
    :param caminho_pdf: Caminho absoluto do arquivo PDF de destino.
    """
    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # Desativa caixas de diálogo do Word
        
        # 17 corresponde à constante wdFormatPDF do Word
        wdFormatPDF = 17
        doc = word.Documents.Open(os.path.abspath(caminho_docx))
        doc.SaveAs(os.path.abspath(caminho_pdf), FileFormat=wdFormatPDF)
    finally:
        if doc is not None:
            try:
                doc.Close(0)  # 0 = wdDoNotSaveChanges
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def processar_documento_para_pdf(caminho_origem: str, caminho_destino_pdf: str, mapa_substituicoes: dict):
    """
    Carrega o modelo do Word, efetua as substituições das chaves e salva diretamente
    no formato PDF no caminho de destino. O arquivo intermediário .docx é gerado
    na pasta temporária do sistema e imediatamente excluído após a conversão.

    :param caminho_origem: Caminho do arquivo modelo .docx.
    :param caminho_destino_pdf: Caminho onde o arquivo .pdf final será salvo.
    :param mapa_substituicoes: Dicionário com as substituições a serem feitas.
    """
    doc = docx.Document(caminho_origem)

    # 1. Substitui em parágrafos normais
    for paragrafo in doc.paragraphs:
        substituir_texto_em_paragrafo(paragrafo, mapa_substituicoes)

    # 2. Substitui em tabelas (linhas, colunas e parágrafos dentro de células)
    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                for paragrafo in celula.paragraphs:
                    substituir_texto_em_paragrafo(paragrafo, mapa_substituicoes)

    # Cria arquivo docx intermediário na pasta temporária do sistema
    fd, caminho_temp_docx = tempfile.mkstemp(suffix=".docx")
    os.close(fd)

    try:
        doc.save(caminho_temp_docx)
        converter_docx_para_pdf_direto(caminho_temp_docx, caminho_destino_pdf)
    finally:
        if os.path.exists(caminho_temp_docx):
            try:
                os.remove(caminho_temp_docx)
            except OSError:
                pass


class AppAutomacaoAtesto(ctk.CTk):
    """
    Janela Principal da Interface Gráfica construída com CustomTkinter.
    """

    MESES = [
        "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
        "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"
    ]

    def __init__(self):
        super().__init__()

        # Configurações básicas da janela
        self.title("Automação de Mala Direta - Atesto & Relatório")
        self.geometry("560x640")
        self.resizable(False, False)

        # Tema visual moderno
        ctk.set_appearance_mode("System")  # Segue o modo do Windows (Dark/Light)
        ctk.set_default_color_theme("blue")

        self._construir_interface()

    def _construir_interface(self):
        """Monta os componentes visuais do formulário."""
        # Cabeçalho
        frame_cabecalho = ctk.CTkFrame(self, fg_color="transparent")
        frame_cabecalho.pack(pady=(20, 10), padx=30, fill="x")

        lbl_titulo = ctk.CTkLabel(
            frame_cabecalho,
            text="Geração de Documentos de Atesto",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        lbl_titulo.pack()

        lbl_subtitulo = ctk.CTkLabel(
            frame_cabecalho,
            text="Preencha os dados abaixo para preencher os modelos do Word.",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        lbl_subtitulo.pack(pady=(2, 0))

        # Container do Formulário
        frame_form = ctk.CTkFrame(self, corner_radius=10)
        frame_form.pack(pady=10, padx=30, fill="both", expand=True)

        # 1. Mês e Ano de Referência em colunas lado a lado
        frame_periodo_ref = ctk.CTkFrame(frame_form, fg_color="transparent")
        frame_periodo_ref.pack(fill="x", padx=20, pady=(15, 8))

        # Mês
        frame_mes = ctk.CTkFrame(frame_periodo_ref, fg_color="transparent")
        frame_mes.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkLabel(frame_mes, text="Mês de Referência:", font=ctk.CTkFont(weight="bold")).pack(anchor="w")

        mes_atual_idx = datetime.datetime.now().month - 1
        self.cb_mes = ctk.CTkComboBox(frame_mes, values=self.MESES)
        self.cb_mes.set(self.MESES[mes_atual_idx])
        self.cb_mes.pack(fill="x", pady=(4, 0))

        # Ano
        frame_ano = ctk.CTkFrame(frame_periodo_ref, fg_color="transparent")
        frame_ano.pack(side="right", fill="x", expand=True, padx=(10, 0))
        ctk.CTkLabel(frame_ano, text="Ano de Referência:", font=ctk.CTkFont(weight="bold")).pack(anchor="w")

        ano_atual = datetime.datetime.now().year
        anos_disponiveis = [str(a) for a in range(ano_atual - 2, ano_atual + 5)]
        self.cb_ano = ctk.CTkComboBox(frame_ano, values=anos_disponiveis)
        self.cb_ano.set(str(ano_atual))
        self.cb_ano.pack(fill="x", pady=(4, 0))

        # 2. Campo Fatura
        self.entry_fatura = self._criar_campo_texto(frame_form, "Número da Fatura:", "Ex: 12345")

        # 3. Campo Período
        self.entry_periodo = self._criar_campo_texto(frame_form, "Período:", "Ex: Conforme processo no SIGADOC")

        # 4. Campo Data do Atesto
        data_hoje_str = datetime.datetime.now().strftime("%d/%m/%Y")
        self.entry_data = self._criar_campo_texto(frame_form, "Data do Atesto:", "Ex: 23/10/2026", valor_padrao=data_hoje_str)

        # 5. Campo Valor
        self.entry_valor = self._criar_campo_texto(frame_form, "Valor:", "Ex: R$ 1.500,00")

        # Botão Gerar Documentos em Destaque
        self.btn_gerar = ctk.CTkButton(
            self,
            text="🚀 Gerar Documentos",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=45,
            command=self._acao_gerar_documentos
        )
        self.btn_gerar.pack(pady=20, padx=30, fill="x")

    def _criar_campo_texto(self, parent, label_text: str, placeholder: str, valor_padrao: str = "") -> ctk.CTkEntry:
        """Cria e posiciona um rótulo e um campo Entry padronizados."""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", padx=20, pady=6)

        lbl = ctk.CTkLabel(frame, text=label_text, font=ctk.CTkFont(weight="bold"))
        lbl.pack(anchor="w")

        entry = ctk.CTkEntry(frame, placeholder_text=placeholder, height=36)
        if valor_padrao:
            entry.insert(0, valor_padrao)
        entry.pack(fill="x", pady=(3, 0))
        return entry

    def _acao_gerar_documentos(self):
        """Valida os dados, confirma eventuais sobrescritas do mês e gera os arquivos."""
        mes = self.cb_mes.get().strip().upper()
        ano = self.cb_ano.get().strip()
        fatura = self.entry_fatura.get().strip()
        periodo = self.entry_periodo.get().strip()
        data_original = self.entry_data.get().strip()
        valor = self.entry_valor.get().strip()

        # Validação de preenchimento obrigatório
        if not all([mes, ano, fatura, periodo, data_original, valor]):
            messagebox.showwarning("Campos Obrigatórios", "Por favor, preencha todos os campos do formulário antes de continuar.")
            return

        diretorio_base = obter_diretorio_base()

        # Validação da existência dos modelos originais
        caminho_modelo_atesto = os.path.join(diretorio_base, "modelo_atesto.docx")
        caminho_modelo_relatorio = os.path.join(diretorio_base, "modelo_relatorio.docx")

        if not os.path.exists(caminho_modelo_atesto):
            messagebox.showerror("Modelo Ausente", f"Não foi encontrado o arquivo:\n{caminho_modelo_atesto}\nCertifique-se de que ele está na mesma pasta do programa.")
            return

        if not os.path.exists(caminho_modelo_relatorio):
            messagebox.showerror("Modelo Ausente", f"Não foi encontrado o arquivo:\n{caminho_modelo_relatorio}\nCertifique-se de que ele está na mesma pasta do programa.")
            return

        # 1. Tratamento da data para atender ao padrão de nomes de arquivos do Windows
        # Substitui barras '/' por '_'
        data_tratada = data_original.replace("/", "_").replace("\\", "_")

        # 2. Definição estrita das pastas por Ano e Mês: <diretorio_base>/<ANO>/<MES>/
        pasta_ano = os.path.join(diretorio_base, ano)
        pasta_mes = os.path.join(pasta_ano, mes)

        # 3. Nomes finais dos documentos em PDF
        nome_atesto = f"Atesto_{data_tratada}.pdf"
        nome_relatorio = f"Relatorio_{data_tratada}.pdf"

        caminho_final_atesto = os.path.join(pasta_mes, nome_atesto)
        caminho_final_relatorio = os.path.join(pasta_mes, nome_relatorio)

        # 4. Verificação de existência prévia: APENAS na pasta do mês específico selecionado
        pasta_ja_existe = os.path.exists(pasta_mes)
        arquivos_ja_existem = os.path.exists(caminho_final_atesto) or os.path.exists(caminho_final_relatorio)

        if pasta_ja_existe and (arquivos_ja_existem or len(os.listdir(pasta_mes)) > 0):
            confirmar = messagebox.askyesno(
                "Confirmação de Sobrescrita",
                f"A pasta do mês '{mes}' no ano '{ano}' já existe com documentos anteriores.\n\n"
                f"Deseja substituir/sobrescrever os documentos PDF apenas deste mês ({mes}/{ano})?\n\n"
                f"(Os demais meses de {ano} permanecerão totalmente inalterados)."
            )
            if not confirmar:
                return  # Usuário optou por não sobrescrever, cancela operação de forma segura

        # 5. Criação do diretório específico do mês (se ainda não existir)
        os.makedirs(pasta_mes, exist_ok=True)

        # 6. Dicionário de mapeamento de tags
        mapa_substituicoes = {
            "{{FATURA}}": fatura,
            "{{PERIODO}}": periodo,
            "{{DATA}}": data_original,  # Mantém formato original legível no documento
            "{{VALOR}}": valor
        }

        # Desabilita o botão temporariamente durante a conversão
        self.btn_gerar.configure(state="disabled", text="⏳ Gerando PDFs...")
        self.update()

        try:
            # Processa e converte diretamente para PDF
            processar_documento_para_pdf(caminho_modelo_atesto, caminho_final_atesto, mapa_substituicoes)
            processar_documento_para_pdf(caminho_modelo_relatorio, caminho_final_relatorio, mapa_substituicoes)

            # Feedback de sucesso ao usuário
            resposta = messagebox.askyesno(
                "Sucesso!",
                f"Documentos PDF gerados com sucesso na pasta:\n{ano}/{mes}\n\n"
                f"• {nome_atesto}\n"
                f"• {nome_relatorio}\n\n"
                f"Deseja abrir a pasta agora no Explorador de Arquivos?"
            )
            if resposta:
                os.startfile(pasta_mes)

        except Exception as erro:
            messagebox.showerror("Erro ao Gerar Documentos", f"Ocorreu um erro durante a criação dos PDFs:\n{str(erro)}")
        finally:
            self.btn_gerar.configure(state="normal", text="🚀 Gerar Documentos")


if __name__ == "__main__":
    app = AppAutomacaoAtesto()
    app.mainloop()
