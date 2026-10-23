# Mail Merge & Document Automation Desktop Application (Docx to PDF)

> Robust Python desktop application engineered for automated contractual mail-merge workflows, multi-run XML placeholder injection, and direct PDF compilation via Windows COM automation.

---

## 🌐 Language Navigation
- [English Overview](#-english-overview)
- [Visão Geral em Português](#-visão-geral-em-português)

---

## 🇺🇸 English Overview

### 🎯 Problem Statement & Solution
In institutional and corporate contract supervision, generating monthly fiscal attestation documents and compliance reports requires repetitive manual data entry into Word templates. This manual process is prone to formatting discrepancies, human error, and inconsistent directory structuring.

This application provides a modern graphical desktop interface (GUI) built with **CustomTkinter** that automates the generation of required fiscal documents directly into clean, distribution-ready **PDFs** without intermediate residual files.

### ⚙️ Key Technical Highlights
- **Word XML Run Fragmentation Handling (`python-docx`):** Solves the common issue where Word splits template placeholders (e.g., `{{FATURA}}`) across multiple internal formatting XML runs. The algorithm detects keys across the whole paragraph and preserves original styling without inheriting adjacent bold labels.
- **Isolated Directory Partitioning:** Automatically categorizes outputs into `<YEAR>/<MONTH>/` (e.g., `2026/AUGUST/`). Modifying or regenerating documents for a specific month strictly targets that directory, leaving all other periods untouched.
- **Interactive Collision & Overwrite Guards:** Protects existing institutional records with proactive prompt confirmation before overwriting previous generations.
- **Headless Word COM Pipeline (`pywin32`):** Direct Windows COM automation (`Word.Application`) for faithful DOCX to PDF conversion, completely eliminating UI freezing and console output dependency errors in windowed (`--windowed`) standalone executables.
- **Native OS Compatibility:** Sanitizes dates for Windows filesystem standards (e.g., converting `/` to `_` in file naming while retaining readable formatting inside the generated document).

### 🛠️ Tech Stack
- **Python 3.13+**
- **CustomTkinter:** Modern GUI with reactive styling and system theme auto-adaptation.
- **python-docx:** XML-level template parsing and token substitution.
- **pywin32 (Windows COM):** High-fidelity document conversion engine.
- **PyInstaller:** Native standalone Windows binary compilation.

### 🚀 Getting Started & Execution
1. **Clone the repository:**
   ```bash
   git clone https://github.com/your-username/automacao_atesto.git
   cd automacao_atesto
   ```
2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```
3. **Templates ready to use:**
   - The repository includes pre-configured sample templates: `modelo_atesto.example.docx` and `modelo_relatorio.example.docx`.
   - The application automatically uses them if present, or you can supply your own `modelo_atesto.docx` and `modelo_relatorio.docx`.
4. **Run application:**
   ```bash
   python app.py
   ```
5. **Compile to Standalone `.exe`:**
   ```powershell
   python -m PyInstaller --noconfirm --onefile --windowed --name "GeradorAtesto" app.py
   ```

---

## 🇧🇷 Visão Geral em Português

### 🎯 Problema e Solução
Na fiscalização e acompanhamento de contratos institucionais e corporativos, a emissão periódica de atestos de prestação de serviços e relatórios de execução exige o preenchimento manual repetitivo de modelos do Word. Essa rotina gera retrabalho, perda de formatação e riscos de erros de digitação.

Esta aplicação desktop com interface visual moderna (**CustomTkinter**) automatiza o preenchimento de variáveis contratuais diretamente em **PDFs finais**, estruturados e higienizados, sem deixar resíduos de arquivos intermediários.

### ⚙️ Diferenciais Técnicos
- **Tratamento de Fragmentação de Runs no Word (`python-docx`):** Contorna o comportamento do Microsoft Word que divide chaves como `{{FATURA}}` em múltiplos nós XML adjacentes. A substituição é feita com precisão cirúrgica, mantendo o rótulo original em negrito e o texto preenchido sem formatação excessiva.
- **Isolamento Estrito por Ano e Mês:** Organiza os relatórios em pastas `<ANO>/<MES>/`. Alterações ou substituições de um mês específico (ex: Agosto) afetam exclusivamente aquele mês, preservando intocados todos os demais meses do ano.
- **Proteção Ativa contra Sobrescrita:** Diálogo de confirmação preventiva antes de substituir documentos de meses já expedidos anteriormente.
- **Pipeline de Conversão via Windows COM (`pywin32`):** Conversão nativa pelo motor do Microsoft Word sem dependência de emuladores de console ou bibliotecas de progresso de terminal, garantindo 100% de estabilidade no executável `--windowed`.
- **Higienização para Sistema de Arquivos Windows:** Converte automaticamente barras `/` da data em `_` nos nomes de arquivo, preservando a data com `/` no corpo do documento.

### 🛠️ Tecnologias Utilizadas
- **Python 3.13+**
- **CustomTkinter**
- **python-docx**
- **pywin32 (Windows COM)**
- **PyInstaller**

### 🚀 Como Executar e Compilar
1. **Instalar dependências:**
   ```bash
   pip install -r requirements.txt
   ```
2. **Modelos de exemplo inclusos e prontos para teste:**
   - O projeto já acompanha os modelos de teste descaracterizados: `modelo_atesto.example.docx` e `modelo_relatorio.example.docx`.
   - A aplicação os reconhece automaticamente de imediato, ou você pode fornecer seus próprios `modelo_atesto.docx` e `modelo_relatorio.docx`.
3. **Executar em desenvolvimento:**
   ```bash
   python app.py
   ```
4. **Compilar o executável autônomo (.exe):**
   ```powershell
   python -m PyInstaller --noconfirm --onefile --windowed --name "GeradorAtesto" app.py
   ```
