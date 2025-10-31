# Extrator de Cursos em PDFs – TCDF

Aplicação em **Python** com interface moderna (**ttkbootstrap**) para extrair informações de “Papeletas de Benefícios – AQ”/relacionados (TCDF).  
Faz leitura estruturada dos PDFs, detecta cursos por faixa de Y, interpreta **checkboxes** (Presencial/Misto/À distância), e exporta tudo para **Excel (.xlsx)**.

> **Destaque**: extração robusta de **Lotação** mesmo quando o valor quebra de linha e vem acompanhado de “Ramal: …” – o código limpa e mantém apenas a lotação (ex.: `SECOF`).

---

## ✨ Funcionalidades
- Interface moderna (tema claro/escuro) com barra de progresso, hover nas linhas e “auto-fit” das colunas.
- Extração de: **arquivo, requerente, cargo, lotação, curso_titulo, curso_horas, modalidade**.
- Interpretação dos **checkboxes** por coordenadas X (Presencial/Misto/À distância).
- Filtros por **faixa Y** e **páginas** (ex.: “1” ou “1,2”).
- Exporta automaticamente para **`dados_extraidos.xlsx`** (ou nome customizado).
- Opção de exportar **PNGs de depuração** com overlays (faixa Y/colunas).

---

## 🧩 Requisitos
- Python **3.10+** (funcionando também no 3.13)
- Windows (testado); tk/ttk já vem com o Python oficial
- Dependências do `requirements.txt`

---

## 🚀 Instalação rápida (Windows)
```bash
git clone https://github.com/Disrrptt/extrator-cursos-tcdf.git
cd <extrator-cursos-tcdf>
setup.bat
