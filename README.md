# Stellantis — Auditor de Especificações IPC (VFs × Indications & Parameters Table)

Sistema automatizado de auditoria bidirecional de especificações de engenharia do **IPC (*Instrument Panel Cluster*)** para prevenção de **RARs (*Requests for Action / Clarification*)**.

---

## 🎯 As 2 Regras Centrais de Auditoria

O sistema lê diretamente os documentos de **VFs (*Vehicle Functions*)** em formato `.docx` ou `.pdf` (incluindo listas consolidadas com **mais de 1.000 páginas**) e realiza o cruzamento bidirecional contra as planilhas Extra-DOORS do IPC:

1. **VF chama Indications (e vice-versa):**
   > *Todas as indicações das VFs que impactam o IPC (seção `Indication Management`) estão listadas dentro do documento de `Indications.xlsx`? **E vice-versa** (existem indicações órfãs na planilha sem VF correspondente)?*
   - Normaliza automaticamente diferenças de grafia entre o texto da VF (`HAND BRAKE.V1`) e a planilha (`HAND BRAKE V1`).

2. **VF chama Parâmetros `[SUPPLIER_EOL_2]` (e vice-versa):**
   > *Todos os parâmetros `[SUPPLIER_EOL_2]` / `[SUPPLIER EOL]` de `IPC` de cada VF (seção `Configuration Parameters Table`) estão listados no `Parameters_Table.xlsx` com `First Trial Value` válido? **E vice-versa** (existem parâmetros órfãos na planilha sem VF correspondente)?*
   - Segrega automaticamente parâmetros `PROXI` (alocados em documento específico da Stellantis) e parâmetros de outros módulos (`BCM`, `ORC`, `BSM`, etc.) para eliminar falsos positivos.

---

## 📚 Documentação Completa do Projeto

- 📄 **[Briefing Executivo (Contexto, Problema dos RARs e ROI)](docs/BRIEFING_EXECUTIVO.md)**
- ⚙️ **[Especificação Técnica e de Arquitetura (SPECS)](docs/ESPECIFICACAO_TECNICA_SPECS.md)**
- 📖 **[Manual Passo a Passo do Usuário](docs/MANUAL_DO_USUARIO.md)**

---

## 🚀 Como Instalar e Iniciar (Em 2 Cliques no Windows)

### 1. Instalar Dependências (Apenas na 1ª vez)
Dê dois cliques em:
```bat
instalar_requisitos.bat
```
*(Cria automaticamente o ambiente virtual `.venv`, instala as bibliotecas de `requirements.txt`, gera os pacotes de exemplo e executa os testes automatizados).*

### 2. Iniciar o Sistema Web
Dê dois cliques em:
```bat
iniciar_sistema.bat
```
O navegador abrirá automaticamente em **http://127.0.0.1:8000**, onde você pode:
- **Enviar manualmente os arquivos de qualquer projeto** (`363`, `376`, `J3U`, `F1H`, etc.):
  1. Documento(s) de VF (`.docx` ou `.pdf`, único de 1.000+ páginas ou múltiplos arquivos)
  2. Planilha `Indications.xlsx` (*Section C - HMI Requirements*)
  3. Planilha `Parameters_Table.xlsx` (*Section B - Software Requirements*)
- **Informar o caminho de uma pasta local ou de rede** no padrão Stellantis.
- **Testar imediatamente com os pacotes de demonstração inclusos:**
  - ⚡ **Exemplo Rápido (6 VFs)**
  - 📚 **Exemplo Completo em Escala Real (1.215 Páginas / 45 VFs)**
- **Exportar relatórios completos** em **Excel (`.xlsx`)** e **PDF (`.pdf`)**.

---

## 🧪 Como Executar os Testes Automatizados

```powershell
.\.venv\Scripts\python.exe -m pytest -v
```
A suíte de testes (`tests/test_auditor.py`) valida:
- Normalização de nomenclatura (`.V1` $\leftrightarrow$ ` V1`) e extração de IDs de VF;
- Filtro estrito de escopo (`Component = IPC` & `Implementation = SUPPLIER_EOL_2` vs. `PROXI` e outras ECUs);
- Leitura híbrida de tabelas e textos corridos no padrão DOORS;
- Auditoria de pacote em escala real com **1.215 páginas e 45 VFs** em menos de 3 segundos.
