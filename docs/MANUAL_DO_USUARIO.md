# Manual do Usuário — Auditor de Especificações IPC (Stellantis)

Este manual explica passo a passo como instalar, abrir e utilizar o sistema para validar qualquer pacote de especificações do **IPC (Instrument Panel Cluster)**.

---

## O Que o Sistema Verifica Automaticamente?
Ao enviar seus arquivos, o sistema responde imediatamente às duas perguntas críticas de engenharia:
1. **VF chama Indications:** *Todas as indicações das VFs que impactam IPC estão listadas dentro do documento de `Indications`? **E vice-versa?***
2. **VF chama Parâmetros:** *Todos os parâmetros `[SUPPLIER_EOL_2]` de `IPC` de cada VF estão listados no `Parameters_Table`? **E vice-versa?***

---

## 1. Instalação (Apenas na Primeira Vez)
1. Abra a pasta do projeto no Windows Explorer.
2. Dê dois cliques no arquivo **`instalar_requisitos.bat`**.
3. O instalador criará automaticamente o ambiente isolado (`.venv`), instalará todas as bibliotecas necessárias e validará o sistema.

---

## 2. Como Iniciar o Sistema no Dia a Dia
1. Dê dois cliques no arquivo **`iniciar_sistema.bat`**.
2. Uma janela do navegador abrirá automaticamente no endereço **http://127.0.0.1:8000**.
   *(Mantenha a janela preta do terminal aberta enquanto estiver usando o painel).*

---

## 3. Como Analisar os Arquivos do Seu Projeto (Passo a Passo)

No topo da tela inicial (**PASSO 1**), você pode escolher entre duas formas de enviar seus documentos:

### Opção A — Selecionar os 3 Arquivos Manualmente no Computador
1. **Nome / Código do Projeto:** Digite o nome do programa que está auditando (ex.: `Projeto 363 - MY27`, `Projeto 376`, `Projeto J3U`, etc.).
2. **Arquivo 1 — Documento(s) de VF (`.docx` ou `.pdf`):**
   - Clique na primeira caixa e selecione o documento consolidado de VFs (mesmo que tenha **mais de 1.000 páginas**) ou selecione múltiplos arquivos `.docx`/`.pdf` individuais segurando a tecla `Ctrl`.
3. **Arquivo 2 — Planilha de Indicações (`Indications.xlsx`):**
   - Clique na segunda caixa e selecione a planilha da *Section C - HMI Requirements*.
4. **Arquivo 3 — Planilha de Parâmetros (`Parameters_Table.xlsx`):**
   - Clique na terceira caixa e selecione a planilha da *Section B - Software Requirements*.
5. Clique no botão azul **"🔍 Ler Arquivos e Verificar Consistência Agora"**.

### Opção B — Informar o Caminho de uma Pasta Local ou de Rede
- Caso seus arquivos já estejam organizados na árvore de pastas padrão da Stellantis, basta colar o caminho da pasta no campo **"📁 Prefere informar o caminho de uma pasta local/rede?"** e clicar em **"Ler Pasta"**.

---

## 4. Como Ler os Resultados no Dashboard

Logo abaixo da área de envio, o painel exibe **4 Cartões Resumo** e **5 Abas**:

1. **Aba `🚨 O Que Precisa Ser Corrigido (Resumo)`:**
   - Mostra **apenas os erros encontrados** (indicações ou parâmetros faltando na planilha, itens órfãos sobrando na planilha, códigos de VF trocados ou valores `First Trial` em branco).
   - Você pode marcar a caixinha **"Feito (`✓`)"** à medida que for corrigindo cada linha na sua planilha.
2. **Aba `📋 1. VF chama Indications (E vice-versa)`:**
   - Mostra todas as indicações das VFs que impactam o IPC comparadas com o documento `Indications.xlsx`.
3. **Aba `⚙️ 2. VF chama Parâmetros [SUPPLIER_EOL_2] (E vice-versa)`:**
   - Mostra todos os parâmetros `IPC` + `SUPPLIER_EOL_2` comparados com o documento `Parameters_Table.xlsx`.
4. **Aba `📘 VFs Lidas no Documento`:**
   - Mostra o resumo individual de cada VF encontrada dentro do Word/PDF enviado.
5. **Aba `🛡️ Parâmetros Ignorados (PROXI / Outros Módulos)`:**
   - Lista transparante de todos os parâmetros `PROXI` ou de outras ECUs (`BCM`, `ORC`, etc.) que o sistema separou automaticamente para não gerar falsos erros.

---

## 5. Como Exportar o Relatório em Excel (`.xlsx`) ou PDF (`.pdf`)
No canto superior direito da barra azul do topo, clique em:
- **`📊 Baixar Relatório Excel (.xlsx)`:** Baixa uma planilha completa com 4 abas formatadas e coloridas por gravidade.
- **`📑 Baixar Relatório PDF (.pdf)`:** Baixa um relatório executivo pronto para enviar por e-mail ou anexar em reuniões de fechamento de pacote.
