# Briefing Executivo — Auditor de Consistência IPC (VFs × Especificações Extra-DOORS)

## 1. Objetivo Central do Projeto (As 2 Regras de Ouro)
O objetivo do sistema é automatizar as duas verificações bidirecionais de consistência do **IPC (*Instrument Panel Cluster*)** antes da liberação de pacotes ao fornecedor, evitando a abertura de **RARs**:

> 1. **VF chama Indications:** *Todas as indicações das VFs que impactam IPC estão listadas dentro do documento de `Indications`? **E vice-versa?***
> 2. **VF chama Parâmetros:** *Todos os parâmetros `[SUPPLIER_EOL_2]` de `IPC` de cada VF estão listados no `Parameters_Table`? **E vice-versa?***

---

## 2. Visão Geral e Contexto de Negócio
No desenvolvimento de sistemas Eletroeletrônicos (EE) automotivos da **Stellantis**, as especificações funcionais do veículo são divididas em duas frentes complementares elaboradas por equipes distintas:
1. **Especificações Sistêmicas — VFs (*Vehicle Functions*):** Documentos extensos (frequentemente superiores a **1.000 páginas** no pacote consolidado `.docx` ou `.pdf` exportado do DOORS) que descrevem o comportamento de cada função do veículo (`Airbag`, `Hand Brake`, `TPMS`, `Seat Belt Reminder`, `Wiper & Washer`, etc.).
2. **Especificações de Componente (Extra-DOORS) — IPC (*Instrument Panel Cluster*):** Planilhas técnicas Excel (`.xlsx`) que consolidam todas as indicações visuais/acústicas (**Section C — HMI Requirements / `Indications.xlsx`**) e os parâmetros de calibração de fim de linha do fornecedor (**Section B — Software Requirements / `Parameters_Table.xlsx`**).

---

## 2. O Problema (Geração de RARs)
Como os documentos de Sistema (**VFs**) e os documentos de Componente (**Planilhas do IPC**) são atualizados em ciclos paralelos por diferentes engenheiros e plataformas (`363`, `376`, `J3U`, `F1H`, etc.), ocorrem desalinhamentos frequentes no fechamento de pacotes de software:
- **Indicação chamada na VF, mas esquecida na planilha `Indications.xlsx`.**
- **Indicação "órfã" na planilha `Indications.xlsx`** (herdada por *carryover* de projetos anteriores, mas sem nenhuma VF que a chame no projeto atual).
- **Divergência no código da VF** registrado na coluna `VFs` da planilha em relação ao documento da VF.
- **Parâmetro `IPC` + `SUPPLIER EOL` exigido na VF, mas ausente na `Parameters_Table.xlsx`.**
- **Parâmetro "órfão" na `Parameters_Table.xlsx`** que não é utilizado por nenhuma VF do pacote.
- **Valor inicial (`First Trial Value`) em branco ou divergente** entre a VF e a planilha de parâmetros.

Quando o pacote é entregue ao fornecedor do cluster com qualquer uma dessas inconsistências, o fornecedor abre **RARs (*Requests for Action / Clarification*)**, gerando retrabalho de engenharia, necessidade de *Change Requests (CRs)* e atrasos no cronograma de validação de software.

---

## 3. A Solução Desenvolvida
O **Auditor de Especificações IPC** é uma aplicação web local e ferramenta de linha de comando (CLI) em Python que realiza a **leitura automatizada e auditoria bidirecional** entre os documentos de VF (`.docx` ou `.pdf`, inclusive pacotes de **1.000+ páginas**) e as planilhas `Indications.xlsx` e `Parameters_Table.xlsx`.

### Principais Diferenciais da Solução
1. **100% Independente de Projeto:** Pode ser utilizado em qualquer programa automotivo (`363`, `376`, `J3U`, etc.) via **Upload Manual** no navegador ou **Varredura de Pasta**.
2. **Leitura Inteligente de VFs Extensas (1.000+ Páginas):** Ignora automaticamente dezenas de capítulos e tabelas distratoras (tabelas de sinais CAN-FD, máquinas de estado, códigos DTC, requisitos gerais) e extrai exclusivamente a seção `Indication Management` e a tabela `Configuration Parameters Table`.
3. **Normalização Automática de Nomenclatura:** Reconhece que `HAND BRAKE.V1` (com ponto no texto da VF) e `HAND BRAKE V1` (com espaço na planilha `Indications`) representam a mesma indicação, evitando alarmes falsos.
4. **Filtro Inteligente de Escopo de Parâmetros (`IPC` + `SUPPLIER EOL`):** Filtra apenas os parâmetros cujo módulo (`Component`) é `IPC` e cuja implementação (`Implementation`) é `SUPPLIER EOL` / `SUPPLIER EOL_2`, segregando automaticamente parâmetros `PROXI` (alocados em outro documento da Stellantis) e parâmetros de outras ECUs (`BCM`, `ORC`, `BSM`).
5. **Relatórios Prontos para Ação:** Gera instantaneamente relatórios executivos e técnicos em **Excel (`.xlsx` com 4 abas formatadas)** e **PDF (`.pdf`)**.

---

## 4. Resultados e Ganhos Esperados (ROI de Engenharia)
| Indicador | Processo Manual Anterior | Com o Auditor Automatizado IPC |
| :--- | :--- | :--- |
| **Tempo de conferência de um pacote de 1.200 páginas (45 VFs)** | 2 a 4 dias úteis de revisão manual | **Menos de 3 segundos** |
| **Risco de erro humano por fadiga visual** | Alto (centenas de linhas e páginas) | **Zero (checagem determinística 100% bidirecional)** |
| **Falsos positivos com parâmetros `PROXI` ou de outras ECUs** | Frequente em conferências manuais | **Eliminado pelo filtro automático `IPC + SUPPLIER EOL`** |
| **Prevenção de RARs antes do envio ao fornecedor** | Reativa (após abertura de RAR) | **100% Preventiva (antes da liberação oficial)** |
