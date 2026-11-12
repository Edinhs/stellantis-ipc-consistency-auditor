let CURRENT_REPORT = null;
let SELECTED_VF_FILTER = "ALL";
let RESOLVED_KEYS = new Set();

const SIMPLE_STATUS = {
  OK: { label: "✅ Tudo Certo (OK)", cls: "pill-ok" },
  MISSING_IN_SHEET: { label: "❌ Falta na Planilha", cls: "pill-danger" },
  ORPHAN_IN_SHEET: { label: "👻 Sobrando na Planilha (Órfão)", cls: "pill-danger" },
  VF_MISMATCH: { label: "⚠️ Código da VF Diferente", cls: "pill-warn" },
  VALUE_MISMATCH_OR_EMPTY: { label: "⚠️ Valor First Trial Vazio / Diferente", cls: "pill-warn" },
};

function showToast(msg) {
  const container = document.getElementById("toastContainer");
  if (!container) return;
  const el = document.createElement("div");
  el.className = "toast";
  el.textContent = msg;
  container.appendChild(el);
  setTimeout(() => {
    el.style.opacity = "0";
    setTimeout(() => el.remove(), 250);
  }, 3000);
}

function toggleTheme() {
  const html = document.documentElement;
  const next = html.getAttribute("data-theme") === "dark" ? "light" : "dark";
  html.setAttribute("data-theme", next);
  const btn = document.getElementById("themeToggleBtn");
  if (btn) btn.textContent = next === "dark" ? "☀️" : "🌙";
  localStorage.setItem("stellantis_ipc_theme", next);
}

function toggleUploadSection() {
  const sec = document.getElementById("uploadSection");
  if (!sec) return;
  sec.classList.toggle("hidden");
  sec.scrollIntoView({ behavior: "smooth" });
}

function updateFileBox(inputId, labelId, boxId) {
  const inp = document.getElementById(inputId);
  const lbl = document.getElementById(labelId);
  const box = document.getElementById(boxId);
  if (!inp || !lbl || !inp.files.length) return;
  if (inp.files.length === 1) {
    lbl.textContent = `✓ Selecionado: ${inp.files[0].name}`;
  } else {
    lbl.textContent = `✓ ${inp.files.length} arquivos selecionados`;
  }
  if (box) box.classList.add("file-selected");
}

function switchTab(tabId) {
  document.querySelectorAll(".tab-btn").forEach((b) => {
    b.classList.toggle("active", b.dataset.tab === tabId);
  });
  document.querySelectorAll(".tab-content").forEach((sec) => {
    sec.classList.toggle("active", sec.id === tabId);
  });
}

function setChipFilter(groupId, hiddenInputId, val, callbackFn) {
  const group = document.getElementById(groupId);
  if (group) {
    group.querySelectorAll(".filter-chip").forEach((btn) => {
      btn.classList.toggle("active", btn.dataset.val === val);
    });
  }
  const hidden = document.getElementById(hiddenInputId);
  if (hidden) hidden.value = val;
  if (typeof callbackFn === "function") callbackFn();
}

function selectGlobalVf(vfId) {
  SELECTED_VF_FILTER = vfId;
  renderGlobalVfPills();
  renderProblemsSummaryTable();
  renderRule1Table();
  renderRule2Table();
  renderVfsExplorer();
  renderFilteredTable();
}

async function runUploadAudit(e) {
  e.preventDefault();
  const vfFiles = document.getElementById("vfFilesInput").files;
  const indFile = document.getElementById("indFileInput").files[0];
  const parFile = document.getElementById("parFileInput").files[0];
  const projName = document.getElementById("projectNameInput").value.trim() || "Projeto IPC Auditado";

  if (!vfFiles.length || !indFile || !parFile) {
    alert("Por favor, selecione os 3 arquivos: VFs (.docx/.pdf), Planilha Indications (.xlsx) e Planilha Parameters (.xlsx).");
    return;
  }

  const formData = new FormData();
  for (const f of vfFiles) {
    formData.append("vf_files", f);
  }
  formData.append("indications_file", indFile);
  formData.append("parameters_file", parFile);
  formData.append("project_name", projName);

  setBannerLoading(`Lendo e analisando arquivos de "${projName}"...`);
  try {
    const res = await fetch("/api/audit/upload", {
      method: "POST",
      body: formData,
    });
    if (!res.ok) {
      const errObj = await res.json().catch(() => ({ detail: res.statusText }));
      throw new Error(errObj.detail || "Falha ao ler arquivos enviados");
    }
    const data = await res.json();
    RESOLVED_KEYS.clear();
    applyReport(data);
    showToast(`Arquivos do projeto "${projName}" analisados com sucesso!`);
  } catch (err) {
    alert("Erro ao analisar arquivos: " + err.message);
  }
}

async function runFolderAudit() {
  const folderPath = document.getElementById("folderPathInput").value.trim();
  const projName = document.getElementById("projectNameInput").value.trim() || folderPath.split(/[\\/]/).pop() || "Projeto Auditado";
  if (!folderPath) {
    alert("Informe o caminho da pasta do projeto.");
    return;
  }
  setBannerLoading(`Lendo pasta ${folderPath}...`);
  try {
    const res = await fetch("/api/audit/folder", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        folder_path: folderPath,
        prefer_format: "docx",
        project_name: projName,
      }),
    });
    if (!res.ok) {
      const errObj = await res.json().catch(() => ({ detail: res.statusText }));
      throw new Error(errObj.detail || "Falha ao ler diretório");
    }
    const data = await res.json();
    RESOLVED_KEYS.clear();
    applyReport(data);
    showToast("Pasta lida e analisada com sucesso!");
  } catch (err) {
    alert("Erro ao ler pasta: " + err.message);
  }
}

async function runMockAudit(preferFormat = "docx") {
  setBannerLoading("Carregando exemplo rápido de demonstração (6 VFs)...");
  try {
    const res = await fetch(`/api/audit/mock?prefer_format=${encodeURIComponent(preferFormat)}`, {
      method: "POST",
    });
    if (!res.ok) throw new Error(await res.text());
    const data = await res.json();
    data.project_name = "Projeto de Demonstração (Amostra Rápida - 6 VFs)";
    RESOLVED_KEYS.clear();
    applyReport(data);
    showToast("Exemplo rápido (6 VFs) carregado!");
  } catch (err) {
    alert("Erro ao carregar exemplo: " + err.message);
  }
}

async function runLargeScaleAudit(preferFormat = "pdf") {
  setBannerLoading("Lendo exemplo completo de 1.215 páginas (45 VFs)...");
  try {
    const res = await fetch(`/api/audit/large?prefer_format=${encodeURIComponent(preferFormat)}`, {
      method: "POST",
    });
    if (!res.ok) throw new Error(await res.text());
    const data = await res.json();
    data.project_name = "Projeto de Demonstração Completo (VF List de 1.215 Páginas / 45 VFs)";
    RESOLVED_KEYS.clear();
    applyReport(data);
    const secs = data.files_analyzed?.elapsed_seconds ?? "";
    showToast(`1.215 páginas (45 VFs) lidas e analisadas em ${secs}s!`);
  } catch (err) {
    alert("Erro ao carregar exemplo de 1.215 páginas: " + err.message);
  }
}

function setBannerLoading(msg) {
  document.getElementById("metaMode").textContent = "LENDO ARQUIVOS...";
  document.getElementById("metaProjectName").textContent = msg;
}

function getUniqueProblemsList() {
  if (!CURRENT_REPORT) return [];
  // Evita duplicar o mesmo VF_MISMATCH ou VALUE_MISMATCH que aparece nas duas direções
  const problems = [];

  for (const f of CURRENT_REPORT.rule1_findings) {
    if (f.status === "OK") continue;
    if (f.status === "VF_MISMATCH" && f.direction === "SHEET_TO_VF") continue;
    problems.push({
      area: "🚨 Indicações (HMI)",
      rule: "RULE_1",
      finding: f,
    });
  }

  for (const f of CURRENT_REPORT.rule2_findings) {
    if (f.status === "OK") continue;
    if (f.status === "VALUE_MISMATCH_OR_EMPTY" && f.direction === "SHEET_TO_VF") continue;
    problems.push({
      area: "⚙️ Parâmetros (SW EOL)",
      rule: "RULE_2",
      finding: f,
    });
  }

  return problems;
}

function getFindingKey(f) {
  return `${f.rule}|${f.normalized_name}|${f.vf_id}|${f.status}`;
}

function toggleResolveFinding(key) {
  if (RESOLVED_KEYS.has(key)) {
    RESOLVED_KEYS.delete(key);
  } else {
    RESOLVED_KEYS.add(key);
  }
  updateKpiCards();
  renderProblemsSummaryTable();
}

function updateKpiCards() {
  if (!CURRENT_REPORT) return;
  const k = CURRENT_REPORT.summary_kpis;
  const allProblems = getUniqueProblemsList();
  const resolvedCnt = allProblems.filter((p) => RESOLVED_KEYS.has(getFindingKey(p.finding))).length;
  const pendingProblems = Math.max(0, k.total_unique_rar_risks - resolvedCnt);

  // Card 1: Total de Problemas
  document.getElementById("kpiTotalRar").textContent = `${pendingProblems} problema(s)`;
  document.getElementById("kpiRarBreakdown").textContent =
    resolvedCnt > 0
      ? `${resolvedCnt} já marcado(s) como resolvido(s) de ${k.total_unique_rar_risks} encontrados`
      : `${k.rule1_unique_issues} em Indicações + ${k.rule2_unique_issues} em Parâmetros`;

  // Card 2: Indicações
  const pillR1 = document.getElementById("pillRule1Status");
  pillR1.textContent = k.rule1_unique_issues === 0 ? "✅ Tudo OK" : `⚠️ ${k.rule1_unique_issues} para corrigir`;
  pillR1.className = `simple-pill ${k.rule1_unique_issues === 0 ? "pill-ok" : "pill-warn"}`;
  document.getElementById("kpiRule1Issues").textContent = `${k.rule1_unique_issues} divergência(s)`;
  document.getElementById("kpiRule1SimpleList").innerHTML = `
    <div>• <strong>${k.rule1_missing_in_sheet}</strong> estão na VF, mas <strong>faltam na planilha</strong></div>
    <div>• <strong>${k.rule1_orphan_in_sheet}</strong> estão <strong>sobrando na planilha</strong> (sem VF)</div>
    <div>• <strong>${k.rule1_vf_mismatch}</strong> estão com <strong>código de VF diferente</strong></div>
  `;

  // Card 3: Parâmetros
  const pillR2 = document.getElementById("pillRule2Status");
  pillR2.textContent = k.rule2_unique_issues === 0 ? "✅ Tudo OK" : `⚠️ ${k.rule2_unique_issues} para corrigir`;
  pillR2.className = `simple-pill ${k.rule2_unique_issues === 0 ? "pill-ok" : "pill-warn"}`;
  document.getElementById("kpiRule2Issues").textContent = `${k.rule2_unique_issues} divergência(s)`;
  document.getElementById("kpiRule2SimpleList").innerHTML = `
    <div>• <strong>${k.rule2_missing_in_sheet}</strong> estão na VF, mas <strong>faltam na planilha</strong></div>
    <div>• <strong>${k.rule2_orphan_in_sheet}</strong> estão <strong>sobrando na planilha</strong> (sem VF)</div>
    <div>• <strong>${k.rule2_value_mismatch}</strong> estão com <strong>valor First Trial vazio/diferente</strong></div>
  `;

  // Card 4: O que o sistema leu
  document.getElementById("kpiConformityBadge").textContent = `${k.conformity_rate_pct}% Conforme`;
  document.getElementById("kpiVfsCount").textContent = `${k.total_vfs_analyzed} VFs lidas`;
  document.getElementById("kpiReadSummaryList").innerHTML = `
    <div>• <strong>${k.total_vf_indications}</strong> indicações nas VFs vs. <strong>${k.total_sheet_indications}</strong> na planilha</div>
    <div>• <strong>${k.total_vf_params_ipc_eol}</strong> parâmetros EOL nas VFs vs. <strong>${k.total_sheet_params}</strong> na planilha</div>
    <div>• <strong>${k.total_vf_params_filtered_out}</strong> parâmetros ignorados (PROXI / outros módulos)</div>
  `;

  // Badges nas abas
  document.getElementById("badgeProblems").textContent = pendingProblems;
  document.getElementById("badgeRule1").textContent = k.total_vf_indications + k.rule1_orphan_in_sheet;
  document.getElementById("badgeRule2").textContent = k.total_vf_params_ipc_eol + k.rule2_orphan_in_sheet;
  document.getElementById("badgeVfs").textContent = k.total_vfs_analyzed;
  document.getElementById("badgeFiltered").textContent = k.total_vf_params_filtered_out;
}

function renderGlobalVfPills() {
  if (!CURRENT_REPORT) return;
  const container = document.getElementById("globalVfPills");
  const allBtn = `
    <button class="vf-pill-btn ${SELECTED_VF_FILTER === "ALL" ? "active" : ""}" onclick="selectGlobalVf('ALL')">
      <span>Mostrar Todas (${CURRENT_REPORT.vfs_overview.length} VFs)</span>
    </button>
  `;
  const vfBtns = CURRENT_REPORT.vfs_overview
    .map((vf) => {
      const issuesCnt = vf.indications_issues + vf.parameters_issues;
      const dotCls = issuesCnt === 0 ? "ok" : "issue";
      const activeCls = SELECTED_VF_FILTER === vf.vf_id ? "active" : "";
      return `
        <button class="vf-pill-btn ${activeCls}" onclick="selectGlobalVf('${escapeHtml(vf.vf_id)}')" title="${escapeHtml(vf.title)}">
          <span class="vf-pill-dot ${dotCls}"></span>
          <span>${escapeHtml(vf.vf_id)}</span>
        </button>
      `;
    })
    .join("");
  container.innerHTML = allBtn + vfBtns;
}

function applyReport(report) {
  CURRENT_REPORT = report;
  SELECTED_VF_FILTER = "ALL";
  const fInfo = report.files_analyzed || {};

  document.getElementById("metaMode").textContent = "PROJETO CARREGADO";
  document.getElementById("metaProjectName").textContent = `${report.project_name}`;
  document.getElementById("metaFilesInfo").innerHTML =
    `<strong>Planilha Indicações:</strong> <code>${escapeHtml(fInfo.indications_file || "-")}</code> &nbsp;|&nbsp; ` +
    `<strong>Planilha Parâmetros:</strong> <code>${escapeHtml(fInfo.parameters_file || "-")}</code>`;

  updateKpiCards();
  renderGlobalVfPills();
  renderProblemsSummaryTable();
  renderRule1Table();
  renderRule2Table();
  renderVfsExplorer();
  renderFilteredTable();
}

function matchesGlobalVf(f) {
  if (SELECTED_VF_FILTER === "ALL") return true;
  return (
    f.vf_id.includes(SELECTED_VF_FILTER) ||
    (f.sheet_vf_ref && f.sheet_vf_ref.toUpperCase().includes(SELECTED_VF_FILTER))
  );
}

function renderProblemsSummaryTable() {
  if (!CURRENT_REPORT) return;
  const search = (document.getElementById("probSearch")?.value || "").trim().toLowerCase();
  const problems = getUniqueProblemsList().filter((item) => {
    const f = item.finding;
    if (!matchesGlobalVf(f)) return false;
    if (search) {
      const hay = `${item.area} ${f.item_name} ${f.vf_id} ${f.sheet_vf_ref} ${f.details} ${f.recommendation}`.toLowerCase();
      if (!hay.includes(search)) return false;
    }
    return true;
  });

  const tbody = document.getElementById("problemsTableBody");
  if (!problems.length) {
    tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;padding:30px;color:var(--ok);font-weight:700;">🎉 Nenhum problema encontrado para este filtro! Tudo está 100% consistente.</td></tr>`;
    return;
  }

  tbody.innerHTML = problems
    .map((item) => {
      const f = item.finding;
      const st = SIMPLE_STATUS[f.status] || { label: f.status, cls: "pill-warn" };
      const key = getFindingKey(f);
      const isResolved = RESOLVED_KEYS.has(key);
      const vfDisplay = f.vf_id !== "NENHUMA" ? f.vf_id : `Planilha: ${f.sheet_vf_ref}`;

      return `
        <tr class="audit-row ${isResolved ? "resolved-row" : ""}">
          <td style="text-align:center;">
            <input type="checkbox" class="check-resolve" ${isResolved ? "checked" : ""} onchange="toggleResolveFinding('${escapeHtml(key)}')" title="Marcar como corrigido" />
          </td>
          <td><span class="area-badge">${item.area}</span></td>
          <td><span class="status-pill ${st.cls}">${st.label}</span></td>
          <td><strong><code>${escapeHtml(f.item_name)}</code></strong></td>
          <td><strong>${escapeHtml(vfDisplay)}</strong></td>
          <td>${escapeHtml(f.details)}</td>
          <td><strong>${escapeHtml(f.recommendation)}</strong></td>
        </tr>
      `;
    })
    .join("");
}

function renderRule1Table() {
  if (!CURRENT_REPORT) return;
  const statusFilter = document.getElementById("r1FilterStatus").value;
  const search = document.getElementById("r1Search").value.trim().toLowerCase();

  // Deduplicar itens OK e VF_MISMATCH que aparecem nas duas direções para deixar a leitura limpa
  const rows = CURRENT_REPORT.rule1_findings.filter((f) => {
    if (f.direction === "SHEET_TO_VF" && f.status !== "ORPHAN_IN_SHEET") return false;
    if (!matchesGlobalVf(f)) return false;
    if (statusFilter === "ISSUES_ONLY" && f.status === "OK") return false;
    if (statusFilter !== "ALL" && statusFilter !== "ISSUES_ONLY" && f.status !== statusFilter) return false;
    if (search) {
      const hay = `${f.item_name} ${f.normalized_name} ${f.vf_id} ${f.sheet_vf_ref} ${f.details}`.toLowerCase();
      if (!hay.includes(search)) return false;
    }
    return true;
  });

  const tbody = document.getElementById("r1TableBody");
  if (!rows.length) {
    tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;padding:26px;color:var(--text-muted);">Nenhum registro encontrado para este filtro.</td></tr>`;
    return;
  }

  tbody.innerHTML = rows
    .map((f) => {
      const st = SIMPLE_STATUS[f.status] || { label: f.status, cls: "pill-warn" };
      return `
        <tr>
          <td><span class="status-pill ${st.cls}">${st.label}</span></td>
          <td><strong>${escapeHtml(f.item_name)}</strong></td>
          <td><strong>${escapeHtml(f.vf_id)}</strong></td>
          <td><code>${escapeHtml(f.sheet_vf_ref)}</code></td>
          <td>${escapeHtml(f.details)}</td>
          <td>${escapeHtml(f.recommendation)}</td>
        </tr>
      `;
    })
    .join("");
}

function renderRule2Table() {
  if (!CURRENT_REPORT) return;
  const statusFilter = document.getElementById("r2FilterStatus").value;
  const search = document.getElementById("r2Search").value.trim().toLowerCase();

  const rows = CURRENT_REPORT.rule2_findings.filter((f) => {
    if (f.direction === "SHEET_TO_VF" && f.status !== "ORPHAN_IN_SHEET") return false;
    if (!matchesGlobalVf(f)) return false;
    if (statusFilter === "ISSUES_ONLY" && f.status === "OK") return false;
    if (statusFilter !== "ALL" && statusFilter !== "ISSUES_ONLY" && f.status !== statusFilter) return false;
    if (search) {
      const hay = `${f.item_name} ${f.vf_id} ${f.sheet_vf_ref} ${f.details}`.toLowerCase();
      if (!hay.includes(search)) return false;
    }
    return true;
  });

  const tbody = document.getElementById("r2TableBody");
  if (!rows.length) {
    tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;padding:26px;color:var(--text-muted);">Nenhum registro encontrado para este filtro.</td></tr>`;
    return;
  }

  tbody.innerHTML = rows
    .map((f) => {
      const st = SIMPLE_STATUS[f.status] || { label: f.status, cls: "pill-warn" };
      const vfVal = f.extra?.vf_first_trial ?? "-";
      const shVal = f.extra?.sheet_first_trial ?? "-";
      const unit = f.extra?.unit ? ` ${f.extra.unit}` : "";
      return `
        <tr>
          <td><span class="status-pill ${st.cls}">${st.label}</span></td>
          <td><strong><code>${escapeHtml(f.item_name)}</code></strong></td>
          <td><strong>${escapeHtml(f.vf_id)}</strong></td>
          <td>
            <div>Na VF: <strong>${escapeHtml(String(vfVal))}${escapeHtml(unit)}</strong></div>
            <div style="font-size:0.78rem;color:var(--text-secondary);">Na Planilha: <strong>${escapeHtml(String(shVal))}</strong></div>
          </td>
          <td>${escapeHtml(f.details)}</td>
          <td>${escapeHtml(f.recommendation)}</td>
        </tr>
      `;
    })
    .join("");
}

function renderVfsExplorer() {
  if (!CURRENT_REPORT) return;
  const container = document.getElementById("vfsExplorerContainer");
  const filteredVfs = CURRENT_REPORT.vfs_overview.filter(
    (vf) => SELECTED_VF_FILTER === "ALL" || vf.vf_id === SELECTED_VF_FILTER
  );

  container.innerHTML = filteredVfs
    .map((vf) => {
      const statusBadge =
        vf.status === "OK"
          ? `<span class="status-pill pill-ok">✅ Sem Divergências</span>`
          : `<span class="status-pill pill-danger">⚠️ ${vf.indications_issues + vf.parameters_issues} Problema(s)</span>`;

      const indRows = vf.indications
        .map(
          (i) => `
          <tr>
            <td><strong>${escapeHtml(i.raw_name)}</strong></td>
            <td>${escapeHtml(i.condition)}</td>
            <td>${escapeHtml(i.telltale)}</td>
          </tr>`
        )
        .join("");

      const parRows = vf.parameters.length
        ? vf.parameters
            .map((p) => {
              const scopeBadge = p.is_ipc_supplier_eol
                ? `<span class="status-pill pill-ok">Auditado (IPC + EOL)</span>`
                : `<span class="status-pill pill-warn">Ignorado (${escapeHtml(p.component)} / ${escapeHtml(p.implementation)})</span>`;
              return `
              <tr>
                <td><code>${escapeHtml(p.name)}</code></td>
                <td>${escapeHtml(p.first_trial_value)} ${escapeHtml(p.unit)}</td>
                <td>${escapeHtml(p.component)}</td>
                <td>${escapeHtml(p.implementation)}</td>
                <td>${scopeBadge}</td>
              </tr>`;
            })
            .join("")
        : `<tr><td colspan="5" style="text-align:center;color:var(--text-muted);padding:10px;">Esta VF não possui parâmetros de configuração (tabela vazia).</td></tr>`;

      return `
        <div class="vf-card">
          <div class="vf-card-header">
            <div>
              <h4>${escapeHtml(vf.title)}</h4>
              <div style="font-size:0.78rem;color:var(--text-muted);margin-top:2px;">
                Arquivo lido: <code>${escapeHtml(vf.source_file)}</code>
              </div>
            </div>
            ${statusBadge}
          </div>

          <div style="font-size:0.82rem;font-weight:700;margin-bottom:4px;">Indicações chamadas nesta VF (${vf.indications_count})</div>
          <table class="mini-table">
            <thead>
              <tr><th>Nome na VF</th><th>Quando Ativa (Condição)</th><th>Comportamento Telltale</th></tr>
            </thead>
            <tbody>${indRows}</tbody>
          </table>

          <div style="font-size:0.82rem;font-weight:700;margin-bottom:4px;">
            Parâmetros desta VF (${vf.parameters_ipc_eol} auditados no IPC · ${vf.parameters_filtered_out} ignorados)
          </div>
          <table class="mini-table">
            <thead>
              <tr><th>Parâmetro</th><th>Valor First Trial</th><th>Módulo</th><th>Tipo</th><th>Situação</th></tr>
            </thead>
            <tbody>${parRows}</tbody>
          </table>
        </div>
      `;
    })
    .join("");
}

function renderFilteredTable() {
  if (!CURRENT_REPORT) return;
  const tbody = document.getElementById("filteredTableBody");
  const list = (CURRENT_REPORT.filtered_out_parameters || []).filter(
    (p) => SELECTED_VF_FILTER === "ALL" || p.vf_id === SELECTED_VF_FILTER
  );
  if (!list.length) {
    tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;padding:24px;color:var(--text-muted);">Nenhum parâmetro ignorado para este filtro.</td></tr>`;
    return;
  }
  tbody.innerHTML = list
    .map(
      (p) => `
      <tr>
        <td><strong>${escapeHtml(p.vf_id)}</strong></td>
        <td><strong><code>${escapeHtml(p.name)}</code></strong></td>
        <td>${escapeHtml(p.first_trial_value)} ${escapeHtml(p.unit)}</td>
        <td><span class="area-badge">${escapeHtml(p.component)}</span></td>
        <td><span class="status-pill pill-warn">${escapeHtml(p.implementation)}</span></td>
        <td>${escapeHtml(p.filter_reason)}</td>
      </tr>
    `
    )
    .join("");
}

function escapeHtml(str) {
  return String(str ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

window.addEventListener("DOMContentLoaded", () => {
  const savedTheme = localStorage.getItem("stellantis_ipc_theme");
  if (savedTheme === "dark") {
    document.documentElement.setAttribute("data-theme", "dark");
    const btn = document.getElementById("themeToggleBtn");
    if (btn) btn.textContent = "☀️";
  }
  runMockAudit("docx");
});
