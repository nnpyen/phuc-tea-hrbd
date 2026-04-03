/* global window, document, localStorage */

(function initApp() {
  "use strict";

  const STORAGE_KEY = "hrbd_dashboard_2026_v2";
  const LEGACY_STORAGE_KEY = "hrbd_dashboard_2026_v1";
  const UI_KEY = "hrbd_dashboard_2026_ui_v2";
  const CLOUD_CFG_KEY = "hrbd_dashboard_2026_cloud_v1";
  const ALL_MONTH_VALUE = "ALL";

  const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const QUARTER_MONTHS = {
    Q1: ["Jan", "Feb", "Mar"],
    Q2: ["Apr", "May", "Jun"],
    Q3: ["Jul", "Aug", "Sep"],
    Q4: ["Oct", "Nov", "Dec"],
  };

  const PAGE_META = {
    dashboard: {
      title: "Dashboard Tổng",
      subtitle: "Tổng hợp OKR năm, Target vs Actual theo tháng, quý, năm.",
    },
    "quarter-plan": {
      title: "Kế Hoạch Quý",
      subtitle: "Theo dõi mục tiêu quý, tiến độ thực tế và hành động ưu tiên.",
    },
    "pillar-hr": {
      title: "Human Resource",
      subtitle: "Chi tiết KPI Human Resource theo tháng/quý và phân tích định hướng.",
    },
    "pillar-vn": {
      title: "National Franchise Development",
      subtitle: "Theo dõi pipeline, phát triển cửa hàng và chăm sóc đối tác.",
    },
    "pillar-intl": {
      title: "International Franchise Development",
      subtitle: "Giám sát triển khai quốc tế, rủi ro vận hành và SLA.",
    },
    "import-center": {
      title: "Trang Import",
      subtitle: "Nhập OKR/KPI theo năm-quý-tháng làm nguồn dữ liệu gốc cho toàn bộ dashboard.",
    },
    "partner-info": {
      title: "Thông Tin Đối Tác",
      subtitle: "Đồng bộ CRM_IMPORT realtime và theo dõi summary đối tác cần chăm sóc.",
    },
    "partner-care": {
      title: "Kế Hoạch Chăm Sóc Đối Tác",
      subtitle: "Theo dõi touchpoint chi tiết đến từng đối tác và hiệu quả thực thi.",
    },
    "intl-problem": {
      title: "International Problem Record",
      subtitle: "Quản trị sự cố quốc tế theo root cause, SLA và hành động khắc phục.",
    },
  };

  const PILLARS = {
    HR: {
      code: "HR",
      label: "Human Resource",
      color: "#0f8c95",
      weightKey: "HR",
    },
    VN: {
      code: "VN",
      label: "National Franchise Development",
      color: "#1aaab4",
      weightKey: "VN",
    },
    INTL: {
      code: "INTL",
      label: "International Franchise Development",
      color: "#7e88e4",
      weightKey: "INTL",
    },
  };

  const baseData = deepClone(window.__PLAN_2026_DATA__ || {});
  let state = loadState(baseData);
  let uiState = loadUiState();
  let importCareDraft = { month: "", rows: [] };
  let cloudConfig = loadCloudConfig();
  let supabaseClient = null;
  let cloudPushTimer = null;
  let cloudPushInFlight = false;
  let queuedPushReason = "";
  let suppressCloudPush = false;

  ensureStateShape();
  buildPillarPages();
  initSelectors();
  initNav();
  initImportHandlers();
  initCloudSyncHandlers();
  initPartnerInfoHandlers();
  initPartnerCareHandlers();
  initIntlProblemHandlers();
  initGlobalHandlers();
  initPillarKpiHandlers();
  renderAll();
  tryCloudBootstrap();

  function ensureStateShape() {
    if (!state.meta) state.meta = {};
    if (!state.meta.weights) state.meta.weights = { HR: 0.45, VN: 0.45, INTL: 0.1 };
    if (!state.settings) state.settings = { thresholds: { good: 0.9, watch: 0.75 } };
    if (!state.settings.thresholds) state.settings.thresholds = { good: 0.9, watch: 0.75 };
    if (!state.annualPlan) state.annualPlan = [];
    if (!state.quarterPlan) state.quarterPlan = [];
    if (!state.scorecards) state.scorecards = {};
    if (!state.partnerCare) state.partnerCare = { monthlyPlan: [], logs: [], partnerMaster: [], partnerMonthlyPlans: [], sync: {} };
    if (!state.partnerCare.monthlyPlan) state.partnerCare.monthlyPlan = [];
    if (!state.partnerCare.logs) state.partnerCare.logs = [];
    if (!state.partnerCare.partnerMaster) state.partnerCare.partnerMaster = [];
    if (!state.partnerCare.partnerMonthlyPlans) state.partnerCare.partnerMonthlyPlans = [];
    if (!state.partnerCare.crmImportRows) state.partnerCare.crmImportRows = [];
    if (!state.partnerCare.crmSummary) state.partnerCare.crmSummary = {};
    if (!state.partnerCare.sync) state.partnerCare.sync = {};
    if (!cleanText(state.partnerCare.sync.endpoint)) state.partnerCare.sync.endpoint = "";
    if (!cleanText(state.partnerCare.sync.lastSyncAt)) state.partnerCare.sync.lastSyncAt = "";
    if (asNumber(state.partnerCare.sync.lastCount) === null) state.partnerCare.sync.lastCount = 0;
    if (!state.intlProblems) state.intlProblems = [];
    if (!state.version) state.version = 1;

    if (!state.importHub) {
      state.importHub = defaultImportHub();
    }

    ["HR", "VN", "INTL"].forEach((key) => {
      if (!state.scorecards[key]) state.scorecards[key] = { kpiCodes: [], months: [] };
      if (!state.scorecards[key].kpiCodes) state.scorecards[key].kpiCodes = [];
      if (!state.scorecards[key].months) state.scorecards[key].months = [];
      MONTHS.forEach((month) => {
        if (!state.scorecards[key].months.find((r) => r.month === month)) {
          state.scorecards[key].months.push({ month, quarter: quarterOf(month), pillarScore: null, kpis: {} });
        }
      });
    });

    migrateLegacyToImportHubIfNeeded();
    normalizePartnerCareData();
    normalizeCrmImportData();
    const cleanedLegacyAnnual = normalizeImportHub();
    if (cleanedLegacyAnnual) {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
      } catch (error) {
        // no-op
      }
    }

    if (cleanText(uiState.month).toUpperCase() === "ALL") uiState.month = ALL_MONTH_VALUE;
    if (!uiState.currentPage || !PAGE_META[uiState.currentPage]) uiState.currentPage = "dashboard";
    if (!uiState.month || (uiState.month !== ALL_MONTH_VALUE && !MONTHS.includes(uiState.month))) {
      uiState.month = guessDefaultMonth();
    }
    if (!uiState.quarter || !QUARTER_MONTHS[uiState.quarter]) uiState.quarter = quarterOf(uiState.month);
  }

  function defaultImportHub() {
    return {
      annualOkrs: [],
      quarterlyOkrs: [],
      monthlyOkrs: [],
      monthlyPlans: [],
      monthlyKpis: [],
    };
  }

  function isLegacySeededAnnualOkr(row) {
    const id = cleanText(row?.id);
    const code = cleanText(row?.code).toUpperCase();
    return /^A-\d+-/.test(id) && /^(HR|VN|INT|INTL)\d+$/.test(code);
  }

  function migrateLegacyToImportHubIfNeeded() {
    const hub = state.importHub;
    if (hub.annualOkrs.length || hub.quarterlyOkrs.length || hub.monthlyKpis.length) return;

    state.quarterPlan.forEach((item, index) => {
      const quarter = normalizeQuarter(item.quarter);
      const pillarKey = inferPillarFromLabel(item.pillar, item.kpiName || item.objective);
      const okrId = `QO-${index + 1}-${safeId(item.kpiName || item.objective || "QOKR")}`;
      hub.quarterlyOkrs.push({
        id: okrId,
        quarter,
        pillarKey,
        code: `Q-${index + 1}`,
        title: cleanText(item.kpiName) || cleanText(item.objective) || `Quarterly OKR ${index + 1}`,
        target: asNumber(item.targetQuarter),
        unit: cleanText(item.unit) || "%",
        linkedAnnualOkrId: "",
        notes: cleanText(item.notes),
      });
    });

    const annualMap = {};
    state.annualPlan.forEach((item) => {
      annualMap[item.kpiCode] = item;
    });

    ["HR", "VN", "INTL"].forEach((pillarKey) => {
      const scorecard = state.scorecards[pillarKey];
      MONTHS.forEach((month) => {
        const monthRow = scorecard.months.find((m) => m.month === month) || { kpis: {} };
        const monthOkrId = `MO-${pillarKey}-${month}`;
        const quarter = quarterOf(month);
        const linkedQuarterly = hub.quarterlyOkrs.find(
          (row) => row.pillarKey === pillarKey && row.quarter === quarter
        );
        let targetSum = 0;
        let hasAnyKpi = false;
        (scorecard.kpiCodes || []).forEach((code) => {
          const cell = monthRow.kpis?.[code] || {};
          const target = asNumber(cell.target);
          const actual = asNumber(cell.actual);
          if (target === null && actual === null) return;
          hasAnyKpi = true;
          targetSum += target || 0;
          hub.monthlyKpis.push({
            id: `MK-${pillarKey}-${month}-${safeId(code)}`,
            month,
            pillarKey,
            monthlyOkrId: monthOkrId,
            kpiCode: code,
            kpiName: cleanText(annualMap[code]?.kpiName) || code,
            target,
            actual,
            unit: cleanText(annualMap[code]?.unit) || "%",
            owner: cleanText(annualMap[code]?.owner),
            status: "",
            result: "",
            notes: "",
          });
        });

        if (targetSum > 0 || hasAnyKpi) {
          hub.monthlyOkrs.push({
            id: monthOkrId,
            month,
            pillarKey,
            title: `${PILLARS[pillarKey].label} ${month}`,
            target: targetSum > 0 ? targetSum : null,
            unit: "points",
            linkedQuarterlyOkrId: linkedQuarterly?.id || "",
            notes: "",
          });
        }
      });
    });
  }

  function normalizeImportHub() {
    let cleanedLegacyAnnual = false;
    const hub = state.importHub;
    ["annualOkrs", "quarterlyOkrs", "monthlyOkrs", "monthlyPlans", "monthlyKpis"].forEach((key) => {
      if (!Array.isArray(hub[key])) hub[key] = [];
    });

    hub.annualOkrs = hub.annualOkrs.map((row, idx) => ({
      id: cleanText(row.id) || `A-${idx + 1}-${Date.now()}`,
      pillarKey: normalizePillar(row.pillarKey || row.pillar || "HR"),
      code: cleanText(row.code),
      title: cleanText(row.title),
      target: asNumber(row.target),
      unit: cleanText(row.unit) || "%",
      notes: cleanText(row.notes),
    }));

    const removedLegacyAnnualIds = new Set(
      hub.annualOkrs.filter((row) => isLegacySeededAnnualOkr(row)).map((row) => cleanText(row.id))
    );
    if (removedLegacyAnnualIds.size) {
      cleanedLegacyAnnual = true;
      hub.annualOkrs = hub.annualOkrs.filter((row) => !removedLegacyAnnualIds.has(cleanText(row.id)));
    }

    hub.quarterlyOkrs = hub.quarterlyOkrs.map((row, idx) => ({
      id: cleanText(row.id) || `QO-${idx + 1}-${Date.now()}`,
      quarter: normalizeQuarter(row.quarter),
      pillarKey: normalizePillar(row.pillarKey || row.pillar || "HR"),
      code: cleanText(row.code),
      title: cleanText(row.title),
      target: asNumber(row.target),
      unit: cleanText(row.unit) || "%",
      linkedAnnualOkrId: removedLegacyAnnualIds.has(cleanText(row.linkedAnnualOkrId))
        ? ""
        : cleanText(row.linkedAnnualOkrId),
      notes: cleanText(row.notes),
    }));

    hub.monthlyOkrs = hub.monthlyOkrs.map((row, idx) => ({
      id: cleanText(row.id) || `MO-${idx + 1}-${Date.now()}`,
      month: normalizeMonth(row.month),
      pillarKey: normalizePillar(row.pillarKey || row.pillar || "HR"),
      title: cleanText(row.title),
      target: asNumber(row.target),
      unit: cleanText(row.unit) || "%",
      linkedQuarterlyOkrId: cleanText(row.linkedQuarterlyOkrId),
      notes: cleanText(row.notes),
    }));

    hub.monthlyPlans = hub.monthlyPlans.map((row, idx) => ({
      id: cleanText(row.id) || `MP-${idx + 1}-${Date.now()}`,
      month: normalizeMonth(row.month),
      pillarKey: normalizePillar(row.pillarKey || row.pillar || "HR"),
      linkedMonthlyOkrId: cleanText(row.linkedMonthlyOkrId),
      action: cleanText(row.action),
      owner: cleanText(row.owner),
      due: cleanText(row.due),
      status: cleanText(row.status),
      result: cleanText(row.result || row.notes),
      notes: cleanText(row.notes),
    }));

    hub.monthlyKpis = hub.monthlyKpis.map((row, idx) => ({
      id: cleanText(row.id) || `MK-${idx + 1}-${Date.now()}`,
      month: normalizeMonth(row.month),
      pillarKey: normalizePillar(row.pillarKey || row.pillar || "HR"),
      monthlyOkrId: cleanText(row.monthlyOkrId),
      kpiCode: cleanText(row.kpiCode),
      kpiName: cleanText(row.kpiName),
      target: asNumber(row.target),
      actual: asNumber(row.actual),
      unit: cleanText(row.unit) || "%",
      owner: cleanText(row.owner),
      status: cleanText(row.status),
      result: cleanText(row.result || row.outcome || row.notes),
      notes: cleanText(row.notes),
    }));

    return cleanedLegacyAnnual;
  }

  function guessDefaultMonth() {
    const monthIndex = new Date().getMonth();
    const candidate = MONTHS[monthIndex] || "Jan";
    if (aggregateByMonths([candidate]).targetWeight > 0) return candidate;
    for (let i = MONTHS.length - 1; i >= 0; i -= 1) {
      const month = MONTHS[i];
      if (aggregateByMonths([month]).targetWeight > 0) return month;
    }
    return "Jan";
  }

  function normalizePartnerCareData() {
    const care = state.partnerCare;
    if (!care) return;

    care.partnerMaster = (care.partnerMaster || [])
      .map((row) => normalizePartnerMasterRow(row))
      .filter((row) => cleanText(row.partnerKey));

    care.partnerMonthlyPlans = (care.partnerMonthlyPlans || [])
      .map((row, idx) => ({
        id: cleanText(row.id) || `PCP-${idx + 1}-${Date.now()}`,
        month: normalizeMonth(row.month),
        careDay: normalizeCareDay(row.careDay ?? row.day ?? row.dayCare ?? row.date),
        dayDone: normalizeCareDay(row.dayDone),
        partnerKey: cleanText(row.partnerKey).toUpperCase(),
        partnerName: cleanText(row.partnerName),
        partnerType: normalizePartnerType(row.partnerType),
        touchpointType: cleanText(row.touchpointType),
        channel: cleanText(row.channel),
        owner: cleanText(row.owner),
        targetTouchpoints: Math.max(0, asNumber(row.targetTouchpoints) ?? 1),
        status: cleanText(row.status),
        result: cleanText(row.result || row.notes),
        response: cleanText(row.response),
        nextAction: cleanText(row.nextAction),
      }))
      .filter((row) => cleanText(row.partnerKey));
  }

  function normalizeCrmImportData() {
    const care = state.partnerCare;
    if (!care) return;
    care.crmImportRows = (care.crmImportRows || [])
      .map((row, idx) => normalizeCrmImportRow(row, idx))
      .filter((row) => cleanText(row.partnerCode) || cleanText(row.partnerKey));

    const summary = care.crmSummary || {};
    care.crmSummary = {
      totalPartner: asNumber(summary.totalPartner) || 0,
      existingEligible: asNumber(summary.existingEligible) || 0,
      newCount: asNumber(summary.newCount) || 0,
      noTouchThisYear: asNumber(summary.noTouchThisYear) || 0,
      activePartner: asNumber(summary.activePartner) || 0,
    };
  }

  function buildPillarPages() {
    const containers = {
      HR: document.getElementById("page-pillar-hr"),
      VN: document.getElementById("page-pillar-vn"),
      INTL: document.getElementById("page-pillar-intl"),
    };

    Object.values(PILLARS).forEach((pillar) => {
      const el = containers[pillar.code];
      if (!el) return;
      el.innerHTML = `
        <article class="panel glass">
          <div class="panel-head">
            <h3>OKR ${pillar.label}: Target vs Actual</h3>
            <p>Mỗi OKR là 1 dòng tiêu đề + 4 block kết quả</p>
          </div>
          <div id="${pillar.code}-okr-cards" class="cards-grid annual-okr-grid"></div>
        </article>

        <article class="panel glass">
          <div class="panel-head">
            <h3 id="${pillar.code}-month-title">KPI tháng ${uiState.month}</h3>
            <p>KPI theo OKR tháng (target + actual)</p>
          </div>
          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Tháng</th>
                  <th>KPI</th>
                  <th>Target</th>
                  <th>Actual</th>
                  <th>Tỷ lệ đạt</th>
                  <th>Kết quả</th>
                </tr>
              </thead>
              <tbody id="${pillar.code}-month-table"></tbody>
            </table>
          </div>
        </article>

        <article class="panel glass">
          <div class="panel-head">
            <h3>Kế hoạch tháng (${pillar.label})</h3>
            <p>Kế hoạch tháng theo OKR tháng tại kỳ đang chọn</p>
          </div>
          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Tháng</th>
                  <th>OKR liên kết</th>
                  <th>Action</th>
                  <th>Owner</th>
                  <th>Due</th>
                  <th>Status</th>
                  <th>Kết quả</th>
                </tr>
              </thead>
              <tbody id="${pillar.code}-monthly-plan-table"></tbody>
            </table>
          </div>
        </article>

        <article class="panel glass">
          <div class="panel-head">
            <h3>Nhận định ${pillar.label}</h3>
            <p>Nhận định, nhận xét, phân tích và hành động tháng/quý</p>
          </div>
          <ul id="${pillar.code}-insight" class="insight-list"></ul>
        </article>
      `;
    });
  }

  function initSelectors() {
    populateMonthSelect("month-select", { includeAll: true });
    populateMonthSelect("ip-month");
    populateMonthSelect("monthly-okr-month");
    populateMonthSelect("monthly-plan-month");
    populateMonthSelect("monthly-kpi-month");
    populateMonthSelect("import-care-month");

    populateQuarterSelect("quarter-select");
    populateQuarterSelect("quarterly-okr-quarter");

    populatePillarSelect("annual-okr-pillar");
    populatePillarSelect("quarterly-okr-pillar");
    populatePillarSelect("monthly-okr-pillar");
    populatePillarSelect("monthly-plan-pillar");
    populatePillarSelect("monthly-kpi-pillar");

    document.getElementById("month-select").value = uiState.month;
    document.getElementById("quarter-select").value = uiState.quarter;
    document.getElementById("ip-month").value = uiState.month === ALL_MONTH_VALUE ? guessDefaultMonth() : uiState.month;

    const formDefaultMonth = uiState.month === ALL_MONTH_VALUE ? guessDefaultMonth() : uiState.month;
    ["monthly-okr-month", "monthly-plan-month", "monthly-kpi-month"].forEach((id) => {
      document.getElementById(id).value = formDefaultMonth;
    });
    const importCareMonth = document.getElementById("import-care-month");
    if (importCareMonth) {
      importCareMonth.value = formDefaultMonth;
      importCareDraft.month = formDefaultMonth;
      importCareDraft.rows = getPartnerCarePlanRowsByMonth(formDefaultMonth).map((row) => ({ ...row }));
    }

    [
      "annual-okr-pillar",
      "quarterly-okr-pillar",
      "monthly-okr-pillar",
      "monthly-plan-pillar",
      "monthly-kpi-pillar",
    ].forEach((id) => {
      document.getElementById(id).value = "HR";
    });

    document.getElementById("quarterly-okr-quarter").value = uiState.quarter;

    refreshQuarterlyAnnualLinkOptions();
    refreshMonthlyQuarterlyLinkOptions();
    refreshMonthlyPlanLinkOptions();
    refreshMonthlyKpiLinkOptions();
    renderImportCareSetup();
  }

  function initNav() {
    document.getElementById("page-nav").addEventListener("click", (event) => {
      const button = event.target.closest(".nav-item");
      if (!button) return;
      const pageKey = button.dataset.page;
      if (!PAGE_META[pageKey]) return;
      uiState.currentPage = pageKey;
      saveUiState();
      renderPageVisibility();
    });
  }

  function initGlobalHandlers() {
    document.getElementById("month-select").addEventListener("change", (event) => {
      uiState.month = event.target.value;
      if (uiState.month !== ALL_MONTH_VALUE) {
        uiState.quarter = quarterOf(uiState.month);
        document.getElementById("quarter-select").value = uiState.quarter;
        document.getElementById("ip-month").value = uiState.month;
      }
      saveUiState();
      renderAll();
    });

    document.getElementById("quarter-select").addEventListener("change", (event) => {
      uiState.quarter = event.target.value;
      saveUiState();
      renderAll();
    });

    document.getElementById("save-local-btn").addEventListener("click", () => {
      persistState("Đã lưu local");
    });
  }

  function initPillarKpiHandlers() {
    ["HR", "VN", "INTL"].forEach((pillarKey) => {
      const table = document.getElementById(`${pillarKey}-month-table`);
      if (!table) return;
      table.addEventListener("change", (event) => {
        const rowEl = event.target.closest("tr[data-kpi-id]");
        if (!rowEl) return;
        const rowId = cleanText(rowEl.getAttribute("data-kpi-id"));
        if (!rowId) return;
        const row = state.importHub.monthlyKpis.find((item) => cleanText(item.id) === rowId);
        if (!row) return;

        if (event.target.matches("[data-kpi-actual]")) {
          row.actual = parseNumeric(event.target.value);
        } else if (event.target.matches("[data-kpi-result]")) {
          row.result = cleanText(event.target.value);
        } else {
          return;
        }

        persistState("Đã cập nhật KPI từ trang Pillar");
        renderAll();
      });

      const planTable = document.getElementById(`${pillarKey}-monthly-plan-table`);
      if (!planTable) return;
      planTable.addEventListener("change", (event) => {
        const rowEl = event.target.closest("tr[data-plan-id]");
        if (!rowEl) return;
        const rowId = cleanText(rowEl.getAttribute("data-plan-id"));
        if (!rowId) return;
        const row = state.importHub.monthlyPlans.find((item) => cleanText(item.id) === rowId);
        if (!row) return;

        if (event.target.matches("[data-plan-status]")) {
          row.status = cleanText(event.target.value);
        } else if (event.target.matches("[data-plan-result]")) {
          row.result = cleanText(event.target.value);
        } else {
          return;
        }

        persistState("Đã cập nhật kế hoạch tháng từ trang Pillar");
        renderAll();
      });
    });
  }

  function initImportHandlers() {
    document.getElementById("annual-okr-form").addEventListener("submit", (event) => {
      event.preventDefault();
      const row = {
        id: uid("A"),
        pillarKey: normalizePillar(document.getElementById("annual-okr-pillar").value),
        code: cleanText(document.getElementById("annual-okr-code").value),
        title: cleanText(document.getElementById("annual-okr-title").value),
        target: parseNumeric(document.getElementById("annual-okr-target").value),
        unit: cleanText(document.getElementById("annual-okr-unit").value) || "%",
        notes: cleanText(document.getElementById("annual-okr-notes").value),
      };
      if (!row.title || row.target === null) {
        setSyncStatus("Thiếu thông tin OKR năm");
        return;
      }
      state.importHub.annualOkrs.push(row);
      persistState("Đã thêm OKR năm");
      event.target.reset();
      document.getElementById("annual-okr-pillar").value = "HR";
      renderAll();
    });

    document.getElementById("quarterly-okr-form").addEventListener("submit", (event) => {
      event.preventDefault();
      const row = {
        id: uid("QO"),
        quarter: normalizeQuarter(document.getElementById("quarterly-okr-quarter").value),
        pillarKey: normalizePillar(document.getElementById("quarterly-okr-pillar").value),
        code: cleanText(document.getElementById("quarterly-okr-code").value),
        title: cleanText(document.getElementById("quarterly-okr-title").value),
        target: parseNumeric(document.getElementById("quarterly-okr-target").value),
        unit: cleanText(document.getElementById("quarterly-okr-unit").value) || "%",
        linkedAnnualOkrId: cleanText(document.getElementById("quarterly-okr-link-annual").value),
        notes: cleanText(document.getElementById("quarterly-okr-notes").value),
      };
      if (!row.title || row.target === null || !row.linkedAnnualOkrId) {
        setSyncStatus("OKR quý cần nhập đủ và link OKR năm");
        return;
      }
      state.importHub.quarterlyOkrs.push(row);
      persistState("Đã thêm OKR quý");
      event.target.reset();
      document.getElementById("quarterly-okr-quarter").value = uiState.quarter;
      document.getElementById("quarterly-okr-pillar").value = "HR";
      refreshQuarterlyAnnualLinkOptions();
      refreshMonthlyQuarterlyLinkOptions();
      renderAll();
    });

    document.getElementById("monthly-okr-form").addEventListener("submit", (event) => {
      event.preventDefault();
      const row = {
        id: uid("MO"),
        month: normalizeMonth(document.getElementById("monthly-okr-month").value),
        pillarKey: normalizePillar(document.getElementById("monthly-okr-pillar").value),
        linkedQuarterlyOkrId: cleanText(document.getElementById("monthly-okr-link-quarterly").value),
        title: cleanText(document.getElementById("monthly-okr-title").value),
        target: parseNumeric(document.getElementById("monthly-okr-target").value),
        unit: cleanText(document.getElementById("monthly-okr-unit").value) || "%",
        notes: cleanText(document.getElementById("monthly-okr-notes").value),
      };
      if (!row.title || row.target === null || !row.linkedQuarterlyOkrId) {
        setSyncStatus("OKR tháng cần nhập đủ và link OKR quý");
        return;
      }
      state.importHub.monthlyOkrs.push(row);
      persistState("Đã thêm OKR tháng");
      event.target.reset();
      document.getElementById("monthly-okr-month").value = uiState.month;
      document.getElementById("monthly-okr-pillar").value = "HR";
      refreshMonthlyQuarterlyLinkOptions();
      refreshMonthlyPlanLinkOptions();
      refreshMonthlyKpiLinkOptions();
      renderAll();
    });

    document.getElementById("monthly-plan-form").addEventListener("submit", (event) => {
      event.preventDefault();
      const row = {
        id: uid("MP"),
        month: normalizeMonth(document.getElementById("monthly-plan-month").value),
        pillarKey: normalizePillar(document.getElementById("monthly-plan-pillar").value),
        linkedMonthlyOkrId: cleanText(document.getElementById("monthly-plan-link-okr").value),
        action: cleanText(document.getElementById("monthly-plan-action").value),
        owner: cleanText(document.getElementById("monthly-plan-owner").value),
        due: cleanText(document.getElementById("monthly-plan-due").value),
        status: cleanText(document.getElementById("monthly-plan-status").value),
        result: cleanText(document.getElementById("monthly-plan-result").value),
        notes: "",
      };
      if (!row.action) {
        setSyncStatus("Thiếu kế hoạch tháng");
        return;
      }
      state.importHub.monthlyPlans.push(row);
      persistState("Đã thêm kế hoạch tháng");
      event.target.reset();
      document.getElementById("monthly-plan-month").value = uiState.month;
      document.getElementById("monthly-plan-pillar").value = "HR";
      refreshMonthlyPlanLinkOptions();
      renderAll();
    });

    document.getElementById("monthly-kpi-form").addEventListener("submit", (event) => {
      event.preventDefault();
      const row = {
        id: uid("MK"),
        month: normalizeMonth(document.getElementById("monthly-kpi-month").value),
        pillarKey: normalizePillar(document.getElementById("monthly-kpi-pillar").value),
        monthlyOkrId: cleanText(document.getElementById("monthly-kpi-link-month-okr").value),
        kpiName: cleanText(document.getElementById("monthly-kpi-name").value),
        target: parseNumeric(document.getElementById("monthly-kpi-target").value),
        actual: null,
        unit: cleanText(document.getElementById("monthly-kpi-unit").value) || "%",
        status: "",
        result: "",
      };
      if (!row.kpiName || row.target === null) {
        setSyncStatus("Thiếu thông tin KPI tháng");
        return;
      }
      state.importHub.monthlyKpis.push(row);
      persistState("Đã thêm KPI tháng");
      event.target.reset();
      document.getElementById("monthly-kpi-month").value = uiState.month;
      document.getElementById("monthly-kpi-pillar").value = "HR";
      refreshMonthlyKpiLinkOptions();
      renderAll();
    });

    document.getElementById("quarterly-okr-pillar").addEventListener("change", refreshQuarterlyAnnualLinkOptions);
    document.getElementById("monthly-okr-month").addEventListener("change", refreshMonthlyQuarterlyLinkOptions);
    document.getElementById("monthly-okr-pillar").addEventListener("change", refreshMonthlyQuarterlyLinkOptions);
    document.getElementById("monthly-plan-month").addEventListener("change", refreshMonthlyPlanLinkOptions);
    document.getElementById("monthly-plan-pillar").addEventListener("change", refreshMonthlyPlanLinkOptions);
    document.getElementById("monthly-kpi-month").addEventListener("change", refreshMonthlyKpiLinkOptions);
    document.getElementById("monthly-kpi-pillar").addEventListener("change", refreshMonthlyKpiLinkOptions);

    const careMonthSelect = document.getElementById("import-care-month");
    if (careMonthSelect) {
      careMonthSelect.addEventListener("change", (event) => {
        const month = normalizeMonth(event.target.value);
        importCareDraft.month = month;
        renderImportCareSetup();
      });
    }

    const carePartnerSelect = document.getElementById("import-care-partner-key");
    if (carePartnerSelect) {
      carePartnerSelect.addEventListener("change", syncImportCareFormFromPartnerSelection);
    }

    const careForm = document.getElementById("import-care-form");
    if (careForm) {
      careForm.addEventListener("submit", (event) => {
        event.preventDefault();
        const month = normalizeMonth(document.getElementById("import-care-month").value);
        const partnerKey = cleanText(document.getElementById("import-care-partner-key").value).toUpperCase();
        const careDay = normalizeCareDay(document.getElementById("import-care-day").value);
        if (!partnerKey) {
          setSyncStatus("Chọn mã đối tác trước khi thêm danh sách chăm sóc");
          return;
        }
        if (careDay === null) {
          setSyncStatus("Ngày chăm sóc cần từ 1 đến 31");
          return;
        }
        const partner = getPartnerMasterByKey(partnerKey);
        if (!partner) {
          setSyncStatus("Mã đối tác chưa có trong CRM master. Vui lòng sync CRM_IMPORT trước.");
          return;
        }
        state.partnerCare.partnerMonthlyPlans.push({
          id: uid("PCP"),
          month,
          careDay,
          dayDone: null,
          partnerKey,
          partnerName: cleanText(partner.partnerName),
          partnerType: normalizePartnerType(partner.partnerType),
          owner: cleanText(document.getElementById("import-care-owner").value) || cleanText(partner.owner),
          touchpointType: cleanText(document.getElementById("import-care-touchpoint").value),
          channel: cleanText(document.getElementById("import-care-channel").value) || "Call/message",
          status: "Planned",
          result: "",
          response: "",
          targetTouchpoints: 1,
          nextAction: "",
        });
        importCareDraft.month = month;
        loadImportCareDraftMonth(month);
        document.getElementById("import-care-day").value = "";
        document.getElementById("import-care-touchpoint").value = "";
        document.getElementById("import-care-channel").value = "Call/message";
        uiState.month = month;
        uiState.quarter = quarterOf(month);
        saveUiState();
        persistState(`Đã thêm 1 đối tác vào danh sách chăm sóc tháng ${month}`);
        renderAll();
      });
    }

    const careDraftTable = document.getElementById("import-care-draft-table");
    if (careDraftTable) {
      careDraftTable.addEventListener("click", (event) => {
        const button = event.target.closest("button[data-import-care-remove]");
        if (!button) return;
        const rowId = cleanText(button.getAttribute("data-import-care-remove"));
        state.partnerCare.partnerMonthlyPlans = state.partnerCare.partnerMonthlyPlans.filter(
          (row) => cleanText(row.id) !== rowId
        );
        persistState("Đã xóa dòng khỏi danh sách chăm sóc tháng");
        renderAll();
      });
    }

    wireRemoveHandlers();
  }

  function initCloudSyncHandlers() {
    const urlInput = document.getElementById("sb-url");
    const anonInput = document.getElementById("sb-anon-key");
    const tableInput = document.getElementById("sb-table");
    const keyInput = document.getElementById("sb-state-key");
    const autoSyncInput = document.getElementById("sb-auto-sync");

    if (urlInput) {
      urlInput.addEventListener("change", () => {
        cloudConfig.url = cleanText(urlInput.value);
        invalidateSupabaseClient();
        saveCloudConfig("Đã lưu Supabase URL");
        renderCloudSyncPanel();
      });
    }

    if (anonInput) {
      anonInput.addEventListener("change", () => {
        cloudConfig.anonKey = cleanText(anonInput.value);
        invalidateSupabaseClient();
        saveCloudConfig("Đã lưu Supabase Anon Key");
        renderCloudSyncPanel();
      });
    }

    if (tableInput) {
      tableInput.addEventListener("change", () => {
        cloudConfig.table = cleanText(tableInput.value) || "dashboard_state";
        saveCloudConfig("Đã lưu tên bảng Supabase");
        renderCloudSyncPanel();
      });
    }

    if (keyInput) {
      keyInput.addEventListener("change", () => {
        cloudConfig.stateKey = cleanText(keyInput.value) || "global";
        saveCloudConfig("Đã lưu state key");
        renderCloudSyncPanel();
      });
    }

    if (autoSyncInput) {
      autoSyncInput.addEventListener("change", () => {
        cloudConfig.autoSync = !!autoSyncInput.checked;
        saveCloudConfig(cloudConfig.autoSync ? "Đã bật auto sync Supabase" : "Đã tắt auto sync Supabase");
        renderCloudSyncPanel();
      });
    }

    const saveBtn = document.getElementById("sb-save-config-btn");
    if (saveBtn) {
      saveBtn.addEventListener("click", () => {
        cloudConfig.url = cleanText(urlInput?.value);
        cloudConfig.anonKey = cleanText(anonInput?.value);
        cloudConfig.table = cleanText(tableInput?.value) || "dashboard_state";
        cloudConfig.stateKey = cleanText(keyInput?.value) || "global";
        cloudConfig.autoSync = !!autoSyncInput?.checked;
        invalidateSupabaseClient();
        saveCloudConfig("Đã lưu cấu hình Supabase");
        renderCloudSyncPanel();
      });
    }

    const testBtn = document.getElementById("sb-test-btn");
    if (testBtn) {
      testBtn.addEventListener("click", async () => {
        await testSupabaseConnection();
      });
    }

    const pullBtn = document.getElementById("sb-pull-btn");
    if (pullBtn) {
      pullBtn.addEventListener("click", async () => {
        await pullStateFromCloud();
      });
    }

    const pushBtn = document.getElementById("sb-push-btn");
    if (pushBtn) {
      pushBtn.addEventListener("click", async () => {
        await pushStateToCloud("manual");
      });
    }
  }

  function wireRemoveHandlers() {
    document.getElementById("annual-okr-table").addEventListener("click", (event) => {
      const button = event.target.closest("button[data-remove-annual]");
      if (!button) return;
      removeAnnualOkr(button.getAttribute("data-remove-annual"));
      persistState("Đã xóa OKR năm");
      renderAll();
    });

    document.getElementById("quarterly-okr-table").addEventListener("click", (event) => {
      const button = event.target.closest("button[data-remove-quarterly-okr]");
      if (!button) return;
      removeQuarterlyOkr(button.getAttribute("data-remove-quarterly-okr"));
      persistState("Đã xóa OKR quý");
      renderAll();
    });

    document.getElementById("monthly-okr-table").addEventListener("click", (event) => {
      const button = event.target.closest("button[data-remove-monthly-okr]");
      if (!button) return;
      removeMonthlyOkr(button.getAttribute("data-remove-monthly-okr"));
      persistState("Đã xóa OKR tháng");
      renderAll();
    });

    document.getElementById("monthly-plan-table").addEventListener("click", (event) => {
      const button = event.target.closest("button[data-remove-monthly-plan]");
      if (!button) return;
      state.importHub.monthlyPlans = state.importHub.monthlyPlans.filter(
        (row) => row.id !== button.getAttribute("data-remove-monthly-plan")
      );
      persistState("Đã xóa kế hoạch tháng");
      renderAll();
    });

    document.getElementById("monthly-kpi-table").addEventListener("click", (event) => {
      const button = event.target.closest("button[data-remove-monthly-kpi]");
      if (!button) return;
      state.importHub.monthlyKpis = state.importHub.monthlyKpis.filter(
        (row) => row.id !== button.getAttribute("data-remove-monthly-kpi")
      );
      persistState("Đã xóa KPI tháng");
      renderAll();
    });
  }

  function refreshQuarterlyAnnualLinkOptions() {
    const pillarKey = normalizePillar(document.getElementById("quarterly-okr-pillar").value);
    const select = document.getElementById("quarterly-okr-link-annual");
    const rows = state.importHub.annualOkrs.filter((row) => row.pillarKey === pillarKey);
    select.innerHTML = "";
    select.appendChild(makeOption("", "-- Không link --"));
    rows.forEach((row) => {
      select.appendChild(makeOption(row.id, `${row.code || row.id} | ${row.title}`));
    });
  }

  function refreshMonthlyQuarterlyLinkOptions() {
    const month = normalizeMonth(document.getElementById("monthly-okr-month").value);
    const quarter = quarterOf(month);
    const pillarKey = normalizePillar(document.getElementById("monthly-okr-pillar").value);
    const select = document.getElementById("monthly-okr-link-quarterly");
    const rows = state.importHub.quarterlyOkrs.filter(
      (row) => row.quarter === quarter && row.pillarKey === pillarKey
    );
    select.innerHTML = "";
    select.appendChild(makeOption("", "-- Không link --"));
    rows.forEach((row) => {
      select.appendChild(makeOption(row.id, `${row.code || row.id} | ${row.title}`));
    });
  }

  function refreshMonthlyPlanLinkOptions() {
    const month = normalizeMonth(document.getElementById("monthly-plan-month").value);
    const pillarKey = normalizePillar(document.getElementById("monthly-plan-pillar").value);
    const select = document.getElementById("monthly-plan-link-okr");
    const rows = state.importHub.monthlyOkrs.filter((row) => row.month === month && row.pillarKey === pillarKey);
    select.innerHTML = "";
    select.appendChild(makeOption("", "-- Không link --"));
    rows.forEach((row) => {
      select.appendChild(makeOption(row.id, `${row.title} (${formatMetric(row.target, row.unit)})`));
    });
  }

  function removeAnnualOkr(id) {
    state.importHub.annualOkrs = state.importHub.annualOkrs.filter((row) => row.id !== id);
    const quarterlyToRemove = state.importHub.quarterlyOkrs
      .filter((row) => row.linkedAnnualOkrId === id)
      .map((row) => row.id);
    quarterlyToRemove.forEach((quarterlyId) => removeQuarterlyOkr(quarterlyId));
  }

  function removeQuarterlyOkr(id) {
    state.importHub.quarterlyOkrs = state.importHub.quarterlyOkrs.filter((row) => row.id !== id);
    const monthlyToRemove = state.importHub.monthlyOkrs
      .filter((row) => row.linkedQuarterlyOkrId === id)
      .map((row) => row.id);
    monthlyToRemove.forEach((monthlyId) => removeMonthlyOkr(monthlyId));
  }

  function removeMonthlyOkr(id) {
    state.importHub.monthlyOkrs = state.importHub.monthlyOkrs.filter((row) => row.id !== id);
    state.importHub.monthlyPlans = state.importHub.monthlyPlans.filter((row) => row.linkedMonthlyOkrId !== id);
    state.importHub.monthlyKpis = state.importHub.monthlyKpis.filter((row) => row.monthlyOkrId !== id);
  }

  function refreshMonthlyKpiLinkOptions() {
    const month = normalizeMonth(document.getElementById("monthly-kpi-month").value);
    const pillarKey = normalizePillar(document.getElementById("monthly-kpi-pillar").value);
    const select = document.getElementById("monthly-kpi-link-month-okr");
    const rows = state.importHub.monthlyOkrs.filter((row) => row.month === month && row.pillarKey === pillarKey);
    select.innerHTML = "";
    select.appendChild(makeOption("", "-- Không link --"));
    rows.forEach((row) => {
      select.appendChild(makeOption(row.id, `${row.title} (${formatMetric(row.target, row.unit)})`));
    });
  }

  function initPartnerInfoHandlers() {
    const endpointInput = document.getElementById("pc-sync-endpoint");
    if (endpointInput) {
      endpointInput.addEventListener("change", (event) => {
        state.partnerCare.sync.endpoint = cleanText(event.target.value);
        persistState("Đã lưu endpoint Google Sheets");
        renderAll();
      });
    }

    const syncButton = document.getElementById("pc-sync-btn");
    if (syncButton) {
      syncButton.addEventListener("click", async () => {
        await syncPartnerMasterFromGoogleSheets();
      });
    }
  }

  function initPartnerCareHandlers() {
    const typeFilter = document.getElementById("pc-filter-type");
    if (typeFilter) typeFilter.addEventListener("change", renderPartnerCare);

    const statusFilter = document.getElementById("pc-filter-status");
    if (statusFilter) statusFilter.addEventListener("change", renderPartnerCare);

    const careRosterTable = document.getElementById("pc-care-roster-table");
    if (careRosterTable) {
      careRosterTable.addEventListener("change", (event) => {
        const rowEl = event.target.closest("tr[data-pc-care-id]");
        if (!rowEl) return;
        const rowId = cleanText(rowEl.getAttribute("data-pc-care-id"));
        if (!rowId) return;
        const row = (state.partnerCare.partnerMonthlyPlans || []).find((item) => cleanText(item.id) === rowId);
        if (!row) return;

        if (event.target.matches("[data-pc-care-status]")) {
          row.status = cleanText(event.target.value) || "Planned";
        } else if (event.target.matches("[data-pc-care-day-done]")) {
          row.dayDone = normalizeCareDay(event.target.value);
        } else if (event.target.matches("[data-pc-care-result]")) {
          row.result = cleanText(event.target.value);
        } else if (event.target.matches("[data-pc-care-response]")) {
          row.response = cleanText(event.target.value);
        } else {
          return;
        }

        persistState("Đã cập nhật actual chăm sóc đối tác");
        renderPartnerCare();
      });

      careRosterTable.addEventListener("click", (event) => {
        const button = event.target.closest("button[data-pc-care-commit]");
        if (!button) return;
        const rowEl = button.closest("tr[data-pc-care-id]");
        if (!rowEl) return;
        const rowId = cleanText(rowEl.getAttribute("data-pc-care-id"));
        if (!rowId) return;
        const row = (state.partnerCare.partnerMonthlyPlans || []).find((item) => cleanText(item.id) === rowId);
        if (!row) return;

        const errors = [];
        if (!cleanText(row.status)) errors.push("Status");
        if (normalizeCareDay(row.dayDone) === null) errors.push("Day Done");
        if (!cleanText(row.result)) errors.push("Result");
        if (!cleanText(row.response)) errors.push("Response");
        if (errors.length) {
          setSyncStatus(`Thiếu dữ liệu trước khi ghi nhận: ${errors.join(", ")}`);
          return;
        }

        upsertPartnerCareLogFromPlan(row);
        persistState(`Đã ghi nhận log chăm sóc: ${cleanText(row.partnerKey)} (${cleanText(row.month)})`);
        renderPartnerCare();
      });
    }
  }

  function initIntlProblemHandlers() {
    document.getElementById("ip-form").addEventListener("submit", (event) => {
      event.preventDefault();
      const payload = {
        date: parseNumeric(document.getElementById("ip-date").value),
        month: document.getElementById("ip-month").value,
        case: cleanText(document.getElementById("ip-case").value),
        result: cleanText(document.getElementById("ip-result").value),
        status: document.getElementById("ip-status").value,
        category: cleanText(document.getElementById("ip-category").value),
        rootCause: cleanText(document.getElementById("ip-root-cause").value),
        slaHours: parseNumeric(document.getElementById("ip-sla").value),
        resolvedHours: parseNumeric(document.getElementById("ip-resolved").value),
        owner: "",
        nextAction: "",
      };
      if (!payload.month || !payload.case) {
        setSyncStatus("Cần nhập ít nhất Month + Case");
        return;
      }
      state.intlProblems.push(payload);
      persistState("Đã thêm case International");
      renderIntlProblem();
      event.target.reset();
      document.getElementById("ip-month").value = uiState.month;
    });
  }

  function renderAll() {
    renderPageVisibility();
    renderDashboard();
    renderQuarterPlan();
    renderPillar("HR");
    renderPillar("VN");
    renderPillar("INTL");
    renderImportCenter();
    renderPartnerInfo();
    renderPartnerCare();
    renderIntlProblem();

    document.getElementById("month-select").value = uiState.month;
    document.getElementById("quarter-select").value = uiState.quarter;
    setSyncStatus("Dữ liệu đã sẵn sàng");
  }

  function renderPageVisibility() {
    document.querySelectorAll(".nav-item").forEach((item) => {
      item.classList.toggle("active", item.dataset.page === uiState.currentPage);
    });

    document.querySelectorAll(".page").forEach((page) => page.classList.remove("active"));

    const pageMap = {
      dashboard: "page-dashboard",
      "quarter-plan": "page-quarter-plan",
      "pillar-hr": "page-pillar-hr",
      "pillar-vn": "page-pillar-vn",
      "pillar-intl": "page-pillar-intl",
      "import-center": "page-import-center",
      "partner-info": "page-partner-info",
      "partner-care": "page-partner-care",
      "intl-problem": "page-intl-problem",
    };

    const pageId = pageMap[uiState.currentPage] || "page-dashboard";
    const pageEl = document.getElementById(pageId);
    if (pageEl) pageEl.classList.add("active");

    const meta = PAGE_META[uiState.currentPage] || PAGE_META.dashboard;
    document.getElementById("page-title").textContent = meta.title;
    document.getElementById("page-subtitle").textContent = meta.subtitle;
  }

  function renderDashboard() {
    const ytdMonths = getMonthsUpToUiSelection(uiState.month);
    const scopeMonths = getScopeMonthsByUiSelection(uiState.month);

    renderDashboardAnnualOkrCards(ytdMonths, getUiMonthLabel(uiState.month));

    const monthBars = document.getElementById("dash-month-bars");
    monthBars.innerHTML = "";
    MONTHS.forEach((month) => {
      const agg = aggregateByMonths([month]);
      const row = document.createElement("div");
      row.className = "bar-item";
      row.innerHTML = `
        <span class="month">${month}</span>
        <div class="bar-shell"><div class="bar-fill" style="width:${Math.min(100, (agg.rate || 0) * 100)}%"></div></div>
        <span class="bar-text">${formatPercent(agg.rate)}</span>
      `;
      monthBars.appendChild(row);
    });

    const breakdown = document.getElementById("dash-pillar-breakdown");
    breakdown.innerHTML = "";
    const pillarMonthStats =
      uiState.month === ALL_MONTH_VALUE ? getPillarRatesByScope("months", scopeMonths) : getPillarRatesByScope("month", uiState.month);
    Object.keys(PILLARS).forEach((pillarKey) => {
      const pillar = PILLARS[pillarKey];
      const rate = pillarMonthStats[pillarKey] || 0;
      const contribution = rate * (state.meta.weights?.[pillar.weightKey] || 0);
      const item = document.createElement("div");
      item.className = "pillar-line";
      item.innerHTML = `
        <p>${pillar.label}</p>
        <div class="mini-bar"><span style="width:${Math.min(100, rate * 100)}%;background:${pillar.color};"></span></div>
        <p>${formatPercent(rate)} | ${formatPercent(contribution)}</p>
      `;
      breakdown.appendChild(item);
    });
  }

  function renderDashboardAnnualOkrCards(ytdMonths, uptoMonth) {
    const container = document.getElementById("dash-annual-okr-cards");
    if (!container) return;

    const monthsSet = new Set(ytdMonths);
    container.innerHTML = "";

    if (!state.importHub.annualOkrs.length) {
      container.innerHTML = `
        <article class="metric-card glass annual-okr-card">
          <div class="annual-okr-main">
            <p class="metric-label">OKR Năm 2026</p>
            <h3>Chưa có dữ liệu</h3>
          </div>
          <div class="annual-okr-stat">
            <small>Target</small>
            <strong>-</strong>
          </div>
          <div class="annual-okr-stat">
            <small>Actual lũy kế</small>
            <strong>-</strong>
          </div>
          <div class="annual-okr-stat">
            <small>Tỷ lệ đạt</small>
            <strong>-</strong>
          </div>
        </article>
      `;
      return;
    }

    state.importHub.annualOkrs.forEach((annual) => {
      const target = asNumber(annual.target);
      const linkedKpis = getKpisForAnnualOkr(annual, Array.from(monthsSet));

      let actual = 0;
      let actualCount = 0;
      linkedKpis.forEach((kpi) => {
        const value = asNumber(kpi.actual);
        if (value === null) return;
        actual += value;
        actualCount += 1;
      });

      const hasActualData = actualCount > 0;
      const rate = target && target > 0 ? actual / target : null;
      const pillar = PILLARS[annual.pillarKey];
      const targetText = formatMetric(target, annual.unit);
      const actualText = hasActualData ? formatMetric(actual, annual.unit) : "-";
      const rateText = rate === null ? "-" : formatPercent(rate);

      const card = document.createElement("article");
      card.className = "metric-card glass annual-okr-card";
      card.innerHTML = `
        <div class="annual-okr-main">
          <p class="metric-label">${safeText(pillar?.label || annual.pillarKey || "Pillar")} ${
            annual.code ? `• ${safeText(annual.code)}` : ""
          }</p>
          <h3>${safeText(annual.title || annual.code || "OKR năm")}</h3>
        </div>
        <div class="annual-okr-stat">
          <small>Target</small>
          <strong>${targetText}</strong>
        </div>
        <div class="annual-okr-stat">
          <small>Actual lũy kế (${safeText(uptoMonth)})</small>
          <strong>${actualText}</strong>
        </div>
        <div class="annual-okr-stat">
          <small>Tỷ lệ đạt</small>
          <strong>${rateText}</strong>
        </div>
      `;
      container.appendChild(card);
    });
  }

  function renderQuarterPlan() {
    const quarter = uiState.quarter;
    const quarterMonths = QUARTER_MONTHS[quarter] || [];
    const quarterlyOkrs = state.importHub.quarterlyOkrs.filter((row) => row.quarter === quarter);
    const getActualByQuarterlyOkr = (quarterlyOkrId) => {
      const quarterlyOkr = state.importHub.quarterlyOkrs.find((row) => row.id === quarterlyOkrId);
      if (!quarterlyOkr) return 0;
      return getKpisForQuarterlyOkr(quarterlyOkr, quarterMonths).reduce((sum, row) => sum + (asNumber(row.actual) || 0), 0);
    };

    const rows = quarterlyOkrs.map((okr) => {
      const target = asNumber(okr.target) || 0;
      const actual = getActualByQuarterlyOkr(okr.id);
      const rate = target > 0 ? actual / target : null;
      return {
        quarter,
        pillarLabel: PILLARS[okr.pillarKey]?.label || okr.pillarKey,
        title: cleanText(okr.title),
        target,
        actual,
        rate,
        unit: okr.unit || "",
        notes: cleanText(okr.notes),
      };
    });

    const cards = document.getElementById("qplan-okr-cards");
    cards.innerHTML = "";
    if (!rows.length) {
      cards.innerHTML = `
        <article class="metric-card glass annual-okr-card">
          <div class="annual-okr-main">
            <p class="metric-label">${safeText(quarter)}</p>
            <h3>Chưa có OKR quý</h3>
          </div>
          <div class="annual-okr-stat">
            <small>Target</small>
            <strong>-</strong>
          </div>
          <div class="annual-okr-stat">
            <small>Actual lũy kế quý</small>
            <strong>-</strong>
          </div>
          <div class="annual-okr-stat">
            <small>Tỷ lệ đạt</small>
            <strong>-</strong>
          </div>
        </article>
      `;
    } else {
      rows.forEach((row) => {
        const card = document.createElement("article");
        card.className = "metric-card glass annual-okr-card";
        card.innerHTML = `
          <div class="annual-okr-main">
            <p class="metric-label">${safeText(row.quarter)} • ${safeText(row.pillarLabel)}</p>
            <h3>${safeText(row.title || "OKR quý")}</h3>
          </div>
          <div class="annual-okr-stat">
            <small>Target</small>
            <strong>${formatMetric(row.target, row.unit)}</strong>
          </div>
          <div class="annual-okr-stat">
            <small>Actual lũy kế quý</small>
            <strong>${formatMetric(row.actual, row.unit)}</strong>
          </div>
          <div class="annual-okr-stat">
            <small>Tỷ lệ đạt</small>
            <strong>${row.rate === null ? "-" : formatPercent(row.rate)}</strong>
          </div>
        `;
        cards.appendChild(card);
      });
    }

    const table = document.getElementById("quarter-plan-table");
    table.innerHTML = "";
    if (!rows.length) {
      table.innerHTML = `<tr><td colspan="7" class="empty-row">Chưa có OKR quý cho ${quarter}</td></tr>`;
    } else {
      rows.forEach((row) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${safeText(row.quarter)}</td>
          <td>${safeText(row.pillarLabel)}</td>
          <td>${safeText(row.title)}</td>
          <td>${formatMetric(row.target, row.unit)}</td>
          <td>${formatMetric(row.actual, row.unit)}</td>
          <td>${row.rate === null ? "-" : formatPercent(row.rate)}</td>
          <td>${safeText(row.notes)}</td>
        `;
        table.appendChild(tr);
      });
    }

    const targetQuarter = rows.reduce((sum, row) => sum + row.target, 0);
    const actualQuarter = rows.reduce((sum, row) => sum + row.actual, 0);
    const rateQuarter = targetQuarter > 0 ? actualQuarter / targetQuarter : 0;

    const unlinkedAnnual = quarterlyOkrs.filter((row) => !cleanText(row.linkedAnnualOkrId)).length;
    const insight = [
      `Nhận định: ${quarter} có ${quarterlyOkrs.length} OKR quý, tỷ lệ đạt hiện tại ${formatPercent(rateQuarter)}.`,
      `Nhận xét: Có ${unlinkedAnnual} OKR quý chưa link OKR năm.`,
      "Phân tích: Mức độ đạt quý đang đo bằng actual KPI của các OKR tháng link về OKR quý.",
      "Đề xuất hành động: Chuẩn hóa link Năm -> Quý -> Tháng đầy đủ để đảm bảo đo lường nhất quán.",
    ];
    renderInsightList("qplan-insight", insight);
  }

  function renderPillar(pillarKey) {
    const pillar = PILLARS[pillarKey];
    if (!pillar) return;

    const month = uiState.month;
    const scopeMonths = getScopeMonthsByUiSelection(month);
    const scopeLabel = month === ALL_MONTH_VALUE ? "Tất cả (Jan-Dec)" : month;
    const monthKpis = getMonthlyKpis(pillarKey, month);
    const forecastRate = forecastByPillar(pillarKey, month);

    const annualOkrs = state.importHub.annualOkrs.filter((row) => row.pillarKey === pillarKey);
    const okrCards = document.getElementById(`${pillarKey}-okr-cards`);
    okrCards.innerHTML = "";
    if (!annualOkrs.length) {
      okrCards.innerHTML = `
        <article class="panel glass">
          <div class="panel-head">
            <h3>${safeText(pillar.label)}</h3>
            <p>Chưa có OKR năm cho pillar này</p>
          </div>
        </article>
      `;
    } else {
      const forecastMonths = getMonthsUpToUiSelection(month);
      annualOkrs.forEach((annual) => {
        const kpisInScope = getKpisForAnnualOkr(annual, scopeMonths);
        const monthlyOkrsInScope = getMonthlyOkrsForAnnualOkr(annual, scopeMonths);
        const monthlyOkrTargets = monthlyOkrsInScope.map((row) => asNumber(row.target)).filter((val) => val !== null);
        let targetAvg = monthlyOkrTargets.length
          ? monthlyOkrTargets.reduce((sum, val) => sum + val, 0) / monthlyOkrTargets.length
          : null;
        if (targetAvg === null) {
          const targetValues = kpisInScope.map((row) => asNumber(row.target)).filter((val) => val !== null);
          targetAvg = targetValues.length ? targetValues.reduce((sum, val) => sum + val, 0) / targetValues.length : null;
        }
        const kpiRateValues = kpisInScope
          .map((row) => {
            const target = asNumber(row.target);
            const actual = asNumber(row.actual);
            if (target === null || target <= 0 || actual === null) return null;
            return actual / target;
          })
          .filter((val) => val !== null);
        const actualAvgRate = kpiRateValues.length
          ? kpiRateValues.reduce((sum, val) => sum + val, 0) / kpiRateValues.length
          : null;

        const totalTarget = kpisInScope.reduce((sum, row) => sum + Math.max(0, asNumber(row.target) || 0), 0);
        const totalActual = kpisInScope.reduce((sum, row) => sum + Math.max(0, asNumber(row.actual) || 0), 0);
        const rate = totalTarget > 0 ? totalActual / totalTarget : null;

        let forecastBase = actualAvgRate;
        if (month !== ALL_MONTH_VALUE) {
          const monthlyRateAverages = forecastMonths
            .map((m) => {
              const monthRateValues = getKpisForAnnualOkr(annual, [m])
                .map((row) => {
                  const target = asNumber(row.target);
                  const actual = asNumber(row.actual);
                  if (target === null || target <= 0 || actual === null) return null;
                  return actual / target;
                })
                .filter((val) => val !== null);
              if (!monthRateValues.length) return null;
              return monthRateValues.reduce((sum, val) => sum + val, 0) / monthRateValues.length;
            })
            .filter((val) => val !== null);
          forecastBase = monthlyRateAverages.length
            ? monthlyRateAverages.reduce((sum, val) => sum + val, 0) / monthlyRateAverages.length
            : null;
        }
        const forecastByOkr = forecastBase;

        const row = document.createElement("article");
        row.className = "panel glass";
        row.innerHTML = `
          <div class="panel-head">
            <h3>${safeText(annual.title || "OKR")}</h3>
            <p>${safeText(pillar.label)} • Đơn vị: ${safeText(annual.unit || "%")}</p>
          </div>
          <div class="cards-grid four">
            <article class="metric-card glass">
              <p class="metric-label">Target tháng (OKR)</p>
              <h3>${targetAvg === null ? "-" : formatMetric(targetAvg, annual.unit)}</h3>
              <p class="metric-note">${safeText(scopeLabel)}</p>
            </article>
            <article class="metric-card glass">
              <p class="metric-label">Actual tháng (KPI)</p>
              <h3>${actualAvgRate === null ? "-" : formatPercent(actualAvgRate)}</h3>
              <p class="metric-note">TB tỷ lệ đạt KPI trong kỳ</p>
            </article>
            <article class="metric-card glass">
              <p class="metric-label">Tỷ lệ đạt tháng</p>
              <h3>${rate === null ? "-" : formatPercent(rate)}</h3>
              <p class="metric-note">Tổng Actual KPI / Tổng Target KPI</p>
            </article>
            <article class="metric-card glass">
              <p class="metric-label">Forecast Cuối Năm</p>
              <h3>${forecastByOkr === null ? "-" : formatPercent(forecastByOkr)}</h3>
              <p class="metric-note">Run-rate theo actual hiện tại</p>
            </article>
          </div>
        `;
        okrCards.appendChild(row);
      });
    }

    setText(`${pillarKey}-month-title`, `KPI ${scopeLabel}`);

    const kpiTable = document.getElementById(`${pillarKey}-month-table`);
    kpiTable.innerHTML = "";
    if (!monthKpis.length) {
      kpiTable.innerHTML = `<tr><td colspan="6" class="empty-row">Chưa có KPI trong kỳ cho ${pillar.label}</td></tr>`;
    } else {
      monthKpis
        .slice()
        .sort((a, b) => {
          const monthDiff = MONTHS.indexOf(cleanText(a.month)) - MONTHS.indexOf(cleanText(b.month));
          if (monthDiff !== 0) return monthDiff;
          return cleanText(a.kpiName).localeCompare(cleanText(b.kpiName));
        })
        .forEach((kpi) => {
        const target = asNumber(kpi.target);
        const actual = asNumber(kpi.actual);
        const ratio = target && target > 0 ? actual / target : null;
        const actualInput = actual === null ? "" : String(actual);
        const resultValue = cleanText(kpi.result || kpi.notes);
        const tr = document.createElement("tr");
        tr.setAttribute("data-kpi-id", cleanText(kpi.id));
        tr.innerHTML = `
          <td>${safeText(kpi.month)}</td>
          <td>${safeText(kpi.kpiName)}</td>
          <td>${formatMetric(target, kpi.unit)}</td>
          <td>
            <input
              class="kpi-cell-input"
              data-kpi-actual
              type="number"
              step="any"
              value="${safeText(actualInput)}"
              placeholder="Nhập actual"
            />
          </td>
          <td>${ratio === null ? "-" : formatPercent(ratio)}</td>
          <td>
            <input
              class="kpi-cell-input"
              data-kpi-result
              type="text"
              value="${safeText(resultValue)}"
              placeholder="Text hoặc https://..."
            />
            ${renderKpiResultPreview(resultValue)}
          </td>
        `;
        kpiTable.appendChild(tr);
        });
    }

    const monthlyPlans = state.importHub.monthlyPlans.filter(
      (row) => row.pillarKey === pillarKey && scopeMonths.includes(row.month)
    );
    const planTable = document.getElementById(`${pillarKey}-monthly-plan-table`);
    planTable.innerHTML = "";
    if (!monthlyPlans.length) {
      planTable.innerHTML = `<tr><td colspan="7" class="empty-row">Chưa có kế hoạch tháng</td></tr>`;
    } else {
      monthlyPlans.forEach((plan) => {
        const linkedOkr = state.importHub.monthlyOkrs.find((row) => row.id === plan.linkedMonthlyOkrId);
        const resultValue = cleanText(plan.result || plan.notes);
        const tr = document.createElement("tr");
        tr.setAttribute("data-plan-id", cleanText(plan.id));
        tr.innerHTML = `
          <td>${safeText(plan.month)}</td>
          <td>${safeText(linkedOkr?.title || "")}</td>
          <td>${safeText(plan.action)}</td>
          <td>${safeText(plan.owner)}</td>
          <td>${safeText(plan.due)}</td>
          <td>
            <input
              class="kpi-cell-input"
              data-plan-status
              type="text"
              value="${safeText(cleanText(plan.status))}"
              placeholder="Planned/In Progress/Done"
            />
          </td>
          <td>
            <input
              class="kpi-cell-input"
              data-plan-result
              type="text"
              value="${safeText(resultValue)}"
              placeholder="Text hoặc https://..."
            />
            ${renderKpiResultPreview(resultValue)}
          </td>
        `;
        planTable.appendChild(tr);
      });
    }

    const periodAgg = aggregatePillarByMonths(pillarKey, scopeMonths);
    const quarterRate = aggregatePillarByMonths(pillarKey, QUARTER_MONTHS[uiState.quarter] || []).rate;
    const insight = [
      `Nhận định: ${pillar.label} kỳ ${scopeLabel} đạt ${formatPercent(periodAgg.rate)} trên dữ liệu KPI trong kỳ.`,
      `Nhận xét: Quý ${uiState.quarter} của pillar đang ở mức ${formatPercent(quarterRate)}, forecast năm ${formatPercent(forecastRate)}.`,
      `Phân tích: Có ${monthKpis.length} KPI và ${monthlyPlans.length} action plan tháng trong kỳ hiện tại.`,
      "Đề xuất hành động: Tập trung xử lý KPI có tỷ lệ đạt thấp, và bám tiến độ action plan theo owner mỗi tuần.",
    ];
    renderInsightList(`${pillarKey}-insight`, insight);
  }

  function renderImportCenter() {
    const rows = state.importHub.monthlyKpis;
    const missingTarget = rows.filter((row) => asNumber(row.target) === null).length;
    const missingActual = rows.filter((row) => asNumber(row.target) !== null && asNumber(row.actual) === null).length;

    setText("import-data-rows", rows.length);
    setText("import-missing-target", missingTarget);
    setText("import-missing-actual", missingActual);
    setText("import-version", state.version || 1);

    renderAnnualOkrTable();
    renderQuarterlyOkrTable();
    renderMonthlyOkrTable();
    renderMonthlyPlanTable();
    renderMonthlyKpiTable();
    renderCloudSyncPanel();

    refreshQuarterlyAnnualLinkOptions();
    refreshMonthlyQuarterlyLinkOptions();
    refreshMonthlyPlanLinkOptions();
    refreshMonthlyKpiLinkOptions();
    renderImportCareSetup();
  }

  function renderCloudSyncPanel() {
    const urlInput = document.getElementById("sb-url");
    const anonInput = document.getElementById("sb-anon-key");
    const tableInput = document.getElementById("sb-table");
    const keyInput = document.getElementById("sb-state-key");
    const autoSyncInput = document.getElementById("sb-auto-sync");

    if (urlInput && urlInput !== document.activeElement) urlInput.value = cleanText(cloudConfig.url);
    if (anonInput && anonInput !== document.activeElement) anonInput.value = cleanText(cloudConfig.anonKey);
    if (tableInput && tableInput !== document.activeElement) tableInput.value = cleanText(cloudConfig.table);
    if (keyInput && keyInput !== document.activeElement) keyInput.value = cleanText(cloudConfig.stateKey);
    if (autoSyncInput) autoSyncInput.checked = !!cloudConfig.autoSync;

    setText("sb-last-pull", cloudConfig.lastPullAt ? formatDateTime(cloudConfig.lastPullAt) : "-");
    setText("sb-last-push", cloudConfig.lastPushAt ? formatDateTime(cloudConfig.lastPushAt) : "-");
    setText("sb-cloud-mode", isCloudSyncConfigured() ? "Đã cấu hình" : "Chưa cấu hình");
  }

  function renderPartnerInfo() {
    const syncMeta = state.partnerCare.sync || {};
    const endpointInput = document.getElementById("pc-sync-endpoint");
    if (endpointInput && endpointInput !== document.activeElement) endpointInput.value = cleanText(syncMeta.endpoint);
    setText("pc-sync-last", syncMeta.lastSyncAt ? formatDateTime(syncMeta.lastSyncAt) : "-");
    setText("pc-sync-count", asNumber(syncMeta.lastCount) || 0);

    const rows = (state.partnerCare.crmImportRows || []).map((row, idx) => normalizeCrmImportRow(row, idx));
    const fallbackSummary = buildCrmSummaryFromRows(rows);
    const storedSummary = state.partnerCare.crmSummary || {};
    const summary = {
      totalPartner: asNumber(storedSummary.totalPartner) || fallbackSummary.totalPartner,
      existingEligible: asNumber(storedSummary.existingEligible) || fallbackSummary.existingEligible,
      newCount: asNumber(storedSummary.newCount) || fallbackSummary.newCount,
      noTouchThisYear: asNumber(storedSummary.noTouchThisYear) || fallbackSummary.noTouchThisYear,
      activePartner: asNumber(storedSummary.activePartner) || fallbackSummary.activePartner,
    };

    setText("pi-total-partner", formatNumber(summary.totalPartner, 0));
    setText("pi-existing-eligible", formatNumber(summary.existingEligible, 0));
    setText("pi-new-count", formatNumber(summary.newCount, 0));
    setText("pi-no-touch", formatNumber(summary.noTouchThisYear, 0));
    setText("pi-active-partner", formatNumber(summary.activePartner, 0));

    const table = document.getElementById("pi-crm-table");
    if (!table) return;
    table.innerHTML = "";
    if (!rows.length) {
      table.innerHTML = `<tr><td colspan="12" class="empty-row">Chưa có dữ liệu CRM_IMPORT. Bấm Sync CRM_IMPORT để tải dữ liệu.</td></tr>`;
      return;
    }

    rows
      .slice()
      .sort((a, b) => cleanText(a.partnerCode || a.partnerKey).localeCompare(cleanText(b.partnerCode || b.partnerKey)))
      .forEach((row) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${safeText(row.rowId ? String(row.rowId) : "")}</td>
          <td>${safeText(row.partnerCode || row.partnerKey)}</td>
          <td>${safeText(row.partnerName)}</td>
          <td>${safeText(row.partnerTypeLabel)}</td>
          <td>${safeText(row.owner)}</td>
          <td>${safeText(row.branchName)}</td>
          <td>${safeText(row.region)}</td>
          <td>${safeText(row.statusText)}</td>
          <td>${safeText(formatSheetDate(row.contractDate))}</td>
          <td>${safeText(formatSheetDate(row.openingDate))}</td>
          <td>${safeText(formatSheetDate(row.closingDate))}</td>
          <td>${safeText(row.eligibleExistingThisYear === 1 ? "1" : "0")}</td>
        `;
        table.appendChild(tr);
      });
  }

  function renderImportCareSetup() {
    const monthSelect = document.getElementById("import-care-month");
    const table = document.getElementById("import-care-draft-table");
    if (!monthSelect || !table) return;

    const selectedMonth = normalizeMonth(monthSelect.value || importCareDraft.month || getPartnerCareMonth());
    loadImportCareDraftMonth(selectedMonth);
    monthSelect.value = selectedMonth;

    refreshImportCarePartnerOptions();
    syncImportCareFormFromPartnerSelection();

    table.innerHTML = "";
    if (!importCareDraft.rows.length) {
      table.innerHTML = `<tr><td colspan="8" class="empty-row">Chưa có dòng setup cho tháng ${selectedMonth}.</td></tr>`;
      return;
    }

    importCareDraft.rows
      .slice()
      .sort((a, b) => {
        const dayA = normalizeCareDay(a.careDay) ?? 99;
        const dayB = normalizeCareDay(b.careDay) ?? 99;
        if (dayA !== dayB) return dayA - dayB;
        return cleanText(a.partnerKey).localeCompare(cleanText(b.partnerKey));
      })
      .forEach((row) => {
        const tr = document.createElement("tr");
        tr.setAttribute("data-import-care-id", cleanText(row.id));
        const careDay = normalizeCareDay(row.careDay);
        tr.innerHTML = `
          <td>${safeText(careDay === null ? "-" : String(careDay))}</td>
          <td>${safeText(row.partnerKey || "-")}</td>
          <td>${safeText(normalizePartnerType(row.partnerType))}</td>
          <td>${safeText(row.partnerName || "-")}</td>
          <td>${safeText(row.touchpointType || "-")}</td>
          <td>${safeText(row.channel || "-")}</td>
          <td>${safeText(row.owner || "-")}</td>
          <td><button class="btn btn-ghost" data-import-care-remove="${safeText(row.id)}">Xóa</button></td>
        `;
        table.appendChild(tr);
      });
  }

  function loadImportCareDraftMonth(month) {
    const selectedMonth = normalizeMonth(month);
    importCareDraft.month = selectedMonth;
    importCareDraft.rows = getPartnerCarePlanRowsByMonth(selectedMonth).map((row) => ({ ...row }));
    const monthSelect = document.getElementById("import-care-month");
    if (monthSelect) monthSelect.value = selectedMonth;
  }

  function refreshImportCarePartnerOptions() {
    const select = document.getElementById("import-care-partner-key");
    if (!select) return;
    const current = cleanText(select.value).toUpperCase();
    const partners = getPartnerMasterRecords().filter((row) => isPartnerActiveStatus(row.status));
    select.innerHTML = "";
    select.appendChild(makeOption("", "-- Chọn mã đối tác --"));
    partners.forEach((row) => {
      const key = cleanText(row.partnerKey).toUpperCase();
      if (!key) return;
      const label = row.partnerName ? `${key} - ${row.partnerName}` : key;
      select.appendChild(makeOption(key, label));
    });
    if (current && partners.some((row) => cleanText(row.partnerKey).toUpperCase() === current)) {
      select.value = current;
    }
  }

  function syncImportCareFormFromPartnerSelection() {
    const partnerKey = cleanText(document.getElementById("import-care-partner-key")?.value).toUpperCase();
    const partner = getPartnerMasterByKey(partnerKey);
    const nameInput = document.getElementById("import-care-partner-name");
    const typeInput = document.getElementById("import-care-partner-type");
    const ownerInput = document.getElementById("import-care-owner");
    const channelInput = document.getElementById("import-care-channel");

    if (!partnerKey) {
      if (nameInput) nameInput.value = "";
      if (typeInput) typeInput.value = "";
      return;
    }

    if (nameInput) nameInput.value = cleanText(partner?.partnerName);
    if (typeInput) typeInput.value = cleanText(partner?.partnerType ? normalizePartnerType(partner.partnerType) : "");
    if (ownerInput && !cleanText(ownerInput.value)) ownerInput.value = cleanText(partner?.owner);
    if (channelInput && !cleanText(channelInput.value)) channelInput.value = "Call/message";
  }

  function renderAnnualOkrTable() {
    const body = document.getElementById("annual-okr-table");
    body.innerHTML = "";
    if (!state.importHub.annualOkrs.length) {
      body.innerHTML = `<tr><td colspan="6" class="empty-row">Chưa có OKR năm</td></tr>`;
      return;
    }

    state.importHub.annualOkrs.forEach((row) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${safeText(PILLARS[row.pillarKey]?.label || row.pillarKey)}</td>
        <td>${safeText(row.code)}</td>
        <td>${safeText(row.title)}</td>
        <td>${formatMetric(row.target, row.unit)}</td>
        <td>${safeText(row.unit)}</td>
        <td><button class="btn btn-ghost" data-remove-annual="${safeText(row.id)}">Xóa</button></td>
      `;
      body.appendChild(tr);
    });
  }

  function renderQuarterlyOkrTable() {
    const body = document.getElementById("quarterly-okr-table");
    body.innerHTML = "";
    if (!state.importHub.quarterlyOkrs.length) {
      body.innerHTML = `<tr><td colspan="8" class="empty-row">Chưa có OKR quý</td></tr>`;
      return;
    }

    state.importHub.quarterlyOkrs
      .slice()
      .sort((a, b) => QUARTER_ORDER(a.quarter) - QUARTER_ORDER(b.quarter))
      .forEach((row) => {
        const linkedAnnual = state.importHub.annualOkrs.find((item) => item.id === row.linkedAnnualOkrId);
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${safeText(row.quarter)}</td>
          <td>${safeText(PILLARS[row.pillarKey]?.label || row.pillarKey)}</td>
          <td>${safeText(row.code)}</td>
          <td>${safeText(row.title)}</td>
          <td>${formatMetric(row.target, row.unit)}</td>
          <td>${safeText(row.unit)}</td>
          <td>${safeText(linkedAnnual?.title || "")}</td>
          <td><button class="btn btn-ghost" data-remove-quarterly-okr="${safeText(row.id)}">Xóa</button></td>
        `;
        body.appendChild(tr);
      });
  }

  function renderMonthlyOkrTable() {
    const body = document.getElementById("monthly-okr-table");
    body.innerHTML = "";
    const rows = state.importHub.monthlyOkrs.filter((row) => asNumber(row.target) !== null || cleanText(row.notes) || cleanText(row.title));
    if (!rows.length) {
      body.innerHTML = `<tr><td colspan="7" class="empty-row">Chưa có OKR tháng</td></tr>`;
      return;
    }

    rows
      .slice()
      .sort((a, b) => MONTHS.indexOf(a.month) - MONTHS.indexOf(b.month))
      .forEach((row) => {
        const linkedQuarterly = state.importHub.quarterlyOkrs.find((item) => item.id === row.linkedQuarterlyOkrId);
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${safeText(row.month)}</td>
          <td>${safeText(PILLARS[row.pillarKey]?.label || row.pillarKey)}</td>
          <td>${safeText(linkedQuarterly?.title || "")}</td>
          <td>${safeText(row.title)}</td>
          <td>${formatMetric(row.target, row.unit)}</td>
          <td>${safeText(row.unit)}</td>
          <td><button class="btn btn-ghost" data-remove-monthly-okr="${safeText(row.id)}">Xóa</button></td>
        `;
        body.appendChild(tr);
      });
  }

  function renderMonthlyPlanTable() {
    const body = document.getElementById("monthly-plan-table");
    body.innerHTML = "";
    if (!state.importHub.monthlyPlans.length) {
      body.innerHTML = `<tr><td colspan="9" class="empty-row">Chưa có kế hoạch tháng</td></tr>`;
      return;
    }

    state.importHub.monthlyPlans
      .slice()
      .sort((a, b) => MONTHS.indexOf(a.month) - MONTHS.indexOf(b.month))
      .forEach((row) => {
        const linkedOkr = state.importHub.monthlyOkrs.find((item) => item.id === row.linkedMonthlyOkrId);
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${safeText(row.month)}</td>
          <td>${safeText(PILLARS[row.pillarKey]?.label || row.pillarKey)}</td>
          <td>${safeText(linkedOkr?.title || "")}</td>
          <td>${safeText(row.action)}</td>
          <td>${safeText(row.owner)}</td>
          <td>${safeText(row.due)}</td>
          <td>${safeText(row.status)}</td>
          <td>${renderKpiResultPreview(row.result || row.notes)}</td>
          <td><button class="btn btn-ghost" data-remove-monthly-plan="${safeText(row.id)}">Xóa</button></td>
        `;
        body.appendChild(tr);
      });
  }

  function renderMonthlyKpiTable() {
    const body = document.getElementById("monthly-kpi-table");
    body.innerHTML = "";
    if (!state.importHub.monthlyKpis.length) {
      body.innerHTML = `<tr><td colspan="10" class="empty-row">Chưa có KPI tháng</td></tr>`;
      return;
    }

    state.importHub.monthlyKpis
      .slice()
      .sort((a, b) => MONTHS.indexOf(a.month) - MONTHS.indexOf(b.month))
      .forEach((row) => {
        const linkedOkr = state.importHub.monthlyOkrs.find((item) => item.id === row.monthlyOkrId);
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${safeText(row.month)}</td>
          <td>${safeText(PILLARS[row.pillarKey]?.label || row.pillarKey)}</td>
          <td>${safeText(linkedOkr?.title || "")}</td>
          <td>${safeText(row.kpiName)}</td>
          <td>${formatMetric(row.target, row.unit)}</td>
          <td>${formatMetric(row.actual, row.unit)}</td>
          <td>${safeText(row.unit)}</td>
          <td>${safeText(row.status)}</td>
          <td>${renderKpiResultPreview(row.result || row.notes)}</td>
          <td><button class="btn btn-ghost" data-remove-monthly-kpi="${safeText(row.id)}">Xóa</button></td>
        `;
        body.appendChild(tr);
      });
  }

  function renderPartnerCare() {
    const month = getPartnerCareMonth();
    const monthLabel = uiState.month === ALL_MONTH_VALUE ? `${month} (tự chọn khi bộ lọc đang là Tất cả)` : month;
    const typeFilter = cleanText(document.getElementById("pc-filter-type")?.value);
    const statusFilter = cleanText(document.getElementById("pc-filter-status")?.value);
    const careRowsAll = getPartnerCarePlanRowsByMonth(month);
    const careRowsFiltered = careRowsAll.filter((row) => {
      if (typeFilter && cleanText(normalizePartnerType(row.partnerType)) !== typeFilter) return false;
      if (statusFilter && cleanText(row.status) !== statusFilter) return false;
      return true;
    });

    const logsInMonth = state.partnerCare.logs.filter((log) => cleanText(log.month) === month);
    const filtered = logsInMonth.filter((log) => {
      if (typeFilter && cleanText(log.partnerType) !== typeFilter) return false;
      if (statusFilter && cleanText(log.status) !== statusFilter) return false;
      return true;
    });

    const legacyPlan = state.partnerCare.monthlyPlan.find((item) => cleanText(item.month) === month);
    const plannedFromSheet = (asNumber(legacyPlan?.plannedNew) || 0) + (asNumber(legacyPlan?.plannedExisting) || 0);
    const hasCareRows = careRowsAll.length > 0;
    const planned = hasCareRows ? careRowsAll.length : plannedFromSheet;
    const done = hasCareRows
      ? careRowsAll.filter((row) => cleanText(row.status).toLowerCase() === "done").length
      : filtered.filter((log) => cleanText(log.status).toLowerCase() === "done").length;
    const compliance = planned > 0 ? done / planned : 0;
    const partnerSet = new Set(
      (hasCareRows ? careRowsAll : filtered).map((item) => `${safeText(item.partnerKey)}__${safeText(item.partnerName)}`)
    );

    setText("pc-planned", formatNumber(planned, 0));
    setText("pc-done", formatNumber(done, 0));
    setText("pc-compliance", formatPercent(compliance));
    setText("pc-partners", partnerSet.size);

    const partnerSummaryMap = {};
    const partnerSummarySource = hasCareRows ? careRowsFiltered : filtered;
    partnerSummarySource.forEach((row) => {
      const key = `${safeText(row.partnerKey)}__${safeText(row.partnerName)}`;
      if (!partnerSummaryMap[key]) {
        partnerSummaryMap[key] = {
          partnerKey: safeText(row.partnerKey),
          partnerName: safeText(row.partnerName),
          partnerType: safeText(normalizePartnerType(row.partnerType)),
          total: 0,
          done: 0,
        };
      }
      partnerSummaryMap[key].total += 1;
      if (cleanText(row.status).toLowerCase() === "done") partnerSummaryMap[key].done += 1;
    });
    const partnerRows = Object.values(partnerSummaryMap).sort((a, b) => b.total - a.total);

    const partnerTable = document.getElementById("pc-partner-table");
    partnerTable.innerHTML = "";
    if (!partnerRows.length) {
      partnerTable.innerHTML = `<tr><td colspan="5" class="empty-row">Không có dữ liệu phù hợp bộ lọc</td></tr>`;
    } else {
      partnerRows.forEach((row) => {
        const doneRate = row.total > 0 ? row.done / row.total : 0;
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${safeText(row.partnerKey)} - ${safeText(row.partnerName)}</td>
          <td>${safeText(row.partnerType)}</td>
          <td>${row.total}</td>
          <td>${row.done}</td>
          <td>${formatPercent(doneRate)}</td>
        `;
        partnerTable.appendChild(tr);
      });
    }

    const careRosterTable = document.getElementById("pc-care-roster-table");
    if (careRosterTable) {
      careRosterTable.innerHTML = "";
      if (!careRowsFiltered.length) {
        careRosterTable.innerHTML = `<tr><td colspan="12" class="empty-row">Chưa có danh sách cần chăm sóc tháng ${monthLabel}.</td></tr>`;
      } else {
        careRowsFiltered
          .slice()
          .sort((a, b) => (normalizeCareDay(a.careDay) || 99) - (normalizeCareDay(b.careDay) || 99))
          .forEach((row) => {
            const tr = document.createElement("tr");
            tr.setAttribute("data-pc-care-id", cleanText(row.id));
            const statusValue = cleanText(row.status) || "Planned";
            const statusOptions = ["Planned", "Done", "Missed", "Rescheduled"]
              .map(
                (value) => `<option value="${value}"${statusValue === value ? " selected" : ""}>${value}</option>`
              )
              .join("");
            const statusHtml =
              statusValue && !["Planned", "Done", "Missed", "Rescheduled"].includes(statusValue)
                ? `${statusOptions}<option value="${safeText(statusValue)}" selected>${safeText(statusValue)}</option>`
                : statusOptions;
            tr.innerHTML = `
              <td>${safeText(normalizeCareDay(row.careDay) === null ? "" : String(normalizeCareDay(row.careDay)))}</td>
              <td>${safeText(row.partnerKey)}</td>
              <td>${safeText(normalizePartnerType(row.partnerType))}</td>
              <td>${safeText(row.partnerName)}</td>
              <td>${safeText(row.touchpointType)}</td>
              <td>${safeText(row.channel)}</td>
              <td>${safeText(row.owner)}</td>
              <td><select class="kpi-cell-input" data-pc-care-status>${statusHtml}</select></td>
              <td>
                <input
                  class="kpi-cell-input"
                  data-pc-care-day-done
                  type="number"
                  min="1"
                  max="31"
                  step="1"
                  value="${safeText(normalizeCareDay(row.dayDone) === null ? "" : String(normalizeCareDay(row.dayDone)))}"
                />
              </td>
              <td><input class="kpi-cell-input" data-pc-care-result type="text" value="${safeText(row.result)}" /></td>
              <td><input class="kpi-cell-input" data-pc-care-response type="text" value="${safeText(row.response)}" /></td>
              <td><button class="btn btn-primary" type="button" data-pc-care-commit>Ghi nhận</button></td>
            `;
            careRosterTable.appendChild(tr);
          });
      }
    }

    const logTable = document.getElementById("pc-log-table");
    logTable.innerHTML = "";
    if (!filtered.length) {
      logTable.innerHTML = `<tr><td colspan="9" class="empty-row">Không có log</td></tr>`;
    } else {
      filtered.slice(0, 120).forEach((log) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${safeText(log.date)}</td>
          <td>${safeText(log.month)}</td>
          <td>${safeText(log.partnerKey)} - ${safeText(log.partnerName)}</td>
          <td>${safeText(log.partnerType)}</td>
          <td>${safeText(log.touchpointType)}</td>
          <td>${safeText(log.owner)}</td>
          <td>${safeText(log.status)}</td>
          <td>${truncateText(safeText(log.result), 120)}</td>
          <td>${truncateText(safeText(log.response), 120)}</td>
        `;
        logTable.appendChild(tr);
      });
    }

    const doneFromSummarySource = partnerSummarySource.filter((row) => cleanText(row.status).toLowerCase() === "done").length;
    const doneRate = partnerSummarySource.length > 0 ? doneFromSummarySource / partnerSummarySource.length : 0;
    const topPartner = partnerRows[0];
    const insights = [
      `Nhận định: Tháng ${monthLabel} đạt compliance ${formatPercent(compliance)} (${done}/${formatNumber(planned, 0)} planned).`,
      `Nhận xét: Tỷ lệ hoàn thành trong danh sách đã lọc là ${formatPercent(doneRate)}; ${
        topPartner ? `partner tương tác nhiều nhất là ${topPartner.partnerName}.` : "chưa có partner nổi bật."
      }`,
      `Phân tích: Có ${careRowsAll.length} đối tác trong danh sách cần chăm sóc tháng ${monthLabel}, ${filtered.length} log phát sinh thực tế.`,
      "Đề xuất hành động: Mỗi tuần rà danh sách cần chăm sóc trước, sau đó cập nhật log Done để theo dõi compliance chính xác.",
    ];
    renderInsightList("pc-insight", insights);
  }

  function renderIntlProblem() {
    const rows = [...state.intlProblems];
    const total = rows.length;
    const openCount = rows.filter((item) => {
      const status = cleanText(item.status).toLowerCase();
      return status && status !== "closed" && status !== "resolved";
    }).length;
    const breachCount = rows.filter((item) => {
      const sla = asNumber(item.slaHours);
      const resolved = asNumber(item.resolvedHours);
      return sla !== null && resolved !== null && resolved > sla;
    }).length;
    const resolvedHours = rows.map((r) => asNumber(r.resolvedHours)).filter((v) => v !== null);
    const avgResolve =
      resolvedHours.length > 0 ? resolvedHours.reduce((sum, val) => sum + val, 0) / resolvedHours.length : 0;

    setText("ip-total-case", total);
    setText("ip-open-case", openCount);
    setText("ip-breach-case", breachCount);
    setText("ip-avg-resolve", formatNumber(avgResolve, 1));

    const categoryMap = {};
    rows.forEach((row) => {
      const key = cleanText(row.category) || "Uncategorized";
      categoryMap[key] = (categoryMap[key] || 0) + 1;
    });
    const topCategory = Object.entries(categoryMap).sort((a, b) => b[1] - a[1])[0];

    const insights = [
      `Nhận định: Có ${total} case quốc tế, trong đó ${openCount} case đang mở.`,
      `Nhận xét: ${
        topCategory ? `Nhóm vấn đề nổi bật là "${topCategory[0]}" với ${topCategory[1]} case.` : "Chưa đủ dữ liệu để phân nhóm."
      }`,
      `Phân tích: SLA breach hiện ${breachCount} case, thời gian xử lý trung bình ${formatNumber(avgResolve, 1)} giờ.`,
      "Đề xuất hành động: Chuẩn hóa taxonomy root cause, đặt ngưỡng cảnh báo SLA và review post-mortem cho mọi case breach.",
    ];
    renderInsightList("ip-insight", insights);

    const table = document.getElementById("ip-table");
    table.innerHTML = "";
    if (!rows.length) {
      table.innerHTML = `<tr><td colspan="9" class="empty-row">Chưa có case quốc tế</td></tr>`;
      return;
    }

    rows
      .sort((a, b) => MONTHS.indexOf(cleanText(a.month)) - MONTHS.indexOf(cleanText(b.month)))
      .forEach((row) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${safeText(row.date)}</td>
          <td>${safeText(row.month)}</td>
          <td>${safeText(row.case)}</td>
          <td>${safeText(row.category)}</td>
          <td>${safeText(row.rootCause)}</td>
          <td>${safeText(row.status)}</td>
          <td>${formatMetric(row.slaHours, "")}</td>
          <td>${formatMetric(row.resolvedHours, "")}</td>
          <td>${safeText(row.result)}</td>
        `;
        table.appendChild(tr);
      });
  }

  function getPartnerCareMonth() {
    if (uiState.month !== ALL_MONTH_VALUE) return uiState.month;
    return guessDefaultMonth();
  }

  function getPartnerCarePlanRowsByMonth(month) {
    const selectedMonth = normalizeMonth(month);
    return (state.partnerCare.partnerMonthlyPlans || [])
      .filter((row) => cleanText(row.month) === selectedMonth)
      .map((row, idx) => ({
        id: cleanText(row.id) || `PCP-${idx + 1}-${Date.now()}`,
        month: selectedMonth,
        careDay: normalizeCareDay(row.careDay),
        dayDone: normalizeCareDay(row.dayDone),
        partnerKey: cleanText(row.partnerKey).toUpperCase(),
        partnerName: cleanText(row.partnerName),
        partnerType: normalizePartnerType(row.partnerType),
        owner: cleanText(row.owner),
        touchpointType: cleanText(row.touchpointType),
        channel: cleanText(row.channel) || "Call/message",
        status: cleanText(row.status) || "Planned",
        result: cleanText(row.result),
        response: cleanText(row.response),
        targetTouchpoints: Math.max(0, asNumber(row.targetTouchpoints) ?? 1),
        nextAction: cleanText(row.nextAction),
      }));
  }

  function upsertPartnerCareLogFromPlan(planRow) {
    if (!planRow) return;
    if (!Array.isArray(state.partnerCare.logs)) state.partnerCare.logs = [];

    const sourcePlanId = cleanText(planRow.id);
    const logPayload = {
      sourcePlanId,
      date: normalizeCareDay(planRow.dayDone),
      month: normalizeMonth(planRow.month),
      partnerKey: cleanText(planRow.partnerKey).toUpperCase(),
      partnerName: cleanText(planRow.partnerName),
      partnerType: normalizePartnerType(planRow.partnerType),
      touchpointType: cleanText(planRow.touchpointType),
      owner: cleanText(planRow.owner),
      status: cleanText(planRow.status),
      result: cleanText(planRow.result),
      response: cleanText(planRow.response),
      updatedAt: new Date().toISOString(),
    };

    const existingIndex = state.partnerCare.logs.findIndex((log) => cleanText(log.sourcePlanId) === sourcePlanId);
    if (existingIndex >= 0) {
      const previous = state.partnerCare.logs[existingIndex] || {};
      state.partnerCare.logs[existingIndex] = {
        ...previous,
        ...logPayload,
      };
      return;
    }

    state.partnerCare.logs.push({
      id: uid("PCL"),
      ...logPayload,
    });
  }

  function getPartnerPlanFormMonth() {
    const select = document.getElementById("pc-plan-month");
    const selected = cleanText(select?.value);
    if (MONTHS.includes(selected)) return selected;
    return getPartnerCareMonth();
  }

  function getPartnerMasterByKey(partnerKey) {
    const key = cleanText(partnerKey).toUpperCase();
    if (!key) return null;
    return getPartnerMasterRecords().find((row) => cleanText(row.partnerKey).toUpperCase() === key) || null;
  }

  function getPartnerRuleTargetForMonth(month, partnerType) {
    const monthIndex = MONTHS.indexOf(normalizeMonth(month));
    if (monthIndex < 0) return 0;
    const yearlyTarget = getYearlyTouchpointTarget(partnerType);
    const ruleMonths = getRuleMonthsByYearTarget(yearlyTarget);
    return ruleMonths.has(monthIndex) ? 1 : 0;
  }

  function refreshPartnerPlanPartnerOptions() {
    const select = document.getElementById("pc-plan-partner-key");
    if (!select) return;
    const current = cleanText(select.value).toUpperCase();
    const partnerRows = getPartnerMasterRecords();
    select.innerHTML = "";
    select.appendChild(makeOption("", "-- Chọn mã đối tác --"));
    partnerRows.forEach((row) => {
      const key = cleanText(row.partnerKey).toUpperCase();
      if (!key) return;
      const label = row.partnerName ? `${key} - ${row.partnerName}` : key;
      select.appendChild(makeOption(key, label));
    });
    if (current && partnerRows.some((row) => cleanText(row.partnerKey).toUpperCase() === current)) {
      select.value = current;
    }
  }

  function syncPartnerPlanFormAutofill() {
    const partnerKey = cleanText(document.getElementById("pc-plan-partner-key")?.value).toUpperCase();
    const month = getPartnerPlanFormMonth();
    const partner = getPartnerMasterByKey(partnerKey);
    const partnerNameInput = document.getElementById("pc-plan-partner-name");
    const partnerTypeInput = document.getElementById("pc-plan-partner-type");
    const ownerInput = document.getElementById("pc-plan-owner");
    const channelInput = document.getElementById("pc-plan-channel");
    const statusInput = document.getElementById("pc-plan-status");
    if (!partnerKey) {
      if (partnerNameInput) partnerNameInput.value = "";
      if (partnerTypeInput) partnerTypeInput.value = "";
      if (ownerInput) ownerInput.value = "";
      return;
    }

    if (partnerNameInput) partnerNameInput.value = cleanText(partner?.partnerName);
    if (partnerTypeInput) {
      const typeValue = cleanText(partner?.partnerType);
      partnerTypeInput.value = typeValue ? normalizePartnerType(typeValue) : "";
    }
    if (ownerInput) ownerInput.value = cleanText(partner?.owner);

    if (channelInput && !cleanText(channelInput.value)) channelInput.value = "Call/message";
    if (statusInput && !cleanText(statusInput.value)) {
      statusInput.value = getPartnerRuleTargetForMonth(month, partner?.partnerType) > 0 ? "Planned" : "";
    }
  }

  function isActivePartnerPlanRow(row) {
    const targetTouchpoints = asNumber(row?.targetTouchpoints);
    const hasDetail =
      normalizeCareDay(row?.careDay) !== null ||
      normalizeCareDay(row?.dayDone) !== null ||
      !!cleanText(row?.touchpointType) ||
      !!cleanText(row?.channel) ||
      !!cleanText(row?.status) ||
      !!cleanText(row?.result) ||
      !!cleanText(row?.response);
    if (targetTouchpoints === null) return true;
    return targetTouchpoints > 0 || hasDetail;
  }

  function normalizePartnerType(value) {
    const txt = normalizeCompareText(value);
    if (!txt) return "Existed";
    if (txt.includes("new") || txt.includes("moi")) return "New";
    if (txt.includes("exist") || txt.includes("hien huu") || txt.includes("current")) return "Existed";
    return "Existed";
  }

  function resolveLifecycleTypeFromDates(contractDateValue, openingDateValue) {
    const endOfMonth = new Date();
    endOfMonth.setHours(23, 59, 59, 999);
    endOfMonth.setMonth(endOfMonth.getMonth() + 1, 0);

    const parsedDates = [parseFlexibleDate(contractDateValue), parseFlexibleDate(openingDateValue)]
      .filter((date) => date instanceof Date && !Number.isNaN(date.getTime()))
      .sort((a, b) => a.getTime() - b.getTime());

    if (!parsedDates.length) return "";

    const nonFutureDates = parsedDates.filter((date) => date.getTime() <= endOfMonth.getTime());
    const anchorDate = nonFutureDates.length ? nonFutureDates[0] : parsedDates[0];

    const startOfNewWindow = new Date(endOfMonth);
    startOfNewWindow.setDate(endOfMonth.getDate() - 90);
    if (anchorDate.getTime() >= startOfNewWindow.getTime() && anchorDate.getTime() <= endOfMonth.getTime()) {
      return "New";
    }

    const startOfYear = new Date(endOfMonth.getFullYear(), 0, 1);
    if (anchorDate.getTime() < startOfYear.getTime()) return "Existed";

    return "";
  }

  function normalizeCrmPartnerType(value) {
    const txt = normalizeCompareText(value);
    if (!txt) return "";
    if (txt.includes("notouch")) return "NoTouchThisYear";
    if (txt.includes("do not touch")) return "DoNotTouchThisYear";
    if (txt.includes("new") || txt.includes("moi")) return "New";
    if (txt.includes("exist") || txt.includes("hien huu") || txt.includes("current")) return "Existed";
    return cleanText(value);
  }

  function normalizeCrmImportRow(row, idx = 0) {
    const partnerCode = cleanText(
      readAny(row, ["partnerCode", "partner_code", "Partner_Code", "partnerKey", "partner_key", "code", "id"])
    ).toUpperCase();
    const partnerKey = cleanText(
      readAny(row, ["partnerKey", "partner_key", "Partner_Key", "partnerCode", "partner_code", "Partner_Code"])
    ).toUpperCase();
    const contractDateRaw = readAny(row, ["contractDate", "contract_date", "Contract_Date"]);
    const openingDateRaw = readAny(row, ["openingDate", "opening_date", "Opening_Date"]);
    const rawPartnerType = readAny(row, [
      "partnerType",
      "partner_type",
      "partner_type_new_existing",
      "new_existed",
      "new_existing",
      "cohort",
      "current_type",
      "type",
      "segment",
    ]);
    const normalizedPartnerType = normalizeCrmPartnerType(rawPartnerType);
    const explicitLifecycleType = ["New", "Existed"].includes(normalizedPartnerType) ? normalizedPartnerType : "";
    const lifecycleFromDates = resolveLifecycleTypeFromDates(contractDateRaw, openingDateRaw);
    const partnerTypeLabel =
      lifecycleFromDates ||
      explicitLifecycleType ||
      normalizedPartnerType ||
      normalizePartnerType(rawPartnerType);
    const partnerNameRaw = cleanText(
      readAny(row, [
        "partnerName",
        "partner_name",
        "Partner_Name",
        "partner",
        "partner_full_name",
        "branchName",
        "branch_name",
        "name",
      ])
    );
    const ownerRaw = cleanText(
      readAny(row, [
        "owner",
        "Owner",
        "pic",
        "account_owner",
        "accountManager",
        "account_manager",
        "am",
        "bdm",
        "rm",
        "sales_owner",
        "franchise_owner",
        "manager",
        "nguoi_phu_trach",
        "nguoi_quan_ly",
      ])
    );
    const statusText = cleanText(
      readAny(row, [
        "statusText",
        "status_text",
        "Status_Text",
        "status",
        "statusActive",
        "status_active",
        "partner_status",
        "operating_status",
        "operatingStatus",
        "trang_thai",
        "tinh_trang",
      ])
    );
    const closingDateRaw = readAny(row, ["closingDate", "closing_date", "Closing_Date"]);
    const resolvedPartnerName = partnerNameRaw || ownerRaw;
    return {
      rowId: asNumber(readAny(row, ["rowId", "row_id", "Row_ID", "index"])) || idx + 1,
      branchName: cleanText(readAny(row, ["branchName", "branch_name", "Branch_Name", "store_name", "chi_nhanh"])),
      region: cleanText(readAny(row, ["region", "Region", "area", "khu_vuc"])),
      phone: cleanText(readAny(row, ["phone", "Phone"])),
      partnerCode: partnerCode || partnerKey,
      owner: normalizeOwnerValue(ownerRaw, resolvedPartnerName),
      partnerKey: partnerKey || partnerCode,
      contractDate: contractDateRaw,
      openingDate: openingDateRaw,
      closingDate: closingDateRaw,
      statusText: inferStatusValue(statusText, closingDateRaw),
      contractYear: asNumber(readAny(row, ["contractYear", "contract_year", "Contract_Year"])),
      contractMonth: asNumber(readAny(row, ["contractMonth", "contract_month", "Contract_Month"])),
      partnerFirstOpen: readAny(row, ["partnerFirstOpen", "partner_first_open", "Partner_First_Open"]),
      eligibleExistingThisYear: asNumber(
        readAny(row, ["eligibleExistingThisYear", "eligible_existing_this_year", "Eligible_Existing_ThisYear"])
      ),
      partnerName: resolvedPartnerName,
      partnerTypeLabel,
    };
  }

  function getObjectPropertyByNames(obj, keys) {
    if (!obj || typeof obj !== "object") return undefined;
    for (const key of keys) {
      if (key in obj) return obj[key];
    }
    const map = {};
    Object.keys(obj).forEach((key) => {
      map[normalizeCompareText(key).replace(/\s+/g, "")] = obj[key];
    });
    for (const key of keys) {
      const normalized = normalizeCompareText(key).replace(/\s+/g, "");
      if (normalized in map) return map[normalized];
    }
    return undefined;
  }

  function rowsFromValues(values) {
    if (!Array.isArray(values) || !Array.isArray(values[0])) return [];
    const headers = values[0].map((item) => cleanText(item));
    return values.slice(1).map((row) => {
      const obj = {};
      headers.forEach((header, idx) => {
        obj[header] = row[idx];
      });
      return obj;
    });
  }

  function coerceRowsFromAny(raw) {
    if (!raw) return [];
    if (Array.isArray(raw)) {
      if (!raw.length) return [];
      if (Array.isArray(raw[0])) return rowsFromValues(raw);
      if (typeof raw[0] === "object") return raw;
      return [];
    }
    if (typeof raw !== "object") return [];

    const directRows = getObjectPropertyByNames(raw, ["rows", "data", "items", "records"]);
    if (directRows) {
      const parsed = coerceRowsFromAny(directRows);
      if (parsed.length) return parsed;
    }

    const values = getObjectPropertyByNames(raw, ["values"]);
    if (values) {
      const parsed = coerceRowsFromAny(values);
      if (parsed.length) return parsed;
    }

    return [];
  }

  function extractCrmImportRowsFromPayload(payload) {
    if (!payload || typeof payload !== "object") return coerceRowsFromAny(payload);
    const direct = getObjectPropertyByNames(payload, [
      "crmImport",
      "crm_import",
      "CRM_IMPORT",
      "crmImportRows",
      "crm_import_rows",
    ]);
    if (direct) {
      const rows = coerceRowsFromAny(direct);
      if (rows.length) return rows;
    }

    const sheets = getObjectPropertyByNames(payload, ["sheets", "sheetData", "tabs"]);
    if (sheets && typeof sheets === "object") {
      const crmSheet = getObjectPropertyByNames(sheets, [
        "crmImport",
        "crm_import",
        "CRM_IMPORT",
        "crmImportRows",
        "crm_import_rows",
      ]);
      if (crmSheet) {
        const rows = coerceRowsFromAny(crmSheet);
        if (rows.length) return rows;
      }
    }

    return coerceRowsFromAny(payload);
  }

  function buildCrmSummaryFromRows(rows) {
    const rowMap = new Map();
    rows.forEach((row) => {
      const key = cleanText(row.partnerCode || row.partnerKey);
      if (!key) return;
      if (!rowMap.has(key)) rowMap.set(key, row);
    });
    const uniqueRows = Array.from(rowMap.values());
    const existingEligible = uniqueRows.filter((row) => asNumber(row.eligibleExistingThisYear) === 1).length;
    const newCount = uniqueRows.filter((row) => normalizeCrmPartnerType(row.partnerTypeLabel) === "New").length;
    const noTouchThisYear = uniqueRows.filter((row) => normalizeCrmPartnerType(row.partnerTypeLabel) === "NoTouchThisYear").length;
    const activePartner = uniqueRows.filter((row) => {
      const status = normalizeCompareText(row.statusText);
      if (!status) return true;
      if (status.includes("closed") || status.includes("dong cua")) return false;
      return status.includes("hoat") || status.includes("active") || status.includes("operate");
    }).length;
    return {
      totalPartner: uniqueRows.length,
      existingEligible,
      newCount,
      noTouchThisYear,
      activePartner,
    };
  }

  function extractCrmSummaryFromPayload(payload, rows) {
    const candidate =
      getObjectPropertyByNames(payload, ["crmImportSummary", "crm_summary", "crmImportTotals"]) ||
      getObjectPropertyByNames(getObjectPropertyByNames(payload, ["summary"]) || {}, ["crmImport", "crm_import", "CRM_IMPORT"]) ||
      {};
    const fallback = buildCrmSummaryFromRows(rows);
    return {
      totalPartner: asNumber(readAny(candidate, ["totalPartner", "total_partner"])) || fallback.totalPartner,
      existingEligible:
        asNumber(readAny(candidate, ["existingEligible", "existing_eligible", "totalExistingEligible"])) ||
        fallback.existingEligible,
      newCount: asNumber(readAny(candidate, ["newCount", "new_count", "totalNew"])) || fallback.newCount,
      noTouchThisYear:
        asNumber(readAny(candidate, ["noTouchThisYear", "no_touch_this_year", "totalNoTouchThisYear"])) ||
        fallback.noTouchThisYear,
      activePartner: asNumber(readAny(candidate, ["activePartner", "active_partner"])) || fallback.activePartner,
    };
  }

  function normalizePartnerMasterRow(row) {
    const contractDateRaw = readAny(row, ["contractDate", "contract_date", "Contract_Date"]);
    const openingDateRaw = readAny(row, ["openingDate", "opening_date", "Opening_Date"]);
    const rawPartnerType = readAny(row, [
      "partnerTypeLabel",
      "partnerType",
      "partner_type",
      "partner_type_new_existing",
      "new_existed",
      "new_existing",
      "cohort",
      "current_type",
      "type",
      "loai_doi_tac",
      "segment",
    ]);
    const lifecycleFromDates = resolveLifecycleTypeFromDates(contractDateRaw, openingDateRaw);
    const partnerNameRaw = cleanText(
      readAny(row, [
        "partnerName",
        "partner_name",
        "partner",
        "partner_full_name",
        "branchName",
        "branch_name",
        "name",
        "ten_doi_tac",
      ])
    );
    const ownerRaw = cleanText(
      readAny(row, [
        "owner",
        "Owner",
        "pic",
        "account_owner",
        "accountManager",
        "account_manager",
        "am",
        "bdm",
        "rm",
        "sales_owner",
        "franchise_owner",
        "manager",
        "nguoi_phu_trach",
        "nguoi_quan_ly",
      ])
    );
    const resolvedPartnerName = partnerNameRaw || ownerRaw;
    const closingDateRaw = readAny(row, ["closingDate", "closing_date", "Closing_Date"]);
    const partnerKey = cleanText(
      readAny(row, [
        "partnerKey",
        "partner_key",
        "partnerId",
        "partner_id",
        "partnerCode",
        "partner_code",
        "Partner_Code",
        "code",
        "id",
        "ma_doi_tac",
      ]) ||
        safeId(
          readAny(row, [
            "partnerName",
            "partner_name",
            "partner",
            "partner_full_name",
            "branchName",
            "branch_name",
            "name",
            "ten_doi_tac",
          ])
        ) ||
        ""
    ).toUpperCase();
    return {
      partnerKey,
      partnerName: resolvedPartnerName,
      partnerType: normalizePartnerType(
        lifecycleFromDates || rawPartnerType
      ),
      owner: normalizeOwnerValue(ownerRaw, resolvedPartnerName),
      region: cleanText(readAny(row, ["region", "area", "khu_vuc"])),
      status: inferStatusValue(
        cleanText(
        readAny(row, [
          "status",
          "status_active",
          "statusActive",
          "status_text",
          "statusText",
          "partner_status",
          "active",
          "trang_thai",
          "tinh_trang",
        ])
      ),
        closingDateRaw
      ),
    };
  }

  function normalizeOwnerValue(ownerValue, partnerNameValue) {
    const owner = cleanText(ownerValue);
    if (!owner) return "";
    const ownerNorm = normalizeCompareText(owner).replace(/\s+/g, "");
    if (!ownerNorm || ownerNorm === "nan" || ownerNorm === "null") return "";
    const partnerNorm = normalizeCompareText(partnerNameValue).replace(/\s+/g, "");
    if (partnerNorm && ownerNorm === partnerNorm) return "";
    return owner;
  }

  function inferStatusValue(statusValue, closingDateValue) {
    const status = cleanText(statusValue);
    if (status) return status;
    const closingDate = parseFlexibleDate(closingDateValue);
    if (closingDate && !Number.isNaN(closingDate.getTime())) {
      const today = new Date();
      today.setHours(23, 59, 59, 999);
      if (closingDate.getTime() <= today.getTime()) return "Closed";
    }
    return "Active";
  }

  function isPartnerActiveStatus(statusValue) {
    const status = normalizeCompareText(statusValue);
    if (!status) return false;

    if (
      status.includes("closed") ||
      status.includes("dong cua") ||
      status.includes("inactive") ||
      status.includes("ngung") ||
      status.includes("tam dung") ||
      status.includes("terminated") ||
      status.includes("suspend")
    ) {
      return false;
    }

    return (
      status.includes("active") ||
      status.includes("hoat dong") ||
      status.includes("operat") ||
      status.includes("open") ||
      status.includes("mo cua")
    );
  }

  function readAny(row, keys) {
    if (!row || typeof row !== "object") return "";
    for (const key of keys) {
      if (key in row && cleanText(row[key])) return row[key];
    }
    const lowered = {};
    Object.keys(row).forEach((k) => {
      lowered[normalizeCompareText(k)] = row[k];
    });
    for (const key of keys) {
      const val = lowered[normalizeCompareText(key)];
      if (cleanText(val)) return val;
    }
    return "";
  }

  function buildPartnerMasterFromLogs(logs) {
    const map = new Map();
    (logs || []).forEach((log) => {
      const partnerKey = cleanText(log.partnerKey).toUpperCase();
      if (!partnerKey) return;
      map.set(partnerKey, {
        partnerKey,
        partnerName: cleanText(log.partnerName),
        partnerType: normalizePartnerType(log.partnerType),
        owner: cleanText(log.owner),
        region: "",
        status: "Active",
      });
    });
    return Array.from(map.values()).sort((a, b) => cleanText(a.partnerName).localeCompare(cleanText(b.partnerName)));
  }

  function getPartnerMasterRecords() {
    const map = new Map();
    buildPartnerMasterFromLogs(state.partnerCare.logs).forEach((row) => {
      map.set(cleanText(row.partnerKey).toUpperCase(), row);
    });
    (state.partnerCare.partnerMaster || []).forEach((row) => {
      const normalized = normalizePartnerMasterRow(row);
      const key = cleanText(normalized.partnerKey).toUpperCase();
      if (!key) return;
      const previous = map.get(key) || {};
      map.set(key, { ...previous, ...normalized });
    });
    return Array.from(map.values()).sort((a, b) => cleanText(a.partnerName).localeCompare(cleanText(b.partnerName)));
  }

  function getRuleMonthsByYearTarget(targetPerYear) {
    const target = Math.max(0, Math.floor(asNumber(targetPerYear) || 0));
    if (!target) return new Set();
    const step = 12 / target;
    const set = new Set();
    for (let i = 0; i < target; i += 1) {
      const monthIndex = Math.min(11, Math.floor(i * step));
      set.add(monthIndex);
    }
    return set;
  }

  function getYearlyTouchpointTarget(partnerType) {
    return normalizePartnerType(partnerType) === "New" ? 6 : 2;
  }

  function generatePartnerMonthlyPlansForMonth(month) {
    const selectedMonth = normalizeMonth(month);
    const monthIndex = MONTHS.indexOf(selectedMonth);
    if (monthIndex < 0) return 0;

    const partners = getPartnerMasterRecords();
    if (!partners.length) return 0;

    const currentMonthRows = state.partnerCare.partnerMonthlyPlans.filter((row) => cleanText(row.month) === selectedMonth);
    const existingPartnerSet = new Set(
      currentMonthRows
        .filter((row) => isActivePartnerPlanRow(row))
        .map((row) => cleanText(row.partnerKey).toUpperCase())
    );
    const generatedRows = [];

    partners.forEach((partner) => {
      const partnerKey = cleanText(partner.partnerKey).toUpperCase();
      if (!partnerKey) return;
      if (existingPartnerSet.has(partnerKey)) return;
      const yearlyTarget = getYearlyTouchpointTarget(partner.partnerType);
      const ruleMonths = getRuleMonthsByYearTarget(yearlyTarget);
      if (!ruleMonths.has(monthIndex)) return;
      generatedRows.push({
        id: uid("PCP"),
        month: selectedMonth,
        careDay: null,
        dayDone: null,
        partnerKey,
        partnerName: cleanText(partner.partnerName),
        partnerType: normalizePartnerType(partner.partnerType),
        touchpointType: "",
        channel: "Call/message",
        owner: cleanText(partner.owner),
        targetTouchpoints: 1,
        status: "Planned",
        result: "",
        response: "",
        nextAction: "",
      });
    });

    if (!generatedRows.length) return 0;
    state.partnerCare.partnerMonthlyPlans = [...state.partnerCare.partnerMonthlyPlans, ...generatedRows];
    return generatedRows.length;
  }

  function normalizePartnerPayload(payload) {
    let rows = [];
    if (Array.isArray(payload)) {
      rows = payload;
    } else if (payload && typeof payload === "object") {
      if (Array.isArray(payload.partners)) rows = payload.partners;
      else if (Array.isArray(payload.data)) rows = payload.data;
      else if (Array.isArray(payload.rows)) rows = payload.rows;
      else if (Array.isArray(payload.values) && Array.isArray(payload.values[0])) {
        const headers = payload.values[0].map((h) => cleanText(h));
        rows = payload.values.slice(1).map((vals) => {
          const obj = {};
          headers.forEach((h, idx) => {
            obj[h] = vals[idx];
          });
          return obj;
        });
      }
    }

    const map = new Map();
    rows.forEach((row) => {
      const normalized = normalizePartnerMasterRow(row);
      const key = cleanText(normalized.partnerKey).toUpperCase();
      if (!key) return;
      map.set(key, normalized);
    });
    return Array.from(map.values()).sort((a, b) => cleanText(a.partnerName).localeCompare(cleanText(b.partnerName)));
  }

  async function syncPartnerMasterFromGoogleSheets() {
    const endpointInput = document.getElementById("pc-sync-endpoint");
    const endpoint = cleanText(endpointInput?.value || state.partnerCare.sync.endpoint);
    if (!endpoint) {
      setSyncStatus("Cần nhập endpoint Google Sheets để sync đối tác");
      return;
    }

    state.partnerCare.sync.endpoint = endpoint;
    setSyncStatus("Đang sync đối tác từ Google Sheets...");
    try {
      const response = await fetch(endpoint, { method: "GET", cache: "no-store" });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const payload = await response.json();
      const crmRawRows = extractCrmImportRowsFromPayload(payload);
      const crmRows = crmRawRows.map((row, idx) => normalizeCrmImportRow(row, idx)).filter((row) => cleanText(row.partnerCode));
      const crmSummary = extractCrmSummaryFromPayload(payload, crmRows);
      const partners = normalizePartnerPayload(crmRows.length ? crmRows : payload);
      if (!partners.length) throw new Error("Không tìm thấy dữ liệu đối tác hợp lệ");

      state.partnerCare.crmImportRows = crmRows;
      state.partnerCare.crmSummary = crmSummary;
      state.partnerCare.partnerMaster = partners;
      state.partnerCare.sync.lastSyncAt = new Date().toISOString();
      state.partnerCare.sync.lastCount = partners.length;

      persistState(`Đã sync ${partners.length} đối tác từ Google Sheets`);
      renderAll();
    } catch (error) {
      setSyncStatus(`Sync đối tác thất bại: ${cleanText(error.message) || "Unknown error"}`);
    }
  }

  function aggregateByMonths(months) {
    const set = new Set(months);
    let targetWeight = 0;
    let achievedWeight = 0;
    let recordCount = 0;
    let missingActualCount = 0;

    state.importHub.monthlyKpis.forEach((row) => {
      if (!set.has(row.month)) return;
      const target = asNumber(row.target);
      const actual = asNumber(row.actual);
      if (target === null || target <= 0) return;
      recordCount += 1;
      targetWeight += target;
      if (actual === null) {
        missingActualCount += 1;
      } else {
        achievedWeight += actual;
      }
    });

    return {
      targetWeight,
      achievedWeight,
      rate: targetWeight > 0 ? achievedWeight / targetWeight : 0,
      recordCount,
      missingActualCount,
      coverageLabel: recordCount > 0 ? `${recordCount - missingActualCount}/${recordCount} KPI có actual` : "0/0",
    };
  }

  function aggregatePillarByMonths(pillarKey, months) {
    const set = new Set(months);
    let targetWeight = 0;
    let achievedWeight = 0;
    let recordCount = 0;
    let missingActualCount = 0;

    state.importHub.monthlyKpis.forEach((row) => {
      if (row.pillarKey !== pillarKey) return;
      if (!set.has(row.month)) return;
      const target = asNumber(row.target);
      const actual = asNumber(row.actual);
      if (target === null || target <= 0) return;
      recordCount += 1;
      targetWeight += target;
      if (actual === null) {
        missingActualCount += 1;
      } else {
        achievedWeight += actual;
      }
    });

    return {
      targetWeight,
      achievedWeight,
      rate: targetWeight > 0 ? achievedWeight / targetWeight : 0,
      recordCount,
      missingActualCount,
      coverageLabel: recordCount > 0 ? `${recordCount - missingActualCount}/${recordCount} KPI có actual` : "0/0",
    };
  }

  function getPillarRatesByScope(scope, key) {
    if (scope === "month") {
      return {
        HR: aggregatePillarByMonths("HR", [key]).rate,
        VN: aggregatePillarByMonths("VN", [key]).rate,
        INTL: aggregatePillarByMonths("INTL", [key]).rate,
      };
    }
    if (scope === "months") {
      const months = Array.isArray(key) ? key : MONTHS;
      return {
        HR: aggregatePillarByMonths("HR", months).rate,
        VN: aggregatePillarByMonths("VN", months).rate,
        INTL: aggregatePillarByMonths("INTL", months).rate,
      };
    }
    const months = QUARTER_MONTHS[key] || [];
    return {
      HR: aggregatePillarByMonths("HR", months).rate,
      VN: aggregatePillarByMonths("VN", months).rate,
      INTL: aggregatePillarByMonths("INTL", months).rate,
    };
  }

  function uniqueRowsById(rows) {
    const map = new Map();
    rows.forEach((row) => {
      const id = cleanText(row?.id);
      if (!id) return;
      if (!map.has(id)) map.set(id, row);
    });
    return Array.from(map.values());
  }

  function isKpiMatchingAnnualOkr(kpi, annualOkr) {
    const annualCode = cleanText(annualOkr?.code).toUpperCase();
    const annualTitle = cleanText(annualOkr?.title).toLowerCase();
    const kpiCode = cleanText(kpi?.kpiCode).toUpperCase();
    const kpiName = cleanText(kpi?.kpiName).toLowerCase();
    if (annualCode && kpiCode && annualCode === kpiCode) return true;
    if (annualTitle && kpiName && annualTitle === kpiName) return true;
    return false;
  }

  function normalizeCompareText(value) {
    return cleanText(value)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, " ")
      .trim();
  }

  function isMonthlyOkrMatchingAnnualOkr(monthlyOkr, annualOkr) {
    const annualTitle = normalizeCompareText(annualOkr?.title);
    const annualCode = normalizeCompareText(annualOkr?.code);
    const monthlyTitle = normalizeCompareText(monthlyOkr?.title);
    if (!monthlyTitle) return false;
    if (annualTitle && (monthlyTitle.includes(annualTitle) || annualTitle.includes(monthlyTitle))) return true;
    if (annualCode && monthlyTitle.includes(annualCode)) return true;
    return false;
  }

  function getKpisForAnnualOkr(annualOkr, months) {
    const monthsSet = new Set(months);
    const quarterlyIds = new Set(
      state.importHub.quarterlyOkrs.filter((row) => row.linkedAnnualOkrId === annualOkr.id).map((row) => row.id)
    );
    const linkedMonthlyIds = new Set(
      state.importHub.monthlyOkrs
        .filter((row) => quarterlyIds.has(row.linkedQuarterlyOkrId) && monthsSet.has(row.month))
        .map((row) => row.id)
    );
    const linkedKpis = state.importHub.monthlyKpis.filter(
      (row) => linkedMonthlyIds.has(row.monthlyOkrId) && monthsSet.has(row.month)
    );
    const fallbackByIdentity = state.importHub.monthlyKpis.filter(
      (row) => row.pillarKey === annualOkr.pillarKey && monthsSet.has(row.month) && isKpiMatchingAnnualOkr(row, annualOkr)
    );
    return uniqueRowsById([...linkedKpis, ...fallbackByIdentity]);
  }

  function getMonthlyOkrsForAnnualOkr(annualOkr, months) {
    const monthsSet = new Set(months);
    const quarterlyIds = new Set(
      state.importHub.quarterlyOkrs.filter((row) => row.linkedAnnualOkrId === annualOkr.id).map((row) => row.id)
    );
    const linkedMonthly = state.importHub.monthlyOkrs.filter(
      (row) => quarterlyIds.has(row.linkedQuarterlyOkrId) && monthsSet.has(row.month)
    );
    if (linkedMonthly.length) return uniqueRowsById(linkedMonthly);

    const byKpiIdentityMonthlyIds = new Set(
      state.importHub.monthlyKpis
        .filter(
          (row) => row.pillarKey === annualOkr.pillarKey && monthsSet.has(row.month) && isKpiMatchingAnnualOkr(row, annualOkr)
        )
        .map((row) => cleanText(row.monthlyOkrId))
    );
    if (byKpiIdentityMonthlyIds.size) {
      return uniqueRowsById(
        state.importHub.monthlyOkrs.filter((row) => byKpiIdentityMonthlyIds.has(cleanText(row.id)) && monthsSet.has(row.month))
      );
    }

    const byTitleMatch = state.importHub.monthlyOkrs.filter(
      (row) =>
        row.pillarKey === annualOkr.pillarKey && monthsSet.has(row.month) && isMonthlyOkrMatchingAnnualOkr(row, annualOkr)
    );
    if (byTitleMatch.length) return uniqueRowsById(byTitleMatch);

    const annualInPillar = state.importHub.annualOkrs.filter((row) => row.pillarKey === annualOkr.pillarKey);
    if (annualInPillar.length === 1) {
      return uniqueRowsById(
        state.importHub.monthlyOkrs.filter((row) => row.pillarKey === annualOkr.pillarKey && monthsSet.has(row.month))
      );
    }

    return [];
  }

  function getKpisForQuarterlyOkr(quarterlyOkr, months) {
    const monthsSet = new Set(months);
    const directMonthlyIds = new Set(
      state.importHub.monthlyOkrs
        .filter((row) => row.linkedQuarterlyOkrId === quarterlyOkr.id && monthsSet.has(row.month))
        .map((row) => row.id)
    );
    const samePillarQuarterOkrs = state.importHub.quarterlyOkrs.filter(
      (row) => row.quarter === quarterlyOkr.quarter && row.pillarKey === quarterlyOkr.pillarKey
    );
    if (samePillarQuarterOkrs.length === 1) {
      state.importHub.monthlyOkrs
        .filter((row) => row.pillarKey === quarterlyOkr.pillarKey && monthsSet.has(row.month))
        .forEach((row) => directMonthlyIds.add(row.id));
    }
    const linkedKpis = state.importHub.monthlyKpis.filter(
      (row) => directMonthlyIds.has(row.monthlyOkrId) && monthsSet.has(row.month)
    );
    const linkedAnnual = state.importHub.annualOkrs.find((row) => row.id === quarterlyOkr.linkedAnnualOkrId);
    const fallbackByAnnual = linkedAnnual ? getKpisForAnnualOkr(linkedAnnual, months) : [];
    return uniqueRowsById([...linkedKpis, ...fallbackByAnnual]);
  }

  function buildBestWorstText(rateMap, labelPrefix) {
    const entries = Object.entries(rateMap).map(([key, rate]) => ({
      key,
      label: PILLARS[key].label,
      rate: rate || 0,
    }));
    entries.sort((a, b) => b.rate - a.rate);
    const best = entries[0];
    const worst = entries[entries.length - 1];
    return `${labelPrefix}: Pillar mạnh nhất là ${best.label} (${formatPercent(best.rate)}), cần ưu tiên cải thiện ${worst.label} (${formatPercent(
      worst.rate
    )}).`;
  }

  function forecastAllPillars(uptoMonth) {
    const rates = Object.keys(PILLARS).map((pillarKey) => forecastByPillar(pillarKey, uptoMonth));
    if (!rates.length) return 0;
    return rates.reduce((sum, rate) => sum + rate, 0) / rates.length;
  }

  function forecastByPillar(pillarKey, uptoMonth) {
    const selectedMonths = getMonthsUpToUiSelection(uptoMonth);
    const elapsedMonths = selectedMonths.length || 1;
    const annualTarget = state.importHub.annualOkrs
      .filter((row) => row.pillarKey === pillarKey)
      .reduce((sum, row) => sum + (asNumber(row.target) || 0), 0);

    const ytdMonths = selectedMonths;
    const actualYtd = state.importHub.monthlyKpis
      .filter((row) => row.pillarKey === pillarKey && ytdMonths.includes(row.month))
      .reduce((sum, row) => sum + (asNumber(row.actual) || 0), 0);

    if (annualTarget > 0 && elapsedMonths > 0) {
      if (uptoMonth === ALL_MONTH_VALUE) return actualYtd / annualTarget;
      const forecastAnnualActual = (actualYtd / elapsedMonths) * 12;
      return forecastAnnualActual / annualTarget;
    }

    const targetYtd = state.importHub.monthlyKpis
      .filter((row) => row.pillarKey === pillarKey && ytdMonths.includes(row.month))
      .reduce((sum, row) => sum + (asNumber(row.target) || 0), 0);

    return targetYtd > 0 ? actualYtd / targetYtd : 0;
  }

  function getMonthlyOkrs(pillarKey, month) {
    const months = getScopeMonthsByUiSelection(month);
    return state.importHub.monthlyOkrs.filter((row) => row.pillarKey === pillarKey && months.includes(row.month));
  }

  function getMonthlyKpis(pillarKey, month) {
    const months = getScopeMonthsByUiSelection(month);
    return state.importHub.monthlyKpis.filter((row) => row.pillarKey === pillarKey && months.includes(row.month));
  }

  function getScopeMonthsByUiSelection(monthValue) {
    return monthValue === ALL_MONTH_VALUE ? [...MONTHS] : [monthValue];
  }

  function getMonthsUpToUiSelection(monthValue) {
    if (monthValue === ALL_MONTH_VALUE) return [...MONTHS];
    const monthIndex = MONTHS.indexOf(monthValue);
    if (monthIndex < 0) return [...MONTHS];
    return MONTHS.slice(0, monthIndex + 1);
  }

  function getUiMonthLabel(monthValue) {
    return monthValue === ALL_MONTH_VALUE ? "Tất cả" : monthValue;
  }

  function populateMonthSelect(id, options = {}) {
    const { includeAll = false } = options;
    const select = document.getElementById(id);
    if (!select) return;
    select.innerHTML = "";
    if (includeAll) {
      select.appendChild(makeOption(ALL_MONTH_VALUE, "Tất cả"));
    }
    MONTHS.forEach((month) => select.appendChild(makeOption(month, month)));
  }

  function populateQuarterSelect(id) {
    const select = document.getElementById(id);
    if (!select) return;
    select.innerHTML = "";
    ["Q1", "Q2", "Q3", "Q4"].forEach((q) => select.appendChild(makeOption(q, q)));
  }

  function populatePillarSelect(id) {
    const select = document.getElementById(id);
    if (!select) return;
    select.innerHTML = "";
    Object.keys(PILLARS).forEach((pillarKey) => {
      select.appendChild(makeOption(pillarKey, `${pillarKey} - ${PILLARS[pillarKey].label}`));
    });
  }

  function inferPillarFromLabel(label, code) {
    const txt = `${cleanText(label)} ${cleanText(code)}`.toLowerCase();
    if (txt.includes("international") || txt.startsWith("int")) return "INTL";
    if (txt.includes("national franchise") || txt.startsWith("vn")) return "VN";
    return "HR";
  }

  function normalizePillar(value) {
    const txt = cleanText(value).toUpperCase();
    if (txt === "HR" || txt === "HUMAN RESOURCE") return "HR";
    if (txt === "VN" || txt === "NATIONAL FRANCHISE DEVELOPMENT") return "VN";
    if (txt === "INTL" || txt === "INT" || txt === "INTERNATIONAL FRANCHISE DEVELOPMENT") return "INTL";
    if (txt.includes("INTERNATIONAL")) return "INTL";
    if (txt.includes("NATIONAL") || txt.includes("FRANCHISE")) return "VN";
    return "HR";
  }

  function normalizeMonth(value) {
    const txt = cleanText(value);
    if (MONTHS.includes(txt)) return txt;
    const normalized = txt.slice(0, 3);
    if (MONTHS.includes(normalized)) return normalized;
    return "Jan";
  }

  function normalizeQuarter(value) {
    const txt = cleanText(value).toUpperCase();
    if (!txt) return "Q1";
    if (txt.startsWith("Q")) return txt;
    if (txt === "I") return "Q1";
    if (txt === "II") return "Q2";
    if (txt === "III") return "Q3";
    if (txt === "IV") return "Q4";
    return "Q1";
  }

  function quarterOf(month) {
    const idx = MONTHS.indexOf(month);
    if (idx <= 2) return "Q1";
    if (idx <= 5) return "Q2";
    if (idx <= 8) return "Q3";
    return "Q4";
  }

  function QUARTER_ORDER(quarter) {
    return ["Q1", "Q2", "Q3", "Q4"].indexOf(normalizeQuarter(quarter));
  }

  function classifyRate(rate) {
    const good = state.settings.thresholds.good || 0.9;
    const watch = state.settings.thresholds.watch || 0.75;
    if (rate === null || Number.isNaN(rate)) return { className: "watch", label: "Chưa đủ dữ liệu" };
    if (rate >= good) return { className: "good", label: "Đạt tốt" };
    if (rate >= watch) return { className: "watch", label: "Cần theo dõi" };
    return { className: "bad", label: "Cảnh báo" };
  }

  function isLikelyUrl(value) {
    const txt = cleanText(value);
    return /^https?:\/\/\S+$/i.test(txt);
  }

  function renderKpiResultPreview(value) {
    const txt = cleanText(value);
    if (!txt) return "";
    if (isLikelyUrl(txt)) {
      return `<span class="kpi-result-preview"><a href="${safeText(txt)}" target="_blank" rel="noopener noreferrer">Mở link</a></span>`;
    }
    return `<span class="kpi-result-preview">${safeText(truncateText(txt, 80))}</span>`;
  }

  function formatMetric(value, unit) {
    const number = asNumber(value);
    if (number === null) return "-";
    const normalizedUnit = cleanText(unit).toLowerCase();
    if (normalizedUnit.includes("vnd") || normalizedUnit.includes("đ")) return `${formatNumber(number, 0)} đ`;
    if ((unit || "").includes("%")) return `${formatNumber(number, 1)}%`;
    return formatNumber(number, number >= 100 ? 0 : 2);
  }

  function formatPoints(value) {
    return `${formatNumber(value || 0, 2)} pts`;
  }

  function formatPercent(rate) {
    const pct = (rate || 0) * 100;
    return `${formatNumber(pct, 1)}%`;
  }

  function formatNumber(value, fractionDigits) {
    const num = asNumber(value);
    if (num === null) return "-";
    return new Intl.NumberFormat("vi-VN", {
      minimumFractionDigits: fractionDigits,
      maximumFractionDigits: fractionDigits,
    }).format(num);
  }

  function truncateText(text, length) {
    if (!text) return "";
    return text.length > length ? `${text.slice(0, length)}...` : text;
  }

  function formatDateTime(value) {
    const txt = cleanText(value);
    if (!txt) return "-";
    const date = new Date(txt);
    if (Number.isNaN(date.getTime())) return txt;
    return date.toLocaleString("vi-VN", {
      hour12: false,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
    });
  }

  function formatSheetDate(value) {
    const asNum = asNumber(value);
    if (asNum !== null && asNum > 20000 && asNum < 90000) {
      const epoch = Math.round((asNum - 25569) * 86400 * 1000);
      const date = new Date(epoch);
      if (!Number.isNaN(date.getTime())) return date.toLocaleDateString("vi-VN");
    }
    const txt = cleanText(value);
    if (!txt) return "-";
    const date = new Date(txt);
    if (!Number.isNaN(date.getTime())) return date.toLocaleDateString("vi-VN");
    return txt;
  }

  function parseFlexibleDate(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;

    const num = asNumber(value);
    if (num !== null && num > 20000 && num < 90000) {
      const epoch = Math.round((num - 25569) * 86400 * 1000);
      const dateFromSerial = new Date(epoch);
      if (!Number.isNaN(dateFromSerial.getTime())) return dateFromSerial;
    }

    const txt = cleanText(value);
    if (!txt) return null;

    let match = txt.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
    if (match) {
      const day = Number(match[1]);
      const month = Number(match[2]);
      const year = Number(match[3]);
      const date = new Date(year, month - 1, day);
      if (date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day) return date;
    }

    match = txt.match(/^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})$/);
    if (match) {
      const year = Number(match[1]);
      const month = Number(match[2]);
      const day = Number(match[3]);
      const date = new Date(year, month - 1, day);
      if (date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day) return date;
    }

    const parsed = new Date(txt);
    if (!Number.isNaN(parsed.getTime())) return parsed;
    return null;
  }

  function normalizeCareDay(value) {
    const num = asNumber(parseNumeric(value));
    if (num === null) return null;
    const day = Math.floor(num);
    if (day < 1 || day > 31) return null;
    return day;
  }

  function parseNumeric(value) {
    if (value === null || value === undefined) return null;
    const txt = String(value).trim().replace(/,/g, "");
    if (!txt) return null;
    const num = Number(txt);
    return Number.isNaN(num) ? null : num;
  }

  function asNumber(value) {
    if (typeof value === "number") return Number.isFinite(value) ? value : null;
    if (typeof value === "string") {
      const txt = value.trim();
      if (!txt || txt === "#REF!" || txt === "#DIV/0!" || txt === "#N/A") return null;
      const num = Number(txt.replace(/,/g, ""));
      return Number.isFinite(num) ? num : null;
    }
    return null;
  }

  function cleanText(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
  }

  function safeId(value) {
    return cleanText(value)
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "");
  }

  function uid(prefix) {
    return `${prefix}-${Date.now()}-${Math.floor(Math.random() * 10000)}`;
  }

  function safeText(value) {
    return escapeHtml(cleanText(value));
  }

  function escapeHtml(text) {
    return text
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }

  function makeOption(value, label) {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = label;
    return option;
  }

  function setText(id, value) {
    const el = document.getElementById(id);
    if (!el) return;
    el.textContent = String(value);
  }

  function renderInsightList(id, items) {
    const el = document.getElementById(id);
    if (!el) return;
    el.innerHTML = "";
    items.forEach((item) => {
      const li = document.createElement("li");
      li.textContent = item;
      el.appendChild(li);
    });
  }

  function deepClone(obj) {
    return JSON.parse(JSON.stringify(obj || {}));
  }

  function getRuntimeCloudDefaults() {
    const runtime = window.__HRBD_CLOUD_CONFIG__ || {};
    const url = cleanText(runtime.url || runtime.supabaseUrl);
    const anonKey = cleanText(runtime.anonKey || runtime.supabaseAnonKey);
    if (!url || !anonKey) return null;

    const table = cleanText(runtime.table) || "dashboard_state";
    const stateKey = cleanText(runtime.stateKey) || "global";
    const autoSync = typeof runtime.autoSync === "boolean" ? runtime.autoSync : true;

    return {
      url,
      anonKey,
      table,
      stateKey,
      autoSync,
    };
  }

  function defaultCloudConfig() {
    return {
      provider: "supabase",
      url: "",
      anonKey: "",
      table: "dashboard_state",
      stateKey: "global",
      autoSync: false,
      lastPullAt: "",
      lastPushAt: "",
      lastError: "",
    };
  }

  function loadCloudConfig() {
    const runtimeDefaults = getRuntimeCloudDefaults();
    try {
      const raw = localStorage.getItem(CLOUD_CFG_KEY);
      if (!raw) {
        return {
          ...defaultCloudConfig(),
          ...(runtimeDefaults || {}),
        };
      }
      const parsed = JSON.parse(raw);
      const merged = {
        ...defaultCloudConfig(),
        ...(runtimeDefaults || {}),
        ...(parsed || {}),
      };
      if (runtimeDefaults) {
        if (!cleanText(merged.url)) merged.url = runtimeDefaults.url;
        if (!cleanText(merged.anonKey)) merged.anonKey = runtimeDefaults.anonKey;
        if (!cleanText(merged.table)) merged.table = runtimeDefaults.table;
        if (!cleanText(merged.stateKey)) merged.stateKey = runtimeDefaults.stateKey;
        if (typeof parsed?.autoSync !== "boolean") merged.autoSync = runtimeDefaults.autoSync;
      }
      return merged;
    } catch (error) {
      return {
        ...defaultCloudConfig(),
        ...(runtimeDefaults || {}),
      };
    }
  }

  function saveCloudConfig(statusMessage) {
    try {
      localStorage.setItem(CLOUD_CFG_KEY, JSON.stringify(cloudConfig));
      if (statusMessage) setSyncStatus(statusMessage);
    } catch (error) {
      setSyncStatus("Không thể lưu cấu hình Supabase");
    }
  }

  function invalidateSupabaseClient() {
    supabaseClient = null;
  }

  function isCloudSyncConfigured() {
    return !!cleanText(cloudConfig.url) && !!cleanText(cloudConfig.anonKey);
  }

  function getCloudTableName() {
    return cleanText(cloudConfig.table) || "dashboard_state";
  }

  function getCloudStateKey() {
    return cleanText(cloudConfig.stateKey) || "global";
  }

  function getSupabaseClient() {
    if (!isCloudSyncConfigured()) return null;
    const supabaseLib = window.supabase;
    if (!supabaseLib || typeof supabaseLib.createClient !== "function") {
      throw new Error("Thiếu thư viện Supabase client");
    }
    if (
      supabaseClient &&
      cleanText(supabaseClient.__url) === cleanText(cloudConfig.url) &&
      cleanText(supabaseClient.__key) === cleanText(cloudConfig.anonKey)
    ) {
      return supabaseClient;
    }
    supabaseClient = supabaseLib.createClient(cloudConfig.url, cloudConfig.anonKey, {
      auth: {
        persistSession: false,
        autoRefreshToken: false,
      },
    });
    supabaseClient.__url = cleanText(cloudConfig.url);
    supabaseClient.__key = cleanText(cloudConfig.anonKey);
    return supabaseClient;
  }

  async function testSupabaseConnection() {
    if (!isCloudSyncConfigured()) {
      setSyncStatus("Cần nhập Supabase URL + Anon Key");
      return false;
    }
    try {
      const client = getSupabaseClient();
      const table = getCloudTableName();
      const { error } = await client.from(table).select("id", { head: true, count: "exact" }).limit(1);
      if (error) throw error;
      cloudConfig.lastError = "";
      saveCloudConfig();
      setSyncStatus(`Kết nối Supabase OK (table: ${table})`);
      return true;
    } catch (error) {
      const message = cleanText(error?.message) || "Unknown error";
      cloudConfig.lastError = message;
      saveCloudConfig();
      setSyncStatus(`Kết nối Supabase thất bại: ${message}`);
      return false;
    }
  }

  async function pullStateFromCloud() {
    if (!isCloudSyncConfigured()) {
      setSyncStatus("Cần cấu hình Supabase trước khi tải cloud");
      return false;
    }
    try {
      const client = getSupabaseClient();
      const table = getCloudTableName();
      const stateKey = getCloudStateKey();
      setSyncStatus("Đang tải dữ liệu từ Supabase...");
      const { data, error } = await client
        .from(table)
        .select("payload, updated_at")
        .eq("id", stateKey)
        .maybeSingle();
      if (error) throw error;

      if (!data || !data.payload || typeof data.payload !== "object") {
        setSyncStatus("Cloud chưa có dữ liệu state để tải");
        return false;
      }

      suppressCloudPush = true;
      state = deepClone(data.payload);
      ensureStateShape();
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
      } catch (errorLocal) {
        // no-op
      }
      suppressCloudPush = false;

      cloudConfig.lastPullAt = new Date().toISOString();
      cloudConfig.lastError = "";
      saveCloudConfig();
      renderAll();
      setSyncStatus(`Đã tải cloud state (${formatDateTime(data.updated_at || cloudConfig.lastPullAt)})`);
      return true;
    } catch (error) {
      suppressCloudPush = false;
      const message = cleanText(error?.message) || "Unknown error";
      cloudConfig.lastError = message;
      saveCloudConfig();
      setSyncStatus(`Tải cloud thất bại: ${message}`);
      return false;
    }
  }

  async function pushStateToCloud(reason = "manual") {
    if (!isCloudSyncConfigured()) return false;
    if (suppressCloudPush) return false;

    if (cloudPushInFlight) {
      queuedPushReason = reason;
      return false;
    }

    cloudPushInFlight = true;
    try {
      const client = getSupabaseClient();
      const table = getCloudTableName();
      const stateKey = getCloudStateKey();
      const payload = deepClone(state);
      const nowIso = new Date().toISOString();
      const { error } = await client.from(table).upsert(
        {
          id: stateKey,
          payload,
          updated_at: nowIso,
        },
        {
          onConflict: "id",
        }
      );
      if (error) throw error;

      cloudConfig.lastPushAt = nowIso;
      cloudConfig.lastError = "";
      saveCloudConfig();
      if (reason === "auto") {
        setSyncStatus("Đã auto-sync dữ liệu lên Supabase");
      } else {
        setSyncStatus("Đã đẩy dữ liệu lên Supabase");
      }
      return true;
    } catch (error) {
      const message = cleanText(error?.message) || "Unknown error";
      cloudConfig.lastError = message;
      saveCloudConfig();
      setSyncStatus(`Đẩy dữ liệu thất bại: ${message}`);
      return false;
    } finally {
      cloudPushInFlight = false;
      if (queuedPushReason) {
        const nextReason = queuedPushReason;
        queuedPushReason = "";
        setTimeout(() => {
          pushStateToCloud(nextReason);
        }, 0);
      }
    }
  }

  function scheduleCloudAutoPush() {
    if (!cloudConfig.autoSync || !isCloudSyncConfigured() || suppressCloudPush) return;
    if (cloudPushTimer) clearTimeout(cloudPushTimer);
    cloudPushTimer = setTimeout(() => {
      cloudPushTimer = null;
      pushStateToCloud("auto");
    }, 900);
  }

  async function tryCloudBootstrap() {
    if (!cloudConfig.autoSync || !isCloudSyncConfigured()) return;
    await pullStateFromCloud();
  }

  function persistState(statusMessage) {
    try {
      state.version = (state.version || 0) + 1;
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
      scheduleCloudAutoPush();
      if (statusMessage) setSyncStatus(statusMessage);
    } catch (error) {
      setSyncStatus("Không thể lưu localStorage");
    }
  }

  function setSyncStatus(message) {
    const el = document.getElementById("sync-status");
    if (!el) return;
    const now = new Date();
    el.textContent = `${message} • ${now.toLocaleTimeString("vi-VN", {
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
    })}`;
  }

  function loadState(fallback) {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) return JSON.parse(raw);
      const legacyRaw = localStorage.getItem(LEGACY_STORAGE_KEY);
      if (legacyRaw) return JSON.parse(legacyRaw);
      return deepClone(fallback);
    } catch (error) {
      return deepClone(fallback);
    }
  }

  function saveUiState() {
    try {
      localStorage.setItem(UI_KEY, JSON.stringify(uiState));
    } catch (error) {
      // no-op
    }
  }

  function loadUiState() {
    const fallback = { currentPage: "dashboard", month: "Jan", quarter: "Q1" };
    try {
      const raw = localStorage.getItem(UI_KEY);
      if (!raw) return fallback;
      return { ...fallback, ...JSON.parse(raw) };
    } catch (error) {
      return fallback;
    }
  }
})();
