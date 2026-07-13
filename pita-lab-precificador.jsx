import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import {
  ComposedChart,
  Bar,
  Cell,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
} from "recharts";
import {
  Settings2,
  Save,
  Trash2,
  Package,
  Gauge,
  Wrench,
  TrendingUp,
  PawPrint,
  Boxes,
  Plus,
  X,
  Warehouse,
  DollarSign,
  ClipboardList,
  Pencil,
  RefreshCw,
  Download,
  Upload,
  History,
  Wallet,
} from "lucide-react";

const currency = (n) =>
  (isFinite(n) ? n : 0).toLocaleString("pt-BR", { style: "currency", currency: "BRL" });

// Converte "hh:mm" (ou apenas minutos, para compatibilidade) em total de minutos
const hhmmToMinutes = (str) => {
  if (str === undefined || str === null || str === "") return 0;
  const match = String(str).trim().match(/^(\d{1,4}):([0-5]?\d)$/);
  if (match) {
    const h = parseInt(match[1], 10) || 0;
    const m = parseInt(match[2], 10) || 0;
    return h * 60 + m;
  }
  const n = parseFloat(str);
  return isFinite(n) ? n : 0;
};

// Converte total de minutos em "hh:mm"
const minutesToHHMM = (totalMinutes) => {
  const total = Math.max(0, Math.round(parseFloat(totalMinutes) || 0));
  const h = Math.floor(total / 60);
  const m = total % 60;
  return `${h}:${String(m).padStart(2, "0")}`;
};

const DEFAULT_SETTINGS = {
  printerWattage: 200,
  energyPricePerKwh: 0.75,
  printerValue: 2500,
  printerLifetimeHours: 8000,
  laborRatePerHour: 30,
  failureRatePercent: 10,
};

const SEGMENT_COLORS = {
  material: "#2E9BF0",
  energy: "#1C6FB8",
  depreciation: "#33B7A6",
  labor: "#E8A33D",
  insumos: "#B075E5",
  failures: "#F2545B",
};

const SEGMENT_LABELS = {
  material: "Material",
  energy: "Energia",
  depreciation: "Depreciação",
  labor: "Mão de obra",
  insumos: "Insumos",
  failures: "Índice de falhas",
};

const cardStyle = { background: "#171D26", border: "1px solid #2A3341" };
const muted = { color: "#8B98AA" };

export default function PitaLabPricer() {
  const [activeTab, setActiveTab] = useState("orcamento");
  const [settings, setSettings] = useState(DEFAULT_SETTINGS);
  const [quotes, setQuotes] = useState([]);
  const [insumosCatalog, setInsumosCatalog] = useState([]);
  const [filamentsCatalog, setFilamentsCatalog] = useState([]);
  const [stockItems, setStockItems] = useState([]);
  const [salesHistory, setSalesHistory] = useState([]);
  const [productionLog, setProductionLog] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [saveStatus, setSaveStatus] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [legacyEditWarning, setLegacyEditWarning] = useState(false);
  const [originalEditName, setOriginalEditName] = useState("");
  const [copyPromptOpen, setCopyPromptOpen] = useState(false);
  const [copyNameInput, setCopyNameInput] = useState("");

  const [newInsumo, setNewInsumo] = useState({ name: "", totalPaid: "", quantity: "", unit: "unidade" });
  const [editingInsumoId, setEditingInsumoId] = useState(null);
  const [insumoToAdd, setInsumoToAdd] = useState({ id: "", quantity: "1" });

  const [newFilament, setNewFilament] = useState({ brand: "", color: "", costPerKg: "", stockGrams: "1000" });
  const [editingFilamentId, setEditingFilamentId] = useState(null);
  const [filamentToAdd, setFilamentToAdd] = useState({ id: "", grams: "" });

  const [newStock, setNewStock] = useState({
    name: "",
    category: "",
    material: "",
    printed: "",
    unitCost: "",
    marginPercent: "50",
    unitPrice: "",
    usedInsumos: [],
    usedFilaments: [],
  });
  const [sellingItemId, setSellingItemId] = useState(null);
  const [saleQuantity, setSaleQuantity] = useState("");
  const [saleUnitPrice, setSaleUnitPrice] = useState("");
  const [editingSaleId, setEditingSaleId] = useState(null);
  const [editSaleForm, setEditSaleForm] = useState({ quantity: "", unitPrice: "" });
  const [editingStockQtyId, setEditingStockQtyId] = useState(null);
  const [editStockQtyValue, setEditStockQtyValue] = useState("");
  const [materialSubTab, setMaterialSubTab] = useState("filamentos");
  const [showSettingsModal, setShowSettingsModal] = useState(false);
  const [importStatus, setImportStatus] = useState("");
  const [catalogExportStatus, setCatalogExportStatus] = useState("");
  const fileInputRef = useRef(null);

  const [form, setForm] = useState({
    name: "",
    category: "",
    material: "",
    printMinutes: "",
    laborMinutes: "",
    piecesPerBatch: "1",
    failureRateOverride: "",
    usedInsumos: [],
    usedFilaments: [],
  });

  useEffect(() => {
    (async () => {
      try {
        const s = await window.storage.get("pita-lab-settings");
        if (s?.value) setSettings({ ...DEFAULT_SETTINGS, ...JSON.parse(s.value) });
      } catch (e) {
        /* no settings saved yet, use defaults */
      }
      try {
        const q = await window.storage.get("pita-lab-quotes");
        if (q?.value) setQuotes(JSON.parse(q.value));
      } catch (e) {
        /* no history yet */
      }
      try {
        const i = await window.storage.get("pita-lab-insumos");
        if (i?.value) setInsumosCatalog(JSON.parse(i.value));
      } catch (e) {
        /* no insumos cadastrados ainda */
      }
      try {
        const f = await window.storage.get("pita-lab-filaments");
        if (f?.value) setFilamentsCatalog(JSON.parse(f.value));
      } catch (e) {
        /* no filamentos cadastrados ainda */
      }
      let parsedStock = [];
      try {
        const st = await window.storage.get("pita-lab-stock");
        if (st?.value) {
          parsedStock = JSON.parse(st.value);
          setStockItems(parsedStock);
        }
      } catch (e) {
        /* no estoque cadastrado ainda */
      }
      try {
        const sh = await window.storage.get("pita-lab-sales");
        if (sh?.value) setSalesHistory(JSON.parse(sh.value));
      } catch (e) {
        /* nenhuma venda registrada ainda */
      }
      try {
        const pl = await window.storage.get("pita-lab-production-log");
        if (pl?.value) {
          setProductionLog(JSON.parse(pl.value));
        } else if (parsedStock.length > 0) {
          // primeira vez usando o histórico de produção: preenche com o estoque atual
          const seeded = parsedStock.map((item) => ({
            id: `plog_seed_${item.id}`,
            name: item.name,
            printed: item.printed || 0,
            date: new Date().toISOString(),
          }));
          setProductionLog(seeded);
          await window.storage.set("pita-lab-production-log", JSON.stringify(seeded));
        }
      } catch (e) {
        /* nenhum histórico de produção ainda */
      }
      setLoaded(true);
    })();
  }, []);

  const persistSettings = async (next) => {
    setSettings(next);
    try {
      await window.storage.set("pita-lab-settings", JSON.stringify(next));
    } catch (e) {
      console.error("Falha ao salvar configurações", e);
    }
  };

  const persistQuotes = async (next) => {
    setQuotes(next);
    try {
      await window.storage.set("pita-lab-quotes", JSON.stringify(next));
    } catch (e) {
      console.error("Falha ao salvar histórico", e);
    }
  };

  const persistInsumosCatalog = async (next) => {
    setInsumosCatalog(next);
    try {
      await window.storage.set("pita-lab-insumos", JSON.stringify(next));
    } catch (e) {
      console.error("Falha ao salvar catálogo de insumos", e);
    }
  };

  const persistFilamentsCatalog = async (next) => {
    setFilamentsCatalog(next);
    try {
      await window.storage.set("pita-lab-filaments", JSON.stringify(next));
    } catch (e) {
      console.error("Falha ao salvar catálogo de filamentos", e);
    }
  };

  const persistStock = async (next) => {
    setStockItems(next);
    try {
      await window.storage.set("pita-lab-stock", JSON.stringify(next));
    } catch (e) {
      console.error("Falha ao salvar estoque", e);
    }
  };

  const persistSales = async (next) => {
    setSalesHistory(next);
    try {
      await window.storage.set("pita-lab-sales", JSON.stringify(next));
    } catch (e) {
      console.error("Falha ao salvar histórico de vendas", e);
    }
  };

  const persistProductionLog = async (next) => {
    setProductionLog(next);
    try {
      await window.storage.set("pita-lab-production-log", JSON.stringify(next));
    } catch (e) {
      console.error("Falha ao salvar histórico de produção", e);
    }
  };

  // ---------- Filamentos: catálogo + estoque (gramas) ----------
  const handleAddFilamentToCatalog = () => {
    const price = parseFloat(newFilament.costPerKg);
    if (!newFilament.brand.trim() || !newFilament.color.trim() || !isFinite(price) || price <= 0) return;

    if (editingFilamentId) {
      persistFilamentsCatalog(
        filamentsCatalog.map((f) =>
          f.id === editingFilamentId
            ? { ...f, brand: newFilament.brand.trim(), color: newFilament.color.trim(), costPerKg: price }
            : f
        )
      );
      setEditingFilamentId(null);
    } else {
      const stockGrams = parseFloat(newFilament.stockGrams);
      const entry = {
        id: `f_${Date.now()}`,
        brand: newFilament.brand.trim(),
        color: newFilament.color.trim(),
        costPerKg: price,
        stockQuantity: isFinite(stockGrams) ? stockGrams : 0,
      };
      persistFilamentsCatalog([...filamentsCatalog, entry]);
    }
    setNewFilament({ brand: "", color: "", costPerKg: "", stockGrams: "1000" });
  };

  const handleEditFilament = (item) => {
    setNewFilament({ brand: item.brand, color: item.color, costPerKg: String(item.costPerKg), stockGrams: "" });
    setEditingFilamentId(item.id);
  };

  const handleCancelEditFilament = () => {
    setEditingFilamentId(null);
    setNewFilament({ brand: "", color: "", costPerKg: "", stockGrams: "1000" });
  };

  const handleDeleteFilamentFromCatalog = (id) => {
    persistFilamentsCatalog(filamentsCatalog.filter((f) => f.id !== id));
    if (editingFilamentId === id) handleCancelEditFilament();
  };

  const handleAdjustFilamentStock = (filamentId, delta) => {
    persistFilamentsCatalog(
      filamentsCatalog.map((f) => (f.id === filamentId ? { ...f, stockQuantity: (f.stockQuantity || 0) + delta } : f))
    );
  };

  const handleSetFilamentStock = (filamentId, value) => {
    const parsed = parseFloat(value);
    persistFilamentsCatalog(
      filamentsCatalog.map((f) => (f.id === filamentId ? { ...f, stockQuantity: isFinite(parsed) ? parsed : 0 } : f))
    );
  };

  const handleAddFilamentToQuote = () => {
    if (form.usedFilaments.length >= 4) return;
    const catalogItem = filamentsCatalog.find((f) => f.id === filamentToAdd.id);
    const grams = parseFloat(filamentToAdd.grams);
    if (!catalogItem || !isFinite(grams) || grams <= 0) return;
    setForm({
      ...form,
      usedFilaments: [
        ...form.usedFilaments,
        {
          lineId: `fl_${Date.now()}`,
          filamentId: catalogItem.id,
          brand: catalogItem.brand,
          color: catalogItem.color,
          costPerKg: catalogItem.costPerKg,
          grams,
        },
      ],
    });
    setFilamentToAdd({ id: "", grams: "" });
  };

  const handleRemoveFilamentFromQuote = (lineId) => {
    setForm({ ...form, usedFilaments: form.usedFilaments.filter((l) => l.lineId !== lineId) });
  };

  const handleUpdateFilamentLineGrams = (lineId, value) => {
    const grams = parseFloat(value);
    setForm({
      ...form,
      usedFilaments: form.usedFilaments.map((l) =>
        l.lineId === lineId ? { ...l, grams: isFinite(grams) ? grams : 0 } : l
      ),
    });
  };

  const handleUpdateFilamentLineChoice = (lineId, filamentId) => {
    const catalogItem = filamentsCatalog.find((f) => f.id === filamentId);
    if (!catalogItem) return;
    setForm({
      ...form,
      usedFilaments: form.usedFilaments.map((l) =>
        l.lineId === lineId
          ? { ...l, filamentId: catalogItem.id, brand: catalogItem.brand, color: catalogItem.color, costPerKg: catalogItem.costPerKg }
          : l
      ),
    });
  };

  // ---------- Insumos: catálogo + estoque (unidades) ----------
  const handleAddInsumoToCatalog = () => {
    const totalPaid = parseFloat(newInsumo.totalPaid);
    const quantity = parseFloat(newInsumo.quantity);
    if (!newInsumo.name.trim() || !isFinite(totalPaid) || totalPaid <= 0 || !isFinite(quantity) || quantity <= 0)
      return;

    if (editingInsumoId) {
      persistInsumosCatalog(
        insumosCatalog.map((i) =>
          i.id === editingInsumoId
            ? {
                ...i,
                name: newInsumo.name.trim(),
                totalPaid,
                quantity,
                unitPrice: totalPaid / quantity,
                unit: newInsumo.unit || "unidade",
              }
            : i
        )
      );
      setEditingInsumoId(null);
    } else {
      const entry = {
        id: `i_${Date.now()}`,
        name: newInsumo.name.trim(),
        totalPaid,
        quantity,
        unitPrice: totalPaid / quantity,
        unit: newInsumo.unit || "unidade",
        stockQuantity: quantity,
      };
      persistInsumosCatalog([...insumosCatalog, entry]);
    }
    setNewInsumo({ name: "", totalPaid: "", quantity: "", unit: "unidade" });
  };

  const handleEditInsumo = (item) => {
    setNewInsumo({
      name: item.name,
      totalPaid: String(item.totalPaid ?? item.unitPrice),
      quantity: String(item.quantity ?? 1),
      unit: item.unit || "unidade",
    });
    setEditingInsumoId(item.id);
  };

  const handleCancelEditInsumo = () => {
    setEditingInsumoId(null);
    setNewInsumo({ name: "", totalPaid: "", quantity: "", unit: "unidade" });
  };

  const handleDeleteInsumoFromCatalog = (id) => {
    persistInsumosCatalog(insumosCatalog.filter((i) => i.id !== id));
    if (editingInsumoId === id) handleCancelEditInsumo();
  };

  const handleAdjustInsumoStock = (insumoId, delta) => {
    persistInsumosCatalog(
      insumosCatalog.map((i) => (i.id === insumoId ? { ...i, stockQuantity: (i.stockQuantity || 0) + delta } : i))
    );
  };

  const handleSetInsumoStock = (insumoId, value) => {
    const parsed = parseFloat(value);
    persistInsumosCatalog(
      insumosCatalog.map((i) => (i.id === insumoId ? { ...i, stockQuantity: isFinite(parsed) ? parsed : 0 } : i))
    );
  };

  const handleAddInsumoToQuote = () => {
    const catalogItem = insumosCatalog.find((i) => i.id === insumoToAdd.id);
    const qty = parseFloat(insumoToAdd.quantity);
    if (!catalogItem || !isFinite(qty) || qty <= 0) return;
    setForm({
      ...form,
      usedInsumos: [
        ...form.usedInsumos,
        {
          lineId: `l_${Date.now()}`,
          insumoId: catalogItem.id,
          name: catalogItem.name,
          unitPrice: catalogItem.unitPrice,
          unit: catalogItem.unit,
          quantity: qty,
        },
      ],
    });
    setInsumoToAdd({ id: "", quantity: "1" });
  };

  const handleRemoveInsumoFromQuote = (lineId) => {
    setForm({ ...form, usedInsumos: form.usedInsumos.filter((l) => l.lineId !== lineId) });
  };

  // ---------- Cálculo do orçamento ----------
  const printMinutes = hhmmToMinutes(form.printMinutes);
  const laborMinutes = parseFloat(form.laborMinutes) || 0;
  const printHours = printMinutes / 60;
  const laborHours = laborMinutes / 60;
  const piecesPerBatch = Math.max(1, parseFloat(form.piecesPerBatch) || 1);
  const failureRate =
    form.failureRateOverride !== "" ? parseFloat(form.failureRateOverride) || 0 : settings.failureRatePercent;

  const materialBatch = form.usedFilaments.reduce((sum, l) => sum + (l.grams / 1000) * l.costPerKg, 0);
  const energyBatch = printHours * (settings.printerWattage / 1000) * settings.energyPricePerKwh;
  const depreciationBatch =
    settings.printerLifetimeHours > 0
      ? (settings.printerValue / settings.printerLifetimeHours) * printHours
      : 0;

  const material = materialBatch / piecesPerBatch;
  const energy = energyBatch / piecesPerBatch;
  const depreciation = depreciationBatch / piecesPerBatch;
  const labor = laborHours * settings.laborRatePerHour; // já é por peça, não dilui pelo lote
  const insumosCost = form.usedInsumos.reduce((sum, l) => sum + l.unitPrice * l.quantity, 0);
  const subtotal = material + energy + depreciation + labor + insumosCost;
  const failureValue = subtotal * (failureRate / 100);
  const total = subtotal + failureValue; // custo final de fabricação (sem margem de lucro)
  const batchTotal = (materialBatch + energyBatch + depreciationBatch) + (labor + insumosCost) * piecesPerBatch;

  const segments = [
    { key: "material", value: material },
    { key: "energy", value: energy },
    { key: "depreciation", value: depreciation },
    { key: "labor", value: labor },
    { key: "insumos", value: insumosCost },
    { key: "failures", value: failureValue },
  ];
  const grandTotal = total || 1;

  const canSave = form.name.trim().length > 0 && total > 0;
  const emptyForm = {
    name: "",
    category: "",
    material: "",
    printMinutes: "",
    laborMinutes: "",
    piecesPerBatch: "1",
    failureRateOverride: "",
    usedInsumos: [],
    usedFilaments: [],
  };

  const nameChangedWhileEditing = editingId !== null && form.name.trim() !== originalEditName.trim();

  const buildSnapshot = () => ({
    name: form.name.trim(),
    category: form.category.trim(),
    material: form.material.trim(),
    printMinutes: String(printMinutes),
    laborMinutes: form.laborMinutes,
    piecesPerBatch: form.piecesPerBatch,
    failureRateOverride: String(failureRate),
    usedInsumos: form.usedInsumos,
    usedFilaments: form.usedFilaments,
  });

  const handleUpdateQuote = async () => {
    if (!canSave || !editingId) return;
    const snapshot = buildSnapshot();
    const next = quotes.map((q) =>
      q.id === editingId ? { ...q, ...snapshot, total, subtotal, date: new Date().toISOString() } : q
    );
    await persistQuotes(next);
    setEditingId(null);
    setLegacyEditWarning(false);
    setOriginalEditName("");
    setForm(emptyForm);
    setSaveStatus("Orçamento atualizado!");
    setTimeout(() => setSaveStatus(""), 2000);
  };

  const handleSaveAsNewQuote = async (nameOverride) => {
    if (!canSave) return;
    const finalName = (nameOverride ?? form.name).trim();
    if (!finalName) return;
    const snapshot = { ...buildSnapshot(), name: finalName };
    const entry = { id: `q_${Date.now()}`, ...snapshot, total, subtotal, date: new Date().toISOString() };
    await persistQuotes([entry, ...quotes]);
    setEditingId(null);
    setLegacyEditWarning(false);
    setOriginalEditName("");
    setForm(emptyForm);
    setSaveStatus("Orçamento salvo!");
    setTimeout(() => setSaveStatus(""), 2000);
  };

  const handleOpenCopyPrompt = () => {
    setCopyNameInput(`${form.name.trim()} (cópia)`);
    setCopyPromptOpen(true);
  };

  const handleCancelCopyPrompt = () => {
    setCopyPromptOpen(false);
    setCopyNameInput("");
  };

  const handleConfirmCopyPrompt = async () => {
    const trimmed = copyNameInput.trim();
    if (!trimmed) return;
    await handleSaveAsNewQuote(trimmed);
    setCopyPromptOpen(false);
    setCopyNameInput("");
  };

  const handleEditQuote = (quote) => {
    const isLegacy = quote.printMinutes === undefined;
    setForm({
      name: quote.name,
      category: quote.category ?? "",
      material: quote.material ?? "",
      printMinutes: quote.printMinutes !== undefined && quote.printMinutes !== "" ? minutesToHHMM(quote.printMinutes) : "",
      laborMinutes: quote.laborMinutes ?? "",
      piecesPerBatch: quote.piecesPerBatch ?? "1",
      failureRateOverride: quote.failureRateOverride ?? "",
      usedInsumos: quote.usedInsumos ?? [],
      usedFilaments: quote.usedFilaments ?? [],
    });
    setEditingId(quote.id);
    setOriginalEditName(quote.name);
    setLegacyEditWarning(isLegacy);
  };

  const handleCancelEdit = () => {
    setEditingId(null);
    setLegacyEditWarning(false);
    setOriginalEditName("");
    setForm(emptyForm);
  };

  const handleDelete = async (id) => {
    await persistQuotes(quotes.filter((q) => q.id !== id));
    if (editingId === id) handleCancelEdit();
  };

  // ---------- Estoque de produtos ----------
  const handleAddStockItem = () => {
    const printed = parseFloat(newStock.printed);
    const unitCost = parseFloat(newStock.unitCost);
    const marginPercent = parseFloat(newStock.marginPercent);
    const unitPrice = parseFloat(newStock.unitPrice);
    if (!newStock.name.trim() || !isFinite(printed) || printed < 0) return;

    const usedInsumos = newStock.usedInsumos || [];
    const usedFilaments = newStock.usedFilaments || [];
    const entry = {
      id: `s_${Date.now()}`,
      name: newStock.name.trim(),
      category: (newStock.category || "").trim(),
      material: (newStock.material || "").trim(),
      printed,
      sold: 0,
      unitCost: isFinite(unitCost) ? unitCost : 0,
      marginPercent: isFinite(marginPercent) ? marginPercent : 0,
      unitPrice: isFinite(unitPrice) ? unitPrice : 0,
      usedInsumos,
      usedFilaments,
    };

    if (printed > 0) {
      if (usedInsumos.length > 0) {
        persistInsumosCatalog(
          insumosCatalog.map((cat) => {
            const line = usedInsumos.find((l) => l.insumoId === cat.id);
            if (!line) return cat;
            return { ...cat, stockQuantity: (cat.stockQuantity || 0) - line.quantity * printed };
          })
        );
      }
      if (usedFilaments.length > 0) {
        persistFilamentsCatalog(
          filamentsCatalog.map((cat) => {
            const line = usedFilaments.find((l) => l.filamentId === cat.id);
            if (!line) return cat;
            return { ...cat, stockQuantity: (cat.stockQuantity || 0) - line.grams * printed };
          })
        );
      }
    }

    persistStock([...stockItems, entry]);
    if (printed > 0) {
      persistProductionLog([
        ...productionLog,
        { id: `plog_${Date.now()}`, name: entry.name, printed, date: new Date().toISOString() },
      ]);
    }
    setNewStock({ name: "", category: "", material: "", printed: "", unitCost: "", marginPercent: "50", unitPrice: "", usedInsumos: [], usedFilaments: [] });
  };

  const handleUpdateStockField = (id, field, value) => {
    const parsed = parseFloat(value);
    persistStock(stockItems.map((item) => (item.id === id ? { ...item, [field]: isFinite(parsed) ? parsed : 0 } : item)));
  };

  const handleUpdateStockTextField = (id, field, value) => {
    persistStock(stockItems.map((item) => (item.id === id ? { ...item, [field]: value } : item)));
  };

  const handleUpdateStockPrinted = (id, value) => {
    const parsed = parseFloat(value);
    const newPrinted = isFinite(parsed) ? parsed : 0;
    const item = stockItems.find((s) => s.id === id);
    if (!item) return;
    const delta = newPrinted - item.printed;
    const usedInsumos = item.usedInsumos || [];
    const usedFilaments = item.usedFilaments || [];

    if (delta !== 0) {
      if (usedInsumos.length > 0) {
        persistInsumosCatalog(
          insumosCatalog.map((cat) => {
            const line = usedInsumos.find((l) => l.insumoId === cat.id);
            if (!line) return cat;
            return { ...cat, stockQuantity: (cat.stockQuantity || 0) - line.quantity * delta };
          })
        );
      }
      if (usedFilaments.length > 0) {
        persistFilamentsCatalog(
          filamentsCatalog.map((cat) => {
            const line = usedFilaments.find((l) => l.filamentId === cat.id);
            if (!line) return cat;
            return { ...cat, stockQuantity: (cat.stockQuantity || 0) - line.grams * delta };
          })
        );
      }
    }
    persistStock(stockItems.map((s) => (s.id === id ? { ...s, printed: newPrinted } : s)));
    if (delta !== 0) {
      persistProductionLog([
        ...productionLog,
        { id: `plog_${Date.now()}`, name: item.name, printed: delta, date: new Date().toISOString() },
      ]);
    }
  };

  const handleDeleteStockItem = (id) => {
    persistStock(stockItems.filter((item) => item.id !== id));
  };

  const handleStartEditStockQty = (item) => {
    const inStock = item.printed - item.sold;
    setEditingStockQtyId(item.id);
    setEditStockQtyValue(String(inStock));
  };

  const handleCancelEditStockQty = () => {
    setEditingStockQtyId(null);
    setEditStockQtyValue("");
  };

  const handleSaveEditStockQty = (item) => {
    const newInStock = parseFloat(editStockQtyValue);
    if (!isFinite(newInStock) || newInStock < 0) return;
    const newPrinted = newInStock + item.sold;
    handleUpdateStockPrinted(item.id, String(newPrinted));
    setEditingStockQtyId(null);
    setEditStockQtyValue("");
  };

  const handleRegisterSale = (id, quantityStr, unitPriceStr) => {
    const item = stockItems.find((s) => s.id === id);
    if (!item) return;
    const qty = parseFloat(quantityStr);
    const available = item.printed - item.sold;
    if (!isFinite(qty) || qty <= 0 || qty > available) return;
    const parsedPrice = parseFloat(unitPriceStr);
    const unitPriceUsed = isFinite(parsedPrice) && parsedPrice >= 0 ? parsedPrice : item.unitPrice;

    const saleEntry = {
      id: `sale_${Date.now()}`,
      stockItemId: item.id,
      itemName: item.name,
      quantity: qty,
      unitPrice: unitPriceUsed,
      catalogUnitPrice: item.unitPrice,
      unitCost: item.unitCost,
      revenue: qty * unitPriceUsed,
      cost: qty * item.unitCost,
      profit: qty * (unitPriceUsed - item.unitCost),
      date: new Date().toISOString(),
    };

    persistStock(stockItems.map((s) => (s.id === id ? { ...s, sold: s.sold + qty } : s)));
    persistSales([saleEntry, ...salesHistory]);
    setSellingItemId(null);
    setSaleQuantity("");
    setSaleUnitPrice("");
  };

  // ---------- Backup: exportar / importar todos os dados ----------
  const handleExportData = () => {
    const payload = {
      app: "pita-lab-precificador",
      version: 3,
      exportedAt: new Date().toISOString(),
      settings,
      quotes,
      insumosCatalog,
      filamentsCatalog,
      stockItems,
      salesHistory,
      productionLog,
    };
    try {
      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `pita-lab-backup-${new Date().toISOString().slice(0, 10)}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      setImportStatus("Backup exportado!");
    } catch (e) {
      console.error("Falha ao exportar dados", e);
      setImportStatus("Erro ao exportar.");
    }
    setTimeout(() => setImportStatus(""), 2500);
  };

  // ---------- Exportar catálogo (.xlsx) no template usado para venda ----------
  const handleExportCatalog = () => {
    try {
      const readyItems = stockItems.filter((item) => item.printed - item.sold > 0);

      const headerRow = ["Nome do produto", "Preço (R$)", "Categoria", "Material", "Estoque (unidades)"];
      const dataRows = readyItems.map((item) => [
        item.name,
        item.unitPrice || 0,
        item.category || "",
        item.material || "",
        item.printed - item.sold,
      ]);

      const produtosSheetData = [
        [`Catálogo gerado pelo Pita's Lab em ${new Date().toLocaleDateString("pt-BR")}. Não altere os nomes das colunas.`],
        headerRow,
        ...dataRows,
      ];
      const wsProdutos = XLSX.utils.aoa_to_sheet(produtosSheetData);
      wsProdutos["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }];
      wsProdutos["!cols"] = [{ wch: 32 }, { wch: 12 }, { wch: 18 }, { wch: 14 }, { wch: 18 }];

      const fotosSheetData = [
        ["Como adicionar as fotos dos produtos"],
        [],
        ['1. Volte para a aba "Produtos".'],
        ["2. Insira a foto do jeito clássico: menu Inserir > Imagem > Do arquivo."],
        ["3. Arraste a foto para que ela fique posicionada na MESMA LINHA do produto correspondente."],
        ["4. Uma foto por produto. Não precisa redimensionar ou formatar."],
        ["5. Depois de preencher tudo (fotos), salve o arquivo e envie para onde for usar o catálogo."],
      ];
      const wsFotos = XLSX.utils.aoa_to_sheet(fotosSheetData);
      wsFotos["!cols"] = [{ wch: 90 }];

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, wsProdutos, "Produtos");
      XLSX.utils.book_append_sheet(wb, wsFotos, "Como enviar as fotos");

      XLSX.writeFile(wb, `pita-lab-catalogo-${new Date().toISOString().slice(0, 10)}.xlsx`);
      setCatalogExportStatus(`Catálogo exportado com ${dataRows.length} produto(s)!`);
    } catch (e) {
      console.error("Falha ao exportar catálogo", e);
      setCatalogExportStatus("Erro ao exportar catálogo.");
    }
    setTimeout(() => setCatalogExportStatus(""), 3000);
  };


  const handleImportFileSelected = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (event) => {
      try {
        const data = JSON.parse(event.target.result);
        if (data.settings) await persistSettings({ ...DEFAULT_SETTINGS, ...data.settings });
        if (Array.isArray(data.quotes)) await persistQuotes(data.quotes);
        if (Array.isArray(data.insumosCatalog)) await persistInsumosCatalog(data.insumosCatalog);
        if (Array.isArray(data.filamentsCatalog)) await persistFilamentsCatalog(data.filamentsCatalog);
        if (Array.isArray(data.stockItems)) await persistStock(data.stockItems);
        if (Array.isArray(data.salesHistory)) await persistSales(data.salesHistory);
        if (Array.isArray(data.productionLog)) await persistProductionLog(data.productionLog);
        setImportStatus("Dados importados com sucesso!");
      } catch (err) {
        console.error("Falha ao importar dados", err);
        setImportStatus("Erro: arquivo inválido.");
      }
      setTimeout(() => setImportStatus(""), 3000);
    };
    reader.readAsText(file);
    e.target.value = "";
  };

  const handleDeleteSale = (saleId) => {
    const sale = salesHistory.find((s) => s.id === saleId);
    if (!sale) return;
    persistStock(
      stockItems.map((s) =>
        s.id === sale.stockItemId ? { ...s, sold: Math.max(0, s.sold - sale.quantity) } : s
      )
    );
    persistSales(salesHistory.filter((s) => s.id !== saleId));
  };

  const handleStartEditSale = (sale) => {
    setEditingSaleId(sale.id);
    setEditSaleForm({ quantity: String(sale.quantity), unitPrice: String(sale.unitPrice) });
  };

  const handleCancelEditSale = () => {
    setEditingSaleId(null);
    setEditSaleForm({ quantity: "", unitPrice: "" });
  };

  const handleSaveEditSale = (saleId) => {
    const sale = salesHistory.find((s) => s.id === saleId);
    if (!sale) return;
    const newQty = parseFloat(editSaleForm.quantity);
    const newPrice = parseFloat(editSaleForm.unitPrice);
    if (!isFinite(newQty) || newQty <= 0 || !isFinite(newPrice) || newPrice < 0) return;

    const stockItem = stockItems.find((s) => s.id === sale.stockItemId);
    const qtyDiff = newQty - sale.quantity;

    if (stockItem) {
      const availableBeforeChange = stockItem.printed - stockItem.sold;
      if (qtyDiff > availableBeforeChange) return; // não há peças suficientes em estoque para o aumento
      persistStock(
        stockItems.map((s) => (s.id === stockItem.id ? { ...s, sold: Math.max(0, s.sold + qtyDiff) } : s))
      );
    }

    const unitCost = stockItem ? stockItem.unitCost : sale.unitCost;
    const updatedSale = {
      ...sale,
      quantity: newQty,
      unitPrice: newPrice,
      unitCost,
      revenue: newQty * newPrice,
      cost: newQty * unitCost,
      profit: newQty * (newPrice - unitCost),
    };
    persistSales(salesHistory.map((s) => (s.id === saleId ? updatedSale : s)));
    setEditingSaleId(null);
    setEditSaleForm({ quantity: "", unitPrice: "" });
  };

  // ---------- Fluxo de caixa ----------
  const materialsInventoryValue =
    filamentsCatalog.reduce((sum, f) => sum + ((f.stockQuantity || 0) / 1000) * (f.costPerKg || 0), 0) +
    insumosCatalog.reduce((sum, i) => sum + (i.stockQuantity || 0) * (i.unitPrice || 0), 0);
  const producedGoodsCost = stockItems.reduce((sum, item) => sum + (item.printed || 0) * (item.unitCost || 0), 0);
  const totalInvested = materialsInventoryValue + producedGoodsCost;
  const totalReceived = salesHistory.reduce((sum, s) => sum + (s.revenue || 0), 0);
  const totalSalesProfit = salesHistory.reduce((sum, s) => sum + (s.profit || 0), 0);
  const netCashFlow = totalReceived - totalInvested;

  // Histórico permanente por produto (nome) — não é afetado por vendas, edições ou exclusão de itens do estoque
  const productAggregates = {};
  productionLog.forEach((log) => {
    if (!productAggregates[log.name]) {
      productAggregates[log.name] = { name: log.name, Impressas: 0, Vendidas: 0, Lucro: 0 };
    }
    productAggregates[log.name].Impressas += log.printed || 0;
  });
  salesHistory.forEach((s) => {
    if (!productAggregates[s.itemName]) {
      productAggregates[s.itemName] = { name: s.itemName, Impressas: 0, Vendidas: 0, Lucro: 0 };
    }
    productAggregates[s.itemName].Vendidas += s.quantity || 0;
    productAggregates[s.itemName].Lucro += s.profit || 0;
  });
  const productChartData = Object.values(productAggregates).sort((a, b) => b.Impressas - a.Impressas);

  // Histórico de categorias e materiais já usados (para os dropdowns de sugestão)
  const existingCategories = Array.from(
    new Set([...quotes.map((q) => q.category), ...stockItems.map((s) => s.category)].filter(Boolean))
  ).sort((a, b) => a.localeCompare(b, "pt-BR"));
  const existingMaterials = Array.from(
    new Set([...quotes.map((q) => q.material), ...stockItems.map((s) => s.material)].filter(Boolean))
  ).sort((a, b) => a.localeCompare(b, "pt-BR"));

  const inputStyle = {
    background: "#1F2733",
    border: "1px solid #2A3341",
    color: "#EAF0F6",
  };

  const tabs = [
    { key: "orcamento", label: "Orçamento", icon: <ClipboardList size={16} /> },
    { key: "estoque", label: "Estoque de Produtos", icon: <Warehouse size={16} /> },
    { key: "vendas", label: "Vendas & Caixa", icon: <Wallet size={16} /> },
    { key: "materiais", label: "Estoque de Insumos", icon: <Boxes size={16} /> },
    { key: "custos", label: "Custos-base", icon: <Settings2 size={16} /> },
  ];

  return (
    <div style={{ background: "#0F1319", minHeight: "100vh", color: "#EAF0F6" }} className="w-full px-4 py-6 sm:px-8">
      <div className="max-w-3xl mx-auto">
        {/* Header */}
        <div className="flex items-center justify-between gap-3 mb-1">
          <div className="flex items-center gap-3">
            <div style={{ background: "#2E9BF0" }} className="w-10 h-10 rounded-full flex items-center justify-center shrink-0">
              <PawPrint size={20} color="#0F1319" strokeWidth={2.5} />
            </div>
            <div>
              <h1 className="text-lg font-bold tracking-tight">Pita's Lab</h1>
              <p className="text-xs" style={muted}>
                Calculadora de orçamentos · impressão 3D
              </p>
            </div>
          </div>
          <button
            onClick={() => setShowSettingsModal(true)}
            className="w-9 h-9 rounded-full flex items-center justify-center shrink-0"
            style={{ background: "#171D26", border: "1px solid #2A3341", color: "#8B98AA" }}
            aria-label="Configurações"
          >
            <Settings2 size={17} />
          </button>
        </div>

        {/* Tabs */}
        <div className="grid grid-cols-2 gap-1.5 mt-5 mb-5">
          {tabs.map((t, idx) => (
            <button
              key={t.key}
              onClick={() => setActiveTab(t.key)}
              className={`flex flex-col items-center justify-center gap-1 py-2.5 rounded-xl text-xs font-semibold ${
                idx === tabs.length - 1 && tabs.length % 2 === 1 ? "col-span-2" : ""
              }`}
              style={{
                background: activeTab === t.key ? "#2E9BF0" : "#171D26",
                color: activeTab === t.key ? "#0F1319" : "#8B98AA",
                border: "1px solid #2A3341",
              }}
            >
              {t.icon}
              {t.label}
            </button>
          ))}
        </div>

        {/* ============ ORÇAMENTO ============ */}
        {activeTab === "orcamento" && (
          <>
            {/* Quote form */}
            <div className="rounded-xl p-4 mb-5" style={cardStyle}>
              <label className="text-xs font-medium block mb-1.5" style={muted}>
                Nome da peça
              </label>
              <input
                value={form.name}
                onChange={(e) => setForm({ ...form, name: e.target.value })}
                placeholder="Ex: Chaveiro Border Collie"
                className="w-full rounded-lg px-3 py-2.5 text-sm mb-4 outline-none"
                style={inputStyle}
              />

              <div className="grid grid-cols-2 gap-3 mb-3">
                <SuggestField
                  icon={<ClipboardList size={14} />}
                  label="Categoria"
                  value={form.category}
                  onChange={(v) => setForm({ ...form, category: v })}
                  inputStyle={inputStyle}
                  placeholder="Ex: Decoração"
                  options={existingCategories}
                />
                <SuggestField
                  icon={<Package size={14} />}
                  label="Material"
                  value={form.material}
                  onChange={(v) => setForm({ ...form, material: v })}
                  inputStyle={inputStyle}
                  placeholder="Ex: PLA"
                  options={existingMaterials}
                />
              </div>

              <div className="grid grid-cols-2 gap-3 mb-1">
                <div>
                  <IconField
                    icon={<Gauge size={14} />}
                    label="Tempo de impressão (hh:mm)"
                    type="text"
                    value={form.printMinutes}
                    onChange={(v) => setForm({ ...form, printMinutes: v.replace(/[^0-9:]/g, "") })}
                    inputStyle={inputStyle}
                    placeholder="ex: 1:30"
                  />
                  {form.printMinutes !== "" && (
                    <p className="text-[10px] mt-1" style={muted}>
                      = {printMinutes} minutos ({printHours.toFixed(2).replace(".", ",")}h)
                    </p>
                  )}
                </div>
                <IconField
                  icon={<Boxes size={14} />}
                  label="Peças nesta impressão (build plate)"
                  value={form.piecesPerBatch}
                  onChange={(v) => setForm({ ...form, piecesPerBatch: v })}
                  inputStyle={inputStyle}
                  placeholder="1"
                />
                <IconField
                  icon={<Wrench size={14} />}
                  label="Minutos de mão de obra (por peça)"
                  value={form.laborMinutes}
                  onChange={(v) => setForm({ ...form, laborMinutes: v })}
                  inputStyle={inputStyle}
                />
                <div className="col-span-2">
                  <IconField
                    icon={<TrendingUp size={14} />}
                    label={`Índice de falhas % (padrão ${settings.failureRatePercent})`}
                    value={form.failureRateOverride}
                    onChange={(v) => setForm({ ...form, failureRateOverride: v })}
                    inputStyle={inputStyle}
                    placeholder={String(settings.failureRatePercent)}
                  />
                </div>
              </div>

              {/* Filamentos usados nesta peça (até 4) */}
              <div className="mt-4 pt-4" style={{ borderTop: "1px solid #2A3341" }}>
                <label className="text-xs font-medium mb-1.5 flex items-center gap-1.5" style={muted}>
                  <Package size={14} />
                  Filamentos usados nesta peça ({form.usedFilaments.length}/4)
                </label>

                {filamentsCatalog.length === 0 ? (
                  <p className="text-xs" style={muted}>
                    Cadastre filamentos na aba Estoque de Insumos para poder selecioná-los aqui.
                  </p>
                ) : form.usedFilaments.length >= 4 ? (
                  <p className="text-xs" style={muted}>
                    Limite de 4 filamentos por peça atingido.
                  </p>
                ) : (
                  <div className="flex gap-2 mb-2">
                    <select
                      value={filamentToAdd.id}
                      onChange={(e) => setFilamentToAdd({ ...filamentToAdd, id: e.target.value })}
                      className="flex-1 rounded-lg px-2 py-2 text-sm outline-none"
                      style={inputStyle}
                    >
                      <option value="">Selecione um filamento</option>
                      {filamentsCatalog.map((f) => (
                        <option key={f.id} value={f.id}>
                          {f.brand} — {f.color} ({currency(f.costPerKg)}/kg)
                        </option>
                      ))}
                    </select>
                    <input
                      type="number"
                      value={filamentToAdd.grams}
                      onChange={(e) => setFilamentToAdd({ ...filamentToAdd, grams: e.target.value })}
                      placeholder="Gramas"
                      className="w-20 rounded-lg px-2 py-2 text-sm outline-none"
                      style={inputStyle}
                    />
                    <button
                      onClick={handleAddFilamentToQuote}
                      className="rounded-lg px-3 flex items-center justify-center"
                      style={{ background: "#2E9BF0", color: "#0F1319" }}
                    >
                      <Plus size={16} />
                    </button>
                  </div>
                )}

                {form.usedFilaments.length > 0 && (
                  <div className="flex flex-col gap-1.5 mt-2">
                    {form.usedFilaments.map((l) => (
                      <div key={l.lineId} className="flex items-center gap-1.5 text-xs px-2.5 py-1.5 rounded-lg flex-wrap" style={{ background: "#1F2733" }}>
                        <select
                          value={l.filamentId}
                          onChange={(e) => handleUpdateFilamentLineChoice(l.lineId, e.target.value)}
                          className="flex-1 min-w-[110px] rounded px-1.5 py-1 text-xs outline-none"
                          style={inputStyle}
                        >
                          {filamentsCatalog.map((f) => (
                            <option key={f.id} value={f.id}>
                              {f.brand} — {f.color}
                            </option>
                          ))}
                        </select>
                        <input
                          type="number"
                          value={l.grams}
                          onChange={(e) => handleUpdateFilamentLineGrams(l.lineId, e.target.value)}
                          className="w-16 rounded px-1.5 py-1 text-xs text-right outline-none"
                          style={inputStyle}
                        />
                        <span className="text-[10px]" style={muted}>
                          g
                        </span>
                        <span className="whitespace-nowrap" style={{ color: "#2E9BF0", fontFamily: "ui-monospace, monospace" }}>
                          {currency((l.grams / 1000) * l.costPerKg)}
                        </span>
                        <button onClick={() => handleRemoveFilamentFromQuote(l.lineId)} style={muted}>
                          <X size={13} />
                        </button>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              {/* Insumos usados nesta peça */}
              <div className="mt-4 pt-4" style={{ borderTop: "1px solid #2A3341" }}>
                <label className="text-xs font-medium mb-1.5 flex items-center gap-1.5" style={muted}>
                  <Boxes size={14} />
                  Insumos usados nesta peça
                </label>

                {insumosCatalog.length === 0 ? (
                  <p className="text-xs" style={muted}>
                    Cadastre insumos na aba Materiais para poder adicioná-los aqui.
                  </p>
                ) : (
                  <div className="flex gap-2 mb-2">
                    <select
                      value={insumoToAdd.id}
                      onChange={(e) => setInsumoToAdd({ ...insumoToAdd, id: e.target.value })}
                      className="flex-1 rounded-lg px-2 py-2 text-sm outline-none"
                      style={inputStyle}
                    >
                      <option value="">Selecione um insumo</option>
                      {insumosCatalog.map((i) => (
                        <option key={i.id} value={i.id}>
                          {i.name} ({currency(i.unitPrice)}/{i.unit})
                        </option>
                      ))}
                    </select>
                    <input
                      type="number"
                      value={insumoToAdd.quantity}
                      onChange={(e) => setInsumoToAdd({ ...insumoToAdd, quantity: e.target.value })}
                      placeholder="Qtd"
                      className="w-16 rounded-lg px-2 py-2 text-sm outline-none"
                      style={inputStyle}
                    />
                    <button
                      onClick={handleAddInsumoToQuote}
                      className="rounded-lg px-3 flex items-center justify-center"
                      style={{ background: "#2E9BF0", color: "#0F1319" }}
                    >
                      <Plus size={16} />
                    </button>
                  </div>
                )}

                {form.usedInsumos.length > 0 && (
                  <div className="flex flex-col gap-1.5 mt-2">
                    {form.usedInsumos.map((l) => (
                      <div key={l.lineId} className="flex items-center justify-between text-xs px-3 py-1.5 rounded-lg" style={{ background: "#1F2733" }}>
                        <span>
                          {l.quantity}x {l.name}
                        </span>
                        <div className="flex items-center gap-2">
                          <span style={{ color: "#B075E5", fontFamily: "ui-monospace, monospace" }}>
                            {currency(l.unitPrice * l.quantity)}
                          </span>
                          <button onClick={() => handleRemoveInsumoFromQuote(l.lineId)} style={muted}>
                            <X size={13} />
                          </button>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>

            {/* Result card */}
            <div className="rounded-xl p-5 mb-5" style={cardStyle}>
              <div className="flex items-baseline justify-between mb-1">
                <span className="text-xs font-medium" style={muted}>
                  Custo de fabricação (por peça)
                </span>
                <span className="text-3xl font-bold tabular-nums" style={{ color: "#2E9BF0", fontFamily: "ui-monospace, monospace" }}>
                  {currency(total)}
                </span>
              </div>
              {piecesPerBatch > 1 && (
                <p className="text-xs mb-3" style={muted}>
                  Impressão inteira: {currency(batchTotal)} ÷ {piecesPerBatch} peças na mesma build plate
                </p>
              )}

              <div className="flex w-full h-3 rounded-full overflow-hidden mb-3" style={{ background: "#0F1319" }}>
                {segments.map((s) =>
                  s.value > 0 ? (
                    <div
                      key={s.key}
                      style={{ width: `${(s.value / grandTotal) * 100}%`, background: SEGMENT_COLORS[s.key] }}
                      title={`${SEGMENT_LABELS[s.key]}: ${currency(s.value)}`}
                    />
                  ) : null
                )}
              </div>

              <div className="grid grid-cols-2 gap-y-1.5 gap-x-4 text-xs">
                {segments.map((s) => (
                  <div key={s.key} className="flex items-center justify-between">
                    <span className="flex items-center gap-1.5" style={muted}>
                      <span className="w-2 h-2 rounded-full inline-block" style={{ background: SEGMENT_COLORS[s.key] }} />
                      {SEGMENT_LABELS[s.key]}
                    </span>
                    <span style={{ color: "#EAF0F6", fontFamily: "ui-monospace, monospace" }}>{currency(s.value)}</span>
                  </div>
                ))}
              </div>

              {editingId && (
                <div className="flex flex-col gap-1 text-xs mt-4 px-3 py-2 rounded-lg" style={{ background: "#1F2733", color: "#8B98AA" }}>
                  <div className="flex items-center justify-between">
                    <span className="flex items-center gap-1.5">
                      <Pencil size={12} />
                      Editando orçamento salvo
                    </span>
                    <button onClick={handleCancelEdit} style={{ color: "#E85D5D" }} className="font-medium">
                      Cancelar
                    </button>
                  </div>
                  {legacyEditWarning && (
                    <p style={{ color: "#E8A33D" }}>
                      ⚠ Este orçamento foi salvo antes da função de edição completa — só o nome pôde ser recuperado.
                    </p>
                  )}
                  {nameChangedWhileEditing && (
                    <p style={{ color: "#2E9BF0" }}>
                      Nome alterado — ao salvar, isso vira um orçamento novo (o original não será sobrescrito).
                    </p>
                  )}
                </div>
              )}

              {editingId && !nameChangedWhileEditing ? (
                <div className="flex gap-2 mt-4">
                  <button
                    onClick={handleUpdateQuote}
                    disabled={!canSave}
                    className="flex-1 rounded-lg py-2.5 text-sm font-semibold flex items-center justify-center gap-2 transition-opacity"
                    style={{
                      background: canSave ? "#2E9BF0" : "#2A3341",
                      color: canSave ? "#0F1319" : "#8B98AA",
                      cursor: canSave ? "pointer" : "not-allowed",
                    }}
                  >
                    <RefreshCw size={16} />
                    {saveStatus || "Atualizar orçamento"}
                  </button>
                  <button
                    onClick={handleOpenCopyPrompt}
                    disabled={!canSave}
                    className="flex-1 rounded-lg py-2.5 text-sm font-semibold flex items-center justify-center gap-2 transition-opacity"
                    style={{
                      background: "#1F2733",
                      border: canSave ? "1px solid #2E9BF0" : "1px solid #2A3341",
                      color: canSave ? "#2E9BF0" : "#8B98AA",
                      cursor: canSave ? "pointer" : "not-allowed",
                    }}
                  >
                    <Save size={16} />
                    Salvar novo
                  </button>
                </div>
              ) : (
                <button
                  onClick={() => handleSaveAsNewQuote()}
                  disabled={!canSave}
                  className="w-full mt-4 rounded-lg py-2.5 text-sm font-semibold flex items-center justify-center gap-2 transition-opacity"
                  style={{
                    background: canSave ? "#2E9BF0" : "#2A3341",
                    color: canSave ? "#0F1319" : "#8B98AA",
                    cursor: canSave ? "pointer" : "not-allowed",
                  }}
                >
                  <Save size={16} />
                  {saveStatus || "Salvar no histórico"}
                </button>
              )}
            </div>

            {/* History */}
            <div>
              <h2 className="text-sm font-semibold mb-3" style={muted}>
                Histórico {quotes.length > 0 && `(${quotes.length})`}
              </h2>
              {!loaded ? (
                <p className="text-xs" style={muted}>
                  Carregando...
                </p>
              ) : quotes.length === 0 ? (
                <p className="text-xs" style={muted}>
                  Nenhum orçamento salvo ainda. Calcule uma peça acima e salve.
                </p>
              ) : (
                <div className="flex flex-col gap-2">
                  {quotes.map((q) => (
                    <div
                      key={q.id}
                      className="rounded-lg px-4 py-3 flex items-center justify-between"
                      style={{ background: "#171D26", border: editingId === q.id ? "1px solid #2E9BF0" : "1px solid #2A3341" }}
                    >
                      <div>
                        <p className="text-sm font-medium">{q.name}</p>
                        <p className="text-xs" style={muted}>
                          {new Date(q.date).toLocaleDateString("pt-BR")}
                        </p>
                      </div>
                      <div className="flex items-center gap-3">
                        <span className="text-sm font-semibold tabular-nums" style={{ color: "#2E9BF0", fontFamily: "ui-monospace, monospace" }}>
                          {currency(q.total)}
                        </span>
                        <button onClick={() => handleEditQuote(q)} style={{ color: "#2E9BF0" }}>
                          <Pencil size={16} />
                        </button>
                        <button onClick={() => handleDelete(q.id)} style={{ color: "#E85D5D" }}>
                          <Trash2 size={16} />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </>
        )}

        {/* ============ ESTOQUE DE PRODUTOS ============ */}
        {activeTab === "estoque" && (
          <>
            <div className="rounded-xl p-4 mb-5" style={cardStyle}>
              <label className="text-xs font-medium block mb-1.5" style={muted}>
                Cadastrar peça no estoque
              </label>

              {quotes.length > 0 && (
                <div className="mb-3">
                  <label className="text-xs block mb-1" style={muted}>
                    Buscar peça já orçada (preenche nome, categoria, material e custo de fabricação)
                  </label>
                  <select
                    defaultValue=""
                    onChange={(e) => {
                      const quote = quotes.find((q) => q.id === e.target.value);
                      if (quote) {
                        setNewStock({
                          ...newStock,
                          name: quote.name,
                          category: quote.category || "",
                          material: quote.material || "",
                          unitCost: String(quote.total.toFixed(2)),
                          usedInsumos: quote.usedInsumos || [],
                          usedFilaments: quote.usedFilaments || [],
                        });
                      }
                    }}
                    className="w-full rounded-lg px-3 py-2.5 text-sm outline-none"
                    style={inputStyle}
                  >
                    <option value="">Selecionar do histórico de orçamentos...</option>
                    {quotes.map((q) => (
                      <option key={q.id} value={q.id}>
                        {q.name} — {currency(q.total)} ({new Date(q.date).toLocaleDateString("pt-BR")})
                      </option>
                    ))}
                  </select>
                </div>
              )}

              {((newStock.usedInsumos && newStock.usedInsumos.length > 0) ||
                (newStock.usedFilaments && newStock.usedFilaments.length > 0)) && (
                <div className="mb-3 px-3 py-2 rounded-lg text-xs" style={{ background: "#1F2733" }}>
                  <p className="mb-1" style={muted}>
                    Será descontado do estoque de materiais (por unidade impressa):
                  </p>
                  {newStock.usedFilaments?.map((l) => (
                    <p key={l.lineId || l.filamentId} style={{ color: "#2E9BF0" }}>
                      {l.grams}g · {l.brand} {l.color}
                    </p>
                  ))}
                  {newStock.usedInsumos?.map((l) => (
                    <p key={l.lineId || l.insumoId} style={{ color: "#B075E5" }}>
                      {l.quantity}x {l.name}
                    </p>
                  ))}
                </div>
              )}

              <label className="text-xs block mb-1" style={muted}>
                Nome da peça
              </label>
              <input
                value={newStock.name}
                onChange={(e) => setNewStock({ ...newStock, name: e.target.value })}
                placeholder="Ex: Chaveiro Border Collie"
                className="w-full rounded-lg px-3 py-2.5 text-sm mb-3 outline-none"
                style={inputStyle}
              />

              <div className="grid grid-cols-2 gap-2 mb-3">
                <SuggestField
                  label="Categoria"
                  value={newStock.category}
                  onChange={(v) => setNewStock({ ...newStock, category: v })}
                  placeholder="Ex: Decoração"
                  inputStyle={inputStyle}
                  options={existingCategories}
                />
                <SuggestField
                  label="Material"
                  value={newStock.material}
                  onChange={(v) => setNewStock({ ...newStock, material: v })}
                  placeholder="Ex: PLA"
                  inputStyle={inputStyle}
                  options={existingMaterials}
                />
              </div>

              <div className="grid grid-cols-2 gap-2 mb-3">
                <div>
                  <label className="text-xs block mb-1" style={muted}>
                    Impressas
                  </label>
                  <input
                    type="number"
                    value={newStock.printed}
                    onChange={(e) => setNewStock({ ...newStock, printed: e.target.value })}
                    className="w-full rounded-lg px-2 py-2 text-sm outline-none"
                    style={inputStyle}
                  />
                </div>
                <div>
                  <label className="text-xs block mb-1" style={muted}>
                    Custo de fabricação
                  </label>
                  <input
                    type="number"
                    value={newStock.unitCost}
                    onChange={(e) => setNewStock({ ...newStock, unitCost: e.target.value })}
                    className="w-full rounded-lg px-2 py-2 text-sm outline-none"
                    style={inputStyle}
                  />
                </div>
                <div>
                  <label className="text-xs block mb-1" style={muted}>
                    Margem de lucro (%)
                  </label>
                  <input
                    type="number"
                    value={newStock.marginPercent}
                    onChange={(e) => setNewStock({ ...newStock, marginPercent: e.target.value })}
                    className="w-full rounded-lg px-2 py-2 text-sm outline-none"
                    style={inputStyle}
                  />
                </div>
                <div>
                  <label className="text-xs block mb-1" style={muted}>
                    Preço de venda (final)
                  </label>
                  <input
                    type="number"
                    value={newStock.unitPrice}
                    onChange={(e) => setNewStock({ ...newStock, unitPrice: e.target.value })}
                    className="w-full rounded-lg px-2 py-2 text-sm outline-none"
                    style={inputStyle}
                  />
                </div>
              </div>

              {parseFloat(newStock.unitCost) > 0 && (
                <div className="flex items-center justify-between text-xs mb-3 px-3 py-2 rounded-lg" style={{ background: "#1F2733" }}>
                  <span style={muted}>Preço sugerido (custo + margem):</span>
                  <div className="flex items-center gap-2">
                    <span style={{ color: "#43C97E", fontFamily: "ui-monospace, monospace" }}>
                      {currency((parseFloat(newStock.unitCost) || 0) * (1 + (parseFloat(newStock.marginPercent) || 0) / 100))}
                    </span>
                    <button
                      onClick={() =>
                        setNewStock({
                          ...newStock,
                          unitPrice: (
                            (parseFloat(newStock.unitCost) || 0) *
                            (1 + (parseFloat(newStock.marginPercent) || 0) / 100)
                          ).toFixed(2),
                        })
                      }
                      className="text-xs font-medium px-2 py-1 rounded"
                      style={{ background: "#2E9BF0", color: "#0F1319" }}
                    >
                      Usar
                    </button>
                  </div>
                </div>
              )}

              {parseFloat(newStock.unitCost) > 0 && parseFloat(newStock.unitPrice) > 0 && (
                <div className="flex items-center justify-between text-xs mb-3 px-3 py-2 rounded-lg" style={{ background: "#1F2733" }}>
                  <span style={muted}>Lucro por peça (no preço de venda final):</span>
                  <span
                    style={{
                      color: (parseFloat(newStock.unitPrice) || 0) - (parseFloat(newStock.unitCost) || 0) >= 0 ? "#43C97E" : "#E85D5D",
                      fontFamily: "ui-monospace, monospace",
                    }}
                  >
                    {currency((parseFloat(newStock.unitPrice) || 0) - (parseFloat(newStock.unitCost) || 0))}
                    {" ("}
                    {(parseFloat(newStock.unitPrice) || 0) - (parseFloat(newStock.unitCost) || 0) >= 0 ? "+" : ""}
                    {((((parseFloat(newStock.unitPrice) || 0) - (parseFloat(newStock.unitCost) || 0)) / (parseFloat(newStock.unitCost) || 1)) * 100).toFixed(0)}
                    {"%)"}
                  </span>
                </div>
              )}

              <button
                onClick={handleAddStockItem}
                className="w-full rounded-lg py-2.5 text-sm font-semibold flex items-center justify-center gap-2"
                style={{ background: "#2E9BF0", color: "#0F1319" }}
              >
                <Plus size={16} />
                Adicionar ao estoque
              </button>
              <p className="text-xs mt-2" style={muted}>
                O custo vem do orçamento (já com índice de falhas). A margem calcula um preço sugerido; o preço de
                venda final é o valor real usado para calcular o lucro.
              </p>
            </div>

            {stockItems.length > 0 && (
              <div className="grid grid-cols-2 gap-3 mb-5">
                <div className="rounded-xl p-4" style={cardStyle}>
                  <p className="text-xs mb-1" style={muted}>
                    Peças em estoque
                  </p>
                  <p className="text-xl font-bold" style={{ fontFamily: "ui-monospace, monospace" }}>
                    {stockItems.reduce((s, i) => s + (i.printed - i.sold), 0)}
                  </p>
                </div>
                <div className="rounded-xl p-4" style={cardStyle}>
                  <p className="text-xs mb-1" style={muted}>
                    Lucro total
                  </p>
                  <p className="text-xl font-bold" style={{ color: "#43C97E", fontFamily: "ui-monospace, monospace" }}>
                    {currency(stockItems.reduce((s, i) => s + i.sold * (i.unitPrice - i.unitCost), 0))}
                  </p>
                </div>
              </div>
            )}

            {(() => {
              const readyItems = stockItems
                .filter((i) => i.printed - i.sold > 0)
                .sort((a, b) => (b.printed - b.sold) - (a.printed - a.sold));
              const maxReady = readyItems.reduce((m, i) => Math.max(m, i.printed - i.sold), 0) || 1;
              return (
                <div className="mb-5">
                  <div className="flex items-center justify-between mb-3">
                    <h2 className="text-sm font-semibold flex items-center gap-1.5">
                      <PawPrint size={15} style={{ color: "#2E9BF0" }} />
                      Pronta entrega {readyItems.length > 0 && `(${readyItems.length})`}
                    </h2>
                    <span className="text-xs" style={muted}>
                      Disponível para venda agora
                    </span>
                  </div>

                  {!loaded ? (
                    <p className="text-xs" style={muted}>
                      Carregando...
                    </p>
                  ) : readyItems.length === 0 ? (
                    <div className="rounded-xl p-4 text-xs" style={{ ...cardStyle, ...muted }}>
                      Nenhuma peça pronta para entrega no momento — tudo vendido ou aguardando impressão.
                    </div>
                  ) : (
                    <div className="grid grid-cols-2 gap-3">
                      {readyItems.map((item) => {
                        const inStock = item.printed - item.sold;
                        const isLow = inStock <= 2;
                        const barPercent = Math.max(8, (inStock / maxReady) * 100);
                        return (
                          <div
                            key={item.id}
                            className="rounded-xl p-3.5 flex flex-col justify-between"
                            style={{ background: "#171D26", border: isLow ? "1px solid #E8A33D" : "1px solid #2A3341" }}
                          >
                            <div className="flex items-start justify-between gap-2 mb-2">
                              <p className="text-sm font-semibold leading-snug">{item.name}</p>
                              <span
                                className="shrink-0 text-[10px] font-bold px-2 py-1 rounded-full"
                                style={{
                                  background: isLow ? "rgba(232,163,61,0.15)" : "rgba(67,201,126,0.15)",
                                  color: isLow ? "#E8A33D" : "#43C97E",
                                }}
                              >
                                {isLow ? "estoque baixo" : "disponível"}
                              </span>
                            </div>

                            <div className="flex items-baseline justify-between mb-2">
                              <div className="flex items-baseline gap-1.5">
                                <span
                                  className="text-2xl font-bold tabular-nums"
                                  style={{ color: isLow ? "#E8A33D" : "#EAF0F6", fontFamily: "ui-monospace, monospace" }}
                                >
                                  {inStock}
                                </span>
                                <span className="text-xs" style={muted}>
                                  {inStock === 1 ? "unidade" : "unidades"}
                                </span>
                              </div>
                              {editingStockQtyId !== item.id && (
                                <button onClick={() => handleStartEditStockQty(item)} style={muted} title="Editar quantidade em estoque">
                                  <Pencil size={13} />
                                </button>
                              )}
                            </div>

                            {editingStockQtyId === item.id ? (
                              <div className="flex flex-col gap-1.5 mb-2">
                                <label className="text-[10px]" style={muted}>
                                  Quantidade disponível em estoque
                                </label>
                                <input
                                  type="number"
                                  autoFocus
                                  min="0"
                                  value={editStockQtyValue}
                                  onChange={(e) => setEditStockQtyValue(e.target.value)}
                                  className="w-full rounded-lg px-2 py-1.5 text-xs outline-none"
                                  style={inputStyle}
                                />
                                <p className="text-[10px]" style={muted}>
                                  Isso ajusta as peças impressas registradas (impressas = disponível + vendidas) e
                                  atualiza os gráficos.
                                </p>
                                <div className="flex gap-1.5">
                                  <button
                                    onClick={() => handleSaveEditStockQty(item)}
                                    className="flex-1 rounded-lg py-1.5 text-xs font-semibold"
                                    style={{ background: "#43C97E", color: "#0F1319" }}
                                  >
                                    Salvar
                                  </button>
                                  <button
                                    onClick={handleCancelEditStockQty}
                                    className="rounded-lg px-3 py-1.5 text-xs font-medium"
                                    style={{ background: "#2A3341", color: "#8B98AA" }}
                                  >
                                    Cancelar
                                  </button>
                                </div>
                              </div>
                            ) : (
                              <>
                                <div className="w-full h-1.5 rounded-full overflow-hidden mb-2" style={{ background: "#0F1319" }}>
                                  <div
                                    className="h-full rounded-full"
                                    style={{ width: `${barPercent}%`, background: isLow ? "#E8A33D" : "#2E9BF0" }}
                                  />
                                </div>

                                <div className="flex items-center justify-between mb-2.5">
                                  <div className="flex items-center gap-1 text-xs" style={{ color: "#43C97E", fontFamily: "ui-monospace, monospace" }}>
                                    <DollarSign size={11} />
                                    {currency(item.unitPrice)}
                                  </div>
                                </div>

                                {sellingItemId === item.id ? (
                                  <div className="flex flex-col gap-1.5">
                                    <div className="flex items-center justify-between text-[10px] px-0.5" style={muted}>
                                      <span>Custo: {currency(item.unitCost)}</span>
                                      <span>Cadastrado: {currency(item.unitPrice)}</span>
                                    </div>
                                    <input
                                      type="number"
                                      min="1"
                                      max={inStock}
                                      value={saleQuantity}
                                      onChange={(e) => setSaleQuantity(e.target.value)}
                                      placeholder={`Qtd (máx ${inStock})`}
                                      className="w-full rounded-lg px-2 py-1.5 text-xs outline-none"
                                      style={inputStyle}
                                    />
                                    <div className="relative">
                                      <span className="absolute left-2 top-1/2 -translate-y-1/2 text-[10px]" style={muted}>
                                        R$
                                      </span>
                                      <input
                                        type="number"
                                        autoFocus
                                        value={saleUnitPrice}
                                        onChange={(e) => setSaleUnitPrice(e.target.value)}
                                        placeholder="Valor de venda (com desconto se houver)"
                                        className="w-full rounded-lg pl-7 pr-2 py-1.5 text-xs outline-none"
                                        style={inputStyle}
                                      />
                                    </div>
                                    {parseFloat(saleUnitPrice) > 0 && parseFloat(saleUnitPrice) < item.unitPrice && (
                                      <p className="text-[10px]" style={{ color: "#E8A33D" }}>
                                        Desconto de {currency(item.unitPrice - parseFloat(saleUnitPrice))} em relação ao valor cadastrado.
                                      </p>
                                    )}
                                    {saleUnitPrice !== "" && isFinite(parseFloat(saleUnitPrice)) && (
                                      <div className="flex items-center justify-between text-[10px] px-0.5 py-1 rounded" style={{ background: "#0F1319" }}>
                                        <span style={muted}>Lucro por peça (neste valor):</span>
                                        <span
                                          style={{
                                            color: parseFloat(saleUnitPrice) - item.unitCost >= 0 ? "#43C97E" : "#F2545B",
                                            fontFamily: "ui-monospace, monospace",
                                          }}
                                        >
                                          {currency(parseFloat(saleUnitPrice) - item.unitCost)}
                                          {parseFloat(saleQuantity) > 0 &&
                                            ` · total ${currency((parseFloat(saleUnitPrice) - item.unitCost) * parseFloat(saleQuantity))}`}
                                        </span>
                                      </div>
                                    )}
                                    <div className="flex gap-1.5">
                                      <button
                                        onClick={() => handleRegisterSale(item.id, saleQuantity, saleUnitPrice)}
                                        className="flex-1 rounded-lg py-1.5 text-xs font-semibold"
                                        style={{ background: "#43C97E", color: "#0F1319" }}
                                      >
                                        Confirmar
                                      </button>
                                      <button
                                        onClick={() => {
                                          setSellingItemId(null);
                                          setSaleQuantity("");
                                          setSaleUnitPrice("");
                                        }}
                                        className="rounded-lg px-3 py-1.5 text-xs font-medium"
                                        style={{ background: "#2A3341", color: "#8B98AA" }}
                                      >
                                        Cancelar
                                      </button>
                                    </div>
                                  </div>
                                ) : (
                                  <button
                                    onClick={() => {
                                      setSellingItemId(item.id);
                                      setSaleQuantity("");
                                      setSaleUnitPrice(String(item.unitPrice));
                                    }}
                                    className="w-full rounded-lg py-1.5 text-xs font-semibold flex items-center justify-center gap-1.5"
                                    style={{ background: "#2E9BF0", color: "#0F1319" }}
                                  >
                                    <DollarSign size={12} />
                                    Realizar venda
                                  </button>
                                )}
                              </>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              );
            })()}

            <div className="rounded-xl p-4 mb-5" style={cardStyle}>
              <p className="text-sm font-semibold mb-1 flex items-center gap-1.5">
                <Download size={15} style={{ color: "#2E9BF0" }} />
                Exportar catálogo (.xlsx)
              </p>
              <p className="text-xs mb-3" style={muted}>
                Gera uma planilha no mesmo modelo usado no catálogo, com nome, preço, categoria, material e
                estoque disponível — só das peças que têm unidades em "Pronta entrega" agora.
              </p>
              <button
                onClick={handleExportCatalog}
                className="w-full rounded-lg py-2.5 text-sm font-semibold flex items-center justify-center gap-2"
                style={{ background: "#2E9BF0", color: "#0F1319" }}
              >
                <Download size={16} />
                {catalogExportStatus || "Exportar catálogo"}
              </button>
            </div>

            <div>
              <h2 className="text-sm font-semibold mb-3" style={muted}>
                Peças impressas {stockItems.length > 0 && `(${stockItems.length})`}
              </h2>
              {!loaded ? (
                <p className="text-xs" style={muted}>
                  Carregando...
                </p>
              ) : stockItems.length === 0 ? (
                <p className="text-xs" style={muted}>
                  Nenhuma peça cadastrada ainda. Adicione uma acima.
                </p>
              ) : (
                <div className="flex flex-col gap-3">
                  {stockItems.map((item) => {
                    const inStock = item.printed - item.sold;
                    const profit = item.sold * (item.unitPrice - item.unitCost);
                    const idleCapital = inStock * item.unitCost;
                    const suggestedPrice = item.unitCost * (1 + (item.marginPercent || 0) / 100);
                    return (
                      <div key={item.id} className="rounded-xl p-4" style={cardStyle}>
                        <div className="flex items-center justify-between mb-3">
                          <p className="text-sm font-semibold">{item.name}</p>
                          <button onClick={() => handleDeleteStockItem(item.id)} style={{ color: "#E85D5D" }}>
                            <Trash2 size={16} />
                          </button>
                        </div>

                        <div className="grid grid-cols-2 gap-2 mb-2">
                          <SuggestField
                            label="Categoria"
                            value={item.category || ""}
                            onChange={(v) => handleUpdateStockTextField(item.id, "category", v)}
                            placeholder="Ex: Decoração"
                            inputStyle={inputStyle}
                            options={existingCategories}
                            small
                          />
                          <SuggestField
                            label="Material"
                            value={item.material || ""}
                            onChange={(v) => handleUpdateStockTextField(item.id, "material", v)}
                            placeholder="Ex: PLA"
                            inputStyle={inputStyle}
                            options={existingMaterials}
                            small
                          />
                        </div>

                        <div className="grid grid-cols-4 gap-2 mb-2">
                          <StockField label="Impressas" value={item.printed} onChange={(v) => handleUpdateStockPrinted(item.id, v)} inputStyle={inputStyle} />
                          <StockField label="Vendidas" value={item.sold} onChange={(v) => handleUpdateStockField(item.id, "sold", v)} inputStyle={inputStyle} />
                          <StockField label="Custo fabr." value={item.unitCost} onChange={(v) => handleUpdateStockField(item.id, "unitCost", v)} inputStyle={inputStyle} />
                          <StockField label="Margem %" value={item.marginPercent || 0} onChange={(v) => handleUpdateStockField(item.id, "marginPercent", v)} inputStyle={inputStyle} />
                        </div>

                        <div className="flex items-center gap-2 mb-3">
                          <div className="flex-1 flex items-center justify-between text-xs px-3 py-1.5 rounded-lg" style={{ background: "#1F2733" }}>
                            <span style={muted}>Sugerido:</span>
                            <span style={{ color: "#43C97E", fontFamily: "ui-monospace, monospace" }}>{currency(suggestedPrice)}</span>
                          </div>
                          <div className="flex-1">
                            <StockField label="Venda final" value={item.unitPrice} onChange={(v) => handleUpdateStockField(item.id, "unitPrice", v)} inputStyle={inputStyle} />
                          </div>
                        </div>

                        <div className="flex items-center justify-between text-xs pt-2" style={{ borderTop: "1px solid #2A3341" }}>
                          <span style={muted}>
                            Em estoque: <b style={{ color: "#EAF0F6" }}>{inStock}</b>
                          </span>
                          <span style={muted}>
                            Parado em estoque: <b style={{ color: "#EAF0F6", fontFamily: "ui-monospace, monospace" }}>{currency(idleCapital)}</b>
                          </span>
                          <span className="flex items-center gap-1" style={muted}>
                            <DollarSign size={12} />
                            Lucro:{" "}
                            <b style={{ color: profit >= 0 ? "#43C97E" : "#E85D5D", fontFamily: "ui-monospace, monospace" }}>{currency(profit)}</b>
                          </span>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          </>
        )}

        {/* ============ VENDAS & CAIXA ============ */}
        {activeTab === "vendas" && (
          <>
            {/* Fluxo de caixa */}
            <div className="rounded-xl p-5 mb-4" style={cardStyle}>
              <div className="flex items-baseline justify-between mb-1">
                <span className="text-xs font-medium flex items-center gap-1.5" style={muted}>
                  <Wallet size={13} />
                  Saldo do fluxo de caixa
                </span>
                <span
                  className="text-3xl font-bold tabular-nums"
                  style={{ color: netCashFlow >= 0 ? "#43C97E" : "#F2545B", fontFamily: "ui-monospace, monospace" }}
                >
                  {currency(netCashFlow)}
                </span>
              </div>
              <p className="text-xs" style={muted}>
                {netCashFlow >= 0
                  ? "Você já recuperou tudo que investiu em materiais e produção, e está no lucro."
                  : "Ainda recuperando o valor investido em materiais e peças produzidas."}
              </p>
            </div>

            <div className="grid grid-cols-2 gap-3 mb-3">
              <div className="rounded-xl p-4" style={cardStyle}>
                <p className="text-xs mb-1" style={muted}>
                  Estoque de matéria-prima
                </p>
                <p className="text-lg font-bold" style={{ fontFamily: "ui-monospace, monospace" }}>
                  {currency(materialsInventoryValue)}
                </p>
                <p className="text-[10px] mt-0.5" style={muted}>
                  Filamentos + insumos não usados
                </p>
              </div>
              <div className="rounded-xl p-4" style={cardStyle}>
                <p className="text-xs mb-1" style={muted}>
                  Custo de produção
                </p>
                <p className="text-lg font-bold" style={{ fontFamily: "ui-monospace, monospace" }}>
                  {currency(producedGoodsCost)}
                </p>
                <p className="text-[10px] mt-0.5" style={muted}>
                  Todas as peças já impressas
                </p>
              </div>
              <div className="rounded-xl p-4" style={{ background: "#171D26", border: "1px solid #2E9BF0" }}>
                <p className="text-xs mb-1" style={muted}>
                  Total investido
                </p>
                <p className="text-lg font-bold" style={{ color: "#2E9BF0", fontFamily: "ui-monospace, monospace" }}>
                  {currency(totalInvested)}
                </p>
                <p className="text-[10px] mt-0.5" style={muted}>
                  Materiais + produção
                </p>
              </div>
              <div className="rounded-xl p-4" style={{ background: "#171D26", border: "1px solid #43C97E" }}>
                <p className="text-xs mb-1" style={muted}>
                  Recebido em vendas
                </p>
                <p className="text-lg font-bold" style={{ color: "#43C97E", fontFamily: "ui-monospace, monospace" }}>
                  {currency(totalReceived)}
                </p>
                <p className="text-[10px] mt-0.5" style={muted}>
                  {salesHistory.length} {salesHistory.length === 1 ? "venda registrada" : "vendas registradas"}
                </p>
              </div>
            </div>

            {salesHistory.length > 0 && (
              <div className="rounded-xl p-4 mb-5 flex items-center justify-between" style={cardStyle}>
                <span className="text-xs" style={muted}>
                  Lucro realizado nas vendas (receita − custo)
                </span>
                <span
                  className="text-sm font-semibold tabular-nums"
                  style={{ color: totalSalesProfit >= 0 ? "#43C97E" : "#F2545B", fontFamily: "ui-monospace, monospace" }}
                >
                  {currency(totalSalesProfit)}
                </span>
              </div>
            )}

            {/* Gráfico por produto */}
            <h2 className="text-sm font-semibold mb-1 flex items-center gap-1.5" style={muted}>
              <TrendingUp size={15} />
              Produtos: impressas x vendidas x lucro
            </h2>
            <p className="text-[10px] mb-3" style={muted}>
              Histórico acumulado de todos os tempos — não some quando você vende, edita ou remove uma peça do
              estoque.
            </p>
            {!loaded ? (
              <p className="text-xs mb-5" style={muted}>
                Carregando...
              </p>
            ) : productChartData.length === 0 ? (
              <p className="text-xs mb-5" style={muted}>
                Nenhuma peça cadastrada no estoque ainda. Cadastre peças na aba "Estoque de Produtos" para ver o
                gráfico.
              </p>
            ) : (
              <div className="rounded-xl p-4 mb-5" style={cardStyle}>
                <ResponsiveContainer width="100%" height={Math.max(260, productChartData.length * 55)}>
                  <ComposedChart data={productChartData} layout="vertical" margin={{ top: 4, right: 12, left: 4, bottom: 4 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#2A3341" horizontal={false} />
                    <XAxis type="number" tick={{ fill: "#8B98AA", fontSize: 10 }} allowDecimals={false} axisLine={{ stroke: "#2A3341" }} />
                    <YAxis
                      type="category"
                      dataKey="name"
                      tick={{ fill: "#EAF0F6", fontSize: 11 }}
                      width={110}
                      axisLine={{ stroke: "#2A3341" }}
                    />
                    <Tooltip content={<ChartTooltip />} cursor={{ fill: "rgba(255,255,255,0.03)" }} />
                    <Legend wrapperStyle={{ fontSize: 11, color: "#8B98AA" }} />
                    <Bar dataKey="Impressas" fill="#2E9BF0" radius={[0, 4, 4, 0]} barSize={12} />
                    <Bar dataKey="Vendidas" fill="#43C97E" radius={[0, 4, 4, 0]} barSize={12} />
                  </ComposedChart>
                </ResponsiveContainer>

                <div className="mt-4 pt-4" style={{ borderTop: "1px solid #2A3341" }}>
                  <p className="text-xs font-medium mb-2" style={muted}>
                    Lucro por produto (receita − custo das vendas)
                  </p>
                  <ResponsiveContainer width="100%" height={Math.max(200, productChartData.length * 40)}>
                    <ComposedChart data={productChartData} layout="vertical" margin={{ top: 4, right: 12, left: 4, bottom: 4 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#2A3341" horizontal={false} />
                      <XAxis type="number" tick={{ fill: "#8B98AA", fontSize: 10 }} tickFormatter={(v) => currency(v)} axisLine={{ stroke: "#2A3341" }} />
                      <YAxis type="category" dataKey="name" tick={{ fill: "#EAF0F6", fontSize: 11 }} width={110} axisLine={{ stroke: "#2A3341" }} />
                      <Tooltip content={<ChartTooltip />} cursor={{ fill: "rgba(255,255,255,0.03)" }} />
                      <Bar dataKey="Lucro" radius={[0, 4, 4, 0]} barSize={14}>
                        {productChartData.map((d, idx) => (
                          <Cell key={idx} fill={d.Lucro >= 0 ? "#43C97E" : "#F2545B"} />
                        ))}
                      </Bar>
                    </ComposedChart>
                  </ResponsiveContainer>
                </div>
              </div>
            )}

            {/* Histórico de vendas */}
            <h2 className="text-sm font-semibold mb-3 flex items-center gap-1.5" style={muted}>
              <History size={15} />
              Histórico de vendas {salesHistory.length > 0 && `(${salesHistory.length})`}
            </h2>
            {!loaded ? (
              <p className="text-xs" style={muted}>
                Carregando...
              </p>
            ) : salesHistory.length === 0 ? (
              <p className="text-xs" style={muted}>
                Nenhuma venda registrada ainda. Registre vendas na aba "Estoque de Produtos", na seção Pronta
                entrega.
              </p>
            ) : (
              <div className="flex flex-col gap-2">
                {salesHistory.map((s) => {
                  const discounted = s.catalogUnitPrice != null && s.unitPrice < s.catalogUnitPrice;
                  const isEditing = editingSaleId === s.id;
                  return (
                    <div key={s.id} className="rounded-lg px-4 py-3" style={{ background: "#171D26", border: isEditing ? "1px solid #2E9BF0" : "1px solid #2A3341" }}>
                      <div className="flex items-center justify-between mb-1">
                        <p className="text-sm font-medium">{s.itemName}</p>
                        {!isEditing && (
                          <div className="flex items-center gap-3">
                            <span className="text-sm font-semibold tabular-nums" style={{ color: "#43C97E", fontFamily: "ui-monospace, monospace" }}>
                              {currency(s.revenue)}
                            </span>
                            <button onClick={() => handleStartEditSale(s)} style={{ color: "#2E9BF0" }} title="Editar venda">
                              <Pencil size={14} />
                            </button>
                            <button onClick={() => handleDeleteSale(s.id)} style={{ color: "#E85D5D" }} title="Desfazer venda">
                              <Trash2 size={15} />
                            </button>
                          </div>
                        )}
                      </div>

                      {isEditing ? (
                        <div className="flex flex-col gap-2 mt-1">
                          <div className="grid grid-cols-2 gap-2">
                            <div>
                              <label className="text-[10px] block mb-1" style={muted}>
                                Quantidade
                              </label>
                              <input
                                type="number"
                                value={editSaleForm.quantity}
                                onChange={(e) => setEditSaleForm({ ...editSaleForm, quantity: e.target.value })}
                                className="w-full rounded-lg px-2 py-1.5 text-xs outline-none"
                                style={inputStyle}
                              />
                            </div>
                            <div>
                              <label className="text-[10px] block mb-1" style={muted}>
                                Valor unitário (R$)
                              </label>
                              <input
                                type="number"
                                value={editSaleForm.unitPrice}
                                onChange={(e) => setEditSaleForm({ ...editSaleForm, unitPrice: e.target.value })}
                                className="w-full rounded-lg px-2 py-1.5 text-xs outline-none"
                                style={inputStyle}
                              />
                            </div>
                          </div>
                          <div className="flex gap-2">
                            <button
                              onClick={() => handleSaveEditSale(s.id)}
                              className="flex-1 rounded-lg py-1.5 text-xs font-semibold"
                              style={{ background: "#43C97E", color: "#0F1319" }}
                            >
                              Salvar
                            </button>
                            <button
                              onClick={handleCancelEditSale}
                              className="rounded-lg px-3 py-1.5 text-xs font-medium"
                              style={{ background: "#2A3341", color: "#8B98AA" }}
                            >
                              Cancelar
                            </button>
                          </div>
                        </div>
                      ) : (
                        <>
                          <div className="flex items-center justify-between text-xs" style={muted}>
                            <span>
                              {s.quantity}x {currency(s.unitPrice)}/un
                              {discounted && (
                                <span style={{ color: "#E8A33D" }}> · com desconto (cadastrado {currency(s.catalogUnitPrice)})</span>
                              )}
                            </span>
                            <span>{new Date(s.date).toLocaleDateString("pt-BR")}</span>
                          </div>
                          <div className="flex items-center justify-between text-[10px] mt-1" style={muted}>
                            <span>Custo: {currency(s.cost)}</span>
                            <span style={{ color: s.profit >= 0 ? "#43C97E" : "#F2545B" }}>Lucro: {currency(s.profit)}</span>
                          </div>
                        </>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </>
        )}

        {/* ============ MATERIAIS (filamentos + insumos: catálogo e estoque) ============ */}
        {activeTab === "materiais" && (
          <>
            <p className="text-xs mb-4" style={muted}>
              O estoque diminui automaticamente quando você registra peças impressas na aba "Estoque de Produtos".
              Use os campos de estoque abaixo para registrar novas compras ou corrigir a quantidade manualmente.
            </p>

            {/* --- Sub-abas: Filamentos / Insumos agrupados --- */}
            <div className="flex gap-1.5 mb-4 p-1 rounded-lg" style={{ background: "#171D26", border: "1px solid #2A3341" }}>
              <button
                onClick={() => setMaterialSubTab("filamentos")}
                className="flex-1 py-2 rounded-md text-xs font-semibold flex items-center justify-center gap-1.5"
                style={{
                  background: materialSubTab === "filamentos" ? "#2E9BF0" : "transparent",
                  color: materialSubTab === "filamentos" ? "#0F1319" : "#8B98AA",
                }}
              >
                <Package size={14} />
                Filamentos {filamentsCatalog.length > 0 && `(${filamentsCatalog.length})`}
              </button>
              <button
                onClick={() => setMaterialSubTab("insumos")}
                className="flex-1 py-2 rounded-md text-xs font-semibold flex items-center justify-center gap-1.5"
                style={{
                  background: materialSubTab === "insumos" ? "#2E9BF0" : "transparent",
                  color: materialSubTab === "insumos" ? "#0F1319" : "#8B98AA",
                }}
              >
                <Boxes size={14} />
                Insumos {insumosCatalog.length > 0 && `(${insumosCatalog.length})`}
              </button>
            </div>

            {/* --- Filamentos --- */}
            {materialSubTab === "filamentos" && (
              <>
                <div className="rounded-xl p-4 mb-3" style={cardStyle}>
                  {editingFilamentId && (
                    <div className="flex items-center justify-between text-xs mb-2 px-3 py-2 rounded-lg" style={{ background: "#1F2733", color: "#8B98AA" }}>
                      <span className="flex items-center gap-1.5">
                        <Pencil size={12} />
                        Editando filamento
                      </span>
                      <button onClick={handleCancelEditFilament} style={{ color: "#E85D5D" }} className="font-medium">
                        Cancelar
                      </button>
                    </div>
                  )}

                  <div className="grid grid-cols-6 gap-2 mb-2">
                    <input
                      value={newFilament.brand}
                      onChange={(e) => setNewFilament({ ...newFilament, brand: e.target.value })}
                      placeholder="Marca"
                      className="col-span-2 rounded-lg px-3 py-2 text-sm outline-none"
                      style={inputStyle}
                    />
                    <input
                      value={newFilament.color}
                      onChange={(e) => setNewFilament({ ...newFilament, color: e.target.value })}
                      placeholder="Cor"
                      className="col-span-2 rounded-lg px-3 py-2 text-sm outline-none"
                      style={inputStyle}
                    />
                    <input
                      type="number"
                      value={newFilament.costPerKg}
                      onChange={(e) => setNewFilament({ ...newFilament, costPerKg: e.target.value })}
                      placeholder="R$/kg"
                      className="col-span-2 rounded-lg px-3 py-2 text-sm outline-none"
                      style={inputStyle}
                    />
                  </div>

                  {!editingFilamentId && (
                    <input
                      type="number"
                      value={newFilament.stockGrams}
                      onChange={(e) => setNewFilament({ ...newFilament, stockGrams: e.target.value })}
                      placeholder="Gramas em estoque (ex: 1000 para um rolo novo)"
                      className="w-full rounded-lg px-3 py-2 text-sm outline-none mb-2"
                      style={inputStyle}
                    />
                  )}

                  <button
                    onClick={handleAddFilamentToCatalog}
                    className="w-full rounded-lg py-2 text-sm font-semibold flex items-center justify-center gap-2"
                    style={{ background: "#2E9BF0", color: "#0F1319" }}
                  >
                    {editingFilamentId ? <RefreshCw size={16} /> : <Plus size={16} />}
                    {editingFilamentId ? "Salvar alterações" : "Adicionar filamento"}
                  </button>
                </div>

                {filamentsCatalog.length === 0 ? (
                  <p className="text-xs mb-2" style={muted}>
                    Nenhum filamento cadastrado ainda.
                  </p>
                ) : (
                  <div className="rounded-xl overflow-hidden mb-2" style={{ border: "1px solid #2A3341" }}>
                    <div className="overflow-x-auto">
                      <table className="w-full text-xs" style={{ borderCollapse: "collapse" }}>
                        <thead>
                          <tr style={{ background: "#1F2733" }}>
                            <th className="text-left px-3 py-2 font-medium whitespace-nowrap" style={muted}>
                              Filamento
                            </th>
                            <th className="text-right px-3 py-2 font-medium whitespace-nowrap" style={muted}>
                              R$/kg
                            </th>
                            <th className="text-right px-3 py-2 font-medium whitespace-nowrap" style={muted}>
                              Estoque (g)
                            </th>
                            <th className="text-center px-3 py-2 font-medium whitespace-nowrap" style={muted}>
                              Ações
                            </th>
                          </tr>
                        </thead>
                        <tbody>
                          {filamentsCatalog.map((f, idx) => {
                            const stock = f.stockQuantity || 0;
                            const isLow = stock <= 0;
                            return (
                              <tr key={f.id} style={{ background: idx % 2 === 0 ? "#12161D" : "#171D26", borderTop: "1px solid #2A3341" }}>
                                <td className="px-3 py-2 whitespace-nowrap">
                                  <p className="font-semibold" style={{ color: isLow ? "#F2545B" : "#EAF0F6" }}>
                                    {f.brand} — {f.color}
                                  </p>
                                  {isLow && (
                                    <p className="text-[10px]" style={{ color: "#F2545B" }}>
                                      ⚠ estoque zerado
                                    </p>
                                  )}
                                </td>
                                <td className="px-3 py-2 text-right tabular-nums" style={{ fontFamily: "ui-monospace, monospace", color: "#8B98AA" }}>
                                  {currency(f.costPerKg)}
                                </td>
                                <td className="px-3 py-2">
                                  <div className="flex items-center justify-end gap-1">
                                    <input
                                      type="number"
                                      value={stock}
                                      onChange={(e) => handleSetFilamentStock(f.id, e.target.value)}
                                      className="w-16 rounded px-1.5 py-1 text-xs text-right outline-none"
                                      style={{ ...inputStyle, color: isLow ? "#F2545B" : "#EAF0F6" }}
                                    />
                                    <button
                                      onClick={() => handleAdjustFilamentStock(f.id, 100)}
                                      className="px-1.5 py-1 rounded text-[10px] font-medium whitespace-nowrap"
                                      style={{ background: "#2A3341", color: "#EAF0F6" }}
                                    >
                                      +100
                                    </button>
                                    <button
                                      onClick={() => handleAdjustFilamentStock(f.id, 1000)}
                                      className="px-1.5 py-1 rounded text-[10px] font-medium whitespace-nowrap"
                                      style={{ background: "#2A3341", color: "#EAF0F6" }}
                                    >
                                      +1k
                                    </button>
                                  </div>
                                </td>
                                <td className="px-3 py-2">
                                  <div className="flex items-center justify-center gap-2.5">
                                    <button onClick={() => handleEditFilament(f)} style={{ color: "#2E9BF0" }}>
                                      <Pencil size={13} />
                                    </button>
                                    <button onClick={() => handleDeleteFilamentFromCatalog(f.id)} style={{ color: "#E85D5D" }}>
                                      <Trash2 size={13} />
                                    </button>
                                  </div>
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </>
            )}

            {/* --- Insumos --- */}
            {materialSubTab === "insumos" && (
              <>
                <div className="rounded-xl p-4 mb-3" style={cardStyle}>
                  {editingInsumoId && (
                    <div className="flex items-center justify-between text-xs mb-2 px-3 py-2 rounded-lg" style={{ background: "#1F2733", color: "#8B98AA" }}>
                      <span className="flex items-center gap-1.5">
                        <Pencil size={12} />
                        Editando insumo
                      </span>
                      <button onClick={handleCancelEditInsumo} style={{ color: "#E85D5D" }} className="font-medium">
                        Cancelar
                      </button>
                    </div>
                  )}

                  <div className="grid grid-cols-6 gap-2 mb-1">
                    <input
                      value={newInsumo.name}
                      onChange={(e) => setNewInsumo({ ...newInsumo, name: e.target.value })}
                      placeholder="Nome (ex: Ímã 5mm)"
                      className="col-span-3 rounded-lg px-3 py-2 text-sm outline-none"
                      style={inputStyle}
                    />
                    <input
                      type="number"
                      value={newInsumo.totalPaid}
                      onChange={(e) => setNewInsumo({ ...newInsumo, totalPaid: e.target.value })}
                      placeholder="Valor pago (R$)"
                      className="col-span-2 rounded-lg px-3 py-2 text-sm outline-none"
                      style={inputStyle}
                    />
                    <input
                      type="number"
                      value={newInsumo.quantity}
                      onChange={(e) => setNewInsumo({ ...newInsumo, quantity: e.target.value })}
                      placeholder="Qtd. unid."
                      className="col-span-1 rounded-lg px-2 py-2 text-sm outline-none"
                      style={inputStyle}
                    />
                  </div>

                  {parseFloat(newInsumo.totalPaid) > 0 && parseFloat(newInsumo.quantity) > 0 && (
                    <p className="text-xs mb-2" style={muted}>
                      Preço por unidade:{" "}
                      <b style={{ color: "#B075E5", fontFamily: "ui-monospace, monospace" }}>
                        {currency(parseFloat(newInsumo.totalPaid) / parseFloat(newInsumo.quantity))}
                      </b>
                    </p>
                  )}

                  <button
                    onClick={handleAddInsumoToCatalog}
                    className="w-full rounded-lg py-2 text-sm font-semibold flex items-center justify-center gap-2"
                    style={{ background: "#2E9BF0", color: "#0F1319" }}
                  >
                    {editingInsumoId ? <RefreshCw size={16} /> : <Plus size={16} />}
                    {editingInsumoId ? "Salvar alterações" : "Adicionar insumo"}
                  </button>
                </div>

                {insumosCatalog.length === 0 ? (
                  <p className="text-xs" style={muted}>
                    Nenhum insumo cadastrado ainda. Adicione parafusos, ímãs, LEDs, tintas, embalagens etc.
                  </p>
                ) : (
                  <div className="rounded-xl overflow-hidden" style={{ border: "1px solid #2A3341" }}>
                    <div className="overflow-x-auto">
                      <table className="w-full text-xs" style={{ borderCollapse: "collapse" }}>
                        <thead>
                          <tr style={{ background: "#1F2733" }}>
                            <th className="text-left px-3 py-2 font-medium whitespace-nowrap" style={muted}>
                              Insumo
                            </th>
                            <th className="text-right px-3 py-2 font-medium whitespace-nowrap" style={muted}>
                              R$/un
                            </th>
                            <th className="text-right px-3 py-2 font-medium whitespace-nowrap" style={muted}>
                              Estoque
                            </th>
                            <th className="text-center px-3 py-2 font-medium whitespace-nowrap" style={muted}>
                              Ações
                            </th>
                          </tr>
                        </thead>
                        <tbody>
                          {insumosCatalog.map((i, idx) => {
                            const stock = i.stockQuantity || 0;
                            const isLow = stock <= 0;
                            return (
                              <tr key={i.id} style={{ background: idx % 2 === 0 ? "#12161D" : "#171D26", borderTop: "1px solid #2A3341" }}>
                                <td className="px-3 py-2 whitespace-nowrap">
                                  <p className="font-semibold" style={{ color: isLow ? "#F2545B" : "#EAF0F6" }}>
                                    {i.name}
                                  </p>
                                  {isLow && (
                                    <p className="text-[10px]" style={{ color: "#F2545B" }}>
                                      ⚠ estoque zerado
                                    </p>
                                  )}
                                </td>
                                <td className="px-3 py-2 text-right tabular-nums whitespace-nowrap" style={{ fontFamily: "ui-monospace, monospace", color: "#8B98AA" }}>
                                  {currency(i.unitPrice)}/{i.unit}
                                </td>
                                <td className="px-3 py-2">
                                  <div className="flex items-center justify-end gap-1">
                                    <input
                                      type="number"
                                      value={stock}
                                      onChange={(e) => handleSetInsumoStock(i.id, e.target.value)}
                                      className="w-16 rounded px-1.5 py-1 text-xs text-right outline-none"
                                      style={{ ...inputStyle, color: isLow ? "#F2545B" : "#EAF0F6" }}
                                    />
                                    <button
                                      onClick={() => handleAdjustInsumoStock(i.id, 1)}
                                      className="px-1.5 py-1 rounded text-[10px] font-medium whitespace-nowrap"
                                      style={{ background: "#2A3341", color: "#EAF0F6" }}
                                    >
                                      +1
                                    </button>
                                    <button
                                      onClick={() => handleAdjustInsumoStock(i.id, 10)}
                                      className="px-1.5 py-1 rounded text-[10px] font-medium whitespace-nowrap"
                                      style={{ background: "#2A3341", color: "#EAF0F6" }}
                                    >
                                      +10
                                    </button>
                                  </div>
                                </td>
                                <td className="px-3 py-2">
                                  <div className="flex items-center justify-center gap-2.5">
                                    <button onClick={() => handleEditInsumo(i)} style={{ color: "#2E9BF0" }}>
                                      <Pencil size={13} />
                                    </button>
                                    <button onClick={() => handleDeleteInsumoFromCatalog(i.id)} style={{ color: "#E85D5D" }}>
                                      <Trash2 size={13} />
                                    </button>
                                  </div>
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </>
            )}
          </>
        )}

        {/* ============ CUSTOS-BASE ============ */}
        {activeTab === "custos" && (
          <>
            <p className="text-xs mb-4" style={muted}>
              Esses valores alimentam automaticamente o cálculo de custo na aba Orçamento.
            </p>
            <div className="rounded-xl p-4 grid grid-cols-2 gap-3" style={cardStyle}>
              <Field label="Potência impressora (W)" value={settings.printerWattage} onChange={(v) => persistSettings({ ...settings, printerWattage: v })} inputStyle={inputStyle} />
              <Field label="Energia (R$/kWh)" value={settings.energyPricePerKwh} onChange={(v) => persistSettings({ ...settings, energyPricePerKwh: v })} inputStyle={inputStyle} />
              <Field label="Valor da impressora (R$)" value={settings.printerValue} onChange={(v) => persistSettings({ ...settings, printerValue: v })} inputStyle={inputStyle} />
              <Field label="Vida útil (horas)" value={settings.printerLifetimeHours} onChange={(v) => persistSettings({ ...settings, printerLifetimeHours: v })} inputStyle={inputStyle} />
              <Field label="Mão de obra (R$/hora)" value={settings.laborRatePerHour} onChange={(v) => persistSettings({ ...settings, laborRatePerHour: v })} inputStyle={inputStyle} />
              <div className="col-span-2">
                <Field label="Índice de falhas padrão (%)" value={settings.failureRatePercent} onChange={(v) => persistSettings({ ...settings, failureRatePercent: v })} inputStyle={inputStyle} />
              </div>
            </div>
          </>
        )}
      </div>

      {/* ============ MODAL: NOME DO NOVO ORÇAMENTO (SALVAR COMO CÓPIA) ============ */}
      {copyPromptOpen && (
        <div
          className="fixed inset-0 flex items-center justify-center p-4"
          style={{ background: "rgba(15,19,25,0.75)", zIndex: 50 }}
          onClick={handleCancelCopyPrompt}
        >
          <div
            className="w-full max-w-sm rounded-xl p-5"
            style={{ background: "#171D26", border: "1px solid #2A3341" }}
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between mb-3">
              <h2 className="text-sm font-semibold flex items-center gap-2">
                <Save size={16} style={{ color: "#2E9BF0" }} />
                Salvar como novo orçamento
              </h2>
              <button onClick={handleCancelCopyPrompt} style={muted}>
                <X size={18} />
              </button>
            </div>

            <p className="text-xs mb-3" style={muted}>
              Defina um nome para este novo orçamento. O orçamento original não será alterado.
            </p>

            <label className="text-xs font-medium block mb-1.5" style={muted}>
              Nome da peça
            </label>
            <input
              autoFocus
              value={copyNameInput}
              onChange={(e) => setCopyNameInput(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter") handleConfirmCopyPrompt();
              }}
              placeholder="Ex: Chaveiro Border Collie (variação)"
              className="w-full rounded-lg px-3 py-2.5 text-sm mb-4 outline-none"
              style={inputStyle}
            />

            <div className="flex gap-2">
              <button
                onClick={handleCancelCopyPrompt}
                className="flex-1 rounded-lg py-2.5 text-sm font-medium"
                style={{ background: "#2A3341", color: "#8B98AA" }}
              >
                Cancelar
              </button>
              <button
                onClick={handleConfirmCopyPrompt}
                disabled={!copyNameInput.trim()}
                className="flex-1 rounded-lg py-2.5 text-sm font-semibold flex items-center justify-center gap-2"
                style={{
                  background: copyNameInput.trim() ? "#2E9BF0" : "#2A3341",
                  color: copyNameInput.trim() ? "#0F1319" : "#8B98AA",
                  cursor: copyNameInput.trim() ? "pointer" : "not-allowed",
                }}
              >
                <Save size={16} />
                Salvar
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ============ MODAL DE CONFIGURAÇÕES ============ */}
      {showSettingsModal && (
        <div
          className="fixed inset-0 flex items-center justify-center p-4"
          style={{ background: "rgba(15,19,25,0.75)", zIndex: 50 }}
          onClick={() => setShowSettingsModal(false)}
        >
          <div
            className="w-full max-w-sm rounded-xl p-5"
            style={{ background: "#171D26", border: "1px solid #2A3341" }}
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between mb-4">
              <h2 className="text-sm font-semibold flex items-center gap-2">
                <Settings2 size={16} style={{ color: "#2E9BF0" }} />
                Configurações
              </h2>
              <button onClick={() => setShowSettingsModal(false)} style={muted}>
                <X size={18} />
              </button>
            </div>

            <div className="rounded-lg p-3.5 mb-3" style={{ background: "#1F2733" }}>
              <p className="text-xs font-medium mb-1">Backup de dados</p>
              <p className="text-xs mb-3" style={muted}>
                Exporte um arquivo com todos os dados salvos (custos-base, orçamentos, filamentos, insumos e
                estoque) ou importe um backup anterior.
              </p>

              <button
                onClick={handleExportData}
                className="w-full rounded-lg py-2.5 text-sm font-semibold flex items-center justify-center gap-2 mb-2"
                style={{ background: "#2E9BF0", color: "#0F1319" }}
              >
                <Download size={16} />
                Exportar dados (.json)
              </button>

              <button
                onClick={() => fileInputRef.current?.click()}
                className="w-full rounded-lg py-2.5 text-sm font-semibold flex items-center justify-center gap-2"
                style={{ background: "#2A3341", color: "#EAF0F6" }}
              >
                <Upload size={16} />
                Importar dados (.json)
              </button>
              <input
                ref={fileInputRef}
                type="file"
                accept="application/json"
                onChange={handleImportFileSelected}
                className="hidden"
              />

              {importStatus && (
                <p className="text-xs mt-2 text-center" style={{ color: importStatus.startsWith("Erro") ? "#F2545B" : "#43C97E" }}>
                  {importStatus}
                </p>
              )}

              <p className="text-[10px] mt-3" style={muted}>
                ⚠ Importar um backup substitui os dados atuais do app pelos dados do arquivo.
              </p>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function Field({ label, value, onChange, inputStyle }) {
  return (
    <div>
      <label className="text-xs block mb-1" style={muted}>
        {label}
      </label>
      <input
        type="number"
        value={value}
        onChange={(e) => onChange(parseFloat(e.target.value) || 0)}
        className="w-full rounded-lg px-3 py-2 text-sm outline-none"
        style={inputStyle}
      />
    </div>
  );
}

function IconField({ icon, label, value, onChange, inputStyle, placeholder, type = "number", list }) {
  return (
    <div>
      <label className="text-xs mb-1 flex items-center gap-1.5" style={muted}>
        {icon}
        {label}
      </label>
      <input
        type={type}
        value={value}
        placeholder={placeholder}
        onChange={(e) => onChange(e.target.value)}
        list={list}
        className="w-full rounded-lg px-3 py-2 text-sm outline-none"
        style={inputStyle}
      />
    </div>
  );
}

function SuggestField({ icon, label, value, onChange, options, inputStyle, placeholder, small }) {
  const [isOpen, setIsOpen] = useState(false);
  const safeValue = value || "";
  const filtered = (options || []).filter((o) => o.toLowerCase().includes(safeValue.toLowerCase()));

  return (
    <div className="relative">
      <label className={`${small ? "text-[10px]" : "text-xs"} mb-1 flex items-center gap-1.5`} style={muted}>
        {icon}
        {label}
      </label>
      <input
        type="text"
        value={safeValue}
        placeholder={placeholder}
        onChange={(e) => onChange(e.target.value)}
        onFocus={() => setIsOpen(true)}
        onBlur={() => setTimeout(() => setIsOpen(false), 150)}
        className={`w-full rounded-lg outline-none ${small ? "px-2 py-1.5 text-xs" : "px-3 py-2 text-sm"}`}
        style={inputStyle}
      />
      {isOpen && filtered.length > 0 && (
        <div
          className="absolute z-20 left-0 right-0 mt-1 rounded-lg overflow-hidden max-h-40 overflow-y-auto"
          style={{ background: "#1F2733", border: "1px solid #2A3341" }}
        >
          {filtered.map((opt) => (
            <button
              key={opt}
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => {
                onChange(opt);
                setIsOpen(false);
              }}
              className={`w-full text-left ${small ? "px-2 py-1.5 text-xs" : "px-3 py-2 text-sm"}`}
              style={{ color: "#EAF0F6" }}
            >
              {opt}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

function StockField({ label, value, onChange, inputStyle }) {
  return (
    <div>
      <label className="text-[10px] block mb-1" style={muted}>
        {label}
      </label>
      <input
        type="number"
        value={value}
        onChange={(e) => onChange(e.target.value)}
        className="w-full rounded-lg px-2 py-1.5 text-xs outline-none"
        style={inputStyle}
      />
    </div>
  );
}

function ChartTooltip({ active, payload, label }) {
  if (!active || !payload || !payload.length) return null;
  return (
    <div className="rounded-lg px-3 py-2 text-xs" style={{ background: "#171D26", border: "1px solid #2A3341", color: "#EAF0F6" }}>
      <p className="font-semibold mb-1">{label}</p>
      {payload.map((p) => (
        <p key={p.dataKey} style={{ color: p.color }}>
          {p.name}: {p.dataKey === "Lucro" ? currency(p.value) : p.value}
        </p>
      ))}
    </div>
  );
}

