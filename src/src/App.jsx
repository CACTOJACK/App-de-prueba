import { useState, useEffect, useMemo } from "react";
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend } from "recharts";
import { Cloud, Sun, CloudRain, Plus, Trash2, TrendingUp, Package, RefreshCw, Upload } from "lucide-react";
import * as XLSX from "xlsx";

const LAT = -33.45, LON = -70.66; // Santiago por defecto

function excelDateToISO(value) {
  if (value instanceof Date) {
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, "0");
    const d = String(value.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  if (typeof value === "number") {
    const d = XLSX.SSF.parse_date_code(value);
    const mm = String(d.m).padStart(2, "0");
    const dd = String(d.d).padStart(2, "0");
    return `${d.y}-${mm}-${dd}`;
  }
  if (typeof value === "string") {
    const iso = value.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return value.slice(0, 10);
    const dmy = value.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (dmy) {
      const [, d, m, y] = dmy;
      return `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
    }
  }
  return null;
}

function linReg(points) {
  const n = points.length;
  if (n < 2) return { m: 0, b: points[0]?.y || 0, r2: 0 };
  const sumX = points.reduce((a, p) => a + p.x, 0);
  const sumY = points.reduce((a, p) => a + p.y, 0);
  const sumXY = points.reduce((a, p) => a + p.x * p.y, 0);
  const sumX2 = points.reduce((a, p) => a + p.x * p.x, 0);
  const denom = n * sumX2 - sumX * sumX;
  if (denom === 0) return { m: 0, b: sumY / n, r2: 0 };
  const m = (n * sumXY - sumX * sumY) / denom;
  const b = (sumY - m * sumX) / n;
  const meanY = sumY / n;
  const ssTot = points.reduce((a, p) => a + (p.y - meanY) ** 2, 0);
  const ssRes = points.reduce((a, p) => a + (p.y - (m * p.x + b)) ** 2, 0);
  const r2 = ssTot === 0 ? 0 : 1 - ssRes / ssTot;
  return { m, b, r2 };
}

function WeatherIcon({ code, size = 20 }) {
  if (code <= 1) return <Sun size={size} className="text-amber-500" />;
  if (code <= 3) return <Cloud size={size} className="text-slate-400" />;
  return <CloudRain size={size} className="text-blue-400" />;
}

export default function MoteForecastApp() {
  const [forecast, setForecast] = useState([]);
  const [loadingWeather, setLoadingWeather] = useState(true);
  const [weatherError, setWeatherError] = useState(null);

  const [orders, setOrders] = useState([
    { id: 1, fecha: "2026-01-10", mote: 14, harina: 6 },
    { id: 2, fecha: "2026-01-15", mote: 19, harina: 8 },
    { id: 3, fecha: "2026-01-22", mote: 10, harina: 4.5 },
    { id: 4, fecha: "2026-02-02", mote: 17, harina: 7 },
    { id: 5, fecha: "2026-02-10", mote: 7, harina: 3 },
  ]);
  const [form, setForm] = useState({ fecha: "", mote: "", harina: "" });
  const [pasteText, setPasteText] = useState("");
  const [pasteTempText, setPasteTempText] = useState("");
  const [tempImportInfo, setTempImportInfo] = useState(null);
  const [uploadError, setUploadError] = useState(null);
  const [uploadInfo, setUploadInfo] = useState(null);

  function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    setUploadError(null);
    setUploadInfo(null);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: "array", cellDates: true });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
        const newOrders = [];
        let skipped = 0;
        for (const row of rows) {
          if (!row || row.length < 3) continue;
          const [rawFecha, rawMote, rawHarina] = row;
          const fecha = excelDateToISO(rawFecha);
          const mote = Number(rawMote);
          const harina = Number(rawHarina);
          if (!fecha || isNaN(mote) || isNaN(harina)) {
            skipped++;
            continue;
          }
          newOrders.push({ id: Date.now() + Math.random(), fecha, mote, harina });
        }
        if (newOrders.length === 0) {
          setUploadError("No se encontraron filas válidas (fecha, mote, harina) en el archivo.");
        } else {
          setOrders((o) => [...o, ...newOrders]);
          setUploadInfo(`Se importaron ${newOrders.length} pedidos${skipped ? ` (se omitieron ${skipped} filas no válidas, probablemente encabezados)` : ""}.`);
        }
      } catch (err) {
        setUploadError("No se pudo leer el archivo. Verifica que sea un .xlsx o .xls válido.");
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  }

  const [tempByDate, setTempByDate] = useState({});
  const [loadingHist, setLoadingHist] = useState(false);
  const [histError, setHistError] = useState(null);

  useEffect(() => {
    async function fetchForecast() {
      try {
        setLoadingWeather(true);
        const res = await fetch(
          `https://api.open-meteo.com/v1/forecast?latitude=${LAT}&longitude=${LON}&daily=temperature_2m_max,weathercode&timezone=America%2FSantiago&forecast_days=7`
        );
        if (!res.ok) {
          const t = await res.text().catch(() => "");
          throw new Error(`HTTP ${res.status} ${t}`.slice(0, 200));
        }
        const data = await res.json();
        setForecast(
          data.daily.time.map((t, i) => ({
            fecha: t,
            temp: Math.round(data.daily.temperature_2m_max[i]),
            code: data.daily.weathercode[i],
          }))
        );
        setWeatherError(null);
      } catch (err) {
        setWeatherError("No se pudo conectar al servicio de pronóstico: " + (err?.message || "error desconocido"));
      } finally {
        setLoadingWeather(false);
      }
    }
    fetchForecast();
  }, []);

  async function fetchHistoricalTemps(dates) {
    if (dates.length === 0) return;
    setLoadingHist(true);
    setHistError(null);
    try {
      const sorted = [...dates].sort();
      const startDate = sorted[0];
      const todayISO = new Date().toISOString().slice(0, 10);
      const endDate = sorted[sorted.length - 1] > todayISO ? todayISO : sorted[sorted.length - 1];
      const res = await fetch(
        `https://api.open-meteo.com/v1/archive?latitude=${LAT}&longitude=${LON}&start_date=${startDate}&end_date=${endDate}&daily=temperature_2m_max&timezone=America%2FSantiago`
      );
      if (!res.ok) {
        const t = await res.text().catch(() => "");
        throw new Error(`HTTP ${res.status} ${t}`.slice(0, 200));
      }
      const data = await res.json();
      const map = {};
      data.daily.time.forEach((t, i) => {
        map[t] = Math.round(data.daily.temperature_2m_max[i] * 10) / 10;
      });
      setTempByDate((prev) => ({ ...prev, ...map }));
    } catch (err) {
      setHistError("No se pudo obtener el historial climático: " + (err?.message || "error desconocido"));
    } finally {
      setLoadingHist(false);
    }
  }

  useEffect(() => {
    const missing = orders.map((o) => o.fecha).filter((f) => f && !(f in tempByDate));
    if (missing.length > 0) fetchHistoricalTemps(missing);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [orders]);

  const dataWithTemp = useMemo(
    () =>
      orders
        .filter((o) => tempByDate[o.fecha] !== undefined)
        .map((o) => ({ ...o, temp: tempByDate[o.fecha] })),
    [orders, tempByDate]
  );

  const moteModel = useMemo(
    () => linReg(dataWithTemp.map((h) => ({ x: h.temp, y: Number(h.mote) }))),
    [dataWithTemp]
  );
  const harinaModel = useMemo(
    () => linReg(dataWithTemp.map((h) => ({ x: h.temp, y: Number(h.harina) }))),
    [dataWithTemp]
  );

  function predict(model, temp) {
    return Math.max(0, Math.round((model.m * temp + model.b) * 10) / 10);
  }

  function importTemps() {
    const lines = pasteTempText.split("\n").map((l) => l.trim()).filter(Boolean);
    const map = {};
    let count = 0;
    for (const line of lines) {
      const parts = line.split(/[,;\t]/).map((p) => p.trim());
      if (parts.length < 2) continue;
      let [fecha, temp] = parts;
      if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}$/.test(fecha)) {
        const [d, m, y] = fecha.split(/[\/\-]/);
        fecha = `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
      }
      if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) continue;
      const t = Number(temp.replace(",", "."));
      if (isNaN(t)) continue;
      map[fecha] = t;
      count++;
    }
    if (count > 0) {
      setTempByDate((prev) => ({ ...prev, ...map }));
      setTempImportInfo(`Se cargaron ${count} temperaturas manualmente.`);
      setPasteTempText("");
    } else {
      setTempImportInfo("No se reconoció ninguna fila válida. Usa el formato: fecha, temperatura");
    }
  }

  function addEntry() {
    if (!form.fecha || form.mote === "" || form.harina === "") return;
    setOrders((o) => [
      ...o,
      { id: Date.now(), fecha: form.fecha, mote: Number(form.mote), harina: Number(form.harina) },
    ]);
    setForm({ fecha: "", mote: "", harina: "" });
  }

  function removeEntry(id) {
    setOrders((o) => o.filter((e) => e.id !== id));
  }

  function importPaste() {
    const lines = pasteText.split("\n").map((l) => l.trim()).filter(Boolean);
    const newOrders = [];
    for (const line of lines) {
      const parts = line.split(/[,;\t]/).map((p) => p.trim());
      if (parts.length < 3) continue;
      const [fecha, mote, harina] = parts;
      if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) continue;
      const moteN = Number(mote);
      const harinaN = Number(harina);
      if (isNaN(moteN) || isNaN(harinaN)) continue;
      newOrders.push({ id: Date.now() + Math.random(), fecha, mote: moteN, harina: harinaN });
    }
    if (newOrders.length > 0) {
      setOrders((o) => [...o, ...newOrders]);
      setPasteText("");
    }
  }

  const chartData = dataWithTemp
    .map((h) => ({ temp: h.temp, mote: Number(h.mote), harina: Number(h.harina) }))
    .sort((a, b) => a.temp - b.temp);

  const todayForecast = forecast[0];
  const pendingCount = orders.length - dataWithTemp.length;

  return (
    <div className="min-h-screen bg-gradient-to-b from-amber-50 to-orange-50 p-4 pb-10">
      <div className="max-w-xl mx-auto space-y-5">
        <header className="text-center pt-2">
          <h1 className="text-2xl font-bold text-amber-900">🥤 Mote con Huesillos</h1>
          <p className="text-sm text-amber-700">Pronóstico de pedido según el clima</p>
        </header>

        {todayForecast && (
          <div className="bg-amber-900 text-white rounded-2xl p-5 shadow-lg">
            <div className="flex items-center justify-between mb-3">
              <div>
                <p className="text-amber-200 text-xs uppercase tracking-wide">Hoy</p>
                <p className="text-3xl font-bold">{todayForecast.temp}°C</p>
              </div>
              <WeatherIcon code={todayForecast.code} size={40} />
            </div>
            <div className="grid grid-cols-2 gap-3 mt-2">
              <div className="bg-amber-800/60 rounded-xl p-3">
                <p className="text-amber-200 text-xs">Mote sugerido</p>
                <p className="text-xl font-bold">{predict(moteModel, todayForecast.temp)} kg</p>
              </div>
              <div className="bg-amber-800/60 rounded-xl p-3">
                <p className="text-amber-200 text-xs">Harina sugerida</p>
                <p className="text-xl font-bold">{predict(harinaModel, todayForecast.temp)} kg</p>
              </div>
            </div>
          </div>
        )}

        <div className="bg-white rounded-2xl p-4 shadow-sm border border-amber-100">
          <div className="flex items-center gap-2 mb-2">
            <TrendingUp size={18} className="text-amber-700" />
            <h2 className="font-semibold text-amber-900">Fórmula aprendida</h2>
            {loadingHist && <RefreshCw size={14} className="animate-spin text-amber-400 ml-auto" />}
            {!loadingHist && pendingCount > 0 && (
              <button
                onClick={() => fetchHistoricalTemps(orders.map((o) => o.fecha).filter((f) => f && !(f in tempByDate)))}
                className="ml-auto text-xs text-amber-700 underline"
              >
                Reintentar fechas pendientes
              </button>
            )}
          </div>
          {dataWithTemp.length >= 2 ? (
            <div className="text-sm text-slate-600 space-y-1">
              <p>
                Mote = <span className="font-mono text-amber-800">{moteModel.b.toFixed(1)} + {moteModel.m.toFixed(2)} × °C</span>{" "}
                <span className="text-xs text-slate-400">(R²={moteModel.r2.toFixed(2)})</span>
              </p>
              <p>
                Harina = <span className="font-mono text-amber-800">{harinaModel.b.toFixed(1)} + {harinaModel.m.toFixed(2)} × °C</span>{" "}
                <span className="text-xs text-slate-400">(R²={harinaModel.r2.toFixed(2)})</span>
              </p>
              <p className="text-xs text-slate-400 pt-1">
                Por cada grado extra: +{moteModel.m.toFixed(2)} kg mote, +{harinaModel.m.toFixed(2)} kg harina
              </p>
              <p className="text-xs text-slate-400">
                Calculado con {dataWithTemp.length} registros con temperatura confirmada
                {pendingCount > 0 && ` (${pendingCount} pendientes de clima)`}.
              </p>
            </div>
          ) : (
            <p className="text-sm text-slate-500">
              Agrega al menos 2 pedidos con fecha para que la app busque la temperatura histórica y calcule la fórmula.
            </p>
          )}
          {histError && <p className="text-xs text-red-500 mt-1">{histError}</p>}
        </div>

        <div className="bg-white rounded-2xl p-4 shadow-sm border border-amber-100">
          <div className="flex items-center gap-2 mb-3">
            <Package size={18} className="text-amber-700" />
            <h2 className="font-semibold text-amber-900">Próximos 7 días</h2>
          </div>
          {loadingWeather && <p className="text-sm text-slate-500">Cargando clima...</p>}
          {weatherError && <p className="text-sm text-red-500">{weatherError}</p>}
          <div className="space-y-2">
            {forecast.map((d) => (
              <div key={d.fecha} className="flex items-center justify-between text-sm border-b border-amber-50 last:border-0 pb-2 last:pb-0">
                <div className="flex items-center gap-2 w-24">
                  <WeatherIcon code={d.code} />
                  <span className="text-slate-600">
                    {new Date(d.fecha + "T00:00:00").toLocaleDateString("es-CL", { weekday: "short", day: "numeric" })}
                  </span>
                </div>
                <span className="font-semibold text-slate-700 w-12">{d.temp}°C</span>
                <span className="text-amber-700 text-right">{predict(moteModel, d.temp)} kg mote</span>
                <span className="text-orange-600 text-right">{predict(harinaModel, d.temp)} kg harina</span>
              </div>
            ))}
          </div>
        </div>

        {dataWithTemp.length >= 2 && (
          <div className="bg-white rounded-2xl p-4 shadow-sm border border-amber-100">
            <h2 className="font-semibold text-amber-900 mb-3">Temperatura vs. Ventas</h2>
            <ResponsiveContainer width="100%" height={220}>
              <LineChart data={chartData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#fde6c4" />
                <XAxis dataKey="temp" label={{ value: "°C", position: "insideBottom", offset: -2 }} fontSize={12} />
                <YAxis fontSize={12} label={{ value: "kg", angle: -90, position: "insideLeft" }} />
                <Tooltip />
                <Legend />
                <Line type="monotone" dataKey="mote" stroke="#b45309" strokeWidth={2} name="Mote (kg)" />
                <Line type="monotone" dataKey="harina" stroke="#ea580c" strokeWidth={2} name="Harina (kg)" />
              </LineChart>
            </ResponsiveContainer>
          </div>
        )}

        <div className="bg-white rounded-2xl p-4 shadow-sm border border-amber-100">
          <h2 className="font-semibold text-amber-900 mb-1">Ingresar temperaturas manualmente</h2>
          <p className="text-xs text-slate-400 mb-2">
            Si el servicio de clima no conecta, busca tú el histórico (ej. en la Dirección Meteorológica de Chile)
            y pégalo aquí. Formato por línea: fecha, temperatura máxima
          </p>
          <textarea
            value={pasteTempText}
            onChange={(e) => setPasteTempText(e.target.value)}
            placeholder={"01-03-2026, 24\n08-03-2026, 27"}
            className="w-full border border-amber-200 rounded-lg px-2 py-1.5 text-sm font-mono h-20"
          />
          <button
            onClick={importTemps}
            className="w-full mt-2 bg-amber-700 hover:bg-amber-800 text-white rounded-lg py-2 text-sm font-medium transition-colors"
          >
            Cargar temperaturas
          </button>
          {tempImportInfo && <p className="text-xs text-green-600 mt-2">{tempImportInfo}</p>}
        </div>

        <div className="bg-white rounded-2xl p-4 shadow-sm border border-amber-100">
          <h2 className="font-semibold text-amber-900 mb-1">Subir Excel de pedidos</h2>
          <p className="text-xs text-slate-400 mb-2">Columnas en orden: Fecha, Mote (kg), Harina (kg). La primera fila puede ser encabezado.</p>
          <label className="w-full flex items-center justify-center gap-2 bg-amber-700 hover:bg-amber-800 text-white rounded-lg py-2.5 text-sm font-medium cursor-pointer transition-colors">
            <Upload size={16} /> Subir archivo .xlsx
            <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} className="hidden" />
          </label>
          {uploadInfo && <p className="text-xs text-green-600 mt-2">{uploadInfo}</p>}
          {uploadError && <p className="text-xs text-red-500 mt-2">{uploadError}</p>}
        </div>

        <div className="bg-white rounded-2xl p-4 shadow-sm border border-amber-100">
          <h2 className="font-semibold text-amber-900 mb-1">Pegar varios pedidos a la vez</h2>
          <p className="text-xs text-slate-400 mb-2">Formato por línea: AAAA-MM-DD, kilos mote, kilos harina</p>
          <textarea
            value={pasteText}
            onChange={(e) => setPasteText(e.target.value)}
            placeholder={"2026-03-01, 12, 5\n2026-03-08, 18, 7.5"}
            className="w-full border border-amber-200 rounded-lg px-2 py-1.5 text-sm font-mono h-20"
          />
          <button
            onClick={importPaste}
            className="w-full mt-2 bg-orange-600 hover:bg-orange-700 text-white rounded-lg py-2 text-sm font-medium transition-colors"
          >
            Importar pedidos pegados
          </button>
        </div>

        <div className="bg-white rounded-2xl p-4 shadow-sm border border-amber-100">
          <h2 className="font-semibold text-amber-900 mb-3">Registro de pedidos (uno por uno)</h2>
          <div className="grid grid-cols-3 gap-2 mb-3">
            <input
              type="date"
              value={form.fecha}
              onChange={(e) => setForm((f) => ({ ...f, fecha: e.target.value }))}
              className="border border-amber-200 rounded-lg px-2 py-1.5 text-sm"
            />
            <input
              type="number"
              placeholder="Mote kg"
              value={form.mote}
              onChange={(e) => setForm((f) => ({ ...f, mote: e.target.value }))}
              className="border border-amber-200 rounded-lg px-2 py-1.5 text-sm"
            />
            <input
              type="number"
              placeholder="Harina kg"
              value={form.harina}
              onChange={(e) => setForm((f) => ({ ...f, harina: e.target.value }))}
              className="border border-amber-200 rounded-lg px-2 py-1.5 text-sm"
            />
          </div>
          <button
            onClick={addEntry}
            className="w-full bg-amber-700 hover:bg-amber-800 text-white rounded-lg py-2 text-sm font-medium flex items-center justify-center gap-1 mb-3 transition-colors"
          >
            <Plus size={16} /> Agregar pedido
          </button>

          <div className="space-y-1 max-h-64 overflow-y-auto">
            {orders
              .slice()
              .sort((a, b) => new Date(b.fecha) - new Date(a.fecha))
              .map((h) => {
                const t = tempByDate[h.fecha];
                return (
                  <div key={h.id} className="flex items-center justify-between text-sm bg-amber-50 rounded-lg px-3 py-2">
                    <span className="text-slate-600 w-20">{h.fecha}</span>
                    <span className="font-medium text-slate-700 w-12">
                      {t !== undefined ? `${t}°C` : <RefreshCw size={12} className="animate-spin inline text-amber-400" />}
                    </span>
                    <span className="text-amber-700">{h.mote}kg mote</span>
                    <span className="text-orange-600">{h.harina}kg harina</span>
                    <button onClick={() => removeEntry(h.id)} className="text-slate-400 hover:text-red-500">
                      <Trash2 size={14} />
                    </button>
                  </div>
                );
              })}
          </div>
        </div>

        <p className="text-center text-xs text-slate-400 pt-2">
          Clima actual y pronóstico: Open-Meteo · Historial climático: Open-Meteo Archive · Ubicación: Santiago
        </p>
      </div>
    </div>
  );
}
