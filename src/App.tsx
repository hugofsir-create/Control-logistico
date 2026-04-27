import React, { useState, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { format, differenceInDays, isPast, isToday, parse, isValid, startOfDay } from 'date-fns';
import { es } from 'date-fns/locale';
import { 
  Package, 
  AlertCircle, 
  FileUp, 
  TrendingUp, 
  Search,
  BarChart3,
  Calendar,
  Clock,
  MessageSquare,
  Send,
  Loader2,
  X
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { cn } from './lib/utils';
import { Delivery, DeliveryStats } from './types';
import { askAssistant } from './services/aiService';

export default function App() {
  const [deliveries, setDeliveries] = useState<Delivery[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [isAiOpen, setIsAiOpen] = useState(false);
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);
  const [aiInput, setAiInput] = useState('');
  const [aiMessages, setAiMessages] = useState<{ role: 'user' | 'assistant', content: string }[]>([]);
  const [isAiLoading, setIsAiLoading] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Sistema mejorado de parseo de fechas de Excel
  const parseExcelDate = (val: any) => {
    if (val === undefined || val === null || val === '') return new Date();
    
    let resolvedDate: Date | null = null;

    // 1. Si ya es un objeto Date (por cellDates: true)
    if (val instanceof Date) {
      resolvedDate = val;
    }
    
    // 2. Si es un número (formato serial de Excel como 45291)
    else if (typeof val === 'number') {
      try {
        const dateObj = XLSX.SSF.parse_date_code(val);
        resolvedDate = new Date(dateObj.y, dateObj.m - 1, dateObj.d);
      } catch (e) {}
    }
    
    // 3. Si es un string (e.g. "27/12/26")
    else if (typeof val === 'string') {
      const cleaned = val.trim();
      const formats = ['dd/MM/yyyy', 'd/M/yyyy', 'dd/MM/yy', 'd/M/yy', 'yyyy-MM-dd', 'MM/dd/yyyy'];
      
      for (const f of formats) {
        const parsed = parse(cleaned, f, new Date());
        if (isValid(parsed) && !isNaN(parsed.getTime())) {
          resolvedDate = parsed;
          break;
        }
      }

      if (!resolvedDate) {
        const d = new Date(cleaned);
        if (isValid(d)) resolvedDate = d;
      }
    }

    if (resolvedDate && isValid(resolvedDate)) {
      // FIX: Forzar años de 2 dígitos a 2000+ si están en el rango lógico del negocio
      let year = resolvedDate.getFullYear();
      if (year < 100) {
        // Si el año es < 50 asumimos 2000, si es > 50 asumimos 1900
        year += (year < 50 ? 2000 : 1900);
        resolvedDate.setFullYear(year);
      }
      return startOfDay(resolvedDate);
    }
    
    return startOfDay(new Date());
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement> | React.DragEvent) => {
    let file: File | null = null;
    if ('files' in e.target && e.target.files) {
      file = e.target.files[0];
    } else if ('dataTransfer' in e) {
      e.preventDefault();
      file = e.dataTransfer.files[0];
    }

    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        // raw: false permite que XLSX-JS nos devuelva los valores formateados según el Excel si es posible
        const wb = XLSX.read(bstr, { type: 'binary', cellDates: true, cellNF: true, cellText: true });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        
        // Obtenemos los datos con header: "A" para mapear por letras fijas.
        const data = XLSX.utils.sheet_to_json(ws, { header: "A", raw: true }) as any[];

        const today = startOfDay(new Date());

        // Saltamos la fila de encabezados
        const mappedData: Delivery[] = data.slice(1).map((row, index) => {
          // Columna E: Creación, F: Empresa, G: Comprobante, H: Cliente, J: Localidad, Y: Vencimiento
          const creationDate = parseExcelDate(row['E']);
          const promisedDate = parseExcelDate(row['Y']);
          
          const company = String(row['F'] || '').trim();
          const receiptNumber = String(row['G'] || '').trim();
          const customerName = String(row['H'] || 'Cliente Genérico').trim();
          const locality = String(row['J'] || '').trim();
          
          const daysDiff = differenceInDays(promisedDate, today);
          let priority: Delivery['priority'] = 'Media';
          if (daysDiff < 0) priority = 'Crítica';
          else if (daysDiff <= 2) priority = 'Alta';
          else if (daysDiff > 7) priority = 'Baja';

          return {
            id: receiptNumber || `TEMP-${index}`,
            receiptNumber: receiptNumber,
            company,
            customer: customerName,
            destination: row['I'] || 'Sin dirección', 
            locality,
            promisedDate,
            creationDate,
            status: (daysDiff < 0 ? 'Demorado' : 'Pendiente') as Delivery['status'],
            priority
          };
        }).filter(d => d.receiptNumber && d.receiptNumber !== "undefined" && d.receiptNumber !== "null");

        setDeliveries(mappedData);
      } catch (err) {
        console.error("Error al procesar archivo:", err);
      }
    };
    reader.readAsBinaryString(file);
  };

  const sortedDeliveries = useMemo(() => {
    return [...deliveries]
      .filter(d => 
        d.customer.toLowerCase().includes(searchTerm.toLowerCase()) ||
        d.receiptNumber.toLowerCase().includes(searchTerm.toLowerCase())
      )
      .sort((a, b) => a.promisedDate.getTime() - b.promisedDate.getTime());
  }, [deliveries, searchTerm]);

  const stats: DeliveryStats = useMemo(() => {
    const today = startOfDay(new Date());
    const delayed = deliveries.filter(d => isPast(d.promisedDate) && !isToday(d.promisedDate));
    const critical = deliveries.filter(d => {
      const diff = differenceInDays(d.promisedDate, today);
      return diff >= 0 && diff <= 1;
    });
    
    const totalDelayDays = delayed.reduce((acc, d) => acc + Math.abs(differenceInDays(today, d.promisedDate)), 0);

    return {
      totalPending: deliveries.length,
      totalDelayed: delayed.length,
      averageDelayDays: delayed.length > 0 ? Math.round(totalDelayDays / delayed.length) : 0,
      criticalDeliveries: critical.length
    };
  }, [deliveries]);

  const delayBreakdown = useMemo(() => {
    if (deliveries.length === 0) return [];
    const counts = { low: 0, medium: 0, high: 0 };
    deliveries.forEach(d => {
      const diff = differenceInDays(new Date(), d.promisedDate);
      if (diff > 5) counts.high++;
      else if (diff > 0) counts.medium++;
      else counts.low++;
    });
    const total = deliveries.length;
    return [
      { name: 'Menos de 24h', value: Math.round((counts.low/total)*100) || 0, color: '#10B981' },
      { name: '1-3 Días', value: Math.round((counts.medium/total)*100) || 0, color: '#F59E0B' },
      { name: 'Crítico (>5 Días)', value: Math.round((counts.high/total)*100) || 0, color: '#EF4444' }
    ];
  }, [deliveries]);

  const handleAiSend = async () => {
    if (!aiInput.trim()) return;
    const userMsg = aiInput;
    setAiInput('');
    setAiMessages(prev => [...prev, { role: 'user', content: userMsg }]);
    setIsAiLoading(true);

    try {
      const response = await askAssistant(userMsg, deliveries);
      setAiMessages(prev => [...prev, { role: 'assistant', content: response }]);
    } catch (err) {
      setAiMessages(prev => [...prev, { role: 'assistant', content: "Lo siento, tuve un problema analizando los datos." }]);
    } finally {
      setIsAiLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-[#0A0A0B] text-slate-300 font-sans selection:bg-amber-500/30 flex flex-col overflow-hidden h-screen">
      <header className="h-16 shrink-0 border-b border-slate-800 flex items-center justify-between px-8 bg-[#0F1115] z-50">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 bg-amber-500 rounded flex items-center justify-center text-black font-bold">L</div>
          <h1 className="text-xl font-semibold tracking-tight text-white">
            Control<span className="text-amber-500">Logístico</span>
          </h1>
          <button 
            onClick={() => setIsSidebarOpen(!isSidebarOpen)}
            className="ml-4 p-2 bg-[#1A1D23] rounded-md border border-slate-700 text-slate-400 hover:text-white transition-all active:scale-95 group"
            title={isSidebarOpen ? "Ocultar panel de métricas" : "Mostrar panel de métricas"}
          >
            {isSidebarOpen ? <X className="w-4 h-4 group-hover:rotate-90 transition-transform" /> : <BarChart3 className="w-4 h-4" />}
          </button>
          <div className="ml-6 px-3 py-1 bg-slate-800/50 rounded-full border border-slate-700 flex items-center gap-2">
            <span className="text-[10px] text-slate-500 font-bold uppercase tracking-widest">Hoy:</span>
            <span className="text-[10px] text-amber-500 font-bold">{format(new Date(), 'dd/MM/yyyy')}</span>
          </div>
        </div>
        
        <div className="flex items-center gap-6">
          <div className="hidden md:flex items-center gap-2 bg-[#1A1D23] px-3 py-1.5 rounded-md border border-slate-700">
            <span className="text-[10px] text-slate-500 uppercase font-bold tracking-wider">ARCHIVO ACTIVO:</span>
            <span className="text-[10px] text-slate-300 italic truncate max-w-[150px]">
              {deliveries.length > 0 ? "Logistica_Activa.xlsx" : "Sin archivo"}
            </span>
          </div>
          
          <button 
            onClick={() => fileInputRef.current?.click()}
            className="px-4 py-1.5 bg-amber-500 text-black rounded-md font-bold text-xs uppercase tracking-widest hover:bg-amber-400 transition-all flex items-center gap-2 active:scale-95 shadow-lg shadow-amber-500/10"
          >
            <FileUp className="w-3.5 h-3.5" />
            <span>IMPORTAR</span>
          </button>
          <input 
            type="file" 
            ref={fileInputRef} 
            onChange={handleFileUpload} 
            accept=".xlsx, .xls, .csv" 
            className="hidden" 
          />
        </div>
      </header>

      <div className="flex-1 flex overflow-hidden">
        {deliveries.length === 0 ? (
          <div className="flex-1 flex flex-col items-center justify-center p-8">
            <div 
              onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
              onDragLeave={() => setIsDragging(false)}
              onDrop={(e) => { e.preventDefault(); setIsDragging(false); handleFileUpload(e); }}
              className={cn(
                "w-full max-w-2xl h-[400px] border-2 border-dashed rounded-2xl flex flex-col items-center justify-center transition-all",
                isDragging ? "border-amber-500 bg-amber-500/5 scale-[1.01]" : "border-slate-800 bg-[#0F1115]"
              )}
            >
              <div className="w-16 h-16 bg-[#1A1D23] rounded-xl flex items-center justify-center mb-6 border border-slate-700">
                <FileUp className="w-8 h-8 text-slate-500" />
              </div>
              <h2 className="text-xl font-semibold mb-2 text-white">LISTO PARA INTEGRACIÓN</h2>
              <p className="text-slate-500 text-center max-w-sm px-6">
                Arrastra tu archivo Excel aquí. Se procesarán: E (Creación), F (Empresa), G (Comprobante), H (Cliente), J (Localidad) e Y (Vencimiento).
              </p>
              <button 
                onClick={() => fileInputRef.current?.click()}
                className="mt-8 text-amber-500 font-bold text-xs uppercase tracking-widest hover:text-amber-400 transition-all border border-amber-500/20 px-6 py-3 rounded-md bg-amber-500/5 hover:bg-amber-500/10"
              >
                BUSCAR ARCHIVO
              </button>
            </div>
          </div>
        ) : (
          <main className="flex-1 grid grid-cols-12 gap-0 overflow-hidden relative">
            {/* Sidebar Metrics */}
            <AnimatePresence mode="wait">
              {isSidebarOpen && (
                <motion.aside 
                  initial={{ x: -20, opacity: 0 }}
                  animate={{ x: 0, opacity: 1 }}
                  exit={{ x: -20, opacity: 0 }}
                  className="col-span-4 border-r border-slate-800 bg-[#0F1115] p-6 flex flex-col gap-6 overflow-y-auto shrink-0 h-full"
                >
                  <h2 className="text-xs font-bold uppercase tracking-widest text-slate-500 mb-2">Métricas Operativas</h2>
                  
                  <div className="grid grid-cols-2 gap-4">
                <div className="bg-[#1A1D23] p-4 rounded-lg border border-slate-800 transition-colors hover:border-slate-700 group">
                  <div className="text-[10px] text-slate-500 mb-1 uppercase tracking-wider font-bold">Total Pendientes</div>
                  <div className="text-2xl font-bold text-white tabular-nums">{stats.totalPending}</div>
                  <div className="text-[10px] text-amber-500 mt-1 flex items-center gap-1">
                    <TrendingUp className="w-3 h-3" />
                    <span>Activas ahora</span>
                  </div>
                </div>
                <div className="bg-[#1A1D23] p-4 rounded-lg border border-slate-800 transition-colors hover:border-slate-700">
                  <div className="text-[10px] text-slate-500 mb-1 uppercase tracking-wider font-bold">Demora Promedio</div>
                  <div className="text-2xl font-bold text-white tabular-nums">{stats.averageDelayDays}d</div>
                  <div className={cn("text-[10px] mt-1", stats.averageDelayDays > 0 ? "text-red-500" : "text-emerald-500")}>
                    {stats.averageDelayDays > 0 ? "Acción Requerida" : "Sistema Saludable"}
                  </div>
                </div>
                <div className="bg-[#1A1D23] p-4 rounded-lg border border-slate-800 transition-colors hover:border-slate-700">
                  <div className="text-[10px] text-slate-500 mb-1 uppercase tracking-wider font-bold">Eficiencia</div>
                  <div className="text-2xl font-bold text-white">94.2%</div>
                  <div className="text-[10px] text-emerald-500 mt-1">Zona Pro</div>
                </div>
                <div className="bg-[#1A1D23] p-4 rounded-lg border border-slate-800 transition-colors hover:border-slate-700">
                  <div className="text-[10px] text-slate-500 mb-1 uppercase tracking-wider font-bold">Críticas</div>
                  <div className="text-2xl font-bold text-red-500 tabular-nums">{stats.criticalDeliveries}</div>
                  <div className="text-[10px] text-slate-500 mt-1">Próximas 24h</div>
                </div>
              </div>

              <div className="mt-4 flex-1">
                <h3 className="text-xs font-bold uppercase tracking-widest text-slate-500 mb-4">Desglose de Demoras</h3>
                <div className="space-y-4">
                  {delayBreakdown.map((item) => (
                    <div key={item.name}>
                      <div className="flex justify-between text-[11px] mb-1 font-medium">
                        <span className="text-slate-400">{item.name}</span>
                        <span className="text-white">{item.value}%</span>
                      </div>
                      <div className="w-full bg-slate-800 h-1.5 rounded-full overflow-hidden">
                        <motion.div 
                          initial={{ width: 0 }}
                          animate={{ width: `${item.value}%` }}
                          transition={{ duration: 1, ease: "easeOut" }}
                          style={{ backgroundColor: item.color }}
                          className="h-full rounded-full shadow-[0_0_8px_rgba(0,0,0,0.5)]"
                        />
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              <div className="mt-8 space-y-4 shrink-0">
                <div className="p-4 bg-amber-500/5 border border-amber-500/10 rounded-lg">
                  <p className="text-[11px] text-amber-500 font-bold uppercase tracking-widest mb-1 flex items-center gap-2">
                    <AlertCircle className="w-3 h-3" />
                    Insight IA
                  </p>
                  <p className="text-xs text-slate-400 leading-relaxed">
                    Prioriza las {stats.criticalDeliveries} entregas inminentes para optimizar la ruta logística hoy.
                  </p>
                </div>
                <button 
                  onClick={() => setIsAiOpen(true)}
                  className="w-full py-3 bg-amber-500 text-black font-bold rounded-md hover:bg-amber-400 transition-all uppercase text-[10px] tracking-widest shadow-lg shadow-amber-500/10 active:scale-[0.98] flex items-center justify-center gap-2"
                >
                  <MessageSquare className="w-3.5 h-3.5" />
                  Consultar con Asistente IA
                </button>
              </div>
            </motion.aside>
          )}
        </AnimatePresence>

            {/* Delivery Table Area */}
            <section className={cn(
              "flex flex-col bg-[#0A0A0B] overflow-hidden transition-all duration-300",
              isSidebarOpen ? "col-span-8" : "col-span-12"
            )}>
              <div className="px-6 py-4 flex justify-between items-center border-b border-slate-800 bg-[#0F1115]">
                <h2 className="text-xs font-bold uppercase tracking-widest text-slate-500">Prioridad</h2>
                <div className="flex items-center gap-4">
                  <div className="relative">
                    <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-slate-500" />
                    <input 
                      type="text" 
                      placeholder="Buscar comprobante..."
                      className="bg-[#1A1D23] border border-slate-700 text-xs text-slate-300 pl-9 pr-4 py-1.5 rounded-md focus:outline-none focus:border-amber-500/50 transition-colors w-48"
                      value={searchTerm}
                      onChange={(e) => setSearchTerm(e.target.value)}
                    />
                  </div>
                  <span className="px-2 py-1 rounded text-[10px] bg-emerald-500/10 text-emerald-500 border border-emerald-500/20 font-bold uppercase tracking-wider">Antigüedad</span>
                </div>
              </div>
              
              <div className="flex-1 overflow-y-auto">
                <table className="w-full text-left border-collapse">
                  <thead className="sticky top-0 bg-[#0F1115] text-[10px] uppercase tracking-wider text-slate-500 z-20">
                    <tr className="border-b border-slate-800">
                      <th className="px-6 py-4 font-bold uppercase">Comprobante</th>
                      <th className="px-6 py-4 font-bold uppercase">Empresa</th>
                      <th className="px-6 py-4 font-bold uppercase">Cliente / Ciudad</th>
                      <th className="px-6 py-4 font-bold uppercase text-center">Tiempos</th>
                      <th className="px-6 py-4 font-bold text-right uppercase">Estado Operativo</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/30">
                    {sortedDeliveries.map((delivery) => {
                      const today = startOfDay(new Date());
                      const isDelayed = isPast(delivery.promisedDate) && !isToday(delivery.promisedDate);
                      const isUrg = isToday(delivery.promisedDate);
                      const diffInDaysTotal = differenceInDays(today, delivery.promisedDate);

                      return (
                        <tr 
                          key={delivery.id} 
                          className={cn(
                            "group transition-colors",
                            isDelayed ? "bg-red-500/[0.03] hover:bg-red-500/[0.06]" : 
                            isUrg ? "bg-amber-500/[0.03] hover:bg-amber-500/[0.06]" : 
                            "hover:bg-slate-800/30"
                          )}
                        >
                          <td className="px-6 py-4 font-mono text-sm text-white font-medium">{delivery.receiptNumber}</td>
                          <td className="px-6 py-4 text-xs font-semibold text-slate-400">
                            {delivery.company}
                          </td>
                          <td className="px-6 py-4">
                            <div className="text-sm text-slate-200 font-medium">{delivery.customer}</div>
                            <div className="text-[10px] text-slate-500 uppercase tracking-tighter">{delivery.locality}</div>
                          </td>
                          <td className="px-6 py-4">
                            <div className="flex flex-col gap-1 items-center">
                              <div className="flex flex-col items-center p-1.5 bg-slate-800/30 rounded border border-slate-700/50 w-24">
                                <span className="text-[9px] text-slate-500 font-bold">EMISIÓN</span>
                                <span className="text-[11px] text-slate-300">{format(delivery.creationDate, 'dd/MM/yyyy')}</span>
                              </div>
                              <div className={cn(
                                "flex flex-col items-center p-1.5 rounded border w-24 mt-1",
                                isDelayed ? "bg-red-500/10 border-red-500/20" : isUrg ? "bg-amber-500/10 border-amber-500/20" : "bg-emerald-500/10 border-emerald-500/20"
                              )}>
                                <span className="text-[9px] text-slate-500 font-bold uppercase">Vencimiento</span>
                                <span className={cn(
                                  "text-[11px] font-bold",
                                  isDelayed ? "text-red-400" : isUrg ? "text-amber-400" : "text-emerald-400"
                                )}>{format(delivery.promisedDate, 'dd/MM/yyyy')}</span>
                              </div>
                            </div>
                          </td>
                          <td className="px-6 py-4 text-right">
                            {isDelayed ? (
                              <div className="flex flex-col items-end">
                                <span className="text-sm font-bold text-red-500 uppercase tracking-wider">{Math.abs(diffInDaysTotal)} DÍAS ATRASO</span>
                                <span className="text-[10px] text-slate-500 mt-1">Plazo de entrega: {differenceInDays(delivery.promisedDate, delivery.creationDate)} días</span>
                                <span className="text-[10px] text-slate-500 uppercase font-bold mt-1">Prioridad: {delivery.priority}</span>
                              </div>
                            ) : (
                              <div className="flex flex-col items-end">
                                <span className={cn("text-sm font-bold uppercase tracking-wider", isUrg ? "text-amber-500" : "text-emerald-500")}>
                                  {isUrg ? "VENCE HOY" : `FALTAN ${Math.abs(diffInDaysTotal)} DÍAS`}
                                </span>
                                <span className="text-[10px] text-slate-500 mt-1">Plazo total: {differenceInDays(delivery.promisedDate, delivery.creationDate)} días</span>
                                <span className="text-[10px] text-slate-300 uppercase mt-1">Estado: {delivery.status}</span>
                              </div>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              <footer className="shrink-0 border-t border-slate-800 bg-[#0F1115] px-6 py-4 flex items-center justify-between">
                <span className="text-[10px] text-slate-500 uppercase tracking-widest font-bold">
                  PEDIDOS: {deliveries.length}
                </span>
                <div className="flex gap-4">
                  <button className="text-[10px] uppercase font-bold tracking-widest text-slate-500 hover:text-white transition-colors">Anterior</button>
                  <div className="flex gap-2 items-center">
                    <span className="w-6 h-6 rounded bg-slate-800 border border-slate-700 flex items-center justify-center text-[10px] text-white font-bold">1</span>
                  </div>
                  <button className="text-[10px] uppercase font-bold tracking-widest text-slate-500 hover:text-white transition-colors">Siguiente</button>
                </div>
              </footer>
            </section>

            {/* IA Assistant Panel */}
            <AnimatePresence>
              {isAiOpen && (
                <motion.div 
                  initial={{ x: '100%' }}
                  animate={{ x: 0 }}
                  exit={{ x: '100%' }}
                  transition={{ type: 'spring', damping: 25, stiffness: 200 }}
                  className="absolute right-0 top-0 h-full w-[400px] bg-[#0F1115] border-l border-slate-800 shadow-2xl z-[60] flex flex-col"
                >
                  <div className="p-4 border-b border-slate-800 flex items-center justify-between bg-[#1A1D23]">
                    <div className="flex items-center gap-2">
                      <div className="w-2 h-2 bg-amber-500 rounded-full animate-pulse" />
                      <h3 className="font-bold text-xs uppercase tracking-widest text-white">LogiAssistant IA</h3>
                    </div>
                    <button onClick={() => setIsAiOpen(false)} className="p-1 hover:text-white transition-colors">
                      <X className="w-4 h-4" />
                    </button>
                  </div>
                  
                  <div className="flex-1 overflow-y-auto p-6 space-y-6">
                    {aiMessages.length === 0 && (
                      <div className="h-full flex flex-col items-center justify-center text-center opacity-50 space-y-4">
                        <MessageSquare className="w-12 h-12 text-slate-700" />
                        <p className="text-xs max-w-[200px]">Pregúntame sobre el estado de las entregas, retrasos o clientes prioritarios.</p>
                      </div>
                    )}
                    {aiMessages.map((msg, i) => (
                      <div key={i} className={cn(
                        "flex flex-col gap-2",
                        msg.role === 'user' ? "items-end" : "items-start"
                      )}>
                        <div className={cn(
                          "max-w-[85%] p-3 rounded-2xl text-xs leading-relaxed",
                          msg.role === 'user' 
                            ? "bg-amber-500 text-black font-medium" 
                            : "bg-[#1A1D23] text-slate-300 border border-slate-700"
                        )}>
                          {msg.content}
                        </div>
                      </div>
                    ))}
                    {isAiLoading && (
                      <div className="flex items-center gap-2 text-slate-500 italic text-[10px]">
                        <Loader2 className="w-3 h-3 animate-spin" />
                        Analizando base de datos logística...
                      </div>
                    )}
                  </div>

                  <div className="p-4 border-t border-slate-800 bg-[#0F1115]">
                    <div className="relative">
                      <input 
                        type="text" 
                        value={aiInput}
                        onChange={(e) => setAiInput(e.target.value)}
                        onKeyDown={(e) => e.key === 'Enter' && handleAiSend()}
                        placeholder="Pregunta algo..."
                        className="w-full bg-[#1A1D23] border border-slate-700 rounded-xl py-3 pl-4 pr-12 text-xs focus:outline-none focus:border-amber-500 transition-colors"
                      />
                      <button 
                        onClick={handleAiSend}
                        disabled={isAiLoading}
                        className="absolute right-2 top-1/2 -translate-y-1/2 p-2 bg-amber-500 text-black rounded-lg hover:bg-amber-400 disabled:opacity-50 transition-all"
                      >
                        <Send className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>
          </main>
        )}
      </div>
    </div>
  );
}
