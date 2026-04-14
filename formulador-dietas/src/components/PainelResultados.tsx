import { useState } from 'react';
import { ChevronDown, ChevronRight } from 'lucide-react';
import type { ResultadoDieta } from '../types';
import { REFERENCIAS_LACTACAO, getStatus, statusColor, statusDot } from '../utils/referencias';
import { formatarValor } from '../utils/calculos';

interface Props {
  resultado: ResultadoDieta;
}

interface NutrienteRowProps {
  chave: string;
  valor: number;
}

function NutrienteRow({ chave, valor }: NutrienteRowProps) {
  const ref = REFERENCIAS_LACTACAO[chave];
  if (!ref) return null;
  const status = getStatus(valor, ref);
  const color = statusColor(status);
  const dot = statusDot(status);
  const valorFormatado = formatarValor(valor, ref.unidade);
  const refStr = ref.min !== undefined && ref.max !== undefined
    ? `${formatarValor(ref.min, ref.unidade)} – ${formatarValor(ref.max, ref.unidade)}`
    : ref.min !== undefined
    ? `≥ ${formatarValor(ref.min, ref.unidade)}`
    : ref.max !== undefined
    ? `≤ ${formatarValor(ref.max, ref.unidade)}`
    : ref.ref ?? '—';

  return (
    <div className={`flex items-center justify-between px-3 py-1.5 rounded text-sm ${color}`}>
      <span className="font-medium">{dot} {ref.label}</span>
      <div className="text-right">
        <span className="font-mono font-bold">{valorFormatado}</span>
        <span className="text-xs ml-2 opacity-70">{refStr}</span>
      </div>
    </div>
  );
}

interface SecaoProps {
  titulo: string;
  chaves: string[];
  resultado: ResultadoDieta;
  defaultOpen?: boolean;
}

function Secao({ titulo, chaves, resultado, defaultOpen = false }: SecaoProps) {
  const [open, setOpen] = useState(defaultOpen);

  // contar alertas
  const alertas = chaves.filter(k => {
    const ref = REFERENCIAS_LACTACAO[k];
    if (!ref) return false;
    const v = resultado[k as keyof ResultadoDieta] as number;
    const s = getStatus(v, ref);
    return s === 'critico_alto' || s === 'critico_baixo';
  }).length;

  const avisos = chaves.filter(k => {
    const ref = REFERENCIAS_LACTACAO[k];
    if (!ref) return false;
    const v = resultado[k as keyof ResultadoDieta] as number;
    const s = getStatus(v, ref);
    return s === 'alto' || s === 'baixo';
  }).length;

  return (
    <div className="border border-gray-200 rounded-lg overflow-hidden">
      <button
        onClick={() => setOpen(o => !o)}
        className="w-full flex items-center justify-between px-3 py-2 bg-gray-50 hover:bg-gray-100 transition-colors text-left"
      >
        <span className="text-sm font-semibold text-gray-700">{titulo}</span>
        <div className="flex items-center gap-1.5">
          {alertas > 0 && (
            <span className="bg-red-100 text-red-700 text-xs font-bold px-1.5 py-0.5 rounded-full">
              {alertas}🔴
            </span>
          )}
          {avisos > 0 && (
            <span className="bg-yellow-100 text-yellow-700 text-xs font-bold px-1.5 py-0.5 rounded-full">
              {avisos}🟡
            </span>
          )}
          {open ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
        </div>
      </button>
      {open && (
        <div className="p-2 flex flex-col gap-1">
          {chaves.map(k => (
            <NutrienteRow key={k} chave={k} valor={resultado[k as keyof ResultadoDieta] as number} />
          ))}
        </div>
      )}
    </div>
  );
}

export default function PainelResultados({ resultado }: Props) {
  const { totalKgMS, cmsExigida, leite_potencial_nel, leite_potencial_prot } = resultado;
  const pctCMS = cmsExigida > 0 ? (totalKgMS / cmsExigida) * 100 : 0;

  const barColor = pctCMS < 85 ? 'bg-red-500' : pctCMS < 95 ? 'bg-yellow-500' : pctCMS <= 110 ? 'bg-green-500' : 'bg-orange-500';

  return (
    <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm flex flex-col gap-3">
      <h2 className="text-sm font-bold text-gray-700 flex items-center gap-2">
        📊 Resultados
      </h2>

      {/* Produção — sempre visível */}
      <div className="bg-blue-50 border border-blue-200 rounded-lg p-3">
        <div className="text-xs font-bold text-blue-700 mb-2">🥛 Produção</div>
        <div className="flex flex-col gap-1.5">
          <div>
            <div className="flex justify-between text-xs text-blue-600 mb-0.5">
              <span>CMS real vs exigida</span>
              <span className="font-mono font-bold">{totalKgMS.toFixed(1)} / {cmsExigida.toFixed(1)} kg</span>
            </div>
            <div className="h-2 bg-blue-200 rounded-full overflow-hidden">
              <div
                className={`h-full rounded-full transition-all ${barColor}`}
                style={{ width: `${Math.min(pctCMS, 120)}%` }}
              />
            </div>
            <div className="text-xs text-blue-500 text-right mt-0.5">{pctCMS.toFixed(0)}%</div>
          </div>
          <div className="flex justify-between text-xs">
            <span className="text-gray-600">Leite pot. NEl</span>
            <span className="font-mono font-bold text-green-700">{leite_potencial_nel.toFixed(1)} kg/d</span>
          </div>
          <div className="flex justify-between text-xs">
            <span className="text-gray-600">Leite pot. Proteína</span>
            <span className="font-mono font-bold text-purple-700">{leite_potencial_prot.toFixed(1)} kg/d</span>
          </div>
        </div>
      </div>

      {/* Seções expansíveis */}
      <Secao titulo="⚡ Energia & Carboidratos" chaves={['nel', 'ndt', 'cnf', 'amido', 'amido_deg']} resultado={resultado} />
      <Secao titulo="🧬 Proteína" chaves={['pb', 'pdr', 'pndr', 'met', 'lys']} resultado={resultado} defaultOpen />
      <Secao titulo="🌾 Fibra" chaves={['fdn', 'efdn', 'fdnf', 'fda']} resultado={resultado} defaultOpen />
      <Secao titulo="🔬 Gordura" chaves={['ee', 'ee_insat']} resultado={resultado} />
      <Secao titulo="🧂 Macrominerais" chaves={['ca', 'p', 'mg', 'k', 's', 'na', 'cl']} resultado={resultado} />
      <Secao titulo="💊 Microminerais" chaves={['co', 'cu', 'mn_min', 'zn', 'se', 'i', 'fe']} resultado={resultado} />
      <Secao titulo="🌟 Vitaminas & Aditivos" chaves={['vit_a', 'vit_d3', 'vit_e', 'biotina', 'monensina', 'cr', 'levedura']} resultado={resultado} />
    </div>
  );
}
