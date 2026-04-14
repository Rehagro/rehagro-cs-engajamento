import type { ResultadoDieta } from '../types';
import { REFERENCIAS_LACTACAO, getStatus, statusColor, statusDot } from '../utils/referencias';

interface Props {
  resultado: ResultadoDieta;
  precoLeite: number;
}

function IndicCard({ label, valor, chave, unidade }: { label: string; valor: string; chave?: string; unidade?: string }) {
  let colorClass = 'bg-gray-50 border-gray-200';
  if (chave) {
    const ref = REFERENCIAS_LACTACAO[chave];
    if (ref) {
      // Not ideal but fine for simple cases
    }
  }
  return (
    <div className={`border rounded-lg p-3 ${colorClass}`}>
      <div className="text-xs text-gray-500 mb-0.5">{label}</div>
      <div className="font-mono font-bold text-gray-800">{valor}</div>
      {unidade && <div className="text-xs text-gray-400">{unidade}</div>}
    </div>
  );
}

function IndicStatus({ chave, resultado }: { chave: string; resultado: ResultadoDieta }) {
  const ref = REFERENCIAS_LACTACAO[chave];
  if (!ref) return null;
  const valor = resultado[chave as keyof ResultadoDieta] as number;
  const status = getStatus(valor, ref);
  const color = statusColor(status);
  const dot = statusDot(status);

  let valorStr = '';
  if (chave === 'fdnf_kg_pv') valorStr = (valor * 100).toFixed(3) + '%';
  else if (chave === 'pct_forragem_ms') valorStr = (valor * 100).toFixed(1) + '%';
  else if (chave === 'dcad') valorStr = valor.toFixed(0);
  else valorStr = valor.toFixed(2);

  const refStr = ref.min !== undefined && ref.max !== undefined
    ? `${ref.min} – ${ref.max}`
    : ref.min !== undefined ? `≥ ${ref.min}`
    : ref.max !== undefined ? `≤ ${ref.max}`
    : ref.ref ?? '—';

  return (
    <div className={`border rounded-lg p-3 ${color}`}>
      <div className="text-xs font-medium mb-0.5">{dot} {ref.label}</div>
      <div className="font-mono font-bold text-lg">{valorStr}</div>
      <div className="text-xs opacity-70">{ref.unidade} · ref: {refStr}</div>
    </div>
  );
}

export default function Indicadores({ resultado, precoLeite }: Props) {
  const { custoTotal, custoKgMS, custoLitro, leite_potencial_nel, leite_potencial_prot, kPf, kPc, kPl } = resultado;
  const lucro = precoLeite > 0
    ? (precoLeite * resultado.totalKgMS /* approximation using CMS as proxy */) - custoTotal
    : null;

  return (
    <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
      <h2 className="text-sm font-bold text-gray-700 mb-3">📈 Indicadores & Custos</h2>

      <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-2 mb-3">
        <IndicStatus chave="fdnf_kg_pv" resultado={resultado} />
        <IndicStatus chave="pct_forragem_ms" resultado={resultado} />
        <IndicStatus chave="fdn8_amido_deg" resultado={resultado} />
        <IndicStatus chave="lis_met" resultado={resultado} />
        <IndicStatus chave="ca_p" resultado={resultado} />
        <IndicStatus chave="dcad" resultado={resultado} />
      </div>

      <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-7 gap-2">
        <IndicCard label="Custo R$/d" valor={`R$ ${custoTotal.toFixed(2)}`} />
        <IndicCard label="Custo R$/kg MS" valor={`R$ ${custoKgMS.toFixed(3)}`} />
        <IndicCard label="Custo R$/litro" valor={`R$ ${custoLitro.toFixed(3)}`} />
        {lucro !== null && (
          <IndicCard label="Receita leite R$/d" valor={`R$ ${(precoLeite * resultado.totalKgMS).toFixed(2)}`} />
        )}
        <IndicCard label="Leite pot. NEl" valor={`${leite_potencial_nel.toFixed(1)} kg/d`} />
        <IndicCard label="Leite pot. Prot." valor={`${leite_potencial_prot.toFixed(1)} kg/d`} />
        <div className="border border-gray-200 rounded-lg p-3 bg-gray-50 col-span-1">
          <div className="text-xs text-gray-500 mb-1">Taxas de Passagem</div>
          <div className="text-xs font-mono space-y-0.5">
            <div>kPf: <span className="font-bold">{(kPf * 100).toFixed(2)}%/h</span></div>
            <div>kPc: <span className="font-bold">{(kPc * 100).toFixed(2)}%/h</span></div>
            <div>kPl: <span className="font-bold">{(kPl * 100).toFixed(2)}%/h</span></div>
          </div>
        </div>
      </div>
    </div>
  );
}
