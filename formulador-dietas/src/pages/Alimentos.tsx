import { useState } from 'react';
import { Plus, Pencil, Trash2, X, Check } from 'lucide-react';
import { useDieta } from '../context/DietaContext';
import type { Alimento } from '../types';

const ALIMENTO_VAZIO: Alimento = {
  nome: '', custo: null, classificacao: 'Energético', tipo: 'C',
  ms: 0.88, pb: 0, pdr: null, pndr: null, fdn: null, efdn: null,
  mn8: null, mn19: null, fdnf: null, fda: null, nel: null, ndt: null,
  ee: null, ee_insat: null, cinza: null, cnf: null, amido: null, kd_amido: null,
  met: null, lys: null, ca: null, p: null, mg: null, k: null, s: null,
  na: null, cl: null, co: null, cu: null, mn_min: null, zn: null, se: null,
  i: null, fe: null, vit_a: null, vit_d3: null, vit_e: null,
  biotina: null, monensina: null, cr: null, levedura: null,
  prot_a: null, prot_b: null, prot_c: null, kd_prot: null,
  rup_digest: null, cp_digest: null, ndf_digest: null, fat_digest: null,
  lisina_pct: null, met_pct: null,
};

function Campo({ label, field, valor, onChange, tipo = 'number', opcoes }: {
  label: string;
  field: string;
  valor: unknown;
  onChange: (v: unknown) => void;
  tipo?: string;
  opcoes?: string[];
}) {
  if (tipo === 'select' && opcoes) {
    return (
      <div className="flex flex-col gap-0.5">
        <label className="text-xs text-gray-500">{label}</label>
        <select
          value={String(valor ?? '')}
          onChange={e => onChange(e.target.value)}
          className="border border-gray-300 rounded px-2 py-1 text-sm focus:outline-none focus:ring-1 focus:ring-green-500"
        >
          {opcoes.map(o => <option key={o} value={o}>{o}</option>)}
        </select>
      </div>
    );
  }
  return (
    <div className="flex flex-col gap-0.5">
      <label className="text-xs text-gray-500">{label}</label>
      <input
        type={tipo}
        value={valor === null || valor === undefined ? '' : String(valor)}
        placeholder={tipo === 'number' ? 'null' : ''}
        onChange={e => {
          if (tipo === 'number') {
            onChange(e.target.value === '' ? null : parseFloat(e.target.value));
          } else {
            onChange(e.target.value);
          }
        }}
        className="border border-gray-300 rounded px-2 py-1 text-sm focus:outline-none focus:ring-1 focus:ring-green-500 w-full"
      />
    </div>
  );
}

function FormAlimento({ inicial, onSalvar, onCancelar }: {
  inicial: Alimento;
  onSalvar: (a: Alimento) => void;
  onCancelar: () => void;
}) {
  const [form, setForm] = useState<Alimento>(inicial);
  const set = (key: keyof Alimento) => (v: unknown) => setForm(f => ({ ...f, [key]: v }));

  const campos: { label: string; key: keyof Alimento; tipo?: string; opcoes?: string[] }[] = [
    { label: 'Nome', key: 'nome', tipo: 'text' },
    { label: 'Custo R$/kg MN', key: 'custo' },
    { label: 'Classificação', key: 'classificacao', tipo: 'select', opcoes: ['Energético', 'Proteico', 'Volumoso', 'Mineral', 'Aditivo', 'Outro'] },
    { label: 'Tipo', key: 'tipo', tipo: 'select', opcoes: ['C', 'F', 'M'] },
    { label: 'MS (0–1)', key: 'ms' },
    { label: 'PB (0–1)', key: 'pb' },
    { label: 'PNDR (0–1)', key: 'pndr' },
    { label: 'FDN (0–1)', key: 'fdn' },
    { label: 'FDA (0–1)', key: 'fda' },
    { label: 'NEl Mcal/kg', key: 'nel' },
    { label: 'NDT (0–1)', key: 'ndt' },
    { label: 'EE (0–1)', key: 'ee' },
    { label: 'Cinzas (0–1)', key: 'cinza' },
    { label: 'Amido (0–1)', key: 'amido' },
    { label: 'kd Amido %/h', key: 'kd_amido' },
    { label: 'mn8 (0–1)', key: 'mn8' },
    { label: 'mn19 (0–1)', key: 'mn19' },
    { label: 'Ca (0–1)', key: 'ca' },
    { label: 'P (0–1)', key: 'p' },
    { label: 'Mg (0–1)', key: 'mg' },
    { label: 'K (0–1)', key: 'k' },
    { label: 'S (0–1)', key: 's' },
    { label: 'Na (0–1)', key: 'na' },
    { label: 'Cl (0–1)', key: 'cl' },
    { label: 'Co mg/kg', key: 'co' },
    { label: 'Cu mg/kg', key: 'cu' },
    { label: 'Mn mg/kg', key: 'mn_min' },
    { label: 'Zn mg/kg', key: 'zn' },
    { label: 'Se mg/kg', key: 'se' },
    { label: 'I mg/kg', key: 'i' },
    { label: 'Fe mg/kg', key: 'fe' },
    { label: 'Vit A UI/kg', key: 'vit_a' },
    { label: 'Vit D3 UI/kg', key: 'vit_d3' },
    { label: 'Vit E UI/kg', key: 'vit_e' },
    { label: 'Biotina mg/kg', key: 'biotina' },
    { label: 'Monensina mg/kg', key: 'monensina' },
    { label: 'Cr mg/kg', key: 'cr' },
    { label: 'Levedura UFC/kg', key: 'levedura' },
  ];

  return (
    <div className="fixed inset-0 bg-black/50 z-50 flex items-start justify-center overflow-y-auto py-8">
      <div className="bg-white rounded-xl shadow-xl w-full max-w-2xl mx-4">
        <div className="flex items-center justify-between p-4 border-b border-gray-200">
          <h2 className="font-bold text-gray-800">{inicial.nome ? 'Editar Alimento' : 'Novo Alimento'}</h2>
          <button onClick={onCancelar} className="p-1 hover:bg-gray-100 rounded"><X size={18} /></button>
        </div>
        <div className="p-4 grid grid-cols-2 md:grid-cols-3 gap-3 max-h-[60vh] overflow-y-auto">
          {campos.map(c => (
            <Campo
              key={c.key}
              label={c.label}
              field={c.key}
              valor={form[c.key]}
              onChange={set(c.key)}
              tipo={c.tipo}
              opcoes={c.opcoes}
            />
          ))}
        </div>
        <div className="flex justify-end gap-2 p-4 border-t border-gray-200">
          <button onClick={onCancelar} className="px-4 py-2 text-sm text-gray-600 hover:bg-gray-100 rounded-lg">Cancelar</button>
          <button
            onClick={() => form.nome ? onSalvar(form) : alert('Nome obrigatório')}
            className="flex items-center gap-1.5 px-4 py-2 bg-green-600 text-white text-sm rounded-lg hover:bg-green-700"
          >
            <Check size={15} /> Salvar
          </button>
        </div>
      </div>
    </div>
  );
}

export default function Alimentos() {
  const { alimentos, adicionarAlimento, editarAlimento, excluirAlimento } = useDieta();
  const [busca, setBusca] = useState('');
  const [filtroTipo, setFiltroTipo] = useState<'todos' | 'C' | 'F' | 'M'>('todos');
  const [editando, setEditando] = useState<Alimento | null>(null);
  const [novo, setNovo] = useState(false);

  const filtrados = alimentos.filter(a =>
    (filtroTipo === 'todos' || a.tipo === filtroTipo) &&
    a.nome.toLowerCase().includes(busca.toLowerCase())
  );

  const tipoLabel = (t: string) => t === 'C' ? 'Concentrado' : t === 'F' ? 'Forragem' : t === 'M' ? 'Mineral' : t;
  const tipoBg = (t: string) => t === 'C' ? 'bg-blue-100 text-blue-700' : t === 'F' ? 'bg-green-100 text-green-700' : 'bg-purple-100 text-purple-700';

  return (
    <div className="max-w-[1400px] mx-auto px-4 py-6">
      {(editando || novo) && (
        <FormAlimento
          inicial={editando ?? ALIMENTO_VAZIO}
          onSalvar={a => {
            if (editando) editarAlimento(editando.nome, a);
            else adicionarAlimento(a);
            setEditando(null);
            setNovo(false);
          }}
          onCancelar={() => { setEditando(null); setNovo(false); }}
        />
      )}

      <div className="flex items-center justify-between mb-4 flex-wrap gap-2">
        <h1 className="text-xl font-bold text-gray-800">🥩 Banco de Alimentos</h1>
        <button
          onClick={() => setNovo(true)}
          className="flex items-center gap-1.5 px-4 py-2 bg-green-600 text-white text-sm rounded-lg hover:bg-green-700"
        >
          <Plus size={15} /> Novo Alimento
        </button>
      </div>

      <div className="flex gap-2 mb-4 flex-wrap">
        <input
          type="text"
          placeholder="Buscar..."
          value={busca}
          onChange={e => setBusca(e.target.value)}
          className="border border-gray-300 rounded-lg px-3 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-green-500 w-64"
        />
        {(['todos', 'C', 'F', 'M'] as const).map(t => (
          <button
            key={t}
            onClick={() => setFiltroTipo(t)}
            className={`px-3 py-1.5 rounded-lg text-sm font-medium transition-colors ${
              filtroTipo === t ? 'bg-green-600 text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'
            }`}
          >
            {t === 'todos' ? 'Todos' : `${t === 'C' ? '🌽' : t === 'F' ? '🌾' : '🧂'} ${tipoLabel(t)}`}
          </button>
        ))}
        <span className="text-sm text-gray-400 self-center">{filtrados.length} alimentos</span>
      </div>

      <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm">
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 border-b border-gray-200">
              <tr>
                <th className="text-left px-4 py-2.5 font-semibold text-gray-600">Nome</th>
                <th className="text-center px-3 py-2.5 font-semibold text-gray-600">Tipo</th>
                <th className="text-left px-3 py-2.5 font-semibold text-gray-600">Classificação</th>
                <th className="text-right px-3 py-2.5 font-semibold text-gray-600">MS %</th>
                <th className="text-right px-3 py-2.5 font-semibold text-gray-600">PB %</th>
                <th className="text-right px-3 py-2.5 font-semibold text-gray-600">FDN %</th>
                <th className="text-right px-3 py-2.5 font-semibold text-gray-600">NEl</th>
                <th className="text-right px-3 py-2.5 font-semibold text-gray-600">Amido %</th>
                <th className="text-right px-3 py-2.5 font-semibold text-gray-600">Custo R$/kg</th>
                <th className="px-3 py-2.5" />
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filtrados.map(a => (
                <tr key={a.nome} className="hover:bg-gray-50">
                  <td className="px-4 py-2 font-medium text-gray-800">{a.nome}</td>
                  <td className="px-3 py-2 text-center">
                    <span className={`inline-block text-xs font-bold px-2 py-0.5 rounded-full ${tipoBg(a.tipo)}`}>
                      {tipoLabel(a.tipo)}
                    </span>
                  </td>
                  <td className="px-3 py-2 text-gray-600">{a.classificacao}</td>
                  <td className="px-3 py-2 text-right font-mono text-gray-700">{(a.ms * 100).toFixed(1)}</td>
                  <td className="px-3 py-2 text-right font-mono text-gray-700">{(a.pb * 100).toFixed(2)}</td>
                  <td className="px-3 py-2 text-right font-mono text-gray-700">{a.fdn !== null ? (a.fdn * 100).toFixed(1) : '—'}</td>
                  <td className="px-3 py-2 text-right font-mono text-gray-700">{a.nel !== null ? a.nel.toFixed(3) : a.ndt !== null ? ((0.0245 * a.ndt * 100) - 0.12).toFixed(3) : '—'}</td>
                  <td className="px-3 py-2 text-right font-mono text-gray-700">{a.amido !== null ? (a.amido * 100).toFixed(1) : '—'}</td>
                  <td className="px-3 py-2 text-right font-mono text-gray-700">{a.custo !== null ? a.custo.toFixed(3) : '—'}</td>
                  <td className="px-3 py-2">
                    <div className="flex gap-1 justify-end">
                      <button
                        onClick={() => setEditando(a)}
                        className="p-1.5 text-gray-400 hover:text-blue-600 hover:bg-blue-50 rounded transition-colors"
                        title="Editar"
                      >
                        <Pencil size={14} />
                      </button>
                      <button
                        onClick={() => {
                          if (confirm(`Excluir "${a.nome}"?`)) excluirAlimento(a.nome);
                        }}
                        className="p-1.5 text-gray-400 hover:text-red-600 hover:bg-red-50 rounded transition-colors"
                        title="Excluir"
                      >
                        <Trash2 size={14} />
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
