import { useState, useEffect } from 'react';
import { Upload, FileText, Check, Download, RefreshCw, AlertCircle, Trash2, Plus, Zap, CheckCircle2, XCircle } from 'lucide-react';

declare global {
  interface Window {
    pdfjsLib: any;
  }
}

const getGradeColorClass = (grau: string) => {
  const g = grau?.toUpperCase().trim() || '';
  if (g === '1' || g === 'PERIGOSO' || g === '2' || g === 'DEFICIENTE') return 'text-red-500 font-bold';
  if (g === '3' || g === 'PRECISA MELHORAR') return 'text-yellow-600 font-bold';
  if (g === '4' || g === 'NORMAL') return 'text-green-600 font-bold';
  if (g === '5' || g === 'DESTACOU-SE') return 'text-blue-500 font-bold';
  if (g === '6') return 'text-blue-800 font-bold';
  return 'text-slate-900 font-medium';
};

const MetaInput = ({ label, value, onChange, placeholder, maxLength, widthClass = "w-full", disabled = false }: any) => (
  <div className={widthClass}>
    <label className="block text-[11px] font-bold text-slate-500 uppercase tracking-wider mb-1">{label}</label>
    <input 
      type="text" value={value || ''} onChange={onChange} placeholder={placeholder} maxLength={maxLength} disabled={disabled}
      className={`w-full border border-slate-200 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition font-medium ${disabled ? 'bg-slate-100 text-slate-400 cursor-not-allowed opacity-70' : 'bg-slate-50 text-slate-800 hover:bg-white focus:bg-white'}`}
    />
  </div>
);

const MetaSelect = ({ label, value, onChange, options, widthClass = "w-full" }: any) => (
  <div className={widthClass}>
    <label className="block text-[11px] font-bold text-slate-500 uppercase tracking-wider mb-1">{label}</label>
    <select value={value || ''} onChange={onChange} className="w-full bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 focus:bg-white outline-none transition text-slate-800 font-medium appearance-none">
      <option value="" disabled>Selecione...</option>
      {options.map((opt: string) => <option key={opt} value={opt}>{opt}</option>)}
    </select>
  </div>
);

const MetaTextarea = ({ label, value, onChange, placeholder, widthClass = "w-full" }: any) => (
  <div className={widthClass}>
    <label className="block text-[11px] font-bold text-slate-500 uppercase tracking-wider mb-1">{label}</label>
    <textarea value={value || ''} onChange={onChange} placeholder={placeholder} className="w-full bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 focus:bg-white outline-none transition text-slate-800 font-medium resize-y min-h-[60px] leading-relaxed" />
  </div>
);

// ---------------------------------------------------------------------------------
// ARQUITETURA DEFINITIVA: MÉTODOS DE EXTRAÇÃO ESPACIAL
// ---------------------------------------------------------------------------------

const classifyToken = (text: string) => {
  if (/^(PR|RC|RM|RO)$/i.test(text)) return 'fase';
  if (/^(N\/A|N\/O|NR|--|AN\/)$/i.test(text)) return 'ignore';
  if (/^[1-6]$/.test(text)) return 'grau';
  // Reconhece números mesmo colados com traço ex: "1-"
  if (/^\d{1,3}[-–—.:]?$/.test(text)) return 'numero'; 
  // Reconhece combo encavalado ex: "RM/5"
  if (/^(PR|RC|RM|RO|--)[/\-]?(1|2|3|4|5|6|N\/A|N\/O|NR|--)$/i.test(text)) return 'fase_grau';
  return 'texto';
};

const detectColumns = (lines: any[][]) => {
  const xs: number[] = [];
  lines.forEach((l: any[]) => l.forEach((t: any) => {
    // Filtra os eixos X somente dos tokens que são notas ou fases (âncoras seguras)
    const type = classifyToken(t.text.replace(/^(\d{1,3})[-–—.:]?([A-Za-zÀ-ÿ])/i, '$1 $2').split(/\s+/)[0]);
    if (type === 'fase' || type === 'grau' || type === 'fase_grau') {
      xs.push(t.x);
    }
  }));
  
  if (xs.length === 0) return { fase: 9999, grau: 9999 };

  xs.sort((a: number, b: number) => a - b);
  const clusters: number[][] = [];
  
  xs.forEach((x: number) => {
    let found = false;
    for (let c of clusters) {
      if (Math.abs(c[0] - x) < 35) { 
        c.push(x);
        found = true;
        break;
      }
    }
    if (!found) clusters.push([x]);
  });

  const centers = clusters.map((c: number[]) => c.reduce((a, b) => a + b, 0) / c.length).sort((a, b) => a - b);

  return {
    fase: centers.length >= 2 ? centers[centers.length - 2] - 15 : (centers.length === 1 ? centers[0] - 30 : 9999),
    grau: centers.length >= 1 ? centers[centers.length - 1] - 15 : 9999
  };
};

const buildLines = (items: any[]) => {
  const map = new Map<number, any[]>();
  items.forEach((i: any) => {
    const y = Math.round(i.y / 4) * 4; 
    if (!map.has(y)) map.set(y, []);
    map.get(y)!.push(i);
  });
  return Array.from(map.entries())
    .sort((a, b) => b[0] - a[0]) 
    .map(([_, line]) => line.sort((a: any, b: any) => a.x - b.x));
};

const mergeBrokenLines = (lines: any[][]) => {
  const merged: any[][] = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const firstText = line[0]?.text || '';
    
    // Agora aceita que a linha comece com "1 - " ou "10Tráfego"
    if (!/^\s*\d{1,3}[-–—.:]?\b/.test(firstText) && merged.length > 0) {
      merged[merged.length - 1].push(...line);
      merged[merged.length - 1].sort((a: any, b: any) => a.x - b.x); 
    } else {
      merged.push([...line]);
    }
  }
  return merged;
};

const extractItemsFromLines = (lines: any[][], columns: any) => {
  const items: any[] = [];
  
  lines.forEach((line: any[]) => {
    let numero = '';
    let nome = '';
    let fase = '--';
    let grau = '';

    line.forEach((token: any) => {
      // Separa strings grudadas tipo "10Tráfego"
      const cleanText = token.text.replace(/^(\d{1,3})[-–—.:]?\s*([A-Za-zÀ-ÿ])/i, '$1 $2');
      const subTokens = cleanText.split(/\s+/);

      subTokens.forEach((subText: string) => {
        const type = classifyToken(subText);

        if (type === 'fase_grau') {
          const pgMatch = subText.match(/^(PR|RC|RM|RO|--)[/\-]?(1|2|3|4|5|6|N\/A|N\/O|NR|--)$/i);
          if (pgMatch) {
            fase = pgMatch[1].toUpperCase();
            grau = pgMatch[2].toUpperCase();
          }
        } else if (type === 'fase') {
          fase = subText.toUpperCase();
        } else if (type === 'grau') {
          grau = subText.toUpperCase();
        } else if (type === 'numero' && !numero) {
          numero = subText.replace(/[-–—.:]/g, ''); // Limpa o traço do número
        } else if (type === 'texto' || type === 'ignore') {
          if (subText.replace(/[-–—.:\s]/g, '').length > 0) {
            // Bloqueio Espacial: Só adiciona ao nome se estiver à esquerda da Fase/Grau
            const limitX = Math.min(columns.fase, columns.grau);
            if (token.x < limitX + 20) {
              nome += (nome ? ' ' : '') + subText;
            }
          }
        }
      });
    });

    nome = nome.replace(/^[-–—.:\s]+|[-–—.:\s]+$/g, '').trim();

    // Com o bug do traço resolvido, agora TODOS os itens passarão nesse if
    if (numero && nome.length >= 2) {
      items.push({ id: crypto.randomUUID(), numero, nome, fase, grau, comentario: '' });
    }
  });

  return items;
};

// ---------------------------------------------------------------------------------
// APLICAÇÃO PRINCIPAL
// ---------------------------------------------------------------------------------

export default function App() {
  const [status, setStatus] = useState('idle'); 
  const [items, setItems] = useState<any[]>([]);
  const [errorMsg, setErrorMsg] = useState('');
  const [showModal, setShowModal] = useState(false);
  const [modalState, setModalState] = useState('sending');
  const [modalMessage, setModalMessage] = useState('');

  const WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbxLlUKIeUnaLW2VeWOpIG5ZtrrAFy_Qg9YQTq5fG4HrMUg7kt196zcFAt4jOjBrMsEE/exec";

  const [meta, setMeta] = useState({ esquadrilha: '', aluno1p: '', instrutor: '', fase: '', aeronave: '', data: '', missao: '', grauMissao: '', tipoMissao: 'Normal', pousos: '', hdep: '', tev: '', parecer: '' });

  useEffect(() => {
    document.body.style.display = 'block'; document.body.style.margin = '0'; document.documentElement.style.backgroundColor = '#f8fafc';
    const rootNode = document.getElementById('root'); if (rootNode) rootNode.style.width = '100%';
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js';
    script.onload = () => { window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js'; };
    document.body.appendChild(script);
  }, []);

  const updateMeta = (field: string, value: string) => {
    if (field === 'aluno1p' || field === 'instrutor') value = value.toUpperCase().slice(0, 3);
    setMeta(prev => { const newMeta = { ...prev, [field]: value }; if (field === 'tipoMissao' && (value === 'Abortiva' || value === 'Extra')) newMeta.grauMissao = ''; return newMeta; });
  };

  const processStructuredData = (rawItems: any[], fullText: string) => {
    
    // EXTRAÇÃO DE METADADOS
    const cleanHeader = fullText.substring(0, 1500).replace(/[\n\r]/g, ' ').replace(/\s{2,}/g, ' ');
    const matchGrauMissao = cleanHeader.match(/GRAU\s*(\d{1,2})/i);
    let tipoMissaoDetectado = 'Normal';
    if (cleanHeader.match(/\bExtra\b/i)) tipoMissaoDetectado = 'Extra';
    else if (cleanHeader.match(/\bRevis[ãa]o\b/i)) tipoMissaoDetectado = 'Revisão';
    else if (cleanHeader.match(/\bAbortiva\b/i) || cleanHeader.match(/\bVMET\b/i) || cleanHeader.match(/\bVMAT\b/i)) tipoMissaoDetectado = 'Abortiva';

    const times = Array.from(cleanHeader.matchAll(/\b(\d{2}:\d{2})\b/g));
    let hdep = '', tev = '';
    if (times.length >= 2) { hdep = times[0][1]; tev = times[times.length - 1][1]; } 
    else if (times.length === 1) { if (/TEV[^\d]*\d{2}:\d{2}/i.test(cleanHeader)) tev = times[0][1]; else hdep = times[0][1]; }

    setMeta(prev => ({
      ...prev,
      missao: (cleanHeader.match(/\b((?:VMAT\s+|VMET\s+)?[A-Z]{2,4}-[A-Z0-9]{1,3})\b/i)?.[1] || '').toUpperCase(),
      data: cleanHeader.match(/\b(\d{2}\/\d{2}\/\d{4})\b/)?.[1] || '',
      fase: (cleanHeader.match(/FASE:\s*(.*?)(?=\s*ALUNO:|\s*INSTRUTOR:|\s*AERONAVE:|\s*NORMAL\b|\s*GRAU\b|$)/i)?.[1] || '').replace(/["\n\r]/g, '').trim().toUpperCase(),
      aeronave: cleanHeader.match(/(?:AERONAVE)[\s:]*(\d{4})\b/i)?.[1] || cleanHeader.match(/\b(13\d{2}|14\d{2})\b/)?.[1] || '',
      pousos: cleanHeader.match(/POUSOS[\s:]*(\d{1,2})\b/i)?.[1] || '',
      grauMissao: (tipoMissaoDetectado === 'Abortiva' || tipoMissaoDetectado === 'Extra') ? '' : (matchGrauMissao ? matchGrauMissao[1] : ''),
      tipoMissao: tipoMissaoDetectado, hdep, tev,
      parecer: (fullText.match(/Recomendações\/Parecer:\s*([\s\S]*?)(?=\bCiente\b|INSTRUTOR|Autoridade|Ass\. Digital|$)/i)?.[1] || '').replace(/[\n\r]/g, ' ').replace(/\s{2,}/g, ' ').trim()
    }));

    // ISOLAR A TABELA DE MANOBRAS E APLICAR OS MÉTODOS DO BIZU
    const normalized = rawItems.map((i: any) => ({ text: i.text.trim(), x: i.x, y: i.y })).filter((i: any) => i.text.length > 0);
    const allLines = buildLines(normalized);

    let tableLines: any[][] = [];
    let isTable = false;

    allLines.forEach((line: any[]) => {
      const text = line.map((t: any) => t.text).join(' ');
      if (/\b1\s*[-–—]?\s*(Partida|Voo sob Capota)/i.test(text)) isTable = true;
      if (/Itens Afetivos/i.test(text) || /Comentários:/i.test(text)) isTable = false;
      if (isTable) tableLines.push(line);
    });

    const leftLines: any[][] = [];
    const rightLines: any[][] = [];
    tableLines.forEach((line: any[]) => {
      const left = line.filter((t: any) => t.x < 300);
      const right = line.filter((t: any) => t.x >= 300);
      if (left.length > 0) leftLines.push(left);
      if (right.length > 0) rightLines.push(right);
    });

    const leftMerged = mergeBrokenLines(leftLines);
    const rightMerged = mergeBrokenLines(rightLines);

    const leftCols = detectColumns(leftMerged);
    const rightCols = detectColumns(rightMerged);

    // EXTRAI A TABELA COMPLETA (ITENS COM E SEM COMENTÁRIO)
    const tableItems = [
      ...extractItemsFromLines(leftMerged, leftCols),
      ...extractItemsFromLines(rightMerged, rightCols)
    ];

    // EXTRAÇÃO DE ITENS AFETIVOS
    const afetivos: any[] = [];
    const idxAfetivos = fullText.search(/Itens Afetivos/i);
    const idxComentarios = fullText.search(/Comentários:/i);
    if (idxAfetivos !== -1) {
      const section = fullText.substring(idxAfetivos, idxComentarios !== -1 ? idxComentarios : fullText.length);
      const afetivoRegex = /([A-Za-zÀ-ÿ0-9\s\-\(\)\.,\/\u0300-\u036f]+?)\s+(NORMAL|DESTACOU-SE|PRECISA MELHORAR|DEFICIENTE|ABAIXO DO PADRÃO|PERIGOSO|N\/O|N\/A|NR|--)\b/gi;
      let matchA;
      while ((matchA = afetivoRegex.exec(section)) !== null) {
        let nomeA = matchA[1].trim().replace(/^Itens Afetivos.*?Cognitivos:/i, '').replace(/^[-–—.:\s]+/, '');
        let grauA = matchA[2].toUpperCase().replace(/\s+/g, '');
        if (nomeA.length > 3 && !/^\d/.test(nomeA)) {
            afetivos.push({ id: crypto.randomUUID(), numero: '', nome: nomeA, fase: '--', grau: ['--', 'N/O', 'N/A', 'NR'].includes(grauA) ? '' : grauA, comentario: '' });
        }
      }
    }

    // EXTRAÇÃO DE COMENTÁRIOS
    const comments: any[] = [];
    const commentRegex = /(?:^|\s)(0?[1-9]|[1-5][0-9])\s*[-–—]\s*([A-Za-zÀ-ÿ0-9\s/]+?)(?:\s*\(\s*([^)]+)\s*\)\s*:?|\s*:)/gi;
    const matches: any[] = [];
    let commentMatch;
    
    while ((commentMatch = commentRegex.exec(fullText)) !== null) {
      matches.push({ 
        index: commentMatch.index, 
        length: commentMatch[0].length, 
        numero: commentMatch[1].trim(), 
        nome: commentMatch[2].trim(), 
        insideParens: commentMatch[3] ? commentMatch[3].toUpperCase() : '' 
      });
    }

    for (let i = 0; i < matches.length; i++) {
      const current = matches[i];
      const start = current.index + current.length;
      let end = i + 1 < matches.length ? matches[i+1].index : fullText.length;
      const assIdx = fullText.indexOf('Ass. Digital', start);
      const recIdx = fullText.indexOf('Recomendações/Parecer:', start);
      if (assIdx !== -1 && assIdx < end) end = assIdx;
      if (recIdx !== -1 && recIdx < end) end = recIdx;

      let fase = '--', grau = '';
      if (current.insideParens.includes('/')) {
         const parts = current.insideParens.split('/'); fase = parts[0].trim() || '--'; grau = parts[1].trim();
      } else if (current.insideParens === 'PR') fase = 'PR';
      else if (/^[1-6]$/.test(current.insideParens)) grau = current.insideParens;
      else fase = current.insideParens;

      if (['--', 'N/O', 'N/A', 'NR', 'AN/'].includes(grau.replace(/\s+/g, ''))) grau = '';
      
      comments.push({ numero: current.numero, nome: current.nome, fase, grau, comentario: fullText.substring(start, end).replace(/[\n\r]/g, ' ').replace(/\s{2,}/g, ' ').trim() });
    }

    // MERGE INTELIGENTE (Tabela + Afetivos + Comentários)
    const similarity = (a: string, b: string) => {
      a = a.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]/g, '');
      b = b.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]/g, '');
      if (a.includes(b) || b.includes(a)) return 1;
      let score = 0; const wordsA = a.split(' ').filter(w => w.length > 2); const wordsB = b.split(' ').filter(w => w.length > 2);
      if (wordsA.length === 0 || wordsB.length === 0) return 0;
      wordsA.forEach(w => { if (wordsB.includes(w)) score++; }); return score / Math.max(wordsA.length, wordsB.length);
    };

    const finalItems = [...tableItems, ...afetivos]; 
    comments.forEach((c: any) => {
      let target = finalItems.find((it: any) => it.numero === c.numero);
      if (!target) {
        const bestMatch = finalItems.map((it: any) => ({ it, score: similarity(it.nome, c.nome) })).sort((a: any, b: any) => b.score - a.score)[0];
        if (bestMatch && bestMatch.score > 0.4) target = bestMatch.it;
      }
      if (target) {
        target.comentario = c.comentario || target.comentario;
        if (c.fase !== '--' && c.fase) target.fase = c.fase;
        if (c.grau) target.grau = c.grau;
        if (!target.numero) target.numero = c.numero;
      } else {
        finalItems.push({ id: crypto.randomUUID(), numero: c.numero, nome: c.nome, fase: c.fase || '--', grau: c.grau || '', comentario: c.comentario });
      }
    });

    const cleanResults = finalItems
      .filter((i: any) => i.nome.length > 2 && !/esquadrão/i.test(i.nome))
      .sort((a: any, b: any) => parseInt(a.numero || '999') - parseInt(b.numero || '999'));

    if (cleanResults.length > 0) { setItems(cleanResults); setStatus('reviewing'); setErrorMsg(''); } 
    else { setErrorMsg('Não foi possível extrair dados estruturados.'); setStatus('idle'); }
  };

  const handleFileUpload = async (e: any) => {
    const file = e.target.files[0];
    if (!file) return; e.target.value = ''; setStatus('loading'); setErrorMsg('');
    if (file.type !== 'application/pdf') { setErrorMsg('Selecione um PDF válido.'); setStatus('idle'); return; }
    if (!window.pdfjsLib) { setErrorMsg('A biblioteca PDF.js não carregou.'); setStatus('idle'); return; }

    try {
      const arrayBuffer = await file.arrayBuffer();
      const pdf = await window.pdfjsLib.getDocument({ data: arrayBuffer }).promise;
      const rawItems: any[] = [];
      let fullText = '';

      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        
        const pageOffsetY = (pdf.numPages - i) * 3000; 

        rawItems.push(...textContent.items.map((it: any) => ({ 
          text: it.str.trim(), 
          x: it.transform[4], 
          y: it.transform[5] + pageOffsetY 
        })));
        
        let sortedForText = [...textContent.items].sort((a: any, b: any) => {
          if (Math.abs(a.transform[5] - b.transform[5]) > 5) return b.transform[5] - a.transform[5];
          return a.transform[4] - b.transform[4];
        });
        
        let lastY = null;
        for (const it of sortedForText) {
           if (lastY !== null && Math.abs(it.transform[5] - lastY) > 5) fullText += '\n';
           else if (lastY !== null) fullText += ' ';
           fullText += it.str; lastY = it.transform[5];
        }
        fullText += '\n\n';
      }
      processStructuredData(rawItems, fullText);
    } catch (err) { setErrorMsg('Erro fatal ao processar o PDF.'); setStatus('idle'); }
  };

  const updateItem = (id: string, field: string, value: string) => setItems(items.map((item: any) => item.id === id ? { ...item, [field]: value } : item));
  const removeItem = (id: string) => setItems(items.filter((item: any) => item.id !== id));
  const addNewItem = () => setItems([...items, { id: crypto.randomUUID(), numero: '', nome: 'Novo Item', fase: '--', grau: '', comentario: '' }]);

  const buildPayload = () => items.map((item: any) => ({ data: meta.data, esquadrilha: meta.esquadrilha, missao: meta.missao, grauMissao: (meta.tipoMissao === 'Abortiva' || meta.tipoMissao === 'Extra') ? '' : meta.grauMissao, aluno1p: meta.aluno1p, instrutor: meta.instrutor, faseMissao: meta.fase, aeronave: meta.aeronave, hdep: meta.hdep, pousos: meta.pousos, tev: meta.tev, parecer: meta.parecer, numero: item.numero, nome: item.nome, faseItem: item.fase, grau: item.grau, comentario: item.comentario, tipoMissao: meta.tipoMissao }));

  const exportCSV = () => {
    const hdrs = ['Data', 'Esquadrilha', 'Missão', 'Grau Missão', '1P / AL', 'IN', 'Fase Missão', 'Anv', 'H.Dep', 'Pousos', 'TEV', 'Parecer', 'Nº Item', 'Nome', 'Fase Item', 'Grau/Menção', 'Comentário', 'Tipo'];
    const csvContent = [ hdrs.join(','), ...buildPayload().map(r => `"${r.data}","${r.esquadrilha}","${r.missao}","${r.grauMissao}","${r.aluno1p}","${r.instrutor}","${r.faseMissao}","${r.aeronave}","${r.hdep}","${r.pousos}","${r.tev}","${(r.parecer || '').replace(/"/g, '""')}","${r.numero}","${r.nome}","${r.faseItem}","${r.grau}","${(r.comentario || '').replace(/"/g, '""')}","${r.tipoMissao}"`) ].join('\n');
    const blob = new Blob(['\ufeff' + csvContent], { type: 'text/csv;charset=utf-8;' });
    const lnk = document.createElement('a'); lnk.href = URL.createObjectURL(blob); lnk.setAttribute('download', `Ficha_${meta.missao || 'Extracao'}_${new Date().toISOString().slice(0,10)}.csv`); lnk.click();
  };

  const sendWebhook = async () => {
    if (!meta.esquadrilha) { setErrorMsg('Selecione a Esquadrilha.'); window.scrollTo(0, 0); return; }
    if (!meta.aluno1p || !meta.instrutor) { setErrorMsg('Preencha os Trigramas (1P e IN).'); window.scrollTo(0, 0); return; }
    setModalState('sending'); setShowModal(true);
    try {
      const res = await fetch(WEBHOOK_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(buildPayload()) });
      if (res.ok) { setModalState('success'); setModalMessage('Banco atualizado com sucesso!'); } 
      else throw new Error('Falha.');
    } catch { setModalState('error'); setModalMessage('Erro na rede. Tente exportar o CSV.'); }
  };

  const resetAfterSuccess = () => { setShowModal(false); if (modalState === 'success') { setStatus('idle'); setItems([]); setMeta({ esquadrilha: '', aluno1p: '', instrutor: '', fase: '', aeronave: '', data: '', missao: '', grauMissao: '', tipoMissao: 'Normal', pousos: '', hdep: '', tev: '', parecer: '' }); } };

  return (
    <div className="min-h-screen w-full bg-slate-50 text-slate-800 font-sans p-4 md:p-8 relative">
      {showModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/70 backdrop-blur-sm px-4">
          <div className="bg-white rounded-2xl shadow-2xl p-8 max-w-sm w-full text-center flex flex-col items-center">
            {modalState === 'sending' && (<><RefreshCw size={44} className="animate-spin text-blue-600 mb-6" /><h2 className="text-2xl font-bold mb-1">Enviando Dados</h2><p className="text-slate-500 mb-3">Gravando dados no banco...</p></>)}
            {modalState === 'success' && (<><CheckCircle2 size={52} className="text-green-500 mb-6" /><h2 className="text-2xl font-bold mb-1">Missão Concluída!</h2><p className="text-slate-500 mb-8">{modalMessage}</p><button onClick={resetAfterSuccess} className="w-full bg-slate-800 text-white font-bold py-3.5 px-6 rounded-xl text-lg">Processar Nova Ficha</button></>)}
            {modalState === 'error' && (<><XCircle size={52} className="text-red-500 mb-6" /><h2 className="text-2xl font-bold mb-1">Ops, Falhou</h2><p className="text-slate-500 mb-8">{modalMessage}</p><button onClick={() => setShowModal(false)} className="w-full bg-slate-100 text-slate-800 font-bold py-3 px-6 rounded-xl">Tentar Novamente</button></>)}
          </div>
        </div>
      )}
      <div className={`max-w-7xl mx-auto transition-opacity ${showModal ? 'opacity-25' : 'opacity-100'}`}>
        <header className="mb-8 bg-white p-6 rounded-xl shadow-sm border flex items-center gap-5">
          <div className="w-14 h-14 bg-blue-600 rounded-lg flex items-center justify-center text-white shrink-0"><FileText size={32} /></div>
          <div><h1 className="text-3xl font-bold tracking-tight">Extrator de Fichas (EIA)</h1><p className="text-slate-500 text-lg">Processamento ultrarrápido offline.</p></div>
        </header>
        {errorMsg && <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-lg text-red-700 font-medium flex gap-2"><AlertCircle className="text-red-500 shrink-0"/>{errorMsg}</div>}
        
        {status === 'idle' && (
          <div className="flex justify-center"><div className="bg-white p-16 rounded-2xl shadow-sm border flex flex-col items-center text-center w-full max-w-3xl border-slate-200 hover:border-blue-300 transition hover:shadow-lg"><Upload size={48} className="text-blue-600 mb-7" /><h2 className="text-2xl font-bold mb-4">Selecione o PDF da Ficha</h2><p className="text-slate-500 mb-8 max-w-md">Processamento local instantâneo. Arquitetura Espacial e Classificador de Tokens Integrados.</p><label className="bg-blue-600 hover:bg-blue-700 text-white px-9 py-4.5 rounded-2xl text-xl font-bold cursor-pointer flex items-center gap-3.5 transition"><FileText size={26} /> Importar Arquivo <input type="file" accept="application/pdf" className="hidden" onClick={(e: any) => e.target.value = ''} onChange={handleFileUpload} /></label></div></div>
        )}
        
        {status === 'loading' && (
          <div className="bg-white p-20 rounded-2xl shadow-sm border flex flex-col items-center text-center"><RefreshCw className="text-blue-600 animate-spin mb-5" size={44} /><h2 className="text-2xl font-bold">Processando...</h2><p className="text-slate-500 mt-2">Extraindo dados velozmente...</p></div>
        )}
        
        {status === 'reviewing' && (
          <div className="space-y-6">
            <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-7">
              <h3 className="text-xl font-bold mb-5 flex items-center gap-2.5 border-b pb-4 border-slate-100"><FileText className="text-blue-500" size={22} /> Cabeçalho e Metadados</h3>
              <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-5">
                <MetaSelect label="Esquadrilha *" value={meta.esquadrilha} onChange={(e: any) => updateMeta('esquadrilha', e.target.value)} options={['Antares', 'Vega', 'Castor', 'Sirius']} />
                <MetaInput label="1P / Aluno *" value={meta.aluno1p} onChange={(e: any) => updateMeta('aluno1p', e.target.value)} maxLength={3} placeholder="MTA" />
                <MetaInput label="Instrutor (IN) *" value={meta.instrutor} onChange={(e: any) => updateMeta('instrutor', e.target.value)} maxLength={3} placeholder="MOT" />
                <MetaInput label="Missão" value={meta.missao} onChange={(e: any) => updateMeta('missao', e.target.value)} />
                <MetaInput label="Grau Missão" value={meta.grauMissao} onChange={(e: any) => updateMeta('grauMissao', e.target.value)} disabled={meta.tipoMissao === 'Abortiva' || meta.tipoMissao === 'Extra'} />
                <MetaSelect label="Tipo Missão" value={meta.tipoMissao} onChange={(e: any) => updateMeta('tipoMissao', e.target.value)} options={['Normal', 'Abortiva', 'Extra', 'Revisão']} />
                <MetaInput label="Data" value={meta.data} onChange={(e: any) => updateMeta('data', e.target.value)} />
                <MetaInput label="Fase" value={meta.fase} onChange={(e: any) => updateMeta('fase', e.target.value)} />
                <MetaInput label="Aeronave" value={meta.aeronave} onChange={(e: any) => updateMeta('aeronave', e.target.value)} />
                <MetaInput label="H. Dep" value={meta.hdep} onChange={(e: any) => updateMeta('hdep', e.target.value)} />
                <MetaInput label="Pousos" value={meta.pousos} onChange={(e: any) => updateMeta('pousos', e.target.value)} />
                <MetaInput label="TEV" value={meta.tev} onChange={(e: any) => updateMeta('tev', e.target.value)} />
              </div>
              <div className="mt-5 pt-5 border-t border-slate-100"><MetaTextarea label="Parecer do Comandante" value={meta.parecer} onChange={(e: any) => updateMeta('parecer', e.target.value)} placeholder="Visão geral do despacho final..." /></div>
            </div>
            
            <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
              <div className="p-6 border-b flex flex-col lg:flex-row justify-between gap-4 bg-slate-50/50">
                <h2 className="text-xl font-bold flex items-center gap-2.5"><Check className="text-green-500" size={26} /> {items.length} Itens Identificados</h2>
                <div className="flex flex-wrap gap-3">
                  <button onClick={() => { setStatus('idle'); setItems([]); }} className="px-5 py-3 text-sm font-semibold border bg-white hover:bg-slate-50 rounded-xl">Cancelar</button>
                  <button onClick={exportCSV} className="px-5 py-3 text-sm font-semibold border bg-white hover:bg-slate-50 rounded-xl flex items-center gap-2"><Download size={18} /> Exportar CSV</button>
                  <button onClick={sendWebhook} className="px-9 py-3 text-lg text-white bg-blue-600 hover:bg-blue-700 rounded-xl font-bold flex items-center gap-2.5 transition active:scale-95"><Zap size={22} /> Enviar Banco de Dados</button>
                </div>
              </div>
              <div className="p-7 overflow-x-auto">
                <table className="w-full text-left border-collapse min-w-[950px]">
                  <thead><tr className="bg-slate-50 border-b border-slate-200 text-xs text-slate-500 uppercase font-bold tracking-wider"><th className="p-3 w-16">Nº</th><th className="p-3">Nome da Manobra / Item</th><th className="p-3 w-20 text-center">Fase</th><th className="p-3 w-24 text-center">Grau</th><th className="p-3">Comentário Estruturado</th><th className="p-3 w-12 text-center"></th></tr></thead>
                  <tbody className="text-sm">
                    {items.map((it) => (
                      <tr key={it.id} className="border-b border-slate-100 hover:bg-slate-50/40 text-slate-800">
                        <td className="p-3 font-mono font-bold text-slate-400"><input value={it.numero} onChange={(e) => updateItem(it.id, 'numero', e.target.value)} className="w-full bg-transparent px-1" placeholder="--" /></td>
                        <td className="p-3 font-medium"><input value={it.nome} onChange={(e) => updateItem(it.id, 'nome', e.target.value)} className={`w-full bg-transparent px-1 ${getGradeColorClass(it.grau)}`} placeholder="Item..." /></td>
                        <td className="p-3 text-center font-bold text-blue-700"><input value={it.fase} onChange={(e) => updateItem(it.id, 'fase', e.target.value)} className="w-full bg-transparent px-1 text-center" placeholder="--" /></td>
                        <td className="p-3 text-center font-extrabold"><input value={it.grau} onChange={(e) => updateItem(it.id, 'grau', e.target.value)} className={`w-full bg-transparent px-1 uppercase text-center ${getGradeColorClass(it.grau)}`} placeholder="--" /></td>
                        <td className="p-3"><textarea value={it.comentario} onChange={(e) => updateItem(it.id, 'comentario', e.target.value)} className="w-full bg-transparent border-slate-200 focus:border-blue-300 focus:bg-white text-slate-700 border px-2.5 py-1.5 rounded-lg resize-y min-h-[44px] text-xs leading-relaxed" placeholder="Comentário..." /></td>
                        <td className="p-3 text-center"><button onClick={() => removeItem(it.id)} className="p-2 hover:bg-red-50 hover:text-red-500 text-slate-300 rounded-lg"><Trash2 size={16} /></button></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <div className="mt-5 flex justify-center border-t border-slate-100 pt-5"><button onClick={addNewItem} className="flex items-center gap-2 font-semibold text-blue-600 hover:text-blue-800 px-5 py-2.5 bg-blue-50/50 rounded-lg active:scale-95 transition"><Plus size={18} /> Adicionar Manualmente</button></div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
