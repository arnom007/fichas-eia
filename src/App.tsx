import { useState, useEffect } from 'react';
import { Upload, FileText, Check, Download, RefreshCw, AlertCircle, Trash2, Plus, Zap, CheckCircle2, XCircle } from 'lucide-react';

declare global {
  interface Window {
    pdfjsLib: any;
  }
}

const getGradeColorClass = (grau: string) => {
  const g = grau.toUpperCase().trim();
  if (g === '1' || g === 'PERIGOSO' || g === '2' || g === 'DEFICIENTE') return 'text-red-500 font-bold';
  if (g === '3' || g === 'PRECISA MELHORAR') return 'text-yellow-600 font-bold';
  if (g === '4' || g === 'NORMAL') return 'text-green-600 font-bold';
  if (g === '5' || g === 'DESTACOU-SE') return 'text-blue-500 font-bold';
  if (g === '6') return 'text-blue-800 font-bold';
  return 'text-slate-900 font-medium';
};

const MetaInput = ({ label, value, onChange, placeholder, maxLength, widthClass = "w-full", disabled = false, title = "" }: any) => (
  <div className={widthClass} title={title}>
    <label className="block text-[11px] font-bold text-slate-500 uppercase tracking-wider mb-1">{label}</label>
    <input 
      type="text" 
      value={value} 
      onChange={onChange} 
      placeholder={placeholder}
      maxLength={maxLength}
      disabled={disabled}
      className={`w-full border border-slate-200 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition font-medium ${disabled ? 'bg-slate-100 text-slate-400 cursor-not-allowed opacity-70' : 'bg-slate-50 text-slate-800 hover:bg-white focus:bg-white'}`}
    />
  </div>
);

const MetaSelect = ({ label, value, onChange, options, widthClass = "w-full" }: any) => (
  <div className={widthClass}>
    <label className="block text-[11px] font-bold text-slate-500 uppercase tracking-wider mb-1">{label}</label>
    <select 
      value={value} 
      onChange={onChange} 
      className="w-full bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 focus:bg-white outline-none transition text-slate-800 font-medium appearance-none"
    >
      <option value="" disabled>Selecione...</option>
      {options.map((opt: string) => <option key={opt} value={opt}>{opt}</option>)}
    </select>
  </div>
);

const MetaTextarea = ({ label, value, onChange, placeholder, widthClass = "w-full" }: any) => (
  <div className={widthClass}>
    <label className="block text-[11px] font-bold text-slate-500 uppercase tracking-wider mb-1">{label}</label>
    <textarea 
      value={value} 
      onChange={onChange} 
      placeholder={placeholder}
      className="w-full bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 focus:bg-white outline-none transition text-slate-800 font-medium resize-y min-h-[60px] leading-relaxed"
    />
  </div>
);

const normalizeStringForMatch = (str: string) => {
  return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]/g, '');
};

export default function App() {
  const [status, setStatus] = useState('idle'); 
  const [items, setItems] = useState<any[]>([]);
  const [errorMsg, setErrorMsg] = useState('');
  
  const [showModal, setShowModal] = useState(false);
  const [modalState, setModalState] = useState('sending');
  const [modalMessage, setModalMessage] = useState('');

  const WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbxLlUKIeUnaLW2VeWOpIG5ZtrrAFy_Qg9YQTq5fG4HrMUg7kt196zcFAt4jOjBrMsEE/exec";

  const [meta, setMeta] = useState({
    esquadrilha: '', aluno1p: '', instrutor: '', fase: '', aeronave: '', data: '', missao: '',
    grauMissao: '', tipoMissao: 'Normal', pousos: '', hdep: '', tev: '', parecer: ''
  });

  useEffect(() => {
    document.body.style.display = 'block';
    document.body.style.margin = '0';
    document.documentElement.style.backgroundColor = '#f8fafc';
    const rootNode = document.getElementById('root');
    if (rootNode) rootNode.style.width = '100%';

    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js';
    script.onload = () => { window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js'; };
    document.body.appendChild(script);
  }, []);

  const updateMeta = (field: string, value: string) => {
    if (field === 'aluno1p' || field === 'instrutor') value = value.toUpperCase().slice(0, 3);
    setMeta(prev => {
      const newMeta = { ...prev, [field]: value };
      if (field === 'tipoMissao' && (value === 'Abortiva' || value === 'Extra')) newMeta.grauMissao = '';
      return newMeta;
    });
  };

  const processTextData = (text: string) => {
    // 1. Limpeza Bruta do Lixo de PDF
    const safeText = text
      .replace(/MATERIAL DE ACESSO RESTRITO/gi, '')
      .replace(/Art\. 44.*?2012/gi, '')
      .replace(/--- PAGE \d+ ---/gi, '')
      .replace(/\b\d+\s+de\s+\d+\b/gi, '')
      .replace(/COMANDO DA AERONÁUTICA/gi, '')
      .replace(/1 ESQUADRÃO DE INSTRUÇÃO AÉREA/gi, '')
      .replace(/T-27 BÁSICO 20\d{2}/gi, '')
      .replace(/PROT[\s.:]*\d+/gi, '')
      .replace(/Itens Afetivos - Cognitivos:/gi, 'Itens Afetivos:');

    // 2. ZONAMENTO - Separar as áreas com precisão militar
    const idxTableStart = safeText.search(/\b1\s*[-–—]?\s*(?:Partida|Voo sob Capota)/i);
    const idxAfetivos = safeText.search(/Itens Afetivos/i);
    const idxComentarios = safeText.search(/Comentários:/i);

    if (idxTableStart === -1) {
      setErrorMsg('Não foi possível identificar o início da tabela de manobras na ficha.');
      setStatus('idle');
      return;
    }

    // Isolar os blocos
    const headerText = safeText.substring(0, idxTableStart).replace(/[\n\r]/g, ' ').replace(/\s{2,}/g, ' ');
    const tableText = safeText.substring(idxTableStart, idxAfetivos !== -1 ? idxAfetivos : (idxComentarios !== -1 ? idxComentarios : safeText.length));
    const afetivosText = idxAfetivos !== -1 ? safeText.substring(idxAfetivos, idxComentarios !== -1 ? idxComentarios : safeText.length) : '';
    const comentariosText = idxComentarios !== -1 ? safeText.substring(idxComentarios) : '';

    // =====================================================================
    // ETAPA 0: METADADOS (Somente no headerText)
    // =====================================================================
    const matchGrauMissao = headerText.match(/GRAU\s*(\d{1,2})/i);
    const grauMissao = matchGrauMissao ? matchGrauMissao[1] : '';

    let tipoMissaoDetectado = 'Normal';
    if (headerText.match(/\bExtra\b/i)) tipoMissaoDetectado = 'Extra';
    else if (headerText.match(/\bRevis[ãa]o\b/i)) tipoMissaoDetectado = 'Revisão';
    else if (headerText.match(/\bAbortiva\b/i) || headerText.match(/\bVMET\b/i) || headerText.match(/\bVMAT\b/i)) tipoMissaoDetectado = 'Abortiva';

    const matchMissao = headerText.match(/\b((?:VMAT\s+|VMET\s+)?[A-Z]{2,4}-[A-Z0-9]{1,3})\b/i);
    const missao = matchMissao ? matchMissao[1].toUpperCase() : '';

    const matchData = headerText.match(/\b(\d{2}\/\d{2}\/\d{4})\b/);
    const data = matchData ? matchData[1] : '';

    const times = Array.from(headerText.matchAll(/\b(\d{2}:\d{2})\b/g));
    let hdep = '', tev = '';
    if (times.length >= 2) { hdep = times[0][1]; tev = times[times.length - 1][1]; } 
    else if (times.length === 1) { if (/TEV[^\d]*\d{2}:\d{2}/i.test(headerText)) tev = times[0][1]; else hdep = times[0][1]; }

    let fase = '';
    const matchFase = headerText.match(/FASE:\s*(.*?)(?=\s*ALUNO:|\s*INSTRUTOR:|\s*AERONAVE:|\s*NORMAL\b|\s*GRAU\b|$)/i);
    if (matchFase) fase = matchFase[1].replace(/["\n\r]/g, '').replace(/^[-:]+|[-:]+$/g, '').trim().toUpperCase();

    const matchAeronave = headerText.match(/(?:AERONAVE)[\s:]*(\d{4})\b/i) || headerText.match(/\b(13\d{2}|14\d{2})\b/);
    const aeronave = matchAeronave ? matchAeronave[1] : '';

    const matchPousos = headerText.match(/POUSOS[\s:]*(\d{1,2})\b/i);
    const pousos = matchPousos ? matchPousos[1] : '';

    const parecerMatch = safeText.match(/Recomendações\/Parecer:\s*([\s\S]*?)(?=\bCiente\b|INSTRUTOR|Autoridade|Ass\. Digital|$)/i);
    const parecerStr = parecerMatch ? parecerMatch[1].replace(/[\n\r]/g, ' ').replace(/\s{2,}/g, ' ').trim() : '';

    setMeta(prev => ({ esquadrilha: prev.esquadrilha, aluno1p: prev.aluno1p, instrutor: prev.instrutor, fase, aeronave, data, missao, grauMissao: (tipoMissaoDetectado === 'Abortiva' || tipoMissaoDetectado === 'Extra') ? '' : grauMissao, tipoMissao: tipoMissaoDetectado, pousos, hdep, tev, parecer: parecerStr }));

    // =====================================================================
    // ETAPA 1: O ESQUELETO DA TABELA (Fatiamento Inteligente)
    // =====================================================================
    const extractedSkeletonItems: any[] = [];
    
    // 1. Achata a tabela inteira e garante espaçamento após números 
    // Ex: "10Tráfego" vira "10 Tráfego"
    let cleanTableText = tableText.replace(/[\n\r]/g, ' ').replace(/\s{2,}/g, ' ');
    cleanTableText = cleanTableText.replace(/\b([1-9]|[1-5][0-9])\b([A-Za-zÀ-ÿ])/g, '$1 $2');

    // 2. Fatiador Cirúrgico: Quebra exatamente antes de cada número de manobra (1 a 59)
    const chunkRegex = /(?=\b(?:[1-9]|[1-5][0-9])\b\s*[-–—]?\s+[A-Za-zÀ-ÿ])/;
    const chunks = cleanTableText.split(chunkRegex);

    for (let chunk of chunks) {
      chunk = chunk.trim();
      if (!chunk) continue;

      // Extrai o Número e o Resto
      const numMatch = chunk.match(/^([1-9]|[1-5][0-9])\b\s*[-–—]?\s+(.*)$/);
      if (!numMatch) continue;

      const num = numMatch[1];
      const rest = numMatch[2];

      // Busca a Fase e o Grau estritamente no final da string do item
      const endMatch = rest.match(/(.*?)\s+(PR|(?:RC|RM|RO|--)\s*(?:[1-6]|N\/O|N\/A|NR|A\s*N\/|--)|(?:RC|RM|RO|--)|(?:[1-6]|N\/O|N\/A|NR|A\s*N\/|--))\s*$/i);
      
      let name = rest;
      let rawPG = '';

      if (endMatch) {
          name = endMatch[1].trim();
          rawPG = endMatch[2].trim().toUpperCase();
      }

      name = name.replace(/^[-–—.:\s]+|[-–—.:\s]+$/g, '');

      let faseItem = '--';
      let grauItem = '';

      // Interpreta a Nota/Fase (Resolve o problema da VI-01)
      if (rawPG === 'PR') {
          faseItem = 'PR';
      } else if (/^(RC|RM|RO|--)$/.test(rawPG)) {
          faseItem = rawPG;
      } else if (/^([1-6]|N\/O|N\/A|NR|A\s*N\/)$/.test(rawPG.replace(/\s+/g, ''))) {
          grauItem = rawPG.replace(/\s+/g, '');
      } else {
          const pgParts = rawPG.split(/\s+/);
          if (pgParts.length >= 2) {
              faseItem = pgParts[0];
              grauItem = pgParts[1].replace(/\s+/g, '');
          }
      }

      if (['--', 'N/O', 'N/A', 'NR', 'AN/', 'NÃOOBSERVADO'].includes(grauItem)) grauItem = '';

      extractedSkeletonItems.push({ id: crypto.randomUUID(), numero: num, nome: name, fase: faseItem, grau: grauItem, comentario: '' });
    }

    // =====================================================================
    // ETAPA 2: ITENS AFETIVOS
    // =====================================================================
    if (afetivosText) {
      const cleanAfetivos = afetivosText.replace(/[\n\r]/g, ' ').replace(/\s{2,}/g, ' ');
      const afetivoRegex = /([A-Za-zÀ-ÿ0-9\s\-\(\)\.,\/\u0300-\u036f]+?)\s+(NORMAL|DESTACOU-SE|PRECISA MELHORAR|DEFICIENTE|ABAIXO DO PADRÃO|PERIGOSO|N\/O|N\/A|NR|--)\b/gi;
      let matchA;

      while ((matchA = afetivoRegex.exec(cleanAfetivos)) !== null) {
        let name = matchA[1].trim();
        const grade = matchA[2].toUpperCase().replace(/\s+/g, '');

        if (name.length > 3 && !/^\d/.test(name) && !name.toLowerCase().includes('itens afetivos') && !name.toLowerCase().includes('cognitivos')) {
          extractedSkeletonItems.push({ 
            id: crypto.randomUUID(), numero: '', nome: name, fase: '--', 
            grau: ['--', 'N/O', 'N/A', 'NR'].includes(grade) ? '' : grade, comentario: '' 
          });
        }
      }
    }

    // =====================================================================
    // ETAPA 3: O MATCH DOS COMENTÁRIOS
    // =====================================================================
    if (comentariosText) {
      const cleanCommentsTextBlock = comentariosText.replace(/[\n\r]/g, ' ').replace(/\s{2,}/g, ' ');
      const allCommentMatches: any[] = [];
      const commentRegex = /(?:^|\s)(0?[1-9]|[1-5][0-9])\s*[-–—]\s*([A-Za-zÀ-ÿ0-9\s/]+?)(?:\s*\(\s*([^)]+)\s*\)\s*:?|\s*:)/gi;
      let matchC;
      
      while ((matchC = commentRegex.exec(cleanCommentsTextBlock)) !== null) {
        allCommentMatches.push({ index: matchC.index, length: matchC[0].length, numero: matchC[1].trim(), nome: matchC[2].trim() });
      }

      allCommentMatches.sort((a, b) => a.index - b.index);

      for (let i = 0; i < allCommentMatches.length; i++) {
        const cMatch = allCommentMatches[i];
        const startIndex = cMatch.index + cMatch.length;
        let endIndex = (i + 1 < allCommentMatches.length) ? allCommentMatches[i + 1].index : cleanCommentsTextBlock.length;
        
        const assDigitalIndex = cleanCommentsTextBlock.indexOf('Ass. Digital', startIndex);
        const recomendacoesIndex = cleanCommentsTextBlock.indexOf('Recomendações/Parecer:', startIndex);
        if (assDigitalIndex !== -1 && assDigitalIndex < endIndex) endIndex = assDigitalIndex;
        if (recomendacoesIndex !== -1 && recomendacoesIndex < endIndex) endIndex = recomendacoesIndex;

        const comentarioText = cleanCommentsTextBlock.substring(startIndex, endIndex).trim();

        // Faz o Match com o Esqueleto
        let targetItem = extractedSkeletonItems.find(item => item.numero === cMatch.numero);
        if (!targetItem) {
          targetItem = extractedSkeletonItems.find(item => {
             const n1 = normalizeStringForMatch(item.nome);
             const n2 = normalizeStringForMatch(cMatch.nome);
             return (n1.length > 3 && n2.length > 3) && (n1.includes(n2) || n2.includes(n1));
          });
        }

        if (targetItem) {
          targetItem.comentario = comentarioText;
          if (!targetItem.numero) targetItem.numero = cMatch.numero;
        } else {
          // Se o item tem comentário, mas sumiu da tabela por alguma anomalia, adiciona mesmo assim
          extractedSkeletonItems.push({ id: crypto.randomUUID(), numero: cMatch.numero, nome: cMatch.nome, fase: '--', grau: '', comentario: comentarioText });
        }
      }
    }

    extractedSkeletonItems.sort((a: any, b: any) => {
      const numA = parseInt(a.numero);
      const numB = parseInt(b.numero);
      if (isNaN(numA) && isNaN(numB)) return 0;
      if (isNaN(numA)) return 1; 
      if (isNaN(numB)) return -1;
      return numA - numB;
    });

    if (extractedSkeletonItems.length > 0) { setItems(extractedSkeletonItems); setStatus('reviewing'); setErrorMsg(''); } 
    else { setErrorMsg('Não foi possível extrair os itens. O PDF pode não estar no padrão esperado.'); setStatus('idle'); }
  };

  const handleFileUpload = async (e: any) => {
    const file = e.target.files[0];
    if (!file) return;
    e.target.value = ''; setStatus('loading'); setErrorMsg('');
    if (file.type === 'application/pdf') {
      if (!window.pdfjsLib) { setErrorMsg('Biblioteca PDF não carregada. Recarregue a página.'); setStatus('idle'); return; }
      try {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await window.pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        let fullParsedText = '';
        
        for (let i = 1; i <= pdf.numPages; i++) {
          const page = await pdf.getPage(i);
          const textContent = await page.getTextContent();
          const itemsArray = textContent.items as any[];
          
          // Leitura linear clássica e segura (Sem X/Y Guessing)
          itemsArray.sort((a, b) => {
             if (Math.abs(a.transform[5] - b.transform[5]) > 5) {
                 return b.transform[5] - a.transform[5];
             }
             return a.transform[4] - b.transform[4];
          });

          let lastY = null;
          let pageText = '';
          for (const item of itemsArray) {
             if (lastY !== null && Math.abs(item.transform[5] - lastY) > 5) {
                 pageText += '\n';
             } else if (lastY !== null) {
                 pageText += ' ';
             }
             pageText += item.str;
             lastY = item.transform[5];
          }
          fullParsedText += pageText + '\n\n';
        }
        processTextData(fullParsedText);
      } catch (err) { setErrorMsg('Erro fatal ao ler o PDF.'); setStatus('idle'); }
    } else { setErrorMsg('Por favor, selecione um arquivo PDF válido.'); setStatus('idle'); }
  };

  const updateItem = (id: string, field: string, value: string) => setItems(items.map((item: any) => item.id === id ? { ...item, [field]: value } : item));
  const removeItem = (id: string) => setItems(items.filter((item: any) => item.id !== id));
  const addNewItem = () => setItems([...items, { id: crypto.randomUUID(), numero: '', nome: 'Novo Item Manual', fase: '', grau: '', comentario: '' }]);

  const buildPayload = () => items.map(item => ({ data: meta.data, esquadrilha: meta.esquadrilha, missao: meta.missao, grauMissao: (meta.tipoMissao === 'Abortiva' || meta.tipoMissao === 'Extra') ? '' : meta.grauMissao, aluno1p: meta.aluno1p, instrutor: meta.instrutor, faseMissao: meta.fase, aeronave: meta.aeronave, hdep: meta.hdep, pousos: meta.pousos, tev: meta.tev, parecer: meta.parecer, numero: item.numero, nome: item.nome, faseItem: item.fase, grau: item.grau, comentario: item.comentario, tipoMissao: meta.tipoMissao }));

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

  const resetAfterSuccess = () => { setShowModal(false); if (modalState === 'success') { setStatus('idle'); setItems([]); } };

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
          <div><h1 className="text-3xl font-bold tracking-tight">Extrator de Fichas (EIA)</h1><p className="text-slate-500 text-lg">Tratamento inteligente de dados de instrução.</p></div>
        </header>
        {errorMsg && <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-lg text-red-700 font-medium flex gap-2"><AlertCircle className="text-red-500 shrink-0"/>{errorMsg}</div>}
        {status === 'idle' && (
          <div className="flex justify-center"><div className="bg-white p-16 rounded-2xl shadow-sm border flex flex-col items-center text-center w-full max-w-3xl border-slate-200 hover:border-blue-300 transition hover:shadow-lg"><Upload size={48} className="text-blue-600 mb-7" /><h2 className="text-2xl font-bold mb-4">Selecione o PDF da Ficha</h2><p className="text-slate-500 mb-8 max-w-md">Envie o arquivo PDF original. O sistema analisa tabelas superiores, afetivos e comentários automaticamente.</p><label className="bg-blue-600 hover:bg-blue-700 text-white px-9 py-4.5 rounded-2xl text-xl font-bold cursor-pointer flex items-center gap-3.5 transition"><FileText size={26} /> Importar Arquivo <input type="file" accept="application/pdf" className="hidden" onClick={(e: any) => e.target.value = ''} onChange={handleFileUpload} /></label></div></div>
        )}
        {status === 'loading' && (
          <div className="bg-white p-20 rounded-2xl shadow-sm border flex flex-col items-center text-center"><RefreshCw className="text-blue-600 animate-spin mb-5" size={44} /><h2 className="text-2xl font-bold">Analisando Estrutura</h2><p className="text-slate-500">Separando colunas, criando esqueleto e vinculando comentários...</p></div>
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
                <h2 className="text-xl font-bold flex items-center gap-2.5"><Check className="text-green-500" size={26} /> {items.length} Itens Extraídos e Vinculados</h2>
                <div className="flex flex-wrap gap-3">
                  <button onClick={() => { setStatus('idle'); setItems([]); }} className="px-5 py-3 text-sm font-semibold border bg-white hover:bg-slate-50 rounded-xl">Voltar / Cancelar</button>
                  <button onClick={exportCSV} className="px-5 py-3 text-sm font-semibold border bg-white hover:bg-slate-50 rounded-xl flex items-center gap-2"><Download size={18} /> Exportar CSV</button>
                  <button onClick={sendWebhook} className="px-9 py-3 text-lg text-white bg-blue-600 hover:bg-blue-700 rounded-xl font-bold flex items-center gap-2.5 transition active:scale-95"><Zap size={22} /> Enviar Banco de Dados</button>
                </div>
              </div>
              <div className="p-7 overflow-x-auto">
                <table className="w-full text-left border-collapse min-w-[950px]">
                  <thead><tr className="bg-slate-50 border-b border-slate-200 text-xs text-slate-500 uppercase font-bold tracking-wider"><th className="p-3 w-16">Nº</th><th className="p-3">Nome da Manobra / Item</th><th className="p-3 w-20 text-center">Fase</th><th className="p-3 w-24 text-center">Grau</th><th className="p-3">Comentário Vinculado</th><th className="p-3 w-12 text-center"></th></tr></thead>
                  <tbody className="text-sm">
                    {items.map((it) => (
                      <tr key={it.id} className="border-b border-slate-100 hover:bg-slate-50/40 text-slate-800">
                        <td className="p-3 font-mono font-bold text-slate-400"><input value={it.numero} onChange={(e) => updateItem(it.id, 'numero', e.target.value)} className="w-full bg-transparent px-1" placeholder="--" /></td>
                        <td className="p-3 font-medium"><input value={it.nome} onChange={(e) => updateItem(it.id, 'nome', e.target.value)} className={`w-full bg-transparent px-1 ${getGradeColorClass(it.grau)}`} placeholder="Nome do item manual..." /></td>
                        <td className="p-3 text-center font-bold text-blue-700"><input value={it.fase} onChange={(e) => updateItem(it.id, 'fase', e.target.value)} className="w-full bg-transparent px-1 text-center" placeholder="--" /></td>
                        <td className="p-3 text-center font-extrabold"><input value={it.grau} onChange={(e) => updateItem(it.id, 'grau', e.target.value)} className={`w-full bg-transparent px-1 uppercase text-center ${getGradeColorClass(it.grau)}`} placeholder="--" /></td>
                        <td className="p-3"><textarea value={it.comentario} onChange={(e) => updateItem(it.id, 'comentario', e.target.value)} className="w-full bg-transparent border-slate-200 focus:border-blue-300 focus:bg-white text-slate-700 border px-2.5 py-1.5 rounded-lg resize-y min-h-[44px] text-xs leading-relaxed" placeholder="Adicionar comentário..." /></td>
                        <td className="p-3 text-center"><button onClick={() => removeItem(it.id)} className="p-2 hover:bg-red-50 hover:text-red-500 text-slate-300 rounded-lg"><Trash2 size={16} /></button></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <div className="mt-5 flex justify-center border-t border-slate-100 pt-5"><button onClick={addNewItem} className="flex items-center gap-2 font-semibold text-blue-600 hover:text-blue-800 px-5 py-2.5 bg-blue-50/50 rounded-lg active:scale-95 transition"><Plus size={18} /> Adicionar Item Manual</button></div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
