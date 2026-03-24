import { useState, useEffect } from 'react';
import { Upload, FileText, Check, Download, RefreshCw, AlertCircle, Trash2, Plus, Zap, CheckCircle2, XCircle } from 'lucide-react';

// Diz ao TypeScript que a biblioteca pdfjsLib existe globalmente na janela do navegador
declare global {
  interface Window {
    pdfjsLib: any;
  }
}

// Função para definir a cor baseada no grau/nota
const getGradeColorClass = (grau: string) => {
  const g = grau.toUpperCase().trim();
  if (g === '1' || g === 'PERIGOSO' || g === '2' || g === 'DEFICIENTE') return 'text-red-500 font-bold';
  if (g === '3' || g === 'PRECISA MELHORAR') return 'text-yellow-600 font-bold';
  if (g === '4' || g === 'NORMAL') return 'text-green-600 font-bold';
  if (g === '5' || g === 'DESTACOU-SE') return 'text-blue-500 font-bold';
  if (g === '6') return 'text-blue-800 font-bold';
  return 'text-slate-900 font-medium'; // Padrão e Vazio
};

// Componente MetaInput
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

export default function App() {
  const [status, setStatus] = useState('idle'); 
  const [items, setItems] = useState<any[]>([]);
  const [errorMsg, setErrorMsg] = useState('');
  
  const [showModal, setShowModal] = useState(false);
  const [modalState, setModalState] = useState('sending'); 
  const [modalMessage, setModalMessage] = useState('');

  const WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbxLlUKIeUnaLW2VeWOpIG5ZtrrAFy_Qg9YQTq5fG4HrMUg7kt196zcFAt4jOjBrMsEE/exec";

  const [meta, setMeta] = useState({
    esquadrilha: '',
    aluno1p: '',
    instrutor: '',
    fase: '',
    aeronave: '',
    data: '',
    missao: '',
    grauMissao: '', 
    tipoMissao: 'Normal', 
    pousos: '',
    hdep: '',
    tev: '',
    parecer: ''
  });

  useEffect(() => {
    document.body.style.display = 'block';
    document.body.style.margin = '0';
    document.documentElement.style.backgroundColor = '#f8fafc'; 
    
    const rootNode = document.getElementById('root');
    if (rootNode) {
      rootNode.style.width = '100%';
      rootNode.style.minHeight = '100vh';
      rootNode.style.maxWidth = 'none'; 
      rootNode.style.padding = '0';
      rootNode.style.margin = '0';
      rootNode.style.textAlign = 'left';
    }

    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js';
    script.onload = () => {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';
    };
    document.body.appendChild(script);
  }, []);

  const updateMeta = (field: string, value: string) => {
    if (field === 'aluno1p' || field === 'instrutor') {
      value = value.toUpperCase().slice(0, 3);
    }
    
    setMeta(prev => {
      const newMeta = { ...prev, [field]: value };
      if (field === 'tipoMissao' && (value === 'Abortiva' || value === 'Extra')) {
        newMeta.grauMissao = '';
      }
      return newMeta;
    });
  };

  const processTextData = (text: string) => {
    // 1. LIMPEZA GLOBAL
    // Aniquila os falsos positivos (Art 44, páginas, cabeçalhos perdidos) em todo o documento!
    let globalCleanText = text
      .replace(/MATERIAL DE ACESSO RESTRITO/gi, '')
      .replace(/Art\. 44 e Art\. 45 do Decreto.*?2012/gi, '')
      .replace(/--- PAGE \d+ ---/gi, '')
      .replace(/\b\d+\s+de\s+\d+\b/gi, '')
      .replace(/COMANDO DA AERONÁUTICA/gi, '')
      .replace(/1 ESQUADRÃO DE INSTRUÇÃO AÉREA/gi, '')
      .replace(/T-27 BÁSICO 20\d{2}/gi, '')
      .replace(/PROT[\s.:]*\d+/gi, ''); 

    // 2. EXTRAÇÃO DE METADADOS
    const headerEndIdx = globalCleanText.search(/1\s*-|Itens Afetivos|Comentários:/i);
    const headerText = headerEndIdx !== -1 ? globalCleanText.substring(0, headerEndIdx) : globalCleanText;
    const singleLineHeader = headerText.replace(/[\n\r]/g, ' ').replace(/\s{2,}/g, ' ');

    const matchGrauMissao = singleLineHeader.match(/GRAU\s*(\d{1,2})/i);
    const grauMissao = matchGrauMissao ? matchGrauMissao[1] : '';

    let tipoMissaoDetectado = 'Normal';
    if (singleLineHeader.match(/\bExtra\b/i)) tipoMissaoDetectado = 'Extra';
    else if (singleLineHeader.match(/\bRevis[ãa]o\b/i)) tipoMissaoDetectado = 'Revisão';
    else if (singleLineHeader.match(/\bAbortiva\b/i) || singleLineHeader.match(/\bVMET\b/i) || singleLineHeader.match(/\bVMAT\b/i)) {
      tipoMissaoDetectado = 'Abortiva';
    }

    const matchMissao = singleLineHeader.match(/\b((?:VMAT\s+|VMET\s+)?[A-Z]{2,4}-[A-Z0-9]{1,3})\b/i);
    const missao = matchMissao ? matchMissao[1].toUpperCase() : '';

    const matchData = singleLineHeader.match(/\b(\d{2}\/\d{2}\/\d{4})\b/);
    const data = matchData ? matchData[1] : '';

    const times = Array.from(singleLineHeader.matchAll(/\b(\d{2}:\d{2})\b/g));
    let hdep = '', tev = '';
    if (times.length >= 2) {
      hdep = times[0][1];
      tev = times[times.length - 1][1]; 
    } else if (times.length === 1) {
      if (/TEV[^\d]*\d{2}:\d{2}/i.test(singleLineHeader)) tev = times[0][1];
      else hdep = times[0][1];
    }

    let fase = '';
    const matchFase = singleLineHeader.match(/FASE:\s*(.*?)(?=\s*ALUNO:|\s*INSTRUTOR:|\s*AERONAVE:|\s*NORMAL\b|\s*GRAU\b|$)/i);
    if (matchFase) {
      fase = matchFase[1].replace(/["\n\r]/g, '').replace(/^[-:]+|[-:]+$/g, '').trim().toUpperCase();
    }

    const matchAeronave = singleLineHeader.match(/(?:AERONAVE)[\s:]*(\d{4})\b/i) || singleLineHeader.match(/\b(13\d{2}|14\d{2})\b/);
    const aeronave = matchAeronave ? matchAeronave[1] : '';

    const matchPousos = singleLineHeader.match(/POUSOS[\s:]*(\d{1,2})\b/i);
    const pousos = matchPousos ? matchPousos[1] : '';

    const parecerMatch = globalCleanText.match(/Recomendações\/Parecer:\s*([\s\S]*?)(?=Ciente|INSTRUTOR do voo subsequente|Autoridade Competente|Ass\. Digital|$)/i);
    let parecerStr = parecerMatch ? parecerMatch[1].replace(/\n/g, ' ').replace(/\s{2,}/g, ' ').trim() : '';

    setMeta(prev => ({
      esquadrilha: prev.esquadrilha, 
      aluno1p: prev.aluno1p,         
      instrutor: prev.instrutor,     
      fase, aeronave, data, missao,
      grauMissao: (tipoMissaoDetectado === 'Abortiva' || tipoMissaoDetectado === 'Extra') ? '' : grauMissao,
      tipoMissao: tipoMissaoDetectado,
      pousos, hdep, tev,
      parecer: parecerStr
    }));

    // 3. EXTRAÇÃO DAS TABELAS (O "Fatiador Blindado")
    const extractedItemsMap = new Map();

    const idxAfetivos = globalCleanText.search(/Itens Afetivos/i);
    const idxComentarios = globalCleanText.search(/Comentários:/i);

    let tableStartIndex = 0;
    const firstItemRegex = /(?:^|\n|\r|\s)\b(1\s*-?\s*[A-Za-zÀ-ÿ])/;
    const firstItemMatch = globalCleanText.substring(0, idxAfetivos !== -1 ? idxAfetivos : globalCleanText.length).match(firstItemRegex);
    if (firstItemMatch) {
      tableStartIndex = firstItemMatch.index! + firstItemMatch[0].indexOf('1');
    }

    let tableEndIndex = globalCleanText.length;
    if (idxAfetivos !== -1) tableEndIndex = idxAfetivos;
    else if (idxComentarios !== -1) tableEndIndex = idxComentarios;

    let cleanTableText = globalCleanText.substring(tableStartIndex, tableEndIndex).replace(/[\n\r]/g, ' ').replace(/\s{2,}/g, ' ').trim();

    const itemBoundaryRegex = /(?:^|\s)(0?[1-9]|[1-9][0-9])\s*[-–—]?\s+(?=[A-Za-zÀ-ÿ])/g;
    let matchT1;
    let lastIndex = 0;
    let lastNum = null;
    const chunks = [];

    while ((matchT1 = itemBoundaryRegex.exec(cleanTableText)) !== null) {
      if (lastNum !== null) {
        chunks.push({ num: lastNum, text: cleanTableText.substring(lastIndex, matchT1.index).trim() });
      }
      lastNum = matchT1[1];
      lastIndex = matchT1.index + matchT1[0].length;
      itemBoundaryRegex.lastIndex = matchT1.index + matchT1[1].length; 
    }
    if (lastNum !== null) {
      chunks.push({ num: lastNum, text: cleanTableText.substring(lastIndex).trim() });
    }

    for (const chunk of chunks) {
      let name = chunk.text;
      let faseItem = '--';
      let grauItem = '';

      const endMatchRegex = /(?:\s+)(?:(PR)|(RC\b|RM\b|RO\b|--)\s+([1-6]\b|N\/O\b|N\/A\b|NR\b|A\s*N\/|--)|([1-6]\b|N\/O\b|N\/A\b|NR\b|A\s*N\/|--))\s*$/i;
      const matchPG = name.match(endMatchRegex);

      if (matchPG) {
        name = name.substring(0, matchPG.index).trim(); 
        if (matchPG[1]) {
          faseItem = 'PR';
        } else if (matchPG[2]) {
          faseItem = matchPG[2].toUpperCase();
          grauItem = matchPG[3].toUpperCase();
        } else if (matchPG[4]) {
          faseItem = '--';
          grauItem = matchPG[4].toUpperCase();
        }
      }

      const grauNorm = grauItem.replace(/\s+/g, '');
      if (['--', 'N/O', 'N/A', 'NR', 'AN/', 'NÃOOBSERVADO'].includes(grauNorm)) {
        grauItem = '';
      }

      extractedItemsMap.set(chunk.num, {
        id: crypto.randomUUID(),
        numero: chunk.num,
        nome: name,
        fase: faseItem,
        grau: grauItem,
        comentario: ''
      });
    }

    // 4. ITENS AFETIVOS
    let affectiveAreaText = '';
    if (idxAfetivos !== -1) {
      const affEndIndex = idxComentarios !== -1 ? idxComentarios : globalCleanText.length;
      affectiveAreaText = globalCleanText.substring(idxAfetivos, affEndIndex).replace(/[\n\r]/g, ' ').replace(/\s{2,}/g, ' ');
      
      const affectiveRegex = /(?:^|\s)([A-Za-zÀ-ÿ0-9\s\-\(\)\.,\/\u0300-\u036f]{4,80}?)\s+(NORMAL|DESTACOU-SE|NÃO OBSERVADO|PRECISA MELHORAR|DEFICIENTE|ABAIXO DO PADRÃO|PERIGOSO|N\/O|N\/A|NR|--)(?=\s|$|\b\d{1,2})/gi;
      let matchT2;
      
      while ((matchT2 = affectiveRegex.exec(affectiveAreaText)) !== null) {
        let itemName = matchT2[1].trim();
        let itemGrau = matchT2[2].toUpperCase();

        const itemGrauNorm = itemGrau.replace(/\s+/g, '');
        if (['--', 'N/O', 'N/A', 'NR', 'NÃOOBSERVADO'].includes(itemGrauNorm)) {
          itemGrau = '';
        }

        if (itemName.length > 3 && !/^\d/.test(itemName) && !itemName.toLowerCase().includes('itens afetivos') && !itemName.toLowerCase().includes('cognitivos')) {
          const grauTextToRemove = matchT2[2].toUpperCase();
          itemName = itemName.replace(new RegExp(`\\b${grauTextToRemove}\\b`, 'i'), '').trim();

          if(itemName){
            extractedItemsMap.set(itemName, {
              id: crypto.randomUUID(),
              numero: '',
              nome: itemName,
              fase: '--',
              grau: itemGrau, 
              comentario: ''
            });
          }
        }
      }
    }

    let finalItems = Array.from(extractedItemsMap.values());

    // 5. COMENTÁRIOS (O LEITOR UNIVERSAL BLINDADO)
    if (idxComentarios !== -1) {
      let cleanComments = globalCleanText.substring(idxComentarios + 12).replace(/[\n\r]/g, ' ').replace(/\s{2,}/g, ' ');

      const allCommentMatches: any[] = [];

      // NOVO REGEX UNIFICADO: Pega TUDO o que estiver dentro dos parênteses (notas, fases, siglas) sem engasgar
      const unifiedCommentRegex = /(?:^|\W)(\d{1,2})\s*[-–—]\s*(.+?)\s*\(\s*([^)]+)\s*\)\s*:?/gi;
      let matchC;
      
      while ((matchC = unifiedCommentRegex.exec(cleanComments)) !== null) {
        const num = matchC[1].trim();
        let rawFaseC = '--';
        let rawGrauC = '';
        const insideParens = matchC[3].toUpperCase();

        // Interpretador inteligente do conteúdo dos parênteses
        if (insideParens.includes('/')) {
           const parts = insideParens.split('/');
           rawFaseC = parts[0].trim() || '--';
           rawGrauC = parts[1].trim();
        } else {
           const val = insideParens.trim();
           if (val === 'PR') {
              rawFaseC = 'PR';
           } else if (/^[1-6]$/.test(val) || ['NORMAL', 'DESTACOU-SE', 'NÃO OBSERVADO', 'PRECISA MELHORAR', 'DEFICIENTE', 'ABAIXO DO PADRÃO', 'PERIGOSO', 'N/O', 'N/A', 'NR'].includes(val)) {
              rawGrauC = val;
           } else {
              rawFaseC = val;
           }
        }

        if (rawFaseC === 'PR') rawGrauC = '';
        const rawGrauCNorm = rawGrauC.replace(/\s+/g, '');
        if (['--', 'N/O', 'N/A', 'NR', 'AN/', 'NÃOOBSERVADO', ''].includes(rawGrauCNorm)) {
          rawGrauC = '';
        }

        allCommentMatches.push({
          index: matchC.index,
          length: matchC[0].length,
          numero: num,
          fase: rawFaseC,
          grau: rawGrauC
        });
      }

      allCommentMatches.sort((a, b) => a.index - b.index);

      for (let i = 0; i < allCommentMatches.length; i++) {
        const currentMatch = allCommentMatches[i];
        const startIndex = currentMatch.index + currentMatch.length;
        let endIndex;

        if (i + 1 < allCommentMatches.length) {
          endIndex = allCommentMatches[i + 1].index;
        } else {
          const assDigitalIndex = cleanComments.indexOf('Ass. Digital', startIndex);
          const recomendacoesIndex = cleanComments.indexOf('Recomendações/Parecer:', startIndex);
          
          if (assDigitalIndex !== -1 && recomendacoesIndex !== -1) endIndex = Math.min(assDigitalIndex, recomendacoesIndex);
          else if (assDigitalIndex !== -1) endIndex = assDigitalIndex;
          else if (recomendacoesIndex !== -1) endIndex = recomendacoesIndex;
          else endIndex = cleanComments.length;
        }

        let comentarioText = cleanComments.substring(startIndex, endIndex).trim();

        let foundItem = finalItems.find((item: any) => item.numero === currentMatch.numero);
        if (foundItem) {
          foundItem.comentario = comentarioText;
        }
      }
    }

    finalItems.sort((a: any, b: any) => {
      const numA = parseInt(a.numero);
      const numB = parseInt(b.numero);
      if (isNaN(numA) && isNaN(numB)) return 0;
      if (isNaN(numA)) return 1; 
      if (isNaN(numB)) return -1;
      return numA - numB;
    });

    if (finalItems.length > 0) {
      setItems(finalItems);
      setStatus('reviewing');
      setErrorMsg('');
    } else {
      setErrorMsg('Não foi possível extrair os itens. Verifique o arquivo enviado.');
      setStatus('idle');
    }
  };

  const handleFileUpload = async (e: any) => {
    const file = e.target.files[0];
    if (!file) return;

    e.target.value = ''; 
    setStatus('loading');
    setErrorMsg('');

    if (file.type === 'application/pdf') {
      if (!window.pdfjsLib) {
        setErrorMsg('Biblioteca de leitura ainda está carregando. Tente novamente em 2 segundos.');
        setStatus('idle');
        return;
      }

      try {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await window.pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        let fullText = '';
        
        for (let i = 1; i <= pdf.numPages; i++) {
          const page = await pdf.getPage(i);
          const textContent = await page.getTextContent();
          
          const items = textContent.items;
          items.sort((a: any, b: any) => {
            if (Math.abs(a.transform[5] - b.transform[5]) > 5) {
              return b.transform[5] - a.transform[5];
            }
            return a.transform[4] - b.transform[4];
          });

          let lastY = null;
          let pageText = '';
          for (let item of items) {
            if (lastY !== null && Math.abs(lastY - item.transform[5]) > 5) {
              pageText += '\n';
            } else if (lastY !== null) {
              pageText += ' '; 
            }
            pageText += item.str.trim();
            lastY = item.transform[5];
          }
          pageText = pageText.replace(/ {2,}/g, ' ');
          fullText += pageText + '\n\n';
        }
        processTextData(fullText);
      } catch (err) {
        console.error(err);
        setErrorMsg('Erro ao ler o arquivo PDF. Verifique se o arquivo não está corrompido.');
        setStatus('idle');
      }
    } else {
      setErrorMsg('Por favor, selecione um arquivo PDF válido.');
      setStatus('idle');
    }
  };

  const updateItem = (id: string, field: string, value: string) => {
    setItems(items.map((item: any) => item.id === id ? { ...item, [field]: value } : item));
  };

  const removeItem = (id: string) => {
    setItems(items.filter((item: any) => item.id !== id));
  };

  const addNewItem = () => {
    setItems([...items, {
      id: crypto.randomUUID(),
      numero: '',
      nome: 'Novo Item',
      fase: '',
      grau: '',
      comentario: ''
    }]);
  };

  const buildPayloadData = () => {
    return items.map(item => ({
      data: meta.data,
      esquadrilha: meta.esquadrilha,
      missao: meta.missao,
      grauMissao: (meta.tipoMissao === 'Abortiva' || meta.tipoMissao === 'Extra') ? '' : meta.grauMissao,
      aluno1p: meta.aluno1p,
      instrutor: meta.instrutor,
      faseMissao: meta.fase,
      aeronave: meta.aeronave,
      hdep: meta.hdep,
      pousos: meta.pousos,
      tev: meta.tev,
      parecer: meta.parecer,
      numero: item.numero,
      nome: item.nome,
      faseItem: item.fase,
      grau: item.grau,
      comentario: item.comentario,
      tipoMissao: meta.tipoMissao 
    }));
  };

  const exportToCSV = () => {
    const headers = [
      'Data da Missão', 'Esquadrilha', 'Missão', 'Grau da Missão', '1P / AL', 'IN', 'Fase da Missão', 'Aeronave', 'H. Dep', 'Pousos', 'TEV',
      'Parecer/Recomendações', 'Nº do Item', 'Nome da Manobra/Item', 'Fase do Item', 'Grau/Menção', 'Comentário', 'Tipo de Missão'
    ];
    
    const payload = buildPayloadData();

    const csvContent = [
      headers.join(','),
      ...payload.map(row => {
        const cleanComentario = row.comentario ? row.comentario.replace(/"/g, '""') : ''; 
        const cleanParecer = row.parecer ? row.parecer.replace(/"/g, '""') : ''; 
        return `"${row.data}","${row.esquadrilha}","${row.missao}","${row.grauMissao}","${row.aluno1p}","${row.instrutor}","${row.faseMissao}","${row.aeronave}","${row.hdep}","${row.pousos}","${row.tev}","${cleanParecer}","${row.numero}","${row.nome}","${row.faseItem}","${row.grau}","${cleanComentario}","${row.tipoMissao}"`;
      })
    ].join('\n');

    const blob = new Blob(['\ufeff' + csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', `Ficha_${meta.missao || 'Extraida'}_${new Date().toISOString().slice(0,10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const sendToGoogleSheets = async () => {
    if (!meta.esquadrilha) {
      setErrorMsg('Atenção: Por favor, selecione a Esquadrilha antes de enviar para a base de dados.');
      window.scrollTo(0, 0);
      return;
    }
    if (!meta.aluno1p || !meta.instrutor) {
       setErrorMsg('Atenção: Preencha o trigrama do 1P/AL e do IN antes de enviar.');
       window.scrollTo(0, 0);
       return;
    }

    setModalState('sending');
    setShowModal(true);

    try {
      const payload = buildPayloadData();
      
      const response = await fetch(WEBHOOK_URL, {
        method: 'POST',
        headers: {
          'Content-Type': 'text/plain;charset=utf-8',
        },
        body: JSON.stringify(payload)
      });

      if (response.ok) {
        setModalState('success');
        setModalMessage('Ficha salva com sucesso no Banco de Instrução!');
      } else {
        throw new Error('Falha na comunicação com a planilha.');
      }
    } catch (error) {
      console.error(error);
      setModalState('error');
      setModalMessage('Ocorreu um erro de rede. Verifique a sua conexão de internet e tente novamente.');
    }
  };

  const closeModalAndReset = () => {
    setShowModal(false);
    if (modalState === 'success') {
      setStatus('idle');
      setItems([]);
    }
  };

  const isGrauDisabled = meta.tipoMissao === 'Abortiva' || meta.tipoMissao === 'Extra';

  return (
    <div className="min-h-screen w-full bg-slate-50 text-slate-800 font-sans p-4 md:p-8 relative">
      
      {showModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-sm px-4 transition-all">
          <div className="bg-white rounded-2xl shadow-2xl p-8 max-w-sm w-full text-center flex flex-col items-center transform scale-100 animate-in fade-in zoom-in duration-200">
            
            {modalState === 'sending' && (
              <>
                <div className="w-20 h-20 bg-blue-50 text-blue-600 rounded-full flex items-center justify-center mb-6">
                  <RefreshCw size={40} className="animate-spin" />
                </div>
                <h2 className="text-2xl font-bold text-slate-800 mb-2">Enviando Dados...</h2>
                <p className="text-slate-500 mb-4">Salvando todos os itens da ficha de voo no banco de instrução.</p>
                <div className="bg-amber-50 text-amber-700 border border-amber-200 text-sm font-medium px-4 py-3 rounded-lg flex items-center justify-center gap-2 w-full">
                  <AlertCircle size={18} />
                  Por favor, não feche esta página.
                </div>
              </>
            )}

            {modalState === 'success' && (
              <>
                <div className="w-20 h-20 bg-green-50 text-green-500 rounded-full flex items-center justify-center mb-6">
                  <CheckCircle2 size={48} />
                </div>
                <h2 className="text-2xl font-bold text-slate-800 mb-2">Concluído!</h2>
                <p className="text-slate-500 mb-8">{modalMessage}</p>
                <button 
                  onClick={closeModalAndReset}
                  className="w-full bg-slate-800 hover:bg-slate-900 text-white font-bold py-3 px-6 rounded-xl transition shadow-lg"
                >
                  Inserir Próxima Ficha
                </button>
              </>
            )}

            {modalState === 'error' && (
              <>
                <div className="w-20 h-20 bg-red-50 text-red-500 rounded-full flex items-center justify-center mb-6">
                  <XCircle size={48} />
                </div>
                <h2 className="text-2xl font-bold text-slate-800 mb-2">Erro no Envio</h2>
                <p className="text-slate-500 mb-8">{modalMessage}</p>
                <button 
                  onClick={() => setShowModal(false)}
                  className="w-full bg-slate-200 hover:bg-slate-300 text-slate-800 font-bold py-3 px-6 rounded-xl transition"
                >
                  Fechar e Tentar Novamente
                </button>
              </>
            )}

          </div>
        </div>
      )}

      <div className={`max-w-6xl mx-auto transition-opacity ${showModal ? 'opacity-30 pointer-events-none' : 'opacity-100'}`}>
        
        <header className="mb-8 bg-white p-6 rounded-xl shadow-sm border border-slate-200">
          <div className="flex items-center gap-4">
            <div className="w-12 h-12 bg-blue-600 rounded-lg flex items-center justify-center text-white shrink-0">
              <FileText size={28} />
            </div>
            <div>
              <h1 className="text-2xl font-bold text-slate-900">Extrator de Fichas de Voo</h1>
              <p className="text-slate-500">Sistema automatizado para instrução aérea.</p>
            </div>
          </div>
        </header>

        {errorMsg && (
          <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-lg flex items-start gap-3 text-red-700">
            <AlertCircle className="shrink-0 mt-0.5" size={20} />
            <p>{errorMsg}</p>
          </div>
        )}

        {status === 'idle' && (
          <div className="flex justify-center">
            <div className="bg-white p-10 md:p-16 rounded-xl shadow-sm border border-slate-200 flex flex-col items-center justify-center text-center w-full max-w-2xl">
              <div className="w-20 h-20 bg-blue-50 text-blue-600 rounded-full flex items-center justify-center mb-6">
                <Upload size={40} />
              </div>
              <h2 className="text-2xl font-semibold mb-3">Inserir Arquivo PDF</h2>
              <p className="text-slate-500 text-base mb-8 max-w-md">
                Selecione o arquivo da ficha de voo gerada pelo sistema. A leitura será feita de forma automática e à prova de falhas.
              </p>
              <label className="bg-blue-600 hover:bg-blue-700 transition text-white px-8 py-4 rounded-xl font-medium cursor-pointer shadow-sm text-lg flex items-center gap-3">
                <FileText size={24} />
                Selecionar PDF
                <input type="file" accept="application/pdf" className="hidden" onClick={(e: any) => { e.target.value = '' }} onChange={handleFileUpload} />
              </label>
            </div>
          </div>
        )}

        {status === 'loading' && (
          <div className="bg-white p-16 rounded-xl shadow-sm border border-slate-200 flex flex-col items-center justify-center text-center">
            <RefreshCw className="text-blue-600 animate-spin mb-4" size={40} />
            <h2 className="text-xl font-semibold">Analisando a Ficha...</h2>
            <p className="text-slate-500">Extraindo cabeçalho, tabelas e comentários.</p>
          </div>
        )}

        {status === 'reviewing' && (
          <div className="space-y-6">
            <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
              <div className="flex items-center justify-between mb-4 border-b border-slate-100 pb-3">
                <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                  <FileText className="text-blue-500" size={20} /> Dados da Missão
                </h3>
                <span className="text-xs font-semibold bg-blue-50 text-blue-600 px-3 py-1 rounded-full">
                  Esses dados acompanharão todas as linhas no banco de dados
                </span>
              </div>
              
              <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-4">
                <MetaSelect 
                  label="Esquadrilha *" 
                  value={meta.esquadrilha} 
                  onChange={(e: any) => updateMeta('esquadrilha', e.target.value)} 
                  options={['Antares', 'Vega', 'Castor', 'Sirius']} 
                />
                <MetaInput label="1P / Aluno *" value={meta.aluno1p} onChange={(e: any) => updateMeta('aluno1p', e.target.value)} maxLength={3} placeholder="MTA" />
                <MetaInput label="Instrutor (IN) *" value={meta.instrutor} onChange={(e: any) => updateMeta('instrutor', e.target.value)} maxLength={3} placeholder="HNI" />
                <MetaInput label="Missão" value={meta.missao} onChange={(e: any) => updateMeta('missao', e.target.value)} />
                <MetaInput 
                  label="Grau da Missão" 
                  value={isGrauDisabled ? '' : meta.grauMissao} 
                  onChange={(e: any) => updateMeta('grauMissao', e.target.value)} 
                  disabled={isGrauDisabled}
                  title={isGrauDisabled ? 'Missões Abortivas ou Extras não possuem grau.' : ''}
                />
                <MetaSelect 
                  label="Tipo de Missão" 
                  value={meta.tipoMissao} 
                  onChange={(e: any) => updateMeta('tipoMissao', e.target.value)} 
                  options={['Normal', 'Abortiva', 'Extra', 'Revisão']} 
                />
                <MetaInput label="Data" value={meta.data} onChange={(e: any) => updateMeta('data', e.target.value)} />
                <MetaInput label="Fase" value={meta.fase} onChange={(e: any) => updateMeta('fase', e.target.value)} />
                <MetaInput label="Aeronave" value={meta.aeronave} onChange={(e: any) => updateMeta('aeronave', e.target.value)} />
                <MetaInput label="H. Dep" value={meta.hdep} onChange={(e: any) => updateMeta('hdep', e.target.value)} />
                <MetaInput label="Pousos" value={meta.pousos} onChange={(e: any) => updateMeta('pousos', e.target.value)} />
                <MetaInput label="TEV" value={meta.tev} onChange={(e: any) => updateMeta('tev', e.target.value)} />
              </div>

              <div className="mt-4 pt-4 border-t border-slate-100">
                <MetaTextarea 
                  label="Parecer do Comandante / Recomendações" 
                  value={meta.parecer} 
                  onChange={(e: any) => updateMeta('parecer', e.target.value)} 
                  placeholder="Se houver recomendações no fim da ficha, aparecerão aqui..." 
                />
              </div>
            </div>

            <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
              <div className="p-6 border-b border-slate-200 flex flex-col lg:flex-row lg:items-center justify-between gap-4 bg-slate-50/50">
                <div>
                  <h2 className="text-xl font-semibold text-slate-900 flex items-center gap-2">
                    <Check className="text-green-500" size={24} /> 
                    {items.length} Itens Encontrados
                  </h2>
                  <p className="text-slate-500 text-sm mt-1">Revise as avaliações antes de enviar para a base.</p>
                </div>
                <div className="flex flex-wrap items-center gap-3">
                  <button 
                    onClick={() => { setStatus('idle'); setErrorMsg(''); setItems([]); }}
                    className="px-4 py-3 text-sm font-medium text-slate-600 bg-white border border-slate-300 hover:bg-slate-50 rounded-xl transition"
                  >
                    Voltar / Cancelar
                  </button>
                  <button 
                    onClick={exportToCSV}
                    className="px-4 py-3 text-sm font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-xl transition flex items-center gap-2"
                  >
                    <Download size={18} /> Exportar (CSV)
                  </button>
                  <button 
                    onClick={sendToGoogleSheets}
                    className="px-8 py-3 text-base font-bold text-white bg-blue-600 hover:bg-blue-700 rounded-xl transition shadow-lg flex items-center gap-2"
                  >
                    <Zap size={20} /> Enviar para o Banco
                  </button>
                </div>
              </div>

              <div className="p-6 overflow-x-auto">
                <table className="w-full text-left border-collapse min-w-[800px]">
                  <thead>
                    <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 text-sm">
                      <th className="p-3 font-medium w-16">Nº</th>
                      <th className="p-3 font-medium w-48">Nome da Manobra/Item</th>
                      <th className="p-3 font-medium w-20">Fase</th>
                      <th className="p-3 font-medium w-24">Grau</th>
                      <th className="p-3 font-medium">Comentário</th>
                      <th className="p-3 font-medium w-12 text-center">Ações</th>
                    </tr>
                  </thead>
                  <tbody className="align-top text-sm">
                    {items.map((item) => {
                      const gradeColorClass = getGradeColorClass(item.grau);
                      return (
                        <tr key={item.id} className="border-b border-slate-100 hover:bg-slate-50/50 transition">
                          <td className="p-3">
                            <input 
                              type="text" 
                              value={item.numero} 
                              onChange={(e) => updateItem(item.id, 'numero', e.target.value)}
                              className="w-full bg-transparent border-b border-transparent hover:border-slate-300 focus:border-blue-500 focus:bg-white outline-none px-1 py-1 text-slate-600"
                            />
                          </td>
                          <td className="p-3">
                            <input 
                              type="text" 
                              value={item.nome} 
                              onChange={(e) => updateItem(item.id, 'nome', e.target.value)}
                              className={`w-full bg-transparent border-b border-transparent hover:border-slate-300 focus:border-blue-500 focus:bg-white outline-none px-1 py-1 ${gradeColorClass}`}
                            />
                          </td>
                          <td className="p-3">
                            <input 
                              type="text" 
                              value={item.fase} 
                              onChange={(e) => updateItem(item.id, 'fase', e.target.value)}
                              className="w-full bg-transparent border-b border-transparent hover:border-slate-300 focus:border-blue-500 focus:bg-white outline-none px-1 py-1 text-center text-slate-600"
                            />
                          </td>
                          <td className="p-3">
                            <input 
                              type="text" 
                              value={item.grau} 
                              disabled={item.fase === 'PR'}
                              title={item.fase === 'PR' ? "Itens PR não recebem nota." : ""}
                              onChange={(e) => updateItem(item.id, 'grau', e.target.value)}
                              className={`w-full bg-transparent border-b border-transparent hover:border-slate-300 focus:border-blue-500 focus:bg-white outline-none px-1 py-1 uppercase ${gradeColorClass} ${item.fase === 'PR' ? 'cursor-not-allowed opacity-50 bg-slate-100/50' : ''}`}
                            />
                          </td>
                          <td className="p-3">
                            <textarea 
                              value={item.comentario} 
                              onChange={(e) => updateItem(item.id, 'comentario', e.target.value)}
                              placeholder="Sem comentários para este item"
                              className={`w-full bg-transparent border border-transparent hover:border-slate-300 focus:border-blue-500 focus:bg-white focus:shadow-sm outline-none px-2 py-1 rounded resize-y min-h-[40px] leading-relaxed ${!item.comentario ? 'text-slate-400 italic' : 'text-slate-700'}`}
                            />
                          </td>
                          <td className="p-3 text-center">
                            <button 
                              onClick={() => removeItem(item.id)}
                              className="p-2 text-slate-400 hover:text-red-500 hover:bg-red-50 rounded transition"
                              title="Remover Item"
                            >
                              <Trash2 size={16} />
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                
                <div className="mt-4 flex justify-center border-t border-slate-100 pt-4">
                  <button 
                    onClick={addNewItem}
                    className="flex items-center gap-2 text-sm text-blue-600 hover:text-blue-800 font-medium px-4 py-2 hover:bg-blue-50 rounded-lg transition"
                  >
                    <Plus size={16} /> Adicionar Item Manualmente
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

      </div>
    </div>
  );
}
