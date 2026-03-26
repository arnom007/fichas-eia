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

// Função auxiliar para normalizar strings e facilitar o Match
const normalizeString = (str: string) => {
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
    // 1. EXTRAÇÃO DO CABEÇALHO (Os primeiros caracteres até começar a tabela ou itens afetivos)
    const headerEndIdx = text.search(/\b1\s*-|Itens Afetivos|Comentários:/i);
    const headerText = headerEndIdx !== -1 ? text.substring(0, headerEndIdx) : text;
    const cleanHeader = headerText.replace(/[\n\r]/g, ' ').replace(/\s{2,}/g, ' ').replace(/\|\|\|/g, '');

    const matchGrauMissao = cleanHeader.match(/GRAU\s*(\d{1,2})/i);
    const grauMissao = matchGrauMissao ? matchGrauMissao[1] : '';

    let tipoMissaoDetectado = 'Normal';
    if (cleanHeader.match(/\bExtra\b/i)) tipoMissaoDetectado = 'Extra';
    else if (cleanHeader.match(/\bRevis[ãa]o\b/i)) tipoMissaoDetectado = 'Revisão';
    else if (cleanHeader.match(/\bAbortiva\b/i) || cleanHeader.match(/\bVMET\b/i) || cleanHeader.match(/\bVMAT\b/i)) tipoMissaoDetectado = 'Abortiva';

    const matchMissao = cleanHeader.match(/\b((?:VMAT\s+|VMET\s+)?[A-Z]{2,4}-[A-Z0-9]{1,3})\b/i);
    const missao = matchMissao ? matchMissao[1].toUpperCase() : '';

    const matchData = cleanHeader.match(/\b(\d{2}\/\d{2}\/\d{4})\b/);
    const data = matchData ? matchData[1] : '';

    const times = Array.from(cleanHeader.matchAll(/\b(\d{2}:\d{2})\b/g));
    let hdep = '', tev = '';
    if (times.length >= 2) { hdep = times[0][1]; tev = times[times.length - 1][1]; } 
    else if (times.length === 1) { if (/TEV[^\d]*\d{2}:\d{2}/i.test(cleanHeader)) tev = times[0][1]; else hdep = times[0][1]; }

    let fase = '';
    const matchFase = cleanHeader.match(/FASE:\s*(.*?)(?=\s*ALUNO:|\s*INSTRUTOR:|\s*AERONAVE:|\s*NORMAL\b|\s*GRAU\b|$)/i);
    if (matchFase) fase = matchFase[1].replace(/["\n\r]/g, '').replace(/^[-:]+|[-:]+$/g, '').trim().toUpperCase();

    const matchAeronave = cleanHeader.match(/(?:AERONAVE)[\s:]*(\d{4})\b/i) || cleanHeader.match(/\b(13\d{2}|14\d{2})\b/);
    const aeronave = matchAeronave ? matchAeronave[1] : '';

    const matchPousos = cleanHeader.match(/POUSOS[\s:]*(\d{1,2})\b/i);
    const pousos = matchPousos ? matchPousos[1] : '';

    const parecerMatch = text.match(/Recomendações\/Parecer:\s*([\s\S]*?)(?=\bCiente\b|INSTRUTOR|Autoridade|Ass\. Digital|$)/i);
    const parecerStr = parecerMatch ? parecerMatch[1].replace(/\|\|\|/g, '').replace(/[\n\r]/g, ' ').replace(/\s{2,}/g, ' ').trim() : '';

    setMeta(prev => ({
      esquadrilha: prev.esquadrilha, aluno1p: prev.aluno1p, instrutor: prev.instrutor,
      fase, aeronave, data, missao, grauMissao: (tipoMissaoDetectado === 'Abortiva' || tipoMissaoDetectado === 'Extra') ? '' : grauMissao,
      tipoMissao: tipoMissaoDetectado, pousos, hdep, tev, parecer: parecerStr
    }));

    // 2. PARSING ESTRUTURADO LINHA A LINHA (COM O SEPARADOR |||)
    const extractedItemsArray: any[] = [];
    const lines = text.split('\n');

    let isAffectiveArea = false;
    let isCommentsArea = false;
    let accumulatedComments = '';

    for (let i = 0; i < lines.length; i++) {
      let line = lines[i].trim();
      if (!line) continue;

      // Remoção de lixo de rodapé no meio da linha
      line = line.replace(/MATERIAL DE ACESSO RESTRITO/gi, '').replace(/Art\. 44.*?2012/gi, '').replace(/--- PAGE \d+ ---/gi, '');

      if (line.match(/Itens Afetivos/i)) { isAffectiveArea = true; continue; }
      if (line.match(/Comentários:/i)) { isCommentsArea = true; isAffectiveArea = false; continue; }

      // Se entrou nos comentários, apenas guarda o bloco de texto para a Etapa 3
      if (isCommentsArea) {
        accumulatedComments += line.replace(/\|\|\|/g, ' ') + ' \n ';
        continue;
      }

      // LÓGICA DE ITENS AFETIVOS
      if (isAffectiveArea) {
        const parts = line.split('|||').map((p: string) => p.trim()).filter((p: string) => p);
        if (parts.length >= 2) {
          const name = parts[0].trim();
          const grade = parts[parts.length - 1].toUpperCase().replace(/\s+/g, '');
          
          if (/^(NORMAL|DESTACOU-SE|PRECISAMELHORAR|DEFICIENTE|ABAIXODOPADRÃO|PERIGOSO)$/.test(grade)) {
             extractedItemsArray.push({
               id: crypto.randomUUID(),
               numero: '', // Itens afetivos na tabela não tem número
               nome: name,
               fase: '--',
               grau: ['--', 'N/O', 'N/A', 'NR'].includes(grade) ? '' : grade,
               comentario: ''
             });
          }
        }
      } 
      // LÓGICA DE MANOBRAS E ITENS DA TABELA PRINCIPAL
      else {
        const parts = line.split('|||').map((p: string) => p.trim()).filter((p: string) => p);
        
        // Se a linha foi perfeitamente separada pelo cálculo espacial (Nome ||| RM 5)
        if (parts.length >= 2) {
          const maneuverPart = parts[0].trim();
          const gradePart = parts[parts.length - 1].trim().toUpperCase();

          const numMatch = maneuverPart.match(/^(0?[1-9]|[1-5][0-9])\s*[-–—]?\s*(.*)$/);
          if (numMatch) {
            const num = numMatch[1];
            const name = numMatch[2].replace(/^[-–—.:\s]+|[-–—.:\s]+$/g, '').trim();

            let faseItem = '--';
            let grauItem = '';

            // O último bloco extraído pelo ||| contém a Fase e/ou o Grau
            const pgParts = gradePart.split(/\s+/);
            
            // Lógica inteligente para Fase e Grau (Atende a Ficha VI-01)
            if (pgParts.length === 2) {
                faseItem = pgParts[0].toUpperCase();
                grauItem = pgParts[1].toUpperCase();
            } else if (pgParts.length === 1) {
                const singlePart = pgParts[0].toUpperCase();
                if (singlePart === 'PR') {
                    faseItem = 'PR';
                } else if (/^(RC|RM|RO|--)$/.test(singlePart)) {
                    faseItem = singlePart;
                } else {
                    // Se for apenas número (como na VI-01: "5")
                    grauItem = singlePart;
                }
            }

            if (['--', 'N/O', 'N/A', 'NR', 'AN/', 'NÃOOBSERVADO'].includes(grauItem.replace(/\s+/g, ''))) grauItem = '';

            extractedItemsArray.push({
              id: crypto.randomUUID(),
              numero: num,
              nome: name,
              fase: faseItem,
              grau: grauItem,
              comentario: ''
            });
          }
        }
      }
    }

    // =====================================================================
    // ETAPA 3: LER COMENTÁRIOS E FAZER O MATCH
    // =====================================================================
    if (accumulatedComments) {
      const cleanComments = accumulatedComments.replace(/[\n\r]/g, ' ').replace(/\s{2,}/g, ' ');

      const allCommentMatches: any[] = [];
      // Regex busca o Título do comentário no formato "Nº - Nome (...):"
      const unifiedCommentRegex = /(?:^|\s)(0?[1-9]|[1-5][0-9])\s*[-–—]\s*([A-Za-zÀ-ÿ0-9\s/]+?)(?:\s*\(\s*([^)]+)\s*\)\s*:?|\s*:)/gi;
      let matchC;
      
      while ((matchC = unifiedCommentRegex.exec(cleanComments)) !== null) {
        allCommentMatches.push({
          index: matchC.index,
          length: matchC[0].length,
          numero: matchC[1].trim(),
          nome: matchC[2].trim()
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

        const comentarioText = cleanComments.substring(startIndex, endIndex).trim();

        // 🟢 O GRANDE MATCH (Busca no array populado nas etapas 1 e 2)
        // Tentativa 1: Tenta achar exatamente pelo Número
        let targetItem = extractedItemsArray.find(item => item.numero === currentMatch.numero);

        // Tentativa 2: Se não achou pelo número, tenta achar pelo Nome normalizado
        if (!targetItem) {
          targetItem = extractedItemsArray.find(item => {
             const n1 = normalizeString(item.nome);
             const n2 = normalizeString(currentMatch.nome);
             return n1.includes(n2) || n2.includes(n1);
          });
        }

        // Se encontrou o item, insere o comentário nele
        if (targetItem) {
          targetItem.comentario = comentarioText;
          // Atualiza o número caso o item afetivo estivesse sem número na tabela
          if (!targetItem.numero) targetItem.numero = currentMatch.numero;
        } else {
          // Fallback de segurança: se o item foi comentado mas o PDF comeu ele da tabela
          extractedItemsArray.push({
            id: crypto.randomUUID(),
            numero: currentMatch.numero,
            nome: currentMatch.nome,
            fase: '--',
            grau: '',
            comentario: comentarioText
          });
        }
      }
    }

    // Ordenação final
    extractedItemsArray.sort((a: any, b: any) => {
      const numA = parseInt(a.numero);
      const numB = parseInt(b.numero);
      if (isNaN(numA) && isNaN(numB)) return 0;
      if (isNaN(numA)) return 1; 
      if (isNaN(numB)) return -1;
      return numA - numB;
    });

    if (extractedItemsArray.length > 0) {
      setItems(extractedItemsArray);
      setStatus('reviewing');
      setErrorMsg('');
    } else {
      setErrorMsg('Não foi possível extrair os itens. O PDF pode não estar no padrão esperado.');
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
        setErrorMsg('Carregando biblioteca... Tente novamente em 2 segundos.');
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
          
          const items = textContent.items as any[];
          
          // Ordenação primária por Eixo Y
          items.sort((a, b) => {
            if (Math.abs(a.transform[5] - b.transform[5]) > 5) {
              return b.transform[5] - a.transform[5];
            }
            return a.transform[4] - b.transform[4];
          });

          // Agrupamento de Linhas
          const lines: any[][] = [];
          let currentLine: any[] = [];
          let currentY = items.length > 0 ? items[0].transform[5] : 0;

          for (const item of items) {
            if (Math.abs(item.transform[5] - currentY) > 5) {
              lines.push(currentLine);
              currentLine = [item];
              currentY = item.transform[5];
            } else {
              currentLine.push(item);
            }
          }
          if (currentLine.length > 0) lines.push(currentLine);

          // O MOTOR DE SEPARAÇÃO ESPACIAL DE COLUNAS (Substitui o espaço grande por '|||')
          for (const line of lines) {
            line.sort((a, b) => a.transform[4] - b.transform[4]);
            let lineStr = '';
            for (let j = 0; j < line.length; j++) {
              lineStr += line[j].str;
              if (j < line.length - 1) {
                const rightEdge = line[j].transform[4] + (line[j].width || 0);
                const nextLeftEdge = line[j + 1].transform[4];
                const gap = nextLeftEdge - rightEdge;
                
                // SE O BURACO FÍSICO FOR GRANDE (COLUNA), INSERE O SEPARADOR DE PROTEÇÃO '|||'
                if (gap > 35) {
                  lineStr += ' ||| ';
                } else {
                  if (!lineStr.endsWith(' ')) lineStr += ' ';
                }
              }
            }
            fullText += lineStr.trim() + '\n';
          }
          fullText += '\n\n';
        }
        processTextData(fullText);
      } catch (err) {
        console.error(err);
        setErrorMsg('Erro ao ler o arquivo PDF.');
        setStatus('idle');
      }
    } else {
      setErrorMsg('Por favor, selecione um arquivo PDF válido.');
      setStatus('idle');
    }
  };

  const updateItem = (id: string, field: string, value: string) => setItems(items.map((item: any) => item.id === id ? { ...item, [field]: value } : item));
  const removeItem = (id: string) => setItems(items.filter((item: any) => item.id !== id));
  const addNewItem = () => setItems([...items, { id: crypto.randomUUID(), numero: '', nome: 'Novo Item', fase: '', grau: '', comentario: '' }]);

  const buildPayloadData = () => {
    return items.map(item => ({
      data: meta.data, esquadrilha: meta.esquadrilha, missao: meta.missao,
      grauMissao: (meta.tipoMissao === 'Abortiva' || meta.tipoMissao === 'Extra') ? '' : meta.grauMissao,
      aluno1p: meta.aluno1p, instrutor: meta.instrutor, faseMissao: meta.fase, aeronave: meta.aeronave,
      hdep: meta.hdep, pousos: meta.pousos, tev: meta.tev, parecer: meta.parecer,
      numero: item.numero, nome: item.nome, faseItem: item.fase, grau: item.grau,
      comentario: item.comentario, tipoMissao: meta.tipoMissao
    }));
  };

  const exportToCSV = () => {
    const headers = ['Data da Missão', 'Esquadrilha', 'Missão', 'Grau da Missão', '1P / AL', 'IN', 'Fase da Missão', 'Aeronave', 'H. Dep', 'Pousos', 'TEV', 'Parecer/Recomendações', 'Nº do Item', 'Nome da Manobra/Item', 'Fase do Item', 'Grau/Menção', 'Comentário', 'Tipo de Missão'];
    const payload = buildPayloadData();
    const csvContent = [
      headers.join(','),
      ...payload.map(row => `"${row.data}","${row.esquadrilha}","${row.missao}","${row.grauMissao}","${row.aluno1p}","${row.instrutor}","${row.faseMissao}","${row.aeronave}","${row.hdep}","${row.pousos}","${row.tev}","${(row.parecer || '').replace(/"/g, '""')}","${row.numero}","${row.nome}","${row.faseItem}","${row.grau}","${(row.comentario || '').replace(/"/g, '""')}","${row.tipoMissao}"`)
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
    if (!meta.esquadrilha) { setErrorMsg('Por favor, selecione a Esquadrilha.'); window.scrollTo(0, 0); return; }
    if (!meta.aluno1p || !meta.instrutor) { setErrorMsg('Preencha o trigrama do 1P/AL e do IN.'); window.scrollTo(0, 0); return; }

    setModalState('sending'); setShowModal(true);
    try {
      const response = await fetch(WEBHOOK_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(buildPayloadData()) });
      if (response.ok) { setModalState('success'); setModalMessage('Salvo com sucesso!'); } 
      else throw new Error('Falha.');
    } catch (error) {
      setModalState('error'); setModalMessage('Erro de conexão. Tente novamente.');
    }
  };

  const closeModalAndReset = () => { setShowModal(false); if (modalState === 'success') { setStatus('idle'); setItems([]); } };
  const isGrauDisabled = meta.tipoMissao === 'Abortiva' || meta.tipoMissao === 'Extra';

  return (
    <div className="min-h-screen w-full bg-slate-50 text-slate-800 font-sans p-4 md:p-8 relative">
      {showModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-sm px-4">
          <div className="bg-white rounded-2xl shadow-2xl p-8 max-w-sm w-full text-center flex flex-col items-center">
            {modalState === 'sending' && (<><RefreshCw size={40} className="animate-spin text-blue-600 mb-6" /><h2 className="text-2xl font-bold mb-2">Enviando...</h2><p className="text-slate-500 mb-4">Salvando dados no banco.</p></>)}
            {modalState === 'success' && (<><CheckCircle2 size={48} className="text-green-500 mb-6" /><h2 className="text-2xl font-bold mb-2">Concluído!</h2><p className="text-slate-500 mb-8">{modalMessage}</p><button onClick={closeModalAndReset} className="w-full bg-slate-800 text-white font-bold py-3 px-6 rounded-xl">Próxima Ficha</button></>)}
            {modalState === 'error' && (<><XCircle size={48} className="text-red-500 mb-6" /><h2 className="text-2xl font-bold mb-2">Erro</h2><p className="text-slate-500 mb-8">{modalMessage}</p><button onClick={() => setShowModal(false)} className="w-full bg-slate-200 text-slate-800 font-bold py-3 px-6 rounded-xl">Fechar</button></>)}
          </div>
        </div>
      )}

      <div className={`max-w-6xl mx-auto transition-opacity ${showModal ? 'opacity-30 pointer-events-none' : 'opacity-100'}`}>
        <header className="mb-8 bg-white p-6 rounded-xl shadow-sm border border-slate-200">
          <div className="flex items-center gap-4">
            <div className="w-12 h-12 bg-blue-600 rounded-lg flex items-center justify-center text-white shrink-0"><FileText size={28} /></div>
            <div><h1 className="text-2xl font-bold">Extrator de Fichas de Voo</h1><p className="text-slate-500">Sistema automatizado para instrução aérea.</p></div>
          </div>
        </header>

        {errorMsg && <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-lg text-red-700"><AlertCircle className="inline-block mr-2" size={20} />{errorMsg}</div>}

        {status === 'idle' && (
          <div className="flex justify-center">
            <div className="bg-white p-10 md:p-16 rounded-xl shadow-sm border border-slate-200 flex flex-col items-center justify-center text-center w-full max-w-2xl">
              <Upload size={40} className="text-blue-600 mb-6" />
              <h2 className="text-2xl font-semibold mb-3">Inserir Arquivo PDF</h2>
              <p className="text-slate-500 text-base mb-8 max-w-md">Selecione o arquivo da ficha de voo gerada pelo sistema. A leitura será feita de forma automática e à prova de falhas.</p>
              <label className="bg-blue-600 hover:bg-blue-700 text-white px-8 py-4 rounded-xl font-medium cursor-pointer flex items-center gap-3">
                <FileText size={24} /> Selecionar PDF
                <input type="file" accept="application/pdf" className="hidden" onClick={(e: any) => e.target.value = ''} onChange={handleFileUpload} />
              </label>
            </div>
          </div>
        )}

        {status === 'loading' && (
          <div className="bg-white p-16 rounded-xl flex flex-col items-center text-center"><RefreshCw className="text-blue-600 animate-spin mb-4" size={40} /><h2 className="text-xl font-semibold">Analisando...</h2></div>
        )}

        {status === 'reviewing' && (
          <div className="space-y-6">
            <div className="bg-white rounded-xl shadow-sm border p-6">
              <h3 className="text-lg font-bold mb-4 flex items-center gap-2 border-b pb-3"><FileText className="text-blue-500" size={20} /> Dados da Missão</h3>
              <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-4">
                <MetaSelect label="Esquadrilha *" value={meta.esquadrilha} onChange={(e: any) => updateMeta('esquadrilha', e.target.value)} options={['Antares', 'Vega', 'Castor', 'Sirius']} />
                <MetaInput label="1P / Aluno *" value={meta.aluno1p} onChange={(e: any) => updateMeta('aluno1p', e.target.value)} maxLength={3} />
                <MetaInput label="Instrutor (IN) *" value={meta.instrutor} onChange={(e: any) => updateMeta('instrutor', e.target.value)} maxLength={3} />
                <MetaInput label="Missão" value={meta.missao} onChange={(e: any) => updateMeta('missao', e.target.value)} />
                <MetaInput label="Grau da Missão" value={isGrauDisabled ? '' : meta.grauMissao} onChange={(e: any) => updateMeta('grauMissao', e.target.value)} disabled={isGrauDisabled} />
                <MetaSelect label="Tipo de Missão" value={meta.tipoMissao} onChange={(e: any) => updateMeta('tipoMissao', e.target.value)} options={['Normal', 'Abortiva', 'Extra', 'Revisão']} />
                <MetaInput label="Data" value={meta.data} onChange={(e: any) => updateMeta('data', e.target.value)} />
                <MetaInput label="Fase" value={meta.fase} onChange={(e: any) => updateMeta('fase', e.target.value)} />
                <MetaInput label="Aeronave" value={meta.aeronave} onChange={(e: any) => updateMeta('aeronave', e.target.value)} />
                <MetaInput label="H. Dep" value={meta.hdep} onChange={(e: any) => updateMeta('hdep', e.target.value)} />
                <MetaInput label="Pousos" value={meta.pousos} onChange={(e: any) => updateMeta('pousos', e.target.value)} />
                <MetaInput label="TEV" value={meta.tev} onChange={(e: any) => updateMeta('tev', e.target.value)} />
              </div>
              <div className="mt-4 pt-4 border-t"><MetaTextarea label="Parecer do Comandante" value={meta.parecer} onChange={(e: any) => updateMeta('parecer', e.target.value)} /></div>
            </div>

            <div className="bg-white rounded-xl shadow-sm border overflow-hidden">
              <div className="p-6 border-b flex flex-col lg:flex-row justify-between gap-4 bg-slate-50/50">
                <div><h2 className="text-xl font-semibold flex items-center gap-2"><Check className="text-green-500" size={24} /> {items.length} Itens Encontrados</h2></div>
                <div className="flex flex-wrap gap-3">
                  <button onClick={() => { setStatus('idle'); setItems([]); }} className="px-4 py-3 text-sm border hover:bg-slate-50 rounded-xl">Voltar</button>
                  <button onClick={exportToCSV} className="px-4 py-3 text-sm border hover:bg-slate-50 rounded-xl flex items-center gap-2"><Download size={18} /> Exportar CSV</button>
                  <button onClick={sendToGoogleSheets} className="px-8 py-3 text-base text-white bg-blue-600 hover:bg-blue-700 rounded-xl flex items-center gap-2"><Zap size={20} /> Enviar Banco</button>
                </div>
              </div>

              <div className="p-6 overflow-x-auto">
                <table className="w-full text-left border-collapse min-w-[800px]">
                  <thead><tr className="bg-slate-50 border-b text-sm"><th className="p-3 w-16">Nº</th><th className="p-3 w-48">Nome</th><th className="p-3 w-20">Fase</th><th className="p-3 w-24">Grau</th><th className="p-3">Comentário</th><th className="p-3 w-12 text-center">Ações</th></tr></thead>
                  <tbody className="text-sm">
                    {items.map((item) => (
                      <tr key={item.id} className="border-b hover:bg-slate-50/50">
                        <td className="p-3"><input value={item.numero} onChange={(e) => updateItem(item.id, 'numero', e.target.value)} className="w-full bg-transparent px-1 outline-none focus:bg-white focus:ring-1 focus:ring-blue-500 rounded" /></td>
                        <td className="p-3"><input value={item.nome} onChange={(e) => updateItem(item.id, 'nome', e.target.value)} className={`w-full bg-transparent px-1 outline-none focus:bg-white focus:ring-1 focus:ring-blue-500 rounded ${getGradeColorClass(item.grau)}`} /></td>
                        <td className="p-3"><input value={item.fase} onChange={(e) => updateItem(item.id, 'fase', e.target.value)} className="w-full bg-transparent px-1 text-center outline-none focus:bg-white focus:ring-1 focus:ring-blue-500 rounded" /></td>
                        <td className="p-3">
                          <input 
                            value={item.grau} 
                            disabled={item.fase === 'PR'}
                            title={item.fase === 'PR' ? "Itens PR não recebem nota." : ""}
                            onChange={(e) => updateItem(item.id, 'grau', e.target.value)} 
                            className={`w-full bg-transparent px-1 uppercase outline-none focus:bg-white focus:ring-1 focus:ring-blue-500 rounded ${getGradeColorClass(item.grau)} ${item.fase === 'PR' ? 'opacity-50 cursor-not-allowed' : ''}`} 
                          />
                        </td>
                        <td className="p-3"><textarea value={item.comentario} onChange={(e) => updateItem(item.id, 'comentario', e.target.value)} className={`w-full bg-transparent border px-2 py-1 rounded resize-y min-h-[40px] outline-none focus:bg-white focus:ring-1 focus:ring-blue-500 ${!item.comentario ? 'text-slate-400 italic' : 'text-slate-700'}`} placeholder="Sem comentários" /></td>
                        <td className="p-3 text-center"><button onClick={() => removeItem(item.id)} className="p-2 hover:text-red-500 rounded transition"><Trash2 size={16} /></button></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <div className="mt-4 flex justify-center border-t pt-4"><button onClick={addNewItem} className="flex items-center gap-2 text-sm text-blue-600 hover:text-blue-800 px-4 py-2"><Plus size={16} /> Adicionar Manualmente</button></div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
