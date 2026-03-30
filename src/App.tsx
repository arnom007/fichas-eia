import { useState, useEffect } from 'react';
import { Upload, FileText, Check, Download, RefreshCw, AlertCircle, Trash2, Plus, Zap, CheckCircle2, XCircle } from 'lucide-react';
import { GoogleGenerativeAI } from '@google/generative-ai';

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

const MetaInput = ({ label, value, onChange, placeholder, maxLength, widthClass = "w-full", disabled = false, title = "" }: any) => (
  <div className={widthClass} title={title}>
    <label className="block text-[11px] font-bold text-slate-500 uppercase tracking-wider mb-1">{label}</label>
    <input 
      type="text" 
      value={value || ''} 
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
      value={value || ''} 
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
      value={value || ''} 
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

  // Função que envia o texto extraído para a Inteligência Artificial
  const processWithAI = async (fullText: string) => {
    const apiKey = import.meta.env.VITE_GEMINI;
    
    if (!apiKey) {
      setErrorMsg('Chave da API do Gemini não encontrada. Configure a variável VITE_GEMINI no Vercel.');
      setStatus('idle');
      return;
    }

    try {
      const genAI = new GoogleGenerativeAI(apiKey);
      const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });

      const prompt = `
        Você é um assistente especializado em organizar dados de instrução aérea da AFA.
        Abaixo está o texto bruto extraído de um PDF de uma Ficha de Voo. O formato original do PDF tem colunas e comentários no final.
        Sua tarefa é ler este texto embaralhado, extrair os metadados da missão e estruturar cada item avaliado associando a manobra com seu respectivo grau, fase e comentário.

        Instruções de extração:
        1. Para o 'tipoMissao', use "Abortiva" se encontrar VMET ou VMAT. Se houver "Extra", use "Extra". Caso contrário, "Normal".
        2. Os 'itens' devem incluir tanto as manobras numeradas quanto os Itens Afetivos/Cognitivos (estes não têm número, deixe vazio).
        3. Para a 'fase' do item, use PR, RC, RM, RO, ou '--'.
        4. Associe o comentário correto a cada item. Se não houver comentário, deixe vazio "".

        Responda APENAS com um objeto JSON válido, sem formatação Markdown (\`\`\`json), estritamente neste formato:
        {
          "meta": {
            "missao": "",
            "aluno1p": "",
            "instrutor": "",
            "grauMissao": "",
            "tipoMissao": "Normal",
            "data": "",
            "fase": "",
            "aeronave": "",
            "hdep": "",
            "tev": "",
            "pousos": "",
            "parecer": ""
          },
          "itens": [
            {
              "numero": "1",
              "nome": "Partida",
              "fase": "RC",
              "grau": "5",
              "comentario": "Bom acionamento."
            }
          ]
        }

        Aqui está o texto bruto do PDF:
        ---
        ${fullText}
      `;

      const result = await model.generateContent(prompt);
      const response = await result.response;
      let textResponse = response.text();

      // Limpeza de possíveis blocos de código Markdown
      textResponse = textResponse.replace(/```json/g, '').replace(/```/g, '').trim();

      const parsedData = JSON.parse(textResponse);

      // Atualizar os estados da aplicação com os dados limpos da IA
      if (parsedData.meta) {
        setMeta({
          esquadrilha: '', // Esquadrilha não vem no PDF, usuário preenche
          aluno1p: parsedData.meta.aluno1p?.slice(0, 3) || '',
          instrutor: parsedData.meta.instrutor?.slice(0, 3) || '',
          fase: parsedData.meta.fase || '',
          aeronave: parsedData.meta.aeronave || '',
          data: parsedData.meta.data || '',
          missao: parsedData.meta.missao || '',
          grauMissao: parsedData.meta.grauMissao || '',
          tipoMissao: parsedData.meta.tipoMissao || 'Normal',
          pousos: parsedData.meta.pousos || '',
          hdep: parsedData.meta.hdep || '',
          tev: parsedData.meta.tev || '',
          parecer: parsedData.meta.parecer || ''
        });
      }

      if (parsedData.itens && Array.isArray(parsedData.itens)) {
        const itensComId = parsedData.itens.map((it: any) => ({
          id: crypto.randomUUID(),
          numero: it.numero || '',
          nome: it.nome || '',
          fase: it.fase || '--',
          grau: it.grau || '',
          comentario: it.comentario || ''
        }));
        setItems(itensComId);
      }

      setStatus('reviewing');
      setErrorMsg('');

    } catch (err) {
      console.error(err);
      setErrorMsg('Falha ao processar os dados com a IA. Verifique se o documento é válido.');
      setStatus('idle');
    }
  };

  const handleFileUpload = async (e: any) => {
    const file = e.target.files[0];
    if (!file) return;

    e.target.value = '';
    setStatus('loading');
    setErrorMsg('');

    if (file.type !== 'application/pdf') {
      setErrorMsg('Por favor, selecione um arquivo PDF válido.');
      setStatus('idle');
      return;
    }

    if (!window.pdfjsLib) {
      setErrorMsg('A biblioteca PDF.js não carregou corretamente. Recarregue a página.');
      setStatus('idle');
      return;
    }

    try {
      const arrayBuffer = await file.arrayBuffer();
      const pdf = await window.pdfjsLib.getDocument({ data: arrayBuffer }).promise;
      let fullText = '';

      // Extrai o texto puro de todas as páginas para enviar para a IA
      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        
        // Ordenação visual básica para facilitar a leitura da IA
        let sortedItems = [...textContent.items].sort((a, b) => {
          if (Math.abs(a.transform[5] - b.transform[5]) > 5) return b.transform[5] - a.transform[5];
          return a.transform[4] - b.transform[4];
        });
        
        let lastY = null;
        for (const it of sortedItems) {
           if (lastY !== null && Math.abs(it.transform[5] - lastY) > 5) fullText += '\n';
           else if (lastY !== null) fullText += ' ';
           fullText += it.str;
           lastY = it.transform[5];
        }
        fullText += '\n\n';
      }

      await processWithAI(fullText);

    } catch (err) {
      console.error(err);
      setErrorMsg('Erro ao ler o documento PDF.');
      setStatus('idle');
    }
  };

  const updateItem = (id: string, field: string, value: string) => setItems(items.map((item: any) => item.id === id ? { ...item, [field]: value } : item));
  const removeItem = (id: string) => setItems(items.filter((item: any) => item.id !== id));
  const addNewItem = () => setItems([...items, { id: crypto.randomUUID(), numero: '', nome: 'Novo Item Manual', fase: '--', grau: '', comentario: '' }]);

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
          <div><h1 className="text-3xl font-bold tracking-tight">Extrator com Inteligência Artificial</h1><p className="text-slate-500 text-lg">Processamento avançado de dados de instrução.</p></div>
        </header>
        {errorMsg && <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-lg text-red-700 font-medium flex gap-2"><AlertCircle className="text-red-500 shrink-0"/>{errorMsg}</div>}
        
        {status === 'idle' && (
          <div className="flex justify-center"><div className="bg-white p-16 rounded-2xl shadow-sm border flex flex-col items-center text-center w-full max-w-3xl border-slate-200 hover:border-blue-300 transition hover:shadow-lg"><Upload size={48} className="text-blue-600 mb-7" /><h2 className="text-2xl font-bold mb-4">Selecione o PDF da Ficha</h2><p className="text-slate-500 mb-8 max-w-md">O documento será analisado pelo Google Gemini para estruturação automática de dados e comentários.</p><label className="bg-blue-600 hover:bg-blue-700 text-white px-9 py-4.5 rounded-2xl text-xl font-bold cursor-pointer flex items-center gap-3.5 transition"><FileText size={26} /> Importar Arquivo <input type="file" accept="application/pdf" className="hidden" onClick={(e: any) => e.target.value = ''} onChange={handleFileUpload} /></label></div></div>
        )}
        
        {status === 'loading' && (
          <div className="bg-white p-20 rounded-2xl shadow-sm border flex flex-col items-center text-center"><RefreshCw className="text-blue-600 animate-spin mb-5" size={44} /><h2 className="text-2xl font-bold">Processamento IA Ativo</h2><p className="text-slate-500 mt-2">Enviando texto para análise e reconstrução de contexto...</p></div>
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
