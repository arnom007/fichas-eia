import { useState, useEffect, useRef } from 'react';
import { Upload, FileText, Check, Download, RefreshCw, AlertCircle, Trash2, Plus, Zap, CheckCircle2, XCircle, Bug } from 'lucide-react';

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
// EXTRAÇÃO DE METADADOS DO CABEÇALHO
// ---------------------------------------------------------------------------------

interface PdfToken {
  text: string;
  x: number;
  y: number;
  w: number;
  page: number;
}

const extractMetadata = (tokens: PdfToken[]) => {
  const meta: any = {};
  
  // Trabalha apenas com tokens da primeira página, região do cabeçalho (y > 670)
  const headerTokens = tokens.filter(t => t.page === 1 && t.y > 670);
  
  // Busca token que segue um label
  const findValueAfterLabel = (labelPattern: RegExp, yTolerance = 5) => {
    const labelToken = headerTokens.find(t => labelPattern.test(t.text.trim()));
    if (!labelToken) return '';
    
    // Pega todos os tokens na mesma linha, à direita do label
    const sameLineTokens = headerTokens
      .filter(t => Math.abs(t.y - labelToken.y) <= yTolerance && t.x > labelToken.x + labelToken.w - 5)
      .sort((a, b) => a.x - b.x);
    
    // Retorna o primeiro token não-vazio que não seja outro label
    for (const t of sameLineTokens) {
      const txt = t.text.trim();
      if (txt.length > 0 && !/^(MISS[AÃ]O:|DATA:|POUSOS:|TEV:|H\.\s*DEP|PROT|AERONAVE|INSTRUTOR|ALUNO|FASE)$/i.test(txt)) {
        return txt;
      }
    }
    return '';
  };

  // Fase da missão (PRÉ-SOLO, INSTRUMENTO BÁSICO, etc.)
  const faseToken = headerTokens.find(t => /^FASE:/i.test(t.text.trim()));
  if (faseToken) {
    const faseValues = headerTokens
      .filter(t => Math.abs(t.y - faseToken.y) <= 3 && t.x > faseToken.x + 20 && t.x < 350)
      .sort((a, b) => a.x - b.x)
      .map(t => t.text.trim())
      .filter(t => t.length > 0);
    meta.fase = faseValues.join(' ').trim();
  }

  // Missão (PS-11, VI-01, etc.)
  meta.missao = findValueAfterLabel(/^MISS[AÃ]O:/i);
  
  // Data
  meta.data = findValueAfterLabel(/^DATA:/i);
  
  // Pousos
  meta.pousos = findValueAfterLabel(/^POUSOS:/i);
  
  // TEV
  meta.tev = findValueAfterLabel(/^TEV:/i);
  
  // H. DEP
  meta.hdep = findValueAfterLabel(/^H\.\s*DEP/i);
  
  // Aeronave
  meta.aeronave = findValueAfterLabel(/^AERONAVE/i);
  
  // Grau de missão - aparece no canto superior direito com "GRAU" como label
  const grauLabel = tokens.find(t => t.page === 1 && t.y > 750 && /^GRAU$/i.test(t.text.trim()));
  if (grauLabel) {
    const grauValue = tokens.find(t => 
      t.page === 1 && 
      Math.abs(t.y - grauLabel.y) <= 5 && 
      t.x > grauLabel.x + 20 &&
      /^[1-6]$/.test(t.text.trim())
    );
    if (grauValue) {
      meta.grauMissao = grauValue.text.trim();
    }
  }

  // Tipo missão (Normal, Abortiva, etc.) - aparece próximo à fase
  if (faseToken) {
    const tipoToken = headerTokens.find(t => 
      Math.abs(t.y - faseToken.y) <= 3 && 
      t.x > 300 &&
      /^(Normal|Abortiva|Extra|Revis[aã]o)$/i.test(t.text.trim())
    );
    if (tipoToken) {
      const tipo = tipoToken.text.trim();
      meta.tipoMissao = tipo.charAt(0).toUpperCase() + tipo.slice(1).toLowerCase();
    }
  }

  return meta;
};

// ---------------------------------------------------------------------------------
// EXTRAÇÃO DE ITENS DE AVALIAÇÃO (LAYOUT DUAS COLUNAS)
// ---------------------------------------------------------------------------------

interface ExtractedItem {
  id: string;
  numero: string;
  nome: string;
  fase: string;
  grau: string;
  comentario: string;
  confidence?: number;
}

const extractEvaluationItems = (tokens: PdfToken[]): ExtractedItem[] => {
  const items: ExtractedItem[] = [];
  
  // Filtra tokens da área de itens (primeira página, entre y ~450 e y ~660)
  // Identifica a região de itens: abaixo do cabeçalho, acima da seção "Itens Afetivos"
  const afetivosToken = tokens.find(t => t.page === 1 && /Itens Afetivos/i.test(t.text));
  const afetivosY = afetivosToken ? afetivosToken.y : 450;
  
  // Tokens da região de itens numéricos
  const itemTokens = tokens.filter(t => 
    t.page === 1 && 
    t.y < 665 && 
    t.y > afetivosY &&
    t.text.trim().length > 0
  );

  if (itemTokens.length === 0) return items;

  // Detecta o divisor entre coluna esquerda e direita
  // A coluna esquerda tem itens com x < ~300, a direita com x >= ~300
  const COLUMN_DIVIDER = 300;
  
  // Separa tokens por coluna
  const leftTokens = itemTokens.filter(t => t.x < COLUMN_DIVIDER);
  const rightTokens = itemTokens.filter(t => t.x >= COLUMN_DIVIDER);

  // Para cada coluna, define as zonas X esperadas
  // Coluna esquerda: numero ~38-48, nome ~52, fase ~264, grau ~290
  // Coluna direita: numero ~309, nome ~323, fase ~536, grau ~561

  const extractColumnItems = (
    colTokens: PdfToken[],
    numXRange: [number, number],
    nameXMin: number,
    faseXRange: [number, number],
    grauXRange: [number, number]
  ): ExtractedItem[] => {
    const result: ExtractedItem[] = [];
    
    // Agrupa tokens por linha Y (com tolerância de 8px)
    const lineMap = new Map<number, PdfToken[]>();
    colTokens.forEach(t => {
      // Encontra ou cria a linha mais próxima
      let foundKey: number | null = null;
      for (const key of lineMap.keys()) {
        if (Math.abs(key - t.y) <= 8) {
          foundKey = key;
          break;
        }
      }
      if (foundKey !== null) {
        lineMap.get(foundKey)!.push(t);
      } else {
        lineMap.set(t.y, [t]);
      }
    });

    // Ordena linhas por Y decrescente (de cima para baixo no PDF)
    const sortedLines = Array.from(lineMap.entries())
      .sort((a, b) => b[0] - a[0])
      .map(([_, tokens]) => tokens.sort((a, b) => a.x - b.x));

    let currentItem: ExtractedItem | null = null;

    for (const lineTokens of sortedLines) {
      // Verifica se esta linha começa um novo item (tem número na zona de número)
      const numToken = lineTokens.find(t => {
        const x = t.x;
        const txt = t.text.trim();
        return x >= numXRange[0] && x <= numXRange[1] && /^\d{1,3}$/.test(txt);
      });

      if (numToken) {
        // Salva item anterior
        if (currentItem && currentItem.numero) {
          result.push(currentItem);
        }
        
        currentItem = {
          id: crypto.randomUUID(),
          numero: numToken.text.trim(),
          nome: '',
          fase: '--',
          grau: '',
          comentario: ''
        };

        // Extrai nome (tokens na zona de nome)
        const nameTokens = lineTokens.filter(t => {
          const txt = t.text.trim();
          if (txt.length === 0) return false;
          if (t === numToken) return false;
          if (/^(PR|RC|RM|RO)$/i.test(txt)) return false;
          if (/^[1-6]$/.test(txt) && t.x >= faseXRange[0] - 30) return false;
          if (t.x >= nameXMin - 5 && t.x < faseXRange[0] - 10) return true;
          return false;
        });
        
        const rawName = nameTokens.map(t => t.text.trim()).join(' ')
          .replace(/^[-–—.:\s]+/, '')
          .trim();
        currentItem.nome = rawName;

        // Extrai fase (PR, RC, RM, RO)
        const faseToken = lineTokens.find(t => {
          const txt = t.text.trim();
          return t.x >= faseXRange[0] - 15 && t.x <= faseXRange[1] + 15 && /^(PR|RC|RM|RO)$/i.test(txt);
        });
        if (faseToken) {
          currentItem.fase = faseToken.text.trim().toUpperCase();
        }

        // Extrai grau (1-6)
        const grauToken = lineTokens.find(t => {
          const txt = t.text.trim();
          return t.x >= grauXRange[0] - 15 && t.x <= grauXRange[1] + 15 && /^[1-6]$/.test(txt);
        });
        if (grauToken) {
          currentItem.grau = grauToken.text.trim();
        }
      } else if (currentItem) {
        // Linha de continuação (nome quebrado em duas linhas)
        const continuationTokens = lineTokens.filter(t => {
          const txt = t.text.trim();
          if (txt.length === 0) return false;
          if (/^(PR|RC|RM|RO)$/i.test(txt)) return false;
          if (/^[1-6]$/.test(txt) && t.x >= faseXRange[0] - 30) return false;
          if (t.x >= nameXMin - 5 && t.x < faseXRange[0] - 10) return true;
          return false;
        });
        
        if (continuationTokens.length > 0) {
          const contText = continuationTokens.map(t => t.text.trim()).join(' ')
            .replace(/^[-–—.:\s]+/, '')
            .trim();
          if (contText.length > 0) {
            currentItem.nome = (currentItem.nome + ' ' + contText).trim();
          }
        }

        // Pode ter fase/grau nesta linha de continuação
        if (!currentItem.fase || currentItem.fase === '--') {
          const faseToken = lineTokens.find(t => {
            const txt = t.text.trim();
            return t.x >= faseXRange[0] - 15 && t.x <= faseXRange[1] + 15 && /^(PR|RC|RM|RO)$/i.test(txt);
          });
          if (faseToken) {
            currentItem.fase = faseToken.text.trim().toUpperCase();
          }
        }
        if (!currentItem.grau) {
          const grauToken = lineTokens.find(t => {
            const txt = t.text.trim();
            return t.x >= grauXRange[0] - 15 && t.x <= grauXRange[1] + 15 && /^[1-6]$/.test(txt);
          });
          if (grauToken) {
            currentItem.grau = grauToken.text.trim();
          }
        }
      }
    }

    // Salva o último item
    if (currentItem && currentItem.numero) {
      result.push(currentItem);
    }

    return result;
  };

  // Extrai itens da coluna esquerda
  const leftItems = extractColumnItems(
    leftTokens,
    [30, 55],   // numXRange
    46,          // nameXMin
    [255, 280],  // faseXRange
    [280, 300]   // grauXRange
  );

  // Extrai itens da coluna direita
  const rightItems = extractColumnItems(
    rightTokens,
    [300, 320],  // numXRange
    315,         // nameXMin
    [525, 545],  // faseXRange
    [550, 575]   // grauXRange
  );

  items.push(...leftItems, ...rightItems);
  
  return items;
};

// ---------------------------------------------------------------------------------
// EXTRAÇÃO DE ITENS AFETIVOS-COGNITIVOS
// ---------------------------------------------------------------------------------

const AFFECTIVE_GRADES: Record<string, string> = {
  'PERIGOSO': '1',
  'DEFICIENTE': '2',
  'PRECISA MELHORAR': '3',
  'NORMAL': '4',
  'DESTACOU-SE': '5',
};

const extractAffectiveItems = (tokens: PdfToken[]): ExtractedItem[] => {
  const items: ExtractedItem[] = [];
  
  // Encontra a seção "Itens Afetivos - Cognitivos"
  const sectionToken = tokens.find(t => t.page === 1 && /Itens Afetivos/i.test(t.text));
  if (!sectionToken) return items;

  // Encontra o limite inferior: "Comentários:" que marca o fim da seção afetiva
  const commentsSectionToken = tokens.find(t => 
    t.page === 1 && 
    t.y < sectionToken.y && 
    /^Coment[áa]rios:$/i.test(t.text.trim())
  );
  const bottomY = commentsSectionToken ? commentsSectionToken.y : sectionToken.y - 250;

  // Lista conhecida de itens afetivos
  const KNOWN_AFFECTIVE = [
    'Comentários Gerais',
    'Debriefing',
    'Preparo de Missão',
    'Raciocínio Espacial',
    'Adaptação à Dinâmica de Voo',
    'Adaptação Fisiológica à Atividade Aérea',
    'Progresso na Instrução',
    'Reação aos Comentários',
    'Interesse na Instrução',
    'Iniciativa',
    'Aplicação de NPA',
    'Conhecimento dos Procedimentos de Emergência',
    'Conhecimento Teórico',
    'Briefing',
    'Mentalidade de Segurança',
  ];

  // Pega tokens na região afetiva
  const affTokens = tokens.filter(t =>
    t.page === 1 &&
    t.y < sectionToken.y &&
    t.y > bottomY &&
    t.text.trim().length > 0
  );

  // Agrupa por linha
  const lineMap = new Map<number, PdfToken[]>();
  affTokens.forEach(t => {
    let foundKey: number | null = null;
    for (const key of lineMap.keys()) {
      if (Math.abs(key - t.y) <= 4) {
        foundKey = key;
        break;
      }
    }
    if (foundKey !== null) {
      lineMap.get(foundKey)!.push(t);
    } else {
      lineMap.set(t.y, [t]);
    }
  });

  const sortedLines = Array.from(lineMap.entries())
    .sort((a, b) => b[0] - a[0])
    .map(([_, tokens]) => tokens.sort((a, b) => a.x - b.x));

  let itemCounter = 100; // Começar numeração em 100 para não conflitar com itens normais

  for (const lineTokens of sortedLines) {
    // Procura o nome do item afetivo
    const nameTokens = lineTokens.filter(t => t.x < 300);
    const gradeToken = lineTokens.find(t => 
      t.x >= 350 && 
      /^(PERIGOSO|DEFICIENTE|PRECISA MELHORAR|NORMAL|DESTACOU-SE)$/i.test(t.text.trim())
    );

    if (nameTokens.length === 0 || !gradeToken) continue;

    const name = nameTokens.map(t => t.text.trim()).filter(t => t.length > 0).join(' ').trim();
    
    // Ignora se não é um item afetivo conhecido (fuzzy match)
    const isAffective = KNOWN_AFFECTIVE.some(known => {
      const normalize = (s: string) => s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
      return normalize(name).includes(normalize(known).substring(0, 8));
    });
    
    if (!isAffective) continue;

    const gradeText = gradeToken.text.trim().toUpperCase();
    const gradeNum = AFFECTIVE_GRADES[gradeText] || '4';

    items.push({
      id: crypto.randomUUID(),
      numero: String(itemCounter++),
      nome: name,
      fase: '--',
      grau: gradeNum,
      comentario: gradeText,
    });
  }

  return items;
};

// ---------------------------------------------------------------------------------
// EXTRAÇÃO DE COMENTÁRIOS
// ---------------------------------------------------------------------------------

const extractComments = (tokens: PdfToken[]): Map<string, string> => {
  const comments = new Map<string, string>();

  // Encontra seção de comentários
  const commentsSections = tokens.filter(t => /^Coment[áa]rios:$/i.test(t.text.trim()));
  
  if (commentsSections.length === 0) return comments;

  // Coleta todos tokens de comentários (na seção de comentários e página 2)
  const commentTokens = tokens.filter(t => {
    // Tokens abaixo da seção "Comentários:" na página 1
    const commentSection = commentsSections.find(cs => cs.page === t.page);
    if (commentSection && t.y < commentSection.y && t.y > 50) {
      // Exclui assinaturas e rodapés
      if (/^(Ass\. Digital|MATERIAL DE ACESSO|Art\. 44|Recomenda|INSTRUTOR do voo|Autoridade Competente|\d+ de \d+)/i.test(t.text.trim())) return false;
      return true;
    }
    // Tokens na página 2 que são continuação de comentários
    if (t.page === 2 && t.y > 300) {
      const p2CommentSection = tokens.find(cs => cs.page === 2 && /^Coment[áa]rios:$/i.test(cs.text.trim()));
      if (p2CommentSection && t.y <= p2CommentSection.y) return true;
      // Ou se não tem seção de comentários na p2, mas tem tokens de comentário
      if (!p2CommentSection) {
        if (/^(Ass\. Digital|MATERIAL DE ACESSO|Art\. 44|Recomenda|INSTRUTOR do voo|Autoridade Competente|\d+ de \d+|Ciente)/i.test(t.text.trim())) return false;
        return true;
      }
    }
    return false;
  }).sort((a, b) => {
    if (a.page !== b.page) return a.page - b.page;
    if (Math.abs(a.y - b.y) > 5) return b.y - a.y;
    return a.x - b.x;
  });

  // Reconstrói texto de comentários
  let currentItemNum = '';
  let currentText = '';

  for (const token of commentTokens) {
    const txt = token.text.trim();
    if (txt.length === 0) continue;

    // Detecta início de comentário de item: "N -" ou "NN -"
    const itemMatch = txt.match(/^(\d{1,3})\s*[-–—]/);
    if (itemMatch) {
      // Salva comentário anterior
      if (currentItemNum && currentText.trim()) {
        comments.set(currentItemNum, currentText.trim());
      }
      currentItemNum = itemMatch[1];
      // O resto do texto após "N -" pode estar no mesmo token ou no próximo
      const afterDash = txt.replace(/^\d{1,3}\s*[-–—]\s*/, '').trim();
      currentText = afterDash;
      continue;
    }

    // Token que é apenas o número do item (será seguido por " -" no próximo token)
    if (/^\d{1,3}$/.test(txt) && token.x < 50) {
      if (currentItemNum && currentText.trim()) {
        comments.set(currentItemNum, currentText.trim());
      }
      currentItemNum = txt;
      currentText = '';
      continue;
    }

    // Continuação do comentário
    if (currentItemNum) {
      // Trata o header do comentário como parte do texto (ex: "Nivelamento (RO/5):")
      if (currentText.length === 0 && token.x > 50) {
        currentText = txt;
      } else {
        currentText += (currentText.length > 0 ? ' ' : '') + txt;
      }
    }
  }

  // Salva o último
  if (currentItemNum && currentText.trim()) {
    comments.set(currentItemNum, currentText.trim());
  }

  // Remove o cabeçalho dos comentários: "Nome do Item (FASE/GRAU):" → fica só o texto
  for (const [key, value] of comments.entries()) {
    // Padrão: "Nivelamento (RO/5): texto real" ou "Comentários Gerais (/NORMAL): texto"
    const cleaned = value
      .replace(/^[^(]*\([^)]*\)\s*:\s*/i, '')  // Remove "Nome (FASE/GRAU): "
      .replace(/^[-–—]\s*/, '')                  // Remove traço residual
      .trim();
    if (cleaned.length > 0) {
      comments.set(key, cleaned);
    }
  }

  return comments;
};

// ---------------------------------------------------------------------------------
// EXTRAÇÃO DO PARECER
// ---------------------------------------------------------------------------------

const extractParecer = (tokens: PdfToken[]): string => {
  // Procura "Recomendações/Parecer:" (normalmente na página 2)
  const parecerLabel = tokens.find(t => /Recomenda[çc][õo]es\/Parecer/i.test(t.text));
  if (!parecerLabel) return '';

  // Pega tokens abaixo do label na mesma página
  const parecerTokens = tokens.filter(t => 
    t.page === parecerLabel.page &&
    t.y < parecerLabel.y &&
    t.y > parecerLabel.y - 80 &&
    t.x < 400 &&
    t.text.trim().length > 0 &&
    !/^(Ass\. Digital|Ciente|INSTRUTOR|Autoridade|Cap |Maj |Ten |MATERIAL|Art\.)/.test(t.text.trim()) &&
    !/\bem \d{2}\/\d{2}\/\d{2}/.test(t.text.trim()) &&
    !/^[A-Z]{2,}\s*-\s*(Maj|Cap|Ten|1.*Ten|2.*Ten|Cel|Cad)\b/i.test(t.text.trim())
  ).sort((a, b) => b.y - a.y || a.x - b.x);

  return parecerTokens.map(t => t.text.trim()).filter(t => t.length > 0).join(' ').trim();
};

// ---------------------------------------------------------------------------------
// CÁLCULO DE CONFIANÇA E VALIDAÇÃO
// ---------------------------------------------------------------------------------

const computeConfidence = (item: ExtractedItem) => {
  let score = 0;
  if (item.numero) score += 0.3;
  if (item.nome.length > 5) score += 0.3;
  if (item.grau) score += 0.2;
  if (item.fase && item.fase !== '--') score += 0.1;
  if (item.comentario) score += 0.1;
  return score;
};

// ---------------------------------------------------------------------------------
// VISUALIZADOR DE DEBUG
// ---------------------------------------------------------------------------------

const PdfViewer = ({ pdf, tokens }: { pdf: any, tokens: PdfToken[] }) => {
  const canvasRef = useRef<HTMLCanvasElement>(null);

  useEffect(() => {
    if (!pdf || !tokens.length) return;
    
    const render = async () => {
      const page = await pdf.getPage(1);
      const viewport = page.getViewport({ scale: 1.5 });
      const canvas = canvasRef.current;
      if (!canvas) return;
      
      const ctx = canvas.getContext('2d');
      if (!ctx) return;

      canvas.width = viewport.width;
      canvas.height = viewport.height;

      await page.render({ canvasContext: ctx, viewport }).promise;

      ctx.strokeStyle = 'red';
      ctx.fillStyle = 'red';
      ctx.font = '10px Arial';

      tokens.forEach((t: any, i: number) => {
        const x = t.x * 1.5;
        const y = viewport.height - (t.y * 1.5);
        
        ctx.strokeRect(x, y - 12, 30, 12);
        ctx.fillText(`${i}`, x, y);
      });
    };

    render();
  }, [pdf, tokens]);

  return (
    <div className="w-full overflow-auto border border-slate-300 rounded-xl shadow-inner bg-slate-100 p-4">
      <canvas ref={canvasRef} className="mx-auto bg-white shadow-md" />
    </div>
  );
};

// ---------------------------------------------------------------------------------
// COMPONENTE PRINCIPAL (APP)
// ---------------------------------------------------------------------------------

export default function App() {
  const [status, setStatus] = useState('idle'); 
  const [items, setItems] = useState<any[]>([]);
  const [errorMsg, setErrorMsg] = useState('');
  const [showModal, setShowModal] = useState(false);
  const [modalState, setModalState] = useState('sending');
  const [modalMessage, setModalMessage] = useState('');
  
  const [debugMode, setDebugMode] = useState(false);
  const [rawPdfTokens, setRawPdfTokens] = useState<PdfToken[]>([]);
  const [pdfDocument, setPdfDocument] = useState<any>(null);

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

  const handleFileUpload = async (e: any) => {
    const file = e.target.files[0];
    if (!file) return; e.target.value = ''; setStatus('loading'); setErrorMsg(''); setDebugMode(false);
    if (file.type !== 'application/pdf') { setErrorMsg('Formato inválido. Selecione um documento PDF.'); setStatus('idle'); return; }
    if (!window.pdfjsLib) { setErrorMsg('Módulo de leitura indisponível. Recarregue a página.'); setStatus('idle'); return; }

    try {
      const arrayBuffer = await file.arrayBuffer();
      const pdf = await window.pdfjsLib.getDocument({ data: arrayBuffer }).promise;
      setPdfDocument(pdf);
      
      const allTokens: PdfToken[] = [];

      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        
        for (const it of textContent.items) {
          const text = it.str;
          if (text === undefined || text === null) continue;
          allTokens.push({
            text: text,
            x: it.transform[4],
            y: it.transform[5],
            w: it.width || 0,
            page: i
          });
        }
      }

      // Debug: salva tokens da página 1 para visualização
      setRawPdfTokens(allTokens.filter(t => t.page === 1));

      // --- EXTRAÇÃO PRINCIPAL ---
      
      // 1. Metadados do cabeçalho
      const metadata = extractMetadata(allTokens);
      
      // 2. Parecer
      const parecer = extractParecer(allTokens);
      
      // 3. Itens de avaliação (tabela duas colunas)
      const evalItems = extractEvaluationItems(allTokens);
      
      // 4. Itens afetivos-cognitivos
      const affItems = extractAffectiveItems(allTokens);
      
      // 5. Comentários
      const comments = extractComments(allTokens);
      
      // 6. Associa comentários aos itens
      for (const item of evalItems) {
        const comment = comments.get(item.numero);
        if (comment) {
          item.comentario = comment;
        }
      }
      
      // Combina todos os itens
      const allItems = [...evalItems, ...affItems];
      
      // Ordena por número
      allItems.sort((a, b) => {
        const numA = parseInt(a.numero) || 999;
        const numB = parseInt(b.numero) || 999;
        return numA - numB;
      });

      // Calcula confiança
      allItems.forEach(item => {
        item.confidence = computeConfidence(item);
      });

      // Atualiza metadados
      setMeta(prev => ({
        ...prev,
        ...metadata,
        parecer: parecer || prev.parecer,
      }));

      if (allItems.length > 0) {
        setItems(allItems);
        setStatus('reviewing');
        setErrorMsg('');
      } else {
        setErrorMsg('Falha na extração. Verifique a legibilidade do arquivo.');
        setStatus('idle');
      }
    } catch (err) {
      console.error('Erro ao processar PDF:', err);
      setErrorMsg('Falha no processamento. Documento corrompido ou inacessível.');
      setStatus('idle');
    }
  };

  const updateItem = (id: string, field: string, value: string) => setItems(items.map((item: any) => item.id === id ? { ...item, [field]: value } : item));
  const removeItem = (id: string) => setItems(items.filter((item: any) => item.id !== id));
  const addNewItem = () => setItems([...items, { id: crypto.randomUUID(), numero: '', nome: 'Registro Manual', fase: '--', grau: '', comentario: '', confidence: 1.0 }]);

  const buildPayload = () => items.map((item: any) => ({ data: meta.data, esquadrilha: meta.esquadrilha, missao: meta.missao, grauMissao: (meta.tipoMissao === 'Abortiva' || meta.tipoMissao === 'Extra') ? '' : meta.grauMissao, aluno1p: meta.aluno1p, instrutor: meta.instrutor, faseMissao: meta.fase, aeronave: meta.aeronave, hdep: meta.hdep, pousos: meta.pousos, tev: meta.tev, parecer: meta.parecer, numero: item.numero, nome: item.nome, faseItem: item.fase, grau: item.grau, comentario: item.comentario, tipoMissao: meta.tipoMissao }));

  const exportCSV = () => {
    const hdrs = ['Data', 'Esquadrilha', 'Missão', 'Grau Missão', '1P / AL', 'IN', 'Fase Missão', 'Anv', 'H.Dep', 'Pousos', 'TEV', 'Parecer', 'Nº Item', 'Nome', 'Fase Item', 'Grau/Menção', 'Comentário', 'Tipo'];
    const csvContent = [ hdrs.join(','), ...buildPayload().map(r => `"${r.data}","${r.esquadrilha}","${r.missao}","${r.grauMissao}","${r.aluno1p}","${r.instrutor}","${r.faseMissao}","${r.aeronave}","${r.hdep}","${r.pousos}","${r.tev}","${(r.parecer || '').replace(/\"/g, '""')}","${r.numero}","${r.nome}","${r.faseItem}","${r.grau}","${(r.comentario || '').replace(/\"/g, '""')}","${r.tipoMissao}"`) ].join('\n');
    const blob = new Blob(['\ufeff' + csvContent], { type: 'text/csv;charset=utf-8;' });
    const lnk = document.createElement('a'); lnk.href = URL.createObjectURL(blob); lnk.setAttribute('download', `Ficha_${meta.missao || 'Extracao'}_${new Date().toISOString().slice(0,10)}.csv`); lnk.click();
  };

  const sendWebhook = async () => {
    if (!meta.esquadrilha) { setErrorMsg('Obrigatório informar a Esquadrilha.'); window.scrollTo(0, 0); return; }
    if (!meta.aluno1p || !meta.instrutor) { setErrorMsg('Obrigatório informar trigramas de Aluno e Instrutor.'); window.scrollTo(0, 0); return; }
    setModalState('sending'); setShowModal(true);
    try {
      const res = await fetch(WEBHOOK_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(buildPayload()) });
      if (res.ok) { setModalState('success'); setModalMessage('Integração concluída com sucesso.'); } 
      else throw new Error('A requisição falhou.');
    } catch { setModalState('error'); setModalMessage('Erro de conexão. A exportação manual em CSV continua disponível.'); }
  };

  const resetAfterSuccess = () => { setShowModal(false); if (modalState === 'success') { setStatus('idle'); setItems([]); setPdfDocument(null); setMeta({ esquadrilha: '', aluno1p: '', instrutor: '', fase: '', aeronave: '', data: '', missao: '', grauMissao: '', tipoMissao: 'Normal', pousos: '', hdep: '', tev: '', parecer: '' }); } };

  return (
    <div className="min-h-screen w-full bg-slate-50 text-slate-800 font-sans p-4 md:p-8 relative">
      {showModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/70 backdrop-blur-sm px-4">
          <div className="bg-white rounded-2xl shadow-2xl p-8 max-w-sm w-full text-center flex flex-col items-center">
            {modalState === 'sending' && (<><RefreshCw size={44} className="animate-spin text-blue-600 mb-6" /><h2 className="text-2xl font-bold mb-1">Processando</h2><p className="text-slate-500 mb-3">Sincronizando com a base de dados...</p></>)}
            {modalState === 'success' && (<><CheckCircle2 size={52} className="text-green-500 mb-6" /><h2 className="text-2xl font-bold mb-1">Finalizado</h2><p className="text-slate-500 mb-8">{modalMessage}</p><button onClick={resetAfterSuccess} className="w-full bg-slate-800 text-white font-bold py-3.5 px-6 rounded-xl text-lg">Nova Extração</button></>)}
            {modalState === 'error' && (<><XCircle size={52} className="text-red-500 mb-6" /><h2 className="text-2xl font-bold mb-1">Aviso</h2><p className="text-slate-500 mb-8">{modalMessage}</p><button onClick={() => setShowModal(false)} className="w-full bg-slate-100 text-slate-800 font-bold py-3 px-6 rounded-xl">Reconectar</button></>)}
          </div>
        </div>
      )}
      
      <div className={`max-w-7xl mx-auto transition-opacity ${showModal ? 'opacity-25' : 'opacity-100'}`}>
        <header className="mb-8 bg-white p-6 rounded-xl shadow-sm border flex items-center justify-between gap-5">
          <div className="flex items-center gap-5">
            <div className="w-14 h-14 bg-blue-600 rounded-lg flex items-center justify-center text-white shrink-0"><FileText size={32} /></div>
            <div><h1 className="text-3xl font-bold tracking-tight">Análise Estruturada EIA</h1><p className="text-slate-500 text-lg">Processamento de Fichas de Avaliação.</p></div>
          </div>
          {status === 'reviewing' && (
            <button onClick={() => setDebugMode(!debugMode)} className={`p-3 rounded-full transition ${debugMode ? 'bg-red-100 text-red-600' : 'bg-slate-100 text-slate-400 hover:bg-slate-200'}`} title="Módulo de Depuração Visual">
              <Bug size={24} />
            </button>
          )}
        </header>

        {errorMsg && <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-lg text-red-700 font-medium flex gap-2"><AlertCircle className="text-red-500 shrink-0"/>{errorMsg}</div>}
        
        {status === 'idle' && (
          <div className="flex justify-center"><div className="bg-white p-16 rounded-2xl shadow-sm border flex flex-col items-center text-center w-full max-w-3xl border-slate-200 hover:border-blue-300 transition hover:shadow-lg"><Upload size={48} className="text-blue-600 mb-7" /><h2 className="text-2xl font-bold mb-4">Seleção de Documento</h2><p className="text-slate-500 mb-8 max-w-md">Importação de Fichas em formato PDF. A arquitetura detecta automaticamente a estrutura do documento.</p><label className="bg-blue-600 hover:bg-blue-700 text-white px-9 py-4.5 rounded-2xl text-xl font-bold cursor-pointer flex items-center gap-3.5 transition"><FileText size={26} /> Carregar PDF <input type="file" accept="application/pdf" className="hidden" onClick={(e: any) => e.target.value = ''} onChange={handleFileUpload} /></label></div></div>
        )}
        
        {status === 'loading' && (
          <div className="bg-white p-20 rounded-2xl shadow-sm border flex flex-col items-center text-center"><RefreshCw className="text-blue-600 animate-spin mb-5" size={44} /><h2 className="text-2xl font-bold">Extração em Andamento</h2><p className="text-slate-500 mt-2">Mapeamento vetorial e reconciliação de dados ativados...</p></div>
        )}
        
        {status === 'reviewing' && (
          <div className="space-y-6 relative z-10">
            {debugMode && (
              <div className="mb-6">
                <h3 className="text-lg font-bold mb-3 flex items-center gap-2"><Bug size={20}/> Visualização de OCR</h3>
                <PdfViewer pdf={pdfDocument} tokens={rawPdfTokens} />
              </div>
            )}

            <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-7">
              <h3 className="text-xl font-bold mb-5 flex items-center gap-2.5 border-b pb-4 border-slate-100"><FileText className="text-blue-500" size={22} /> Cabeçalho da Missão</h3>
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
              <div className="mt-5 pt-5 border-t border-slate-100"><MetaTextarea label="Parecer do Comandante" value={meta.parecer} onChange={(e: any) => updateMeta('parecer', e.target.value)} placeholder="Registro de observações operacionais..." /></div>
            </div>
            
            <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
              <div className="p-6 border-b flex flex-col lg:flex-row justify-between gap-4 bg-slate-50/50">
                <h2 className="text-xl font-bold flex items-center gap-2.5"><Check className="text-green-500" size={26} /> Tabela de Avaliação ({items.length} itens)</h2>
                <div className="flex flex-wrap gap-3">
                  <button onClick={() => { setStatus('idle'); setItems([]); setRawPdfTokens([]); setPdfDocument(null); setDebugMode(false); }} className="px-5 py-3 text-sm font-semibold border bg-white hover:bg-slate-50 rounded-xl">Descartar</button>
                  <button onClick={exportCSV} className="px-5 py-3 text-sm font-semibold border bg-white hover:bg-slate-50 rounded-xl flex items-center gap-2"><Download size={18} /> Exportar CSV</button>
                  <button onClick={sendWebhook} className="px-9 py-3 text-lg text-white bg-blue-600 hover:bg-blue-700 rounded-xl font-bold flex items-center gap-2.5 transition active:scale-95"><Zap size={22} /> Salvar Registros</button>
                </div>
              </div>
              <div className="p-7 overflow-x-auto">
                <table className="w-full text-left border-collapse min-w-[950px] relative z-20">
                  <thead><tr className="bg-slate-50 border-b border-slate-200 text-xs text-slate-500 uppercase font-bold tracking-wider"><th className="p-3 w-8 text-center" title="Acurácia da Extração">C.</th><th className="p-3 w-16">Item</th><th className="p-3">Manobra / Competência</th><th className="p-3 w-20 text-center">Fase</th><th className="p-3 w-24 text-center">Grau</th><th className="p-3">Observações Mapeadas</th><th className="p-3 w-12 text-center">Ações</th></tr></thead>
                  <tbody className="text-sm">
                    {items.map((it) => (
                      <tr key={it.id} className="border-b border-slate-100 hover:bg-slate-50/40 text-slate-800">
                        <td className="p-3 text-center">
                          <div className={`w-3 h-3 rounded-full mx-auto ${it.confidence > 0.6 ? 'bg-green-400' : (it.confidence > 0.4 ? 'bg-yellow-400' : 'bg-red-500 animate-pulse')}`} title={`Acurácia calculada: ${Math.round(it.confidence * 100)}%`}></div>
                        </td>
                        <td className="p-3 font-mono font-bold text-slate-400"><input value={it.numero} onChange={(e) => updateItem(it.id, 'numero', e.target.value)} className="w-full bg-transparent px-1" placeholder="--" /></td>
                        <td className="p-3 font-medium"><input value={it.nome} onChange={(e) => updateItem(it.id, 'nome', e.target.value)} className={`w-full bg-transparent px-1 ${getGradeColorClass(it.grau)}`} placeholder="Item de avaliação..." /></td>
                        <td className="p-3 text-center font-bold text-blue-700"><input value={it.fase} onChange={(e) => updateItem(it.id, 'fase', e.target.value)} className="w-full bg-transparent px-1 text-center" placeholder="--" /></td>
                        <td className="p-3 text-center font-extrabold"><input value={it.grau} onChange={(e) => updateItem(it.id, 'grau', e.target.value)} className={`w-full bg-transparent px-1 uppercase text-center ${getGradeColorClass(it.grau)}`} placeholder="--" /></td>
                        <td className="p-3"><textarea value={it.comentario} onChange={(e) => updateItem(it.id, 'comentario', e.target.value)} className="w-full bg-transparent border-slate-200 focus:border-blue-300 focus:bg-white text-slate-700 border px-2.5 py-1.5 rounded-lg resize-y min-h-[44px] text-xs leading-relaxed" placeholder="Adicionar dados..." /></td>
                        <td className="p-3 text-center"><button onClick={() => removeItem(it.id)} className="p-2 hover:bg-red-50 hover:text-red-500 text-slate-300 rounded-lg"><Trash2 size={16} /></button></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <div className="mt-5 flex justify-center border-t border-slate-100 pt-5 relative z-20"><button onClick={addNewItem} className="flex items-center gap-2 font-semibold text-blue-600 hover:text-blue-800 px-5 py-2.5 bg-blue-50/50 rounded-lg active:scale-95 transition"><Plus size={18} /> Inserir Linha Manual</button></div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
