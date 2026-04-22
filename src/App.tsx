import { useState, useEffect, useRef } from 'react';
import { Upload, FileText, Check, Download, RefreshCw, AlertCircle, Trash2, Plus, Zap, CheckCircle2, XCircle, Bug, X } from 'lucide-react';

declare global {
  interface Window {
    pdfjsLib: any;
  }
}

// ---------------------------------------------------------------------------------
// UTILITÁRIO: Cor do texto baseada no grau/menção
// ---------------------------------------------------------------------------------
const getGradeColorClass = (grau: string) => {
  const g = grau?.toUpperCase().trim() || '';
  if (g === '1' || g === 'PERIGOSO' || g === '2' || g === 'DEFICIENTE') return 'text-red-500 font-bold';
  if (g === '3' || g === 'PRECISA MELHORAR') return 'text-yellow-600 font-bold';
  if (g === '4' || g === 'NORMAL') return 'text-green-600 font-bold';
  if (g === '5' || g === 'DESTACOU-SE') return 'text-blue-500 font-bold';
  if (g === '6') return 'text-blue-800 font-bold';
  return 'text-slate-900 font-medium';
};

// ---------------------------------------------------------------------------------
// COMPONENTES DE FORMULÁRIO REUTILIZÁVEIS
// ---------------------------------------------------------------------------------

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

    items.push({
      id: crypto.randomUUID(),
      numero: String(itemCounter++),
      nome: name,
      fase: '--',
      grau: gradeText,
      comentario: '',
    });
  }

  return items;
};

// ---------------------------------------------------------------------------------
// EXTRAÇÃO DE COMENTÁRIOS
// ---------------------------------------------------------------------------------

const extractComments = (tokens: PdfToken[]): Map<string, string> => {
  const comments = new Map<string, string>();

  // Encontra todas as seções "Comentários:" em qualquer página
  const commentsSections = tokens.filter(t => /^Coment[áa]rios:$/i.test(t.text.trim()));
  
  if (commentsSections.length === 0) return comments;

  // Coleta tokens de comentários de TODAS as páginas
  // Detecta limites inferiores: "Recomendações/Parecer:" marca o fim dos comentários na página
  const parecerBoundaries = tokens.filter(t => /Recomenda[çc][õo]es\/Parecer/i.test(t.text));

  const commentTokens = tokens.filter(t => {
    const txt = t.text.trim();
    if (txt.length === 0) return false;
    
    // Exclui assinaturas, rodapés e metadados em qualquer página
    if (/^(Ass\. Digital|MATERIAL DE ACESSO|Art\. 44|Recomenda|INSTRUTOR do voo|Autoridade Competente|\d+ de \d+|Ciente|Prossegue)/i.test(txt)) return false;
    // Exclui nomes com posto/graduação (assinaturas)
    if (/^[A-Z]{2,}.*[-–—]\s*(Maj|Cap|Ten|1.*Ten|2.*Ten|Cel|Cad|Cmte)\b/i.test(txt)) return false;
    if (/[-–—]\s*(Maj|Cap|Ten|Cel|Cad|Cmte)\s+(QOAV|CFOAV)/i.test(txt)) return false;
    // Exclui tokens com datas de assinatura
    if (/\bem \d{2}\/\d{2}\/\d{2}/.test(txt)) return false;

    // Verifica se está abaixo do "Recomendações/Parecer:" na mesma página (= fora da zona de comentários)
    const parecerOnPage = parecerBoundaries.find(p => p.page === t.page);
    if (parecerOnPage && t.y <= parecerOnPage.y) return false;

    // Verifica se está numa zona de comentários em qualquer página
    const pageCommentSection = commentsSections.find(cs => cs.page === t.page);
    if (pageCommentSection && t.y < pageCommentSection.y && t.y > 50) {
      return true;
    }

    // Para páginas que continuam comentários da página anterior
    // Usa y < 760 para excluir cabeçalho da página (onde fica o "1 de 2")
    if (!pageCommentSection && t.page > 1 && t.y > 50 && t.y < 760) {
      const prevPageHasComments = commentsSections.some(cs => cs.page < t.page);
      if (prevPageHasComments) {
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
      .replace(/\s+\d+\s+de\s+\d+\s*$/i, '')    // Remove índice de página residual ("1 de 2")
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
// INTERFACES E TIPOS PARA O SISTEMA MULTI-FICHA
// ---------------------------------------------------------------------------------

// Metadados do cabeçalho de cada ficha (esquadrilha, aluno, missão, etc.)
// Cada campo corresponde a um dado extraído do PDF ou preenchido manualmente
interface MetaData {
  esquadrilha: string;  // Esquadrilha do aluno (Antares, Vega, Castor, Sirius)
  aluno1p: string;      // Trigrama do 1P / Aluno (ex: MTA) — OBRIGATÓRIO
  instrutor: string;    // Trigrama do Instrutor IN (ex: MOT) — OBRIGATÓRIO
  fase: string;         // Fase da missão (PRÉ-SOLO, INSTRUMENTO BÁSICO, etc.)
  aeronave: string;     // Matrícula / código da aeronave
  data: string;         // Data da missão
  missao: string;       // Código da missão (ex: PS-01, VI-03)
  grauMissao: string;   // Grau geral da missão (1-6)
  tipoMissao: string;   // Tipo: Normal, Abortiva, Extra, Revisão
  pousos: string;       // Quantidade de pousos
  hdep: string;         // Hora de decolagem
  tev: string;          // Tempo efetivo de voo
  parecer: string;      // Parecer/recomendações do comandante
}

// Cada aba (tab) = uma ficha PDF carregada com seus dados extraídos
interface FichaTab {
  id: string;                // Identificador único da aba
  fileName: string;          // Nome do arquivo PDF original (ex: "Ficha_MTA_PS-01.pdf")
  meta: MetaData;            // Metadados extraídos do cabeçalho
  items: any[];              // Itens de avaliação (notas, competências, afetivos)
  pdfDocument: any;          // Referência ao documento PDF (para o visualizador debug)
  rawPdfTokens: PdfToken[];  // Tokens brutos da página 1 (para debug)
}

// Status individual do envio de cada ficha (usado no popup de progresso)
interface SendStatus {
  fileName: string;                                     // Nome do arquivo
  status: 'waiting' | 'sending' | 'success' | 'error';  // Estado atual do envio
  message?: string;                                      // Mensagem descritiva (erro, etc.)
}

// Valores padrão para metadados de uma nova ficha
const INITIAL_META: MetaData = {
  esquadrilha: '', aluno1p: '', instrutor: '', fase: '',
  aeronave: '', data: '', missao: '', grauMissao: '',
  tipoMissao: 'Normal', pousos: '', hdep: '', tev: '', parecer: ''
};

// ---------------------------------------------------------------------------------
// COMPONENTE PRINCIPAL (APP) — VERSÃO MULTI-FICHA COM SISTEMA DE ABAS
// ---------------------------------------------------------------------------------

export default function App() {

  // ============================================================================
  // ESTADO PRINCIPAL
  // ============================================================================

  // Array de fichas carregadas — cada ficha é uma aba com seus próprios dados
  const [tabs, setTabs] = useState<FichaTab[]>([]);

  // Índice da aba ativa (a que está sendo visualizada/editada no momento)
  const [activeTabIndex, setActiveTabIndex] = useState(0);

  // Status geral: 'idle' = aguardando upload, 'loading' = processando PDFs, 'reviewing' = revisando fichas
  const [status, setStatus] = useState<string>('idle');

  // Mensagem de erro exibida no topo da página
  const [errorMsg, setErrorMsg] = useState('');

  // Modo debug: exibe visualização dos tokens extraídos do PDF
  const [debugMode, setDebugMode] = useState(false);

  // Progresso do carregamento de múltiplos PDFs (ex: "Processando 3 de 5")
  const [loadingProgress, setLoadingProgress] = useState({ current: 0, total: 0 });

  // ============================================================================
  // ESTADO DOS POPUPS
  // ============================================================================

  // Popup de confirmação: "Todas as fichas verificadas?"
  const [showConfirmDialog, setShowConfirmDialog] = useState(false);

  // Popup de progresso do envio (barra de progresso + tasklist)
  const [showProgressDialog, setShowProgressDialog] = useState(false);

  // Lista de status individuais de envio de cada ficha
  const [sendStatuses, setSendStatuses] = useState<SendStatus[]>([]);

  // Porcentagem geral do envio (0 a 100)
  const [sendProgress, setSendProgress] = useState(0);

  // Indica se o envio de todas as fichas foi concluído
  const [sendComplete, setSendComplete] = useState(false);

  // Resumo final: quantas fichas foram enviadas com sucesso / com erro
  const [sendSummary, setSendSummary] = useState({ success: 0, errors: 0 });

  // ============================================================================
  // CONSTANTES
  // ============================================================================

  // URL do script Google Apps Script que recebe os dados da ficha
  // >>> EDITE ESTA URL SE MUDAR O SCRIPT DO GOOGLE <<<
  const WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbxLlUKIeUnaLW2VeWOpIG5ZtrrAFy_Qg9YQTq5fG4HrMUg7kt196zcFAt4jOjBrMsEE/exec";

  // ============================================================================
  // INICIALIZAÇÃO (carrega a biblioteca PDF.js do CDN)
  // ============================================================================

  useEffect(() => {
    document.body.style.display = 'block';
    document.body.style.margin = '0';
    document.documentElement.style.backgroundColor = '#f8fafc';
    const rootNode = document.getElementById('root');
    if (rootNode) rootNode.style.width = '100%';

    // Carrega a biblioteca PDF.js para extrair texto dos PDFs
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js';
    script.onload = () => {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc =
        'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';
    };
    document.body.appendChild(script);
  }, []);

  // ============================================================================
  // ATALHO: A FICHA (ABA) ATIVA
  // ============================================================================

  // Retorna a ficha atualmente visível, ou null se não houver nenhuma carregada
  const activeTab = tabs.length > 0 ? tabs[activeTabIndex] : null;

  // ============================================================================
  // PROCESSAMENTO DE UM ÚNICO ARQUIVO PDF
  // ============================================================================

  // Extrai tokens, metadados, itens e comentários de um arquivo PDF.
  // Retorna um objeto FichaTab pronto para exibição, ou null se a extração falhar.
  const processFile = async (file: File): Promise<FichaTab | null> => {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await window.pdfjsLib.getDocument({ data: arrayBuffer }).promise;

    // Extrai todos os tokens de texto de todas as páginas do PDF
    const allTokens: PdfToken[] = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const textContent = await page.getTextContent();
      for (const it of textContent.items) {
        const text = it.str;
        if (text === undefined || text === null) continue;
        allTokens.push({
          text, x: it.transform[4], y: it.transform[5],
          w: it.width || 0, page: i
        });
      }
    }

    // --- EXTRAÇÃO PRINCIPAL (usa as mesmas funções de sempre) ---
    const metadata = extractMetadata(allTokens);
    const parecer = extractParecer(allTokens);
    const evalItems = extractEvaluationItems(allTokens);
    const affItems = extractAffectiveItems(allTokens);
    const comments = extractComments(allTokens);

    // Associa comentários aos itens de avaliação (match direto por número)
    for (const item of evalItems) {
      const comment = comments.get(item.numero);
      if (comment) item.comentario = comment;
    }

    // Associa comentários aos itens afetivos (match por nome no cabeçalho)
    const normalize = (s: string) =>
      s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();

    // Re-extrai comentários brutos (sem limpeza de cabeçalho) para fazer match afetivo
    const rawComments = new Map<string, string>();
    {
      let curNum = '';
      let curTxt = '';
      const commentsSections2 = allTokens.filter((t: PdfToken) =>
        /^Coment[áa]rios:$/i.test(t.text.trim()));
      const parecerBoundaries2 = allTokens.filter((t: PdfToken) =>
        /Recomenda[çc][õo]es\/Parecer/i.test(t.text));

      const allCommentTokens = allTokens.filter((t: PdfToken) => {
        const txt = t.text.trim();
        if (txt.length === 0) return false;
        if (/^(Ass\. Digital|MATERIAL DE ACESSO|Art\. 44|Recomenda|INSTRUTOR do voo|Autoridade Competente|\d+ de \d+|Ciente|Prossegue)/i.test(txt)) return false;
        if (/^[A-Z]{2,}.*[-–—]\s*(Maj|Cap|Ten|1.*Ten|2.*Ten|Cel|Cad|Cmte)\b/i.test(txt)) return false;
        if (/[-–—]\s*(Maj|Cap|Ten|Cel|Cad|Cmte)\s+(QOAV|CFOAV)/i.test(txt)) return false;
        if (/\bem \d{2}\/\d{2}\/\d{2}/.test(txt)) return false;
        const parecerOnPage = parecerBoundaries2.find((p: PdfToken) => p.page === t.page);
        if (parecerOnPage && t.y <= parecerOnPage.y) return false;
        const pageCS = commentsSections2.find((cs: PdfToken) => cs.page === t.page);
        if (pageCS && t.y < pageCS.y && t.y > 50) return true;
        if (!pageCS && t.page > 1 && t.y > 50 && t.y < 760) {
          if (commentsSections2.some((cs: PdfToken) => cs.page < t.page)) return true;
        }
        return false;
      }).sort((a: PdfToken, b: PdfToken) => {
        if (a.page !== b.page) return a.page - b.page;
        if (Math.abs(a.y - b.y) > 5) return b.y - a.y;
        return a.x - b.x;
      });

      for (const token of allCommentTokens) {
        const txt = token.text.trim();
        if (!txt) continue;
        const m = txt.match(/^(\d{1,3})\s*[-–—]/);
        if (m) {
          if (curNum && curTxt.trim()) rawComments.set(curNum, curTxt.trim());
          curNum = m[1];
          curTxt = txt.replace(/^\d{1,3}\s*[-–—]\s*/, '').trim();
        } else if (/^\d{1,3}$/.test(txt) && token.x < 50) {
          if (curNum && curTxt.trim()) rawComments.set(curNum, curTxt.trim());
          curNum = txt;
          curTxt = '';
        } else if (curNum) {
          curTxt += (curTxt ? ' ' : '') + txt;
        }
      }
      if (curNum && curTxt.trim()) rawComments.set(curNum, curTxt.trim());
    }

    // Associa comentários brutos aos itens afetivos pelo nome no cabeçalho
    for (const [commentNum, rawText] of rawComments.entries()) {
      // Pula se já foi associado a um item de avaliação normal
      if (evalItems.some(i => i.numero === commentNum)) continue;
      const headerMatch = rawText.match(/^([^(/]+)/);
      if (!headerMatch) continue;
      const commentItemName = normalize(headerMatch[1]);
      const matchedAff = affItems.find(aff => {
        const affName = normalize(aff.nome);
        return commentItemName.includes(affName.substring(0, 8)) ||
               affName.includes(commentItemName.substring(0, 8));
      });
      if (matchedAff) {
        const cleaned = rawText
          .replace(/^[^(]*\([^)]*\)\s*:\s*/i, '')
          .replace(/^[-–—]\s*/, '')
          .replace(/\s+\d+\s+de\s+\d+\s*$/i, '')
          .trim();
        if (cleaned.length > 0) matchedAff.comentario = cleaned;
      }
    }

    // Combina todos os itens, ordena por número e calcula confiança
    const allItems = [...evalItems, ...affItems];
    allItems.sort((a, b) => (parseInt(a.numero) || 999) - (parseInt(b.numero) || 999));
    allItems.forEach(item => { item.confidence = computeConfidence(item); });

    // Se nenhum item foi extraído, o PDF provavelmente não é válido
    if (allItems.length === 0) return null;

    // Retorna a ficha pronta para exibição
    return {
      id: crypto.randomUUID(),
      fileName: file.name,
      meta: { ...INITIAL_META, ...metadata, parecer: parecer || '' },
      items: allItems,
      pdfDocument: pdf,
      rawPdfTokens: allTokens.filter(t => t.page === 1),
    };
  };

  // ============================================================================
  // UPLOAD DE MÚLTIPLOS PDFs (seleção de vários arquivos de uma vez)
  // ============================================================================

  const handleFileUpload = async (e: any) => {
    const files = Array.from(e.target.files) as File[];
    if (!files.length) return;
    e.target.value = ''; // Permite re-selecionar os mesmos arquivos depois

    // Verifica se todos os arquivos selecionados são PDF
    const invalidFile = files.find(f => f.type !== 'application/pdf');
    if (invalidFile) {
      setErrorMsg(`Arquivo "${invalidFile.name}" não é PDF. Selecione apenas documentos PDF.`);
      return;
    }

    // Verifica se a biblioteca PDF.js está carregada
    if (!window.pdfjsLib) {
      setErrorMsg('Módulo de leitura indisponível. Recarregue a página.');
      return;
    }

    setStatus('loading');
    setErrorMsg('');
    setDebugMode(false);
    setLoadingProgress({ current: 0, total: files.length });

    const newTabs: FichaTab[] = [];
    const errors: string[] = [];

    // Processa cada arquivo PDF individualmente
    for (let fileIdx = 0; fileIdx < files.length; fileIdx++) {
      // Atualiza o indicador de progresso ("Processando 3 de 5...")
      setLoadingProgress({ current: fileIdx + 1, total: files.length });

      try {
        const result = await processFile(files[fileIdx]);
        if (result) {
          newTabs.push(result);
        } else {
          errors.push(files[fileIdx].name);
        }
      } catch (err) {
        console.error(`Erro ao processar ${files[fileIdx].name}:`, err);
        errors.push(files[fileIdx].name);
      }
    }

    // Adiciona as novas abas ao array existente
    // (permite clicar "Adicionar" para carregar mais fichas depois)
    if (newTabs.length > 0) {
      setTabs(prev => {
        const allTabs = [...prev, ...newTabs];
        // Se não havia abas antes, seleciona a primeira nova
        if (prev.length === 0) setActiveTabIndex(0);
        return allTabs;
      });
      setStatus('reviewing');
    }

    // Se algum arquivo falhou na extração, exibe aviso
    if (errors.length > 0) {
      setErrorMsg(`Falha na extração de: ${errors.join(', ')}`);
      if (newTabs.length === 0 && tabs.length === 0) setStatus('idle');
    }
  };

  // ============================================================================
  // MANIPULAÇÃO DE METADADOS E ITENS (sempre da aba ativa)
  // ============================================================================

  // Atualiza um campo dos metadados da aba ativa
  const updateMeta = (field: string, value: string) => {
    // Trigramas são sempre maiúsculos e limitados a 3 caracteres
    if (field === 'aluno1p' || field === 'instrutor') {
      value = value.toUpperCase().slice(0, 3);
    }

    setTabs(prev => prev.map((tab, i) => {
      if (i !== activeTabIndex) return tab;
      const newMeta = { ...tab.meta, [field]: value };
      // Missões Abortivas ou Extras não possuem grau de missão
      if (field === 'tipoMissao' && (value === 'Abortiva' || value === 'Extra')) {
        newMeta.grauMissao = '';
      }
      return { ...tab, meta: newMeta };
    }));
  };

  // Atualiza um campo de um item de avaliação na aba ativa
  const updateItem = (id: string, field: string, value: string) => {
    setTabs(prev => prev.map((tab, i) => {
      if (i !== activeTabIndex) return tab;
      return { ...tab, items: tab.items.map(
        (item: any) => item.id === id ? { ...item, [field]: value } : item
      )};
    }));
  };

  // Remove um item de avaliação da aba ativa
  const removeItem = (id: string) => {
    setTabs(prev => prev.map((tab, i) => {
      if (i !== activeTabIndex) return tab;
      return { ...tab, items: tab.items.filter((item: any) => item.id !== id) };
    }));
  };

  // Adiciona um item manual (em branco) na aba ativa
  const addNewItem = () => {
    setTabs(prev => prev.map((tab, i) => {
      if (i !== activeTabIndex) return tab;
      return { ...tab, items: [...tab.items, {
        id: crypto.randomUUID(), numero: '', nome: 'Registro Manual',
        fase: '--', grau: '', comentario: '', confidence: 1.0
      }]};
    }));
  };

  // ============================================================================
  // GERENCIAMENTO DE ABAS (fechar, navegar)
  // ============================================================================

  // Fecha (remove) uma aba específica
  const closeTab = (indexToClose: number) => {
    const newTabs = tabs.filter((_, i) => i !== indexToClose);

    if (newTabs.length === 0) {
      // Sem fichas restantes — volta ao estado inicial (tela de upload)
      setTabs([]);
      setStatus('idle');
      setActiveTabIndex(0);
      setDebugMode(false);
      setErrorMsg('');
    } else {
      setTabs(newTabs);
      // Ajusta o índice da aba ativa para não sair do range
      if (activeTabIndex >= newTabs.length) {
        setActiveTabIndex(newTabs.length - 1);
      } else if (activeTabIndex > indexToClose) {
        setActiveTabIndex(activeTabIndex - 1);
      }
      // Se activeTabIndex === indexToClose, a próxima aba assume a mesma posição
    }
  };

  // ============================================================================
  // CONSTRUÇÃO DO PAYLOAD E EXPORTAÇÃO CSV
  // ============================================================================

  // Constrói o payload de envio para UMA ficha específica
  // (um array de objetos, um por item de avaliação)
  const buildPayloadForTab = (tab: FichaTab) => {
    return tab.items.map((item: any) => ({
      data: tab.meta.data,
      esquadrilha: tab.meta.esquadrilha,
      missao: tab.meta.missao,
      grauMissao: (tab.meta.tipoMissao === 'Abortiva' || tab.meta.tipoMissao === 'Extra')
        ? '' : tab.meta.grauMissao,
      aluno1p: tab.meta.aluno1p,
      instrutor: tab.meta.instrutor,
      faseMissao: tab.meta.fase,
      aeronave: tab.meta.aeronave,
      hdep: tab.meta.hdep,
      pousos: tab.meta.pousos,
      tev: tab.meta.tev,
      parecer: tab.meta.parecer,
      numero: item.numero,
      nome: item.nome,
      faseItem: item.fase,
      grau: item.grau,
      comentario: item.comentario,
      tipoMissao: tab.meta.tipoMissao,
    }));
  };

  // Exporta CSV de TODAS as fichas (todas as abas) em um único arquivo
  const exportCSV = () => {
    const hdrs = ['Data', 'Esquadrilha', 'Missão', 'Grau Missão', '1P / AL', 'IN',
      'Fase Missão', 'Anv', 'H.Dep', 'Pousos', 'TEV', 'Parecer',
      'Nº Item', 'Nome', 'Fase Item', 'Grau/Menção', 'Comentário', 'Tipo'];

    // Coleta dados de TODAS as fichas de todas as abas
    const allPayloads = tabs.flatMap(tab => buildPayloadForTab(tab));

    const csvContent = [
      hdrs.join(','),
      ...allPayloads.map(r =>
        `"${r.data}","${r.esquadrilha}","${r.missao}","${r.grauMissao}",` +
        `"${r.aluno1p}","${r.instrutor}","${r.faseMissao}","${r.aeronave}",` +
        `"${r.hdep}","${r.pousos}","${r.tev}",` +
        `"${(r.parecer || '').replace(/"/g, '""')}",` +
        `"${r.numero}","${r.nome}","${r.faseItem}","${r.grau}",` +
        `"${(r.comentario || '').replace(/"/g, '""')}","${r.tipoMissao}"`
      )
    ].join('\n');

    const blob = new Blob(['\ufeff' + csvContent], { type: 'text/csv;charset=utf-8;' });
    const lnk = document.createElement('a');
    lnk.href = URL.createObjectURL(blob);
    lnk.setAttribute('download', `Fichas_${tabs.length}_${new Date().toISOString().slice(0, 10)}.csv`);
    lnk.click();
  };

  // ============================================================================
  // VALIDAÇÃO DE CAMPOS OBRIGATÓRIOS (antes do envio)
  // ============================================================================

  // Percorre TODAS as fichas e encontra a primeira com campos obrigatórios faltando.
  // Campos obrigatórios: Esquadrilha, Aluno (1P), Instrutor (IN).
  // Retorna null se todas estão válidas.
  // >>> EDITE AQUI SE QUISER ADICIONAR MAIS CAMPOS OBRIGATÓRIOS <<<
  const findFirstInvalidTab = (): { index: number; message: string } | null => {
    for (let i = 0; i < tabs.length; i++) {
      const tab = tabs[i];
      if (!tab.meta.esquadrilha) {
        return { index: i, message: `"${tab.fileName}" — Esquadrilha não informada.` };
      }
      if (!tab.meta.aluno1p) {
        return { index: i, message: `"${tab.fileName}" — Trigrama do Aluno (1P) não informado.` };
      }
      if (!tab.meta.instrutor) {
        return { index: i, message: `"${tab.fileName}" — Trigrama do Instrutor (IN) não informado.` };
      }
    }
    return null; // Tudo OK
  };

  // ============================================================================
  // FLUXO DE ENVIO: Validação → Confirmação → Progresso → Concluído
  // ============================================================================

  // Passo 1: Chamado ao clicar "Salvar Registros"
  // Valida os campos e, se tudo OK, mostra popup de confirmação
  const handleSendClick = () => {
    setErrorMsg('');

    // Verifica se todos os campos obrigatórios estão preenchidos
    const invalid = findFirstInvalidTab();
    if (invalid) {
      // Salta para a aba com problema e exibe mensagem de erro
      setActiveTabIndex(invalid.index);
      setErrorMsg(invalid.message);
      window.scrollTo(0, 0);
      return;
    }

    // Todas as fichas estão válidas — mostra popup "Todas as fichas verificadas?"
    setShowConfirmDialog(true);
  };

  // Passo 2: Chamado ao confirmar no popup
  // Envia TODAS as fichas sequencialmente para o Google Sheets
  // (Sequencial para evitar conflito com o LockService do Google Apps Script)
  const sendAllFichas = async () => {
    setShowConfirmDialog(false);
    setShowProgressDialog(true);
    setSendComplete(false);
    setSendProgress(0);

    // Inicializa o status de cada ficha como "aguardando"
    const initialStatuses: SendStatus[] = tabs.map(tab => ({
      fileName: tab.fileName,
      status: 'waiting' as const,
    }));
    setSendStatuses(initialStatuses);

    let successCount = 0;
    let errorCount = 0;

    // Envia uma ficha por vez (sequencial)
    for (let i = 0; i < tabs.length; i++) {
      // Marca a ficha atual como "enviando"
      setSendStatuses(prev => prev.map((s, idx) =>
        idx === i ? { ...s, status: 'sending' } : s
      ));

      try {
        const payload = buildPayloadForTab(tabs[i]);
        const res = await fetch(WEBHOOK_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify(payload),
        });

        if (res.ok) {
          // Sucesso: marca como enviado
          successCount++;
          setSendStatuses(prev => prev.map((s, idx) =>
            idx === i ? { ...s, status: 'success', message: 'Enviado com sucesso' } : s
          ));
        } else {
          throw new Error(`HTTP ${res.status}`);
        }
      } catch (err: any) {
        // Erro: marca como falha
        errorCount++;
        setSendStatuses(prev => prev.map((s, idx) =>
          idx === i ? { ...s, status: 'error', message: err.message || 'Erro desconhecido' } : s
        ));
      }

      // Atualiza a barra de progresso (porcentagem)
      setSendProgress(Math.round(((i + 1) / tabs.length) * 100));
    }

    // Envio concluído — salva o resumo final
    setSendSummary({ success: successCount, errors: errorCount });
    setSendComplete(true);
  };

  // Passo 3: Limpa tudo após o envio e volta à tela inicial
  const resetAfterSuccess = () => {
    setShowProgressDialog(false);
    setSendComplete(false);
    setSendProgress(0);
    setSendStatuses([]);
    setTabs([]);
    setActiveTabIndex(0);
    setStatus('idle');
    setDebugMode(false);
    setErrorMsg('');
  };

  // ============================================================================
  // RENDERIZAÇÃO (JSX)
  // ============================================================================

  return (
    <div className="min-h-screen w-full bg-slate-50 text-slate-800 font-sans p-4 md:p-8 relative">

      {/* ================================================================== */}
      {/* POPUP DE CONFIRMAÇÃO: "Todas as fichas verificadas?"               */}
      {/* Aparece quando o usuário clica "Salvar Registros" e tudo está OK   */}
      {/* ================================================================== */}
      {showConfirmDialog && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/70 backdrop-blur-sm px-4">
          <div className="bg-white rounded-2xl shadow-2xl p-8 max-w-md w-full text-center">
            {/* Ícone decorativo */}
            <div className="w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center mx-auto mb-6">
              <FileText size={32} className="text-blue-600" />
            </div>

            <h2 className="text-2xl font-bold mb-2">Todas as fichas verificadas?</h2>
            <p className="text-slate-500 mb-8">
              {tabs.length} ficha{tabs.length > 1 ? 's' : ''} será{tabs.length > 1 ? 'ão' : ''} enviada{tabs.length > 1 ? 's' : ''} para a base de dados.
            </p>

            {/* Botões: Cancelar / Confirmar e Enviar */}
            <div className="flex gap-3">
              <button
                onClick={() => setShowConfirmDialog(false)}
                className="flex-1 px-6 py-3.5 border border-slate-200 rounded-xl font-semibold text-slate-700 hover:bg-slate-50 transition"
              >
                Cancelar
              </button>
              <button
                onClick={sendAllFichas}
                className="flex-1 px-6 py-3.5 bg-blue-600 text-white rounded-xl font-bold hover:bg-blue-700 transition active:scale-95"
              >
                Confirmar e Enviar
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ================================================================== */}
      {/* POPUP DE PROGRESSO DO ENVIO                                        */}
      {/* Mostra barra de progresso, porcentagem e tasklist de cada ficha    */}
      {/* Quando finaliza, mostra "✅ Concluído"                            */}
      {/* ================================================================== */}
      {showProgressDialog && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/70 backdrop-blur-sm px-4">
          <div className="bg-white rounded-2xl shadow-2xl p-8 max-w-lg w-full">

            {sendComplete ? (
              /* ---------- ENVIO CONCLUÍDO ---------- */
              <>
                <div className="text-center mb-6">
                  {/* Ícone de conclusão: ✅ se tudo OK, ⚠️ se houve erros */}
                  <div className="text-5xl mb-4">{sendSummary.errors === 0 ? '✅' : '⚠️'}</div>
                  <h2 className="text-2xl font-bold">
                    {sendSummary.errors === 0 ? 'Concluído' : 'Concluído com Erros'}
                  </h2>
                  <p className="text-slate-500 mt-2">
                    {sendSummary.errors === 0
                      ? `${sendSummary.success} ficha${sendSummary.success > 1 ? 's' : ''} enviada${sendSummary.success > 1 ? 's' : ''} com sucesso.`
                      : `${sendSummary.success} de ${tabs.length} fichas enviadas. ${sendSummary.errors} erro${sendSummary.errors > 1 ? 's' : ''}.`
                    }
                  </p>
                </div>

                {/* Lista de resultados por ficha */}
                <div className="max-h-60 overflow-y-auto space-y-2 mb-6">
                  {sendStatuses.map((s, i) => (
                    <div key={i} className={`flex items-center gap-3 px-4 py-2.5 rounded-lg text-sm ${
                      s.status === 'success' ? 'bg-green-50' : 'bg-red-50'
                    }`}>
                      {s.status === 'success'
                        ? <CheckCircle2 size={18} className="text-green-500 shrink-0" />
                        : <XCircle size={18} className="text-red-500 shrink-0" />
                      }
                      <span className="font-medium truncate flex-1">{s.fileName}</span>
                      <span className="text-xs text-slate-400 shrink-0">
                        {s.status === 'success' ? 'Enviado ✓' : (s.message || 'Erro')}
                      </span>
                    </div>
                  ))}
                </div>

                {/* Botão para iniciar nova extração (limpa tudo) */}
                <button
                  onClick={resetAfterSuccess}
                  className="w-full bg-slate-800 text-white font-bold py-3.5 px-6 rounded-xl text-lg hover:bg-slate-900 transition"
                >
                  Nova Extração
                </button>
              </>
            ) : (
              /* ---------- ENVIO EM ANDAMENTO ---------- */
              <>
                <div className="text-center mb-6">
                  <RefreshCw size={44} className="animate-spin text-blue-600 mx-auto mb-4" />
                  <h2 className="text-2xl font-bold">Enviando Fichas</h2>
                  <p className="text-slate-500 mt-1">Sincronizando com a base de dados...</p>
                </div>

                {/* Barra de progresso com porcentagem */}
                <div className="mb-2 flex justify-between text-sm font-semibold text-slate-600">
                  <span>
                    {sendStatuses.filter(s => s.status === 'success' || s.status === 'error').length} de {tabs.length}
                  </span>
                  <span>{sendProgress}%</span>
                </div>
                <div className="w-full bg-slate-200 rounded-full h-3 mb-6 overflow-hidden">
                  <div
                    className="bg-blue-600 h-full rounded-full transition-all duration-500 ease-out"
                    style={{ width: `${sendProgress}%` }}
                  />
                </div>

                {/* Tasklist: status individual de cada ficha */}
                <div className="max-h-60 overflow-y-auto space-y-2">
                  {sendStatuses.map((s, i) => (
                    <div key={i} className={`flex items-center gap-3 px-4 py-2.5 rounded-lg text-sm transition-colors ${
                      s.status === 'success' ? 'bg-green-50' :
                      s.status === 'sending' ? 'bg-blue-50' :
                      s.status === 'error' ? 'bg-red-50' : 'bg-slate-50'
                    }`}>
                      {/* Ícone animado por status */}
                      {s.status === 'waiting' && <span className="text-slate-400 shrink-0">⏳</span>}
                      {s.status === 'sending' && <RefreshCw size={16} className="animate-spin text-blue-500 shrink-0" />}
                      {s.status === 'success' && <CheckCircle2 size={16} className="text-green-500 shrink-0" />}
                      {s.status === 'error' && <XCircle size={16} className="text-red-500 shrink-0" />}

                      {/* Nome do arquivo */}
                      <span className="font-medium truncate flex-1">{s.fileName}</span>

                      {/* Texto de status */}
                      <span className="text-xs text-slate-400 shrink-0">
                        {s.status === 'waiting' && 'Aguardando...'}
                        {s.status === 'sending' && 'Enviando...'}
                        {s.status === 'success' && 'Enviado ✓'}
                        {s.status === 'error' && (s.message || 'Erro')}
                      </span>
                    </div>
                  ))}
                </div>
              </>
            )}
          </div>
        </div>
      )}

      {/* ================================================================== */}
      {/* CONTEÚDO PRINCIPAL DA PÁGINA                                       */}
      {/* ================================================================== */}
      <div className={`max-w-7xl mx-auto transition-opacity ${(showConfirmDialog || showProgressDialog) ? 'opacity-25' : 'opacity-100'}`}>

        {/* ----- HEADER DO APP ----- */}
        <header className="mb-8 bg-white p-6 rounded-xl shadow-sm border flex items-center justify-between gap-5">
          <div className="flex items-center gap-5">
            <div className="w-14 h-14 bg-blue-600 rounded-lg flex items-center justify-center text-white shrink-0">
              <FileText size={32} />
            </div>
            <div>
              <h1 className="text-3xl font-bold tracking-tight">Análise Estruturada EIA</h1>
              <p className="text-slate-500 text-lg">Processamento de Fichas de Avaliação.</p>
            </div>
          </div>

          {/* Botão de debug (só aparece no modo de revisão) */}
          {status === 'reviewing' && (
            <button
              onClick={() => setDebugMode(!debugMode)}
              className={`p-3 rounded-full transition ${debugMode ? 'bg-red-100 text-red-600' : 'bg-slate-100 text-slate-400 hover:bg-slate-200'}`}
              title="Módulo de Depuração Visual"
            >
              <Bug size={24} />
            </button>
          )}
        </header>

        {/* ----- MENSAGEM DE ERRO (aparece no topo quando há campos faltando) ----- */}
        {errorMsg && (
          <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-lg text-red-700 font-medium flex gap-2">
            <AlertCircle className="text-red-500 shrink-0" />{errorMsg}
          </div>
        )}

        {/* ===== TELA INICIAL — SELEÇÃO DE PDFs ===== */}
        {status === 'idle' && (
          <div className="flex justify-center">
            <div className="bg-white p-16 rounded-2xl shadow-sm border flex flex-col items-center text-center w-full max-w-3xl border-slate-200 hover:border-blue-300 transition hover:shadow-lg">
              <Upload size={48} className="text-blue-600 mb-7" />
              <h2 className="text-2xl font-bold mb-4">Seleção de Documentos</h2>
              <p className="text-slate-500 mb-8 max-w-md">
                Selecione uma ou múltiplas fichas em formato PDF.
                Cada ficha será processada e exibida em uma aba separada para revisão.
              </p>
              {/* Botão de upload — aceita múltiplos arquivos PDF */}
              <label className="bg-blue-600 hover:bg-blue-700 text-white px-9 py-4.5 rounded-2xl text-xl font-bold cursor-pointer flex items-center gap-3.5 transition active:scale-95">
                <FileText size={26} /> Carregar PDFs
                <input
                  type="file"
                  accept="application/pdf"
                  multiple
                  className="hidden"
                  onClick={(e: any) => e.target.value = ''}
                  onChange={handleFileUpload}
                />
              </label>
            </div>
          </div>
        )}

        {/* ===== TELA DE CARREGAMENTO (processando os PDFs) ===== */}
        {status === 'loading' && (
          <div className="bg-white p-20 rounded-2xl shadow-sm border flex flex-col items-center text-center">
            <RefreshCw className="text-blue-600 animate-spin mb-5" size={44} />
            <h2 className="text-2xl font-bold">Extração em Andamento</h2>
            <p className="text-slate-500 mt-2">
              {loadingProgress.total > 1
                ? `Processando ficha ${loadingProgress.current} de ${loadingProgress.total}...`
                : 'Mapeamento vetorial e reconciliação de dados ativados...'
              }
            </p>
            {/* Barra de progresso do carregamento (só aparece com 2+ arquivos) */}
            {loadingProgress.total > 1 && (
              <div className="w-full max-w-sm mt-6">
                <div className="w-full bg-slate-200 rounded-full h-2.5 overflow-hidden">
                  <div
                    className="bg-blue-600 h-full rounded-full transition-all duration-300"
                    style={{ width: `${Math.round((loadingProgress.current / loadingProgress.total) * 100)}%` }}
                  />
                </div>
              </div>
            )}
          </div>
        )}

        {/* ===== TELA DE REVISÃO — ABAS + CONTEÚDO DA FICHA ATIVA ===== */}
        {status === 'reviewing' && activeTab && (
          <div className="space-y-0">

            {/* ---------------------------------------------------------- */}
            {/* BARRA DE ABAS (TAB BAR)                                    */}
            {/* Cada aba = um PDF carregado. Clique para navegar entre elas */}
            {/* ---------------------------------------------------------- */}
            <div className="flex items-center bg-white border border-slate-200 rounded-t-xl overflow-hidden">
              {/* Container scrollável com as abas */}
              <div className="flex-1 flex overflow-x-auto">
                {tabs.map((tab, i) => (
                  <button
                    key={tab.id}
                    onClick={() => { setActiveTabIndex(i); setErrorMsg(''); }}
                    className={`group flex items-center gap-2 px-5 py-3.5 text-sm font-semibold whitespace-nowrap border-b-2 transition-all shrink-0 ${
                      i === activeTabIndex
                        ? 'border-blue-600 bg-blue-50/60 text-blue-700'
                        : 'border-transparent text-slate-500 hover:text-slate-700 hover:bg-slate-50'
                    }`}
                  >
                    <FileText size={14} className={i === activeTabIndex ? 'text-blue-500' : 'text-slate-400'} />
                    {/* Nome do arquivo sem a extensão .pdf */}
                    <span className="max-w-[180px] truncate">
                      {tab.fileName.replace(/\.pdf$/i, '')}
                    </span>
                    {/* Botão X para fechar a aba (aparece ao passar o mouse) */}
                    <span
                      onClick={(e) => { e.stopPropagation(); closeTab(i); }}
                      className="ml-1 p-0.5 rounded-full opacity-0 group-hover:opacity-100 hover:bg-red-100 hover:text-red-500 transition cursor-pointer"
                    >
                      <X size={14} />
                    </span>
                  </button>
                ))}
              </div>

              {/* Botão "+" para adicionar mais fichas (abre o seletor de arquivos) */}
              <label
                className="px-4 py-3.5 border-l border-slate-200 cursor-pointer text-slate-400 hover:text-blue-600 hover:bg-blue-50 transition flex items-center gap-1.5 text-sm font-semibold shrink-0"
                title="Adicionar mais fichas"
              >
                <Plus size={16} />
                <span className="hidden md:inline">Adicionar</span>
                <input
                  type="file"
                  accept="application/pdf"
                  multiple
                  className="hidden"
                  onClick={(e: any) => e.target.value = ''}
                  onChange={handleFileUpload}
                />
              </label>
            </div>

            {/* Indicador textual: "Ficha 2 de 5 • Ficha_MTA_PS-02.pdf" */}
            <div className="bg-white border-x border-slate-200 px-5 py-2 text-xs text-slate-400 font-semibold uppercase tracking-wider">
              Ficha {activeTabIndex + 1} de {tabs.length} • {activeTab.fileName}
            </div>

            {/* ----- VISUALIZADOR DEBUG (se ativado via botão no header) ----- */}
            {debugMode && (
              <div className="bg-white border-x border-slate-200 px-6 pt-4 pb-6">
                <h3 className="text-lg font-bold mb-3 flex items-center gap-2">
                  <Bug size={20} /> Visualização de OCR
                </h3>
                <PdfViewer pdf={activeTab.pdfDocument} tokens={activeTab.rawPdfTokens} />
              </div>
            )}

            {/* ---------------------------------------------------------- */}
            {/* METADADOS DO CABEÇALHO (da aba/ficha ativa)                */}
            {/* ---------------------------------------------------------- */}
            <div className="bg-white border-x border-slate-200 p-7">
              <h3 className="text-xl font-bold mb-5 flex items-center gap-2.5 border-b pb-4 border-slate-100">
                <FileText className="text-blue-500" size={22} /> Cabeçalho da Missão
              </h3>

              <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-5">
                {/* >>> Campos com * são obrigatórios para o envio <<< */}
                <MetaSelect label="Esquadrilha *" value={activeTab.meta.esquadrilha}
                  onChange={(e: any) => updateMeta('esquadrilha', e.target.value)}
                  options={['Antares', 'Vega', 'Castor', 'Sirius']} />
                <MetaInput label="1P / Aluno *" value={activeTab.meta.aluno1p}
                  onChange={(e: any) => updateMeta('aluno1p', e.target.value)} maxLength={3} placeholder="MTA" />
                <MetaInput label="Instrutor (IN) *" value={activeTab.meta.instrutor}
                  onChange={(e: any) => updateMeta('instrutor', e.target.value)} maxLength={3} placeholder="MOT" />
                <MetaInput label="Missão" value={activeTab.meta.missao}
                  onChange={(e: any) => updateMeta('missao', e.target.value)} />
                <MetaInput label="Grau Missão" value={activeTab.meta.grauMissao}
                  onChange={(e: any) => updateMeta('grauMissao', e.target.value)}
                  disabled={activeTab.meta.tipoMissao === 'Abortiva' || activeTab.meta.tipoMissao === 'Extra'} />
                <MetaSelect label="Tipo Missão" value={activeTab.meta.tipoMissao}
                  onChange={(e: any) => updateMeta('tipoMissao', e.target.value)}
                  options={['Normal', 'Abortiva', 'Extra', 'Revisão']} />
                <MetaInput label="Data" value={activeTab.meta.data}
                  onChange={(e: any) => updateMeta('data', e.target.value)} />
                <MetaInput label="Fase" value={activeTab.meta.fase}
                  onChange={(e: any) => updateMeta('fase', e.target.value)} />
                <MetaInput label="Aeronave" value={activeTab.meta.aeronave}
                  onChange={(e: any) => updateMeta('aeronave', e.target.value)} />
                <MetaInput label="H. Dep" value={activeTab.meta.hdep}
                  onChange={(e: any) => updateMeta('hdep', e.target.value)} />
                <MetaInput label="Pousos" value={activeTab.meta.pousos}
                  onChange={(e: any) => updateMeta('pousos', e.target.value)} />
                <MetaInput label="TEV" value={activeTab.meta.tev}
                  onChange={(e: any) => updateMeta('tev', e.target.value)} />
              </div>

              {/* Parecer do Comandante (campo de texto livre) */}
              <div className="mt-5 pt-5 border-t border-slate-100">
                <MetaTextarea label="Parecer do Comandante" value={activeTab.meta.parecer}
                  onChange={(e: any) => updateMeta('parecer', e.target.value)}
                  placeholder="Registro de observações operacionais..." />
              </div>
            </div>

            {/* ---------------------------------------------------------- */}
            {/* TABELA DE AVALIAÇÃO (itens da aba/ficha ativa)             */}
            {/* ---------------------------------------------------------- */}
            <div className="bg-white rounded-b-2xl shadow-sm border border-slate-200 border-t-0 overflow-hidden">
              {/* Cabeçalho da tabela com botões de ação */}
              <div className="p-6 border-b flex flex-col lg:flex-row justify-between gap-4 bg-slate-50/50">
                <h2 className="text-xl font-bold flex items-center gap-2.5">
                  <Check className="text-green-500" size={26} />
                  Tabela de Avaliação ({activeTab.items.length} itens)
                </h2>
                <div className="flex flex-wrap gap-3">
                  {/* Descartar TODAS as fichas e voltar à tela inicial */}
                  <button
                    onClick={() => { setTabs([]); setStatus('idle'); setActiveTabIndex(0); setDebugMode(false); setErrorMsg(''); }}
                    className="px-5 py-3 text-sm font-semibold border bg-white hover:bg-slate-50 rounded-xl"
                  >
                    Descartar Tudo
                  </button>

                  {/* Exportar CSV de TODAS as fichas (todas as abas) */}
                  <button
                    onClick={exportCSV}
                    className="px-5 py-3 text-sm font-semibold border bg-white hover:bg-slate-50 rounded-xl flex items-center gap-2"
                  >
                    <Download size={18} /> Exportar CSV ({tabs.length})
                  </button>

                  {/* Enviar TODAS as fichas para o Google Sheets */}
                  <button
                    onClick={handleSendClick}
                    className="px-9 py-3 text-lg text-white bg-blue-600 hover:bg-blue-700 rounded-xl font-bold flex items-center gap-2.5 transition active:scale-95"
                  >
                    <Zap size={22} /> Salvar Registros ({tabs.length})
                  </button>
                </div>
              </div>

              {/* Corpo da tabela com os itens de avaliação */}
              <div className="p-7 overflow-x-auto">
                <table className="w-full text-left border-collapse min-w-[950px] relative z-20">
                  <thead>
                    <tr className="bg-slate-50 border-b border-slate-200 text-xs text-slate-500 uppercase font-bold tracking-wider">
                      <th className="p-3 w-8 text-center" title="Acurácia da Extração">C.</th>
                      <th className="p-3 w-16">Item</th>
                      <th className="p-3">Manobra / Competência</th>
                      <th className="p-3 w-20 text-center">Fase</th>
                      <th className="p-3 w-24 text-center">Grau</th>
                      <th className="p-3">Observações Mapeadas</th>
                      <th className="p-3 w-12 text-center">Ações</th>
                    </tr>
                  </thead>
                  <tbody className="text-sm">
                    {activeTab.items.map((it: any) => (
                      <tr key={it.id} className="border-b border-slate-100 hover:bg-slate-50/40 text-slate-800">
                        {/* Indicador visual de confiança (verde/amarelo/vermelho) */}
                        <td className="p-3 text-center">
                          <div
                            className={`w-3 h-3 rounded-full mx-auto ${
                              it.confidence > 0.6 ? 'bg-green-400' :
                              (it.confidence > 0.4 ? 'bg-yellow-400' : 'bg-red-500 animate-pulse')
                            }`}
                            title={`Acurácia calculada: ${Math.round(it.confidence * 100)}%`}
                          />
                        </td>
                        {/* Número do item */}
                        <td className="p-3 font-mono font-bold text-slate-400">
                          <input value={it.numero} onChange={(e) => updateItem(it.id, 'numero', e.target.value)}
                            className="w-full bg-transparent px-1" placeholder="--" />
                        </td>
                        {/* Nome da manobra/competência */}
                        <td className="p-3 font-medium">
                          <input value={it.nome} onChange={(e) => updateItem(it.id, 'nome', e.target.value)}
                            className={`w-full bg-transparent px-1 ${getGradeColorClass(it.grau)}`}
                            placeholder="Item de avaliação..." />
                        </td>
                        {/* Fase do item (PR, RC, RM, RO) */}
                        <td className="p-3 text-center font-bold text-blue-700">
                          <input value={it.fase} onChange={(e) => updateItem(it.id, 'fase', e.target.value)}
                            className="w-full bg-transparent px-1 text-center" placeholder="--" />
                        </td>
                        {/* Grau do item (1-6 ou menção textual) */}
                        <td className="p-3 text-center font-extrabold">
                          <input value={it.grau} onChange={(e) => updateItem(it.id, 'grau', e.target.value)}
                            className={`w-full bg-transparent px-1 uppercase text-center ${getGradeColorClass(it.grau)}`}
                            placeholder="--" />
                        </td>
                        {/* Comentário/observação do item */}
                        <td className="p-3">
                          <textarea value={it.comentario} onChange={(e) => updateItem(it.id, 'comentario', e.target.value)}
                            className="w-full bg-transparent border-slate-200 focus:border-blue-300 focus:bg-white text-slate-700 border px-2.5 py-1.5 rounded-lg resize-y min-h-[44px] text-xs leading-relaxed"
                            placeholder="Adicionar dados..." />
                        </td>
                        {/* Botão para excluir o item */}
                        <td className="p-3 text-center">
                          <button onClick={() => removeItem(it.id)}
                            className="p-2 hover:bg-red-50 hover:text-red-500 text-slate-300 rounded-lg">
                            <Trash2 size={16} />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>

                {/* Botão para inserir item manual na tabela */}
                <div className="mt-5 flex justify-center border-t border-slate-100 pt-5 relative z-20">
                  <button onClick={addNewItem}
                    className="flex items-center gap-2 font-semibold text-blue-600 hover:text-blue-800 px-5 py-2.5 bg-blue-50/50 rounded-lg active:scale-95 transition">
                    <Plus size={18} /> Inserir Linha Manual
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
