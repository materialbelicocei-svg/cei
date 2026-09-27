/**
 * RESERVA DE ARMAMENTOS CEI - 100% JAVASCRIPT / TYPESCRIPT INTEGRADO
 * Sincronização direta com a API do Google Sheets v4 e modo offline com cache local.
 */
import './style.css';
import {
  initGoogleAuth,
  signInWithGoogle,
  logoutGoogle,
  getGoogleAccessToken,
  getActiveSpreadsheetId,
  setActiveSpreadsheetId,
  DEFAULT_SHEET_ID,
  fetchSpreadsheetMetadata,
  readSheetRows,
  insertRowInSheet,
  updateCellWithIdCheck,
  deleteRowWithIdCheck,
  createSheetTab,
  renameSheetTab,
  deleteSheetTab,
  generateRowId,
  listSpreadsheetsFromDrive
} from './sheetsService.ts';
import initialBookData from './initialBookData.json';
import ExcelJS from 'exceljs/dist/exceljs.min.js';

// ============================================================================
// CONFIGURAÇÃO DO GOOGLE SHEETS E SIMULADOR DO GOOGLE.SCRIPT.RUN
// ============================================================================
const SHEET_ID_FIXO = getActiveSpreadsheetId();
const COL_ASS = [6, 7, 9, 10];
let currentUserGoogle = null;
let currentGoogleToken = null;

function isAbaOculta(nome) {
  if (!nome) return true;
  const u = String(nome).toUpperCase().trim().replace(/[\s_-]+/g, '');
  return (
    u === 'LOG' ||
    u === 'LOGS' ||
    u === 'APAGADOS' ||
    u === 'LOGAPAGADOS' ||
    u === 'LOGSAPAGADOS' ||
    u === 'LOGAPAGADO' ||
    u === 'LOGSAPAGADO' ||
    u === 'LIXEIRA' ||
    u === 'LIXO' ||
    u.startsWith('LOG') ||
    u.includes('APAGADO')
  );
}

function getStoredAbas() {
  try {
    const raw = localStorage.getItem('CEI_ABAS');
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length > 0) {
        const filtradas = parsed.filter(a => !isAbaOculta(a));
        if (filtradas.length > 0) return filtradas;
      }
    }
  } catch (e) {}

  // Carrega abas do livro de cautelas XLSX se disponível
  const bookSheets = Object.keys(initialBookData || {}).filter(a => !isAbaOculta(a));
  const padrao = bookSheets.length > 0 ? bookSheets : ['CAUTELAS', 'GERAL', 'SERVIÇO', 'INSTRUÇÃO'];
  localStorage.setItem('CEI_ABAS', JSON.stringify(padrao));
  return padrao;
}

function setStoredAbas(abas) {
  localStorage.setItem('CEI_ABAS', JSON.stringify(abas));
}

// ============================================================================
// REPOSITÓRIO RESILIENTE DE ASSINATURAS EM MEMÓRIA
// ============================================================================
const assinaturasMemoriaMap = new Map();

function registrarAssinaturaMemoria(id, nome, data, nomeAba, campo, base64) {
  if (!base64 || base64.length < 30) return;
  const aba = nomeAba || abaAtual || 'CAUTELAS';
  if (id) assinaturasMemoriaMap.set(`${id}_${campo}`, base64);
  if (nome && data) {
    const key = `${aba}_${nome.trim().toUpperCase()}_${data.trim()}_${campo}`;
    assinaturasMemoriaMap.set(key, base64);
  }
  if (nome) {
    const keyNome = `${aba}_${nome.trim().toUpperCase()}_${campo}`;
    assinaturasMemoriaMap.set(keyNome, base64);
  }
}

// Inicializa o mapa com as assinaturas do livro de cautelas
if (initialBookData) {
  Object.keys(initialBookData).forEach(sheetName => {
    (initialBookData[sheetName] || []).forEach((rowItem) => {
      ['f', 'g', 'i', 'j'].forEach(c => {
        if (rowItem[c] && rowItem[c].length > 30) {
          registrarAssinaturaMemoria(rowItem.id, rowItem.nome, rowItem.data, sheetName, c, rowItem[c]);
        }
      });
    });
  });
}

function mesclarComAssinaturasDoLivro(lista, nomeAba) {
  const targetAba = nomeAba || abaAtual || 'CAUTELAS';
  return (lista || []).map((item) => {
    const clone = Object.assign({}, item);

    ['f', 'g', 'i', 'j'].forEach(c => {
      if (!clone[c] || clone[c].length < 30) {
        // 1. Busca por ID exclusivo
        let sig = clone.id ? assinaturasMemoriaMap.get(`${clone.id}_${c}`) : null;
        // 2. Busca por Nome + Data
        if (!sig && clone.nome && clone.data) {
          sig = assinaturasMemoriaMap.get(`${targetAba}_${clone.nome.trim().toUpperCase()}_${clone.data.trim()}_${c}`);
        }
        // 3. Busca por Nome
        if (!sig && clone.nome) {
          sig = assinaturasMemoriaMap.get(`${targetAba}_${clone.nome.trim().toUpperCase()}_${c}`);
        }

        if (sig && sig.length > 30) {
          clone[c] = sig;
        }
      } else {
        registrarAssinaturaMemoria(clone.id, clone.nome, clone.data, targetAba, c, clone[c]);
      }
    });

    return clone;
  });
}

function getStoredAbaDados(nomeAba) {
  const chave = 'cei_' + (nomeAba || 'CAUTELAS');
  let dados = [];
  try {
    const raw = localStorage.getItem(chave);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length > 0) {
        dados = parsed;
      }
    }
  } catch (e) {}

  if (!dados || dados.length === 0) {
    if (initialBookData && initialBookData[nomeAba] && initialBookData[nomeAba].length > 0) {
      dados = initialBookData[nomeAba].map((item, idx) => ({
        ...item,
        row: idx + 2
      }));
    } else if (nomeAba === 'CAUTELAS' || nomeAba === 'GERAL') {
      if (initialBookData && initialBookData['CAUTELAS'] && initialBookData['CAUTELAS'].length > 0) {
        dados = initialBookData['CAUTELAS'].map((item, idx) => ({
          ...item,
          row: idx + 2
        }));
      }
    }
  }

  // Garante que todas as assinaturas do livro de cautelas sejam sempre restauradas
  dados = mesclarComAssinaturasDoLivro(dados, nomeAba);
  return dados;
}

function setStoredAbaDados(nomeAba, dados) {
  const chave = 'cei_' + (nomeAba || 'GERAL');
  try {
    // Tenta salvar no localStorage versão otimizada (sem duplicar bytes de imagens já presentes em memória)
    const leve = (dados || []).map(d => {
      const item = Object.assign({}, d);
      // Se tiver imagem base64 pesada (> 1000 caracteres), remove do localStorage para não estourar a quota de 5MB
      if (item.f && item.f.length > 500) item.f = '';
      if (item.g && item.g.length > 500) item.g = '';
      if (item.i && item.i.length > 500) item.i = '';
      if (item.j && item.j.length > 500) item.j = '';
      return item;
    });
    localStorage.setItem(chave, JSON.stringify(leve));
  } catch (e) {
    console.warn('LocalStorage quota atingida, dados mantidos em memória RAM com segurança.');
  }
}

// Hash SHA-256 criptográfico de via única (a senha real nunca fica visível no código nem ao inspecionar elemento)
const HASH_SENHA_PADRAO = 'aafaaef9f8fc3f8c0bec772926ed1d3cc6a8a00e60289db9aee94374e572cdaa';

async function gerarHashSenha(texto) {
  if (!texto) return '';
  const encoder = new TextEncoder();
  const data = encoder.encode(String(texto).trim());
  const hashBuffer = await crypto.subtle.digest('SHA-256', data);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
}

// Objeto simulador do google.script.run
window.google = window.google || {};
window.google.script = window.google.script || {};
window.google.script.run = {
  _success: null,
  _failure: null,
  withSuccessHandler: function(fn) {
    const clone = Object.assign({}, this);
    clone._success = fn;
    return clone;
  },
  withFailureHandler: function(fn) {
    const clone = Object.assign({}, this);
    clone._failure = fn;
    return clone;
  },
  verificarSenhaServidor: async function(s) {
    try {
      if (!s) { if (this._success) this._success(false); return; }
      const hashDigitado = await gerarHashSenha(s);
      const hashSalvo = localStorage.getItem('HASH_SENHA_SISTEMA') || HASH_SENHA_PADRAO;
      const ok = (hashDigitado === hashSalvo);
      if (this._success) this._success(ok);
    } catch (e) {
      if (this._failure) this._failure(e);
    }
  },
  getAbas: async function() {
    try {
      const token = await getGoogleAccessToken();
      const sheetId = getActiveSpreadsheetId();
      if (token && sheetId) {
        try {
          const meta = await fetchSpreadsheetMetadata(sheetId, token);
          if (meta.sheets && meta.sheets.length > 0) {
            // Oculta abas de logs/apagados/lixeira do menu visível
            const filtradas = meta.sheets.filter(s => !isAbaOculta(s));
            const finais = filtradas.length > 0 ? filtradas : ['GERAL'];
            setStoredAbas(finais);
            if (this._success) this._success(finais);
            return;
          }
        } catch (e) {
          console.warn('Erro ao buscar abas do Sheets, usando local:', e);
        }
      }
      const abas = getStoredAbas().filter(s => !isAbaOculta(s));
      if (this._success) this._success(abas.length > 0 ? abas : ['GERAL']);
    } catch (e) {
      if (this._failure) this._failure(e);
    }
  },
  criarAba: async function(n) {
    try {
      const token = await getGoogleAccessToken();
      const sheetId = getActiveSpreadsheetId();
      if (token && sheetId) {
        try {
          await createSheetTab(sheetId, n, token);
        } catch (e) {
          console.warn('Erro ao criar aba no Sheets:', e);
        }
      }
      const abas = getStoredAbas();
      if (!abas.includes(n)) {
        abas.push(n);
        setStoredAbas(abas);
        setStoredAbaDados(n, []);
      }
      if (this._success) this._success(n);
    } catch (e) {
      if (this._failure) this._failure(e);
    }
  },
  renomearAba: async function(a, b) {
    try {
      const token = await getGoogleAccessToken();
      const sheetId = getActiveSpreadsheetId();
      if (token && sheetId) {
        try {
          await renameSheetTab(sheetId, a, b, token);
        } catch (e) {
          console.warn('Erro ao renomear aba no Sheets:', e);
        }
      }
      const abas = getStoredAbas();
      const idx = abas.indexOf(a);
      if (idx > -1) {
        abas[idx] = b;
        setStoredAbas(abas);
        const dados = getStoredAbaDados(a);
        setStoredAbaDados(b, dados);
        localStorage.removeItem('cei_' + a);
      }
      if (this._success) this._success(b);
    } catch (e) {
      if (this._failure) this._failure(e);
    }
  },
  deletarAba: async function(n) {
    try {
      const abas = getStoredAbas();
      if (abas.length <= 1) {
        throw new Error('⛔ Não pode apagar a única aba');
      }
      const token = await getGoogleAccessToken();
      const sheetId = getActiveSpreadsheetId();
      if (token && sheetId) {
        try {
          await deleteSheetTab(sheetId, n, token);
        } catch (e) {
          console.warn('Erro ao apagar aba no Sheets:', e);
        }
      }
      const novas = abas.filter(x => x !== n);
      setStoredAbas(novas);
      localStorage.removeItem('cei_' + n);
      if (this._success) this._success(novas);
    } catch (e) {
      if (this._failure) this._failure(e);
    }
  },
  getPlanilhaCompleta: async function(nomeAba, mesFiltro) {
    try {
      if (mesFiltro === undefined || mesFiltro === null) mesFiltro = new Date().getMonth();
      const token = await getGoogleAccessToken();
      const sheetId = getActiveSpreadsheetId();
      let dados = [];

      if (!token) {
        // Alerta amigável no status bar
        status('ℹ Conecte o Google no topo para sincronizar em tempo real com a planilha');
        dados = getStoredAbaDados(nomeAba);
      } else if (token && sheetId) {
        try {
          const rowsFromSheet = await readSheetRows(sheetId, nomeAba || 'GERAL', token);
          if (rowsFromSheet && rowsFromSheet.length > 0) {
            // Mescla as linhas da planilha com as assinaturas gráficas do banco
            dados = mesclarComAssinaturasDoLivro(rowsFromSheet, nomeAba);
            setStoredAbaDados(nomeAba, dados);
          } else {
            dados = getStoredAbaDados(nomeAba);
          }
        } catch (err) {
          console.error('Erro na leitura da planilha do Google:', err);
          status('❌ ' + (err.message || 'Erro ao sincronizar com Google Sheets'));
          dados = getStoredAbaDados(nomeAba);
        }
      } else {
        dados = getStoredAbaDados(nomeAba);
      }

      if (mesFiltro !== -1) {
        dados = dados.filter(l => {
          const ds = l.data;
          let m = -1;
          if (typeof ds === 'string') {
            const p = ds.split('/');
            if (p.length === 3) m = parseInt(p[1], 10) - 1;
          }
          return m === -1 || m === mesFiltro;
        });
      }
      if (this._success) this._success(dados);
    } catch (e) {
      if (this._failure) this._failure(e);
    }
  },
  criarNovaLinhaComAss: async function(nomeAba, dataRet, nome, armas, equip, mun, obs, dataDev, assF, assG, assI, assJ, customId) {
    try {
      const dFormatada = dataRet ? (dataRet.includes('/') ? dataRet : dataRet.split('-').reverse().join('/')) : new Date().toLocaleDateString('pt-BR');
      const devFormatada = dataDev ? (dataDev.includes('/') ? dataDev : dataDev.split('-').reverse().join('/')) : '';
      const munStr = String(mun || '').replace(/[^0-9]/g, '');
      const uniqueId = customId || generateRowId();

      const token = await getGoogleAccessToken();
      const sheetId = getActiveSpreadsheetId();
      let realRow = 2;

      if (token && sheetId) {
        try {
          const res = await insertRowInSheet(
            sheetId,
            nomeAba || 'GERAL',
            {
              id: uniqueId,
              data: dFormatada,
              nome: nome || '',
              arma: armas || '',
              equip: equip || '',
              mun: munStr,
              dataDev: devFormatada,
              assF: assF || '',
              assG: assG || '',
              assI: assI || '',
              assJ: assJ || '',
              obs: obs || ''
            },
            token
          );
          realRow = res.row;
        } catch (err) {
          throw new Error('Não foi possível gravar o novo registro no Google Sheets: ' + (err.message || err));
        }
      }

      const dados = getStoredAbaDados(nomeAba);
      dados.forEach(r => { if (typeof r.row === 'number') r.row++; });

      const novaLinha = {
        id: uniqueId,
        row: realRow,
        data: dFormatada,
        nome: nome,
        arma: armas || '',
        equip: equip || '',
        mun: munStr,
        f: assF || '',
        g: assG || '',
        dataDev: devFormatada,
        i: assI || '',
        j: assJ || '',
        obs: obs || ''
      };

      dados.unshift(novaLinha);
      setStoredAbaDados(nomeAba, dados);
      if (this._success) this._success({ ok: true, row: realRow, id: uniqueId });
    } catch (e) {
      if (this._failure) this._failure(e);
    }
  },
  salvarNaCelula: async function(base64, row, col, nomeAba, nomeMilitar, targetId) {
    try {
      if (!base64 || base64.length < 100) {
        if (this._success) this._success({ ok: false, skip: true, msg: 'Assinatura vazia' });
        return;
      }
      const token = await getGoogleAccessToken();
      const sheetId = getActiveSpreadsheetId();
      if (token && sheetId && row > 1) {
        try {
          // Atualiza célula com trava de segurança por ID único!
          await updateCellWithIdCheck(sheetId, nomeAba || 'GERAL', targetId, row, col, base64, token);
        } catch (err) {
          throw new Error('Não foi possível sincronizar a assinatura: ' + (err.message || err));
        }
      }
      const dados = getStoredAbaDados(nomeAba);
      const idx = dados.findIndex(x => targetId ? x.id === targetId : x.row === row);
      if (idx > -1) {
        const campo = col === 6 ? 'f' : col === 7 ? 'g' : col === 9 ? 'i' : 'j';
        dados[idx][campo] = base64;
        registrarAssinaturaMemoria(targetId || dados[idx].id, dados[idx].nome, dados[idx].data, nomeAba, campo, base64);
        setStoredAbaDados(nomeAba, dados);
      }
      if (this._success) this._success({ ok: true, msg: '✅ Assinatura salva com sucesso' });
    } catch (e) {
      if (this._failure) this._failure(e);
    }
  },
  salvarTextoInline: async function(nomeAba, row, col, val, targetId) {
    try {
      let novo = val;
      if (val && val.toString().trim() !== '' && !val.toString().includes('(editado)')) {
        novo = val + ' (editado)';
      }
      const token = await getGoogleAccessToken();
      const sheetId = getActiveSpreadsheetId();
      if (token && sheetId && row > 1) {
        try {
          await updateCellWithIdCheck(sheetId, nomeAba || 'GERAL', targetId, row, col, novo, token);
        } catch (err) {
          throw new Error('Não foi possível atualizar o registro no Google Sheets: ' + (err.message || err));
        }
      }
      const dados = getStoredAbaDados(nomeAba);
      const idx = dados.findIndex(x => targetId ? x.id === targetId : x.row === row);
      if (idx > -1) {
        if (col === 1) {
          dados[idx].data = novo.replace(' (editado)', '').split('-').reverse().join('/');
        } else if (col === 2) {
          dados[idx].nome = novo;
        } else if (col === 3) {
          dados[idx].arma = novo;
        } else if (col === 4) {
          dados[idx].equip = novo;
        } else if (col === 5) {
          dados[idx].mun = novo;
        } else if (col === 8) {
          dados[idx].dataDev = novo.replace(' (editado)', '').split('-').reverse().join('/');
        } else if (col === 11) {
          dados[idx].obs = novo;
        }
        setStoredAbaDados(nomeAba, dados);
      }
      if (this._success) this._success({ ok: true });
    } catch (e) {
      if (this._failure) this._failure(e);
    }
  },
  apagarLinha: async function(nomeAba, row, targetId) {
    try {
      const token = await getGoogleAccessToken();
      const sheetId = getActiveSpreadsheetId();
      if (token && sheetId && row > 1) {
        try {
          await deleteRowWithIdCheck(sheetId, nomeAba || 'GERAL', targetId, row, token);
        } catch (err) {
          throw new Error('Não foi possível apagar o registro no Google Sheets: ' + (err.message || err));
        }
      }
      const dados = getStoredAbaDados(nomeAba);
      const idx = dados.findIndex(x => (targetId && x.id === targetId) || x.row === row);
      if (idx > -1) {
        dados.splice(idx, 1);
        setStoredAbaDados(nomeAba, dados);
      }
      if (this._success) this._success(true);
    } catch (e) {
      if (this._failure) this._failure(e);
    }
  },
  criarBackupPlanilha: function() {
    setTimeout(() => {
      try {
        const abas = getStoredAbas();
        const backup = { data: new Date().toISOString(), abas: {} };
        abas.forEach(a => { backup.abas[a] = getStoredAbaDados(a); });

        const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        const d = new Date().toISOString().replace(/[:.]/g, '-');
        const nomeArquivo = 'BACKUP_RESERVA_CEI_' + d + '.json';
        link.href = url;
        link.download = nomeArquivo;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);

        if (this._success) this._success('✅ Backup: ' + nomeArquivo);
      } catch (e) {
        if (this._failure) this._failure(e);
      }
    }, 60);
  },
  criarBackupSilencioso: function() {
    try {
      const abas = getStoredAbas();
      const backup = { timestamp: Date.now(), dataHora: new Date().toLocaleString('pt-BR'), abas: {} };
      abas.forEach(a => { backup.abas[a] = getStoredAbaDados(a); });
      localStorage.setItem('CEI_AUTO_BACKUP', JSON.stringify(backup));
      localStorage.setItem('CEI_LAST_BACKUP_TIME', String(Date.now()));
      return true;
    } catch (e) {
      return false;
    }
  },
  getUrlExportacao: async function(nomeAba, formato) {
    try {
      const dados = getStoredAbaDados(nomeAba);

      async function toDataUrlFromSource(src) {
        if (!src) return '';
        const s = String(src).trim();
        const looksLikeRawBase64 = /^[A-Za-z0-9+/=\r\n]+$/.test(s) && s.length > 200 && !s.includes('http') && !s.includes('=IMAGE');

        if (s.startsWith('data:image/')) return s;
        if (looksLikeRawBase64) return `data:image/png;base64,${s}`;
        if (s.startsWith('http://') || s.startsWith('https://')) {
          try {
            const res = await fetch(s);
            const blob = await res.blob();
            return await new Promise((resolve) => {
              const reader = new FileReader();
              reader.onloadend = () => resolve(reader.result || '');
              reader.readAsDataURL(blob);
            });
          } catch (e) {
            console.warn('Falha ao converter URL de assinatura em dataURL:', e);
            return '';
          }
        }

        if (s.startsWith('=IMAGE(') || s.startsWith('=image(')) {
          const m = s.match(/"([^"]+)"/) || s.match(/'([^']+)'/);
          if (m && m[1]) return toDataUrlFromSource(m[1]);
          return '';
        }

        if (s.length > 50 && !s.includes(' ') && !s.includes('http') && !s.includes('=')) {
          return 'data:image/png;base64,' + s;
        }

        return '';
      }

      async function obterBase64ParaExportacao(r, campo) {
        if (!r) return '';
        var val = r[campo];

        if (r.id) {
          var memId = assinaturasMemoriaMap.get(r.id + '_' + campo);
          if (memId) {
            const convertido = await toDataUrlFromSource(memId);
            if (convertido) return convertido;
          }
        }
        if (r.nome && r.data) {
          var key = (abaAtual || 'CAUTELAS') + '_' + r.nome.trim().toUpperCase() + '_' + r.data.trim() + '_' + campo;
          var memKey = assinaturasMemoriaMap.get(key);
          if (memKey) {
            const convertido = await toDataUrlFromSource(memKey);
            if (convertido) return convertido;
          }
        }
        if (r.nome) {
          var keyNome = (abaAtual || 'CAUTELAS') + '_' + r.nome.trim().toUpperCase() + '_' + campo;
          var memNome = assinaturasMemoriaMap.get(keyNome);
          if (memNome) {
            const convertido = await toDataUrlFromSource(memNome);
            if (convertido) return convertido;
          }
        }

        if (!val) return '';
        return toDataUrlFromSource(val);
      }

      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet(nomeAba || 'CAUTELAS');

      worksheet.columns = [
        { header: 'DATA RETIRADA', key: 'data', width: 16 },
        { header: 'GRADUAÇÃO / NOME', key: 'nome', width: 30 },
        { header: 'ARMAMENTOS', key: 'arma', width: 28 },
        { header: 'EQUIPAMENTOS', key: 'equip', width: 28 },
        { header: 'MUNIÇÃO', key: 'mun', width: 14 },
        { header: 'ASSINATURA RETIRADA', key: 'assF', width: 24 },
        { header: 'ASS. ARMEIRO ENTREGA', key: 'assG', width: 24 },
        { header: 'DATA DEVOLUÇÃO', key: 'dataDev', width: 16 },
        { header: 'ASSINATURA DEVOLUÇÃO', key: 'assI', width: 24 },
        { header: 'ASS. ARMEIRO RECEBEU', key: 'assJ', width: 24 },
        { header: 'OBSERVAÇÕES', key: 'obs', width: 28 },
      ];

      const headerRow = worksheet.getRow(1);
      headerRow.font = { bold: true, color: { argb: 'FFFFFF' }, size: 11 };
      headerRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0B5FFF' }
      };
      headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
      headerRow.height = 28;

      for (let idx = 0; idx < dados.length; idx++) {
        const r = dados[idx];
        const rowNum = idx + 2;
        const row = worksheet.addRow({
          data: r.data || '',
          nome: r.nome || '',
          arma: r.arma || '',
          equip: r.equip || '',
          mun: r.mun || '',
          dataDev: r.dataDev || '',
          obs: r.obs || ''
        });

        row.height = 65;
        row.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

        const camposAss = [
          { campo: 'f', col: 5 },
          { campo: 'g', col: 6 },
          { campo: 'i', col: 8 },
          { campo: 'j', col: 9 },
        ];

        for (const { campo, col } of camposAss) {
          const b64 = await obterBase64ParaExportacao(r, campo);
          if (b64 && b64.startsWith('data:image/')) {
            try {
              const imageId = workbook.addImage({
                base64: b64.includes(',') ? b64.split(',')[1] : b64,
                extension: 'png',
              });
              worksheet.addImage(imageId, {
                tl: { col: col, row: rowNum - 1 },
                ext: { width: 140, height: 58 },
                editAs: 'oneCell'
              });
            } catch (err) {
              console.warn('Erro ao inserir assinatura no XLSX:', err);
            }
          }
        }
      }

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      if (this._success) this._success(url);
    } catch (e) {
      if (this._failure) this._failure(e);
    }
  },
  enviarCopiaPorEmail: function(emailDestino, nomeAba) {
    setTimeout(() => {
      try {
        const dados = getStoredAbaDados(nomeAba);
        const assunto = encodeURIComponent(`Cópia - Reserva CEI - Aba ${nomeAba} (${new Date().toLocaleDateString('pt-BR')})`);
        let resumo = `RELATÓRIO DE CAUTELAS - RESERVA CEI\n`;
        resumo += `Aba: ${nomeAba}\n`;
        resumo += `Total de Cautelas: ${dados.length}\n`;
        resumo += `Data de Emissão: ${new Date().toLocaleString('pt-BR')}\n\n`;
        resumo += `LISTA DE REGISTROS:\n`;
        resumo += `----------------------------------------\n`;
        dados.slice(0, 50).forEach((r, idx) => {
          resumo += `${idx + 1}. [${r.data || '-'}] ${r.nome || 'SEM NOME'}\n`;
          if (r.arma) resumo += `   Armas: ${r.arma}\n`;
          if (r.equip) resumo += `   Equip: ${r.equip}\n`;
          if (r.mun) resumo += `   Mun: ${r.mun}\n`;
          if (r.dataDev) resumo += `   Devolução: ${r.dataDev}\n`;
          resumo += `\n`;
        });
        const corpo = encodeURIComponent(resumo);
        window.location.href = `mailto:${emailDestino}?subject=${assunto}&body=${corpo}`;
        if (this._success) this._success("Email preparado e aberto para " + emailDestino);
      } catch (e) {
        if (this._failure) this._failure(e);
      }
    }, 60);
  }
};

// ============================================================================
// VARIÁVEIS DE ESTADO E CATÁLOGO MILITAR ORIGINAL
// ============================================================================
var carregandoDados = false;
var tentativasErradas = parseInt(localStorage.getItem('tentativasReserva') || '0', 10);
var bloqueioAte = parseInt(localStorage.getItem('bloqueioReserva') || '0', 10);
var abaAtual = '', selRow = 0, selCol = 0, modoNovo = null;
var tempAss = { F: '', G: '', I: '', J: '' };
var editandoInfo = { row: 0, col: 0, valorAtual: '', tipo: 'texto' };
var pendingEdit = null;
var selecaoAtualTipo = '';
var valoresSelecionados = { armas: {}, equip: {}, linhasEdicao: {} };
var linhaEditandoAtiva = 0;
var linhaParaApagar = 0;
var dadosBrutos = [];
var filaEnvio = [];
var salvandoAssinatura = false;

var catalogoItens = {
  armas: [
    { cat: "Fuzis, Carabinas e Mosquetões", nome: "Mosquefal", pool: ["10280", "20153", "34273", "67501", "71639", "72054", "73381", "73427", "74210", "81163", "81193", "81199", "81352", "81807", "82416", "82747", "82770", "82775", "85549", "85555", "85620", "86231", "86393", "86467", "86882", "87652", "88126", "88206", "88253", "88407", "88833", "89026", "89195", "91162", "91248", "92425", "94511", "94605", "94643", "96038", "96544", "97038", "98653", "99161", "99651", "10433", "15775", "61781", "65225", "65561", "70638", "73181", "77220", "82559", "82768", "85222", "85681", "86306", "86526", "86939", "88439", "88802", "88830", "89074", "89124", "91139", "91197", "91198", "91489", "92212", "94613", "96880", "97348", "97998", "99099", "105729", "79069", "81486", "87557", "96813", "85766", "96505", "97481"] },
    { cat: "Fuzis, Carabinas e Mosquetões", nome: "Carabina IA2", pool: ["JFA 04155", "JFA 07945", "JFA 03884", "JFA 00826", "JFA12472", "JFA 12410", "JFA 12010", "JFA 07947", "JFA 117769", "JFA 01174", "JFA 00980", "JFA 00546", "JFA 05658", "JFA 01106", "JFA 02922", "JFA 05263", "JFA 12049", "JFA 01282", "JFA 04017"] },
    { cat: "Fuzis, Carabinas e Mosquetões", nome: "Carabina MD 97", pool: ["JAA04893", "JCA00288", "JAA02050", "JCA000413", "JAA04907", "JEA00280", "JAA03371"] },
    { cat: "Fuzis, Carabinas e Mosquetões", nome: "Mosquetão", pool: ["7365", "8761", "8308", "8488", "8089", "4468", "4403"] },
    { cat: "Espingardas (Calibre 12)", nome: "Calibre 12 (0.586)", pool: ["123935", "123953", "123967", "124676", "123916", "123913", "106565", "109008", "106453", "123912", "123904", "107914"] },
    { cat: "Espingardas (Calibre 12)", nome: "Calibre 12 Benelli", pool: ["091-W20", "010-C20", "009-K20", "007-H20", "093-Y20", "076-X20", "090-V20", "005-F20", "075-W20", "006-G20", "064-T20"] },
    { cat: "Espingardas (Calibre 12)", nome: "Calibre 12 Military", pool: ["KZB5258660", "KZB5258630", "KZB5258651"] },
    { cat: "Pistolas e Revólveres", nome: "Pistola.040 PT100", pool: ["SWA 25984", "SSJ34686", "STI 71917", "SSJ 34531", "STI 71686", "SVA 46552", "STI 71631", "STI 71572", "STI 71959", "SWA 25001", "SSJ 34696", "SWA 25125", "SWA 23608", "STI 71691", "STI 71633", "STI 71699", "SWA 23618", "SVI 72017", "STI 71929", "STI 71731", "STI 71523", "STI 71579", "SVI 71990", "SSJ 34629", "SWA 25081", "SWA 23688", "SVA 46505", "STI 71609"] },
    { cat: "Pistolas e Revólveres", nome: "Revólver.38", pool: ["OF285665", "OF286583", "J606961", "J606971", "J158943", "J158937", "J292016", "J292064", "J607794", "J607814"] },
    { cat: "Pistolas e Revólveres", nome: "Beretta APX", pool: ["AA-055498B", "AA-109648B", "AA-110346B", "AA-110136B", "AA-059924B", "AA-059925B", "AA110218B", "AA109596B", "AA109597B", "AA109921B", "AA109684B", "AA110031B", "AA110370B", "AA110146B", "AA059912B", "AA110283B", "AA109705B", "AA110288B", "AA109594B", "AA109896B", "AA110321B", "AA109868B", "AA110248B", "AA109974B", "AA109631B", "AA109805B", "AA109678B", "AA109739B", "AA110291B", "AA110318B", "AA110050B", "AA110070B", "AA110051B", "AA110224B", "AA109686B"] },
    { cat: "Submetralhadoras", nome: "Submetralhadora MT.040", pool: ["JAA04894"] },
    { cat: "Simulacros (Airsoft)", nome: "Pistola Airsoft", pool: ["2018-041355", "2018-045875", "2018-041357", "2018-045938", "2018-041369", "2018-045942", "2018-041416", "2018-045949", "2018-041418", "2018-045952", "2018-041419", "2018-046014", "2018-041421", "2018-046016", "2018-041423", "2018-046024", "2018-041438", "2018-046026", "2018-041447", "2018-046027", "2018-041452", "2018-046064", "2018-044339", "2018-046067", "2018-044357", "2018-046177", "2018-044359", "2018-046185", "2018-044364", "2018-046192"] },
    { cat: "Simulacros (Airsoft)", nome: "Rifle Airsoft AR15", pool: ["RNSA-AZUL-380022", "RNSA-AZUL-380622", "RNSA-AZUL-392122", "RNSA-AZUL-392222", "RNSA-AZUL-392322", "RNSA-VERDE-369422", "RNSA-VERDE-369522", "RNSA-VERDE-377122", "RNSA-VERDE-378122", "RNSA-VERDE-378222", "RNSA-VERMELHO-400422", "RNSA-VERMELHO-400922", "RNSA-VERMELHO-401122", "RNSA-VERMELHO-401222", "RNSA-VERMELHO-406922", "RNSA-AZUL-393222", "RNSA-AZUL-393322", "RNSA-AZUL-393522", "RNSA-AZUL-393622", "RNSA-AZUL-393822", "RNSA-VERDE-369622", "RNSA-VERDE-378322", "RNSA-VERDE-378422", "RNSA-VERDE-379722", "RNSA-VERDE-379822", "RNSA-VERMELHO-407022", "RNSA-VERMELHO-407122", "RNSA-VERMELHO-407222", "RNSA-VERMELHO-407922", "RNSA-VERMELHO-408022"] },
    { cat: "Carregadores", nome: "Carregador Calibre 5.56", pool: [] },
    { cat: "Carregadores", nome: "Carregador PT 100", pool: [] },
    { cat: "Carregadores", nome: "Carregador Fuzil Airsoft", pool: [] },
    { cat: "Carregadores", nome: "Carregador pistola Airsoft", pool: [] },
    { cat: "Carregadores", nome: "Carregador MT.040", pool: [] },
    { cat: "Carregadores", nome: "Carregador .9 mm", pool: [] }
  ],
  equip: [
    { cat: "Equipamentos", nome: "Algemas", pool: ["02156", "02189", "02382", "02440", "02520", "02521", "02536", "02544", "02545", "02546", "02579", "02580", "02585", "02587", "02588", "02589", "02592", "02593", "02594", "02597", "02598", "02666", "02671", "0298", "414", "02828"] },
    { cat: "Equipamentos", nome: "Rádio HT", pool: [] },
    { cat: "Equipamentos", nome: "Colete Balístico", pool: [] },
    { cat: "Equipamentos", nome: "Escudo Balístico", pool: [] },
    { cat: "Equipamentos", nome: "Lanterna Tática", pool: [] },
    { cat: "Equipamentos", nome: "Capacete", pool: [] },
    { cat: "Equipamentos", nome: "Colete Refletivo", pool: [] },
    { cat: "Equipamentos", nome: "Bandoleira", pool: [] },
    { cat: "Equipamentos", nome: "Porta carregador", pool: [] },
    { cat: "Equipamentos", nome: "Porta Tonfa", pool: [] },
    { cat: "Equipamentos", nome: "Bandeiras", pool: [] },
    { cat: "Equipamentos", nome: "Carregador de Bateria Airsoft", pool: [] },
    { cat: "Equipamentos", nome: "Carregador de Bateria HT", pool: [] },
    { cat: "Equipamentos", nome: "Base HT", pool: [] },
    { cat: "Equipamentos", nome: "Fone de ouvido PTT", pool: [] },
    { cat: "Equipamentos", nome: "Tonfa Bp-60", pool: [] },
    { cat: "Equipamentos", nome: "Tonfa Bp-90", pool: [] },
    { cat: "Equipamentos", nome: "Bastão", pool: [] }
  ]
};

// ============================================================================
// FUNÇÕES DE FILA E PERSISTÊNCIA LOCAL
// ============================================================================
function getFilaKey() { return 'filaEnvio_' + (abaAtual || 'GERAL'); }
function getChaveLocal() { return 'cei_' + (abaAtual || 'GERAL'); }
function carregarFila() {
  try { filaEnvio = JSON.parse(localStorage.getItem(getFilaKey()) || '[]'); }
  catch (e) { filaEnvio = []; }
  return filaEnvio;
}
function salvarFila() {
  try { localStorage.setItem(getFilaKey(), JSON.stringify(filaEnvio)); } catch (e) {}
  atualizarIndicadorFila();
}
function addFila(op) {
  if (op.type === 'ass' && (!op.data || !op.data.img || op.data.img.length < 100)) return;
  carregarFila();
  op.id = Date.now() + '_' + Math.random().toString(36).substr(2, 5);
  filaEnvio.push(op);
  salvarFila();
}
function atualizarIndicadorFila() {
  var fila = carregarFila();
  var b = document.getElementById('filaBadge');
  if (!b) {
    b = document.createElement('div');
    b.id = 'filaBadge';
    b.style.cssText = 'position:fixed;bottom:12px;right:12px;background:#0f172a;color:#fff;padding:8px 14px;border-radius:20px;font-size:11px;font-weight:900;z-index:999';
    document.body.appendChild(b);
  }
  b.style.display = fila.length > 0 ? 'block' : 'none';
  if (fila.length > 0) b.innerText = '🔄 ' + fila.length + ' enviando...';
}
function carregarLocal() {
  try {
    var r = localStorage.getItem(getChaveLocal());
    var a = r ? JSON.parse(r) : [];
    var lista = Array.isArray(a) ? a : [];
    return mesclarComAssinaturasDoLivro(lista, abaAtual);
  } catch (e) {
    return mesclarComAssinaturasDoLivro([], abaAtual);
  }
}
function salvarLocal() {
  try { localStorage.setItem(getChaveLocal(), JSON.stringify(dadosBrutos)); } catch (e) {}
}
function status(msg) {
  var s = document.getElementById('statusBar');
  if (!s) return;
  s.style.display = 'block';
  s.innerText = msg;
  clearTimeout(s._t);
  s._t = setTimeout(function () { s.style.display = 'none'; }, 4000);
}
function verificarBloqueio() {
  if (Date.now() < bloqueioAte) {
    var s = Math.ceil((bloqueioAte - Date.now()) / 1000);
    var e = document.getElementById('loginErro');
    if (e) e.innerText = 'Bloqueado por ' + s + 's';
    var inp = document.getElementById('senhaInput');
    if (inp) inp.disabled = true;
    setTimeout(verificarBloqueio, 1000);
    return true;
  }
  var inp = document.getElementById('senhaInput');
  if (inp) inp.disabled = false;
  return false;
}

function inicializarApp() {
  loadAbas();
  var hoje = new Date().toISOString().split('T')[0];
  var el = document.getElementById('dataRet');
  if (el) el.value = hoje;
  var fm = document.getElementById('filtroMes');
  if (fm) fm.value = new Date().getMonth();

  if (window._backupInterval) clearInterval(window._backupInterval);
  window._backupInterval = setInterval(function () {
    try {
      google.script.run.criarBackupSilencioso();
      status('☁️ Backup automático de 1h salvo');
    } catch (e) {}
  }, 3600000); // Executa pontualmente a cada 1 hora (3.600.000 ms)
}

function verificarSenha() {
  if (verificarBloqueio()) return;
  var dig = document.getElementById('senhaInput').value.trim();
  var err = document.getElementById('loginErro');
  var btn = document.getElementById('btnEntrar');
  if (!dig) { err.innerText = "Digite a senha"; return; }
  btn.innerHTML = "Verificando...";
  btn.disabled = true;

  google.script.run.withSuccessHandler(function (ok) {
    btn.innerHTML = "ENTRAR NO SISTEMA";
    btn.disabled = false;
    if (ok) {
      localStorage.setItem('authReserva', '1');
      localStorage.setItem('tentativasReserva', '0');
      document.getElementById('loginOverlay').style.display = 'none';
      inicializarApp();
    } else {
      tentativasErradas++;
      localStorage.setItem('tentativasReserva', tentativasErradas);
      if (tentativasErradas >= 5) {
        bloqueioAte = Date.now() + 60000;
        localStorage.setItem('bloqueioReserva', bloqueioAte);
        err.innerText = '5 erros! Bloqueado 1 min';
        verificarBloqueio();
      } else {
        err.innerText = 'Senha incorreta! ' + tentativasErradas + '/5';
      }
      document.getElementById('senhaInput').value = '';
    }
  }).withFailureHandler(function (e) {
    btn.innerHTML = "ENTRAR NO SISTEMA";
    btn.disabled = false;
    err.innerText = e.message;
  }).verificarSenhaServidor(dig);
}

function uniaoDados(server, local) {
  var mapaLocalPorId = {};
  var mapaLocalPorNomeData = {};
  var mapaLocalPorNome = {};

  (local || []).forEach(function (l) {
    if (!l) return;
    if (l.id) mapaLocalPorId[l.id] = l;
    if (l.nome && l.data) {
      var k = l.nome.trim().toUpperCase() + '|' + l.data.trim();
      mapaLocalPorNomeData[k] = l;
    }
    if (l.nome) {
      mapaLocalPorNome[l.nome.trim().toUpperCase()] = l;
    }
  });

  var resultado = (server || []).map(function (s, sIdx) {
    var item = Object.assign({}, s);
    if (!item.row) item.row = sIdx + 2;

    // Busca dados locais correspondentes pelo ID, Nome+Data ou Nome
    var l = null;
    if (item.id && mapaLocalPorId[item.id]) {
      l = mapaLocalPorId[item.id];
    } else if (item.nome && item.data) {
      var k = item.nome.trim().toUpperCase() + '|' + item.data.trim();
      l = mapaLocalPorNomeData[k] || mapaLocalPorNome[item.nome.trim().toUpperCase()];
    } else if (item.nome) {
      l = mapaLocalPorNome[item.nome.trim().toUpperCase()];
    }

    if (l) {
      // Se o servidor retornou célula vazia mas temos a assinatura local válida do MESMO militar, restaura!
      if (l.f && l.f.length > 30 && (!item.f || item.f.length < 30)) item.f = l.f;
      if (l.g && l.g.length > 30 && (!item.g || item.g.length < 30)) item.g = l.g;
      if (l.i && l.i.length > 30 && (!item.i || item.i.length < 30)) item.i = l.i;
      if (l.j && l.j.length > 30 && (!item.j || item.j.length < 30)) item.j = l.j;
      if (l.dataDev && (!item.dataDev || item.dataDev.trim() === '')) item.dataDev = l.dataDev;
      if (l.obs && (!item.obs || item.obs.trim() === '')) item.obs = l.obs;
      if (l.id && !item.id) item.id = l.id;
    }

    return item;
  });

  var idsNoResultado = new Set(resultado.map(function (r) { return r.id; }).filter(Boolean));

  // Adiciona itens locais que ainda não foram criados ou sincronizados no servidor
  (local || []).forEach(function (l) {
    if (!l) return;
    if (l.row < 0 || l._temp) {
      var jaExiste = resultado.some(function (s) {
        return (l.id && s.id === l.id) ||
               (s.nome && l.nome && s.nome.trim().toUpperCase() === l.nome.trim().toUpperCase() && s.data === l.data && s.arma === l.arma);
      });
      if (!jaExiste) {
        resultado.unshift(l);
        if (l.id) idsNoResultado.add(l.id);
      }
    }
  });

  return resultado;
}

function loadDados() {
  carregarFila();
  var local = carregarLocal();
  dadosBrutos = local;
  render(dadosBrutos);
  filtrarPorNome();
  status('💾 Local: ' + dadosBrutos.length + ' (fila: ' + filaEnvio.length + ')');
  atualizarIndicadorFila();
}

function filtrarPorNome() {
  var buscaEl = document.getElementById('buscaNome');
  var termo = buscaEl ? buscaEl.value.trim().toLowerCase() : '';
  var mesEl = document.getElementById('filtroMes');
  var mesFiltro = mesEl ? parseInt(mesEl.value, 10) : -1;

  var filtrada = dadosBrutos.filter(function (r) {
    // Filtro por mês automático da data de retirada
    if (mesFiltro !== -1 && r.data) {
      var partes = r.data.split('/');
      if (partes.length === 3) {
        var m = parseInt(partes[1], 10) - 1;
        if (m !== mesFiltro) return false;
      }
    }
    if (!termo) return true;
    return (r.nome && r.nome.toLowerCase().indexOf(termo) > -1) ||
           (r.arma && r.arma.toLowerCase().indexOf(termo) > -1) ||
           (r.equip && r.equip.toLowerCase().indexOf(termo) > -1);
  });
  render(filtrada);
}

function formatarAssinaturaSrc(val) {
  if (!val) return '';
  var s = String(val).trim();
  if (s.startsWith('=IMAGE(') || s.startsWith('=image(')) {
    var m = s.match(/"([^"]+)"/) || s.match(/'([^']+)'/);
    if (m && m[1]) return m[1];
  }
  if (s.startsWith('data:image/') || s.startsWith('http://') || s.startsWith('https://')) {
    return s;
  }
  if (s.length > 50 && !s.includes(' ')) {
    return 'data:image/png;base64,' + s;
  }
  return s;
}

async function prepararAssinaturaParaSheets(base64, nomeArquivo) {
  if (!base64 || typeof base64 !== 'string') return '';
  const raw = base64.trim();
  const looksLikeRawBase64 = /^[A-Za-z0-9+/=\r\n]+$/.test(raw) && raw.length > 200 && !raw.includes('http') && !raw.includes('=IMAGE');
  const base64Normalizado = looksLikeRawBase64 ? `data:image/png;base64,${raw}` : raw;
  const token = await getGoogleAccessToken();
  if (!token) return base64Normalizado;
  const valor = await uploadSignatureToDriveAndGetFormula(base64Normalizado, nomeArquivo || 'assinatura', token);
  return valor || '';
}

function render(lista) {
  var tb = document.getElementById('tbody');
  if (!tb) return;
  tb.innerHTML = '';
  if (!lista || lista.length == 0) {
    tb.innerHTML = '<tr><td colspan="12">Nenhum registro. Clique Sincronizar para buscar do Sheets.</td></tr>';
    return;
  }
  lista.forEach(function (r) {
    var tr = document.createElement('tr');
    tr.id = 'linha-' + r.row;

    var tdData = document.createElement('td');
    tdData.className = 'sticky';
    tdData.style.left = '0';
    tdData.innerHTML = '<div class="date-trigger">' + esc(r.data) + '</div>';
    tdData.querySelector('div').addEventListener('click', function () { abrirModalData(r.row, 1, r.data); });

    var tdNome = document.createElement('td');
    tdNome.className = 'sticky2';
    tdNome.style.left = '135px';
    tdNome.innerHTML = '<div class="text-trigger">' + esc(r.nome) + '</div>';
    tdNome.addEventListener('click', function () { solicitarEdicaoComConfirmacao(r.row, 2, r.nome, null, 'texto'); });

    var tdArma = document.createElement('td');
    tdArma.innerHTML = '<div class="select-trigger">' + esc(r.arma) + '</div>';
    tdArma.addEventListener('click', function () { solicitarEdicaoComConfirmacao(r.row, 3, r.arma, null, 'arma'); });

    var tdEquip = document.createElement('td');
    tdEquip.innerHTML = '<div class="select-trigger">' + esc(r.equip) + '</div>';
    tdEquip.addEventListener('click', function () { solicitarEdicaoComConfirmacao(r.row, 4, r.equip, null, 'equip'); });

    var tdMun = document.createElement('td');
    tdMun.innerHTML = '<div class="num-trigger">' + esc(r.mun) + '</div>';
    tdMun.querySelector('div').addEventListener('click', function () { abrirModalNum(r.row, 5, r.mun); });

    var srcF = formatarAssinaturaSrc(r.f);
    var tdF = document.createElement('td');
    tdF.id = 'cel-' + r.row + '-6';
    if (srcF && srcF.length > 20) {
      tdF.innerHTML = '<img class="thumb" src="' + srcF + '"><div style="font-size:8px;color:#10b981;font-weight:900">BLOQUEADO</div>';
    } else {
      tdF.innerHTML = '<button class="btn-ass">ASSINAR F</button>';
      tdF.querySelector('button').addEventListener('click', function () { abrir(r.row, 6); });
    }

    var srcG = formatarAssinaturaSrc(r.g);
    var tdG = document.createElement('td');
    tdG.id = 'cel-' + r.row + '-7';
    if (srcG && srcG.length > 20) {
      tdG.innerHTML = '<img class="thumb" src="' + srcG + '"><div style="font-size:8px;color:#10b981;font-weight:900">BLOQUEADO</div>';
    } else {
      tdG.innerHTML = '<button class="btn-ass">ASSINAR G</button>';
      tdG.querySelector('button').addEventListener('click', function () { abrir(r.row, 7); });
    }

    var tdDataDev = document.createElement('td');
    tdDataDev.innerHTML = '<div class="date-trigger">' + esc(r.dataDev) + '</div>';
    tdDataDev.querySelector('div').addEventListener('click', function () { abrirModalData(r.row, 8, r.dataDev); });

    var srcI = formatarAssinaturaSrc(r.i);
    var tdI = document.createElement('td');
    tdI.id = 'cel-' + r.row + '-9';
    if (srcI && srcI.length > 20) {
      tdI.innerHTML = '<img class="thumb" src="' + srcI + '"><div style="font-size:8px;color:#92400e;font-weight:900">BLOQUEADO</div>';
    } else {
      tdI.innerHTML = '<button class="btn-ass dev">ASSINAR I</button>';
      tdI.querySelector('button').addEventListener('click', function () { abrir(r.row, 9); });
    }

    var srcJ = formatarAssinaturaSrc(r.j);
    var tdJ = document.createElement('td');
    tdJ.id = 'cel-' + r.row + '-10';
    if (srcJ && srcJ.length > 20) {
      tdJ.innerHTML = '<img class="thumb" src="' + srcJ + '"><div style="font-size:8px;color:#92400e;font-weight:900">BLOQUEADO</div>';
    } else {
      tdJ.innerHTML = '<button class="btn-ass dev">ASSINAR J</button>';
      tdJ.querySelector('button').addEventListener('click', function () { abrir(r.row, 10); });
    }

    var tdObs = document.createElement('td');
    tdObs.innerHTML = '<div class="text-trigger">' + esc(r.obs) + '</div>';
    tdObs.addEventListener('click', function () { solicitarEdicaoComConfirmacao(r.row, 11, r.obs, null, 'texto'); });

    var tdDel = document.createElement('td');
    tdDel.innerHTML = '<button class="btn-del">Apagar Linha</button>';
    tdDel.querySelector('button').addEventListener('click', function () { solicitarApagar(r.row); });

    tr.appendChild(tdData);
    tr.appendChild(tdNome);
    tr.appendChild(tdArma);
    tr.appendChild(tdEquip);
    tr.appendChild(tdMun);
    tr.appendChild(tdF);
    tr.appendChild(tdG);
    tr.appendChild(tdDataDev);
    tr.appendChild(tdI);
    tr.appendChild(tdJ);
    tr.appendChild(tdObs);
    tr.appendChild(tdDel);

    tb.appendChild(tr);
  });
}

function toInputDate(d) {
  if (!d) return '';
  var limpo = d.replace(' (editado)', '');
  var p = limpo.split('/');
  if (p.length == 3) return p[2] + '-' + p[1] + '-' + p[0];
  return '';
}

function esc(t) {
  return t ? String(t).replace(/</g, '&lt;') : '';
}

function solicitarEdicaoComConfirmacao(row, col, valorAtual, novoValorDireto, tipo) {
  pendingEdit = { row: row, col: col, valorAtual: valorAtual, novoValorDireto: novoValorDireto, tipo: tipo };
  var m = document.getElementById('modalConfirmEdit');
  if (m) m.style.display = 'flex';
}

function confirmarEdicaoComConfirmacao() {
  var mc = document.getElementById('modalConfirmEdit');
  if (mc) mc.style.display = 'none';
  if (!pendingEdit) return;
  var row = pendingEdit.row, col = pendingEdit.col, valorAtual = pendingEdit.valorAtual, tipo = pendingEdit.tipo;
  pendingEdit = null;

  if (tipo === 'arma') { abrirSeletor('armas', true, row); return; }
  if (tipo === 'equip') { abrirSeletor('equip', true, row); return; }
  if (col == 1 || col == 8) { abrirModalData(row, col, valorAtual); return; }
  if (col == 5) { abrirModalNum(row, col, valorAtual); return; }

  editandoInfo = { row: row, col: col, valorAtual: valorAtual, tipo: 'texto' };
  var tit = document.getElementById('tituloModalTexto');
  if (tit) tit.innerText = col == 2 ? 'Editar NOME' : col == 11 ? 'Editar OBS' : 'Editar';
  var inp = document.getElementById('inputModalTexto');
  if (inp) inp.value = (valorAtual || '').replace(' (editado)', '');
  var mt = document.getElementById('modalTexto');
  if (mt) mt.style.display = 'flex';
}

function abrirModalData(row, col, valorAtual) {
  if (pendingEdit && pendingEdit.row == row && pendingEdit.col == col) {
  } else if (valorAtual && valorAtual.trim() !== '' && !pendingEdit) {
    pendingEdit = { row: row, col: col, valorAtual: valorAtual, tipo: 'data' };
    var mc = document.getElementById('modalConfirmEdit');
    if (mc) mc.style.display = 'flex';
    return;
  }
  editandoInfo = { row: row, col: col, valorAtual: valorAtual, tipo: 'data' };
  var input = document.getElementById('inputModalData');
  if (input) input.value = toInputDate(valorAtual);
  var tit = document.getElementById('tituloModalData');
  if (tit) tit.innerText = col == 1 ? 'Editar DATA RETIRADA' : 'Editar DATA DEVOLUÇÃO';
  var md = document.getElementById('modalData');
  if (md) md.style.display = 'flex';
}

function salvarModalData() {
  var v = document.getElementById('inputModalData').value;
  var row = editandoInfo.row, col = editandoInfo.col;
  var md = document.getElementById('modalData');
  if (md) md.style.display = 'none';

  var idx = dadosBrutos.findIndex(function (x) { return x.row == row; });
  var targetId = idx > -1 ? dadosBrutos[idx].id : '';
  if (idx > -1) {
    if (col == 1) dadosBrutos[idx].data = v ? v.split('-').reverse().join('/') : '';
    else dadosBrutos[idx].dataDev = v ? v.split('-').reverse().join('/') : '';
    salvarLocal();
    render(dadosBrutos);
  }
  addFila({ type: 'texto', data: { row: row, col: col, val: v, id: targetId } });
  status('⏳ Salvando...');
  setTimeout(processarFila, 100);
}

function abrirModalNum(row, col, valorAtual) {
  if (pendingEdit && pendingEdit.row == row && pendingEdit.col == col) {
  } else if (valorAtual && valorAtual.trim() !== '' && !pendingEdit) {
    pendingEdit = { row: row, col: col, valorAtual: valorAtual, tipo: 'num' };
    var mc = document.getElementById('modalConfirmEdit');
    if (mc) mc.style.display = 'flex';
    return;
  }
  editandoInfo = { row: row, col: col, valorAtual: valorAtual, tipo: 'num' };
  var input = document.getElementById('inputModalNum');
  if (input) input.value = (valorAtual || '').replace(' (editado)', '').replace(/[^0-9]/g, '');
  var mn = document.getElementById('modalNum');
  if (mn) mn.style.display = 'flex';
}

function salvarModalNum() {
  var v = document.getElementById('inputModalNum').value;
  var row = editandoInfo.row, col = editandoInfo.col;
  var mn = document.getElementById('modalNum');
  if (mn) mn.style.display = 'none';

  var idx = dadosBrutos.findIndex(function (x) { return x.row == row; });
  var targetId = idx > -1 ? dadosBrutos[idx].id : '';
  if (idx > -1) {
    dadosBrutos[idx].mun = v;
    salvarLocal();
    render(dadosBrutos);
  }
  addFila({ type: 'texto', data: { row: row, col: col, val: v, id: targetId } });
  setTimeout(processarFila, 100);
}

function salvarModalTexto() {
  var v = document.getElementById('inputModalTexto').value;
  var row = editandoInfo.row, col = editandoInfo.col;
  var mt = document.getElementById('modalTexto');
  if (mt) mt.style.display = 'none';

  var idx = dadosBrutos.findIndex(function (x) { return x.row == row; });
  var targetId = idx > -1 ? dadosBrutos[idx].id : '';
  if (idx > -1) {
    if (col == 2) dadosBrutos[idx].nome = v;
    else if (col == 11) dadosBrutos[idx].obs = v;
    salvarLocal();
    render(dadosBrutos);
  }
  addFila({ type: 'texto', data: { row: row, col: col, val: v, id: targetId } });
  setTimeout(processarFila, 100);
}

function solicitarApagar(row) {
  linhaParaApagar = row;
  var el = document.getElementById('linhaApagarNum');
  if (el) el.innerText = row;
  var inp = document.getElementById('inputApagarConfirm');
  if (inp) inp.value = '';
  var ma = document.getElementById('modalApagar');
  if (ma) ma.style.display = 'flex';
}

function confirmarApagar() {
  var dig = document.getElementById('inputApagarConfirm').value.trim();
  if (dig !== 'APAGAR') { alert('Digite APAGAR'); return; }
  var ma = document.getElementById('modalApagar');
  if (ma) ma.style.display = 'none';

  var idx = dadosBrutos.findIndex(function (x) { return x.row == linhaParaApagar; });
  var targetId = idx > -1 ? dadosBrutos[idx].id : '';
  if (idx > -1) {
    dadosBrutos.splice(idx, 1);
    salvarLocal();
    render(dadosBrutos);
    // Executa backup automático imediatamente ao apagar
    google.script.run.criarBackupSilencioso();
    status('🗑 Linha apagada e backup automático gerado com sucesso');
  }
  addFila({ type: 'apagar', data: { row: linhaParaApagar, id: targetId } });
  setTimeout(processarFila, 100);
}

function criarLinha() {
  var dataRet = document.getElementById('dataRet').value;
  var nome = document.getElementById('nome').value.trim();
  var armasBtn = document.getElementById('btnArmas').innerText;
  var equipBtn = document.getElementById('btnEquip').innerText;
  var mun = document.getElementById('mun').value;
  var obs = document.getElementById('obs').value;
  var dataDev = document.getElementById('dataDev').value;

  if (!nome) { alert('Nome obrigatório'); return; }

  var armas = armasBtn.indexOf('Selecionar') > -1 ? '' : armasBtn;
  var equip = equipBtn.indexOf('Selecionar') > -1 ? '' : equipBtn;
  var tempRow = -(Date.now());
  var novoId = generateRowId();

  var novo = {
    id: novoId,
    row: tempRow,
    data: dataRet ? dataRet.split('-').reverse().join('/') : new Date().toLocaleDateString('pt-BR'),
    nome: nome,
    arma: armas,
    equip: equip,
    mun: mun,
    f: tempAss.F,
    g: tempAss.G,
    dataDev: dataDev ? dataDev.split('-').reverse().join('/') : '',
    i: tempAss.I,
    j: tempAss.J,
    obs: obs,
    _temp: true
  };

  dadosBrutos.unshift(novo);
  salvarLocal();
  render(dadosBrutos);

  addFila({
    type: 'create',
    data: {
      id: novoId,
      dataRet: dataRet,
      nome: nome,
      armas: armas,
      equip: equip,
      mun: mun,
      obs: obs,
      dataDev: dataDev,
      assF: tempAss.F,
      assG: tempAss.G,
      assI: tempAss.I,
      assJ: tempAss.J,
      tempRow: tempRow
    }
  });

  limparFormulario();
  status('⚡ Local + enviando...');
  setTimeout(processarFila, 200);
}

function limparFormulario() {
  var dr = document.getElementById('dataRet');
  if (dr) dr.value = new Date().toISOString().split('T')[0];
  var n = document.getElementById('nome');
  if (n) n.value = '';
  var m = document.getElementById('mun');
  if (m) m.value = '';
  var ob = document.getElementById('obs');
  if (ob) ob.value = '';
  var dd = document.getElementById('dataDev');
  if (dd) dd.value = '';
  var ba = document.getElementById('btnArmas');
  if (ba) ba.innerText = 'Selecionar Armamentos';
  var be = document.getElementById('btnEquip');
  if (be) be.innerText = 'Selecionar Equipamentos';

  valoresSelecionados = { armas: {}, equip: {}, linhasEdicao: {} };
  tempAss = { F: '', G: '', I: '', J: '' };

  document.querySelectorAll('.btn-ass').forEach(function (b) {
    b.classList.remove('has');
    b.innerHTML = 'ASSINAR ' + b.id.replace('btn', '');
  });
}

async function processarFila() {
  carregarFila();
  if (filaEnvio.length == 0) return;
  if (!navigator.onLine) { status('⚠ Offline ' + filaEnvio.length); return; }

  var item = filaEnvio[0];
  status('📤 Enviando ' + filaEnvio.length + '...');

  if (item.type == 'create') {
    const assF = await prepararAssinaturaParaSheets(item.data.assF, 'assF_' + (item.data.nome || 'novo'));
    const assG = await prepararAssinaturaParaSheets(item.data.assG, 'assG_' + (item.data.nome || 'novo'));
    const assI = await prepararAssinaturaParaSheets(item.data.assI, 'assI_' + (item.data.nome || 'novo'));
    const assJ = await prepararAssinaturaParaSheets(item.data.assJ, 'assJ_' + (item.data.nome || 'novo'));

    const uploadFalhou = [
      [item.data.assF, assF],
      [item.data.assG, assG],
      [item.data.assI, assI],
      [item.data.assJ, assJ]
    ].some(function ([original, convertido]) { return original && !convertido; });
    if (uploadFalhou) {
      status('⚠ Upload da assinatura falhou; mantendo na fila para tentar novamente.');
      setTimeout(processarFila, 5000);
      return;
    }

    google.script.run.withSuccessHandler(function (r) {
      var realRow = typeof r === 'number' ? r : (r && r.row ? r.row : 2);
      var idx = dadosBrutos.findIndex(function (x) { return x.row == item.data.tempRow; });
      if (idx > -1) {
        dadosBrutos[idx].row = realRow;
        delete dadosBrutos[idx]._temp;
        if (realRow == 2) {
          dadosBrutos.forEach(function (r, i) {
            if (i != idx && r.row >= 2) r.row++;
          });
        }
        salvarLocal();
      }
      filaEnvio.shift();
      salvarFila();
      if (filaEnvio.length > 0) setTimeout(processarFila, 400);
      else status('✅ Tudo gravado');
    }).withFailureHandler(function () {
      status('❌ Falha ao gravar o registro; ele permanece na fila para tentar novamente.');
      setTimeout(processarFila, 5000);
    }).criarNovaLinhaComAss(
      abaAtual,
      item.data.dataRet,
      item.data.nome,
      item.data.armas,
      item.data.equip,
      item.data.mun,
      item.data.obs,
      item.data.dataDev,
      assF,
      assG,
      assI,
      assJ
    );
  } else if (item.type == 'ass') {
    const valorParaSheets = await prepararAssinaturaParaSheets(item.data.img, 'assinatura_' + item.data.row + '_' + item.data.col);
    if (!valorParaSheets) {
      status('⚠ Upload da assinatura falhou; mantendo na fila para tentar novamente.');
      setTimeout(processarFila, 5000);
      return;
    }
    google.script.run.withSuccessHandler(function () {
      filaEnvio.shift();
      salvarFila();
      if (filaEnvio.length > 0) setTimeout(processarFila, 400);
      else status('✅ Assinatura gravada');
    }).withFailureHandler(function () {
      status('❌ Falha ao gravar assinatura; mantendo na fila para tentar novamente.');
      setTimeout(processarFila, 5000);
    }).salvarNaCelula(valorParaSheets, item.data.row, item.data.col, abaAtual, item.data.nomeMilitar, item.data.id);
  } else if (item.type == 'texto') {
    google.script.run.withSuccessHandler(function () {
      filaEnvio.shift();
      salvarFila();
      if (filaEnvio.length > 0) setTimeout(processarFila, 400);
    }).withFailureHandler(function () {
      status('❌ Falha ao salvar a edição; mantendo na fila para tentar novamente.');
      setTimeout(processarFila, 5000);
    }).salvarTextoInline(abaAtual, item.data.row, item.data.col, item.data.val, item.data.id);
  } else if (item.type == 'apagar') {
    google.script.run.withSuccessHandler(function () {
      filaEnvio.shift();
      salvarFila();
      if (filaEnvio.length > 0) setTimeout(processarFila, 400);
    }).withFailureHandler(function () {
      status('❌ Falha ao apagar o registro; mantendo na fila para tentar novamente.');
      setTimeout(processarFila, 5000);
    }).apagarLinha(abaAtual, item.data.row, item.data.id);
  }
}

function sincronizarManual() {
  var btn = document.getElementById('btnSinc');
  if (btn) { btn.disabled = true; btn.innerText = '⏳ Enviando fila...'; }
  carregarFila();
  var localAntes = carregarLocal() || dadosBrutos.slice();

  var enviarTudo = function (cb) {
    if (filaEnvio.length == 0) { cb(); return; }
    var it = filaEnvio[0];
    var ok = function () {
      filaEnvio.shift();
      salvarFila();
      setTimeout(function () { enviarTudo(cb); }, 400);
    };

    if (it.type == 'create') {
      google.script.run.withSuccessHandler(function (r) {
        var real = typeof r === 'number' ? r : r.row;
        var idx = dadosBrutos.findIndex(function (x) { return x.row == it.data.tempRow; });
        if (idx > -1) {
          dadosBrutos[idx].row = real;
          delete dadosBrutos[idx]._temp;
          if (real == 2) {
            dadosBrutos.forEach(function (r, i) {
              if (i != idx && r.row >= 2) r.row++;
            });
          }
          salvarLocal();
        }
        ok();
      }).withFailureHandler(function (e) {
        if (btn) { btn.disabled = false; btn.innerText = '🔄 Sincronizar'; }
        status('❌ Falha ao gravar registro: ' + (e && e.message ? e.message : 'tente novamente'));
      }).criarNovaLinhaComAss(
        abaAtual,
        it.data.dataRet,
        it.data.nome,
        it.data.armas,
        it.data.equip,
        it.data.mun,
        it.data.obs,
        it.data.dataDev,
        it.data.assF,
        it.data.assG,
        it.data.assI,
        it.data.assJ,
        it.data.id
      );
    } else if (it.type == 'ass') {
      google.script.run.withSuccessHandler(ok).withFailureHandler(function (e) {
        if (btn) { btn.disabled = false; btn.innerText = '🔄 Sincronizar'; }
        status('❌ Falha ao sincronizar assinatura: ' + (e && e.message ? e.message : 'tente novamente'));
      }).salvarNaCelula(it.data.img, it.data.row, it.data.col, abaAtual, it.data.nomeMilitar, it.data.id);
    } else if (it.type == 'texto') {
      google.script.run.withSuccessHandler(ok).withFailureHandler(function (e) {
        if (btn) { btn.disabled = false; btn.innerText = '🔄 Sincronizar'; }
        status('❌ Falha ao salvar edição: ' + (e && e.message ? e.message : 'tente novamente'));
      }).salvarTextoInline(abaAtual, it.data.row, it.data.col, it.data.val, it.data.id);
    } else {
      google.script.run.withSuccessHandler(ok).withFailureHandler(function (e) {
        if (btn) { btn.disabled = false; btn.innerText = '🔄 Sincronizar'; }
        status('❌ Falha ao apagar registro: ' + (e && e.message ? e.message : 'tente novamente'));
      }).apagarLinha(abaAtual, it.data.row, it.data.id);
    }
  };

  enviarTudo(function () {
    if (btn) btn.innerText = '🔍 Buscando...';
    status('🔍 Buscando e fazendo UNIÃO...');
    var mesEl = document.getElementById('filtroMes');
    var mes = mesEl ? parseInt(mesEl.value, 10) : -1;

    google.script.run.withSuccessHandler(function (server) {
      if (btn) { btn.disabled = false; btn.innerText = '🔄 Sincronizar'; }
      server = server || [];
      var uniao = uniaoDados(server, localAntes);
      dadosBrutos = uniao;
      salvarLocal();
      render(dadosBrutos);
      filtrarPorNome();
      status('✔ UNIÃO: ' + server.length + ' gravados = ' + uniao.length);
    }).withFailureHandler(function (e) {
      if (btn) { btn.disabled = false; btn.innerText = '🔄 Sincronizar'; }
      status('❌ ' + e.message);
    }).getPlanilhaCompleta(abaAtual, mes);
  });
}

function loadAbas() {
  google.script.run.withSuccessHandler(function (abas) {
    var t = document.getElementById('tabs');
    if (!t) return;
    t.innerHTML = '';
    var podeApagar = abas.length > 1;

    abas.forEach(function (n) {
      var d = document.createElement('div');
      d.className = 'tab' + (n == abaAtual ? ' active' : '');
      var delBtn = podeApagar ? '<span class="tab-del">X</span>' : '';
      d.innerHTML = '<span class="tab-nome">' + n + '</span> <span class="tab-edit" style="opacity:0.7;margin-left:4px">✏</span> ' + delBtn;

      d.querySelector('.tab-nome').addEventListener('click', function () { trocarAba(n); });
      var editEl = d.querySelector('.tab-edit');
      if (editEl) editEl.addEventListener('click', function () { editarAba(n); });
      var delEl = d.querySelector('.tab-del');
      if (delEl) delEl.addEventListener('click', function () { apagarAba(n); });

      t.appendChild(d);
    });

    if (!abaAtual && abas.length) abaAtual = abas[0];
    if (abaAtual) loadDados();
  }).getAbas();
}

function trocarAba(nome) {
  abaAtual = nome;
  loadDados();
  loadAbas();
}

let abaParaEditarNomeAntigo = '';

function editarAba(old) {
  abaParaEditarNomeAntigo = old;
  const input = document.getElementById('inputNomeEditarAba');
  if (input) input.value = old;
  const m = document.getElementById('modalEditarAba');
  if (m) m.style.display = 'flex';
  if (input) setTimeout(() => { input.focus(); input.select(); }, 100);
}

function confirmarEdicaoNomeAba() {
  const input = document.getElementById('inputNomeEditarAba');
  const novo = input ? input.value.trim().toUpperCase() : '';
  if (!novo || novo === abaParaEditarNomeAntigo) {
    const m = document.getElementById('modalEditarAba');
    if (m) m.style.display = 'none';
    return;
  }
  const m = document.getElementById('modalEditarAba');
  if (m) m.style.display = 'none';

  status('⏳ Renomeando aba...');
  google.script.run.withSuccessHandler(function () {
    abaAtual = novo;
    loadAbas();
    status('✔ Aba renomeada para ' + novo);
  }).withFailureHandler(function(e) {
    status('❌ Erro ao renomear: ' + e.message);
  }).renomearAba(abaParaEditarNomeAntigo, novo);
}

let abaParaApagarNome = '';

function apagarAba(nome) {
  abaParaApagarNome = nome;
  const tit = document.getElementById('nomeAbaApagar');
  if (tit) tit.innerText = nome;
  const inp = document.getElementById('inputApagarAbaConfirm');
  if (inp) inp.value = '';
  const m = document.getElementById('modalApagarAba');
  if (m) m.style.display = 'flex';
}

function confirmarExclusaoAba() {
  const inp = document.getElementById('inputApagarAbaConfirm');
  const dig = inp ? inp.value.trim() : '';
  if (dig !== 'APAGAR') {
    status('⚠️ Digite APAGAR para confirmar a exclusão');
    return;
  }
  const m = document.getElementById('modalApagarAba');
  if (m) m.style.display = 'none';

  status('⏳ Apagando aba ' + abaParaApagarNome + '...');
  google.script.run.withSuccessHandler(function (abas) {
    status('✔ Aba ' + abaParaApagarNome + ' excluída com sucesso');
    abaAtual = abas[0] || '';
    loadAbas();
  }).withFailureHandler(function (e) {
    status('❌ Erro: ' + e.message);
  }).deletarAba(abaParaApagarNome);
}

function salvarCopiaLocal() {
  if (!abaAtual) { status("⚠️ Selecione uma aba primeiro"); return; }
  status('⏳ Gerando planilha oficial do Excel (.xlsx) com assinaturas incorporadas...');
  google.script.run.withSuccessHandler(function (url) {
    var link = document.createElement('a');
    link.href = url;
    link.download = `RESERVA_CEI_${abaAtual}.xlsx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    status('📥 Planilha Excel (.xlsx) baixada com sucesso!');
  }).withFailureHandler(function (e) {
    status('❌ Não foi possível gerar o arquivo: ' + (e && e.message ? e.message : 'tente novamente'));
  }).getUrlExportacao(abaAtual, 'xlsx');
}

function enviarCopiaEmail() {
  var email = document.getElementById("emailCopia").value.trim();
  if (!email) { status("⚠️ Digite um email para enviar a cópia"); return; }
  status('✉️ Preparando cópia para ' + email + '...');
  google.script.run.withSuccessHandler(function (msg) {
    status('✅ ' + msg);
  }).withFailureHandler(function (e) {
    status('❌ Não foi possível enviar a cópia: ' + (e && e.message ? e.message : 'tente novamente'));
  }).enviarCopiaPorEmail(email, abaAtual);
}

function criarBackupDrive() {
  status('⏳ Gerando backup da reserva...');
  google.script.run.withSuccessHandler(function (m) {
    status(m);
  }).criarBackupPlanilha();
}

function novaAba() {
  const input = document.getElementById('inputNomeNovaAba');
  if (input) input.value = '';
  const m = document.getElementById('modalNovaAba');
  if (m) m.style.display = 'flex';
  if (input) setTimeout(() => input.focus(), 100);
}

function confirmarCriacaoNovaAba() {
  const input = document.getElementById('inputNomeNovaAba');
  const nome = input ? input.value.trim().toUpperCase() : '';
  if (!nome) {
    status('⚠️ Digite o nome da aba');
    return;
  }
  const m = document.getElementById('modalNovaAba');
  if (m) m.style.display = 'none';

  status('⏳ Criando aba ' + nome + '...');
  google.script.run.withSuccessHandler(function (n) {
    abaAtual = n;
    loadAbas();
    status('✔ Aba ' + n + ' criada com sucesso');
  }).withFailureHandler(function(e) {
    status('❌ Erro ao criar aba: ' + e.message);
  }).criarAba(nome);
}

// ============================================================================
// CANVAS DE ASSINATURA
// ============================================================================
var canvas = null, ctx = null;

function initCanvas() {
  canvas = document.getElementById('c');
  if (!canvas) return;
  var rect = canvas.parentElement.getBoundingClientRect();
  canvas.width = rect.width * 2;
  canvas.height = 760;
  var ctx2 = canvas.getContext('2d', { willReadFrequently: true });
  ctx2.scale(2, 2);
  ctx = ctx2;
  ctx.lineWidth = 2;
  ctx.lineCap = 'round';
  ctx.strokeStyle = '#000';
  ctx.lineJoin = 'round';
}

function resize() {
  if (!canvas || !ctx) return;
  var r = canvas.getBoundingClientRect();
  var dpr = window.devicePixelRatio || 1;
  canvas.width = r.width * dpr;
  canvas.height = r.height * dpr;
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.scale(dpr, dpr);
  ctx.lineWidth = 4;
  ctx.lineCap = 'round';
  ctx.lineJoin = 'round';
  ctx.strokeStyle = '#000';
}

var des = false, last = {};

function pos(e) {
  if (!canvas) return { x: 0, y: 0 };
  var r = canvas.getBoundingClientRect();
  var t = e.touches ? e.touches[0] : e;
  return { x: t.clientX - r.left, y: t.clientY - r.top };
}

function isLinhaValida(row) {
  var item = dadosBrutos.find(function (x) { return x.row == row; });
  if (!item) return false;
  if (item._temp) return false;
  if (!item.nome || item.nome.trim() === '') return false;
  if (item.row <= 0) return false;
  return true;
}

function abrir(row, col) {
  var lista = dadosBrutos.find(function (x) { return x.row == row; });
  if (!lista) { alert('⛔ Linha não encontrada! Sincronize.'); return; }
  if (!isLinhaValida(row)) { alert('⛔ Não pode assinar linha em branco ou temporária! Preencha o nome primeiro e sincronize.'); return; }
  if ((col == 6 && lista.f) || (col == 7 && lista.g) || (col == 9 && lista.i) || (col == 10 && lista.j)) {
    alert('⚠ Já assinado e bloqueado!');
    return;
  }
  if (lista.nome && lista.nome.trim() === '') { alert('⛔ Linha em branco bloqueada!'); return; }

  selRow = row; selCol = col; modoNovo = null;
  var mt = document.getElementById('mt');
  if (mt) mt.innerText = 'Assinar linha ' + row + ' - ' + lista.nome;
  var mo = document.getElementById('modal');
  if (mo) mo.style.display = 'flex';

  setTimeout(function () {
    resize();
    if (ctx) ctx.clearRect(0, 0, canvas.width, canvas.height);
  }, 100);
}

function abrirNovo(letra) {
  var nomeInput = document.getElementById('nome');
  if (!nomeInput || !nomeInput.value.trim()) {
    alert('⛔ Preencha GRADUAÇÃO/NOME antes de assinar!');
    if (nomeInput) nomeInput.focus();
    return;
  }
  modoNovo = letra;
  var mt = document.getElementById('mt');
  if (mt) mt.innerText = 'Assinar ' + letra + ' - NOVO CADASTRO';
  var mo = document.getElementById('modal');
  if (mo) mo.style.display = 'flex';

  setTimeout(function () {
    resize();
    if (ctx) ctx.clearRect(0, 0, canvas.width, canvas.height);
  }, 100);
}

function fechar() {
  var mo = document.getElementById('modal');
  if (mo) mo.style.display = 'none';
  modoNovo = null;
}

function getTrim() {
  if (!canvas || !ctx) return null;
  var w = canvas.width, h = canvas.height;
  var data = ctx.getImageData(0, 0, w, h).data;
  var minX = w, minY = h, maxX = 0, maxY = 0, achou = false;

  for (var y = 0; y < h; y += 2) {
    for (var x = 0; x < w; x += 2) {
      if (data[(y * w + x) * 4 + 3] > 10) {
        achou = true;
        if (x < minX) minX = x;
        if (x > maxX) maxX = x;
        if (y < minY) minY = y;
        if (y > maxY) maxY = y;
      }
    }
  }
  if (!achou) return null;

  var margem = 20;
  minX = Math.max(0, minX - margem);
  minY = Math.max(0, minY - margem);
  maxX = Math.min(w, maxX + margem);
  maxY = Math.min(h, maxY + margem);

  var cw = maxX - minX, ch = maxY - minY;
  var tmp = document.createElement('canvas');
  tmp.width = cw; tmp.height = ch;
  tmp.getContext('2d', { willReadFrequently: true }).drawImage(canvas, minX, minY, cw, ch, 0, 0, cw, ch);
  return tmp;
}

function aplicarDataDevolucaoAutomatica(row) {
  var hojeISO = new Date().toISOString().split('T')[0];
  var hojeBR = hojeISO.split('-').reverse().join('/');
  var idx = dadosBrutos.findIndex(function (x) { return x.row == row; });
  var precisa = false;

  if (idx > -1) {
    if (!dadosBrutos[idx].dataDev || dadosBrutos[idx].dataDev.trim() === '') {
      dadosBrutos[idx].dataDev = hojeBR;
      precisa = true;
    }
  } else {
    var el = document.getElementById('dataDev');
    if (el && !el.value) { el.value = hojeISO; precisa = true; }
  }

  if (precisa) {
    salvarLocal();
    render(dadosBrutos);
    addFila({ type: 'texto', data: { row: row, col: 8, val: hojeISO } });
    setTimeout(processarFila, 100);
  }
}

function salvarAss() {
  var t = getTrim();
  if (!t) { alert('Assine!'); return; }
  var b = t.toDataURL('image/png');
  var s = document.getElementById('st');
  if (s) s.innerHTML = 'Salvando...';

  if (modoNovo) {
    tempAss[modoNovo] = b;
    var btn = document.getElementById('btn' + modoNovo);
    if (btn) {
      btn.innerHTML = '<img class="thumb-instant" src="' + b + '" style="max-height:50px">ASSINADO ✅';
      btn.classList.add('has');
    }
    if (modoNovo === 'I' || modoNovo === 'J') {
      var el = document.getElementById('dataDev');
      if (el && !el.value) {
        var hoje = new Date().toISOString().split('T')[0];
        el.value = hoje;
      }
    }
    status('✔ Assinatura capturada');
    setTimeout(function () { fechar(); }, 400);
  } else {
    var idx = dadosBrutos.findIndex(function (x) { return x.row == selRow; });
    var campo = selCol == 6 ? 'f' : selCol == 7 ? 'g' : selCol == 9 ? 'i' : 'j';
    var nomeMilitarLinha = idx > -1 ? dadosBrutos[idx].nome : '';
    var linhaId = idx > -1 ? dadosBrutos[idx].id : '';

    if (idx > -1) {
      if (!isLinhaValida(selRow)) {
        alert('⛔ Linha inválida bloqueada! Não foi possível salvar assinatura.');
        fechar();
        return;
      }
      dadosBrutos[idx][campo] = b;
      var dataMilitarLinha = idx > -1 ? dadosBrutos[idx].data : '';
      registrarAssinaturaMemoria(linhaId, nomeMilitarLinha, dataMilitarLinha, abaAtual, campo, b);
      salvarLocal();
      render(dadosBrutos);
      // Realiza backup automático silencioso ao salvar assinatura
      google.script.run.criarBackupSilencioso();
    }
    if (selCol == 9 || selCol == 10) {
      aplicarDataDevolucaoAutomatica(selRow);
    }
    fechar();
    addFila({ type: 'ass', data: { row: selRow, col: selCol, img: b, nomeMilitar: nomeMilitarLinha, id: linhaId } });
    status('✍ Assinatura salva com sucesso (backup automático gerado)');
    setTimeout(processarFila, 100);
  }
}

// ============================================================================
// SELETOR DE ARMAMENTOS E EQUIPAMENTOS
// ============================================================================
function abrirSeletor(tipo, isEdicao, rowNum) {
  selecaoAtualTipo = tipo;
  linhaEditandoAtiva = rowNum;
  var tit = document.getElementById('tituloModalSelecao');
  if (tit) tit.innerText = tipo === 'armas' ? 'Selecionar Armamentos' : 'Selecionar Equipamentos';

  var container = document.getElementById('listaOpcoesContainer');
  if (!container) return;
  container.innerHTML = '';

  var buscaCat = document.getElementById('buscaItemCatalogo');
  if (buscaCat) buscaCat.value = '';

  var lista = tipo === 'armas' ? catalogoItens.armas : catalogoItens.equip;
  var valoresAtuais = {};

  if (isEdicao && rowNum > 0) {
    var linha = dadosBrutos.find(function (x) { return x.row == rowNum; });
    if (linha) {
      var texto = tipo === 'armas' ? linha.arma : linha.equip;
      if (texto) {
        texto.split('|').forEach(function (p) {
          var trimmed = p.trim();
          if (!trimmed) return;
          var m = trimmed.match(/^(?:(\d+)\s+)?(.+?)(?:\s*\[Reg:\s*(.+?)\])?\s*$/);
          if (m) {
            var qtd = m[1] || '1';
            var nome = (m[2] || '').trim();
            var reg = (m[3] || '').trim();
            if (nome) valoresAtuais[nome] = { qtd: qtd, reg: reg };
          }
        });
      }
    }
  } else {
    // Para nova cautela (linha 0), restaura o que o usuário já havia digitado anteriormente!
    if (valoresSelecionados[tipo]) {
      Object.keys(valoresSelecionados[tipo]).forEach(function(k) {
        valoresAtuais[k] = Object.assign({}, valoresSelecionados[tipo][k]);
      });
    }
    // Também extrai do texto atual do botão caso já esteja preenchido
    var btnAtual = document.getElementById(tipo === 'armas' ? 'btnArmas' : 'btnEquip');
    if (btnAtual && btnAtual.innerText && btnAtual.innerText.indexOf('Selecionar') === -1) {
      btnAtual.innerText.split('|').forEach(function (p) {
        var trimmed = p.trim();
        if (!trimmed) return;
        var m = trimmed.match(/^(?:(\d+)\s+)?(.+?)(?:\s*\[Reg:\s*(.+?)\])?\s*$/);
        if (m) {
          var qtd = m[1] || '1';
          var nome = (m[2] || '').trim();
          var reg = (m[3] || '').trim();
          if (nome && !valoresAtuais[nome]) valoresAtuais[nome] = { qtd: qtd, reg: reg };
        }
      });
    }
  }

  // Preenche também o item livre caso houvesse sido digitado
  var nomeLivre = document.getElementById('nomeItemLivre');
  var qtdLivre = document.getElementById('qtdItemLivre');
  var regLivre = document.getElementById('regItemLivre');
  if (nomeLivre && (!isEdicao || rowNum <= 0)) {
    // mantém o que estava digitado
  } else if (nomeLivre && isEdicao) {
    nomeLivre.value = '';
    if (qtdLivre) qtdLivre.value = '';
    if (regLivre) regLivre.value = '';
  }

  var ultimaCat = '';
  lista.forEach(function (item) {
    if (item.cat && item.cat !== ultimaCat) {
      ultimaCat = item.cat;
      var d = document.createElement('div');
      d.className = 'categoria-titulo';
      d.innerText = ultimaCat;
      container.appendChild(d);
    }

    var card = document.createElement('div');
    card.className = 'card-item';
    card.dataset.nome = item.nome.toLowerCase();

    var nomeDiv = document.createElement('div');
    nomeDiv.className = 'card-item-nome';
    nomeDiv.innerText = item.nome;

    var camposDiv = document.createElement('div');
    camposDiv.className = 'card-item-campos';

    var regWrap = document.createElement('div');
    regWrap.className = 'reg-container';

    var inputReg = document.createElement('input');
    inputReg.type = 'text';
    inputReg.className = 'reg-input';
    inputReg.placeholder = 'Nº Reg';
    inputReg.dataset.nome = item.nome;
    if (valoresAtuais[item.nome] && valoresAtuais[item.nome].reg) inputReg.value = valoresAtuais[item.nome].reg;

    var box = document.createElement('div');
    box.className = 'sugestoes-box';

    if (item.pool && item.pool.length > 0) {
      inputReg.addEventListener('input', function () {
        var termo = this.value.toLowerCase();
        var filtrados = item.pool.filter(function (p) { return p.toLowerCase().indexOf(termo) > -1; }).slice(0, 8);
        if (!termo || filtrados.length == 0) { box.style.display = 'none'; return; }
        box.innerHTML = '';
        filtrados.forEach(function (s) {
          var el = document.createElement('div');
          el.className = 'sugestao-item';
          el.innerText = s;
          el.addEventListener('click', function () {
            inputReg.value = s;
            box.style.display = 'none';
          });
          box.appendChild(el);
        });
        box.style.display = 'block';
      });

      inputReg.addEventListener('blur', function () {
        setTimeout(function () { box.style.display = 'none'; }, 200);
      });
    }

    var qtdInput = document.createElement('input');
    qtdInput.type = 'number';
    qtdInput.className = 'qtd-input';
    qtdInput.placeholder = 'Qtd';
    qtdInput.dataset.nome = item.nome;
    if (valoresAtuais[item.nome] && valoresAtuais[item.nome].qtd) qtdInput.value = valoresAtuais[item.nome].qtd;

    var temValor = !!(valoresAtuais[item.nome] && (valoresAtuais[item.nome].reg || valoresAtuais[item.nome].qtd));
    if (temValor) {
      card.classList.add('item-preenchido');
      card.style.borderColor = '#0B5FFF';
      card.style.background = '#f0f7ff';
    }

    var btnLimparCard = document.createElement('button');
    btnLimparCard.type = 'button';
    btnLimparCard.className = 'btn-limpar-card';
    btnLimparCard.innerText = '✕ Limpar';
    btnLimparCard.title = 'Limpar este item';
    btnLimparCard.style.cssText = 'padding:6px 10px;font-size:11px;font-weight:800;background:#fee2e2;color:#ef4444;border:1px solid #fca5a5;border-radius:8px;cursor:pointer;display:' + (temValor ? 'inline-block' : 'none');
    btnLimparCard.addEventListener('click', function(e) {
      e.stopPropagation();
      inputReg.value = '';
      qtdInput.value = '';
      card.classList.remove('item-preenchido');
      card.style.borderColor = '#e2e8f0';
      card.style.background = '#fff';
      btnLimparCard.style.display = 'none';
      if (valoresAtuais[item.nome]) delete valoresAtuais[item.nome];
    });

    var sincronizarVisualCard = function() {
      var rVal = inputReg.value.trim();
      var qVal = qtdInput.value.trim();
      var preenchido = !!(rVal || qVal);
      if (preenchido) {
        card.style.borderColor = '#0B5FFF';
        card.style.background = '#f0f7ff';
        btnLimparCard.style.display = 'inline-block';
        if (!qVal) qtdInput.value = '1';
      } else {
        card.style.borderColor = '#e2e8f0';
        card.style.background = '#fff';
        btnLimparCard.style.display = 'none';
      }
    };

    inputReg.addEventListener('input', sincronizarVisualCard);
    qtdInput.addEventListener('input', sincronizarVisualCard);

    regWrap.appendChild(inputReg);
    regWrap.appendChild(box);
    camposDiv.appendChild(regWrap);
    camposDiv.appendChild(qtdInput);
    camposDiv.appendChild(btnLimparCard);

    card.appendChild(nomeDiv);
    card.appendChild(camposDiv);
    container.appendChild(card);
  });

  var ms = document.getElementById('modalSelecao');
  if (ms) ms.style.display = 'flex';

  setTimeout(function () {
    var f = document.getElementById('buscaItemCatalogo');
    if (f) f.focus();
  }, 100);
}

function filtrarCatalogo() {
  var termo = (document.getElementById('buscaItemCatalogo').value || '').toLowerCase().trim();
  var container = document.getElementById('listaOpcoesContainer');
  if (!container) return;
  var cards = container.querySelectorAll('.card-item');

  cards.forEach(function (c) {
    var nome = c.dataset.nome || '';
    var inputReg = c.querySelector('.reg-input');
    var poolMatch = false;

    if (inputReg) {
      var nomeItem = inputReg.dataset.nome || '';
      var item = catalogoItens[selecaoAtualTipo].find(function (x) { return x.nome === nomeItem; });
      if (item && item.pool) {
        poolMatch = item.pool.some(function (p) { return p.toLowerCase().indexOf(termo) > -1; });
      }
    }

    var match = nome.indexOf(termo) > -1 || poolMatch || termo === '';
    c.style.display = match ? '' : 'none';
  });

  var currentCat = null;
  var visibleCount = 0;
  var all = Array.from(container.children);
  all.forEach(function (el) {
    if (el.classList.contains('categoria-titulo')) {
      if (currentCat !== null) { currentCat.style.display = visibleCount > 0 ? '' : 'none'; }
      currentCat = el;
      visibleCount = 0;
    } else if (el.classList.contains('card-item')) {
      if (el.style.display !== 'none') visibleCount++;
    }
  });
  if (currentCat) currentCat.style.display = visibleCount > 0 ? '' : 'none';
}

function fecharSeletor() {
  var ms = document.getElementById('modalSelecao');
  if (ms) ms.style.display = 'none';
}

function confirmarSelecao() {
  var container = document.getElementById('listaOpcoesContainer');
  if (!container) return;
  var resumo = [];
  var temp = {};

  container.querySelectorAll('input.qtd-input').forEach(function (inp) {
    var q = parseInt(inp.value, 10);
    var nome = inp.dataset.nome;
    var regInput = null;
    container.querySelectorAll('input.reg-input').forEach(function (r) {
      if (r.dataset.nome === nome) regInput = r;
    });
    var reg = regInput ? regInput.value.trim() : '';
    if ((q && q > 0) || reg) {
      temp[nome] = { qtd: q || '', reg: reg };
      resumo.push((q ? q + ' ' : '') + nome + (reg ? ' [Reg: ' + reg + ']' : ''));
    }
  });

  // Item não listado (livre)
  var nomeLivre = document.getElementById('nomeItemLivre');
  var qtdLivre = document.getElementById('qtdItemLivre');
  var regLivre = document.getElementById('regItemLivre');
  if (nomeLivre && nomeLivre.value.trim()) {
    var qL = qtdLivre ? qtdLivre.value : '';
    var rL = regLivre ? regLivre.value.trim() : '';
    resumo.push((qL ? qL + ' ' : '') + nomeLivre.value.trim() + (rL ? ' [Reg: ' + rL + ']' : ''));
  }

  var final = resumo.join(' | ') || 'Selecionar';

  if (linhaEditandoAtiva == 0) {
    valoresSelecionados[selecaoAtualTipo] = temp;
    var ba = document.getElementById(selecaoAtualTipo === 'armas' ? 'btnArmas' : 'btnEquip');
    if (ba) ba.innerText = final || 'Selecionar';
  } else {
    var idx = dadosBrutos.findIndex(function (x) { return x.row == linhaEditandoAtiva; });
    var targetId = idx > -1 ? dadosBrutos[idx].id : '';
    if (idx > -1) {
      if (selecaoAtualTipo === 'armas') dadosBrutos[idx].arma = final;
      else dadosBrutos[idx].equip = final;
      salvarLocal();
      render(dadosBrutos);
    }
    addFila({ type: 'texto', data: { row: linhaEditandoAtiva, col: selecaoAtualTipo === 'armas' ? 3 : 4, val: final, id: targetId } });
    setTimeout(processarFila, 100);
  }
  fecharSeletor();
}

function setupClickFora() {
  var m = document.getElementById('modal');
  if (m) m.addEventListener('click', function (e) { if (e.target.id === 'modal') { salvarAss(); } });
  var ms = document.getElementById('modalSelecao');
  if (ms) ms.addEventListener('click', function (e) { if (e.target.id === 'modalSelecao') { confirmarSelecao(); } });
  var mt = document.getElementById('modalTexto');
  if (mt) mt.addEventListener('click', function (e) { if (e.target.id === 'modalTexto') { salvarModalTexto(); } });
  var md = document.getElementById('modalData');
  if (md) md.addEventListener('click', function (e) { if (e.target.id === 'modalData') { salvarModalData(); } });
  var mn = document.getElementById('modalNum');
  if (mn) mn.addEventListener('click', function (e) { if (e.target.id === 'modalNum') { salvarModalNum(); } });
}

// ============================================================================
// INICIALIZAÇÃO NO DOMCONTENTLOADED
// ============================================================================
document.addEventListener('DOMContentLoaded', function () {
  initCanvas();
  setupClickFora();

  if (canvas) {
    canvas.addEventListener('mousedown', function (e) { des = true; last = pos(e); });
    canvas.addEventListener('touchstart', function (e) { e.preventDefault(); des = true; last = pos(e); }, { passive: false });
    canvas.addEventListener('mousemove', function (e) {
      if (!des) return;
      var p = pos(e);
      if (!ctx) return;
      ctx.beginPath();
      ctx.moveTo(last.x, last.y);
      ctx.lineTo(p.x, p.y);
      ctx.stroke();
      last = p;
    });
    canvas.addEventListener('touchmove', function (e) {
      e.preventDefault();
      if (!des) return;
      var p = pos(e);
      if (!ctx) return;
      ctx.beginPath();
      ctx.moveTo(last.x, last.y);
      ctx.lineTo(p.x, p.y);
      ctx.stroke();
      last = p;
    }, { passive: false });
    window.addEventListener('mouseup', function () { des = false; });
    window.addEventListener('touchend', function () { des = false; });
  }

  var el;
  el = document.getElementById('btnEntrar');
  if (el) el.addEventListener('click', verificarSenha);
  el = document.getElementById('senhaInput');
  if (el) el.addEventListener('keydown', function (e) { if (e.key === 'Enter') verificarSenha(); });
  el = document.getElementById('btnNovaAba');
  if (el) el.addEventListener('click', novaAba);
  el = document.getElementById('btnSalvarCopia');
  if (el) el.addEventListener('click', salvarCopiaLocal);
  el = document.getElementById('btnEnviarEmail');
  if (el) el.addEventListener('click', enviarCopiaEmail);
  el = document.getElementById('btnBackup');
  if (el) el.addEventListener('click', criarBackupDrive);
  el = document.getElementById('btnSinc');
  if (el) el.addEventListener('click', sincronizarManual);

  el = document.getElementById('btnArmas');
  if (el) el.addEventListener('click', function () { abrirSeletor('armas', false, 0); });
  el = document.getElementById('btnEquip');
  if (el) el.addEventListener('click', function () { abrirSeletor('equip', false, 0); });

  el = document.getElementById('btnF');
  if (el) el.addEventListener('click', function () { abrirNovo('F'); });
  el = document.getElementById('btnG');
  if (el) el.addEventListener('click', function () { abrirNovo('G'); });
  el = document.getElementById('btnI');
  if (el) el.addEventListener('click', function () { abrirNovo('I'); });
  el = document.getElementById('btnJ');
  if (el) el.addEventListener('click', function () { abrirNovo('J'); });

  el = document.getElementById('btnAdicionar');
  if (el) el.addEventListener('click', criarLinha);
  el = document.getElementById('btnFecharModal');
  if (el) el.addEventListener('click', fechar);
  el = document.getElementById('btnSalvarAss');
  if (el) el.addEventListener('click', salvarAss);

  el = document.getElementById('btnFecharSeletor');
  if (el) el.addEventListener('click', fecharSeletor);
  el = document.getElementById('btnConfirmarSelecao');
  if (el) el.addEventListener('click', confirmarSelecao);

  el = document.getElementById('btnFecharTexto');
  if (el) el.addEventListener('click', function () { var mt = document.getElementById('modalTexto'); if (mt) mt.style.display = 'none'; });
  el = document.getElementById('btnSalvarTexto');
  if (el) el.addEventListener('click', salvarModalTexto);

  el = document.getElementById('btnFecharData');
  if (el) el.addEventListener('click', function () { var md = document.getElementById('modalData'); if (md) md.style.display = 'none'; });
  el = document.getElementById('btnSalvarData');
  if (el) el.addEventListener('click', salvarModalData);

  el = document.getElementById('btnFecharNum');
  if (el) el.addEventListener('click', function () { var mn = document.getElementById('modalNum'); if (mn) mn.style.display = 'none'; });
  el = document.getElementById('btnSalvarNum');
  if (el) el.addEventListener('click', salvarModalNum);

  el = document.getElementById('btnCancelarApagar');
  if (el) el.addEventListener('click', function () { var m = document.getElementById('modalApagar'); if (m) m.style.display = 'none'; });
  el = document.getElementById('btnConfirmarApagar');
  if (el) el.addEventListener('click', confirmarApagar);

  el = document.getElementById('btnCancelarApagarAba');
  if (el) el.addEventListener('click', function () { var m = document.getElementById('modalApagarAba'); if (m) m.style.display = 'none'; });
  el = document.getElementById('btnConfirmarApagarAba');
  if (el) el.addEventListener('click', confirmarExclusaoAba);

  el = document.getElementById('btnCancelarConfirmEdit');
  if (el) el.addEventListener('click', function () { var m = document.getElementById('modalConfirmEdit'); if (m) m.style.display = 'none'; pendingEdit = null; });
  el = document.getElementById('btnConfirmarConfirmEdit');
  if (el) el.addEventListener('click', confirmarEdicaoComConfirmacao);

  el = document.getElementById('filtroMes');
  if (el) el.addEventListener('change', function () { loadDados(); });
  el = document.getElementById('buscaNome');
  if (el) el.addEventListener('input', filtrarPorNome);
  el = document.getElementById('buscaItemCatalogo');
  if (el) el.addEventListener('input', filtrarCatalogo);

  el = document.getElementById('btnConfigSheet');
  if (el) {
    el.addEventListener('click', async function() {
      const atual = getActiveSpreadsheetId();
      const input = document.getElementById('inputPlanilhaLinkOuId');
      if (input) input.value = atual;
      const statusDiv = document.getElementById('statusPlanilhaConexao');
      const listContainer = document.getElementById('listaPlanilhasDrive');
      
      const m = document.getElementById('modalSelecionarPlanilha');
      if (m) m.style.display = 'flex';

      if (!currentGoogleToken) {
        if (statusDiv) {
          statusDiv.style.color = '#f59e0b';
          statusDiv.innerText = '⚠️ Conecte a conta Google acima para listar suas planilhas do Drive automaticamente.';
        }
        if (listContainer) {
          listContainer.innerHTML = '<div style="text-align:center;padding:16px;color:#64748b;font-size:12px">Clique em <b>Conectar Google</b> para listar seus arquivos do Drive ou cole o link abaixo.</div>';
        }
      } else {
        if (statusDiv) {
          statusDiv.style.color = '#3b82f6';
          statusDiv.innerText = '🔍 Buscando planilhas no seu Google Drive...';
        }
        if (listContainer) {
          listContainer.innerHTML = '<div style="text-align:center;padding:16px;color:#64748b;font-size:12px">Carregando planilhas do Drive...</div>';
        }
        try {
          const files = await listSpreadsheetsFromDrive(currentGoogleToken);
          if (statusDiv) {
            statusDiv.style.color = '#10b981';
            statusDiv.innerText = '✔ Encontradas ' + files.length + ' planilhas no Google Drive:';
          }
          if (listContainer) {
            if (files.length === 0) {
              listContainer.innerHTML = '<div style="text-align:center;padding:16px;color:#64748b;font-size:12px">Nenhuma planilha encontrada no Drive. Cole o Link abaixo.</div>';
            } else {
              listContainer.innerHTML = '';
              files.forEach(f => {
                const item = document.createElement('div');
                const isSelected = (f.id === atual);
                item.style.cssText = `padding:10px 12px;margin-bottom:6px;background:${isSelected ? '#dbeafe' : '#fff'};border:1px solid ${isSelected ? '#3b82f6' : '#cbd5e1'};border-radius:8px;cursor:pointer;display:flex;justify-content:space-between;align-items:center;transition:background 0.2s`;
                item.innerHTML = `<div style="text-align:left"><div style="font-weight:800;font-size:13px;color:#0f172a">📊 ${f.name}</div><div style="font-size:10px;color:#64748b;font-family:monospace">${f.id}</div></div><button class="btn-s" style="padding:6px 12px;font-size:11px;background:${isSelected ? '#10b981' : '#3b82f6'}">${isSelected ? 'ATIVA' : 'Selecionar'}</button>`;
                item.addEventListener('click', () => {
                  if (input) input.value = f.id;
                  setActiveSpreadsheetId(f.id);
                  if (m) m.style.display = 'none';
                  status('✔ Planilha selecionada: ' + f.name);
                  loadAbas();
                  sincronizarManual();
                });
                listContainer.appendChild(item);
              });
            }
          }
        } catch (err) {
          if (statusDiv) {
            statusDiv.style.color = '#ef4444';
            statusDiv.innerText = 'Erro ao carregar Drive: ' + err.message;
          }
        }
      }
    });
  }

  el = document.getElementById('btnFecharSelecionarPlanilha');
  if (el) {
    el.addEventListener('click', function() {
      const m = document.getElementById('modalSelecionarPlanilha');
      if (m) m.style.display = 'none';
    });
  }

  el = document.getElementById('btnConfirmarSelecionarPlanilha');
  if (el) {
    el.addEventListener('click', async function() {
      const input = document.getElementById('inputPlanilhaLinkOuId');
      const statusDiv = document.getElementById('statusPlanilhaConexao');
      let valor = input ? input.value.trim() : '';
      if (!valor) {
        if (statusDiv) { statusDiv.style.color = '#ef4444'; statusDiv.innerText = 'Digite ou cole o Link ou ID da planilha.'; }
        return;
      }
      // Extrai ID caso tenha colado link completo do Google Sheets
      const match = valor.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (match && match[1]) valor = match[1];

      setActiveSpreadsheetId(valor);
      const m = document.getElementById('modalSelecionarPlanilha');
      if (m) m.style.display = 'none';

      status('⏳ Carregando dados da planilha ' + valor.substring(0, 8) + '...');
      // Se não estiver conectado no Google, convida a conectar
      if (!currentGoogleToken) {
        try {
          const res = await signInWithGoogle();
          currentUserGoogle = res.user;
          currentGoogleToken = res.accessToken;
          atualizarBotaoGoogle();
        } catch (e) {
          console.warn('Login não concluído:', e);
        }
      }
      loadAbas();
      sincronizarManual();
    });
  }

  el = document.getElementById('btnCancelarNovaAba');
  if (el) el.addEventListener('click', function() { var m = document.getElementById('modalNovaAba'); if (m) m.style.display = 'none'; });
  el = document.getElementById('btnConfirmarNovaAba');
  if (el) el.addEventListener('click', confirmarCriacaoNovaAba);
  el = document.getElementById('inputNomeNovaAba');
  if (el) el.addEventListener('keydown', function(e) { if (e.key === 'Enter') confirmarCriacaoNovaAba(); });

  el = document.getElementById('btnCancelarEditarAba');
  if (el) el.addEventListener('click', function() { var m = document.getElementById('modalEditarAba'); if (m) m.style.display = 'none'; });
  el = document.getElementById('btnConfirmarEditarAba');
  if (el) el.addEventListener('click', confirmarEdicaoNomeAba);
  el = document.getElementById('inputNomeEditarAba');
  if (el) el.addEventListener('keydown', function(e) { if (e.key === 'Enter') confirmarEdicaoNomeAba(); });

  el = document.getElementById('btnGoogleAuth');
  if (el) {
    el.addEventListener('click', async function() {
      try {
        if (currentGoogleToken) {
          await logoutGoogle();
          currentGoogleToken = null;
          currentUserGoogle = null;
          atualizarBotaoGoogle();
          status('Desconectado do Google');
        } else {
          status('⏳ Conectando conta Google...');
          const res = await signInWithGoogle();
          currentUserGoogle = res.user;
          currentGoogleToken = res.accessToken;
          atualizarBotaoGoogle();
          status('✅ Conectado com ' + (currentUserGoogle.email || 'Google'));
          loadAbas();
          sincronizarManual();
        }
      } catch (err) {
        console.error('Erro de autenticação Google:', err);
        status('❌ Erro no login Google: ' + (err.message || 'Falha'));
      }
    });
  }

  function atualizarBotaoGoogle() {
    const textEl = document.getElementById('googleAuthText');
    const btnGoogle = document.getElementById('btnGoogleAuth');
    if (!textEl || !btnGoogle) return;
    if (currentGoogleToken && currentUserGoogle) {
      textEl.innerText = (currentUserGoogle.displayName || currentUserGoogle.email?.split('@')[0] || 'Conectado');
      btnGoogle.style.background = '#dcfce7';
      btnGoogle.style.borderColor = '#10b981';
      btnGoogle.style.color = '#15803d';
      btnGoogle.title = 'Clique para desconectar (' + (currentUserGoogle.email || '') + ')';
    } else {
      textEl.innerText = 'Conectar Google';
      btnGoogle.style.background = '#fff';
      btnGoogle.style.borderColor = '#cbd5e1';
      btnGoogle.style.color = '#1e293b';
      btnGoogle.title = 'Clique para conectar sua conta Google e sincronizar com o Sheets';
    }
  }

  // Inicializa monitor de autenticação Google
  initGoogleAuth(
    function(user, token) {
      currentUserGoogle = user;
      currentGoogleToken = token;
      atualizarBotaoGoogle();
    },
    function() {
      currentUserGoogle = null;
      currentGoogleToken = null;
      atualizarBotaoGoogle();
    }
  );

  if (localStorage.getItem('authReserva') === '1') {
    var lo = document.getElementById('loginOverlay');
    if (lo) lo.style.display = 'none';
    inicializarApp();
  } else {
    if (tentativasErradas >= 5) verificarBloqueio();
  }
});
