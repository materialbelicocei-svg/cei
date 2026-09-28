/**
 * Integração Google Sheets via Firebase Auth e REST API v4
 * Com trava de segurança por ID único e preservação estrita de integridade.
 */
import { initializeApp } from 'firebase/app';
import { getAuth, signInWithPopup, GoogleAuthProvider, onAuthStateChanged, signOut, type User } from 'firebase/auth';
import firebaseConfig from '../firebase-applet-config.json';

const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);

const provider = new GoogleAuthProvider();
provider.addScope('https://www.googleapis.com/auth/spreadsheets');
provider.addScope('https://www.googleapis.com/auth/drive.readonly');
provider.addScope('https://www.googleapis.com/auth/drive.file');
provider.setCustomParameters({
  prompt: 'consent'
});

let cachedAccessToken: string | null = null;
let signInRequest: Promise<{ user: User; accessToken: string }> | null = null;

// Id padrão da planilha ou configurável
export const DEFAULT_SHEET_ID = '1n0F-SlK4BuIjI8wzCEETfI-TXmjFBd7VFkisfzgVM1o';
const LEGACY_SHEET_ID = '19PV34qUCreU5xamyD-p-tssw0Ri5AHx8gD-vReF2m-0';

export function getActiveSpreadsheetId(): string {
  const storedId = localStorage.getItem('CEI_SHEET_ID');
  if (!storedId || storedId === LEGACY_SHEET_ID) {
    localStorage.setItem('CEI_SHEET_ID', DEFAULT_SHEET_ID);
    return DEFAULT_SHEET_ID;
  }
  return storedId;
}

export function setActiveSpreadsheetId(id: string) {
  if (id && id.trim()) {
    localStorage.setItem('CEI_SHEET_ID', id.trim());
  }
}

export const initGoogleAuth = (
  onSuccess?: (user: User, token: string) => void,
  onFailure?: () => void
) => {
  return onAuthStateChanged(auth, async (user: User | null) => {
    if (user && cachedAccessToken) {
      if (onSuccess) onSuccess(user, cachedAccessToken);
    } else {
      if (onFailure) onFailure();
    }
  });
};

export const signInWithGoogle = (): Promise<{ user: User; accessToken: string }> => {
  if (signInRequest) return signInRequest;

  signInRequest = (async () => {
    try {
      const result = await signInWithPopup(auth, provider);
      const credential = GoogleAuthProvider.credentialFromResult(result);
      if (!credential?.accessToken) {
        throw new Error('Falha ao obter token de acesso do Google.');
      }
      cachedAccessToken = credential.accessToken;
      return { user: result.user, accessToken: cachedAccessToken };
    } catch (error) {
      console.error('Erro no login do Google:', error);
      throw error;
    } finally {
      signInRequest = null;
    }
  })();

  return signInRequest;
};

export const getGoogleAccessToken = async (): Promise<string | null> => {
  return cachedAccessToken;
};

export const logoutGoogle = async () => {
  await signOut(auth);
  cachedAccessToken = null;
};

// ============================================================================
// GOOGLE SHEETS API V4 CALLS COM SUPORTE A ID ÚNICO E TRAVA DE SEGURANÇA
// ============================================================================

export interface SheetRowData {
  id: string; // ID único permanente
  row: number; // Linha atual na planilha física
  data: string;
  nome: string;
  arma: string;
  equip: string;
  mun: string;
  f: string;
  g: string;
  dataDev: string;
  i: string;
  j: string;
  obs: string;
  _temp?: boolean;
}

/**
 * Lista planilhas Google e arquivos XLSX diretamente do Google Drive do usuário
 */
export async function listSpreadsheetsFromDrive(accessToken: string): Promise<{ id: string; name: string; modifiedTime?: string; mimeType?: string }[]> {
  try {
    const q = encodeURIComponent("(mimeType='application/vnd.google-apps.spreadsheet' or mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or name contains 'cautela' or name contains 'livro' or name contains 'reserva' or name contains '.xlsx' or name contains '.xls') and trashed=false");
    const fields = encodeURIComponent('files(id, name, mimeType, modifiedTime)');
    const url = `https://www.googleapis.com/drive/v3/files?q=${q}&fields=${fields}&orderBy=modifiedTime desc&pageSize=100&supportsAllDrives=true&includeItemsFromAllDrives=true`;
    
    let res = await fetch(url, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });

    if (!res.ok) {
      // Fallback simples sem filtro restrito
      const fallbackUrl = `https://www.googleapis.com/drive/v3/files?q=trashed%3Dfalse&fields=${fields}&orderBy=modifiedTime desc&pageSize=50&supportsAllDrives=true&includeItemsFromAllDrives=true`;
      res = await fetch(fallbackUrl, {
        headers: { Authorization: `Bearer ${accessToken}` }
      });
    }

    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      console.warn('Falha na consulta Drive v3:', res.status, err);
      throw new Error(err.error?.message || `Erro ${res.status} ao acessar o Google Drive.`);
    }

    const data = await res.json();
    const files = (data.files || []).map((f: any) => ({
      id: f.id,
      name: f.name || 'Planilha Sem Título',
      mimeType: f.mimeType,
      modifiedTime: f.modifiedTime
    }));

    return files;
  } catch (err: any) {
    console.warn('Erro ao listar planilhas do Drive:', err);
    throw err;
  }
}

export async function uploadLocalSpreadsheetAsGoogleSheet(
  file: File,
  accessToken: string
): Promise<{ id: string; name: string }> {
  if (!file || !/\.xlsx$/i.test(file.name)) {
    throw new Error('Selecione um arquivo .xlsx válido.');
  }

  const boundary = 'cei_xlsx_' + Date.now().toString(36);
  const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const metadata = {
    name: file.name.replace(/\.xlsx$/i, '') || file.name,
    mimeType: 'application/vnd.google-apps.spreadsheet'
  };
  const body = new Blob([
    `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n`,
    `--${boundary}\r\nContent-Type: ${mimeType}\r\n\r\n`,
    file,
    `\r\n--${boundary}--`
  ], { type: `multipart/related; boundary=${boundary}` });

  const response = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,mimeType', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': `multipart/related; boundary=${boundary}`
    },
    body
  });

  if (!response.ok) {
    const error = await response.json().catch(() => ({}));
    throw new Error(error.error?.message || `Falha ao enviar o arquivo ao Drive (HTTP ${response.status}).`);
  }

  const createdFile = await response.json();
  if (!createdFile.id || createdFile.mimeType !== 'application/vnd.google-apps.spreadsheet') {
    throw new Error('O Drive não converteu o arquivo em uma planilha Google.');
  }

  return { id: createdFile.id, name: createdFile.name || metadata.name };
}

export async function restoreLocalSpreadsheetSignatures(
  spreadsheetId: string,
  sheetName: string,
  images: { row: number; col: number; dataUrl: string }[],
  accessToken: string
): Promise<number> {
  const data: { range: string; values: string[][] }[] = [];

  for (const image of images) {
    const formula = await uploadSignatureToDriveAndGetFormula(
      image.dataUrl,
      `assinatura_${sheetName}_${image.row}_${image.col}`,
      accessToken
    );
    if (!formula) {
      throw new Error(`Não foi possível enviar a assinatura da linha ${image.row}.`);
    }

    const colLetter = String.fromCharCode(64 + image.col);
    const safeName = sheetName.replace(/'/g, "''");
    data.push({
      range: `'${safeName}'!${colLetter}${image.row}`,
      values: [[formula]]
    });
  }

  if (data.length === 0) return 0;

  const url = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values:batchUpdate`;
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ valueInputOption: 'USER_ENTERED', data })
  });

  if (!response.ok) {
    const error = await response.json().catch(() => ({}));
    throw new Error(error.error?.message || 'Não foi possível vincular as assinaturas à nova planilha.');
  }

  return data.length;
}

/**
 * Gera um ID único e resistente a colisões
 */
export function generateRowId(): string {
  return 'CEI-' + Date.now().toString(36).toUpperCase() + '-' + Math.random().toString(36).substring(2, 7).toUpperCase();
}

/**
 * Busca os metadados da planilha e os nomes das abas (Sheets)
 */
export async function fetchSpreadsheetMetadata(spreadsheetId: string, accessToken: string) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}?fields=sheets.properties`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `Erro ${res.status} ao acessar a planilha`);
  }
  const data = await res.json();
  const sheets = (data.sheets || []).map((s: any) => s.properties?.title || '').filter(Boolean);
  return { sheets, raw: data };
}

function formatSheetDate(value: any): string {
  if (value === undefined || value === null || String(value).trim() === '') return '';
  const text = String(value).trim();
  const serial = typeof value === 'number' || /^\d{4,6}(?:\.\d+)?$/.test(text) ? Number(value) : NaN;

  if (Number.isFinite(serial) && serial >= 20000 && serial <= 100000) {
    const date = new Date(Date.UTC(1899, 11, 30) + Math.floor(serial) * 86400000);
    const day = String(date.getUTCDate()).padStart(2, '0');
    const month = String(date.getUTCMonth() + 1).padStart(2, '0');
    return `${day}/${month}/${date.getUTCFullYear()}`;
  }

  const isoDate = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoDate) return `${isoDate[3]}/${isoDate[2]}/${isoDate[1]}`;
  return text;
}

/**
 * Lê todas as linhas de uma aba (A1:L ou A2:L). Coluna L armazena o ID ÚNICO.
 */
export async function readSheetRows(
  spreadsheetId: string,
  sheetName: string,
  accessToken: string
): Promise<SheetRowData[]> {
  const safeName = sheetName.replace(/'/g, "''");
  const range = `'${safeName}'!A1:L`;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}?valueRenderOption=FORMULA`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    const msg = err.error?.message || `Erro HTTP ${res.status} ao acessar a planilha`;
    throw new Error(msg);
  }
  const data = await res.json();
  const allValues: any[][] = data.values || [];
  if (allValues.length === 0) return [];

  // Verifica se a primeira linha é cabeçalho
  let startIndex = 0;
  const firstRow = allValues[0] || [];
  const firstCell = String(firstRow[0] || '').toUpperCase();
  if (firstCell.includes('DATA') || firstCell.includes('RETIRADA') || firstCell.includes('HORA') || firstCell.includes('GRAD')) {
    startIndex = 1;
  }

  const rows: SheetRowData[] = [];
  const rowsNeedingId: { rowNumber: number; id: string }[] = [];

  for (let idx = startIndex; idx < allValues.length; idx++) {
    const cols = allValues[idx] || [];
    const rowNumber = idx + 1; // Linha física real na planilha 1-indexed

    const hasData = cols.some((c: any) => c !== undefined && c !== null && String(c).trim() !== '');
    if (!hasData) continue;

    // Coluna L (índice 11) armazena o ID Único
    let idUnico = cols[11] ? String(cols[11]).trim() : '';
    const nomeMilitar = cols[1] ? String(cols[1]).trim() : '';

    // Se o registro foi cancelado/apagado, mantém a linha física mas oculta da lista ativa para não deslocar índices
    if (nomeMilitar === '[REGISTRO CANCELADO]' || idUnico.endsWith('_CANCELADO')) {
      continue;
    }

    if (!idUnico) {
      idUnico = 'CEI-R' + rowNumber + '-' + (cols[0] ? String(cols[0]).replace(/[^0-9]/g, '') : '') + '-' + (cols[1] ? String(cols[1]).replace(/[^a-zA-Z0-9]/g, '').substring(0, 6) : '');
      rowsNeedingId.push({ rowNumber, id: idUnico });
    }

    rows.push({
      id: idUnico,
      row: rowNumber,
      data: formatSheetDate(cols[0]),
      nome: nomeMilitar,
      arma: cols[2] ? String(cols[2]).trim() : '',
      equip: cols[3] ? String(cols[3]).trim() : '',
      mun: cols[4] ? String(cols[4]).trim() : '',
      f: cols[5] ? String(cols[5]).trim() : '',
      g: cols[6] ? String(cols[6]).trim() : '',
      dataDev: formatSheetDate(cols[7]),
      i: cols[8] ? String(cols[8]).trim() : '',
      j: cols[9] ? String(cols[9]).trim() : '',
      obs: cols[10] ? String(cols[10]).trim() : ''
    });
  }

  // Se houver IDs gerados automaticamente, salva sem travar o usuário
  if (rowsNeedingId.length > 0) {
    (async () => {
      try {
        const batchData = rowsNeedingId.map(r => ({
          range: `'${safeName}'!L${r.rowNumber}`,
          values: [[r.id]]
        }));
        await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values:batchUpdate`, {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            valueInputOption: 'USER_ENTERED',
            data: batchData
          })
        });
      } catch (e) {
        console.warn('Persistência assíncrona de ID no Sheets ignorada:', e);
      }
    })();
  }

  return rows;
}

/**
 * Faz upload de uma imagem Base64 para o Google Drive e retorna a fórmula =IMAGE("url")
 * Assim o Google Sheets exibe a imagem real desenhada na célula em vez do texto base64.
 */
export async function uploadSignatureToDriveAndGetFormula(
  base64Data: string,
  nomeArquivo: string,
  accessToken: string
): Promise<string> {
  if (!base64Data || typeof base64Data !== 'string') return '';

  const raw = base64Data.trim();
  const looksLikeRawBase64 = /^[A-Za-z0-9+/=\r\n]+$/.test(raw) && raw.length > 200 && !raw.includes('http') && !raw.includes('=IMAGE');
  const normalized = looksLikeRawBase64 ? `data:image/png;base64,${raw}` : raw;

  if (normalized.startsWith('=IMAGE(') || normalized.startsWith('=image(') || normalized.startsWith('http://') || normalized.startsWith('https://')) {
    return normalized;
  }

  if (!normalized.startsWith('data:image/')) {
    return '';
  }

  try {
    const parts = normalized.split(',');
    if (parts.length < 2) return '';
    const mimeMatch = parts[0].match(/:(.*?);/);
    const mimeType = mimeMatch ? mimeMatch[1] : 'image/png';
    const byteCharacters = atob(parts[1]);
    const byteNumbers = new Uint8Array(byteCharacters.length);
    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i);
    }
    const blob = new Blob([byteNumbers], { type: mimeType });

    // Upload direto binário para o Google Drive
    const fileName = (nomeArquivo || 'assinatura_cautela') + '_' + Date.now() + '.png';
    const uploadUrl = `https://www.googleapis.com/upload/drive/v3/files?uploadType=media`;

    const uploadRes = await fetch(uploadUrl, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': mimeType
      },
      body: blob
    });

    if (!uploadRes.ok) {
      const errTxt = await uploadRes.text();
      console.warn('Falha no upload da imagem no Drive:', uploadRes.status, errTxt);
      return '';
    }

    const fileData = await uploadRes.json();
    const fileId = fileData.id;

    if (fileId) {
      // 1. Define o nome do arquivo
      try {
        await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}`, {
          method: 'PATCH',
          headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({ name: fileName })
        });
      } catch (e) {}

      const imageUrl = `https://drive.google.com/thumbnail?id=${fileId}&sz=w1000`;
      return `=IMAGE("${imageUrl}", 1)`;
    }
  } catch (err) {
    console.warn('Erro ao enviar imagem ao Drive:', err);
  }

  return '';
}

/**
 * Insere uma nova linha com ID Único exclusivo no topo da planilha (imediatamente na linha 2)
 */
export async function insertRowInSheet(
  spreadsheetId: string,
  sheetName: string,
  rowItem: {
    id?: string;
    data: string;
    nome: string;
    arma: string;
    equip: string;
    mun: string;
    dataDev: string;
    assF: string;
    assG: string;
    assI: string;
    assJ: string;
    obs: string;
  },
  accessToken: string
): Promise<{ ok: boolean; row: number; id: string }> {
  const meta = await fetchSpreadsheetMetadata(spreadsheetId, accessToken);
  const targetSheet = meta.raw.sheets?.find((s: any) => s.properties?.title === sheetName);
  const sheetIdNum = targetSheet?.properties?.sheetId ?? 0;
  const rowId = rowItem.id || generateRowId();

  // Converte assinaturas Base64 em fórmulas =IMAGE() no Google Sheets
  let assFVal = rowItem.assF || '';
  let assGVal = rowItem.assG || '';
  let assIVal = rowItem.assI || '';
  let assJVal = rowItem.assJ || '';

  if (assFVal) assFVal = await uploadSignatureToDriveAndGetFormula(assFVal, `${rowItem.nome}_assF`, accessToken);
  if (assGVal) assGVal = await uploadSignatureToDriveAndGetFormula(assGVal, `${rowItem.nome}_assG`, accessToken);
  if (assIVal) assIVal = await uploadSignatureToDriveAndGetFormula(assIVal, `${rowItem.nome}_assI`, accessToken);
  if (assJVal) assJVal = await uploadSignatureToDriveAndGetFormula(assJVal, `${rowItem.nome}_assJ`, accessToken);

  // Inserir 1 linha vazia na posição 2 (índice 1) para não sobrescrever nenhum dado existente!
  const batchUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}:batchUpdate`;
  const insertRequest = {
    requests: [
      {
        insertDimension: {
          range: {
            sheetId: sheetIdNum,
            dimension: 'ROWS',
            startIndex: 1, // Linha 2 física
            endIndex: 2
          },
          inheritFromBefore: false
        }
      }
    ]
  };

  const batchRes = await fetch(batchUrl, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(insertRequest)
  });

  if (!batchRes.ok) {
    console.warn('Fallback para appendRow...');
    return appendRowFallback(spreadsheetId, sheetName, { ...rowItem, id: rowId, assF: assFVal, assG: assGVal, assI: assIVal, assJ: assJVal }, accessToken);
  }

  // Grava os valores exatamente na linha 2 criada, incluindo o ID na coluna L (A2:L2)
  const range = `'${sheetName.replace(/'/g, "''")}'!A2:L2`;
  const updateUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`;
  
  const values = [
    [
      rowItem.data || '',
      rowItem.nome || '',
      rowItem.arma || '',
      rowItem.equip || '',
      rowItem.mun || '',
      assFVal,
      assGVal,
      rowItem.dataDev || '',
      assIVal,
      assJVal,
      rowItem.obs || '',
      rowId
    ]
  ];

  const putRes = await fetch(updateUrl, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ values })
  });

  if (!putRes.ok) {
    const err = await putRes.json().catch(() => ({}));
    throw new Error(err.error?.message || 'Falha ao gravar valores na linha 2');
  }

  return { ok: true, row: 2, id: rowId };
}

async function appendRowFallback(
  spreadsheetId: string,
  sheetName: string,
  rowItem: any,
  accessToken: string
) {
  const rowId = rowItem.id || generateRowId();
  const range = `'${sheetName.replace(/'/g, "''")}'!A:L`;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`;
  const values = [
    [
      rowItem.data || '',
      rowItem.nome || '',
      rowItem.arma || '',
      rowItem.equip || '',
      rowItem.mun || '',
      rowItem.assF || '',
      rowItem.assG || '',
      rowItem.dataDev || '',
      rowItem.assI || '',
      rowItem.assJ || '',
      rowItem.obs || '',
      rowId
    ]
  ];

  const res = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ values })
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || 'Falha ao adicionar linha no Sheets');
  }
  const result = await res.json();
  const updatedRange = result.updates?.updatedRange || 'A2';
  const match = updatedRange.match(/!A(\d+)/);
  const row = match ? parseInt(match[1], 10) : 2;
  return { ok: true, row, id: rowId };
}

/**
 * Encontra a linha atual de um registro no Sheets buscando pelo ID ÚNICO
 * TRAVA DE SEGURANÇA: Garante que nunca alterará a linha de outro registro!
 */
export async function findRowById(
  spreadsheetId: string,
  sheetName: string,
  targetId: string,
  accessToken: string
): Promise<number | null> {
  if (!targetId) return null;
  const range = `'${sheetName.replace(/'/g, "''")}'!L2:L`;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}?valueRenderOption=FORMATTED_VALUE`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });
  if (!res.ok) return null;
  const data = await res.json();
  const values: any[][] = data.values || [];
  for (let i = 0; i < values.length; i++) {
    const val = values[i] && values[i][0] ? String(values[i][0]).trim() : '';
    if (val === targetId) {
      return i + 2;
    }
  }
  return null;
}

/**
 * Atualiza uma célula específica com conferência de ID para nunca sobrescrever outro registro
 */
export async function updateCellWithIdCheck(
  spreadsheetId: string,
  sheetName: string,
  targetId: string,
  fallbackRow: number,
  colNumber: number,
  value: string,
  accessToken: string
) {
  // Trava de segurança: localizar a linha exata do ID na planilha
  let actualRow = fallbackRow;
  if (targetId) {
    const found = await findRowById(spreadsheetId, sheetName, targetId, accessToken);
    if (found) actualRow = found;
  }

  // Se for coluna de assinatura e for base64, converte em imagem do Google Drive com fórmula =IMAGE()
  let finalVal = value;
  if ((colNumber === 6 || colNumber === 7 || colNumber === 9 || colNumber === 10) && value) {
    finalVal = await uploadSignatureToDriveAndGetFormula(value, `ass_col${colNumber}_row${actualRow}`, accessToken);
    if (!finalVal) {
      throw new Error('Não foi possível enviar a assinatura; a célula não foi alterada.');
    }
  }

  const colLetter = String.fromCharCode(64 + colNumber);
  const cellRange = `'${sheetName.replace(/'/g, "''")}'!${colLetter}${actualRow}`;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(cellRange)}?valueInputOption=USER_ENTERED`;

  const res = await fetch(url, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ values: [[finalVal]] })
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `Erro ao atualizar célula ${cellRange}`);
  }
  return true;
}

/**
 * Apaga/cancela uma linha no Google Sheets com conferência de ID sem deslocar índices físicos
 */
export async function deleteRowWithIdCheck(
  spreadsheetId: string,
  sheetName: string,
  targetId: string,
  fallbackRow: number,
  accessToken: string
) {
  let actualRow = fallbackRow;
  if (targetId) {
    const found = await findRowById(spreadsheetId, sheetName, targetId, accessToken);
    if (found) actualRow = found;
  }

  // Em vez de deleteDimension (que desloca índices físicos de todas as linhas abaixo),
  // limpamos e marcamos a linha com [REGISTRO CANCELADO], mantendo o espaço físico e a estabilidade total dos índices!
  const range = `'${sheetName.replace(/'/g, "''")}'!A${actualRow}:L${actualRow}`;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`;

  const canceledValues = [
    [
      '', // A: Data Ret
      '[REGISTRO CANCELADO]', // B: Nome
      '', // C: Armas
      '', // D: Equip
      '', // E: Mun
      '', // F: Ass F
      '', // G: Ass G
      '', // H: Data Dev
      '', // I: Ass I
      '', // J: Ass J
      'Registro apagado pelo operador', // K: Obs
      targetId ? `${targetId}_CANCELADO` : `CEI-R${actualRow}_CANCELADO` // L: ID
    ]
  ];

  const res = await fetch(url, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ values: canceledValues })
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `Erro ao marcar linha ${actualRow} como cancelada no Sheets`);
  }
  return true;
}

/**
 * Cria uma nova aba na planilha do Google Sheets
 */
export async function createSheetTab(
  spreadsheetId: string,
  title: string,
  accessToken: string
) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}:batchUpdate`;
  const request = {
    requests: [
      {
        addSheet: {
          properties: {
            title: title
          }
        }
      }
    ]
  };

  const res = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(request)
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `Erro ao criar nova aba ${title}`);
  }

  // Cabeçalho incluindo a coluna de ID
  const headerRange = `'${title.replace(/'/g, "''")}'!A1:L1`;
  const headerUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(headerRange)}?valueInputOption=USER_ENTERED`;
  await fetch(headerUrl, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      values: [[
        'DATA RETIRADA',
        'GRADUAÇÃO/NOME',
        'ARMAMENTOS',
        'EQUIPAMENTOS',
        'MUNIÇÃO',
        'ASS. USUÁRIO RET.',
        'ASS. ARMEIRO RET.',
        'DATA DEV.',
        'ASS. USUÁRIO DEV.',
        'ASS. ARMEIRO DEV.',
        'OBS',
        'ID_REGISTRO'
      ]]
    })
  }).catch(() => {});

  return true;
}

/**
 * Renomeia uma aba no Google Sheets
 */
export async function renameSheetTab(
  spreadsheetId: string,
  oldTitle: string,
  newTitle: string,
  accessToken: string
) {
  const meta = await fetchSpreadsheetMetadata(spreadsheetId, accessToken);
  const targetSheet = meta.raw.sheets?.find((s: any) => s.properties?.title === oldTitle);
  if (!targetSheet) throw new Error(`Aba "${oldTitle}" não encontrada no Sheets`);
  const sheetIdNum = targetSheet.properties.sheetId;

  const url = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}:batchUpdate`;
  const request = {
    requests: [
      {
        updateSheetProperties: {
          properties: {
            sheetId: sheetIdNum,
            title: newTitle
          },
          fields: 'title'
        }
      }
    ]
  };

  const res = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(request)
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `Erro ao renomear aba no Sheets`);
  }
  return true;
}

/**
 * Deleta uma aba no Google Sheets
 */
export async function deleteSheetTab(
  spreadsheetId: string,
  title: string,
  accessToken: string
) {
  const meta = await fetchSpreadsheetMetadata(spreadsheetId, accessToken);
  const targetSheet = meta.raw.sheets?.find((s: any) => s.properties?.title === title);
  if (!targetSheet) throw new Error(`Aba "${title}" não encontrada no Sheets`);
  const sheetIdNum = targetSheet.properties.sheetId;

  const url = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}:batchUpdate`;
  const request = {
    requests: [
      {
        deleteSheet: {
          sheetId: sheetIdNum
        }
      }
    ]
  };

  const res = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(request)
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `Erro ao apagar aba no Sheets`);
  }
  return true;
}
