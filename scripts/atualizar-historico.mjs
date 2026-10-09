// ============================================================
//  Registro de alteracoes dos chamados do SISCON
//
//  Compara a planilha atual (Solicitacoes.xlsx) com o ultimo estado
//  guardado em historico.json e registra o que mudou: prazo,
//  classificacao, prioridade, modulo e sistema. Tambem guarda o
//  "estado original" (primeira vez que cada chamado foi visto).
//
//  Uso (na pasta do projeto):
//    node scripts/atualizar-historico.mjs
//
//  Opcoes:
//    --xlsx=caminho.xlsx       planilha a ler (padrao: Solicitacoes.xlsx)
//    --historico=caminho.json  arquivo de historico (padrao: historico.json)
//
//  Rodar duas vezes com a mesma planilha nao duplica nada.
//  Nao precisa instalar nada: usa a biblioteca ja existente em vendor/.
// ============================================================

import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import fs from 'node:fs';
import crypto from 'node:crypto';

const require = createRequire(import.meta.url);
const AQUI = path.dirname(fileURLToPath(import.meta.url));
const RAIZ = path.resolve(AQUI, '..');
const XLSX = require(path.join(RAIZ, 'vendor', 'xlsx.full.min.js'));

// posicao das colunas no export "Solicitacoes" do SISCON
const COL = {
  protocolo: 0, incluidoPor: 1, solicitante: 2, classificacao: 3, sistema: 4,
  modulo: 5, resumo: 6, situacao: 7, prazoEntrega: 8, ultimaAlteracao: 9,
  fila: 10, prioridade: 11, dataInclusao: 12, prazoSla: 13, prazoDesejado: 14,
};

// campos que o historico acompanha
export const CAMPOS = ['prazo', 'classificacao', 'prioridade', 'modulo', 'sistema'];

// situacoes em que o chamado ja nao esta mais "em aberto" para efeito de prazo
const ENCERRADAS = ['liberado para cliente', 'resolvido', 'faturado', 'cancelado', 'homologando'];

const semAcento = (s) => String(s).normalize('NFD').replace(/[̀-ͯ]/g, '');
const chave = (v) => (v == null ? '' : semAcento(String(v).trim()).toLowerCase());
const txt = (v) => (v == null || String(v).trim() === '' ? null : String(v).trim());
const p2 = (n) => String(n).padStart(2, '0');

// qualquer data da planilha -> 'AAAA-MM-DD HH:MM' (horario de parede, sem fuso)
function fmtData(v) {
  if (v == null || v === '') return null;
  let d = null;
  if (v instanceof Date) d = v;
  else if (typeof v === 'number') {
    const c = XLSX.SSF.parse_date_code(v);
    if (!c) return null;
    d = new Date(c.y, c.m - 1, c.d, c.H || 0, c.M || 0, Math.round(c.S || 0));
  } else if (typeof v === 'string') {
    const m = v.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:[ ,]+(\d{1,2}):(\d{2}))?/);
    if (m) {
      const ano = m[3].length === 2 ? '20' + m[3] : m[3];
      d = new Date(+ano, +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0));
    }
  }
  if (!d || isNaN(d.getTime())) return null;
  d = new Date(Math.round(d.getTime() / 60000) * 60000); // tira ruido de segundos
  return `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())} ${p2(d.getHours())}:${p2(d.getMinutes())}`;
}

const comoUTC = (s) => Date.parse(s.replace(' ', 'T') + ':00Z');
const diasEntre = (a, b) => Math.round(((comoUTC(b) - comoUTC(a)) / 86400000) * 10) / 10;

// ISO local com fuso, ex.: 2026-10-09T10:57:00-03:00
export function agoraLocalISO(d = new Date()) {
  const off = -d.getTimezoneOffset();
  const sinal = off >= 0 ? '+' : '-';
  const a = Math.abs(off);
  return `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())}T${p2(d.getHours())}:${p2(d.getMinutes())}:${p2(d.getSeconds())}${sinal}${p2(Math.floor(a / 60))}:${p2(a % 60)}`;
}
const horaDeParede = (iso) => iso.slice(0, 16).replace('T', ' ');

export function lerPlanilha(buf) {
  const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const m = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  let h = m.findIndex((r) => r && r.some((c) => chave(c) === 'protocolo'));
  if (h < 0) h = 0;
  const out = {};
  for (let i = h + 1; i < m.length; i++) {
    const r = m[i];
    if (!r) continue;
    const p = txt(r[COL.protocolo]);
    if (!p) continue;
    out[p] = {
      resumo: txt(r[COL.resumo]) || '',
      sistema: txt(r[COL.sistema]),
      modulo: txt(r[COL.modulo]),
      classificacao: txt(r[COL.classificacao]),
      prioridade: txt(r[COL.prioridade]),
      situacao: txt(r[COL.situacao]),
      prazo: fmtData(r[COL.prazoSla]) || fmtData(r[COL.prazoEntrega]), // mesmo criterio do painel
    };
  }
  return out;
}

function iguais(campo, a, b) {
  if (campo === 'prazo' || campo === 'prioridade') return (a ?? null) === (b ?? null);
  return chave(a) === chave(b); // classificacao/modulo/sistema: ignora so maiuscula/acento
}

export function novoHistorico() {
  return { versao: 1, campos: CAMPOS, snapshots: [], chamados: {}, eventos: [] };
}

const instantaneo = (l) => ({ ...Object.fromEntries(CAMPOS.map((c) => [c, l[c] ?? null])), situacao: l.situacao ?? null });

// Aplica uma "foto" da planilha ao historico. Devolve um resumo do que mudou.
export function aplicarSnapshot(hist, linhas, em, rotulo = '') {
  const protos = Object.keys(linhas).sort();
  const hash = crypto
    .createHash('sha1')
    .update(JSON.stringify(protos.map((p) => [p, ...CAMPOS.map((c) => chave(linhas[p][c])), chave(linhas[p].situacao)])))
    .digest('hex')
    .slice(0, 12);

  const ultimo = hist.snapshots[hist.snapshots.length - 1];
  if (ultimo && ultimo.hash === hash) return { novo: false, total: protos.length };

  const agora = horaDeParede(em);
  const resumo = { novo: true, total: protos.length, chamadosNovos: 0, eventos: 0, porCampo: {} };

  for (const p of protos) {
    const l = linhas[p];
    const c = hist.chamados[p];
    if (!c) {
      hist.chamados[p] = {
        resumo: l.resumo,
        sistema: l.sistema,
        original: { ...instantaneo(l), vistoEm: em },
        atual: instantaneo(l),
      };
      resumo.chamadosNovos++;
      continue;
    }
    for (const campo of CAMPOS) {
      const antes = c.atual[campo] ?? null;
      const depois = l[campo] ?? null;
      if (iguais(campo, antes, depois)) continue;
      const ev = {
        em,
        protocolo: p,
        campo,
        antes,
        depois,
        situacaoAntes: c.atual.situacao,
        situacaoDepois: l.situacao ?? null,
      };
      if (campo === 'prazo') {
        if (antes && depois) ev.deltaDias = diasEntre(antes, depois);
        // o prazo antigo ja tinha estourado e o chamado ainda estava em aberto
        ev.jaVencido = !!(antes && antes < agora && !ENCERRADAS.includes(chave(c.atual.situacao)));
      }
      hist.eventos.push(ev);
      resumo.eventos++;
      resumo.porCampo[campo] = (resumo.porCampo[campo] || 0) + 1;
    }
    c.resumo = l.resumo || c.resumo;
    c.sistema = l.sistema || c.sistema;
    c.atual = instantaneo(l);
    delete c.ausenteDesde;
  }
  for (const p of Object.keys(hist.chamados)) {
    if (!linhas[p] && !hist.chamados[p].ausenteDesde) hist.chamados[p].ausenteDesde = em;
  }
  hist.snapshots.push({ em, rotulo, total: protos.length, hash });
  return resumo;
}

// uma linha por chamado e por evento, para o Git mostrar diferencas pequenas
export function serializar(h) {
  const j = (o) => JSON.stringify(o);
  return (
    '{\n"versao":' + h.versao +
    ',\n"campos":' + j(h.campos) +
    ',\n"snapshots":[\n' + h.snapshots.map(j).join(',\n') +
    '\n],\n"chamados":{\n' + Object.entries(h.chamados).map(([k, v]) => j(k) + ':' + j(v)).join(',\n') +
    '\n},\n"eventos":[\n' + h.eventos.map(j).join(',\n') +
    '\n]\n}\n'
  );
}

function arg(nome, padrao) {
  const a = process.argv.find((x) => x.startsWith('--' + nome + '='));
  return a ? a.slice(nome.length + 3) : padrao;
}

function main() {
  const xlsx = path.resolve(arg('xlsx', path.join(RAIZ, 'Solicitacoes.xlsx')));
  const arqHist = path.resolve(arg('historico', path.join(RAIZ, 'historico.json')));

  if (!fs.existsSync(xlsx)) {
    console.error('[erro] nao encontrei a planilha: ' + xlsx);
    process.exit(1);
  }
  const linhas = lerPlanilha(fs.readFileSync(xlsx));
  const total = Object.keys(linhas).length;
  if (!total) {
    console.error('[erro] nenhum chamado encontrado na planilha (coluna "Protocolo" existe?). Historico nao alterado.');
    process.exit(1);
  }

  const hist = fs.existsSync(arqHist) ? JSON.parse(fs.readFileSync(arqHist, 'utf8')) : novoHistorico();
  const anterior = Object.keys(hist.chamados).length;
  if (anterior && total < anterior * 0.8) {
    console.warn(`[aviso] a planilha tem ${total} chamados e o historico ja conhece ${anterior}. Confira se e o arquivo certo.`);
  }

  const r = aplicarSnapshot(hist, linhas, agoraLocalISO(), path.basename(xlsx));
  if (!r.novo) {
    console.log(`[historico] sem mudancas desde o ultimo registro (${total} chamados). Nada a gravar.`);
    return;
  }
  fs.writeFileSync(arqHist, serializar(hist), 'utf8');
  const detalhe = Object.entries(r.porCampo).map(([c, n]) => `${c}: ${n}`).join(', ');
  console.log(`[historico] gravado. ${total} chamados, ${r.chamadosNovos} novos, ${r.eventos} alteracoes registradas${detalhe ? ' (' + detalhe + ')' : ''}.`);
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) main();
