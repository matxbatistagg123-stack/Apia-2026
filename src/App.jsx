import React, { useState, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, PieChart, Pie, Cell, AreaChart, Area,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  RadialBarChart, RadialBar, LineChart, Line
} from "recharts";
import {
  LayoutDashboard, Users, CalendarX, Clock, Award,
  FileBarChart, Download, Search, Upload, LogOut, Settings,
  AlertTriangle, CheckCircle2, ChevronRight, UserCheck, UserX, Plane, HeartPulse, X, Activity,
  Menu, Palmtree, Calendar
} from "lucide-react";
import { auth, db } from "./firebase";
import { onAuthStateChanged, signOut, signInWithEmailAndPassword, createUserWithEmailAndPassword } from "firebase/auth";
import { doc, getDoc, setDoc, onSnapshot, serverTimestamp } from "firebase/firestore";

/* ====== IDENTIDADE VISUAL APIA ====== */
const NAVY = "#043A66";
const NAVY_DARK = "#032A4A";
const NAVY_SOFT = "#EAF1F7";
const NAVY_MED = "#1E6CA8";
const OK = "#1D9E75";
const WARN = "#EF9F27";
const DANGER = "#E24B4A";
const LIMITE_EXTRA = 40 * 60; // 40h em minutos

/* ====== HELPERS ====== */
const hhmmToMin = (v) => {
  if (v == null) return 0;
  const m = String(v).trim().match(/^(\d{1,3}):(\d{2})/);
  return m ? parseInt(m[1], 10) * 60 + parseInt(m[2], 10) : 0;
};
const minToHHMM = (min) => `${Math.floor(min / 60)}h${String(Math.round(min % 60)).padStart(2, "0")}`;
const norm = (s) => String(s ?? "").trim();
const isVazio = (v) => { const s = norm(v).toLowerCase(); return !s || s === "nan" || s === "00:00"; };
const safeGet = (obj, key) => {
  if (!obj || typeof key !== "string") return undefined;
  if (key === "__proto__" || key === "constructor" || key === "prototype") return undefined;
  return Reflect.get(obj, key);
};

const normalizeToLocalMidnight = (d) => {
  if (!d || isNaN(d.getTime())) return null;
  if (d.getUTCHours() === 0 && d.getUTCMinutes() === 0 && d.getUTCSeconds() === 0) {
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
};

const ensureDate = (v) => {
  if (!v) return null;
  if (v instanceof Date) return normalizeToLocalMidnight(v);
  if (typeof v.toDate === "function") return normalizeToLocalMidnight(v.toDate());
  if (v.seconds != null) return normalizeToLocalMidnight(new Date(v.seconds * 1000));
  const s = norm(v);
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (m) return normalizeToLocalMidnight(new Date(parseInt(m[3],10), parseInt(m[2],10)-1, parseInt(m[1],10)));
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : normalizeToLocalMidnight(d);
};

const compressAndSavePhoto = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement("canvas");
        const MAX_WIDTH = 180;
        const MAX_HEIGHT = 180;
        let width = img.width;
        let height = img.height;

        if (width > height) {
          if (width > MAX_WIDTH) {
            height *= MAX_WIDTH / width;
            width = MAX_WIDTH;
          }
        } else {
          if (height > MAX_HEIGHT) {
            width *= MAX_HEIGHT / height;
            height = MAX_HEIGHT;
          }
        }
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, width, height);
        const dataUrl = canvas.toDataURL("image/jpeg", 0.7);
        resolve(dataUrl);
      };
      img.onerror = () => reject(new Error("Erro ao carregar a imagem."));
      img.src = e.target.result;
    };
    reader.onerror = () => reject(new Error("Erro ao ler o arquivo."));
    reader.readAsDataURL(file);
  });
};

const translations = {
  pt: {
    anexePlanilha: "Anexe a planilha de ponto",
    arrasteArquivo: "Arraste o arquivo ",
    selecionarPlanilha: "Selecionar planilha",
    verDemonstracao: "Ver com dados de demonstração",
    acessoRestrito: "Acesso restrito ao Administrador.",
    nenhumaPlanilha: "Nenhuma planilha importada ainda",
    distribuicaoEfetivo: "Distribuição do efetivo",
    faltas: "Faltas",
    faltasMin: "faltas",
    horasPerdidas: "horas perdidas",
    pessoas: "pessoas",
    nenhumaFaltaPeriodo: "Nenhuma falta no período.",
    verTodasAs: "Ver todas as ",
    rankingAbonos: "Ranking de abonos",
    nenhumAbonoPeriodo: "Nenhum abono no período.",
    verRankingCompleto: "Ver ranking completo ",
    horasExtrasFunc: "Horas extras por funcionário",
    limiteMensal: "Limite 40h/mês",
    totalLabel: "Total: ",
    nenhumaExtraPeriodo: "Nenhuma hora extra no período.",
    verTodosOs: "Ver todos os ",
    relatorioPronto: "Relatório pronto para apresentação — indicadores calculados automaticamente.",
    ninguemSituacaoPeriodo: "Ninguém nesta situação no período.",
    ninguemSituacao: "Ninguém nesta situação.",
    totalFaltas: "Total de faltas",
    horasPerdidasCard: "Horas perdidas",
    pessoasFalta: "Pessoas com falta",
    faltasColaborador: "Faltas por colaborador",
    nenhumaFaltaPeriodoCelebration: "Nenhuma falta no período. 🎉",
    cliqueMotivo: "Clique em um motivo para ver os colaboradores",
    totalHoras: "Total de horas",
    funcionarios: "Funcionários",
    emAtencao: "Em atenção (≥75%)",
    noLimite: "No limite (40h)",
    ausenciasTotais: "Ausências totais",
    faltasAtestados: "faltas + atestados",
    noPeriodo: "no período",
    atestadosMedicos: "Atestados médicos",
    evolucaoAbsenteismo: "Evolução do absenteísmo por semana",
    detalhamentoSemana: "Detalhamento por semana",
    semana: "Semana",
    atestados: "Atestados",
    diasUteis: "Dias úteis",
    semDatasValidas: "Sem datas válidas na planilha para calcular semanas.",
    resumoGeral: "Resumo geral do período",
    todosIndicadores: "Todos os indicadores consolidados — pronto para copiar.",
    menu: "Menu",
    sistemaGestao: "Sistema de Gestão de Ponto",
    login: "Login",
    acesseConta: "Acesse sua conta para continuar",
    usuario: "Usuário",
    senha: "Senha",
    esqueceuSenha: "Esqueceu a senha?",
    entrar: "Entrar ",
    cadastreSe: "Cadastre-se",
    verificacao: "Verificação",
    digiteCodigo: "Digite o código de verificação em duas etapas.",
    verificarEntrar: "Verificar e entrar",
    criarConta: "Criar conta",
    cadastreAcesso: "Cadastre seu acesso ao sistema",
    nome: "Nome",
    confirmarSenha: "Confirmar senha",
    perfilAcesso: "Perfil de acesso",
    cadastrar: "Cadastrar",
    fazerLogin: "Fazer login",
    buscar: "Buscar",
    ativos: "Ativos",
    ferias: "Férias",
    afastadosInss: "Afastados INSS",
    demitidos: "Demitidos",
    novaplanilha: "Nova planilha",
    sair: "Sair",
    registros: "registros",
    colaboradores: "colaboradores",
    saudeGeral: "SAÚDE GERAL",
    todas: "Todas",
    copiarResumo: "Copiar resumo para slides",
    copiarResumoCompleto: "Copiar resumo completo para slides",
    indiceGeral: "Índice geral",
    voltar: "Voltar",
    meuPerfil: "Meu Perfil",
    cliqueFotoAlterar: "Clique na foto para alterar",
    nomeCompleto: "Nome Completo",
    gerente: "Gerente",
    administrador: "Administrador",
    cancelar: "Cancelar",
    salvarAlteracoes: "Salvar Alterações",
    codigoSegurancaAdmin: "Código de Segurança Administrador",
    digiteCodigoPonto: "Digite o código de segurança",
    podeFazerMais: "Pode fazer mais ",
    limiteAtingido: "Limite atingido",
    buscarColaboradorCargo: "Buscar por colaborador ou cargo...",
    totalExtra: "Total extra: ",
    dia: "Dia ",
    nenhumRegistroExtra: "Nenhum registro de hora extra.",
    nenhumaFeriasImportada: "Nenhuma planilha de férias importada",
    importePlanilhaFeriasAbaImportar: "Importe a planilha de férias na aba Importar para visualizar os dados.",
    irParaImportar: "Ir para Importar",
    saldo: "Saldo:",
    limite: "Limite:",
    totalFuncionarios: "Total de funcionários",
    vencemEm30Dias: "Vencem em 30 dias",
    planilhaLabel: "Planilha: ",
    anexePlanilhaFerias: "Anexe a planilha de férias",
    arrasteArquivoFerias: "Arraste o arquivo .xlsx ou clique. Filtra automaticamente para obras 948 e 935.",
    selecionarPlanilhaFerias: "Selecionar planilha de férias",
    feriasVencidas: "Férias vencidas",
    buscarNomeFuncao: "Buscar por nome ou função...",
    ninguemEncontradoFiltro: "Ninguém encontrado com esse filtro.",
    nenhumFuncionarioEncontradoSecoes: "Nenhum funcionário encontrado nas seções 948 / 935.",
    arrasteOuCliqueSelecionar: "Arraste o arquivo .xlsx ou clique para selecionar."
  }
};

const t = (key) => {
  const langSet = Reflect.get(translations, "pt") || translations.pt;
  return Reflect.get(langSet, key) || key;
};

/* ====== PARSER ====== */
function parseSheet(rows) {
  const keys = Object.keys(rows[0] || {});
  const find = (...c) => {
    // Fase 1: prefere colunas que COMEÇAM com o termo (evita CODFUNCAO ao buscar FUNCAO)
    for (const x of c) {
      const match = keys.find((k) => k.toUpperCase().replace(/\s/g, "").startsWith(x));
      if (match) return match;
    }
    // Fase 2: fallback para colunas que CONTÊM o termo
    for (const x of c) {
      const match = keys.find((k) => k.toUpperCase().replace(/\s/g, "").includes(x));
      if (match) return match;
    }
    return undefined;
  };
  const K = {
    chapa: find("CHAPA"), nome: find("NOME"), obra: find("SEDE_OBRA", "OBRA"),
    funcao: find("FUNCAO_EPOCA", "FUNCAO"), situacao: find("CODSITUACAO", "SITUACAO"),
    falta: find("FALTA"), abono: find("ABONO"), extra: find("EXTRA"), data: find("DATA"),
  };
  const fmtData = (v) => {
    if (v instanceof Date) return `${String(v.getDate()).padStart(2, "0")}/${String(v.getMonth() + 1).padStart(2, "0")}`;
    const s = norm(v);
    const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})/);
    if (m) return `${m[1].padStart(2, "0")}/${m[2].padStart(2, "0")}`;
    return s;
  };
  const getTs = (v) => {
    if (v instanceof Date) return v.getTime();
    const s = norm(v);
    const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (m) return new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10)).getTime();
    const m2 = s.match(/^(\d{1,2})[\/\-](\d{1,2})/);
    if (m2) return new Date(new Date().getFullYear(), parseInt(m2[2], 10) - 1, parseInt(m2[1], 10)).getTime();
    return Date.parse(s) || 0;
  };
  return rows.filter((r) => norm(safeGet(r, K.chapa)) && norm(safeGet(r, K.nome))).map((r) => ({
    chapa: norm(safeGet(r, K.chapa)), nome: norm(safeGet(r, K.nome)), obra: norm(safeGet(r, K.obra)),
    funcao: norm(safeGet(r, K.funcao)) || "—", situacao: norm(safeGet(r, K.situacao)).toUpperCase(),
    faltaMin: hhmmToMin(safeGet(r, K.falta)), extraMin: hhmmToMin(safeGet(r, K.extra)),
    abono: isVazio(safeGet(r, K.abono)) ? null : norm(safeGet(r, K.abono)),
    data: fmtData(safeGet(r, K.data)), ts: getTs(safeGet(r, K.data)),
  }));
}

/* ====== PARSER FÉRIAS ====== */
function parseFeriasSheet(rows) {
  const keys = Object.keys(rows[0] || {});
  const find = (...c) => {
    for (const x of c) {
      const match = keys.find((k) => k.toUpperCase().replace(/\s/g, "").startsWith(x));
      if (match) return match;
    }
    for (const x of c) {
      const match = keys.find((k) => k.toUpperCase().replace(/\s/g, "").includes(x));
      if (match) return match;
    }
    return undefined;
  };
  const K = {
    nome: find("NOME"),
    secao: find("SECAO", "SEÇÃO", "SEÇAO", "SECÃO"),
    funcao: find("FUNCAO_EPOCA", "FUNCAO", "FUNÇÃO"),
    saldo: find("SALDO"),
    dataLimite: find("DATA_LIMITE", "DATALIMITE", "LIMITE"),
  };
  const parseDataLimite = (v) => {
    if (v instanceof Date) return normalizeToLocalMidnight(v);
    if (typeof v === "number" || (!isNaN(v) && !isNaN(parseFloat(v)))) {
      const num = Number(v);
      if (num > 30000 && num < 60000) {
        return normalizeToLocalMidnight(new Date(Math.round((num - 25569) * 86400 * 1000)));
      }
    }
    const s = norm(v);
    const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (m) return normalizeToLocalMidnight(new Date(parseInt(m[3],10), parseInt(m[2],10)-1, parseInt(m[1],10)));
    const m2 = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (m2) return normalizeToLocalMidnight(new Date(parseInt(m2[1],10), parseInt(m2[2],10)-1, parseInt(m2[3],10)));
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : normalizeToLocalMidnight(d);
  };
  return rows
    .filter((r) => {
      const nome = norm(safeGet(r, K.nome));
      if (!nome) return false;
      const secao = norm(safeGet(r, K.secao));
      return secao.includes("948") || secao.includes("935");
    })
    .map((r) => {
      const dl = parseDataLimite(safeGet(r, K.dataLimite));
      return {
        nome: norm(safeGet(r, K.nome)),
        secao: norm(safeGet(r, K.secao)),
        funcao: norm(safeGet(r, K.funcao)) || "—",
        saldo: norm(safeGet(r, K.saldo)) || "—",
        dataLimite: dl,
        dataLimiteStr: dl ? `${String(dl.getDate()).padStart(2,"0")}/${String(dl.getMonth()+1).padStart(2,"0")}/${dl.getFullYear()}` : norm(safeGet(r, K.dataLimite)) || "—",
      };
    })
    .sort((a, b) => {
      if (!a.dataLimite && !b.dataLimite) return 0;
      if (!a.dataLimite) return 1;
      if (!b.dataLimite) return -1;
      return a.dataLimite.getTime() - b.dataLimite.getTime();
    });
}

/* ====== AGREGAÇÕES ====== */
function aggregate(data) {
  const total = data.length;
  const chapas = new Map();
  data.forEach((d) => {
    if (!chapas.has(d.chapa)) chapas.set(d.chapa, []);
    chapas.get(d.chapa).push(d);
  });

  // EFETIVO por situação (contagem de PESSOAS únicas) + listas
  const sitPessoa = new Map([["A", new Set()], ["F", new Set()], ["P", new Set()], ["D", new Set()]]);
  const sitListas = new Map([["A", new Map()], ["F", new Map()], ["P", new Map()], ["D", new Map()]]);
  data.forEach((d) => {
    if (sitPessoa.has(d.situacao)) {
      sitPessoa.get(d.situacao).add(d.chapa);
      sitListas.get(d.situacao).set(d.chapa, { nome: d.nome, funcao: d.funcao, obra: d.obra });
    }
  });
  const efetivo = {
    A: sitPessoa.get("A").size, F: sitPessoa.get("F").size, P: sitPessoa.get("P").size, D: sitPessoa.get("D").size,
  };
  const efetivoListas = {
    A: Array.from(sitListas.get("A").values()), F: Array.from(sitListas.get("F").values()),
    P: Array.from(sitListas.get("P").values()), D: Array.from(sitListas.get("D").values()),
  };
  const funcionarios = chapas.size;

  // FALTAS por pessoa
  const faltasPorPessoa = Array.from(chapas.values())
    .map((regs) => {
      const fs = regs.filter((r) => r.faltaMin > 0);
      const firstReg = regs.at(0);
      return { nome: firstReg.nome, funcao: firstReg.funcao, obra: firstReg.obra,
        qtd: fs.length, horas: fs.reduce((s, r) => s + r.faltaMin, 0) };
    })
    .filter((p) => p.qtd > 0)
    .sort((a, b) => b.qtd - a.qtd || b.horas - a.horas);
  const totalFaltas = faltasPorPessoa.reduce((s, p) => s + p.qtd, 0);
  const totalHorasFalta = faltasPorPessoa.reduce((s, p) => s + p.horas, 0);

  // ABONO ranking
  const abonoCount = new Map();
  data.forEach((d) => {
    if (d.abono) {
      abonoCount.set(d.abono, (abonoCount.get(d.abono) || 0) + 1);
    }
  });
  const abonoRank = Array.from(abonoCount.entries())
    .map(([motivo, qtd]) => ({ motivo, qtd }))
    .sort((a, b) => b.qtd - a.qtd);
  const totalAbonos = abonoRank.reduce((s, a) => s + a.qtd, 0);

  // EXTRA por funcionário
  const extraPorPessoa = Array.from(chapas.values())
    .map((regs) => {
      const firstReg = regs.at(0);
      return { chapa: firstReg.chapa, nome: firstReg.nome, funcao: firstReg.funcao, obra: firstReg.obra,
        min: regs.reduce((s, r) => s + r.extraMin, 0) };
    })
    .filter((p) => p.min > 0)
    .sort((a, b) => b.min - a.min);
  const totalExtraMin = extraPorPessoa.reduce((s, p) => s + p.min, 0);

  const obras = [...new Set(data.map((d) => d.obra).filter(Boolean))];

  // ABSENTEÍSMO por semana = (faltas + atestados médicos) / efetivo ativo
  const inicioSemana = (ts) => {
    const d = new Date(ts);
    const dia = (d.getDay() + 6) % 7; // 0 = segunda
    d.setDate(d.getDate() - dia); d.setHours(0, 0, 0, 0);
    return d;
  };
  const semanas = new Map();
  data.forEach((d) => {
    if (!d.ts) return;
    const ini = inicioSemana(d.ts);
    const key = ini.getTime();
    if (!semanas.has(key)) {
      semanas.set(key, {
        key, ini,
        label: `${String(ini.getDate()).padStart(2, "0")}/${String(ini.getMonth() + 1).padStart(2, "0")}`,
        faltas: 0, atestados: 0, diasSet: new Set(),
      });
    }
    const semObj = semanas.get(key);
    if (d.faltaMin > 0) semObj.faltas++;
    if (d.abono === "Atestado Médico") semObj.atestados++;
    const wd = new Date(d.ts).getDay(); // 0=dom ... 6=sab
    if (wd >= 1 && wd <= 5) semObj.diasSet.add(new Date(d.ts).toDateString());
  });
  const efetivoAtivo = efetivo.A || 1;
  const ddmm = (dt) => `${String(dt.getDate()).padStart(2, "0")}/${String(dt.getMonth() + 1).padStart(2, "0")}`;
  const absSemanas = Array.from(semanas.values()).sort((x, y) => x.key - y.key).map((s, i) => {
    const fim = new Date(s.ini); fim.setDate(fim.getDate() + 4); // segunda → sexta
    const diasUteis = s.diasSet.size || 1;
    const total = s.faltas + s.atestados;
    return {
      ...s, num: i + 1, semLabel: `Sem. ${i + 1}`,
      range: `${ddmm(s.ini)} a ${ddmm(fim)}`,
      diasUteis, total,
      indice: (total / (efetivoAtivo * diasUteis)) * 100,
    };
  });
  const absTotalAusencias = absSemanas.reduce((s, w) => s + w.total, 0);
  const absTotalDiasUteis = absSemanas.reduce((s, w) => s + w.diasUteis, 0) || 1;
  const absIndiceGeral = (absTotalAusencias / (efetivoAtivo * absTotalDiasUteis)) * 100;
  const absTotalFaltas = absSemanas.reduce((s, w) => s + w.faltas, 0);
  const absTotalAtestados = absSemanas.reduce((s, w) => s + w.atestados, 0);

  return { total, funcionarios, efetivo, efetivoListas, faltasPorPessoa, totalFaltas, totalHorasFalta,
    abonoRank, totalAbonos, extraPorPessoa, totalExtraMin, obras,
    absSemanas, absIndiceGeral, absTotalAusencias, absTotalFaltas, absTotalAtestados, efetivoAtivo };
}

// detalha as pessoas de um motivo de abono específico
function detalheAbono(data, motivo) {
  const porPessoa = new Map();
  data.forEach((d) => {
    if (d.abono !== motivo) return;
    if (!porPessoa.has(d.chapa)) {
      porPessoa.set(d.chapa, { nome: d.nome, funcao: d.funcao, obra: d.obra, datas: [] });
    }
    if (d.data) {
      porPessoa.get(d.chapa).datas.push(d.data);
    }
  });
  return Array.from(porPessoa.values())
    .map((p) => ({ ...p, dias: p.datas.length, datas: p.datas.sort().join(", ") }))
    .sort((a, b) => b.dias - a.dias);
}

/* ====== UI ====== */
const Initials = ({ name, photo, bg = NAVY, size = 36 }) => {
  if (photo) {
    return <img src={photo} alt={name} className="rounded-full object-cover shrink-0" style={{ width: size, height: size }} />;
  }
  const ini = norm(name).split(/\s+/).slice(0, 2).map((w) => w[0]).join("").toUpperCase();
  return <div className="flex items-center justify-center rounded-full text-white font-medium shrink-0"
    style={{ width: size, height: size, background: bg, fontSize: size * 0.33 }}>{ini}</div>;
};
const Card = ({ children, className = "", style = {} }) => (
  <div className={`bg-white rounded-2xl anim-card ${className}`}
    style={{ boxShadow: "0 1px 3px rgba(4,58,102,0.06)", ...style }}>{children}</div>
);

// anel de "saúde geral" do efetivo (% de ativos) com legenda
function SaudeGeralGauge({ a, sitPie }) {
  const totalEfetivo = a.efetivo.A + a.efetivo.F + a.efetivo.P + a.efetivo.D;
  const pct = totalEfetivo ? Math.round((a.efetivo.A / totalEfetivo) * 100) : 0;
  const R = 52, C = 2 * Math.PI * R;
  const fill = (pct / 100) * C;
  return (
    <div className="flex items-center gap-6 w-full">
      <div style={{ width: 150, height: 150, position: "relative", flexShrink: 0 }}>
        <svg width="150" height="150" viewBox="0 0 140 140">
          <circle cx="70" cy="70" r={R} fill="none" stroke="#EDF1F5" strokeWidth="14" />
          <circle cx="70" cy="70" r={R} fill="none" stroke={OK} strokeWidth="14" strokeLinecap="round"
            strokeDasharray={`${fill} ${C}`} transform="rotate(-90 70 70)" />
        </svg>
        <div style={{ position: "absolute", inset: 0, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center" }}>
          <span className="font-bold" style={{ color: NAVY_DARK, fontSize: 32, lineHeight: 1 }}>{pct}%</span>
          <span style={{ color: "#94A3B8", fontSize: 10, letterSpacing: 1, marginTop: 4 }}>{t("saudeGeral")}</span>
        </div>
      </div>
      <div className="flex-1 flex flex-col gap-3">
        {sitPie.map((d) => (
          <div key={d.name} className="flex items-center justify-between">
            <span className="flex items-center gap-2 text-sm text-gray-600">
              <span style={{ width: 10, height: 10, borderRadius: 3, background: d.color }} />{d.name}</span>
            <span className="font-semibold text-sm" style={{ color: NAVY_DARK }}>{d.value}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

function ApiaLogo({ light }) {
  const c = light ? "#fff" : NAVY;
  return (
    <div className="flex items-center gap-2">
      <svg width="34" height="34" viewBox="0 0 100 100">
        <circle cx="50" cy="50" r="48" fill={c} />
        <path d="M50 22 L78 74 L22 74 Z" fill="none" stroke={light ? NAVY : "#fff"} strokeWidth="6" />
        <path d="M50 42 L64 68 L36 68 Z" fill={light ? NAVY : "#fff"} />
      </svg>
      <span className="font-bold tracking-tight" style={{ color: c, fontSize: 22 }}>ÁPIA</span>
    </div>
  );
}

/* ====== UPLOAD ====== */
function UploadScreen({ onData, onDemo }) {
  const inputRef = useRef();
  const [drag, setDrag] = useState(false);
  const [err, setErr] = useState("");
  const handleFile = (file) => {
    setErr("");
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const sheetName = wb.SheetNames.at(0);
        const sheet = Reflect.get(wb.Sheets, sheetName);
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
        const data = parseSheet(rows);
        if (!data.length) throw new Error();
        onData(data, file.name);
      } catch { setErr("Não consegui ler esta planilha. Confira o formato do relatório de ponto."); }
    };
    reader.readAsArrayBuffer(file);
  };
  return (
    <div className="min-h-screen flex items-center justify-center p-6" style={{ background: NAVY }}>
      <div className="w-full max-w-lg text-center">
        <div className="flex justify-center mb-8"><ApiaLogo light /></div>
        <div onDragOver={(e) => { e.preventDefault(); setDrag(true); }} onDragLeave={() => setDrag(false)}
          onDrop={(e) => { e.preventDefault(); setDrag(false); if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); }}
          onClick={() => inputRef.current.click()}
          className="bg-white rounded-3xl p-12 cursor-pointer"
          style={{ border: drag ? `3px dashed ${NAVY_MED}` : "3px dashed transparent" }}>
          <div className="flex justify-center mb-5">
            <div className="rounded-2xl p-5" style={{ background: NAVY_SOFT }}><Upload size={40} color={NAVY} /></div>
          </div>
          <h2 className="text-xl font-semibold mb-2" style={{ color: NAVY_DARK }}>{t("anexePlanilha")}</h2>
          <p className="text-gray-500 text-sm mb-6">{t("arrasteArquivo")}<b>.xlsx</b> ou clique para selecionar.<br />O relatório é gerado automaticamente.</p>
          <span className="inline-block text-white font-medium px-6 py-3 rounded-xl text-sm" style={{ background: NAVY }}>{t("selecionarPlanilha")}</span>
          <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden"
            onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])} />
        </div>
        {err && <p className="text-white text-sm mt-4 bg-red-500/30 rounded-lg py-2 px-3">{err}</p>}
        <button onClick={onDemo} className="text-white/70 text-sm mt-6 underline hover:text-white">{t("verDemonstracao")}</button>
      </div>
    </div>
  );
}

/* ====== DEMO ====== */
function demoData() {
  const funcs = ["OFICIAL DE SERVICOS GERAIS", "PEDREIRO", "ARMADOR", "SERVENTE", "CARPINTEIRO"];
  const motivos = ["Aguardando Crachá", "Atestado Médico", "Exames Periódicos", "Liberalidade Empresa", "Abono Chefia"];
  const out = [];
  for (let i = 0; i < 291; i++) {
    const obra = i % 2 ? "OBRA - 0948" : "OBRA - 0935";
    const funcao = funcs.at(i % funcs.length);
    let sit = "A";
    if (i < 6) sit = "D"; else if (i < 16) sit = "P"; else if (i < 18) sit = "F";
    for (let d = 0; d < 9; d++) {
      const faltou = sit === "A" && Math.random() < 0.05;
      const dt = new Date(2026, 4, 21 + d);
      out.push({ chapa: String(100000 + i), nome: `Funcionário ${i + 1}`, obra, funcao, situacao: sit,
        faltaMin: faltou ? (Math.random() < 0.5 ? 480 : 540) : 0,
        extraMin: sit === "A" && Math.random() < 0.25 ? Math.floor(Math.random() * 90) : 0,
        abono: Math.random() < 0.06 ? motivos.at(Math.floor(Math.random() * motivos.length)) : null,
        data: `${String(dt.getDate()).padStart(2, "0")}/05`, ts: dt.getTime() });
    }
  }
  return out;
}

/* ====== DASHBOARD ====== */
function Dashboard({ data, fileName, conta, onData, onDemo, onReset, onLogout, onUpdateProfile, feriasData, feriasFileName, onFeriasData }) {
  const isAdmin = conta?.role === "Administrador";
  const [obraFilter, setObraFilter] = useState("Todas");
  const [nav, setNav] = useState("dashboard");
  const [abonoSel, setAbonoSel] = useState(null);
  const [efetivoSel, setEfetivoSel] = useState(null);
  const [extraSel, setExtraSel] = useState(null);
  const [showProfile, setShowProfile] = useState(false);
  const [newName, setNewName] = useState(conta?.nome || "");
  const [savingProfile, setSavingProfile] = useState(false);
  const [profErr, setProfErr] = useState("");
  const fileInputRef = useRef(null);
  const [tempRole, setTempRole] = useState(conta?.role || "Gerente");
  const [adminCode, setAdminCode] = useState("");
  const prevShowProfile = useRef(false);

  useEffect(() => {
    if (showProfile && !prevShowProfile.current && conta) {
      setNewName(conta.nome || "");
      setTempRole(conta.role || "Gerente");
      setAdminCode("");
      setProfErr("");
    }
    prevShowProfile.current = showProfile;
  }, [showProfile, conta]);

  const handlePhotoChange = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      setSavingProfile(true);
      setProfErr("");
      const compressedBase64 = await compressAndSavePhoto(file);
      await onUpdateProfile({ photo: compressedBase64 });
    } catch (err) {
      setProfErr("Erro ao carregar ou processar imagem.");
    } finally {
      setSavingProfile(false);
    }
  };

  const handleSaveProfile = async (e) => {
    e.preventDefault();
    if (!newName.trim()) { setProfErr("O nome não pode ficar em branco."); return; }
    
    // Validação de segurança: exige o código se estiver mudando para Administrador
    if (tempRole === "Administrador" && conta?.role !== "Administrador") {
      if (adminCode.trim().toLowerCase() !== "ponto") {
        setProfErr("Código de segurança incorreto. Não foi possível assumir o cargo de Administrador.");
        return;
      }
    }

    try {
      setSavingProfile(true);
      setProfErr("");
      await onUpdateProfile({ nome: newName.trim(), role: tempRole });
      setShowProfile(false);
    } catch (err) {
      setProfErr("Erro ao salvar o perfil.");
    } finally {
      setSavingProfile(false);
    }
  };

  const filtered = useMemo(() => obraFilter === "Todas" ? data : data.filter((d) => d.obra === obraFilter), [data, obraFilter]);
  const a = useMemo(() => aggregate(filtered), [filtered]);

  const efetivoCards = [
    { key: "A", label: "Ativos", desc: "Efetivo active", icon: UserCheck, color: OK, value: a.efetivo.A },
    { key: "F", label: "Férias", desc: "Pessoas de férias", icon: Plane, color: NAVY_MED, value: a.efetivo.F },
    { key: "P", label: "Afastados INSS", desc: "Afastamento INSS", icon: HeartPulse, color: WARN, value: a.efetivo.P },
    { key: "D", label: "Demitidos", desc: "Desligados", icon: UserX, color: DANGER, value: a.efetivo.D },
  ];
  const sitPie = efetivoCards.filter((c) => c.value > 0).map((c) => ({ name: c.label, value: c.value, color: c.color }));

  const exportar = () => {
    const txt = [
      `RELATÓRIO DE FREQUÊNCIA — ÁPIA`, `Arquivo: ${fileName} · Filtro: ${obraFilter}`, ``,
      `EFETIVO (CODSITUACAO):`,
      `  Ativos (A): ${a.efetivo.A}`, `  Férias (F): ${a.efetivo.F}`,
      `  Afastados INSS (P): ${a.efetivo.P}`, `  Demitidos (D): ${a.efetivo.D}`, ``,
      `FALTAS: ${a.totalFaltas} faltas · ${minToHHMM(a.totalHorasFalta)} perdidas · ${a.faltasPorPessoa.length} pessoas`,
      ...a.faltasPorPessoa.slice(0, 10).map((p, i) => `  ${i + 1}. ${p.nome}: ${p.qtd} falta(s) — ${minToHHMM(p.horas)}`), ``,
      `ABONOS (ranking): ${a.totalAbonos} total`,
      ...a.abonoRank.map((r, i) => `  ${i + 1}. ${r.motivo}: ${r.qtd}`), ``,
      `HORAS EXTRAS (limite 40h/mês): ${minToHHMM(a.totalExtraMin)} total · ${a.extraPorPessoa.length} pessoas`,
      ...a.extraPorPessoa.slice(0, 10).map((p, i) => `  ${i + 1}. ${p.nome}: ${minToHHMM(p.min)} (${Math.round(p.min / LIMITE_EXTRA * 100)}% do limite)`),
    ].join("\n");
    navigator.clipboard.writeText(txt);
    alert("Resumo copiado! Cole nos seus slides.");
  };

  const navItems = [
    { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { id: "efetivo", label: "Efetivo", icon: Users },
    { id: "ferias", label: "Férias", icon: Palmtree },
    { id: "faltas", label: "Faltas", icon: CalendarX },
    { id: "abono", label: "Abonos", icon: Award },
    { id: "extras", label: "Horas extras", icon: Clock },
    { id: "absenteismo", label: "Absenteísmo", icon: Activity },
    { id: "relatorios", label: "Relatórios", icon: FileBarChart },
    ...(isAdmin ? [{ id: "importar", label: "Importar", icon: Upload }] : []),
  ];
  const navTitulo = { dashboard: "Relatório de frequência", efetivo: "Efetivo por situação",
    ferias: "Controle de férias", faltas: "Controle de faltas", abono: "Abonos", extras: "Horas extras",
    absenteismo: "Absenteísmo", relatorios: "Relatórios", importar: "Importar planilha" };

  /* ---- bloco de horas extras (lista completa, sem espaço em branco) ---- */
  const ExtraStatus = (min) => {
    const pct = Math.min((min / LIMITE_EXTRA) * 100, 100);
    const status = min >= LIMITE_EXTRA ? { t: "Limite atingido", c: DANGER } :
      pct >= 75 ? { t: "Atenção", c: WARN } : { t: "Normal", c: OK };
    return { pct, status };
  };

  const ExtraRow = ({ p, i, rank }) => {
    const { pct, status } = ExtraStatus(p.min);
    const remainingMin = Math.max(0, LIMITE_EXTRA - p.min);
    const remainingStr = remainingMin > 0 ? `${t("podeFazerMais")}${minToHHMM(remainingMin)}` : t("limiteAtingido");

    return (
      <div onClick={() => setExtraSel(p)} className="flex items-center gap-3 py-2 cursor-pointer row-hov" style={{ borderBottom: "0.5px solid #EDF1F5" }}>
        {rank && <span className="font-bold text-gray-400 w-5 text-sm text-center shrink-0">{i + 1}</span>}
        <Initials name={p.nome} bg={NAVY} />
        <div className="flex-1 min-w-0">
          <div className="flex items-center justify-between">
            <div className="min-w-0 flex-1">
              <p className="font-medium text-sm truncate" style={{ color: NAVY_DARK }}>{p.nome}</p>
              <p className="text-xs text-gray-400 truncate mt-0.5">
                {p.funcao} · <span className="font-medium" style={{ color: remainingMin > 0 ? (remainingMin <= 600 ? WARN : OK) : DANGER }}>{remainingStr}</span>
              </p>
            </div>
            <span className="font-semibold text-sm ml-2 shrink-0" style={{ color: NAVY_DARK }}>{minToHHMM(p.min)}</span>
          </div>
          <div className="flex items-center gap-2 mt-1.5">
            <div className="rounded-full h-1.5 flex-1 overflow-hidden" style={{ background: "#EDF1F5" }}>
              <div className="h-full rounded-full bar-fill" style={{ width: `${pct}%`, background: status.c }} />
            </div>
            <span className="text-xs font-medium shrink-0" style={{ color: status.c, minWidth: 90, textAlign: "right" }}>{status.t}</span>
          </div>
        </div>
      </div>
    );
  };

  /* ---- conteúdo de cada aba ---- */
  const renderConteudo = () => {
    if (nav === "efetivo") return <AbaEfetivo a={a} efetivoCards={efetivoCards} sitPie={sitPie} onSel={setEfetivoSel} />;
    if (nav === "ferias") return <AbaFerias feriasData={feriasData || []} feriasFileName={feriasFileName || ""} onNav={() => setNav("importar")} />;
    if (nav === "faltas") return <AbaFaltas a={a} />;
    if (nav === "abono") return <AbaAbono a={a} onSel={setAbonoSel} />;
    if (nav === "extras") return <AbaExtras a={a} ExtraRow={ExtraRow} />;
    if (nav === "absenteismo") return <AbaAbsenteismo a={a} />;
    if (nav === "relatorios") return <AbaRelatorios a={a} exportar={exportar} />;
    if (nav === "importar") return isAdmin ? <AbaImportar
      onImport={(d, n) => { onData(d, n); setNav("dashboard"); }}
      onDemo={() => { onDemo(); setNav("dashboard"); }} temDados={data.length > 0} fileName={fileName}
      onFeriasImport={(d, n) => { onFeriasData(d, n); setNav("ferias"); }} temFeriasData={(feriasData || []).length > 0} feriasFileName={feriasFileName || ""} />
      : <Card className="p-8 text-center"><p className="text-gray-400 text-sm">{t("acessoRestrito")}</p></Card>;

    /* ===== DASHBOARD (visão geral) ===== */
    return (
      <>
        {data.length === 0 && (
          <Card className="p-5 mb-4 flex items-center justify-between flex-wrap gap-3" style={{ borderLeft: `4px solid ${NAVY}`, borderRadius: 12 }}>
            <div className="flex items-center gap-3">
              <div className="rounded-xl p-2.5" style={{ background: NAVY_SOFT }}><Upload size={20} color={NAVY} /></div>
              <div>
                <p className="font-medium text-sm" style={{ color: NAVY_DARK }}>{t("nenhumaPlanilha")}</p>
                <p className="text-xs text-gray-400">{isAdmin
                  ? "Os indicadores estão zerados. Importe a planilha de ponto para preencher o relatório."
                  : "Os indicadores estão zerados. Aguarde o Administrador importar a planilha de ponto."}</p>
              </div>
            </div>
            {isAdmin && (
              <button onClick={() => setNav("importar")} className="py-2.5 px-5 rounded-xl btn-primary text-white font-medium text-sm flex items-center gap-2" style={{ background: NAVY }}>
                <Upload size={16} /> {t("selecionarPlanilha")}</button>
            )}
          </Card>
        )}        <div className="grid grid-cols-4 gap-4 mb-4 stagger">
          {efetivoCards.map((c) => (
            <Card key={c.key} className="p-5 cursor-pointer transition-transform hover:-translate-y-0.5"
              style={{ borderTop: `3px solid ${c.color}` }}>
              <div onClick={() => setEfetivoSel(c.key)}>
                <div className="flex items-start justify-between mb-2">
                  <span className="text-gray-500 font-medium" style={{ fontSize: 13 }}>{c.label}</span>
                </div>
                <span className="font-bold leading-none" style={{ color: NAVY_DARK, fontSize: 34 }}>{c.value}</span>
                <p className="text-gray-400 mt-2 flex items-center gap-1" style={{ fontSize: 12 }}>{c.desc} · <b style={{ color: c.color }}>{c.key}</b>
                  <ChevronRight size={12} color="#94A3B8" className="ml-auto" /></p>
              </div>
            </Card>
          ))}
        </div>

        <div className="grid gap-4 mb-4" style={{ gridTemplateColumns: "1fr 1.4fr" }}>
          <Card className="p-6 flex flex-col">
            <h3 className="font-semibold mb-4" style={{ color: NAVY_DARK, fontSize: 16 }}>{t("distribuicaoEfetivo")}</h3>
            <div className="flex-1 flex items-center">
              <SaudeGeralGauge a={a} sitPie={sitPie} />
            </div>
          </Card>

          <Card className="p-6">
            <div className="flex items-center justify-between mb-4">
              <h3 className="font-semibold" style={{ color: NAVY_DARK, fontSize: 16 }}>{t("faltas")}</h3>
              <div className="flex gap-4 text-right">
                <div><p className="font-bold" style={{ color: DANGER, fontSize: 20 }}>{a.totalFaltas}</p><p className="text-xs text-gray-400">{t("faltasMin")}</p></div>
                <div><p className="font-bold" style={{ color: NAVY_DARK, fontSize: 20 }}>{minToHHMM(a.totalHorasFalta)}</p><p className="text-xs text-gray-400">{t("horasPerdidas")}</p></div>
                <div><p className="font-bold" style={{ color: NAVY_DARK, fontSize: 20 }}>{a.faltasPorPessoa.length}</p><p className="text-xs text-gray-400">{t("pessoas")}</p></div>
              </div>
            </div>
            <div className="flex flex-col gap-2">
              {a.faltasPorPessoa.slice(0, 5).map((p, i) => (
                <div key={i} className="flex items-center gap-3">
                  <span className="font-bold text-gray-400 w-5 text-sm">{i + 1}</span>
                  <Initials name={p.nome} bg={DANGER} />
                  <div className="flex-1 min-w-0"><p className="font-medium text-sm truncate" style={{ color: NAVY_DARK }}>{p.nome}</p>
                    <p className="text-xs text-gray-400 truncate">{p.funcao}</p></div>
                  <span className="font-semibold text-sm px-2.5 py-1 rounded-lg" style={{ background: "#FDECEC", color: DANGER }}>
                    {p.qtd} falta{p.qtd > 1 ? "s" : ""} · {minToHHMM(p.horas)}</span>
                </div>
              ))}
              {a.faltasPorPessoa.length === 0 && <p className="text-sm text-gray-400">{t("nenhumaFaltaPeriodo")}</p>}
            </div>
            {a.faltasPorPessoa.length > 5 && (
              <button onClick={() => setNav("faltas")} className="text-sm font-medium mt-3 flex items-center gap-1" style={{ color: NAVY }}>
                {t("verTodasAs")}{a.faltasPorPessoa.length} {t("pessoas")} <ChevronRight size={14} /></button>)}
          </Card>
        </div>

        <div className="grid gap-4" style={{ gridTemplateColumns: "1fr 1.4fr" }}>
          <Card className="p-6 flex flex-col">
            <div className="flex items-center justify-between mb-4">
              <h3 className="font-semibold" style={{ color: NAVY_DARK, fontSize: 16 }}>{t("rankingAbonos")}</h3>
              <span className="font-bold" style={{ color: NAVY, fontSize: 20 }}>{a.totalAbonos}</span>
            </div>
            <div className="flex flex-col flex-1 justify-between" style={{ gap: 10 }}>
              {a.abonoRank.slice(0, 6).map((r, i) => {
                const max = a.abonoRank[0].qtd;
                return (
                  <div key={i} onClick={() => setAbonoSel(r.motivo)}
                    className="cursor-pointer rounded-lg px-2 py-1.5 -mx-2 hover:bg-gray-50">
                    <div className="flex items-center gap-2 mb-1.5">
                      <span className="flex items-center justify-center rounded-lg text-white text-xs font-bold w-5 h-5"
                        style={{ background: i === 0 ? "#D4AF37" : i === 1 ? "#9CA3AF" : i === 2 ? "#B87333" : NAVY_MED }}>{i + 1}</span>
                      <span className="text-sm flex-1" style={{ color: NAVY_DARK }}>{r.motivo}</span>
                      <span className="font-semibold text-sm" style={{ color: NAVY_DARK }}>{r.qtd}</span>
                      <ChevronRight size={14} color="#94A3B8" />
                    </div>
                    <div className="rounded-full h-2 ml-7 overflow-hidden" style={{ background: "#EDF1F5" }}>
                      <div className="h-full rounded-full bar-fill" style={{ width: `${(r.qtd / max) * 100}%`, background: NAVY }} />
                    </div>
                  </div>
                );
              })}
              {a.abonoRank.length === 0 && <p className="text-sm text-gray-400">{t("nenhumAbonoPeriodo")}</p>}
            </div>
            {a.abonoRank.length > 6 && (
              <button onClick={() => setNav("abono")} className="text-sm font-medium mt-4 flex items-center gap-1" style={{ color: NAVY }}>
                {t("verRankingCompleto")}<ChevronRight size={14} /></button>)}
          </Card>

          <Card className="p-6 flex flex-col">
            <div className="flex items-center justify-between mb-1">
              <h3 className="font-semibold" style={{ color: NAVY_DARK, fontSize: 16 }}>{t("horasExtrasFunc")}</h3>
              <span className="text-xs px-2 py-1 rounded-lg" style={{ background: NAVY_SOFT, color: NAVY }}>{t("limiteMensal")}</span>
            </div>
            <p className="text-gray-400 text-xs mb-3">{t("totalLabel")}{minToHHMM(a.totalExtraMin)} · {a.extraPorPessoa.length} {t("funcionarios")}</p>
            <div className="flex flex-col flex-1 justify-between">
              {a.extraPorPessoa.slice(0, 6).map((p, i) => <ExtraRow key={i} p={p} i={i} />)}
              {a.extraPorPessoa.length === 0 && <p className="text-sm text-gray-400">{t("nenhumaExtraPeriodo")}</p>}
            </div>
            {a.extraPorPessoa.length > 6 && (
              <button onClick={() => setNav("extras")} className="text-sm font-medium mt-4 flex items-center gap-1" style={{ color: NAVY }}>
                {t("verTodosOs")}{a.extraPorPessoa.length} {t("funcionarios")} <ChevronRight size={14} /></button>)}
          </Card>
        </div>
      </>
    );
  };

  return (
    <div className="min-h-screen flex" style={{ background: "#F4F6F9", fontFamily: "system-ui, sans-serif" }}>
      <aside className="w-60 bg-white flex flex-col p-5 shrink-0" style={{ borderRight: "1px solid #EDF1F5" }}>
        <div className="mb-6"><ApiaLogo /></div>
        <div className="relative mb-6">
          <Search size={16} className="absolute left-3 top-2.5 text-gray-400" />
          <input placeholder={t("buscar")} className="w-full pl-9 pr-3 py-2 rounded-xl text-sm outline-none" style={{ background: "#F4F6F9" }} />
        </div>
        <nav className="flex flex-col gap-1 flex-1">
          {navItems.map((it) => {
            const active = nav === it.id;
            return <button key={it.id} onClick={() => setNav(it.id)}
              className="nav-item flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-medium text-left"
              style={{ background: active ? NAVY_SOFT : "transparent", color: active ? NAVY : "#5B6B7B" }}>
              <it.icon size={18} /> {it.label}</button>;
          })}
        </nav>
        <div className="pt-4 mt-4" style={{ borderTop: "1px solid #EDF1F5" }}>
          <div onClick={() => setShowProfile(true)} className="flex items-center gap-3 mb-3 cursor-pointer p-2 -mx-2 rounded-xl hover:bg-gray-100 transition-colors">
            <Initials name={conta?.nome || "ÁPIA"} photo={conta?.photo} />
            <div className="min-w-0 flex-1">
              <p className="font-semibold text-sm truncate" style={{ color: NAVY_DARK }}>{conta?.nome || "Gestão de Ponto"}</p>
              <span className="text-xs px-2 py-0.5 rounded-full truncate inline-block" style={{ background: isAdmin ? NAVY_SOFT : "#F0E9D8", color: isAdmin ? NAVY : "#8A6D1A" }}>{conta?.role || "—"}</span>
            </div>
          </div>
          {isAdmin && (
            <button onClick={onReset} className="flex items-center gap-3 px-1 py-2 text-sm" style={{ color: "#5B6B7B" }}>
              <Upload size={16} /> {t("novaplanilha")}</button>
          )}
          <button onClick={onLogout} className="flex items-center gap-3 px-1 py-2 text-sm" style={{ color: DANGER }}>
            <LogOut size={16} /> {t("sair")}</button>
        </div>
      </aside>

      <main className="flex-1 p-6 overflow-auto" style={{ position: "relative" }}>
        <div className="flex items-center justify-between mb-6 flex-wrap gap-3">
          <div>
            <h1 className="font-bold" style={{ color: NAVY_DARK, fontSize: 22 }}>{Reflect.get(navTitulo, nav)}</h1>
            <p className="text-gray-400 text-sm">{fileName} · {a.funcionarios} {t("colaboradores")} · {a.total} {t("registros")}</p>
          </div>
          <div className="flex items-center gap-2 flex-wrap">
            {[t("todas"), ...a.obras].map((o) => (
              <button key={o} onClick={() => setObraFilter(o === t("todas") ? "Todas" : o)} className="px-4 py-2 rounded-xl text-sm font-medium"
                style={{ background: (o === t("todas") && obraFilter === "Todas") || obraFilter === o ? NAVY : "#fff", color: (o === t("todas") && obraFilter === "Todas") || obraFilter === o ? "#fff" : "#5B6B7B",
                  boxShadow: (o === t("todas") && obraFilter === "Todas") || obraFilter === o ? "none" : "0 1px 2px rgba(4,58,102,0.06)" }}>{o}</button>
            ))}
          </div>
        </div>

        <div key={nav} className="anim-fade">{renderConteudo()}</div>

        {abonoSel && (() => {
          const lista = detalheAbono(filtered, abonoSel);
          const totalDias = lista.reduce((s, p) => s + p.dias, 0);
          return (
            <div onClick={() => setAbonoSel(null)} className="flex items-center justify-center p-4 anim-fade"
              style={{ position: "fixed", inset: 0, background: "rgba(3,42,74,0.45)", zIndex: 100 }}>
              <div onClick={(e) => e.stopPropagation()} className="bg-white rounded-2xl w-full anim-scale"
                style={{ maxWidth: 640, maxHeight: "85vh", display: "flex", flexDirection: "column" }}>
                <div className="p-5 flex items-center justify-between" style={{ borderBottom: "1px solid #EDF1F5" }}>
                  <div>
                    <h3 className="font-semibold" style={{ color: NAVY_DARK, fontSize: 17 }}>{abonoSel}</h3>
                    <p className="text-gray-400 text-sm">{lista.length} colaborador(es) · {totalDias} dia(s)</p>
                  </div>
                  <button onClick={() => setAbonoSel(null)} className="rounded-lg p-2 hover:bg-gray-100" style={{ color: "#5B6B7B" }}><X size={20} /></button>
                </div>
                <div className="p-5 overflow-auto flex flex-col gap-2">
                  {lista.map((p, i) => (
                    <div key={i} className="flex items-center gap-3 p-2 rounded-xl" style={{ background: p.dias >= 5 ? "#FAEEDA" : "#F8FAFC" }}>
                      <span className="font-bold text-gray-400 w-5 text-sm text-center">{i + 1}</span>
                      <Initials name={p.nome} bg={NAVY} />
                      <div className="flex-1 min-w-0">
                        <p className="font-medium text-sm truncate" style={{ color: NAVY_DARK }}>{p.nome}</p>
                        <p className="text-xs text-gray-400 truncate">{p.funcao} · {p.obra}</p>
                        <p className="text-xs text-gray-400 truncate">{p.datas}</p>
                      </div>
                      <span className="font-semibold text-sm px-2.5 py-1 rounded-lg shrink-0"
                        style={{ background: p.dias >= 5 ? "#F5E0BC" : NAVY_SOFT, color: p.dias >= 5 ? "#854F0B" : NAVY }}>
                        {p.dias} dia{p.dias > 1 ? "s" : ""}</span>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          );
        })()}

        {efetivoSel && (() => {
          const metaMap = { A: { t: "Ativos", c: OK }, F: { t: "Férias", c: NAVY_MED },
            P: { t: "Afastamento INSS", c: WARN }, D: { t: "Demitidos", c: DANGER } };
          const meta = Reflect.get(metaMap, efetivoSel);
          const lista = [...(Reflect.get(a.efetivoListas, efetivoSel) || [])].sort((x, y) => x.nome.localeCompare(y.nome));
          return (
            <div onClick={() => setEfetivoSel(null)} className="flex items-center justify-center p-4 anim-fade"
              style={{ position: "fixed", inset: 0, background: "rgba(3,42,74,0.45)", zIndex: 100 }}>
              <div onClick={(e) => e.stopPropagation()} className="bg-white rounded-2xl w-full anim-scale"
                style={{ maxWidth: 640, maxHeight: "85vh", display: "flex", flexDirection: "column" }}>
                <div className="p-5 flex items-center justify-between" style={{ borderBottom: "1px solid #EDF1F5" }}>
                  <div className="flex items-center gap-3">
                    <span className="w-3 h-8 rounded-full" style={{ background: meta.c }} />
                    <div>
                      <h3 className="font-semibold" style={{ color: NAVY_DARK, fontSize: 17 }}>{meta.t} <span className="text-gray-400 font-normal">({efetivoSel})</span></h3>
                      <p className="text-gray-400 text-sm">{lista.length} colaborador(es)</p>
                    </div>
                  </div>
                  <button onClick={() => setEfetivoSel(null)} className="rounded-lg p-2 hover:bg-gray-100" style={{ color: "#5B6B7B" }}><X size={20} /></button>
                </div>
                <div className="p-5 overflow-auto flex flex-col gap-2">
                  {lista.map((p, i) => (
                    <div key={i} className="flex items-center gap-3 p-2 rounded-xl" style={{ background: "#F8FAFC" }}>
                      <span className="font-bold text-gray-400 w-6 text-sm text-center shrink-0">{i + 1}</span>
                      <Initials name={p.nome} bg={meta.c} />
                      <div className="flex-1 min-w-0">
                        <p className="font-medium text-sm truncate" style={{ color: NAVY_DARK }}>{p.nome}</p>
                        <p className="text-xs text-gray-400 truncate">{p.funcao} · {p.obra}</p>
                      </div>
                    </div>
                  ))}
                  {lista.length === 0 && <p className="text-sm text-gray-400">{t("ninguemSituacaoPeriodo")}</p>}
                </div>
              </div>
            </div>
          );
        })()}

        {extraSel && (() => {
          const lista = filtered.filter((d) => d.chapa === extraSel.chapa && d.extraMin > 0).sort((x, y) => x.ts - y.ts);
          const totalMin = lista.reduce((s, d) => s + d.extraMin, 0);
          return (
            <div onClick={() => setExtraSel(null)} className="flex items-center justify-center p-4 anim-fade"
              style={{ position: "fixed", inset: 0, background: "rgba(3,42,74,0.45)", zIndex: 100 }}>
              <div onClick={(e) => e.stopPropagation()} className="bg-white rounded-2xl w-full anim-scale"
                style={{ maxWidth: 640, maxHeight: "85vh", display: "flex", flexDirection: "column" }}>
                <div className="p-5 flex items-center justify-between" style={{ borderBottom: "1px solid #EDF1F5" }}>
                  <div>
                    <h3 className="font-semibold" style={{ color: NAVY_DARK, fontSize: 17 }}>{extraSel.nome}</h3>
                    <p className="text-gray-400 text-sm">{extraSel.funcao} · {extraSel.obra}</p>
                    <p className="text-xs font-semibold mt-1" style={{ color: NAVY }}>{t("totalExtra")}{minToHHMM(totalMin)}</p>
                  </div>
                  <button onClick={() => setExtraSel(null)} className="rounded-lg p-2 hover:bg-gray-100" style={{ color: "#5B6B7B" }}><X size={20} /></button>
                </div>
                <div className="p-5 overflow-auto flex flex-col gap-2">
                  {lista.map((d, i) => (
                    <div key={i} className="flex items-center justify-between p-3 rounded-xl" style={{ background: "#F8FAFC" }}>
                      <div className="flex items-center gap-3">
                        <span className="font-bold text-gray-400 text-sm w-5 text-center">{i + 1}</span>
                        <div>
                          <p className="font-semibold text-sm" style={{ color: NAVY_DARK }}>{t("dia")}{d.data}</p>
                        </div>
                      </div>
                      <span className="font-semibold text-sm px-2.5 py-1 rounded-lg shrink-0"
                        style={{ background: NAVY_SOFT, color: NAVY }}>
                        +{minToHHMM(d.extraMin)}
                      </span>
                    </div>
                  ))}
                  {lista.length === 0 && <p className="text-sm text-gray-400 py-4 text-center">{t("nenhumRegistroExtra")}</p>}
                </div>
              </div>
            </div>
          );
        })()}

        {showProfile && (
          <div onClick={() => setShowProfile(false)} className="flex items-center justify-center p-4 anim-fade"
            style={{ position: "fixed", inset: 0, background: "rgba(3,42,74,0.45)", zIndex: 110 }}>
            <div onClick={(e) => e.stopPropagation()} className="bg-white rounded-2xl w-full anim-scale p-6"
              style={{ maxWidth: 420 }}>
              <div className="flex items-center justify-between mb-5" style={{ borderBottom: "1px solid #EDF1F5", paddingBottom: 12 }}>
                <h3 className="font-semibold text-lg" style={{ color: NAVY_DARK }}>{t("meuPerfil")}</h3>
                <button onClick={() => setShowProfile(false)} className="rounded-lg p-1.5 hover:bg-gray-100" style={{ color: "#5B6B7B" }}><X size={20} /></button>
              </div>

              <div className="flex flex-col items-center gap-4 mb-6">
                <div className="relative group cursor-pointer" onClick={() => fileInputRef.current.click()}>
                  <Initials name={conta?.nome || "ÁPIA"} photo={conta?.photo} size={84} />
                  <div className="absolute inset-0 bg-black/40 rounded-full flex items-center justify-center opacity-0 group-hover:opacity-100 transition-opacity duration-200">
                    <Upload size={20} color="#fff" />
                  </div>
                </div>
                <input type="file" ref={fileInputRef} accept="image/*" onChange={handlePhotoChange} className="hidden" />
                <div className="text-center">
                  <p className="text-xs text-gray-400">{t("cliqueFotoAlterar")}</p>
                </div>
              </div>

              <form onSubmit={handleSaveProfile} className="flex flex-col gap-4">
                <div>
                  <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">{t("nomeCompleto")}</label>
                  <input value={newName} onChange={(e) => setNewName(e.target.value)} disabled={savingProfile} className="w-full px-3 py-2.5 rounded-xl text-sm outline-none border" style={{ borderColor: "#D8E0E8", color: NAVY_DARK }} />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">{t("usuario")}</label>
                  <input value={conta?.user || ""} disabled className="w-full px-3 py-2.5 rounded-xl text-sm outline-none border bg-gray-50" style={{ borderColor: "#D8E0E8", color: "#94A3B8" }} />
                </div>

                <div>
                  <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">{t("perfilAcesso")}</label>
                  <div className="flex gap-2">
                    <button
                      type="button"
                      onClick={() => setTempRole("Administrador")}
                      disabled={savingProfile}
                      className="flex-1 py-2 px-3 rounded-xl border text-sm font-semibold transition-all duration-200"
                      style={{
                        background: tempRole === "Administrador" ? NAVY : "transparent",
                        color: tempRole === "Administrador" ? "#fff" : NAVY_DARK,
                        borderColor: tempRole === "Administrador" ? NAVY : "#D8E0E8"
                      }}
                    >
                      {t("administrador")}
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setTempRole("Gerente");
                        setAdminCode("");
                        setProfErr("");
                      }}
                      disabled={savingProfile}
                      className="flex-1 py-2 px-3 rounded-xl border text-sm font-semibold transition-all duration-200"
                      style={{
                        background: tempRole === "Gerente" ? NAVY : "transparent",
                        color: tempRole === "Gerente" ? "#fff" : NAVY_DARK,
                        borderColor: tempRole === "Gerente" ? NAVY : "#D8E0E8"
                      }}
                    >
                      {t("gerente")}
                    </button>
                  </div>
                </div>

                {tempRole === "Administrador" && conta?.role !== "Administrador" && (
                  <div className="anim-up">
                    <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">{t("codigoSegurancaAdmin")}</label>
                    <input
                      type="password"
                      value={adminCode}
                      onChange={(e) => setAdminCode(e.target.value)}
                      disabled={savingProfile}
                      placeholder={t("digiteCodigoPonto")}
                      className="w-full px-3 py-2.5 rounded-xl text-sm outline-none border"
                      style={{ borderColor: "#D8E0E8", color: NAVY_DARK }}
                    />
                  </div>
                )}

                {profErr && <p className="text-red-500 text-xs mt-1 bg-red-50 p-2 rounded-lg">{profErr}</p>}

                <div className="flex gap-3 mt-4">
                  <button type="button" onClick={() => setShowProfile(false)} disabled={savingProfile} className="flex-1 py-2.5 rounded-xl text-sm font-medium text-gray-600 bg-gray-100 hover:bg-gray-200">{t("cancelar")}</button>
                  <button type="submit" disabled={savingProfile} className="flex-1 py-2.5 rounded-xl text-sm font-medium text-white btn-primary" style={{ background: NAVY }}>{savingProfile ? "Salvando..." : t("salvarAlteracoes")}</button>
                </div>
              </form>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

/* ====== ABAS ====== */
function AbaEfetivo({ a, efetivoCards, sitPie, onSel }) {
  const blocos = [
    { key: "F", titulo: "Férias", color: NAVY_MED, lista: a.efetivoListas.F },
    { key: "P", titulo: "Afastamento INSS", color: WARN, lista: a.efetivoListas.P },
    { key: "D", titulo: "Demitidos", color: DANGER, lista: a.efetivoListas.D },
  ];
  return (
    <>
      <div className="grid grid-cols-4 gap-4 mb-4 stagger">
        {efetivoCards.map((c) => (
          <Card key={c.key} className="p-5 cursor-pointer transition-transform hover:-translate-y-0.5"
            style={{ borderTop: `3px solid ${c.color}` }}>
            <div onClick={() => onSel(c.key)}>
              <div className="flex items-start justify-between mb-2">
                <span className="text-gray-500 font-medium" style={{ fontSize: 13 }}>{c.label}</span>
              </div>
              <span className="font-bold leading-none" style={{ color: NAVY_DARK, fontSize: 34 }}>{c.value}</span>
              <p className="text-gray-400 mt-2 flex items-center gap-1" style={{ fontSize: 12 }}>{c.desc} · <b style={{ color: c.color }}>{c.key}</b>
                <ChevronRight size={12} color="#94A3B8" className="ml-auto" /></p>
            </div>
          </Card>
        ))}
      </div>
      <div className="grid gap-4" style={{ gridTemplateColumns: "1fr 1fr 1fr" }}>
        {blocos.map((b) => (
          <Card key={b.key} className="p-6">
            <div className="flex items-center justify-between mb-4">
              <h3 className="font-semibold" style={{ color: NAVY_DARK, fontSize: 15 }}>{b.titulo}</h3>
              <span className="font-bold" style={{ color: b.color, fontSize: 18 }}>{b.lista.length}</span>
            </div>
            <div className="flex flex-col gap-3">
              {b.lista.map((p, i) => (
                <div key={i} className="flex items-center gap-3">
                  <Initials name={p.nome} bg={b.color} />
                  <div className="min-w-0"><p className="font-medium text-sm truncate" style={{ color: NAVY_DARK }}>{p.nome}</p>
                    <p className="text-xs text-gray-400 truncate">{p.funcao}</p></div>
                </div>
              ))}
              {b.lista.length === 0 && <p className="text-sm text-gray-400">{t("ninguemSituacao")}</p>}
            </div>
          </Card>
        ))}
      </div>
    </>
  );
}

function AbaFaltas({ a }) {
  const [searchTerm, setSearchTerm] = useState("");

  const filteredFaltas = useMemo(() => {
    const term = searchTerm.trim().toLowerCase();
    if (!term) return a.faltasPorPessoa;
    return a.faltasPorPessoa.filter((p) =>
      p.nome.toLowerCase().includes(term) ||
      p.funcao.toLowerCase().includes(term)
    );
  }, [a.faltasPorPessoa, searchTerm]);

  return (
    <>
      <div className="grid grid-cols-3 gap-4 mb-4 stagger">
        <Card className="p-5"><p className="text-gray-500 text-sm mb-1">{t("totalFaltas")}</p><p className="font-bold" style={{ color: DANGER, fontSize: 32 }}>{a.totalFaltas}</p></Card>
        <Card className="p-5"><p className="text-gray-500 text-sm mb-1">{t("horasPerdidasCard")}</p><p className="font-bold" style={{ color: NAVY_DARK, fontSize: 32 }}>{minToHHMM(a.totalHorasFalta)}</p></Card>
        <Card className="p-5"><p className="text-gray-500 text-sm mb-1">{t("pessoasFalta")}</p><p className="font-bold" style={{ color: NAVY_DARK, fontSize: 32 }}>{a.faltasPorPessoa.length}</p></Card>
      </div>
      <Card className="p-6">
        <h3 className="font-semibold mb-4" style={{ color: NAVY_DARK, fontSize: 16 }}>{t("faltasColaborador")}</h3>

        {/* Barra de Pesquisa */}
        <div className="relative mb-5">
          <Search size={18} className="absolute left-3 top-3.5 text-gray-400" />
          <input
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            placeholder={t("buscarColaboradorCargo")}
            className="w-full pl-10 pr-3 py-2.5 rounded-xl text-sm outline-none border"
            style={{ borderColor: "#D8E0E8", color: NAVY_DARK }}
          />
        </div>

        <div className="flex flex-col">
          {filteredFaltas.map((p, i) => (
            <div key={i} className="flex items-center gap-3 py-2.5" style={{ borderBottom: "0.5px solid #EDF1F5" }}>
              <span className="font-bold text-gray-400 w-6 text-sm text-center shrink-0">{i + 1}</span>
              <Initials name={p.nome} bg={DANGER} />
              <div className="flex-1 min-w-0"><p className="font-semibold text-sm truncate" style={{ color: NAVY_DARK }}>{p.nome}</p>
                <p className="text-xs text-gray-400 truncate">{p.funcao} · {p.obra}</p></div>
              <div className="text-right shrink-0">
                <span className="font-semibold text-sm px-2.5 py-1 rounded-lg" style={{ background: "#FDECEC", color: DANGER }}>{p.qtd} falta{p.qtd > 1 ? "s" : ""}</span>
                <p className="text-xs text-gray-400 mt-1">{minToHHMM(p.horas)} perdidas</p>
              </div>
            </div>
          ))}
          {filteredFaltas.length === 0 && (
            <p className="text-sm text-gray-400 py-4 text-center">
              {searchTerm ? t("ninguemSituacao") : t("nenhumaFaltaPeriodoCelebration")}
            </p>
          )}
        </div>
      </Card>
    </>
  );
}

function AbaAbono({ a, onSel }) {
  const max = a.abonoRank.length ? a.abonoRank[0].qtd : 1;
  return (
    <Card className="p-6">
      <div className="flex items-center justify-between mb-1">
        <h3 className="font-semibold" style={{ color: NAVY_DARK, fontSize: 16 }}>{t("rankingAbonos")}</h3>
        <span className="font-bold" style={{ color: NAVY, fontSize: 22 }}>{a.totalAbonos}</span>
      </div>
      <p className="text-gray-400 text-xs mb-4">{t("cliqueMotivo")}</p>
      <div className="flex flex-col gap-3">
        {a.abonoRank.map((r, i) => (
          <div key={i} onClick={() => onSel(r.motivo)} className="cursor-pointer rounded-lg px-2 py-1.5 -mx-2 hover:bg-gray-50">
            <div className="flex items-center gap-2 mb-1">
              <span className="flex items-center justify-center rounded-lg text-white text-xs font-bold w-5 h-5"
                style={{ background: i === 0 ? "#D4AF37" : i === 1 ? "#9CA3AF" : i === 2 ? "#B87333" : NAVY_MED }}>{i + 1}</span>
              <span className="text-sm flex-1" style={{ color: NAVY_DARK }}>{r.motivo}</span>
              <span className="font-semibold text-sm" style={{ color: NAVY_DARK }}>{r.qtd}</span>
              <ChevronRight size={14} color="#94A3B8" />
            </div>
            <div className="rounded-full h-2 ml-7 overflow-hidden" style={{ background: "#EDF1F5" }}>
              <div className="h-full rounded-full bar-fill" style={{ width: `${(r.qtd / max) * 100}%`, background: NAVY }} />
            </div>
          </div>
        ))}
        {a.abonoRank.length === 0 && <p className="text-sm text-gray-400">{t("nenhumAbonoPeriodo")}</p>}
      </div>
    </Card>
  );
}

function AbaExtras({ a, ExtraRow }) {
  const [searchTerm, setSearchTerm] = useState("");
  const acima = a.extraPorPessoa.filter((p) => p.min >= LIMITE_EXTRA).length;
  const atencao = a.extraPorPessoa.filter((p) => p.min >= LIMITE_EXTRA * 0.75 && p.min < LIMITE_EXTRA).length;

  const filteredExtras = useMemo(() => {
    const term = searchTerm.trim().toLowerCase();
    if (!term) return a.extraPorPessoa;
    return a.extraPorPessoa.filter((p) =>
      p.nome.toLowerCase().includes(term) ||
      p.funcao.toLowerCase().includes(term)
    );
  }, [a.extraPorPessoa, searchTerm]);

  return (
    <>
      <div className="grid grid-cols-4 gap-4 mb-4 stagger">
        <Card className="p-5"><p className="text-gray-500 text-sm mb-1">{t("totalHoras")}</p><p className="font-bold" style={{ color: NAVY_DARK, fontSize: 30 }}>{minToHHMM(a.totalExtraMin)}</p></Card>
        <Card className="p-5"><p className="text-gray-500 text-sm mb-1">{t("funcionarios")}</p><p className="font-bold" style={{ color: NAVY_DARK, fontSize: 30 }}>{a.extraPorPessoa.length}</p></Card>
        <Card className="p-5"><p className="text-gray-500 text-sm mb-1">{t("emAtencao")}</p><p className="font-bold" style={{ color: WARN, fontSize: 30 }}>{atencao}</p></Card>
        <Card className="p-5"><p className="text-gray-500 text-sm mb-1">{t("noLimite")}</p><p className="font-bold" style={{ color: DANGER, fontSize: 30 }}>{acima}</p></Card>
      </div>
      <Card className="p-6">
        <div className="flex items-center justify-between mb-4">
          <h3 className="font-semibold" style={{ color: NAVY_DARK, fontSize: 16 }}>{t("horasExtrasFunc")}</h3>
          <span className="text-xs px-2 py-1 rounded-lg" style={{ background: NAVY_SOFT, color: NAVY }}>{t("limiteMensal")}</span>
        </div>

        {/* Barra de Pesquisa */}
        <div className="relative mb-5">
          <Search size={18} className="absolute left-3 top-3.5 text-gray-400" />
          <input
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            placeholder={t("buscarColaboradorCargo")}
            className="w-full pl-10 pr-3 py-2.5 rounded-xl text-sm outline-none border"
            style={{ borderColor: "#D8E0E8", color: NAVY_DARK }}
          />
        </div>

        <div className="flex flex-col">
          {filteredExtras.map((p, i) => <ExtraRow key={i} p={p} i={i} rank />)}
          {filteredExtras.length === 0 && <p className="text-sm text-gray-400 py-4 text-center">{t("ninguemSituacao")}</p>}
        </div>
      </Card>
    </>
  );
}

function AbaAbsenteismo({ a }) {
  const sem = a.absSemanas;
  const corIndice = (v) => v >= 40 ? DANGER : v >= 15 ? WARN : OK;
  const txtIndice = (v) => v >= 40 ? "Crítico" : v >= 15 ? "Atenção" : "Saudável";
  const idxGeral = a.absIndiceGeral;
  return (
    <>
      {/* Cards de topo */}
      <div className="grid grid-cols-4 gap-4 mb-4 stagger">
        <Card className="p-5" style={{ borderTop: `3px solid ${corIndice(idxGeral)}` }}>
          <p className="text-gray-500 text-sm mb-1">{t("indiceGeral")}</p>
          <p className="font-bold leading-none" style={{ color: corIndice(idxGeral), fontSize: 34 }}>{idxGeral.toFixed(1)}%</p>
          <p className="text-xs mt-2 font-medium" style={{ color: corIndice(idxGeral) }}>{txtIndice(idxGeral)}</p>
        </Card>
        <Card className="p-5"><p className="text-gray-500 text-sm mb-1">{t("ausenciasTotais")}</p>
          <p className="font-bold leading-none" style={{ color: NAVY_DARK, fontSize: 34 }}>{a.absTotalAusencias}</p>
          <p className="text-xs text-gray-400 mt-2">{t("faltasAtestados")}</p></Card>
        <Card className="p-5"><p className="text-gray-500 text-sm mb-1">{t("faltas")}</p>
          <p className="font-bold leading-none" style={{ color: DANGER, fontSize: 34 }}>{a.absTotalFaltas}</p>
          <p className="text-xs text-gray-400 mt-2">{t("noPeriodo")}</p></Card>
        <Card className="p-5"><p className="text-gray-500 text-sm mb-1">{t("atestadosMedicos")}</p>
          <p className="font-bold leading-none" style={{ color: WARN, fontSize: 34 }}>{a.absTotalAtestados}</p>
          <p className="text-xs text-gray-400 mt-2">{t("noPeriodo")}</p></Card>
      </div>

      {/* Gráfico de evolução */}
      <Card className="p-6 mb-4">
        <div className="flex items-center justify-between mb-1">
          <h3 className="font-semibold" style={{ color: NAVY_DARK, fontSize: 16 }}>{t("evolucaoAbsenteismo")}</h3>
          <span className="text-xs px-2 py-1 rounded-lg" style={{ background: NAVY_SOFT, color: NAVY }}>(Faltas + atestados) ÷ ({a.efetivoAtivo} ativos × dias úteis)</span>
        </div>
        <p className="text-gray-400 text-xs mb-4">Índice semanal sobre os dias-homem disponíveis (efetivo ativo × dias úteis)</p>
        <div className="flex gap-4 mb-3 text-xs">
          <span className="flex items-center gap-1"><span style={{ width: 10, height: 10, borderRadius: 2, background: NAVY }} />Índice (%)</span>
          <span className="flex items-center gap-1" style={{ color: OK }}>● até 15% saudável</span>
          <span className="flex items-center gap-1" style={{ color: WARN }}>● 15–40% atenção</span>
          <span className="flex items-center gap-1" style={{ color: DANGER }}>● acima de 40% crítico</span>
        </div>
        <div style={{ height: 240 }}>
          <ResponsiveContainer>
            <LineChart data={sem} margin={{ left: -18, right: 10, top: 6 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#EDF1F5" vertical={false} />
              <XAxis dataKey="semLabel" tick={{ fontSize: 12, fill: "#5B6B7B" }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: "#94A3B8" }} axisLine={false} tickLine={false} unit="%" />
              <Tooltip
                formatter={(v) => [`${v.toFixed(1)}%`, "Índice"]}
                labelFormatter={(l, p) => p && p[0] ? `${l} · ${p[0].payload.range}` : l}
                contentStyle={{ background: "#fff", borderRadius: 12, border: "1px solid #E2E8F0" }}
              />
              <Line
                type="linear"
                dataKey="indice"
                stroke={NAVY}
                strokeWidth={3}
                dot={{ r: 5, stroke: NAVY, strokeWidth: 2, fill: "#fff" }}
                activeDot={{ r: 7 }}
              />
            </LineChart>
          </ResponsiveContainer>
        </div>
      </Card>

      {/* Tabela por semana */}
      <Card className="p-6">
        <h3 className="font-semibold mb-4" style={{ color: NAVY_DARK, fontSize: 16 }}>{t("detalhamentoSemana")}</h3>
        <table className="w-full" style={{ fontSize: 14 }}>
          <thead>
            <tr style={{ borderBottom: `2px solid ${NAVY}` }}>
              <th className="text-left font-medium py-2" style={{ color: NAVY_DARK }}>{t("semana")}</th>
              <th className="text-center font-medium py-2" style={{ color: NAVY_DARK }}>{t("faltas")}</th>
              <th className="text-center font-medium py-2" style={{ color: NAVY_DARK }}>{t("atestados")}</th>
              <th className="text-center font-medium py-2" style={{ color: NAVY_DARK }}>{t("diasUteis")}</th>
              <th className="text-right font-medium py-2" style={{ color: NAVY_DARK }}>Índice</th>
            </tr>
          </thead>
          <tbody>
            {sem.map((w, i) => (
              <tr key={i} style={{ borderBottom: "0.5px solid #EDF1F5" }}>
                <td className="py-3">
                  <span className="font-medium" style={{ color: NAVY_DARK }}>{w.semLabel}</span>
                  <span className="text-gray-400 ml-2 text-sm">{w.range}</span>
                </td>
                <td className="text-center py-3" style={{ color: "#5B6B7B" }}>{w.faltas}</td>
                <td className="text-center py-3" style={{ color: "#5B6B7B" }}>{w.atestados}</td>
                <td className="text-center py-3" style={{ color: "#5B6B7B" }}>{w.diasUteis}</td>
                <td className="text-right py-3">
                  <span className="font-semibold px-2.5 py-1 rounded-lg"
                    style={{ background: corIndice(w.indice) + "22", color: corIndice(w.indice) }}>{w.indice.toFixed(1)}%</span>
                </td>
              </tr>
            ))}
            {sem.length === 0 && <tr><td colSpan={5} className="py-4 text-gray-400 text-sm">{t("semDatasValidas")}</td></tr>}
          </tbody>
        </table>
        <p className="text-xs text-gray-400 mt-4">O índice usa os dias-homem disponíveis (efetivo ativo × dias úteis da semana), por isso semanas de tamanhos diferentes ficam comparáveis entre si.</p>
      </Card>
    </>
  );
}

/* ====== ABA FÉRIAS ====== */
function AbaFerias({ feriasData, feriasFileName, onNav }) {
  const [searchTerm, setSearchTerm] = useState("");

  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  const getUrgencia = (dl) => {
    if (!dl) return { label: "Sem data", color: "#94A3B8", bg: "#F1F5F9" };
    const diff = Math.round((dl.getTime() - hoje.getTime()) / (1000 * 60 * 60 * 24));
    if (diff < 0) return { label: "Vencida", color: "#fff", bg: DANGER };
    if (diff === 0) return { label: "Vence hoje", color: "#fff", bg: DANGER };
    if (diff <= 30) return { label: `${diff}d restantes`, color: DANGER, bg: "#FDECEC" };
    if (diff <= 90) return { label: `${diff}d restantes`, color: "#B45309", bg: "#FEF3C7" };
    return { label: `${diff}d restantes`, color: OK, bg: "#E1F5EE" };
  };

  const normalizedFeriasData = useMemo(() => {
    return (feriasData || []).map((p) => {
      const dl = ensureDate(p.dataLimite);
      return {
        ...p,
        dataLimite: dl,
        dataLimiteStr: dl ? `${String(dl.getDate()).padStart(2, "0")}/${String(dl.getMonth() + 1).padStart(2, "0")}/${dl.getFullYear()}` : p.dataLimiteStr || "—"
      };
    }).sort((a, b) => {
      if (!a.dataLimite && !b.dataLimite) return 0;
      if (!a.dataLimite) return 1;
      if (!b.dataLimite) return -1;
      return a.dataLimite.getTime() - b.dataLimite.getTime();
    });
  }, [feriasData]);

  const filteredData = useMemo(() => {
    const term = searchTerm.trim().toLowerCase();
    if (!term) return normalizedFeriasData;
    return normalizedFeriasData.filter((p) =>
      p.nome.toLowerCase().includes(term) ||
      p.funcao.toLowerCase().includes(term)
    );
  }, [normalizedFeriasData, searchTerm]);

  const data948 = useMemo(() => filteredData.filter((p) => p.secao.includes("948")), [filteredData]);
  const data935 = useMemo(() => filteredData.filter((p) => p.secao.includes("935")), [filteredData]);

  const vencidas = normalizedFeriasData.filter((p) => p.dataLimite && p.dataLimite < hoje).length;
  const proximas = normalizedFeriasData.filter((p) => {
    if (!p.dataLimite) return false;
    const diff = Math.ceil((p.dataLimite.getTime() - hoje.getTime()) / (1000 * 60 * 60 * 24));
    return diff >= 0 && diff <= 30;
  }).length;

  if (!feriasData.length) {
    return (
      <Card className="p-8 text-center">
        <div className="flex justify-center mb-4">
          <div className="rounded-2xl p-4" style={{ background: NAVY_SOFT }}><Palmtree size={40} color={NAVY} /></div>
        </div>
        <h3 className="font-semibold mb-2" style={{ color: NAVY_DARK, fontSize: 18 }}>{t("nenhumaFeriasImportada")}</h3>
        <p className="text-gray-400 text-sm mb-4">{t("importePlanilhaFeriasAbaImportar")}</p>
        <button onClick={onNav} className="py-2.5 px-5 rounded-xl text-white font-medium text-sm" style={{ background: NAVY }}>
          {t("irParaImportar")}
        </button>
      </Card>
    );
  }

  const renderRow = (p, i) => {
    const urg = getUrgencia(p.dataLimite);
    return (
      <div key={i} className="flex items-center gap-3 py-3" style={{ borderBottom: "0.5px solid #EDF1F5" }}>
        <span className="font-bold text-gray-400 w-6 text-sm text-center shrink-0">{i + 1}</span>
        <Initials name={p.nome} bg={NAVY_MED} />
        <div className="flex-1 min-w-0">
          <p className="font-semibold text-sm truncate" style={{ color: NAVY_DARK }}>{p.nome}</p>
          <p className="text-xs text-gray-400 truncate">{p.funcao}</p>
        </div>
        <div className="text-right shrink-0 flex flex-col items-end gap-1">
          <div className="flex items-center gap-2">
            <span className="text-xs text-gray-400">{t("saldo")}</span>
            <span className="font-semibold text-sm" style={{ color: NAVY_DARK }}>{p.saldo}</span>
          </div>
          <div className="flex items-center gap-2">
            <span className="text-xs text-gray-400">{t("limite")}</span>
            <span className="font-medium text-xs px-2 py-0.5 rounded-full" style={{ background: urg.bg, color: urg.color }}>{p.dataLimiteStr}</span>
          </div>
          <span className="text-xs font-medium px-2 py-0.5 rounded-full" style={{ background: urg.bg, color: urg.color }}>{urg.label}</span>
        </div>
      </div>
    );
  };

  const SecaoBlock = ({ titulo, dados }) => (
    dados.length > 0 && (
      <Card className="p-6 mb-4">
        <div className="flex items-center justify-between mb-4">
          <h3 className="font-semibold" style={{ color: NAVY_DARK, fontSize: 16 }}>{titulo}</h3>
          <span className="text-xs px-2.5 py-1 rounded-lg font-medium" style={{ background: NAVY_SOFT, color: NAVY }}>{dados.length} funcionário{dados.length > 1 ? "s" : ""}</span>
        </div>
        <div className="flex flex-col">{dados.map(renderRow)}</div>
      </Card>
    )
  );

  return (
    <>
      <div className="grid grid-cols-3 gap-4 mb-4 stagger">
        <Card className="p-5">
          <p className="text-gray-500 text-sm mb-1">{t("totalFuncionarios")}</p>
          <p className="font-bold" style={{ color: NAVY_DARK, fontSize: 32 }}>{feriasData.length}</p>
        </Card>
        <Card className="p-5" style={{ borderTop: `3px solid ${DANGER}` }}>
          <p className="text-gray-500 text-sm mb-1">{t("feriasVencidas")}</p>
          <p className="font-bold" style={{ color: DANGER, fontSize: 32 }}>{vencidas}</p>
        </Card>
        <Card className="p-5" style={{ borderTop: `3px solid ${WARN}` }}>
          <p className="text-gray-500 text-sm mb-1">{t("vencemEm30Dias")}</p>
          <p className="font-bold" style={{ color: WARN, fontSize: 32 }}>{proximas}</p>
        </Card>
      </div>

      <Card className="p-4 mb-4">
        <div className="relative">
          <Search size={18} className="absolute left-3 top-3 text-gray-400" />
          <input
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            placeholder={t("buscarNomeFuncao")}
            className="w-full pl-10 pr-3 py-2.5 rounded-xl text-sm outline-none border"
            style={{ borderColor: "#D8E0E8", color: NAVY_DARK }}
          />
        </div>
      </Card>

      <SecaoBlock titulo="OBRA — 948" dados={data948} />
      <SecaoBlock titulo="OBRA — 935" dados={data935} />

      {filteredData.length === 0 && (
        <Card className="p-8 text-center">
          <p className="text-gray-400 text-sm">{searchTerm ? t("ninguemEncontradoFiltro") : t("nenhumFuncionarioEncontradoSecoes")}</p>
        </Card>
      )}

      <p className="text-xs text-gray-400 text-center mt-2">{t("planilhaLabel")}{feriasFileName}</p>
    </>
  );
}

/* ====== ABA IMPORTAR ====== */
function AbaImportar({ onImport, onDemo, temDados, fileName, onFeriasImport, temFeriasData, feriasFileName }) {
  const inputRef = useRef();
  const feriasInputRef = useRef();
  const [drag, setDrag] = useState(false);
  const [dragF, setDragF] = useState(false);
  const [err, setErr] = useState("");
  const [errF, setErrF] = useState("");

  const handleFile = (file) => {
    setErr("");
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const sheetName = wb.SheetNames.at(0);
        const sheet = Reflect.get(wb.Sheets, sheetName);
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
        const dados = parseSheet(rows);
        if (!dados.length) throw new Error();
        onImport(dados, file.name);
      } catch { setErr("Não consegui ler esta planilha. Confira se é o relatório de ponto no formato esperado."); }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleFeriasFile = (file) => {
    setErrF("");
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const sheetName = wb.SheetNames.at(0);
        const sheet = Reflect.get(wb.Sheets, sheetName);
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
        const dados = parseFeriasSheet(rows);
        if (!dados.length) throw new Error("Nenhum dado encontrado para seções 948/935.");
        onFeriasImport(dados, file.name);
      } catch (ex) { setErrF(ex.message || "Não consegui ler esta planilha. Confira se é o relatório de férias no formato esperado."); }
    };
    reader.readAsArrayBuffer(file);
  };

  return (
    <div className="flex justify-center">
      <div className="w-full" style={{ maxWidth: 600 }}>
        {/* --- Planilha de Ponto --- */}
        <h3 className="font-semibold mb-3 flex items-center gap-2" style={{ color: NAVY_DARK, fontSize: 16 }}>
          <FileBarChart size={20} color={NAVY} /> Planilha de Ponto
        </h3>
        {temDados && (
          <div className="mb-3 p-3 rounded-xl text-sm flex items-center gap-2" style={{ background: "#E1F5EE", color: OK }}>
            <CheckCircle2 size={16} /> Planilha ativa: <b>{fileName}</b>. Importe outra para substituir.
          </div>
        )}
        <div onDragOver={(e) => { e.preventDefault(); setDrag(true); }} onDragLeave={() => setDrag(false)}
          onDrop={(e) => { e.preventDefault(); setDrag(false); if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); }}
          onClick={() => inputRef.current.click()}
          className="bg-white rounded-2xl p-8 cursor-pointer text-center"
          style={{ border: drag ? `3px dashed ${NAVY_MED}` : "3px dashed #E1E8EF" }}>
          <div className="flex justify-center mb-4">
            <div className="rounded-xl p-3" style={{ background: NAVY_SOFT }}><Upload size={28} color={NAVY} /></div>
          </div>
          <h2 className="text-base font-semibold mb-1" style={{ color: NAVY_DARK }}>{t("anexePlanilha")}</h2>
          <p className="text-gray-400 text-xs mb-4">{t("arrasteOuCliqueSelecionar")}</p>
          <span className="inline-block text-white font-medium px-5 py-2.5 rounded-xl text-sm" style={{ background: NAVY }}>{t("selecionarPlanilha")}</span>
          <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden"
            onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])} />
        </div>
        {err && <p className="text-sm mt-3 rounded-lg py-2 px-3" style={{ background: "#FDECEC", color: DANGER }}>{err}</p>}
        <div className="text-center mt-3 mb-8">
          <button onClick={onDemo} className="text-sm underline" style={{ color: NAVY }}>{t("verDemonstracao")}</button>
        </div>

        {/* --- Planilha de Férias --- */}
        <div style={{ borderTop: "1px solid #EDF1F5", paddingTop: 24 }}>
          <h3 className="font-semibold mb-3 flex items-center gap-2" style={{ color: NAVY_DARK, fontSize: 16 }}>
            <Palmtree size={20} color={NAVY_MED} /> Planilha de Férias
          </h3>
          {temFeriasData && (
            <div className="mb-3 p-3 rounded-xl text-sm flex items-center gap-2" style={{ background: "#E1F5EE", color: OK }}>
              <CheckCircle2 size={16} /> Planilha de férias ativa: <b>{feriasFileName}</b>. Importe outra para substituir.
            </div>
          )}
          <div onDragOver={(e) => { e.preventDefault(); setDragF(true); }} onDragLeave={() => setDragF(false)}
            onDrop={(e) => { e.preventDefault(); setDragF(false); if (e.dataTransfer.files[0]) handleFeriasFile(e.dataTransfer.files[0]); }}
            onClick={() => feriasInputRef.current.click()}
            className="bg-white rounded-2xl p-8 cursor-pointer text-center"
            style={{ border: dragF ? `3px dashed ${NAVY_MED}` : "3px dashed #E1E8EF" }}>
            <div className="flex justify-center mb-4">
              <div className="rounded-xl p-3" style={{ background: "#FEF3C7" }}><Palmtree size={28} color="#B45309" /></div>
            </div>
            <h2 className="text-base font-semibold mb-1" style={{ color: NAVY_DARK }}>{t("anexePlanilhaFerias")}</h2>
            <p className="text-gray-400 text-xs mb-4">{t("arrasteArquivoFerias")}</p>
            <span className="inline-block text-white font-medium px-5 py-2.5 rounded-xl text-sm" style={{ background: NAVY_MED }}>{t("selecionarPlanilhaFerias")}</span>
            <input ref={feriasInputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden"
              onChange={(e) => e.target.files[0] && handleFeriasFile(e.target.files[0])} />
          </div>
          {errF && <p className="text-sm mt-3 rounded-lg py-2 px-3" style={{ background: "#FDECEC", color: DANGER }}>{errF}</p>}
        </div>
      </div>
    </div>
  );
}

function AbaRelatorios({ a, exportar }) {
  return (
    <Card className="p-6">
      <h3 className="font-semibold mb-1" style={{ color: NAVY_DARK, fontSize: 16 }}>{t("resumoGeral")}</h3>
      <p className="text-gray-400 text-sm mb-5">{t("todosIndicadores")}</p>
      <div className="grid grid-cols-2 gap-3 mb-5">
        {[
          ["Ativos (A)", a.efetivo.A, OK], ["Férias (F)", a.efetivo.F, NAVY_MED],
          ["Afastados INSS (P)", a.efetivo.P, WARN], ["Demitidos (D)", a.efetivo.D, DANGER],
          ["Total de faltas", a.totalFaltas, DANGER], ["Horas perdidas", minToHHMM(a.totalHorasFalta), NAVY_DARK],
          ["Total de abonos", a.totalAbonos, NAVY], ["Horas extras totais", minToHHMM(a.totalExtraMin), NAVY_DARK],
        ].map(([l, v, c], i) => (
          <div key={i} className="flex items-center justify-between p-3 rounded-xl" style={{ background: "#F8FAFC" }}>
            <span className="text-sm text-gray-600">{l}</span>
            <span className="font-bold text-sm" style={{ color: c }}>{v}</span>
          </div>
        ))}
      </div>
      <button onClick={exportar} className="w-full py-3 rounded-xl btn-primary text-white font-medium text-sm flex items-center justify-center gap-2" style={{ background: NAVY }}>
        <Download size={16} /> {t("copiarResumoCompleto")}</button>
    </Card>
  );
}

/* ====== LOGIN / CADASTRO ====== */
function AuthScreen({ onAuth }) {
  const [mode, setMode] = useState("login");   // login | register
  const [step, setStep] = useState("cred");     // cred | token
  const [user, setUser] = useState("");
  const [pass, setPass] = useState("");
  const [token, setToken] = useState("");
  const [err, setErr] = useState("");
  const [msg, setMsg] = useState("");
  const [rNome, setRNome] = useState("");
  const [rUser, setRUser] = useState("");
  const [rPass, setRPass] = useState("");
  const [rPass2, setRPass2] = useState("");
  const [rRole, setRRole] = useState("Gerente");
  const [pendingUser, setPendingUser] = useState(null);

  const inputStyle = {
    width: "100%", padding: "11px 13px", borderRadius: 12, fontSize: 14,
    border: "1px solid #D8E0E8", outline: "none", background: "#fff", color: NAVY_DARK,
  };
  const Label = ({ children }) => <label className="block text-sm font-medium mb-1.5" style={{ color: NAVY_DARK }}>{children}</label>;

  const doCred = async () => {
    if (!user.trim() || !pass) { setErr("Preencha todos os campos."); return; }
    try {
      setErr("");
      setMsg("");
      const email = `${user.trim().toLowerCase()}@apia.com.br`;
      const userCredential = await signInWithEmailAndPassword(auth, email, pass);
      const uid = userCredential.user.uid;
      const userDoc = await getDoc(doc(db, "users", uid));
      
      if (userDoc.exists()) {
        const uData = { uid, ...userDoc.data() };
        if (uData.role === "Administrador") {
          // Administrador exige o código de verificação
          setPendingUser(uData);
          setStep("token");
        } else {
          // Gerente faz login direto, sem código!
          onAuth(uData);
        }
      } else {
        setErr("Usuário autenticado, mas perfil não encontrado.");
      }
    } catch (e) {
      console.error("Erro no login:", e);
      setErr("Usuário ou senha incorretos.");
    }
  };

  const doToken = () => {
    if (token.trim().toLowerCase() !== "ponto") { setErr("Código de verificação inválido."); return; }
    if (pendingUser) {
      onAuth(pendingUser);
      setPendingUser(null);
    } else {
      setErr("Sessão expirada. Tente logar novamente.");
      setStep("cred");
    }
  };

  const doRegister = async () => {
    if (!rNome.trim() || !rUser.trim() || !rPass) { setErr("Preencha todos os campos."); return; }
    if (rPass.length < 6) { setErr("A senha deve ter ao menos 6 caracteres."); return; }
    if (rPass !== rPass2) { setErr("As senhas não conferem."); return; }
    
    try {
      setErr("");
      const email = `${rUser.trim().toLowerCase()}@apia.com.br`;
      const userCredential = await createUserWithEmailAndPassword(auth, email, rPass);
      const uid = userCredential.user.uid;
      const profile = { user: rUser.trim().toLowerCase(), nome: rNome.trim(), role: rRole };
      await setDoc(doc(db, "users", uid), profile);
      
      setMode("login");
      setStep("cred");
      setErr("");
      setMsg("Conta criada! Faça login para entrar.");
      setUser(rUser.trim().toLowerCase());
      setPass("");
      setRNome("");
      setRUser("");
      setRPass("");
      setRPass2("");
      setRRole("Gerente");
    } catch (e) {
      if (e.code === "auth/email-already-in-use") {
        setErr("Este usuário já existe.");
      } else {
        setErr("Erro ao criar conta. Tente novamente.");
      }
    }
  };

  return (
    <div className="min-h-screen flex" style={{ fontFamily: "system-ui, sans-serif" }}>
      {/* PAINEL ESQUERDO — marca */}
      <div className="hidden md:flex flex-col justify-between p-10 relative overflow-hidden" style={{ width: "44%",
        background: `radial-gradient(circle at 20% 30%, #0A4E84 0%, ${NAVY} 45%, ${NAVY_DARK} 100%)` }}>
        {/* textura de partículas */}
        <svg className="absolute inset-0 w-full h-full" style={{ opacity: 0.5, animation: "floatParticles 6s ease-in-out infinite alternate" }} preserveAspectRatio="none">
          {Array.from({ length: 60 }).map((_, i) => {
            const x = (i * 137.5) % 100, y = (i * 53.7) % 100, r = (i % 4) + 1;
            return <circle key={i} cx={`${x}%`} cy={`${y}%`} r={r}
              fill={i % 3 === 0 ? "#E8A24A" : "#CFE0F0"} opacity={0.12 + (i % 5) * 0.04} />;
          })}
        </svg>
        <div className="relative flex items-center gap-2 text-white/90">
          <Menu size={20} /><span className="text-sm font-medium">{t("menu")}</span>
        </div>
        <div className="relative flex flex-col items-center">
          <div className="anim-pop" style={{ transform: "scale(1.6)" }}><ApiaLogo light /></div>
          <p className="text-white/60 text-sm mt-8 tracking-wide anim-fade" style={{ animationDelay: ".25s" }}>{t("sistemaGestao")}</p>
        </div>
      </div>

      {/* PAINEL DIREITO — formulário */}
      <div className="flex-1 flex items-center justify-center p-6 bg-white">
        <div className="w-full anim-up" style={{ maxWidth: 380 }}>
          <div className="flex md:hidden justify-center mb-8"><ApiaLogo /></div>

          {mode === "login" && step === "cred" && (
            <>
              <h2 className="font-bold mb-1" style={{ color: NAVY_DARK, fontSize: 28 }}>{t("login")}</h2>
              <p className="text-gray-400 text-sm mb-8">{t("acesseConta")}</p>
              <div className="mb-4"><Label>{t("usuario")}</Label>
                <input style={inputStyle} value={user} placeholder="seu.usuario"
                  onChange={(e) => setUser(e.target.value)} onKeyDown={(e) => e.key === "Enter" && doCred()} /></div>
              <div className="mb-2"><Label>{t("senha")}</Label>
                <input style={inputStyle} type="password" value={pass} placeholder="••••••••"
                  onChange={(e) => setPass(e.target.value)} onKeyDown={(e) => e.key === "Enter" && doCred()} /></div>
              <div className="flex justify-end mb-5">
                <button className="text-xs" style={{ color: NAVY_MED }}>{t("esqueceuSenha")}</button>
              </div>
              <button onClick={doCred} className="w-full py-3 rounded-xl btn-primary text-white font-medium text-sm flex items-center justify-center gap-2" style={{ background: NAVY }}>
                {t("entrar")} <ChevronRight size={16} /></button>
              <div className="flex items-center gap-3 my-6"><div className="flex-1 h-px" style={{ background: "#EDF1F5" }} /><span className="text-xs text-gray-400">ou</span><div className="flex-1 h-px" style={{ background: "#EDF1F5" }} /></div>
              <p className="text-center text-sm text-gray-500">Não tem uma conta?
                <button onClick={() => { setMode("register"); setErr(""); setMsg(""); }} className="font-semibold ml-1" style={{ color: NAVY_MED }}>{t("cadastreSe")}</button></p>
            </>
          )}

          {mode === "login" && step === "token" && (
            <>
              <button onClick={() => { setStep("cred"); setToken(""); setErr(""); }} className="text-sm mb-4 flex items-center gap-1" style={{ color: NAVY_MED }}>‹ {t("voltar")}</button>
              <h2 className="font-bold mb-1" style={{ color: NAVY_DARK, fontSize: 28 }}>{t("verificacao")}</h2>
              <p className="text-gray-400 text-sm mb-6">{t("digiteCodigo")}</p>
              <div className="mb-5"><Label>{t("verificacao")}</Label>
                <input style={{ ...inputStyle, letterSpacing: 4, textAlign: "center", fontSize: 18 }} value={token} placeholder="Código"
                  onChange={(e) => setToken(e.target.value)} onKeyDown={(e) => e.key === "Enter" && doToken()} /></div>
              <button onClick={doToken} className="w-full py-3 rounded-xl btn-primary text-white font-medium text-sm" style={{ background: NAVY }}>{t("verificarEntrar")}</button>
            </>
          )}

          {mode === "register" && (
            <>
              <h2 className="font-bold mb-1" style={{ color: NAVY_DARK, fontSize: 28 }}>{t("criarConta")}</h2>
              <p className="text-gray-400 text-sm mb-6">{t("cadastreAcesso")}</p>
              <div className="mb-3"><Label>{t("nome")}</Label>
                <input style={inputStyle} value={rNome} placeholder="Seu nome" onChange={(e) => setRNome(e.target.value)} /></div>
              <div className="mb-3"><Label>{t("usuario")}</Label>
                <input style={inputStyle} value={rUser} placeholder="seu.usuario" onChange={(e) => setRUser(e.target.value)} /></div>
              <div className="mb-3"><Label>{t("senha")}</Label>
                <input style={inputStyle} type="password" value={rPass} placeholder="mínimo 6 caracteres" onChange={(e) => setRPass(e.target.value)} /></div>
              <div className="mb-5"><Label>{t("confirmarSenha")}</Label>
                <input style={inputStyle} type="password" value={rPass2} placeholder="repita a senha"
                  onChange={(e) => setRPass2(e.target.value)} onKeyDown={(e) => e.key === "Enter" && doRegister()} /></div>
              <div className="mb-5"><Label>{t("perfilAcesso")}</Label>
                <div className="grid grid-cols-2 gap-2">
                  {["Administrador", "Gerente"].map((r) => (
                    <button key={r} onClick={() => setRRole(r)} className="py-2.5 rounded-xl text-sm font-medium transition-colors"
                      style={{ background: rRole === r ? NAVY : "#fff", color: rRole === r ? "#fff" : "#5B6B7B",
                        border: `1px solid ${rRole === r ? NAVY : "#D8E0E8"}` }}>{r}</button>
                  ))}
                </div>
                <p className="text-xs text-gray-400 mt-2">{rRole === "Administrador" ? "Acesso total: importa e gerencia os dados." : "Somente visualização dos dados importados."}</p>
              </div>
              <button onClick={doRegister} className="w-full py-3 rounded-xl btn-primary text-white font-medium text-sm" style={{ background: NAVY }}>{t("cadastrar")}</button>
              <p className="text-center text-sm text-gray-500 mt-6">Já tem conta?
                <button onClick={() => { setMode("login"); setErr(""); }} className="font-semibold ml-1" style={{ color: NAVY_MED }}>{t("fazerLogin")}</button></p>
            </>
          )}

          {err && <p className="text-sm mt-4 rounded-lg py-2 px-3" style={{ background: "#FDECEC", color: DANGER }}>{err}</p>}
          {msg && <p className="text-sm mt-4 rounded-lg py-2 px-3" style={{ background: "#E1F5EE", color: OK }}>{msg}</p>}
        </div>
      </div>
    </div>
  );
}

/* ====== EFEITOS / ANIMAÇÕES GLOBAIS ====== */
function GlobalStyles() {
  return (
    <style>{`
      @keyframes fadeUp { from { opacity: 0; transform: translateY(14px); } to { opacity: 1; transform: none; } }
      @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
      @keyframes scaleIn { from { opacity: 0; transform: scale(.95); } to { opacity: 1; transform: none; } }
      @keyframes slideInLeft { from { opacity: 0; transform: translateX(-24px); } to { opacity: 1; transform: none; } }
      @keyframes popIn { 0% { opacity: 0; transform: scale(.8); } 60% { transform: scale(1.04); } 100% { opacity: 1; transform: scale(1); } }
      @keyframes floatParticles { from { transform: translateY(0); } to { transform: translateY(-12px); } }

      .anim-card { animation: fadeUp .55s cubic-bezier(.22,1,.36,1) both; transition: transform .25s ease, box-shadow .25s ease; }
      .anim-card:hover { transform: translateY(-3px); box-shadow: 0 12px 30px rgba(4,58,102,.12) !important; }
      .anim-fade { animation: fadeIn .4s ease both; }
      .anim-up { animation: fadeUp .5s cubic-bezier(.22,1,.36,1) both; }
      .anim-scale { animation: scaleIn .32s cubic-bezier(.22,1,.36,1) both; }
      .anim-pop { animation: popIn .5s cubic-bezier(.22,1,.36,1) both; }
      .anim-slide { animation: slideInLeft .45s cubic-bezier(.22,1,.36,1) both; }

      .bar-fill { transition: width 1s cubic-bezier(.22,1,.36,1); }

      .row-hov { transition: background-color .18s ease, transform .18s ease; border-radius: 10px; }
      .row-hov:hover { background: #F6F9FC; transform: translateX(2px); }

      button { transition: transform .14s ease, filter .18s ease, background-color .2s ease, color .2s ease, box-shadow .2s ease; }
      button:active { transform: scale(.97); }
      .btn-primary:hover { filter: brightness(1.12); box-shadow: 0 6px 18px rgba(4,58,102,.28); }

      input { transition: border-color .2s ease, box-shadow .2s ease; }
      input:focus { border-color: ${NAVY_MED} !important; box-shadow: 0 0 0 3px rgba(30,108,168,.15); }

      .nav-item { transition: background-color .2s ease, color .2s ease, transform .15s ease; }
      .nav-item:hover { transform: translateX(2px); }

      /* stagger automático dos cards dentro de grids */
      .stagger > * { animation: fadeUp .5s cubic-bezier(.22,1,.36,1) both; }
      .stagger > *:nth-child(1){animation-delay:.04s}
      .stagger > *:nth-child(2){animation-delay:.10s}
      .stagger > *:nth-child(3){animation-delay:.16s}
      .stagger > *:nth-child(4){animation-delay:.22s}
      .stagger > *:nth-child(5){animation-delay:.28s}
      .stagger > *:nth-child(6){animation-delay:.34s}
    `}</style>
  );
}

/* ====== ROOT ====== */
export default function App() {
  const [conta, setConta] = useState(null);
  const [data, setData] = useState([]);
  const [fileName, setFileName] = useState("Nenhuma planilha importada");
  const [feriasData, setFeriasData] = useState([]);
  const [feriasFileName, setFeriasFileName] = useState("Nenhuma planilha de férias");
  const [loading, setLoading] = useState(true);

  // Escutar estado de autenticação e perfil em tempo real
  useEffect(() => {
    let unsubUserDoc = null;
    const unsubAuth = onAuthStateChanged(auth, (firebaseUser) => {
      if (unsubUserDoc) {
        unsubUserDoc();
        unsubUserDoc = null;
      }

      if (firebaseUser) {
        // Tentar ler perfil cacheado no localStorage para carregamento instantâneo
        const cached = localStorage.getItem(`apia_profile_${firebaseUser.uid}`);
        if (cached) {
          try {
            setConta(JSON.parse(cached));
            setLoading(false);
          } catch (e) {
            console.error("Erro ao ler cache do perfil:", e);
          }
        } else {
          // Se não houver cache, define estado padrão imediato para pular tela de carregamento
          setConta({ uid: firebaseUser.uid, nome: firebaseUser.displayName || "Usuário", role: "Gerente" });
          setLoading(false);
        }

        try {
          unsubUserDoc = onSnapshot(doc(db, "users", firebaseUser.uid), (docSnap) => {
            if (docSnap.exists()) {
              const uData = { uid: firebaseUser.uid, ...docSnap.data() };
              setConta(uData);
              localStorage.setItem(`apia_profile_${firebaseUser.uid}`, JSON.stringify(uData));
            } else {
              const uData = { uid: firebaseUser.uid, nome: firebaseUser.displayName || "Usuário", role: "Gerente" };
              setConta(uData);
              localStorage.setItem(`apia_profile_${firebaseUser.uid}`, JSON.stringify(uData));
            }
            setLoading(false);
          }, (err) => {
            console.error("Erro em tempo real no onSnapshot:", err);
            setConta({ uid: firebaseUser.uid, nome: firebaseUser.displayName || "Usuário", role: "Gerente" });
            setLoading(false);
          });
        } catch (e) {
          console.error("Erro síncrono ao registrar listener de usuário:", e);
          setConta({ uid: firebaseUser.uid, nome: firebaseUser.displayName || "Usuário", role: "Gerente" });
          setLoading(false);
        }
      } else {
        setConta(null);
        setLoading(false);
      }
    });
    return () => {
      unsubAuth();
      if (unsubUserDoc) unsubUserDoc();
    };
  }, []);

  // Escutar dados sincronizados da planilha ativa
  useEffect(() => {
    if (!conta) return;

    // Tentar ler cache local da planilha ativa
    const cachedSheet = localStorage.getItem("apia_active_sheet");
    if (cachedSheet) {
      try {
        const sData = JSON.parse(cachedSheet);
        setData(sData.data || []);
        setFileName(sData.fileName || "Nenhuma planilha importada");
      } catch (e) {
        console.error("Erro ao ler cache da planilha:", e);
      }
    }

    const unsubSheet = onSnapshot(doc(db, "dashboard", "active_sheet"), (docSnap) => {
      if (docSnap.exists()) {
        const sheetData = docSnap.data();
        setData(sheetData.data || []);
        setFileName(sheetData.fileName || "Nenhuma planilha importada");
        localStorage.setItem("apia_active_sheet", JSON.stringify({
          data: sheetData.data || [],
          fileName: sheetData.fileName || "Nenhuma planilha importada"
        }));
      } else {
        setData([]);
        setFileName("Nenhuma planilha importada");
        localStorage.removeItem("apia_active_sheet");
      }
    });
    return () => unsubSheet();
  }, [conta]);

  // Escutar dados sincronizados da planilha de férias ativa
  useEffect(() => {
    if (!conta) return;

    // Tentar ler cache local da planilha de férias
    const cachedFerias = localStorage.getItem("apia_active_ferias");
    if (cachedFerias) {
      try {
        const fData = JSON.parse(cachedFerias);
        setFeriasData(fData.data || []);
        setFeriasFileName(fData.fileName || "Nenhuma planilha de férias");
      } catch (e) {
        console.error("Erro ao ler cache de férias:", e);
      }
    }

    const unsubFerias = onSnapshot(doc(db, "dashboard", "active_ferias"), (docSnap) => {
      if (docSnap.exists()) {
        const sheetData = docSnap.data();
        setFeriasData(sheetData.data || []);
        setFeriasFileName(sheetData.fileName || "Nenhuma planilha de férias");
        localStorage.setItem("apia_active_ferias", JSON.stringify({
          data: sheetData.data || [],
          fileName: sheetData.fileName || "Nenhuma planilha de férias"
        }));
      } else {
        setFeriasData([]);
        setFeriasFileName("Nenhuma planilha de férias");
        localStorage.removeItem("apia_active_ferias");
      }
    });
    return () => unsubFerias();
  }, [conta]);

  const handleDataUpdate = async (newData, name) => {
    if (conta?.role !== "Administrador") return;
    try {
      await setDoc(doc(db, "dashboard", "active_sheet"), {
        data: newData,
        fileName: name,
        updatedAt: serverTimestamp(),
        updatedBy: conta.uid
      });
    } catch (e) {
      console.error("Erro ao salvar dados no Firebase:", e);
      alert("Erro ao sincronizar dados com o Firebase.");
    }
  };

  const handleFeriasUpdate = async (newData, name) => {
    if (conta?.role !== "Administrador") return;
    try {
      await setDoc(doc(db, "dashboard", "active_ferias"), {
        data: newData,
        fileName: name,
        updatedAt: serverTimestamp(),
        updatedBy: conta.uid
      });
    } catch (e) {
      console.error("Erro ao salvar férias no Firebase:", e);
      alert("Erro ao sincronizar dados de férias com o Firebase.");
    }
  };

  const handleReset = async () => {
    if (conta?.role !== "Administrador") return;
    try {
      await setDoc(doc(db, "dashboard", "active_sheet"), {
        data: [],
        fileName: "Nenhuma planilha importada",
        updatedAt: serverTimestamp(),
        updatedBy: conta.uid
      });
      await setDoc(doc(db, "dashboard", "active_ferias"), {
        data: [],
        fileName: "Nenhuma planilha de férias",
        updatedAt: serverTimestamp(),
        updatedBy: conta.uid
      });
    } catch (e) {
      console.error("Erro ao limpar dados no Firebase:", e);
    }
  };

  const handleLogout = async () => {
    try {
      await signOut(auth);
    } catch (e) {
      console.error("Erro ao sair:", e);
    }
  };

  const handleUpdateProfile = async (profileData) => {
    if (!conta) return;
    setConta(prev => ({ ...prev, ...profileData }));
    localStorage.setItem(`apia_profile_${conta.uid}`, JSON.stringify({ ...conta, ...profileData }));
    
    // Grava em background sem bloquear o fechamento do modal
    const userRef = doc(db, "users", conta.uid);
    setDoc(userRef, profileData, { merge: true }).catch((e) => {
      console.warn("Erro ao sincronizar perfil com Firestore (salvo localmente):", e);
    });
  };

  if (loading) {
    return (
      <div className="min-h-screen flex flex-col items-center justify-center relative overflow-hidden" style={{ background: "#043A66", fontFamily: "system-ui, sans-serif" }}>
        {/* Efeito de brilho de fundo */}
        <div className="absolute w-[500px] h-[500px] rounded-full blur-[120px] opacity-20" style={{ background: "#1E6CA8", top: "20%", left: "30%" }} />
        
        <div className="flex flex-col items-center gap-6 relative z-10">
          {/* Container circular com bordas e efeito de pulso */}
          <div className="relative flex items-center justify-center p-6 rounded-full bg-white/5 border border-white/10 shadow-2xl animate-pulse">
            {/* Anel giratório */}
            <div className="absolute inset-0 rounded-full border-2 border-t-white border-r-white/20 border-b-white/20 border-l-white/20 animate-spin" style={{ margin: "-4px" }} />
            
            {/* Logo grande da Ápia animada */}
            <svg width="64" height="64" viewBox="0 0 100 100" className="animate-bounce" style={{ animationDuration: '2s' }}>
              <circle cx="50" cy="50" r="48" fill="#fff" />
              <path d="M50 22 L78 74 L22 74 Z" fill="none" stroke="#043A66" strokeWidth="6" />
              <path d="M50 42 L64 68 L36 68 Z" fill="#043A66" />
            </svg>
          </div>
          
          {/* Identidade visual e texto de carregamento */}
          <div className="text-center">
            <h2 className="text-white font-bold text-lg tracking-wide uppercase">ÁPIA</h2>
            <p className="text-white/60 text-xs mt-1 font-medium tracking-widest uppercase">Gestão de Ponto</p>
          </div>
          
          {/* Bouncing dots de carregamento */}
          <div className="flex items-center gap-2 mt-2">
            <span className="text-white/80 text-sm font-medium">Carregando sistema</span>
            <span className="flex gap-1">
              <span className="w-1.5 h-1.5 rounded-full bg-white animate-bounce" style={{ animationDuration: '0.8s', animationDelay: '0ms' }} />
              <span className="w-1.5 h-1.5 rounded-full bg-white animate-bounce" style={{ animationDuration: '0.8s', animationDelay: '150ms' }} />
              <span className="w-1.5 h-1.5 rounded-full bg-white animate-bounce" style={{ animationDuration: '0.8s', animationDelay: '300ms' }} />
            </span>
          </div>
        </div>
      </div>
    );
  }

  return (
    <>
      <GlobalStyles />
      {!conta ? (
        <AuthScreen onAuth={(c) => setConta(c)} />
      ) : (
        <Dashboard
          data={data}
          fileName={fileName}
          conta={conta}
          onData={handleDataUpdate}
          onDemo={() => handleDataUpdate(demoData(), "Dados de demonstração")}
          onReset={handleReset}
          onLogout={handleLogout}
          onUpdateProfile={handleUpdateProfile}
          feriasData={feriasData}
          feriasFileName={feriasFileName}
          onFeriasData={handleFeriasUpdate}
        />
      )}
    </>
  );
}
