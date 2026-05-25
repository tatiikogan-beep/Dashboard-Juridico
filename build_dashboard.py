import base64
import os
import urllib.request
import streamlit as st

# ââ Embed external JS/CSS libraries inline for offline/Teams compatibility ââ
def _fetch_lib(url, fallback=""):
    import urllib.request
    try:
        with urllib.request.urlopen(url, timeout=15) as r:
            return r.read().decode("utf-8")
    except Exception:
        return fallback

XLSX_JS        = _fetch_lib("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js")
CHART_JS       = _fetch_lib("https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js")
DATALABELS_JS  = _fetch_lib("https://cdnjs.cloudflare.com/ajax/libs/chartjs-plugin-datalabels/2.2.0/chartjs-plugin-datalabels.min.js")


logo_path = 'igsa_b64.txt'
if os.path.exists(logo_path):
        with open(logo_path) as f:
                    LOGO = 'data:image/png;base64,' + f.read().strip()
else:
        LOGO = ''

# Build the HTML in Python parts to avoid any escaping issues
CSS = """
:root{
  --c9:#6B0010;--c8:#8B0E1A;--c7:#A81020;--c6:#C01020;--c5:#D42030;
  --c4:#E05060;--c3:#E89090;--c2:#F5C8CC;--c1:#FDE8EA;--c0:#FFF5F6;
  --gd:#C4A85A;--gd2:#E8D8A0;
  --bg:#F7F4F4;--wh:#FFF;--tx:#1A0810;--mu:#7A5060;--bd:#EDD8DA;
  --sh:0 2px 8px rgba(100,0,20,.1);--shm:0 4px 20px rgba(100,0,20,.15);
  --r:8px;--r2:14px;
}
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Inter',sans-serif;background:var(--bg);color:var(--tx);font-size:14px}
h1,h2,.sf{font-family:'Libre Baskerville',serif}
/* CONFIG */
#cfg{min-height:100vh;background:#fff;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:32px 20px}
.cfgh{text-align:center;margin-bottom:24px}
.cfgh img{height:84px;margin:0 auto 12px;display:block;filter:drop-shadow(0 4px 12px rgba(0,0,0,.4))}
.cfgh h1{color:var(--c9);font-size:1.7rem}
.cfgh p{color:var(--mu);font-size:.82rem;margin-top:5px}
.card{background:#fff;border-radius:var(--r2);box-shadow:0 8px 40px rgba(0,0,0,.28);padding:32px;width:100%;max-width:840px}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:24px}
.stl{font-family:'Libre Baskerville',serif;font-size:.88rem;font-weight:700;color:var(--c7);padding-bottom:7px;border-bottom:2px solid var(--c1);margin-bottom:12px;display:flex;align-items:center;gap:7px}
.sn{width:20px;height:20px;background:var(--c7);color:#fff;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:.65rem;font-weight:700;flex-shrink:0}
.dz{border:2px dashed var(--c3);border-radius:var(--r);padding:14px 10px;text-align:center;cursor:pointer;background:var(--c0);transition:.2s;position:relative;overflow:hidden}
.dz:hover{border-color:var(--c6);background:var(--c1)}
.dz.ok{border-style:solid;border-color:var(--c7);background:var(--c1)}
.dz input[type=file]{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
.dzi{font-size:1.5rem;margin-bottom:4px}
.dzl{font-size:.75rem;color:var(--c7);font-weight:600}
.dzs{font-size:.68rem;color:var(--mu);margin-top:2px}
.dzst{font-size:.68rem;color:var(--c8);font-weight:700;margin-top:4px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis}
.up3{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:16px}
.fld{margin-bottom:12px}
.fld label{display:block;font-size:.72rem;font-weight:600;color:var(--mu);margin-bottom:4px;text-transform:uppercase;letter-spacing:.04em}
.fld input{width:100%;padding:9px 12px;border:1.5px solid var(--bd);border-radius:var(--r);font-family:'Inter',sans-serif;font-size:.86rem;color:var(--tx);background:#fff;transition:border-color .2s;outline:none}
.fld input:focus{border-color:var(--c6)}
.fld small{display:block;font-size:.67rem;color:var(--mu);margin-top:3px}
.bgen{width:100%;padding:13px;background:var(--c7);color:#fff;border:none;border-radius:var(--r);font-family:'Inter',sans-serif;font-size:.92rem;font-weight:700;cursor:pointer;letter-spacing:.04em;transition:.2s;margin-top:4px}
.bgen:hover{background:var(--c9)}
.err{color:#900;font-size:.75rem;padding:8px 12px;background:#fff0f0;border-radius:6px;margin-top:8px;border-left:3px solid #c00;display:none}
/* TOOLBAR */
#dsh,#rpt{display:none}
.tbar{background:var(--c9);padding:9px 22px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:200;gap:10px}
.thi{font-size:.82rem;color:var(--c3);font-weight:500;letter-spacing:.01em}
.tbtns{display:flex;gap:7px}
.tb{padding:6px 13px;border-radius:6px;font-size:.72rem;font-weight:700;cursor:pointer;border:none;transition:.2s;letter-spacing:.03em}
.tbk{background:rgba(255,255,255,.15);color:#fff}.tbk:hover{background:rgba(255,255,255,.28)}
.trl{background:var(--gd);color:var(--c9)}.trl:hover{background:var(--gd2)}
.tpd{background:var(--c3);color:var(--c9)}.tpd:hover{background:#fff}
/* DASH */
.dw{padding:0 26px 36px;max-width:1360px;margin:0 auto}
.dh{display:flex;align-items:center;justify-content:space-between;padding:22px 0 16px;border-bottom:2px solid var(--c1);margin-bottom:20px;gap:14px}
.dhl img{height:62px}
.dht{text-align:center;flex:1;padding:0 14px}
.dmt{font-size:1.6rem;font-weight:700;color:var(--c9)}
.dsb{font-size:.7rem;color:var(--mu);margin-top:4px;text-transform:uppercase;letter-spacing:.05em}
.clo img{max-height:56px;max-width:120px;object-fit:contain}
.clph{width:110px;height:46px;border:1.5px solid var(--bd);border-radius:var(--r);display:flex;align-items:center;justify-content:center;font-size:.65rem;color:var(--mu);text-align:center}
.warn{background:#FFF8E0;border:1.5px solid #C4A020;border-radius:var(--r);padding:10px 14px;margin-bottom:14px}
.warn h4{color:#7A6000;font-size:.76rem;font-weight:700;margin-bottom:5px}
.warn li{font-size:.7rem;color:#7A6000;padding:2px 0 2px 16px;position:relative;list-style:none}
.warn li::before{content:'â ';position:absolute;left:0}
.sbn{background:var(--c8);color:#fff;padding:9px 16px;border-radius:var(--r);margin-bottom:14px;display:flex;align-items:center;gap:8px}
.sbn h2{font-family:'Libre Baskerville',serif;font-size:1.05rem;font-weight:700;letter-spacing:.03em}
.kr{display:grid;grid-template-columns:repeat(auto-fit,minmax(128px,1fr));gap:12px;margin-bottom:12px}
.kp{background:var(--wh);border-radius:var(--r);padding:14px 14px 10px;box-shadow:var(--sh);border-top:3px solid var(--c6)}
.kl{font-size:.62rem;color:var(--mu);text-transform:uppercase;letter-spacing:.06em;font-weight:600}
.kv{font-family:'Libre Baskerville',serif;font-size:2rem;color:var(--c8);font-weight:700;line-height:1;margin-top:2px}
.kd{font-size:.64rem;color:var(--mu);margin-top:2px}
.cg{display:grid;gap:12px;margin-bottom:12px}
.c3{grid-template-columns:1fr 1fr 1fr}
.c2{grid-template-columns:1fr 1fr}
.c1{grid-template-columns:1fr}
.cc{background:var(--wh);border-radius:var(--r);padding:16px;box-shadow:var(--sh)}
.ct{font-size:.67rem;color:var(--mu);text-transform:uppercase;letter-spacing:.06em;font-weight:600;margin-bottom:12px;padding-bottom:7px;border-bottom:1px solid var(--c0)}
.ch{position:relative}
.sd{height:2px;background:var(--c1);margin:20px 0}
/* REPORT */
.rw{max-width:1100px;margin:0 auto;padding:28px 36px;background:#fff;min-height:100vh}
.rh{display:flex;align-items:center;justify-content:space-between;padding-bottom:14px;border-bottom:3px solid var(--c7);margin-bottom:20px}
.rh img{height:64px}
.rtb h1{font-family:'Libre Baskerville',serif;font-size:1.3rem;color:var(--c9);text-align:right}
.rtb p{font-size:.72rem;color:var(--mu);text-align:right;margin-top:3px}
.rnote{background:#FFF8E0;border-left:3px solid var(--gd);padding:8px 14px;border-radius:4px;margin-bottom:16px;font-size:.71rem;color:#6A5000}
.rbtns{display:flex;gap:10px;margin-bottom:18px}
.rb{padding:9px 18px;border-radius:var(--r);font-size:.78rem;font-weight:700;cursor:pointer;border:none;transition:.2s}
.rxl{background:var(--gd);color:var(--c9)}.rxl:hover{background:var(--gd2)}
.rpf{background:var(--c7);color:#fff}.rpf:hover{background:var(--c9)}
.rks{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:20px}
.rk{background:var(--c0);border-radius:var(--r);padding:12px;border-top:3px solid var(--c6);text-align:center}
.rk .l{font-size:.6rem;color:var(--mu);text-transform:uppercase;letter-spacing:.04em;font-weight:600}
.rk .v{font-family:'Libre Baskerville',serif;font-size:1.65rem;color:var(--c8);font-weight:700}
.rs{font-family:'Libre Baskerville',serif;font-size:.92rem;color:var(--c8);margin:18px 0 8px;padding-bottom:5px;border-bottom:2px solid var(--c2)}
.pt{width:100%;border-collapse:collapse;font-size:.71rem}
.pt th{background:var(--c8);color:#fff;padding:7px 8px;text-align:left;font-weight:600;font-size:.65rem;text-transform:uppercase;letter-spacing:.04em;white-space:nowrap}
.pt td{padding:6px 8px;border-bottom:1px solid var(--bd);vertical-align:top}
.pt tr:nth-child(even) td{background:var(--c0)}
@media print{
  #cfg,.tbar,.rbtns{display:none!important}
  body[data-print=dash] #rpt{display:none!important}
  body[data-print=dash] #dsh{display:block!important}
  body[data-print=report] #dsh{display:none!important}
  body[data-print=report] #rpt{display:block!important}
  *{-webkit-print-color-adjust:exact;print-color-adjust:exact}
  .cc,.kp,.rk{break-inside:avoid}
  .rw{padding:14px}
  .dw{padding:10px;max-width:none}
}
@media(max-width:860px){
  .grid2,.up3,.c3,.c2{grid-template-columns:1fr}
  .kr{grid-template-columns:repeat(2,1fr)}
}
"""

# JavaScript - written carefully without embedded HTML in template literals
JS = r"""
const ST={proc:null,serv:null,aud:null,dec:null,logo:null};
const CH={};


const TAX = {
// AUDIÃNCIA
'AudiÃªncia':'AudiÃªncia','AudiÃªncia - CÃ­vel':'AudiÃªncia','AudiÃªncia / AudiÃªncia CÃ­vel':'AudiÃªncia',
'AudiÃªncia / ConciliaÃ§Ã£o':'AudiÃªncia','AudiÃªncia / Inaugural':'AudiÃªncia',
'AudiÃªncia / InstruÃ§Ã£o':'AudiÃªncia','AudiÃªncia / InstruÃ§Ã£o e Julgamento':'AudiÃªncia',
'AudiÃªncia / Julgamento':'AudiÃªncia','AudiÃªncia / JustificaÃ§Ã£o PrÃ©via':'AudiÃªncia',
'AudiÃªncia / Una':'AudiÃªncia','AudiÃªncia Una':'AudiÃªncia',
'ServiÃ§o / AudiÃªncia':'AudiÃªncia','Pauta de Julgamento':'AudiÃªncia',
'Workflow / Registro de audiÃªncia':'AudiÃªncia','Workflow / Sinalizar audiÃªncia':'AudiÃªncia',
'ServiÃ§o / ConciliaÃ§Ã£o Extrajudicial':'AudiÃªncia',
// REUNIÃO
'ReuniÃ£o':'ReuniÃ£o','ReuniÃ£o / Cliente':'ReuniÃ£o','ReuniÃ£o / Coordenadores':'ReuniÃ£o',
'ReuniÃ£o / Prospect':'ReuniÃ£o','ReuniÃ£o / SÃ³cios':'ReuniÃ£o','ReuniÃ£o / Treinamento':'ReuniÃ£o',
'ReuniÃ£o de CaptaÃ§Ã£o':'ReuniÃ£o','ReuniÃ£o de Consultoria':'ReuniÃ£o','ReuniÃ£o Interna':'ReuniÃ£o',
'ReuniÃ£o / ReuniÃ£o':'ReuniÃ£o','ServiÃ§o / ReuniÃ£o':'ReuniÃ£o',
'ServiÃ§o / Acompanhamento em ReuniÃ£o':'Acompanhamento em ReuniÃ£o',
// CONSULTA
'ServiÃ§o / Consulta':'Consulta','ServiÃ§o / OrientaÃ§Ã£o':'Consulta',
// DESPACHO
'Despacho com Magistrado':'Despacho com Magistrado',
// DILIGÃNCIA
'DiligÃªncia externa':'DiligÃªncia','ServiÃ§o / DiligÃªncia':'DiligÃªncia',
'Diversos / DiligÃªncia':'DiligÃªncia',
// TREINAMENTO
'TREINAMENTO':'Treinamento','ElaboraÃ§Ã£o de Plano de Treinamento':'ElaboraÃ§Ã£o de Plano de Treinamento',
// PALESTRA
'ServiÃ§o / Palestra':'Palestra',
// PARECER
'ServiÃ§o / Parecer':'Parecer','ServiÃ§o / Parecer Complexo':'Parecer',
// ACOMPANHAMENTO
'ServiÃ§o / Acompanhamento em Depoimento/Oitiva':'Acompanhamento em Depoimento/Oitiva',
// DEMANDA SEFAZ
'ServiÃ§o / Demanda Sefaz':'Demanda Sefaz',
// SERVIÃOS ESPECÃFICOS
'ServiÃ§o / ServiÃ§os Redesim':'ServiÃ§os Redesim','ServiÃ§o / ServiÃ§os Siscoex':'ServiÃ§os Siscoex',
'ServiÃ§os Redesim':'ServiÃ§os Redesim','ServiÃ§os Siscoex':'ServiÃ§os Siscoex',
// ANÃLISE CONTRATUAL/SOCIETÃRIA
'ServiÃ§o / AnÃ¡lise de Aditivo':'AnÃ¡lise de Aditivo',
'ServiÃ§o / AnÃ¡lise de Contrato':'AnÃ¡lise de Contrato',
'ServiÃ§o / RevisÃ£o Contratual':'AnÃ¡lise de Contrato',
'ServiÃ§o / AnÃ¡lise de Distrato':'AnÃ¡lise de Distrato',
'ServiÃ§o / AnÃ¡lise de Inst. SocietÃ¡rio':'AnÃ¡lise de Inst. SocietÃ¡rio',
// ALTERAÃÃO
'ServiÃ§o / AlteraÃ§Ã£o AGE':'AlteraÃ§Ã£o AGE','ServiÃ§o / AlteraÃ§Ã£o Contratual':'AlteraÃ§Ã£o Contratual',
'ServiÃ§o / AlteraÃ§Ã£o Estatuto':'AlteraÃ§Ã£o Estatuto',
'ServiÃ§o / AlteraÃ§Ã£o Inst. SocietÃ¡rio':'AlteraÃ§Ã£o Inst. SocietÃ¡rio',
// ELABORAÃÃO CONTRATUAL/SOCIETÃRIA
'ServiÃ§o / ElaboraÃ§Ã£o de Aditivo':'ElaboraÃ§Ã£o de Aditivo',
'ServiÃ§o / Protocolo de Aditivo':'ElaboraÃ§Ã£o de Aditivo',
'Aditivo contratual de clÃ¡usulas de ProteÃ§Ã£o de Dados':'ElaboraÃ§Ã£o de Aditivo',
'ServiÃ§o / ElaboraÃ§Ã£o de clÃ¡usula':'ElaboraÃ§Ã£o de ClÃ¡usula',
'ElaboraÃ§Ã£o de clÃ¡usula contratual':'ElaboraÃ§Ã£o de Contrato',
'ServiÃ§o / ElaboraÃ§Ã£o de Contrato':'ElaboraÃ§Ã£o de Contrato',
'ServiÃ§o / Contratual Comum':'ElaboraÃ§Ã£o de Contrato',
'ServiÃ§o / Contratual Complexo':'ElaboraÃ§Ã£o de Contrato',
'ServiÃ§o / Contratual ExtraordinÃ¡rio':'ElaboraÃ§Ã£o de Contrato',
'ServiÃ§o':'ElaboraÃ§Ã£o de Contrato',
'ServiÃ§o / ElaboraÃ§Ã£o de Distrato':'ElaboraÃ§Ã£o de Distrato',
'ServiÃ§o / ElaboraÃ§Ã£o de doc societÃ¡rio':'ElaboraÃ§Ã£o de Inst. SocietÃ¡rio',
'ServiÃ§o / ElaboraÃ§Ã£o de Inst. SocietÃ¡rio':'ElaboraÃ§Ã£o de Inst. SocietÃ¡rio',
'ServiÃ§o / ElaboraÃ§Ã£o de Minuta':'ElaboraÃ§Ã£o de Minuta',
'ServiÃ§o / ElaboraÃ§Ã£o de NotificaÃ§Ã£o.':'ElaboraÃ§Ã£o de NotificaÃ§Ã£o',
'ServiÃ§o / ElaboraÃ§Ã£o de OfÃ­cio':'ElaboraÃ§Ã£o de OfÃ­cio',
'ServiÃ§o / ElaboraÃ§Ã£o de Requerimento':'ElaboraÃ§Ã£o de Requerimento',
'ServiÃ§o / ElaboraÃ§Ã£o de Termo de Compromisso':'ElaboraÃ§Ã£o de Termo de Compromisso',
'ELABORAÃÃO DE TERMO DE COMPROMISSO':'ElaboraÃ§Ã£o de Termo de Compromisso',
'ElaboraÃ§Ã£o de Termo':'ElaboraÃ§Ã£o de Termo',
// LGPD
'AnÃ¡lise de Inventario de processos com dados pessoais':'AnÃ¡lise de InventÃ¡rio de Processos com Dados Pessoais',
'AnÃ¡lise de PolÃ­tica de Privacidade':'AnÃ¡lise de PolÃ­tica de Privacidade',
'AnÃ¡lise de PolÃ­tica de SeguranÃ§a da InformaÃ§Ã£o':'AnÃ¡lise de PolÃ­tica de SeguranÃ§a da InformaÃ§Ã£o',
'ANÃLISE DE TERMO DE COMPROMISSO':'AnÃ¡lise de Termo de Compromisso',
'Atendimento Titular de Dados -DPO':'Atendimento Titular de Dados DPO',
'AtualizaÃ§Ã£o Be Compliance':'AtualizaÃ§Ã£o Be Compliance',
'ElaboraÃ§Ã£o de Aviso de cookies':'ElaboraÃ§Ã£o de Aviso de Cookies',
'ElaboraÃ§Ã£o de Aviso de Privacidade':'ElaboraÃ§Ã£o de Aviso de Privacidade',
'ElaboraÃ§Ã£o de Plano de NotificaÃ§Ãµes':'ElaboraÃ§Ã£o de Plano de NotificaÃ§Ãµes',
'ElaboraÃ§Ã£o de PolÃ­tica':'ElaboraÃ§Ã£o de PolÃ­tica',
'ElaboraÃ§Ã£o de PolÃ­tica de Privacidade':'ElaboraÃ§Ã£o de PolÃ­tica de Privacidade',
'ElaboraÃ§Ã£o de PolÃ­tica de SeguranÃ§a da InformaÃ§Ã£o':'ElaboraÃ§Ã£o de PolÃ­tica de SeguranÃ§a da InformaÃ§Ã£o',
'ElaboraÃ§Ã£o de relatÃ³rio da anÃ¡lise dos riscos':'ElaboraÃ§Ã£o de RelatÃ³rio da AnÃ¡lise dos Riscos',
'ElaboraÃ§Ã£o de ROPA':'ElaboraÃ§Ã£o de ROPA',
'ElaboraÃ§Ã£o de termos de uso':'ElaboraÃ§Ã£o de Termos de Uso',
// ELABORAÃÃO DE PEÃA PROCESSUAL
'Prazo':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / AÃ§Ã£o RescisÃ³ria-RazÃµes Finais (art. 973 CPC)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Acompanhar Julgamento.':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Aditivo do contrato':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Agravo contra decisÃ£o que inadmite na origem Resp ou Rext - ContrarrazÃµes (art. 1.042, Â§ 3Âº, CPC)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Agravo contra decisÃ£o que inadmite na origem Resp ou Rext (art. 1.042 c/c art. 1.003, Â§ 5Âº CPC)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Agravo de Instrumento':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Agravo de Instrumento Civel pz em dobro (Art. 525 c/c 191 CPC)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Agravo de Instrumento Trabalhista':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Agravo de Instrumento Trabalhista (art. 897, b CLT)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Agravo de PetiÃ§Ã£o (art. 897, Â§ 1Âº, CLT)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Agravo de PetiÃ§Ã£o (art. 897, a, CLT)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Agravo Interno - ContrarrazÃµes (art. 1.021, Â§ 2Âº, CPC)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Agravo Interno de decisÃ£o TRT':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Agravo Interno/Regimental':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Agravo Retido':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / AIRO':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / AIRR':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ApelaÃ§Ã£o - ContrarrazÃµes (art. 1.010, Â§ 1Âº, CPC)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ApelaÃ§Ã£o Civel':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ApresentaÃ§Ã£o De-Documentos':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Ato processual sem prazo fixado (art. 218, Â§ 3Âº CPC)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ContestaÃ§Ã£o':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'ContestaÃ§Ã£o (art. 190 c/c 240, III CPC - Prazo em dobro)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ContestaÃ§Ã£o (art. 335 CPC)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ContestaÃ§Ã£o Trabalhista':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Contra RazÃµes em Recurso Trabalhista (4d)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ContrarrazÃµes':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ContrarrazÃµes Embargos Ã  ExecuÃ§Ã£o':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ContrarrazÃµes RO/RR':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Cumprimento de SentenÃ§a':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Defesa Administrativa':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Defesa Administrativa - SEMACE':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Defesa Penal':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Defesa Penal - Arts. 396 e 396 A do CPP':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Elaborar e protocolar Agravo para CARF':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Elaborar e Protocolar ED para CARF':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Embargos Ã  ExecuÃ§Ã£o':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Embargos Ã  ExecuÃ§Ã£o Fiscal':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Embargos Ã  ExecuÃ§Ã£o Trabalhista':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Embargos a ExecuÃ§Ã£o Trabalhista (art. 884 CLT)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Embargos de DeclaraÃ§Ã£o':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Embargos de DeclaraÃ§Ã£o (art. 1.023/1.024 CPC)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Embargos de DivergÃªncia (art. 1.003 CPC)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Embargos de Terceiro':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Embargos MonitÃ³rios':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Emenda da petiÃ§Ã£o inicial na tutela antecipada em carÃ¡ter antecedente denegada (art. 303, Â§ 6Âº, CPC)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Emendar Inicial':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Especificar provas e temas sobre provas':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ExceÃ§Ã£o de PrÃ©-Executividade':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ExceÃ§Ã£o PrÃ© Executividade Trabalhista':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / HabilitaÃ§Ã£o - ManifestaÃ§Ã£o dos Requeridos (art. 690, CPC)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ImpugnaÃ§Ã£o aos Embargos MonitÃ³rios':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ImpugnaÃ§Ã£o Auto de InfraÃ§Ã£o SRFB':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ImpugnaÃ§Ã£o de CÃ¡lculos Trabalhista 8d':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / IndicaÃ§Ã£o de bens Ã-Penhora':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / IndicaÃ§Ã£o de testemunhas (CPC art. 357) ou assistente tÃ©cnico e apresentaÃ§Ã£o de quesitos Ã  perÃ­cia (art. 465, CPC)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Juntar Documentos':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Liminar':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Mandado de SeguranÃ§a':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ManfestaÃ§Ã£o Diversa (20d)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Manif Trab 48h':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ManifestaÃ§Ã£o diversa':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ManifestaÃ§Ã£o Diversa (30d)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ManifestaÃ§Ã£o Diversa 10d':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ManifestaÃ§Ã£o Diversa 15d':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ManifestaÃ§Ã£o Diversa 5d':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ManifestaÃ§Ã£o Diversa 8d':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ManifestaÃ§Ã£o do sÃ³cio ou da pessoa jurÃ­dica no pedido de desconsideraÃ§Ã£o da personalidade jurÃ­dica (art. 135, CPC)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ManifestaÃ§Ã£o PrÃ©via - Improbidade Administrativa':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ManifestaÃ§Ã£o sobre documentos (art. 437, Â§ 1Âº CPC)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / ManifestaÃ§Ã£o Sobre Laudo Pericial CONAT - CE':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / NomeaÃ§Ã£o de bens ou depÃ³sito em ExecuÃ§Ã£o Trabalhista':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Pagamento de acordo':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Pagamento de custas sob pena de cancelamento da distribuiÃ§Ã£o (art. 290, CPC)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Pagamento em execuÃ§Ã£o trabalhista -15d':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Pagamento em execuÃ§Ã£o trabalhista -48h':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Pedido de esclarecimentos ou correÃ§Ãµes da decisÃ£o de saneamento e organizaÃ§Ã£o do processo (art. 357, Â§ 1Âº, CPC)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Pedido Inicial- Emendar ou Completar (art. 321 e 801 CPC)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo para apresentaÃ§Ã£o de rol de testemunhas (prazo mÃ¡ximo - art. 357, Â§ 4Âº, CPC)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Protocolar apelaÃ§Ã£o':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Protocolar prova depÃ³sito 30% em ExecuÃ§Ã£o':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / PROTOCOLAR prova depÃ³sito parcelamento em ExecuÃ§Ã£o 1':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / PROTOCOLAR prova depÃ³sito parcelamento em ExecuÃ§Ã£o 2':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / PROTOCOLAR prova depÃ³sito parcelamento em ExecuÃ§Ã£o 3':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / PROTOCOLAR prova depÃ³sito parcelamento em ExecuÃ§Ã£o 4':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / PROTOCOLAR prova depÃ³sito parcelamento em ExecuÃ§Ã£o 5':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / PROTOCOLAR prova depÃ³sito parcelamento em ExecuÃ§Ã£o 6':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Protocolar Recurso decisao TJ TRF':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Protocolar Recurso HierÃ¡rquico DRF':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Protocolar Recurso sobre decisÃ£o de Relator e Colegiados TJ TRF':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Protocolar Recurso sobre DecisÃ£o STJ STF':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Protocolar recurso sobre DecisÃ£o TJ TRF de inadmissibilidade de REsp / RE':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Protocolar recurso sobre decisÃ£o TST (4d)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Protocolar recurso sobre sentenÃ§a':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Protocolar Recurso voluntÃ¡rio para CARF':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Protocolar Recurso Voluntario Sefaz':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Protocolar REsp para CARF':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Protocolo Penal 2d corridos D-0':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / RazÃµes Finais Trabalhista (10 dias)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / RazÃµes Finais Trabalhista (5 dias)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Recurso Adesivo':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Recurso Administrativo':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Recurso de Revista':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Recurso Especial':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Recurso Especial - CARF':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Recurso Especial ou ExtraordinÃ¡rio - ContrarrazÃµes (art. 1.030, caput, CPC)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Recurso ExtraordinÃ¡rio':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Recurso ExtraordinÃ¡rio PAF SEFAZ':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Recurso Inominado':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Recurso OrdinÃ¡rio':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Recurso OrdinÃ¡rio (JECF)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Recurso OrdinÃ¡rio em Habeas Corpus':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Recurso OrdinÃ¡rio-Trabalhista (art. 895 CLT)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Recurso VoluntÃ¡rio':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Recurso VoluntÃ¡rio - SEFIN - FOR':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / RÃ©plica':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / RÃ©plica a contestaÃ§Ã£o - (art 350 e 351 CPC)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Requerer o que entender De-Direito':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Prazo / Vista aos Autos (art. 107, II CPC)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'RazÃµes finais escritas - Prazos sucessivos autor rÃ©u MinistÃ©rio PÃºblico (art. 364 Â§ 2Âº CPC)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Protocolo':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Protocolo / Protocolar REsp Sefaz':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'ServiÃ§os/COOP':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'ServiÃ§o / NotificaÃ§Ã£o Extrajudicial':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'ServiÃ§o / Requerimento':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'ServiÃ§o / Requerimento Administrativo':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'ServiÃ§o / Requerimento Complexo':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'ElaboraÃ§Ã£o de ManifestaÃ§Ã£o':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Diversos / Elaborar apelaÃ§Ã£o':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Diversos / ManifestaÃ§Ã£o Diversa (1d)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Diversos / ManifestaÃ§Ã£o diversa (2d)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Diversos / ManifestaÃ§Ã£o diversa (5d)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar Agravo':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar agravo de petiÃ§Ã£o':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar Agravo Interno - negativa de segmento a RE e REsp (duas peÃ§as)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar Agravo Interno - negativa de segmento a RE ou REsp':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar Agravo Interno TST':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar apelaÃ§Ã£o':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar ApelaÃ§Ã£o Penal':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar ARE e AREsp (duas peÃ§as)':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar ARE ou AREsp':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar ED':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar ED Penal':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar Embargos Ã  ExecuÃ§Ã£o':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar Embargos de DivergÃªncia':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar Emenda a inicial':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar Mandado de SeguranÃ§a':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar ManifestaÃ§Ã£o Diversa 10d':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar ManifestaÃ§Ã£o Diversa 15d':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar ManifestaÃ§Ã£o Diversa 20d':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar ManifestaÃ§Ã£o Diversa 30d':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar ManifestaÃ§Ã£o Diversa 5d':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar ManifestaÃ§Ã£o Diversa 8 d':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar Pet Custas complementares':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar RazÃµes Finais':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar RE':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar RE e REsp':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar Recurso de SentenÃ§a Penal em Juizado':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar Recurso HierÃ¡rquico Lei 9.874/99':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar Recurso para Juizado':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar Recurso Sefin':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar Recurso voluntÃ¡rio para CARF':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar Recurso VoluntÃ¡rio Sefaz':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar Replica Ã  contestaÃ§Ã£o':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar REsp':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar REsp para CARF':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar RO':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Elaborar: Especificar provas e temas sobre provas':'ElaboraÃ§Ã£o de PeÃ§a Processual',
// CONTRATOS ESPECIAIS
'Diversos / Reajuste contratual':'Reajuste contratual',
'Diversos / RenovaÃ§Ã£o contratual':'RenovaÃ§Ã£o contratual',
// ATIVIDADE INTERNA
'Alvara':'Atividade Interna','ANÃLISE DE DOCUMENTAÃÃO':'Atividade Interna',
'ANÃLISE DE PROPOSTA':'Atividade Interna','Andamento':'Atividade Interna',
'Atendimento/LigaÃ§Ã£o':'Atividade Interna','Canal de DenÃºncia':'Atividade Interna',
'CriaÃ§Ã£o de conteÃºdo':'Atividade Interna','Diversos':'Atividade Interna',
'Diversos / Contato TelefÃ´nico':'Atividade Interna',
'Diversos / Informar sentenÃ§a ao cliente':'Atividade Interna',
'Diversos / Resposta de E-mail':'Atividade Interna',
'Edital':'Atividade Interna','ElaboraÃ§Ã£o de certificados':'Atividade Interna',
'ElaboraÃ§Ã£o de material treinamento':'Atividade Interna',
'Estudo de Caso':'Atividade Interna','IntimaÃ§Ã£o':'Atividade Interna',
'IntimaÃ§Ã£o EletrÃ´nica':'Atividade Interna','IntimaÃ§Ã£o EletrÃ´nica / VerificaÃ§Ã£o':'Atividade Interna',
'IntimaÃ§Ãµes Iprazos':'Atividade Interna','Pesquisas/Estudos':'Atividade Interna',
'Pesquisas/Estudos / ElaboraÃ§Ã£o de Nova Tese':'Atividade Interna',
'Pesquisas/Estudos / ElaboraÃ§Ã£o de RelatÃ³rio':'Atividade Interna',
'PublicaÃ§Ã£o':'Atividade Interna','PublicaÃ§Ã£o / IntimaÃ§Ã£o':'Atividade Interna',
'PublicaÃ§Ã£o / VerificaÃ§Ã£o':'Atividade Interna','PublicaÃ§Ãµes Iprazos':'Atividade Interna',
'RelatÃ³rio de Processos':'Atividade Interna','Resposta de E-mail':'Atividade Interna',
'RevisÃ£o':'Atividade Interna','RevisÃ£o / Agravo contra decisÃ£o que inadmite na origem Resp ou Rext (art. 1.042 c/c art. 1.003, Â§ 5Âº CPC)':'Atividade Interna',
'Saneamento de relatÃ³rio':'Atividade Interna','Tratativa de Acordo':'Atividade Interna',
'Venda':'Atividade Interna','Verificar protocolo':'Atividade Interna',
'Verificar protocolo / ApelaÃ§Ã£o':'Atividade Interna',
'VISTO JURÃDICO':'Atividade Interna',
'Workflow / Informar acÃ³rdÃ£o ao cliente':'Atividade Interna',
'Workflow / Enviar ao cliente petiÃ§Ã£o inicial RT e documentos':'Atividade Interna',
'Workflow / Enviar sentenÃ§a ao cliente (2d)':'Atividade Interna',
'Workflow / Enviar sentenÃ§a ao cliente (3d)':'Atividade Interna',
'Workflow / Informar sentenÃ§a ao cliente (4d)':'Atividade Interna',
'Workflow / Analisar se hÃ¡ honorÃ¡rios contratuais':'Atividade Interna',
'Workflow / AnÃ¡lise de execuÃ§Ã£o de honorÃ¡rios (20d)':'Atividade Interna',
'Workflow / Atos de Impulso (15d)':'Atividade Interna','Workflow / Atos de Impulso (5d)':'Atividade Interna',
'Workflow / ConferÃªncia (4d)':'Atividade Interna','Workflow / Conferencia (7d)':'Atividade Interna',
'Workflow / Decidir providÃªncias PAT':'Atividade Interna',
'Workflow / Habilitar Dra. Imaculada no Processo':'Atividade Interna',
'Workflow / InterlocutÃ³ria CPC':'Atividade Interna','Workflow / InterlocutÃ³ria trabalhista':'Atividade Interna',
'Workflow / InterlocutÃ³ria Trabalhista em AudiÃªncia':'Atividade Interna',
'Workflow / Memoriais (25d)':'Atividade Interna',
'Workflow / Pagamento de parcelamento execuÃ§Ã£o CPC Art. 916':'Atividade Interna',
'Workflow / Parcelamento CPC Art. 916 em execuÃ§Ã£o parcela 1':'Atividade Interna',
'Workflow / Parcelamento CPC Art. 916 em execuÃ§Ã£o parcela 2':'Atividade Interna',
'Workflow / Parcelamento CPC Art. 916 em execuÃ§Ã£o parcela 3':'Atividade Interna',
'Workflow / Parcelamento CPC Art. 916 em execuÃ§Ã£o parcela 4':'Atividade Interna',
'Workflow / Parcelamento CPC Art. 916 em execuÃ§Ã£o parcela 5':'Atividade Interna',
'Workflow / Parcelamento CPC Art. 916 em execuÃ§Ã£o parcela 6':'Atividade Interna',
'Workflow / Protocolo (14d)':'Atividade Interna','Workflow / Protocolo (2d)':'Atividade Interna',
'Workflow / Protocolo (4d)':'Atividade Interna','Workflow / Protocolo (7d)':'Atividade Interna',
'Workflow / Protocolo (9d)':'Atividade Interna',
'Workflow / Providenciar REsp Sefaz':'Atividade Interna',
'Workflow / ProvidÃªncias (1d)':'Atividade Interna','Workflow / ProvidÃªncias (2d)':'Atividade Interna',
'Workflow / Providencias em ExecuÃ§Ã£o Trabalhista':'Atividade Interna',
'Workflow / Saneamento de dados (10d)':'Atividade Interna','Workflow / Saneamento de dados (1d)':'Atividade Interna',
'Workflow / Saneamento de dados (20d)':'Atividade Interna','Workflow / Saneamento de dados (35d)':'Atividade Interna',
'Workflow / Sinalizar':'Atividade Interna','Workflow / Sinalizar PerÃ­cia':'Atividade Interna',
'Workflow / AcordÃ£o TJ TRF':'Atividade Interna','Workflow / AcÃ³rdÃ£o TRT':'Atividade Interna',
'Workflow / ContrarrazÃµes Trabalhista':'Atividade Interna',
'Workflow / Decidir providÃªncias PAT':'Atividade Interna',
'Workflow / DecisÃ£o complexa':'Atividade Interna','Workflow / DecisÃ£o Presidencia TRT':'Atividade Interna',
'Workflow / DecisÃ£o STJ/STF':'Atividade Interna','Workflow / DecisÃ£o TJ TRF sobre REsp e RE':'Atividade Interna',
'Workflow / DecisÃ£o TST':'Atividade Interna',
'Workflow / Elaborar Pet Custas complementares':'ElaboraÃ§Ã£o de PeÃ§a Processual',
'Workflow / Gerar taxa recurso Sefaz':'Atividade Interna',
'Workflow / Providenciar REsp Sefaz':'Atividade Interna',
'Workflow / Registro de audiÃªncia':'AudiÃªncia',
'Memoriais':'Atividade Interna',
'ElaboraÃ§Ã£o de PolÃ­tica de SeguranÃ§a da InformaÃ§Ã£o':'ElaboraÃ§Ã£o de PolÃ­tica de SeguranÃ§a da InformaÃ§Ã£o',
'ServiÃ§o / AnÃ¡lise de ProcuraÃ§Ã£o':'Atividade Interna',
'ServiÃ§o / ElaboraÃ§Ã£o de ProcuraÃ§Ã£o':'Atividade Interna',
'ServiÃ§o / ElaboraÃ§Ã£o de Proposta':'Atividade Interna',
'ServiÃ§o / ElaboraÃ§Ã£o de Recibo':'Atividade Interna',
'Pagamento SolidÃ¡rio/SubsidiÃ¡rio':'Atividade Interna',
'HonorÃ¡rios - Verificar / Monitorar':'Atividade Interna',
'CobranÃ§a de honorarios':'Atividade Interna',
'Diversos / Reajuste contratual':'Reajuste contratual',
'Diversos / RenovaÃ§Ã£o contratual':'RenovaÃ§Ã£o contratual',
'Diversos / Ida Ã  Delegacia':'Atividade Interna','Diversos / Ida Ã  Secretaria':'Atividade Interna',
'Diversos / CJ Atualizar Sistema de Cliente':'Atividade Interna',
'Diversos / CJ SolicitaÃ§Ãµes equipe tÃ©cnica':'Atividade Interna',
'Diversos / CJ Verificar TrÃ¢nsito em Julgado/ Cadastrar eSocial':'Atividade Interna',
'Diversos / ET SolicitaÃ§Ãµes CJ':'Atividade Interna',
'Diversos / Encaminhar inicial e solicitar documentos / habilitar no processo':'Atividade Interna',
'Diversos / Verificar Processo Especial':'Atividade Interna',
'Diversos / Verificar se contrato Ã© por ato e fazer a cobranÃ§a':'Atividade Interna',
'ReuniÃ£o / Prospect':'ReuniÃ£o',
  'AudiÃªncia / AudiÃªncia - CÃ­vel':'AudiÃªncia',
  'ServiÃ§o /':'ElaboraÃ§Ã£o de Contrato',
  'AudiÃªncia /':'AudiÃªncia',
  'ReuniÃ£o /':'ReuniÃ£o',
  'ReuniÃ£o / ReuniÃ£o':'ReuniÃ£o',
};

// Natureza simplification map (from services spreadsheet full name â short name)
const NAT_MAP = {
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea administrativa':'SocietÃ¡ria',
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea operacional / Ambiental':'Ambiental',
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea operacional / CÃ­vel':'CÃ­vel',
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea operacional / CÃ­vel / Consultas':'Consultas',
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea operacional / CÃ­vel / Processos':'CÃ­vel',
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea operacional / Contratual':'Contratual',
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea operacional / Empresarial / SocietÃ¡ria':'SocietÃ¡ria',
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea operacional / Hospitalar':'Hospitalar',
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea operacional / Hospitalar / Pareceres':'Hospitalar',
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea operacional / LGPD':'LGPD',
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea operacional / OperaÃ§Ãµes e Negocios':'OperaÃ§Ãµes e NegÃ³cios',
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea operacional / Penal':'Penal',
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea operacional / Penal / Processos':'Penal',
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea operacional / RegulatÃ³rio':'RegulatÃ³rio',
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea operacional / Trabalhista':'Trabalhista',
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea operacional / Trabalhista / Consultas':'Consultas',
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea operacional / Trabalhista / Pareceres':'Trabalhista',
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea operacional / Trabalhista / Processos':'Trabalhista',
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea operacional / TributÃ¡rio':'TributÃ¡rio',
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea operacional / TributÃ¡rio / Consultas':'Consultas',
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea operacional / TributÃ¡rio / Processos':'TributÃ¡rio',
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea operacional / Direito Administrativo':'Direito Administrativo',
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea operacional / Direito Administrativo / Processos':'Direito Administrativo',
'IMACULADA GORDIANO SOCIEDADE DE ADVOGADOS / Ãrea operacional / Consumidor':'Consumidor',
};

// Robust year extractor â handles dd/MM/yyyy, yyyy-MM-dd, Date objects, any format
function parseYear(d){
  if(!d) return NaN;
  if(d instanceof Date) return isNaN(d.getTime()) ? NaN : d.getFullYear();
  var s = String(d).trim();
  if(!s || s==='null' || s==='undefined') return NaN;
  var m = s.match(/\b(19\d{2}|20\d{2})\b/);
  return m ? parseInt(m[1], 10) : NaN;
}

function uT(t){
  if(!t || String(t).trim()==='' || t==='nan') return 'Consulta';
  var raw = String(t).trim();
  if(TAX[raw]) return TAX[raw];
  // Normalize new export format: 'Tipo - Subtipo' or 'Tipo - ' (trailing dash)
  var s = raw.replace(/\s*-\s*$/, '').trim();   // strip trailing ' - ' or ' -'
  if(TAX[s]) return TAX[s];
  s = s.replace(/ - /, ' / ');                    // replace first ' - ' with ' / '
  if(TAX[s]) return TAX[s];
  return raw;
}
function uNat(n){
  if(!n || String(n).trim()==='') return 'NÃ£o informado';
  return NAT_MAP[String(n).trim()] || String(n).trim();
}

function extractNat(assunto){
  if(!assunto || String(assunto).trim()==='') return 'NÃ£o classificado';
  var s = String(assunto).trim().toUpperCase();
  // Must start with CONSULTIVO
  if(s.indexOf('CONSULTIVO') !== 0) return 'NÃ£o classificado';
  // Remove leading CONSULTIVO and optional spaces
  var rest = s.replace(/^CONSULTIVO\s*/, '');
  // Extract first word (the nature keyword)
  var word = rest.split(/[\s:,/\-]/)[0] || '';
  if(!word) return 'NÃ£o classificado';
  // Normalization map
  var norm = {
    'CONTRATOS':'Contratos','CONTRATUAL':'Contratos',
    'TRIBUTARIO':'TributÃ¡rio','TRIBUTARIA':'TributÃ¡rio','TRIBUTÃRIA':'TributÃ¡rio','TRIBUTÃRIO':'TributÃ¡rio',
    'CIVEL':'CÃ­vel','CÃVEL':'CÃ­vel',
    'REGULATORIO':'RegulatÃ³rio','REGULATÃRIO':'RegulatÃ³rio',
    'TRABALHISTA':'Trabalhista',
    'HOSPITALAR':'Hospitalar',
    'AMBIENTAL':'Ambiental',
    'DIGITAL':'Digital',
    'EMPRESARIAL':'Empresarial',
    'CIVIL':'CÃ­vel'
  };
  if(norm[word]) return norm[word];
  // Capitalize first letter, rest lowercase
  return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
}


function uT(t){
  if(!t || String(t).trim()==='' || t==='nan') return 'Consulta';
  var raw = String(t).trim();
  if(TAX[raw]) return TAX[raw];
  // Normalize new export format: 'Tipo - Subtipo' or 'Tipo - ' (trailing dash)
  var s = raw.replace(/\s*-\s*$/, '').trim();   // strip trailing ' - ' or ' -'
  if(TAX[s]) return TAX[s];
  s = s.replace(/ - /, ' / ');                    // replace first ' - ' with ' / '
  if(TAX[s]) return TAX[s];
  return raw;
}

function showErr(m){
  var el = document.getElementById('errbox');
  el.textContent = m;
  el.style.display = 'block';
}
function clrErr(){
  var el = document.getElementById('errbox');
  el.textContent = '';
  el.style.display = 'none';
}
function showWarns(list){
  var el = document.getElementById('warnblock');
  if(!list || list.length===0){ el.innerHTML=''; return; }
  var html = '<div class="warn"><h4>AtenÃ§Ã£o â '+list.length+' inconsistÃªncia(s)</h4><ul>';
  list.forEach(function(w){ html += '<li>'+w+'</li>'; });
  html += '</ul></div>';
  el.innerHTML = html;
}

function loadF(inp, key){
  var f = inp.files[0];
  if(!f) return;
  var reader = new FileReader();
  reader.onload = function(e){
    try{
      var wb = XLSX.read(new Uint8Array(e.target.result), {type:'array', cellDates:true});
      var ws = wb.Sheets[wb.SheetNames[0]];
      var rows = XLSX.utils.sheet_to_json(ws, {defval:null, raw:false});
      if(key==='serv' && rows.length>0){
        var firstVals = Object.values(rows[0]);
        if(firstVals.indexOf('Pasta')>=0 || firstVals.indexOf('Natureza')>=0){
          var oldKeys = Object.keys(rows[0]);
          var newKeys = oldKeys.map(function(k){ return rows[0][k]; });
          rows = rows.slice(1).map(function(r){
            var o={};
            oldKeys.forEach(function(ok,i){ if(newKeys[i]) o[newKeys[i]]=r[ok]; });
            return o;
          });
        }
      }
      ST[key] = rows;
      document.getElementById('dz-'+key).classList.add('ok');
      var ids = {proc:'sp', serv:'ss', aud:'sa', dec:'sd'};
      document.getElementById(ids[key]).textContent = 'â '+f.name+' ('+rows.length+')';
    }catch(ex){
      showErr('Erro ao ler '+f.name+': '+ex.message);
    }
  };
  reader.readAsArrayBuffer(f);
}

function loadLogo(inp){
  var f = inp.files[0]; if(!f) return;
  var r = new FileReader();
  r.onload = function(e){
    ST.logo = e.target.result;
    var p = document.getElementById('lprev');
    p.src = e.target.result; p.style.display='block';
    document.getElementById('sl').textContent = 'â '+f.name;
    document.getElementById('dz-logo').classList.add('ok');
  };
  r.readAsDataURL(f);
}

var COLORS = ['#8B0E1A','#A81020','#C01020','#D42030','#E05060','#E89090','#F5C8CC','#C4A85A','#A89860','#6B6040'];
var GOLD = '#C4A85A';

function mkBar(id, labels, vals, horiz, colors){
  if(CH[id]) CH[id].destroy();
  var bg = colors || vals.map(function(_,i){ return COLORS[i]||COLORS[COLORS.length-1]; });
  var dlPlugin = {};
  try{
    if(window.ChartDataLabels){
      dlPlugin = {datalabels:{
        anchor: horiz?'end':'end', align: horiz?'end':'top',
        color: '#333', font:{size:10, weight:'700', family:"Inter,sans-serif"},
        formatter: function(v){ return v>0?v:''; },
        clamp:true, clip:false,
        padding: horiz?{right:6}:{top:2}
      }};
    }
  }catch(e){}
  CH[id] = new Chart(document.getElementById(id), {
    type:'bar',
    data:{labels:labels, datasets:[{data:vals, backgroundColor:bg, borderRadius:4, borderSkipped:false}]},
    options:{
      indexAxis: horiz?'y':'x', responsive:true, maintainAspectRatio:false,
      plugins: Object.assign({legend:{display:false}}, dlPlugin),
      scales:{
        x:{grid:{color:horiz?'#F5E8EA':'transparent'}, ticks:{color:'#9A5060',font:{size:10}}, border:{display:false}},
        y:{grid:{color:horiz?'transparent':'#F5E8EA'}, ticks:{color:'#4A2030',font:{size:10}}, border:{display:false}}
      },
      layout:{padding:{right:horiz?40:0, top:24}}
    }
  });
}

function mkDonut(id, labels, vals){
  if(CH[id]) CH[id].destroy();
  var dlPlugin = {};
  try{
    if(window.ChartDataLabels){
      dlPlugin = {datalabels:{
        color:'#fff', font:{size:10, weight:'700'},
        formatter:function(v,ctx){
          var t=ctx.dataset.data.reduce(function(a,b){return a+b;},0);
          return v>0 ? Math.round(v/t*100)+'%' : '';
        }
      }};
    }
  }catch(e){}
  CH[id] = new Chart(document.getElementById(id), {
    type:'doughnut',
    data:{labels:labels, datasets:[{data:vals, backgroundColor:['#8B0E1A','#E89090','#C4A85A'], borderWidth:2, borderColor:'#fff'}]},
    options:{
      cutout:'56%', responsive:true, maintainAspectRatio:false,
      plugins: Object.assign({legend:{position:'bottom', labels:{font:{family:'Inter,sans-serif',size:10},color:'#4A2030',padding:8}},
        tooltip:{callbacks:{label:function(ctx){return ' '+ctx.label+': '+ctx.parsed;}}}
      }, dlPlugin)
    }
  });
}

function fDate(v){
  if(!v) return '';
  try{ var d=new Date(v); return isNaN(d)?String(v):d.toLocaleDateString('pt-BR'); }catch(e){return String(v);}
}
function fVal(v){
  var n=parseFloat(v);
  return isNaN(n)||n===0?'R$ 0,00':'R$ '+n.toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});
}
function fValBlankAsZero(v){
  var n=parseFloat(v);
  return isNaN(n)?'R$ 0,00':'R$ '+n.toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});
}
function esc(v){
  return String(v||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function getField(rows, candidates){
  if(!rows||rows.length===0) return null;
  for(var i=0;i<candidates.length;i++){
    if(candidates[i] in rows[0]) return candidates[i];
  }
  return null;
}

function filterClients(rows, field, clients){
  if(!rows||!field||!clients) return rows||[];
  return rows.filter(function(r){ return r[field] && clients.indexOf(String(r[field]).trim())>=0; });
}

function generate(){
  clrErr();
  
  var title = document.getElementById('ttl').value.trim() || 'Dashboard JurÃ­dico';
  var cliRaw = document.getElementById('cli').value.trim();
  var yrRaw = document.getElementById('ayr').value.trim();
  var clients = cliRaw ? cliRaw.split(';').map(function(s){return s.trim();}).filter(Boolean) : null;

  Object.keys(CH).forEach(function(k){ try{CH[k].destroy();}catch(e){} });
  var warns = [];

  var procFld = getField(ST.proc, ['Cliente principal','cliente principal']);
  var servFld = getField(ST.serv, ['Cliente principal','cliente principal']);
  var audFld  = ST.aud ? getField(ST.aud,  ['Cliente Processo','Cliente processo']) : null;

  if(ST.proc && clients && !procFld) warns.push('Campo "Cliente principal" nÃ£o encontrado na planilha de Processos.');
  if(ST.serv && clients && !servFld) warns.push('Campo "Cliente principal" nÃ£o encontrado na planilha de ServiÃ§os.');
  if(ST.aud && clients && !audFld) warns.push('Campo "Cliente Processo" nÃ£o encontrado na planilha de AudiÃªncias/Prazos.');

  var proc = ST.proc ? ((clients && procFld) ? filterClients(ST.proc, procFld, clients) : ST.proc) : [];
  var serv = ST.serv ? ((clients && servFld) ? filterClients(ST.serv, servFld, clients) : ST.serv) : [];
  var aud  = ST.aud ? ((clients && audFld) ? filterClients(ST.aud, audFld, clients) : ST.aud) : [];

  // (proc.length check removed - handled by procblock visibility)

  // Processes â Ativo+Suspenso = Ativo; ignore Inativo
  var active   = proc.filter(function(r){ return r['Status']==='Ativo'||r['Status']==='Suspenso'; });
  var archived = proc.filter(function(r){ return r['Status']==='Arquivado'; });

  var natC={};
  active.forEach(function(r){ var n=r['Natureza']||'N/A'; natC[n]=(natC[n]||0)+1; });
  var natSort = Object.entries(natC).sort(function(a,b){return b[1]-a[1];});

  var compC={};
  active.forEach(function(r){ var c=r[procFld||'Cliente principal']||'N/A'; compC[c]=(compC[c]||0)+1; });
  var compSort = Object.entries(compC).sort(function(a,b){return b[1]-a[1];}).slice(0,8);
  var multi = compSort.length > 1;

  var regY={};
  proc.filter(function(r){ return r['Status']==='Ativo'||r['Status']==='Suspenso'||r['Status']==='Arquivado'; })
      .forEach(function(r){
        var d=r['Data do cadastro']; if(!d) return;
        var y=parseYear(d); if(y>=2008&&y<=2030) regY[y]=(regY[y]||0)+1;
      });

  var archY={};
  archived.forEach(function(r){
    var d=r['Data do encerramento']; if(!d) return;
    var y=parseYear(d); if(y>=2008&&y<=2030) archY[y]=(archY[y]||0)+1;
  });

  // Services
  // Detect date column name (old files use 'Data', new exports use 'Data do cadastro')
  var servDateFld = (serv.length>0 && serv[0]['Data do cadastro']!==undefined) ? 'Data do cadastro' : 'Data';
  var servY={};
  serv.forEach(function(r){
    var d=r[servDateFld]; if(!d) return;
    var y=parseYear(d); if(y>=2012&&y<=2030) servY[y]=(servY[y]||0)+1;
  });
  var syArr = Object.keys(servY).map(Number).sort();
  var refY = yrRaw ? parseInt(yrRaw) : (syArr[syArr.length-1] || new Date().getFullYear());
  var refYCount = servY[refY] || 0;
  if(refYCount===0 && serv.length>0) warns.push('Nenhum serviÃ§o encontrado para '+refY+'. Deixe o campo em branco para usar o ano mais recente.');

  // AudiÃªncias â current year only
  var nAud=0, nPrz=0, nPer=0;
  if(aud.length===0 && ST.aud && clients) warns.push('Nenhum registro de Aud/Prazos/PerÃ­cias encontrado para os clientes. Verifique os nomes no campo "Cliente Processo".');
  aud.forEach(function(r){
    var d=r['Data de inÃ­cio']; if(!d) return;
    if(parseYear(d)!==refY) return;
    var t=String(r['Tipo']||'');
    if(t==='AudiÃªncia') nAud++;
    else if(t==='Prazo') nPrz++;
    else if(t==='PerÃ­cia'||t==='Pauta de Julgamento') nPer++;
  });


  // ââ Item 2: Cadastrados vs Encerrados por mÃªs no ano vigente ââ
  var MONTHS = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  var cmpRegMon = new Array(12).fill(0);
  var cmpArcMon = new Array(12).fill(0);
  proc.filter(function(r){return r['Status']==='Ativo'||r['Status']==='Suspenso'||r['Status']==='Arquivado';})
      .forEach(function(r){
        var d=r['Data do cadastro']; if(!d) return;
        var dt=d instanceof Date?d:new Date(String(d).trim());
        if(isNaN(dt.getTime())) return;
        if(parseYear(d)===refY) cmpRegMon[dt.getMonth()]++;
      });
  proc.filter(function(r){return r['Status']==='Arquivado';})
      .forEach(function(r){
        var d=r['Data do encerramento']; if(!d) return;
        var dt=d instanceof Date?d:new Date(String(d).trim());
        if(isNaN(dt.getTime())) return;
        if(parseYear(d)===refY) cmpArcMon[dt.getMonth()]++;
      });

  // ââ Item 3: DistribuiÃ§Ã£o de fases (ativos+suspensos) ââ
  var faseC = {};
  active.forEach(function(r){
    var f = r['Fase'] || 'NÃ£o informado';
    faseC[f] = (faseC[f]||0)+1;
  });
  var faseSort = Object.entries(faseC).sort(function(a,b){return b[1]-a[1];});

  // ââ Item 4: Passivo processual â probabilidade ProvÃ¡vel ââ
  var provTotalCausa = 0, provTotalEnvolvido = 0, provQtd = 0;
  active.forEach(function(r){
    var prob = String(r['Tipo da probabilidade atual']||'').trim();
    if(prob !== 'ProvÃ¡vel' && prob !== 'Perda' && prob !== 'Perda ProvÃ¡vel') return;
    var causa = parseFloat(r['Valor da causa']) || 0;
    var envol  = parseFloat(r['Valor envolvido']) || 0;
    if(causa > 0 || envol > 0){
      provTotalCausa    += causa;
      provTotalEnvolvido += envol;
      provQtd++;
    }
  });
  var provReducao = provTotalCausa - provTotalEnvolvido;

  // ââ Item 5: AudiÃªncias Ãºltimos 3 anos ââ
  var audYears = [refY-2, refY-1, refY];
  var audByYear = {};
  audYears.forEach(function(y){ audByYear[y]=0; });
  aud.forEach(function(r){
    var d=r['Data de inÃ­cio']; if(!d) return;
    var y=parseYear(d);
    if(audYears.indexOf(y)>=0 && String(r['Tipo']||'')==='AudiÃªncia') audByYear[y]=(audByYear[y]||0)+1;
  });

  var sCompY={};
  serv.forEach(function(r){
    var d=r[servDateFld]; if(!d) return;
    if(parseYear(d)!==refY) return;
    var c=r[servFld||'Cliente principal']||'N/A'; sCompY[c]=(sCompY[c]||0)+1;
  });
  var sCYS = Object.entries(sCompY).sort(function(a,b){return b[1]-a[1];}).slice(0,8);

  var sTipoY={};
  serv.forEach(function(r){
    var d=r[servDateFld]; if(!d) return;
    if(parseYear(d)!==refY) return;
    var t=uT(r['Tipo']); if(t==='Atividade Interna') return;
    sTipoY[t]=(sTipoY[t]||0)+1;
  });
  var sTYS = Object.entries(sTipoY).sort(function(a,b){return b[1]-a[1];}).slice(0,10);
  // Services by Natureza (refYear)
  var sNatY={};
  serv.forEach(function(r){
    var d=r[servDateFld]; if(!d) return;
    if(parseYear(d)!==refY) return;
    var n=extractNat(r['Assunto']); sNatY[n]=(sNatY[n]||0)+1;
  });
  var sNYS = Object.entries(sNatY).sort(function(a,b){return b[1]-a[1];}).slice(0,12);


  showWarns(warns);

  // Header
  document.getElementById('dtitle').textContent = title;
  document.getElementById('dsub').textContent = 'Dashboard JurÃ­dico Executivo â ReferÃªncia '+refY;
  // Build services row dynamically â hide company chart for single client
  var sMulti = sCYS.length > 1;
  var rowserv = document.getElementById('rowserv');
  rowserv.className = 'cg ' + (sMulti ? 'c3' : 'c2');
  var compCard = sMulti
    ? '<div class="cc"><div class="ct" id="sct1">â</div><div class="ch" style="height:210px"><canvas id="chsc"></canvas></div></div>'
    : '';
  rowserv.innerHTML = compCard +
    '<div class="cc"><div class="ct" id="sct2">â</div><div class="ch" style="height:210px"><canvas id="chst"></canvas></div></div>' +
    '<div class="cc"><div class="ct" id="sct3">â</div><div class="ch" style="height:210px"><canvas id="chsn"></canvas></div></div>';

  if(sMulti) document.getElementById('sct1').textContent = 'Total de ServiÃ§os por Empresa Realizados em '+refY;
  document.getElementById('sct2').textContent = 'Total de ServiÃ§os Realizados em '+refY+' â As 10 Atividades Mais Executadas';
  document.getElementById('sct3').textContent = 'Total de ServiÃ§os Realizados em '+refY+' por Natureza';
  if(ST.logo){
    document.getElementById('clogoimg').src=ST.logo;
    document.getElementById('clogoimg').style.display='block';
    document.getElementById('clph').style.display='none';
  }

  // KPI row 1
  document.getElementById('kpi1').innerHTML =
    kpiCard('c7','Processos Ativos',active.length,'Ativos') +
    kpiCard('c8x','Processos Arquivados',archived.length,'Encerrados') +
    kpiCard('c4','Total da Carteira',active.length+archived.length,'Ativos e Encerrados') +
    (ST.serv ? kpiCard('gd','Serviços em '+refY,refYCount,'Serviços realizados') : '');

  var audLbl = ST.aud ? 'Clientes filtrados' : 'Planilha nÃ£o carregada';
  document.getElementById('kpi2').innerHTML =
    kpiCard('c6','AudiÃªncias em '+refY,nAud,audLbl) +
    kpiCard('c5','Prazos em '+refY,nPrz,audLbl) +
    kpiCard('gd','PerÃ­cias em '+refY,nPer,audLbl);

  // Top charts row
  var rtop = document.getElementById('rowtop');
  rtop.className = 'cg '+(multi?'c3':'c2');
  var compHtml = multi ? '<div class="cc"><div class="ct">Processos Ativos por Empresa</div><div class="ch" style="height:210px"><canvas id="chcmp"></canvas></div></div>' : '';
  rtop.innerHTML =
    '<div class="cc"><div class="ct">DistribuiÃ§Ã£o por Status</div><div class="ch" style="height:210px"><canvas id="chpie"></canvas></div></div>' +
    '<div class="cc"><div class="ct">Processos Ativos por Natureza</div><div class="ch" style="height:210px"><canvas id="chnat"></canvas></div></div>' +
    compHtml;

  // Draw charts
  mkDonut('chpie', ['Ativos','Arquivados'], [active.length, archived.length]);
  mkBar('chnat', natSort.map(function(e){return e[0];}), natSort.map(function(e){return e[1];}), true,
    natSort.map(function(_,i){return COLORS[i]||COLORS[COLORS.length-1];}));
  if(multi && document.getElementById('chcmp')){
    mkBar('chcmp',
      compSort.map(function(e){return e[0].length>22?e[0].substring(0,22)+'â¦':e[0];}),
      compSort.map(function(e){return e[1];}), true,
      compSort.map(function(_,i){return COLORS[i]||COLORS[COLORS.length-1];}));
  }

  var ryK = Object.keys(regY).map(Number).sort();
  mkBar('chry', ryK, ryK.map(function(y){return regY[y]||0;}), false,
    ryK.map(function(y){return y>=2017?'#8B0E1A':'#E89090';}));

  var ayK = Object.keys(archY).map(Number).sort();
  mkBar('chay', ayK, ayK.map(function(y){return archY[y]||0;}), false,
    ayK.map(function(){return GOLD;}));

  // ââ Item 2a: Cadastrados por mÃªs no ano vigente ââ
  var todayMon = new Date().getMonth(); // 0-based, e.g. 4 = Mai
  var cmpRegLabels = MONTHS.slice(0, todayMon + 1);
  var cmpRegData   = cmpRegMon.slice(0, todayMon + 1);
  document.getElementById('ctcmp').textContent = 'Processos Cadastrados em '+refY+' (por mÃªs)';
  if(CH['chcmp2']) CH['chcmp2'].destroy();
  CH['chcmp2'] = new Chart(document.getElementById('chcmp2'), {
    type:'bar',
    data:{
      labels: cmpRegLabels,
      datasets:[
        {label:'Cadastrados', data:cmpRegData,
          backgroundColor:'#8B0E1A',borderRadius:3,borderSkipped:false}
      ]
    },
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{
        legend:{position:'bottom',labels:{font:{family:'Inter,sans-serif',size:10},color:'#4A2030',padding:8}},
        datalabels:{anchor:'end',align:'top',color:'#333',font:{size:9,weight:'700'},
          formatter:function(v){return v>0?v:'';},clamp:true,clip:false}},
      scales:{
        x:{grid:{display:false},ticks:{color:'#4A2030',font:{size:10}},border:{display:false}},
        y:{grid:{color:'#F5E8EA'},ticks:{color:'#9A5060',font:{size:10},stepSize:1},border:{display:false}}},
      layout:{padding:{top:24}}}
  });

  // ââ Item 2b: Encerrados por mÃªs no ano vigente (atÃ© hoje) ââ
  var cmpArcLabels = MONTHS.slice(0, todayMon + 1);
  var cmpArcData   = cmpArcMon.slice(0, todayMon + 1);
  document.getElementById('ctcmp3').textContent = 'Processos Encerrados em '+refY+' (por mÃªs)';
  if(CH['chcmp3']) CH['chcmp3'].destroy();
  CH['chcmp3'] = new Chart(document.getElementById('chcmp3'), {
    type:'bar',
    data:{
      labels: cmpArcLabels,
      datasets:[
        {label:'Encerrados', data:cmpArcData,
          backgroundColor:GOLD,borderRadius:3,borderSkipped:false}
      ]
    },
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{
        legend:{position:'bottom',labels:{font:{family:'Inter,sans-serif',size:10},color:'#4A2030',padding:8}},
        datalabels:{anchor:'end',align:'top',color:'#333',font:{size:9,weight:'700'},
          formatter:function(v){return v>0?v:'';},clamp:true,clip:false}},
      scales:{
        x:{grid:{display:false},ticks:{color:'#4A2030',font:{size:10}},border:{display:false}},
        y:{grid:{color:'#F5E8EA'},ticks:{color:'#9A5060',font:{size:10},stepSize:1},border:{display:false}}},
      layout:{padding:{top:24}}}
  });

  // ââ Item 3: Fases processuais ââ
  if(CH['chfase']) CH['chfase'].destroy();
  CH['chfase'] = new Chart(document.getElementById('chfase'), {
    type:'bar',
    data:{labels:faseSort.map(function(e){return e[0];}),
      datasets:[{data:faseSort.map(function(e){return e[1];}),
        backgroundColor:faseSort.map(function(_,i){return COLORS[i]||COLORS[COLORS.length-1];}),
        borderRadius:4,borderSkipped:false}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},
        datalabels:{anchor:'end',align:'end',color:'#333',font:{size:10,weight:'700'},
          formatter:function(v){return v;},clamp:true,clip:false,padding:{right:6}}},
      scales:{x:{grid:{color:'#F5E8EA'},ticks:{color:'#9A5060',font:{size:10}},border:{display:false}},
              y:{grid:{display:false},ticks:{color:'#4A2030',font:{size:11}},border:{display:false}}},
      layout:{padding:{right:40}}}
  });

  // ââ Item 4: Passivo processual ââ
  (function(){
    var fmt = function(v){return 'R$ '+v.toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});};
    var pct = provTotalCausa>0 ? Math.round(provReducao/provTotalCausa*100) : 0;
    var html = '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-top:8px">';
    html += '<div style="background:var(--c0);border-radius:8px;padding:14px;text-align:center;border-top:3px solid var(--c6)">';
    html += '<div style="font-size:.62rem;color:var(--mu);text-transform:uppercase;letter-spacing:.05em;font-weight:600">Valor da Causa</div>';
    html += '<div style="font-family:Libre Baskerville,serif;font-size:1.1rem;color:var(--c8);font-weight:700;margin-top:4px">'+fmt(provTotalCausa)+'</div>';
    html += '<div style="font-size:.65rem;color:var(--mu);margin-top:3px">'+provQtd+' processo'+(provQtd!==1?'s':'')+'</div></div>';
    html += '<div style="background:var(--c0);border-radius:8px;padding:14px;text-align:center;border-top:3px solid var(--gd)">';
    html += '<div style="font-size:.62rem;color:var(--mu);text-transform:uppercase;letter-spacing:.05em;font-weight:600">Valor Envolvido</div>';
    html += '<div style="font-family:Libre Baskerville,serif;font-size:1.1rem;color:var(--c8);font-weight:700;margin-top:4px">'+fmt(provTotalEnvolvido)+'</div>';
    html += '<div style="font-size:.65rem;color:var(--mu);margin-top:3px">ExposiÃ§Ã£o real</div></div>';
    html += '<div style="background:#EAF5EA;border-radius:8px;padding:14px;text-align:center;border-top:3px solid #4A8A4A">';
    html += '<div style="font-size:.62rem;color:#2A5A2A;text-transform:uppercase;letter-spacing:.05em;font-weight:600">ReduÃ§Ã£o Obtida</div>';
    html += '<div style="font-family:Libre Baskerville,serif;font-size:1.1rem;color:#2A5A2A;font-weight:700;margin-top:4px">'+fmt(provReducao)+'</div>';
    html += '<div style="font-size:.65rem;color:#2A5A2A;margin-top:3px">'+pct+'% de reduÃ§Ã£o</div></div>';
    html += '</div>';
    if(provQtd===0) html='<div style="padding:20px;text-align:center;color:var(--mu);font-size:.82rem">Nenhum processo com probabilidade de perda e valores cadastrados.</div>';
    document.getElementById('passivo-content').innerHTML=html;
  })();

  // ââ Item 5: AudiÃªncias Ãºltimos 3 anos ââ
  document.getElementById('ctaud3').textContent = 'AudiÃªncias Realizadas â '+audYears[0]+' a '+audYears[2];
  if(CH['chaud3']) CH['chaud3'].destroy();
  CH['chaud3'] = new Chart(document.getElementById('chaud3'), {
    type:'bar',
    data:{labels:audYears.map(String),
      datasets:[{data:audYears.map(function(y){return audByYear[y]||0;}),
        backgroundColor:audYears.map(function(y,i){return i===2?'#8B0E1A':i===1?'#C01020':'#E89090';}),
        borderRadius:5,borderSkipped:false}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},
        datalabels:{anchor:'end',align:'top',color:'#333',font:{size:14,weight:'700'},
          formatter:function(v){return v||0;},clamp:true,clip:false}},
      scales:{x:{grid:{display:false},ticks:{color:'#4A2030',font:{size:12,weight:'600'}},border:{display:false}},
              y:{grid:{color:'#F5E8EA'},ticks:{color:'#9A5060',font:{size:10}},border:{display:false}}},
      layout:{padding:{top:30}}}
  });

  mkBar('chsy', syArr, syArr.map(function(y){return servY[y]||0;}), false,
    syArr.map(function(y){return y===refY?'#8B0E1A':'#E89090';}));

  if(sMulti && document.getElementById('chsc')){
    mkBar('chsc',
      sCYS.map(function(e){return e[0].length>22?e[0].substring(0,22)+'â¦':e[0];}),
      sCYS.map(function(e){return e[1];}), true,
      sCYS.map(function(_,i){return COLORS[i]||COLORS[COLORS.length-1];}));
  }

  mkBar('chst',
    sTYS.map(function(e){return e[0];}),
    sTYS.map(function(e){return e[1];}), true,
    sTYS.map(function(_,i){return COLORS[i]||COLORS[COLORS.length-1];}));

  mkBar('chsn',
    sNYS.map(function(e){return e[0];}),
    sNYS.map(function(e){return e[1];}), true,
    sNYS.map(function(_,i){return COLORS[i]||COLORS[COLORS.length-1];}));

  // ââ DecisÃµes: classification & chart ââ
  (function(){
    var decBlock = document.getElementById('decblock');
    if(!ST.dec || ST.dec.length===0){ decBlock.innerHTML=''; return; }

    // Detect columns from first row
    var sample = ST.dec[0];
    var keys = Object.keys(sample);
    // Column positions (0-based): Natureza=0, Polo=3, ResultadoDecisao=15, DataPublicacao=21
    var poloFld    = keys[3]  || '';
    var resultFld  = keys[15] || '';
    var dataPubFld = keys[21] || '';
    var clientFld  = keys[2]  || '';

    // Filter by clients if provided
    var decRows = clients
      ? ST.dec.filter(function(r){ return r[clientFld] && clients.indexOf(String(r[clientFld]).trim())>=0; })
      : ST.dec;

    if(decRows.length===0){ decBlock.innerHTML=''; return; }

    // Classify decisions
    var nFav=0, nDesfav=0, nAcordo=0;
    var minDate=null, maxDate=null;

    // ââ Text-based classification using decision verbs + pole rule ââ
    function classifyDecision(polo, texto){
      var polo_l = String(polo||'').trim().toLowerCase();
      var polAt = polo_l==='ativo'||polo_l==='autor';
      var polPas= polo_l==='passivo'||polo_l==='rÃ©u'||polo_l==='reu';
      if(!polAt && !polPas) return null;

      var t = String(texto||'').toLowerCase();
      if(!t||t==='nan'||t==='null') return null;

      // ExtinÃ§Ã£o
      if(/declaro\s+extint|extingo\s+o\s+processo|extinta\s+a\s+execu|extinÃ§Ã£o.*processo/.test(t))
        return 'extinto';

      // Parcial procedÃªncia
      if(/parcialmente\s+proced|julgo\s+parcial|procedente\s+em\s+parte|em\s+parte\s+procedente/.test(t)){
        return polAt ? 'favoravel_parcial' : 'desfavoravel_parcial';
      }
      // ProcedÃªncia
      if(/julgo\s+proced[ea]nte(?!\s*ncia)|julgado\s+proced(?!Ãªncia)|julgar\s+proced(?!Ãªncia)|procedÃªncia\s+do|dou\s+provimento|concedo\s+(parcialmente\s+)?a\s+seguran|condeno\s+[ao]\s+r[eÃ©u]|acolho\s+os?\s+pedidos?/.test(t)){
        return polAt ? 'favoravel' : 'desfavoravel';
      }
      // ImprocedÃªncia (inclui "REJEITAR preliminar...julgar IMPROCEDENTES", "decide...IMPROCEDENTES")
      if(/julgo\s+improced[ea]nte|julgado\s+improced|julgar\s+improced|improcedÃªncia\s+do|improcedentes\s+os\s+pedidos|totalmente\s+improcedentes|pedidos?.*improced|rejeito\s+os?\s+pedidos?|decide.*julgar\s+improced|decide.*improcedente|julgar\s+os\s+pedidos.*improced/.test(t)){
        return polAt ? 'desfavoravel' : 'favoravel';
      }
      return null; // verb not identified
    }

    decRows.forEach(function(r){
      var datePub = r[dataPubFld];
      if(datePub){
        var dp = datePub instanceof Date ? datePub : new Date(String(datePub).split('/').reverse().join('-'));
        if(!isNaN(dp.getTime())){
          if(!minDate||dp<minDate) minDate=dp;
          if(!maxDate||dp>maxDate) maxDate=dp;
        }
      }
      var cls = classifyDecision(r[poloFld], r[keys[keys.length-1]]);
      if(cls==='favoravel'||cls==='favoravel_parcial')            nFav++;
      else if(cls==='desfavoravel'||cls==='desfavoravel_parcial') nDesfav++;
      else if(cls==='extinto')                                    nAcordo++;
    });

    var total = nFav+nDesfav+nAcordo;
    if(total===0){ decBlock.innerHTML=''; return; }

    // Period label
    var fmtD=function(d){ return d?d.toLocaleDateString('pt-BR',{month:'short',year:'numeric'}):''; };
    var period = (minDate&&maxDate)
      ? (minDate.getMonth()===maxDate.getMonth()&&minDate.getFullYear()===maxDate.getFullYear()
          ? fmtD(minDate)
          : fmtD(minDate)+' a '+fmtD(maxDate))
      : '';
    var chartTitle = 'Resultado das DecisÃµes'+(period?' â '+period:'');

    // Build chart HTML
    decBlock.innerHTML = '<div class="sd"></div><div class="sbn" style="background:var(--c9)"><h2>ðï¸ DecisÃµes Judiciais</h2></div>' +
      '<div class="cg c2" style="margin-bottom:12px">' +
      '<div class="cc"><div class="ct">'+chartTitle+'</div>' +
      '<div class="ch" style="height:220px"><canvas id="chdec"></canvas></div></div>' +
      '<div class="cc"><div class="ct">Resultado por Tipo</div>' +
      '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;padding:12px 0" id="dec-kpis"></div></div>' +
      '</div>';

    // KPI cards
    var pctFav   = Math.round(nFav/total*100);
    var pctDesfav= Math.round(nDesfav/total*100);
    var pctAcordo= nAcordo>0?Math.round(nAcordo/total*100):0;
    var kpiHtml =
      '<div style="background:#EAF5EA;border-radius:8px;padding:14px;text-align:center;border-top:3px solid #4A8A4A">'+
        '<div style="font-size:.6rem;color:#2A5A2A;text-transform:uppercase;letter-spacing:.05em;font-weight:600">FavorÃ¡veis</div>'+
        '<div style="font-family:Libre Baskerville,serif;font-size:1.8rem;color:#2A5A2A;font-weight:700">'+nFav+'</div>'+
        '<div style="font-size:.65rem;color:#2A5A2A">'+pctFav+'% do total</div></div>'+
      '<div style="background:var(--c0);border-radius:8px;padding:14px;text-align:center;border-top:3px solid var(--c7)">'+
        '<div style="font-size:.6rem;color:var(--mu);text-transform:uppercase;letter-spacing:.05em;font-weight:600">DesfavorÃ¡veis</div>'+
        '<div style="font-family:Libre Baskerville,serif;font-size:1.8rem;color:var(--c8);font-weight:700">'+nDesfav+'</div>'+
        '<div style="font-size:.65rem;color:var(--mu)">'+pctDesfav+'% do total</div></div>'+
      (nAcordo>0?
        '<div style="background:#FFF8E0;border-radius:8px;padding:14px;text-align:center;border-top:3px solid var(--gd)">'+
          '<div style="font-size:.6rem;color:#7A6000;text-transform:uppercase;letter-spacing:.05em;font-weight:600">Extintos</div>'+
          '<div style="font-family:Libre Baskerville,serif;font-size:1.8rem;color:#7A6000;font-weight:700">'+nAcordo+'</div>'+
          '<div style="font-size:.65rem;color:#7A6000">'+pctAcordo+'% do total</div></div>'
        :'<div style="background:var(--c0);border-radius:8px;padding:14px;text-align:center;border-top:3px solid var(--bd)">'+
          '<div style="font-size:.6rem;color:var(--mu);text-transform:uppercase;letter-spacing:.05em;font-weight:600">Total Analisado</div>'+
          '<div style="font-family:Libre Baskerville,serif;font-size:1.8rem;color:var(--c8);font-weight:700">'+total+'</div>'+
          '<div style="font-size:.65rem;color:var(--mu)">decisÃµes classificadas</div></div>');
    document.getElementById('dec-kpis').innerHTML = kpiHtml;

    // Chart
    if(CH['chdec']) CH['chdec'].destroy();
    var decLabels = ['FavorÃ¡veis','DesfavorÃ¡veis'];
    var decVals   = [nFav, nDesfav];
    var decColors = ['#4A8A4A','#8B0E1A'];
    if(nAcordo>0){ decLabels.push('Extintos'); decVals.push(nAcordo); decColors.push('#C4A85A'); }
    CH['chdec'] = new Chart(document.getElementById('chdec'),{
      type:'doughnut',
      data:{labels:decLabels,datasets:[{data:decVals,backgroundColor:decColors,borderWidth:3,borderColor:'#fff'}]},
      options:{cutout:'55%',responsive:true,maintainAspectRatio:false,
        plugins:{
          legend:{position:'bottom',labels:{font:{family:'Inter,sans-serif',size:11},color:'#4A2030',padding:10}},
          datalabels:{color:'#fff',font:{size:11,weight:'700'},
            formatter:function(v,ctx){
              var t=ctx.dataset.data.reduce(function(a,b){return a+b;},0);
              return v>0?Math.round(v/t*100)+'%':'';
            }},
          tooltip:{callbacks:{label:function(ctx){return ' '+ctx.label+': '+ctx.parsed+' ('+Math.round(ctx.parsed/total*100)+'%)';}}}
        }}
    });
  })();

  window._rpt = {proc:proc, active:active, archived:archived, natSort:natSort, title:title, refY:refY, nAud:nAud, nPrz:nPrz, nPer:nPer};

  // === Ocultar seções sem planilha ===
  var procBlock = document.getElementById('procblock');
  var audBlock  = document.getElementById('audblock');
  var servBlock = document.getElementById('servblock');
  if(procBlock){ procBlock.style.display = ST.proc ? '' : 'none'; }
  if(audBlock){  audBlock.style.display  = ST.aud  ? '' : 'none'; }
  if(servBlock){ servBlock.style.display = ST.serv ? '' : 'none'; }

  document.getElementById('cfg').style.display = 'none';
  document.getElementById('dsh').style.display = 'block';
  window.scrollTo(0,0);
}

function kpiCard(col, lbl, val, det){
  var colors = {'c7':'var(--c7)','c8x':'#8A6060','c4':'var(--c4)','gd':'var(--gd)','c6':'var(--c6)','c5':'var(--c5)'};
  var bc = colors[col] || 'var(--c6)';
  return '<div class="kp" style="border-top-color:'+bc+'">' +
    '<div class="kl">'+lbl+'</div>' +
    '<div class="kv">'+val+'</div>' +
    '<div class="kd">'+det+'</div></div>';
}

function showReport(){
  var d = window._rpt; if(!d) return;
  document.getElementById('rtitle').textContent = 'RelatÃ³rio de Processos â '+d.title;
  var now = new Date().toLocaleDateString('pt-BR',{day:'2-digit',month:'long',year:'numeric'});
  document.getElementById('rsub').textContent = 'Gerado em: '+now;
  document.getElementById('rfdate').textContent = now;
  document.getElementById('rnote').textContent = 'Processos com status Ativo e Suspenso, ordenados por Natureza. Os campos Data da DistribuiÃ§Ã£o e ContrÃ¡rio principal sÃ£o preenchidos quando disponÃ­veis na planilha.';
  document.getElementById('rkpis').innerHTML =
    '<div class="rk"><div class="l">Ativos</div><div class="v">'+d.active.length+'</div></div>' +
    '<div class="rk"><div class="l">Arquivados</div><div class="v">'+d.archived.length+'</div></div>' +
    '<div class="rk"><div class="l">Total</div><div class="v">'+(d.active.length+d.archived.length)+'</div></div>' +
    '<div class="rk"><div class="l">AudiÃªncias '+d.refY+'</div><div class="v">'+d.nAud+'</div></div>' +
    '<div class="rk"><div class="l">Prazos '+d.refY+'</div><div class="v">'+d.nPrz+'</div></div>';

  var nh = '<table class="pt"><thead><tr><th>Natureza</th><th style="text-align:right">Qtd.</th><th style="text-align:right">%</th></tr></thead><tbody>';
  d.natSort.forEach(function(e){
    nh += '<tr><td>'+esc(e[0])+'</td><td style="text-align:right;font-weight:700">'+e[1]+'</td><td style="text-align:right">'+Math.round(e[1]/d.active.length*100)+'%</td></tr>';
  });
  nh += '</tbody></table>';
  document.getElementById('rnat').innerHTML = nh;

  var sorted = d.active.slice().sort(function(a,b){return (a['Natureza']||'').localeCompare(b['Natureza']||'');});
  var tbody = document.getElementById('rtbody'); tbody.innerHTML='';
  sorted.forEach(function(r,i){
    var dtD = r['Data da distribuiÃ§Ã£o'] || r['Data do cadastro'];
    var tr = '<tr>' +
      '<td>'+(i+1)+'</td>' +
      '<td><b>'+esc(r['Natureza'])+'</b></td>' +
      '<td title="'+esc(r['AÃ§Ã£o'])+'">'+esc((r['AÃ§Ã£o']||'').substring(0,35))+'</td>' +
      '<td style="white-space:nowrap">'+fDate(dtD)+'</td>' +
      '<td>'+esc((r['Cliente principal']||'').substring(0,28))+'</td>' +
      '<td>'+esc(r['PosiÃ§Ã£o']||'')+'</td>' +
      '<td>'+esc((r['ContrÃ¡rio principal']||'').substring(0,28))+'</td>' +
      '<td>'+esc((r['ÃrgÃ£o']||'').substring(0,25))+'</td>' +
      '<td>'+esc(r['Cidade']||'')+'</td><td>'+esc(r['UF']||'')+'</td>' +
      '<td style="text-align:right;white-space:nowrap">'+fVal(r['Valor da causa'])+'</td>' +
      '<td style="text-align:right;white-space:nowrap">'+fVal(r['Valor envolvido'])+'</td>' +
      '<td>'+esc(r['Tipo da probabilidade atual']||'')+'</td>' +
      '<td>'+esc(r['Faixa de probabilidade atual']||'')+'</td>' +
      '<td>'+esc(r['ClassificaÃ§Ã£o do Processo']||'Comum')+'</td>' +
      '</tr>';
    tbody.innerHTML += tr;
  });
  document.getElementById('dsh').style.display = 'none';
  document.getElementById('rpt').style.display = 'block';
  window.scrollTo(0,0);
}

function exportXLS(){
  var d = window._rpt; if(!d) return;
  var sorted = d.active.slice().sort(function(a,b){return (a['Natureza']||'').localeCompare(b['Natureza']||'');});

  var H='#8B0E1A', H2='#A81020', ALT='#FDE8EA', GC='#C4A85A', BD='#EDD8DA';
  var co='style="padding:5px 8px;border:1px solid '+BD+';font-family:Calibri;font-size:9pt;vertical-align:top;background:';
  var coH='style="padding:7px 8px;background:'+H+';color:#fff;font-weight:700;border:1px solid #6B0010;font-family:Calibri;font-size:9pt;text-align:left"';

  var headers=['Natureza','AÃ§Ã£o','Data da DistribuiÃ§Ã£o','Cliente principal','PosiÃ§Ã£o','ContrÃ¡rio principal','ÃrgÃ£o','Cidade','UF','Valor da causa (R$)','Valor envolvido (R$)','Tipo da probabilidade atual','Faixa de probabilidade atual','ClassificaÃ§Ã£o do Processo'];
  var hRow = '<tr>'+headers.map(function(h){return '<th '+coH+'>'+h+'</th>';}).join('')+'</tr>';

  var dRows = sorted.map(function(r,i){
    var bg = i%2 ? ALT : '#fff';
    var cols=[
      '<b style="color:'+H+'">'+esc(r['Natureza'])+'</b>',
      esc(r['AÃ§Ã£o']),
      fDate(r['Data da distribuiÃ§Ã£o']||r['Data do cadastro']),
      esc(r['Cliente principal']),
      esc(r['PosiÃ§Ã£o']),
      esc(r['ContrÃ¡rio principal']||''),
      esc(r['ÃrgÃ£o']),
      esc(r['Cidade']),
      esc(r['UF']),
      fValBlankAsZero(r['Valor da causa']).replace('R$ ',''),
      fValBlankAsZero(r['Valor envolvido']).replace('R$ ',''),
      esc(r['Tipo da probabilidade atual']),
      esc(r['Faixa de probabilidade atual']),
      esc(r['ClassificaÃ§Ã£o do Processo']||'Comum')
    ];
    return '<tr>'+cols.map(function(v){return '<td '+co+bg+'">'+(v||'')+'</td>';}).join('')+'</tr>';
  }).join('');

  var natRows = d.natSort.map(function(e,i){
    var bg = i%2?ALT:'#fff';
    return '<tr><td style="padding:6px 10px;border:1px solid '+BD+';background:'+bg+';font-family:Calibri;font-weight:600">'+esc(e[0])+'</td><td style="padding:6px 10px;border:1px solid '+BD+';background:'+bg+';text-align:center;font-weight:700;color:'+H+'">'+e[1]+'</td><td style="padding:6px 10px;border:1px solid '+BD+';background:'+bg+';text-align:center">'+Math.round(e[1]/d.active.length*100)+'%</td></tr>';
  }).join('');

  var proc_table =
    '<table>' +
    '<tr><td colspan="'+headers.length+'" style="padding:12px;background:'+H+';color:#fff;font-size:14pt;font-weight:700;font-family:Calibri">'+esc(d.title)+' â Processos Ativos</td></tr>' +
    '<tr><td colspan="'+headers.length+'" style="padding:5px 12px;background:'+H2+';color:#fff;font-size:9pt;font-family:Calibri">Total: '+sorted.length+' processos Â· Gerado em '+new Date().toLocaleDateString('pt-BR')+'</td></tr>' +
    hRow + dRows + '</table>';

  var res_table =
    '<table>' +
    '<tr><td colspan="3" style="padding:12px;background:'+H+';color:#fff;font-size:14pt;font-weight:700;font-family:Calibri">'+esc(d.title.toUpperCase())+' â RESUMO</td></tr>' +
    '<tr><td colspan="3" style="padding:6px 12px;background:'+H2+';color:#fff;font-size:10pt;font-family:Calibri">Gerado em: '+new Date().toLocaleDateString('pt-BR',{day:'2-digit',month:'long',year:'numeric'})+'</td></tr>' +
    '<tr><td colspan="3" style="padding:4px"></td></tr>' +
    '<tr><td style="padding:8px 12px;background:'+GC+';color:#fff;font-weight:700;border:1px solid '+BD+';font-family:Calibri">INDICADOR</td><td style="padding:8px 12px;background:'+GC+';color:#fff;font-weight:700;border:1px solid '+BD+';font-family:Calibri">VALOR</td><td style="background:'+GC+';border:1px solid '+BD+'"></td></tr>' +
    '<tr><td style="padding:6px 12px;border:1px solid '+BD+';font-family:Calibri">Processos Ativos</td><td style="padding:6px 12px;border:1px solid '+BD+';font-weight:700;color:'+H+';text-align:center;font-family:Calibri">'+d.active.length+'</td><td style="border:1px solid '+BD+'"></td></tr>' +
    '<tr><td style="padding:6px 12px;border:1px solid '+BD+';background:'+ALT+';font-family:Calibri">Processos Arquivados</td><td style="padding:6px 12px;border:1px solid '+BD+';background:'+ALT+';font-weight:700;color:'+H+';text-align:center;font-family:Calibri">'+d.archived.length+'</td><td style="border:1px solid '+BD+';background:'+ALT+'"></td></tr>' +
    '<tr><td style="padding:6px 12px;border:1px solid '+BD+';font-family:Calibri">Total Carteira</td><td style="padding:6px 12px;border:1px solid '+BD+';font-weight:700;color:'+H+';text-align:center;font-family:Calibri">'+(d.active.length+d.archived.length)+'</td><td style="border:1px solid '+BD+'"></td></tr>' +
    '<tr><td colspan="3" style="padding:4px"></td></tr>' +
    '<tr><td style="padding:8px;background:'+H2+';color:#fff;font-weight:700;border:1px solid '+BD+';font-family:Calibri">NATUREZA</td><td style="padding:8px;background:'+H2+';color:#fff;font-weight:700;border:1px solid '+BD+';font-family:Calibri;text-align:center">QTD.</td><td style="padding:8px;background:'+H2+';color:#fff;font-weight:700;border:1px solid '+BD+';font-family:Calibri;text-align:center">%</td></tr>' +
    natRows + '</table>';

  var xls = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">\n<head><meta charset="UTF-8">\n<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets>\n<x:ExcelWorksheet><x:Name>Processos Ativos</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>\n<x:ExcelWorksheet><x:Name>Resumo</x:Name></x:ExcelWorksheet>\n</x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->\n</head>\n<body>\n<div id="Processos Ativos">'+proc_table+'</div>\n<div id="Resumo">'+res_table+'</div>\n</body>\n</html>';

  var blob = new Blob(['\ufeff'+xls], {type:'application/vnd.ms-excel;charset=utf-8'});
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href=url; a.download=(d.title||'Relatorio').replace(/[^a-zA-Z0-9\u00C0-\u024F ]/g,'_')+' - Processos Ativos.xls';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(function(){URL.revokeObjectURL(url);}, 5000);
}

function goBack(){ document.getElementById('dsh').style.display='none'; document.getElementById('cfg').style.display='flex'; }
function hideReport(){ document.getElementById('rpt').style.display='none'; document.getElementById('dsh').style.display='block'; }
function printDash(){ document.body.setAttribute('data-print','dash'); window.print(); setTimeout(function(){document.body.removeAttribute('data-print');},2000); }
function shareDash(){
  var nm=document.getElementById("dtitle")?document.getElementById("dtitle").innerText.replace(/[^\w\s-]/g,"").trim().substring(0,50):"dashboard";
  var clone=document.documentElement.cloneNode(true);
  var cloneCfg=clone.querySelector("#cfg");
  if(cloneCfg)cloneCfg.style.display="none";
  var cloneDsh=clone.querySelector("#dsh");
  if(cloneDsh)cloneDsh.style.display="block";
  var cloneTbar=clone.querySelector(".tbar");
  if(cloneTbar)cloneTbar.style.display="none";
  var chartData={};
  for(var id in CH){
    try{
      var cfg=CH[id].config;
      chartData[id]={type:cfg.type,data:JSON.parse(JSON.stringify(cfg.data)),options:JSON.parse(JSON.stringify(cfg.options||{}))};
    }catch(e){}
  }
  var reinitScript=clone.ownerDocument?clone.ownerDocument.createElement("script"):document.createElement("script");
  reinitScript.textContent='window.addEventListener("load",function(){var D='+JSON.stringify(chartData)+';for(var id in D){var el=document.getElementById(id);if(!el)continue;var ctx=el.getContext("2d");new Chart(ctx,{type:D[id].type,data:D[id].data,options:D[id].options});} });';
  var cloneBody=clone.querySelector("body");
  if(cloneBody)cloneBody.appendChild(reinitScript);
  var html="<!DOCTYPE html>"+clone.outerHTML;
  var blob=new Blob([html],{type:"text/html"});
  var a=document.createElement("a");
  a.href=URL.createObjectURL(blob);
  a.download=nm+".html";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(a.href);
}

function printReport(){ document.body.setAttribute('data-print','report'); window.print(); setTimeout(function(){document.body.removeAttribute('data-print');},2000); }

window.addEventListener('load', function(){
  try{ if(window.ChartDataLabels) Chart.register(ChartDataLabels); }catch(e){}
});
"""

# Build complete HTML
html = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Dashboard JurÃ­dico â IGSA</title>
<style>
""" + CSS + """
</style>
</head>
<body>

<!-- CONFIG -->
<div id="cfg">
<div class="cfgh">
  <img src=\"""" + LOGO + """\" alt="IGSA">
  <h1 class="sf">Gerador de Dashboard JurÃ­dico</h1>
  <p>Imaculada Gordiano Sociedade de Advogados</p>
</div>
<div class="card">
<div class="grid2">
<div>
  <div class="stl"><span class="sn">1</span>Planilhas de dados</div>
  <div class="up3">
    <div class="dz" id="dz-proc">
      <input type="file" accept=".xlsx,.xls" onchange="loadF(this,'proc')">
      <div class="dzi">ð</div><div class="dzl">Processos</div>
      <div class="dzs">.xlsx / .xls</div><div class="dzst" id="sp">Nenhum arquivo</div>
    </div>
    <div class="dz" id="dz-serv">
      <input type="file" accept=".xlsx,.xls" onchange="loadF(this,'serv')">
      <div class="dzi">âï¸</div><div class="dzl">ServiÃ§os</div>
      <div class="dzs">.xlsx / .xls</div><div class="dzst" id="ss">Nenhum arquivo</div>
    </div>
    <div class="dz" id="dz-aud">
      <input type="file" accept=".xlsx,.xls" onchange="loadF(this,'aud')">
      <div class="dzi">ðï¸</div><div class="dzl">Aud / Prazos / PerÃ­cias</div>
      <div class="dzs">.xlsx / .xls</div><div class="dzst" id="sa">Nenhum arquivo</div>
    </div>
    <div class="dz" id="dz-dec">
      <input type="file" accept=".xlsx,.xls" onchange="loadF(this,'dec')">
      <div class="dzi">âï¸</div><div class="dzl">DecisÃµes <em style="font-size:.62rem;font-weight:400">(opcional)</em></div>
      <div class="dzs">.xlsx / .xls</div><div class="dzst" id="sd">Nenhum arquivo</div>
    </div>
  </div>
  <div class="stl"><span class="sn">2</span>Logo do cliente / grupo</div>
  <div class="dz" id="dz-logo" style="display:flex;align-items:center;gap:12px;text-align:left">
    <input type="file" accept="image/*" onchange="loadLogo(this)">
    <div style="font-size:1.5rem">ð¢</div>
    <div style="flex:1"><div class="dzl">Carregar logo do cliente</div>
    <div class="dzs">PNG, JPG, SVG</div><div class="dzst" id="sl">Nenhuma imagem</div></div>
    <img id="lprev" style="display:none;max-height:42px;max-width:80px;border-radius:4px" src="" alt="">
  </div>
</div>
<div>
  <div class="stl"><span class="sn">3</span>ConfiguraÃ§Ãµes</div>
  <div class="fld"><label>TÃ­tulo do dashboard</label>
    <input type="text" id="ttl" placeholder="Ex: Alscience Metrologia">
  </div>
  <div class="fld"><label>Cliente(s) â separe por ; para mÃºltiplos</label>
    <input type="text" id="cli" placeholder="Ex: EMPRESA LTDA.;SÃCIO NOME">
    <small>ð¡ Nomes idÃªnticos ao campo "Cliente principal" / "Cliente Processo" das planilhas</small>
  </div>
  <div class="fld"><label>Ano de referÃªncia dos serviÃ§os <em style="font-weight:300;text-transform:none">(opcional)</em></label>
    <input type="text" id="ayr" placeholder="Deixe em branco para usar o ano mais recente">
  </div>
  <div id="errbox" class="err"></div>
  <button class="bgen" onclick="generate()">â¶ Gerar Dashboard</button>
  <div id="warnblock"></div>
  <div style="text-align:center;font-size:.65rem;color:var(--mu);margin-top:7px">Processamento 100% local â sem envio de dados externos</div>
</div>
</div>
</div>
</div>

<!-- DASHBOARD -->
<div id="dsh">
<div class="tbar">
  <button class="tb tbk" onclick="goBack()">â Voltar</button>
  <span class="thi">Imaculada Gordiano Sociedade de Advogados Â· Dashboard JurÃ­dico</span>
  <div class="tbtns">
    <button class="tb trl" onclick="showReport()">ð RelatÃ³rio de Processos</button>
    <button class="tb tpd" onclick="printDash()">â¬ Exportar PDF</button>
        <button class="tb tbk" onclick="shareDash()" title="Baixar HTML para compartilhar">ð¤ Compartilhar</button>
  </div>
</div>
<div class="dw">
  <div class="dh">
    <div class="dhl"><img src=\"""" + LOGO + """\" alt="IGSA"></div>
    <div class="dht">
      <div class="dmt sf" id="dtitle">â</div>
      <div class="dsb" id="dsub">Dashboard JurÃ­dico Executivo</div>
    </div>
    <div class="clo">
      <img id="clogoimg" src="" style="display:none" alt="">
      <div id="clph" class="clph">Logo do Cliente</div>
    </div>
  </div>
<div id="procblock">
  <div class="sbn"><h2>ð Processos Judiciais</h2></div>
  <div class="kr" id="kpi1"></div>
<div id="audblock">
  <div class="kr" id="kpi2" style="margin-bottom:18px"></div>
  <div class="cg c3" id="rowtop"></div>
  <div class="cg c2" style="margin-bottom:12px">
    <div class="cc"><div class="ct">Processos Cadastrados por Ano</div><div class="ch" style="height:190px"><canvas id="chry"></canvas></div></div>
    <div class="cc"><div class="ct">Processos Encerrados por Ano</div><div class="ch" style="height:190px"><canvas id="chay"></canvas></div></div>
  </div>
  <div class="cg c3" style="margin-bottom:12px">
    <div class="cc"><div class="ct" id="ctcmp">â</div><div class="ch" style="height:200px"><canvas id="chcmp2"></canvas></div></div>
    <div class="cc"><div class="ct" id="ctcmp3">â</div><div class="ch" style="height:200px"><canvas id="chcmp3"></canvas></div></div>
    <div class="cc"><div class="ct">DistribuiÃ§Ã£o por Fase Processual</div><div class="ch" style="height:200px"><canvas id="chfase"></canvas></div></div>
  </div>
  <div class="cg c2" style="margin-bottom:12px">
    <div class="cc" id="passivo-card">
      <div class="ct">ReduÃ§Ã£o do Passivo Processual â Probabilidade de Perda</div>
      <div id="passivo-content" style="padding:8px 0"></div>
    </div>
    <div class="cc"><div class="ct" id="ctaud3">â</div><div class="ch" style="height:200px"><canvas id="chaud3"></canvas></div></div>
  </div>
</div><!-- /audblock -->
  <div id="decblock"></div>
</div><!-- /procblock -->
<div id="servblock">
  <div class="sd"></div>
  <div class="sbn" style="background:var(--c7)"><h2>âï¸ ServiÃ§os Realizados</h2></div>
  <div class="cg c1" style="margin-bottom:12px">
    <div class="cc"><div class="ct">Total de ServiÃ§os por Ano</div><div class="ch" style="height:150px"><canvas id="chsy"></canvas></div></div>
  </div>
  <div id="rowserv"></div>
</div><!-- /servblock -->
</div>
</div>

<!-- RELATÃRIO -->
<div id="rpt">
<div class="tbar">
  <button class="tb tbk" onclick="hideReport()">â Voltar ao Dashboard</button>
  <span class="thi" style="color:var(--c3)">RelatÃ³rio de Processos â Imaculada Gordiano</span>
  <button class="tb tpd" onclick="printReport()">â¬ Exportar PDF</button>
</div>
<div class="rw">
  <div class="rh">
    <img src=\"""" + LOGO + """\" alt="IGSA">
    <div class="rtb">
      <h1 id="rtitle">RelatÃ³rio de Processos</h1>
      <p id="rsub">â</p>
    </div>
  </div>
  <div class="rnote" id="rnote"></div>
  <div class="rbtns">
    <button class="rb rxl" onclick="exportXLS()">â¬ Exportar Excel (.xls)</button>
    <button class="rb rpf" onclick="printReport()">â¬ Exportar PDF (Imprimir)</button>
  </div>
  <div class="rks" id="rkpis"></div>
  <div class="rs">Processos Ativos e Suspensos por Natureza</div>
  <div id="rnat"></div>
  <div class="rs">Lista Completa â Processos Ativos e Suspensos</div>
  <table class="pt" id="rptable">
    <thead><tr>
      <th>#</th><th>Natureza</th><th>AÃ§Ã£o</th><th>Data DistribuiÃ§Ã£o</th>
      <th>Cliente principal</th><th>PosiÃ§Ã£o</th><th>ContrÃ¡rio principal</th>
      <th>ÃrgÃ£o</th><th>Cidade</th><th>UF</th>
      <th>Valor da causa</th><th>Valor envolvido</th>
      <th>Probabilidade</th><th>Faixa</th><th>ClassificaÃ§Ã£o</th>
    </tr></thead>
    <tbody id="rtbody"></tbody>
  </table>
  <div style="margin-top:32px;padding-top:12px;border-top:1px solid var(--bd);font-size:.63rem;color:var(--mu);text-align:center">
    Imaculada Gordiano Sociedade de Advogados Â· Gerado em <span id="rfdate"></span>
  </div>
</div>
</div>

    <script>""" + XLSX_JS + """</script>
    <script>""" + CHART_JS + """</script>
    <script>""" + DATALABELS_JS + """</script>
<script>
""" + JS + """
</script>
</body>
</html>"""

# Display in Streamlit
st.set_page_config(page_title="Dashboard Jur\u00eddico - IGSA", layout="wide")
st.components.v1.html(html, height=900, scrolling=True)


# ââ Ferramenta: Corrigir HTML antigo (sem grÃ¡ficos) ââââââââââââââââââââââââââ
with st.sidebar:
    st.markdown("---")
    st.markdown("### ð§ Corrigir HTML existente")
    st.caption("FaÃ§a upload de um HTML gerado antes do fix para embutir as bibliotecas automaticamente.")
    uploaded_fix = st.file_uploader("Selecione o arquivo HTML", type=["html"], key="fix_uploader")
    if uploaded_fix is not None:
        with st.spinner("Baixando bibliotecas e corrigindo... (pode levar ~20s)"):
            import re as _re
            _html = uploaded_fix.read().decode("utf-8")
            _xlsx   = _fetch_lib("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js")
            _chart  = _fetch_lib("https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js")
            _dlbl   = _fetch_lib("https://cdnjs.cloudflare.com/ajax/libs/chartjs-plugin-datalabels/2.2.0/chartjs-plugin-datalabels.min.js")
            _html = _re.sub(r'<script src="https://cdnjs\.cloudflare\.com/ajax/libs/xlsx/[^"]+"></script>',
                            f'<script>{_xlsx}</script>', _html)
            _html = _re.sub(r'<script src="https://cdnjs\.cloudflare\.com/ajax/libs/Chart\.js/[^"]+"></script>',
                            f'<script>{_chart}</script>', _html)
            _html = _re.sub(r'<script src="https://cdnjs\.cloudflare\.com/ajax/libs/chartjs-plugin-datalabels/[^"]+"></script>',
                            f'<script>{_dlbl}</script>', _html)
            _html = _re.sub(r'<link[^>]*fonts\.googleapis\.com[^>]*>', '', _html)
            _out_name = uploaded_fix.name.replace(".html", "_OFFLINE.html")
            st.success("â HTML corrigido com sucesso!")
            st.download_button(
                label="â¬ï¸ Baixar HTML corrigido (funciona offline e no Teams)",
                data=_html.encode("utf-8"),
                file_name=_out_name,
                mime="text/html"
            )


# ââ Ferramenta: Publicar dashboard no Netlify ââââââââââââââââââââââââââââââââââââââââââââ
with st.sidebar:
    st.markdown("---")
    st.markdown("### ð Publicar no Netlify")
    st.caption("Publique o HTML gerado e receba um link para enviar por e-mail ou WhatsApp.")

    _netlify_token = ""
    try:
        _netlify_token = st.secrets.get("NETLIFY_TOKEN", "")
    except Exception:
        pass
    if not _netlify_token:
        _netlify_token = st.text_input("Token do Netlify (Personal Access Token)", type="password", key="netlify_token_input")
        st.caption("Crie seu token em: app.netlify.com")

    _html_to_publish = st.file_uploader("Selecione o HTML do dashboard", type=["html"], key="netlify_uploader")

    if _html_to_publish is not None and _netlify_token:
        if st.button("\U0001f4e4 Publicar e gerar link", key="netlify_publish_btn"):
            import urllib.request as _ur2
            import json as _json2
            import zipfile as _zf
            import io as _io2

            with st.spinner("Publicando no Netlify... aguarde"):
                try:
                    _html_bytes = _html_to_publish.read()

                    # netlify.toml para forÃ§ar Content-Type correto
                    _toml_lines = [
                        "[[headers]]",
                        '  for = "/*"',
                        "  [headers.values]",
                        '    Content-Type = "text/html; charset=utf-8"'
                    ]
                    _toml_bytes = "\n".join(_toml_lines).encode("utf-8")

                    # Cria ZIP com index.html + netlify.toml
                    _zip_buf = _io2.BytesIO()
                    with _zf.ZipFile(_zip_buf, "w", _zf.ZIP_DEFLATED) as _zobj:
                        _zobj.writestr("index.html", _html_bytes)
                        _zobj.writestr("netlify.toml", _toml_bytes)
                    _zip_data = _zip_buf.getvalue()

                    # Passo 1: Criar site
                    _req_site = _ur2.Request(
                        "https://api.netlify.com/api/v1/sites",
                        data=_json2.dumps({}).encode("utf-8"),
                        headers={
                            "Authorization": f"Bearer {_netlify_token}",
                            "Content-Type": "application/json"
                        },
                        method="POST"
                    )
                    with _ur2.urlopen(_req_site) as _r:
                        _site = _json2.loads(_r.read().decode("utf-8"))
                    _site_id = _site["id"]

                    # Passo 2: Deploy via ZIP
                    _req_deploy = _ur2.Request(
                        f"https://api.netlify.com/api/v1/sites/{_site_id}/deploys",
                        data=_zip_data,
                        headers={
                            "Authorization": f"Bearer {_netlify_token}",
                            "Content-Type": "application/zip"
                        },
                        method="POST"
                    )
                    with _ur2.urlopen(_req_deploy) as _r:
                        _deploy = _json2.loads(_r.read().decode("utf-8"))

                    # Polling atÃ© ficar pronto
                    import time as _time
                    _deploy_id = _deploy["id"]
                    _status = _deploy
                    for _ in range(15):
                        _time.sleep(2)
                        _req_check = _ur2.Request(
                            f"https://api.netlify.com/api/v1/deploys/{_deploy_id}",
                            headers={"Authorization": f"Bearer {_netlify_token}"}
                        )
                        with _ur2.urlopen(_req_check) as _r:
                            _status = _json2.loads(_r.read().decode("utf-8"))
                        if _status.get("state") in ("ready", "current"):
                            break

                    _link = _status.get("ssl_url") or _status.get("url") or _site.get("ssl_url") or _site.get("url", "")
                    st.success("\u2705 Dashboard publicado com sucesso!")
                    st.markdown("### \U0001f517 Link para enviar ao cliente:")
                    st.code(_link, language=None)
                    st.caption("Copie o link acima e envie por e-mail ou WhatsApp.")
                except Exception as _e:
                    st.error(f"Erro ao publicar: {str(_e)}")
    elif _html_to_publish is not None and not _netlify_token:
        st.warning("\u26a0\ufe0f Token do Netlify n\u00e3o encontrado nos secrets.")
