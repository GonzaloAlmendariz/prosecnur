# =============================================================================
# Tema visual de reporte_interactivo()
# - CSS parametrizable
# - JS navbar pill
# - Paleta de colores de la app
# =============================================================================

`%||%` <- get0("%||%", ifnotfound = function(x, y) if (!is.null(x)) x else y)

# -----------------------------------------------------------------------------
# Paleta visual por defecto
# -----------------------------------------------------------------------------

reporte_interactivo_theme_default <- function() {
  list(
    color_primario      = "#002457",
    color_fondo_app     = "#f5f6fa",
    color_borde         = "#e6e9f2",
    color_texto         = "#1f2933",
    color_texto_suave   = "#5f6b7a",
    color_superficie    = "#ffffff",
    color_superficie_2  = "#fafbff",
    color_header_tabla  = "#f1f3f9"
  )
}

# -----------------------------------------------------------------------------
# Helpers internos de tema
# -----------------------------------------------------------------------------

.css_escape <- function(x) {
  x <- as.character(x)[1]
  x <- gsub("\\\\", "\\\\\\\\", x)
  x <- gsub("\"", "\\\\\"", x)
  x
}

.theme_merge <- function(theme_app = NULL) {
  base <- reporte_interactivo_theme_default()
  if (is.null(theme_app)) return(base)

  nm <- intersect(names(theme_app), names(base))
  if (length(nm)) {
    base[nm] <- theme_app[nm]
  }
  base
}

# -----------------------------------------------------------------------------
# CSS de la app
# -----------------------------------------------------------------------------

reporte_interactivo_theme_css <- function(theme_app = NULL) {

  th <- .theme_merge(theme_app)

  color_primario     <- .css_escape(th$color_primario)
  color_fondo_app    <- .css_escape(th$color_fondo_app)
  color_borde        <- .css_escape(th$color_borde)
  color_texto        <- .css_escape(th$color_texto)
  color_texto_suave  <- .css_escape(th$color_texto_suave)
  color_superficie   <- .css_escape(th$color_superficie)
  color_superficie_2 <- .css_escape(th$color_superficie_2)
  color_header_tabla <- .css_escape(th$color_header_tabla)

  shadow_soft <- "rgba(0, 36, 87, 0.06)"
  shadow_med  <- "rgba(0, 36, 87, 0.07)"
  shadow_low  <- "rgba(0, 36, 87, 0.04)"
  focus_ring  <- "rgba(0, 36, 87, 0.15)"
  prim_005    <- "rgba(0, 36, 87, 0.05)"
  prim_006    <- "rgba(0, 36, 87, 0.06)"
  prim_020    <- "rgba(0, 36, 87, 0.20)"
  prim_035    <- "rgba(0, 36, 87, 0.35)"
  prim_085    <- "rgba(230, 233, 242, 0.85)"
  pill_border <- "rgba(0, 36, 87, 0.10)"

  css <- "
/* ============================================================
   ====== Base ======
   ============================================================ */
body { background: __COLOR_FONDO_APP__; color: __COLOR_TEXTO__; padding-top: 10px; }
.container-fluid { max-width: 1400px; padding-left: 15px; padding-right: 15px; }

/* ============================================================
   ====== Tipografía ======
   ============================================================ */
h2, h3, h4 { font-weight: 800; color: __COLOR_PRIMARIO__; }
.title { font-weight: 900; color: __COLOR_PRIMARIO__; }

/* ============================================================
   ====== Sidebar ======
   ============================================================ */
.well, .sidebarPanel {
  background: __COLOR_SUPERFICIE__ !important;
  border: 1px solid __COLOR_BORDE__ !important;
  border-radius: 16px !important;
  box-shadow: 0 12px 28px __SHADOW_SOFT__;
}
.sidebar h3 { margin-top: 0; color: __COLOR_PRIMARIO__; }
.sidebar p  { color: __COLOR_TEXTO_SUAVE__; font-size: 13px; }
.sidebar hr { border-top: 1px solid #edf0f7; }

/* ============================================================
   ====== Inputs ======
   ============================================================ */
.selectize-input, .form-control {
  border-radius: 12px !important;
  border: 1px solid __COLOR_BORDE__ !important;
  box-shadow: none !important;
  font-size: 13px;
}
.selectize-input.focus, .form-control:focus {
  border-color: __COLOR_PRIMARIO__ !important;
  box-shadow: 0 0 0 3px __FOCUS_RING__ !important;
}

/* ============================================================
   ====== Botones ======
   ============================================================ */
.btn {
  border-radius: 12px !important;
  border: 1px solid __COLOR_BORDE__ !important;
  background: __COLOR_SUPERFICIE__ !important;
  font-weight: 700;
  color: __COLOR_PRIMARIO__ !important;
}
.btn:hover {
  background: __PRIM_005__ !important;
  border-color: __COLOR_PRIMARIO__ !important;
}

/* ============================================================
   ====== Cards ======
   ============================================================ */
.cardbox {
  background: __COLOR_SUPERFICIE__;
  border: 1px solid __COLOR_BORDE__;
  border-radius: 18px;
  box-shadow: 0 14px 34px __SHADOW_MED__;
  padding: 12px;
}

/* ============================================================
   ====== Layout spacing ======
   ============================================================ */
.row { margin-left: -10px; margin-right: -10px; }
.col-sm-6, .col-sm-12, .col-sm-9, .col-sm-3 { padding-left: 10px; padding-right: 10px; }

/* ============================================================
   ====== Header con logo ======
   ============================================================ */
.topbar{
  background:__COLOR_SUPERFICIE__;
  border:1px solid __COLOR_BORDE__;
  border-radius:18px;
  box-shadow:0 14px 34px __SHADOW_MED__;
  padding:16px 18px;
  margin-top: 6px;
  margin-bottom:14px;
  display:flex;
  align-items:center;
  justify-content:space-between;
  gap:14px;
}
.topbar-title{
  font-size:28px;
  font-weight:900;
  color:__COLOR_PRIMARIO__;
  line-height:1.12;
  flex: 1 1 auto;
  padding-top: 2px;
}
.topbar-logo{
  height:52px;
  max-width:240px;
  object-fit:contain;
  display:block;
  flex: 0 0 auto;
  margin-right: 2px;
}

/* ============================================================
   ====== Card header (editorial) ======
   ============================================================ */
.cardbox-header{
  padding:10px 12px 6px 12px;
  border-bottom:1px solid #edf0f7;
  margin:-12px -12px 10px -12px;
}
.cardbox-title{
  font-size:18px;
  font-weight:900;
  color:__COLOR_PRIMARIO__;
  line-height:1.15;
  margin:0;
}
.cardbox-subtitle{
  margin-top:4px;
  font-size:12px;
  color:__COLOR_TEXTO_SUAVE__;
}

/* ============================================================
   ====== Plotly ======
   ============================================================ */
.plot-container, .svg-container { width: 100% !important; }
.plotly .main-svg { overflow: visible !important; }
.plotly text{ font-weight:800 !important; }
.plotly .hoverlayer .hovertext{
  font-family: Arial, sans-serif !important;
  border-radius: 10px !important;
}

/* ============================================================
   ====== DataTable ======
   ============================================================ */
table.dataTable { border-collapse: collapse !important; table-layout: fixed !important; width: 100% !important; }
table.dataTable thead th{
  background:__COLOR_HEADER_TABLA__;
  color:__COLOR_PRIMARIO__;
  font-weight:800;
  border-bottom: 1px solid #dfe5f2 !important;
  border-right: 1px solid #dfe5f2 !important;
  text-align: center !important;
  vertical-align: middle !important;
}
table.dataTable tbody td{
  font-size:12px;
  color:__COLOR_TEXTO__;
  border-bottom: 1px solid #edf0f7 !important;
  border-right: 1px solid #edf0f7 !important;
  text-align: center !important;
  vertical-align: middle !important;
  white-space: normal !important;
  word-wrap: break-word !important;
  overflow-wrap: anywhere !important;
}
table.dataTable tbody tr:hover td{
  background: __COLOR_SUPERFICIE_2__ !important;
}

/* ============================================================
   ====== Toggle ======
   ============================================================ */
.toggle-row{
  display:flex;
  align-items:center;
  justify-content:space-between;
  gap:10px;
  margin-top: 10px;
  margin-bottom: 10px;
}
.toggle-label{
  font-size: 12px;
  color: __COLOR_TEXTO_SUAVE__;
  font-weight: 700;
  white-space: nowrap;
}
.switch {
  position: relative;
  display: inline-block;
  width: 52px;
  height: 28px;
  flex: 0 0 auto;
}
.switch input { display:none; }
.slider {
  position: absolute;
  cursor: pointer;
  top: 0; left: 0; right: 0; bottom: 0;
  background-color: __COLOR_BORDE__;
  transition: .25s;
  border-radius: 999px;
  border: 1px solid #dfe5f2;
}
.slider:before {
  position: absolute;
  content: \"\";
  height: 22px;
  width: 22px;
  left: 3px;
  bottom: 2.5px;
  background-color: white;
  transition: .25s;
  border-radius: 50%;
  box-shadow: 0 6px 14px rgba(0,0,0,0.12);
}
input:checked + .slider {
  background-color: __PRIM_020__;
  border-color: __PRIM_035__;
}
input:checked + .slider:before {
  transform: translateX(23px);
}

/* ============================================================
   ====== Diccionario ======
   ============================================================ */
.dicc-kv{
  display:grid;
  grid-template-columns: 92px 1fr;
  gap: 6px 10px;
  font-size: 12px;
  color: __COLOR_TEXTO__;
}
.dicc-k{
  color: __COLOR_TEXTO_SUAVE__;
  font-weight: 800;
}
.dicc-v{
  color: __COLOR_TEXTO__;
  font-weight: 600;
  word-break: break-word;
}

/* ============================================================
   ====== KPI BLOCK ======
   ============================================================ */
.kpi-block{
  display:flex;
  flex-direction:column;
  gap:10px;
  padding-bottom: 6px;
}

.kpi-block-title{
  font-size:14px;
  font-weight:900;
  color:__COLOR_PRIMARIO__;
  line-height:1.15;
  margin:0;
}

.kpi-block-subtitle{
  margin-top:4px;
  font-size:12px;
  color:__COLOR_TEXTO_SUAVE__;
}

.kpi-n-chip{
  width:100%;
  padding:18px 14px;
  border:1px solid #edf0f7;
  border-radius:16px;
  background:__COLOR_SUPERFICIE_2__;
  display:flex;
  align-items:center;
  justify-content:center;
}

.kpi-n-text{
  font-size:16px;
  font-weight:700;
  color:__COLOR_PRIMARIO__;
  letter-spacing:0.01em;
  line-height: 1.2;
  max-width: 100%;
  white-space: normal;
  word-break: break-word;
  width: 100% !important;
  text-align: center !important;
}

.kpi-grid{
  display:flex;
  gap:12px;
  width:100%;
  align-items:stretch;
}

.kpi-cell{
  flex:1 1 0;
  border:1px solid #edf0f7;
  border-radius:16px;
  padding:8px 8px 10px 8px;
  background:__COLOR_SUPERFICIE__;
  overflow: hidden;
  display: flex;
  flex-direction: column;
  align-items: stretch;
  justify-content: flex-start;
  width: 100%;
  box-sizing: border-box;
}

.kpi-legend{
  margin-top:8px !important;
  display:flex;
  flex-wrap:wrap;
  gap:4px 10px;
  justify-content:center !important;
  font-size:10px;
  color:__COLOR_TEXTO_SUAVE__;
  line-height:1.25 !important;
  white-space: normal !important;
  padding: 0 8px 10px 8px !important;
}

.kpi-legend-item{
  display:inline-flex;
  align-items:center;
  gap:6px;
}

.kpi-legend-swatch{
  display:inline-block;
  width:10px;
  height:10px;
  border-radius:3px;
}

.kpi-cell .plotly .gtitle,
.kpi-cell .plotly .g-gtitle,
.kpi-cell .plotly text{
  white-space: normal !important;
}

.kpi-cell .plotly{
  overflow: hidden !important;
}

.kpi-donut-title{
  font-size: 14px;
  font-weight: 900;
  color: __COLOR_PRIMARIO__;
  text-align: center;
  line-height: 1.15;
  margin: 4px 6px 2px 6px;
  white-space: normal;
  overflow-wrap: anywhere;
  word-break: break-word;
}

.kpi-profile-row{
  display:flex;
  gap:12px;
  align-items:stretch;
}

.kpi-n-card{
  flex: 0 0 42%;
  min-width: 320px;
  border:1px solid #edf0f7;
  border-radius:16px;
  background:__COLOR_SUPERFICIE__;
  padding:12px;
  display:flex;
  flex-direction:column;
  justify-content:center;
  align-items: center;
  text-align: center;
  overflow: hidden;
  width: 100% !important;
  max-width: 100% !important;
  box-sizing: border-box !important;
}

.kpi-n-card .kpi-block-title{
  margin:0 0 8px 0;
}

.kpi-donuts{
  flex: 1 1 auto;
  display:flex;
  gap:12px;
  align-items:stretch;
}

.kpi-donuts .kpi-cell{
  flex:1 1 0;
  min-width: 260px;
}

/* ============================================================
   ====== RESUMEN SECCIÓN ======
   ============================================================ */
.section-summary{
  display:flex;
  flex-direction:column;
  gap:10px;
}

.summary-row{
  border:1px solid #edf0f7;
  border-radius:16px;
  background:__COLOR_SUPERFICIE__;
  padding:10px 12px;
  box-shadow: 0 10px 22px __SHADOW_LOW__;
}

.summary-row-title{
  font-size:13px;
  font-weight:900;
  color:__COLOR_PRIMARIO__;
  line-height:1.2;
  margin:0 0 6px 0;
  overflow-wrap:anywhere;
}

.summary-row-subtitle{
  font-size:11px;
  color:__COLOR_TEXTO_SUAVE__;
  font-weight:700;
  margin:0 0 8px 0;
}

.summary-row-plot{
  height:84px;
  overflow:hidden;
}

.summary-row-plot:has(.sm-card-inner){
  height: auto !important;
  overflow: visible !important;
}

.sm-card-inner{
  display: flex;
  flex-direction: column;
  gap: 12px;
  height: auto !important;
  overflow: visible !important;
}

.sm-option-block{
  height: auto !important;
  overflow: visible !important;
}

/* ============================================================
   Sidebar KPI stack
   ============================================================ */
.kpi-sidebar-stack{
  display: flex;
  flex-direction: column;
  gap: 12px;
  align-items: stretch;
  width: 100%;
  box-sizing: border-box;
}

.kpi-sidebar-stack .kpi-profile-row{ display:block !important; }
.kpi-sidebar-stack .kpi-donuts{ display:block !important; }

.kpi-sidebar-stack .kpi-n-card{
  flex: 0 0 auto !important;
  min-width: 0 !important;
  width: 100% !important;
  max-width: 100% !important;
  box-sizing: border-box !important;
  align-items: center !important;
  justify-content: center !important;
  padding: 12px 12px !important;
  border-radius: 16px !important;
}

.kpi-sidebar-stack .kpi-n-chip{
  width: 100% !important;
  max-width: 100% !important;
  box-sizing: border-box !important;
  margin: 0 !important;
  justify-content: center !important;
}

.kpi-sidebar-stack .kpi-n-text{
  width: 100% !important;
  text-align: center !important;
  max-width: 100% !important;
  white-space: normal !important;
  word-break: break-word !important;
  font-weight: 900 !important;
  font-size: 18px !important;
}

.kpi-sidebar-stack .kpi-cell{
  width: 100% !important;
  max-width: 100% !important;
  min-width: 0 !important;
  box-sizing: border-box !important;
  margin: 0 !important;
  overflow: hidden !important;
  height: auto !important;
  min-height: 340px !important;
  padding-bottom: 14px !important;
}

.kpi-sidebar-stack .plotly.html-widget,
.kpi-sidebar-stack .plot-container,
.kpi-sidebar-stack .svg-container{
  width: 100% !important;
  max-width: 100% !important;
}

#kpi_plot_1, #kpi_plot_2{
  height: 220px !important;
  min-height: 220px !important;
}

#kpi_plot_1 .plot-container,
#kpi_plot_2 .plot-container,
#kpi_plot_1 .svg-container,
#kpi_plot_2 .svg-container{
  height: 220px !important;
  min-height: 220px !important;
}

.sidebarPanel .cardbox{
  overflow: hidden !important;
}

/* ============================================================
   NAVBAR COMO TOGGLE
   ============================================================ */
.navbar{
  background: transparent !important;
  border: 0 !important;
  box-shadow: none !important;
  margin-bottom: 14px !important;
  min-height: auto !important;
  padding-left: 0 !important;
  padding-right: 0 !important;
}

.navbar > .container-fluid{
  padding-left: 15px !important;
  padding-right: 15px !important;
}

.navbar .nav{
  margin-left: 0 !important;
  border-bottom: 0 !important;
  box-shadow: none !important;
  padding-bottom: 6px;
}

.navbar .nav.navbar-nav{
  position: relative;
  display: inline-flex !important;
  align-items: center;
  gap: 2px;
  padding: 4px;
  border-radius: 999px;
  background: __PRIM_085__;
  border: 1px solid __COLOR_BORDE__;
  box-shadow: 0 10px 22px __SHADOW_SOFT__;
  margin: 0 !important;
  padding-left: 10px;
}

.navbar .nav.navbar-nav::before{
  content: \"\";
  position: absolute;
  top: 3px;
  left: 3px;
  height: calc(100% - 6px);
  width: var(--pill-w, 0px);
  transform: translateX(var(--pill-x, 0px));
  border-radius: 999px;
  background: __COLOR_SUPERFICIE__;
  border: 1px solid __PILL_BORDER__;
  box-shadow: 0 10px 24px rgba(0,0,0,0.10);
  transition: transform 220ms cubic-bezier(.2,.9,.2,1),
              width 220ms cubic-bezier(.2,.9,.2,1);
  z-index: 0;
}

.navbar .nav > li{
  position: relative;
  z-index: 1;
}

.navbar .nav > li > a{
  background: transparent !important;
  border: 0 !important;
  box-shadow: none !important;
  color: __COLOR_PRIMARIO__ !important;
  font-weight: 900 !important;
  font-size: 13px;
  padding: 8px 14px !important;
  border-radius: 999px;
  line-height: 1;
  border-bottom: 0 !important;
}

.navbar .nav > li > a:hover{
  background: __PRIM_006__ !important;
}

.navbar .nav > li.active > a,
.navbar .nav > li.active > a:hover,
.navbar .nav > li.active > a:focus{
  background: transparent !important;
  border: 0 !important;
  box-shadow: none !important;
  outline: none !important;
  color: __COLOR_PRIMARIO__ !important;
}

.col-sm-3, .col-sm-9{
  padding-left: 10px;
  padding-right: 10px;
}
"

repl <- c(
  "__COLOR_PRIMARIO__"      = color_primario,
  "__COLOR_FONDO_APP__"     = color_fondo_app,
  "__COLOR_BORDE__"         = color_borde,
  "__COLOR_TEXTO__"         = color_texto,
  "__COLOR_TEXTO_SUAVE__"   = color_texto_suave,
  "__COLOR_SUPERFICIE__"    = color_superficie,
  "__COLOR_SUPERFICIE_2__"  = color_superficie_2,
  "__COLOR_HEADER_TABLA__"  = color_header_tabla,
  "__SHADOW_SOFT__"         = shadow_soft,
  "__SHADOW_MED__"          = shadow_med,
  "__SHADOW_LOW__"          = shadow_low,
  "__FOCUS_RING__"          = focus_ring,
  "__PRIM_005__"            = prim_005,
  "__PRIM_006__"            = prim_006,
  "__PRIM_020__"            = prim_020,
  "__PRIM_035__"            = prim_035,
  "__PRIM_085__"            = prim_085,
  "__PILL_BORDER__"         = pill_border
)

for (pat in names(repl)) {
  css <- gsub(pat, repl[[pat]], css, fixed = TRUE)
}

shiny::tags$style(shiny::HTML(css))
}

# -----------------------------------------------------------------------------
# JS de la app
# -----------------------------------------------------------------------------

reporte_interactivo_theme_js <- function() {
  shiny::tags$script(shiny::HTML("
(function(){
  function getNav(){
    return document.querySelector('.navbar .nav.navbar-nav') ||
           document.querySelector('.navbar .nav');
  }

  function getActiveLink(nav){
    if(!nav) return null;
    return nav.querySelector('li.active > a') ||
           nav.querySelector('li.active a') ||
           nav.querySelector('a[aria-selected=\"true\"]');
  }

  function updatePill(){
    var nav = getNav();
    if(!nav) return;
    var active = getActiveLink(nav);
    if(!active) return;

    var navRect = nav.getBoundingClientRect();
    var aRect   = active.getBoundingClientRect();

    var x = (aRect.left - navRect.left);
    var w = aRect.width;

    nav.style.setProperty('--pill-x', x + 'px');
    nav.style.setProperty('--pill-w', w + 'px');
  }

  function bindNavClicks(){
    document.addEventListener('click', function(e){
      var a = e.target && (e.target.closest ? e.target.closest('.navbar a') : null);
      if(!a) return;

      setTimeout(updatePill, 0);
      setTimeout(updatePill, 50);
      setTimeout(updatePill, 120);
    }, true);
  }

  function observeActiveChanges(){
    var nav = getNav();
    if(!nav || !window.MutationObserver) return;

    var obs = new MutationObserver(function(muts){
      var should = muts.some(function(m){
        return m.type === 'attributes' || m.type === 'childList';
      });
      if(should){
        window.requestAnimationFrame(updatePill);
      }
    });

    obs.observe(nav, {
      subtree: true,
      childList: true,
      attributes: true,
      attributeFilter: ['class','style','aria-selected']
    });
  }

  document.addEventListener('DOMContentLoaded', function(){
    setTimeout(updatePill, 80);
    setTimeout(updatePill, 200);
    bindNavClicks();
    observeActiveChanges();
  });

  document.addEventListener('shown.bs.tab', function(){
    setTimeout(updatePill, 0);
  });

  if(window.Shiny){
    document.addEventListener('shiny:value', function(){
      setTimeout(updatePill, 0);
      setTimeout(updatePill, 80);
    });

    document.addEventListener('shiny:connected', function(){
      setTimeout(updatePill, 120);
    });
  }

  window.addEventListener('resize', function(){
    updatePill();
  });
})();
"))
}
