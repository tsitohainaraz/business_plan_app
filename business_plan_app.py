import streamlit as st
import io
import os
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import openpyxl
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                              GradientFill)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference, LineChart
from openpyxl.chart.series import DataPoint

# ─── Page config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Créateur de Business Plan",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── Custom CSS ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');

  html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

  .main { background: #f8fafc; }

  .hero {
    background: linear-gradient(135deg, #0f2027, #203a43, #2c5364);
    border-radius: 20px;
    padding: 50px 40px;
    text-align: center;
    color: white;
    margin-bottom: 30px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.2);
  }
  .hero h1 { font-size: 2.8em; font-weight: 700; margin-bottom: 10px; }
  .hero p  { font-size: 1.15em; opacity: 0.85; }

  .step-card {
    background: white;
    border-radius: 16px;
    padding: 30px;
    box-shadow: 0 4px 20px rgba(0,0,0,0.07);
    border-left: 5px solid #2c5364;
    margin-bottom: 20px;
  }
  .step-card h2 { color: #2c5364; font-size: 1.5em; margin-bottom: 5px; }
  .step-card .sub { color: #64748b; font-size: 0.9em; margin-bottom: 20px; }

  .progress-bar {
    background: #e2e8f0;
    border-radius: 10px;
    height: 10px;
    margin-bottom: 30px;
    overflow: hidden;
  }
  .progress-fill {
    background: linear-gradient(90deg, #2c5364, #203a43);
    height: 100%;
    border-radius: 10px;
    transition: width .4s ease;
  }

  .step-indicator {
    display: flex;
    justify-content: center;
    gap: 10px;
    margin-bottom: 30px;
    flex-wrap: wrap;
  }
  .step-dot {
    width: 36px; height: 36px;
    border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    font-size: 0.8em; font-weight: 600;
    border: 2px solid #cbd5e1;
    color: #94a3b8;
  }
  .step-dot.active  { background: #2c5364; color: white; border-color: #2c5364; }
  .step-dot.done    { background: #22c55e; color: white; border-color: #22c55e; }

  .info-box {
    background: #eff6ff;
    border: 1px solid #bfdbfe;
    border-radius: 10px;
    padding: 12px 16px;
    font-size: 0.88em;
    color: #1e40af;
    margin-bottom: 16px;
  }

  .success-box {
    background: linear-gradient(135deg, #f0fdf4, #dcfce7);
    border: 1px solid #86efac;
    border-radius: 16px;
    padding: 30px;
    text-align: center;
    margin-bottom: 20px;
  }
  .success-box h2 { color: #15803d; font-size: 2em; }
  .success-box p  { color: #166534; }

  div[data-testid="stButton"] > button {
    border-radius: 10px !important;
    font-weight: 600 !important;
    padding: 10px 28px !important;
    transition: all .2s !important;
  }
  div[data-testid="stButton"] > button:first-child {
    background: linear-gradient(135deg, #2c5364, #203a43) !important;
    color: white !important;
    border: none !important;
  }
  div[data-testid="stButton"] > button:first-child:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(44,83,100,0.4) !important;
  }

  div[data-testid="stDownloadButton"] > button {
    border-radius: 10px !important;
    font-weight: 600 !important;
    padding: 12px 30px !important;
    width: 100%;
  }
</style>
""", unsafe_allow_html=True)

# ─── Session state init ──────────────────────────────────────────────────────
if "step" not in st.session_state:
    st.session_state.step = 0

if "data" not in st.session_state:
    st.session_state.data = {}

STEPS = [
    "Accueil",
    "Votre Idée",
    "Votre Projet",
    "Étude de Marché",
    "Produit / Service",
    "Équipe",
    "Financement",
    "Projections",
    "Statut Juridique",
    "Récapitulatif",
]

TOTAL_STEPS = len(STEPS)

# ─── Helpers ─────────────────────────────────────────────────────────────────
def next_step():
    st.session_state.step += 1

def prev_step():
    st.session_state.step -= 1

def nav_buttons(show_prev=True):
    c1, c2, c3 = st.columns([1, 3, 1])
    if show_prev:
        with c1:
            if st.button("◀ Précédent"):
                prev_step()
    with c3:
        if st.button("Suivant ▶", key=f"next_{st.session_state.step}"):
            next_step()

def render_progress():
    step = st.session_state.step
    pct = int((step / (TOTAL_STEPS - 1)) * 100)
    st.markdown(f"""
    <div class="progress-bar">
      <div class="progress-fill" style="width:{pct}%"></div>
    </div>""", unsafe_allow_html=True)

    dots = ""
    for i, name in enumerate(STEPS):
        if i < step:
            cls = "done"
            icon = "✓"
        elif i == step:
            cls = "active"
            icon = str(i + 1)
        else:
            cls = ""
            icon = str(i + 1)
        dots += f'<div class="step-dot {cls}" title="{name}">{icon}</div>'
    st.markdown(f'<div class="step-indicator">{dots}</div>', unsafe_allow_html=True)

def field(label, key, placeholder="", multiline=False, help=""):
    val = st.session_state.data.get(key, "")
    if multiline:
        result = st.text_area(label, value=val, placeholder=placeholder, help=help, height=120)
    else:
        result = st.text_input(label, value=val, placeholder=placeholder, help=help)
    st.session_state.data[key] = result
    return result

def num_field(label, key, min_val=0, max_val=10_000_000, step=1000, default=0, fmt="%d"):
    val = st.session_state.data.get(key, default)
    result = st.number_input(label, min_value=min_val, max_value=max_val,
                             value=int(val), step=step, format=fmt)
    st.session_state.data[key] = result
    return result

def select_field(label, key, options, help=""):
    val = st.session_state.data.get(key, options[0])
    idx = options.index(val) if val in options else 0
    result = st.selectbox(label, options, index=idx, help=help)
    st.session_state.data[key] = result
    return result

def radio_field(label, key, options):
    val = st.session_state.data.get(key, options[0])
    idx = options.index(val) if val in options else 0
    result = st.radio(label, options, index=idx, horizontal=True)
    st.session_state.data[key] = result
    return result

def info(txt):
    st.markdown(f'<div class="info-box">💡 {txt}</div>', unsafe_allow_html=True)

# ─── STEPS ───────────────────────────────────────────────────────────────────

def step_0_welcome():
    st.markdown("""
    <div class="hero">
      <h1>🚀 Créateur de Business Plan</h1>
      <p>Répondez aux questions étape par étape et obtenez automatiquement<br>
         un <strong>Business Plan Word</strong> et un <strong>Plan Financier Excel</strong> professionnels.</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        <div style="background:white;border-radius:12px;padding:20px;text-align:center;box-shadow:0 4px 15px rgba(0,0,0,0.06)">
        <div style="font-size:2.2em">📋</div>
        <strong>9 sections</strong><br>
        <small>de questions guidées</small>
        </div>""", unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div style="background:white;border-radius:12px;padding:20px;text-align:center;box-shadow:0 4px 15px rgba(0,0,0,0.06)">
        <div style="font-size:2.2em">📄</div>
        <strong>Document Word</strong><br>
        <small>Business plan complet</small>
        </div>""", unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div style="background:white;border-radius:12px;padding:20px;text-align:center;box-shadow:0 4px 15px rgba(0,0,0,0.06)">
        <div style="font-size:2.2em">📊</div>
        <strong>Fichier Excel</strong><br>
        <small>Projections financières</small>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    _, c, _ = st.columns([2, 1, 2])
    with c:
        if st.button("Commencer ▶", key="start"):
            next_step()


def step_1_idea():
    st.markdown("""<div class="step-card">
    <h2>💡 Étape 1 — Votre Idée</h2>
    <div class="sub">Définissez clairement le problème que vous résolvez et votre solution.</div>
    </div>""", unsafe_allow_html=True)

    info("Une bonne idée répond à un besoin du marché non assouvi. Soyez précis et concis.")

    field("Quel problème ou besoin résolvez-vous ?", "probleme",
          "Ex : Les PME perdent du temps à gérer leurs factures manuellement…", multiline=True)
    field("Quelle est votre solution concrète ?", "solution",
          "Ex : Une application SaaS qui automatise la facturation en 3 clics…", multiline=True)

    col1, col2 = st.columns(2)
    with col1:
        select_field("Quel est votre domaine d'activité ?", "domaine",
                     ["Tech / Numérique", "Commerce / Retail", "Services aux entreprises",
                      "Services aux particuliers", "Restauration / Food", "Artisanat",
                      "Santé / Bien-être", "Éducation / Formation", "Immobilier", "Autre"])
    with col2:
        radio_field("Votre cible principale ?", "cible", ["B2C (Particuliers)", "B2B (Entreprises)", "Les deux"])

    field("Avez-vous identifié des concurrents ? Lesquels ?", "concurrents",
          "Ex : Freebe, Indy, Pennylane…", multiline=True)
    field("Quelle est votre différence par rapport à eux ?", "differenciateur",
          "Ex : Plus simple, moins cher, spécialisé pour les artisans…", multiline=True)

    nav_buttons(show_prev=False)


def step_2_project():
    st.markdown("""<div class="step-card">
    <h2>🏢 Étape 2 — Votre Projet</h2>
    <div class="sub">Présentez votre projet et vos motivations.</div>
    </div>""", unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        field("Nom de votre entreprise / projet", "nom_entreprise", "Ex : MonStartup SAS")
        field("Prénom et nom du porteur de projet", "porteur", "Ex : Marie Dupont")
    with col2:
        field("Ville de création", "ville", "Ex : Paris")
        field("Date de création prévue", "date_creation", "Ex : Septembre 2026")

    field("Décrivez votre projet en 2-3 phrases (pitch)", "pitch",
          "Ex : Nous aidons les artisans à gérer leur comptabilité en 5 minutes par semaine…", multiline=True)

    radio_field("Vous lancez-vous seul ou à plusieurs ?", "seul_ou_plusieurs",
                ["Seul(e)", "À 2", "À 3 et plus"])

    field("Pourquoi ce projet ? Quelle est votre motivation personnelle ?", "motivation",
          "Ex : Ancien artisan, j'ai vécu cette douleur moi-même…", multiline=True)

    field("Quelles compétences / expériences avez-vous dans ce domaine ?", "competences_porteur",
          "Ex : 5 ans d'expérience en comptabilité, diplôme d'expert-comptable…", multiline=True)

    nav_buttons()


def step_3_market():
    st.markdown("""<div class="step-card">
    <h2>📈 Étape 3 — Étude de Marché</h2>
    <div class="sub">Prouvez qu'il existe une vraie opportunité sur votre marché.</div>
    </div>""", unsafe_allow_html=True)

    info("L'étude de marché est essentielle pour convaincre vos investisseurs et valider votre concept.")

    col1, col2 = st.columns(2)
    with col1:
        field("Taille estimée de votre marché", "taille_marche",
              "Ex : 500 000 PME en France, marché de 2Md€")
        field("Tendances de votre marché", "tendances",
              "Ex : Marché en croissance de 15%/an, digitalisation accélérée post-Covid")
    with col2:
        select_field("Maturité du marché", "maturite_marche",
                     ["Émergent (nouveau marché)", "En croissance", "Mature / Établi", "En déclin"])
        field("Nombre de clients potentiels visés la 1ère année", "nb_clients_an1",
              "Ex : 50 clients PME")

    field("Qui sont vos 3 principaux concurrents directs ?", "concurrents_details",
          "Nom — Part de marché — Prix — Points forts / faibles…", multiline=True)
    field("Qui sont vos concurrents indirects ?", "concurrents_indirects",
          "Ex : Excel, experts-comptables traditionnels…", multiline=True)
    field("Quelle est votre stratégie pour acquérir vos premiers clients ?", "strategie_acquisition",
          "Ex : Réseaux sociaux, prospection directe, partenariats, SEO…", multiline=True)

    nav_buttons()


def step_4_product():
    st.markdown("""<div class="step-card">
    <h2>📦 Étape 4 — Produit / Service</h2>
    <div class="sub">Décrivez ce que vous vendez, comment et à quel prix.</div>
    </div>""", unsafe_allow_html=True)

    field("Description détaillée de votre produit ou service", "produit_detail",
          "Fonctionnalités, avantages, comment ça marche…", multiline=True)
    field("Comment est-il produit ou livré ?", "production",
          "Ex : Développé en interne, sous-traité, SaaS accessible en ligne…", multiline=True)

    col1, col2 = st.columns(2)
    with col1:
        select_field("Modèle économique principal", "modele_eco",
                     ["Abonnement mensuel / annuel", "Vente à l'unité",
                      "Commission / % du CA", "Freemium", "Sur mesure / devis",
                      "Marketplace", "Autre"])
        field("Prix de vente unitaire ou mensuel (€)", "prix_vente",
              "Ex : 49€/mois ou 299€ à l'unité")
    with col2:
        field("Coût de revient par unité/mois (€)", "cout_revient",
              "Ex : 8€/mois de coûts variables")
        field("Marge brute estimée (%)", "marge_brute",
              "Ex : 83%")

    field("Quelles sont vos ressources clés ?", "ressources_cles",
          "Ex : Technologie propriétaire, réseau de partenaires, marque…", multiline=True)
    field("Quels sont vos fournisseurs principaux ?", "fournisseurs",
          "Ex : AWS pour l'hébergement, Stripe pour les paiements…")

    nav_buttons()


def step_5_team():
    st.markdown("""<div class="step-card">
    <h2>👥 Étape 5 — L'Équipe</h2>
    <div class="sub">Présentez les fondateurs et les compétences clés du projet.</div>
    </div>""", unsafe_allow_html=True)

    info("Les investisseurs investissent d'abord dans des personnes. Mettez en valeur votre équipe !")

    nb_str = st.session_state.data.get("seul_ou_plusieurs", "Seul(e)")
    nb_fondateurs = 1 if "Seul" in nb_str else (2 if "À 2" in nb_str else 3)

    for i in range(1, nb_fondateurs + 1):
        st.markdown(f"**👤 Fondateur {i}**")
        c1, c2 = st.columns(2)
        with c1:
            field(f"Nom complet", f"fondateur_{i}_nom", "Ex : Marie Dupont")
            field(f"Rôle dans l'entreprise", f"fondateur_{i}_role",
                  "Ex : CEO / Développement commercial")
        with c2:
            field(f"Formation / Diplôme", f"fondateur_{i}_formation",
                  "Ex : Master Finance, HEC Paris")
            field(f"Expérience pertinente", f"fondateur_{i}_experience",
                  "Ex : 7 ans en comptabilité PME", multiline=False)

        field(f"Compétences clés apportées au projet", f"fondateur_{i}_competences",
              "Ex : Tech, commercial, finance, opérations…")
        st.markdown("---")

    field("Avez-vous des conseillers, mentors ou investisseurs-clés ?", "advisors",
          "Ex : Jean Martin (ex-DG BNP), membre du board…")
    field("Avez-vous besoin de recruter ? Pour quels postes ?", "recrutements",
          "Ex : 1 développeur back-end dès le 3ème mois")

    nav_buttons()


def step_6_financing():
    st.markdown("""<div class="step-card">
    <h2>💰 Étape 6 — Financement</h2>
    <div class="sub">Comment allez-vous financer votre projet ?</div>
    </div>""", unsafe_allow_html=True)

    info("Depuis janvier 2026, l'ACRE n'est plus accordée automatiquement. Pensez à faire la demande auprès de l'URSSAF.")

    col1, col2 = st.columns(2)
    with col1:
        num_field("Apports personnels (€)", "apports_perso", step=500)
        num_field("Love money (famille/amis) (€)", "love_money", step=500)
    with col2:
        num_field("Emprunt bancaire souhaité (€)", "emprunt", step=1000)
        num_field("Financement participatif / subventions (€)", "subventions", step=500)

    radio_field("Êtes-vous demandeur d'emploi (ARE) ?", "are",
                ["Non", "Oui — je maintiens l'ARE", "Oui — je prends l'ARCE"])
    radio_field("Pensez-vous être éligible à l'ACRE ?", "acre",
                ["Je ne sais pas", "Oui", "Non"])
    field("Autres sources de financement ?", "autres_financements",
          "Ex : Business angels, aides régionales, prix startup…")

    num_field("Budget de lancement total estimé (€)", "budget_lancement", step=1000)
    field("À quoi servira principalement ce budget ?", "usage_budget",
          "Ex : 40% tech, 30% marketing, 20% salaires, 10% frais divers", multiline=True)

    nav_buttons()


def step_7_projections():
    st.markdown("""<div class="step-card">
    <h2>📊 Étape 7 — Projections Financières</h2>
    <div class="sub">Estimez vos revenus et charges sur 3 ans.</div>
    </div>""", unsafe_allow_html=True)

    info("Ces projections alimenteront automatiquement votre fichier Excel avec graphiques.")

    st.markdown("**🎯 Chiffre d'affaires prévisionnel**")
    c1, c2, c3 = st.columns(3)
    with c1:
        num_field("CA Année 1 (€)", "ca_an1", max_val=50_000_000, step=5000)
    with c2:
        num_field("CA Année 2 (€)", "ca_an2", max_val=50_000_000, step=5000)
    with c3:
        num_field("CA Année 3 (€)", "ca_an3", max_val=50_000_000, step=5000)

    st.markdown("**📉 Charges mensuelles fixes**")
    c1, c2, c3 = st.columns(3)
    with c1:
        num_field("Salaires / rémunération (€/mois)", "charges_salaires", step=200)
        num_field("Loyer / siège social (€/mois)", "charges_loyer", step=100)
    with c2:
        num_field("Logiciels / abonnements (€/mois)", "charges_logiciels", step=50)
        num_field("Marketing / publicité (€/mois)", "charges_marketing", step=100)
    with c3:
        num_field("Comptabilité / juridique (€/mois)", "charges_compta", step=50)
        num_field("Autres charges fixes (€/mois)", "autres_charges", step=100)

    st.markdown("**📈 Charges variables (% du CA)**")
    c1, c2 = st.columns(2)
    with c1:
        num_field("Coût des ventes / COGS (%)", "cogs_pct", min_val=0, max_val=100, step=1, default=20, fmt="%d")
    with c2:
        num_field("Commissions / frais paiement (%)", "commission_pct", min_val=0, max_val=30, step=1, default=3, fmt="%d")

    nav_buttons()


def step_8_legal():
    st.markdown("""<div class="step-card">
    <h2>⚖️ Étape 8 — Statut Juridique</h2>
    <div class="sub">Choisissez la structure adaptée à votre situation.</div>
    </div>""", unsafe_allow_html=True)

    statut_options = [
        "Micro-entreprise (Auto-entrepreneur)",
        "SASU — Société par Actions Simplifiée Unipersonnelle",
        "EURL — Entreprise Unipersonnelle à Responsabilité Limitée",
        "SAS — Société par Actions Simplifiée",
        "SARL — Société à Responsabilité Limitée",
        "Je ne sais pas encore",
    ]
    select_field("Statut juridique envisagé", "statut_juridique", statut_options,
                 help="Basé sur votre situation et le guide Legalstart 2026")

    c1, c2 = st.columns(2)
    with c1:
        num_field("Capital social envisagé (€)", "capital_social", step=100, default=1000)
        field("Adresse du siège social", "siege_social",
              "Ex : 15 rue de Rivoli, 75001 Paris")
    with c2:
        select_field("Type de domiciliation", "domiciliation",
                     ["Domicile personnel", "Local commercial",
                      "Société de domiciliation", "Pépinière d'entreprises"])
        field("Régime fiscal choisi", "regime_fiscal",
              "Ex : IS (Impôt sur les Sociétés), IR (Impôt sur le Revenu)")

    field("Avez-vous des questions ou précisions sur votre statut ?", "questions_statut",
          "Ex : J'hésite entre SASU et micro-entreprise car…", multiline=True)

    nav_buttons()


def step_9_recap():
    d = st.session_state.data

    st.markdown("""<div class="success-box">
    <h2>🎉 Votre Business Plan est prêt !</h2>
    <p>Téléchargez maintenant votre <strong>document Word</strong> et votre <strong>fichier Excel</strong> professionnels.</p>
    </div>""", unsafe_allow_html=True)

    # Summary
    with st.expander("📋 Récapitulatif de vos réponses", expanded=False):
        cols = st.columns(2)
        items = list(d.items())
        half = len(items) // 2
        for i, (k, v) in enumerate(items):
            with cols[0] if i < half else cols[1]:
                if v:
                    st.markdown(f"**{k}**: {v}")

    st.markdown("---")
    c1, c2 = st.columns(2)

    with c1:
        st.markdown("### 📄 Business Plan Word")
        st.caption("Document professionnel avec mise en page élaborée, sections détaillées et synthèse executive.")
        docx_bytes = generate_word(d)
        st.download_button(
            label="⬇️ Télécharger le Business Plan (.docx)",
            data=docx_bytes,
            file_name=f"BusinessPlan_{d.get('nom_entreprise','Projet').replace(' ','_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    with c2:
        st.markdown("### 📊 Plan Financier Excel")
        st.caption("Projections sur 3 ans, graphiques de CA et résultat, tableaux des charges et financement.")
        xlsx_bytes = generate_excel(d)
        st.download_button(
            label="⬇️ Télécharger le Plan Financier (.xlsx)",
            data=xlsx_bytes,
            file_name=f"PlanFinancier_{d.get('nom_entreprise','Projet').replace(' ','_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.markdown("---")
    if st.button("🔄 Recommencer un nouveau projet"):
        st.session_state.step = 0
        st.session_state.data = {}
        st.rerun()


# ─── WORD GENERATOR ──────────────────────────────────────────────────────────
def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def add_colored_heading(doc, text, level=1, color="2C5364"):
    p = doc.add_heading(text, level=level)
    for run in p.runs:
        run.font.color.rgb = RGBColor.from_string(color)
    return p

def add_info_box(doc, text, bg="EFF6FF", border_color="2563EB"):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    set_cell_bg(cell, bg)
    cell.width = Inches(6)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.left_indent = Inches(0.15)
    run = p.add_run(text)
    run.font.size = Pt(10)
    run.font.italic = True
    run.font.color.rgb = RGBColor.from_string("1E40AF")
    doc.add_paragraph()

def add_kv_table(doc, rows_data, header_color="2C5364"):
    table = doc.add_table(rows=len(rows_data), cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (label, value) in enumerate(rows_data):
        row = table.rows[i]
        # Label cell
        lc = row.cells[0]
        lc.width = Inches(2.2)
        set_cell_bg(lc, "F1F5F9" if i % 2 == 0 else "E2E8F0")
        p = lc.paragraphs[0]
        run = p.add_run(label)
        run.bold = True
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor.from_string(header_color)
        # Value cell
        vc = row.cells[1]
        vc.width = Inches(4.1)
        set_cell_bg(vc, "FFFFFF" if i % 2 == 0 else "F8FAFC")
        vp = vc.paragraphs[0]
        vr = vp.add_run(str(value or "—"))
        vr.font.size = Pt(10)
    doc.add_paragraph()


def generate_word(d):
    doc = Document()

    # Page margins
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(3)
        section.right_margin = Cm(2.5)

    # Document default styles
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(11)

    # ── COVER PAGE ─────────────────────────────────────────────────────────
    # Top accent line
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    run = p.add_run("─" * 80)
    run.font.color.rgb = RGBColor.from_string("2C5364")
    run.font.size = Pt(8)

    doc.add_paragraph()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("BUSINESS PLAN")
    run.font.size = Pt(36)
    run.font.bold = True
    run.font.color.rgb = RGBColor.from_string("0F2027")

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(d.get("nom_entreprise", "Mon Projet") or "Mon Projet")
    run.font.size = Pt(26)
    run.font.bold = True
    run.font.color.rgb = RGBColor.from_string("2C5364")

    doc.add_paragraph()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(d.get("pitch", "") or "")
    run.font.size = Pt(13)
    run.font.italic = True
    run.font.color.rgb = RGBColor.from_string("475569")

    doc.add_paragraph()
    doc.add_paragraph()

    # Meta info box
    meta = [
        ("Porteur de projet", d.get("porteur", "—")),
        ("Ville", d.get("ville", "—")),
        ("Date de création prévue", d.get("date_creation", "—")),
        ("Statut juridique envisagé", d.get("statut_juridique", "—")),
        ("Domaine", d.get("domaine", "—")),
        ("Date du document", datetime.now().strftime("%d/%m/%Y")),
    ]
    table = doc.add_table(rows=len(meta), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for i, (k, v) in enumerate(meta):
        r = table.rows[i]
        lc = r.cells[0]
        lc.width = Inches(2.5)
        set_cell_bg(lc, "2C5364")
        lp = lc.paragraphs[0]
        lr = lp.add_run(k)
        lr.bold = True
        lr.font.color.rgb = RGBColor.from_string("FFFFFF")
        lr.font.size = Pt(10)

        vc = r.cells[1]
        vc.width = Inches(3.8)
        set_cell_bg(vc, "F8FAFC")
        vp = vc.paragraphs[0]
        vr = vp.add_run(str(v))
        vr.font.size = Pt(10)

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("─" * 80)
    run.font.color.rgb = RGBColor.from_string("2C5364")
    run.font.size = Pt(8)

    doc.add_page_break()

    # ── SOMMAIRE ────────────────────────────────────────────────────────────
    add_colored_heading(doc, "Sommaire", 1)
    sections_toc = [
        "1. Executive Summary",
        "2. L'Idée et le Problème Résolu",
        "3. Étude de Marché",
        "4. Produit / Service",
        "5. Équipe",
        "6. Stratégie Commerciale",
        "7. Plan de Financement",
        "8. Projections Financières",
        "9. Statut Juridique",
        "10. Conclusion",
    ]
    for s in sections_toc:
        p = doc.add_paragraph(s, style="List Bullet")
        p.runs[0].font.size = Pt(11)

    doc.add_page_break()

    # ── 1. EXECUTIVE SUMMARY ────────────────────────────────────────────────
    add_colored_heading(doc, "1. Executive Summary", 1)
    add_info_box(doc, "L'Executive Summary doit convaincre en 2 minutes. Il résume l'essentiel de votre projet.")

    summary_rows = [
        ("Nom du projet", d.get("nom_entreprise", "—")),
        ("Porteur(s)", d.get("porteur", "—")),
        ("Problème résolu", d.get("probleme", "—")),
        ("Solution proposée", d.get("solution", "—")),
        ("Cible", d.get("cible", "—")),
        ("Modèle économique", d.get("modele_eco", "—")),
        ("Marché visé", d.get("taille_marche", "—")),
        ("Financement recherché", f"{int(d.get('emprunt', 0) or 0):,} € d'emprunt + {int(d.get('subventions', 0) or 0):,} € de subventions"),
        ("CA Année 1 / 3", f"{int(d.get('ca_an1', 0) or 0):,} € → {int(d.get('ca_an3', 0) or 0):,} €"),
        ("Statut envisagé", d.get("statut_juridique", "—")),
    ]
    add_kv_table(doc, summary_rows)

    # ── 2. L'IDÉE ────────────────────────────────────────────────────────────
    add_colored_heading(doc, "2. L'Idée et le Problème Résolu", 1)

    add_colored_heading(doc, "2.1 Le problème identifié", 2, "203A43")
    p = doc.add_paragraph(d.get("probleme", "Non renseigné") or "Non renseigné")
    p.runs[0].font.size = Pt(11)

    add_colored_heading(doc, "2.2 Notre solution", 2, "203A43")
    p = doc.add_paragraph(d.get("solution", "Non renseigné") or "Non renseigné")
    p.runs[0].font.size = Pt(11)

    add_colored_heading(doc, "2.3 Notre différenciateur", 2, "203A43")
    p = doc.add_paragraph(d.get("differenciateur", "Non renseigné") or "Non renseigné")
    p.runs[0].font.size = Pt(11)

    idea_rows = [
        ("Domaine d'activité", d.get("domaine", "—")),
        ("Cible principale", d.get("cible", "—")),
        ("Concurrents identifiés", d.get("concurrents", "—")),
    ]
    add_kv_table(doc, idea_rows)

    # ── 3. ÉTUDE DE MARCHÉ ──────────────────────────────────────────────────
    add_colored_heading(doc, "3. Étude de Marché", 1)

    add_colored_heading(doc, "3.1 Taille et maturité du marché", 2, "203A43")
    market_rows = [
        ("Taille du marché", d.get("taille_marche", "—")),
        ("Maturité", d.get("maturite_marche", "—")),
        ("Tendances", d.get("tendances", "—")),
        ("Clients visés — Année 1", d.get("nb_clients_an1", "—")),
    ]
    add_kv_table(doc, market_rows)

    add_colored_heading(doc, "3.2 Analyse concurrentielle", 2, "203A43")
    p = doc.add_paragraph("Concurrents directs :")
    p.runs[0].bold = True
    p = doc.add_paragraph(d.get("concurrents_details", "Non renseigné") or "Non renseigné")
    p.runs[0].font.size = Pt(11)

    p = doc.add_paragraph("Concurrents indirects :")
    p.runs[0].bold = True
    p = doc.add_paragraph(d.get("concurrents_indirects", "Non renseigné") or "Non renseigné")
    p.runs[0].font.size = Pt(11)

    add_colored_heading(doc, "3.3 Stratégie d'acquisition clients", 2, "203A43")
    p = doc.add_paragraph(d.get("strategie_acquisition", "Non renseigné") or "Non renseigné")
    p.runs[0].font.size = Pt(11)

    # ── 4. PRODUIT / SERVICE ─────────────────────────────────────────────────
    add_colored_heading(doc, "4. Produit / Service", 1)

    add_colored_heading(doc, "4.1 Description", 2, "203A43")
    p = doc.add_paragraph(d.get("produit_detail", "Non renseigné") or "Non renseigné")
    p.runs[0].font.size = Pt(11)

    product_rows = [
        ("Production / Livraison", d.get("production", "—")),
        ("Modèle économique", d.get("modele_eco", "—")),
        ("Prix de vente", d.get("prix_vente", "—")),
        ("Coût de revient", d.get("cout_revient", "—")),
        ("Marge brute estimée", d.get("marge_brute", "—")),
        ("Ressources clés", d.get("ressources_cles", "—")),
        ("Fournisseurs", d.get("fournisseurs", "—")),
    ]
    add_kv_table(doc, product_rows)

    # ── 5. ÉQUIPE ────────────────────────────────────────────────────────────
    add_colored_heading(doc, "5. L'Équipe", 1)

    nb_str = d.get("seul_ou_plusieurs", "Seul(e)")
    nb = 1 if "Seul" in nb_str else (2 if "À 2" in nb_str else 3)
    for i in range(1, nb + 1):
        add_colored_heading(doc, f"5.{i} Fondateur {i} — {d.get(f'fondateur_{i}_nom', '—')}", 2, "203A43")
        team_rows = [
            ("Rôle", d.get(f"fondateur_{i}_role", "—")),
            ("Formation", d.get(f"fondateur_{i}_formation", "—")),
            ("Expérience", d.get(f"fondateur_{i}_experience", "—")),
            ("Compétences clés", d.get(f"fondateur_{i}_competences", "—")),
        ]
        add_kv_table(doc, team_rows)

    if d.get("advisors"):
        add_colored_heading(doc, "5.x Conseillers & Advisors", 2, "203A43")
        doc.add_paragraph(d["advisors"])

    if d.get("recrutements"):
        add_colored_heading(doc, "5.x Recrutements prévus", 2, "203A43")
        doc.add_paragraph(d["recrutements"])

    # ── 6. STRATÉGIE COMMERCIALE ─────────────────────────────────────────────
    add_colored_heading(doc, "6. Stratégie Commerciale", 1)
    doc.add_paragraph(d.get("strategie_acquisition", "Non renseigné") or "Non renseigné")
    doc.add_paragraph()
    p = doc.add_paragraph("Motivation du porteur de projet :")
    p.runs[0].bold = True
    doc.add_paragraph(d.get("motivation", "Non renseigné") or "Non renseigné")

    # ── 7. PLAN DE FINANCEMENT ───────────────────────────────────────────────
    add_colored_heading(doc, "7. Plan de Financement", 1)

    total_financement = (int(d.get("apports_perso", 0) or 0)
                         + int(d.get("love_money", 0) or 0)
                         + int(d.get("emprunt", 0) or 0)
                         + int(d.get("subventions", 0) or 0))

    fin_rows = [
        ("Apports personnels", f"{int(d.get('apports_perso', 0) or 0):,} €"),
        ("Love money", f"{int(d.get('love_money', 0) or 0):,} €"),
        ("Emprunt bancaire", f"{int(d.get('emprunt', 0) or 0):,} €"),
        ("Subventions / financement participatif", f"{int(d.get('subventions', 0) or 0):,} €"),
        ("Autres financements", d.get("autres_financements", "—")),
        ("TOTAL FINANCEMENT", f"{total_financement:,} €"),
        ("Budget de lancement estimé", f"{int(d.get('budget_lancement', 0) or 0):,} €"),
        ("Usage du budget", d.get("usage_budget", "—")),
        ("ARE / ARCE", d.get("are", "—")),
        ("ACRE", d.get("acre", "—")),
    ]
    add_kv_table(doc, fin_rows)

    # ── 8. PROJECTIONS FINANCIÈRES ───────────────────────────────────────────
    add_colored_heading(doc, "8. Projections Financières", 1)
    add_info_box(doc, "Les projections détaillées (mensuel, trimestriel) sont disponibles dans le fichier Excel joint.")

    ca1 = int(d.get("ca_an1", 0) or 0)
    ca2 = int(d.get("ca_an2", 0) or 0)
    ca3 = int(d.get("ca_an3", 0) or 0)
    charges_fixes_mois = sum([
        int(d.get("charges_salaires", 0) or 0),
        int(d.get("charges_loyer", 0) or 0),
        int(d.get("charges_logiciels", 0) or 0),
        int(d.get("charges_marketing", 0) or 0),
        int(d.get("charges_compta", 0) or 0),
        int(d.get("autres_charges", 0) or 0),
    ])
    charges_fixes_an = charges_fixes_mois * 12
    cogs_pct = int(d.get("cogs_pct", 20) or 20) / 100
    comm_pct = int(d.get("commission_pct", 3) or 3) / 100

    def resultat(ca):
        charges_var = ca * (cogs_pct + comm_pct)
        return ca - charges_var - charges_fixes_an

    r1, r2, r3 = resultat(ca1), resultat(ca2), resultat(ca3)

    # Summary table
    proj_table = doc.add_table(rows=5, cols=4)
    proj_table.style = "Table Grid"
    proj_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    headers = ["", "Année 1", "Année 2", "Année 3"]
    for ci, h in enumerate(headers):
        cell = proj_table.rows[0].cells[ci]
        set_cell_bg(cell, "2C5364")
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.bold = True
        r.font.color.rgb = RGBColor.from_string("FFFFFF")
        r.font.size = Pt(10)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    data_rows = [
        ("Chiffre d'affaires", f"{ca1:,} €", f"{ca2:,} €", f"{ca3:,} €"),
        ("Charges fixes", f"{charges_fixes_an:,} €", f"{charges_fixes_an:,} €", f"{charges_fixes_an:,} €"),
        ("Charges variables", f"{int(ca1 * (cogs_pct + comm_pct)):,} €",
         f"{int(ca2 * (cogs_pct + comm_pct)):,} €",
         f"{int(ca3 * (cogs_pct + comm_pct)):,} €"),
        ("Résultat estimé",
         f"{int(r1):,} €", f"{int(r2):,} €", f"{int(r3):,} €"),
    ]

    for ri, (label, v1, v2, v3) in enumerate(data_rows, start=1):
        row = proj_table.rows[ri]
        bg = "F1F5F9" if ri % 2 == 0 else "FFFFFF"
        for ci, val in enumerate([label, v1, v2, v3]):
            cell = row.cells[ci]
            set_cell_bg(cell, bg)
            p = cell.paragraphs[0]
            run = p.add_run(val)
            run.font.size = Pt(10)
            if ci == 0:
                run.bold = True
                run.font.color.rgb = RGBColor.from_string("2C5364")
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if ci > 0 else WD_ALIGN_PARAGRAPH.LEFT

    doc.add_paragraph()

    # ── 9. STATUT JURIDIQUE ──────────────────────────────────────────────────
    add_colored_heading(doc, "9. Statut Juridique", 1)

    legal_rows = [
        ("Statut choisi", d.get("statut_juridique", "—")),
        ("Fondateurs", d.get("seul_ou_plusieurs", "—")),
        ("Capital social", f"{int(d.get('capital_social', 0) or 0):,} €"),
        ("Siège social", d.get("siege_social", "—")),
        ("Type de domiciliation", d.get("domiciliation", "—")),
        ("Régime fiscal", d.get("regime_fiscal", "—")),
        ("Statut ARE/ARCE", d.get("are", "—")),
        ("ACRE", d.get("acre", "—")),
    ]
    add_kv_table(doc, legal_rows)

    if d.get("questions_statut"):
        add_colored_heading(doc, "Notes sur le statut juridique", 2, "203A43")
        doc.add_paragraph(d["questions_statut"])

    # ── 10. CONCLUSION ───────────────────────────────────────────────────────
    add_colored_heading(doc, "10. Conclusion", 1)

    nom = d.get("nom_entreprise", "ce projet") or "ce projet"
    porteur = d.get("porteur", "l'équipe fondatrice") or "l'équipe fondatrice"
    p = doc.add_paragraph(
        f"{nom} représente une opportunité concrète de répondre à un besoin identifié sur le marché. "
        f"Porté par {porteur}, ce projet s'appuie sur une vision claire, une stratégie adaptée "
        f"et des projections financières réalistes. "
        f"Nous sommes convaincus de sa viabilité et de son potentiel de croissance."
    )
    p.runs[0].font.size = Pt(11)

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Document généré le " + datetime.now().strftime("%d/%m/%Y à %H:%M"))
    run.font.size = Pt(9)
    run.font.italic = True
    run.font.color.rgb = RGBColor.from_string("94A3B8")

    # Save to bytes
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ─── EXCEL GENERATOR ─────────────────────────────────────────────────────────
def excel_header_style(cell, text, bg="2C5364", fg="FFFFFF", bold=True, size=11):
    cell.value = text
    cell.font = Font(name="Arial", bold=bold, color=fg, size=size)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def excel_label_style(cell, text, bold=False, color="0F2027", size=10):
    cell.value = text
    cell.font = Font(name="Arial", bold=bold, color=color, size=size)
    cell.alignment = Alignment(vertical="center")

def excel_value_style(cell, value, num_format='#,##0 "€"', color="000000", bold=False):
    cell.value = value
    cell.font = Font(name="Arial", color=color, size=10, bold=bold)
    cell.number_format = num_format
    cell.alignment = Alignment(horizontal="right", vertical="center")

def thin_border():
    thin = Side(style="thin", color="CBD5E1")
    return Border(left=thin, right=thin, top=thin, bottom=thin)

def add_border(ws, min_row, max_row, min_col, max_col):
    thin = Side(style="thin", color="CBD5E1")
    b = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                             min_col=min_col, max_col=max_col):
        for cell in row:
            cell.border = b


def generate_excel(d):
    wb = openpyxl.Workbook()

    ca1 = int(d.get("ca_an1", 0) or 0)
    ca2 = int(d.get("ca_an2", 0) or 0)
    ca3 = int(d.get("ca_an3", 0) or 0)
    charges_fixes_mois = sum([
        int(d.get("charges_salaires", 0) or 0),
        int(d.get("charges_loyer", 0) or 0),
        int(d.get("charges_logiciels", 0) or 0),
        int(d.get("charges_marketing", 0) or 0),
        int(d.get("charges_compta", 0) or 0),
        int(d.get("autres_charges", 0) or 0),
    ])
    charges_fixes_an = charges_fixes_mois * 12
    cogs_pct = int(d.get("cogs_pct", 20) or 20) / 100
    comm_pct = int(d.get("commission_pct", 3) or 3) / 100

    nom = d.get("nom_entreprise", "Mon Projet") or "Mon Projet"

    # ── SHEET 1 : DASHBOARD ─────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "📊 Dashboard"

    ws1.column_dimensions["A"].width = 35
    ws1.column_dimensions["B"].width = 18
    ws1.column_dimensions["C"].width = 18
    ws1.column_dimensions["D"].width = 18
    ws1.column_dimensions["E"].width = 4
    ws1.column_dimensions["F"].width = 28
    ws1.column_dimensions["G"].width = 22

    # Title banner
    ws1.merge_cells("A1:G1")
    c = ws1["A1"]
    c.value = f"BUSINESS PLAN — {nom.upper()}"
    c.font = Font(name="Arial", bold=True, size=18, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor="0F2027")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 45

    ws1.merge_cells("A2:G2")
    c = ws1["A2"]
    c.value = f"Porteur : {d.get('porteur','—')}   |   Statut : {d.get('statut_juridique','—')}   |   Document : {datetime.now().strftime('%d/%m/%Y')}"
    c.font = Font(name="Arial", size=10, color="FFFFFF", italic=True)
    c.fill = PatternFill("solid", fgColor="203A43")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[2].height = 22

    # KPI Cards
    ws1.row_dimensions[4].height = 20
    ws1.merge_cells("A4:A4")
    ws1["A4"].value = "🎯 KPIs CLÉS"
    ws1["A4"].font = Font(name="Arial", bold=True, size=12, color="2C5364")

    kpi_data = [
        ("Chiffre d'affaires — Année 1", ca1, '#,##0 "€"', "22C55E"),
        ("Chiffre d'affaires — Année 2", ca2, '#,##0 "€"', "22C55E"),
        ("Chiffre d'affaires — Année 3", ca3, '#,##0 "€"', "22C55E"),
        ("Charges fixes annuelles", charges_fixes_an, '#,##0 "€"', "EF4444"),
        ("Résultat estimé — Année 1", ca1 - ca1 * (cogs_pct + comm_pct) - charges_fixes_an, '#,##0 "€"', "2563EB"),
        ("Résultat estimé — Année 3", ca3 - ca3 * (cogs_pct + comm_pct) - charges_fixes_an, '#,##0 "€"', "2563EB"),
        ("Capital social", int(d.get("capital_social", 0) or 0), '#,##0 "€"', "7C3AED"),
        ("Budget de lancement", int(d.get("budget_lancement", 0) or 0), '#,##0 "€"', "D97706"),
    ]

    for i, (label, val, fmt, color) in enumerate(kpi_data):
        row = 5 + i
        ws1.row_dimensions[row].height = 24
        lc = ws1.cell(row=row, column=1, value=label)
        lc.font = Font(name="Arial", size=10, color="374151")
        lc.fill = PatternFill("solid", fgColor="F8FAFC" if i % 2 == 0 else "F1F5F9")
        lc.alignment = Alignment(vertical="center")
        lc.border = thin_border()

        vc = ws1.cell(row=row, column=2, value=int(val))
        vc.font = Font(name="Arial", bold=True, size=11, color=color)
        vc.number_format = fmt
        vc.alignment = Alignment(horizontal="right", vertical="center")
        vc.fill = PatternFill("solid", fgColor="F8FAFC" if i % 2 == 0 else "F1F5F9")
        vc.border = thin_border()

    # Infos projet
    ws1["F4"].value = "📋 INFOS PROJET"
    ws1["F4"].font = Font(name="Arial", bold=True, size=12, color="2C5364")

    info_items = [
        ("Domaine", d.get("domaine", "—")),
        ("Cible", d.get("cible", "—")),
        ("Modèle économique", d.get("modele_eco", "—")),
        ("Marge brute estimée", d.get("marge_brute", "—")),
        ("Maturité du marché", d.get("maturite_marche", "—")),
        ("Nb clients — An 1", d.get("nb_clients_an1", "—")),
        ("Statut juridique", d.get("statut_juridique", "—")),
        ("Domiciliation", d.get("domiciliation", "—")),
    ]
    for i, (k, v) in enumerate(info_items):
        row = 5 + i
        kc = ws1.cell(row=row, column=6, value=k)
        kc.font = Font(name="Arial", bold=True, size=9, color="374151")
        kc.fill = PatternFill("solid", fgColor="EFF6FF")
        kc.border = thin_border()
        kc.alignment = Alignment(vertical="center")

        vc = ws1.cell(row=row, column=7, value=str(v))
        vc.font = Font(name="Arial", size=9, color="000000")
        vc.border = thin_border()
        vc.alignment = Alignment(vertical="center", wrap_text=True)

    # ── SHEET 2 : P&L 3 ANS ─────────────────────────────────────────────────
    ws2 = wb.create_sheet("📈 P&L 3 Ans")
    ws2.column_dimensions["A"].width = 38
    ws2.column_dimensions["B"].width = 20
    ws2.column_dimensions["C"].width = 20
    ws2.column_dimensions["D"].width = 20
    ws2.column_dimensions["E"].width = 20

    ws2.merge_cells("A1:E1")
    c = ws2["A1"]
    c.value = f"COMPTE DE RÉSULTAT PRÉVISIONNEL — {nom.upper()}"
    c.font = Font(name="Arial", bold=True, size=14, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor="0F2027")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 38

    headers = ["", "Année 1", "Année 2", "Année 3", "Variation An1→An3"]
    for ci, h in enumerate(headers, 1):
        c = ws2.cell(row=2, column=ci, value=h)
        c.font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2C5364")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border()
    ws2.row_dimensions[2].height = 28

    cogs1 = int(ca1 * cogs_pct)
    cogs2 = int(ca2 * cogs_pct)
    cogs3 = int(ca3 * cogs_pct)
    comm1 = int(ca1 * comm_pct)
    comm2 = int(ca2 * comm_pct)
    comm3 = int(ca3 * comm_pct)
    mb1 = ca1 - cogs1 - comm1
    mb2 = ca2 - cogs2 - comm2
    mb3 = ca3 - cogs3 - comm3

    sal = int(d.get("charges_salaires", 0) or 0) * 12
    loyer = int(d.get("charges_loyer", 0) or 0) * 12
    log = int(d.get("charges_logiciels", 0) or 0) * 12
    mkt = int(d.get("charges_marketing", 0) or 0) * 12
    compta = int(d.get("charges_compta", 0) or 0) * 12
    autres = int(d.get("autres_charges", 0) or 0) * 12

    res1 = mb1 - charges_fixes_an
    res2 = mb2 - charges_fixes_an
    res3 = mb3 - charges_fixes_an

    def pct_change(v1, v3):
        if v1 == 0:
            return "—"
        return f"+{((v3-v1)/v1*100):.0f}%" if v3 > v1 else f"{((v3-v1)/v1*100):.0f}%"

    pl_rows = [
        ("CHIFFRE D'AFFAIRES", ca1, ca2, ca3, pct_change(ca1, ca3), "2C5364", True, True),
        ("  Coût des ventes (COGS)", -cogs1, -cogs2, -cogs3, "", "EF4444", False, False),
        ("  Commissions / frais paiement", -comm1, -comm2, -comm3, "", "EF4444", False, False),
        ("MARGE BRUTE", mb1, mb2, mb3, pct_change(mb1, mb3), "203A43", True, True),
        ("  Salaires / rémunération", -sal, -sal, -sal, "", "94A3B8", False, False),
        ("  Loyer / domiciliation", -loyer, -loyer, -loyer, "", "94A3B8", False, False),
        ("  Logiciels / abonnements", -log, -log, -log, "", "94A3B8", False, False),
        ("  Marketing / publicité", -mkt, -mkt, -mkt, "", "94A3B8", False, False),
        ("  Comptabilité / juridique", -compta, -compta, -compta, "", "94A3B8", False, False),
        ("  Autres charges fixes", -autres, -autres, -autres, "", "94A3B8", False, False),
        ("TOTAL CHARGES FIXES", -charges_fixes_an, -charges_fixes_an, -charges_fixes_an, "", "374151", True, False),
        ("RÉSULTAT NET ESTIMÉ", res1, res2, res3, pct_change(res1, res3), "22C55E" if res1 >= 0 else "EF4444", True, True),
    ]

    for ri, (label, v1, v2, v3, chg, color, bold, is_total) in enumerate(pl_rows, start=3):
        ws2.row_dimensions[ri + 2].height = 22
        row_num = ri + 2

        bg = "F0FDF4" if is_total and v1 >= 0 else ("FEF2F2" if is_total and v1 < 0 else ("F8FAFC" if ri % 2 == 0 else "FFFFFF"))

        lc = ws2.cell(row=row_num, column=1, value=label)
        lc.font = Font(name="Arial", bold=bold, size=10, color=color if is_total else "374151")
        lc.fill = PatternFill("solid", fgColor=bg)
        lc.border = thin_border()
        lc.alignment = Alignment(vertical="center")

        for ci, val in enumerate([v1, v2, v3], start=2):
            vc = ws2.cell(row=row_num, column=ci, value=val)
            vc.font = Font(name="Arial", bold=bold, size=10,
                          color=color if is_total else ("22C55E" if isinstance(val, int) and val >= 0 else "EF4444"))
            vc.number_format = '#,##0 "€";[Red]-#,##0 "€"'
            vc.fill = PatternFill("solid", fgColor=bg)
            vc.border = thin_border()
            vc.alignment = Alignment(horizontal="right", vertical="center")

        cc = ws2.cell(row=row_num, column=5, value=chg)
        cc.font = Font(name="Arial", bold=bold, size=10, color=color)
        cc.fill = PatternFill("solid", fgColor=bg)
        cc.border = thin_border()
        cc.alignment = Alignment(horizontal="center", vertical="center")

    # Chart CA + Résultat
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = "CA et Résultat sur 3 ans (€)"
    chart.y_axis.title = "Euros (€)"
    chart.x_axis.title = "Années"
    chart.width = 18
    chart.height = 12

    ca_row = 5
    res_row = 16

    cats = Reference(ws2, min_col=2, max_col=4, min_row=2, max_row=2)
    data_ca = Reference(ws2, min_col=2, max_col=4, min_row=ca_row, max_row=ca_row)
    data_res = Reference(ws2, min_col=2, max_col=4, min_row=res_row, max_row=res_row)

    chart.add_data(data_ca, titles_from_data=False)
    chart.add_data(data_res, titles_from_data=False)
    chart.set_categories(cats)
    from openpyxl.chart.series import SeriesLabel
    chart.series[0].title = SeriesLabel(v="Chiffre d'affaires")
    chart.series[1].title = SeriesLabel(v="Résultat net")

    ws2.add_chart(chart, "A19")

    # ── SHEET 3 : CHARGES DÉTAILLÉES ─────────────────────────────────────────
    ws3 = wb.create_sheet("💸 Charges Détail")
    ws3.column_dimensions["A"].width = 35
    ws3.column_dimensions["B"].width = 20
    ws3.column_dimensions["C"].width = 20
    ws3.column_dimensions["D"].width = 20

    ws3.merge_cells("A1:D1")
    c = ws3["A1"]
    c.value = "DÉTAIL DES CHARGES MENSUELLES ET ANNUELLES"
    c.font = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor="203A43")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 35

    for ci, h in enumerate(["Poste de charge", "Mensuel (€)", "Annuel (€)", "% des charges fixes"], 1):
        c = ws3.cell(row=2, column=ci, value=h)
        c.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2C5364")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border()

    charges_detail = [
        ("Salaires / rémunération", int(d.get("charges_salaires", 0) or 0)),
        ("Loyer / domiciliation", int(d.get("charges_loyer", 0) or 0)),
        ("Logiciels / abonnements", int(d.get("charges_logiciels", 0) or 0)),
        ("Marketing / publicité", int(d.get("charges_marketing", 0) or 0)),
        ("Comptabilité / juridique", int(d.get("charges_compta", 0) or 0)),
        ("Autres charges fixes", int(d.get("autres_charges", 0) or 0)),
    ]

    for ri, (label, monthly) in enumerate(charges_detail, start=3):
        annual = monthly * 12
        pct = (annual / charges_fixes_an * 100) if charges_fixes_an > 0 else 0
        bg = "F8FAFC" if ri % 2 == 0 else "FFFFFF"

        lc = ws3.cell(row=ri, column=1, value=label)
        lc.font = Font(name="Arial", size=10, color="374151")
        lc.fill = PatternFill("solid", fgColor=bg)
        lc.border = thin_border()

        mc = ws3.cell(row=ri, column=2, value=monthly)
        mc.number_format = '#,##0 "€"'
        mc.font = Font(name="Arial", size=10, color="000000")
        mc.fill = PatternFill("solid", fgColor=bg)
        mc.border = thin_border()
        mc.alignment = Alignment(horizontal="right")

        ac = ws3.cell(row=ri, column=3, value=annual)
        ac.number_format = '#,##0 "€"'
        ac.font = Font(name="Arial", size=10, color="000000")
        ac.fill = PatternFill("solid", fgColor=bg)
        ac.border = thin_border()
        ac.alignment = Alignment(horizontal="right")

        pc = ws3.cell(row=ri, column=4, value=pct / 100)
        pc.number_format = "0.0%"
        pc.font = Font(name="Arial", size=10, color="2C5364")
        pc.fill = PatternFill("solid", fgColor=bg)
        pc.border = thin_border()
        pc.alignment = Alignment(horizontal="right")

    # Total row
    total_row = len(charges_detail) + 3
    ws3.cell(row=total_row, column=1, value="TOTAL CHARGES FIXES").font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    ws3.cell(row=total_row, column=1).fill = PatternFill("solid", fgColor="2C5364")
    ws3.cell(row=total_row, column=1).border = thin_border()

    ws3.cell(row=total_row, column=2, value=charges_fixes_mois).number_format = '#,##0 "€"'
    ws3.cell(row=total_row, column=2).font = Font(name="Arial", bold=True, color="FFFFFF")
    ws3.cell(row=total_row, column=2).fill = PatternFill("solid", fgColor="2C5364")
    ws3.cell(row=total_row, column=2).border = thin_border()
    ws3.cell(row=total_row, column=2).alignment = Alignment(horizontal="right")

    ws3.cell(row=total_row, column=3, value=charges_fixes_an).number_format = '#,##0 "€"'
    ws3.cell(row=total_row, column=3).font = Font(name="Arial", bold=True, color="FFFFFF")
    ws3.cell(row=total_row, column=3).fill = PatternFill("solid", fgColor="2C5364")
    ws3.cell(row=total_row, column=3).border = thin_border()
    ws3.cell(row=total_row, column=3).alignment = Alignment(horizontal="right")

    ws3.cell(row=total_row, column=4, value=1.0).number_format = "0.0%"
    ws3.cell(row=total_row, column=4).font = Font(name="Arial", bold=True, color="FFFFFF")
    ws3.cell(row=total_row, column=4).fill = PatternFill("solid", fgColor="2C5364")
    ws3.cell(row=total_row, column=4).border = thin_border()
    ws3.cell(row=total_row, column=4).alignment = Alignment(horizontal="right")

    # Pie chart
    from openpyxl.chart import PieChart
    pie = PieChart()
    pie.title = "Répartition des charges fixes"
    pie.width = 14
    pie.height = 10

    data_ref = Reference(ws3, min_col=3, min_row=3, max_row=total_row - 1)
    labels_ref = Reference(ws3, min_col=1, min_row=3, max_row=total_row - 1)
    pie.add_data(data_ref)
    pie.set_categories(labels_ref)

    ws3.add_chart(pie, "F3")

    # ── SHEET 4 : FINANCEMENT ─────────────────────────────────────────────────
    ws4 = wb.create_sheet("💰 Financement")
    ws4.column_dimensions["A"].width = 35
    ws4.column_dimensions["B"].width = 20
    ws4.column_dimensions["C"].width = 22

    ws4.merge_cells("A1:C1")
    c = ws4["A1"]
    c.value = "PLAN DE FINANCEMENT"
    c.font = Font(name="Arial", bold=True, size=14, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor="0F2027")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[1].height = 38

    for ci, h in enumerate(["Source de financement", "Montant (€)", "% du total"], 1):
        c = ws4.cell(row=2, column=ci, value=h)
        c.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2C5364")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border()

    fin_sources = [
        ("Apports personnels", int(d.get("apports_perso", 0) or 0)),
        ("Love money (famille / amis)", int(d.get("love_money", 0) or 0)),
        ("Emprunt bancaire", int(d.get("emprunt", 0) or 0)),
        ("Subventions / crowdfunding", int(d.get("subventions", 0) or 0)),
    ]
    total_fin = sum(v for _, v in fin_sources)

    for ri, (label, val) in enumerate(fin_sources, start=3):
        pct = (val / total_fin) if total_fin > 0 else 0
        bg = "F8FAFC" if ri % 2 == 0 else "FFFFFF"

        lc = ws4.cell(row=ri, column=1, value=label)
        lc.font = Font(name="Arial", size=10, color="374151")
        lc.fill = PatternFill("solid", fgColor=bg)
        lc.border = thin_border()

        vc = ws4.cell(row=ri, column=2, value=val)
        vc.number_format = '#,##0 "€"'
        vc.font = Font(name="Arial", size=10, color="22C55E" if val > 0 else "94A3B8")
        vc.fill = PatternFill("solid", fgColor=bg)
        vc.border = thin_border()
        vc.alignment = Alignment(horizontal="right")

        pc = ws4.cell(row=ri, column=3, value=pct)
        pc.number_format = "0.0%"
        pc.font = Font(name="Arial", size=10, color="2C5364")
        pc.fill = PatternFill("solid", fgColor=bg)
        pc.border = thin_border()
        pc.alignment = Alignment(horizontal="right")

    # Total
    tr = len(fin_sources) + 3
    for ci, (val, fmt) in enumerate([(  "TOTAL FINANCEMENT", "@"),
                                      (total_fin, '#,##0 "€"'),
                                      (1.0, "0.0%")], start=1):
        c = ws4.cell(row=tr, column=ci, value=val)
        c.font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="203A43")
        c.border = thin_border()
        c.number_format = fmt
        if ci > 1:
            c.alignment = Alignment(horizontal="right")

    # Budget
    ws4.cell(row=tr + 2, column=1, value="Budget de lancement estimé")
    ws4.cell(row=tr + 2, column=1).font = Font(name="Arial", bold=True, size=11, color="2C5364")
    ws4.cell(row=tr + 2, column=2, value=int(d.get("budget_lancement", 0) or 0))
    ws4.cell(row=tr + 2, column=2).number_format = '#,##0 "€"'
    ws4.cell(row=tr + 2, column=2).font = Font(name="Arial", bold=True, size=11, color="D97706")
    ws4.cell(row=tr + 2, column=2).alignment = Alignment(horizontal="right")

    gap = total_fin - int(d.get("budget_lancement", 0) or 0)
    ws4.cell(row=tr + 3, column=1, value="Solde (financement - budget)")
    ws4.cell(row=tr + 3, column=1).font = Font(name="Arial", size=10, color="374151")
    ws4.cell(row=tr + 3, column=2, value=gap)
    ws4.cell(row=tr + 3, column=2).number_format = '#,##0 "€"'
    ws4.cell(row=tr + 3, column=2).font = Font(name="Arial", bold=True, size=10,
                                               color="22C55E" if gap >= 0 else "EF4444")
    ws4.cell(row=tr + 3, column=2).alignment = Alignment(horizontal="right")

    # ── SHEET 5 : CA MENSUEL An1 ─────────────────────────────────────────────
    ws5 = wb.create_sheet("📅 CA Mensuel An1")
    ws5.column_dimensions["A"].width = 18
    ws5.column_dimensions["B"].width = 18
    ws5.column_dimensions["C"].width = 18
    ws5.column_dimensions["D"].width = 18

    ws5.merge_cells("A1:D1")
    c = ws5["A1"]
    c.value = f"PROJECTION MENSUELLE — ANNÉE 1 — {nom.upper()}"
    c.font = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor="0F2027")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws5.row_dimensions[1].height = 35

    mois = ["Jan", "Fév", "Mar", "Avr", "Mai", "Jun",
            "Jul", "Aoû", "Sep", "Oct", "Nov", "Déc"]

    for ci, h in enumerate(["Mois", "CA Mensuel (€)", "Charges Fixes (€)", "Résultat (€)"], 1):
        c = ws5.cell(row=2, column=ci, value=h)
        c.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2C5364")
        c.alignment = Alignment(horizontal="center")
        c.border = thin_border()

    # Ramp-up model : montée en puissance progressive
    ramp = [0.04, 0.06, 0.07, 0.08, 0.08, 0.08, 0.08, 0.09, 0.09, 0.1, 0.11, 0.12]
    for ri, (m, r) in enumerate(zip(mois, ramp), start=3):
        ca_m = int(ca1 * r)
        charges_m = charges_fixes_mois
        res_m = ca_m - int(ca_m * (cogs_pct + comm_pct)) - charges_m
        bg = "F8FAFC" if ri % 2 == 0 else "FFFFFF"

        ws5.cell(row=ri, column=1, value=m).font = Font(name="Arial", bold=True, size=10, color="2C5364")
        ws5.cell(row=ri, column=1).fill = PatternFill("solid", fgColor=bg)
        ws5.cell(row=ri, column=1).border = thin_border()

        for ci, val in enumerate([ca_m, charges_m, res_m], start=2):
            c = ws5.cell(row=ri, column=ci, value=val)
            c.number_format = '#,##0 "€"'
            c.font = Font(name="Arial", size=10,
                         color="22C55E" if (ci == 4 and val >= 0) else ("EF4444" if (ci == 4 and val < 0) else "000000"))
            c.fill = PatternFill("solid", fgColor=bg)
            c.border = thin_border()
            c.alignment = Alignment(horizontal="right")

    # Totals
    trow = 15
    ws5.cell(row=trow, column=1, value="TOTAL").font = Font(name="Arial", bold=True, color="FFFFFF")
    ws5.cell(row=trow, column=1).fill = PatternFill("solid", fgColor="203A43")
    ws5.cell(row=trow, column=1).border = thin_border()

    for ci in range(2, 5):
        col_letter = get_column_letter(ci)
        c = ws5.cell(row=trow, column=ci)
        c.value = f"=SUM({col_letter}3:{col_letter}14)"
        c.number_format = '#,##0 "€"'
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="203A43")
        c.border = thin_border()
        c.alignment = Alignment(horizontal="right")

    # Line chart CA mensuel
    line = LineChart()
    line.title = "Évolution du CA mensuel — Année 1"
    line.y_axis.title = "Euros (€)"
    line.x_axis.title = "Mois"
    line.width = 20
    line.height = 12

    data_ref = Reference(ws5, min_col=2, max_col=4, min_row=2, max_row=14)
    cats_ref = Reference(ws5, min_col=1, min_row=3, max_row=14)
    line.add_data(data_ref, titles_from_data=True)
    line.set_categories(cats_ref)

    ws5.add_chart(line, "A17")

    # ── Save ────────────────────────────────────────────────────────────────
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─── ROUTER ──────────────────────────────────────────────────────────────────
def main():
    step = st.session_state.step

    if step > 0:
        render_progress()

    steps_map = {
        0: step_0_welcome,
        1: step_1_idea,
        2: step_2_project,
        3: step_3_market,
        4: step_4_product,
        5: step_5_team,
        6: step_6_financing,
        7: step_7_projections,
        8: step_8_legal,
        9: step_9_recap,
    }

    fn = steps_map.get(step, step_0_welcome)
    fn()


if __name__ == "__main__":
    main()
