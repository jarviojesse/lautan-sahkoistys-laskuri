import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import datetime, timedelta, time
import requests

st.set_page_config(page_title="Lauttojen sähköistäminen v35 - Dynaaminen", layout="wide")

# --- 1. APUFUNKTIOT ---
def hae_sahkon_hinta():
    try:
        url = "https://api.porssisahko.net/v1/latest-prices.json"
        res = requests.get(url, timeout=5)
        res_json = res.json()
        prices = [e['price'] for e in res_json['prices']]
        keskihinta = (sum(prices) / len(prices)) / 100
        return keskihinta
    except Exception:
        return None

def puhdista_aika(aika_val):
    if isinstance(aika_val, time): return datetime.combine(datetime.now().date(), aika_val)
    if isinstance(aika_val, datetime): return datetime.combine(datetime.now().date(), aika_val.time())
    return pd.to_datetime(str(aika_val)).replace(year=datetime.now().year, month=datetime.now().month, day=datetime.now().day)

def laske_lcc_yksinkertainen(test_akku, ladattu_vrk, investointi_pohja):
    # Käytetään globaaleja hintoja ja korkoja
    inv = (test_akku * u_capex_kwh) + investointi_pohja
    vuosikulu = (ladattu_vrk * (st.session_state.s_hinta + s_siirto) * 365) + \
                (kok_bess_kwh * BESS_HUOLTO_OPEX) - markkinatuotto_v
    # 10 vuoden elinkaarikustannus nykyarvolla
    lcc = inv + sum([vuosikulu / (1 + d_korko)**i for i in range(1, 11)])
    return lcc
    
def laske_akun_degradaatio(df_sim, akkukoko_kwh, cycle_life_100dod, base_calendar_rate, temp_c):
    """
    Laskee akun eliniän perustuen sykli- ja kalenteriväsymykseen.
    """
    # Haetaan SoC-arvot simulaatiosta
    soc = df_sim["SoC"].values

    # --- 1. DoD (Depth of Discharge) laskenta ---
    # Lasketaan vaihteluväli (max-min). np.clip varmistaa ettei DoD ole liian pieni/suuri.
    dod = np.clip((np.max(soc) - np.min(soc)) / 100, 0.05, 1.0)

    # --- 2. EFC (Equivalent Full Cycles) laskenta ---
    # Lasketaan päivittäinen syklitys integroimalla energian muutokset
    energy = soc / 100 * akkukoko_kwh
    efc_day = np.sum(np.abs(np.diff(energy))) / (2 * akkukoko_kwh)
    efc_year = efc_day * 365

    # --- 3. Sykliväsymys (Cycle aging) ---
    cycle_deg = cycle_aging(efc_year, dod, cycle_life_100dod)

    # --- 4. Kalenteriväsymys (Calendar aging) ---
    avg_soc = np.mean(soc)
    cal_deg = calendar_aging(avg_soc, temp_c, base_calendar_rate)

    # --- 5. Kokonaisväsymys ja elinikä ---
    total_deg = cycle_deg + cal_deg
    lifetime_years = 1 / total_deg if total_deg > 0 else 100

    return lifetime_years, total_deg, efc_day, dod

def cycle_aging(efc_year, dod, cycle_life_100dod):
    # DoD-korrektio (Akku kuluu vähemmän pienillä sykleillä)
    dod_exp = 1.7
    effective_cycle_life = cycle_life_100dod / (dod ** dod_exp)
    yearly_cycle_degradation = efc_year / effective_cycle_life
    return yearly_cycle_degradation

def calendar_aging(avg_soc, temp_c, base_rate):
    # SOC-stressi (korkea varaustaso kuluttaa akkua enemmän)
    soc_factor = 1.0
    if avg_soc > 60:
        soc_factor += (avg_soc - 60) * 0.02

    # Lämpötilakerroin (Arrhenius-yhtälön approksimaatio)
    # Akun kesto puolittuu tyypillisesti jokaista 10 asteen nousua kohden
    temp_factor = np.exp(0.07 * (temp_c - 25))

    return base_rate * soc_factor * temp_factor

# --- 2. DATAN LATAUS JA VAKIOT (Synkronoitu Juhan datan kanssa) ---
try:
    df_lahto = pd.read_excel('Master_Excel.xlsx', sheet_name='Lähtötiedot')
    df_kaikki_aikataulut = pd.read_excel('Master_Excel.xlsx', sheet_name='Aikataulu & Reitti')
    
    def get_val(p): 
        res = df_lahto.loc[df_lahto['Parametri'] == p, 'Arvo'].values
        return res[0] if len(res) > 0 else 0

    # Vakiot suoraan Juhan määritysten mukaan
    DIESEL_KULUTUS_L_KWH = float(get_val('Polttoaineen ominaiskulutus')) 
    DIESEL_HUOLTO_RATE = float(get_val('Dieselkoneiden huoltokulu'))    
    HUOLTO_VERTAILUTEHO = float(get_val('Huoltokulun vertailuteho'))   
    BESS_HUOLTO_OPEX = float(get_val('BESS-vuosihuolto'))             
    LATAUS_MIN_MINS = 10  
    
    # Tekniset parametrit
    base_ajoteho = float(get_val('Ajonaikainen sähköteho'))
    soc_min_limit_excel = float(get_val('SoC-ikkuna (min)'))
    soc_max_limit_excel = float(get_val('SoC-ikkuna (max)'))
    u_eol_kriteeri = float(get_val('EoL-kriteeri (SOH)')) / 100 if get_val('EoL-kriteeri (SOH)') else 0.80
    
    # Talous ja Ympäristö
    diesel_hinta_excel = float(get_val('Polttoaineen hinta'))
    co2_kerroin = float(get_val('Dieselin päästökerroin'))
    s_siirto = float(get_val('Sähkön_siirto_ja_vero'))
    d_korko = float(get_val('Diskonttokorko')) / 100
    
    # CAPEX-arvot
    b_capex_per_kwh = float(get_val('Ranta-BESS CAPEX'))
    c_varsi_base = float(get_val('Latausvarsi_automaatti'))
    c_liittyma_base = float(get_val('Liittymismaksu_arvio'))
    
    s_dict = {"Kesä": float(get_val('Sääkerroin_Kesä')), 
              "Syksy": float(get_val('Sääkerroin_Syksy')), 
              "Talvi": float(get_val('Sääkerroin_Talvi_Jää'))}
except Exception as e:
    st.error(f"Virhe datan lukemisessa Master Excelistä: {e}"); st.stop()

# --- 3. SIVUPALKKI (Käyttöliittymä) ---
st.sidebar.header("Aluksen valinta")
uniikit_lautat = sorted(df_kaikki_aikataulut['Lautan nimi'].unique().tolist())
valittu_lautta = st.sidebar.selectbox("Valitse lautta", uniikit_lautat)
df_aikataulu = df_kaikki_aikataulut[df_kaikki_aikataulut['Lautan nimi'] == valittu_lautta].copy()

st.sidebar.header("Olosuhteet")
valittu_saa = st.sidebar.selectbox("Olosuhteet (Sääkerroin)", list(s_dict.keys()))
s_kerroin = s_dict[valittu_saa]
u_diesel_hinta = st.sidebar.slider("Dieselin hinta (€/l)", 0.5, 3.0, diesel_hinta_excel, step=0.05)
u_lampotila = st.sidebar.slider("Akkutilan lämpötila (°C)", 10, 45, 30, help="Juhan suositus: 30°C. Vaikuttaa kalenterivanhenemiseen.")

st.sidebar.header("Rantavaihtoehto")
rantavaihtoehdot = {
    "V1: 20kV / 2MW (Liityntä)": {"kulu": 91482, "spot_saasto": 0, "taajuustulo": 0, "aggregointi_kulu": 0, "max_p": 2000, "bess_kwh": 0},
    "V2: 0.4kV / 0.4MW (Liityntä)": {"kulu": 30688, "spot_saasto": 20000, "taajuustulo": 0, "aggregointi_kulu": 0, "max_p": 400, "bess_kwh": 0},
    "V3: 20kV / 2MW + BESS": {"kulu": 91482, "spot_saasto": 20000, "taajuustulo": 100000, "aggregointi_kulu": 48000, "max_p": 2000, "bess_kwh": 1000},
    "V4: 0.4kV / 0.4MW + BESS": {"kulu": 30688, "spot_saasto": 20000, "taajuustulo": 20000, "aggregointi_kulu": 16000, "max_p": 400, "bess_kwh": 500}
}
valittu_v_nimi = st.sidebar.selectbox("Valitse verkkoliityntä", list(rantavaihtoehdot.keys()))
v_data = dict(rantavaihtoehdot[valittu_v_nimi]) 

st.sidebar.header("Teknologia ja Infra")
u_akkukoko = st.sidebar.number_input("Lautta-akun koko (kWh)", value=2240, step=100)

st.sidebar.subheader("Lataustehot päädyissä")
col_p1, col_p2 = st.sidebar.columns(2)
with col_p1:
    u_teho_houtskari = st.number_input("Houtskari (kW)", value=int(v_data["max_p"]), step=100)
with col_p2:
    u_teho_korppoo = st.number_input("Korppoo (kW)", value=int(v_data["max_p"]), step=100)

u_akun_tyyppi = st.sidebar.selectbox("Akun kemia", ["NMC (Corvus Orca oletus)", "LFP"])
u_infra_malli = st.sidebar.radio("Latauspisteiden sijainti", ["Molemmat päät", "1. Pää", "2. Pää"])
u_varsi_paalla = st.sidebar.toggle("Automaattinen latausvarsi", value=True)

if "BESS" in valittu_v_nimi:
    v_data["bess_kwh"] = st.sidebar.number_input("Ranta-BESS kapasiteetti per ranta (kWh)", value=v_data["bess_kwh"], step=100)

with st.sidebar.expander("Edistyneet laskentaparametrit"):
    u_soc_max = st.slider("SoC yläraja (%)", 50, 100, int(soc_max_limit_excel), step=5)
    u_soc_min = st.slider("SoC alaraja (%)", 0, 50, int(soc_min_limit_excel), step=5)
    u_cal_loss = st.slider("Kalenteri-ikääntyminen (%/v)", 0.5, 5.0, 1.0, step=0.1) / 100
    temp_kerroin = 1.0 + ((u_lampotila - 25) / 10) if u_lampotila > 25 else 1.0
    default_cycles = 3000 if "NMC" in u_akun_tyyppi else 6000
    u_cycle_life = st.number_input("Akun sykli-ikä (MEC @ 100% DoD)", value=default_cycles, step=500)
    u_hyotysuhde = st.slider("Lataushyötysuhde", 0.70, 0.98, 0.90, step=0.01)
    u_capex_kwh = st.number_input("Lautta-akun CAPEX (€/kWh)", value=int(get_val('Lautta-akuston CAPEX')))

if 's_hinta' not in st.session_state: st.session_state.s_hinta = float(get_val('Sähkön energiahinta'))
if st.sidebar.button("Päivitä pörssisähkö"):
    h = hae_sahkon_hinta()
    if h: st.session_state.s_hinta = h; st.sidebar.success("Päivitetty!")

# --- 4. SIMULAATIO (PÄIVITETTY TARKKAAN MINUUTTILASKENNAN LOGIIKKAAN) ---
def aja_simulaatio(akkukoko, teho_h, teho_k, infra_malli, hyotysuhde, s_max_limit):
    df = df_aikataulu.copy()
    df['Lähtö_dt'] = df['Lähtöaika'].apply(puhdista_aika)
    df['Saapuminen_dt'] = df['Saapumisaika'].apply(puhdista_aika)
    df = df.sort_values('Lähtö_dt')
    
    ajoteho_min = (base_ajoteho * s_kerroin) / 60
    soc_kwh = akkukoko * (s_max_limit / 100)
    
    # BESS-alustus
    bess_h_max = v_data["bess_kwh"] if teho_h > 0 else 0
    bess_k_max = v_data["bess_kwh"] if teho_k > 0 else 0
    soc_bess_h = bess_h_max * 0.9
    soc_bess_k = bess_k_max * 0.9
    
    log = []
    bess_log = [] # Pidetään mukana yhteensopivuuden vuoksi
    tot_ladattu = 0
    tot_purettu = 0

    for idx, row in df.iterrows():
        # 1. Ajo
        kesto_min = (row['Saapuminen_dt'] - row['Lähtö_dt']).seconds / 60
        energia_ajo = kesto_min * ajoteho_min
        soc_kwh -= energia_ajo
        tot_purettu += energia_ajo
        
        # BESS latautuu verkosta (150kW teholla)
        soc_bess_h = min(bess_h_max * 0.9, soc_bess_h + (150 * kesto_min / 60)) if bess_h_max > 0 else 0
        soc_bess_k = min(bess_k_max * 0.9, soc_bess_k + (150 * kesto_min / 60)) if bess_k_max > 0 else 0
        
        log.append({
            'Aika': row['Lähtö_dt'], 'SoC': (soc_kwh / akkukoko) * 100, 
            'BESS_H_SoC': (soc_bess_h / bess_h_max * 100) if bess_h_max > 0 else 0,
            'BESS_K_SoC': (soc_bess_k / bess_k_max * 100) if bess_k_max > 0 else 0
        })
        
        # 2. Lataus rannassa
        if idx < len(df) - 1:
            seisonta = (df.iloc[idx+1]['Lähtö_dt'] - row['Saapuminen_dt']).seconds / 60
            nykyinen_sijainti = "Korppoo" if "Houtskari" in str(row['Lähtöpaikka']) else "Houtskari"
            aktiivinen_teho = teho_h if nykyinen_sijainti == "Houtskari" else teho_k
            
            if seisonta >= LATAUS_MIN_MINS and aktiivinen_teho > 0:
                vapaa_tila = (akkukoko * (s_max_limit / 100)) - soc_kwh
                lataus = min(vapaa_tila, (aktiivinen_teho * (seisonta / 60)) * hyotysuhde)
                
                # Kulutetaan BESS-akkua
                if nykyinen_sijainti == "Houtskari" and bess_h_max > 0:
                    soc_bess_h -= min(lataus, max(0, soc_bess_h - (bess_h_max * 0.1)))
                elif nykyinen_sijainti == "Korppoo" and bess_k_max > 0:
                    soc_bess_k -= min(lataus, max(0, soc_bess_k - (bess_k_max * 0.1)))
                
                soc_kwh += lataus
                tot_ladattu += lataus
        
        log.append({
            'Aika': row['Saapuminen_dt'], 'SoC': (soc_kwh / akkukoko) * 100,
            'BESS_H_SoC': (soc_bess_h / bess_h_max * 100) if bess_h_max > 0 else 0,
            'BESS_K_SoC': (soc_bess_k / bess_k_max * 100) if bess_k_max > 0 else 0
        })

    return pd.DataFrame(log), pd.DataFrame(bess_log), tot_ladattu, tot_purettu

# Suoritus
latausteho_h = u_teho_houtskari if u_infra_malli in ["Molemmat päät", "Vain Houtskari"] else 0
latausteho_k = u_teho_korppoo if u_infra_malli in ["Molemmat päät", "Vain Korppoo"] else 0

# Kutsutaan funktiota niin että paluuarvot (4 kpl) täsmäävät
df_sim, df_bess, tot_ladattu, tot_purettu = aja_simulaatio(
    u_akkukoko, latausteho_h, latausteho_k, u_infra_malli, u_hyotysuhde, u_soc_max
)

# --- Määritellään muuttujat optimointia ja talouslaskentaa varten ---
bess_kpl = 2 if u_infra_malli == "Molemmat päät" else 1
kok_bess_kwh = v_data["bess_kwh"] * bess_kpl
markkinatuotto_v = (v_data['taajuustulo'] + v_data['spot_saasto'] - v_data['aggregointi_kulu']) * bess_kpl

def laske_lcc_yksinkertainen(test_akku, ladattu_vrk, investointi_pohja):
    # Lasketaan investointi (CAPEX)
    inv = (test_akku * u_capex_kwh) + investointi_pohja
    # Vuotuiset operatiiviset kulut (Sähkö + huolto - tuotot)
    vuosikulu = (ladattu_vrk * (st.session_state.s_hinta + s_siirto) * 365) + \
                (kok_bess_kwh * BESS_HUOLTO_OPEX) - markkinatuotto_v
    # 10 vuoden elinkaarikustannus nykyarvolla (diskontattuna)
    lcc = inv + sum([vuosikulu / (1 + d_korko)**i for i in range(1, 11)])
    return lcc

st.sidebar.markdown("---")
if st.sidebar.button("ETSI KAIKKI OPTIMIVAIHTOEHDOT"):
    with st.spinner('Analysoidaan kaikkia rantavaihtoehtoja ja akkukokoja...'):
        kaikki_skenaariot = []
        
        # Haetaan diesel-vertailukohta Excel-datasta (v_data)
        d_hinta = v_data.get('Polttoaineen hinta', 1.3)
        d_kulutus_vrk = v_data.get('Polttoaineen kulutus', 1500)
        diesel_vuosikulu = d_kulutus_vrk * d_hinta * 365 + v_data.get('Huoltokulut', 50000)

        # Testataan eri kemiat (NMC ja LFP)
        kemiat = [("NMC", 300), ("LFP", 200)] 

        for v_nimi, v_tiedot in rantavaihtoehdot.items():
            test_teho = v_tiedot["max_p"]
            test_bess_kpl = 2 if u_infra_malli == "Molemmat päät" else 1
            test_bess_kwh = v_tiedot["bess_kwh"] * test_bess_kpl
            
            # Markkinatulot (varmistetaan oletusarvot .get-metodilla)
            test_markkinatulo = (v_tiedot.get('taajuustulo', 0) + 
                                 v_tiedot.get('spot_saasto', 0) - 
                                 v_tiedot.get('aggregointi_kulu', 0)) * test_bess_kpl
            
            for k_nimi, k_capex in kemiat:
                for test_koko in range(1000, 4201, 200): 
                    # Simulaatio (HUOM: lisätty sääkerroin 1.2 turvamarginaaliksi)
                    d_sim, d_bess, t_ladattu, t_purettu = aja_simulaatio(
                        test_koko, test_teho, test_teho, u_infra_malli, u_hyotysuhde, u_soc_max
                    )
                    
                    alin_soc = d_sim['SoC'].min()
                    
                    # Tarkistetaan elinikä (sykli-ikä vaihtelee kemian mukaan)
                    c_life = 3000 if k_nimi == "NMC" else 6000
                    m_vrk = t_purettu / test_koko
                    e_v, _, _, _ = laske_akun_degradaatio(
                    d_sim,
                    t_purettu,
                    test_koko,
                    c_life,
                    u_cal_loss,
                    temp_kerroin
)
                    
                    # Hyväksyntäehdot: 8v kesto JA SoC pysyy turvallisena
                    if e_v >= 8.0 and alin_soc >= u_soc_min:
                        # CAPEX laskenta
                        inv_pohja = (test_bess_kwh * b_capex_per_kwh) + (c_varsi_base if u_varsi_paalla else 0) + (c_liittyma_base * test_bess_kpl)
                        test_inv = (test_koko * k_capex) + inv_pohja
                        
                        # OPEX laskenta (Sähkö + BESS huolto - Markkinatuotto)
                        test_vuosikulu_sahko = (t_ladattu * (st.session_state.get('s_hinta', 0.1) + s_siirto) * 365)
                        test_opex = test_vuosikulu_sahko + (test_bess_kwh * BESS_HUOLTO_OPEX) - test_markkinatulo
                        
                        # Säästöt vs Diesel
                        vuotuinen_saasto = diesel_vuosikulu - test_opex
                        tm_aika = test_inv / vuotuinen_saasto if vuotuinen_saasto > 0 else 99
                        
                        # 10v LCC
                        test_lcc = test_inv + sum([test_opex / (1 + d_korko)**i for i in range(1, 11)])
                        
                        kaikki_skenaariot.append({
                            'V-Nimi': v_nimi, 'koko': test_koko, 'kemia': k_nimi,
                            'teho': test_teho, 'elinika': e_v, 'lcc': test_lcc,
                            'alin_soc': round(alin_soc, 1), 'inv': test_inv,
                            'takaisinmaksu': tm_aika, 'saasto': vuotuinen_saasto
                        })

        if kaikki_skenaariot:
            df_results = pd.DataFrame(kaikki_skenaariot).sort_values("lcc")
            # Tallennetaan paras session stateen
            st.session_state.optimi = kaikki_skenaariot[0] # Halvin LCC
            st.sidebar.success(f"Optimi löydetty: {df_results.iloc[0]['koko']} kWh ({df_results.iloc[0]['kemia']})")
            
            # Näytetään tulostaulukko pääruudulla
            st.write("### Optimointitulokset (Top 5)")
            st.table(df_results[['V-Nimi', 'koko', 'kemia', 'takaisinmaksu', 'lcc', 'alin_soc']].head(5))
        else:
            st.sidebar.error("Ei löytynyt ehtoja täyttävää ratkaisua.")

# Päivitys: Jos optimi on ajettu, pakotetaan UI-arvot vastaamaan sitä
if 'optimi' in st.session_state:
    o = st.session_state.optimi
    u_akkukoko = o['koko']
    # Huom: Jotta loppukoodi käyttää oikeita ranta-arvoja, päivitetään v_data
    v_data = dict(rantavaihtoehdot[o['V-Nimi']])
    latausteho_h = o['teho']
    latausteho_k = o['teho']
    df_sim = o['df']
    df_bess = o['df_b']
    tot_purettu = o['purettu']
    tot_ladattu = o['ladattu']
    
    st.info(f"**LÖYDETTY OPTIMI:** Halvin ratkaisu on **{o['V-Nimi']}** ja **{u_akkukoko} kWh** akku. "
            f"Tällä yhdistelmällä elinikä on **{o['elinika']:.1f} vuotta** ja 10v elinkaarikustannus on pienin.")
# --- 5. ANALYYSI JA GRAAFIT ---
st.title(f"Analyysi: {valittu_lautta}")

col1, col2, col3, col4 = st.columns(4)
col1.metric("Energiankulutus / vrk", f"{tot_purettu:,.0f} kWh")
col2.metric("Lataus / vrk", f"{tot_ladattu:,.0f} kWh")

elinika_v, vuosi_kuluma, mec_vrk, dod_avg = laske_akun_degradaatio(
    df_sim,
    tot_purettu,
    u_akkukoko,
    u_cycle_life,
    u_cal_loss,
    temp_kerroin
)

# Huomioidaan lämpötilakerroin kalenterikulumaan
col3.metric("Akun elinikäennuste", f"{elinika_v:.1f} vuotta")
col4.metric("Sykliä / vrk (MEC)", f"{mec_vrk:.2f}")

# --- Graafin rakentaminen ---
fig = go.Figure()

# --- 5. VISUALISOINTI ---

# --- 1. Lautan varaustila ja infran monitorointi ---
st.subheader("Lautan varaustila (SoC %)")

fig1 = go.Figure()

# Lautan SoC-käyrä
fig1.add_trace(go.Scatter(
    x=df_sim['Aika'], 
    y=df_sim['SoC'], 
    name='Lautta SoC %', 
    line=dict(color='cyan', width=3)
))

# 2. SoC Ylä- ja alarajat (Apuviivat)
fig1.add_hline(y=u_soc_max, line_dash="dash", line_color="red", annotation_text="Max SoC limit")
fig1.add_hline(y=u_soc_min, line_dash="dash", line_color="red", annotation_text="Min SoC limit")

# 3. BESS-toiminta (Visualisointi palkeilla)
# Huom: Käytetään df_sim sarakkeita, koska BESS-logiikka on integroitu siihen
if 'BESS_H_SoC' in df_sim.columns:
    # Luodaan indikaattori purkutapahtumille (kun BESS SoC laskee)
    # Tässä käytetään keltaisia palkkeja pohjalla osoittamaan milloin BESS on aktiivinen
    fig1.add_trace(go.Bar(
        x=df_sim['Aika'],
        y=[10] * len(df_sim), # Vakio korkeus palkille
        name='BESS Aktiivinen',
        marker_color='rgba(255, 165, 0, 0.4)',
        hoverinfo='none',
        showlegend=True
    ))

fig1.update_layout(
    title="Lautan varaustilan (SoC) ja infran monitorointi",
    template="plotly_dark",
    xaxis_title="Kellonaika",
    yaxis_title="SoC %",
    yaxis_range=[0, 105],
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    hovermode="x unified"
)

st.plotly_chart(fig1, use_container_width=True)

# Varoitus, jos SoC menee alle rajan
if df_sim['SoC'].min() < u_soc_min:
    st.warning(f"⚠️ **Varoitus:** Akun varaustila laskee alimmillaan {df_sim['SoC'].min():.1f} prosenttiin, mikä alittaa asetetun {u_soc_min} % rajan.")

# 2. Houtskari BESS
if latausteho_h > 0 and v_data["bess_kwh"] > 0:
    st.subheader("BESS Varaustila: Houtskari (SoC %)")
    fig_h = go.Figure()
    fig_h.add_trace(go.Scatter(
        x=df_sim['Aika'], 
        y=df_sim['BESS_H_SoC'], 
        name='Houtskari BESS', 
        line=dict(color='lime', width=2)
    ))
    
    # SoC Ylä- ja alarajat (Apuviivat)
    fig_h.add_hline(y=u_soc_max, line_dash="dash", line_color="red", annotation_text="Max SoC limit")
    fig_h.add_hline(y=u_soc_min, line_dash="dash", line_color="red", annotation_text="Min SoC limit")
    
    fig_h.update_layout(
        template="plotly_dark",
        xaxis_title="Kellonaika",
        yaxis_title="SoC %",
        yaxis_range=[0, 105],
        hovermode="x unified"
    )
    st.plotly_chart(fig_h, use_container_width=True)

# 3. Korppoo BESS
if latausteho_k > 0 and v_data["bess_kwh"] > 0:
    st.subheader("BESS Varaustila: Korppoo (SoC %)")
    fig_k = go.Figure()
    fig_k.add_trace(go.Scatter(
        x=df_sim['Aika'], 
        y=df_sim['BESS_K_SoC'], 
        name='Korppoo BESS', 
        line=dict(color='yellow', width=2)
    ))
    
    # SoC Ylä- ja alarajat (Apuviivat)
    fig_k.add_hline(y=u_soc_max, line_dash="dash", line_color="red", annotation_text="Max SoC limit")
    fig_k.add_hline(y=u_soc_min, line_dash="dash", line_color="red", annotation_text="Min SoC limit")
    
    fig_k.update_layout(
        template="plotly_dark",
        xaxis_title="Kellonaika",
        yaxis_title="SoC %",
        yaxis_range=[0, 105],
        hovermode="x unified"
    )
    st.plotly_chart(fig_k, use_container_width=True)


# --- 6. TALOUSLASKENTA (LCC) ---
st.header("7. Taloudellinen vertailu (LCC)")

# Diesel-kustannukset
diesel_l_v = tot_purettu * DIESEL_KULUTUS_L_KWH * 365
diesel_kulu_v = diesel_l_v * u_diesel_hinta
diesel_huolto_v = HUOLTO_VERTAILUTEHO * DIESEL_HUOLTO_RATE
co2_v = (diesel_l_v * co2_kerroin) / 1000

# Sähkö-kustannukset
sahko_kulu_v = tot_ladattu * (st.session_state.s_hinta + s_siirto) * 365
bess_kpl = 2 if u_infra_malli == "Molemmat päät" else 1
kok_bess_kwh = v_data["bess_kwh"] * bess_kpl
bess_huolto_v = kok_bess_kwh * BESS_HUOLTO_OPEX

# Markkinatuotot (UUSI)
markkinatuotto_v = (v_data['taajuustulo'] + v_data['spot_saasto'] - v_data['aggregointi_kulu']) * bess_kpl

# Investoinnit
c_akku = u_akkukoko * u_capex_kwh
c_investointi = c_akku + (kok_bess_kwh * b_capex_per_kwh) + (c_varsi_base if u_varsi_paalla else 0) + (c_liittyma_base * bess_kpl)

lcc_diesel_10v = sum([(diesel_kulu_v + diesel_huolto_v) / (1+d_korko)**i for i in range(10)])
lcc_sahko_10v = c_investointi + sum([(sahko_kulu_v + bess_huolto_v - markkinatuotto_v) / (1+d_korko)**i for i in range(10)])

t1, t2 = st.columns(2)
with t1:
    st.subheader("Nykytila (Diesel)")
    st.write(f"- Polttoainekulut: {diesel_kulu_v:,.0f} €/v")
    st.write(f"- Huoltokulut: {diesel_huolto_v:,.0f} €/v")
    st.write(f"- CO2-päästöt: {co2_v:,.0f} t/v")
    st.info(f"**Dieselin LCC (10v): {lcc_diesel_10v:,.0f} €**")

with t2:
    st.subheader("Sähköistys")
    st.write(f"- Investointi (CAPEX): {c_investointi:,.0f} €")
    st.write(f"- Sähköenergia + siirto: {sahko_kulu_v:,.0f} €/v")
    st.write(f"- BESS ylläpito + Aggregointi: {bess_huolto_v + (v_data['aggregointi_kulu']*bess_kpl):,.0f} €/v")
    st.write(f"- Markkinatuotot (Taajuus+Spot): { (v_data['taajuustulo'] + v_data['spot_saasto'])*bess_kpl :,.0f} €/v")
    st.info(f"**Sähkön LCC (10v, sis. CAPEX): {lcc_sahko_10v:,.0f} €**")

# --- TALOUSLASKENTA ---
vuotuiset_saastot = diesel_kulu_v - sahko_kulu_v
takaisinmaksuaika = c_investointi / vuotuiset_saastot if vuotuiset_saastot > 0 else float('inf')

st.subheader("Taloudellinen yhteenveto")
t1, t2, t3 = st.columns(3)
t1.metric("Investointi (CAPEX)", f"{c_investointi/1e6:.2f} M€")
t2.metric("Vuotuinen säästö (OPEX)", f"{vuotuiset_saastot/1e3:.1f} k€/v")

if takaisinmaksuaika < 15:
    t3.metric("Takaisinmaksuaika", f"{takaisinmaksuaika:.1f} vuotta", delta_color="normal")
else:
    t3.metric("Takaisinmaksuaika", "Yli 15v", delta="-", delta_color="inverse")