import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import datetime, timedelta, time
import requests
import itertools

st.set_page_config(page_title="Lauttojen sähköistäminen v35 - Dynaaminen", layout="wide")

# =========================================================
# 0. VAKIOT (EI RIIPPUVAISIA USER INPUTISTA)
# =========================================================

rantavaihtoehdot_default = {
    "V1: 20kV / 2MW (Liityntä)": {"kulu": 91482, "spot_saasto": 0, "taajuustulo": 0, "aggregointi_kulu": 0, "max_p": 2000, "bess_kwh": 0},
    "V2: 0.4kV / 0.4MW (Liityntä)": {"kulu": 30688, "spot_saasto": 20000, "taajuustulo": 0, "aggregointi_kulu": 0, "max_p": 400, "bess_kwh": 0},
    "V3: 20kV / 2MW + BESS": {"kulu": 91482, "spot_saasto": 20000, "taajuustulo": 100000, "aggregointi_kulu": 48000, "max_p": 2000, "bess_kwh": 1000},
    "V4: 0.4kV / 0.4MW + BESS": {"kulu": 30688, "spot_saasto": 20000, "taajuustulo": 20000, "aggregointi_kulu": 16000, "max_p": 400, "bess_kwh": 500}
}

# =========================================================
# 1. APUFUNKTIOT
# =========================================================

def puhdista_aika(aika_val):
    if isinstance(aika_val, time):
        return datetime.combine(datetime.now().date(), aika_val)
    if isinstance(aika_val, datetime):
        return datetime.combine(datetime.now().date(), aika_val.time())
    return pd.to_datetime(str(aika_val)).replace(year=datetime.now().year,
                                                 month=datetime.now().month,
                                                 day=datetime.now().day)

def hae_sahkon_hinta():
    try:
        url = "https://api.porssisahko.net/v1/latest-prices.json"
        res = requests.get(url, timeout=5)
        res_json = res.json()
        prices = [e['price'] for e in res_json['prices']]
        return (sum(prices) / len(prices)) / 100
    except:
        return None

# ---------------- SIMULAATIO ----------------

def aja_simulaatio(akkukoko, teho_h, teho_k, infra_malli, hyotysuhde, s_max_limit):
    df = df_aikataulu.copy()
    df['Lähtö_dt'] = df['Lähtöaika'].apply(puhdista_aika)
    df['Saapuminen_dt'] = df['Saapumisaika'].apply(puhdista_aika)
    df = df.sort_values('Lähtö_dt')

    ajoteho_min = (base_ajoteho * s_kerroin) / 60
    soc_kwh = akkukoko * (s_max_limit / 100)

    bess_h_max = v_data["bess_kwh"] if teho_h > 0 else 0
    bess_k_max = v_data["bess_kwh"] if teho_k > 0 else 0

    soc_bess_h = bess_h_max * 0.9
    soc_bess_k = bess_k_max * 0.9

    log = []
    tot_ladattu = 0
    tot_purettu = 0

    for idx, row in df.iterrows():
        kesto_min = (row['Saapuminen_dt'] - row['Lähtö_dt']).seconds / 60
        energia_ajo = kesto_min * ajoteho_min
        soc_kwh -= energia_ajo
        tot_purettu += energia_ajo

        log.append({
            'Aika': row['Lähtö_dt'],
            'SoC': (soc_kwh / akkukoko) * 100
        })

        if idx < len(df) - 1:
            seisonta = (df.iloc[idx+1]['Lähtö_dt'] - row['Saapuminen_dt']).seconds / 60
            vapaa = (akkukoko * (s_max_limit / 100)) - soc_kwh

            lataus = min(vapaa, (teho_h * (seisonta / 60)) * hyotysuhde)

            soc_kwh += lataus
            tot_ladattu += lataus

        log.append({
            'Aika': row['Saapuminen_dt'],
            'SoC': (soc_kwh / akkukoko) * 100
        })

    return pd.DataFrame(log), None, tot_ladattu, tot_purettu


# ---------------- DEGRADAATIO ----------------

def laske_akun_degradaatio(df_sim, akkukoko, cycle_life, base_rate, temp_c):
    soc = df_sim["SoC"].values
    dod = np.clip((np.max(soc) - np.min(soc)) / 100, 0.05, 1.0)

    efc_year = 300
    cycle_deg = efc_year / cycle_life
    cal_deg = base_rate * np.exp(0.07 * (temp_c - 25))

    total = cycle_deg + cal_deg
    life = 1 / total if total > 0 else 100

    return life, total, efc_year, dod


# ---------------- OPTIMOINTI ----------------

def optimoi_jarjestelma(rantavaihtoehdot):
    tulokset = []

    for v_nimi, v_data_local in rantavaihtoehdot.items():
        for koko in range(1200, 3000, 400):

            dummy_soc = 50

            if dummy_soc < 20:
                continue

            tulokset.append({
                "v": v_nimi,
                "koko": koko,
                "lcc": koko * 200
            })

    if not tulokset:
        return None, None

    df = pd.DataFrame(tulokset).sort_values("lcc")
    return df.iloc[0].to_dict(), df


# =========================================================
# 2. DATA (ladataan ennen UI:ta)
# =========================================================

df_lahto = pd.read_excel('Master_Excel.xlsx', sheet_name='Lähtötiedot')
df_kaikki_aikataulut = pd.read_excel('Master_Excel.xlsx', sheet_name='Aikataulu & Reitti')

def get_val(p):
    res = df_lahto.loc[df_lahto['Parametri'] == p, 'Arvo'].values
    return res[0] if len(res) > 0 else 0

base_ajoteho = float(get_val('Ajonaikainen sähköteho'))
diesel_hinta_excel = float(get_val('Polttoaineen hinta'))
s_siirto = float(get_val('Sähkön_siirto_ja_vero'))
d_korko = float(get_val('Diskonttokorko')) / 100

# =========================================================
# 3. SIDEBAR (KAIKKI INPUTIT ENNEN LASKENTAA)
# =========================================================

st.sidebar.header("Alus")
uniikit_lautat = df_kaikki_aikataulut['Lautan nimi'].unique()
valittu_lautta = st.sidebar.selectbox("Valitse", uniikit_lautat)
df_aikataulu = df_kaikki_aikataulut[df_kaikki_aikataulut['Lautan nimi'] == valittu_lautta]

st.sidebar.header("Ranta")
valittu_ranta = st.sidebar.selectbox("Ranta", list(rantavaihtoehdot_default.keys()))
v_data = rantavaihtoehdot_default[valittu_ranta]

st.sidebar.header("Akku")
u_akkukoko = st.sidebar.number_input("kWh", value=2000)

st.sidebar.header("Simulaatio")
u_soc_max = st.sidebar.slider("Max SoC", 50, 100, 90)
u_soc_min = st.sidebar.slider("Min SoC", 0, 50, 20)
u_lampotila = st.sidebar.slider("Lämpötila", 10, 45, 30)

s_kerroin = 1.0

# =========================================================
# 4. SUORITUS
# =========================================================

if st.sidebar.button("Optimoi"):
    optimi, df_all = optimoi_jarjestelma(rantavaihtoehdot_default)
    st.write(optimi)

df_sim, _, tot_ladattu, tot_purettu = aja_simulaatio(
    u_akkukoko,
    v_data["max_p"],
    v_data["max_p"],
    "Molemmat",
    0.9,
    u_soc_max
)

life, _, _, _ = laske_akun_degradaatio(df_sim, u_akkukoko, 3000, 0.01, u_lampotila)

# =========================================================
# 5. VISUALISOINTI
# =========================================================

st.title("Lauttasimulaatio")

st.line_chart(df_sim.set_index("Aika")["SoC"])
st.metric("Elinikä (v)", round(life, 1))