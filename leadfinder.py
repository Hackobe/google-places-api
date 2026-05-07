import streamlit as st
import pandas as pd
import requests
from io import BytesIO

st.set_page_config(
    page_title="Lead Finder | Albama Web",
    page_icon="📍",
    layout="wide"
)

st.markdown("""
<style>

.stApp {
    background: linear-gradient(135deg, #0B0F17 0%, #111827 55%, #0A0D12 100%);
    color: #F8FAFC;
}

.block-container {
    max-width: 1180px;
    padding-top: 2.5rem;
    padding-bottom: 4rem;
}

.hero {
    padding: 42px;
    border-radius: 26px;
    background: rgba(17, 24, 39, 0.92);
    border: 1px solid rgba(148, 163, 184, 0.16);
    box-shadow: 0 20px 60px rgba(0,0,0,0.28);
    margin-bottom: 36px;
}

.badge {
    display: inline-block;
    padding: 8px 14px;
    border-radius: 999px;
    background: rgba(148,163,184,0.10);
    border: 1px solid rgba(148,163,184,0.20);
    color: #D6DEE9;
    font-size: 12px;
    font-weight: 700;
    margin-bottom: 20px;
}

.hero-title {
    font-size: 52px;
    line-height: 1.05;
    font-weight: 850;
    color: white;
    margin-bottom: 18px;
    letter-spacing: -1.5px;
}

.hero-title span {
    color: #93C5FD;
}

.hero-desc {
    max-width: 760px;
    color: #CBD5E1;
    font-size: 17px;
    line-height: 1.7;
}

div[data-testid="stTextInput"] label,
div[data-testid="stCheckbox"] label {
    color: #D6DEE9 !important;
    font-weight: 650 !important;
    font-size: 13px !important;
}

div[data-testid="stTextInput"] [data-baseweb="input"] {
    background: rgba(15, 23, 42, 0.85) !important;
    border: 1px solid rgba(148,163,184,0.18) !important;
    border-radius: 14px !important;
    overflow: hidden !important;
}

div[data-testid="stTextInput"] input {
    background: transparent !important;
    border: none !important;
    color: white !important;
    height: 52px !important;
    padding: 0 16px !important;
    font-size: 15px !important;
}

div[data-testid="stTextInput"] input:focus {
    box-shadow: none !important;
}

.stButton button {
    height: 52px;
    border-radius: 14px;
    border: 1px solid rgba(147,197,253,0.25);
    background: linear-gradient(135deg, #2563EB 0%, #3B82F6 100%);
    color: white;
    font-size: 15px;
    font-weight: 750;
    padding: 0 26px;
    box-shadow: 0 12px 28px rgba(37,99,235,0.28);
    transition: 0.18s ease;
}

.stButton button:hover {
    transform: translateY(-1px);
    background: linear-gradient(135deg, #1D4ED8 0%, #2563EB 100%);
}

.stDownloadButton button {
    height: 52px;
    border-radius: 14px;
    background: #0F766E;
    color: white;
    border: none;
    font-weight: 750;
}

.info-card {
    padding: 24px;
    border-radius: 18px;
    background: rgba(15, 23, 42, 0.92);
    border: 1px solid rgba(148,163,184,0.16);
}

.info-card strong {
    display: block;
    font-size: 30px;
    margin-bottom: 6px;
    color: white;
}

.info-card span {
    color: #94A3B8;
    font-size: 14px;
}

div[data-testid="stDataFrame"] {
    border-radius: 18px;
    overflow: hidden;
    border: 1px solid rgba(148,163,184,0.16);
}

</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
<div class="badge">ALBAMA WEB · Lead Finder</div>
<div class="hero-title">Google Maps <span>Lead Finder</span></div>
<div class="hero-desc">Wyszukuj lokalne firmy z Google Places API, filtruj potencjalne okazje sprzedażowei eksportuj gotową bazę leadów do Excela.</div>
</div>
""", unsafe_allow_html=True)

api_key = st.secrets["GOOGLE_API_KEY"]

col1, col2, col3, col4 = st.columns([1.1, 1.1, 0.9, 0.9])

with col1:
    branza = st.text_input(
        "Branża",
        value="barber"
    )

with col2:
    miasto = st.text_input(
        "Miasto",
        value="Warszawa"
    )

with col3:
    min_opinie_input = st.text_input(
        "Minimalna liczba opinii",
        value="0"
    )

    try:
        min_opinie = int(min_opinie_input)
    except:
        min_opinie = 0

with col4:
    liczba_leadow_input = st.text_input(
        "Liczba leadów",
        value="20"
    )

    try:
        liczba_leadow = int(liczba_leadow_input)
    except:
        liczba_leadow = 20

    liczba_leadow = max(1, min(20, liczba_leadow))

tylko_bez_strony = st.checkbox(
    "Pokaż tylko firmy bez strony www"
)

search_button = st.button("Szukaj leadów")


def search_places(api_key, query, liczba_leadow):

    url = "https://places.googleapis.com/v1/places:searchText"

    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": api_key,
        "X-Goog-FieldMask": (
            "places.displayName,"
            "places.formattedAddress,"
            "places.nationalPhoneNumber,"
            "places.websiteUri,"
            "places.rating,"
            "places.userRatingCount,"
            "places.googleMapsUri"
        )
    }

    payload = {
        "textQuery": query,
        "languageCode": "pl",
        "maxResultCount": liczba_leadow
    }

    response = requests.post(
        url,
        headers=headers,
        json=payload
    )

    if response.status_code != 200:
        st.error("Błąd API")
        st.code(response.text)
        return []

    data = response.json()

    return data.get("places", [])


def convert_to_excel(df):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Leady"
        )

    return output.getvalue()


if search_button:

    if not api_key:

        st.warning("Wpisz najpierw API Key.")

    else:

        query = f"{branza} {miasto}"

        with st.spinner("Szukam leadów w Google Maps..."):

            places = search_places(
                api_key,
                query,
                liczba_leadow
            )

        results = []

        for place in places:

            name = place.get(
                "displayName",
                {}
            ).get("text", "")

            address = place.get(
                "formattedAddress",
                ""
            )

            phone = place.get(
                "nationalPhoneNumber",
                ""
            )

            website = place.get(
                "websiteUri",
                ""
            )

            rating = place.get(
                "rating",
                ""
            )

            reviews = place.get(
                "userRatingCount",
                0
            )

            maps = place.get(
                "googleMapsUri",
                ""
            )

            if reviews < min_opinie:
                continue

            if tylko_bez_strony and website:
                continue

            results.append({
                "Branża": branza,
                "Miasto": miasto,
                "Nazwa firmy": name,
                "Adres": address,
                "Telefon": phone,
                "Strona www": website,
                "Ocena": rating,
                "Liczba opinii": reviews,
                "Google Maps": maps,
                "Status": "Nie dzwonione",
                "Notatka": ""
            })

        if results:

            df = pd.DataFrame(results)

            total = len(df)
            bez_www = df["Strona www"].eq("").sum()
            z_tel = df["Telefon"].ne("").sum()

            st.markdown("## Wyniki")

            col_a, col_b, col_c = st.columns(3)

            with col_a:
                st.markdown(f"""
                <div class="info-card">
                    <strong>{total}</strong>
                    <span>Znalezione leady</span>
                </div>
                """, unsafe_allow_html=True)

            with col_b:
                st.markdown(f"""
                <div class="info-card">
                    <strong>{bez_www}</strong>
                    <span>Firmy bez strony www</span>
                </div>
                """, unsafe_allow_html=True)

            with col_c:
                st.markdown(f"""
                <div class="info-card">
                    <strong>{z_tel}</strong>
                    <span>Leady z numerem telefonu</span>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("### Lista firm")

            st.dataframe(
                df,
                use_container_width=True,
                height=500
            )

            excel_file = convert_to_excel(df)

            st.download_button(
                label="Pobierz bazę leadów Excel",
                data=excel_file,
                file_name=f"leady_{branza}_{miasto}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        else:
            st.warning("Brak wyników dla tych filtrów.")