"""
CompeteIQ DEMO — لوحة تحليل المنافسين / Competitor Analysis Tool
بيانات تجريبية مدمجة مباشرة — لا يحتاج scraping ولا ملفات خارجية
Embedded demo data — no scraping or external files needed
"""
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import time

from analyzers import analyze_prices, analyze_reviews, analyze_seo, find_opportunities
from excel_exporter import export_to_excel

# ── Page Config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="CompeteIQ — Competitor Analysis",
    page_icon="🕵️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Translations ─────────────────────────────────────────────────────────────
T = {
    "ar": {
        "title": "🕵️ CompeteIQ",
        "subtitle": "أداة تحليل المنافسين للمتاجر الإلكترونية — أسعار · تقييمات · SEO · فرص تنافسية",
        "demo_badge": "🎯 نسخة تجريبية — Demo Version",
        "settings": "⚙️ إعدادات الأداة",
        "urls_label": "🔗 روابط المنتجات (رابط واحد لكل سطر)",
        "urls_help": "أدخل الروابط التجريبية أو جرب تغييرها.",
        "demo_info": "💡 هذه نسخة تجريبية تعرض بيانات نموذجية.\nالنسخة الكاملة تستخرج بيانات حقيقية من أي متجر إلكتروني.",
        "platforms_title": "📋 المنصات المدعومة في النسخة الكاملة:",
        "platforms": [
            "✅ Amazon (.sa, .ae, .com)",
            "✅ Jarir",
            "✅ Extra",
            "✅ أي متجر يدعم JSON-LD",
            "⚠️ SHEIN / AliExpress (قد لا يعمل على جميع الصفحات)",
        ],
        "lang_toggle": "🌐 اللغة / Language",
        "start_btn": "🚀 ابدأ التحليل",
        "new_analysis": "🔄 تحليل جديد",
        "warn_no_urls": "⚠️ أدخل رابطاً واحداً على الأقل.",
        "progress_text": "🔍 جارٍ تحليل المنتج {0} من {1}...",
        "done_text": "✅ اكتمل التحليل!",
        "kpi_products": "📦 المنتجات المُحللة",
        "kpi_avg_price": "💰 متوسط السعر",
        "kpi_avg_rating": "⭐ متوسط التقييم",
        "kpi_total_reviews": "💬 إجمالي المراجعات",
        "kpi_available": "🛒 متوفر للشراء",
        "unit_product": "منتج",
        "unit_review": "مراجعة",
        "unit_out_of_5": "من 5",
        "tabs": ["📋 نظرة عامة", "💰 تحليل الأسعار", "⭐ التقييمات", "🔍 SEO", "💡 الفرص"],
        "overview_title": "جميع المنتجات المُحللة",
        "col_platform": "المنصة",
        "col_title": "العنوان",
        "col_price": "السعر (SAR)",
        "col_rating": "التقييم ⭐",
        "col_reviews": "المراجعات 💬",
        "col_availability": "التوفر",
        "col_seller": "البائع",
        "pie_title": "توزيع المنتجات حسب المنصة",
        "price_compare_title": "مقارنة الأسعار",
        "price_chart_title": "مقارنة الأسعار مع المتوسط",
        "avg_label": "المتوسط",
        "min_label": "الأدنى",
        "max_label": "الأعلى",
        "col_product": "المنتج",
        "col_vs_avg": "الفرق عن المتوسط",
        "col_vs_avg_pct": "الفرق %",
        "col_position": "الموقع التنافسي",
        "reviews_chart_title": "مقارنة التقييمات ودرجة الثقة",
        "social_proof": "درجة الثقة الاجتماعية",
        "best_rated": "**⭐ الأعلى تقييماً:**",
        "most_reviewed": "**💬 الأكثر مراجعات:**",
        "shared_keywords_title": "الكلمات المفتاحية المشتركة",
        "shared_keywords_chart": "أكثر الكلمات المفتاحية استخداماً بين المنافسين",
        "col_keyword": "الكلمة المفتاحية",
        "col_keyword_count": "عدد المنتجات",
        "title_length_title": "طول عناوين المنتجات",
        "title_length_chart": "طول العنوان (التوصية: 80-150 حرف)",
        "col_chars": "عدد الأحرف",
        "min_recommend": "الحد الأدنى (80)",
        "max_recommend": "الحد الأقصى (150)",
        "opportunities_title": "الفرص التنافسية المكتشفة",
        "radar_title": "الخريطة التنافسية",
        "radar_chart_title": "الخريطة التنافسية متعددة الأبعاد",
        "radar_categories": ["التقييم", "القيمة مقابل السعر", "المراجعات", "التوفر"],
        "export_title": "📥 تصدير التقرير",
        "export_btn": "⬇️ تحميل تقرير Excel الكامل",
        "export_caption": "التقرير يتضمن: نظرة عامة · تحليل الأسعار · التقييمات · SEO · الفرص التنافسية",
        "footer": "CompeteIQ Demo · تحليل المنافسين للمتاجر الإلكترونية",
        "yes": "متوفر",
        "no": "غير متوفر",
        "insight_price_gap": "💡 يوجد فارق سعري بين المنتجات — راجع استراتيجية التسعير",
        "insight_unavailable": "🛒 بعض المنتجات غير متوفرة — فرصة لتلبية الطلب الموجود",
        "insight_low_rating": "⚠️ متوسط التقييمات منخفض — فرصة للتميز بجودة أعلى",
        "insight_balanced": "✅ السوق متوازن — ركز على جودة المنتج وخدمة العملاء",
    },
    "en": {
        "title": "🕵️ CompeteIQ",
        "subtitle": "E-commerce competitor analysis tool — Prices · Ratings · SEO · Opportunities",
        "demo_badge": "🎯 Demo Version",
        "settings": "⚙️ Tool Settings",
        "urls_label": "🔗 Product URLs (one per line)",
        "urls_help": "Enter demo URLs or try changing them.",
        "demo_info": "💡 This is a demo showing sample data.\nThe full version extracts real data from any online store.",
        "platforms_title": "📋 Supported platforms in the full version:",
        "platforms": [
            "✅ Amazon (.sa, .ae, .com)",
            "✅ Jarir",
            "✅ Extra",
            "✅ Any store with JSON-LD support",
            "⚠️ SHEIN / AliExpress (may not work on all pages)",
        ],
        "lang_toggle": "🌐 Language / اللغة",
        "start_btn": "🚀 Start Analysis",
        "new_analysis": "🔄 New Analysis",
        "warn_no_urls": "⚠️ Please enter at least one URL.",
        "progress_text": "🔍 Analyzing product {0} of {1}...",
        "done_text": "✅ Analysis complete!",
        "kpi_products": "📦 Products Analyzed",
        "kpi_avg_price": "💰 Average Price",
        "kpi_avg_rating": "⭐ Average Rating",
        "kpi_total_reviews": "💬 Total Reviews",
        "kpi_available": "🛒 In Stock",
        "unit_product": "products",
        "unit_review": "reviews",
        "unit_out_of_5": "out of 5",
        "tabs": ["📋 Overview", "💰 Price Analysis", "⭐ Reviews", "🔍 SEO", "💡 Opportunities"],
        "overview_title": "All Analyzed Products",
        "col_platform": "Platform",
        "col_title": "Title",
        "col_price": "Price (SAR)",
        "col_rating": "Rating ⭐",
        "col_reviews": "Reviews 💬",
        "col_availability": "Availability",
        "col_seller": "Seller",
        "pie_title": "Product Distribution by Platform",
        "price_compare_title": "Price Comparison",
        "price_chart_title": "Price Comparison vs Average",
        "avg_label": "Average",
        "min_label": "Lowest",
        "max_label": "Highest",
        "col_product": "Product",
        "col_vs_avg": "Diff vs Average",
        "col_vs_avg_pct": "Diff %",
        "col_position": "Competitive Position",
        "reviews_chart_title": "Ratings & Trust Score Comparison",
        "social_proof": "Social Proof Score",
        "best_rated": "**⭐ Highest Rated:**",
        "most_reviewed": "**💬 Most Reviewed:**",
        "shared_keywords_title": "Shared Keywords",
        "shared_keywords_chart": "Most Common Keywords Among Competitors",
        "col_keyword": "Keyword",
        "col_keyword_count": "Product Count",
        "title_length_title": "Product Title Length",
        "title_length_chart": "Title Length (Recommended: 80-150 chars)",
        "col_chars": "Characters",
        "min_recommend": "Min (80)",
        "max_recommend": "Max (150)",
        "opportunities_title": "Discovered Opportunities",
        "radar_title": "Competitive Map",
        "radar_chart_title": "Multi-Dimensional Competitive Map",
        "radar_categories": ["Rating", "Value for Price", "Reviews", "Availability"],
        "export_title": "📥 Export Report",
        "export_btn": "⬇️ Download Full Excel Report",
        "export_caption": "Report includes: Overview · Price Analysis · Reviews · SEO · Opportunities",
        "footer": "CompeteIQ Demo · E-commerce Competitor Analysis Tool",
        "yes": "In Stock",
        "no": "Out of Stock",
        "insight_price_gap": "💡 There's a price gap between products — review your pricing strategy",
        "insight_unavailable": "🛒 Some products are out of stock — opportunity to meet existing demand",
        "insight_low_rating": "⚠️ Average rating is low — opportunity to stand out with better quality",
        "insight_balanced": "✅ Market is balanced — focus on product quality and customer service",
    },
}

if "lang" not in st.session_state:
    st.session_state.lang = "ar"

lang = st.session_state.lang
t = T[lang]
is_rtl = lang == "ar"

# ── CSS ───────────────────────────────────────────────────────────────────────
direction = "rtl" if is_rtl else "ltr"
border_side = "right" if is_rtl else "left"

st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Arabic:wght@300;400;500;600;700&display=swap');

  html, body, [class*="css"] {{ font-family: 'IBM Plex Sans Arabic', sans-serif; }}
  .main .block-container {{ direction: {direction}; }}
  [data-testid="stSidebar"] > div {{ direction: {direction}; }}

  [data-testid="stAppDeployButton"] {{ display: none !important; }}
  #MainMenu {{ visibility: hidden !important; }}
  footer {{ visibility: hidden !important; }}
  header {{ background-color: transparent !important; }}

  [data-testid="collapsedControl"] {{
    position: fixed !important;
    top: 15px !important;
    left: 15px !important;
    display: flex !important;
    visibility: visible !important;
    opacity: 1 !important;
    z-index: 999999 !important;
    background-color: #1e293b !important;
    border: 1px solid #334155 !important;
    border-radius: 8px !important;
  }}

  .main-header {{
    background: linear-gradient(135deg, #1a1f3a 0%, #2d3875 50%, #1a1f3a 100%);
    border-radius: 16px; padding: 2rem 2.5rem; margin-bottom: 2rem;
    box-shadow: 0 8px 32px rgba(26,31,58,0.4); text-align: center;
  }}
  .main-header h1 {{ color: #fff; font-size: 2.4rem; font-weight: 700; margin: 0; }}
  .main-header p  {{ color: #a5b4fc; font-size: 1.05rem; margin: 0.5rem 0 0; }}

  .demo-badge {{
    background: linear-gradient(90deg, #f59e0b, #ef4444);
    color: white; border-radius: 20px; padding: 0.3rem 1rem;
    font-size: 0.85rem; font-weight: 700; display: inline-block;
    margin-bottom: 1rem;
  }}

  .metric-card {{
    background: #fff; border: 1px solid #e8ecf8; border-radius: 14px;
    padding: 1.2rem 1.5rem; text-align: center;
    box-shadow: 0 2px 12px rgba(0,0,0,0.06);
    transition: transform 0.2s, box-shadow 0.2s;
  }}
  .metric-card:hover {{ transform: translateY(-3px); box-shadow: 0 8px 24px rgba(79,142,247,0.15); }}
  .metric-card .label {{ color: #6b7280; font-size: 0.85rem; font-weight: 500; margin-bottom: 0.4rem; }}
  .metric-card .value {{ color: #1a1f3a; font-size: 1.9rem; font-weight: 700; }}
  .metric-card .sub   {{ color: #9ca3af; font-size: 0.78rem; margin-top: 0.2rem; }}

  .section-title {{
    font-size: 1.25rem; font-weight: 700; color: #1a1f3a;
    border-{border_side}: 4px solid #4f8ef7; padding-{border_side}: 0.75rem; margin: 1.5rem 0 1rem;
  }}
  .insight-card {{
    background: linear-gradient(135deg, #f0f4ff, #fff);
    border: 1px solid #c7d7fd; border-radius: 12px;
    padding: 1rem 1.25rem; margin-bottom: 0.75rem;
    font-size: 0.95rem; color: #1e293b; line-height: 1.6;
  }}
  .stTabs [data-baseweb="tab"] {{ font-size: 0.95rem; font-weight: 600; }}
  .stButton > button {{ border-radius: 10px; font-weight: 600; font-size: 1rem; padding: 0.6rem 2rem; }}
  [data-testid="stSidebar"] {{ background: #f8faff; border-{('left' if is_rtl else 'right')}: 1px solid #e8ecf8; }}
</style>
""", unsafe_allow_html=True)

# ── Demo Scenarios (Noon removed → Jarir / Extra) ────────────────────────────
DEMO_SCENARIOS = {
    "https://www.amazon.sa/dp/B0CKLKNBPY": [
        {
            "platform": "Amazon",
            "url": "https://www.amazon.sa/dp/B0CKLKNBPY",
            "title": "سامسونج جالاكسي S24 الترا 256 جيجا تيتانيوم أسود",
            "title_en": "Samsung Galaxy S24 Ultra 256GB Titanium Black",
            "price": 4299.0, "currency": "SAR", "rating": 4.6,
            "reviews_count": 2847, "availability": "متوفر",
            "seller": "Amazon.sa", "keywords": ["سامسونج", "جالاكسي", "s24", "الترا", "هاتف", "ذكي"],
            "error": None,
        },
        {
            "platform": "Jarir",
            "url": "https://www.jarir.com/samsung-galaxy-s24",
            "title": "سامسونج جالاكسي S24 الترا - تيتانيوم أسود 256GB",
            "title_en": "Samsung Galaxy S24 Ultra - Titanium Black 256GB",
            "price": 4499.0, "currency": "SAR", "rating": 4.5,
            "reviews_count": 634, "availability": "متوفر",
            "seller": "جرير", "keywords": ["سامسونج", "s24", "الترا", "256", "تيتانيوم"],
            "error": None,
        },
        {
            "platform": "Extra",
            "url": "https://www.extra.com/samsung-galaxy-s24",
            "title": "سامسونج جالاكسي S24 الترا 256 جيجا",
            "title_en": "Samsung Galaxy S24 Ultra 256GB",
            "price": 4199.0, "currency": "SAR", "rating": 4.4,
            "reviews_count": 1203, "availability": "متوفر",
            "seller": "Extra", "keywords": ["samsung", "galaxy", "s24", "ultra", "smartphone"],
            "error": None,
        },
    ],
    "https://www.amazon.sa/dp/B0FQFV3HJ2": [
        {
            "platform": "Amazon",
            "url": "https://www.amazon.sa/dp/B0FQFV3HJ2",
            "title": "آيفون 15 برو ماكس 256 جيجا - تيتانيوم طبيعي",
            "title_en": "iPhone 15 Pro Max 256GB - Natural Titanium",
            "price": 5299.0, "currency": "SAR", "rating": 4.8,
            "reviews_count": 5621, "availability": "متوفر",
            "seller": "Apple Store", "keywords": ["آيفون", "15", "برو", "ماكس", "تيتانيوم", "أبل"],
            "error": None,
        },
        {
            "platform": "Extra",
            "url": "https://www.extra.com/iphone-15-pro-max",
            "title": "آيفون 15 برو ماكس 256 جيجا - تيتانيوم طبيعي",
            "title_en": "Apple iPhone 15 Pro Max 256GB Natural Titanium",
            "price": 5199.0, "currency": "SAR", "rating": 4.7,
            "reviews_count": 892, "availability": "متوفر",
            "seller": "Extra", "keywords": ["apple", "iphone", "15", "pro", "max", "titanium"],
            "error": None,
        },
        {
            "platform": "Jarir",
            "url": "https://www.jarir.com/iphone-15-pro-max",
            "title": "آيفون 15 برو ماكس 256GB تيتانيوم طبيعي",
            "title_en": "iPhone 15 Pro Max 256GB Natural Titanium",
            "price": 5399.0, "currency": "SAR", "rating": 4.6,
            "reviews_count": 1544, "availability": "غير متوفر",
            "seller": "جرير", "keywords": ["iphone", "15", "pro", "max", "256", "natural"],
            "error": None,
        },
    ],
    "https://www.amazon.sa/dp/B0FQFYMLS5": [
        {
            "platform": "Amazon",
            "url": "https://www.amazon.sa/dp/B0FQFYMLS5",
            "title": "ماك بوك برو 14 إنش M3 - 8GB RAM 512GB SSD",
            "title_en": "MacBook Pro 14-inch M3 - 8GB RAM 512GB SSD",
            "price": 6999.0, "currency": "SAR", "rating": 4.9,
            "reviews_count": 3102, "availability": "متوفر",
            "seller": "Amazon.sa", "keywords": ["ماك", "بوك", "برو", "m3", "لابتوب", "أبل"],
            "error": None,
        },
        {
            "platform": "Jarir",
            "url": "https://www.jarir.com/macbook-pro-14",
            "title": "ماك بوك برو 14 إنش M3 8GB 512GB - رمادي",
            "title_en": "MacBook Pro 14-inch M3 Chip 8GB 512GB - Space Gray",
            "price": 7199.0, "currency": "SAR", "rating": 4.8,
            "reviews_count": 421, "availability": "متوفر",
            "seller": "جرير", "keywords": ["macbook", "pro", "14", "m3", "512", "apple"],
            "error": None,
        },
        {
            "platform": "Extra",
            "url": "https://www.extra.com/macbook-pro-14-m3",
            "title": "ماك بوك برو 14 إنش M3 8GB 512GB رمادي",
            "title_en": "Apple MacBook Pro 14\" M3 8GB 512GB Space Gray",
            "price": 6899.0, "currency": "SAR", "rating": 4.7,
            "reviews_count": 267, "availability": "متوفر",
            "seller": "Extra", "keywords": ["apple", "macbook", "pro", "m3", "laptop", "512gb"],
            "error": None,
        },
    ],
}

DEFAULT_URLS = "\n".join(DEMO_SCENARIOS.keys())

# ── Session State ─────────────────────────────────────────────────────────────
if "results" not in st.session_state:
    st.session_state.results = []
if "scraped" not in st.session_state:
    st.session_state.scraped = False

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    # Language toggle — first element
    lang_choice = st.selectbox(
        t["lang_toggle"],
        options=["ar", "en"],
        format_func=lambda x: "العربية" if x == "ar" else "English",
        index=0 if st.session_state.lang == "ar" else 1,
    )
    if lang_choice != st.session_state.lang:
        st.session_state.lang = lang_choice
        st.rerun()

    st.markdown(f"### {t['settings']}")
    st.markdown("---")

    urls_input = st.text_area(
        t["urls_label"],
        value=DEFAULT_URLS,
        height=180,
        help=t["urls_help"],
    )

    st.markdown("---")
    st.info(t["demo_info"])

    st.markdown(f"**{t['platforms_title']}**")
    for p in t["platforms"]:
        st.caption(p)

    st.markdown("---")
    analyze_btn = st.button(t["start_btn"], type="primary", use_container_width=True)

    if st.session_state.scraped:
        if st.button(t["new_analysis"], use_container_width=True):
            st.session_state.results = []
            st.session_state.scraped = False
            st.rerun()

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="main-header">
  <h1>{t['title']}</h1>
  <p>{t['subtitle']}</p>
</div>
""", unsafe_allow_html=True)

st.markdown(f'<div style="text-align:center"><span class="demo-badge">{t["demo_badge"]}</span></div>', unsafe_allow_html=True)

# ── Main Logic ────────────────────────────────────────────────────────────────
if analyze_btn:
    urls = [u.strip() for u in urls_input.strip().splitlines() if u.strip()]

    if not urls:
        st.warning(t["warn_no_urls"])
    else:
        results = []
        progress_bar = st.progress(0, text=t["progress_text"].format(1, len(urls)))

        for i, url in enumerate(urls):
            for pct in range(0, 34):
                time.sleep(0.01)
                progress_bar.progress(
                    int((i / len(urls)) * 100 + pct / 3),
                    text=t["progress_text"].format(i + 1, len(urls))
                )

            if url in DEMO_SCENARIOS:
                results.extend(DEMO_SCENARIOS[url])
            else:
                results.extend(DEMO_SCENARIOS[list(DEMO_SCENARIOS.keys())[0]])

        progress_bar.progress(100, text=t["done_text"])
        time.sleep(0.5)

        st.session_state.results = results
        st.session_state.scraped = True
        st.rerun()

# ── Dashboard ─────────────────────────────────────────────────────────────────
if st.session_state.scraped and st.session_state.results:
    products = st.session_state.results
    successful = [p for p in products if not p.get("error")]

    title_key = "title_en" if lang == "en" else "title"

    prices_ok  = [p["price"] for p in successful if p.get("price")]
    ratings_ok = [p["rating"] for p in successful if p.get("rating")]
    reviews_ok = [p["reviews_count"] for p in successful if p.get("reviews_count")]

    def avail_label(p):
        return t["yes"] if p.get("availability") == "متوفر" else t["no"]

    # ── KPI Cards ──────────────────────────────────────────────────────────
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.markdown(f"""<div class="metric-card">
            <div class="label">{t['kpi_products']}</div>
            <div class="value">{len(successful)}</div>
            <div class="sub">{t['unit_product']}</div></div>""", unsafe_allow_html=True)
    with c2:
        avg_p = round(sum(prices_ok)/len(prices_ok), 1) if prices_ok else "N/A"
        st.markdown(f"""<div class="metric-card">
            <div class="label">{t['kpi_avg_price']}</div>
            <div class="value">{avg_p}</div>
            <div class="sub">SAR</div></div>""", unsafe_allow_html=True)
    with c3:
        avg_r = round(sum(ratings_ok)/len(ratings_ok), 2) if ratings_ok else "N/A"
        st.markdown(f"""<div class="metric-card">
            <div class="label">{t['kpi_avg_rating']}</div>
            <div class="value">{avg_r}</div>
            <div class="sub">{t['unit_out_of_5']}</div></div>""", unsafe_allow_html=True)
    with c4:
        total_rev = sum(reviews_ok) if reviews_ok else 0
        st.markdown(f"""<div class="metric-card">
            <div class="label">{t['kpi_total_reviews']}</div>
            <div class="value">{total_rev:,}</div>
            <div class="sub">{t['unit_review']}</div></div>""", unsafe_allow_html=True)
    with c5:
        avail = sum(1 for p in successful if p.get("availability") == "متوفر")
        st.markdown(f"""<div class="metric-card">
            <div class="label">{t['kpi_available']}</div>
            <div class="value">{avail}/{len(successful)}</div>
            <div class="sub">{t['unit_product']}</div></div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Tabs ───────────────────────────────────────────────────────────────
    tab1, tab2, tab3, tab4, tab5 = st.tabs(t["tabs"])

    # Tab 1 — Overview
    with tab1:
        st.markdown(f'<div class="section-title">{t["overview_title"]}</div>', unsafe_allow_html=True)
        df = pd.DataFrame([{
            t["col_platform"]: p.get("platform",""),
            t["col_title"]: p.get(title_key, p.get("title",""))[:70],
            t["col_price"]: p.get("price"),
            t["col_rating"]: p.get("rating"),
            t["col_reviews"]: p.get("reviews_count"),
            t["col_availability"]: avail_label(p),
            t["col_seller"]: p.get("seller",""),
        } for p in successful])
        st.dataframe(df, use_container_width=True, hide_index=True)

        platform_counts = df[t["col_platform"]].value_counts().reset_index()
        platform_counts.columns = [t["col_platform"], t["col_keyword_count"]]
        fig = px.pie(platform_counts, values=t["col_keyword_count"], names=t["col_platform"],
                     title=t["pie_title"],
                     color_discrete_sequence=["#4f8ef7","#22c55e","#f59e0b","#ef4444","#8b5cf6"])
        fig.update_layout(font_family="IBM Plex Sans Arabic")
        st.plotly_chart(fig, use_container_width=True)

    # Tab 2 — Price
    with tab2:
        price_data = analyze_prices(successful)
        c1, c2, c3 = st.columns(3)
        c1.metric(t["min_label"],  f"{price_data['min']} SAR")
        c2.metric(t["avg_label"],  f"{price_data['average']} SAR")
        c3.metric(t["max_label"],  f"{price_data['max']} SAR")

        st.markdown(f'<div class="section-title">{t["price_compare_title"]}</div>', unsafe_allow_html=True)
        price_df = pd.DataFrame([{
            t["col_product"]: p.get(title_key, p.get("title",""))[:45]+"...",
            t["col_price"]: p.get("price"),
            t["col_position"]: p.get("position",""),
        } for p in price_data["products"] if p.get("price")])

        color_map = {
            "🥇 الأرخص": "#22c55e", "🔴 الأغلى": "#ef4444",
            "✅ أقل من المتوسط": "#4f8ef7", "⚠️ أعلى من المتوسط": "#f59e0b",
            "❓ غير محدد": "#9ca3af",
        }
        fig2 = px.bar(price_df, x=t["col_price"], y=t["col_product"],
                      color=t["col_position"], color_discrete_map=color_map,
                      orientation="h", title=t["price_chart_title"], text=t["col_price"])
        fig2.add_vline(x=price_data["average"], line_dash="dash", line_color="#6366f1",
                       annotation_text=f"{t['avg_label']}: {price_data['average']} SAR")
        fig2.update_layout(font_family="IBM Plex Sans Arabic",
                           height=max(300, len(price_df)*70))
        st.plotly_chart(fig2, use_container_width=True)

        pos_df = pd.DataFrame([{
            t["col_product"]: p.get(title_key, p.get("title",""))[:60],
            t["col_price"]: p.get("price"),
            t["col_vs_avg"]: p.get("vs_avg"),
            t["col_vs_avg_pct"]: f"{p.get('vs_avg_pct',0)}%",
            t["col_position"]: p.get("position",""),
        } for p in price_data["products"]])
        st.dataframe(pos_df, use_container_width=True, hide_index=True)

    # Tab 3 — Reviews
    with tab3:
        rev_data = analyze_reviews(successful)
        ranked   = rev_data["ranked"]
        rev_df   = pd.DataFrame([{
            t["col_product"]: p.get(title_key, p.get("title",""))[:45]+"...",
            t["col_rating"]: p.get("rating"),
            t["col_reviews"]: p.get("reviews_count"),
            t["social_proof"]: p.get("social_proof_score"),
        } for p in ranked if p.get("rating")])

        fig3 = go.Figure()
        fig3.add_trace(go.Bar(name=t["social_proof"],
                              x=rev_df[t["col_product"]], y=rev_df[t["social_proof"]],
                              marker_color="#4f8ef7"))
        fig3.add_trace(go.Scatter(name=t["col_rating"],
                                  x=rev_df[t["col_product"]], y=rev_df[t["col_rating"]],
                                  mode="lines+markers", yaxis="y2",
                                  line=dict(color="#f59e0b", width=3),
                                  marker=dict(size=10)))
        fig3.update_layout(yaxis2=dict(overlaying="y", side="right", title=t["unit_out_of_5"]),
                           font_family="IBM Plex Sans Arabic",
                           title=t["reviews_chart_title"])
        st.plotly_chart(fig3, use_container_width=True)

        col1, col2 = st.columns(2)
        with col1:
            best = rev_data.get("best_rated", {})
            if best:
                st.success(f"{t['best_rated']} {best.get(title_key, best.get('title',''))[:50]} — {best.get('rating',0)}/5")
        with col2:
            most = rev_data.get("most_reviewed", {})
            if most:
                st.info(f"{t['most_reviewed']} {most.get(title_key, most.get('title',''))[:50]} — {most.get('reviews_count',0):,}")

    # Tab 4 — SEO
    with tab4:
        seo_data   = analyze_seo(successful)
        shared_kws = seo_data.get("top_shared_keywords", [])

        if shared_kws:
            st.markdown(f'<div class="section-title">{t["shared_keywords_title"]}</div>', unsafe_allow_html=True)
            kw_df = pd.DataFrame(shared_kws)
            kw_df.columns = [t["col_keyword"], t["col_keyword_count"]]
            fig4  = px.bar(kw_df.head(15), x=t["col_keyword_count"], y=t["col_keyword"], orientation="h",
                           color=t["col_keyword_count"], color_continuous_scale="Blues",
                           title=t["shared_keywords_chart"])
            fig4.update_layout(font_family="IBM Plex Sans Arabic")
            st.plotly_chart(fig4, use_container_width=True)

        st.markdown(f'<div class="section-title">{t["title_length_title"]}</div>', unsafe_allow_html=True)
        tl_df = pd.DataFrame(seo_data.get("title_lengths", []))
        if not tl_df.empty:
            tl_df.columns = [t["col_product"], t["col_chars"]]
            fig5 = px.bar(tl_df, x=t["col_product"], y=t["col_chars"], color=t["col_chars"],
                          color_continuous_scale="Viridis",
                          title=t["title_length_chart"])
            fig5.add_hline(y=80,  line_dash="dot", line_color="green",  annotation_text=t["min_recommend"])
            fig5.add_hline(y=150, line_dash="dot", line_color="red",    annotation_text=t["max_recommend"])
            fig5.update_layout(font_family="IBM Plex Sans Arabic")
            st.plotly_chart(fig5, use_container_width=True)

    # Tab 5 — Opportunities
    with tab5:
        opportunities = find_opportunities(successful)

        # Translate insight messages if English
        insight_map_ar_to_key = {
            "💡 يوجد فارق سعري بين المنتجات — راجع استراتيجية التسعير": "insight_price_gap",
            "🛒 بعض المنتجات غير متوفرة — فرصة لتلبية الطلب الموجود": "insight_unavailable",
            "⚠️ متوسط التقييمات منخفض — فرصة للتميز بجودة أعلى": "insight_low_rating",
            "✅ السوق متوازن — ركز على جودة المنتج وخدمة العملاء": "insight_balanced",
        }
        if lang == "en":
            opportunities = [t.get(insight_map_ar_to_key.get(o, ""), o) for o in opportunities]

        st.markdown(f'<div class="section-title">{t["opportunities_title"]}</div>', unsafe_allow_html=True)
        for opp in opportunities:
            st.markdown(f'<div class="insight-card">{opp}</div>', unsafe_allow_html=True)

        # Radar chart
        if len(successful) >= 2 and prices_ok and ratings_ok:
            st.markdown(f'<div class="section-title">{t["radar_title"]}</div>', unsafe_allow_html=True)
            radar_data = []
            for p in successful:
                if p.get("price") and p.get("rating"):
                    radar_data.append({
                        "name": p.get(title_key, p.get("title",""))[:30],
                        t["radar_categories"][0]: (p.get("rating") or 0) * 20,
                        t["radar_categories"][1]: max(0, 100 - (
                            ((p.get("price",0) - min(prices_ok)) /
                             (max(prices_ok) - min(prices_ok) + 0.01)) * 100)),
                        t["radar_categories"][2]: min(100, ((p.get("reviews_count") or 0) /
                                              max(reviews_ok or [1])) * 100),
                        t["radar_categories"][3]: 100 if p.get("availability") == "متوفر" else 0,
                    })

            if radar_data:
                categories = t["radar_categories"]
                colors     = ["#4f8ef7","#22c55e","#f59e0b","#ef4444","#8b5cf6"]
                fig6       = go.Figure()
                for i, d in enumerate(radar_data):
                    fig6.add_trace(go.Scatterpolar(
                        r=[d[c] for c in categories], theta=categories,
                        fill="toself", name=d["name"],
                        opacity=0.7, line=dict(color=colors[i % len(colors)])
                    ))
                fig6.update_layout(
                    polar=dict(radialaxis=dict(visible=True, range=[0,100])),
                    title=t["radar_chart_title"],
                    font_family="IBM Plex Sans Arabic"
                )
                st.plotly_chart(fig6, use_container_width=True)

    # ── Export ─────────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown(f'<div class="section-title">{t["export_title"]}</div>', unsafe_allow_html=True)
    col_exp1, col_exp2 = st.columns([2, 4])
    with col_exp1:
        price_analysis = analyze_prices(successful)
        rev_analysis   = analyze_reviews(successful)
        seo_analysis   = analyze_seo(successful)
        opps           = find_opportunities(successful)
        excel_bytes    = export_to_excel(successful, price_analysis, rev_analysis, seo_analysis, opps)
        filename       = f"CompeteIQ_Demo_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        st.download_button(
            label=t["export_btn"],
            data=excel_bytes, file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", use_container_width=True,
        )
    with col_exp2:
        st.caption(t["export_caption"])

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    f"<div style='text-align:center;color:#9ca3af;font-size:0.8rem;'>"
    f"{t['footer']}"
    f"</div>", unsafe_allow_html=True
)
