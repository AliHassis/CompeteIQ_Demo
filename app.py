"""
CompeteIQ DEMO — لوحة تحليل المنافسين
بيانات تجريبية مدمجة مباشرة — لا يحتاج scraping ولا ملفات خارجية
"""
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import time
import io

from analyzers import analyze_prices, analyze_reviews, analyze_seo, find_opportunities
from excel_exporter import export_to_excel

# ── Page Config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="CompeteIQ — تحليل المنافسين",
    page_icon="🕵️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Arabic:wght@300;400;500;600;700&display=swap');
  html, body, [class*="css"] { font-family: 'IBM Plex Sans Arabic', sans-serif; direction: rtl; }

  .main-header {
    background: linear-gradient(135deg, #1a1f3a 0%, #2d3875 50%, #1a1f3a 100%);
    border-radius: 16px; padding: 2rem 2.5rem; margin-bottom: 2rem;
    box-shadow: 0 8px 32px rgba(26,31,58,0.4); text-align: center;
  }
  .main-header h1 { color: #fff; font-size: 2.4rem; font-weight: 700; margin: 0; }
  .main-header p  { color: #a5b4fc; font-size: 1.05rem; margin: 0.5rem 0 0; }

  .demo-badge {
    background: linear-gradient(90deg, #f59e0b, #ef4444);
    color: white; border-radius: 20px; padding: 0.3rem 1rem;
    font-size: 0.85rem; font-weight: 700; display: inline-block;
    margin-bottom: 1rem;
  }

  .metric-card {
    background: #fff; border: 1px solid #e8ecf8; border-radius: 14px;
    padding: 1.2rem 1.5rem; text-align: center;
    box-shadow: 0 2px 12px rgba(0,0,0,0.06);
    transition: transform 0.2s, box-shadow 0.2s;
  }
  .metric-card:hover { transform: translateY(-3px); box-shadow: 0 8px 24px rgba(79,142,247,0.15); }
  .metric-card .label { color: #6b7280; font-size: 0.85rem; font-weight: 500; margin-bottom: 0.4rem; }
  .metric-card .value { color: #1a1f3a; font-size: 1.9rem; font-weight: 700; }
  .metric-card .sub   { color: #9ca3af; font-size: 0.78rem; margin-top: 0.2rem; }

  .section-title {
    font-size: 1.25rem; font-weight: 700; color: #1a1f3a;
    border-right: 4px solid #4f8ef7; padding-right: 0.75rem; margin: 1.5rem 0 1rem;
  }
  .insight-card {
    background: linear-gradient(135deg, #f0f4ff, #fff);
    border: 1px solid #c7d7fd; border-radius: 12px;
    padding: 1rem 1.25rem; margin-bottom: 0.75rem;
    font-size: 0.95rem; color: #1e293b; line-height: 1.6;
  }
  .stTabs [data-baseweb="tab"] { font-size: 0.95rem; font-weight: 600; }
  .stButton > button { border-radius: 10px; font-weight: 600; font-size: 1rem; padding: 0.6rem 2rem; }
  [data-testid="stSidebar"] { background: #f8faff; border-left: 1px solid #e8ecf8; }
</style>
""", unsafe_allow_html=True)

# ── البيانات التجريبية المدمجة ────────────────────────────────────────────────
# ثلاثة سيناريوهات — كل رابط يعطي مجموعة بيانات مختلفة
DEMO_SCENARIOS = {
    "https://www.amazon.sa/dp/B0CKLKNBPY": [
        {
            "platform": "Amazon",
            "url": "https://www.amazon.sa/dp/B0CKLKNBPY",
            "title": "سامسونج جالاكسي S24 الترا 256 جيجا تيتانيوم أسود",
            "price": 4299.0, "currency": "SAR", "rating": 4.6,
            "reviews_count": 2847, "availability": "متوفر",
            "seller": "Amazon.sa", "keywords": ["سامسونج", "جالاكسي", "s24", "الترا", "هاتف", "ذكي"],
            "error": None,
        },
        {
            "platform": "Noon",
            "url": "https://www.noon.com/ar-sa/samsung-s24",
            "title": "Samsung Galaxy S24 Ultra 256GB Titanium Black",
            "price": 4199.0, "currency": "SAR", "rating": 4.4,
            "reviews_count": 1203, "availability": "متوفر",
            "seller": "Noon", "keywords": ["samsung", "galaxy", "s24", "ultra", "smartphone"],
            "error": None,
        },
        {
            "platform": "Jarir",
            "url": "https://www.jarir.com/samsung-galaxy-s24",
            "title": "سامسونج جالاكسي S24 الترا - تيتانيوم أسود 256GB",
            "price": 4499.0, "currency": "SAR", "rating": 4.5,
            "reviews_count": 634, "availability": "متوفر",
            "seller": "جرير", "keywords": ["سامسونج", "s24", "الترا", "256", "تيتانيوم"],
            "error": None,
        },
    ],
    "https://www.amazon.sa/dp/B0FQFV3HJ2": [
        {
            "platform": "Amazon",
            "url": "https://www.amazon.sa/dp/B0FQFV3HJ2",
            "title": "آيفون 15 برو ماكس 256 جيجا - تيتانيوم طبيعي",
            "price": 5299.0, "currency": "SAR", "rating": 4.8,
            "reviews_count": 5621, "availability": "متوفر",
            "seller": "Apple Store", "keywords": ["آيفون", "15", "برو", "ماكس", "تيتانيوم", "أبل"],
            "error": None,
        },
        {
            "platform": "Extra",
            "url": "https://www.extra.com/iphone-15-pro-max",
            "title": "Apple iPhone 15 Pro Max 256GB Natural Titanium",
            "price": 5199.0, "currency": "SAR", "rating": 4.7,
            "reviews_count": 892, "availability": "متوفر",
            "seller": "Extra", "keywords": ["apple", "iphone", "15", "pro", "max", "titanium"],
            "error": None,
        },
        {
            "platform": "Noon",
            "url": "https://www.noon.com/ar-sa/iphone-15-pro-max",
            "title": "iPhone 15 Pro Max 256GB - Natural Titanium",
            "price": 5399.0, "currency": "SAR", "rating": 4.6,
            "reviews_count": 1544, "availability": "غير متوفر",
            "seller": "Noon", "keywords": ["iphone", "15", "pro", "max", "256", "natural"],
            "error": None,
        },
    ],
    "https://www.amazon.sa/dp/B0FQFYMLS5": [
        {
            "platform": "Amazon",
            "url": "https://www.amazon.sa/dp/B0FQFYMLS5",
            "title": "ماك بوك برو 14 إنش M3 - 8GB RAM 512GB SSD",
            "price": 6999.0, "currency": "SAR", "rating": 4.9,
            "reviews_count": 3102, "availability": "متوفر",
            "seller": "Amazon.sa", "keywords": ["ماك", "بوك", "برو", "m3", "لابتوب", "أبل"],
            "error": None,
        },
        {
            "platform": "Jarir",
            "url": "https://www.jarir.com/macbook-pro-14",
            "title": "MacBook Pro 14-inch M3 Chip 8GB 512GB - Space Gray",
            "price": 7199.0, "currency": "SAR", "rating": 4.8,
            "reviews_count": 421, "availability": "متوفر",
            "seller": "جرير", "keywords": ["macbook", "pro", "14", "m3", "512", "apple"],
            "error": None,
        },
        {
            "platform": "Extra",
            "url": "https://www.extra.com/macbook-pro-14-m3",
            "title": "Apple MacBook Pro 14\" M3 8GB 512GB Space Gray",
            "price": 6899.0, "currency": "SAR", "rating": 4.7,
            "reviews_count": 267, "availability": "متوفر",
            "seller": "Extra", "keywords": ["apple", "macbook", "pro", "m3", "laptop", "512gb"],
            "error": None,
        },
    ],
}

# الروابط الافتراضية للعرض
DEFAULT_URLS = "\n".join(DEMO_SCENARIOS.keys())

# ── Session State ─────────────────────────────────────────────────────────────
if "results" not in st.session_state:
    st.session_state.results = []
if "scraped" not in st.session_state:
    st.session_state.scraped = False

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1>🕵️ CompeteIQ</h1>
  <p>أداة تحليل المنافسين للمتاجر الإلكترونية — أسعار · تقييمات · SEO · فرص تنافسية</p>
</div>
""", unsafe_allow_html=True)

st.markdown('<div style="text-align:center"><span class="demo-badge">🎯 نسخة تجريبية — Demo Version</span></div>', unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ إعدادات الأداة")
    st.markdown("---")

    urls_input = st.text_area(
        "🔗 روابط المنتجات (رابط واحد لكل سطر)",
        value=DEFAULT_URLS,
        height=180,
        help="أدخل الروابط التجريبية أو جرب تغييرها."
    )

    st.markdown("---")
    st.info("💡 هذه نسخة تجريبية تعرض بيانات نموذجية.\nالنسخة الكاملة تستخرج بيانات حقيقية من أي متجر إلكتروني.")

    st.markdown("**📋 المنصات المدعومة في النسخة الكاملة:**")
    for p in ["✅ Amazon (.sa, .ae, .com)", "✅ Noon", "✅ Jarir", "✅ Extra", "✅ أي متجر إلكتروني"]:
        st.caption(p)

    st.markdown("---")
    analyze_btn = st.button("🚀 ابدأ التحليل", type="primary", use_container_width=True)

    if st.session_state.scraped:
        if st.button("🔄 تحليل جديد", use_container_width=True):
            st.session_state.results = []
            st.session_state.scraped = False
            st.rerun()

# ── Main Logic ────────────────────────────────────────────────────────────────
if analyze_btn:
    urls = [u.strip() for u in urls_input.strip().splitlines() if u.strip()]

    if not urls:
        st.warning("⚠️ أدخل رابطاً واحداً على الأقل.")
    else:
        results = []
        progress_bar = st.progress(0, text="جارٍ تحليل المنتجات...")

        for i, url in enumerate(urls):
            # محاكاة تحليل حقيقي
            for pct in range(0, 34):
                time.sleep(0.01)
                progress_bar.progress(
                    int((i / len(urls)) * 100 + pct / 3),
                    text=f"🔍 جارٍ تحليل المنتج {i+1} من {len(urls)}..."
                )

            # جلب البيانات التجريبية
            if url in DEMO_SCENARIOS:
                results.extend(DEMO_SCENARIOS[url])
            else:
                # أي رابط آخر يعطي نفس بيانات أول سيناريو
                results.extend(DEMO_SCENARIOS[list(DEMO_SCENARIOS.keys())[0]])

        progress_bar.progress(100, text="✅ اكتمل التحليل!")
        time.sleep(0.5)

        st.session_state.results = results
        st.session_state.scraped = True
        st.rerun()

# ── Dashboard ─────────────────────────────────────────────────────────────────
if st.session_state.scraped and st.session_state.results:
    products  = st.session_state.results
    successful = [p for p in products if not p.get("error")]

    prices_ok  = [p["price"] for p in successful if p.get("price")]
    ratings_ok = [p["rating"] for p in successful if p.get("rating")]
    reviews_ok = [p["reviews_count"] for p in successful if p.get("reviews_count")]

    # ── KPI Cards ──────────────────────────────────────────────────────────
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.markdown(f"""<div class="metric-card">
            <div class="label">📦 المنتجات المُحللة</div>
            <div class="value">{len(successful)}</div>
            <div class="sub">منتج</div></div>""", unsafe_allow_html=True)
    with c2:
        avg_p = round(sum(prices_ok)/len(prices_ok), 1) if prices_ok else "N/A"
        st.markdown(f"""<div class="metric-card">
            <div class="label">💰 متوسط السعر</div>
            <div class="value">{avg_p}</div>
            <div class="sub">ريال سعودي</div></div>""", unsafe_allow_html=True)
    with c3:
        avg_r = round(sum(ratings_ok)/len(ratings_ok), 2) if ratings_ok else "N/A"
        st.markdown(f"""<div class="metric-card">
            <div class="label">⭐ متوسط التقييم</div>
            <div class="value">{avg_r}</div>
            <div class="sub">من 5</div></div>""", unsafe_allow_html=True)
    with c4:
        total_rev = sum(reviews_ok) if reviews_ok else 0
        st.markdown(f"""<div class="metric-card">
            <div class="label">💬 إجمالي المراجعات</div>
            <div class="value">{total_rev:,}</div>
            <div class="sub">مراجعة</div></div>""", unsafe_allow_html=True)
    with c5:
        avail = sum(1 for p in successful if p.get("availability") == "متوفر")
        st.markdown(f"""<div class="metric-card">
            <div class="label">🛒 متوفر للشراء</div>
            <div class="value">{avail}/{len(successful)}</div>
            <div class="sub">منتج</div></div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Tabs ───────────────────────────────────────────────────────────────
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📋 نظرة عامة", "💰 تحليل الأسعار",
        "⭐ التقييمات", "🔍 SEO", "💡 الفرص"
    ])

    # Tab 1 — Overview
    with tab1:
        st.markdown('<div class="section-title">جميع المنتجات المُحللة</div>', unsafe_allow_html=True)
        df = pd.DataFrame([{
            "المنصة": p.get("platform",""),
            "العنوان": p.get("title","")[:70],
            "السعر (SAR)": p.get("price"),
            "التقييم ⭐": p.get("rating"),
            "المراجعات 💬": p.get("reviews_count"),
            "التوفر": p.get("availability",""),
            "البائع": p.get("seller",""),
        } for p in successful])
        st.dataframe(df, use_container_width=True, hide_index=True)

        platform_counts = df["المنصة"].value_counts().reset_index()
        platform_counts.columns = ["المنصة", "العدد"]
        fig = px.pie(platform_counts, values="العدد", names="المنصة",
                     title="توزيع المنتجات حسب المنصة",
                     color_discrete_sequence=["#4f8ef7","#22c55e","#f59e0b","#ef4444","#8b5cf6"])
        fig.update_layout(font_family="IBM Plex Sans Arabic")
        st.plotly_chart(fig, use_container_width=True)

    # Tab 2 — Price
    with tab2:
        price_data = analyze_prices(successful)
        c1, c2, c3 = st.columns(3)
        c1.metric("الأدنى",   f"{price_data['min']} SAR")
        c2.metric("المتوسط",  f"{price_data['average']} SAR")
        c3.metric("الأعلى",   f"{price_data['max']} SAR")

        st.markdown('<div class="section-title">مقارنة الأسعار</div>', unsafe_allow_html=True)
        price_df = pd.DataFrame([{
            "المنتج": p.get("title","")[:45]+"...",
            "السعر": p.get("price"),
            "الموقع التنافسي": p.get("position",""),
        } for p in price_data["products"] if p.get("price")])

        color_map = {
            "🥇 الأرخص": "#22c55e", "🔴 الأغلى": "#ef4444",
            "✅ أقل من المتوسط": "#4f8ef7", "⚠️ أعلى من المتوسط": "#f59e0b",
            "❓ غير محدد": "#9ca3af",
        }
        fig2 = px.bar(price_df, x="السعر", y="المنتج",
                      color="الموقع التنافسي", color_discrete_map=color_map,
                      orientation="h", title="مقارنة الأسعار مع المتوسط", text="السعر")
        fig2.add_vline(x=price_data["average"], line_dash="dash", line_color="#6366f1",
                       annotation_text=f"المتوسط: {price_data['average']} SAR")
        fig2.update_layout(font_family="IBM Plex Sans Arabic",
                           height=max(300, len(price_df)*70))
        st.plotly_chart(fig2, use_container_width=True)

        pos_df = pd.DataFrame([{
            "المنتج": p.get("title","")[:60],
            "السعر": p.get("price"),
            "الفرق عن المتوسط": p.get("vs_avg"),
            "الفرق %": f"{p.get('vs_avg_pct',0)}%",
            "الموقع التنافسي": p.get("position",""),
        } for p in price_data["products"]])
        st.dataframe(pos_df, use_container_width=True, hide_index=True)

    # Tab 3 — Reviews
    with tab3:
        rev_data = analyze_reviews(successful)
        ranked   = rev_data["ranked"]
        rev_df   = pd.DataFrame([{
            "المنتج": p.get("title","")[:45]+"...",
            "التقييم": p.get("rating"),
            "المراجعات": p.get("reviews_count"),
            "درجة الثقة": p.get("social_proof_score"),
        } for p in ranked if p.get("rating")])

        fig3 = go.Figure()
        fig3.add_trace(go.Bar(name="درجة الثقة الاجتماعية",
                              x=rev_df["المنتج"], y=rev_df["درجة الثقة"],
                              marker_color="#4f8ef7"))
        fig3.add_trace(go.Scatter(name="التقييم ⭐",
                                  x=rev_df["المنتج"], y=rev_df["التقييم"],
                                  mode="lines+markers", yaxis="y2",
                                  line=dict(color="#f59e0b", width=3),
                                  marker=dict(size=10)))
        fig3.update_layout(yaxis2=dict(overlaying="y", side="right", title="التقييم (5)"),
                           font_family="IBM Plex Sans Arabic",
                           title="مقارنة التقييمات ودرجة الثقة")
        st.plotly_chart(fig3, use_container_width=True)

        col1, col2 = st.columns(2)
        with col1:
            best = rev_data.get("best_rated", {})
            if best:
                st.success(f"**⭐ الأعلى تقييماً:** {best.get('title','')[:50]} — {best.get('rating',0)}/5")
        with col2:
            most = rev_data.get("most_reviewed", {})
            if most:
                st.info(f"**💬 الأكثر مراجعات:** {most.get('title','')[:50]} — {most.get('reviews_count',0):,} مراجعة")

    # Tab 4 — SEO
    with tab4:
        seo_data   = analyze_seo(successful)
        shared_kws = seo_data.get("top_shared_keywords", [])

        if shared_kws:
            st.markdown('<div class="section-title">الكلمات المفتاحية المشتركة</div>', unsafe_allow_html=True)
            kw_df = pd.DataFrame(shared_kws)
            fig4  = px.bar(kw_df.head(15), x="count", y="keyword", orientation="h",
                           color="count", color_continuous_scale="Blues",
                           title="أكثر الكلمات المفتاحية استخداماً بين المنافسين",
                           labels={"count":"عدد المنتجات","keyword":"الكلمة المفتاحية"})
            fig4.update_layout(font_family="IBM Plex Sans Arabic")
            st.plotly_chart(fig4, use_container_width=True)

        st.markdown('<div class="section-title">طول عناوين المنتجات</div>', unsafe_allow_html=True)
        tl_df = pd.DataFrame(seo_data.get("title_lengths", []))
        if not tl_df.empty:
            fig5 = px.bar(tl_df, x="title", y="length", color="length",
                          color_continuous_scale="Viridis",
                          title="طول العنوان (التوصية: 80-150 حرف)",
                          labels={"length":"عدد الأحرف","title":"المنتج"})
            fig5.add_hline(y=80,  line_dash="dot", line_color="green",  annotation_text="الحد الأدنى (80)")
            fig5.add_hline(y=150, line_dash="dot", line_color="red",    annotation_text="الحد الأقصى (150)")
            fig5.update_layout(font_family="IBM Plex Sans Arabic")
            st.plotly_chart(fig5, use_container_width=True)

    # Tab 5 — Opportunities
    with tab5:
        opportunities = find_opportunities(successful)
        st.markdown('<div class="section-title">الفرص التنافسية المكتشفة</div>', unsafe_allow_html=True)
        for opp in opportunities:
            st.markdown(f'<div class="insight-card">{opp}</div>', unsafe_allow_html=True)

        # Radar chart
        if len(successful) >= 2 and prices_ok and ratings_ok:
            st.markdown('<div class="section-title">الخريطة التنافسية</div>', unsafe_allow_html=True)
            radar_data = []
            for p in successful:
                if p.get("price") and p.get("rating"):
                    radar_data.append({
                        "المنتج": p.get("title","")[:30],
                        "التقييم": (p.get("rating") or 0) * 20,
                        "القيمة مقابل السعر": max(0, 100 - (
                            ((p.get("price",0) - min(prices_ok)) /
                             (max(prices_ok) - min(prices_ok) + 0.01)) * 100)),
                        "المراجعات": min(100, ((p.get("reviews_count") or 0) /
                                              max(reviews_ok or [1])) * 100),
                        "التوفر": 100 if p.get("availability") == "متوفر" else 0,
                    })

            if radar_data:
                categories = ["التقييم","القيمة مقابل السعر","المراجعات","التوفر"]
                colors     = ["#4f8ef7","#22c55e","#f59e0b","#ef4444","#8b5cf6"]
                fig6       = go.Figure()
                for i, d in enumerate(radar_data):
                    fig6.add_trace(go.Scatterpolar(
                        r=[d[c] for c in categories], theta=categories,
                        fill="toself", name=d["المنتج"],
                        opacity=0.7, line=dict(color=colors[i % len(colors)])
                    ))
                fig6.update_layout(
                    polar=dict(radialaxis=dict(visible=True, range=[0,100])),
                    title="الخريطة التنافسية متعددة الأبعاد",
                    font_family="IBM Plex Sans Arabic"
                )
                st.plotly_chart(fig6, use_container_width=True)

    # ── Export ─────────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown('<div class="section-title">📥 تصدير التقرير</div>', unsafe_allow_html=True)
    col_exp1, col_exp2 = st.columns([2, 4])
    with col_exp1:
        price_analysis = analyze_prices(successful)
        rev_analysis   = analyze_reviews(successful)
        seo_analysis   = analyze_seo(successful)
        opps           = find_opportunities(successful)
        excel_bytes    = export_to_excel(successful, price_analysis, rev_analysis, seo_analysis, opps)
        filename       = f"CompeteIQ_Demo_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        st.download_button(
            label="⬇️ تحميل تقرير Excel الكامل",
            data=excel_bytes, file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", use_container_width=True,
        )
    with col_exp2:
        st.caption("التقرير يتضمن: نظرة عامة · تحليل الأسعار · التقييمات · SEO · الفرص التنافسية")

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    "<div style='text-align:center;color:#9ca3af;font-size:0.8rem;'>"
    "CompeteIQ Demo · تحليل المنافسين للمتاجر الإلكترونية"
    "</div>", unsafe_allow_html=True
)
