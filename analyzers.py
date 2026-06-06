"""
Analyzers — Demo Version
نسخة تجريبية مبسّطة | النسخة الكاملة متوفرة عند الشراء
"""
from __future__ import annotations
from typing import Any


# ── Price Analysis ──────────────────────────────────────────────────────────

def analyze_prices(products: list[dict]) -> dict[str, Any]:
    prices = [p["price"] for p in products if p.get("price") is not None]
    if not prices:
        return {"error": "No price data available"}

    avg      = round(sum(prices) / len(prices), 2)
    min_p    = min(prices)
    max_p    = max(prices)
    currency = next((p.get("currency", "SAR") for p in products if p.get("price")), "SAR")

    enriched = []
    for p in products:
        entry = {**p}
        if p.get("price") is not None:
            if p["price"] == min_p:
                entry["position"] = "🥇 الأرخص"
            elif p["price"] == max_p:
                entry["position"] = "🔴 الأغلى"
            elif p["price"] < avg:
                entry["position"] = "✅ أقل من المتوسط"
            else:
                entry["position"] = "⚠️ أعلى من المتوسط"
            entry["vs_avg"]     = round(p["price"] - avg, 2)
            entry["vs_avg_pct"] = round(((p["price"] - avg) / avg) * 100, 1)
        else:
            entry["position"]   = "❓ غير محدد"
            entry["vs_avg"]     = None
            entry["vs_avg_pct"] = None
        enriched.append(entry)

    return {
        "average":  avg,
        "min":      min_p,
        "max":      max_p,
        "spread":   round(max_p - min_p, 2),
        "currency": currency,
        "products": enriched,
    }


# ── Review / Rating Analysis ────────────────────────────────────────────────

def analyze_reviews(products: list[dict]) -> dict[str, Any]:
    scored = []
    for p in products:
        # Demo: درجة ثقة مبسّطة = التقييم × 10
        score = round((p.get("rating") or 0) * 10, 1)
        scored.append({**p, "social_proof_score": score})

    scored.sort(key=lambda x: x["social_proof_score"], reverse=True)

    ratings = [p["rating"] for p in products if p.get("rating")]
    reviews = [p["reviews_count"] for p in products if p.get("reviews_count")]

    return {
        "ranked":        scored,
        "avg_rating":    round(sum(ratings) / len(ratings), 2) if ratings else None,
        "total_reviews": sum(reviews) if reviews else 0,
        "best_rated":    max(scored, key=lambda x: x.get("rating") or 0, default={}),
        "most_reviewed": max(scored, key=lambda x: x.get("reviews_count") or 0, default={}),
    }


# ── SEO / Keyword Analysis ──────────────────────────────────────────────────

def analyze_seo(products: list[dict]) -> dict[str, Any]:
    # Demo: نعرض الكلمات المفتاحية كما هي بدون تحليل تكراري
    all_keywords: list[str] = []
    for p in products:
        all_keywords.extend(p.get("keywords", []))

    seen  = {}
    for kw in all_keywords:
        seen[kw] = seen.get(kw, 0) + 1

    top_shared = [
        {"keyword": kw, "count": cnt}
        for kw, cnt in sorted(seen.items(), key=lambda x: -x[1])
        if cnt >= 2
    ][:15]

    title_lengths = [
        {"title": p.get("title", "")[:50], "length": len(p.get("title", ""))}
        for p in products
    ]

    return {
        "top_shared_keywords":     top_shared,
        "unique_keywords":         {},
        "title_lengths":           title_lengths,
        "total_keywords_analyzed": len(all_keywords),
    }


# ── Opportunity Finder ──────────────────────────────────────────────────────

def find_opportunities(products: list[dict]) -> list[str]:
    # Demo: رسائل ثابتة بناءً على البيانات الظاهرة فقط
    insights = []

    prices = [p["price"] for p in products if p.get("price")]
    if prices and (max(prices) - min(prices)) > min(prices) * 0.1:
        insights.append(
            f"💡 يوجد فارق سعري بين المنتجات — راجع استراتيجية التسعير"
        )

    unavailable = [p for p in products if p.get("availability") == "غير متوفر"]
    if unavailable:
        insights.append("🛒 بعض المنتجات غير متوفرة — فرصة لتلبية الطلب الموجود")

    ratings = [p["rating"] for p in products if p.get("rating")]
    if ratings and (sum(ratings) / len(ratings)) < 4.2:
        insights.append("⚠️ متوسط التقييمات منخفض — فرصة للتميز بجودة أعلى")

    if not insights:
        insights.append("✅ السوق متوازن — ركز على جودة المنتج وخدمة العملاء")

    return insights
