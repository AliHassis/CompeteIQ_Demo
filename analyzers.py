"""
Analyzers for price, reviews, and SEO keywords.
All functions operate on a list of product dicts returned by scrapers.
"""
from __future__ import annotations
import re
from collections import Counter
from typing import Any


# ── Price Analysis ──────────────────────────────────────────────────────────

def analyze_prices(products: list[dict]) -> dict[str, Any]:
    """
    Returns price stats and competitive position for each product.
    """
    prices = [p["price"] for p in products if p.get("price") is not None]
    if not prices:
        return {"error": "No price data available"}

    avg = sum(prices) / len(prices)
    min_p = min(prices)
    max_p = max(prices)
    spread = max_p - min_p
    currency = next((p.get("currency", "N/A") for p in products if p.get("price")), "N/A")

    enriched = []
    for p in products:
        entry = {**p}
        if p.get("price") is not None:
            diff = p["price"] - avg
            entry["vs_avg"] = round(diff, 2)
            entry["vs_avg_pct"] = round((diff / avg) * 100, 1) if avg else 0
            if p["price"] == min_p:
                entry["position"] = "🥇 الأرخص"
            elif p["price"] == max_p:
                entry["position"] = "🔴 الأغلى"
            elif p["price"] < avg:
                entry["position"] = "✅ أقل من المتوسط"
            else:
                entry["position"] = "⚠️ أعلى من المتوسط"
        else:
            entry["vs_avg"] = None
            entry["vs_avg_pct"] = None
            entry["position"] = "❓ غير محدد"
        enriched.append(entry)

    return {
        "average": round(avg, 2),
        "min": min_p,
        "max": max_p,
        "spread": round(spread, 2),
        "currency": currency,
        "products": enriched,
    }


# ── Review / Rating Analysis ────────────────────────────────────────────────

def analyze_reviews(products: list[dict]) -> dict[str, Any]:
    """
    Ranks products by social proof (rating × log(reviews)).
    """
    import math

    scored = []
    for p in products:
        rating = p.get("rating") or 0
        reviews = p.get("reviews_count") or 0
        # Social proof score: weighted combination
        score = rating * math.log1p(reviews)
        scored.append({**p, "social_proof_score": round(score, 3)})

    scored.sort(key=lambda x: x["social_proof_score"], reverse=True)

    ratings = [p["rating"] for p in products if p.get("rating")]
    reviews = [p["reviews_count"] for p in products if p.get("reviews_count")]

    return {
        "ranked": scored,
        "avg_rating": round(sum(ratings) / len(ratings), 2) if ratings else None,
        "total_reviews": sum(reviews) if reviews else 0,
        "best_rated": max(scored, key=lambda x: x.get("rating") or 0, default={}),
        "most_reviewed": max(scored, key=lambda x: x.get("reviews_count") or 0, default={}),
    }


# ── SEO / Keyword Analysis ──────────────────────────────────────────────────

def analyze_seo(products: list[dict]) -> dict[str, Any]:
    """
    Extracts common and unique keywords across all products,
    plus title length and meta-keyword density stats.
    """
    all_keywords: list[str] = []
    for p in products:
        all_keywords.extend(p.get("keywords", []))

    # Normalize
    normalized = [re.sub(r"[^\w\s]", "", kw).strip().lower() for kw in all_keywords]
    normalized = [k for k in normalized if len(k) > 2]

    freq = Counter(normalized)
    top_shared = [
        {"keyword": kw, "count": cnt, "found_in": cnt}
        for kw, cnt in freq.most_common(20)
        if cnt >= 2  # shared by ≥2 products = competitive keyword
    ]
    unique_per_product = {}
    for p in products:
        title = p.get("title", "")
        own_kws = set(re.sub(r"[^\w\s]", "", k).strip().lower() for k in p.get("keywords", []))
        others_kws = set(
            re.sub(r"[^\w\s]", "", k).strip().lower()
            for op in products if op is not p
            for k in op.get("keywords", [])
        )
        unique_per_product[title[:40]] = list(own_kws - others_kws)[:5]

    title_lengths = [
        {"title": p.get("title", "")[:50], "length": len(p.get("title", ""))}
        for p in products
    ]

    return {
        "top_shared_keywords": top_shared,
        "unique_keywords": unique_per_product,
        "title_lengths": title_lengths,
        "total_keywords_analyzed": len(normalized),
    }


# ── Opportunity Finder ──────────────────────────────────────────────────────

def find_opportunities(products: list[dict]) -> list[str]:
    """
    Returns actionable insights based on the collected data.
    """
    insights = []
    prices = [p["price"] for p in products if p.get("price")]
    ratings = [p["rating"] for p in products if p.get("rating")]
    reviews = [p["reviews_count"] for p in products if p.get("reviews_count")]

    if prices:
        spread = max(prices) - min(prices)
        if spread > max(prices) * 0.15:
            insights.append(
                f"💡 فجوة سعرية كبيرة ({round(spread, 1)}) — فرصة لتسعير تنافسي بين الأرخص والمتوسط"
            )

    if ratings:
        avg_r = sum(ratings) / len(ratings)
        if avg_r < 4.0:
            insights.append(
                f"⚠️ متوسط التقييمات منخفض ({round(avg_r, 1)}/5) — المنافسون ضعفاء في الجودة، فرصة للتميز"
            )

    if reviews:
        low_review_products = [p for p in products if (p.get("reviews_count") or 0) < 50]
        if len(low_review_products) > len(products) / 2:
            insights.append(
                "📈 أغلب المنتجات لديها مراجعات قليلة — السوق غير مشبع، فرصة للدخول"
            )

    unavailable = [p for p in products if p.get("availability") == "غير متوفر"]
    if unavailable:
        names = ", ".join(p.get("title", "")[:30] for p in unavailable[:2])
        insights.append(f"🛒 منتجات غير متوفرة: {names} — طلب موجود بدون عرض")

    if not insights:
        insights.append("✅ السوق تنافسي ومتوازن — ركز على جودة المنتج وخدمة العملاء للتميز")

    return insights
