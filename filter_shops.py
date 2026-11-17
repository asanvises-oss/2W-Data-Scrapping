#!/usr/bin/env python3
"""
Thailand Motorcycle Tire Shop Cleaner & Classifier
- Deduplicates Google Maps scraper results (CID / Data ID / Link / Title+Address)
- Filters 2-Wheeler (Motorcycle) tire shops vs 4-Wheeler (Car/Truck) tire shops using:
  1. Business Category rules
  2. Shop Title keyword rules
  3. Customer Review text & Description keyword rules + Tire Brand mention extraction
  4. Optional Stage-2 Gemini Vision AI check on storefront thumbnail for ambiguous shops
- Assigns shop tiers (Tier 1 / Tier 2 / Tier 3)
- Exports UTF-8-SIG CSV ready for both Microsoft Excel and GitHub table preview
"""

import argparse
import base64
import json
import os
import re
import time
import urllib.request
import urllib.error
import pandas as pd


# --- KEYWORD DICTIONARIES ---
POS_2W_KEYWORDS = [
    "มอเตอร์ไซค์", "มอไซค์", "จยย", "จักรยานยนต์", "บิ๊กไบค์", "bigbike", "big bike",
    "moto", "motorcycle", "rider", "biker", "scooter", "vespa", "เวสป้า",
    "lambretta", "แลมเบรตต้า", "วิบาก", "motocross", "2 ล้อ", "สองล้อ",
    "ยางควิก", "ปะยางมอไซค์", "เปลี่ยนยางมอไซค์", "ซ่อมมอไซค์", "อะไหล่มอไซค์",
    "irc", "ไออาร์ซี", "metzeler", "เมทเซเลอร์", "battlax", "diablo", "rosso",
    "city grip", "pilot street", "road 5", "road 6", "power gp", "anakee",
    "camel", "คาเมล", "nd rubber", "เอ็นดี"
]

NEG_4W_KEYWORDS = [
    "รถยนต์", "รถเก๋ง", "รถกระบะ", "กระบะ", "ออฟโรด", "offroad", "4x4", "4wd",
    "suv", "รถบรรทุก", "สิบล้อ", "หกล้อ", "รถไถ", "รถตู้",
    "b-quik", "บี-ควิก", "บีควิก", "cockpit", "ค็อกพิท", "tyreplus", "ไทร์พลัส",
    "a.c.t", "autobacs", "ออโต้แบคส์", "fit auto", "ฟิต ออโต้", "ก.เจริญยางยนต์",
    "yokohama", "โยโกฮาม่า", "toyo", "โตโย", "falken", "ฟอลเคน", "hankook", "ฮันกุก",
    "kumho", "คัมโฮ", "goodyear", "กู๊ดเยียร์", "bfgoodrich", "otani", "โอตานิ",
    "linglong", "sailun", "nexen", "firestone",
    "ตั้งศูนย์", "ล้อแม็ก", "ซ่อมแม็ก", "ช่วงล่าง", "โช๊คอัพ", "คาร์แคร์", "car care",
    "ประดับยนต์", "ฟิล์มกรองแสง", "แอร์รถยนต์", "แบตเตอรี่รถยนต์"
]

REVIEW_2W_KEYWORDS = [
    "มอเตอร์ไซค์", "มอไซค์", "จยย", "บิ๊กไบค์", "bigbike", "scooter",
    "เปลี่ยนยางมอไซค์", "เปลี่ยนยางมอเตอร์ไซค์", "ปะยางมอไซค์", "ปะยางรถเล็ก",
    "ตั้งโซ่", "สเตอร์", "ถ่วงล้อยางมอไซค์", "ยางนอก", "ยางใน",
    "wave", "เวฟ", "pcx", "nmax", "xmax", "forza", "aerox", "click", "คลิก",
    "scoopy", "fino", "grand filano", "giorno", "lead", "adv", "cbr", "ninja",
    "r15", "r3", "mt-03", "mt-07", "cb500", "cb650", "vespa", "เวสป้า", "msx", "monkey"
]

REVIEW_4W_KEYWORDS = [
    "ตั้งศูนย์ถ่วงล้อ", "ตั้งศูนย์", "สลับยาง", "ยางกระบะ", "รถเก๋ง",
    "ช่วงล่าง", "ปะยางรถยนต์", "เปลี่ยนยางรถยนต์", "รถsuv", "รถตู้"
]

TIRE_BRANDS = {
    "Michelin": ["michelin", "มิชลิน", "city grip", "pilot street", "road 5", "road 6", "power 5", "power 6", "anakee"],
    "Pirelli": ["pirelli", "พิเรลลี่", "diablo", "rosso", "angel scooter", "angel gt"],
    "IRC": ["irc", "ไออาร์ซี", "iz-s", "wing"],
    "Bridgestone": ["bridgestone", "บริดจสโตน", "battlax"],
    "Dunlop": ["dunlop", "ดันลอป", "tt900", "sportmax"],
    "Maxxis": ["maxxis", "แม็กซิส"],
    "Metzeler": ["metzeler", "เมทเซเลอร์", "roadtec", "sportec"],
    "Continental": ["continental", "คอนติเนนทอล", "contiscoot", "contiroad"],
    "Deestone": ["deestone", "ดีสโตน"],
    "Camel": ["camel", "คาเมล"],
    "ND Rubber": ["nd rubber", "เอ็นดี"]
}

PROVINCES_77 = [
    "กรุงเทพมหานคร", "กระบี่", "กาญจนบุรี", "กาฬสินธุ์", "กำแพงเพชร", "ขอนแก่น",
    "จันทบุรี", "ฉะเชิงเทรา", "ชลบุรี", "ชัยนาท", "ชัยภูมิ", "ชุมพร", "เชียงราย",
    "เชียงใหม่", "ตรัง", "ตราด", "ตาก", "นครนายก", "นครปฐม", "นครพนม",
    "นครราชสีมา", "นครศรีธรรมราช", "นครสวรรค์", "นนทบุรี", "นราธิวาส", "น่าน",
    "บึงกาฬ", "บุรีรัมย์", "ปทุมธานี", "ประจวบคีรีขันธ์", "ปราจีนบุรี", "ปัตตานี",
    "พระนครศรีอยุธยา", "พะเยา", "พังงา", "พัทลุง", "พิจิตร", "พิษณุโลก",
    "เพชรบุรี", "เพชรบูรณ์", "แพร่", "ภูเก็ต", "มหาสารคาม", "มุกดาหาร",
    "แม่ฮ่องสอน", "ยโสธร", "ยะลา", "ร้อยเอ็ด", "ระนอง", "ระยอง", "ราชบุรี",
    "ลพบุรี", "ลำปาง", "ลำพูน", "เลย", "ศรีสะเกษ", "สกลนคร", "สงขลา", "สตูล",
    "สมุทรปราการ", "สมุทรสงคราม", "สมุทรสาคร", "สระแก้ว", "สระบุรี", "สิงห์บุรี",
    "สุโขทัย", "สุพรรณบุรี", "สุราษฎร์ธานี", "สุรินทร์", "หนองคาย", "หนองบัวลำภู",
    "อ่างทอง", "อำนาจเจริญ", "อุดรธานี", "อุตรดิตถ์", "อุทัยธานี", "อุบลราชธานี"
]


def count_matches(text: str, keywords: list[str]) -> list[str]:
    if not text:
        return []
    text_lower = text.lower()
    return [kw for kw in keywords if kw in text_lower]


def extract_brands(text: str) -> str:
    if not text:
        return ""
    text_lower = text.lower()
    found = []
    for brand, patterns in TIRE_BRANDS.items():
        if any(p in text_lower for p in patterns):
            found.append(brand)
    return ", ".join(found)


def extract_province(address: str) -> str:
    if not address:
        return ""
    if "กรุงเทพ" in address or "กทม" in address:
        return "กรุงเทพมหานคร"
    for prov in PROVINCES_77:
        if prov in address:
            return prov
    return ""


def classify_shop_with_gemini_vision(image_url: str, title: str, api_key: str) -> str:
    """Uses Gemini 2.5 Flash Vision via REST API to inspect storefront thumbnail."""
    if not image_url or not image_url.startswith("http"):
        return "UNKNOWN"
    try:
        req_img = urllib.request.Request(
            image_url,
            headers={"User-Agent": "Mozilla/5.0"}
        )
        with urllib.request.urlopen(req_img, timeout=10) as resp:
            img_bytes = resp.read()
        img_b64 = base64.b64encode(img_bytes).decode("utf-8")

        prompt = (
            f"You are analyzing a Google Maps storefront photo for a shop in Thailand named '{title}'. "
            "Determine whether this shop services/sells motorcycle tires (2-wheelers), car/truck tires (4-wheelers), or both. "
            "Look at storefront signs, banners, parked vehicles, and tire racks. "
            "Reply with ONLY ONE of these exact labels: 2W_ONLY, HYBRID_2W_4W, 4W_ONLY, or UNKNOWN."
        )

        payload = {
            "contents": [{
                "parts": [
                    {"text": prompt},
                    {"inline_data": {"mime_type": "image/jpeg", "data": img_b64}}
                ]
            }],
            "generationConfig": {"temperature": 0.0, "maxOutputTokens": 20}
        }

        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
        req = urllib.request.Request(
            url,
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=15) as resp:
            result = json.loads(resp.read().decode("utf-8"))
            text = result["candidates"][0]["content"]["parts"][0]["text"].strip().upper()
            for label in ["2W_ONLY", "HYBRID_2W_4W", "4W_ONLY"]:
                if label in text:
                    return label
    except Exception:
        pass
    return "UNKNOWN"


def classify_row(row: pd.Series) -> tuple[str, str]:
    """Returns (classification, reason)"""
    title = str(row.get("title", "") or "")
    cat = str(row.get("category", "") or "")
    cats_all = str(row.get("categories", "") or "")
    cat_combined = f"{cat} {cats_all}".lower()
    reviews = f"{row.get('user_reviews', '') or ''} {row.get('descriptions', '') or ''} {row.get('about', '') or ''}"

    title_2w = count_matches(title, POS_2W_KEYWORDS)
    title_4w = count_matches(title, NEG_4W_KEYWORDS)
    rev_2w = count_matches(reviews, REVIEW_2W_KEYWORDS)
    rev_4w = count_matches(reviews, REVIEW_4W_KEYWORDS)

    # 1. Explicit 4W exclusion in title
    if title_4w and not title_2w:
        if rev_2w:
            return "HYBRID_2W_4W", f"4W title ({title_4w[0]}) but 2W reviews ({rev_2w[0]})"
        return "4W_EXCLUDE", f"4W keyword in title: {title_4w[0]}"

    # 2. Explicit 2W category
    is_2w_cat = any(k in cat_combined for k in ["รถจักรยานยนต์", "motorcycle", "scooter", "moped", "vespa"])
    is_4w_cat = any(k in cat_combined for k in ["รถยนต์", "auto ", "car ", "truck", "wheel alignment", "ตั้งศูนย์", "รถบรรทุก"])
    is_generic_tire_cat = ("tire shop" in cat_combined) or ("ร้านยาง" in cat_combined)

    if is_2w_cat:
        if title_4w or (len(rev_4w) >= 2 and not rev_2w):
            return "HYBRID_2W_4W", "2W category with 4W signals"
        return "2W_CONFIRMED", f"2W category: {cat}"

    # 3. Explicit 2W keyword in title
    if title_2w:
        if title_4w or is_4w_cat:
            return "HYBRID_2W_4W", f"2W title ({title_2w[0]}) + 4W signal"
        return "2W_CONFIRMED", f"2W title keyword: {title_2w[0]}"

    # 4. Review text evidence
    if rev_2w and not rev_4w:
        return "2W_CONFIRMED", f"2W review keyword: {rev_2w[0]}"
    if rev_2w and rev_4w:
        return "HYBRID_2W_4W", f"Mixed reviews ({rev_2w[0]} + {rev_4w[0]})"
    if rev_4w and not rev_2w:
        return "4W_EXCLUDE", f"4W review keyword: {rev_4w[0]}"

    # 5. Category-based fallback
    if is_4w_cat and not is_generic_tire_cat:
        return "4W_EXCLUDE", f"4W category: {cat}"

    if is_generic_tire_cat:
        return "AMBIGUOUS", "Generic tire shop without clear 2W/4W keywords"

    return "4W_EXCLUDE", f"Unrelated category: {cat}"


def assign_tier(row: pd.Series) -> str:
    cls = row["classification"]
    if cls not in ["2W_CONFIRMED", "HYBRID_2W_4W"]:
        return "Excluded"

    title = str(row.get("title", "") or "").lower()
    cat = str(row.get("category", "") or "").lower()
    website = str(row.get("website", "") or "").strip()
    try:
        rev_count = int(float(row.get("review_count", 0) or 0))
    except (ValueError, TypeError):
        rev_count = 0

    has_tire_word = any(w in title for w in ["ยาง", "tire", "tyre"]) or ("ขายยาง" in cat)

    if cls == "HYBRID_2W_4W":
        return "Tier 1 - Hybrid Tire Dealer" if rev_count >= 25 else "Tier 2 - Hybrid Shop"

    if has_tire_word and rev_count >= 20:
        return "Tier 1 - Motorcycle Tire Specialist"
    if has_tire_word or rev_count >= 10 or website:
        return "Tier 2 - Moto Parts & Tire Service"
    return "Tier 3 - General Moto Repair / Small Shop"


def main():
    parser = argparse.ArgumentParser(description="Clean & filter motorcycle tire shops CSV")
    parser.add_argument("--input", default="raw_results.csv", help="Input raw CSV path")
    parser.add_argument("--output", default="clean_motorcycle_tire_shops.csv", help="Output clean CSV path")
    parser.add_argument("--use-vision", action="store_true", help="Use Gemini Vision API for ambiguous shops")
    parser.add_argument("--max-vision", type=int, default=100, help="Max Gemini Vision API calls")
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Input file not found: {args.input}")
        return

    df = pd.read_csv(args.input, dtype=str).fillna("")
    print(f"Loaded {len(df)} raw rows from {args.input}")

    # 1. Deduplicate
    def make_dedup_key(r):
        if r.get("cid"):
            return f"cid:{r['cid']}"
        if r.get("data_id"):
            return f"did:{r['data_id']}"
        if r.get("link"):
            return f"link:{r['link'].split('?')[0]}"
        return f"ta:{r.get('title','')}_{r.get('address','')}"

    df["_dedup_key"] = df.apply(make_dedup_key, axis=1)
    df = df.drop_duplicates(subset=["_dedup_key"]).copy()
    print(f"After deduplication: {len(df)} unique places")

    # 2. Classify Stage 1 (Category + Title + Reviews)
    results = df.apply(classify_row, axis=1)
    df["classification"] = [r[0] for r in results]
    df["match_reason"] = [r[1] for r in results]

    # 3. Optional Stage 2 (Gemini Vision on AMBIGUOUS thumbnails)
    api_key = os.environ.get("GEMINI_API_KEY", "").strip()
    if args.use_vision and api_key:
        amb_idx = df[df["classification"] == "AMBIGUOUS"].index[:args.max_vision]
        print(f"Running Gemini Vision on {len(amb_idx)} ambiguous shops...")
        for idx in amb_idx:
            thumb = df.at[idx, "thumbnail"] if "thumbnail" in df.columns else ""
            title = df.at[idx, "title"] if "title" in df.columns else ""
            v_label = classify_shop_with_gemini_vision(thumb, title, api_key)
            if v_label in ["2W_ONLY", "HYBRID_2W_4W"]:
                df.at[idx, "classification"] = "2W_CONFIRMED" if v_label == "2W_ONLY" else "HYBRID_2W_4W"
                df.at[idx, "match_reason"] = f"Gemini Vision: {v_label}"
            elif v_label == "4W_ONLY":
                df.at[idx, "classification"] = "4W_EXCLUDE"
                df.at[idx, "match_reason"] = "Gemini Vision: 4W_ONLY"
            time.sleep(0.3)

    # 4. Keep confirmed 2W and Hybrid shops
    keep_df = df[df["classification"].isin(["2W_CONFIRMED", "HYBRID_2W_4W"])].copy()

    # 5. Enrich with Tier, Province, and Mentioned Brands
    keep_df["shop_tier"] = keep_df.apply(assign_tier, axis=1)
    addr_col = "complete_address" if "complete_address" in keep_df.columns else "address"
    keep_df["province"] = keep_df[addr_col].apply(extract_province) if addr_col in keep_df.columns else ""

    text_for_brands = (
        keep_df.get("title", pd.Series("", index=keep_df.index)) + " " +
        keep_df.get("user_reviews", pd.Series("", index=keep_df.index)) + " " +
        keep_df.get("descriptions", pd.Series("", index=keep_df.index))
    )
    keep_df["mentioned_tire_brands"] = text_for_brands.apply(extract_brands)

    # 6. Select & order clean columns
    desired_cols = [
        "shop_tier", "classification", "title", "province", "category",
        "review_rating", "review_count", "mentioned_tire_brands",
        "phone", "website", "address", "latitude", "longitude",
        "link", "thumbnail", "match_reason", "cid"
    ]
    final_cols = [c for c in desired_cols if c in keep_df.columns]
    out_df = keep_df[final_cols].sort_values(
        by=["shop_tier", "province", "title"], ascending=[True, True, True]
    )

    out_df.to_csv(args.output, index=False, encoding="utf-8-sig")
    print(f"Saved {len(out_df)} verified 2W/Hybrid tire shops to {args.output}")


if __name__ == "__main__":
    main()
