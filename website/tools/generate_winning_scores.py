#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sri Lanka Winning Product Index (SL-WPI — 100-Point Sea Cargo Model)
Evaluates every product on a 100-point scale tailored for the Sri Lankan
e-commerce / Cash-on-Delivery (COD) market importing via Sea Cargo.

4 Pillars (0-25 pts each, sum == winScore):
  1. freight (0-25): Freight & Sea Cargo Efficiency (unit CBM & unit KG)
  2. margin  (0-25): Profit Margin & SL Retail Sweet Spot (RMB factory cost vs LKR COD retail)
  3. demand  (0-25): Local Consumer Demand & Viral Appeal in Sri Lanka
  4. risk    (0-25): Operational & COD Return Risk (durability, transit safety, ticket risk)
"""

import json
import os
import re
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


def compute_sl_wpi(product):
    """
    Compute the 100-point Sri Lanka Winning Product Index (SL-WPI) for a product dict.
    Mutates and returns `product` with:
      - winScore (int 0..100)
      - winRate (str "XX%")
      - winTier ("mega" | "strong" | "niche" | "low")
      - winBreakdown ({ freight, margin, demand, risk })
      - winReason ({ en, si, zh, es })
    """
    price = float(product.get("price") or 0.0)
    pcs = max(1, int(product.get("pcsPerCarton") or 1))
    cbm = float(product.get("cbm") or 0.12)
    kg_raw = product.get("kg")
    kg = float(kg_raw) if (kg_raw is not None and kg_raw != "") else None

    unit_cbm = cbm / pcs
    if kg is not None and kg > 0:
        unit_kg = kg / pcs
    else:
        # Estimate unit weight from volume (~180 kg/m³ for foldable organizers)
        unit_kg = round(unit_cbm * 180.0, 3)

    zh = str(product.get("zh") or "")
    en = str(product.get("en") or "").lower()
    cat = str(product.get("category") or "")

    # -------------------------------------------------------------------------
    # PILLAR 1: Freight & Sea Cargo Efficiency (0 - 25 pts)
    # -------------------------------------------------------------------------
    if unit_cbm <= 0.0006 and unit_kg <= 0.25:
        freight_score = 25
    elif unit_cbm <= 0.0008 and unit_kg <= 0.35:
        freight_score = 24
    elif unit_cbm <= 0.0012 and unit_kg <= 0.45:
        freight_score = 22
    elif unit_cbm <= 0.0018 and unit_kg <= 0.65:
        freight_score = 20
    elif unit_cbm <= 0.0025 and unit_kg <= 0.85:
        freight_score = 18
    elif unit_cbm <= 0.0035 and unit_kg <= 1.15:
        freight_score = 15
    elif unit_cbm <= 0.0055 and unit_kg <= 1.60:
        freight_score = 12
    elif unit_cbm <= 0.0090:
        freight_score = 9
    else:
        freight_score = 6

    unit_cbm_cm3 = round(unit_cbm * 10000, 1)  # in 0.0001 m³ units for readability
    if freight_score >= 22:
        freight_desc = {
            "en": f"Ultra-compact flat-pack volume ({unit_cbm:.4f} m³/pc, ~{unit_kg:.2f} kg/pc; {pcs} pcs/ctn) maximizes sea cargo container density and minimizes LKR landed freight cost.",
            "si": f"ඉතා කුඩා ඒකක පරිමාව ({unit_cbm:.4f} m³/pc, ~{unit_kg:.2f} kg/pc; පෙට්ටියකට කෑලි {pcs}) නිසා මුහුදු නැව්ගත කිරීමේ ගාස්තුව අවම වී වැඩි ලාභයක් ලබා දෙයි.",
            "zh": f"超紧凑折叠体积（单件 {unit_cbm:.4f} m³，约 {unit_kg:.2f} kg，{pcs}件/箱），海运拼箱/整柜装载率极高，单件分摊海运费极低。",
            "es": f"Volumen plegable ultracompacto ({unit_cbm:.4f} m³/ud, ~{unit_kg:.2f} kg/ud; {pcs} uds/caja) que maximiza la eficiencia del flete marítimo."
        }
    elif freight_score >= 16:
        freight_desc = {
            "en": f"Moderate foldable volume ({unit_cbm:.4f} m³/pc, ~{unit_kg:.2f} kg/pc; {pcs} pcs/ctn) offers balanced sea freight efficiency per unit.",
            "si": f"සාමාන්‍ය ඒකක පරිමාව ({unit_cbm:.4f} m³/pc, ~{unit_kg:.2f} kg/pc; පෙට්ටියකට කෑලි {pcs}) මුහුදු ප්‍රවාහනය සඳහා හොඳ කාර්යක්ෂමතාවයක් පෙන්වයි.",
            "zh": f"适中折叠体积（单件 {unit_cbm:.4f} m³，约 {unit_kg:.2f} kg，{pcs}件/箱），海运柜容利用率良好，运费成本可控。",
            "es": f"Volumen plegable moderado ({unit_cbm:.4f} m³/ud, ~{unit_kg:.2f} kg/ud; {pcs} uds/caja) con eficiencia equilibrada en carga marítima."
        }
    else:
        freight_desc = {
            "en": f"Bulky unit footprint ({unit_cbm:.4f} m³/pc, ~{unit_kg:.2f} kg/pc; only {pcs} pcs/ctn) consumes significant sea container CBM and increases domestic courier tiers.",
            "si": f"විශාල ඒකක පරිමාව ({unit_cbm:.4f} m³/pc, ~{unit_kg:.2f} kg/pc; පෙට්ටියකට කෑලි {pcs} පමණි) නිසා මුහුදු නැව් ගාස්තු සහ කුරියර් ගාස්තු ඉහළ යයි.",
            "zh": f"单件体积/重量较大（{unit_cbm:.4f} m³/件，约 {unit_kg:.2f} kg，仅 {pcs}件/箱），占用较多海运体积及本地快递首重/续重成本。",
            "es": f"Huella voluminosa por unidad ({unit_cbm:.4f} m³/ud, ~{unit_kg:.2f} kg/ud; solo {pcs} uds/caja) que incrementa el costo de flete marítimo y mensajería."
        }

    # -------------------------------------------------------------------------
    # PILLAR 2: Profit Margin & SL Retail Sweet Spot (0 - 25 pts)
    # -------------------------------------------------------------------------
    # In Sri Lanka COD e-commerce, factory price 3.8 - 16.5 RMB (~LKR 165 - 720 FOB)
    # sells effortlessly at LKR 1,950 - 4,900 (or 3/4/6/7-pc sets at LKR 3,500 - 7,500)
    # yielding 3x - 4.5x markup!
    is_set = any(k in zh for k in ("套", "合1", "合一")) or "set" in en
    if 4.0 <= price <= 13.5:
        margin_score = 24 if not is_set else 25
    elif 13.5 < price <= 18.5:
        margin_score = 22 if is_set else 21
    elif 2.5 <= price < 4.0:
        margin_score = 21  # Great as 2-pack / 3-pack bundle on FB/Daraz
    elif 18.5 < price <= 23.5:
        margin_score = 17
    elif 1.0 <= price < 2.5:
        margin_score = 18  # Low single-unit ticket, best sold as multi-pack
    elif 23.5 < price <= 28.5:
        margin_score = 13
    else:
        margin_score = 9   # High factory cost (>28.5 RMB) compresses COD impulse markup

    fob_lkr = round(price * 43.5)
    est_retail_lkr = round((fob_lkr * 3.4) / 100.0) * 100
    est_retail_lkr = max(1490, min(14900, est_retail_lkr))

    if margin_score >= 21:
        margin_desc = {
            "en": f"Factory cost ¥{price:.2f} (~LKR {fob_lkr}) hits Sri Lanka's prime COD sweet spot, supporting a 3x–4x retail price (~LKR {est_retail_lkr:,}) with strong ad-spend ROI.",
            "si": f"කර්මාන්තශාලා මිල ¥{price:.2f} (~රු. {fob_lkr}) ශ්‍රී ලංකාවේ COD වෙළඳපොළට ඉතාමත් ගැලපෙන අතර රු. {est_retail_lkr:,} පමණ සිල්ලර මිලකට (3x–4x ලාභ පරතරයක් සහිතව) අලෙවි කළ හැක.",
            "zh": f"出厂价 ¥{price:.2f}（约 LKR {fob_lkr}）处于斯里兰卡 COD 黄金利润区间，终端零售可达 LKR {est_retail_lkr:,} 左右（3-4倍溢价空间）。",
            "es": f"El costo de fábrica ¥{price:.2f} (~LKR {fob_lkr}) se ubica en el punto ideal de margen COD, permitiendo un precio minorista de ~LKR {est_retail_lkr:,} (margen 3x–4x)."
        }
    elif margin_score >= 16:
        margin_desc = {
            "en": f"Factory cost ¥{price:.2f} (~LKR {fob_lkr}) provides solid 2.5x–3x markup potential (or bundle pricing around LKR {est_retail_lkr:,}) for online retail.",
            "si": f"කර්මාන්තශාලා මිල ¥{price:.2f} (~රු. {fob_lkr}) මගින් 2.5x–3x පමණ හොඳ ලාභ පරතරයක් (සිල්ලර මිල ~රු. {est_retail_lkr:,}) ලබා ගත හැක.",
            "zh": f"出厂价 ¥{price:.2f}（约 LKR {fob_lkr}）具备稳健的 2.5-3 倍加价空间（或组合装定价约 LKR {est_retail_lkr:,}）。",
            "es": f"El costo de fábrica ¥{price:.2f} (~LKR {fob_lkr}) ofrece un margen sólido de 2.5x–3x (precio minorista estimado ~LKR {est_retail_lkr:,})."
        }
    else:
        margin_desc = {
            "en": f"Higher factory cost ¥{price:.2f} (~LKR {fob_lkr}) plus sea freight pushes landed cost up, limiting COD impulse-buy markup flexibility.",
            "si": f"ඉහළ කර්මාන්තශාලා මිල ¥{price:.2f} (~රු. {fob_lkr}) සහ ප්‍රවාහන පිරිවැය නිසා COD පාරිභෝගිකයන් සඳහා ලාභ පරතරය සීමා වේ.",
            "zh": f"出厂单价较高（¥{price:.2f}，约 LKR {fob_lkr}），叠加海运成本后终端售价偏高，限制了 COD 冲动消费的加价倍率。",
            "es": f"El mayor costo de fábrica ¥{price:.2f} (~LKR {fob_lkr}) más el flete reduce la flexibilidad de margen en ventas contra reembolso (COD)."
        }

    # -------------------------------------------------------------------------
    # PILLAR 3: Local Consumer Demand & Viral Appeal (0 - 25 pts)
    # -------------------------------------------------------------------------
    # High viral demand in SL: wardrobe/underwear grid organizers, steel-frame window
    # storage boxes (protects clothes/saris from tropical humidity & dust), cosmetic/travel sets,
    # foldable laundry hampers, shoe boxes, over-the-door hanging organizers.
    demand_score = 17
    if any(k in zh for k in ("内衣", "文胸", "袜", "分隔裤盒", "分格收纳", "衣物收纳格", "7格", "6格", "8格", "16格", "20格", "24格")):
        demand_score = 24
    elif any(k in zh for k in ("洗漱", "化妆", "旅行")):
        demand_score = 23
    elif any(k in zh for k in ("百纳箱", "钢架", "视窗", "透窗", "可视", "鞋盒")):
        demand_score = 22
    elif any(k in zh for k in ("挂袋", "门后", "床底", "防尘罩", "棉被袋", "搬家袋")):
        demand_score = 21
    elif any(k in zh for k in ("拉链箱", "天地盖", "连盖", "双盖", "软盖", "抽屉")):
        demand_score = 20
    elif any(k in zh for k in ("脏衣", "洗衣")):
        demand_score = 19 if price <= 22 else 13
    elif any(k in zh for k in ("收纳筐", "收纳框", "收纳桶", "方盒", "扣扣箱", "整理箱")):
        demand_score = 18
    elif "收纳凳" in zh:
        demand_score = 15 if price <= 22 else 11

    # Adjust slightly by size practicality
    if "150L" in zh or "130L" in zh or "110L" in zh:
        demand_score = max(10, demand_score - 3)
    if is_set and demand_score < 24:
        demand_score = min(25, demand_score + 2)

    if demand_score >= 21:
        demand_desc = {
            "en": "High viral appeal on Facebook, TikTok & Daraz LK — solves everyday wardrobe, sari/garment dust protection, and space-saving organization needs in Sri Lankan homes.",
            "si": "ශ්‍රී ලංකාවේ නිවාසවල ඇඳුම් අල්මාරි පිළිවෙලට තබා ගැනීමට සහ දූවිලි/තෙතමනයෙන් ආරක්ෂා කිරීමට ඉතා ඉහළ ඉල්ලුමක් ඇති (Facebook/TikTok/Daraz) ජනප්‍රිය භාණ්ඩයකි.",
            "zh": "在斯里兰卡 Facebook、TikTok 和 Daraz 平台具有极高爆款潜力，完美解决当地家庭衣柜收纳、防尘防潮及空间节省刚需。",
            "es": "Alto atractivo viral en Facebook, TikTok y Daraz LK: resuelve necesidades diarias de organización de armarios y protección contra el polvo y la humedad."
        }
    elif demand_score >= 16:
        demand_desc = {
            "en": "Steady household utility demand across urban and suburban Sri Lankan households looking for neat bedroom and laundry organization.",
            "si": "නිදන කාමර සහ ගෘහස්ථ භාණ්ඩ පිළිවෙලට තබා ගැනීම සඳහා ශ්‍රී ලාංකික පාරිභෝගිකයන් අතර ස්ථාවර ඉල්ලුමක් පවතී.",
            "zh": "斯里兰卡城市及郊区家庭日常居家整理刚需品，卧室与家政收纳需求稳定。",
            "es": "Demanda doméstica constante en hogares de Sri Lanka para la organización ordenada del dormitorio y la lavandería."
        }
    else:
        demand_desc = {
            "en": "Niche buyer segment in Sri Lanka due to larger footprint or higher price point compared to fast-moving compact organizers.",
            "si": "විශාල ප්‍රමාණය හෝ වැඩි මිල නිසා සාමාන්‍ය එදිනෙදා සංවිධායක පෙට්ටිවලට සාපේක්ෂව විශේෂිත පාරිභෝගික පිරිසකට පමණක් සීමා වේ.",
            "zh": "受限于较大尺寸或较高单价，在斯里兰卡属于细分小众需求，受众面窄于轻便型收纳品。",
            "es": "Segmento de compradores de nicho debido a su mayor tamaño o precio en comparación con organizadores compactos de alta rotación."
        }

    # -------------------------------------------------------------------------
    # PILLAR 4: Operational & COD Return Risk (0 - 25 pts)
    # -------------------------------------------------------------------------
    # Foldable textile items have zero glass/electronic failure!
    # Risk increases only with heavy MDF boards, metal/iron tubes, wheels, or high COD price.
    risk_score = 23
    if any(k in zh for k in ("带轮", "万向轮", "铁管", "铝管")):
        risk_score = 14
    elif "收纳凳" in zh:
        risk_score = 16 if unit_kg <= 2.0 else 12
    elif any(k in zh for k in ("密度板", "双灰板", "纸板", "钢架")):
        risk_score = 20 if unit_kg <= 0.9 else 17
    else:
        risk_score = 24  # Soft foldable fabric/non-woven/PU/PVC/mesh - virtually zero transit damage

    if price > 26:
        risk_score = max(10, risk_score - 4)
    elif price > 18:
        risk_score = max(12, risk_score - 2)

    if risk_score >= 21:
        risk_desc = {
            "en": "100% foldable, non-fragile fabric construction with zero electronics or glass — virtually 0% courier transit damage and very low COD doorstep rejection.",
            "si": "කැඩී බිඳී නොයන නැවිය හැකි රෙදි නිමාව නිසා කුරියර් ප්‍රවාහනයේදී හානි වීමේ අවදානම 0% ක් වන අතර COD ප්‍රතික්ෂේප වීමේ ප්‍රතිශතය ඉතා අවම වේ.",
            "zh": "纯布艺/柔性折叠结构，无玻璃无电子元件，本地快递暴力分拣零破损率，且客单价亲民、COD 拒收率极低。",
            "es": "Construcción textil 100% plegable y no frágil, sin vidrio ni electrónica: 0% de daño en tránsito por mensajería y mínima tasa de rechazo COD."
        }
    elif risk_score >= 16:
        risk_desc = {
            "en": "Sturdy reinforced board/frame structure handles standard courier delivery well, with moderate COD ticket risk.",
            "si": "ශක්තිමත් රාමු/පුවරු නිමාව සාමාන්‍ය කුරියර් ප්‍රවාහනයට ඔරොත්තු දෙන අතර COD ප්‍රතික්ෂේප වීමේ අවදානම සාමාන්‍ය මට්ටමක පවතී.",
            "zh": "内置加固衬板/钢架结构，常规快递派送安全性良好，COD 签收风险适中。",
            "es": "Estructura reforzada con marco/tablero resistente al envío estándar por mensajería, con riesgo moderado en entregas COD."
        }
    else:
        risk_desc = {
            "en": "Heavier rigid components (wheels/tubes/thick MDF) and higher COD retail price require protective outer packaging and carry higher doorstep return risk.",
            "si": "බරැති කොටස් (රෝද/ලෝහ නල/ඝන පුවරු) සහ ඉහළ COD මිල නිසා ආරක්ෂිත ඇසුරුම් අවශ්‍ය වන අතර ඇණවුම් ප්‍රතික්ෂේප වීමේ අවදානම වැඩිය.",
            "zh": "含较重刚性配件（轮子/铁管/厚密度板）且终端 COD 客单价偏高，需加强外包装防护并防范退货风险。",
            "es": "Componentes rígidos más pesados y un mayor precio COD requieren embalaje protector adicional y presentan mayor riesgo de devolución."
        }

    # Exact sum invariant: freight + margin + demand + risk == winScore
    win_score = int(freight_score + margin_score + demand_score + risk_score)
    win_score = max(0, min(100, win_score))

    if win_score >= 85:
        win_tier = "mega"
        reason = {
            "en": f"Mega Winner ({win_score}/100): Exceptional sea-cargo density ({unit_cbm:.4f} m³/pc), high 3x–4x COD retail markup from ¥{price:.2f}, strong Sri Lankan viral home-organization demand, and near-zero courier damage risk.",
            "si": f"සුපිරි ජයග්‍රාහී භාණ්ඩය ({win_score}/100): ඉතා අඩු නැව් ගාස්තුව ({unit_cbm:.4f} m³/pc), ¥{price:.2f} සිට ඉහළ ලාභ පරතරය, ශ්‍රී ලංකාවේ ඉහළම පාරිභෝගික ඉල්ලුම සහ ප්‍රවාහන හානි නොමැති වීම.",
            "zh": f"超级爆款 ({win_score}/100)：极佳海运体积密度（{unit_cbm:.4f} m³/件）、¥{price:.2f} 黄金拿货价带来 3-4 倍 COD 利润、斯里兰卡高热度家居刚需且快递零破损。",
            "es": f"Mega Ganador ({win_score}/100): Densidad óptima de carga marítima ({unit_cbm:.4f} m³/ud), alto margen COD desde ¥{price:.2f}, fuerte demanda viral y mínimo riesgo logístico."
        }
    elif win_score >= 70:
        win_tier = "strong"
        reason = {
            "en": f"Strong Winner ({win_score}/100): Reliable Sri Lankan e-commerce performer combining efficient sea freight ({pcs} pcs/ctn), healthy COD profit margins (¥{price:.2f} FOB), and steady household utility.",
            "si": f"ඉහළ විකුණුම් හැකියාවක් ඇති භාණ්ඩය ({win_score}/100): කාර්යක්ෂම මුහුදු ප්‍රවාහනය (පෙට්ටියකට කෑලි {pcs}), හොඳ COD ලාභ පරතරය (¥{price:.2f}) සහ ස්ථාවර ගෘහස්ථ ඉල්ලුම.",
            "zh": f"强力推荐款 ({win_score}/100)：海运装箱率高（{pcs}件/箱）、¥{price:.2f} 出厂价利润空间稳健，契合斯里兰卡家庭日常收纳需求。",
            "es": f"Fuerte Ganador ({win_score}/100): Sólido desempeño en e-commerce combinando flete marítimo eficiente ({pcs} uds/caja), buen margen COD (¥{price:.2f}) y demanda constante."
        }
    elif win_score >= 50:
        win_tier = "niche"
        reason = {
            "en": f"Niche Product ({win_score}/100): Suitable for targeted catalog expansion or bundle offers; larger unit volume ({unit_cbm:.4f} m³/pc) or higher ¥{price:.2f} cost moderates impulse COD volume.",
            "si": f"විශේෂිත වෙළඳපොළ භාණ්ඩය ({win_score}/100): තෝරාගත් පාරිභෝගිකයන් සඳහා සුදුසුයි; ඒකක පරිමාව ({unit_cbm:.4f} m³/pc) හෝ ¥{price:.2f} පිරිවැය නිසා සාමාන්‍ය අලෙවි වේගයක් පවතී.",
            "zh": f"细分市场款 ({win_score}/100)：适合作为店铺丰富度补充或搭配销售；单件体积（{unit_cbm:.4f} m³）或 ¥{price:.2f} 成本使其更适合精准受众。",
            "es": f"Producto de Nicho ({win_score}/100): Adecuado para expansión de catálogo; su volumen ({unit_cbm:.4f} m³/ud) o costo (¥{price:.2f}) modera la rotación masiva."
        }
    else:
        win_tier = "low"
        reason = {
            "en": f"Low Feasibility ({win_score}/100): Bulky sea-cargo footprint ({unit_cbm:.4f} m³/pc, {pcs} pcs/ctn) and high factory price (¥{price:.2f}) create high landed cost and elevated COD return risk in Sri Lanka.",
            "si": f"අඩු යෝග්‍යතාවයක් ඇති භාණ්ඩය ({win_score}/100): විශාල නැව්ගත කිරීමේ පරිමාව ({unit_cbm:.4f} m³/pc, පෙට්ටියකට කෑලි {pcs}) සහ ඉහළ මිල (¥{price:.2f}) නිසා COD අලෙවිය සඳහා අවදානම වැඩිය.",
            "zh": f"可行性较低 ({win_score}/100)：单件海运体积偏大（{unit_cbm:.4f} m³，仅 {pcs}件/箱）且出厂价较高（¥{price:.2f}），到岸成本与 COD 拒收风险偏高。",
            "es": f"Baja Viabilidad ({win_score}/100): El gran volumen de carga ({unit_cbm:.4f} m³/ud, {pcs} uds/caja) y alto costo (¥{price:.2f}) elevan el costo final y el riesgo COD."
        }

    product["winScore"] = win_score
    product["winRate"] = f"{win_score}%"
    product["winTier"] = win_tier
    product["winReason"] = reason
    product["winBreakdown"] = {
        "freight": {
            "score": freight_score,
            "max": 25,
            "desc": freight_desc
        },
        "margin": {
            "score": margin_score,
            "max": 25,
            "desc": margin_desc
        },
        "demand": {
            "score": demand_score,
            "max": 25,
            "desc": demand_desc
        },
        "risk": {
            "score": risk_score,
            "max": 25,
            "desc": risk_desc
        }
    }
    return product


def score_catalog_file(products_js_path):
    """Load website/data/products.js, score all products, and write back."""
    if not os.path.exists(products_js_path):
        print(f"File not found: {products_js_path}")
        return
    with open(products_js_path, "r", encoding="utf-8") as f:
        text = f.read()
    m = re.search(r"window\.CATALOG\s*=\s*(\{.*\})\s*;\s*$", text, re.DOTALL)
    if not m:
        raise ValueError("Could not parse window.CATALOG in " + products_js_path)
    catalog = json.loads(m.group(1))
    tier_counts = {"mega": 0, "strong": 0, "niche": 0, "low": 0}
    for p in catalog.get("products", []):
        compute_sl_wpi(p)
        tier_counts[p["winTier"]] = tier_counts.get(p["winTier"], 0) + 1
    out_js = "window.CATALOG = " + json.dumps(catalog, ensure_ascii=False, indent=2) + ";\n"
    with open(products_js_path, "w", encoding="utf-8") as f:
        f.write(out_js)
    print(f"Scored {len(catalog.get('products', []))} products. Tier breakdown: {tier_counts}")


if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    target = os.path.join(base_dir, "data", "products.js")
    score_catalog_file(target)
