# -*- coding: utf-8 -*-
"""Классификация автора документа инвентаризации: ревизор / оператор."""
from __future__ import annotations

import datetime
from typing import Any, Dict, List, Optional, Tuple, Union

import pandas as pd

AUTHOR_REVISOR = "ревизор"
AUTHOR_OPERATOR = "оператор"
AUTHOR_UNKNOWN = "неизвестно"

# Ровно 08:00:00 — ревизор (по бизнес-правилу).
REVISOR_TIME = datetime.time(8, 0, 0)


def parse_time_flexible(value: Any) -> Optional[datetime.time]:
    """Извлечь время из строки/datetime/time Excel-вариантов."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, datetime.datetime):
        return value.time().replace(microsecond=0)
    if isinstance(value, datetime.time):
        return value.replace(microsecond=0)
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime().time().replace(microsecond=0)

    s = str(value).strip()
    if not s or s.lower() in ("nan", "nat", "none"):
        return None

    # Полный datetime в строке
    for fmt in (
        "%d.%m.%Y %H:%M:%S",
        "%d.%m.%Y %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%d/%m/%Y %H:%M:%S",
    ):
        try:
            return datetime.datetime.strptime(s, fmt).time()
        except ValueError:
            pass

    # Только время: 8:00:00 / 08:00:00 / 8:00 / 08:00
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.datetime.strptime(s, fmt).time()
        except ValueError:
            pass
    # Без ведущего нуля у часа
    parts = s.replace(".", ":").split(":")
    if 2 <= len(parts) <= 3:
        try:
            h = int(parts[0])
            m = int(parts[1])
            sec = int(parts[2]) if len(parts) == 3 else 0
            if 0 <= h <= 23 and 0 <= m <= 59 and 0 <= sec <= 59:
                return datetime.time(h, m, sec)
        except ValueError:
            return None
    return None


def classify_author(
    dt: Optional[Union[datetime.datetime, datetime.time, pd.Timestamp, str]] = None,
    *,
    time_only: Optional[datetime.time] = None,
) -> str:
    """
    ревизор — время ровно 08:00:00;
    оператор — любое другое распознанное время;
    неизвестно — время не извлечено.
    """
    t = time_only
    if t is None:
        if isinstance(dt, datetime.time):
            t = dt.replace(microsecond=0)
        elif isinstance(dt, (datetime.datetime, pd.Timestamp)):
            t = parse_time_flexible(dt)
        else:
            t = parse_time_flexible(dt)
    if t is None:
        return AUTHOR_UNKNOWN
    t = t.replace(microsecond=0)
    if t == REVISOR_TIME:
        return AUTHOR_REVISOR
    return AUTHOR_OPERATOR


def enrich_author_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Добавить автор_дока и время_дока по дате документа."""
    out = df.copy()
    times: List[Optional[datetime.time]] = []
    authors: List[str] = []
    for val in out.get("дата_док", pd.Series(dtype=object)):
        t = parse_time_flexible(val)
        times.append(t)
        authors.append(classify_author(time_only=t))
    out["время_дока"] = times
    out["автор_дока"] = authors
    return out


def author_doc_lookup(df: pd.DataFrame) -> Dict[Tuple[str, str], str]:
    """(магазин, документ) -> автор."""
    if df.empty or "автор_дока" not in df.columns:
        return {}
    g = (
        df.groupby(["магазин", "документ"], dropna=False)["автор_дока"]
        .agg(lambda s: s.mode().iloc[0] if not s.mode().empty else AUTHOR_UNKNOWN)
    )
    return {(str(k[0]), str(k[1])): str(v) for k, v in g.items()}


def annotate_overlap_authors(overlap_df: pd.DataFrame, df: pd.DataFrame) -> pd.DataFrame:
    """Добавить автора документа излишка и недостачи в пары перекрытия."""
    if overlap_df is None or overlap_df.empty:
        return overlap_df if overlap_df is not None else pd.DataFrame()
    lookup = author_doc_lookup(df)
    out = overlap_df.copy()
    out["автор_изл"] = [
        lookup.get((str(r.get("магазин_изл", "")), str(r.get("док_изл", ""))), AUTHOR_UNKNOWN)
        for _, r in out.iterrows()
    ]
    out["автор_нед"] = [
        lookup.get((str(r.get("магазин_нед", "")), str(r.get("док_нед", ""))), AUTHOR_UNKNOWN)
        for _, r in out.iterrows()
    ]
    return out


def _safe_share(part: float, total: float) -> float:
    return (part / total * 100.0) if total else 0.0


def calc_author_role_stats(df: pd.DataFrame, overlap_df: pd.DataFrame) -> pd.DataFrame:
    """Сводка по ролям на уровне сети."""
    rows = []
    lookup = author_doc_lookup(df)
    for role in (AUTHOR_REVISOR, AUTHOR_OPERATOR, AUTHOR_UNKNOWN):
        sub = df[df["автор_дока"] == role] if "автор_дока" in df.columns else df.iloc[0:0]
        docs = sub["документ"].nunique() if not sub.empty else 0
        sur = float(sub["излишек_сумма"].sum()) if not sub.empty else 0.0
        sh = float(sub["недостача_сумма"].sum()) if not sub.empty else 0.0
        ov = 0.0
        if overlap_df is not None and not overlap_df.empty and "автор_нед" in overlap_df.columns:
            # Пересорт относим к автору документа недостачи (что «закрывается»)
            ov = float(overlap_df.loc[overlap_df["автор_нед"] == role, "перекрытие_сум"].sum())
        elif overlap_df is not None and not overlap_df.empty:
            # Fallback: атрибуция по lookup
            for _, r in overlap_df.iterrows():
                a = lookup.get((str(r.get("магазин_нед", "")), str(r.get("док_нед", ""))), AUTHOR_UNKNOWN)
                if a == role:
                    ov += float(r.get("перекрытие_сум", 0) or 0)
        rows.append({
            "автор": role,
            "документов": docs,
            "излишки": sur,
            "недостачи": sh,
            "перекрытие": ov,
            "сальдо": sur - sh,
        })
    out = pd.DataFrame(rows)
    tot_sur = out["излишки"].sum()
    tot_sh = out["недостачи"].sum()
    tot_ov = out["перекрытие"].sum()
    tot_net = out["сальдо"].sum()
    out["доля_излишки_%"] = out["излишки"].apply(lambda x: _safe_share(x, tot_sur))
    out["доля_недостачи_%"] = out["недостачи"].apply(lambda x: _safe_share(x, tot_sh))
    out["доля_перекрытие_%"] = out["перекрытие"].apply(lambda x: _safe_share(x, tot_ov))
    out["доля_сальдо_%"] = out["сальдо"].apply(lambda x: _safe_share(abs(x), abs(tot_net)) if tot_net else 0.0)
    return out


def calc_author_store_stats(df: pd.DataFrame, overlap_df: pd.DataFrame) -> pd.DataFrame:
    """По магазину: вклад ревизора и оператора."""
    if df.empty:
        return pd.DataFrame()
    rows = []
    ov_by_store_role: Dict[Tuple[str, str], float] = {}
    if overlap_df is not None and not overlap_df.empty:
        annotated = overlap_df if "автор_нед" in overlap_df.columns else annotate_overlap_authors(overlap_df, df)
        for _, r in annotated.iterrows():
            key = (str(r.get("магазин_нед", "")), str(r.get("автор_нед", AUTHOR_UNKNOWN)))
            ov_by_store_role[key] = ov_by_store_role.get(key, 0.0) + float(r.get("перекрытие_сум", 0) or 0)

    for store, grp in df.groupby("магазин"):
        rec: Dict[str, Any] = {"магазин": store}
        for role, prefix in ((AUTHOR_REVISOR, "рев"), (AUTHOR_OPERATOR, "оп"), (AUTHOR_UNKNOWN, "неизв")):
            sub = grp[grp["автор_дока"] == role]
            rec[f"{prefix}_документов"] = int(sub["документ"].nunique()) if not sub.empty else 0
            rec[f"{prefix}_излишки"] = float(sub["излишек_сумма"].sum()) if not sub.empty else 0.0
            rec[f"{prefix}_недостачи"] = float(sub["недостача_сумма"].sum()) if not sub.empty else 0.0
            rec[f"{prefix}_перекрытие"] = float(ov_by_store_role.get((store, role), 0.0))
            rec[f"{prefix}_сальдо"] = rec[f"{prefix}_излишки"] - rec[f"{prefix}_недостачи"]

        tot_sh = rec["рев_недостачи"] + rec["оп_недостачи"] + rec["неизв_недостачи"]
        tot_sur = rec["рев_излишки"] + rec["оп_излишки"] + rec["неизв_излишки"]
        tot_ov = rec["рев_перекрытие"] + rec["оп_перекрытие"] + rec["неизв_перекрытие"]
        tot_net = tot_sur - tot_sh
        rec["излишки"] = tot_sur
        rec["недостачи"] = tot_sh
        rec["перекрытие"] = tot_ov
        rec["сальдо"] = tot_net
        rec["доля_рев_недостачи_%"] = _safe_share(rec["рев_недостачи"], tot_sh)
        rec["доля_оп_недостачи_%"] = _safe_share(rec["оп_недостачи"], tot_sh)
        rec["доля_рев_излишки_%"] = _safe_share(rec["рев_излишки"], tot_sur)
        rec["доля_оп_излишки_%"] = _safe_share(rec["оп_излишки"], tot_sur)
        rec["доля_рев_перекрытие_%"] = _safe_share(rec["рев_перекрытие"], tot_ov)
        rec["доля_оп_перекрытие_%"] = _safe_share(rec["оп_перекрытие"], tot_ov)

        # Кто внёс больший вклад в финансовый результат магазина (по |сальдо| роли)
        rev_abs = abs(rec["рев_сальдо"])
        op_abs = abs(rec["оп_сальдо"])
        if rev_abs == 0 and op_abs == 0:
            rec["доминанта"] = "нет данных"
            rec["доминанта_знак"] = "нейтрально"
        elif rev_abs >= op_abs:
            rec["доминанта"] = AUTHOR_REVISOR
            rec["доминанта_знак"] = "плюс" if rec["рев_сальдо"] >= 0 else "минус"
        else:
            rec["доминанта"] = AUTHOR_OPERATOR
            rec["доминанта_знак"] = "плюс" if rec["оп_сальдо"] >= 0 else "минус"
        rows.append(rec)
    return pd.DataFrame(rows).sort_values("недостачи", ascending=False)


def enrich_store_metrics_with_authors(
    store_metrics: pd.DataFrame,
    author_store: pd.DataFrame,
) -> pd.DataFrame:
    """Добавить колонки авторства к рейтингу магазинов (без изменения store_score)."""
    if store_metrics.empty or author_store.empty:
        return store_metrics
    cols = [
        "магазин",
        "рев_излишки", "оп_излишки",
        "рев_недостачи", "оп_недостачи",
        "рев_перекрытие", "оп_перекрытие",
        "доля_рев_недостачи_%", "доля_оп_недостачи_%",
        "доля_рев_излишки_%", "доля_оп_излишки_%",
        "доминанта", "доминанта_знак",
    ]
    avail = [c for c in cols if c in author_store.columns]
    merged = store_metrics.merge(author_store[avail], on="магазин", how="left")
    for c in avail:
        if c == "магазин" or c in ("доминанта", "доминанта_знак"):
            continue
        if c in merged.columns:
            merged[c] = merged[c].fillna(0)
    if "доминанта" in merged.columns:
        merged["доминанта"] = merged["доминанта"].fillna("нет данных")
    if "доминанта_знак" in merged.columns:
        merged["доминанта_знак"] = merged["доминанта_знак"].fillna("нейтрально")
    return merged


def detect_author_chains(df: pd.DataFrame) -> pd.DataFrame:
    """
    Цепочки пересортов между ролями внутри магазина по одному sku_key:
    недостача у одной роли → излишек у другой (позже), и наоборот.
    Инициатор = автор первого документа в цепочке.
    """
    cols = [
        "магазин", "sku_key", "наименование", "category_group",
        "док1", "дата1", "автор1", "знак1", "сумма1",
        "док2", "дата2", "автор2", "знак2", "сумма2",
        "сумма_цепочки", "тип_цепочки", "инициатор",
    ]
    if df.empty or "автор_дока" not in df.columns or "sku_key" not in df.columns:
        return pd.DataFrame(columns=cols)

    work = df.copy()
    work["дата_док"] = pd.to_datetime(work["дата_док"], errors="coerce")
    chains: List[Dict[str, Any]] = []

    for (store, sku), grp in work.groupby(["магазин", "sku_key"], dropna=False):
        if not sku or str(sku).startswith("__unmatched__"):
            # всё равно анализируем unmatched по ключу
            pass
        docs = (
            grp.groupby("документ", dropna=False)
            .agg(
                дата=("дата_док", "min"),
                автор=("автор_дока", "first"),
                излишек=("излишек_сумма", "sum"),
                недостача=("недостача_сумма", "sum"),
                наименование=("наименование", "first"),
                category_group=("category_group", "first"),
            )
            .reset_index()
            .sort_values("дата")
        )
        if len(docs) < 2:
            continue
        records = docs.to_dict("records")
        used_pairs = set()
        for i in range(len(records)):
            for j in range(i + 1, len(records)):
                a, b = records[i], records[j]
                if a["автор"] == b["автор"]:
                    continue
                if a["автор"] == AUTHOR_UNKNOWN or b["автор"] == AUTHOR_UNKNOWN:
                    continue
                # недостача → излишек
                if float(a["недостача"] or 0) > 0 and float(b["излишек"] or 0) > 0:
                    s1, s2 = float(a["недостача"]), float(b["излишек"])
                    chain_sum = min(s1, s2)
                    if chain_sum <= 0:
                        continue
                    key = (a["документ"], b["документ"], "sh_sur")
                    if key in used_pairs:
                        continue
                    used_pairs.add(key)
                    chains.append({
                        "магазин": store,
                        "sku_key": sku,
                        "наименование": a["наименование"] or b["наименование"],
                        "category_group": a.get("category_group") or b.get("category_group") or "",
                        "док1": a["документ"],
                        "дата1": a["дата"],
                        "автор1": a["автор"],
                        "знак1": "недостача",
                        "сумма1": round(s1, 2),
                        "док2": b["документ"],
                        "дата2": b["дата"],
                        "автор2": b["автор"],
                        "знак2": "излишек",
                        "сумма2": round(s2, 2),
                        "сумма_цепочки": round(chain_sum, 2),
                        "тип_цепочки": f"{a['автор']} недостача → {b['автор']} излишек",
                        "инициатор": a["автор"],
                    })
                # излишек → недостача
                if float(a["излишек"] or 0) > 0 and float(b["недостача"] or 0) > 0:
                    s1, s2 = float(a["излишек"]), float(b["недостача"])
                    chain_sum = min(s1, s2)
                    if chain_sum <= 0:
                        continue
                    key = (a["документ"], b["документ"], "sur_sh")
                    if key in used_pairs:
                        continue
                    used_pairs.add(key)
                    chains.append({
                        "магазин": store,
                        "sku_key": sku,
                        "наименование": a["наименование"] or b["наименование"],
                        "category_group": a.get("category_group") or b.get("category_group") or "",
                        "док1": a["документ"],
                        "дата1": a["дата"],
                        "автор1": a["автор"],
                        "знак1": "излишек",
                        "сумма1": round(s1, 2),
                        "док2": b["документ"],
                        "дата2": b["дата"],
                        "автор2": b["автор"],
                        "знак2": "недостача",
                        "сумма2": round(s2, 2),
                        "сумма_цепочки": round(chain_sum, 2),
                        "тип_цепочки": f"{a['автор']} излишек → {b['автор']} недостача",
                        "инициатор": a["автор"],
                    })

    if not chains:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame(chains).sort_values("сумма_цепочки", ascending=False).reset_index(drop=True)


def author_network_summary(
    role_stats: pd.DataFrame,
    store_stats: pd.DataFrame,
    chains: pd.DataFrame,
) -> Dict[str, Any]:
    """Краткая сводка для дашборда и выводов."""
    def _row(role: str) -> Dict[str, Any]:
        if role_stats.empty:
            return {}
        m = role_stats[role_stats["автор"] == role]
        return m.iloc[0].to_dict() if not m.empty else {}

    rev = _row(AUTHOR_REVISOR)
    op = _row(AUTHOR_OPERATOR)
    unk = _row(AUTHOR_UNKNOWN)

    rev_net = float(rev.get("сальдо", 0) or 0)
    op_net = float(op.get("сальдо", 0) or 0)
    if abs(rev_net) >= abs(op_net):
        network_dominant = AUTHOR_REVISOR
        network_sign = "плюс" if rev_net >= 0 else "минус"
    else:
        network_dominant = AUTHOR_OPERATOR
        network_sign = "плюс" if op_net >= 0 else "минус"

    stores_rev_dom = int((store_stats["доминанта"] == AUTHOR_REVISOR).sum()) if not store_stats.empty else 0
    stores_op_dom = int((store_stats["доминанта"] == AUTHOR_OPERATOR).sum()) if not store_stats.empty else 0

    return {
        "rev_docs": int(rev.get("документов", 0) or 0),
        "op_docs": int(op.get("документов", 0) or 0),
        "unk_docs": int(unk.get("документов", 0) or 0),
        "rev_surplus": float(rev.get("излишки", 0) or 0),
        "op_surplus": float(op.get("излишки", 0) or 0),
        "rev_shortage": float(rev.get("недостачи", 0) or 0),
        "op_shortage": float(op.get("недостачи", 0) or 0),
        "rev_overlap": float(rev.get("перекрытие", 0) or 0),
        "op_overlap": float(op.get("перекрытие", 0) or 0),
        "rev_net": rev_net,
        "op_net": op_net,
        "rev_share_shortage": float(rev.get("доля_недостачи_%", 0) or 0),
        "op_share_shortage": float(op.get("доля_недостачи_%", 0) or 0),
        "rev_share_surplus": float(rev.get("доля_излишки_%", 0) or 0),
        "op_share_surplus": float(op.get("доля_излишки_%", 0) or 0),
        "rev_share_overlap": float(rev.get("доля_перекрытие_%", 0) or 0),
        "op_share_overlap": float(op.get("доля_перекрытие_%", 0) or 0),
        "network_dominant": network_dominant,
        "network_sign": network_sign,
        "stores_rev_dominant": stores_rev_dom,
        "stores_op_dominant": stores_op_dom,
        "chains_count": len(chains) if chains is not None else 0,
        "chains_sum": float(chains["сумма_цепочки"].sum()) if chains is not None and not chains.empty else 0.0,
        "unknown_docs": int(unk.get("документов", 0) or 0),
    }


def format_author_conclusion(author_sum: Dict[str, Any]) -> Tuple[str, str]:
    return (
        "Ревизоры vs Операторы",
        (
            f"Документов ревизора (08:00:00): {author_sum.get('rev_docs', 0)} | "
            f"оператора: {author_sum.get('op_docs', 0)}"
            + (f" | неизвестно: {author_sum.get('unk_docs', 0)}" if author_sum.get("unk_docs") else "")
            + f". Излишки: ревизор {author_sum.get('rev_surplus', 0):,.0f} "
            f"({author_sum.get('rev_share_surplus', 0):.1f}%) / "
            f"оператор {author_sum.get('op_surplus', 0):,.0f} "
            f"({author_sum.get('op_share_surplus', 0):.1f}%). "
            f"Недостачи: ревизор {author_sum.get('rev_shortage', 0):,.0f} "
            f"({author_sum.get('rev_share_shortage', 0):.1f}%) / "
            f"оператор {author_sum.get('op_shortage', 0):,.0f} "
            f"({author_sum.get('op_share_shortage', 0):.1f}%). "
            f"Пересорт: ревизор {author_sum.get('rev_overlap', 0):,.0f} / "
            f"оператор {author_sum.get('op_overlap', 0):,.0f}. "
            f"По сети больший вклад в результат: {author_sum.get('network_dominant')} "
            f"({author_sum.get('network_sign')}). "
            f"Цепочек пересортов ревизор↔оператор: {author_sum.get('chains_count', 0)} "
            f"на {author_sum.get('chains_sum', 0):,.0f} руб. "
            "Детали — листы «Ревизоры vs Операторы» и «Цепочки пересортов»."
        ),
    )
