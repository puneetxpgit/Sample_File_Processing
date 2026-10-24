from __future__ import annotations

import re

from processor.read_inputs import MopDocument, SheetRow


def match_rows(
    rows: list[SheetRow],
    mops: list[MopDocument],
    strategy: str,
) -> tuple[list[tuple[SheetRow, MopDocument | None, str]], list[MopDocument]]:
    """Pair each schedule row with one MOP.

    Fuze ID wins when the document has one. Site name is the fallback used by
    files such as the GILLETT MOP, which has no Fuze ID. One MOP may cover
    several rows for the same site.
    """
    usable = [mop for mop in mops if not mop.read_error]
    pairs: list[tuple[SheetRow, MopDocument | None, str]] = []
    used: set[int] = set()
    considered: set[int] = set()
    for row in rows:
        chosen, reason, candidates = _match_one(row, usable, strategy)
        if chosen is not None:
            used.add(id(chosen))
        if reason == "ambiguous_mop":
            names = ", ".join(mop.path.name for mop in candidates)
            reason = f"ambiguous_mop|{names}"
            considered.update(id(mop) for mop in candidates)
        pairs.append((row, chosen, reason))
    unmatched = [mop for mop in usable if id(mop) not in used and id(mop) not in considered]
    return pairs, unmatched


def _match_one(row: SheetRow, mops: list[MopDocument], strategy: str) -> tuple[MopDocument | None, str, list[MopDocument]]:
    use_id = strategy in {"both", "filename", "content"}
    use_site = strategy in {"both", "filename", "content"}
    if strategy == "filename":
        id_hits = [mop for mop in mops if _filename_has_fuze(mop, row.fuze_id)]
        site_hits = [mop for mop in mops if _filename_has_site(mop, row.site_name)]
    elif strategy == "content":
        id_hits = [mop for mop in mops if row.fuze_id and row.fuze_id in mop.fuze_ids]
        site_hits = [mop for mop in mops if _content_has_site(mop, row.site_name)]
    else:
        id_hits = [
            mop
            for mop in mops
            if (row.fuze_id and row.fuze_id in mop.fuze_ids) or _filename_has_fuze(mop, row.fuze_id)
        ]
        site_hits = [
            mop
            for mop in mops
            if _content_has_site(mop, row.site_name) or _filename_has_site(mop, row.site_name)
        ]

    if row.fuze_id:
        site_hits = [
            mop
            for mop in site_hits
            if not mop.fuze_ids or row.fuze_id in mop.fuze_ids
        ]

    if use_id and len(id_hits) == 1:
        return id_hits[0], "fuze_id", []
    if use_id and len(id_hits) > 1:
        return None, "ambiguous_mop", id_hits
    if use_site and len(site_hits) == 1:
        return site_hits[0], "site_name", []
    if use_site and len(site_hits) > 1:
        return None, "ambiguous_mop", site_hits
    return None, "missing_mop", []


def filename_digit_runs(stem: str) -> list[str]:
    """Digit groups in a filename, excluding a YYYY-MM-DD or YYYY/MM/DD date.

    Joining every digit would let Fuze 4520 match MOP_12345_2026-10-20. A date
    written with separators is not a Fuze ID, so 2026 from 2026-10-20 is left out.
    """
    scrubbed = re.sub(r"\d{4}[-/]\d{1,2}[-/]\d{1,2}", " ", stem)
    return re.findall(r"\d+", scrubbed)


def _filename_has_fuze(mop: MopDocument, fuze_id: str) -> bool:
    if not fuze_id:
        return False
    return fuze_id in filename_digit_runs(mop.path.stem)


def _filename_has_site(mop: MopDocument, site_name: str) -> bool:
    tokens = _tokens(site_name)
    if not tokens:
        return False
    return all(token in _tokens(mop.path.stem) for token in tokens)


def _content_has_site(mop: MopDocument, site_name: str) -> bool:
    if not site_name or not mop.site_name:
        return False
    return _compact(site_name) == _compact(mop.site_name)


def _tokens(value: str) -> list[str]:
    parts = [part for part in re.split(r"[^a-z0-9]+", value.casefold()) if part]
    if len(parts) <= 1:
        return parts
    return [part for part in parts if len(part) >= 2] or parts


def _compact(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", value.casefold())
