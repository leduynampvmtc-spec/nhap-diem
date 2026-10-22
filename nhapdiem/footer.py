"""Footer (chân trang) của 4 file xuất: Mã biểu mẫu | LBH + NHL | Trang x/y — mỗi file có bộ riêng."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict

KEYS = ("dk", "tl", "lt", "tk")
DEFAULT_CODES = {"dk": "BM/CL/05.01", "tl": "BM/CL/05.02", "lt": "BM/CL/05.04a", "tk": ""}
DEFAULT_LBH = {"dk": "06", "tl": "06", "lt": "06", "tk": "06"}
DEFAULT_NHL = {"dk": "02/03/2026", "tl": "02/03/2026", "lt": "02/03/2026", "tk": "02/03/2026"}


@dataclass
class FooterOptions:
    codes: Dict[str, str] = field(default_factory=lambda: dict(DEFAULT_CODES))
    lbh: Dict[str, str] = field(default_factory=lambda: dict(DEFAULT_LBH))
    nhl: Dict[str, str] = field(default_factory=lambda: dict(DEFAULT_NHL))

    def left(self, key: str) -> str:
        return (self.codes.get(key) or "").strip()

    def center(self, key: str) -> str:
        parts = []
        lbh, nhl = (self.lbh.get(key) or "").strip(), (self.nhl.get(key) or "").strip()
        if lbh:
            parts.append(f"LBH: {lbh}")
        if nhl:
            parts.append(f"NHL: {nhl}")
        return "      ".join(parts)

    @classmethod
    def from_opts(cls, o: dict) -> "FooterOptions":
        def g(k, d):
            v = o.get(k)
            return d if v is None else str(v)

        return cls(codes={k: g("ft_" + k, DEFAULT_CODES[k]) for k in KEYS},
                   lbh={k: g(f"ft_lbh_{k}", g("ft_lbh", DEFAULT_LBH[k])) for k in KEYS},
                   nhl={k: g(f"ft_nhl_{k}", g("ft_nhl", DEFAULT_NHL[k])) for k in KEYS})
