"""Footer (chân trang) của 4 file xuất: Mã biểu mẫu | LBH + NHL | Trang x/y."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict

DEFAULT_CODES = {"dk": "BM/CL/05.01", "tl": "BM/CL/05.02", "lt": "BM/CL/05.04a", "tk": ""}


@dataclass
class FooterOptions:
    codes: Dict[str, str] = field(default_factory=lambda: dict(DEFAULT_CODES))
    lbh: str = "06"
    nhl: str = "02/03/2026"

    def left(self, key: str) -> str:
        return (self.codes.get(key) or "").strip()

    def center(self) -> str:
        parts = []
        if self.lbh.strip():
            parts.append(f"LBH: {self.lbh.strip()}")
        if self.nhl.strip():
            parts.append(f"NHL: {self.nhl.strip()}")
        return "      ".join(parts)

    @classmethod
    def from_opts(cls, o: dict) -> "FooterOptions":
        g = lambda k, d: str(o.get(k, d) if o.get(k, None) is not None else d)
        return cls(codes={k: g("ft_" + k, v) for k, v in DEFAULT_CODES.items()},
                   lbh=g("ft_lbh", "06"), nhl=g("ft_nhl", "02/03/2026"))
