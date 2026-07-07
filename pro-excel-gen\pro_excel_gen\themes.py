from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ThemeProfile:
    name: str
    accent: str
    accent_2: str
    accent_3: str
    background: str
    header_bg: str
    header_fg: str
    grid: str
    good: str
    warn: str
    bad: str


THEMES: dict[str, ThemeProfile] = {
    "corporate_formal": ThemeProfile(
        name="corporate_formal",
        accent="#0F4C81",
        accent_2="#2F75B5",
        accent_3="#8FAADC",
        background="#FFFFFF",
        header_bg="#0F4C81",
        header_fg="#FFFFFF",
        grid="#D9E2F3",
        good="#70AD47",
        warn="#ED7D31",
        bad="#C00000",
    ),
    "minimal_clean": ThemeProfile(
        name="minimal_clean",
        accent="#2B2B2B",
        accent_2="#595959",
        accent_3="#A6A6A6",
        background="#FFFFFF",
        header_bg="#F2F2F2",
        header_fg="#222222",
        grid="#E5E5E5",
        good="#5B8C5A",
        warn="#B07D25",
        bad="#A94442",
    ),
    "modern_tech": ThemeProfile(
        name="modern_tech",
        accent="#003049",
        accent_2="#1D70A2",
        accent_3="#78C0E0",
        background="#F7FBFF",
        header_bg="#003049",
        header_fg="#FFFFFF",
        grid="#D8E6F0",
        good="#2A9D8F",
        warn="#F4A261",
        bad="#E63946",
    ),
    "bold_impact": ThemeProfile(
        name="bold_impact",
        accent="#7A1F1F",
        accent_2="#C1121F",
        accent_3="#F28482",
        background="#FFF8F7",
        header_bg="#7A1F1F",
        header_fg="#FFFFFF",
        grid="#F2D7D5",
        good="#588157",
        warn="#D97706",
        bad="#B00020",
    ),
    "elegant_luxury": ThemeProfile(
        name="elegant_luxury",
        accent="#3D2C1E",
        accent_2="#7B5E3B",
        accent_3="#C6A969",
        background="#FFFDF8",
        header_bg="#3D2C1E",
        header_fg="#FDF8ED",
        grid="#E7DDC7",
        good="#728C69",
        warn="#C08B2C",
        bad="#9E2A2B",
    ),
}


def get_theme(name: str | None = None) -> ThemeProfile:
    if not name:
        return THEMES["corporate_formal"]
    return THEMES.get(name, THEMES["corporate_formal"])


def list_themes() -> list[str]:
    return list(THEMES)
