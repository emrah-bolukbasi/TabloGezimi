"""Table Navigation add-on translations; NVDA UI language with English fallback."""
import gettext
from pathlib import Path
import languageHandler

_LOCALE = Path(__file__).resolve().parents[2] / "locale"
_SUPPORTED = {"tr", "en", "es", "ru", "ar", "zh_CN", "de", "fr", "it"}

def _language():
    try:
        value = languageHandler.getLanguage().replace("-", "_")
    except Exception:
        return "en"
    language = value.split("_")[0].lower()
    if language == "zh":
        return "zh_CN"
    return language if language in _SUPPORTED else "en"

def _load():
    language = _language()
    try:
        with (_LOCALE / language / "LC_MESSAGES" / "nvda.mo").open("rb") as stream:
            return gettext.GNUTranslations(stream)
    except (OSError, ValueError):
        with (_LOCALE / "en" / "LC_MESSAGES" / "nvda.mo").open("rb") as stream:
            return gettext.GNUTranslations(stream)

_translator = _load()

def _(message):
    return _translator.gettext(message)
