"""Invoice data, Serbian amount-in-words, and bill-number storage."""

from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
OUTPUT_DIR = APP_DIR / "izdati"
LOGO_PATH = APP_DIR / "assets" / "veritas_logo.png"
STATE_PATH = DATA_DIR / "stanje.json"
LEGACY_COUNTER = APP_DIR.parent / "Next bill number.txt"

MONTHS = (
    "januar",
    "februar",
    "mart",
    "april",
    "maj",
    "jun",
    "jul",
    "avgust",
    "septembar",
    "oktobar",
    "novembar",
    "decembar",
)

_ONES_M = (
    "",
    "jedan",
    "dva",
    "tri",
    "četiri",
    "pet",
    "šest",
    "sedam",
    "osam",
    "devet",
)
_ONES_F = (
    "",
    "jedna",
    "dve",
    "tri",
    "četiri",
    "pet",
    "šest",
    "sedam",
    "osam",
    "devet",
)
_TEENS = (
    "deset",
    "jedanaest",
    "dvanaest",
    "trinaest",
    "četrnaest",
    "petnaest",
    "šesnaest",
    "sedamnaest",
    "osamnaest",
    "devetnaest",
)
_TENS = (
    "",
    "",
    "dvadeset",
    "trideset",
    "četrdeset",
    "pedeset",
    "šezdeset",
    "sedamdeset",
    "osamdeset",
    "devedeset",
)
_HUNDREDS = (
    "",
    "sto",
    "dvesta",
    "trista",
    "četiristo",
    "petsto",
    "šeststo",
    "sedamsto",
    "osamsto",
    "devetsto",
)


@dataclass(frozen=True)
class Party:
    name: str
    address: str
    registration: str
    tax: str


@dataclass(frozen=True)
class Profile:
    key: str
    title: str
    currency: str
    symbol: str
    price_header: str
    total_header: str
    grand_label: str
    place: str
    buyer: Party
    bank_lines: tuple[str, ...]
    payment_hint: str
    use_words: bool


@dataclass
class Line:
    name: str
    qty: Decimal
    price: Decimal

    @property
    def total(self) -> Decimal:
        return (self.qty * self.price).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


@dataclass
class Invoice:
    number: int
    country: str
    issue_date: date
    due_date: date
    place: str
    period: str
    buyer: Party
    lines: list[Line]

    @property
    def profile(self) -> Profile:
        return PROFILES[self.country]

    @property
    def total(self) -> Decimal:
        return sum((line.total for line in self.lines), Decimal("0.00"))


SELLER = Party(
    name="JANKO JOVIČIĆ PR RAČUNARSKO PROGRAMIRANJE VERITAS SOFTWARE",
    address="Mišarska 13, 14000 Valjevo Srbija",
    registration="67882083",
    tax="114825676",
)

ISSUER = {
    "company": "Veritas Software",
    "person": "Janko Jovičić",
    "id_card": "011366776",
    "payment": "transakcijski račun",
    "note": "Radnja nije u sistemu PDV-a",
}

PROFILES: dict[str, Profile] = {
    "slovenija": Profile(
        key="slovenija",
        title="Slovenija",
        currency="EUR",
        symbol="€",
        price_header="CENA €",
        total_header="UKUPNO €",
        grand_label="UKUPNO(EUR)",
        place="Velenje",
        buyer=Party(
            name="IN.SIST d.o.o.",
            address="Koroška cesta 62B, Velenje, 3320 Velenje Slovenia",
            registration="2147645000",
            tax="SI21135991",
        ),
        bank_lines=(
            "OTP banka Srbija akcionarsko društvo Novi Sad",
            "SWIFT adresa: OTPVRS22",
            "Broj računa: RS35325960170009772731",
            "Naziv: JANKO JOVIČIĆ PR RAČUNARSKO PROGRAMIRANJE VERITAS SOFTWARE",
        ),
        payment_hint="Payment: IBAN RS35325960170009772731 · SWIFT OTPVRS22 · OTP banka Srbija",
        use_words=False,
    ),
    "srbija": Profile(
        key="srbija",
        title="Srbija",
        currency="RSD",
        symbol="",
        price_header="CENA",
        total_header="UKUPNO (RSD)",
        grand_label="UKUPNO(RSD)",
        place="Beograd",
        buyer=Party(
            name="Landfreight d.o.o.",
            address="Cara Lazara 5-7, 11000 Beograd",
            registration="21508730",
            tax="111602061",
        ),
        bank_lines=(
            "OTP banka Srbija akcionarsko društvo Novi Sad",
            "Broj računa: 325-9500700221537-34",
            "Naziv: JANKO JOVIČIĆ PR RAČUNARSKO PROGRAMIRANJE VERITAS SOFTWARE",
        ),
        payment_hint="Payment: account 325-9500700221537-34 · OTP banka Srbija",
        use_words=True,
    ),
}


def triplet(n: int, feminine: bool = False) -> str:
    ones = _ONES_F if feminine else _ONES_M
    words: list[str] = []
    hundreds, rest = divmod(n, 100)
    tens, unit = divmod(rest, 10)
    if hundreds:
        words.append(_HUNDREDS[hundreds])
    if tens == 1:
        words.append(_TEENS[unit])
    else:
        if tens:
            words.append(_TENS[tens])
        if unit:
            words.append(ones[unit])
    return " ".join(words)


def _thousands_phrase(n: int) -> str:
    if n == 1:
        return "hiljadu"
    last_two = n % 100
    last = n % 10
    noun = "hiljade" if last in (2, 3, 4) and not 10 <= last_two <= 20 else "hiljada"
    return f"{triplet(n, feminine=True)} {noun}"


def _millions_phrase(n: int) -> str:
    if n == 1:
        return "milion"
    last_two = n % 100
    last = n % 10
    noun = "milion" if last == 1 and not 10 <= last_two <= 20 else "miliona"
    return f"{triplet(n)} {noun}"


def int_to_words(n: int) -> str:
    if n == 0:
        return "nula"
    if n < 0:
        return "minus " + int_to_words(-n)
    millions, rest_n = divmod(n, 1_000_000)
    thousands, rest = divmod(rest_n, 1000)
    parts: list[str] = []
    if millions:
        parts.append(_millions_phrase(millions))
    if thousands:
        parts.append(_thousands_phrase(thousands))
    if rest:
        parts.append(triplet(rest))
    return " ".join(parts)


def _title_sr(text: str) -> str:
    return " ".join(word[:1].upper() + word[1:] for word in text.split())


def _para_phrase(cents: int) -> str:
    if cents == 0:
        return "nula PARE"
    last_two = cents % 100
    last = cents % 10
    if last == 1 and last_two != 11:
        noun = "para"
    elif last in (2, 3, 4) and not 10 <= last_two <= 20:
        noun = "pare"
    else:
        noun = "para"
    return f"{triplet(cents, feminine=True)} {noun}"


def amount_in_words(amount: Decimal, currency: str) -> str:
    amount = abs(amount).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    cents_total = int((amount * 100).to_integral_value(rounding=ROUND_HALF_UP))
    whole, cents = divmod(cents_total, 100)
    return f"Slovima: {_title_sr(int_to_words(whole))} {currency} i {_para_phrase(cents)}"


def format_date(value: date) -> str:
    return value.strftime("%d.%m.%Y")


def month_end(value: date) -> date:
    if value.month == 12:
        next_month = date(value.year + 1, 1, 1)
    else:
        next_month = date(value.year, value.month + 1, 1)
    return next_month - timedelta(days=1)


def default_issue_date(today: date | None = None) -> date:
    """Last day of the current month, matching the existing invoices."""
    return month_end(today or date.today())


def default_due_date(issue: date) -> date:
    """10th day of the month after the invoice date."""
    if issue.month == 12:
        return date(issue.year + 1, 1, 10)
    return date(issue.year, issue.month + 1, 10)


def builtin_customers() -> list[dict]:
    slovenia = PROFILES["slovenija"]
    serbia = PROFILES["srbija"]
    return [
        {
            "id": "insist",
            "naziv": slovenia.buyer.name,
            "adresa": slovenia.buyer.address,
            "mb": slovenia.buyer.registration,
            "pib": slovenia.buyer.tax,
            "mesto": slovenia.place,
            "format": "slovenija",
        },
        {
            "id": "landfreight",
            "naziv": serbia.buyer.name,
            "adresa": serbia.buyer.address,
            "mb": serbia.buyer.registration,
            "pib": serbia.buyer.tax,
            "mesto": serbia.place,
            "format": "srbija",
        },
    ]


def upsert_customer(customers: list[dict], record: dict) -> list[dict]:
    name = record["naziv"].strip().lower()
    country = record["format"]
    for existing in customers:
        same_id = bool(record.get("id")) and existing.get("id") == record["id"]
        same_name = existing.get("naziv", "").strip().lower() == name and existing.get("format") == country
        if same_id or same_name:
            kept_id = existing.get("id") or record.get("id")
            existing.update(
                {
                    "naziv": record["naziv"].strip(),
                    "adresa": record.get("adresa", "").strip(),
                    "mb": record.get("mb", "").strip(),
                    "pib": record.get("pib", "").strip(),
                    "mesto": record.get("mesto", "").strip(),
                    "format": country,
                }
            )
            if kept_id:
                existing["id"] = kept_id
            return customers
    customers.append(
        {
            "id": record.get("id") or f"k{len(customers) + 1}",
            "naziv": record["naziv"].strip(),
            "adresa": record.get("adresa", "").strip(),
            "mb": record.get("mb", "").strip(),
            "pib": record.get("pib", "").strip(),
            "mesto": record.get("mesto", "").strip(),
            "format": country,
        }
    )
    return customers


def period_label(value: date) -> str:
    return f"{MONTHS[value.month - 1]} {value.year}"


def parse_date(text: str) -> date:
    raw = text.strip()
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
    raise ValueError("Enter the date as 30.09.2026.")


def parse_decimal(text: str) -> Decimal:
    raw = (
        text.strip()
        .replace(" ", "")
        .replace("€", "")
        .replace("RSD", "")
        .replace("rsd", "")
        .replace("eur", "")
        .replace("EUR", "")
    )
    if not raw:
        raise ValueError("Enter an amount.")
    if "," in raw and "." in raw:
        if raw.rfind(",") > raw.rfind("."):
            raw = raw.replace(".", "").replace(",", ".")
        else:
            raw = raw.replace(",", "")
    elif "," in raw:
        raw = raw.replace(",", ".")
    try:
        return Decimal(raw)
    except InvalidOperation as exc:
        raise ValueError(f"Invalid number: {text}") from exc


def parse_money(text: str) -> Decimal:
    return parse_decimal(text).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def fmt_money(amount: Decimal) -> str:
    quantized = amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    sign = "-" if quantized < 0 else ""
    whole, frac = f"{abs(quantized):.2f}".split(".")
    groups: list[str] = []
    while whole:
        groups.append(whole[-3:])
        whole = whole[:-3]
    return sign + ",".join(reversed(groups)) + "." + frac


def fmt_qty(amount: Decimal) -> str:
    quantized = amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if quantized == quantized.to_integral():
        return str(int(quantized))
    return fmt_money(quantized)


def money_with_currency(amount: Decimal, profile: Profile) -> str:
    text = fmt_money(amount)
    if profile.symbol:
        return f"{text} {profile.symbol}"
    return f"{text} {profile.currency}"


def default_state() -> dict:
    number = 10
    if LEGACY_COUNTER.exists():
        raw = LEGACY_COUNTER.read_text(encoding="utf-8").strip()
        if raw.isdigit():
            number = int(raw)
    return {
        "sledeci_broj": number,
        "istorija": [],
        "kupci": {},
        "kupci_lista": builtin_customers(),
        "poslednji_format": "slovenija",
    }


def load_state() -> dict:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not STATE_PATH.exists():
        return default_state()
    try:
        state = json.loads(STATE_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return default_state()
    base = default_state()
    if not isinstance(state, dict):
        return base
    had_customers = isinstance(state.get("kupci_lista"), list)
    base.update({key: state[key] for key in base if key in state})
    if not had_customers:
        base["kupci_lista"] = _customers_from_legacy(base.get("kupci"))
    if not isinstance(base["sledeci_broj"], int) or base["sledeci_broj"] < 1:
        base["sledeci_broj"] = default_state()["sledeci_broj"]
    if not isinstance(base["istorija"], list):
        base["istorija"] = []
    if not isinstance(base["kupci"], dict):
        base["kupci"] = {}
    if not isinstance(base.get("kupci_lista"), list):
        base["kupci_lista"] = builtin_customers()
    if base.get("poslednji_format") not in PROFILES:
        base["poslednji_format"] = "slovenija"
    return base


def save_state(state: dict) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    STATE_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    try:
        LEGACY_COUNTER.write_text(str(state["sledeci_broj"]), encoding="utf-8")
    except OSError:
        pass


def remember_invoice(invoice: Invoice, pdf_path: Path) -> dict:
    state = load_state()
    if invoice.number >= int(state["sledeci_broj"]):
        state["sledeci_broj"] = invoice.number + 1
    state["poslednji_format"] = invoice.country
    if not isinstance(state.get("kupci_lista"), list):
        state["kupci_lista"] = builtin_customers()
    upsert_customer(
        state["kupci_lista"],
        {
            "naziv": invoice.buyer.name,
            "adresa": invoice.buyer.address,
            "mb": invoice.buyer.registration,
            "pib": invoice.buyer.tax,
            "mesto": invoice.place,
            "format": invoice.country,
        },
    )
    state["istorija"].insert(
        0,
        {
            "broj": invoice.number,
            "datum": invoice.issue_date.isoformat(),
            "rok": invoice.due_date.isoformat(),
            "format": invoice.profile.title,
            "zemlja": invoice.country,
            "kupac": invoice.buyer.name,
            "iznos": str(invoice.total),
            "valuta": invoice.profile.currency,
            "stavke": [line.name for line in invoice.lines],
            "pdf": str(pdf_path),
        },
    )
    state["istorija"] = state["istorija"][:300]
    save_state(state)
    return state


def _customers_from_legacy(legacy: object) -> list[dict]:
    customers = builtin_customers()
    if not isinstance(legacy, dict):
        return customers
    for key, row in legacy.items():
        if key not in PROFILES or not isinstance(row, dict) or not row.get("naziv"):
            continue
        for customer in customers:
            if customer["format"] == key:
                customer["naziv"] = str(row.get("naziv") or customer["naziv"])
                customer["adresa"] = str(row.get("adresa") or customer["adresa"])
                customer["mb"] = str(row.get("mb") or customer["mb"])
                customer["pib"] = str(row.get("pib") or customer["pib"])
                customer["mesto"] = str(row.get("mesto") or customer["mesto"])
                break
    return customers


def meta_lines(invoice: Invoice) -> list[tuple[str, bool]]:
    return [
        (f"Mesto i mesec prometa: {invoice.period}, {invoice.place}", False),
        (f"Datum računa: {format_date(invoice.issue_date)}", False),
        (f"Rok plaćanja: {format_date(invoice.due_date)}", True),
    ]


def footer_lines() -> list[tuple[str, bool]]:
    return [
        (f"Izdao za: {ISSUER['company']}", False),
        (f"Ime i prezime: {ISSUER['person']}", False),
        (f"Broj lične karte: {ISSUER['id_card']}", False),
        (f"Način plaćanja: {ISSUER['payment']}", False),
        (f"Napomena: {ISSUER['note']}", True),
    ]


def output_pdf_path(invoice: Invoice) -> Path:
    folder = OUTPUT_DIR / str(invoice.issue_date.year)
    folder.mkdir(parents=True, exist_ok=True)
    raw = f"Faktura Veritas - {invoice.profile.title} - {invoice.number} - {invoice.period}"
    safe = "".join(ch for ch in raw if ch not in '\\/:*?"<>|').strip()
    number = 1
    while True:
        suffix = "" if number == 1 else f" ({number})"
        pdf_path = folder / f"{safe}{suffix}.pdf"
        if not pdf_path.exists():
            return pdf_path
        number += 1


if __name__ == "__main__":
    assert amount_in_words(Decimal("20000"), "RSD") == "Slovima: Dvadeset Hiljada RSD i nula PARE"
    assert (
        amount_in_words(Decimal("3160"), "RSD")
        == "Slovima: Tri Hiljade Sto Šezdeset RSD i nula PARE"
    )
    assert amount_in_words(Decimal("60"), "RSD") == "Slovima: Šezdeset RSD i nula PARE"
    assert amount_in_words(Decimal("2000"), "RSD") == "Slovima: Dve Hiljade RSD i nula PARE"
    assert amount_in_words(Decimal("20000.50"), "RSD") == (
        "Slovima: Dvadeset Hiljada RSD i pedeset para"
    )
    assert parse_money("3,160.50") == Decimal("3160.50")
    assert parse_money("3160,5") == Decimal("3160.50")
    assert fmt_money(Decimal("3160")) == "3,160.00"
    assert default_issue_date(date(2026, 10, 12)) == date(2026, 10, 31)
    assert default_due_date(date(2026, 10, 31)) == date(2026, 11, 10)
    assert default_due_date(date(2026, 9, 30)) == date(2026, 10, 10)
    assert default_due_date(date(2026, 12, 31)) == date(2027, 1, 10)
    assert builtin_customers()[0]["format"] == "slovenija"
    assert builtin_customers()[1]["pib"] == "111602061"
    print("model ok")
