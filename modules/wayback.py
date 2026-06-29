import subprocess
from datetime import datetime
from urllib.parse import urlparse


WAYBACK_TIMESTAMP_FORMAT = "%Y%m%d%H%M%S"


def normalize_url(url: str, default_scheme: str = "http") -> str:
    cleaned_url = url.strip()
    if not cleaned_url:
        raise ValueError("URL is required.")
    if not cleaned_url.startswith(("http://", "https://")):
        cleaned_url = f"{default_scheme}://{cleaned_url}"
    return cleaned_url


def domain_from_url(url: str) -> str:
    parsed = urlparse(normalize_url(url))
    return parsed.netloc or parsed.path.split("/")[0]


def get_snapshots(url: str) -> list[str]:
    normalized_url = normalize_url(url)
    try:
        result = subprocess.run(
            ["waybackpack", normalized_url, "--list"],
            capture_output=True,
            text=True,
            check=False,
        )
    except FileNotFoundError as exc:
        raise RuntimeError("waybackpack is not installed or is not on PATH.") from exc

    if result.returncode != 0:
        detail = result.stderr.strip() or "Unknown waybackpack error."
        raise RuntimeError(f"Error fetching snapshots: {detail}")

    return [line.strip() for line in result.stdout.splitlines() if line.strip()]


def extract_date_from_link(link: str) -> datetime:
    parts = link.split("/")
    if len(parts) <= 4:
        raise ValueError(f"Snapshot link does not contain a Wayback timestamp: {link}")
    return datetime.strptime(parts[4], WAYBACK_TIMESTAMP_FORMAT)


def split_date_components(date: datetime) -> tuple[str, str]:
    return date.strftime("%B %d, %Y"), date.strftime("%I:%M:%S %p")


def format_date_columns(date: datetime) -> dict[str, str]:
    return {
        "Month Day, Year": date.strftime("%B %d, %Y"),
        "Hour, Minute, Second AM/PM": date.strftime("%I:%M:%S %p"),
    }


def convert_to_http(link: str) -> str:
    return link.replace("https://", "http://")

