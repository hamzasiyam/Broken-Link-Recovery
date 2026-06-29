"""Wayback Machine URL and snapshot helpers.

This module contains the archive-specific logic that is shared by spreadsheet,
report, and download workflows. It normalizes user-entered URLs, calls
``waybackpack`` to list captures, and extracts capture timestamps from Wayback
snapshot links.
"""

import subprocess
from datetime import datetime
from urllib.parse import urlparse


WAYBACK_TIMESTAMP_FORMAT = "%Y%m%d%H%M%S"


def normalize_url(url: str, default_scheme: str = "http") -> str:
    """Normalize user input into a URL with an explicit scheme.

    Args:
        url: Website URL as a string. It may include ``http://`` or
            ``https://`` already, or it may be a bare domain.
        default_scheme: Scheme string to prepend when ``url`` has no scheme.

    Returns:
        A non-empty URL string that starts with ``http://`` or ``https://``.

    Raises:
        ValueError: If ``url`` is empty or only whitespace.
    """
    # Strip whitespace so accidental spaces in the GUI field do not become part
    # of the URL sent to waybackpack.
    cleaned_url = url.strip()

    # If the user did not type a URL, stop before downstream tools receive bad
    # input.
    if not cleaned_url:
        raise ValueError("URL is required.")

    # If no scheme is provided, default to HTTP because archived sites often
    # have older captures before HTTPS was adopted.
    if not cleaned_url.startswith(("http://", "https://")):
        cleaned_url = f"{default_scheme}://{cleaned_url}"
    return cleaned_url


def domain_from_url(url: str) -> str:
    """Extract the domain portion of a URL for filenames and report labels.

    Args:
        url: Website URL string, with or without a scheme.

    Returns:
        Domain string such as ``example.com``.
    """
    # Normalize first so ``urlparse`` can reliably place the domain in
    # ``netloc`` rather than treating it as a path.
    parsed = urlparse(normalize_url(url))
    return parsed.netloc or parsed.path.split("/")[0]


def get_snapshots(url: str) -> list[str]:
    """Fetch Wayback Machine snapshot links for a URL.

    Args:
        url: Website URL string to pass to ``waybackpack --list``.

    Returns:
        A list of snapshot URL strings. Empty output becomes an empty list.

    Raises:
        RuntimeError: If ``waybackpack`` is missing or exits with an error.
        ValueError: If the URL cannot be normalized.
    """
    # Normalize once before invoking the external command so every caller gets
    # consistent URL handling.
    normalized_url = normalize_url(url)
    try:
        # Capture stdout/stderr because GUI workflows display errors in message
        # boxes instead of streaming terminal output.
        result = subprocess.run(
            ["waybackpack", normalized_url, "--list"],
            capture_output=True,
            text=True,
            check=False,
        )
    except FileNotFoundError as exc:
        # If waybackpack is not installed, raise a clear runtime error that the
        # GUI can show directly to the user.
        raise RuntimeError("waybackpack is not installed or is not on PATH.") from exc

    # If waybackpack returned a non-zero exit code, include stderr so the user
    # can diagnose rate limits, network errors, or invalid URLs.
    if result.returncode != 0:
        detail = result.stderr.strip() or "Unknown waybackpack error."
        raise RuntimeError(f"Error fetching snapshots: {detail}")

    # Remove blank lines so downstream spreadsheet/report code receives only
    # usable snapshot links.
    return [line.strip() for line in result.stdout.splitlines() if line.strip()]


def extract_date_from_link(link: str) -> datetime:
    """Extract a datetime from a Wayback Machine snapshot link.

    Args:
        link: Snapshot URL string whose fifth slash-delimited segment is a
            Wayback timestamp in ``YYYYMMDDHHMMSS`` format.

    Returns:
        A ``datetime`` representing the capture timestamp.

    Raises:
        ValueError: If the link does not contain a timestamp in the expected
        position or the timestamp cannot be parsed.
    """
    # Wayback capture URLs place the timestamp after the "/web/" segment.
    parts = link.split("/")

    # If the URL does not have enough segments, it cannot be a normal capture
    # link and should fail loudly.
    if len(parts) <= 4:
        raise ValueError(f"Snapshot link does not contain a Wayback timestamp: {link}")
    return datetime.strptime(parts[4], WAYBACK_TIMESTAMP_FORMAT)


def split_date_components(date: datetime) -> tuple[str, str]:
    """Split a datetime into display-friendly date and time strings.

    Args:
        date: ``datetime`` object to format.

    Returns:
        Tuple containing ``("Month Day, Year", "Hour:Minute:Second AM/PM")``.
    """
    # Keep report date and time values in separate columns/paragraphs.
    return date.strftime("%B %d, %Y"), date.strftime("%I:%M:%S %p")


def format_date_columns(date: datetime) -> dict[str, str]:
    """Format a datetime for the spreadsheet date columns.

    Args:
        date: ``datetime`` object to format.

    Returns:
        Dictionary with Excel column names mapped to formatted strings.
    """
    # Return a dictionary because pandas can expand this into two named columns.
    return {
        "Month Day, Year": date.strftime("%B %d, %Y"),
        "Hour, Minute, Second AM/PM": date.strftime("%I:%M:%S %p"),
    }


def convert_to_http(link: str) -> str:
    """Convert an HTTPS snapshot URL to HTTP while leaving other text unchanged.

    Args:
        link: Snapshot URL string.

    Returns:
        URL string with ``https://`` replaced by ``http://``.
    """
    # Some recovery workflows prefer HTTP links because older archived assets
    # can reference HTTP paths.
    return link.replace("https://", "http://")
