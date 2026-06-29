"""Proxy environment helpers used by archive download workflows.

The downloader relies on standard ``http_proxy`` and ``https_proxy``
environment variables because command-line tools like ``wget`` can read those
values without needing a custom integration.
"""

import os


def connect_proxy(
    proxy_type: str,
    domain_name: str,
    proxy_port: str,
    proxy_username: str,
    proxy_password: str,
) -> str:
    """Set HTTP and HTTPS proxy environment variables.

    Args:
        proxy_type: Proxy scheme string such as ``"HTTP"`` or ``"SOCKS5"``.
        domain_name: Proxy host/domain string.
        proxy_port: Proxy port string.
        proxy_username: Username string for authenticated proxies.
        proxy_password: Password string for authenticated proxies.

    Returns:
        The proxy URL string that was written into the environment.
    """
    # Build the URL format understood by wget and many HTTP libraries.
    proxy_url = f"{proxy_type.lower()}://{proxy_username}:{proxy_password}@{domain_name}:{proxy_port}"

    # Store both variables so HTTP and HTTPS requests route through the same
    # configured proxy.
    os.environ["http_proxy"] = proxy_url
    os.environ["https_proxy"] = proxy_url
    return proxy_url


def disable_proxy() -> None:
    """Remove proxy environment variables for the current process.

    Args:
        None.

    Returns:
        None. Environment variables are removed if present.
    """
    # ``pop(..., None)`` avoids KeyError if a proxy was already disabled.
    os.environ.pop("http_proxy", None)
    os.environ.pop("https_proxy", None)


def current_proxy_info() -> dict[str, str | None]:
    """Read the active proxy environment values.

    Args:
        None.

    Returns:
        Dictionary with ``http_proxy`` and ``https_proxy`` values, each either a
        string or ``None``.
    """
    # Return a small dictionary so the GUI can print both values consistently.
    return {
        "http_proxy": os.environ.get("http_proxy"),
        "https_proxy": os.environ.get("https_proxy"),
    }
