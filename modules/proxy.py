import os


def connect_proxy(
    proxy_type: str,
    domain_name: str,
    proxy_port: str,
    proxy_username: str,
    proxy_password: str,
) -> str:
    proxy_url = f"{proxy_type.lower()}://{proxy_username}:{proxy_password}@{domain_name}:{proxy_port}"
    os.environ["http_proxy"] = proxy_url
    os.environ["https_proxy"] = proxy_url
    return proxy_url


def disable_proxy() -> None:
    os.environ.pop("http_proxy", None)
    os.environ.pop("https_proxy", None)


def current_proxy_info() -> dict[str, str | None]:
    return {
        "http_proxy": os.environ.get("http_proxy"),
        "https_proxy": os.environ.get("https_proxy"),
    }

