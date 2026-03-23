def log_ok(message: str) -> None:
    print(f"  [✔] {message}")


def log_err(message: str) -> None:
    print(f"  [✗] {message}")


def log_info(message: str) -> None:
    print(f"  [→] {message}")


def log_warn(message: str) -> None:
    print(f"  [!] {message}")


def section(title: str) -> None:
    print(f"\n{'─' * 55}")
    print(f"  {title}")
    print(f"{'─' * 55}")
