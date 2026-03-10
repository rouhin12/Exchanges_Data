"""
Compatibility shim for the standard-library `imghdr` module.

Python 3.13+ removed `imghdr`, but older versions of Streamlit (like 1.19)
still import it. This lightweight shim provides a minimal `what` function
so that `import imghdr` continues to work.

The app does not rely on `imghdr` for any logic, so returning `None` is
sufficient for our usage.
"""

from __future__ import annotations

from typing import IO, Optional


def what(file: str | IO[bytes], h: Optional[bytes] = None) -> None:
    """
    Mimic imghdr.what interface.

    Always returns None, indicating the type cannot be determined.
    This is adequate because the app is not using `imghdr` directly.
    """

    return None

