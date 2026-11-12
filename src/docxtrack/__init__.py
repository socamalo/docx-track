# -*- coding: utf-8 -*-
"""docxtrack - Word tracked-changes engine and CLI."""
from .core import (  # noqa: F401
    Docx,
    DocxError,
    apply_ops,
    accept_all,
    reject_all,
    build_tracked,
    verify,
)

__version__ = "0.1.0"
