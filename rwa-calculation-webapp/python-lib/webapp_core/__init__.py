"""Core logic for the Dataiku category formula & comparison webapp.

Pure modules (no ``dataiku`` dependency): ``sample_parser``, ``formula_translate``,
``compute``, ``compare``. They operate on bytes / DataFrames / plain dicts so they can be
unit-tested locally.

Dataiku-aware modules: ``config_store`` (managed-folder JSON persistence) and ``auth``.
"""

__version__ = "1.0.0"
