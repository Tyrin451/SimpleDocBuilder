from importlib.metadata import PackageNotFoundError, version

from .builder import DocumentConfig, DocumentError, SimpleDocBuilder

try:
    __version__ = version("simpledocbuilder")
except PackageNotFoundError:
    __version__ = "0.1.0"

__all__ = ["SimpleDocBuilder", "DocumentConfig", "DocumentError", "__version__"]
